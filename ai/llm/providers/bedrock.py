"""
ai.llm.providers.bedrock
=========================
Amazon Bedrock implementation of :class:`ai.llm.provider.LLMProvider`.

This is the only module in the entire ``ai`` package permitted to import
``boto3`` or reference Bedrock's request/response wire format. Every
other consumer -- the service facade today, the workflow graph and
tools from Phase 2b onward -- depends only on the provider-agnostic
types in ``ai.llm.provider``.

Uses Bedrock's Converse API (``converse`` / ``converse_stream``) rather
than the older, model-specific ``invoke_model`` API for chat, because
Converse provides one stable request/response shape with native
tool-use support across model families -- see Architecture Plan
Section 7. Embeddings use ``invoke_model`` directly, since Converse
does not cover embedding models.

Credentials are never read by this module for standard AWS IAM
authentication. ``boto3.client(...)`` is constructed with only a
region name for that path; authentication is resolved entirely by
boto3's own default credential chain.

This module additionally supports Amazon Bedrock's bearer-token
authentication mode (``AWS_BEARER_TOKEN_BEDROCK``), used instead of
standard AWS IAM credentials when configured. Precedence and fallback
are handled by AWS's own SDK: when the bearer token environment
variable is present, botocore uses it for Bedrock service calls
automatically; when absent, the standard credential chain is used
exactly as before. See :meth:`BedrockProvider._build_client`.
"""

from __future__ import annotations

import json
import logging
import os
import time
from typing import Any, Iterator, List, Optional

import numpy as np

try:
    import boto3
    from botocore.config import Config as BotoConfig
    from botocore.exceptions import BotoCoreError, ClientError
except ImportError as exc:  # pragma: no cover - exercised only when boto3 is absent
    raise ImportError(
        "boto3 is required to use ai.llm.providers.bedrock.BedrockProvider. "
        "Install it with 'pip install boto3' (see requirements.txt)."
    ) from exc

from ai.llm.provider import (
    LLMProvider,
    LLMProviderError,
    Message,
    Response,
    StreamChunk,
    ToolResultBlock,
    ToolSchema,
    ToolUseRequest,
)
from ai.settings import BedrockSettings

logger = logging.getLogger(__name__)

#: AWS's own standard name for the Bedrock bearer-token env var --
#: must match exactly what botocore itself reads.
_ENV_AWS_BEARER_TOKEN_BEDROCK = "AWS_BEARER_TOKEN_BEDROCK"

#: Bedrock error codes that represent a transient condition worth
#: retrying (throttling, momentary service unavailability) as opposed
#: to a request that will fail identically on every retry (bad
#: arguments, access denied, unknown model).
_RETRYABLE_ERROR_CODES = frozenset(
    {
        "ThrottlingException",
        "ServiceUnavailableException",
        "ModelTimeoutException",
        "InternalServerException",
    }
)


class BedrockProvider(LLMProvider):
    """Amazon Bedrock-backed :class:`LLMProvider`.

    Attributes:
        settings: The Bedrock configuration this provider was
            constructed with.
    """

    def __init__(self, settings: BedrockSettings, *, client: Optional[Any] = None) -> None:
        """Construct a Bedrock-backed provider.

        Args:
            settings: Region, model IDs, and request tuning parameters.
            client: An already-constructed ``boto3`` ``bedrock-runtime``
                client. Exposed as a constructor parameter specifically
                so tests can inject a fake client instead of talking to
                real AWS -- production callers should omit this and let
                the provider construct its own client from ``settings``.
        """
        self.settings = settings
        self._client = client if client is not None else self._build_client(settings)

    @staticmethod
    def _build_client(settings: BedrockSettings) -> Any:
        """Construct a ``boto3`` ``bedrock-runtime`` client for the
        configured region, with a request timeout matching
        ``settings.timeout_seconds``.

        Authentication: if ``settings.bearer_token`` is set, it is
        exported to the ``AWS_BEARER_TOKEN_BEDROCK`` environment
        variable (the name AWS's own SDKs read) so botocore
        authenticates Bedrock calls with it instead of standard AWS IAM
        credentials -- this also covers the case where the token was
        supplied through this project's own settings object (e.g. from
        a secrets manager) rather than already being present in the
        process environment. When ``bearer_token`` is not set, nothing
        changes here and boto3's default credential chain resolves
        authentication exactly as before.
        """
        if settings.bearer_token:
            os.environ[_ENV_AWS_BEARER_TOKEN_BEDROCK] = settings.bearer_token
            logger.debug("Bedrock authentication: using bearer token.")
        else:
            logger.debug("Bedrock authentication: using the default AWS credential chain.")

        boto_config = BotoConfig(
            region_name=settings.region,
            read_timeout=settings.timeout_seconds,
            connect_timeout=settings.timeout_seconds,
        )
        return boto3.client("bedrock-runtime", config=boto_config)

    # ------------------------------------------------------------------
    # Public interface
    # ------------------------------------------------------------------
    def converse(
        self,
        messages: List[Message],
        *,
        system_prompt: str = "",
        tools: Optional[List[ToolSchema]] = None,
    ) -> Response:
        request = self._build_converse_request(messages, system_prompt=system_prompt, tools=tools)
        raw_response = self._call_with_retries(
            lambda: self._client.converse(**request),
            operation_name="converse",
        )
        return self._parse_converse_response(raw_response)

    def converse_stream(
        self,
        messages: List[Message],
        *,
        system_prompt: str = "",
        tools: Optional[List[ToolSchema]] = None,
    ) -> Iterator[StreamChunk]:
        request = self._build_converse_request(messages, system_prompt=system_prompt, tools=tools)
        raw_response = self._call_with_retries(
            lambda: self._client.converse_stream(**request),
            operation_name="converse_stream",
        )
        yield from self._parse_stream_events(raw_response["stream"])

    def embed(self, texts: List[str]) -> List[np.ndarray]:
        # Bedrock's embedding models (e.g. Amazon Titan Text Embeddings)
        # are invoked one text at a time via invoke_model; there is no
        # batch-embedding operation in the API to call instead.
        if not self.settings.embedding_model_id:
            raise LLMProviderError(
                "embed() was called but BEDROCK_EMBEDDING_MODEL_ID is not configured. "
                "Set it before using any feature that requires embeddings.",
                retryable=False,
            )
        embeddings: List[np.ndarray] = []
        for text in texts:
            body = json.dumps({"inputText": text})
            raw_response = self._call_with_retries(
                lambda body=body: self._client.invoke_model(
                    modelId=self.settings.embedding_model_id,
                    body=body,
                ),
                operation_name="invoke_model (embedding)",
            )
            payload = json.loads(raw_response["body"].read())
            embeddings.append(np.array(payload["embedding"], dtype=np.float32))
        return embeddings

    # ------------------------------------------------------------------
    # Request construction
    # ------------------------------------------------------------------
    def _build_converse_request(
        self,
        messages: List[Message],
        *,
        system_prompt: str,
        tools: Optional[List[ToolSchema]],
    ) -> dict:
        """Translate provider-agnostic messages/tools into a Bedrock
        Converse API request dict."""
        request: dict = {
            "modelId": self.settings.inference_profile_arn or self.settings.model_id,
            "messages": [self._message_to_bedrock(m) for m in messages],
            "inferenceConfig": {
                "maxTokens": self.settings.max_tokens,
                "temperature": self.settings.temperature,
            },
        }
        if system_prompt:
            request["system"] = [{"text": system_prompt}]
        if tools:
            request["toolConfig"] = {
                "tools": [self._tool_schema_to_bedrock(t) for t in tools]
            }
        return request

    @staticmethod
    def _message_to_bedrock(message: Message) -> dict:
        """Translate one :class:`Message` into a Bedrock Converse API
        message dict, handling all three content shapes :class:`Message`
        supports (see its docstring for why all three are necessary for
        a correct tool-use round-trip)."""
        if message.tool_use is not None:
            return {
                "role": message.role,
                "content": [
                    {
                        "toolUse": {
                            "toolUseId": message.tool_use.tool_use_id,
                            "name": message.tool_use.name,
                            "input": message.tool_use.arguments,
                        }
                    }
                ],
            }
        if message.tool_result is not None:
            return {
                "role": message.role,
                "content": [
                    {
                        "toolResult": {
                            "toolUseId": message.tool_result.tool_use_id,
                            "content": [{"text": message.tool_result.text}],
                        }
                    }
                ],
            }
        return {"role": message.role, "content": [{"text": message.text or ""}]}

    @staticmethod
    def _tool_schema_to_bedrock(tool: ToolSchema) -> dict:
        return {
            "toolSpec": {
                "name": tool.name,
                "description": tool.description,
                "inputSchema": {"json": tool.input_schema},
            }
        }

    # ------------------------------------------------------------------
    # Response parsing
    # ------------------------------------------------------------------
    @staticmethod
    def _parse_converse_response(raw_response: dict) -> Response:
        message = raw_response["output"]["message"]
        text: Optional[str] = None
        tool_use: Optional[ToolUseRequest] = None
        for block in message.get("content", []):
            if "text" in block:
                text = (text or "") + block["text"]
            elif "toolUse" in block:
                block_tool_use = block["toolUse"]
                tool_use = ToolUseRequest(
                    tool_use_id=block_tool_use["toolUseId"],
                    name=block_tool_use["name"],
                    arguments=block_tool_use.get("input", {}),
                )
        usage = raw_response.get("usage", {})
        return Response(
            text=text,
            tool_use=tool_use,
            stop_reason=raw_response.get("stopReason", "unknown"),
            input_tokens=usage.get("inputTokens", 0),
            output_tokens=usage.get("outputTokens", 0),
        )

    @staticmethod
    def _parse_stream_events(stream: Any) -> Iterator[StreamChunk]:
        """Translate a Bedrock Converse streaming event sequence into
        :class:`StreamChunk` instances.

        Bedrock streams tool-use arguments as incremental JSON-string
        fragments (``contentBlockDelta`` events with ``toolUse.input``
        deltas); this accumulates those fragments per content block and
        only yields a completed :class:`ToolUseRequest` once the block
        closes (``contentBlockStop``), since a partially-parsed JSON
        fragment is not a usable tool call.
        """
        pending_tool_name: Optional[str] = None
        pending_tool_use_id: Optional[str] = None
        pending_tool_input_json = ""

        for event in stream:
            if "contentBlockStart" in event:
                start = event["contentBlockStart"].get("start", {})
                if "toolUse" in start:
                    pending_tool_name = start["toolUse"]["name"]
                    pending_tool_use_id = start["toolUse"]["toolUseId"]
                    pending_tool_input_json = ""

            elif "contentBlockDelta" in event:
                delta = event["contentBlockDelta"].get("delta", {})
                if "text" in delta:
                    yield StreamChunk(text_delta=delta["text"])
                elif "toolUse" in delta:
                    pending_tool_input_json += delta["toolUse"].get("input", "")

            elif "contentBlockStop" in event:
                if pending_tool_name is not None:
                    try:
                        arguments = json.loads(pending_tool_input_json) if pending_tool_input_json else {}
                    except json.JSONDecodeError:
                        logger.warning(
                            "Bedrock streamed a malformed tool-use argument payload for "
                            "tool '%s'; treating arguments as empty.",
                            pending_tool_name,
                        )
                        arguments = {}
                    yield StreamChunk(
                        tool_use=ToolUseRequest(
                            tool_use_id=pending_tool_use_id or "",
                            name=pending_tool_name,
                            arguments=arguments,
                        )
                    )
                    pending_tool_name = None
                    pending_tool_use_id = None
                    pending_tool_input_json = ""

            elif "messageStop" in event:
                yield StreamChunk(is_final=True)

    # ------------------------------------------------------------------
    # Retry / error handling
    # ------------------------------------------------------------------
    def _call_with_retries(self, operation: Any, *, operation_name: str) -> Any:
        """Invoke a zero-argument callable wrapping a boto3 call, retrying
        transient failures with exponential backoff.

        Args:
            operation: A zero-argument callable performing the boto3 call.
            operation_name: Human-readable operation name, used only in
                log messages and error text.

        Returns:
            The boto3 call's return value.

        Raises:
            LLMProviderError: If every retry attempt is exhausted, or if
                the failure is not retryable.
        """
        last_error: Optional[Exception] = None
        for attempt in range(1, self.settings.max_retries + 1):
            try:
                return operation()
            except ClientError as exc:
                error_code = exc.response.get("Error", {}).get("Code", "")
                retryable = error_code in _RETRYABLE_ERROR_CODES
                last_error = exc
                logger.warning(
                    "Bedrock %s failed on attempt %d/%d (code=%s, retryable=%s): %s",
                    operation_name,
                    attempt,
                    self.settings.max_retries,
                    error_code,
                    retryable,
                    exc,
                )
                if not retryable or attempt == self.settings.max_retries:
                    raise LLMProviderError(
                        f"Bedrock {operation_name} failed: {exc}",
                        retryable=retryable,
                    ) from exc
                time.sleep(2 ** (attempt - 1))
            except BotoCoreError as exc:
                # Network-level failures (timeouts, connection errors) --
                # always worth a bounded retry.
                last_error = exc
                logger.warning(
                    "Bedrock %s failed on attempt %d/%d (network error): %s",
                    operation_name,
                    attempt,
                    self.settings.max_retries,
                    exc,
                )
                if attempt == self.settings.max_retries:
                    raise LLMProviderError(
                        f"Bedrock {operation_name} failed after {attempt} attempts: {exc}",
                        retryable=True,
                    ) from exc
                time.sleep(2 ** (attempt - 1))

        # Unreachable in practice (the loop above always returns or
        # raises), but keeps type checkers satisfied and fails loudly
        # rather than returning None if the loop's invariants are ever
        # violated by a future edit.
        raise LLMProviderError(
            f"Bedrock {operation_name} failed after {self.settings.max_retries} attempts: {last_error}",
            retryable=False,
        )
