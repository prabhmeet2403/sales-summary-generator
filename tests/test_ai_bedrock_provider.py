"""
tests/test_ai_bedrock_provider.py
===================================
Tests ``ai.llm.providers.bedrock.BedrockProvider`` against a fake
``boto3`` client -- this is the one place in the AI test suite where
mocking is appropriate (the external network boundary), per
Architecture Plan Section 17's testing philosophy: mock the provider
boundary, never mock business logic. Everything tested here is our own
request-construction and response-parsing code; no real AWS call is
made.

Usage:
    python tests/test_ai_bedrock_provider.py
"""
from __future__ import annotations

import io
import json
import os
import sys
from dataclasses import replace
from pathlib import Path
from typing import Any, Iterator, List

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from botocore.exceptions import ClientError  # noqa: E402

from ai.llm.provider import LLMProviderError, Message, ToolSchema  # noqa: E402
from ai.llm.providers.bedrock import BedrockProvider  # noqa: E402
from ai.settings import BedrockSettings  # noqa: E402

_SETTINGS = BedrockSettings(
    region="us-east-1",
    model_id="anthropic.claude-3-test",
    embedding_model_id="amazon.titan-embed-text-v2",
    inference_profile_arn=None,
    max_tokens=4096,
    temperature=0.2,
    timeout_seconds=60,
    max_retries=3,
)


class _FakeBedrockClient:
    """Stands in for a real ``boto3`` ``bedrock-runtime`` client.

    Each test configures ``converse_response``, ``converse_stream_events``,
    ``invoke_model_response``, or ``raise_errors_then_succeed`` before
    constructing a :class:`BedrockProvider` around this fake, so the
    provider's request-building and response-parsing logic can be
    exercised deterministically.
    """

    def __init__(self) -> None:
        self.converse_response: dict = {}
        self.converse_stream_events: List[dict] = []
        self.invoke_model_response: dict = {}
        self.raise_errors_then_succeed: List[Exception] = []
        self.last_converse_kwargs: dict = {}
        self.call_count = 0

    def converse(self, **kwargs: Any) -> dict:
        self.last_converse_kwargs = kwargs
        self.call_count += 1
        if self.raise_errors_then_succeed:
            raise self.raise_errors_then_succeed.pop(0)
        return self.converse_response

    def converse_stream(self, **kwargs: Any) -> dict:
        self.last_converse_kwargs = kwargs
        self.call_count += 1
        if self.raise_errors_then_succeed:
            raise self.raise_errors_then_succeed.pop(0)
        return {"stream": iter(self.converse_stream_events)}

    def invoke_model(self, **kwargs: Any) -> dict:
        self.last_converse_kwargs = kwargs
        self.call_count += 1
        if self.raise_errors_then_succeed:
            raise self.raise_errors_then_succeed.pop(0)
        body_bytes = json.dumps(self.invoke_model_response).encode("utf-8")
        return {"body": io.BytesIO(body_bytes)}


def _client_error(code: str) -> ClientError:
    return ClientError({"Error": {"Code": code, "Message": f"synthetic {code}"}}, "Converse")


def main() -> int:
    problems: list = []

    # --- converse(): text response parsing ---
    fake = _FakeBedrockClient()
    fake.converse_response = {
        "output": {"message": {"role": "assistant", "content": [{"text": "Revenue is $5,686,374."}]}},
        "stopReason": "end_turn",
        "usage": {"inputTokens": 120, "outputTokens": 15},
    }
    provider = BedrockProvider(_SETTINGS, client=fake)
    response = provider.converse([Message(role="user", text="What is Q1 revenue?")], system_prompt="You are helpful.")
    if response.text != "Revenue is $5,686,374.":
        problems.append(f"converse() text parsing incorrect: {response.text!r}")
    if response.tool_use is not None:
        problems.append("converse() should not report a tool_use for a text-only response")
    if response.input_tokens != 120 or response.output_tokens != 15:
        problems.append("converse() did not correctly parse usage/token counts")
    if fake.last_converse_kwargs.get("modelId") != _SETTINGS.model_id:
        problems.append("converse() did not send the configured model_id")
    if fake.last_converse_kwargs.get("system") != [{"text": "You are helpful."}]:
        problems.append("converse() did not correctly translate system_prompt")

    # --- converse(): tool_use response parsing ---
    fake2 = _FakeBedrockClient()
    fake2.converse_response = {
        "output": {
            "message": {
                "role": "assistant",
                "content": [
                    {"toolUse": {"toolUseId": "tool-1", "name": "revenue_analysis", "input": {"client": "HPE"}}}
                ],
            }
        },
        "stopReason": "tool_use",
        "usage": {"inputTokens": 200, "outputTokens": 30},
    }
    provider2 = BedrockProvider(_SETTINGS, client=fake2)
    response2 = provider2.converse(
        [Message(role="user", text="Show HPE revenue")],
        tools=[ToolSchema(name="revenue_analysis", description="...", input_schema={})],
    )
    if response2.tool_use is None or response2.tool_use.name != "revenue_analysis":
        problems.append("converse() did not correctly parse a tool_use response")
    elif response2.tool_use.arguments != {"client": "HPE"}:
        problems.append(f"converse() tool_use arguments incorrect: {response2.tool_use.arguments!r}")
    if "toolConfig" not in fake2.last_converse_kwargs:
        problems.append("converse() did not include toolConfig when tools were supplied")

    # --- converse_stream(): text delta + tool use + final chunk ---
    fake3 = _FakeBedrockClient()
    fake3.converse_stream_events = [
        {"contentBlockDelta": {"delta": {"text": "Hello"}}},
        {"contentBlockDelta": {"delta": {"text": " there"}}},
        {"contentBlockStart": {"start": {"toolUse": {"toolUseId": "t1", "name": "revenue_analysis"}}}},
        {"contentBlockDelta": {"delta": {"toolUse": {"input": '{"client": '}}}},
        {"contentBlockDelta": {"delta": {"toolUse": {"input": '"HPE"}'}}}},
        {"contentBlockStop": {}},
        {"messageStop": {"stopReason": "end_turn"}},
    ]
    provider3 = BedrockProvider(_SETTINGS, client=fake3)
    chunks = list(provider3.converse_stream([Message(role="user", text="hi")]))
    text_deltas = [c.text_delta for c in chunks if c.text_delta]
    if text_deltas != ["Hello", " there"]:
        problems.append(f"converse_stream() text deltas incorrect: {text_deltas}")
    tool_use_chunks = [c for c in chunks if c.tool_use is not None]
    if len(tool_use_chunks) != 1 or tool_use_chunks[0].tool_use.arguments != {"client": "HPE"}:
        problems.append("converse_stream() did not correctly accumulate and parse streamed tool-use JSON")
    if not chunks[-1].is_final:
        problems.append("converse_stream() should yield a final chunk with is_final=True")

    # --- embed(): request/response shape ---
    fake4 = _FakeBedrockClient()
    fake4.invoke_model_response = {"embedding": [0.1, 0.2, 0.3]}
    provider4 = BedrockProvider(_SETTINGS, client=fake4)
    embeddings = provider4.embed(["hello world"])
    if len(embeddings) != 1 or list(embeddings[0]) != [0.1, 0.2, 0.3]:
        problems.append(f"embed() did not correctly parse the embedding response: {embeddings}")
    if fake4.last_converse_kwargs.get("modelId") != _SETTINGS.embedding_model_id:
        problems.append("embed() did not use the configured embedding_model_id")

    # --- retry: a retryable error followed by success ---
    fake5 = _FakeBedrockClient()
    fake5.raise_errors_then_succeed = [_client_error("ThrottlingException")]
    fake5.converse_response = {
        "output": {"message": {"content": [{"text": "ok"}]}},
        "stopReason": "end_turn",
        "usage": {},
    }
    provider5 = BedrockProvider(_SETTINGS, client=fake5)
    result5 = provider5.converse([Message(role="user", text="hi")])
    if result5.text != "ok":
        problems.append("A retryable error followed by success should still return the successful response")
    if fake5.call_count != 2:
        problems.append(f"Expected exactly 2 calls (1 failure + 1 success), got {fake5.call_count}")

    # --- retry: a non-retryable error must fail immediately, no retry ---
    fake6 = _FakeBedrockClient()
    fake6.raise_errors_then_succeed = [_client_error("ValidationException")]
    provider6 = BedrockProvider(_SETTINGS, client=fake6)
    try:
        provider6.converse([Message(role="user", text="hi")])
        problems.append("A non-retryable error should raise LLMProviderError, not succeed")
    except LLMProviderError as exc:
        if exc.retryable:
            problems.append("A ValidationException should be reported as non-retryable")
        if fake6.call_count != 1:
            problems.append(f"A non-retryable error should not be retried; got {fake6.call_count} calls")

    # --- retry: exhausting all retries raises LLMProviderError(retryable=True) ---
    fake7 = _FakeBedrockClient()
    fake7.raise_errors_then_succeed = [
        _client_error("ThrottlingException"),
        _client_error("ThrottlingException"),
        _client_error("ThrottlingException"),
    ]
    provider7 = BedrockProvider(_SETTINGS, client=fake7)
    try:
        provider7.converse([Message(role="user", text="hi")])
        problems.append("Exhausting all retries should raise LLMProviderError")
    except LLMProviderError as exc:
        if not exc.retryable:
            problems.append("Exhausted throttling retries should still be reported as retryable=True")
        if fake7.call_count != _SETTINGS.max_retries:
            problems.append(f"Expected exactly {_SETTINGS.max_retries} attempts, got {fake7.call_count}")

    # --- bearer token: _build_client exports it to the environment
    #     variable botocore itself reads (confirmed present in the
    #     installed botocore via get_token_from_environment) ---
    saved_env_value = os.environ.pop("AWS_BEARER_TOKEN_BEDROCK", None)
    try:
        settings_with_token = replace(_SETTINGS, bearer_token="test-bearer-token-xyz")
        BedrockProvider._build_client(settings_with_token)
        if os.environ.get("AWS_BEARER_TOKEN_BEDROCK") != "test-bearer-token-xyz":
            problems.append("_build_client did not export bearer_token to AWS_BEARER_TOKEN_BEDROCK")
    finally:
        if saved_env_value is not None:
            os.environ["AWS_BEARER_TOKEN_BEDROCK"] = saved_env_value
        else:
            os.environ.pop("AWS_BEARER_TOKEN_BEDROCK", None)

    # --- backward compatibility: no bearer_token -> no env var set,
    #     construction behaves exactly as before ---
    saved_env_value = os.environ.pop("AWS_BEARER_TOKEN_BEDROCK", None)
    try:
        BedrockProvider._build_client(_SETTINGS)  # bearer_token defaults to None
        if "AWS_BEARER_TOKEN_BEDROCK" in os.environ:
            problems.append("_build_client should not set AWS_BEARER_TOKEN_BEDROCK when bearer_token is None")
    finally:
        if saved_env_value is not None:
            os.environ["AWS_BEARER_TOKEN_BEDROCK"] = saved_env_value

    # --- constructing BedrockSettings the original (pre-bearer-token) way
    #     still works -- positional/keyword construction unaffected ---
    legacy_settings = BedrockSettings(
        region="us-east-1", model_id="m", embedding_model_id="e",
        inference_profile_arn=None, max_tokens=100, temperature=0.1,
        timeout_seconds=10, max_retries=1,
    )
    if legacy_settings.bearer_token is not None:
        problems.append("BedrockSettings constructed without bearer_token should default it to None")

    # --- embed() raises a clear, actionable error when no embedding
    #     model is configured, rather than passing None to boto3 ---
    settings_no_embedding_model = replace(_SETTINGS, embedding_model_id=None)
    provider_no_embedding = BedrockProvider(settings_no_embedding_model, client=_FakeBedrockClient())
    try:
        provider_no_embedding.embed(["hello"])
        problems.append("embed() without embedding_model_id configured should raise LLMProviderError")
    except LLMProviderError as exc:
        if "BEDROCK_EMBEDDING_MODEL_ID" not in exc.message:
            problems.append(f"embed() error message should name the missing setting; got: {exc.message}")

    if problems:
        print("\nFAILURES:")
        for p in problems:
            print(f"  - {p}")
        print(f"\nFAIL - {len(problems)} problem(s).")
        return 1

    print("ALL BEDROCK PROVIDER CHECKS PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
