"""
ai.llm.provider
================
The provider-agnostic interface every LLM backend implements.

This is the single abstraction that makes "the application uses Amazon
Bedrock" a configuration fact rather than a structural one. No code
outside ``ai.llm.providers`` may import a provider SDK (``boto3`` or any
future equivalent) or reference a provider by name -- every caller
(the service facade today; the workflow graph and tools from Phase 2b
onward) depends on :class:`LLMProvider` and the message/response types
defined here, never on a concrete implementation.

See ``Phase2_AI_Assistant_Architecture_Plan_v3.md`` Section 11.2.
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Iterator, List, Literal, Optional

import numpy as np

#: The two conversational roles this platform's messages can carry.
#: Tool-result messages (introduced in Phase 2b, once tools exist) are
#: represented as "user" role messages whose content encodes the tool
#: result, per the common Converse-API-style convention -- this keeps
#: the role vocabulary stable as tool-use is added rather than growing
#: a third role later.
Role = Literal["user", "assistant"]


@dataclass(frozen=True)
class Message:
    """One turn in a conversation sent to an :class:`LLMProvider`.

    Exactly one of ``text``, ``tool_use``, or ``tool_result`` is
    expected to be set, matching the three content shapes a tool-use
    conversation protocol (Bedrock's Converse API, and tool-calling
    protocols generally) requires:

    - A plain text turn (either role).
    - An assistant turn echoing back a tool call the model previously
      requested (``tool_use`` set) -- required when continuing a
      conversation after a tool-use response, so the model sees its own
      prior request in context.
    - A user turn reporting a tool's result (``tool_result`` set),
      linked to the original request via ``tool_use_id`` -- this is
      what a ``tool_use`` stop reason must be followed by; a plain text
      follow-up message does not satisfy the protocol and will be
      rejected by a real Bedrock endpoint.

    Attributes:
        role: Who produced this turn.
        text: The turn's plain text content, if this is a text turn.
        tool_use: The tool call this turn represents, if this is an
            assistant turn echoing back a prior tool-use request.
        tool_result: The tool result this turn reports, if this is a
            user turn responding to a tool-use request.
    """

    role: Role
    text: Optional[str] = None
    tool_use: Optional["ToolUseRequest"] = None
    tool_result: Optional["ToolResultBlock"] = None


@dataclass(frozen=True)
class ToolResultBlock:
    """A tool's result, reported back to the model as part of a
    ``Message`` with ``role="user"``.

    Attributes:
        tool_use_id: The ``tool_use_id`` from the
            :class:`ToolUseRequest` this result answers.
        text: The tool's result, as text (typically a
            :class:`~ai.tools.base.ToolResult`'s ``summary``).
    """

    tool_use_id: str
    text: str


@dataclass(frozen=True)
class ToolSchema:
    """The description of one callable tool, as presented to the model.

    Not used by any caller until Phase 2b (no tools exist yet in Phase
    2a), but part of the interface now so :meth:`LLMProvider.converse`
    and :meth:`LLMProvider.converse_stream` never need a breaking
    signature change when tools are introduced.

    Attributes:
        name: The tool's unique name, matching a
            ``ai.tools.base.BaseTool.name`` (Phase 2b).
        description: A natural-language description the model uses to
            decide when this tool is relevant.
        input_schema: A JSON Schema object describing the tool's
            expected arguments.
    """

    name: str
    description: str
    input_schema: dict


@dataclass(frozen=True)
class ToolUseRequest:
    """A request from the model to invoke a specific tool.

    Attributes:
        tool_use_id: Provider-assigned identifier for this specific
            invocation, required to correlate a later tool result back
            to this request.
        name: The requested tool's name.
        arguments: The arguments the model supplied, already parsed
            from the provider's wire format into a plain ``dict``.
    """

    tool_use_id: str
    name: str
    arguments: dict


@dataclass(frozen=True)
class Response:
    """A complete, non-streaming response from an :class:`LLMProvider`.

    Exactly one of ``text`` or ``tool_use`` is expected to be
    meaningful for a given ``stop_reason`` (a text-completion response
    populates ``text``; a tool-use response populates ``tool_use``),
    but both fields are always present (as ``None`` when not
    applicable) so callers can pattern-match without hasattr checks.

    Attributes:
        text: The model's text output, if this response is a final or
            partial text completion.
        tool_use: The model's requested tool invocation, if this
            response is a tool-use request.
        stop_reason: Provider-reported reason generation stopped
            (e.g. ``"end_turn"``, ``"tool_use"``, ``"max_tokens"``).
        input_tokens: Number of input tokens billed for this request,
            for cost tracking and the audit log (Phase 2d).
        output_tokens: Number of output tokens billed for this request.
    """

    text: Optional[str]
    tool_use: Optional[ToolUseRequest]
    stop_reason: str
    input_tokens: int
    output_tokens: int


@dataclass(frozen=True)
class StreamChunk:
    """One incremental piece of a streamed response.

    Attributes:
        text_delta: A new fragment of text to append to what has
            already been streamed, or ``None`` if this chunk carries no
            text (e.g. a tool-use chunk).
        tool_use: A completed tool-use request, populated only on the
            chunk that finalizes it.
        is_final: ``True`` on the last chunk of the stream.
    """

    text_delta: Optional[str] = None
    tool_use: Optional[ToolUseRequest] = None
    is_final: bool = False


class LLMProviderError(Exception):
    """Raised when a provider call fails in a way callers must handle
    explicitly (throttling, timeout, malformed response) rather than a
    silently swallowed or fabricated result.

    Attributes:
        message: Plain-language description of the failure.
        retryable: Whether the caller may reasonably retry this exact
            request (e.g. ``True`` for throttling, ``False`` for a
            malformed-request error that will fail identically again).
    """

    def __init__(self, message: str, *, retryable: bool = False) -> None:
        super().__init__(message)
        self.message = message
        self.retryable = retryable


class LLMProvider(ABC):
    """Provider-agnostic interface for conversational and embedding LLM calls.

    Every method is provider-agnostic in its signature: callers never
    see a provider-specific request or response shape. A concrete
    implementation (e.g. :class:`ai.llm.providers.bedrock.BedrockProvider`)
    is responsible for translating to and from its own backend's wire
    format entirely internally.
    """

    @abstractmethod
    def converse(
        self,
        messages: List[Message],
        *,
        system_prompt: str = "",
        tools: Optional[List[ToolSchema]] = None,
    ) -> Response:
        """Send a conversation and get one complete response.

        Args:
            messages: The conversation so far, oldest first.
            system_prompt: Instructions prepended as system context.
            tools: Tool schemas the model may choose to invoke. An
                empty list or ``None`` means no tools are offered.

        Returns:
            The model's response.

        Raises:
            LLMProviderError: On a request failure the caller must
                handle (see the class docstring on ``retryable``).
        """
        raise NotImplementedError

    @abstractmethod
    def converse_stream(
        self,
        messages: List[Message],
        *,
        system_prompt: str = "",
        tools: Optional[List[ToolSchema]] = None,
    ) -> Iterator[StreamChunk]:
        """Send a conversation and stream the response incrementally.

        Args:
            messages: The conversation so far, oldest first.
            system_prompt: Instructions prepended as system context.
            tools: Tool schemas the model may choose to invoke.

        Yields:
            :class:`StreamChunk` instances; the final chunk has
            ``is_final=True``.

        Raises:
            LLMProviderError: On a request failure the caller must
                handle.
        """
        raise NotImplementedError

    @abstractmethod
    def embed(self, texts: List[str]) -> List[np.ndarray]:
        """Compute embedding vectors for a batch of texts.

        Not used until semantic search is implemented (Phase 2c); part
        of the interface now so the provider abstraction is complete
        against the approved architecture from the start.

        Args:
            texts: The strings to embed.

        Returns:
            One embedding vector per input string, in the same order.

        Raises:
            LLMProviderError: On a request failure the caller must
                handle.
        """
        raise NotImplementedError
