"""
ai.llm.factory
===============
Selects and constructs the configured :class:`ai.llm.provider.LLMProvider`
implementation.

This is the only place that maps ``AISettings.llm_provider`` (a plain
string) to a concrete provider class. Adding a second provider means
adding one more branch here and a new file under ``ai.llm.providers`` --
no change to any caller of :func:`get_provider`.
"""

from __future__ import annotations

from ai.llm.provider import LLMProvider
from ai.settings import AISettings


class UnknownProviderError(Exception):
    """Raised when ``AISettings.llm_provider`` names a provider with no
    registered implementation."""

    def __init__(self, provider_name: str) -> None:
        super().__init__(
            f"Unknown LLM provider '{provider_name}'. "
            f"Supported providers: {sorted(_SUPPORTED_PROVIDERS)}."
        )
        self.provider_name = provider_name


_SUPPORTED_PROVIDERS = frozenset({"bedrock"})


def get_provider(settings: AISettings) -> LLMProvider:
    """Construct the :class:`LLMProvider` implementation named by
    ``settings.llm_provider``.

    Args:
        settings: Application AI settings, as produced by
            :func:`ai.settings.load_ai_settings`.

    Returns:
        A ready-to-use provider instance.

    Raises:
        UnknownProviderError: If ``settings.llm_provider`` does not name
            a supported, implemented provider.
    """
    provider_name = settings.llm_provider.lower()

    if provider_name == "bedrock":
        # Imported locally rather than at module scope so importing
        # ai.llm.factory does not require boto3 to be installed unless
        # the Bedrock provider is actually selected -- keeps the
        # provider-agnostic parts of ai/ importable in environments
        # that only ever use a different provider.
        from ai.llm.providers.bedrock import BedrockProvider

        return BedrockProvider(settings.bedrock)

    raise UnknownProviderError(settings.llm_provider)
