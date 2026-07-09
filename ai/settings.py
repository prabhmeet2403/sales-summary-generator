"""
ai.settings
===========
Centralized configuration for the AI platform.

Every value that varies by environment or deployment (region, model ID,
timeouts, rate limits) is read from environment variables through this
module, and only through this module -- no other file in ``ai`` calls
``os.environ`` for AI-related configuration. This is what makes it
possible to swap models, regions, or providers without touching any
planner, tool, or provider-implementation code, and to unit-test any of
those consumers against a fake ``AISettings`` instance instead of
mutating the real process environment.

Credentials are never read here for standard AWS IAM authentication
(access key, secret key, session token) -- that continues to be
resolved entirely by ``boto3``'s own default credential chain
(environment variables, a shared credentials file, or an attached IAM
role), and this module never touches those values.

The one exception, by design: Amazon Bedrock also supports a bearer
token authentication mode (the ``AWS_BEARER_TOKEN_BEDROCK`` environment
variable), which AWS's own SDKs read directly from the process
environment. This module reads it too -- centralizing awareness of it
here, consistent with every other Bedrock setting -- but never logs its
value, only whether one is configured.

See ``Phase2_AI_Assistant_Architecture_Plan_v3.md`` Sections 7, 9, and
11.2 for the approved design this module implements.
"""

from __future__ import annotations

import logging
import os
from dataclasses import dataclass
from typing import Mapping, Optional

logger = logging.getLogger(__name__)

#: Environment variable names, centralized as constants so a typo in one
#: place cannot silently create a second, inconsistent name for the same
#: setting elsewhere in the codebase.
ENV_LLM_PROVIDER = "LLM_PROVIDER"
ENV_BEDROCK_REGION = "BEDROCK_REGION"
ENV_BEDROCK_MODEL_ID = "BEDROCK_MODEL_ID"
ENV_BEDROCK_EMBEDDING_MODEL_ID = "BEDROCK_EMBEDDING_MODEL_ID"
ENV_BEDROCK_INFERENCE_PROFILE_ARN = "BEDROCK_INFERENCE_PROFILE_ARN"
ENV_BEDROCK_MAX_TOKENS = "BEDROCK_MAX_TOKENS"
ENV_BEDROCK_TEMPERATURE = "BEDROCK_TEMPERATURE"
ENV_BEDROCK_TIMEOUT_SECONDS = "BEDROCK_TIMEOUT_SECONDS"
ENV_BEDROCK_MAX_RETRIES = "BEDROCK_MAX_RETRIES"
ENV_AI_RATE_LIMIT_PER_MINUTE = "AI_RATE_LIMIT_PER_MINUTE"
#: AWS's own standard name for the Bedrock bearer-token env var --
#: not a "BEDROCK_"-prefixed name of this project's own choosing, since
#: this must match exactly what AWS's SDKs (and this project's own
#: boto3 client) read.
ENV_AWS_BEARER_TOKEN_BEDROCK = "AWS_BEARER_TOKEN_BEDROCK"

#: Non-secret, non-environment-specific defaults. These are deliberately
#: limited to values that are safe to ship as fallbacks (token limits,
#: timeouts, temperature) -- never a region or model ID, which must
#: always come from the environment so a deployment can never silently
#: run against the wrong model or region by omission.
_DEFAULT_LLM_PROVIDER = "bedrock"
_DEFAULT_MAX_TOKENS = 4096
_DEFAULT_TEMPERATURE = 0.2
_DEFAULT_TIMEOUT_SECONDS = 60
_DEFAULT_MAX_RETRIES = 3
_DEFAULT_RATE_LIMIT_PER_MINUTE = 30


class AISettingsError(Exception):
    """Raised when required AI configuration is missing or invalid.

    Carries a plain-language ``message`` naming exactly which
    environment variable is missing or malformed, so a misconfigured
    deployment fails fast with an actionable error instead of a
    ``KeyError``/``ValueError`` traceback pointing at an unrelated line
    deep inside a provider implementation.
    """

    def __init__(self, message: str) -> None:
        super().__init__(message)
        self.message = message


@dataclass(frozen=True)
class BedrockSettings:
    """Configuration for the Amazon Bedrock provider.

    Attributes:
        region: AWS region hosting the Bedrock endpoint, e.g. ``"us-east-1"``.
        model_id: Bedrock model identifier used for chat/tool-use calls.
        embedding_model_id: Bedrock model identifier used for text
            embeddings. Optional -- no code path in this version of the
            application calls ``LLMProvider.embed()`` yet (semantic
            search is not implemented); this is only enforced once an
            embedding-calling feature actually exists.
        inference_profile_arn: Optional cross-region inference profile
            ARN. When set, providers should prefer it over ``model_id``
            for invocation, per Bedrock's cross-region inference feature.
        max_tokens: Default maximum output tokens per response.
        temperature: Default sampling temperature.
        timeout_seconds: Per-request timeout passed to the underlying
            HTTP client.
        max_retries: Maximum retry attempts for transient/throttling
            errors (see ``ai.llm.providers.bedrock``).
        bearer_token: Optional Bedrock API key for bearer-token
            authentication (``AWS_BEARER_TOKEN_BEDROCK``), used instead
            of standard AWS IAM credentials when set. ``None`` when not
            configured, in which case the standard boto3 credential
            chain is used exactly as before -- this field is additive
            and does not change behavior when absent. Never logged;
            code that needs to report configuration state should log
            ``bearer_token is not None``, not the value itself.
    """

    region: str
    model_id: str
    inference_profile_arn: Optional[str]
    max_tokens: int
    temperature: float
    timeout_seconds: int
    max_retries: int
    embedding_model_id: Optional[str] = None
    bearer_token: Optional[str] = None


@dataclass(frozen=True)
class AISettings:
    """Top-level AI platform configuration.

    Attributes:
        llm_provider: Selects which ``ai.llm.provider.LLMProvider``
            implementation ``ai.llm.factory.get_provider`` constructs.
            Only ``"bedrock"`` is implemented as of Phase 2a; the field
            exists so adding a second provider later is a new
            implementation file plus a value for this setting, not a
            structural change (see Architecture Plan Section 11.2).
        bedrock: Bedrock-specific settings. Present regardless of
            ``llm_provider`` so it is always available for the
            embedding calls semantic search will use in Phase 2c
            (embeddings may use Bedrock even if chat uses a different
            provider in the future).
        request_rate_limit_per_minute: Application-level cap on AI
            requests per session, per minute, independent of whatever
            throttling the provider itself enforces (see Architecture
            Plan Section 15, "Rate limiting / abuse prevention").
    """

    llm_provider: str
    bedrock: BedrockSettings
    request_rate_limit_per_minute: int


def _get_str(env: Mapping[str, str], name: str, *, required: bool, default: Optional[str] = None) -> Optional[str]:
    """Read a string environment variable, raising a clear error if a
    required value is absent, or returning ``default`` otherwise."""
    value = env.get(name)
    if value is None or value.strip() == "":
        if required:
            raise AISettingsError(
                f"Required environment variable '{name}' is not set. "
                f"Set it before starting the application (see ai/README.md)."
            )
        return default
    return value


def _get_int(env: Mapping[str, str], name: str, default: int) -> int:
    """Read an integer environment variable, falling back to ``default``
    when absent, and raising a clear error when present but not a valid
    integer (rather than letting a bad value silently become 0 or crash
    somewhere unrelated later)."""
    raw = env.get(name)
    if raw is None or raw.strip() == "":
        return default
    try:
        return int(raw)
    except ValueError as exc:
        raise AISettingsError(
            f"Environment variable '{name}' must be an integer, got {raw!r}."
        ) from exc


def _get_float(env: Mapping[str, str], name: str, default: float) -> float:
    """Read a float environment variable, with the same fail-fast
    behavior as :func:`_get_int`."""
    raw = env.get(name)
    if raw is None or raw.strip() == "":
        return default
    try:
        return float(raw)
    except ValueError as exc:
        raise AISettingsError(
            f"Environment variable '{name}' must be a number, got {raw!r}."
        ) from exc


def load_ai_settings(env: Optional[Mapping[str, str]] = None) -> AISettings:
    """Load :class:`AISettings` from environment variables.

    Args:
        env: Mapping to read variables from. Defaults to ``os.environ``.
            Accepting this as a parameter (rather than reading
            ``os.environ`` unconditionally) is what makes this function
            unit-testable without mutating the real process
            environment -- tests pass a plain ``dict`` instead.

    Returns:
        A fully populated, immutable :class:`AISettings`.

    Raises:
        AISettingsError: If a required variable (``BEDROCK_REGION``,
            ``BEDROCK_MODEL_ID``) is missing, or if a numeric variable
            is present but malformed. ``BEDROCK_EMBEDDING_MODEL_ID`` is
            optional -- see :attr:`BedrockSettings.embedding_model_id`.
    """
    resolved_env: Mapping[str, str] = env if env is not None else os.environ

    llm_provider = _get_str(resolved_env, ENV_LLM_PROVIDER, required=False, default=_DEFAULT_LLM_PROVIDER)
    assert llm_provider is not None  # default is always provided above

    bedrock_settings = BedrockSettings(
        region=_get_str(resolved_env, ENV_BEDROCK_REGION, required=True),  # type: ignore[arg-type]
        model_id=_get_str(resolved_env, ENV_BEDROCK_MODEL_ID, required=True),  # type: ignore[arg-type]
        inference_profile_arn=_get_str(
            resolved_env, ENV_BEDROCK_INFERENCE_PROFILE_ARN, required=False
        ),
        max_tokens=_get_int(resolved_env, ENV_BEDROCK_MAX_TOKENS, _DEFAULT_MAX_TOKENS),
        temperature=_get_float(resolved_env, ENV_BEDROCK_TEMPERATURE, _DEFAULT_TEMPERATURE),
        timeout_seconds=_get_int(resolved_env, ENV_BEDROCK_TIMEOUT_SECONDS, _DEFAULT_TIMEOUT_SECONDS),
        max_retries=_get_int(resolved_env, ENV_BEDROCK_MAX_RETRIES, _DEFAULT_MAX_RETRIES),
        embedding_model_id=_get_str(resolved_env, ENV_BEDROCK_EMBEDDING_MODEL_ID, required=False),
        bearer_token=_get_str(resolved_env, ENV_AWS_BEARER_TOKEN_BEDROCK, required=False),
    )

    settings = AISettings(
        llm_provider=llm_provider,
        bedrock=bedrock_settings,
        request_rate_limit_per_minute=_get_int(
            resolved_env, ENV_AI_RATE_LIMIT_PER_MINUTE, _DEFAULT_RATE_LIMIT_PER_MINUTE
        ),
    )

    logger.debug(
        "AI settings loaded: provider=%s region=%s model_id=%s embedding_model_id=%s "
        "max_tokens=%d temperature=%.2f timeout=%ds max_retries=%d rate_limit=%d/min "
        "auth_mode=%s",
        settings.llm_provider,
        settings.bedrock.region,
        settings.bedrock.model_id,
        settings.bedrock.embedding_model_id,
        settings.bedrock.max_tokens,
        settings.bedrock.temperature,
        settings.bedrock.timeout_seconds,
        settings.bedrock.max_retries,
        settings.request_rate_limit_per_minute,
        "bearer_token" if settings.bedrock.bearer_token else "default_credential_chain",
    )
    return settings
