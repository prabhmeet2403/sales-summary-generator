"""
tests/test_ai_settings.py
===========================
Tests ``ai.settings.load_ai_settings``: required variables raise a
clear error when absent, optional variables fall back to documented
defaults, and numeric parsing fails fast and clearly on malformed
input. Uses an injected ``dict`` rather than mutating the real process
environment, per ``load_ai_settings``'s own design for testability.

Usage:
    python tests/test_ai_settings.py
"""
from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from ai.settings import AISettingsError, load_ai_settings  # noqa: E402

_MINIMAL_VALID_ENV = {
    "BEDROCK_REGION": "us-east-1",
    "BEDROCK_MODEL_ID": "anthropic.claude-3-test",
    "BEDROCK_EMBEDDING_MODEL_ID": "amazon.titan-embed-text-v2",
}


def main() -> int:
    problems: list = []

    # --- happy path: minimal required variables, defaults for the rest ---
    settings = load_ai_settings(_MINIMAL_VALID_ENV)
    if settings.llm_provider != "bedrock":
        problems.append(f"Default llm_provider should be 'bedrock', got {settings.llm_provider!r}")
    if settings.bedrock.region != "us-east-1":
        problems.append("region not read correctly")
    if settings.bedrock.model_id != "anthropic.claude-3-test":
        problems.append("model_id not read correctly")
    if settings.bedrock.max_tokens != 4096:
        problems.append(f"Default max_tokens should be 4096, got {settings.bedrock.max_tokens}")
    if settings.bedrock.temperature != 0.2:
        problems.append(f"Default temperature should be 0.2, got {settings.bedrock.temperature}")
    if settings.bedrock.inference_profile_arn is not None:
        problems.append("inference_profile_arn should default to None when unset")

    # --- required variables missing: must raise AISettingsError, not KeyError ---
    for missing_var in ("BEDROCK_REGION", "BEDROCK_MODEL_ID"):
        env = {k: v for k, v in _MINIMAL_VALID_ENV.items() if k != missing_var}
        try:
            load_ai_settings(env)
            problems.append(f"Missing {missing_var} should raise AISettingsError but did not")
        except AISettingsError as exc:
            if missing_var not in str(exc):
                problems.append(f"Error message for missing {missing_var} should name it; got: {exc}")

    # --- BEDROCK_EMBEDDING_MODEL_ID is optional: absent should NOT raise,
    #     since no current code path calls LLMProvider.embed() ---
    env_without_embedding = {k: v for k, v in _MINIMAL_VALID_ENV.items() if k != "BEDROCK_EMBEDDING_MODEL_ID"}
    try:
        settings_no_embedding = load_ai_settings(env_without_embedding)
        if settings_no_embedding.bedrock.embedding_model_id is not None:
            problems.append("embedding_model_id should default to None when unset")
    except AISettingsError:
        problems.append("BEDROCK_EMBEDDING_MODEL_ID being unset should not raise AISettingsError")

    # --- BEDROCK_EMBEDDING_MODEL_ID still works when explicitly provided ---
    settings_with_embedding = load_ai_settings(_MINIMAL_VALID_ENV)
    if settings_with_embedding.bedrock.embedding_model_id != "amazon.titan-embed-text-v2":
        problems.append("BEDROCK_EMBEDDING_MODEL_ID was not read correctly when present")

    # --- overriding a default via the environment ---
    env_with_overrides = dict(_MINIMAL_VALID_ENV)
    env_with_overrides["BEDROCK_MAX_TOKENS"] = "8000"
    env_with_overrides["BEDROCK_TEMPERATURE"] = "0.7"
    env_with_overrides["LLM_PROVIDER"] = "bedrock"
    overridden = load_ai_settings(env_with_overrides)
    if overridden.bedrock.max_tokens != 8000:
        problems.append("BEDROCK_MAX_TOKENS override was not applied")
    if overridden.bedrock.temperature != 0.7:
        problems.append("BEDROCK_TEMPERATURE override was not applied")

    # --- malformed numeric value: must raise a clear error, not crash obscurely ---
    env_bad_int = dict(_MINIMAL_VALID_ENV)
    env_bad_int["BEDROCK_MAX_TOKENS"] = "not-a-number"
    try:
        load_ai_settings(env_bad_int)
        problems.append("A malformed BEDROCK_MAX_TOKENS should raise AISettingsError but did not")
    except AISettingsError as exc:
        if "BEDROCK_MAX_TOKENS" not in str(exc):
            problems.append(f"Error message should name the malformed variable; got: {exc}")

    # --- no credentials of any kind are ever read by this module ---
    env_with_fake_creds = dict(_MINIMAL_VALID_ENV)
    env_with_fake_creds["AWS_SECRET_ACCESS_KEY"] = "should-never-be-touched"
    settings_ignoring_creds = load_ai_settings(env_with_fake_creds)
    if "should-never-be-touched" in str(vars(settings_ignoring_creds)) + str(vars(settings_ignoring_creds.bedrock)):
        problems.append("load_ai_settings must never read or surface AWS credential environment variables")

    # --- bearer token: absent by default, backward compatible ---
    settings_no_token = load_ai_settings(_MINIMAL_VALID_ENV)
    if settings_no_token.bedrock.bearer_token is not None:
        problems.append("bearer_token should default to None when AWS_BEARER_TOKEN_BEDROCK is unset")

    # --- bearer token: read correctly when present ---
    env_with_token = dict(_MINIMAL_VALID_ENV)
    env_with_token["AWS_BEARER_TOKEN_BEDROCK"] = "test-bearer-token-value"
    settings_with_token = load_ai_settings(env_with_token)
    if settings_with_token.bedrock.bearer_token != "test-bearer-token-value":
        problems.append("AWS_BEARER_TOKEN_BEDROCK was not read into bedrock.bearer_token")
    # Presence of a bearer token must not affect any other field.
    if settings_with_token.bedrock.region != "us-east-1" or settings_with_token.bedrock.model_id != "anthropic.claude-3-test":
        problems.append("Setting a bearer token should not affect unrelated settings fields")

    if problems:
        print("\nFAILURES:")
        for p in problems:
            print(f"  - {p}")
        print(f"\nFAIL - {len(problems)} problem(s).")
        return 1

    print("ALL SETTINGS CHECKS PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
