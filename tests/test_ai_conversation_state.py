"""
tests/test_ai_conversation_state.py
======================================
Tests ``ai.session.ConversationState.merge_filter`` -- the mechanism
behind resolving a follow-up message ("Show margin") using entities
established in an earlier turn ("Compare Q2 vs Q3 for HPE") without the
user repeating them.

Usage:
    python tests/test_ai_conversation_state.py
"""
from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from ai.data.filters import Filter  # noqa: E402
from ai.session import ChatSession, ConversationState  # noqa: E402


def main() -> int:
    problems: list = []

    # --- a fresh state merges to exactly whatever the new filter sets ---
    state = ConversationState()
    merged = state.merge_filter(Filter(client="HPE"))
    if merged.client != "HPE":
        problems.append("Merging into a fresh ConversationState should adopt the new filter's fields")

    # --- an established client carries forward when the new filter doesn't mention one ---
    state.active_filter = merged  # simulate FilteringNode persisting the merged result
    followup_merged = state.merge_filter(Filter())  # "Show margin" resolves no new entities
    if followup_merged.client != "HPE":
        problems.append(
            f"A follow-up with no new client should carry forward the previous client; "
            f"got client={followup_merged.client!r}"
        )

    # --- a new client in the follow-up overrides the old one ---
    state.active_filter = followup_merged
    override_merged = state.merge_filter(Filter(client="Aldevron"))
    if override_merged.client != "Aldevron":
        problems.append("A follow-up naming a new client should override the previously active one")

    # --- fields are merged independently: setting poc doesn't clear an existing client ---
    state.active_filter = Filter(client="HPE", section="staffing_secured")
    partial_merged = state.merge_filter(Filter(poc="Vijay"))
    if partial_merged.client != "HPE" or partial_merged.section != "staffing_secured" or partial_merged.poc != "Vijay":
        problems.append(
            f"merge_filter should combine fields independently, not treat the new filter as a full "
            f"replacement; got {partial_merged}"
        )

    # --- merge_filter does not mutate self.active_filter; caller must assign explicitly ---
    state.active_filter = Filter(client="HPE")
    _ = state.merge_filter(Filter(client="Aldevron"))
    if state.active_filter.client != "HPE":
        problems.append("merge_filter() should not mutate active_filter in place")

    # --- ChatSession starts with a fresh ConversationState and empty messages ---
    session = ChatSession(session_id="s1")
    if session.state.active_filter.client is not None:
        problems.append("A new ChatSession's ConversationState should start with an empty Filter")
    if session.messages:
        problems.append("A new ChatSession should start with no messages")

    if problems:
        print("\nFAILURES:")
        for p in problems:
            print(f"  - {p}")
        print(f"\nFAIL - {len(problems)} problem(s).")
        return 1

    print("ALL CONVERSATION STATE CHECKS PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
