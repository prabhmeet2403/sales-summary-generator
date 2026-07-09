"""
ai.workflow.nodes.response
==============================
Produces the final, coherent, natural-language answer shown to the
user -- the only node the approved architecture designates as always
running, regardless of intent (Architecture Plan v3 Section 1.2).

For an analytical question, this node's prompt to the model is
deliberately constrained to *only* the factual findings already
gathered by ``AnalysisNode`` (each tool's ``ToolResult.summary``) --
the model's job here is narration over already-correct facts, not
independent calculation, which is what keeps a multi-tool answer
"a single coherent response" instead of a list of disconnected tool
outputs, without ever letting the model introduce a number the tools
didn't already produce.

The synthesis prompt additionally asks the model to present that
narration in a consistent, business-oriented structure (Summary / Key
Findings / Business Impact / Next Suggested Questions, with tables for
multi-record findings) -- a presentation-only instruction over the
same already-gathered facts; it does not change what data reaches this
node or how it is gathered. When a tool's own findings already arrive
under their own clear headings (e.g. ``ai.tools.executive_summary``'s
five-section board format), the prompt asks the model to preserve
those headings verbatim instead of substituting the generic template.
"""

from __future__ import annotations

import logging
from typing import ClassVar

from ai.llm.provider import LLMProvider, Message
from ai.workflow.graph import Intent, WorkflowNode, WorkflowState

logger = logging.getLogger(__name__)

_UNSUPPORTED_REQUEST_RESPONSE = (
    "I can't create charts, dashboards, or exported files yet -- that capability is "
    "planned for a later phase. Right now I can answer questions about revenue, margin, "
    "quarter comparisons, and specific clients or POCs. Would you like to ask one of those instead?"
)

_NO_RESULTS_RESPONSE = (
    "I wasn't able to find relevant data for that question. Could you rephrase it, or ask "
    "about revenue, margin, a quarter comparison (e.g. Q2 vs Q3), or a specific client or POC?"
)

_ANALYTICAL_SYNTHESIS_PROMPT = """\
The user asked: "{message}"

Here is the factual data already gathered to answer it:
{findings}

Write the answer for a business/executive audience, using only the \
facts above -- never invent, estimate, or round a number beyond what \
is shown. If the facts above don't fully answer the question, say \
plainly what is available rather than guessing at the rest.

Structure the answer with these markdown headings, in this order, \
using only the ones that genuinely apply to what was asked (skip a \
heading entirely rather than leave it empty or padded with filler):

## Summary
One or two sentences giving the headline answer.

## Key Findings
The specific figures from the facts above. If there are several \
similar records to present (e.g. multiple clients, quarters, or line \
items), format them as a markdown table with clear column headers \
instead of a bulleted list.

## Business Impact
Briefly explain why this matters in business terms -- what it means \
for revenue, margin, forecasting confidence, or reporting accuracy -- \
rather than only restating the numbers. Base this strictly on what \
the facts above support; do not speculate beyond them. For example, \
prefer "27 historical values could not be located, which may affect \
year-over-year comparisons and should be reviewed before finalizing \
the report" over a bare "27 historical values missing."

## Next Suggested Questions
Two or three natural follow-up questions the user could reasonably \
ask next, each on its own line.

Exception: if the facts above already arrive under their own clear \
markdown headings (for example, an Executive Summary's own "## \
Executive Summary" / "## Revenue Highlights" / "## Validation Status" \
/ "## Key Business Observations" / "## Recommended Management \
Attention" sections), keep exactly those headings, in that exact \
order, instead of the Summary/Key Findings/Business Impact template \
above -- only polish the wording within each section, and still use a \
markdown table wherever a section lists several similar records.

Keep the tone factual, concise, and business-appropriate throughout."""

_GENERAL_CONVERSATION_PROMPT = 'Respond briefly and warmly to this message: "{message}"'


class ResponseNode(WorkflowNode):
    """Synthesizes the final answer shown to the user."""

    name: ClassVar[str] = "response"
    display_label: ClassVar[str] = "Finalizing response"

    def __init__(self, provider: LLMProvider) -> None:
        self._provider = provider

    def should_run(self, state: WorkflowState) -> bool:
        return True  # always the last node

    def run(self, state: WorkflowState) -> WorkflowState:
        if state.intent == Intent.UNSUPPORTED_REQUEST:
            state.final_response = _UNSUPPORTED_REQUEST_RESPONSE
            return state

        if state.intent == Intent.ANALYTICAL_QUESTION:
            if not state.analysis_results:
                state.final_response = _NO_RESULTS_RESPONSE
                return state
            findings = "\n".join(result.summary for result in state.analysis_results)
            prompt = _ANALYTICAL_SYNTHESIS_PROMPT.format(message=state.user_message, findings=findings)
        else:
            prompt = _GENERAL_CONVERSATION_PROMPT.format(message=state.user_message)

        response = self._provider.converse([Message(role="user", text=prompt)])
        state.final_response = response.text or "I'm not able to provide an answer right now."
        return state
