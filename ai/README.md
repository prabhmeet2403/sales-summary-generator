# `ai/` — AI-Powered Business Intelligence Platform (Phase 2)

This package is the Phase 2 AI layer built on top of the Phase 1 Sales
Forecast Automation Engine. It is entirely additive: it reads Phase 1's
already-validated output and never re-implements, re-parses, or
recalculates anything Phase 1 already produces.

**Full design documents** (read these for the complete approved
architecture; this README covers what is actually implemented so far):

- `Phase2_AI_Assistant_Architecture_Plan.md` — Revision 1 (original scope)
- `Phase2_AI_Assistant_Architecture_Plan_v2.md` — Revision 2 (planner, DataFrame layer, filters, semantic search, plugin tools)
- `Phase2_AI_Assistant_Architecture_Plan_v3.md` — Revision 3 (**approved, frozen baseline**: workflow graph, Analytics Engine, Dashboard Designer, LLM provider abstraction, service facade, audit logging)

## Implementation status

**Phase 2a (foundation) — implemented.** Settings, the LLM provider
abstraction (Amazon Bedrock), the business data context, the DataFrame
query layer, the Universal Filter Engine, and the service facade.

**Phase 2b (workflow graph, plugin tools, Analytics Engine, AI Assistant UI) — implemented.**
The nine-node workflow graph, the plugin tool architecture with five
analysis tools, the Analytics Engine, `ConversationState`, and a working
"AI Assistant" page in the Streamlit application. See "What's
implemented" below for the full module list.

**Not yet implemented** (Phase 2c onward): chart generation and the
Chart Recommendation Engine, semantic search, the Dashboard Designer,
report templates, the full Explainability panel (today: "Sources Used"
only), multi-level caching beyond what Python's own object lifetime
provides, audit logging, and the REST API layer. Do not assume any
capability beyond what is listed below actually exists in the code.

## What's implemented

### Phase 2a

| Module | Purpose |
|---|---|
| `ai/settings.py` | Centralized, environment-variable-driven configuration. No credentials, model IDs, or regions are ever hardcoded. |
| `ai/llm/provider.py` | The `LLMProvider` abstract interface (`converse`, `converse_stream`, `embed`) and the `Message`/`ToolUseRequest`/`ToolResultBlock` types every provider and tool-use conversation is built from. |
| `ai/llm/factory.py` | Selects and constructs the configured provider from `AISettings.llm_provider`. |
| `ai/llm/providers/bedrock.py` | The Amazon Bedrock implementation. The *only* module permitted to import `boto3`. |
| `ai/context.py` | `BusinessContext` — wraps one successful generation's Phase 1 output for AI consumption. |
| `ai/data/frames.py` | Builds `groups_df`/`monthly_df` as a pure, auditable readout of `GroupSummary`/`MonthlyGroupSummary` — zero recalculation. |
| `ai/data/filters.py` | `Filter` + `apply_filter()` — the one shared filtering mechanism every tool uses. |

### Phase 2b

| Module | Purpose |
|---|---|
| `ai/tools/base.py` | `BaseTool`, `ToolCategory`, `ToolResult` — the plugin interface every tool implements. |
| `ai/tools/registry.py` | Auto-discovers every `BaseTool` subclass under `ai.tools`; dispatches tool calls by name. |
| `ai/tools/schemas.py` | The shared filter JSON-schema fragment every analytical tool's schema embeds. |
| `ai/tools/revenue.py`, `margin.py`, `comparison.py`, `client_lookup.py`, `poc_lookup.py` | The five Phase 2b tools (Revenue Analysis, Margin Analysis, Quarter Comparison, Client Lookup, POC Lookup) — all `ANALYSIS` category, all thin wrappers over `AnalyticsEngine`. |
| `ai/analytics/engine.py` | `AnalyticsEngine` — the single implementation of `kpi`/`rank`/`compare`/`trend`/`aggregate`, the generic operations every tool uses instead of reimplementing its own. |
| `ai/session.py` | `ConversationState` (built on `Filter`) and `ChatSession` — cross-turn memory. |
| `ai/workflow/graph.py` | `WorkflowState`, `WorkflowNode`, `WorkflowGraph`, `ExecutionTrace` — the core graph execution engine. |
| `ai/workflow/factory.py` | Assembles the approved nine-node graph in order. |
| `ai/workflow/nodes/` | The nine nodes: `intent_detection.py`, `planning.py`, `entity_resolution.py`, `filtering.py`, `analysis.py`, `response.py` (all real); `future_stubs.py` (Visualization/Reporting/Export — always decline to run in Phase 2b; see its module docstring). |
| `ai/service.py` | `AIService` — now runs the full workflow graph per `ask()` call. Public signature unchanged from Phase 2a. |
| `ai/ui/chat_page.py` | The "AI Assistant" Streamlit page. |
| `ai/ui/message_render.py` | Renders one chat turn via `st.chat_message`, including the "Sources Used" trail. |
| `ai/ui/progress_display.py` | The live, per-node progress checklist during graph execution. |

## The Phase 1 touch-points

Two, both narrow and both justified individually:

1. **`gui/runner.py`** — `GenerationResult` has six new, optional fields
   (`section_results`, `monthly_section_results`, `rows`, `month_roles`,
   `target_year`, `prior_years`), populated at the end of
   `generate_summary()`'s existing success path from variables it
   already computes. This is the originally-approved Phase 2a
   touch-point; nothing about it changed in Phase 2b.
2. **`app.py`** — one new nav branch (`elif _nav == "ai": ...`), plus a
   one-line guard added to the pre-existing `_resolve_upload_state()`
   function. That guard fixes a latent defect this phase's own nav
   addition exposed: `_resolve_upload_state()` runs on every page (by
   design, to keep sidebar badges accurate everywhere), and one of its
   branches misread "the `master_upload` widget isn't rendered on this
   page" as "the user removed their file," silently clearing the
   generated Summary the moment a user navigated to any other page.
   Confirmed pre-existing and independent of Phase 2 (reproduced by
   navigating to the untouched "Settings" page); fixed with a single
   additional condition on the existing branch, changing nothing about
   what happens when a file is genuinely removed on the Upload page
   itself. See the Phase 2b completion notes for the full before/after
   verification.
3. **`components/sidebar.py`** — one new `_NAV` tuple entry.

Nothing about how a number is calculated, parsed, aggregated, or
written to the generated workbook changed anywhere. All four Phase 1
test files (`tests/compare_with_manual.py`, `tests/test_column_reordering.py`,
`tests/test_future_year_compatibility.py`,
`tests/test_worksheet2_actual_forecast.py`) pass unmodified, verified
after every change in both phases.

## Configuration

Set these environment variables before using the AI Assistant (loaded
via `ai.settings.load_ai_settings()`):

**Local development:** copy `.env.example` to `.env` in the project
root and fill in real values -- `app.py` loads it automatically on
startup via `python-dotenv`. `.env` is gitignored; never commit it.
A real deployment should set these as actual environment variables
(container/platform secrets) and does not need a `.env` file at all --
`load_dotenv()` never overrides a variable that's already set in the
real environment.

| Variable | Required | Default | Purpose |
|---|---|---|---|
| `BEDROCK_REGION` | Yes | — | AWS region hosting the Bedrock endpoint. |
| `BEDROCK_MODEL_ID` | Yes | — | Bedrock model ID for chat/tool-use calls. |
| `BEDROCK_EMBEDDING_MODEL_ID` | No | none | Bedrock model ID for text embeddings. Not required -- no feature in this version calls the embedding API; enforced only once semantic search is implemented. |
| `LLM_PROVIDER` | No | `bedrock` | Selects the provider implementation. Only `bedrock` exists today. |
| `AWS_BEARER_TOKEN_BEDROCK` | No | none | Bedrock API key for bearer-token authentication, used instead of standard AWS IAM credentials when set. See "Authentication" below. |
| `BEDROCK_INFERENCE_PROFILE_ARN` | No | none | Cross-region inference profile ARN, preferred over `BEDROCK_MODEL_ID` when set. |
| `BEDROCK_MAX_TOKENS` | No | `4096` | Default max output tokens. |
| `BEDROCK_TEMPERATURE` | No | `0.2` | Default sampling temperature. |
| `BEDROCK_TIMEOUT_SECONDS` | No | `60` | Per-request timeout. |
| `BEDROCK_MAX_RETRIES` | No | `3` | Max retry attempts for throttling/transient errors. |
| `AI_RATE_LIMIT_PER_MINUTE` | No | `30` | Application-level request cap (defined, not yet enforced anywhere — reserved for a later phase). |

### Authentication

Two supported modes, chosen automatically based on what's configured:

- **Standard AWS IAM credentials** (default): access key/secret key,
  a shared credentials file, or an attached IAM role, resolved
  entirely by `boto3`'s own default credential chain. No AWS
  credential is ever read directly by this codebase for this mode.
- **Bearer token**: set `AWS_BEARER_TOKEN_BEDROCK` to a Bedrock API
  key. When present, it takes precedence for Bedrock calls; when
  absent, standard credentials are used exactly as before. This
  variable's value is read once (by `ai.settings.load_ai_settings`)
  and is never logged — only whether one is configured.

In the Streamlit app, if neither is available, the AI Assistant page
shows a clear "not configured" message rather than crashing.

## Usage

Via the Streamlit app: generate a Summary on the "Upload & Generate"
page, then open "AI Assistant" from the sidebar and ask a question.

Programmatically:

```python
from gui.runner import generate_summary
from ai.context import BusinessContext
from ai.llm.factory import get_provider
from ai.service import AIService
from ai.settings import load_ai_settings

# 1. Run the unmodified Phase 1 pipeline.
result = generate_summary(input_path="master_2026.xlsx", output_dir="output", year=2026)

# 2. Wrap the result for AI consumption.
context = BusinessContext.from_generation_result(result)

# 3. Construct a provider from centralized settings (reads env vars).
provider = get_provider(load_ai_settings())

# 4. Ask a question -- runs the full nine-node workflow graph.
service = AIService(context, provider)
turn = service.ask(session_id="user-session-1", message="Compare Q2 vs Q3 revenue for HPE")
print(turn.text)
print("Sources used:", turn.sources_used)

# 5. A natural follow-up resolves via ConversationState, no repetition needed:
turn2 = service.ask(session_id="user-session-1", message="Show margin instead")
```

## Architectural boundaries (enforced by convention and by tests)

- No module outside `ai.ui` may import `streamlit` (checked by `tests/test_ai_service_facade.py`, which scans every `.py` file under `ai/` except `ai/ui/`).
- No module outside `ai.llm.providers` may import `boto3` or reference a specific LLM provider by name.
- No Phase 1 module may import anything from `ai`.
- `ai/data/frames.py` never performs arithmetic, aggregation, or a business calculation — every column is a direct readout of an existing Phase 1 field, checked mechanically by `tests/test_ai_dataframe_parity.py`.
- `ai/analytics/engine.py` only performs generic shaping operations (sum, rank, group-by, percent-change) — never a business rule — checked against independent pandas computations by `tests/test_ai_analytics_engine.py`.
- Adding a new tool requires no change to `ai/tools/registry.py` or any node — checked by the auto-discovery assertion in `tests/test_ai_tools.py`.

## Testing

All tests follow the same standalone-script convention already
established by Phase 1's test suite (`def main() -> int`, no pytest,
real fixture data, no mocking of business logic):

```bash
# Phase 2a
python tests/test_ai_settings.py              # settings loading, required/optional/malformed vars
python tests/test_ai_dataframe_parity.py       # DataFrame columns match Phase 1 objects exactly
python tests/test_ai_filter_engine.py          # Filter/apply_filter, every field + combinations
python tests/test_ai_context.py                # BusinessContext construction, lookups, error handling
python tests/test_ai_bedrock_provider.py       # BedrockProvider against a fake boto3 client

# Phase 2b
python tests/test_ai_tools.py                  # all five tools + registry auto-discovery
python tests/test_ai_analytics_engine.py       # kpi/rank/compare/trend/aggregate vs. independent pandas
python tests/test_ai_conversation_state.py     # Filter.merge_filter ellipsis-resolution logic
python tests/test_ai_workflow_graph.py         # node ordering, should_run() gating, ExecutionTrace
python tests/test_ai_workflow_nodes.py         # intent parsing, entity matching, planning, stub nodes
python tests/test_ai_service_facade.py         # AIService against a scripted fake LLMProvider
python tests/test_ai_phase2b_integration.py    # the full worked example, end to end
```

The only place any of these tests use a fake/mock is the LLM provider
boundary — the external network call this environment cannot make.
Everything upstream of that boundary (generation, `BusinessContext`,
the DataFrame layer, filtering, the Analytics Engine, every tool, the
graph's own mechanics) is exercised for real, against the real fixture
workbook.

**Manually verified, not part of the automated suite** (requires a live
browser + running Streamlit server): the full UI flow — nav item
presence, empty state before generation, "not configured" state without
Bedrock env vars, the chat interface rendering real context after
generation, and a real chat message failing gracefully at the network
boundary (this sandbox has no AWS access) with a clear, styled error
rather than a crash.

## Known limitations (Phase 2b)

- No chart, dashboard, report, search, or export capability exists yet
  — a request for any of these is classified as `unsupported_request`
  and answered with a fixed, honest message explaining the limitation,
  with zero LLM cost beyond the classification call itself.
- The Planning node's tool-category selection is deterministic (checks
  the registry for at least one `ANALYSIS` tool), not LLM-driven — a
  reasoned choice given every registered tool belongs to the same
  category today; revisit once a second non-trivial category exists to
  choose between (see the module docstring on `ai/workflow/nodes/planning.py`).
- Entity resolution matches client/POC names via substring matching
  against the known, finite vocabulary in this generation's data, not
  an LLM call — deliberately, since the vocabulary is fully enumerable
  and a lookup is strictly more reliable than a model's recall for this
  specific task.
- Cross-turn memory is `ConversationState`'s `Filter` only (the
  documented, worked-example-tested mechanism for "show margin"-style
  follow-ups); raw conversation history is not replayed to every node's
  LLM call. This is a deliberate scope boundary for Phase 2b, not an
  oversight — see Architecture Plan Revision 2 Section 9's distinction
  between grounding context and full history threading.
- `AI_RATE_LIMIT_PER_MINUTE` remains defined but unenforced.
- Historical year-over-year data (`GroupSummary.historical`) is still
  not in `groups_df`, for the same reason as Phase 2a: no current
  consumer needs it.
