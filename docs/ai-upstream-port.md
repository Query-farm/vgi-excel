# AI changes ported from vgi-web-frontend

Source checkout: `~/Development/vgi-web-frontend`, reviewed through `d667e41`
(2026-09-20). Cupola for Excel build: `20260920.2`.

| Upstream commits | Excel adaptation |
| --- | --- |
| `ac04853` | Sonnet 5 default, Opus 5 / Haiku choices, migration of superseded model IDs, model-gated adaptive thinking and effort, bounded output-token setting, signed and redacted thinking preserved during tool rounds. |
| `5974f6d`, `ac04853` | Cache markers on stable tools/system and the growing conversation. Office freezes its catalog prompt for the session, as desktop already did. Regression tests compare complete rendered request prefixes across turns. |
| `1a251e1` | Optional Anthropic workspace ID header and actionable workspace errors. Workspace IDs remain in session storage; API-key storage rules are unchanged. |
| `256637d`, `8c9cc13` | Shared paginated/searchable catalog tools, category registries, exact function descriptions, required-filter explanations for tables and views, decoded VGI tags, bounded/deduplicated examples, and exclusion of private agent-grader tags. Tag helpers, the contract snapshot, and its tests are adapted from upstream. |
| `7d5ef8e` | Copy Markdown tables as HTML and TSV, with a native text-clipboard fallback in WebView2 and persistent inline copy status. |

The Office pane and desktop WebView share the Messages API request builder,
retry policy, SSE parser, and catalog tool execution in `packages/core`. Office
loads the JSON extension before catalog queries to avoid threaded WASM stalls
during extension autoload. The
native WinForms fallback also uses the new default model, cache markers and
model-gated thinking/effort. It retains its existing smaller tool set and
non-streaming interface. Its optional settings are
`VGI_EXCEL_ANTHROPIC_MODEL`, `VGI_EXCEL_ANTHROPIC_EFFORT`, and
`VGI_EXCEL_ANTHROPIC_WORKSPACE_ID`.

Cancellation fills missing tool results instead of deleting signed assistant
blocks. Live conversation history stays unchanged as the cached prefix grows.
On window reload, conversations containing signed thinking resume from the
visible user/assistant transcript: persisted tool results must expire their
process-local IDs, and replaying an edited signed tool exchange is invalid.
This intentionally gives up the old tool transcript on reload while preserving
conversation tabs and the visible answers.

## Features that require a separate Excel implementation

Report editing/narratives (`e46a6bb`, `b9f1f03`, report portions of `5974f6d`),
report layout changes, and chart-image pruning (`dbd7eea`) depend on the web
report/chart surfaces, which Excel does not implement. Excel continues to stage
worksheet writes for explicit user confirmation.

Semantic query compilation and access modes (`a7fa198`, `a5dda23` and subsequent
semantic/report commits) depend on the web app's semantic model/compiler and
governed report dataset layer. Metadata is exposed by Excel's describe tools,
but this port does not claim semantic-only enforcement or add a compiler tool.

The upstream AI tracing, prompt/response telemetry, and usage telemetry are
excluded by Excel's error-only privacy policy. No AI prompts, thinking blocks,
responses, tool results or credentials are sent to Sentry.

## Validation

Run `npm run check`, `npm test`, `npm run build`, `npm run test:ui`, and
`npm run test:office-wasm`. The live WASM suite uses the real VGI engine for AI
metadata tools but mocks Anthropic responses; it makes no paid AI requests.
The unit suites cover split SSE frames/signatures, model gating, workspace
headers, cancellation, cache-prefix stability, metadata and persistence.

Run `tests\run-windows.ps1` on Europa after copying a clean desktop bundle.
Native policy tests check the fallback request format and keep the read-only
bridge boundary intact. No real Anthropic credentials are needed by tests.

The model-discovery update supersedes the original model-ID migration: saved
models are now preserved exactly. AI settings use the Models API with a daily,
credential/workspace-scoped metadata cache and capability-driven options. Bundled
model IDs remain an offline suggestion list, not an automatic upgrade policy.
