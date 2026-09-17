# AGENTS.md — flashcards / cloud tax packet workspace

Lean operating card for Cursor Cloud agents on this repo. Not the Topics kernel.
Organizational tax work here is **not tax advice**.

## Environment facts (do not re-discover)

| Fact | Implication |
|------|-------------|
| This is **not** `C:\Topics` | No `_system/tools`, no `transcript_census.py`, no `plan_link.py`, no Topics suite. Topics skills still run — use the **CLOUD ADAPTER** block at the top of each skill. |
| Monarch MCP is **not** in this cloud session | Do not promise live Monarch tags/MCP. Canonical cloud path: user drops CSV/PDF into Drive → agent profiles. Desktop MCP (`https://api.monarch.com/mcp`) is a separate human setup. |
| Google Drive MCP is available | Prefer Sheets/CSV `textContent`. Binary xlsx `base64Content` truncates — see Drive write policy. |

## Drive write policy (HARD)

Incident: Sep 17 2026 tax turbo — 4 child upload agents, 28 `create_file` attempts, corrupt stubs of **8 / 1526 / 7500** bytes vs 25–34 KB sources; two ERROR loops (972 + 148 msgs) retrying the same inline-base64 path.

**ALWAYS**
1. Upload tabular data as **Google Sheet** via `textContent` (TSV/CSV) or a Doc for prose.
2. After every `create_file`, **verify `fileSize`** (and re-read content). `fileSize` 1 / &lt;2KB for a claimed spreadsheet = corrupt — trash and retry with a smaller/text payload.
3. Keep CapEx/CPA “xlsx” as a **Sheet** or a tiny verified slice (&lt;~20KB) only when binary is required.

**NEVER**
1. Inline multi-KB `base64Content` for full workbooks through `CallDynamicTool` when a Sheet/CSV works.
2. Spawn child agents whose only job is “upload this xlsx via base64.”
3. Declare success without checking returned `fileSize` against the source.

## Plan lock before packet build

If the user has not locked the operating model (e.g. hybrid Monarch + CapEx registers), **do not** mass-create Drive workbooks. Deepen/diagnose first. Incident: user had to say “Lets take a step back and make sure this is a good plan” after an overbuild.

## Monarch (cloud)

Lead with this card when Monarch comes up — do not make the user re-ask “How do we access monarch?”:

1. **Cloud agent:** no Monarch MCP tools here. Ask for a Monarch CSV export in Drive (`Documents/Monarch Money`).
2. **Desktop Cursor:** add official MCP `https://api.monarch.com/mcp` via OAuth in `mcp.json`.
3. **Hybrid tax rule:** Monarch = operating P&L; CapEx + sale Closing Disclosures stay on Excel/Sheets registers.

## Topics skills in this repo

Installed under `.cursor/skills/`: `deepen-plan`, `diagnose`, `implement-plan`, `improve-agents-md`.
On invoke: if `/c/Topics` / `C:/Topics` is missing, print one line — `CLOUD ADAPTER: Topics _system absent — running skill without suite tools` — then continue with Drive/git/local plans only.

## improve-agents-md in this repo

Mine **cloud agent transcripts** via `cursor-cloud` `batch-fetch-details` (includeTranscripts), not `C:\Users\yanti\.cursor\projects\c-Topics\agent-transcripts`. Park extracts under `/home/ubuntu/.cursor/scratch/imp_*`. Commit guidance fixes here (this `AGENTS.md` + skills); do not invent a fake Topics `_system`.
