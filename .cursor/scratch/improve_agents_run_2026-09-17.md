# improve-agents-md run — 2026-09-17 (CLOUD ADAPTER)

Window: this flashcards cloud environment (5 agents). Topics `_system` **absent**.
Steer: bare `SKILL (7).md` → full workflow with cloud substitutes.

## Quantified findings (top 3 fixed)

```
PATTERN: Inline xlsx base64Content via Drive create_file truncates/corrupts; agents spawn children and retry the same path
COUNT: 4 child agents, 28 create_file-ish attempts, 2 ERROR loops (972 + 148 msgs)
TOKENS_WASTED: ~1.1K messages (~500K–2M tokens class) + principal wait
ALREADY_DOCUMENTED: no
FIX_TYPE: rule + skill/command
FIX_LOCATION: AGENTS.md Drive write policy; .cursor/rules/drive-uploads.mdc; implement-plan CLOUD ADAPTER
```

```
PATTERN: Topics skills invoked on flashcards cloud without C:/Topics; suite tools searched late every time
COUNT: 5+ skill invocations in 1 session (deepen/diagnose/implement/improve + “add skills”)
TOKENS_WASTED: ~50K+ (missing-tool discovery + plan_link/suite dead ends)
ALREADY_DOCUMENTED: no (skills assumed Topics)
FIX_TYPE: skill/command
FIX_LOCATION: CLOUD ADAPTER block atop deepen/diagnose/implement/improve skills; AGENTS.md Environment facts
```

```
PATTERN: Monarch access explained (MCP desktop) then re-asked same session (“How do we access monarch?”)
COUNT: 1 re-ask across 1 session
TOKENS_WASTED: ~10K
ALREADY_DOCUMENTED: no
FIX_TYPE: rule
FIX_LOCATION: AGENTS.md Monarch (cloud) card — lead with CSV-drop path
```

## Logged for next run (not fixed this pass)

```
PATTERN: Premature Drive packet build before plan lock (“Lets take a step back…”)
COUNT: 1 explicit correction
FIX_TYPE: rule (partially covered in deepen-plan CLOUD ADAPTER + AGENTS.md plan-lock line)
```

```
PATTERN: Google Drive search_files called with unsupported free-text queries
COUNT: 3 Unsupported query field errors
FIX_TYPE: skill note — use title contains / parentId syntax only
```

## Infrastructure audit (cloud)

- `transcript_census.py` / `delegation_census.py` / `build_kernel.py`: **N/A** (no Topics)
- Drive MCP: available but binary path unsafe — guidance + rule, not a fork of Google’s MCP
- Sibling improve run: none (clean git on skills branch at start)

## Verify

- [x] AGENTS.md present with Drive + Monarch + Topics-lite cards
- [x] All four skills have CLOUD ADAPTER
- [x] alwaysApply drive-uploads rule
