# AGENTS.md — this repo (Cursor Cloud)

Not the Topics kernel. Tax packet work here is **organizational only — not tax advice**.

## 0. First 10 seconds

```
test -d /c/Topics/_system || test -d /workspace/../Topics/_system
```

- **Missing** → print `CLOUD ADAPTER: Topics _system absent` and follow this file + each skill’s CLOUD ADAPTER.
- **Present** → follow Topics skills as written (this card is secondary).

## 1. Environment facts (do not re-discover)

| Fact | Rule |
|------|------|
| Repo is `flashcards`, not `C:\Topics` | No `transcript_census.py`, `plan_link.py`, Topics suite, or `_system` packs |
| Monarch MCP **not** in cloud tool catalog | NEVER claim live Monarch access. Cloud path = CSV/PDF drop in Drive |
| Google Drive MCP **is** available | Sheets/Docs via `textContent` only for multi-KB data — §2 |
| Plans live in `.cursor/plans/` | Prefer workspace plans; `~/.cursor/plans/` is a mirror when present |

## 2. Drive writes (HARD — mechanism, not preference)

**Incident (2026-09-17):** 4 child upload agents, 28 `create_file` attempts, corrupt stubs **8 / 1526 / 7500** bytes vs 25–34 KB sources; ERROR loops of 972 and 148 messages retrying the same inline-`base64Content` path.

| | |
|--|--|
| ✅ INSTEAD | `create_file` with `textContent` (TSV/CSV) → Google Sheet, or prose → Google Doc |
| ⛔ NEVER | Inline multi-KB `base64Content` for full `.xlsx` through `CallDynamicTool` when a Sheet works |
| ⛔ NEVER | Spawn a child agent whose only job is “upload this xlsx via base64” |
| ⛔ NEVER | Declare success without a size/content check |

**After every `create_file` (required visible step):**

```bash
python .cursor/tools/drive_create_verify.py --expected-min-bytes N --got-filesize F --title "…"
```

Exit nonzero ⇒ trash the file, retry as Sheet/`textContent`, do **not** re-inline the same base64.  
`fileSize` of `1` or `< 2048` for a claimed spreadsheet = corrupt (known MCP empty-sheet / truncation symptom).

**Binary exception:** only when the user explicitly needs an `.xlsx` download **and** source ≤ ~20 KB; still run the verifier against returned `fileSize`.

**`search_files` query language (not free text):**

```
title contains 'Monarch'
parentId = '1HqEA5yfN12bAeY5yseTANPKEDBl0DJgT'
title contains 'Airbnb' and mimeType = 'application/pdf'
```

⛔ NEVER pass bare words like `query: "Monarch"` — returns `Unsupported query field`.

## 3. Plan lock before packet build

⛔ NEVER mass-create Drive workbooks / CapEx clones before the operating model is locked (hybrid vs Monarch-only, etc.).  
✅ INSTEAD deepen → diagnose → then implement.  
Incident: user had to say “Lets take a step back and make sure this is a good plan” after an overbuild.

## 4. Monarch (answer once — lead with this card)

When Monarch comes up, lead with this block (do not make the user re-ask “How do we access monarch?”):

1. **This cloud session:** no Monarch MCP tools. Drop a Monarch **CSV export** in Drive folder `Documents/Monarch Money` (`1HqEA5yfN12bAeY5yseTANPKEDBl0DJgT`).
2. **Desktop Cursor:** OAuth official MCP `https://api.monarch.com/mcp` in `mcp.json`.
3. **Hybrid tax rule:** Monarch = recurring operating P&L only. CapEx + sale Closing Disclosures stay on Excel/Sheets registers — never double-enter CapEx as Monarch tags.

## 5. Topics skills in this repo

`.cursor/skills/`: `deepen-plan` · `diagnose` · `implement-plan` · `improve-agents-md`

On every invoke without Topics: one `CLOUD ADAPTER:` line, then continue (Drive / git / `.cursor/plans`). Do not invent a fake `_system`.

**improve-agents-md (cloud):** mine via `cursor-cloud` `list-cloud-agents` + `batch-fetch-details` (`includeTranscripts: true`); one subagent per transcript; park extracts in `/home/ubuntu/.cursor/scratch/imp_*`; write fixes to **this** `AGENTS.md` + skills/rules; push the feature branch.

## 6. Canonical Drive anchors (2025 Taxes)

| Role | Id |
|------|-----|
| 2025 Taxes folder | `1DkF0ealLM2YZ8IkCxBp-_jN9BJRJcxuu` |
| Monarch Money folder | `1HqEA5yfN12bAeY5yseTANPKEDBl0DJgT` |
| CapEx register CANONICAL | `1ojNHX_z94FGMXBFKGcZ7Eta_iyQCrBvNjKbitKUJT4Q` |
| TY2025 CapEx CPA ($53,660) | `1aqJe32LWQwwAXdk7pDdo0MJHBu6mVnM7hbsE5W6aBmE` |
| 524 Ferdinand Sale | `1K6_C7BcRF1PuC1l_MGjYWHBOGS2I9w9L` |
| STR platform vs Monarch recon | `1Yo0IFgaWi_fM6OFxznZH6HX25uwyDMp2puq-6PRA1fY` |

Prefer these ids over fuzzy title search when touching the tax packet.
