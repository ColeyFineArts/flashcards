---
name: deepen-plan
description: "Second-pass review of a plan: challenge assumptions, find gaps, verify with record searches, and fan the plan out to a cross-lab model panel (plan_panel.py) for adversarial advice. Use when a plan exists and needs rigorous review before implementation."
---

# Deepen Plan (Topics)

## CLOUD ADAPTER (flashcards / Cursor Cloud — run FIRST)

If Topics `_system` is missing, print `CLOUD ADAPTER: Topics _system absent — deepen without plan_panel/suite; write findings into the plan file only`.
Skip `plan_panel.py` / OpenRouter fan-out unless those tools exist. Prefer workspace `.cursor/plans/*.plan.md`. Do not mass-create Drive workbooks during deepen — lock the operating model first (incident: user “take a step back” after premature packet build).

---

Force a rigorous second-pass review of a plan before implementation. Challenges assumptions, searches for missed context, identifies gaps — and fans the plan out to a cross-lab advisory panel of OpenRouter models whose returns are triaged as intake.

Usage: `/deepen-plan` (uses most recent plan) or `/deepen-plan <plan_path>`

**Analysis-only -- no code or record changes in this phase. Plan file updates ARE permitted.**

**Plan mode: `/deepen-plan` is an EXECUTION skill (writes to the plan file in §6).** If Plan mode is active when this command is invoked, silently `SwitchMode(target_mode_id="agent")` with a one-line explanation and continue. Do NOT ask the user how to proceed -- the user invoking the command IS the override.

---

## Why This Exists

First-pass plans suffer from:
- **Anchoring bias** -- first interpretation shapes all subsequent analysis
- **Confirmation bias** -- searching for evidence that supports the initial plan
- **Premature closure** -- stopping when "good enough" instead of complete
- **Tunnel vision** -- missing adjacent record entries, related tools, edge cases

Plans without a second pass routinely miss context that forces mid-implementation rewrites. (An earlier version of this line carried an invented "~40%" statistic — exactly the false-precision failure the advice gate's Odds question now bans. No one measured that number.)

---

## 0. Load the Plan

Find the most recent plan. Topics plans live in the user-global plans dir and are named `topics_*.plan.md` (the "Topics " title convention):

```
Glob: `C:/Users/yanti/.cursor/plans/topics_*.plan.md` (sorted by mtime, take first)
Fallback: `C:/Users/yanti/.cursor/plans/*.plan.md`, then `.cursor/plans/*.plan.md` in the workspace
```

Or use the explicit path provided by the user.

**Privacy gate:** before writing anything into the plan, verify the EF gitignore holds:
`git -C "C:/Programming/Epistemic Firewall" check-ignore <plan_path>` must succeed. If it fails, STOP and fix the title/gitignore first (per AGENTS.md "Plans" convention).

**Announce:** "Deepening plan: `<plan_filename>`" — and in the SAME first reply paste the plan-link handover block:

```bash
python C:/Topics/_system/tools/plan_link.py --match "<words from the plan filename>"
```

Its two-line output (bare clickable link + plain Windows path) goes in VERBATIM — never hand-built, never backticked or fenced — and AGAIN as the last lines of this skill's FINAL reply. The user rule ("Plan handover — always give the link") binds every session that touches a plan, but a duty missing from a skill's own announce/gate/output templates drops in practice: sessions build replies from the templates, and /deepen-plan run f2445d54 shipped link-less minutes after that rule landed (Aug 29, 2026 — his third re-ask: "Yet again I don't know it").

Read the entire plan file into context before proceeding.

**If the plan touches a topic's record**, read that topic's hub §A block (hub filename from `_config.json` `hub` -- never assume `HUB.md`) and, for structural work, `_system/SYSTEM_PATTERN.md` WHOLE before reviewing.

---

## 0.5 Launch the Advisory Panel (background -- MANDATORY unless keyless)

Fan the plan out to a cross-lab, closed-book panel of OpenRouter models for adversarial advice. Launch it in the background NOW (`block_until_ms: 0`) so the seats run while you do §1--§2.5 yourself:

```bash
python C:/Topics/_system/tools/plan_panel.py --plan "<plan_path>" --context "<file>" --context "<file>"
```

- **Default seats** (all BARE -- the tool refuses `:online`): `openai/gpt-5.6-sol` (Pass A/B bench leader), `google/gemini-3.1-pro-preview` (8.5 on the Aug 23, 2026 health bench; OpenRouter serves the benched gemini-3.1-pro under the `-preview` id), `z-ai/glm-5.3-flash` (13/13 closed-book audit bench, near-free). Three non-Anthropic labs on purpose: the plan author is usually a Fable/Claude-class session, so cross-lab seats give anti-correlated blind spots (the dossier dual-run lesson). Typical panel cost is cents; each seat prints its usd on stderr.
- **The panel is CLOSED-BOOK -- it sees ONLY what you attach.** `--context <file>` (repeatable) attaches what a serious reviewer would need: the hub §A block, the source of the tool the plan edits, the incident note the plan cites. An unattached referent comes back "NOT ATTACHED -- CANNOT VERIFY"; that is the tool working, not failing.
- The run writes an in-flight research marker: §2.6 triage closes the loop, and `python C:/Topics/_system/tools/research_markers.py --ack-all` runs after the plan file is updated (§6). Do not end the turn while seats are in flight.
- **Exit 3 = no OpenRouter key:** skip the panel, name the skip in the Final Gate row, proceed. The panel augments §1--§3, never replaces them.

---

## 1. Adversarial Questions (MANDATORY)

Answer ALL questions in writing. No skipping. No "N/A". If something does not apply, explain WHY.

### 1.1 Completeness Check

| Question | Your Answer |
|----------|-------------|
| What happens if the primary approach fails? Is there a fallback? | |
| What error conditions are NOT handled in this plan? | |
| Which files are touched that AREN'T mentioned in the plan? | |
| Which checks in the suite could the change turn red? Which tool `--self-test`s cover it? | |
| What configuration dependencies exist (`_config.json`, `_system` protocol files, drive paths J:/Y:/Z:)? | |

### 1.2 Assumption Audit

| Question | Your Answer |
|----------|-------------|
| What did I assume about the record/tooling that I didn't verify with a lookup? | |
| What did I assume about file formats or entry-ID conventions that I didn't verify? | |
| What did I assume is ABSENT from the record without running the absence sweep (`lookup_id.py --absence`)? | |
| What public fact does the plan rest on that hasn't been web-verified (RESEARCH.md rules)? | |

### 1.3 Negative Space Search

| Question | Your Answer |
|----------|-------------|
| What search terms did I NOT try that might reveal related entries or code (vocabulary mismatch -- try `query_wiki.py`)? | |
| What files in the same directory did I NOT read? | |
| What shared tools in `_system/tools/` already do part of what the plan proposes? | |
| Which derived views (wiki/PATHS.md, atlas/artifact rows) would the plan's changes touch? | |

---

## 2. Forced Second Search (MANDATORY)

You MUST run at least 3 NEW searches that are DIFFERENT from the initial research.

⛔ Do NOT use the semantic-search or language-server MCPs here -- they index the EF codebase, not this workspace. Record search runs through the topic's own tools.

### 2.1 Inverse Record Search (REQUIRED)

Search for the OPPOSITE of what you planned -- why might the approach be wrong?

```bash
# From the topic root: discovery search with question-words phrasing the doubt
python ../_system/tools/query_wiki.py "why <planned approach> wrong <subject>"

# Grep the hub + INBOX for standing corrections/rulings on the subject
rg -n -i "<subject terms>" <hub_file> INBOX.md wiki/CHANGELOG.md
```

At least ONE search must be an inverse/contradicting query. `query_wiki.py` returns candidates only -- verify each via `lookup_id.py`, never conclude absence from a miss.

### 2.2 Impact Search (REQUIRED)

Find everything affected by the planned changes:

```bash
# Who references what you're modifying? (tools)
rg -n "<function/tool name>" ../_system/tools/

# Which entries/pages cite the IDs or files the plan touches? (record)
python ../_system/tools/lookup_id.py <ID>     # prints index row + entry + every also-in file
```

### 2.3 Adjacency Search

Find material physically adjacent to your changes:

```bash
# Files in the same directory as files being modified
ls <dir_of_modified_files>

# Suite checks that assert over the same files
rg -l "<filename or pattern>" ../_system/tools/check_*.py
```

### 2.4 Historical Search (REQUIRED)

Check git history AND what past agent sessions learned:

```bash
# Recent changes to files in the plan (run in the repo that owns them: topic root or _system)
git log --oneline -20 -- <files_being_modified>
```

```
# What did past agents learn about this area? (OPTIONAL -- log-rag is EF-hosted;
# if unreachable, skip silently and fall back to wiki/CHANGELOG.md + git log)
CallMcpTool: user-log-rag / query_logs / {"query": "<what the plan modifies>"}
```

### Search Gate

Before proceeding, show results:

| Search Type | Query/Command | Key Findings |
|-------------|---------------|--------------|
| Inverse | ... | ... |
| Impact | ... | ... |
| Adjacency | ... | ... |
| Historical | ... | ... |

**Any row blank = gate failed.** Complete the searches.

---

## 2.5 Regression Prevention Checks (MANDATORY)

These checks are fast local tool calls. Run ALL of them.

### Check A: Suite Baseline

Run the full check suite from the topic root BEFORE any changes, and record the result:

```bash
python ../_system/tools/run_suite.py
```

A pre-existing red check is a finding: the plan must either fix it or explain why it is out of scope. NEVER plan a commit on top of a red suite.

### Check B: Tool Self-Tests

For every `_system/tools/` module the plan touches, run its self-test now to establish the baseline:

```bash
python ../_system/tools/<tool>.py --self-test
```

### Check C: File Size Check

```bash
wc -l <files_being_modified>
```

Tools approaching 1,000 lines (soft threshold, AGENTS.md Aug 15, 2026) need a split plan at an architectural seam, not a "trim docstrings" plan (docstrings are never trimmed to fit).

### Regression Prevention Gate

| Check | Command | Result |
|-------|---------|--------|
| Suite baseline | `run_suite.py` | all green / list reds |
| Tool self-tests | `<tool>.py --self-test` | pass/fail per tool |
| File sizes | `wc -l <files>` | all under 1,000 (soft) |

---

## 2.6 Panel Triage (MANDATORY when the panel ran)

Collect the panel (AwaitShell the background shell if seats are still running), read its `PANEL.md`, and triage EVERY advisor finding -- including the ones you disagree with:

| # | Advisor finding (which seat) | Verdict | Evidence (your own lookup) | Plan change |
|---|------------------------------|---------|----------------------------|-------------|

- Verdicts: **verified** (your lookup confirms it -- the plan changes), **refuted** (your lookup disproves it -- record the refutation so the next reader doesn't re-raise it), **needs-check** (parked as an explicit open item in the plan, never dropped silently).
- **Panel returns are INTAKE, never primaries:** no finding reaches the plan file on an advisor's word alone. Verify with your own Read/rg/`lookup_id.py` -- an advisor can hallucinate a gap as easily as a fix, and its instruction sheet ("a false gap is as bad as a missed one") is a request, not a guarantee.
- Route survivors where they belong: failure histories → §3.5 premortem; challenged decisions → §3 devil's advocate; assumptions and gaps → §5 synthesis and the plan file.
- Advisor DISAGREEMENT is signal, not noise: a decision two seats read opposite ways gets its §3 devil's-advocate paragraph written with extra care.

---

## 3. Challenge the Architecture

For each major decision in the plan, write one paragraph defending the OPPOSITE approach.

### Format

**Plan says:** "<decision>"
**Devil's advocate:** "<why the opposite might be better>"
**Resolution:** "<kept original because X / updated plan because Y>"

You MUST write this for at least 2 major decisions.

---

## 3.5 Premortem (MANDATORY)

Assume the plan was implemented exactly as written and it FAILED — the suite
went red, a record page got mangled, or the principal rejected the outcome.
Write the failure history in past tense, as if reporting what happened:

**It failed because:** "<the single most likely failure, written as fact>"

The already-happened framing is the point — it produces measurably richer,
more specific failure reasons than "what could go wrong?"
(Mitchell/Russo/Pennington 1989, the study under Klein's premortem; primaries
and the cross-audit that killed the oft-quoted "30%" number:
`_system/archive/advice_decision_frameworks_2026-08-24/`). Then, from that
history:

- **Irreversible steps:** which plan steps cannot be cheaply undone (a
  registered entry, a published page, a deleted file)? Those get done LAST,
  after the reversible steps have validated the approach.
- **Tripwire:** the observable mid-implementation condition that means STOP
  and revert (a self-test failing, a diff touching a file the plan never
  named, an entry count off by one) — named before implementation starts.
- **Invented parameters:** every number the plan introduces (a cap, a batch
  size, a timeout, a threshold) traced to a real constraint or removed — an
  untraceable limit is a premortem finding, not a safety margin.

---

## 4. Gap Analysis Checklist

Check each box ONLY if verified. Unchecked = gap to address.

### Coverage
- [ ] Suite baseline recorded (ran `run_suite.py`, noted green/red)
- [ ] Every touched tool has a `--self-test` (TDD applies to tools here)
- [ ] Record entries the plan cites were re-read THIS session via `lookup_id.py` (scope confirmed, nothing registered since changes their premises)

### Dependency Coverage
- [ ] No new pip installs without the §Environment log line in `_system/SYSTEM_PATTERN.md` (installs are pre-approved but MUST be logged)
- [ ] Import order follows conventions (stdlib, third-party, local); tools use stdlib `re`
- [ ] Any tool printing record text starts with the UTF-8 stdout reconfigure line

### File Size Coverage
- [ ] Modified files stay under 1,000 lines (soft threshold)
- [ ] If >1,000 lines expected, plan specifies a split at an architectural seam

### Risk Coverage
- [ ] Plan has rollback procedure (or changes are trivially reversible -- `git revert`, never `reset --hard`/`stash`)
- [ ] Nothing in the plan touches `archive/`, evidence originals, or hash manifests
- [ ] Nothing in the plan hand-edits hub/wiki/`_INDEX.md` outside the ingest/settle-out protocol
- [ ] No step writes scratch files to the `C:\Topics` root

---

## 5. Synthesis

Write a summary using this template:

```markdown
### Deepening Review Results

**New information discovered:**
1. <What you found that wasn't in the original plan>

**Assumptions validated:**
1. <Assumption + how you verified it>

**Assumptions invalidated:**
1. <Assumption + what the reality is + how this changes the plan>

**Gaps identified:**
1. <Gap + how to address it>

**Plan updates required:**
- [ ] <update 1>
- [ ] <update 2>
```

---

## 6. Update the Plan

If you found anything in step 5, you MUST update the plan file using StrReplace.
(Scope note, Aug 29, 2026: plan-file StrReplace stays legal — the html_edit.py
write-path ban covers scoped published HTML only, never `*.plan.md`.)

**DO NOT proceed to implementation with known gaps.**

(Plan files are deliberately untracked -- the EF repo gitignores `topics_*` plans -- so there is no diff-against-git step; re-read the edited sections to verify the updates landed.)

---

## 7. Final Gate

| Checkpoint | Status |
|------------|--------|
| All adversarial questions answered (no blank cells) | |
| At least 1 inverse record search completed (`query_wiki.py` / hub grep) | |
| At least 1 impact search completed (rg / `lookup_id.py`) | |
| At least 1 historical search completed (git log; log-rag if reachable) | |
| At least 2 devil's advocate arguments written | |
| Premortem written (failure history + irreversible steps + tripwire) | |
| Advisory panel launched + EVERY finding triaged, markers acked (or the skip named: no key / API down) | |
| Gap checklist reviewed (all boxes checked or gaps noted) | |
| Plan file updated with findings | |
| Plan-link handover block in the announce reply AND as the final reply's last lines (`plan_link.py` output, verbatim) | |

**ANY unchecked = plan NOT ready.** Go back and complete it.

---

## 8. Output Format

```markdown
## Plan Deepening Complete: <plan_name>

### Adversarial Review
[Answers to all questions from Section 1]

### Second-Pass Searches
| Search Type | Query | Findings |
|-------------|-------|----------|
| Inverse | ... | ... |
| Impact | ... | ... |
| Adjacency | ... | ... |
| Historical | ... | ... |

### Panel Triage
| # | Advisor finding (which seat) | Verdict | Evidence | Plan change |
|---|------------------------------|---------|----------|-------------|

### Devil's Advocate
**Decision 1:** <decision>
**Counterargument:** <argument>
**Resolution:** <kept original / updated plan>

**Decision 2:** <decision>
**Counterargument:** <argument>
**Resolution:** <kept original / updated plan>

### Gaps Found
- <gap 1 and resolution>
- <gap 2 and resolution>

### Plan Updates Made
- <update 1>
- <update 2>

### Final Status
Plan deepening complete. Ready for implementation — run `/implement-plan <plan_path>`
in a FRESH chat on **Grok 4.6 xhigh** (plan-EXECUTION is a benched cheap class, ~1/12
Fable cost at 10/10 must-catch, Aug 30, 2026; MODELS.md. A plan whose execution needs
mid-leg DESIGN judgment stays frontier — say which and why).

### Plan File
[<plan_filename>](file:///C:/Users/yanti/.cursor/plans/<plan_filename>)
`C:\Users\yanti\.cursor\plans\<plan_filename>`
```

The `### Plan File` lines are `plan_link.py` output pasted verbatim — in the real reply the link stays bare (no backticks, no fence) and these are the LAST lines of the message.

---

## Anti-Patterns

| Don't | Do |
|-------|-----|
| Answer "N/A" to adversarial questions | Explain WHY it doesn't apply |
| Run the same searches as initial research | Run DIFFERENT searches |
| Accept initial assumptions unchallenged | Challenge each assumption explicitly |
| Skip devil's advocate for "obvious" decisions | Write counterarguments anyway |
| Declare plan ready with unchecked boxes | Complete all checkpoints |
| Skip this for "simple" plans | Run it every time -- simple plans have hidden complexity |
| Use the semantic-search MCP for record questions | It indexes EF, not Topics -- use `query_wiki.py` + `lookup_id.py` |
| Claim a fact is absent from the record without the sweep | `lookup_id.py --absence "<term>"` -- THE ABSENCE RULE applies to reviews too |
| Claim "search returned nothing" without showing attempts | Show the queries and results, even empty ones |
| Fold a panel finding into the plan unverified | Returns are INTAKE -- verify with your own lookup first (§2.6) |
| Skip the panel to save cost or because the plan "is simple" | It costs cents and exists because self-review anchors (§Why This Exists) |
| End the turn while panel seats are in flight | AwaitShell and triage in-session -- an untriaged return is wasted spend |
| Launch the panel `:online` or add an anthropic seat "for strength" | Closed-book BARE by design; cross-lab seats are the anti-correlation (§0.5) |
