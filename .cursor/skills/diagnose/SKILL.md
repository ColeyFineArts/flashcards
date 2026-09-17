---
name: diagnose
description: "Find the fastest experiment that yields the most learning for a plan or problem. Hypothesis-first diagnostic cycles with execution. Use when a plan needs validation before committing research spend, long migrations, or implementation time."
---

# Diagnose (Topics)

Find the fastest experiment that yields the most learning, then **RUN IT**.

Usage: `/diagnose` (uses most recent plan) or `/diagnose <path>`

**Stay in Agent mode. This is a WRITE session.** You WILL:
1. Write to the plan/analysis file (Diagnostic Ideas Backlog + Section 7 propagation + Diagnostic Session Summary) -- a diagnose that leaves the plan untouched has not done the work
2. Write one-shot diagnostic scripts and run them (see "Where diag scripts live" below)
3. Park what outlives the session -- durable scripts into a tools directory (committed), record findings into the owning topic's `INBOX.md` -- in THIS session, never "for a write-authorized session"
4. **Execute ALL experiments you can -- there is no human to defer to**

If Plan mode is active when `/diagnose` is invoked, silently `SwitchMode(target_mode_id="agent")` with a one-line explanation and continue. Do NOT ask the user how to handle writes -- the user invoking `/diagnose` IS the override.

**Write posture -- every session is write-authorized (his Aug 20, 2026 ruling, which revoked the /ask zero-writes rule; `_system/ASK.md` §6, `DRAFTING.md`, `ARTIFACTS.md`; the per-topic write lock `topic_lock.py` serializes siblings, so the old race rationale is dead).** Never describe this session as "read-only", "not write-authorized", or "not written anywhere", and never defer a write to some later session. What the label costs: run 362e2aa9 (Sep 2, 2026) told itself "read-only against the record" three times, skipped the hub §A read on that reasoning, left `terminal_census.py` in scratch "which needs a write-authorized session", and handed the principal five dated capture lines to route by hand -- the round trip his rules forbid. Write targets, all in-session:

| What | Where | How |
|---|---|---|
| Diagnosis results, plan corrections, backlog, summary | The plan file | Sections 6-7 -- MANDATORY, never optional |
| One-shot scripts, captured outputs | `D:\cursor_agents\scratch\` | Plain files (never `$TEMP` -- temp gets deleted, his ruling Aug 26, 2026; never the `C:\Topics` root) |
| A script whose output was reused/sent or that proves durable | `_system/tools/` (≤1,000 lines soft, `--self-test`, UTF-8 stdout reconfigure) or the owning topic's `tools\` | `_system`: `--self-test` green, then `git -C C:/Topics/_system add tools/<x>.py` + `git -C C:/Topics/_system commit -m "<msg>" -- tools/<x>.py` (a NEW file has no sibling hunks; a pre-existing path gets `commit_hunks.py --root _system list -- <path>` first). Topic: `python ../_system/tools/topic_lock.py close -m "<msg>" -- tools/<x>.py` from the topic root (the atomic suite-gated close) |
| A record finding (false index row, unregistered fact, tension, verified public fact) | The owning topic's `INBOX.md` | `python ../_system/tools/record_edit.py park --text "<dated capture line>" --keep` from the topic root -- raw intake, no ceremony; `--keep` because this session will not run the queue-close that would otherwise consume its own lines. A finding that needs an ENTRY is an ingest (`_system/INGEST.md` WHOLE + the topic's `INTAKE_PROMPT.md`); park it first, then ingest it if the session has the budget, else name it in the summary WITH the park already in place |
| A tool/check defect in `_system` | The plan (Implementation Notes / Constraints) | Written in Section 7; if a topic record is affected as well, park the topic-side line too |

**Experiments measure, they never mutate:** greps, `lookup_id.py`, `query_wiki.py`, counts and dry-runs on COPIES in scratch -- an experiment that edits the hub/wiki it is measuring, or the byte-exact `archive/`, is not an experiment. That is discipline about what an EXPERIMENT does, not a label on the session: the parks, tool commits and plan edits above are the session's writes and they happen.

**Reading is not writing:** a diagnosis that runs against a topic's record starts like every session inside a topic -- `python ../_system/tools/topic_config.py --get hub` names the hub, read its §A before the first experiment. "I'm not authoring anything in the topic" was 362e2aa9's reason to skip it; it is not one.

**Where diag scripts live:** one-shot scratch scripts go in `D:\cursor_agents\scratch\` (never the `C:\Topics` root, never `$TEMP` -- past diag scripts stay searchable for later sessions). If a script's output gets reused, sent, or the script proves durable, park it the same session per the table above -- `_system/tools/` or the topic's `tools\` -- and commit it there.

**Execution rules:**
- **T0-T2 (<=30 min):** Execute immediately. **Chain results** -- if experiment #1's output
  reveals a new cheap question (file to check, pattern to grep, entry to look up), pursue
  it BEFORE moving to experiment #2. Sections 3-5 are a LOOP, not a checklist.
- **T3 (30-120 min):** Execute with a sample/`--limit` to keep runtime under session budget.
- **T4-T5 (hours -- deep-research passes, bulk fetches):** Launch in background
  (`block_until_ms: 0`), then continue with remaining T0-T2 experiments. Read
  `_system/RESEARCH.md` WHOLE before launching any deep-research run.

---

## Section 0: Find the Document

**Detection order (follow exactly):**

1. **User-specified path** -- if `/diagnose <path>`, use that path.
2. **Most recent Topics plan** -- Glob `C:/Users/yanti/.cursor/plans/topics_*.plan.md`, pick most recently modified.
3. **Any recent plan** -- Glob `C:/Users/yanti/.cursor/plans/*.plan.md` and workspace `.cursor/plans/*.plan.md`.
4. **Ask user** -- "No analysis document detected. What file should I diagnose?"

⛔ Do NOT infer the target from open editor tabs or recently viewed files -- in this workspace open files are explicitly NOT task context (workspace rule).

**Announce:** "Diagnosing: `<filename>`" — and in the SAME first reply paste the target's handover block: `python C:/Topics/_system/tools/plan_link.py --match "<filename words>"` prints it for plans (two lines: bare clickable link + plain Windows path — verbatim, never hand-built, never backticked or fenced); a non-plan target gets the same two-line shape by hand. The block repeats as the LAST lines of this skill's final reply. The user rule ("Plan handover — always give the link") binds every session that touches a plan, but a duty missing from a skill's own announce/completion templates drops in practice — /deepen-plan run f2445d54 shipped link-less (Aug 29, 2026, his third re-ask: "Yet again I don't know it").

---

## Section 1: Extract Diagnostic Targets

**For `.md` plan files:** Read the FULL file (`Read(file)` -- no offset, no limit).
**For code files under 500 lines:** Read the whole file.
**For code files over 500 lines:** Use Grep to find targets via these markers:

- **Status:** `Open Question`, `Still open`, `NOT yet`, `needs verification`, `TBD`, `Unknown`, `TODO`, `FIXME`
- **Structure:** `Phase [A-Z]` (check COMPLETE/CANCELLED), `Next Steps`, `Next Session`, `Recommendation`, `Options`
- **Science:** `Hypothesis`, `Testable Prediction`, `Missing Experiment`, `Experiment`, `Question`

**Present diagnostic targets as a numbered list with type classification:**

| Type | Definition | Example |
|------|-----------|---------|
| **HYPOTHESIS** | Core assumption the plan bets on -- must be tested | "the FTS5 index will surface entries the keyword grep misses" |
| **DIAGNOSTIC** | Unknown answer -- requires investigation | "Does the record actually hold a date for the 1976 purchase?" |
| **EXECUTION** | Known task -- just needs doing | "Add the new check to run_suite.py's roster" |

**Classify EVERY target.** This determines what kind of experiments to run.

**Readiness checks are NOT targets.** "Does the tool exist?" / "Is the repo clean?" confirm
prerequisites -- they teach you nothing about whether the approach will work.

| Condition | Action |
|-----------|--------|
| Any HYPOTHESIS targets | **Section 1.5 FIRST** (mandatory). Then DIAGNOSTIC, then EXECUTION risk-probes. |
| All targets DIAGNOSTIC | Normal flow -> Section 2 |
| **ALL targets EXECUTION** | **The plan is fully diagnosed.** Shift to pre-execution risk assessment: for each target ask (1) What could block this? (2) What assumption hasn't been verified? (3) What's the fastest execution sequence? Second-order assumptions ARE diagnostic targets -- an index rebuild assumes `_INDEX.md` regenerates clean, a check change assumes no other topic's repo goes red. |
| All questions resolved | Announce: "No open diagnostic targets found." STOP. |

---

## Section 1.5: Hypothesis Gate (MANDATORY when the plan includes EXPENSIVE work)

**"Expensive" here = deep-research spend, bulk web fetches/downloads, a record-wide
migration, an atlas/artifact republish, or hours of agent time. If the plan proposes any
of these, you MUST test the core hypothesis BEFORE readiness checks.**

**Step 1 -- State the hypothesis in one sentence.** What does the plan bet will improve?
- "A deep-research pass on the museum's archives will surface a pre-1970 reference"
- "Splitting check_text_health into two modules keeps both under 300 lines"
- "The new PATHS coverage check will catch the five skipped stations"

**Step 2 -- Find the cheapest experiment that tests it.** Use this table:

| Plan proposes... | Cheap hypothesis test (T1-T2, <=15 min) |
|---|---|
| A deep-research pass | ONE targeted web search on the narrowest sub-question first -- does any public signal exist at all? |
| A new/changed suite check | Run the draft check against ONE topic (report mode, fix nothing yet); count true/false positives |
| A record-wide migration or reformat | Apply to a COPY of one wiki page in `D:\cursor_agents\scratch\`; diff and eyeball |
| A new tool or tool refactor | Write the `--self-test` FIRST and run it against current behavior (TDD applies to tools) |
| An index/discovery improvement | Run 5 known-answer queries through `query_wiki.py` before/after; count hits |
| A threshold/cap change | Grep how many existing pages/entries would violate the new value today |

**Step 3 -- Run it NOW.** This is your #1 ranked experiment regardless of tier.

**Step 4 -- Record the result.** If hypothesis supported -> proceed. If contradicted ->
plan needs revision BEFORE spending the hours.

---

## Section 2: Check Existing Backlog + Past Experiments

**(a)** Grep the doc for `## Diagnostic Ideas Backlog`. If found, read it.

**(b)** Check completion status using:

| Signal | How |
|--------|-----|
| Doc text | Grep for result text matching the experiment |
| Log-RAG (OPTIONAL -- EF-hosted; if unreachable skip silently) | `CallMcpTool: user-log-rag / query_logs / {"query": "<plan topic> experiment results"}` |
| Session record | Grep the topic's `wiki/CHANGELOG.md` + `git log --oneline -20` in the owning repo |

Mark items as **Completed** / **Superseded** / **Still Pending**.

---

## Sections 3-5 are a LOOP (not a single pass)

**The core diagnostic cycle is: Brainstorm -> Rank -> Execute -> Reassess -> (loop or exit).**

```
+--> Section 3: Brainstorm (new targets only on cycle 2+)
|    Section 4: Rank
|    Section 5: Execute
|    Section 5.5: Reassess --> new T0-T2 questions? --YES--> loop back
|                               |
|                               NO / cycle=3 / context>80%
|                               v
|                          Section 6: Update Backlog
|                          Section 7: Propagate Learnings to Plan
```

Iteration cap: **3 cycles**. Announce "--- Diagnostic Cycle N ---" before each.

---

## Section 3: Brainstorm Experiments (Tier Walk)

**Walk tiers in order for each open target.** T-1 first. If T-1 or T0 answers it, STOP --
do not brainstorm higher tiers for that target.

**On cycles 2+:** Only brainstorm for NEW targets added by the previous cycle's reassessment.
Skip targets already addressed in prior cycles.

⛔ Never use the semantic-search or language-server MCPs here -- they index the EF codebase, not this workspace.

### T-1: Agent Memory (0 min)

Query what past sessions already learned -- they may have run the exact experiment you're planning.

```
CallMcpTool: user-log-rag / query_logs / {"query": "<topic>"}   # OPTIONAL: skip silently if down
```
```bash
rg -n -i "<topic terms>" wiki/CHANGELOG.md
git log --oneline -20
```

### T0: Existing Record/Data (0 min)

Check what is already on disk:
- `_INDEX.md` (grep-only, never read whole) + `lookup_id.py <ID>` for any entry in play
- Hub §B/§D, `INBOX.md`, `wiki/PATHS.md` for tracked subjects
- `files\` and `archive\` listings (originals are byte-exact and untouchable)
- `python ../_system/tools/find_book.py "<title/author>"` BEFORE any book/PDF fetch
- Prior tool output files, `D:\cursor_agents\scratch\` files from this session

### T1: Single Command (1-5 min)

| Investigating... | Run This (from the topic root) |
|---|---|
| Whole-record health | `python ../_system/tools/run_suite.py` |
| One entry's full footprint | `python ../_system/tools/lookup_id.py <ID>` |
| Is X in the record at all? | `python ../_system/tools/lookup_id.py --absence "<term>"` |
| Vocabulary-mismatch discovery | `python ../_system/tools/query_wiki.py "<question words>"` |
| Index freshness / ID resolution | `python ../_system/tools/retrieval_drill.py` |
| Local book/PDF existence | `python ../_system/tools/find_book.py "<title>"` (or `--isbn`) |
| Draft AI-ism sweep | `python ../_system/tools/check_aiisms.py <file>` |
| A single suite check in isolation | `python ../_system/tools/check_<name>.py` |
| Tool behavior | `python ../_system/tools/<tool>.py --self-test` |

### T2: Small Scripted Experiment (5-30 min)

A one-shot script in `D:\cursor_agents\scratch\` (plain `python`, UTF-8 reconfigure if it prints record text):
- Parse/count something across wiki pages (measure, never mutate)
- Dry-run a proposed transformation against a COPY of one page
- Batch 10-20 `lookup_id.py` calls and tabulate

### T3: Full Pass (30-120 min)

Record-wide measuring sweep, a full `build_index.py` + suite run across topics, or an
exhaustive verification pass with a sample cap for directional signal first.

### T4-T5: Large Scale (hours)

Deep-research runs (`deep_research.py` / `openrouter_research.py` per `_system/RESEARCH.md`
-- read it WHOLE first), bulk downloads, full migrations. Launch in background.

---

## Section 4: Rank by Diagnostic Value

Rank using THREE dimensions, in priority order:

1. **Target type** -- HYPOTHESIS > DIAGNOSTIC > EXECUTION risk-probes. Always.
2. **Decision Narrowing** -- How many downstream options does this result eliminate?
3. **Tier** -- Lower tier wins when decision narrowing is equal.

**Present as a markdown table:**

```
| Rank | Experiment | Question It Answers | Type | Tier | Est. Time | Options Eliminated If... | Investigations Cancelled If... |
```

"Options Eliminated" forces concrete reasoning. "Investigations Cancelled" names what becomes
unnecessary -- agents continue planned work after negative results unless this is explicit.

**Self-check:** If >50% of your ranked experiments are readiness checks, you haven't
found the real diagnostic targets. Go back to Section 1.5.

---

## Section 5: Execute Experiments

**Execute ALL ranked experiments, starting from #1.**

**Per-experiment sequence:**

1. State the question (1-2 sentences MAX)
2. Run the command NOW: `timeout <tier-appropriate> python ...` (T1: 30s, T2: 300s, T3+: use `block_until_ms: 0`). Plain `python`, never `python3`, never `python -c`, never `sleep`.
3. Read the output -- **scan for checkable references** (file paths, entry IDs, error strings, counts)
4. State which options are eliminated (1-2 sentences)
5. **If output contains a checkable reference you haven't inspected:** pursue it NOW before the next experiment. This is the iteration trigger.
6. Update the backlog entry to **Done** with the result

**Stop executing when:**
- All experiments in THIS CYCLE are completed or launched, OR
- Session is running low on context (>80% used) -- record remaining in backlog

**Execution strategy by tier:**

| Tier | Action |
|------|--------|
| T-1, T0 | Already done in Section 3. |
| T1 (1-5 min) | Run inline. Wait for result. |
| T2 (5-30 min) | Run inline. Wait for result. |
| T3 (30-120 min) | Run with a sample cap for directional signal. |
| T4-T5 (hours) | Launch with `block_until_ms: 0`. Record PID. Continue with remaining experiments. |

---

## Section 5.5: Reassess & Loop

**IMMEDIATELY after executing -- before writing ANY summary -- fill this table:**

```
| Experiment | Result (1 line) | New Question? | Tier | Action |
```

For EACH completed experiment, fill one row. **"None" is only acceptable for EXPLORATORY
experiments** (where the result genuinely closes a question). For CONFIRMATORY experiments
("is X still true?" -> "yes"), "None" is almost always wrong -- ask "What would BREAK this
assumption during execution?" If you wrote "None" for >50% of rows, you ran the wrong experiments.

**Trigger rules:**

| When you see this in output... | Do THIS immediately |
|-------------------------------|---------------------|
| A file path you haven't read | Read/Grep it (T0, 0 min) |
| An entry ID you haven't resolved | `lookup_id.py <ID>` (T0, 0 min) |
| An error message or traceback | Grep source for the error string (T0, 0 min) |
| A count that's surprisingly high/low | Check the input data or config that produced it (T0-T1) |
| A function name you don't understand | Read it (T0, 0 min) |

**Decision after filling the table:**

| Condition | Action |
|-----------|--------|
| New T0-T2 questions AND cycle < 3 | Loop back to Section 3. |
| New questions but all T3+ | Record in backlog. -> Section 6. |
| No new questions | -> Section 6. |
| Cycle = 3 or context > 80% | Record remaining. -> Section 6. |

---

## Section 6: Update Plan with Backlog

Insert or update a `## Diagnostic Ideas Backlog` section in the plan doc.

```markdown
## Diagnostic Ideas Backlog

*Generated by /diagnose on <date>. Re-run /diagnose to re-rank.*

| # | Experiment | Question | Tier | Est. Time | Cancels If Clear | Status |
|---|-----------|----------|------|-----------|-----------------|--------|
| 1 | ... | ... | T0 | 0 min | ... | **Done: result=X** |
| 2 | ... | ... | T1 | 3 min | ... | **Done: result=Y** |
| 3 | ... | ... | T3 | 45 min | ... | Pending |
```

On re-runs: update statuses, re-rank, add new experiments, prune completed items.

---

## Section 7: Propagate Learnings to Plan

**The backlog tracks experiments. The PLAN tracks strategy.** Propagate what you learned
INTO the plan body.

| When this happened... | Update the plan by... |
|---|---|
| Experiment eliminated an approach | Mark `~~ELIMINATED~~` with 1-line reason WHERE it appears in the plan |
| Open question is now answered | Replace `TBD`/`Unknown`/`Open Question` with the answer and source |
| Phase is no longer viable | Update phase status to `CANCELLED -- <reason>` |
| New constraint discovered | Add to constraints section (create if absent) |
| Results change which phase should execute next | Update phase ordering or add `PRIORITY CHANGED` annotation |
| Dead end confirmed | Add to `### Dead Ends` subsection (create if absent) |
| New implementation instructions emerged | Add to `### Implementation Notes` subsection (create if absent) |

**Write a `### Diagnostic Session Summary` at the end of the plan:**

```markdown
### Diagnostic Session Summary (<date>)

**Cycles run:** N
**Key findings:**
- <finding 1>
- <finding 2>

**Options eliminated:** <list>
**Dead ends confirmed:** <list, or "None">
**Recommended next action:** <1-2 sentences>
**Unresolved questions:** <list, or "None">
```

**Findings that outlive the plan** (a resolved record question, a false index row, a
discovered tension, a verified public fact) do NOT get silently dropped and do NOT get
handed to the principal to route: park each one NOW into the owning topic's `INBOX.md` --
`python ../_system/tools/record_edit.py park --text "<YYYY-MM-DD> <finding, with the
file/ID/tool it names>" --keep` from that topic's root (see the write-posture table at the
top). The summary in chat then reports them as PARKED (topic + line count), not as lines to
paste. A `_system`-only finding (tool stdout shape, a check that should exist) lives in the
plan's Implementation Notes / Constraints -- it is already written by Section 7.

---

## Section 8: Anti-Patterns

**Tier discipline:**

| Don't | Do |
|-------|-----|
| Jump to T4+ before exhausting T-1->T2 | Walk tiers in order. Changelog/git (0 min) -> record on disk (0 min) -> one tool call (3 min) -> full passes |
| Suggest "run a deep-research pass" first | A 3-min `lookup_id.py`/grep answers 60% of questions |
| Fetch a book/PDF before `find_book.py` | Search the local index FIRST -- one miss ≠ absence, try a second phrasing |

**Hypothesis discipline:**

| Don't | Do |
|-------|-----|
| Run 8+ readiness checks before testing the hypothesis | Hypothesis test is ALWAYS experiment #1 |
| "Is X ready?" as a diagnostic experiment | Only "will X actually work?" is diagnostic |
| >50% readiness in ranked experiments | Re-read the plan: what does it ASSUME will work? Test THAT |

**Execution discipline:**

| Don't | Do |
|-------|-----|
| Present commands without running them | You ARE the actor. Run everything. |
| Narrate >200 words between sections | Targets -> ranked table -> results -> backlog |
| Stop after writing backlog without executing | Identifying experiments = 10%. Running = 90%. |
| Call the session "read-only" / "not write-authorized"; hand the principal capture lines to route; defer a park or a tool commit to "a write-authorized session" | Every session is write-authorized (Aug 20, 2026 ruling). Park the finding (`record_edit.py park --text ... --keep`), commit the tool, update the plan -- in THIS session (362e2aa9 did none of the first two, Sep 2, 2026) |
| Mutate the hub/wiki/`archive/` you are measuring as part of an experiment | Experiments measure on the record or on scratch COPIES; the session's writes are the plan, parks, and tool commits |
| Report a fact as absent without the sweep | `lookup_id.py --absence` -- THE ABSENCE RULE |

**Iteration discipline:**

| Don't | Do |
|-------|-----|
| Skip Section 5.5 reassessment table | Fill the table. Every time. |
| Treat ranked list as rigid checklist | Output contains checkable reference -> pursue it |
| Run 10 confirmatory checks -> 10x "None" | Shift to risk-probing: "what breaks this?" |

**Propagation discipline:**

| Don't | Do |
|-------|-----|
| Skip Section 7 plan propagation | Backlog = tracking. Plan = what implementing agent reads. |
| End without Diagnostic Session Summary | 5-line summary saves 5000 tokens of re-discovery |
| Leave answered `TBD`/`Open Question` markers | You answered it. Update the marker. |
| Leave a reused diag script in `D:\cursor_agents\scratch` | Park it in `_system/tools/` (or the topic's `tools\`) and commit it this session -- commands in the write-posture table |

---

## Completion Criteria

You are done when ALL are true:
1. **Hypothesis gate passed** (if plan includes expensive work, core hypothesis tested)
2. All feasible experiments executed (T0-T2 inline, T3 sampled, T4+ launched)
3. Results recorded in backlog with options eliminated
4. Plan body updated with findings (Section 7)
5. Section 5.5 reassessment found no new T0-T2 opportunities (or cycle cap reached)
6. Diagnostic Session Summary written
7. Nothing deferred to another session: every finding that outlives the plan is PARKED in its topic's `INBOX.md`, every reused/durable script is committed to a tools directory, and the reply nowhere calls the session "read-only" or asks the principal to route anything
8. Final reply ends with the diagnosed file's handover block (for plans: `plan_link.py` output, verbatim)

**"I wrote a backlog with 5 pending experiments" is NOT done.** You should have executed
experiments 1-3 and launched experiment 4. The only acceptable "Pending" items are
experiments that genuinely cannot run yet. **"Capture lines for routing" is NOT done
either** -- that is a park you did not run.

---

## Quick Reference: Workspace Commands

```bash
# All from a topic root. Plain python -- no conda, no venv, never python3/python -c.
timeout 30 python ../_system/tools/run_suite.py          # full check suite
timeout 30 python ../_system/tools/lookup_id.py <ID>     # entry + every also-in file
timeout 30 python ../_system/tools/lookup_id.py --absence "<term>"
timeout 30 python ../_system/tools/query_wiki.py "<question words>"
timeout 30 python ../_system/tools/find_book.py "<title or author>"
timeout 30 python ../_system/tools/retrieval_drill.py
timeout 30 python ../_system/tools/<tool>.py --self-test
```
