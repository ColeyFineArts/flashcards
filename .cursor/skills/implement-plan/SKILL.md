---
name: implement-plan
description: "Execute a plan file with batched tasks, verification gates, and test-first tooling. Use when a deepened plan is ready for implementation."
---

# Implement Plan (Topics)

## CLOUD ADAPTER — run FIRST

Topics `_system` missing → print `CLOUD ADAPTER: Topics _system absent — implement via .cursor/plans + Drive/git` and:

1. Skip `plan_link.py` / Topics suite; use workspace `.cursor/plans/*.plan.md`.
2. **Drive:** Sheet/`textContent` only for multi-KB data. After every `create_file`, run `python .cursor/tools/drive_create_verify.py --got-filesize F …` (AGENTS.md §2). ⛔ NEVER spawn child agents to inline xlsx base64.
3. Push feature branch + update PR when the plan says so (cloud is not Topics LOCAL-ONLY).

---

Execute a plan file using batched task execution with verification gates between batches.

Usage: `/implement-plan` (uses most recent plan) or `/implement-plan <plan_path>`

**This is an EXECUTION command. Requires Agent mode (write access).**

If Plan mode is active when the user invokes this, the user is explicitly requesting implementation. Switch to Agent mode and proceed. Do NOT ask "should I switch modes?" -- the command IS the permission.

**Write scope:** the user invoking `/implement-plan` authorizes the writes THE PLAN SPECIFIES -- tool edits, `_system` changes, topic-repo commits the plan calls for. It does NOT suspend record discipline: hub/wiki/`_INDEX.md` content changes still go through the ingest/settle-out protocol, `archive\` and evidence originals stay untouched, and everything stays LOCAL-ONLY (never add a remote, never push).

---

## 1. Load Plan

Find the plan to implement:

```
Glob: `C:/Users/yanti/.cursor/plans/topics_*.plan.md` (sorted by mtime, take first)
Fallback: `C:/Users/yanti/.cursor/plans/*.plan.md`, then workspace `.cursor/plans/*.plan.md`
```

Or use the explicit path provided. ⛔ Never infer the target from open editor tabs (workspace rule: open files are not task context).

**Announce:** "Implementing plan: `<plan_filename>`"

**Same first reply:** paste the plan-link handover block — the two-line output of
`python C:/Topics/_system/tools/plan_link.py --match "<plan filename words>"` (bare
clickable link + plain Windows path), VERBATIM, never hand-built, never backticked or
fenced — and again as the LAST lines of the completion reply. The user rule ("Plan
handover — always give the link") binds every session that touches a plan, but a duty
missing from a skill's own announce/gate/output templates drops in practice: the first
`/implement-plan` run after the sibling plan skills were fixed (Aug 29, 2026, 18:11)
still shipped zero plan links because THIS skill's templates lacked the step.

**Model routing check (same first reply):** plan-EXECUTION legs are a
`cursor-grok-4.6-xhigh` class — benched 10/10 must-catch at ~1/12 Fable cost on the
real $80 html_edit-session replay (Aug 30, 2026; AGENTS.md routing table, receipts in
`_system/MODELS.md`). You know your own model. If this session runs on a Fable/Opus-class
model, add ONE line to the announce — "Routing note: plan execution is benched to
Grok 4.6 xhigh (~1/12 cost); this chat is on <model>. Continuing here — open the next
execution leg on Grok." — then PROCEED (never idle waiting for a model switch he didn't
order; he can Stop and relaunch cheap). The 48h window before the routing row landed
held 4 frontier `/implement-plan` chats at $410, the window's single biggest chat
($226) among them. Plan AUTHORING/deepening judgment stays frontier; a leg that turns
out to need mid-flight DESIGN judgment is not this class (MODELS.md condition 3) — say
so instead if it happens. The continuation footer prints a `! ROUTING` line for the
same mismatch mechanically (`cursor_usage.footer_lines`).

Read the entire plan file. Extract all todos from the frontmatter.

**Load the protocol packs the plan's work requires BEFORE starting** (they are protocol files -- read WHOLE):
- Structural/record work → `_system/SYSTEM_PATTERN.md`
- Any ingest step → `_system/INGEST.md` + the topic's `INTAKE_PROMPT.md`
- Any outbound text → `_system/DRAFTING.md`; a recommendation → `_system/ADVICE.md`;
  a shareable document → `_system/ARTIFACTS.md`
- Any web research beyond a couple of lookups → `_system/RESEARCH.md`
- Work inside a topic → that topic's hub §A (hub filename from `_config.json` `hub`)

---

## 2. Verify Plan Readiness

Before implementing, verify the plan is ready:

| Check | Verify |
|-------|--------|
| Plan has concrete goals section | Required |
| Plan has specific file paths (not "somewhere in _system/") | Required |
| Plan has been deepened (adversarial review, searches) | Recommended |
| Plan todos are all "pending" | Required |
| Owning repo state | `git status` in each repo the plan touches -- a dirty tree means another session's WIP; commit-forward or stop, NEVER stash/reset |

If the plan has NOT been deepened, warn the user:
> "This plan has not been through `/deepen-plan`. Proceeding, but gaps may cause mid-implementation rewrites."

---

## 3. Create Task List

Convert plan todos into a TodoWrite task list. Mark the first task as `in_progress`.

Group tasks into batches of 3 (or fewer if tasks are large/complex).

---

## 4. Execute Batch

For each task in the current batch:

### 4a. Understand Before Modifying

**Before touching ANY file**, read it first. Understand:
- What the current code does
- Who calls it (`rg` for references across `_system/tools/` and topic `tools\`)
- What covers it (which suite checks, which `--self-test`)

⛔ Do NOT use the semantic-search or language-server MCPs -- they index the EF codebase, not this workspace. Record questions go through `lookup_id.py` / `query_wiki.py` / `_INDEX.md` grep.

### 4b. Test-First When Adding Features

Follow the `superpowers:test-driven-development` skill
(`C:\Users\yanti\.cursor\skills\superpowers\skills\test-driven-development\SKILL.md`).
In this workspace "tests" = the check suite + tool self-tests:

When adding a new tool or capability:
1. Write the failing `--self-test` (or a new `check_*.py` assertion) first
2. Implement the feature
3. Verify the self-test passes and the suite stays green

When modifying an existing tool:
1. Run its `--self-test` and `run_suite.py` to establish the baseline
2. Make the change
3. Verify both still pass
4. Extend the self-test to cover the new behavior

### 4c. Implementation Rules

Follow the workspace's standing rules:
- Plain `python` (miniconda base via PATH) -- no conda activation, no venv, never `python3`, never `python -c`/heredocs, never `sleep`
- Use `timeout 30` for commands that might hang
- Never pipe through `| head -N` / `| tail -N` (hook-denied) -- use native flags (`git log -5`, `rg -m 5`) or `tail -n 15 <file>` directly on a file
- Keep tools ≤1,000 lines (soft threshold, AGENTS.md Aug 15, 2026) -- crossing it, split at an architectural seam; NEVER trim docstrings to fit; NEVER whole-file-rewrite an existing tool
- Type hints, meaningful names, PEP 8, comprehensive docstrings
- Tools use stdlib `re`; any tool printing record text starts with `sys.stdout.reconfigure(encoding="utf-8", errors="replace")`
- New pip installs are pre-approved BUT get one line in `_system/SYSTEM_PATTERN.md` §Environment in the same session
- Shared tools live ONCE in `_system/tools/` and run from a topic root reading `_config.json` -- never create a topic-local copy of a shared tool
- No scratch files at the `C:\Topics` root -- session TEXT scratch goes to `D:\cursor_agents\scratch\` (never `$TEMP` -- temp gets deleted, his ruling Aug 26, 2026); downloaded/captured images, PDFs, scans, and zips go to the scrape store (`python ../_system/tools/scrape_dir.py` prints the session folder under `Y:\Images\ScrapedSites\mural-research\`; never delete them)
- After edits, `ReadLints` on modified files

### 4d. Mark Complete

After each task passes verification, mark it as `completed` in TodoWrite.

---

## 5. Between-Batch Gate (MANDATORY)

Before starting the next batch, verify the current batch:

| Check | Command | Evidence |
|-------|---------|----------|
| Suite green | `python ../_system/tools/run_suite.py` (from each affected topic root) | Show output |
| Self-tests pass | `python ../_system/tools/<touched_tool>.py --self-test` | Show output |
| No lint errors | ReadLints tool on modified files | Show output |
| Repo state known | `git status` in each affected repo | Show output |

**Any failure = STOP.** Fix before next batch. Do not accumulate errors across batches. A red check is NEVER committed over.

After verification passes, proceed to next batch. Announce:
> "Batch N complete. All checks pass. Proceeding to batch N+1."

---

## 6. Completion

After all tasks are done:

### 6a. Final Verification

Use the `superpowers:verification-before-completion` skill
(`C:\Users\yanti\.cursor\skills\superpowers\skills\verification-before-completion\SKILL.md`)
to ensure nothing is claimed as done without evidence.

| Verification | Evidence Required |
|--------------|-------------------|
| Suite green in every touched repo | Show `run_suite.py` output |
| Self-tests pass | Show output |
| No lint errors | Show ReadLints output |
| Plan requirements met | Checklist against plan goals |
| All todos completed | Show todo list |
| Plan-link handover block in the announce reply AND as the final reply's last lines | `plan_link.py` output, verbatim |

### 6b. Update Documentation

If the plan changed tooling, conventions, or the environment, update
`_system/SYSTEM_PATTERN.md` (and `C:\Topics\README.md` if user-facing) in the same session.

### 6c. Commit

Commit in EACH repo the plan touched (topic repos and/or `_system` -- they are separate local-only repos), with clear, descriptive messages:

```bash
git add <changed_files> && git commit -m "$(cat <<'EOF'
<commit message>
EOF
)"
```

- ALL suite checks green BEFORE any commit -- no exceptions, no `--no-verify` shortcuts around a red suite
- `git remote -v` must be empty; NEVER push, NEVER add a remote
- Do NOT commit secrets
- **Cheap-seat condition (MODELS.md, binding on grok-class sessions):** when THIS session
  runs on a grok-class model, the leg's diff gets a frontier or audit-queue review
  against the plan's invariants before it counts as merged -- the Aug 30, 2026 bench
  caught every planted error only by code review against the answer key; the seats' own
  green tests missed them all. Topic-repo commits enqueue:
  `python ../_system/tools/audit_queue.py add <topic> --range <before>..<after> --model <your-model> --source implement-plan`.
  For `_system`-only legs, name the pending diff-review in the close so a frontier
  session picks it up.

### 6d. Close With the Significance

Per AGENTS.md: the closing chat reply leads with what the work MEANS in plain terms --
what changed in the story the record/tooling now tells, 2-4 sentences a human reads cold.
File paths, module names, and pipeline mechanics go below the substance, never instead of it.
If any finding belongs in the record and this session is not ingest-authorized for it,
deliver it as dated, paste-ready capture lines in chat for the principal to route.

---

## 7. Output Format

```markdown
## Implementation Complete: <plan_name>

### What this means
<2-4 plain-language sentences: what was found/built and how it moves things forward>

### Execution Summary
- Tasks completed: N/N
- Batches: M
- Plan changes required during implementation: yes/no

### Verification Evidence
run_suite.py / --self-test output showing all green

### Files Changed
| File | Change Type | Repo |
|------|-------------|------|
| _system/tools/foo.py | Modified | _system |
| tools/atlas/bar.py | Added | coyotes |

### Commits
- <repo>: <sha> <message>

### Final Status
Implementation complete. All verifications passed.

### Plan File
[<plan_filename>](file:///C:/Users/yanti/.cursor/plans/<plan_filename>)
`C:\Users\yanti\.cursor\plans\<plan_filename>`
```

The `### Plan File` lines are `plan_link.py` output pasted verbatim — in the real reply
the link stays bare (no backticks, no fence) and these are the LAST lines of the message.

---

## When to Stop and Ask for Help

**STOP executing immediately when:**
- A check fails and the fix is not obvious
- The plan has a gap that blocks the next task
- You hit a missing dependency or environment issue that a logged pip install can't solve
- Implementation reveals the plan's approach is wrong
- A step would require anything in the Hard Safety table (recursive delete, archive edits, history rewrites, remotes) -- these are never done, plan or no plan

**Ask for clarification rather than guessing.** A wrong guess costs more than a question.

---

## When to Update the Plan Mid-Implementation

If implementation reveals plan gaps:
1. STOP implementation
2. Update the plan file with the new information
3. Announce: "Plan gap found: <description>. Updated plan. Continuing."
4. Resume implementation

Small gaps (missing file, extra import): fix inline and note it.
Large gaps (wrong approach, missing feature): stop and discuss with user.

---

## Anti-Patterns

| Don't | Do |
|-------|-----|
| Start coding without reading the target file | Read first, understand, then modify |
| Skip verification because "it's a small change" | Every change gets a suite/self-test verification |
| Run the full suite after every single line change | Run per-task or per-batch |
| Accumulate failures across batches | Fix failures before proceeding |
| Strip docstrings to reduce file size | Split into multiple files instead |
| Ask "should I proceed?" between batches | Verify and proceed autonomously |
| Guess at unclear plan instructions | Stop and ask |
| Commit with the suite red | Fix it in this session, then commit |
| Hand-edit hub/wiki/`_INDEX.md` because the plan says so | Route record content through the ingest protocol |
| Leave a session-built generator in `$TEMP`/`D:\cursor_agents\scratch` once its output is reused | Park it in the serving repo the same session and commit |

---

## Quick Reference: Workspace Commands

```bash
# All from a topic root. Plain python -- no conda, no venv.
python ../_system/tools/run_suite.py                 # full check suite (exit 0 = green)
python ../_system/tools/<tool>.py --self-test        # tool self-test
python ../_system/tools/build_index.py               # rebuild _INDEX.md after page edits
timeout 30 python ../_system/tools/<script>.py       # timeout wrapper for anything that might hang
```
