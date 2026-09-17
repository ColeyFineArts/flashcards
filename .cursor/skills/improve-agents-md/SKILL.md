---
name: improve-agents-md
description: Analyze past agent chat transcripts to find user friction (draft corrections, fact challenges, re-pushes, misread intent) and wasted tokens, then improve BOTH the canonical AGENTS.md/rules/skills/commands AND the Topics programming infrastructure (_system/tools, check suite, hooks). Use when the user asks to review chats, analyze agent sessions, mine transcripts for failure patterns, improve/update AGENTS.md, or improve the workspace tooling based on agent behavior. Any text given with the invocation is the steer for that run.
disable-model-invocation: true
---

# Improve AGENTS.md + Infrastructure From Chat Analysis (Topics)

## CLOUD ADAPTER — run FIRST

```bash
test -f /c/Topics/_system/tools/transcript_census.py || test -f C:/Topics/_system/tools/transcript_census.py
```

Missing → print `CLOUD ADAPTER: Topics _system absent — mining cloud-agent transcripts; writing fixes to this repo AGENTS.md + .cursor/skills|rules/` and use the table below. Present → ignore this block.

| Topics | Cloud substitute |
|--------|------------------|
| `transcript_census.py --since …` | `cursor-cloud` `list-cloud-agents` + `batch-fetch-details` (`includeTranscripts: true`); park `/home/ubuntu/.cursor/scratch/imp_*` |
| Raw JSONL under `agent-transcripts/` | One **subagent per** `…/<bcId>/transcript.json` — never load MB transcripts into the parent |
| `delegation_census.py` | Skip one line if absent; still pattern-scan assistant text |
| Edits in `_system/workspace/` | This repo `AGENTS.md` + `.cursor/skills\|rules\|tools/`; push feature branch + `ManagePullRequest` |
| LOCAL-ONLY no-push | Cloud agents **do** push |

Do not invent a fake Topics `_system`. Quantify format + top-1–3 fix cap still bind. Floor-first: prefer forcing mechanisms (`drive_create_verify.py`, alwaysApply rules) over re-worded prose.

---

Mine past agent conversations for the moments where the principal had to correct,
re-explain, or re-push the agent; quantify them; turn the worst offenders into durable
fixes. This is not a coding workspace: the waste that matters here is NOT failed
commands — it is the principal fact-checking a drafted email line by line, flagging
AI-sounding text, restating his own positions, or asking for the same lookup twice.
Every correction he types is a doc gap, an enforcement gap, or a TOOLING gap; the goal
of a run is to make the next session already know — or be mechanically unable to
repeat — what this one had to be told.

The deliverable is not always prose. This workspace's history shows the strongest fixes
are mechanisms: `lookup_id.py` replaced 240-grep manual chains, `check_aiisms.py`
replaced an ignored ban-list exhortation, the atlas scope/grade/refs checks replaced
eyes-only review. Every run audits BOTH layers: the guidance AND the programming
infrastructure that enforces it.

## Invocation steer

Any text he gives with the invocation (`/improve-agents-md <text>`) is the STEER for
this run: it directs the pass and outranks the default heuristics wherever they
conflict — which transcripts to mine (a named date range, topic, or session set beats
the since-last-run window), which friction category to chase, which finding to fix
first (a steered target gets fixed even if unsteered ranking would rank it below the
top 1–3), which layer to audit (guidance vs tooling). The steer changes WHAT gets
mined and fixed, never HOW: the quantify format, dedup gate, canonical-copies-only
routing, self-tests, and the sibling check still bind. A bare invocation runs the
default full workflow.

## Workflow Checklist

```
- [ ] Step 1: Locate and extract chat transcripts
- [ ] Step 2: Scan for friction patterns + harvest principal facts
- [ ] Step 3: Audit the programming infrastructure (transcript-driven + static)
- [ ] Step 4: Quantify each finding (required format)
- [ ] Step 5: Deduplicate against existing guidance AND tooling (BLOCKING)
- [ ] Step 6: Route and write the fix (canonical copies only)
- [ ] Step 7: Install, verify, commit
```

## Step 1: Locate and Extract Transcripts

Transcripts live at `C:\Users\yanti\.cursor\projects\c-Topics\agent-transcripts\<uuid>\<uuid>.jsonl`
— one subdirectory per chat, one JSONL line per message (`*/subagents/*` are subagent
runs, never the chat). UUID names carry no timestamp; selection orders by file mtime.
**A transcript is NOT the whole chat:** when a long session's context is summarized,
the JSONL is REWRITTEN to the post-summary tail — a 113 KB file can be missing the
session's first hours (the 27bc51e0 labs pass, Sep 1, 2026: the mass-mint order and
36 of its 37 lab registrations predate line 1; the tail alone reads like a chat that
"dropped" the order). Before claiming a chat never did / dropped something, check
line 1 for a `[Previous conversation summary]` marker and cross-check `git log`
over the session's window — a transcript miss is never proof of session behavior.

**Census + extraction is ONE call each — `transcript_census.py` — never an ad-hoc
find/jq/rg loop.** The loop shapes kept zeroing out SILENTLY: Aug 26, 2026, a per-file
jq loop missing its `cd`, with `2>/dev/null` eating every "Could not open file",
printed 0 real messages for all 76 transcripts; Aug 29, 2026, a census printed 0 for a
known-nonempty window when find(1) got rg's backslash paths — both runs nearly skipped
their whole windows. The tool is pathlib-native (no find, no cd, no quoting) and LOUD:
per-file errors print and count, and an empty window or a
files-present-but-nothing-parsed result exits nonzero instead of passing silent zeros
as data.

```bash
python C:/Topics/_system/tools/transcript_census.py --since "$(git -C /c/Topics/_system log -1 --grep improve -i --format=%ci)"
python C:/Topics/_system/tools/transcript_census.py --extract <uuid-prefix>   # one chat's real user messages
```

The `--since` subshell pastes git's own `%ci` time (offset included) — the
census converts an offset-aware bound to local itself (Sep 2, 2026: the
hand-pasted `%ci` string died with a naive-vs-aware TypeError on the first
transcript, before `parse_since` existed).

The census prints real-user-message counts per main-chat transcript, mtime-sorted
(REAL = contains `<user_query>`, excluding system text: the notice shapes "Briefly
inform the user about the task result" / "Perform any necessary follow-up actions"
and "STOP GATE:" hook injections — none of it is his typing; 77% of the Aug 16 run's
extracted "user" messages were notices before the first filter existed). **Triage on
the `real` column** (the Aug 19 run: 338 transcripts, ~300 automated loop passes —
reading loop transcripts as if they were conversations wastes the run): read every
chat with `real >= 3` via `--extract`; `real == 0` rows and near-empty files are
automated passes — pattern-scan them only.

**Window = since the last run.** Find the previous pass's commit
(`git -C /c/Topics/_system log --grep improve -i --oneline`) and scan only sessions with
mtime after it — earlier sessions were already mined, and re-finding their patterns
wastes the run (the Aug 16 pass windowed on the Aug 15 lean-down commit this way).

**Check for a live sibling improve run BEFORE writing anything (BLOCKING).** He
launches `/improve-agents-md` from multiple chats in the same hour (three invocations
within ~35 minutes on Aug 21, 2026 — one was still mid-run, editing `_system` tools,
when the next started), and all runs write the same `_system` files. Two checks:
`git -C /c/Topics/_system status --short` (dirty files this run didn't touch = a
sibling mid-work — never stage, revert, or "fix" them; stage explicit paths only), and
the tail of any in-window transcript still growing (a session actively narrating this
skill's steps owns the findings it is narrating). A live sibling's findings are TAKEN:
scope this run to findings it is not fixing — supplement, never duplicate — and treat
a `_system` `index.lock` as the sibling committing (wait seconds, retry). "We found the
same bug so my fix is harmless" = NO — same-file edits conflict at commit.

Optional pre-scan: the log-rag MCP (`analyze_waste`, `count_pattern`, `query_logs`)
indexes past sessions and can rank candidates before manual extraction. It is
EF-hosted — if unreachable, skip silently (AGENTS.md MCP scope map); never block on it.

**Extract output is the reading surface** (NEVER read a raw JSONL whole — single
lines run to tens of KB — and never use inline Python; `jq` stays fine for ad-hoc
single-file digs). `--extract` prints the text INSIDE `<user_query>` (command
invocations re-inline the full attached SKILL.md BEFORE his text, so a head preview
shows only the skill header and LOSES what he actually typed — the Aug 19 run hit
this twice), each message's `<timestamp>` tag, and long bodies previewed at 400 chars
(`--full` for whole messages). Long messages are pasted evidence/commands — the
preview is enough; short messages are where corrections live: read every one in
full. **Verbatim-duplicated user messages are NOT re-type friction until their
timestamps differ:** the client re-delivers queued messages, so same-minute copies
are transport artifacts — count a repeat as a failed-turn signal only when minutes
passed between copies (Aug 21 run: three same-minute triplicates nearly counted as
stalls; the real Aug 20 galaxy re-pastes were minutes apart).

For assistant-side checks (permission stalling, time-constraint invention,
attribution claims, banned shell patterns, tool tracebacks), use census
`--pattern <regex>`: it counts hits in ASSISTANT-authored content only (text,
thinking, tool_use inputs) — raw rg over a transcript false-positives on the embedded
AGENTS.md/skill text riding every user message; user-side `rg -ci` detections run
against `--extract` output for the same reason. A pattern that STARTS with `/`
(a command name: `/improve-agents-md`) must be written `\/improve-agents-md` —
Git Bash path-converts a leading-slash argument into `C:/Program Files/Git/...`
before Python sees it; the census now refuses that shape with the pattern you
typed (Sep 1, 2026: it printed "0 across 0 chats" over two real invocations
before the guard existed). Park outputs worth keeping in
`D:\cursor_agents\scratch\` under session-unique names (`imp_$$_*.txt` — the scratch
home is SHARED across parallel sessions and two same-minute sessions clobbered fixed
names, Aug 24, 2026; never `$TEMP` — temp gets deleted, his ruling Aug 26, 2026, and
past extractions stay searchable for later runs), never into the workspace. Any
remaining raw-transcript rg keeps the `cd` and the command in ONE call (**shell cwd
does NOT reliably persist between tool calls** — each call can start a fresh shell at
the workspace root, `$$` changes too), and an all-zeros count from a known-nonempty
window is the tool failing, not the data.

## Step 2: Scan for Friction Patterns

| Category | What it looks like in this workspace | Detection idea | Threshold |
|----------|--------------------------------------|----------------|-----------|
| COST_COMPLAINTS | He complains about spend or model routing: "save me money", "burning cash", "why is this expensive agent doing all this", "could be delegated to a cheaper model" | `transcript_census.py --money` over the window (user-side sweep, quotes printed — pasted invoice/record text is noise; read the quotes, keep his asks) | any genuine = TOP-priority steer he already typed. ~20 distinct money asks over 5 days (157 raw hits, 62 chats) never became a finding before this row existed — every category below was drafting/record class (Aug 30, 2026) |
| TAIL_TURNS | Injected "Briefly inform" / STOP GATE turns after a reply — a Shell (usually the footer helper, 30–120 s under load vs the 30 s default block) was still running at turn end; each tail reply bills a full context re-read, and a tail that re-runs the footer chains | `tail_turn_census.py --hours 48` (footer-call outlive rate, chains, cause of the last Shell, IDE billing join; `--since/--until` for before/after) | footer outlive rate >10% or any chain = problem. At discovery: 48% of footer calls, 77% of chats, chains of 10, $185 matched = 5.1% of 48 h spend (Sep 2, 2026) — fix is F-TAIL-TURN (kernel) + the `! TAIL-TURNS` footer advisory; re-run to verify adoption |
| SOURCE_READS | Sessions learn a house tool by Reading its SOURCE (`Read _system/tools/<tool>.py limit=80` at the point of first use, or Grep-then-Read at a deep offset for a subcommand) — one billed step per tool, the 80-line head then riding every later send; cheap loop passes repeat it every pass because each pass is a fresh session | the `SOURCE-READS` block of `delegation_census.py --since-last-improve` (steps, chats, unconverted chats at the footer fire line, toolcard launches) | any window where steps stay flat and launches stay ~0 = the kernel line is not reaching the seats (enforcement gap). Baseline Aug 30–Sep 2, 2026: 4,104 steps / 734 chats (~5.8 per cheap loop pass; record_edit 500, cheap_read 331, scrape_dir 247, topic_lock 243, pdf_text 206), est. ~$505 of billed steps = 8.7% of the window's IDE spend — fix is `toolcard.py` (one bounded call: docstring head + `--help` per tool named) + kernel operating line + `! SOURCE-READS` footer advisory + gate_read.py gate 3 (deny-once with the toolcard hint; unregistered like the rest of gate_read) |
| DRAFT_FACT_ERRORS | Principal challenges a specific in a drafted email: "seems halucinated", "check the accuracy", "how do we know this is true", "are you sure" | `rg -ci -e hallucinat -e halucinat -e "check the accuracy" -e "are you sure" -e "this correct" -e "need to confirm"` | any = problem; highest-value category |
| DRAFT_MARKUP_ROUNDS | Principal pastes the draft back with `[bracketed]` or `(1)(2)(3)` corrections | `rg -c '\[[^\]]{10,}\]'` and `rg -c '\(1\)'` on user msgs | >1 round per email = problem |
| AI_ISM_ESCAPES | "feels very AI", "sounds AIish" — despite the AGENTS.md ban list | `rg -ci -e "very AI" -e AIish -e "AI type"` | any = ENFORCEMENT gap |
| VOICE_VOCAB | Words he wouldn't use or the recipient wouldn't know ("pelike... wtf is it") | manual review of vocabulary complaints | any = problem |
| INCOMPLETE_RETRIEVAL | Re-pushes: "missed the first time", "did any of them get answered", "much more extensive", repeated "deep scan" | `rg -ci -e missed -e "be comprehensive" -e "deep scan" -e "read all the"` | >1 = problem |
| FALSE_ATTRIBUTION | Agent claims the principal set/added/owns something the agent created ("I HAVE NO STANDING RULES") | manual review of "you/your" claims vs git provenance | any = problem |
| TOPIC_CONFUSION | Wrong topic/path resolved; wandering into Sandbox or unrelated dirs | user messages naming paths back ("no its Sam C:\Topics\Sam") | any = problem |
| PRINCIPAL_MISREAD | Advice/drafts against his stated stances (dealer framing, volunteering concessions, repeating advice he already rejected) | manual review — compare to AGENTS.md §The principal | any = problem |
| PERMISSION_STALLING | "Should I continue?" / "Want me to…?" mid-task | `rg -ci -e "should I continue" -e "shall I" -e "want me to"` on assistant msgs | >0 = problem |
| TIME_CONSTRAINT_INVENTION | Agent fabricates a clock to cut scope: "given time constraints", "time budget", "given the time pressure" in assistant/thinking text justifying a spot-check, an early close, or a downgraded fix | `rg -ci -e "time constraint" -e "time budget" -e "interest of time"` — then classify: domain content (counterparty pressure, `--virtual-time-budget`, API credits) is noise | any genuine = problem (his ruling Aug 22, 2026: "there are never time constraints") |
| IGNORED_RULES | Behavior AGENTS.md already forbids | cross-check every finding against AGENTS.md | any = enforcement gap, not doc gap |

Detection commands use repeated `-e` flags, never `\|` alternation: in ripgrep's regex
`\|` is a LITERAL pipe character and the pattern silently matches nothing (verified
Aug 14, 2026 — the old escaped-pipe recipes returned 0 on known-hit text).

**Rough token cost per instance:** draft markup round trip ~10–15K (full re-read +
rewrite + re-sweep), fact-challenge round ~5–10K, re-pushed retrieval ~10K+, false
attribution ~5K plus trust damage. User time costs more than all of it.

**Harvest principal facts (the "understand me" pass).** Separately from friction
counting, collect every correction that reveals a DURABLE fact about the principal —
voice ("Best regards, Jonathan"), vocabulary limits, negotiation stances (collector not
dealer, never volunteer concessions), workflow preferences (inline code-block emails,
not .md files). One-off task corrections ("use $75k not $90k") stay in the topic record;
stable preferences belong in AGENTS.md §The principal. The test: would this correction
apply verbatim in a different topic next month? Yes → candidate rule.

## Step 3: Audit the Programming Infrastructure

The infrastructure = `_system/tools/` (helpers + the check suite + `run_suite.py` +
`install_workspace.py`), the canonical `.cursor/` (skills, commands, rules, hooks),
and topic-local generators (e.g. coyotes `tools/atlas/`). Two passes, both every run.

**A. Transcript-driven** — what actually broke or ground in the sessions scanned:

| Category | What it looks like | Detection idea | Fix shape |
|----------|--------------------|----------------|-----------|
| TOOL_FAILURES | A `_system/tools/*.py` run crashes or misleads: Traceback, UnicodeEncodeError, arg errors, empty output taken at face value | `rg -c -e Traceback -e UnicodeEncodeError -e "usage: "` on assistant-side text near tool names | fix the tool; extend its `--self-test` to cover the bug FIRST (TDD) |
| MANUAL_CHAINS | Long rg/ls sequences doing work an existing tool does (`lookup_id`, `find_book`, `query_wiki` bypassed — the Aug 13 019N-40 NAS sweeps), or the same multi-command sequence hand-built across sessions | count shell calls per question; compare against the tool list | tool exists → discoverability rule in AGENTS.md; none exists → new helper |
| DELEGATION_MISS | Frontier chats grinding mechanical work the MODELS.md routing table benches to cheap seats: scan/photo Reads into context, hand-authored bulk apply sidecars, benched-cheap command classes billing on Fable/Opus | `delegation_census.py` — the Step 3C required call; every flagged row (WRONG-SEAT / GRIND-READS / BULK-APPLY) is a finding candidate | any flagged row = problem — HIS TOP EXPENSE ("the expensive models not outsourcing to the cheaper models is my biggest expense", Aug 30, 2026). Fix = routing row + forcing mechanism (advisory → ledger field → gate), never prose alone |
| SUITE_GAPS | An error class caught by eye (his or a later session's) that a mechanical invariant could catch — the atlas scope/grade/refs checks were all born this way | manual review of each correction: is the violated invariant checkable? | new `check_*.py` wired into the `CHECKS` list in `run_suite.py` (or a topic-local suite) |
| HOOK_GAPS | Forbidden shell patterns actually executed (`\| head`, `python3`, inline Python, `sleep`) | `rg -c` the patterns in assistant-side shell text | strengthen `gate_shell.py` (canonical `.cursor/hooks/`) — registration is currently gated, see the AGENTS.md forbidden table; the pattern list still improves now |
| SKILL_COMMAND_DRIFT | A skill/command step sessions repeatedly deviate from, or that references paths/flags/tools that no longer exist | compare each skill's commands against reality (paths exist? flags exist?) | edit the canonical skill/command |

**B. Static sweep** — cheap convention checks, run as-is from `/c/Topics`:

```bash
wc -l _system/tools/*.py                                        # >1,000 lines = split candidate (soft threshold, Aug 15, 2026)
rg --files-without-match -e "--self-test" _system/tools/*.py    # missing self-test
rg --files-without-match "reconfigure" _system/tools/*.py       # missing UTF-8 stdout guard
ls _system/tools/check_*.py && rg -n -A 10 "CHECKS = " _system/tools/run_suite.py   # on-disk checks vs suite registry
diff _system/workspace/AGENTS.md AGENTS.md; diff -rq _system/workspace/.cursor .cursor   # canonical vs installed drift
```

Static hits are CANDIDATES, not verdicts — verify before filing. Known false-positive
shapes (all verified Aug 16, 2026): self-tests living in a sibling or entry-point tool
(`find_book_selftest.py`; `ingest_email.py --self-test` covers
`ingest_email_formats.py`); the UTF-8 guard inherited transitively (`research_archive`
reconfigures stdout at import, covering `deep_research.py`/`openrouter_research.py`);
`research_loop.sh` is not Python. Library modules with no `__main__` need neither
convention directly. Also treat a `diff -rq` root-vs-canonical hit showing a ROOT-ONLY
file as a real finding: root `.cursor` is not a git repo, so a skill created there is
unversioned — port it into `_system/workspace/` (the make-pdf skill sat unversioned
until the Aug 16 run). A convention miss that never caused observed friction ranks
below any transcript-evidenced finding.

**C. Delegation census — REQUIRED every run, ONE call, results quoted in the
findings:**

```bash
python C:/Topics/_system/tools/delegation_census.py --since-last-improve
```

Dollar-ranked delegation-compliance receipts for the window: WRONG-SEAT chats
(benched-cheap command class billing on frontier), GRIND-READS chats (3+ scan
Reads, zero `cheap_read.py` launches), BULK-APPLY html_edit rows (>= 10 blocks,
conv-attributed), the frontier-vs-cheap billing split with the flagged-$ share,
and his money complaints quoted (`MONEY_RE_SRC`). Each flagged row is a finding
candidate; the flags map to the routing table in `_system/MODELS.md` (binding
decisions + receipts in ONE pack since the 2026-08-31 kernel cutover — the
kernel body carries no routing prose). Mechanism: the per-session advisories
(`! READS`, `! ROUTING`, `! ROUTE`) fire one chat at a time and vanish with the
chat — without the window aggregate, delegation drift is invisible exactly where
this skill looks. Post-cutover this call is also the cutover brief's standing
watch item: with delegation prose living only in a pointer-reached pack, the
census aggregate + footer advisories are the measure of whether delegation
survives the swap. Incident: ~20 distinct save-me-money asks over 5 days (157
raw `--money` hits across 62 chats) while frontier kept grinding mechanical
work — none of it became a finding until he escalated a 15th time
(Aug 30, 2026). "The steer didn't mention cost this run" = NO — the call runs
every pass.

**D. Bench-unit accumulation — check BOTH tallies every run:**
`_system/archive/bench_runs/KERNEL_TALLY.md` (the kernel program scoreboard;
his standing order, Aug 31, 2026 — paired A/B units 1–5 pre-cutover; since
the 2026-08-31 cutover, units are SINGLE-ARM verification replays: the LIVE
kernel bytes against a sealed answer key, one cheap seat — unit 7 kern7art
is the pattern) and
`_system/archive/bench_runs/COMPLY_TALLY.md` (competing-prompt compliance
units — sup/neu/adv variants per rule, `bench_comply.py`; ordered Aug 31,
2026, conv 0c362804 continuation). If NO unit landed on either tally since
the last improve pass, stage + launch + grade ONE unit (cheap grok CLI
seats; grading stays frontier) — prefer the comply unit whose rule this
run's Step 2/3 findings implicate, else a kernel verification unit on the
thinnest-covered fixture class. A comply verdict routes mechanically:
REMINDER-DEPENDENT/BROKEN = enforcement gap (Step 5 ladder — mechanism,
never re-worded prose); HOLDS through adv = pruning evidence; a failing
`NEW:*` baseline = the new rule line lands WITH its evidence. A kernel
verification unit showing a kernel-attributable drop = fix the kernel
SOURCE (`KERNEL_HEADER.md` operating prose, an `enforcement_map.py` row, or
the pack the pointer names), re-assemble, re-run the class; if a drop class
resists a kernel fix, the revert source is
`_system/archive/AGENTS_FULL_2026-08-31.md` — and reverting is HIS decision,
never a bench session's. Pruning a kernel row likewise goes to him as a
decide item, only on extinct-behavior evidence.

## Step 4: Quantify (REQUIRED FORMAT)

For each pattern above threshold, output exactly this block — no prose findings:

```
PATTERN: <one-line description>
COUNT: <occurrences> (across <N> sessions)
TOKENS_WASTED: <count × per-instance estimate>
ALREADY_DOCUMENTED: <yes/no — checked in Step 5>
FIX_TYPE: <rule | tool | check | hook | skill/command>
FIX_LOCATION: <from routing table in Step 6>
```

Rank by TOKENS_WASTED × recurrence, with user-time weight: a category that makes the
principal re-type his own reasoning outranks a pure token burner. Fix the top 1–3 only;
log the rest in the commit message body for the next run. Ten mediocre rules are worse
than one sharp rule — and one working check outranks both.

**Floor-first ranking (composes with the rank above, never replaces it):** at
comparable TOKENS_WASTED, a floor-class fix — a cause of the worst sessions:
mechanical enforcement for a still-violated rule, a cause of the dead-on-arrival
cheap runs, a built-but-unregistered gate — outranks a ceiling-class one (enriching
guidance compliant sessions already follow; that class ranks LAST). A floor
classification must NAME its receipt (the still-violated rule with its measured
hits, the dead-on-arrival cause, the named unregistered gate) — no receipt, no
floor rank; an undefined tie-breaker would let a run label its preferred work
floor-raising. This ranking guidance is itself prose — the weakest ladder rung —
with the parked blind-regression fixture (Step 6) as its designated compensating
control. Evidence: gate_shell/gate_read sat BUILT, self-tested, and UNREGISTERED
while their gated patterns kept executing (~430 head/tail-pipe hits Aug 20–23 plus
~290 in one 90-min Aug 28 window — gate_shell's patterns; 191 redundant re-reads in
a 14-session scan — gate_read's; AGENTS.md forbidden table) and ~25% of cheap runs
arrive dead on arrival (MODELS.md process tax); secondary support, the study behind
the steer: Anthropic's "Agentic coding and persistent returns to expertise"
(June 2026, https://www.anthropic.com/research/claude-code-expertise, verified
Aug 28, 2026) finds most of the gain concentrated at the lower end of the expertise
scale — "the gains come mostly from competence, not mastery" (human expertise
tiers; the session-class analogy is this workspace's adopted steer).

## Step 5: Deduplicate (BLOCKING)

Before adding ANY new guidance:

```bash
rg -i "<pattern keyword>" _system/workspace/AGENTS.md _system/workspace/.cursor/ _system/tools/enforcement_map.py _system/INGEST.md _system/ASK.md _system/DRAFTING.md _system/ADVICE.md _system/ARTIFACTS.md _system/RESEARCH.md _system/SYSTEM_PATTERN.md _system/MODELS.md _system/FORBIDDEN.md _system/PRINCIPAL.md _system/HOOKS.md _system/MAIL.md
```

Before writing ANY new tool or check: `ls _system/tools/` and read the nearest
candidate's docstring — the gap is often discoverability, not absence (a tool that
exists but wasn't reached for needs a pointer rule, not a twin).

- **Already documented** → do NOT duplicate. This is an enforcement gap — escalate up
  the ladder: prose rule → required visible output step → check/hook. A rule violated
  again after it was written never gets the same exhortation re-worded; it gets a
  mechanism. Incident on file: three AI-ism lines escaped drafts the same day the ban
  list was committed — the fix that held was `check_aiisms.py`, not stronger wording.
- **Not documented** → new entry at the lowest ladder rung that plausibly holds.

## Step 6: Route and Write the Fix

**Canonical copies ONLY — never edit root `AGENTS.md` or root `.cursor/`** (they are
installer output; `install_workspace.py` refuses on conflict and your edit dies there).
Since the Aug 31, 2026 kernel cutover the canonical `workspace/AGENTS.md` is ITSELF
assembled output — `KERNEL_HEADER.md` + generated `KERNEL.md`, built by
`python _system/tools/build_kernel.py --assemble` — so never hand-edit it either:
route rulebook edits per the table below, then re-assemble + re-install. The full
pre-cutover rulebook is archived at `_system/archive/AGENTS_FULL_2026-08-31.md`.

| Fix type | Location |
|----------|----------|
| Cross-cutting behavior, session-rulebook prose | `_system/KERNEL_HEADER.md` (operating essentials) or the pack owning the detail (`_system/PRINCIPAL.md`, `_system/DRAFTING.md`, `_system/FORBIDDEN.md`, `_system/MODELS.md`); then `python _system/tools/build_kernel.py --assemble` |
| New/changed Hard-safety or FORBIDDEN row | `_system/tools/enforcement_map.py` (id + needle + shapes; the row TEXT is the needle) → `build_kernel.py --assemble`; `check_enforcement_coverage` reds a stale KERNEL.md |
| Scoped rules / new skills / new commands | `_system/workspace/.cursor/rules\|skills\|commands/` |
| Ingest / ask / drafting / advice / artifact / research protocol behavior | `_system/INGEST.md`, `_system/ASK.md`, `_system/DRAFTING.md`, `_system/ADVICE.md`, `_system/ARTIFACTS.md`, `_system/RESEARCH.md` (read the target file whole first) |
| Structural conventions, environment log | `_system/SYSTEM_PATTERN.md` |
| Model-routing decisions, bench evidence, delegation mechanics | `_system/MODELS.md` (the binding routing table + receipts live here since the cutover; the kernel only points at the pack) |
| Tool bug or missing flag | the tool itself in `_system/tools/` — extend its `--self-test` to reproduce the bug first, then fix |
| Missing helper tooling | `_system/tools/<name>.py` (≤1,000 lines soft, ships with `--self-test` + UTF-8 stdout guard) + one pointer line in AGENTS.md |
| New mechanical invariant | `_system/tools/check_<name>.py` + add to the `CHECKS` list in `run_suite.py`; topic-scoped invariants go in that topic's local suite (the coyotes `tools/atlas/check_*.py` pattern) |
| Shell-pattern gate | `_system/workspace/.cursor/hooks/gate_shell.py` (canonical); do NOT re-register `hooks.json` until the CLI gate in the AGENTS.md forbidden table passes |

**Writing rules that actually work** (apply to every edit):

- **Format: statement + mechanism + at most ONE incident line.** The mechanism (WHY it
  fails) is what prevents rationalization; a rule without a mechanism gets argued away.
  - Good: `Every specific in a draft comes from a grep of the record, not memory — 11 accuracy challenges in the Jul 27–28 sessions came from unchecked specifics.`
  - Bad: `Try to be more accurate in drafts.`
- **Authority language** for discipline rules: NEVER / ALWAYS / MUST, with the ✅
  INSTEAD alternative next to the ⛔ NEVER.
- **Close the observed loophole:** if a transcript shows the agent rationalizing, quote
  that exact rationalization with `= NO.` Only when evidenced.
- **Prefer forcing mechanisms over exhortations** for rules that keep being ignored: a
  named visible step ("state `AI-ism sweep: clean` above the draft") is checkable;
  "be careful about AI-isms" is not; a suite check is stronger than both.
- **Keep the assembled rulebook lean — the byte budget moved to the SOURCES at
  the 2026-08-31 kernel cutover.** The canonical `workspace/AGENTS.md` is now
  assembled output (~13 KB: `KERNEL_HEADER.md` operating essentials +
  generated `KERNEL.md` rows), and every byte is still always-injected context
  in every future session — so hold the sources lean: header additions stay
  one-line operating essentials; enforcement-map rows stay statement + shapes
  (incident narrative goes to `_system/FORBIDDEN.md`); anything longer lives
  in a pack the header points at (`PRINCIPAL.md`, `MODELS.md`, `HOOKS.md`,
  the protocol packs) or a scoped rule file. Packs load on-demand by trigger;
  the evidence that pointer-loading holds is bench units 1–5 plus the comply
  `pipelong` long-context datum (KERNEL_TALLY.md / COMPLY_TALLY.md), and the
  standing post-cutover monitors are the verification units,
  `delegation_census.py`, and the footer advisories. The pre-cutover
  lean-down program (126 KB → 45 KB in three tranches, then a deferred
  "stage 2" vendor-band cut) is RETIRED — superseded by the cutover; its
  narrative lives in the archived rulebook
  (`_system/archive/AGENTS_FULL_2026-08-31.md`) and the packs it filled.

**Writing tools/checks that actually work:** ≤1,000 lines (soft), `--self-test` written first
(TDD, per the workspace user rules), `sys.stdout.reconfigure(encoding="utf-8",
errors="replace")` at the top, stdlib `re`, config via the topic's `_config.json` —
never a hardcoded topic path. A check must FAIL on the incident that motivated it:
reproduce the historical failure in the self-test before trusting the green.

**Verifying BEHAVIORAL fixes (rules/protocol edits with no self-test): blind
regression session.** The TDD equivalent for prompt-space fixes — hand the
principal a ready-to-paste prompt built from the REAL incident input (the draft
he marked up, the question that went wrong), let a FRESH session run it blind
(it must not know it is being tested), then grade the delivered output against
an answer key built from his own corrections. Park the fixture (input + verbatim
answer key + per-run results) in `_system/archive/<name>_regression_<date>/` so
the next rubric/protocol change re-runs it instead of re-deriving it. Grade
honestly WHICH layer caught each item — the record (his corrections registered
by a prior session) vs the new rules/gates on fresh text — and say so; a pass
that mostly proves record-inheritance is still a pass, but the claim must match
(first use: the email deepen gate, `archive/email_gate_regression_2026-08-23/` —
run 1 caught ~1/8 classes, post-fix run 8/8 with zero principal edits).

## Step 7: Install, Verify, Commit

```bash
python C:/Topics/_system/tools/<touched>.py --self-test        # every tool/check touched
python C:/Topics/_system/tools/build_kernel.py --assemble      # if KERNEL_HEADER.md or enforcement_map.py changed — regenerates KERNEL.md + the canonical AGENTS.md
cd /c/Topics/<any-topic> && python ../_system/tools/run_suite.py   # if a check or suite-adjacent tool changed
python C:/Topics/_system/tools/install_workspace.py            # canonical → root
rg "<new pattern keyword>" C:/Topics/AGENTS.md C:/Topics/.cursor/ C:/Topics/_system/<pack>.md   # installed + findable at the fix's actual home
```

The installer's guard compares content, not provenance: after you edit a canonical file
the root copy (old version) now "differs" and the installer refuses, printing the diff.
Confirm the printed diff is EXACTLY your canonical edit — nothing else would be lost —
then re-run with `--force`. If the diff contains anything you didn't write, STOP: the
root copy was hand-edited and must be ported into the canonical copy first.

Re-read edited sections to confirm no truncation. Commit in the `_system` repo
(message style: `AGENTS.md: <summary>` — see `git log`), unfixed findings listed in the
body. LOCAL-ONLY: `git remote -v` must be empty; NEVER push.

## Few-Shot Examples (real, from this workspace's history)

**Good finding (guidance fix):**

```
PATTERN: Drafted emails contain unverified specifics; principal fact-checks line by line
COUNT: 11 challenges + 14 markup round-trips (across 5 sessions)
TOKENS_WASTED: 25 × ~10K = 250K, plus principal trust in every other line
ALREADY_DOCUMENTED: no
FIX_TYPE: rule
FIX_LOCATION: _system/workspace/AGENTS.md — facts-are-lookups rule in email section
```

**Good finding (infrastructure fix):**

```
PATTERN: ID lookups done as manual grep chains across row files; sweeps skipped under fatigue → false absence claims
COUNT: 240+ grep round-trips per marathon session (2 sessions)
TOKENS_WASTED: ~500K, plus false absence claims reaching the principal
ALREADY_DOCUMENTED: yes (THE ABSENCE RULE) — enforcement gap, rule alone did not hold
FIX_TYPE: tool
FIX_LOCATION: _system/tools/lookup_id.py (--absence sweeps all row files mechanically) + AGENTS.md pointer line
```

**Bad finding (reject this):**

```
PATTERN: Agent makes mistakes sometimes
FIX: Improve guidance
```

**Good rule produced from a finding:**

```markdown
Never attribute an artifact, rule, or file to the principal without checking provenance
(`git log` on the file). Twice an agent presented its own creations as his ("I HAVE NO
STANDING RULES. AI (like you add them. I have not added any)") — misattribution makes
him doubt the whole record.
```
