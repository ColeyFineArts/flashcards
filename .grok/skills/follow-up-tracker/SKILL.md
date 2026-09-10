---
name: follow-up-tracker
description: >
  Extract commitments from email, meetings, and notes; maintain a chase list of who
  owes what by when; nudge overdue items and produce a daily loose-ends digest. Use
  when the user asks about follow-ups, open loops, action items, chasing someone, or
  runs /follow-up-tracker.
when-to-use: >
  follow up, chase, open loops, action items, commitments, what am I owed, what did I
  promise, loose ends, outstanding asks, nudge
argument-hint: "[optional: digest | add | nudge <person> | from meeting/email]"
metadata:
  author: executive-assistant
  short-description: Commitment chase list + daily loose-ends digest
---

# Follow-Up Tracker

Close loops. Track who owes what by when — including promises the executive made — and surface what needs a nudge.

## Modes

| Arg / ask | Behavior |
| --- | --- |
| (default) / `digest` | Produce today's loose-ends digest |
| `add` / paste / "from this email" | Extract commitments and add to the list |
| `nudge <person>` | Draft a polite chase for overdue items involving that person |
| meeting name / thread | Mine that source only, then update the list |

## Data store

Persist commitments in `./.grok/ea/commitments.md` (create the folder/file if missing). If the path is not writable, keep an in-session list and remind the user to save it.

Use this structure:

```markdown
# Commitments

| ID | Status | Due | Owner | Owed to | Commitment | Source | Last nudge |
| --- | --- | --- | --- | --- | --- | --- | --- |
| C001 | open | 2026-09-12 | Dana Lee | Exec | Send revised SOW | Email 2026-09-08 "SOW v3" | — |
```

Statuses: `open` · `nudged` · `done` · `dropped`

## Extraction rules

When mining email, notes, or meeting transcripts, capture:

- Explicit promises ("I'll send…", "Can you have this by Friday?")
- Soft commitments with implied owners ("we'll circle back on pricing")
- Decisions that imply a next step

Skip: vague brainstorming, already-completed items, calendar holds with no ask.

For each item set: **Owner**, **Owed to**, **Due** (infer from language; else `TBD` + suggest a date), **Source** (link or precise cite), **Commitment** (one line, verb-first).

## Digest template

```markdown
# Loose ends — <date>

## Overdue
- **C00x** <commitment> — **<Owner>** owed <Owed to> since <due> · Source: …
  - Suggested nudge: <1-line angle>

## Due soon (≤3 days)
- …

## Waiting on others
- …

## You owe
- … (executive's open promises — protect credibility)

## Suggested nudges (drafts)
### To: <Name> re: <topic>
<short chase email/Slack in the exec's voice>
```

## Nudge style

- Assume positive intent; reference the original ask and date.
- Offer an easy out ("still on track for Thursday, or want a new date?").
- One ask per nudge. No guilt, no stacked threads of unrelated items.
- Do not send nudges unless the user explicitly confirms.

## Rules

- Separate **You owe** from **Waiting on others** — both matter; different urgency tone.
- Deduplicate: same owner + same deliverable = one row (update source/due if needed).
- When marking done, keep the row with `done` and a completion note — do not delete history casually.
- If due date is TBD, propose one and ask to confirm rather than leaving it forever open.
- Never invent commitments that were not implied by the source text.
