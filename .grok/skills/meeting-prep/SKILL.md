---
name: meeting-prep
description: >
  Build a concise pre-meeting brief from calendar, email, and docs. Use when the user
  asks to prep for a meeting, brief them before a call, run morning meeting prep,
  summarize attendees, or runs /meeting-prep.
when-to-use: >
  meeting prep, brief me, prepare for my next meeting, morning briefing, attendee
  research, what's this meeting about, prep notes
argument-hint: "[meeting name or time, e.g. next, 2pm, Board sync]"
metadata:
  author: executive-assistant
  short-description: Pre-meeting briefs from calendar + email + docs
---

# Meeting Prep

Produce a one-page brief so the executive walks into the meeting ready. Prefer facts from tools over guesses.

## When invoked

1. Resolve the target meeting(s):
   - Explicit name/time → that meeting
   - "next" / no arg → next upcoming meeting
   - "today" / "morning" → all meetings in the next working day (cap at 5; ask which to expand)
2. Pull calendar details: title, time, location/link, attendees, description, attachments.
3. For each external or non-obvious attendee, gather lightweight context from email/Drive (role, last interaction, open threads). Skip deep research unless asked.
4. Scan recent email (≈14 days) and relevant docs for: decisions pending, commitments, prior meeting notes, attached decks.
5. Draft the brief using the template below. Keep it scannable — bullets over paragraphs.

## Output template

```markdown
# Meeting brief: <Title>
**When:** <day, time, duration> · **Where:** <place or link>
**Goal (inferred):** <1 line — flag if unclear>

## Attendees
- **Name** (org/role if known) — last touch: <date + 1-line context>

## Context
- <2–5 bullets: why this meeting exists, recent thread highlights, relevant docs>

## Open items
- <decisions, asks, blockers still unresolved>

## Suggested talking points
1. ...
2. ...
3. ...

## Watch-outs
- <politics, sensitive topics, missing info, time risks — omit section if none>

## Prep actions (optional)
- [ ] <anything to do before joining>
```

## Rules

- Be specific and attributable ("Email from Dana 3/12: …"). Do not invent attendee bios or deal status.
- If calendar/email tools are unavailable, say what you need and produce a partial brief from what the user provides.
- Default length: ~½–1 page. Expand only if the user asks for deep prep.
- For multi-meeting morning runs: one short card per meeting, then offer a deep brief on any one.
- Never send email or change calendar events unless the user explicitly asks.
- End with one question only if a critical fact is missing (e.g. meeting goal unknown).
