---
name: inbox-triage
description: >
  Triage the executive's inbox into Act / FYI / Delegate / Archive, surface only what
  needs attention, and draft voice-matched replies. Use when the user asks to check
  email, clear the inbox, prioritize messages, draft a reply, or runs /inbox-triage.
when-to-use: >
  inbox, triage email, check my email, prioritize inbox, draft reply, what needs my
  attention, clear my inbox, catch up on email
argument-hint: "[optional: unread | today | from:name | subject keywords]"
metadata:
  author: executive-assistant
  short-description: Inbox triage + voice-matched reply drafts
---

# Inbox Triage

Protect the executive's attention. Surface signal, suppress noise, and draft replies that sound like them.

## When invoked

1. Scope the batch:
   - Default: unread + flagged + anything needing reply in the last 48 hours (cap ~25)
   - Honor filters if given (`from:`, `today`, subject keywords)
2. For each message, classify into exactly one bucket:
   - **Act** — needs the executive's decision, reply, or signature
   - **FYI** — important awareness; no action required
   - **Delegate** — someone else should own it (suggest who)
   - **Archive** — low value / auto-generated / already handled
3. Rank **Act** items by urgency × impact (deadlines, VIP senders, money/legal/people risk first).
4. For top Act items (default top 5), draft a reply in the executive's voice.
5. Present the triage board, then drafts. Do not send anything unless explicitly told to.

## Voice matching

Build a short style card from recent *sent* mail when available:

- Greeting/sign-off habits
- Typical length (terse vs. explanatory)
- Formality, contractions, how they say no / push back / ask for time
- Signature lines to preserve

If no sent history: use crisp, warm-professional prose; ask once if they want a different tone.

## Output template

```markdown
# Inbox triage — <time window>
**Reviewed:** N · **Act:** n · **FYI:** n · **Delegate:** n · **Archive:** n

## Act (do these)
1. **<Subject>** — <From> · <age>
   - Why it matters: <1 line>
   - Suggested move: Reply | Schedule | Decide | Forward
   - Draft: ready below / skip if trivial

## FYI
- **<Subject>** — <From>: <1-line takeaway>

## Delegate
- **<Subject>** → **<owner>**: <1-line brief to forward>

## Archive / skip
- <count> newsletters, receipts, CC threads (list only if asked)

---

## Drafts

### Reply: <Subject>
<draft body in their voice — ready to paste/send>
```

## Draft quality bar

- Lead with the answer or decision; context second.
- Mirror the sender's ask; address every question.
- Include a clear next step and owner when relevant.
- Offer 2 tone variants only when the stakes are high (e.g. decline a VIP, negotiate).
- Flag uncertainty: "Confirm before send — I inferred X from Y."

## Rules

- Never mark read, archive, label, or send without explicit confirmation.
- Prefer fewer, sharper Act items over dumping the whole inbox.
- VIP / legal / HR / security threads: always Act or FYI — never silent Archive.
- If tools are unavailable, triage pasted threads or exports the same way.
- After presenting drafts, ask once: "Send any of these, or revise?"
