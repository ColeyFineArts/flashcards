---
name: 2025 Tax Turbo
overview: Produce a CPA-ready 2025 tax packet matching prior years without re-keying every transaction into Excel — hybrid Monarch (operating P&L) + CapEx/Sale registers + one summary workbook. Diagnosed 2026-09-17.
todos:
  - id: decide-source-of-truth
    content: "DECIDED hybrid — Monarch for recurring operating P&L (LTR/STR/LLC/W-2 feeds); CapEx + Sale as Excel/Sheets event registers; ONE year-end summary workbook. Pure Monarch-only ELIMINATED."
    status: completed
  - id: declutter-drive
    content: "Trash corrupt Turbo stubs (8B id 1IMNMy…, 1526B id 1UZSbV…) + superseded CapEx sheet 18KFS0…; keep CapEx register 1ojNHX… + usable Turbo 1SaLWL… (~27KB) OR Personal Income copy 1ffBOP…; trash extra CapEx clones"
    status: pending
  - id: monarch-connect
    content: Connect/confirm Mercury, BoA, STR payout, kid/joint accounts in Monarch; create tax-bucket tags for OPERATING only (not CapEx double-entry)
    status: pending
  - id: diagnose-monarch
    content: "/diagnose DONE 2026-09-17 — pure Monarch-only CONTRADICTED; hybrid SUPPORTED (see Diagnostic Session Summary)"
    status: completed
  - id: oak-park-2025
    content: Finish 216 Oak Park LTR 2025 (smoke detector TBD); use Rocket 1098 already in Documents/216 folder for interest not full P+I
    status: pending
  - id: grove-capex-split
    content: "TY split DONE in workbook ($53,660 / 2025 = 2 rows only; $268,532 / 2026). Still fill UNKNOWN $9,401.93 dates + 3 missing amounts; resolve $3031.41 dup"
    status: pending
  - id: sale-docs-blocker
    content: "BLOCKER — drop 524 sale Closing Disclosure + prior depreciation schedules into empty 524 Ferdinand Sale folder (Drive reconfirm 2026-09-17: 0 children; no CD elsewhere)"
    status: pending
  - id: grove-purchase-cd
    content: "Also need 827 N Grove purchase Closing Disclosure for initial basis (CapEx action item R154)"
    status: pending
  - id: cpa-packet
    content: Assemble forms + checklist; CPA handles §121 allocation / dep. recapture / STR Sch E vs C
    status: pending
  - id: optional-monarch-mcp
    content: Connect Monarch MCP read-only (https://api.monarch.com/mcp) after declutter — host reachable (HTTP 405); Cursor MCP not auth'd yet
    status: pending
isProject: false
---

# 2025 Tax Turbo Plan

**Goal:** Same CPA-ready year-end packet as prior years, without re-keying every transaction into Excel.

**Disclaimer:** Organizational plan only — not tax, legal, or accounting advice.

## Concrete goals

1. ~~One daily source of truth for transactions (Monarch)~~ → **DECIDED hybrid** (diagnose 2026-09-17): Monarch for recurring operating P&L; CapEx + Sale as event registers.
2. One year-end summary shaped like `Personal Income.xlsx` tabs (not three competing workbooks).
3. Dedicated trackers for (a) 524 Ferdinand sale ~Dec 15, 2025 and (b) 827 N Grove CapEx with **calendar-year split** — TY2025 portion already measured: **$53,660 (2 payments on 2025-12-19)**.
4. Drive folder `2025 Taxes` holds PDFs + checklist for CPA handoff.

## Operating model (DIAGNOSED — hybrid locked)

| Layer | Tool | Job |
|-------|------|-----|
| Capture | Mercury + BoA + cards + Airbnb/VRBO | Feeds |
| Classify (operating) | Monarch Money | Tax-bucket tags for LTR / STR / LLC / income |
| Event registers | ONE Grove CapEx sheet + ONE Home Sale tracker | Basis / sale facts — **not** Monarch P&L |
| Summarize | ONE Personal Income / Tax Turbo workbook | Prior-year tab shape |
| File | CPA | Packet |

**Rule:** Do not maintain three live ledgers. CapEx hand-lists are the **register** (required — multi-account Laura reimbursements). ~~Pure Monarch-only for CapEx~~ **ELIMINATED** — diagnose 2026-09-17.

**Canonical Drive IDs (keep):**
- CapEx register: `1ojNHX_z94FGMXBFKGcZ7Eta_iyQCrBvNjKbitKUJT4Q`
- Usable Turbo xlsx: `1SaLWLzyPSxatdyfO28pwdkKv0lGbkPjw` (27104 B) **or** Personal Income copy `1ffBOPDNo8vwt2LSKZJy__VG37nmIw9sy` (132313 B)
- Sale folder (empty): `1K6_C7BcRF1PuC1l_MGjYWHBOGS2I9w9L`

**Trash candidates:** Turbo stubs `1IMNMy5WtmbbHgG2mtGmqekglui-JIcR0` (8 B), `1UZSbVckykzcZWo2KtEc6yDWvutUk4imY` (1526 B); superseded CapEx `18KFS06ib6MhC2mmg3sqBqRKBPUlLQrkPu3aceRWimNY`.

## Key properties / entities

- W-2 Hindman; Megan income; GCM; art; EPGC LLC (Mercury)
- LTR: 216 N. Oak Park — Documents/216 already has Rocket **1098**
- STR / two-flat: 524 Ferdinand Unit 2 through sale; Unit 1 owner allocation
- New primary: 827 N Grove — Documents/827 already has USB Mortgage **1098**; Water `0245000849-01`; ComEd `7103042419`

## Already captured

- Drive `2025 Taxes` + property subfolders under Documents
- Oak Park 2025 known cash ~$21,423 (smoke detector TBD)
- Grove CapEx running ~$331,594 (excl. duplicate-suspect $3,031.41)
  - **TY2025 included:** $53,660 (16.2%) — JCDA $41,910 + Riverton $11,750 on 2025-12-19 only
  - **TY2026 included:** $268,532 (81.0%)
  - **UNKNOWN / undated:** $9,401.93 (2.8%) — Wayfair $1,189.08 + Laura supplies $7,512.85 + Village permit ~$700
- Official Monarch MCP: `https://api.monarch.com/mcp` (host reachable; Cursor MCP not connected)

---

## Deepening Review Results (2026-09-17)

### New information discovered

1. **`524 Ferdinand Sale` folder has no sale Closing Disclosure** — search found only unrelated art consignment settlements. Sale packet is still the critical gap.
2. **Drive artifact sprawl is worse than stated:** multiple `827 N Grove CapEx Payment Register` copies; multiple `Personal Income 2025 Tax Turbo.xlsx` including **corrupt/tiny files** (8 bytes, 1526 bytes) alongside a ~27KB copy — risk of editing the wrong file.
3. User already filed **2025 1098s** under `Documents/216 N OAK PARK AVE` and `Documents/827 N GROVE AVE` — plan should point there, not ask to re-collect blindly.
4. **§121 duplex allocation is not DIY-optional:** IRS Pub 523 / Treas. Reg. 1.121-1 require allocating gain between residential and separate rental portions; depreciation after May 6, 1997 is never excluded. Confirms CPA handoff for sale math.
5. Corporate return `TaxReturn2023Corporate.pdf` exists on Drive — EPGC entity type must be confirmed against that history (Sch C vs entity return), not assumed.

### Assumptions validated

1. Prior `Personal Income.xlsx` tab shape is the right CPA interface — still present and used historically.
2. Mixed personal/rental sale needs allocation + dep. recapture attention — **web-verified** via IRS Pub 523 (2025) and 26 CFR 1.121-1.
3. Monarch has an official MCP suitable for Cursor — **verified** Monarch help article; not yet connected in this environment.

### Assumptions invalidated

1. ~~"Packet structure is ready; just fill data"~~ → Packet has **duplicate/corrupt Turbo workbooks** and an **empty sale folder**. Declutter is a prerequisite todo.
2. ~~"Excel monthly SUMIFS from Monarch is the speedup"~~ → User already proved they'll paste CapEx/LTR lists into chat/sheets; without a hard Monarch habit, Excel becomes the ledger again. **Source-of-truth decision is blocking.** → **RESOLVED hybrid** (diagnose).
3. ~~"Grove CapEx is a 2025 tax turbo item as one blob"~~ → Most spend is **2026**; TY2025 packet should only include Dec 2025 (and any other 2025) CapEx + purchase basis docs. Full ~$331k register is a multi-year basis log.
4. ~~"Monarch alone can replace CapEx Excel"~~ → **ELIMINATED** — CapEx entered as hand lists; multi-account reimbursements require a register.

### Gaps identified

1. ~~No Monarch connectivity test~~ → **Partial:** host `api.monarch.com` returns HTTP 405 (endpoint exists); Cursor Monarch MCP **not** in dynamic tools catalog — live tag test still blocked.
2. No Ferdinand Unit1/Unit2 allocation % chosen.
3. Smoke detector Amazon amount still missing (Oak Park).
4. Grove missing amounts / undated Laura & Wayfair rows — **itemized** in Diagnostic Ideas Backlog.
5. No OpenRouter `plan_panel.py` on this host — panel skipped.
6. Topics suite/tools N/A — regression gate N/A for this personal-finance plan.
7. **827 purchase Closing Disclosure** still needed for initial basis (separate from 524 sale CD).

### Plan updates required

- [x] Add declutter / corrupt-file cleanup as early todo
- [x] Mark sale CD as BLOCKER
- [x] Require TY2025 vs TY2026 CapEx split
- [x] Add `/diagnose` Monarch hypothesis before more sheet-building
- [x] Point to existing 1098 locations under Documents/
- [x] Premortem + devil's advocate + tripwires (below)
- [x] **/diagnose complete** — hybrid locked; backlog + session summary below

---

## Adversarial Review

### 1.1 Completeness

| Question | Answer |
|----------|--------|
| Primary approach fails — fallback? | Fall back to **Excel-first** like prior years: one refreshed Personal Income workbook + Drive PDFs. Slower, known. Do not invent a third system. **Hybrid already is the middle path.** |
| Error conditions not handled? | Monarch miscategorization; transfer double-counts; CapEx paid from kids' accounts; Laura reimbursements vs original merchant; voided Riverton check #101; cash JCDA/Jesus with weak receipts; corrupt Drive uploads. |
| Files touched not mentioned? | `TaxReturn2023Corporate.pdf`; Merrill tax reporting PDFs; multiple CapEx sheet clones; corrupt Turbo xlsx stubs. |
| Suite / self-tests? | N/A — not a Topics tooling change. Verification = Monarch reconcile + CPA checklist completeness + Drive folder non-empty for sale. |
| Config dependencies? | Monarch account links; Drive folder IDs; optional MCP OAuth; CPA engagement; filing calendar (extension?). |

### 1.2 Assumption audit

| Question | Answer |
|----------|--------|
| Unverified about record? | Assumed sale docs would land in sale folder — **folder still empty of CD** (reconfirmed empty children 2026-09-17). Assumed Turbo xlsx uploads healthy — **some are corrupt**. |
| Formats/conventions? | Assumed STR → Sch E from prior Airbnb PDFs — still CPA confirm. Assumed LLC → Sch C — corporate PDF suggests verify. |
| Assumed absent without sweep? | Sale CD absent from sale folder **and** no title match for Closing Disclosure / ALTA tied to Ferdinand (only art consignment settlements). User may still have CD offline/email. |
| Public fact unverified before? | §121 duplex rules — now verified via IRS Pub 523 / 1.121-1. |

### 1.3 Negative space

| Question | Answer |
|----------|--------|
| Search terms not tried first pass? | Closing Disclosure, ALTA, Form 4562, depreciation schedule, corporate return, corrupt file sizes. |
| Adjacent files unread? | Contents of new `Documents/216` and `Documents/827` beyond 1098 titles; full prior-year tax PDF schedules. |
| Existing tools? | Monarch MCP (official) already does export-like access — better than inventing CSV paste automation first. |
| Derived views? | N/A Topics PATHS/atlas. Drive duplicates are the derived-view hazard. |

---

## Second-Pass Searches

| Search Type | Query | Findings |
|-------------|-------|----------|
| Inverse | Why Monarch-as-ledger fails / DIY §121 | User already hand-entered CapEx outside Monarch; IRS requires duplex allocation + dep. carve-out — DIY gain calc fails |
| Impact | Drive title search tax/Grove/Oak/Ferdinand/1098 | 1098s present under Documents; CapEx register duplicated; sale folder empty of CD; Turbo xlsx duplicates incl. 8-byte file |
| Adjacency | List `2025 Taxes` children | Documents/, New House CapEx/, 524 Ferdinand Sale/, many sheets + xlsx clones |
| Historical | Prior Personal Income.xlsx + 2023/2024 tax folders + this chat | Excel-tab system is proven; 2024 packet pattern exists; this chat created sprawl faster than habits |

---

## Panel Triage

| # | Advisor finding | Verdict | Evidence | Plan change |
|---|-----------------|---------|----------|-------------|
| — | Panel not run | skip named | No OPENROUTER_API_KEY / no plan_panel.py on this cloud host | None — self-review completed §1–§3 |

---

## Devil's Advocate

**Decision 1:** Monarch is the daily ledger; Excel is only a summary.
**Counterargument:** You already maintain CapEx and LTR as hand lists and Drive sheets. Forcing Monarch adds a habit you have not kept. Excel-first with annual Monarch reconcile may match actual behavior better and still beat cold-start from bank PDFs.
**Resolution:** **DIAGNOSED** — pure Monarch-only **ELIMINATED**. **Hybrid locked:** Monarch for operating; CapEx/Sale registers stay; one summary workbook. `decide-source-of-truth` → completed.

**Decision 2:** Build rich Tax Turbo multi-tab workbook + many Google Sheets.
**Counterargument:** Sprawl already produced corrupt uploads and duplicate CapEx registers. Minimum viable packet = prior Personal Income copy + one CapEx sheet + one sale tracker + PDFs.
**Resolution:** **Updated plan** — declutter first; prefer ONE CapEx sheet (canonical `1ojNHX…`) and ONE summary workbook; trash stubs.

---

## Premortem

**It failed because:** The household never consistently tagged Monarch, kept typing into scattered Sheets, handed the CPA three conflicting CapEx totals and an empty sale folder, and the CPA rebuilt everything from bank statements in March — the "turbo" packet was ignored.

- **Irreversible steps:** Filing a return with wrong §121/dep. figures; deleting Drive originals. CapEx categorization mistakes are reversible until filed. **Do sale math with CPA last; declutter and doc collection first.**
- **Tripwire:** STOP if (a) two CapEx registers disagree by >$500 after a cleanup pass, or (b) sale folder still has no Closing Disclosure when starting gain worksheets, or (c) Turbo xlsx on Drive is <10KB (corrupt).
- **Invented parameters:** "Weekly 15 min / monthly 30 min" — soft habit targets, not constraints. "~$75 receipt photo" — heuristic only. Remove any implication these are tax-law thresholds.

---

## Gap Analysis Checklist

### Coverage
- [x] Baseline recorded (Drive listing + IRS Pub check; Topics suite N/A — noted)
- [x] Touched artifacts re-checked this session (2025 Taxes folder, sale folder emptiness, 1098 locations)
- [x] Monarch hypothesis gated via proxy + host reachability (live MCP auth still pending)
- [x] CapEx TY split measured from workbook

### Risk
- [x] Rollback = trash duplicates / revert to Personal Income.xlsx pattern — reversible
- [x] No archive/original PDF rewrite in plan
- [x] Sale DIY gain calc explicitly forbidden
- [x] CapEx calendar-year split added

---

## Recommended next action

1. **Declutter Drive** — trash corrupt Turbo stubs + superseded CapEx; keep canonical IDs above.
2. **Drop 524 Closing Disclosure** (+ dep. schedules) into sale folder — still BLOCKER; also collect **827 purchase CD** for basis.
3. Put only **$53,660 TY2025 CapEx** (2 Dec 19 payments) on the 2025 CPA packet CapEx line; keep full register for multi-year basis.
4. Optionally connect Monarch MCP / tag operating accounts — **not** CapEx double-entry.
5. Refresh ONE summary workbook for CPA.

---

## Constraints (diagnose-added)

1. **Hybrid is mandatory for CapEx** — multi-payer reimbursements (Jake BoA, Megan Chase, Emma Merrill, Phoebe, joint, cash, Amazon card → Laura) will double-count if treated as Monarch-only expenses.
2. **TY2025 CapEx packet slice is small** — do not present ~$331k as a 2025 deduction/basis dump without year columns.
3. **No DIY §121 / dep. recapture worksheets** until sale CD is in the folder.
4. **Corrupt Drive files <10KB** are not editable originals — trash, do not "fix."
5. Live Monarch tag proof still requires Cursor MCP OAuth; do not block declutter/sale-doc collection on it.

### Dead Ends

- ~~Pure Monarch-only ledger for CapEx + LTR + Mercury~~ — contradicted by hand-list behavior + multi-account reimbursements (diagnose 2026-09-17).
- ~~Treating full Grove CapEx ~$331k as TY2025 work~~ — 81% is 2026; only $53,660 dated 2025.

### Implementation Notes

- CapEx year math source: local `/tmp/Personal_Income_2025_Tax_Turbo.xlsx` sheet `New_House_CapEx` (rows 16–18 summary; detail from row 22). Scratch: `~/.cursor/scratch/diag_capex_year_split.json`, `diag_capex_unknown_detail.json`.
- UNKNOWN $9,401.93 breakdown: Wayfair $1,189.08 (R47), Laura Connelly supplies $7,512.85 (R48), Village of Oak Park ~$700 (R81). Missing $: Viking Tile (R46), Crate & Barrel (R49), Amazon basement faucet (R51).
- Monarch host probe: HTTP 405 on GET `https://api.monarch.com/mcp` = endpoint present; MCP not in this session's dynamic tool catalog (`diag_monarch_reachability.json`).
- Drive sale folder `1K6_C7BcRF1PuC1l_MGjYWHBOGS2I9w9L` children query returned empty `{}` on 2026-09-17 reconfirm.

---

## Diagnostic Ideas Backlog

*Generated by /diagnose on 2026-09-17. Re-run /diagnose to re-rank.*

| # | Experiment | Question | Tier | Est. Time | Cancels If Clear | Status |
|---|------------|----------|------|-----------|------------------|--------|
| 1 | Hypothesis gate: Monarch-only vs hybrid (proxy: hand CapEx/LTR lists + multi-account payers; host MCP 405) | Can Monarch classify Oak Park + Grove CapEx + Mercury without parallel Excel? | T1 | 10 min | Pure Monarch-only build / CapEx-in-Monarch automation | **Done: CONTRADICTED pure Monarch-only; SUPPORTED hybrid** |
| 2 | CapEx year split from Turbo `New_House_CapEx` | How much Grove CapEx is TY2025 vs TY2026? | T0 | 2 min | Treating $331k as 2025 packet line | **Done: $53,660 (16.2%) / $268,532 (81.0%) / $9,402 UNKNOWN (2.8%)** |
| 3 | Drive packet health (file sizes + CapEx/Turbo copy counts) | Which files are canonical vs corrupt/duplicate? | T0 | 3 min | Editing wrong Turbo / dual CapEx registers | **Done: 4 CapEx-ish sheets, 3 Turbo copies, 2 corrupt stubs; keep 1ojNHX… + 1SaLWL…/1ffBOP…** |
| 4 | Sale folder children + Drive CD title search | Is 524 Closing Disclosure anywhere on Drive? | T0 | 2 min | Starting gain worksheets without docs | **Done: sale folder 0 files; no Ferdinand CD (only art settlements)** |
| 5 | CapEx UNKNOWN / missing-amount detail | What is the $9.4k undated + which rows lack $? | T0 | 3 min | Guessing TY for incomplete rows | **Done: Wayfair $1189 + Laura $7513 + Village ~$700; missing Viking/C&B/Amazon faucet** |
| 6 | Live Monarch MCP tag smoke (30 days export ↔ Mapping) | Do operating txns tag cleanly once MCP auth'd? | T2 | 15 min | Excel-first for operating too | **Pending — needs Cursor Monarch MCP OAuth** |

---

### Diagnostic Session Summary (2026-09-17)

**Cycles run:** 2

**Key findings:**
- Pure Monarch-only as the CapEx/LTR ledger is **contradicted**; **hybrid** (Monarch operating P&L + CapEx/Sale registers + one summary) is the supported model.
- Grove CapEx TY2025 slice is only **two payments totaling $53,660** (2025-12-19 JCDA + Riverton); ~81% of the register is 2026.
- Drive packet is cluttered: **corrupt Turbo stubs** + CapEx clones; **524 Ferdinand Sale folder still empty**; no Closing Disclosure on Drive for the sale.
- UNKNOWN CapEx $9,401.93 is three undated rows (Wayfair / Laura supplies / Village permit estimate); three rows still lack amounts.

**Options eliminated:** Pure Monarch-only CapEx; treating ~$331k CapEx as a single TY2025 line item; building more sheet tooling before declutter + sale CD.
**Dead ends confirmed:** Monarch-without-register for multi-payer Laura reimbursements.
**Recommended next action:** Declutter Drive to canonical CapEx + one workbook; drop 524 (and 827 purchase) Closing Disclosures; put only $53,660 CapEx on the 2025 packet line.
**Unresolved questions:** Live Monarch tag quality (MCP auth); Unit1/Unit2 §121 allocation %; Oak Park smoke detector $; exact dates for UNKNOWN CapEx rows; EPGC Sch C vs entity return.

---

## Success criteria (unchanged intent)

- No hand rebuild of monthly property ledgers from statements **OR** honest Excel-first fallback chosen → **hybrid chosen**
- Sale + CapEx have dedicated CPA-flagged trackers with docs present
- Next year = duplicate one workbook + clear numbers + keep Mapping/habits
