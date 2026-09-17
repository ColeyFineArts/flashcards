---
name: 2025 Tax Turbo
overview: Produce a CPA-ready 2025 tax packet matching prior years without re-keying every transaction into Excel — hybrid Monarch (operating P&L) + CapEx/Sale registers + one summary workbook. Diagnosed + partially implemented 2026-09-17.
todos:
  - id: decide-source-of-truth
    content: "DECIDED hybrid — Monarch for recurring operating P&L (LTR/STR/LLC/W-2 feeds); CapEx + Sale as Excel/Sheets event registers; ONE year-end summary workbook. Pure Monarch-only ELIMINATED."
    status: completed
  - id: declutter-drive
    content: "DONE 2026-09-17 — numbered topic folders 00–07 + _SUPERSEDED under 2025 Taxes; nested Documents flattened; duplicate STR PDFs trashed; folder map 10zS1YH…"
    status: completed
  - id: diagnose-monarch
    content: "/diagnose DONE 2026-09-17 — pure Monarch-only CONTRADICTED; hybrid SUPPORTED (see Diagnostic Session Summary)"
    status: completed
  - id: grove-capex-ty2025-packet
    content: "DONE — TY2025 CapEx CPA one-pager $53,660 on Drive (sheet 1aqJe32… + xlsx 1NFZPjh…); CapEx year-split README in New House CapEx folder"
    status: completed
  - id: sale-folder-need-doc
    content: "DONE — Closing Statement + purchase CD filed into 03/Platform PDFs (copies 1vvgkDyz… / 10TQQAum…); NEED doc retitled FILED"
    status: completed
  - id: monarch-connect
    content: "PARTIAL 2026-09-17 — CSV export profiled (2,521 TY2025 txns). Accounts present: Mercury, BoA Jake/AdvPlus, Checking STR, cards, Megan Chase, USB 9422/5907. STILL NEED: apply TAX:STR/LTR-OAK/EPGC/W2/HOA-GROVE tags in Monarch app (export can't create tags)."
    status: completed
  - id: monarch-csv-profile
    content: "DONE — CANONICAL CSV 1ElqZTJq… + Operating Snapshot 1yxtwpNG… in 02 Operating; checklist CURRENT 1xKJFKyt… in 01"
    status: completed
  - id: oak-park-2025
    content: "PARTIAL — LTR costs + vacancy accounting sheet 1O4SKBYp… (empty-apt clean 9/30 supports late vacancy). Still NEED exact vacant months + smoke detector $; Rocket 1098 for interest"
    status: pending
  - id: grove-capex-split
    content: "TY split published ($53,660 / 2025). Still fill UNKNOWN $9,401.93 dates + 3 missing amounts; resolve $3031.41 dup. Riverton $11,750 absent from Monarch — register remains required."
    status: pending
  - id: sale-docs-blocker
    content: "DONE (CDs) — Sale CS + both purchase CDs PROVIDED & fact-sheeted. Still NEED prior depreciation schedules (Sch E / 4562) + Unit 1/2 allocation."
    status: completed
  - id: grove-purchase-cd
    content: "DONE — 827 Grove CD PROVIDED; purchase facts sheet 1RuCiJmO… ($670k / close 2025-12-18); radon $2,550 on CD flagged vs CapEx register"
    status: completed
  - id: sale-facts-from-cds
    content: "DONE — 524 sale facts sheet 1p1fRuyT…: sold $600k on 2025-12-18; net $216,844.96; bought 2023-08-09 for $575k"
    status: completed
  - id: income-w2-1099
    content: "DONE — Form W2.pdf + 1099_JacobColey… provided in 07 Income & Bank Docs (1Xgd2Tc4… / 1k9ySc3h…)"
    status: completed
  - id: cpa-packet
    content: "PARTIAL — W2/1099 + CDs + STR + childcare + CapEx $53,660 in Drive; still assemble packet + BrightWheel EIN + dep schedules + CapEx UNKNOWN"
    status: pending
  - id: str-platform-pdfs
    content: "DONE — Airbnb Earnings 2025 ($8,053 gross / $7,691.41 net) + VRBO Payout 2025 ($10,214.30 / $9,395.53) filed CANONICAL in 524 Sale + 2025 Taxes; recon sheet 1Yo0IFga… — USE PLATFORM not Monarch for STR"
    status: completed
  - id: childcare-brightwheel
    content: "DONE — BrightWheel $30,668.55 provider cash; Aug $23,361.05 prepaid for BOTH kids; Kids Empire $233 + Zelle babysitting $100 EXCLUDED from 2441; recon 1CbssclH…"
    status: completed
  - id: optional-monarch-mcp
    content: Connect Monarch MCP read-only (https://api.monarch.com/mcp) after declutter — host reachable (HTTP 405); Cursor MCP not auth'd yet; CSV export unblocks operating P&L without MCP
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

**Drive layout (cleaned 2026-09-17)** under `2025 Taxes` (`1DkF0ealLM2YZ8IkCxBp-_jN9BJRJcxuu`):
- `00 READ ME — Folder Map` → map doc `10zS1YHwErhsUkFzL5qJ5EdbAvpz8EtVUe5-6eVQjLP0`
- `01 Checklist & Plan` → CURRENT checklist `1xKJFKyt…`
- `02 Operating — Monarch` → snapshot + `Monarch Money` CSV
- `03 STR + 524 Ferdinand Sale` → recon + `Platform PDFs + Closing NEED`
- `04 LTR — 216 Oak Park`
- `05 CapEx — 827 Grove`
- `06 Childcare — BrightWheel`
- `07 Income & Bank Docs`
- `_SUPERSEDED`

**Canonical Drive IDs (post-organize 2026-09-17):**
- CapEx register: `1ojNHX_z94FGMXBFKGcZ7Eta_iyQCrBvNjKbitKUJT4Q` (in 05)
- TY2025 CapEx CPA one-pager: `1aqJe32LWQwwAXdk7pDdo0MJHBu6mVnM7hbsE5W6aBmE` ($53,660 ONLY)
- TY2025 CapEx CPA slice xlsx: `1NFZPjhBOpIdLOB5TZfsHqCueDpBtoc_5`
- REFERENCE Personal Income: `1ffBOPDNo8vwt2LSKZJy__VG37nmIw9sy` (in 07)
- Sale/platform folder + status doc: `1K6_C7BcRF1PuC1l_MGjYWHBOGS2I9w9L` / `1fmtxZFK9fNWr7xY0CoN1isOxqv0wVP4oLdUCOZX8GlY`
- CapEx year-split README: `14XLuHT4cLvh4wDbHARM70x0aC09OX7-qwBhMwKtkah8` (05/Year-split notes)
- Monarch Money folder: `1HqEA5yfN12bAeY5yseTANPKEDBl0DJgT` (under 02)
- Monarch CSV export CANONICAL: `1ElqZTJqP4y6Flbh5ajlfLSaIpYJivR-w`
- Monarch TY2025 Operating Snapshot: `1yxtwpNGpHvdP_2VjXVlKlpBCcFlsSRAQm4duqiUEv2w` (in 02; duplicate copy archived)
- Tax Packet Checklist CURRENT: `1xKJFKytSL_vYDnoPquHeMH4qEnD9Zlpes5ejTOH9evE` (in 01)
- Airbnb Earnings 2025 CANONICAL: `1XJI1T8xRjoR_Gs_ZHdWoc39vm8gAW22J` (03/Platform PDFs)
- VRBO Payout 2025 CANONICAL: `1GfFqkAPJIeJ70qk6V2i9se_s4VRBXCje` (03/Platform PDFs)
- STR Platform vs Monarch recon: `1Yo0IFgaWi_fM6OFxznZH6HX25uwyDMp2puq-6PRA1fY` (in 03)
- 524 Sale Closing Statement: `1vvgkDyz8GSV2HufijcG8OKHNO8oZfKHv`
- 524 Purchase 2023 CD: `10TQQAumpK1nx1nXPiRzxE5E8om2UTm0f`
- 827 Grove Purchase CD: `1mWyNNIAIkzO63VeRjiVwqxrLsYBH0AtQ`
- 524 Sale Facts sheet: `1p1fRuyTDqtybIPVQOTm7o39dnZNgQ6ghylQ0i3bScvY`
- 827 Grove Purchase Facts sheet: `1RuCiJmOp8gbKbpnu7X8UuACXkyxddcsbOfMhOwYFvvc`
- Childcare recon: `1CbssclH6xcRW6B2frIu1t-n0yajUcT4JQNA7fLl1deY` (in 06)

**Trashed / archived:** Turbo stubs; CapEx clones; nested empty Documents; empty 827 Grove Purchase; root duplicate Airbnb/VRBO PDFs; prior checklists + old Turbo + snapshot copy → `_SUPERSEDED`.

## Key properties / entities

- W-2 Hindman; Megan income; GCM; art; EPGC LLC (Mercury)
- LTR: 216 N. Oak Park — `04 LTR / 216 N OAK PARK AVE` already has Rocket **1098**
- STR / two-flat: 524 Ferdinand Unit 2 through sale; Unit 1 owner allocation — CDs in `03/Platform PDFs`
- New primary: 827 N Grove — `05 CapEx / 827 N GROVE AVE` has USB Mortgage **1098** + purchase CD; Water `0245000849-01`; ComEd `7103042419`

## Already captured

- **Childcare BrightWheel (2026-09-17, updated):** Provider invoices **$30,668.55**; Aug 4 **$23,361.05** is prepaid/bulk for **both** Phoebe + Emma (charged to Phoebe account). **Kids Empire $233.28** and **Zelle babysitting $100** do **not** qualify for Form 2441. Qualifying Monarch cash ≈ Brightwheel **$30,640.58**. Recon `1CbssclH6xcRW6B2frIu1t-n0yajUcT4JQNA7fLl1deY`


- Drive `2025 Taxes` numbered `00–07` + `_SUPERSEDED` (nested Documents flattened)
- Oak Park 2025 known cash ~$21,423 (smoke detector TBD)
- Grove CapEx running ~$331,594 (excl. duplicate-suspect $3,031.41)
  - **TY2025 included:** $53,660 (16.2%) — JCDA $41,910 + Riverton $11,750 on 2025-12-19 only
  - **TY2026 included:** $268,532 (81.0%)
  - **UNKNOWN / undated:** $9,401.93 (2.8%) — Wayfair $1,189.08 + Laura supplies $7,512.85 + Village permit ~$700
- **STR platform PDFs (2026-09-17):** Airbnb Earnings Report + VRBO Payout Summary for 524 Unit 2
  - Airbnb gross **$8,053** / net payout **$7,691.41** (Megan → Checking 8507); remitted tax $641.20
  - VRBO gross **$10,214.30** / net payout **$9,395.53** (16 res / 46 nights); owner-remit tax $36.30
  - Combined net payout **$17,086.94** — Monarch cash only $12,794.57 (gap exactly = missing Apr/May Airbnb + 5 VRBO payouts)
  - Recon sheet: `1Yo0IFgaWi_fM6OFxznZH6HX25uwyDMp2puq-6PRA1fY`
- **Monarch CSV export** (2026-09-17) in `02/Monarch Money` — **2,521 TY2025 txns** profiled into Operating Snapshot
- **Closing docs (2026-09-17):** 524 sale Closing Statement + 524 purchase 2023 CD + 827 Grove purchase CD (Megan uploads; CANONICAL copies filed — Megan-owned originals may remain at root)
  - **524 sold** 2025-12-18 for **$600,000**; net to seller **$216,844.96**; USB payoff **$340,911.17**; commission **$15,000**; bought 2023-08-09 for **$575,000**
  - **827 bought** 2025-12-18 for **$670,000**; loan **$535,000** @ 5.875%; cash to close **$121,925.63**; radon mitigation **$2,550** on CD (CapEx check)
  - Fact sheets: sale `1p1fRuyT…` / Grove `1RuCiJmO…`
  - STR Airbnb/VRBO cash net **$12,794.57** (25 txns; many Airbnb still in Travel & Vacation)
  - Hindman deposits ~$55.7k + Freeman ~$11.8k; Mercury Business Income ~$282k
  - Monarch Home Improvement only **−$3,870** — CapEx register still authoritative
  - JCDA Check #140 −$41,910 present as Transfer; Riverton $11,750 **absent** from Monarch
- Official Monarch MCP: `https://api.monarch.com/mcp` (host reachable; Cursor MCP not connected; CSV unblocks operating P&L)

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

1. ~~Declutter Drive~~ **DONE** (implement 2026-09-17).
2. ~~Publish TY2025 CapEx $53,660 CPA line~~ **DONE** (sheet + xlsx on Drive).
3. **Drop 524 Closing Disclosure** (+ dep. schedules) into sale folder — still BLOCKER; also collect **827 purchase CD** for basis.
4. Fill CapEx UNKNOWN dates / missing amounts; confirm Laura $3,031.41 dup in bank.
5. Optionally connect Monarch MCP / tag operating accounts — **not** CapEx double-entry.
6. Finish Oak Park LTR (smoke detector $) + assemble remaining CPA packet forms.

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

---

### Implementation Session Summary (2026-09-17)

**Batches:** 3 (declutter+labels; CapEx CPA publish + sale NEED doc; Monarch CSV profile)

**Completed:**
- Drive declutter: 6 files trashed (2 corrupt Turbo, 1 superseded CapEx, 3 CapEx clones)
- Canonical labels applied; TY2025 CapEx CPA one-pager ($53,660) + slice xlsx uploaded
- Sale folder NEED Closing Disclosure checklist created — **CDs now PROVIDED** (sale CS + purchase CDs fact-sheeted)
- CapEx year-split README in New House CapEx folder
- Local implemented workbook with `TY2025_CapEx_CPA` sheet + checklist stamps
- `implement-plan` skill parked under `.cursor/skills/implement-plan/`
- **Monarch CSV profiled:** 7,027 rows / 2,521 TY2025; Operating Snapshot + STR detail published to Drive; CSV renamed CANONICAL; checklist updated
- **STR platform PDFs filed:** Airbnb + VRBO 2025; combined net payout $17,086.94; Monarch understates by $4,292.37 (exact missing payouts identified)
- **Income docs PROVIDED:** Form W-2 + 1099 in `07 Income & Bank Docs`

**Still blocked / pending:**
- Prior depreciation schedules (Sch E / 4562) + Unit 1/2 allocation for 524 sale gain
- CapEx UNKNOWN dates / missing amounts / Laura dup confirm; radon $2,550 on Grove CD vs register
- BrightWheel provider EIN for Form 2441; CPA allocate Aug prepaid across both kids
- Apply TAX:* tags inside Monarch app (export cannot create tags)
- Oak Park smoke detector $; Monarch MCP OAuth; final CPA packet assembly

**Provided (no longer blockers):** W-2, 1099, 524 sale Closing Statement, 524 purchase 2023 CD, 827 Grove purchase CD.

**Verification:** Drive folder re-list shows no corrupt stubs; TY2025 CapEx sheet + CPA xlsx present; sale folder has NEED doc only (no CD yet); Monarch Money folder has CANONICAL CSV + Operating Snapshot + STR sheet. Topics `run_suite.py` N/A for this personal-finance plan.
