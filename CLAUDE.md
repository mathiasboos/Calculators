# Calculators — Project Notes

## What this project is

Offline, dependency-free Swedish pension and salary calculators built as standalone HTML files. Each file is self-contained — no build step, no server, no external dependencies beyond fonts and charting libraries loaded via CDN where needed.

## Tax constants (verify EVERY December for the following year)

### Single national values
- IBB (Inkomstbasbelopp) — 2026: 83,400 kr
- PBB (Prisbasbelopp) — 2026: 59,200 kr
  ⚠ IBB and PBB are DIFFERENT numbers and move independently — don't
  assume one tracks the other
- SKIKT (skiktgräns) — 2026: 643,000 kr
  → threshold on BESKATTNINGSBAR förvärvsinkomst (income AFTER grundavdrag)
  → used directly in the state tax calculation
  → only appears in Net_Salary_2026.html — confirmed not duplicated
- BRYT_UNDER (brytpunkt) — 2026: 660,400 kr/year
  → threshold on GROSS salary (BEFORE grundavdrag) — the figure usually
    quoted in media/Ekonomifakta
  → display-facing only in Net_Salary_2026.html, NOT used in the tax calc
  → relationship: BRYT_UNDER = SKIKT + grundavdrag (gap ≈17,400 kr in 2026)
  → ⚠ DO NOT confuse SKIKT and BRYT_UNDER — check which one a source is
    quoting before editing either constant

### Kommunalskatt (municipal tax) — NOT a single number
- Full ~290-municipality lookup table, sourced from SCB (published
  mid-December each year)
- Update each December: replace the ENTIRE table, not a single value
- Calculation rule: kommunal = floor(bfi × ki) — truncate, do not round
  (Skatteverket rule)

### Jobbskatteavdrag — NOT a flat number where implemented in full
- Full version: two age-banded piecewise formulas
  - Under 66 (jobbavdragUnder): 4 income bands, depends on ki
    (kommunalskatt) and ga (grundavdrag)
  - Over 66 (jobbavdragOver): 3 income bands, flat coefficients, no
    kommunalskatt dependency
  - Band thresholds expressed as PBB multiples (0.91×, 3.24×, 8.08×,
    1.75×, 5.24×) — if PBB changes, kr values shift automatically, but
    the multiplier coefficients themselves (0.3874, 0.251, 0.22, 0.07,
    etc.) can ALSO change in Skatteverket's yearly spec — verify the
    whole function each December, not just constants
  - Source: Skatteverket, published ~December each year

### Grundavdrag — NOT a flat number, two piecewise formulas (most complex constant in this project)

- Input: ffi (fastställd förvärvsinkomst — gross income after any salary
  sacrifice, before grundavdrag itself is applied)
- Rounding: ceil100 — rounds UP to nearest 100 kr
  ⚠ opposite direction from kommunalskatt, which floors — don't assume a
  consistent rounding rule across this file
- Cap: Math.min(ga, ffi) — grundavdrag can never exceed actual income

- Under 66 (grundavdragUnder): 5 income bands (flat floor → rising →
  flat peak → falling → flat floor), thresholds at 0.99×, 2.72×, 3.11×,
  7.88× PBB

- Over 66 (grundavdragOver): ordinary 5-band formula above PLUS an
  additional "förhöjt grundavdrag" top-up with 11 MORE bands of its own
  (thresholds from 0.91× up to 12.84× PBB) — this is the single most
  complex piece of tax logic in the project

- All thresholds expressed as PBB multiples — if PBB changes, kr
  breakpoints shift automatically, but the coefficients themselves
  (0.423, 0.77, 0.293, 0.687, 0.20, 0.057, etc. — 15+ distinct constants
  across both functions) come from Skatteverket's spec and can ALSO
  change each year. Verify the full function, every band, every
  December — do not assume only PBB moved.

- Source: Skatteverket, published ~December each year

### Cross-file consistency — annual update scope differs per file

| What | Net_Salary_2026.html | Salary_Exchange.html | Salary_Exchange_Consumer.html |
|---|---|---|---|
| Kommun lookup table | full table | full table (duplicate) | single default (KOMMUNAL_DEFAULT) |
| Jobbskatteavdrag | full piecewise formula | simplified (brytpunkt check + flat 20% state tax above it) | simplified, same as Salary_Exchange |
| Grundavdrag | full piecewise (both age formulas) | not used — works with marginal rates only, never computes gross-to-net | not used — same reason |
| Brytpunkt constant | BRYT_UNDER = 660,400 (annual) | LIMIT_BRYTPUNKT = 55,033 (monthly, confirmed ≈ matches) | not used |
| SKIKT (skiktgräns) | yes | not used | not used |
| Annual update needed | full re-verify of all tax logic | replace kommun table; verify LIMIT_BRYTPUNKT | update KOMMUNAL_DEFAULT only |

⚠ Updating Net_Salary_2026.html's kommun table does NOT automatically
update Salary_Exchange.html's copy — both must be replaced separately
each December, or the two calculators will disagree.

### Occupational pensions (tjänstepension) — Retirement_Planner.html only

Source: `reference/Typfallsmodellen_v4.8.xlsb` (Pensionsmyndigheten), VBA
modules `Tjänstepensioner` (premiums, payout divisor) and
`TjänstepensionerFörmån` (defined benefit, function `FTJP`). Extract the VBA
with oletools (`olevba`) to read it. When a new Typfallsmodellen version is
published, replace the workbook and re-verify everything below.

- Engine (`occMakeCtx`, `occStep`, `occPremium`, `occHistoryOpening`) is a port of
  the workbook's yearly loop: balance = premium·y^0.5 + prev·y·keptAfterFees
  + prev·inheritance gain, minus yield tax, minus payout·y^0.5. The payout is
  set ONCE at pension start (balance ÷ cohort divisor) and then stays flat in
  constant prices (the workbook indexes it with prisbasbelopp)
- Embedded data, extend/verify every year: `OCC_HIST` (income index, IBB, fund
  return, kept-after-fees, statslåneränta — sheets Nyckeltal / Några tal),
  `OCC_DT` (premiepension divisor by birth year × age, Nyckeltal c27:c71) and
  `OCC_ARV` (inheritance gains, sheet `arv PP`). Forward years use the
  Parameters-tab real return (default 1.7 %, `occReturn`) and
  statslåneränta 2.5 % (`avkskatt` Case Else) → yield tax 0.375 %
- `occPremium` — premium rules per scheme AND year (tlITP1, tlITP2A, SAF_LO,
  tlkap_kl, tlakap_kr, Kapan + PA_indiv, tlPA16). ⚠ PA16 flex steps up yearly
  (2026 is a blended 6.125 % / 31.625 %); PA16 Avd 2's rate depends on birth
  year (<1965). `OCC_SCHEMES` only holds the 2027+ values shown in the UI
- Past wages = today's salary × income index(year) / income index(2026);
  future IBB grows with the real wage growth (workbook convention). Scheme
  accounts without an entered Opening get the history value (`openAuto`)
- `KAPKL_DB_RATES` / `kapklDbRates` — KAP-KL förmånsbestämd cohort rates
  (`KAPKL_f`; the 1966 value 0.268 is copied as-is from the VBA)
- `PA16A2_DB_RATES` / `pa16a2DbRates` — PA16 Avd 2 förmånsbestämd (`tlPA03`)
- `occDbBenefit` — ITP 2 (`tlITP2F`), KAP-KL, PA16 Avd 2, using IBB(start
  age − 1). ITP 2 gets both its own start-age factor (−0.5 %/month before 65,
  +0.6 %/month after) AND the deltal(65)/deltal(age) ratio, as in the workbook
- State income pension (inkomstpension, account scope `inkomst`, `isIpEngineAcc`)
  uses the same approach: `ipPgi` / `ipRatt` (16 % of PGI = 93 % of income up to
  7.5 IBB, credited the year after the income), `ipStep` (VBA `IP_`: pension
  index, admin fee, inheritance gains `OCC_ARV_IP`; from 2015 the right is multiplied by
  gällande index / inkomstindex while that is below 1 — VBA `Bindex`, only 2015–2017 in the data).
  Payout as the workbook's Table 2 (VBA `IP_`, `fnDeltal_IP2`): start year = balance at the start of the
  year ÷ `ipDelningstal(age, by, true)` (mid-year blend of this age and the next, rounded like aaDeltal.bas);
  the right credited in the start year (+ arv, index, admin fee incl. the fee on the whole balance,
  `statState.pend`) is added the year after at that age's divisor; then indexed with the pension index ÷ 1.016.
  Divisor: cohorts 1958–2034 at ages 61–82 use the workbook's own mortality-based figures
  (`DELNINGSTAL_IP_OWN` / `ipOwnDelningstal`: unisex, born 1938–2034 × ages 61–105, generated with the port's
  `DeltalTables.ownIncomePension`), others the published `DELNINGSTAL_IP` (Nyckeltal). The Garantipension
  underlag keeps the published table.
  Balance from the start year on (Portfolio Value, Dashboard totals) = the pension still to come, as the port's
  `finishLedger`: next year's FULL pension (incl. the pending last right) × `ipOwnDelningstal(age + 1)` (`_ipValue`);
  return = the rest of the change. Matches the port's Portfolio value within 0,01 % at every age to 95 (ITP1/1981/68).
  `recalcYearly` rows (work again) keep their run-down balance.
  The Dashboard replacement rate adds the last IP/PP rights at the published divisors to the first year
  (`row.lastRight1`), as Table 1 / kompensationsgrad does (`creditLastPensionRight`).
  Data to extend yearly: `OCC_IP` (pension index, admin fee, PBB,
  förhöjt PBB, lägsta pensionsålder by cohort), `OCC_ARV_IP` and `DELNINGSTAL_IP_OWN`. Forecast pension
  index = income index (no balancing brake), as in the workbook
- Premium pension (premiepension, scope `premie`, `isPpAutoAcc`): only the HISTORY
  uses the workbook (`ppRatt` = 2.5 % of PGI credited the year after, `ppStep` =
  VBA `PP_pbh`: fund yield, `arv PP` gains, RGK interest `PP_RGK`; no manager fee, since the
  workbook's default "efter fondavgifter" makes `PPMavg` 0 — `PP_RETURNS_NET_OF_FEES`, `ppFee` kept).
  Used when Opening is blank/0 (`openAuto`); the projection keeps
  the planner's own growth so equity %, fees, yield tax and fixed return stay
  editable. The right is credited at year end: no return on it in the year it is credited (`_retBase`).
  Data to extend yearly: `PP_RGK` (Några tal C19). Payout = balance ÷ cohort divisor each year
- Automatic openings (`openAuto`: scheme premium rows, income and premium pension) are the balance at
  31 Dec of the year before START_YEAR, and the projection runs START_YEAR as a WHOLE year for them
  (`autoFullYear`: full premiums and return, as the workbook). A typed opening is today's balance and
  keeps the partial first year (projection month). Before this the scheme opening included this year's
  premiums so far but not their return, and the first year got only the rest of the year's return
  (ITP 1 ended ~0,9 % low)
- Garantipension and Inkomstpensionstillägg (`benefitKind` `garanti` / `ipt`,
  `isUnderlagAcc`): both derive from the "underlag" = the income pension the person
  would get if the whole 18.5 % went to the income pension (`underlagSeries`:
  `ipStep` with rate 0.185, balance ÷ divisor at riktåldern, last right NOT
  included, then indexed like the income pension). `gpAmount` = VBA `Bidrag.gp`
  (2026+ levels, single/cohabiting, residence years), `iptAmount` = VBA `tillagg`;
  paid from max(Wdrl start, riktålder `RIKTALDER`). The underlag is notional (no
  balance shows in the assets). Data to extend yearly: `RIKTALDER`, the 2026+ levels
  in `gpAmount`/`iptAmount`. Agrees with the port's yearly table within ~0.3 %
  (the port's summary tile adds the last right, +2.5 %). Not modelled: tilläggspension
  (ATP) in the underlag, bostadstillägg interplay, residence years < 40 for IPT
- Tilläggspension (ATP, `benefitKind` `tillagg`, `atpSeries`/`atpFull`): only for
  people born 1938–1953 (zero for later cohorts, by law and in the workbook).
  ATP points = PGI/förhöjt prisbasbelopp − 1 per year (age < 65, from 1960), average
  of the best 15 years, 0.6·points·PBB + 0.96 (single) / 0.785 (couple) · PBB, scaled
  by years/30 (min 3), × (1 − `andelNya`) (the rest sits in the income pension),
  × the early/late-withdrawal factor TWICE (workbook quirk, kept), indexed like the
  income pension, plus the guarantee top-up from pre-1995 points. `andelNya` also
  scales the income/premium/guarantee rights of those cohorts (ages ≤ 64). Agrees
  with the port exactly (±0.4 %) when the payout year is 2026+ (flat PBB 59 200).
  Not modelled: PGB points (childcare/studies), ATP in the Garantipension underlag
- PGB — pensionsgrundande belopp (section under Assets, `pgbItems`, `pgbCalc`,
  `rightsBase`): items Child (1–4 PGB years, default 4) / Conscription / Studies / Activity-sickness compensation,
  ported from `VBA_go` (PGB block), `Pensionssystemet.pgb_barn` and the workbook sheet
  `PGB`. Child: 4 years (birth year … +3) = max(income year of birth − PGI, 0.75·MPGI − PGI,
  1 IBB), rounded up to 100 kr; conscription: 0.5·MPGI × days/365 (1995–2010 and 2018+,
  120–730 days); studies: terms × grant per term × 1.38 (from 1997) floored to 100 kr;
  sickness/activity: manual amount (sheet says "ange manuellt"). PGI + PGB ≤ 7.5 IBB; ages 16 …
  riktålder, before retirement, born ≥ 1938. PGB feeds income pension, premium pension (history),
  Garantipension underlag and ATP points (`pgbyears` rule: < 5 years use pgi + PGB·185/160 for
  IP/GP, only pgi for PP). Own PGI for the child/SA rules assumes work from the earliest
  contribution start of the IP/PP/underlag accounts. Data to extend yearly: `PGB_DATA` (MPGI,
  study grant per term). The workbook's study/conscription cell formulas could not be read
  (xlsb), so the 1.38 factor and days/365 are inferred from the sheet's cached values; no
  oracle (the HTML port has no PGB inputs). Not modelled: KPI scaling of the first child's PGI
- Assets tab Contribution cell of state accounts: Income pension shows "16 % of PGI (93 % of salary, up to 8,07 IBB)" (`ip.rule`), Premium pension (`isPpAutoAcc`) shows "2,5 % of PGI (93 % of salary, up to 8,07 IBB)"
  (`pp.rule`) with no Fixed/Salary buttons. The text is only an explanation (0,93 × 8,07 is not exactly 7,5 IBB). The projection's forward premium pension contribution is the workbook's `ppRatt(_occ, a, yr) × ppRgk(yr)`:
  2,5 % of PGI (93 % of the PREVIOUS year's income, capped at 7,5 IBB, incl. PGB and `andelNya`), credited at year end in the year AFTER the income (income year `yr−1` inside the account's Contrib window, before retirement,
  so the last right is credited in the retirement year), with the RGK interest of the year; same helper as the history and the Assets cell / `effectiveContrib` (no monthly growth of the contribution, as in `ppStep`).
  Replaced the old editable 2,32 % × salary rate. ITP1/1980/40 000: age 68 total tax 119 154
- "+ Add State Account" menu (`STATE_KINDS`, `addStateAccount`): Inkomstpension and
  Premiepension (auto opening), Garantipension / Inkomstpensionstillägg / Tilläggspension (all calculated,
  see above), Annat (generic)
- Verified against the HTML port of v4.8 (same inputs, inflation 0, growth 1 %):
  premium-only schemes, the DB schemes at pension age ≥ 65 and the income pension
  agree within ~1 %
- Not modelled: SAF-LO's STP part (~3 %), KAP-KL's pre-1998 PFA transition,
  ITP 2 salary-growth caps (`DC_underlag`), the `Rng_ddelat` divisor
  adjustment, AKAP-KR's 6 %-flat rate above 69, and the workbook's behaviour
  for tjänstepension taken before 65 (it pays more than the planner there)

### Plan presets and guided walkthrough — Retirement_Planner.html

- Header dropdown `#plan-preset` → `PLAN_PRESETS` (key → function, add new typical cases
  there plus a `<option>` and `preset.*` i18n keys). Every choice asks for confirmation;
  the plan saved in the browser is untouched until Save
- Its placeholder reads "Begin here to create a plan" and pulses gold/green/blue for 5 s on load
  (`.preset-attn`, CSS only)
- `planClear()` — zero salaries, remove all accounts and PGB items, expense categories kept
  at 0, withdrawal plan reset (`wdPlan` = `WDP_DEFAULT`). Birth / retirement / life-expectancy years are kept
- `planTypical1980()` — ITP 1, born 1980, 40 000 kr/month, retirement at riktålder (68),
  life year 2075, typical pensioner expenses (`EXPENSES_INIT`, 18 400 kr/month), ITP 1 +
  Premiepension + Inkomstpension with automatic (Typfallsmodellen) openings. ITP 1 Contrib End
  is 66 (`CE_MAX`; `occPremium` stops ITP 1 premiums at 66 from 2023, workbook rule)
- Personal tab works in ages: Begin work at age (`#workStartAge`, default 23), Retirement Age
  (`#retirementAge`, tick box `#retTargetAge` = follow riktålder) and Life Expectancy Age
  (`#lifeAge`, default 95). The hidden `#retirementYear` / `#lifeYear` stay the engine's source
  of truth (`persApplyAges` writes them, `persSyncFromYears` restores ages from older saves).
  New state accounts start contributions at the work age; schemes at max(work age, scheme
  from-age); `pgbWorkStart` = birth + work age. Gross/Net Retirement salary rows were removed
- Assets tab contribution window: Contrib End = retirement age (`defaultCe`). Contrib Start depends on
  who creates the row — an account the USER adds (`addAccount`) starts at `START_YEAR`, i.e. the person's
  age today, because it starts now; schemes (`addOccScheme`) use max(work age, scheme from-age) and state
  rows (`addStateAccount`) the begin-work age, since their opening balances are built from the whole income
  history. Numerically the user-added case is a display fix only: the projection starts at `START_YEAR` and
  gates contributions on `yr >= a.cs`, so an earlier `cs` never moved money there — hence no save-version key.
  A user-added row therefore no longer follows Personal-tab changes (`persRelinkAccounts` only moves a `cs`
  still equal to `_csFor`); scheme and state rows still do (trackers set by `persTrack`). "End = retirement
  year" means no contribution in the retirement year itself (fixed contributions; salary/scheme/IP ones were
  already gated)
- Equity share: EVERY account without its own brackets follows Pensionsmyndigheten's default fund
  (AP7 Såfa) — `ppGlideEq` (100 % equities to 55, linear to 33 % at 75) via `usesGlideEq` in
  `calcEquityShare`; the per-account "Equity share" control (Details) switches to eq1–eq3 (`a.ownEq`).
  Excluded: DB, benefit and work rows, and anything with a `fixedRet` (no equity share at all).
  `isPpGlide` = `isPpAcc && usesGlideEq`, kept for the premium-pension wording; `populateMCDefaults`
  follows the same helper. ⚠ The UI names AP7 Såfa ONLY on the premium pension row (`isPpAcc`), where the
  money really is in that fund: chip `acs.eq.default`, note `pp.glide`, tooltip `pp.glide.tip`. Every other
  account gets the same curve under a neutral name — `acs.eq.auto` "Glidepath" / "Nedtrappning",
  `acs.eq.auto.note`, `acs.eq.auto.tip` — because a state-pension fund name on an ISK reads as if the ISK
  were invested in it. `fundSeg` in `acsCard` branches on `isPpAcc(a)`; do NOT re-unify the two sets of
  strings. The internal names (`ppGlideEq`, `usesGlideEq`, `isPpGlide`) stay Såfa-flavoured on purpose. The three bracket AGES live in ONE place, the `#acs-eq-bar` on the ASSETS
  tab above the steps (`ageFrom1`/`ageTo1`/`ageTo2` = 55/63/73): they are edited where the accounts
  are, and the bar shows itself only while some account is on own brackets (`acsSyncEqBar`,
  `acsAnyOwnEq`); the Details link calls `acsFocusBrackets()`, which reveals, scrolls to, flashes and
  focuses it instead of leaving the tab. Keep the ids unique wherever the bar lives — until Nov 2026
  that bar existed twice with the same ids
  (Assets 65/75/85, Projection 55/63/73), so the brackets in force were not the ones the headers named
  and a plan changed behaviour the first time it was saved and reloaded (`collectPlanState` snapshots
  `querySelectorAll('input[id]')` last-wins). Saves carry `eqRulesV: 1`; restoring an older save sets
  `ownEq` on any account whose brackets are NOT one of the three old defaults (0,9/0,4/0,75 · 0,8/0,4/0,7
  · 0,6/0,4/0,3), so typed brackets — 0/0/0 included — survive.
  Typfallsmodellen itself has one fund-return assumption, no glidepath
  Contractual scheme premium rows (`isOccPremiumAcc`) have the same tick box: unticked = the
  Parameters occupational return (1,7 % net); ticked = `engineStep` runs `occStep` with the account's
  own net return (equity % brackets × eq/bond returns − Mgmt %); premiums, arv gains, yield tax and
  history stay Typfallsmodellen's
  Default Mgmt %: Typfallsmodellen's general fund fee 0,20 % (`MGMT_DEFAULT_TM`; 1,9 % before fees → 1,7 % after,
  as the workbook's forecast) on the rows the workbook owns — contractual scheme rows, premium pension
  (income pension keeps 0,10 %, inert) — and **1 %** (`mgmtDefaultPrivate`, from the Parameters `#mgmtFee`
  field, which `addAccount` now actually reads) on private ones: ISK, endowment, private pension, the generic
  occupational rows and salary exchange. Cash 0. `#mgmtFee` also drives the Scenarios drawdown (10450).
  Saves carry `ppFeeV: 1` (older saves lift the old 0,10 % / 0,04 % premium pension defaults to 0,20 %) and
  `mgmtV: 1` (older saves lift the exact literal 0,20 % to 1 % on non-workbook rows only; the guard runs
  after the `ppFeeV` one and excludes premium pension so nothing is lifted twice). Typed fees are never moved
- State pensions are always lifelong: `we = 9998` for every state account (defaults, new accounts,
  restore), the Withdrawals tab's details table shows ∞ as their end
- Premium pension payout is lifelong like Typfallsmodellen (`ppkassa` + `PP_pbh`): each year
  balance ÷ the cohort PP divisor (`occDelningstal`, 2 from age 101), and the projection adds the
  inheritance gains (`occArvGain`, sheet "arv PP", ÷ yield^0.5) to the PP balance every year (`isPpAcc`). The
  payout is not capped by the balance. As `PP_pbh`, the pension is paid during the year and costs payout × yield^0.5
  of the year-end balance (`_wdCost`), and the fee (Mgmt %) is a rate off the return on the OPENING balance
  (`_ppFee`: 1,9 % − 0,20 % = the workbook's 1,7 %; pro-rated in a partial first year), not a charge on the year-end
  value. Before these two the payout drifted above the port at old ages (+2,7 % at 94, born 1981). Shape matches the port (born 1968: 0.97 / 0.81 / 0.60 of the
  age-68 amount at 85 / 95 / 100; born 1980 stays level)
- Default assumptions = Typfallsmodellen's prognosis standard: equities and bonds real return
  1,9 % BEFORE fees (each account deducts its own Mgmt %; 1,7 % after the workbook's average 0,2 %
  fee), inflation 0 %, occupational schemes 1,7 % (already net of fees), salary
  increase 0 %. Salary convention as in Typfallsmodellen: the Personal-tab gross salary is 2025
  money, raised by the 2026 income index (`salBaseUplift` = OCC_HIST 228.08 / 220.23, +3,56 %:
  40 000 → 41 426 kr/month) wherever the engine reads it (`grossSal`); net salary is computed on the
  raised amount; hint `#sal-uplift-hint`. Yearly payouts vs the port's Table 2 (born 1960–1995, retiring
  66–70, ITP 1, to age 99): ITP 1, income and premium pension within 0,06 % at every age; ITP1/1980/40 000 at 68:
  all three +0,01 %, net salary 32 752 vs 32 770. Oracle scripts must enter
  the 2025 salary (port wage ÷ salBaseUplift). The 1,7 % is AFTER fees; the workbook's unticked
  "efter fondavgifter" option deducts a general 0,2 % fee (1,9 % before fees)
- `resetAssumptions()` puts eqReturn / bdReturn / inflation / occReturn / salaryIncrease back to their
  HTML defaults (`defaultValue`); called by `planClear` (so every preset) and the Parameters-tab
  button "↺ Typfallsmodellen defaults". Saves carry `assumptionsV: 2`; restoring an older save moves
  untouched old defaults (eq 4,5/1,7, bonds 0,5/1,7, salary +1,0) to 1,9 / 1,9 / 0,0, typed values stay
- Dashboard Replacement Rate = Typfallsmodellen's kompensationsgrad: pension before tax in the first
  retirement year (wdrlState + wdrlPIns) ÷ final salary before tax (mean of the last 5 working
  years' earnedIncomePre). `repRateNet` (after-tax ratio) is still computed but not shown.
  ITP1/1980/40 000: 74,7 %, same as the port
- Withdrawal rules (Withdrawals tab): `WD_RULES` / `wdRule(a)` / `wdRuleErrors(a)` / `wdDefaultYear(a)` give each
  pension an earliest start age and (P-insurance, Temporärt only) a shortest withdrawal time; a broken rule shows
  a red message in the Withdrawals step (`.wdp-err`; payout periods below the minimum are disabled), the typed value is kept. Verify the ages yearly.
  ITP 1 55/5, ITP 2 55/5, SAF-LO 55/5, KAP-KL 61/10, AKAP-KR 62/10, PA16 Avd 2 61 (born ≤ 1965) or 63 /5,
  PA16 Avd 1 63/5, Individual pension account 55/5 (earliest age / min years). Premium + Income pension: earliest
  target age (riktålder) − 3; Garantipension + Inkomstpensionstillägg: target age
- Wdrl Start of every state and P-insurance account follows the Personal-tab retirement age, raised to its earliest
  allowed age (`wdDefaultYear`). Accounts carry `wsAuto`; `wdSyncStart()` re-applies it from `persApplyAges` /
  `onRetirementYearChange` / restore (a Temporärt account keeps its number of payout years), and choosing an age in the
  Withdrawals steps clears `wsAuto`. New schemes start lifelong (∞); the Individual pension account is lifelong by default. Saves carry `wdRulesV: 2`;
  older saves flag accounts still on the retirement year (or the first version's per-scheme ages) as `wsAuto`
- Contrib End (Assets tab) of every state and P-insurance account with a contribution window follows the Personal-tab
  retirement age (`ceAuto`, `ceSyncEnd()` from `persApplyAges` / `onRetirementYearChange` / restore), capped at the
  scheme's last premium age `CE_MAX` (ITP 1 66, ITP 2 and SAF-LO 65, KAP-KL 67, PA16 Avd 2 69, PA16 Avd 1 67; AKAP-KR and
  the Individual account have no cap; verify with `occPremium` yearly). A red 9 px note "Max age N" (like the "auto" label, absolutely placed so the row stays aligned) shows
  under the field while the cap is the binding limit. Typing a Contrib End stops the following; a typed age above the cap is
  set to the cap (`ceClamped`, note shown). Saves carry `ceRulesV: 1`; older saves flag accounts still on the retirement
  year (or the cap) as `ceAuto`. After-tax accounts keep the older `persRelinkAccounts` behaviour
- Contrib Start (Assets tab) of the contractual premium rows shows a muted 9 px note "Low age N" (`csNote`, `OCC_SCHEMES[..].prem.from`,
  enforced by `occPremium`; verified against the VBA): ITP 1 25, ITP 2 – ITPK 28, SAF-LO 22 (23 in 2022, 24 in 2021, 25 in 2000–2020,
  28 before 2000), KAP-KL 21, PA16 Avd 2 23, AKAP-KR 21 (born ≤ 1985; born 1986+ from 2014 premiums start with the work) and PA16 Avd 1
  (no minimum in the workbook) both use the planner floor 18. Verify yearly
- Expenses tab doughnut (`buildExpenseChart`) is styled like the Withdrawals doughnuts: no segment borders, 40 % hole, dark tooltip,
  and the `pctLabels` inline plugin draws each share inside its segment (hidden below 3 %, dark text on light segments)
- Dashboard charts: Portfolio Value and Portfolio Value by Account Type are in MSEK (`fmtMsekTick`: one decimal, no "M"; comma in Swedish),
  the two income charts use kronor with a thousand separator (`fmtKrTick`, "10 000") and start on Monthly. `taxAdjPort` = After tax Acc +
  Before tax Acc × (1 − kommunalskatt), so it already INCLUDES the after-tax accounts (goal fulfilment, vs. Goal and the Tax-adjusted line all use
  it). In Tax-adjusted view the Before tax Acc line is hidden and the tooltip lists After tax Acc + Before tax Acc after est. tax, which add up to the
  headline. Swedish KPI names: Kompensationsgrad, Uttagsnivå, Kapitalet tar slut
  Portfolio chart per mode: Total = Portfolio value + After tax Acc + Before tax Acc (no Savings goal line, tooltip without goal / vs. Goal);
  Tax-adjusted = tax-adjusted line + Savings goal + After tax Acc + "Before tax Acc, after est. tax" (portBeforeTax × (1 − munTax)), tooltip with goal / vs. Goal.
  Tax chart: Monthly/Annual toggle (`_taxChartMode`, default monthly, `setTaxChartMode`), title "Tax Per Month (kr/mth)" / "Tax Per Year (kr/yr)", kronor axis;
  tooltip total = sum of the series. The tax chart shows four categories (the 7 % Pension fee `cTaxPens` is NOT drawn in the chart or shown in the Projection table / export) and uses the same ones before and after retirement: `row.cTaxMun` (municipal tax after
  jobbskatteavdrag / förvärvsavdrag, incl. begravningsavgift), `row.cTaxState`, `row.cTaxFee` (public service fee: `ps` for salary years, `publicAvg` for pension years), `row.cTaxPens` (general 7 % pension
  fee net of its tax reduction, `pensAvg − pensRed`; it only shows when there is no tax to offset it) and `row.cTaxCap` = Capital tax (the yearly yield taxes `yieldTaxK + yieldTaxP` on after-tax
  and pension accounts). Municipal + state + fee + pension fee add up to `salaryTax + totalTax` (the pension fee is not displayed in the chart or table, so the shown categories can fall slightly short in years it is not offset); capital tax comes on top (it is deducted from the account returns, not from income)
- Pension income tax (`pensionTax`, `xAge`, projection ~"Swedish income tax") = Typfallsmodellen (VBA_go tax block, Skatteregler `avdrag26`, `statlig`,
  `PublicAvg`): floor100(income) − grundavdrag (`grundavdragUnder`; `grundavdragOver` from `xAge(year)` = 66 to 2023, 67 to 2026, 68 to 2028, then target
  age + 1) = taxable income; kommunalskatt (`munTax` input) + begravningsavgift `BEGRAV_AVG` 0,26 % on it, state tax (`stateTax` input) above `SKIKT` 643 000,
  and `PublicAvg` 1 % up to 1,42 IBB. `row.munTax` = kommunal + begravning, `row.stateTax` = statlig + PublicAvg. No jobbskatteavdrag on pensions. The Parameters-tab
  allowance amounts and state thresholds are no longer used for pension tax (thresholds only draw the chart line). Checked against an independent
  re-implementation at ages 66–95 (ITP1/1980: age 68 ≈ 112 700 kr = 30,6 %, from 69 ≈ 70 900 kr = 19,4 %; after the 2026 Table 2 alignment of the income pension 111 901 / 71 598). Verify yearly: `SKIKT`, `PBB`/`IBB`, `BEGRAV_AVG`
  (K_skatt), `PUBLIC_AVG_CAP`, the `xAge` schedule. Pension drawn while working again is taxed together with the salary (see Work after retirement); `taxAdjPort` and the
  "After tax" display modes use flat tax-rate inputs; the workbook's 32,38 % municipal default is a user input here (32,4 %)
- Salary tax before retirement uses the same age rule as the pension tax: `_ageMode = age >= xAge(yr) ? 'over' : 'under'` (workbook `Xage` drives both grundavdrag and
  jobbskatteavdrag; target age + 1 from 2029, so 69 for born 1980). The old fixed "67" made the tax dip at 67. ITP1/1980/40 000: salary tax 102 687 at 65–67 (total 109 812 at 67), retirement year 68 =
  pension tax 111 901 (grundavdrag only ~27 200, no jobbskatteavdrag) + capital tax 7 253 = 119 154, age 69 = pension tax 71 598 after förhöjt grundavdrag (~152 000) = 78 590 total.
  Projection table has a "Salary Tax" column (`y.salaryTax`, also in the export) because the Municipal/State columns are pension tax only (0 before retirement)
  and, right after it, the same four tax-chart categories for ALL years (`cTaxMun/State/Fee/Cap`; the 7 % pension fee `cTaxPens` is computed but not shown anywhere) plus "Total Tax" (their sum); the older pension-only Municipal/State columns in the Pre-Tax group were removed from the table (the export still has them as "Municipal Tax" / "State Tax"). Export also has the new columns
- Work after retirement (Personal tab "Work after retirement": `workAgainSal` 2025 money × `salBaseUplift`, `workAgainFrom` / `workAgainTo` ages inclusive; `workPeriod()` = {fromY, toY, sal} or null,
  invalid = ignored with a red message; `workInPeriod`, `ctxWorking(ctx, y)` = before the first retirement OR inside the period). Typfallsmodellen has an own wage vector by age and taxes wages and pensions together.
  Engine: `occMakeCtx` carries `work`; `occWage` returns the second salary in the period; `ipRatt` / `ppRatt` / `rightsBase` gate on `ctxWorking`; projection `_workingYr` gates contractual premiums and salary-based
  contributions; `occPremiumFor` does not stop premiums at the pension start age while working again (the scheme's own age limits still apply: ITP 1 66, PA16 Avd 1 67, PA16 Avd 2 69, AKAP-KR none);
  contractual rows' Contrib End (`ceBaseYear`) = max(retirement year, period end). `syncWorkRows()` (called from `recalc` and `workAgainChanged`) creates / updates / removes the two managed state accounts
  `workRow` "Income pension 67-70y" / "Premium pension 67-70y" (`accDisplayName`, no delete button, cs/ce = period, ws = first year + 1, `recalcYearly` on the income pension: each year's rights add
  `contrib / divisor(age)` to the annuity). Main IP / PP rows keep ce = retirement, so nothing is counted twice. Withdrawals tab: main Income / Premium pension rows get a second "Share while working" selector
  (`wdShareWork`, `wdShareAt(a, yr, wp)`, `setWorkShare`); premium pension = balance ÷ divisor × share (held-back part stays in the balance), income pension keeps the full annuity and `statState[i].mult`
  releases the held-back part when the share rises: `Δshare × dt0/dt(age) × 1,016^years`. Tax in years with salary AND pension: `_comb = incComputeParams(salary, 0, pension, …, BEGRAV_AVG)` (one grundavdrag, jobbskatteavdrag on the salary);
  `totalTax` stays pension-alone, `salaryTax = combined − pension-alone` so Salary Tax + pension tax = combined tax, `cTax*` categories come from the combined result. Dashboard "Income Sources" and Income vs Expenses show the
  salary whenever `earnedIncomePre > 0`. Not modelled: Garantipension/IPT underlag and ATP points ignore the extra rights; PGB unchanged; the Withdrawals-tab summary estimates use the plain `wdShare`;
  replacement rate = last 5 years before the FIRST retirement. ITP1/1980, retire 65, 20 000 kr/month at 67–70: salary tax 51 178 at 67, 33 102 at 69 (förhöjt grundavdrag from 69)
- Dashboard top KPI row: Saving Goal (Target Portfolio) = `savingGoal` (saving multiplier × annual expenses, sub shows the multiplier), Portfolio Value After Tax at
  Retirement = `taxAdjPort`, Saving Goal Fulfilment = the second ÷ the first, Replacement Rate. The earlier Max Withdrawal Rate / Portfolio Runs Out cards (and their +/- rate
  control, `applyWdrlRate`, the startup 4 % call and `maxSustainableRate`) were removed: per-account Wdrl amounts were overridden by the expense intervals, so the cards did not reflect the plan
- Header: the plan dropdown sits to the left, right after the subtitle. The subtitle (`#hdr-sub`, `updateHeaderSub()` from `recalc`) reads "Constant Prices · Swedish Tax Rules"
  and turns into a red "Inflation x% · Swedish Tax Rules" when the Parameters-tab inflation is not 0. KPI colours: `.tile-value.good` green, `.warn` yellow (`--amber`; `--gold` is the
  brand green in this file), `.bad` red. Saving Goal Fulfilment: green ≥ 100 %, else yellow. Replacement Rate: green ≥ 70 %, yellow 50–70 %, red < 50 %
- KPI tile frame and top bar follow the value colour: `renderMetricTiles` adds `st-good` / `st-warn` / `st-bad` from the tile's `cls` (CSS `.metric-tile.st-*`, green / yellow / red); tiles without `cls` keep the default
- Dashboard monthly-pension cards (`buildPhaseTiles`, `computePensionPhases`, phases = `wdpPhaseBounds`): green when the AFTER-TAX pension (`afterTaxIncome`, the model's own tax) covers the expenses of the same years, red when it does not, whichever
  Before/After tax view the card shows; the sub line lists both figures. Withdrawals chart (`wdChartData`, `buildWithdrawalChart`): Before/After tax buttons in the chart drive the same `_wdTaxMode` as the
  top toggle, default Monthly + Before tax, kronor axis (`fmtKrTick`), CSV / Excel export of exactly the plotted bars (`_wdChartRows`, `;` CSV with BOM, `XLSX`); in After tax view a dashed
  "Spending" line (`row.expenses`, retirement years) is drawn on top (not exported, not in the tooltip total)
- After-tax views all use the model's own pension tax (`_netFactorFor(y)` = 1 − `totalTax` / `taxable` of the year; `_ilsTaxFactor(y)`, `wdTaxFactor(a, y)`): Income vs Expenses,
  Income Sources Over Lifetime (After Tax), the Withdrawals chart/overview/table/exports and the monthly pension cards agree to the krona, and the Parameters-tab municipal / state tax
  drives them. The old flat "Tax rate" inputs (32 %) were removed (a hint "Tax per Parameters tab" replaces them). `taxAdjPort` (capital estimate) still uses the flat municipal rate. Income vs
  Expenses tooltip shows full amounts ("24 555 kr/mth")
- Projection tab table: first two header rows and the Year / Age columns are frozen (`.yr-table-wrap` scrolls both ways, sticky cells with opaque backgrounds, `--yr-h1` = measured height of header row 1 via
  `updateYrStickyOffset`). Portfolio Value tab, "All Accounts": the label column is frozen; the Withdrawal row is split into "Withdr aft tax acc" and "Withdr bef tax acc" (`pvWdrlSplit`: after-tax accounts vs P-insurance + state);
  in Tax adjusted mode the bef row is shown after the model's tax (`_netFactorFor(y)`), and the chart's withdrawal series uses the same split. Contribution is split the same way
  ("Contrib. aft tax acc" / "Contrib. bef tax acc", `pvContribSplit`): in Tax adjusted mode the pre-tax accounts' contributions are valued after tax like their balances (flat municipal rate `taxFactor`).
  The per-account tables ("Show all accounts" and single account) label those two rows by account type and apply the same rules (`pvAccFlow`). Balance rows still use the flat municipal rate
- Portfolio Value tab: the label column is frozen in EVERY table (summary and per-account tables with "Show all accounts"); the retirement-year column is tinted (`#fbf6e8`, `pvRet` / `retBg`) and its Year header
  carries "★ Ret", as in the Projection table. Every table is scrolled on build so the retirement-year column sits in the middle of the visible area (`pvCenterRetirement`, header cell class `pv-ret`; the frozen label column is excluded). Charts mark the retirement year on the x axis with the age (Withdrawals year mode: the year) in BOLD GOLD at the regular tick size, no star: `retStarPlugin` (global, `beforeInit`: for charts with `options._retMark` = one `{idx, text}` or an array, from `retMark(years, mode)`, it wraps the x axis `ticks.color` / `ticks.font` (scriptable) so Chart.js itself draws that tick bold gold with the same size, rotation and position as the others, and marks the tick `major` so autoSkip keeps it; no custom canvas drawing). Tick callbacks just use `retTick(years, idx)`. Wired: the six Dashboard charts, the Withdrawals chart, the Portfolio Value charts, the Scenarios chart (one mark per scenario retirement age) and the Monte Carlo fan and survival charts (start age = index 0; the histogram has value bins, not ages)
- Expenses tab persona buttons Basic / Average / Extra (Bas / Medel / Plus) = `EXPENSE_PERSONAS`,
  minPension's typical retiree budgets (17 600 / 18 400 / 27 300 kr/month,
  https://www.minpension.se/allt-om-pensioner/ta-ut-pension/hur-mycket-pengar-behover-en-pensionar),
  verify yearly. `EXPENSES_INIT` = Average
- "New retirement plan" = `planClear()` + `guideStart()`: floating panel (`GUIDE_STEPS`,
  `guide.*` keys) Personal → Assets → Withdrawals → Expenses, highlighting key fields
  (`.guide-hl`). `goTab(id)` switches tabs from code (`switchTab` no longer needs a click event)
  ⚠ The guide starts from an EMPTY plan, so every path it puts the user on must be able to ADD an
  account. Two ways that broke, both fixed Nov 2026: `guideShow` set `_acsOpen = 0` on the Assets
  step ("all five step headers visible"), but a step's BODY is what renders `acsAddPanel`, so the
  tab showed five headers and no "+ Add" button at all — it now opens step 1. And the Withdrawals
  empty states (`wdp.s1.none`/`s2`/`s3`/`s4.empty`) named Assets as the place to add without a way
  to get there; they now render through `wdpNoneBody(key, step)`, whose button calls `wdpGoAdd(step)`
  (goTab + `acsAddOpen`, which must run LAST since it sets `_acsOpen`; the guide is moved to its
  Assets step with `guideRender`, never `guideShow`, which would reset `_acsOpen`). The button reuses
  the existing `acs.add.{n}` labels, so it needs no new i18n keys. `.guide-hl` is outline-only and the
  panel has no overlay — the guide never blocks clicks, so a report of "I can't press anything" means
  the control is not being rendered.

### Temporary payout ("Temporärt", Withdrawals tab) — Retirement_Planner.html

- Withdrawals steps 2–3 offer Lifelong | 5 | 10 | 15 | 20 years | Other (`wdpSetYears`: n = 0 lifelong, else Temporärt with `we = ws + n − 1`; the internal mode and `isPlannedPIns` keep the old
  "planned" name; `setPInsPayout` is kept for older callers).
- A Temporärt account with a window (`isTempAcc`: planned, `ws` < 9999, `we` < 9998, `we` > `ws`) pays a CALCULATED yearly amount, not a typed one. Start value = balance at `ws` ÷
  `tempPayoutDivisor` = Typfallsmodellen's temporary occupational withdrawal (`tjp_ddeltal`, `rng_Temp_Tjp_Uttag`, N = we − ws + 1 < 30: Σ (1 − m(x+i)) / (1 + r)^i ÷ (1 − m(x)), Makeham m(x) = 0,0002 +
  0,000007·e^(0,1071x), r = the agreement's own rate `tempRate` = `schemeBasis`; checked against the engine for 36 combinations). `zeroTemporaryAccounts()` then bisects `a.wd` with `runProjection()`
  (all accounts in the same runs) so the balance at Wdrl End is 0 (± 1 kr). The step's option chips show Typfallsmodellen's amount per period (balance at `ws` ÷ divisor, before
  the adjustment, `wdpPinsOptions`); the line under them shows the projection's first-year payout. Runs from `recalc` (`scheduleTempZero`, debounced, only when `tempSignature()` — windows,
  balance at `ws`, returns, savings mode — changed) and when the payout is switched.
- Savings in step 4 mode "The same amount every year" (`savingsMode 'even'`, `isTempTaxed`): every after-tax account gets the window evenFrom … evenTo (`wdpApplyEven`, from `recalc`;
  default retirement … life age) and is adjusted the same way, so each ends at 0. In the other modes after-tax accounts are not adjusted ('manual' keeps typed amounts — before the redesign
  they were also zeroed whenever a Temporärt account existed; saved plans carry the zeroed amounts, so they load with the same numbers).

### Assets tab (the plan in five steps) — Retirement_Planner.html

`#tab-accounts`: the asset-overview pies (`buildAssetSummary`, unchanged, three slices by `a.type`), a
totals strip (`dst-total-opening` / `dst-total-contrib`), then five numbered steps in the Withdrawals
tab's shell — `buildAccountsTable()` keeps its name (~20 call sites) but renders `#acs-steps`.
Until Nov 2026 the tab was one 11-column table (`accounts-tbody`) showing every expert field on every row.
1 State pension (`a.type==='state'`, work rows included, read-only) · 2 Occupational (`p-ins && wdpIsOcc`)
· 3 Private pension (`p-ins && !wdpIsOcc`) · 4 Savings (`after`) · 5 Other pension rights (PGB,
`buildPgbSection` renders into step 5's body, so it only exists while that step is open — `recalc` already
guards on `#pgb-section`).
- One step open at a time (`_acsOpen`, `acsOpen`/`acsNext`); closed steps show `acsSumText(n)` (count ·
  value · contribution), refreshed IN PLACE by `updateAccountSubtotals` so typing never costs a re-render.
- A card (`acsCard`) shows at most three fields: savings get an ISK/KF-vs-bank chip (the yield tax is the
  only thing separating them, `acsSavKind`/`acsApplySavKind`), then opening, then saving per month
  (`acsSetMonthly`, stored ×12 in `a.contrib`). The read-only rule blocks are the SAME renderers as before
  (`renderPInsContrib`, `renderOccDbCell`, `renderBenefitCell`), reflowed from `<td>` to `<div class="acs-cell">`.
  Everything else — contribution window, equity share, Mgmt %, Yield Tax %, Fixed Ret %, and the account's
  NAME (labelled, with help and an example) — sits in a per-card `<details>`. `occLockedCells`' six dash
  cells became `acsLockedNote(a)`, one line with the same two numbers.
- Which cards have Details open lives in `window._acsDetails`, a **WeakSet keyed by the account object**:
  an index would point at the wrong card after `removeAccount` splices, and a WeakSet never reaches
  `JSON.stringify(accounts)`.
- Adding an account (`acsAddPanel`/`acsAddGo`): radio cards for what it is (`STATE_KINDS`, `OCC_SCHEME_ORDER`
  plus `individual-occ`, ISK vs bank) with a line each on who it is for, then a **pre-filled, labelled Name
  field** with an example —
  the old table's unlabelled name box was the thing people could not work out. `addAccount` / `addOccScheme` /
  `addStateAccount` keep their signatures; `acsAddGo` identity-diffs `accounts` before/after, because
  `addOccScheme` adds TWO rows and all three splice after the last row of the same type. An untouched name
  keeps its `nameKey` and follows the language switch. Step 2's `individual-occ` is an occupational
  pension outside the collective agreements: created by `addAccount('p-ins', null)`, so it has exactly
  a private pension's fields, then flagged `occGroup: true`, which `wdpIsOcc` honours so it groups with
  the occupational pensions on BOTH tabs. The flag rides along in the saved plan; its absence is the
  old behaviour, so no save-version key.
- Field labels are written out in full, no abbreviations (owner's rule, Nov 2026): `acc.col.opening`
  "Capital today" / "Kapital i dag" (NOT "Opening"; "Portfolio value" was rejected because that is the
  name of a whole tab showing PROJECTED values, while this field is the balance today),
  `acc.opening.after.tax`, `acc.col.cs` / `acc.col.ce` "Contribution start/end (age)",
  `acc.col.mgmt` "Management fee %", `acc.col.yt` "Yield tax %", `acc.col.fr` "Fixed return %",
  `acc.eq.brackets` / `acc.eq.prefix` "Equity share…", `acs.tot.contrib` / `acs.tbl.contrib`
  "Contributed per year", `wdp.peryear` "kr per year" (also shows on the Withdrawals tab), `acs.sum`
  (closed-step summary), `occ.db.from`, `pgb.f.terms` / `pgb.f.grant` / `pgb.f.amount`, and the EN
  `occ.own.eq.tip` / `pp.glide.tip` tooltip texts, which named the field as "Mgmt %".
  ⚠ When checking this rule, scan the WHOLE of `#tab-accounts` — every text node plus every `title`
  and `placeholder` — with all five steps opened, Details open on every card and one PGB item of each
  kind (`child`, `vpl`, `studies`, `sa`). Scanning only the label elements misses the tooltips, the
  closed-step summary line, the defined-benefit cell and the PGB fields, which is exactly how five
  hardcoded `' kr/yr'` strings in the contribution-rule renderers survived a "0 hits" check: they were
  never in the i18n table at all, so Swedish showed "kr/yr" too. They now call `t('wdp.peryear')`.
  (In a JS regex `\b` is ASCII-only, so a pattern like `/\bst\b/` matches the "st" in "i stället" —
  do not trust ad-hoc abbreviation patterns against Swedish without checking the hits.) The SV
  `acc.col.cs/ce` moved from "Bidrag" to "Insättning", matching `acs.f.contrib`. The `acc.col.*` keys
  are now used ONLY by this tab (`acsCard`, `acsSummaryTable`), so they are safe to reword; `acc.col.contrib.yr`
  and `refreshEqHeaders`' `th-eq1/2/3` lookups are dead leftovers of the old 11-column table.
- Every input keeps `data-ai`/`data-field` + `updateAccountField`, and is `onchange`, never `oninput` — the
  panel is rebuilt with `innerHTML`. i18n namespace `acs.*` (EN + SV); `_SAVE_SKIP` carries `/^acs-/`.
  `GUIDE_STEPS`' Assets step highlights `#acs-step-1/2/4/5` and `guideShow` sets `_acsOpen = 0`.
- Under the steps, `acsSummaryTable()` (`#acs-table`) is the read-only at-a-glance view the old table
  used to give: item · value today · in per year, grouped by the five steps and totalled, PGB included
  (`pgbResultHtml(ctx, i).tot` as the value, a dash for the contribution). Value today, NOT a projection,
  so it agrees with the totals strip and the pies; it follows the Before/After tax toggle via
  `acsDispOpening` and renames its header. Built from `buildAccountsTable()` AND from
  `updateAccountSubtotals()` (it holds no inputs, so rebuilding it cannot steal focus). Empty plan →
  renders nothing. Keys `acs.tbl.*`.
- Guided walkthrough: `guide.b2` must describe these five steps — it is the one place outside the tab
  that explains it, and it was left describing the old After-tax / P-Insurance / State / PGB groups for
  one release. The panel is draggable by its `.g-grip` (step + title); the position lives in `_guidePos`
  and is re-applied by `guideRender` because the panel is rebuilt with `innerHTML` on every step, is
  clamped to the viewport on move and on resize, and resets on close. Dragging is armed only above
  767 px. The minimise button (`guideToggleMin`, `.g-collapsed`) works at every width — its rules used
  to sit inside the mobile media query, so on desktop the button was `display:none`.

### ISK / kapitalförsäkring tax-free allowance (fribelopp) — Retirement_Planner.html

`ISK_FRI` = 300 000 kr (constants block, shared with the Income tab). Schablonskatt is charged only on the
part of the COMBINED balance of after-tax accounts **with `yt > 0`** above it — a bank account neither pays
the tax nor consumes the allowance. Implemented in `runProjection`'s after-tax pre-pass as a per-year factor
`_iskF = max(0, base − ISK_FRI) / base`, applied to each account's rate into `atYtEff[i]` and read at five
places: `netRet2`, the `monthlyFV` of this year's deposits, the `yieldTaxK` accumulator and the two
`accDetail` sites (the ordinary one and the deferred `need` settlement). A factor on the RATE, not a sum of
kronor, because the tax sits inside each account's own net return and the factor stays right if one account
is given a different `yt`. The base uses the unscaled rate (the dependence of this year's deposits on the
rate the base sets is third-order). Nothing is rounded — Skatteverket floors the schablonintäkt and the
Income tab does too, but the projection floors nothing anywhere. Held flat in today's money, like every
other threshold in this constant-prices model. Saves carry `iskFriV: 1`; an older save gets a one-off toast
(`toast.iskfri`) because the numbers move. Monte Carlo has its own engine and is unaffected.
Planner-specific (Typfallsmodellen does not model private saving at all) — named in the About tab as
`about.own.8`. Verify every December: `ISK_FRI`, and `SLR_DEFAULT` below.

### Fee and yield-tax assumptions — Retirement_Planner.html

Both Swedish yield taxes are built on the statslåneränta. `SLR_DEFAULT` = 2,5 % (Typfallsmodellen's own
forward assumption, the `Case Else` in `avkskatt` and the fallback inside `occTax`), editable as `#slr` on the
Parameters tab, and the two rate fields below it are **read-only displays** derived from it:
`ytKDefault()` = (SLR + 1 pp) × 30 % = 1,05 % (ISK/KF) and `ytPDefault()` = SLR × 15 % = 0,375 % (pension
insurance), both written by `refreshYieldTaxes()` from `recalc`. New accounts take their `yt` from these;
the engine still reads each account's own `a.yt`, which Details can override. `#ytP`/`#ytK` are in
`_SAVE_SKIP` so a restored plan recomputes them from the saved SLR. The old `params.ytP` / `params.ytK`
were collected and read by nothing and are gone. `occTax` is a workbook formula and is NOT driven by `#slr`.
Note the deliberate difference: the Income tab uses the ACTUAL 2026 schablonränta (`ISK_SCHABLON` = 3,55 %,
SLR 2,55 % + 1 pp → 1,065 %), the projection the 2,5 % long-run assumption → 1,05 %.

### Withdrawals tab (step-by-step plan) — Retirement_Planner.html

The tab is a guided plan: five numbered steps in an accordion (`wdpRender`, one open at a time, `_wdpOpen`; closed steps show a one-line summary;
"Next →"), then a Result card (`wdpRenderResult`), the "Withdrawals by account" chart, the overview pies and a folded, read-only "Details per account"
table (`buildWdTable`, every figure read from the projection: balance at the first payout year, first-year payout, first / last payout age or ∞, left at the end,
average payout per spending phase, in the Before / After tax mode). Every choice writes the accounts' own fields or `wdPlan`; the projection does the rest.
1. State pension start age (`wdpSetStateAge`): chips riktålder − 3 … riktålder + 3 plus "At my retirement age" (= `wsAuto`) and an other-age input; each chip shows the
   first-year state pension per month before tax from one `runProjection()` per age (`wdpStateOptions`, cached by `wdpSig`; adds `row.lastRight1` as Table 1). Applies to every
   non-work-row state account; Garantipension / IPT never before the target age. "More choices" (`wdpStateMore`): start per pension, withdrawal share, share while working
   again, premium pension fund / traditional insurance, custom accounts' divisor or manual amount.
2. Occupational pension (scheme rows, the occupational and salary-exchange rows: `wdpIsOcc`) and 3. private pension (the other P-insurance rows): from age (chip "At retirement"
   = `wsAuto`, input, earliest age from `WD_RULES`) and how long — see Temporary payout below. DB rows: start age only. Link to the Withdrawal strategy tab.
4. Savings (`wdPlan.savingsMode`): `need` (default) = cover what the pensions do not: `runProjection` settles the after-tax accounts in a deferred pass after the year's
   pension tax and earned-income block — draw = max(0, spending − (pension − pension tax) − net salary while working again), ISK / endowment first, then cash, capped by the
   balance, from the retirement year; `portAfterTax`, `totalPort`, `eqPct`, `taxFree`, `afterTaxIncome` are finalised after it; a surplus is not saved. `even` = same amount
   every year (Temporary payout below). `none` = no draws. `manual` = each account's own amount and window (the old behaviour, typed in the step).
5. Spending: `spendAt(age, budget)` = Expenses-tab budget × level (`flatLevel`, or with "More early, less later" `levels[0..2]` of the three phases `wdpPhaseBounds`: retirement
   … life age split in thirds, or the user's boundary ages `splitAges` while `splitAuto` is false). Levels as % of the budget (stepper ± 5 %, input 30–300 %), quick picks Even
   100/100/100, Moderate 110/100/90, Active start 120/100/80. `row.expenses` (Dashboard, Income vs Expenses, Monte Carlo default withdrawal `getWdForAge`) is `spendAt`.
- Result: per phase the average income after tax (`totalAfterTaxIncome`: pensions, savings draws, a salary while working again) vs spending per month, green when the income
  covers the spending in EVERY year of the phase, else "Short from age X". Then the highest sustainable level (`wdpMaxSpend`, replaces Maximize / Estimate All / Optimize to Zero):
  bisection on a scale k of the whole spending profile (`_wdpScale`, k ≤ 8) for income ≥ spending − 1 kr in every retirement year up to the life year; computed from scratch
  (`wdpScheduleMax`, debounced, only while the tab is open and `wdpSig` changed). "Use this level" (`wdpUseMax`) floors level × k to whole % and is offered only when k is
  outside [1, 1,02), so pressing it again changes nothing. Without savings the sentence says what the pensions alone allow (the budget is never set to 0).
- Planner-specific (Typfallsmodellen has no spending phases, no expense-driven drawdown and no max-spending search; its private saving is spread over the payout period, which
  is the `even` mode). All pension amounts, divisors and the tax are Typfallsmodellen's.
- Saves carry `wdPlan` + `wdPlanV: 1` (`wdpRestore`). Older saves: "Age-based expense intervals" on with amounts → phases with the old boundaries and levels = amount ÷ budget
  (rounded to 5 %, 50–150 %), savings mode `need`, toast `wdp.migrated`; intervals off → `manual` (same numbers as before). New plans and presets: `need`, phases off.

### Withdrawal strategy tab ("Uttagsstrategi") — Retirement_Planner.html

A what-if over the plan, built in the planner (`wsRender`, `ws*` functions, texts inline via `wsSay`), on the planner's own engine
(`runProjection`, the Typfallsmodellen port) and read with the rules of the Typfallsmodellen web app
(github.com/mathiasboos/Typfallsmodellen, `apps/web/src/pensionLasts.ts` and `strategy.ts`; latest checked 6d13a3c). Until Nov 2026 the tab was
that app's own panel compiled into the file (`tools/build-strategy.mjs`, ~1 MB); it was replaced because it had its own inputs, one scheme and one
payout time for all pots. Nothing in the tab is saved with the plan.
- Inputs are the plan's, read-only: Person (folded: Personal / Parameters), the pension and savings accounts from Assets, the spending = Expenses
  budget × the Withdrawals spending phases (`row.expenses`). No budget field, no assumptions (Parameters tab). Ages run from the retirement age up to
  and including the life-expectancy age.
- The sandbox `_ws.choice` = {pins: {account index: {ws, n}}, sav: {mode, n, from}} (n = years, 0 lifelong, null = as in the plan), cloned from the plan
  (`wsChoiceFromPlan`) on the first visit, after "Reset to my plan", and when the plan changed while nothing was edited (`_ws.planKey`). Tables:
  pensions (state rows read-only and lifelong; occupational / private P-insurance: from age + Lifelong / 5 / 10 / 15 / 20 / other years, the agreement's
  minimum from `WD_RULES` enforced) and private savings (one use for all: cover the gap, the same amount for N years or to the life age, untouched).
  "Use in my plan" (`wsUseInPlan`) writes it like `wdpSetYears` + `wdPlan.savingsMode`/`evenFrom`/`evenTo`; afterwards the plan's projection equals the
  what-if run (checked: identical yearly income).
- `wsRun(choice)`: snapshot (`wsSnap`) → `wsApply` (windows, then the Temporärt / even amounts zero-adjusted once per account + window in `_ws.zc`,
  `wsZero` = the bisection of `zeroTemporaryAccounts` from Typfallsmodellen's divisor; valid because an account's balance path depends only on its own
  amount) → `runProjection` → `wsSummary` rows (`net` = totalAfterTaxIncome, `gross`, `need`, per kind wdrl `k`, `nf` = `_netFactorFor`) → restore. The
  plan is never changed by it. `wsCompute` (debounced from `recalc` / tab switch while the tab is open, signature `wdpSig` + choice) runs the 5 uniform
  periods, every combination occupational × private × savings over 5/10/15/20/life (≤ 125 runs), the what-if and the plan: ~30–90 ms.
- Best payout time (KPI boxes, pensionLasts `assess` / `evaluate`): every pot over the same period (shorter boxes use an agreement's longer minimum,
  noted); per box the average income a month during the payout, "then X" afterwards, and how long the spending is covered (✓ all the way / lasts to N /
  x of y years); BEST = most years with income ≥ spending − 6 kr, then the most in total, a tie (± 1 kr) to the longer. Coverage is always after tax;
  the Before / After tax switch changes the amounts shown.
- Strategic scenarios (`wsScenario`): Balanced (all lifelong, savings spread to the life age), Your plan (`_projYears`), Optimal (`wsOptimal` =
  `optimalCombo`: most years at/above the spending, then most in total, tie → more lifelong pots, then longer periods), Your what-if when it differs.
  Per spending phase average, average, totals through 80 / 85 / 90 / life age, lowest month, "Balanced pays more from" (`wsOvertaken`), line chart,
  "Try this". Payout-period table: occupational over 5/10/15/20/life with the rest lifelong. Timeline (`wsTimelineSvg`): Typfallsmodellen's Gantt incl.
  the 6d13a3c label fix, periods = the spending phases. CSV / Excel of both tables.
- Insights (`wsInsights`: payout, first year taxed harder, ISK tax, state tax from the what-if, savings early), tax rules (`wsTaxPicture`,
  `WS_TAX_YEARS`) and the checklist (`wsChecklist`, `WS_JAMKNING`) follow strategy.ts. Not ported: housing supplement (bostadstillägg, not modelled) and
  its insight / checklist step; the ISK insight uses the planner's yield tax, now net of the 300 000 kr fribelopp. Verify every December: `WS_TAX_YEARS` (skattefri årsinkomst, ROT/RUT
  maxima) and `WS_JAMKNING`. When the Typfallsmodellen app changes these tabs' rules or texts, port the change by hand.

### Pensionskompassen_offline.html — Typfallsmodellen is the master

All capital and payout figures in the "Uppskattat belopp" mode come from the Typfallsmodellen engine copied (unchanged)
from Retirement_Planner.html into the block `// ═══ Typfallsmodellen-motor` (data tables + `occStep`, `occPremium`,
`ipStep`/`ipRatt`, `ppStep`/`ppRatt`, `ipDelningstal`, `occDelningstal`, `riktalderOf` …). Pensionskompassen's own part is
`pkIncomePension` / `pkPremiePension` / `pkOccupational` / `computeCapital` (scheme picker `#scheme-select`, `PK_SCHEMES`:
ITP 1, SAF-LO, AKAP-KR, PA16 Avd 1 — the pure premium plans; ITP 2, KAP-KL and PA16 Avd 2 have a defined-benefit part that
only Retirement_Planner models). Constant prices (inflation 0 %, salary growth 0 %, 1,7 % real return), salary entered in
2025 money × `salBaseUplift`, work start age `WORK_START_AGE` 23. Income pension is indexed by the pension index ÷ 1,016
each year and premium pension follows the cohort divisor, so the public pension varies by age (`cap.apAt(age)`). Riktålder =
`riktalderOf`, earliest withdrawal = riktålder − 3. Verified against Retirement_Planner (ITP1 / born 1980 / 40 000, start
68): opening balances and payouts at 68, 69, 80 and 90 match to the krona. When the planner's engine tables change, re-copy
the block (the extraction script is not kept; copy the same declarations by name). Manual capital mode annuitises the typed
public capital at the cohort income-pension divisor, flat. The occupational balance keeps earning the 1,7 % return (minus yield tax, `pkTpStep`) while it is drawn down;
"Jämnt fördelat" solves the yearly amount that empties it at 95. Inheritance gains (`occArvGain`) are NOT applied to the
drawn-down balance (it is the person's own capital, not a pooled lifelong annuity); the planner applies them.

Tax card (Före/Efter skatt, `#tax-card`): `pkTaxYear` = the planner's `pensionTax` (grundavdrag by age, kommunalskatt,
begravnings-/kyrkoavgift, statlig skatt over SKIKT, 1 % PublicAvg); the municipality list `KOMMUNALSKATT_2026` and the
church/burial options match Typfallsmodellen. After tax, entered and shown amounts are net (`pkGrossMonthly` /
`pkNetMonthly`, passed to `runSimulation`; the payout tab converts at the payout start age). Verify PBB/IBB/SKIKT and the
municipality list every December.

## Source of truth for formulas

**Typfallsmodellen (`reference/Typfallsmodellen_v4.8.xlsb` + its VBA) is the first source to consult and to implement in Retirement_Planner.html** — pension rules, contributions, tax and rounding follow the workbook exactly; a planner-specific approximation is only acceptable where the workbook has no equivalent, and must be documented as such. Explanatory UI text may round or simplify (e.g. "up to 8,07 IBB"), the calculation must not.

**Credit Pensionsmyndigheten.** Pensionsmyndigheten's licence for Typfallsmodellen: "Du får sprida och modifiera verktyget men du ska ange Pensionsmyndigheten som källa."
(https://www.pensionsmyndigheten.se/statistik-och-rapporter/pensionsmodellen/typfallsmodellen). Retirement_Planner.html therefore has a source line under every tab (`<footer id="src-credit">`, keys `foot.*`) and an
"About" tab (`#tab-about`, keys `about.*`, EN/SV) that names Pensionsmyndigheten as the source, lists what is taken from Typfallsmodellen, what is developed for the planner (charts, budget, Withdrawals planning,
Scenarios, Monte Carlo, Income tab, presets) and the other sources (Skatteverket, SCB, minpension.se). Keep those lists current when features are added or the workbook version changes, and never remove the credit.

When fixing or verifying calculation logic, **treat Excel as the source of truth**. If a formula in an HTML file diverges from the reference Excel model, the HTML is wrong.

## Publishing to pensionslyft.se

`.github/workflows/notify-pensionslyft.yml`: a push to `main` that changes `Retirement_Planner.html` starts the
Pensionsplaneraren sync of github.com/mathiasboos/Pensionslyft right away (workflow_dispatch of its
`pensionsplaneraren.yml`, which fetches, checks and publishes the file). Needs the repository secret
`PENSIONSLYFT_DISPATCH_TOKEN` (fine-grained token, Pensionslyft only, Actions: Read and write); without it the step
skips with a notice and the change waits for the site's own schedule. That schedule is `2,17,32,47 * * * *`,
but GitHub runs schedules best-effort and skips most of them — observed runs are hours apart, not 15 minutes.
So while the secret is unset, do NOT rely on the schedule: after a merge that touches `Retirement_Planner.html`,
dispatch `pensionsplaneraren.yml` in Pensionslyft by hand and watch it finish.

## Git workflow

This folder is a git repository. Commit all changes with clear, descriptive messages. Prefer separate commits for separate concerns (e.g. don't bundle a bug fix with a visual tweak). Use `git revert` rather than rewriting history when undoing changes.
