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
  index, admin fee, inheritance gains `OCC_ARV_IP`), payout = balance incl. the
  last right ÷ cohort divisor (`DELNINGSTAL_IP`), then indexed with the pension
  index ÷ 1.016. Data to extend yearly: `OCC_IP` (pension index, admin fee, PBB,
  förhöjt PBB, lägsta pensionsålder by cohort) and `OCC_ARV_IP`. Forecast pension
  index = income index (no balancing brake), as in the workbook
- Premium pension (premiepension, scope `premie`, `isPpAutoAcc`): only the HISTORY
  uses the workbook (`ppRatt` = 2.5 % of PGI credited the year after, `ppStep` =
  VBA `PP_pbh`: fund yield, `arv PP` gains, RGK interest `PP_RGK`, manager fee
  `ppFee`/`PPMavg`). Used when Opening is blank/0 (`openAuto`); the projection keeps
  the planner's own growth so equity %, fees, yield tax and fixed return stay
  editable. Data to extend yearly: `PP_RGK` (Några tal C19). Agrees with the port
  within ~0.6 % (payout = balance(retY−1)·yield + last right ÷ cohort divisor)
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
  oracle (the HTML port has no PGB inputs). Not modelled: forward premium-pension contributions
  from PGB (planner's own salary-based contribution), KPI scaling of the first child's PGI
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
  at 0, expense intervals reset. Birth / retirement / life-expectancy years are kept
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
- Assets tab contribution window defaults to Personal: Contrib Start = begin-work age (schemes:
  max(work age, scheme from-age), `_csFor`), Contrib End = retirement age (`defaultCe`). Accounts
  still on the old default follow Personal-tab changes (`persRelinkAccounts`, trackers set by
  `persTrack`); edited ones stay. "End = retirement year" means no contribution in the retirement
  year itself (fixed contributions; salary/scheme/IP ones were already gated)
- Premium pension equity: default fund glidepath (`ppGlideEq`: 100 % equities to 55, linear to
  33 % at 75) via `calcEquityShare`; tick "Own equity brackets" (`a.ownEq`) to use eq1–eq3.
  Typfallsmodellen itself has one fund-return assumption, no glidepath
  Contractual scheme premium rows (`isOccPremiumAcc`) have the same tick box: unticked = the
  Parameters occupational return (1,7 % net); ticked = `engineStep` runs `occStep` with the account's
  own net return (equity % brackets × eq/bond returns − Mgmt %); premiums, arv gains, yield tax and
  history stay Typfallsmodellen's
  Default Mgmt % = AP7 Såfa's fee 0,04 % (`mgmt: 0.0004`, check ap7.se yearly); restore lifts the
  old 0,10 % default on non-`ownEq` premium pension accounts
- State pensions are always lifelong: `we = 9998` for every state account (defaults, new accounts,
  restore, Optimize), Withdrawals tab shows ∞ in Wdrl End instead of an input
- Premium pension payout is lifelong like Typfallsmodellen (`ppkassa` + `PP_pbh`): each year
  balance ÷ the cohort PP divisor (`occDelningstal`, 2 from age 101), and the projection adds the
  inheritance gains (`occArvGain`, sheet "arv PP") to the PP balance every year (`isPpAcc`). The
  payout is not capped by the balance. Shape matches the port (born 1968: 0.97 / 0.81 / 0.60 of the
  age-68 amount at 85 / 95 / 100; born 1980 stays level)
- Default assumptions = Typfallsmodellen's prognosis standard: equities and bonds real return
  1,9 % BEFORE fees (each account deducts its own Mgmt %; 1,7 % after the workbook's average 0,2 %
  fee), inflation 0 %, occupational schemes 1,7 % (already net of fees), salary
  increase 0 %. Salary convention as in Typfallsmodellen: the Personal-tab gross salary is 2025
  money, raised by the 2026 income index (`salBaseUplift` = OCC_HIST 228.08 / 220.23, +3,56 %:
  40 000 → 41 426 kr/month) wherever the engine reads it (`grossSal`); net salary is computed on the
  raised amount; hint `#sal-uplift-hint`. ITP1/1980/40 000 vs the port at 68: ITP 1 −0,9 %,
  PP −3,5 % (−2,7 % with fee 0), IP +1,5 %, net salary 32 752 vs 32 770. Oracle scripts must enter
  the 2025 salary (port wage ÷ salBaseUplift). The 1,7 % is AFTER fees; the workbook's unticked
  "efter fondavgifter" option deducts a general 0,2 % fee (1,9 % before fees)
- `resetAssumptions()` puts eqReturn / bdReturn / inflation / occReturn / salaryIncrease back to their
  HTML defaults (`defaultValue`); called by `planClear` (so every preset) and the Parameters-tab
  button "↺ Typfallsmodellen defaults". Saves carry `assumptionsV: 2`; restoring an older save moves
  untouched old defaults (eq 4,5/1,7, bonds 0,5/1,7, salary +1,0) to 1,9 / 1,9 / 0,0, typed values stay
- Dashboard Replacement Rate = Typfallsmodellen's kompensationsgrad: pension before tax in the first
  retirement year (wdrlState + wdrlPIns) ÷ final salary before tax (mean of the last 5 working
  years' earnedIncomePre). `repRateNet` (after-tax ratio) is still computed but not shown.
  ITP1/1980/40 000: 74,2 % vs the port's 74,7 %
- Withdrawal rules (Withdrawals tab): `WD_RULES` / `wdRule(a)` / `wdRuleErrors(a)` / `wdDefaultYear(a)` give each
  pension an earliest start age and (P-insurance, Planned mode only) a shortest withdrawal time; a broken rule shows
  a red `.wd-err` message under the input, the typed value is kept. Verify the ages yearly.
  ITP 1 55/5, ITP 2 55/5, SAF-LO 55/5, KAP-KL 61/10, AKAP-KR 62/10, PA16 Avd 2 61 (born ≤ 1965) or 63 /5,
  PA16 Avd 1 63/5, Individual pension account 55/5 (earliest age / min years). Premium + Income pension: earliest
  target age (riktålder) − 3; Garantipension + Inkomstpensionstillägg: target age
- Wdrl Start of every state and P-insurance account follows the Personal-tab retirement age, raised to its earliest
  allowed age (`wdDefaultYear`). Accounts carry `wsAuto`; `wdSyncStart()` re-applies it from `persApplyAges` /
  `onRetirementYearChange` / restore, and typing a Wdrl Start clears `wsAuto`. New schemes start lifelong (∞); the
  Individual pension account is lifelong by default and has the Planned/Lifelong toggle. Saves carry `wdRulesV: 2`;
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
  tooltip total = sum of the series. The tax chart uses the same five categories before and after retirement: `row.cTaxMun` (municipal tax after
  jobbskatteavdrag / förvärvsavdrag, incl. begravningsavgift), `row.cTaxState`, `row.cTaxFee` (public service fee: `ps` for salary years, `publicAvg` for pension years), `row.cTaxPens` (general 7 % pension
  fee net of its tax reduction, `pensAvg − pensRed`; it only shows when there is no tax to offset it) and `row.cTaxCap` = Capital tax (the yearly yield taxes `yieldTaxK + yieldTaxP` on after-tax
  and pension accounts). Municipal + state + fee + pension fee add up to `salaryTax + totalTax`; capital tax comes on top (it is deducted from the account returns, not from income)
- Pension income tax (`pensionTax`, `xAge`, projection ~"Swedish income tax") = Typfallsmodellen (VBA_go tax block, Skatteregler `avdrag26`, `statlig`,
  `PublicAvg`): floor100(income) − grundavdrag (`grundavdragUnder`; `grundavdragOver` from `xAge(year)` = 66 to 2023, 67 to 2026, 68 to 2028, then target
  age + 1) = taxable income; kommunalskatt (`munTax` input) + begravningsavgift `BEGRAV_AVG` 0,26 % on it, state tax (`stateTax` input) above `SKIKT` 643 000,
  and `PublicAvg` 1 % up to 1,42 IBB. `row.munTax` = kommunal + begravning, `row.stateTax` = statlig + PublicAvg. No jobbskatteavdrag on pensions. The Parameters-tab
  allowance amounts and state thresholds are no longer used for pension tax (thresholds only draw the chart line). Checked against an independent
  re-implementation at ages 66–95 (ITP1/1980: age 68 ≈ 112 700 kr = 30,6 %, from 69 ≈ 70 900 kr = 19,4 %). Verify yearly: `SKIKT`, `PBB`/`IBB`, `BEGRAV_AVG`
  (K_skatt), `PUBLIC_AVG_CAP`, the `xAge` schedule. Not covered: pension drawn while still working is taxed separately from the salary; `taxAdjPort` and the
  "After tax" display modes use flat tax-rate inputs; the workbook's 32,38 % municipal default is a user input here (32,4 %)
- Dashboard top KPI row: Saving Goal (Target Portfolio) = `savingGoal` (saving multiplier × annual expenses, sub shows the multiplier), Portfolio Value After Tax at
  Retirement = `taxAdjPort`, Saving Goal Fulfilment = the second ÷ the first, Replacement Rate. The earlier Max Withdrawal Rate / Portfolio Runs Out cards (and their +/- rate
  control, `applyWdrlRate`, the startup 4 % call and `maxSustainableRate`) were removed: per-account Wdrl amounts are overridden by the expense intervals, so the cards did not reflect the plan
- Header: the plan dropdown sits to the left, right after the subtitle. The subtitle (`#hdr-sub`, `updateHeaderSub()` from `recalc`) reads "Constant Prices · Swedish Tax Rules"
  and turns into a red "Inflation x% · Swedish Tax Rules" when the Parameters-tab inflation is not 0. KPI colours: `.tile-value.good` green, `.warn` yellow (`--amber`; `--gold` is the
  brand green in this file), `.bad` red. Saving Goal Fulfilment: green ≥ 100 %, else yellow. Replacement Rate: green ≥ 70 %, yellow 50–70 %, red < 50 %
- KPI tile frame and top bar follow the value colour: `renderMetricTiles` adds `st-good` / `st-warn` / `st-bad` from the tile's `cls` (CSS `.metric-tile.st-*`, green / yellow / red); tiles without `cls` keep the default
- Dashboard monthly-pension cards (`buildPhaseTiles`, `computePensionPhases`): green when the AFTER-TAX pension (`afterTaxIncome`, the model's own tax) covers the expenses of the same years, red when it does not, whichever
  Before/After tax view the card shows; the sub line lists both figures. Withdrawals chart (`wdChartData`, `buildWithdrawalChart`): Before/After tax buttons in the chart drive the same `_wdTaxMode` as the
  top toggle, default Monthly + Before tax, kronor axis (`fmtKrTick`), CSV / Excel export of exactly the plotted series (`_wdChartRows`, `;` CSV with BOM, `XLSX`)
- After-tax views all use the model's own pension tax (`_netFactorFor(y)` = 1 − `totalTax` / `taxable` of the year; `_ilsTaxFactor(y)`, `wdTaxFactor(a, y)`): Income vs Expenses,
  Income Sources Over Lifetime (After Tax), the Withdrawals chart/overview/table/exports and the monthly pension cards agree to the krona, and the Parameters-tab municipal / state tax
  drives them. The old flat "Tax rate" inputs (32 %) were removed (a hint "Tax per Parameters tab" replaces them). `taxAdjPort` (capital estimate) still uses the flat municipal rate. Income vs
  Expenses tooltip shows full amounts ("24 555 kr/mth")
- Projection tab table: first two header rows and the Year / Age columns are frozen (`.yr-table-wrap` scrolls both ways, sticky cells with opaque backgrounds, `--yr-h1` = measured height of header row 1 via
  `updateYrStickyOffset`). Portfolio Value tab, "All Accounts": the label column is frozen; the Withdrawal row is split into "Withdr aft tax acc" and "Withdr bef tax acc" (`pvWdrlSplit`: after-tax accounts vs P-insurance + state);
  in Tax adjusted mode the bef row is shown after the model's tax (`_netFactorFor(y)`), and the chart's withdrawal series uses the same split. Contribution is split the same way
  ("Contrib. aft tax acc" / "Contrib. bef tax acc", `pvContribSplit`): in Tax adjusted mode the pre-tax accounts' contributions are valued after tax like their balances (flat municipal rate `taxFactor`).
  The per-account tables ("Show all accounts" and single account) label those two rows by account type and apply the same rules (`pvAccFlow`). Balance rows still use the flat municipal rate
- Expenses tab persona buttons Basic / Average / Extra (Bas / Medel / Plus) = `EXPENSE_PERSONAS`,
  minPension's typical retiree budgets (17 600 / 18 400 / 27 300 kr/month,
  https://www.minpension.se/allt-om-pensioner/ta-ut-pension/hur-mycket-pengar-behover-en-pensionar),
  verify yearly. `EXPENSES_INIT` = Average
- "New retirement plan" = `planClear()` + `guideStart()`: floating panel (`GUIDE_STEPS`,
  `guide.*` keys) Personal → Assets → Withdrawals → Expenses, highlighting key fields
  (`.guide-hl`). `goTab(id)` switches tabs from code (`switchTab` no longer needs a click event)

## Source of truth for formulas

When fixing or verifying calculation logic, **treat Excel as the source of truth**. If a formula in an HTML file diverges from the reference Excel model, the HTML is wrong.

## Git workflow

This folder is a git repository. Commit all changes with clear, descriptive messages. Prefer separate commits for separate concerns (e.g. don't bundle a bug fix with a visual tweak). Use `git revert` rather than rewriting history when undoing changes.
