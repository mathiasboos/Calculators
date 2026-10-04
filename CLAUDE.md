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

## Source of truth for formulas

When fixing or verifying calculation logic, **treat Excel as the source of truth**. If a formula in an HTML file diverges from the reference Excel model, the HTML is wrong.

## Git workflow

This folder is a git repository. Commit all changes with clear, descriptive messages. Prefer separate commits for separate concerns (e.g. don't bundle a bug fix with a visual tweak). Use `git revert` rather than rewriting history when undoing changes.
