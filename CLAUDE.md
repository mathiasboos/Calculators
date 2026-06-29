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

### Updated cross-file consistency — grundavdrag scope

| What | Net_Salary_2026.html | Salary_Exchange.html | Salary_Exchange_Consumer.html |
|---|---|---|---|
| Grundavdrag | full piecewise (both age formulas) | not used — file works purely with marginal rates, never computes gross-to-net | not used — same reason |

Annual update scope for grundavdrag: Net_Salary_2026.html ONLY. The two
salary exchange calculators don't compute tax from gross income up, so
they have no grundavdrag dependency at all.

### Cross-file consistency — annual update scope differs per file

| What | Net_Salary_2026.html | Salary_Exchange.html | Salary_Exchange_Consumer.html |
|---|---|---|---|
| Kommun lookup table | full table | full table (duplicate) | single default (KOMMUNAL_DEFAULT) |
| Jobbskatteavdrag | full piecewise formula | simplified (brytpunkt check + flat 20% state tax above it) | simplified, same as Salary_Exchange |
| Brytpunkt constant | BRYT_UNDER = 660,400 (annual) | LIMIT_BRYTPUNKT = 55,033 (monthly, confirmed ≈ matches) | not used |
| SKIKT (skiktgräns) | yes | not used | not used |
| Annual update needed | full re-verify of all tax logic | replace kommun table; verify LIMIT_BRYTPUNKT | update KOMMUNAL_DEFAULT only |

⚠ Updating Net_Salary_2026.html's kommun table does NOT automatically
update Salary_Exchange.html's copy — both must be replaced separately
each December, or the two calculators will disagree.

## Source of truth for formulas

When fixing or verifying calculation logic, **treat Excel as the source of truth**. If a formula in an HTML file diverges from the reference Excel model, the HTML is wrong.

## Git workflow

This folder is a git repository. Commit all changes with clear, descriptive messages. Prefer separate commits for separate concerns (e.g. don't bundle a bug fix with a visual tweak). Use `git revert` rather than rewriting history when undoing changes.
