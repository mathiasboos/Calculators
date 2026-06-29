# Calculators — Project Notes

## What this project is

Offline, dependency-free Swedish pension and salary calculators built as standalone HTML files. Each file is self-contained — no build step, no server, no external dependencies beyond fonts and charting libraries loaded via CDN where needed.

## Key 2026 constants

| Constant | Value |
|---|---|
| IBB (Inkomstbasbelopp) | 83,400 kr |

Add further constants here as they come up (e.g. PBB, takbelopp, skiktgräns).

## Tax constants (change annually — verify each January)

- **IBB (Inkomstbasbelopp):** single national value — see constants table
- **Brytpunkt (state tax threshold):** single national value — see constants table
- **Kommunalskatt:** varies by municipality (28.93%–35.65% in 2026, avg 32.38%)
  → Net_Salary file uses [a lookup table / a single average / user-input field — confirm which]
  → Source: ekonomifakta.se or skatteverket.se, published ~Dec each year
- **Jobbskatteavdrag:** NOT a flat number — income-dependent curve, also depends on kommunalskatt and age (under/over 66)
  → Do not hardcode a single jobbskatteavdrag value; verify the formula/table is re-implemented each year
  → Source: skatteverket.se, "jobbskatteavdrag" tables, published ~Dec each year

## Source of truth for formulas

When fixing or verifying calculation logic, **treat Excel as the source of truth**. If a formula in an HTML file diverges from the reference Excel model, the HTML is wrong.

## Git workflow

This folder is a git repository. Commit all changes with clear, descriptive messages. Prefer separate commits for separate concerns (e.g. don't bundle a bug fix with a visual tweak). Use `git revert` rather than rewriting history when undoing changes.
