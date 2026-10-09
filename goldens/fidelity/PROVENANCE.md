Test fixtures in this directory are copied from xlq
(https://github.com/wms2537/aix, `fixtures/t1/`), which is dual MIT/Apache-2.0
upstream and used here under its MIT terms, this project being MIT-only.
Copied for the Phase 0 corpus spike and reused
here as permanent fidelity regression fixtures because they specifically
exercise the constructs an in-place workbook editor must preserve: `pivot-chart.xlsx`
(charts, pivot tables/caches, a structured table - also, incidentally, a
chart sheet IronCalc's in-memory model silently drops, which is exactly the
case that exposed a named-range scope bug the internal tool's notes
describes) and `macro.xlsm` (a VBA project).

`defined_names.xlsx` is copied from IronCalc's own test suite
(github.com/ironcalc/IronCalc, `tests/calc_tests/`), MIT - a
workbook with workbook- and sheet-scoped defined names and no dropped/exotic
sheet types, used as the "clean" (no chart-sheet quirk) counterpart to
pivot-chart.xlsx in the named-range regression tests.
