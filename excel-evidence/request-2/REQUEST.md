> **Record copy.** The original, with the 11 SpreadsheetBench workbooks it names in `files/`, is `excel-request-2/` at the repo root (not committed; the workbooks are corpus copies). The disagreement list it was drawn from comes from `DISAGREE_DUMP=<path>` on `examples/read_equivalence.rs`.

# What does Microsoft Excel show? Request 2: the differential's unexplained cases

**For:** a session with Microsoft Excel (the same setup as request 1 is ideal: Excel for Mac 16.113.3).
**From:** an internal spreadsheet tool's project, review batch B4 (2026-10-02).
**Files:** `files/` beside this document holds 11 workbooks from SpreadsheetBench, a public benchmark corpus. They are copies, so nothing in them is private.
**Time:** about 30 minutes.

## Why

The internal tool reads `.xlsx` files without Excel. To check its reader, we compare it, cell by cell, with a second reader (IronCalc) over about 5,000 public workbooks. 6,263 cells still disagree, and the harness cannot say which side is right. They come from about ten causes. In most of them the file itself shows that IronCalc is wrong, and we can prove that ourselves.

Excel is needed for the rest. Two causes are what Excel DISPLAYS for a stored number: a number a hair below a .5 rounding tie, and day 0 of 1900. The rest are what Excel makes of a few file constructs. Your answers settle those cases for good.

## Ground rules (please follow exactly)

1. **Set calculation to MANUAL before opening any file:** open a new blank workbook, then Formulas tab, Calculation Options, **Manual**. On a Mac this is also under Excel, Preferences, Calculation, Manually. Do this FIRST. Opening a file in automatic mode can recalculate it and replace the stored values we are asking about.
2. Open each file from `files/`. Do not save over them.
3. When a cell shows `#####`, widen its column until the full text shows. Then record the text exactly as displayed, with commas, signs and percent signs.
4. Record your Excel version: File, Account, About Excel.
5. Where a step says **full recalc**, press Ctrl+Alt+F9 (Mac: Cmd+Option+F9, or Formulas, Calculate Now with Shift held), and record the values again.
6. If a step's answer is an error, a prompt or something surprising, write it down and go on.

---

## Part A: cells in the corpus files

### A1. Numbers a hair below a .5 tie: what does the cell DISPLAY?

The stored value sits within a few units of the last binary place below x.5. The two readers disagree on whether it rounds up or down. Do not recalculate; just read the display.

| file | sheet | cell | stored value | format | the internal tool shows | IronCalc shows | Excel shows |
|---|---|---|---|---|---|---|---|
| `1_55427_answer.xlsx` | KS4 data | K76 | 0.14499999999999999 | `0%` | 15% | 14% | |
| `1_32468_answer.xlsx` | SUMMARY | I20 | -14705083.499999994 | `#,##0` | -14,705,084 | -14,705,083 | |
| `1_50442_answer.xlsx` | RESULTS 1 | G20 | 1322197.4999999995 | `#,##0` | 1,322,198 | 1,322,197 | |
| `1_50442_answer.xlsx` | CALCS | B12 | 1322197.4999999995 | `#,##0` | 1,322,198 | 1,322,197 | |
| `3_49613_input.xlsx` | RESULTS 1 | C33 | -2251.4999999999991 | `#,##0` | -2,252 | -2,251 | |

### A2. Day 0 under a day-of-month format

| file | sheet | cell | stored value | format | the internal tool shows | IronCalc shows | Excel shows |
|---|---|---|---|---|---|---|---|
| `1_118-8_answer.xlsx` | Schedule | H4 | 0 | `d` | 0 | #VALUE! | |
| `1_118-8_answer.xlsx` | Schedule | P28 | 0 | `d` | 0 | #VALUE! | |

### A3. A shared formula whose text is not in the range's first cell

In `1_44017_answer.xlsx`, sheet **Data**, the shared formula covers AD14:AO14. Its text is written in **AE14**, and AD14 holds a formula of its own. Click each cell below and copy the formula exactly as the FORMULA BAR shows it.

| cell | Excel's formula bar |
|---|---|
| AE14 | |
| AF14 | |
| AG14 | |
| AD15 | |
| AE15 | |
| AF15 | |

(the internal tool reads AF14 as `=IF($L14,S14*(1+SUM($M14:INDEX($M14:$P14,(YEARFRAC($L14,AF$9)*$J14+1)))*(AF$9>=$L14)),)`. That is AE14's text moved one column right. IronCalc reads it with AE14's text unmoved: R14 and AE$9.)

### A4. Control characters written as `_xHHHH_`

In `1_545-35_input.xlsx`, sheet **Data**, cell Y3 is stored as the text `˜_x001E_ø` and Z3 as `fï_x0006_`. The file format says `_x001E_` stands for the single character U+001E.
Pick any empty cell on that sheet, type each formula, and record the results:

| formula | Excel gives |
|---|---|
| `=LEN(Y3)` | |
| `=CODE(MID(Y3,2,1))` | |
| `=LEN(Z3)` | |
| `=CODE(MID(Z3,3,1))` | |

(If Excel decodes the escapes, the answers are 3, 30, 3 and 6. If it does not, they are 9, 95, 9 and 95.)

### A5. Do the stored values survive a full recalc? (IronCalc evaluates these to errors)

For each cell, record what it shows on opening (Manual calculation, so these are the stored values). Then do a **full recalc** and record again.

| file | sheet | cells | on opening | after full recalc |
|---|---|---|---|---|
| `1_CF_18814_answer.xlsx` | attendance-Dec | E1, C3, D3 | | |
| `1_11276_answer.xlsx` | ATTENDENCE | F3, G3 | | |
| `1_57716_answer.xlsx` | Calendar | B3, C3, D3, H3, and I3 | | |

### A6. MAX over a cell that has a format but no value

In `1_334-11_answer.xlsx`, sheet **imported Data**, C5 has a format but no value. Type these in an empty cell and record:

| formula | Excel gives |
|---|---|
| `=MAX(C2:C5)` | |
| `=COUNT(C2:C5)` | |
| `=COUNTA(C2:C5)` | |

---

## Part B: a new workbook, to settle the general rule

Part A gives single cases. Part B gives the rule behind them. Create a NEW blank workbook; Automatic calculation is fine here.

### B1. Display of numbers just below a .5 tie

In column A, type each formula. Then give column B the formula `=A1` (and so on down), and give each B cell the format in the table. Record what each B cell SHOWS.

| row | A (formula) | value it builds | B's format | Excel shows in B |
|---|---|---|---|---|
| 1 | `=0.145` | 0.14499999999999999 | `0%` | |
| 2 | `=-14705083.5+3*2^-29` | -14705083.499999994 | `#,##0` | |
| 3 | `=1322197.5-2^-31` | 1322197.4999999995 | `#,##0` | |
| 4 | `=-2251.5+2^-40` | -2251.4999999999991 | `#,##0` | |
| 5 | `=0.5-2^-54` | 0.49999999999999994 | `0` | |
| 6 | `=2.5-2^-51` | 2.4999999999999996 | `0` | |
| 7 | `=1.005` | 1.00499999999999989... | `0.00` | |
| 8 | `=0.285` | 0.28499999999999998... | `0%` | |

### B2. Day 0 and the 1900 leap-year day under date formats

In D1:D6 type the numbers **0, 0.5, 1, 59, 60, 61**. Copy them into columns E to N, then give each column one format:

| column | format |
|---|---|
| E | `d` |
| F | `dd` |
| G | `ddd` |
| H | `dddd` |
| I | `m` |
| J | `mmm` |
| K | `yyyy` |
| L | `m/d/yyyy` |
| M | `d-mmm-yy` |
| N | `dddd d mmmm yyyy` |

Record what every cell SHOWS, as six rows of ten. These are the cases the project's format check leaves out, because "no LibreOffice rendering is Excel's".

---

## What to send back

1. The Excel version and build.
2. The tables, filled in.
3. The Part B workbook, saved as `.xlsx`.
4. Anything surprising: a prompt on opening a file, a formula Excel rewrote, or a value that changed without a recalc.

Thank you. Each answer either confirms the file says what the internal tool reads or tells us Excel shows something else, and either way the case gets settled.
