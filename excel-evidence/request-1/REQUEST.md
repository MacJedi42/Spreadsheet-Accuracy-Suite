# What does Microsoft Excel do? A request for observations

**For:** a session with Microsoft Excel (desktop, Windows or Mac).
**From:** an internal spreadsheet tool's project, review batch B4 (2026-10-02).
**Time needed:** about 20 minutes for Part 1, and about 30 more for Part 2, which is optional.

## Why we are asking

The internal tool is a command-line tool that AI agents use to read and edit `.xlsx`
files without Excel. When it edits a cell, it works out which formulas depend
on that cell. For simple formulas it computes their new values; for the rest
it removes their now-stale cached results.

It knows Excel's behaviour only from Microsoft's documentation, from
LibreOffice, and from cached values in old Excel-saved files. Where these
disagree or say nothing, the tool currently refuses to compute.

Observations from real Excel settle such questions. Without them the tool
must either guess or refuse. You do not need the internal tool or anything else from
this project.

## Ground rules (please follow them exactly)

1. Use a NEW blank workbook. Calculation must be **Automatic**: on the
   Formulas tab, Calculation Options, Automatic.
2. Type every value and formula exactly as written. A formula starts with
   `=`. Do not press Ctrl+Shift+Enter; plain Enter is right.
3. Record what each formula cell SHOWS, exactly as shown: a number, `#N/A`,
   `#NUM!`, `#VALUE!`, `TRUE` or `FALSE`.
4. Record your Excel version: File, Account, About Excel. Note the version,
   the build, and whether it is Windows or Mac.
5. When you finish, SAVE the workbook as `.xlsx` and send it back with your
   tables. We read the saved file's stored values at full precision. Excel
   only displays 15 digits, and Part 2 needs more.
6. If Excel refuses a formula, record the message and move on.

---

## Part 1 (required): does LOOKUP read past a short result vector?

### The question

`LOOKUP(lookup_value, lookup_vector, result_vector)`. Microsoft's LOOKUP
page says the `result_vector` "must be the same size as `lookup_vector`".
It does not say what happens when it is not.

LibreOffice EXTENDS a short result vector, the way Excel's SUMIF documents
resizing its sum range:
- `LOOKUP(3, A1:A3, B1:B2)` returns B3's value, a cell the formula never
  names.
- `LOOKUP(2, A1:A3, B1)` returns C1's: a single cell is extended along the
  ROW.

If Excel does the same, the internal tool must treat those extra cells as inputs.
Otherwise an edit to one of them leaves the formula's cached result stale
without a warning. We also need to know whether Excel's own recalculation
notices an edit to an extended cell.

### Setup: type these values (Sheet1)

Every value is different, so you can tell which cell a formula returned.

| | A | B | C | D | E |
|---|---|---|---|---|---|
| **1** | 1 | 11 | 12 | 13 | 14 |
| **2** | 2 | 21 | 22 | 23 | 24 |
| **3** | 3 | 31 | 32 | 33 | 34 |
| **4** | 4 | 41 | 42 | 43 | 44 |
| **6** | 1 | 2 | 3 | | |
| **7** | 71 | 72 | 73 | 74 | |
| **8** | 81 | 82 | 83 | 84 | |

(Row 5 is empty. Row 6 holds 1, 2, 3 in A6:C6.)

### Formulas: type each in column H, rows 1 to 12

| cell | formula | what it tests | LibreOffice shows |
|---|---|---|---|
| H1 | `=LOOKUP(2,A1:A3,B1:B3)` | control: same size | 21 |
| H2 | `=LOOKUP(2,A1:A3,B1)` | a one-cell result vector | 12 (C1, along the row) |
| H3 | `=LOOKUP(3,A1:A3,B1:B2)` | a result vector one cell short | 31 (B3) |
| H4 | `=LOOKUP(3,A1:A3,B1:C1)` | a horizontal result for a vertical lookup | 13 (D1) |
| H5 | `=LOOKUP(2,A1:A2,B1:B3)` | a result vector LONGER than the lookup | 21 |
| H6 | `=LOOKUP(2,1/(A1:A3=3),B1:B2)` | the common "last match" idiom, short result | 31 |
| H7 | `=LOOKUP(3,A6:C6,A7:B7)` | a horizontal lookup, short result | 73 (C7) |
| H8 | `=LOOKUP(2,A6:C6,A7)` | horizontal, one-cell result | 72 (B7, along the row) |
| H9 | `=SUMIF(A1:A3,">0",B1)` | SUMIF resize (documented), as a check | 63 |
| H10 | `=AVERAGEIF(A1:A3,">0",B1)` | AVERAGEIF resize (documented) | 21 |
| H11 | `=SUMIF(A1:A2,">0",B1:B4)` | a sum range LONGER than the criteria | 32 |
| H12 | `=SUMIF(A1:A3,">0",B1:C1)` | a horizontal sum range for a vertical criteria range | 63 (read as B1:B3) |

### Then: does Excel notice an edit to an extended cell?

Do these steps in order. Record H2, H3, H4, H6 and H9 after each step.
1. **Before:** the values as first computed.
2. Change **B3** from 31 to **99** and press Enter, with Automatic still on.
   Does H3 show 99 at once? Does H6? Does H9 change, to 131?
3. Change **C1** from 12 to **55** and press Enter. Does H2 show 55 at once?
4. Press **Ctrl+Alt+F9**, a full recalculation of everything (on a Mac: Cmd+Option+F9, or Formulas, Calculate Now). Record again.

(The LibreOffice column was measured with LibreOffice 24.2 on this exact layout.)

If a value changes only at step 4, Excel reads that cell but does not track
it as an input. That is just as important for us to know.

### Results for Part 1 (please fill in)

| cell | first shown | after B3 → 99 | after C1 → 55 | after full recalc |
|---|---|---|---|---|
| H1 | | | | |
| H2 | | | | |
| H3 | | | | |
| H4 | | | | |
| H5 | | | | |
| H6 | | | | |
| H7 | | | | |
| H8 | | | | |
| H9 | | | | |
| H10 | | | | |
| H11 | | | | |
| H12 | | | | |

---

## Part 2 (optional, very valuable): settle what the tool now refuses

The internal tool refuses to compute where Excel's behaviour is undocumented and the
evidence it has disagrees. Each line below is one such question. Use a NEW
SHEET, Sheet2, for these, in the same workbook. **Save and return the
workbook:** most of these answers are in the 16th and 17th digits, which
only the saved file shows.

Some inputs must be built with a formula. Excel cuts a TYPED number to 15
digits, and these questions are about values just beside a round number.
`2^-52` is the gap between 1 and the next number a double can hold.

### 2a. Subtracting nearly equal numbers

Excel documents that "should an addition or subtraction operation result in a
value at or very close to zero", it gives 0 (KB "Floating-point arithmetic
may give inaccurate results in Excel"). We need to know how close, and
whether this applies only to the formula's last step.

In Sheet2, A1:A10 hold `=1+k*2^-52` for k = 1, 2, 4, 5, 6, 7, 8, 9, 12, 16, in that
order. A1 is `=1+1*2^-52`, A2 is `=1+2*2^-52`, and so on to A10, which is
`=1+16*2^-52`. B1 holds `1`.

For each row r from 1 to 10, enter:

| column | formula (row r) | what it tests |
|---|---|---|
| C | `=Ar-$B$1` | the subtraction is the formula's last step |
| D | `=(Ar-$B$1)*1` | the subtraction is NOT the last step |
| E | `=SUM(Ar,-$B$1)` | inside SUM |
| F | `=Ar=$B$1` | equality (TRUE or FALSE) |
| G | `=Ar>$B$1` | ordering (TRUE or FALSE) |

Record what C to G show. The saved file gives us C to E's exact values; a
value shown as `0` may not be exactly 0.

### 2b. ROUND near a tie

These inputs are built so that the stored number sits a hair below a
rounding tie. Enter each in Sheet2, column J, and record what it shows:

| cell | formula | the question |
|---|---|---|
| J1 | `=ROUND(2.5-2^-51,0)` | is 2.4999999999999996 rounded to 2 or to 3? |
| J2 | `=ROUND(0.5-2^-54,0)` | is 0.49999999999999994 rounded to 0 or to 1? |
| J3 | `=ROUND(1250-2^-42,-2)` | is 1249.9999999999998 rounded to 1200 or to 1300? |
| J4 | `=ROUND(10+4.96E-13,12)` | 10 or 10.000000000001? |
| J5 | `=ROUND(1234567890123+0.125,2)` | is it .12 or .13 (a 16-digit value)? |
| J6 | `=ROUND(2^53,0)` | 9007199254740992 stays, or does it change? |
| J7 | `=ROUND(1.5*0.15,2)` | control: we expect 0.23 |
| J8 | `=ROUND(1/3,20)` | digits beyond 15 |

### 2c. Other undocumented cases

| cell | formula (Sheet2) | the question |
|---|---|---|
| L1 | `=0^0` | `#NUM!` or 1? |
| L2 | `=(-8)^(1/3)` | `#NUM!` or -2? |
| L3 | `=TRUE=1` | TRUE or FALSE? |
| L4 | `=+TRUE` | is TRUE shown, or 1? |
| L5 | `=Z99=FALSE` | Z99 is empty: TRUE or FALSE? |
| L6 | `=2^-1074` | 0, `#NUM!`, or a tiny number? |
| L7 | `=PMT(0.05,12,1000,0,2)` | type 2: treated as 1, or an error? |
| L8 | `=N1+1`, with N1 holding the TEXT `' 3` (a leading space; type the apostrophe so it is stored as text) | 4 or `#VALUE!`? |
| L9 | `=SUM(P1:Q2)`, with P1 `=10^16`, Q1 `1`, P2 `=-10^16`, Q2 `1` | 1 (added row by row) or 2 (column by column)? |

### Results for Part 2 (please fill in)

Record what each cell SHOWS. We read the exact values from your saved file.

| cell | shows | | cell | shows | | cell | shows |
|---|---|---|---|---|---|---|---|
| C1..C10 | | | J1 | | | L1 | |
| D1..D10 | | | J2 | | | L2 | |
| E1..E10 | | | J3 | | | L3 | |
| F1..F10 | | | J4 | | | L4 | |
| G1..G10 | | | J5 | | | L5 | |
| | | | J6 | | | L6 | |
| | | | J7 | | | L7 | |
| | | | J8 | | | L8 | |
| | | | | | | L9 | |

(For C to G, list the ten values top to bottom, for example
`0, 0, 0, 0, ?, ?, ...`.)

---

## What to send back

1. The Excel version, build, and Windows or Mac.
2. The Part 1 table, and the Part 2 tables if you did Part 2.
3. The saved `.xlsx` file. If you cannot send the file, rename a copy to
   `.zip`, open `xl/worksheets/sheet1.xml` and `sheet2.xml` in a text editor,
   and paste them. Each formula cell's stored value is the `<v>…</v>` after
   its `<f>…</f>`.
4. Anything Excel said or did that surprised you, such as an error message,
   an automatic correction to a formula, or a prompt.

Thank you. Each answer lets the tool compute a value, or flag one, where today
it can only refuse.
