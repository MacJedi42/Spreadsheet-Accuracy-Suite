# review batch B4: Excel observations

## 1. Excel version

- **Microsoft Excel for Mac, Version 16.113.3 (build 26092714)**. AppleScript reports `build` 927; the saved file has `calcId="191029"` and `rupBuild="10927"`.
- **Mac**: macOS 27.0.1 (26A434), Apple Silicon (arm64).
- Calculation was **Automatic** throughout.

**Method note.** Values and formulas were entered with AppleScript, by setting each cell's `formula2` property. That is Range.Formula2, the dynamic-array equivalent of typing into the formula bar. The edit steps were entered the same way. For "Ctrl+Alt+F9" I used AppleScript `calculate full` (Application.CalculateFull). The displayed values are each cell's displayed text (Range.Text). On Sheet2 the columns were autofit first, because at default width Excel shows fewer digits (for example `1.7764E-15`). Stored values come from the saved XML.

Files:
- `excel-observations.xlsx`: the requested workbook, with Part 1 on Sheet1 and Part 2 on Sheet2. Its state is as of step 4.
- `excel-extras.xlsx`: extra probes that were not requested (section 5). They are kept in a separate file so the main workbook stays exactly as specified.

---

## 2. Part 1: LOOKUP with a short result vector

| cell | first shown | after B3 → 99 | after C1 → 55 | after full recalc | stored `<f>` flags |
|---|---|---|---|---|---|
| H1 | 21 | 21 | 21 | 21 | none |
| H2 | 12 | 12 | **55** | 55 | `ca="1"` |
| H3 | 31 | **99** | 99 | 99 | `ca="1"` |
| H4 | 13 | 13 | 13 | 13 | `ca="1"` |
| H5 | 21 | 21 | 21 | 21 | `ca="1"` |
| H6 | 31 | **31 (stale)** | 31 (stale) | **99** | `t="array" ref="H6"`, cell `cm="1"`; **no `ca`** |
| H7 | 73 | 73 | 73 | 73 | `ca="1"` |
| H8 | 72 | 72 | 72 | 72 | `ca="1"` |
| H9 | 63 | **131** | 131 | 131 | `ca="1"` |
| H10 | 21 | 43.6666667 | 43.6666667 | 43.6666667 | `ca="1"` (stored 43.666666666666664) |
| H11 | 32 | 32 | 32 | 32 | none |
| H12 | 63 | 131 | 131 | 131 | `ca="1"` |

### Findings

1. **Excel extends short result vectors, as LibreOffice does.** All twelve first values match the LibreOffice column. A one-cell result vector extends along the row (H2 = C1, H8 = B7). A short vector extends in its own direction (H3 = B3, H7 = C7). A horizontal result for a vertical lookup reads D1 (H4 = 13).
2. **For range arguments, Excel handles this by marking the formula "calculate always".** In the saved file, every LOOKUP whose result vector is a different size from its lookup vector has `ca="1"` on its `<f>`: H2, H3, H4, H5 and H7, H8. H5's result vector is longer than its lookup vector, but it is flagged too. So are the SUMIF/AVERAGEIF cells whose sum range is resized upward or reshaped: H9, H10 and H12. H11, whose sum range shrinks, has no flag, and neither does H1. As a result these formulas are recomputed on every recalc, so edits to extended cells show up at once. That covers B3 (H3, H9, H10, H12) and C1 (H2). Section 5 adds D1, C7 and B7.
   - Excel's own **Trace Precedents** lists only the literal ranges. For example, H3's direct precedents are `$A$1:$A$3,$B$1:$B$2`, with no B3. So the immediate update comes from the always-calculate flag, not from B3 being registered as a precedent.
3. **H6 (`=LOOKUP(2,1/(A1:A3=3),B1:B2)`) reads B3 but does not track it.** After B3 → 99, H6 kept showing 31 through both edit steps. It showed 99 only after the full recalc. Excel saved H6 as a single-cell dynamic-array formula with no `ca` flag. I reproduced this in a fresh workbook (section 5). **In this case Excel's own cached value goes stale until a full recalc.**

---

## 3. Part 2: what the tool now refuses (Sheet2)

### 2a. Subtracting nearly equal numbers

`A` = `1+k*2^-52` for k = 1, 2, 4, 5, 6, 7, 8, 9, 12, 16. All ten show `1`; stored values are `1.0000000000000002` … `1.0000000000000036`.

| col | shown, rows 1–10 |
|---|---|
| C `=Ar-$B$1` | 0, 0, 0, 0, 0, 0, 1.77636E-15, 1.9984E-15, 2.66454E-15, 3.55271E-15 |
| D `=(Ar-$B$1)*1` | 2.22045E-16, 4.44089E-16, 8.88178E-16, 1.11022E-15, 1.33227E-15, 1.55431E-15, 1.77636E-15, 1.9984E-15, 2.66454E-15, 3.55271E-15 |
| E `=SUM(Ar,-$B$1)` | 0, 0, 0, 0, 0, 0, 1.77636E-15, 1.9984E-15, 2.66454E-15, 3.55271E-15 |
| F `=Ar=$B$1` | TRUE, TRUE, TRUE, TRUE, TRUE, TRUE, TRUE, TRUE, TRUE, TRUE |
| G `=Ar>$B$1` | FALSE, FALSE, FALSE, FALSE, FALSE, FALSE, FALSE, FALSE, FALSE, FALSE |

Stored values:

| k | C (`A−B`) | D (`(A−B)*1`) | E (`SUM`) |
|---|---|---|---|
| 1 | 0 | 2.2204460492503131E-16 | 0 |
| 2 | 0 | 4.4408920985006262E-16 | 0 |
| 4 | 0 | 8.8817841970012523E-16 | 0 |
| 5 | 0 | 1.1102230246251565E-15 | 0 |
| 6 | 0 | 1.3322676295501878E-15 | 0 |
| 7 | 0 | 1.5543122344752192E-15 | 0 |
| 8 | 1.7763568394002505E-15 | 1.7763568394002505E-15 | 1.7763568394002505E-15 |
| 9 | 1.9984014443252818E-15 | 1.9984014443252818E-15 | 1.9984014443252818E-15 |
| 12 | 2.6645352591003757E-15 | 2.6645352591003757E-15 | 2.6645352591003757E-15 |
| 16 | 3.5527136788005009E-15 | 3.5527136788005009E-15 | 3.5527136788005009E-15 |

What this shows:
- **The cutoff lies between 7 and 8 units of 2^-52.** At operands near 1, `A−B` is stored as exactly 0 when the true difference is 7·2^-52 or less. From 8·2^-52 (that is, 2^-49) upward, the exact difference is kept. Only magnitude 1 was tested, so it is not yet known whether the cutoff scales with the operands.
- **The zeroing applies only to the formula's last step.** In D the subtraction is not the last step, and no value is zeroed: every exact difference survives `*1`.
- **SUM applies the same rule.** E matches C exactly.
- **`=` and `>` use a wider tolerance.** All ten rows compare equal, even at 16·2^-52 ≈ 3.55E-15, where subtraction gives a nonzero result. So comparisons treat as equal some values whose subtraction does not give 0.

### 2b. ROUND near a tie

| cell | formula | shows | stored | answer |
|---|---|---|---|---|
| J1 | `=ROUND(2.5-2^-51,0)` | 3 | 3 | **3** |
| J2 | `=ROUND(0.5-2^-54,0)` | 1 | 1 | **1** |
| J3 | `=ROUND(1250-2^-42,-2)` | 1300 | 1300 | **1300** |
| J4 | `=ROUND(10+4.96E-13,12)` ¹ | 10 | 10.000000000001 | **10.000000000001** |
| J5 | `=ROUND(1234567890123+0.125,2)` | 1.23457E+12 | 1234567890123.1201 | **.12**: the stored value is the double nearest 1234567890123.12 |
| J6 | `=ROUND(2^53,0)` | 9.0072E+15 | **9007199254740990** | **It changes**: 9007199254740992 becomes 9007199254740990 |
| J7 | `=ROUND(1.5*0.15,2)` | 0.23 | 0.23 | 0.23, as expected |
| J8 | `=ROUND(1/3,20)` | 0.333333333 | 0.33333333333333298 | the double nearest 0.333333333333333 (15 threes); 1/3 itself would be 0.33333333333333331 |

¹ Excel rewrote the formula as `=ROUND(10+0.000000000000496,12)`.

All eight results fit one rule: ROUND first reduces its argument to 15 significant digits, then rounds half away from zero. J4 shows a double rounding: 10.000000000000496 → 10.0000000000005 → 10.000000000001. J5 is an exact tie at the 16th digit (…123.125), and it went down to .12. So that first 15-digit step does not round a tie up. Only one data point covers that case.

### 2c. Other undocumented cases

| cell | formula | shows | stored | note |
|---|---|---|---|---|
| L1 | `=0^0` | **#NUM!** | `t="e"` #NUM! | |
| L2 | `=(-8)^(1/3)` | **-2** | **-1.9999999999999998** | an odd root of a negative base is allowed. The result is not exactly -2, although IEEE `pow(8,1/3)` gives exactly 2.0 |
| L3 | `=TRUE=1` | **FALSE** | `t="b"` 0 | |
| L4 | `=+TRUE` | **TRUE** | `t="b"` 1 | unary plus does not convert to a number |
| L5 | `=Z99=FALSE` | **TRUE** | `t="b"` 1 | |
| L6 | `=2^-1074` | **0** | exactly 0 | a subnormal result is flushed to zero, with no error |
| L7 | `=PMT(0.05,12,1000,0,2)` | **($107.45)** ² | -107.45277144839562 | **type 2 is treated as 1**. Type 0 would give -112.83 |
| L8 | `=N1+1` | **4** | 4 | N1 was confirmed to be the text `" 3"`, with prefix `'` |
| L9 | `=SUM(P1:Q2)` | 2E+16 | 2E+16 | **the test as written cannot tell the two orders apart; see below** |

² Excel automatically applied a currency number format to L7.

**L9: the test as written does not work.** In Excel, unary minus binds tighter than `^`, so `=-10^16` means `(-10)^16`, which is **+1E+16** (likewise `=-2^2` gives 4). With P2 = +1E16, both orders give 2E16. In the extras workbook I used `=-(10^16)` instead, and **L9 gave exactly 1**. So **SUM adds row by row** (P1, Q1, P2, Q2): 1E16+1 rounds to 1E16, minus 1E16 is 0, plus 1 is 1. Column by column would give 2.

---

## 4. Anything surprising

1. **`=-10^16` is +1E16** (negation before exponentiation). Because of this, L9 as specified cannot answer its question; the corrected version is in section 5.
2. **H6 stays stale in Excel itself** until a full recalc (Part 1, finding 3). It is the only mis-sized formula without `ca="1"`.
3. **The saved file flags the extension cases with `ca="1"`.** This applies to LOOKUP with mismatched vectors and to SUMIF/AVERAGEIF whose sum range is resized upward or reshaped. A reader can use this flag as a signal.
4. Excel **rewrote J4's literal** from `4.96E-13` to `0.000000000000496`.
5. Excel **auto-formatted L7** as currency: `($107.45)`.
6. **`(-8)^(1/3)` is not exactly -2**: it is stored as -1.9999999999999998.
7. Excel showed no error messages, prompts or refusals.

---

## 5. Extras (not requested; `excel-extras.xlsx`)

This is the same Part 1 layout in a fresh workbook, with Automatic calculation. The steps below each change one extended cell, then a full recalc follows. Each column shows the values after that step.

| cell | start | D1 → 66 | C7 → 77 | B7 → 88 | B3 → 99 | full recalc |
|---|---|---|---|---|---|---|
| H3 | 31 | 31 | 31 | 31 | **99** | 99 |
| H4 (reads D1) | 13 | **66** | 66 | 66 | 66 | 66 |
| H6 | 31 | 31 | 31 | 31 | **31 (stale)** | **99** |
| H7 (reads C7) | 73 | 73 | **77** | 77 | 77 | 77 |
| H8 (reads B7) | 72 | 72 | 72 | **88** | 88 | 88 |
| H9 / H10 / H12 | 63 / 21 / 63 | same | same | same | 131 / 43.6666667 / 131 | same |

- Edits to the extended cells D1, C7 and B7 show up at once in H4, H7 and H8. All three carry `ca="1"`.
- H6's staleness reproduced.
- The SUM order test was corrected: P1 `=10^16`, Q1 `1`, P2 `=-(10^16)`, Q2 `1`. **`=SUM(P1:Q2)` gave 1**, stored as exactly `1`, which means row by row.
- Precedence checks: `=-10^16` gives 1E+16, `=0-10^16` gives -1E+16, and `=-2^2` gives 4.
