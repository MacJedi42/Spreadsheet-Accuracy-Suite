# What Excel request 3 settles

**The run.** Microsoft Excel for Mac 16.113.3, macOS 27.0.1, locale en_NZ, on 2026-10-02. Excel recalculated 52 workbooks in full and saved them. They hold every curated row of the formula golden (15,046 rows) plus `lookup-shapes.xlsx`.

**The files here:**
- `excel-cache.jsonl`: Excel's outcome for every row. This is the evidence; cite it as "Excel observed (request 3), <row id>".
- `import-report.txt`: the full report from `formula-golden.py --excel-import`, run on branch `gate/excel-export` `4114af7`.
- `files/`: the saved workbooks, with the author and local path blanked.
- `repair/`: Excel's repair logs.
- `RESULTS.md`: the session's notes.

## Against the golden

The golden had 14,553 decided rows. Excel agrees with 14,519 of them, and 81 of those agree only within the gate's tolerance. It disagrees with 33.

- **The snap under parentheses or unary plus.**
  - Cases: `=(A-B)`, `=((A-B))`, `=+(A-B)` and `=((A)-(B))`.
  - Excel KEEPS the residue at 1 to 7 ulps. The golden had 0 (basis `lo`).
  - This is the round-5 reviewer's corpus cell (B4R5-REV-OBS-2), now confirmed on 21 rows.
- **The snap threshold, across binades.**
  - `=A-B-C` over 1, 0.5000000000000002, 0.5000000000000002 keeps -4.44e-16, which is 4 ulps of the larger operand.
  - `=A-B` over 2+2^-49 and 2-2^-51 snaps.
  - So the threshold is not measured in ulps of the larger operand. One rule fits every observation so far, requests 1 and 3 together: **snap iff |result| < 2^(e-49), where e is the binary exponent of the subtraction's LEFT operand.** It still has to be checked against the corpus's 6,990 cached residues (`<scratch>/b4r5v/work/rv-obs/snapscan.py`).
- **`0^-1`** is `#DIV/0!`, not `#NUM!`.
- **ROUND past the 15th digit.** ROUND-rooted rows hold the models' 15-digit decimal exactly; the golden had held LibreOffice's double.
  - `ROUND(0.30000000000000004,16)` = 0.3.
  - `ROUND(1000000000000000.5,1)` = 1e15.
- **Within tolerance:**
  - **PMT** differs from LibreOffice by a few ulps in about 40 rows.
  - **SUM** adds plainly, without LibreOffice's compensation: `SUM` over ten 0.1 cells is 0.9999999999999999.

## The open questions

| Question | Excel's answer |
|---|---|
| Snap position | **Snaps:** the bare root `=A-B`, the end of `=0+A-B`, and a defined name's body, even when used as `=Gap*1` or `=Gap+0` (a name is evaluated like a formula of its own). **Does not snap:** a root wrapped in `(...)`, `+(...)` or `-(...)`; an IF branch (literal, computed or cell condition); `=A-B+0`, `=A-B-0`; `SUM(A-B)`. |
| Snap at other magnitudes | The same ulp count at 1e16, 2^60 and 1e-290. A subnormal residue becomes 0. |
| SUM's additions | **Every SUM's LAST addition snaps, whether or not the SUM is the root**: `SUM(A,-B)*1` and `SUM(A,-B)+0` give 0. Its other additions never snap: `SUM(A,-B,0)` keeps 8.9e-16, and `SUM(3,1e16,-1e16,0)` = 4. Neither of the models' two readings was this. |
| Integer powers | **Right-to-left binary exponentiation**, bit for bit on all 36 rows: n up to 1000 and down to -360, a cell or a literal exponent, and `((1+r)^12-1)*100`. A negative n is 1/x^n, not (1/x)^n. A non-integer exponent (12+2^-49) uses pow. |
| A negative base to a fractional power | A value only when 1/exponent is an odd integer: (-8)^(1/3) = -1.9999999999999998, (-32)^0.2 = -2, (-2)^(1/3) = -1.2599210498948732, (-27)^(-1/3) = -0.33333333333333337. **#NUM!** otherwise: (-8)^0.4, (-1)^0.5, (-8)^(2/3), -2^0.5. |
| A logical against a number | **A logical is greater than every number**: TRUE>1E300 is TRUE, and FALSE>3 is TRUE. |
| A logical against text | **A logical is greater than text**: TRUE>"a" is TRUE. |
| Text against a number | Text is greater than any number, and never equal: "3"=3 is FALSE. |
| Text against text | **Case-insensitive**: "a"="A". Ordered by character ("10"<"9"; "u"<"û"). A trailing space counts ("a "≠"a"). Against a blank cell: ""=blank, and "a">blank. |
| Text in arithmetic (en_NZ) | **Converts:** "3 ", " 3 ", "3.", ".5", "50%", "12:00", "2:15 PM", "$4.00", "(5)" (= -5), "1,000", "1,000.5", "1e3", "1E3", "+3", "-0". **#VALUE!:** "abc", "", " ", "0x10", "TRUE". Locale: DATEVALUE("1/2/2001") reads day first (1 February). |
| Unary plus on text | Keeps the text: `=+A` over "3 " gives the text "3 ". |
| IF over a text condition | `#VALUE!`, except the text "TRUE", which counts as TRUE. |
| ROUND with a text argument | Converts as arithmetic does: "3" gives 3. "abc", "" and "TRUE" give `#VALUE!`. |
| ROUND's digits | **Rounded to 15 significant digits, then truncated toward 0**: 1.5 → 1, -0.5 → 0, -1.5 → -1, 2.9999999999999996 → 3. d ≤ -16 gives 0 for small x; ROUND(5.5e15,-16) = 1e16. |
| PMT's type | Any nonzero type is type 1, including 0.9999999999999999, -0.5 and 1.5. |
| SUM over a logical through `(B1)` or `IF(TRUE,B1)` | Skipped: 0. |
| Equality (119 rows the readings split on) | Excel's answer on every row; the gate must find which reading, if either, gives all 119. |
| Subnormal input cells | Loaded as 0, and saved back as 0. |
| Hostile formulas | Accepted: parentheses up to 250 deep, 200 unary minus signs, 4,001 tokens, stacked `%` (8%% = 0.0008). Removed on load with a repair prompt: 1,000-deep parentheses, number literals of 323 to 400 characters, and `1E+308*10` and `1^1E+308` (an overflowing literal). |
| LOOKUP shapes | A 2-D RESULT vector gives `#N/A`. A 2-D lookup vector searches its first column (taller) or first row (wider). A short result vector extends in its own direction (H11 `LOOKUP(3,A1:B3,D1:D2)` reads D3 = 33), and a one-cell result extends along its row (H1 reads E1). The FREQUENCY idiom extends its result by one (`LOOKUP(1,1/FREQUENCY(5,A6:C6),A7:C7)` reads D7): IC-26 is right that FREQUENCY is one longer. Every extending LOOKUP is saved `ca="1"`. |
| Empty-text inputs | Excel keeps them as text (ISTEXT TRUE on all 22). |
