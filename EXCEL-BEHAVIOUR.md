# Excel behaviour this suite established

**Every rule below names its source:**
- `R1` to `R4` are Excel requests 1 to 4 (`excel-evidence/request-N/`; requests 3 and 4 each have a `SETTLED.md`), and `R5` is their re-run on Windows Excel (`excel-evidence/request-5-windows/SUMMARY.md`);
- `ECMA` is ECMA-376;
- `MS-OI` is Microsoft's implementation notes [MS-OI29500];
- `corpus` means Excel's own caches in Excel-saved corpus files.

"Excel" here is Microsoft Excel for Mac 16.113 (locale en_NZ). **Windows Excel** (Microsoft 365, 16.0.20430, locale en-NZ) re-ran requests 3 and 4 as `R5` and matched it everywhere but the one power value marked below. Where a rule says *undecided*, no source settled it, and the goldens expect a refusal.

## 1. Arithmetic and formulas (formula golden)

### The snap to zero
**The rule:** a near-zero result of addition or subtraction becomes exactly 0, but only in these places, and only when |result| < 2^(e-49), where e is the binary exponent of the LEFT operand (R1, R3, corpus):
- **the bare root of the formula,** as in `=A-B`, and the end of `=0+A-B`;
- **the last addition of every SUM,** whether or not the SUM is the root: `SUM(A,-B)*1` gives 0;
- **a defined name's body,** which snaps as a formula of its own.

**Not inside these:** `(...)`, `+(...)`, `-(...)`, an IF branch, `=A-B+0`, or a SUM's inner additions. `SUM(A,-B,0)` keeps 8.9e-16, and `SUM(3,1e16,-1e16,0)` is 4. LibreOffice snaps at every add/sub, which is wrong.

### Equality and order
- Numbers are equal when they agree to 15 significant digits, and order follows equality (R3).
- Ties at the 16th digit (R4, decided in Excel's answers; the internal model still refuses them):
  - ROUND takes a tie toward zero: `ROUND(1234567890123.375,2)` is 1234567890123.37;
  - equality takes it away from zero: `100000000000000.5 = 100000000000001` is TRUE.

### ROUND
- ROUND works on the value's 15-digit decimal (R3): `ROUND(0.30000000000000004,16)` is 0.3, and `ROUND(1000000000000000.5,1)` is 1e15.
- Its digits argument is rounded to 15 digits and then truncated toward 0: 1.5 becomes 1, -1.5 becomes -1, and 2.9999999999999996 becomes 3.
- Past the decided range of -20 to 42 (R4): d=100 leaves x unchanged, and a large negative d gives 0, except that 1e300 survives d=-21.

### Powers
| case | Excel's rule | source |
|---|---|---|
| an integer exponent | right-to-left binary exponentiation (repeated squaring), bit for bit; x^-n is 1/x^n, not (1/x)^n | R3 |
| a positive base to a fractional exponent | exp(b·ln x), with sqrt at 0.5 and the reciprocal form for b<0 | corpus, R4 |
| a negative base | a value only when 1/b is an odd integer: (-8)^(1/3) is -1.9999999999999998; otherwise #NUM!. For b<0 Excel uses -1/exp(-b·ln\|x\|). **Platform-dependent:** `(-8)^(-1/3)` is -0.5000000000000001 on Mac and -0.5 on Windows | R3, R4, R5 |
| `0^-1` | #DIV/0! | R3 |
| `0^0` | #NUM! | R3 |

### Types
**Comparing types (R3):**
- a logical is greater than every number and every text: `TRUE>1E300` and `FALSE>3` are TRUE;
- text is greater than any number and never equal to one: `"3"=3` is FALSE.

**Comparing text with text (R3):**
- case-insensitively;
- by character: `"10"<"9"`;
- a trailing space counts: `"a "≠"a"`.

**Blanks and logicals:**
- a blank equals `""` and 0;
- SUM over a logical through a reference skips it.

### Text in arithmetic (en_NZ, R3 and R4)
**Converts:**
- `"3 "`, `"50%"`, `"12:00"`, `"2:15 PM"`, `"$4.00"`;
- `"(5)"`, which is -5;
- `"1,000.5"`, `"1e3"`, `"7%"`, `"$1,000"`, `"(1,000)"`, `"1,0000"`.

**Gives #VALUE!:** `"abc"`, `""`, `" "`, `"0x10"`, `"TRUE"`, `"N/A"`, `"May"`.

**Locale-dependent:** DATEVALUE reads day-first in en_NZ.

### Unary minus and precedence
- Unary minus binds tighter than `^`: `=-2^2` is 4, and `=-10^16` is +1E16 (R1).
- Exponentiation chains left to right: `=2^3^2` is 64.

### PMT
- Any nonzero `type` is type 1 (R3).
- Excel's PMT is not reproducible to the last bit: it differs from the closed form by up to 11 units of 2^-52 of its scale. At very high growth (rate 0.5 over 20 to 100 periods) it departs much further, and no rule for that is known (R3, R4).

### LOOKUP
- **A short result vector** extends in its own direction; a one-cell result extends along its row (R1, R3).
- **Shape errors:** a 2-D result vector gives #N/A. A 2-D lookup vector searches its first column if taller, its first row if wider.
- **Recalculation:** Excel saves every extending LOOKUP with `ca="1"` (calculate always).

### Overflow
SUM, AVERAGE, STDEV and VAR whose arithmetic passes about 1.8e308 give #NUM!, AVERAGE included (R4). STDEV and VAR over 1E+200, -1E+200, 0 give #NUM!, because the squares overflow, while SUM and AVERAGE give 0.

### Small cases
- A subnormal input is 0 (R3).
- Percent stacks: `8%%` is 0.0008.
- Excel accepts 250 nested parentheses, 200 unary minus signs and 4,001 tokens.
- On load, Excel repairs away a formula with 1,000-deep parentheses, a 323-to-400-character number literal, or `1E+308*10`.

## 2. Reading the file (decoding golden)

### Positions
- **No address (ECMA 18.3.1.73, 18.3.1.4; R4 confirmed it, no repair):**
  - a `<row>` with no `r` is the previous row + 1;
  - a `<c>` with no `r` is the previous cell's column + 1;
  - a self-closing `<row/>` or `<c/>` counts.
- **A repeated row (R4, no repair):** two `<row r="2">` elements are one row, read as the cells of both.
- **Damage Excel repairs (R4, removed records):**
  - two different cells at one address (Excel kept the first);
  - a cell whose `r` names another row;
  - cells or rows out of order.
- **Byte-identical duplicates:** Excel repairs them too, leaving the same value.

### Values
- **Phonetic runs:** `<rPh>` is not part of a string; Excel shows the base text only (ECMA 18.4.6; R4).
- **Escapes:** `_xHHHH_` escapes decode, `_x0009_` (tab) and `_x000A_` (line feed) included. Excel READS them, although MS-OI 2.1.1747 says Office does not WRITE them (R2, R4).
- **Trimming:** a `<t>` WITHOUT `xml:space="preserve"` loses leading and trailing spaces, tabs and line feeds, in shared and inline strings alike (R4: `' lead'` reads `lead`, `LEN` 4). Internal whitespace is kept, as Excel writes it (corpus). Not decided: a rich-text run's `<t>` with end whitespace.
- **ISO dates (`t="d"`):** they are the serial of that moment in the workbook's date system (ECMA 18.17.4). Excel reads `YYYY-MM-DD` and `YYYY-MM-DDThh:mm[:ss[.fff]]`. Other forms made Excel repair the file (R4).
- **Booleans:** `1` and `0` only (MS-OI 2.1.660). `true` and `false` made Excel repair the file (R4).
- **Numbers that are damage:** `+5`, `NaN`, `INF`, `-INF` and `1e400` in a `<v>` are damage. Excel repaired them into text, or into #NULL!. `<v> 42 </v>`, `1E3`, `.5`, `5.` and `-0` are numbers, and `-0` is 0 (R4).
- **Error codes:** only `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!` and `#N/A`. Others, such as `#SPILL!` or `#FOO!` stored in a file, made Excel repair it into text (`#BUSY!` stayed an error) (R4).

### The package
- **Finding parts:** a workbook's parts are found through its relationships, not by name. That includes a string table at another name, `..` and absolute targets, a relocated workbook part, and strict-namespace relationships (ECMA Part 2; R4).
- **Missing parts:** a missing styles part reads as no styles, with every cell General. A missing string table or sheet part is damage that Excel repairs (R4).
- **Encodings Excel reads (R4):**
  - UTF-8 and UTF-16 with a byte-order mark;
  - BOM-less UTF-16 that declares `encoding="UTF-16"`;
  - parts declared `windows-1252` or `ISO-8859-1` (`Caf\xe9` reads `Café`).
- **Not decided:** other bytes in those single-byte encodings (Excel saw only ASCII and 0xE9).
- **Invalid bytes, cut-off parts and mismatched tags:** these are damage. Excel replaces or removes the part, and refuses a cut-off workbook part outright (R4).

### The 1904 date system
**The rule:** `workbookPr date1904="1"` or `"true"` means serial 0 is 1904-01-01, and there is no fictional 1900-02-29 (ECMA 18.2.28; R4).
- 1462 is 1908-01-02, and 59 is 1904-02-29.
- A 1904 serial is the 1900 serial minus 1462.
- A negative serial gives #NUM! from DAY.

**Strict workbooks** (`conformance="strict"`) draw dates like the ordinary system, despite MS-OI 2.1.585 (R4).

## 3. Display (number-format golden)

- **15 digits:** Excel shows at most 15 significant digits and then zeros (R4, the IC-51 rows). The 15 digits are rounded from the double's exact binary value, and the display then rounds half away from zero at its last place.
- **Undecided ties:** past about 2^41, where LibreOffice draws the binary rounding instead, a 13- or 14-digit display at a tie is not yet decided, and the golden leaves it out.
- **Clock rounding:** Excel rounds a time to the second when the code shows no fractions of a second, and to the shown fraction when it does. A date-only code takes its day at the second: 0.7 ms before midnight is the next day (R4). LibreOffice truncates a time instead.
- **Date functions:** DAY, MONTH and YEAR round the serial to the second too. Past 9999-12-31 at that rounding there is no date: #NUM!, or #VALUE! from TEXT (R4).
- **The 1900 epoch:** serial 0 is 1900-01-00, and serial 60 is the fictional 1900-02-29. LibreOffice is a day off below serial 61.
- **Built-in date formats 14, 15, 16 and 22** are drawn per locale. In en-NZ, id 15 is `dd-mmm-yy` (R2).

## 4. Where LibreOffice is not Excel

These are the places where LibreOffice is wrong as an Excel reference. Use it for anything else only as a second opinion.

**Formulas:**
- it snaps near-zero results at every add/sub;
- SUM over a range counts logicals;
- `-TRUE` keeps a logical type;
- powers use `pow`, not repeated squaring, which is up to 3.6e-12 relative away in finance idioms such as `((1+r)^12-1)*100`;
- it reads `t="d"` zone forms and `true`/`false` booleans that Excel repairs.

**Display:**
- the 1900 epoch below 61;
- a minus sign kept on a one-section negative that rounds to zero;
- `#.##` drops a trailing point that Excel keeps;
- `[$-x-sysdate]` and `[$-F400]`;
- `A/P` and `y`/`yyy`;
- it draws a 16th digit;
- it truncates clock times;
- General's digit count;
- it merges `###` and `#,##` within one workbook;
- a silent General fallback for codes it cannot read;
- a fraction search that never returns on tiny values;
- some 15-digit ties past about 2^41.

**Saving:** its xlsx writer stores 15 significant digits, so a LibreOffice round trip cannot prove a full-precision read.

## 5. What the suite does not decide

These are refused in the goldens, as `undecided:*` or `excel:repaired`. A later Excel request can settle them:
- PMT at high growth;
- text-to-number arithmetic beyond the observed forms;
- ISO date forms with zones, week dates or ordinal dates;
- rich-text runs with end whitespace;
- single-byte encodings beyond ASCII and 0xE9;
- 15-digit display ties past 2^41;
- `date1904` values outside `1`/`0`/`true`/`false`.
