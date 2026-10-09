# What Excel request 4 settles

**The run.** Microsoft Excel for Mac 16.113.4, macOS 27.0.1, locale en_NZ, on 2026-10-08, driven by a session with Excel.

**Decoding.** Excel opened 78 decoding workbooks (`scripts/decoding-golden.py --excel-export`) and saved each one as .xlsx:
- 50 opened cleanly;
- 27 opened only after Excel repaired them (a repair log each);
- 1 was refused: `val-trunc-workbook`, a cut-off `xl/workbook.xml`.

**Formulas.** It recalculated the formula golden's `excel-probe-4` group (94 formulas) and saved it.

**Separators.** Excel's AppleScript returned no decimal or thousands separators on this build. The region is en_NZ, with no Excel override. Nothing here depends on the separators: the product refuses text-to-number (decision 42).

**What is here, scrubbed per trap 43** (author, `x15ac:absPath` and the paths in repair logs blanked, by `scripts/scrub-excel-return.py`):
- `decoding/saved/`: Excel's saved workbooks, its repair logs, and the session's NOTES.txt. The saved `<v>` is the evidence (trap 42).
- `decoding/excel-cache.jsonl` and `import-report.txt`: `--excel-import`'s reading of them.
- `formulas/`: the saved probe workbook, its excel-cache.jsonl, the import report and the settings.

## How a repair counts

**A shape Excel repairs is DAMAGE.** Excel's repaired file shows its recovery, not a reading. The product refuses such a shape (an unreadable cell, or an error naming the part or the address) even where the repaired file shows a value.

**Per cell, in a repaired file:**
- a cell where Excel's saved value equals the rule's is decided (`excel:observed`);
- a cell where it differs, and every undecided cell, is refused (`excel:repaired`).

**The one exception** is a byte-identical repeated cell. Excel repairs it, removing the extra record, which leaves the same value. Reading it as one cell gives exactly what Excel shows, so it stays one cell (the 3 FUSE files' styled blanks).

## What Excel did with each shape

### Read cleanly (no repair): these decide

| shape | Excel |
|---|---|
| every r-less layout: mixed, all r-less, after a self-closing `<row/>`, a self-closing `<c/>`, after a gap, `x:`-prefixed, r-less formulas | positional, exactly as POS-1 and POS-2 say: every decided row agrees |
| two `<row r="2">` elements (`pos-row-repeated`) | **merged**: the second row's B2 = 3 is read beside A2 = 2. A repeated row continues that row; its cells must still ascend (otherwise POS-4/POS-5 apply) |
| `_x0009_` and `_x000A_` in element text | **decoded** to a tab and a line feed. [MS-OI29500] 2.1.1747 a says what Office WRITES; Excel reads them |
| phonetic runs | left out (VAL-1 holds) |
| UTF-16 LE and BE, with a BOM | read |
| **UTF-16 with no BOM**, declared `encoding="UTF-16"` | **read** from the declaration |
| **`encoding="windows-1252"`** (a sharedStrings `Caf\xe9`) | **read**: `Café` |
| **`encoding="ISO-8859-1"`** | **read** |
| a UTF-8 BOM | read |
| sharedStrings and styles at other names, absolute and `..` Targets, strict relationship types, a relocated workbook part | read through the relationships (PKG-1 holds) |
| the 1904 system in every form (`1`, `true`, prefixed); `0`/`false`/absent | as VAL-3 says: every decided row agrees |
| CDATA | read |
| **`<t>` with spaces at its ends and no `xml:space="preserve"`** (`ask-space`) | **TRIMMED**: leading and trailing spaces, tabs and line feeds are dropped, in shared and inline strings alike. `' lead'` reads `lead` (LEN 4), `'\n  x\n'` reads `x`, `'   '` reads `` (LEN 0) |

### Repaired: damage, refused

| shape | the repair log |
|---|---|
| two different cells at one address; an r-less cell landing on an explicit one | "Removed Records: Cell information" (Excel kept the first) |
| two byte-identical cells at one address, also backwards | the same; the value is unchanged (the exception above) |
| a cell out of order; a row out of order; a cell whose r names another row | "Removed Records: Cell information" |
| invalid UTF-8 in a `<v>`, an `<f>`, an inline string, sharedStrings | "Replaced Part ... with XML error" / "Removed Part" |
| a cut-off sheet, sharedStrings, styles or rels; mismatched tags | the part replaced or removed |
| a cut-off workbook part | Excel REFUSED to open the file |
| a sharedStrings relationship to a missing part; `t="s"` cells with no table; an index out of range or not an integer; a sharedStrings part with no relationship | "Removed Records: Cell information" |
| a styles part with no relationship | "Repaired Records" (the cells' `s` dropped: General) |
| a missing sheet part | "Removed Records: Worksheet properties" |
| `t="b"` holding `true`, `false`, `TRUE` or `2` (`val-b4-bools`) | "Repaired Records". Excel shows TRUE/FALSE for the words and text for the others. The file is damaged, so all stay refused |
| `<v>+5</v>`, `NaN`, `INF`, `-INF`, `1e400`, `inf`, `nan`, `1,5`, `0x2` (`val-b4-numbers`) | "Repaired Records". Excel turns `+5`, `NaN`, `INF` and `-INF` into TEXT, and `1e400` into #NULL!. **This contradicts B4's rules:** `+5` was read as 5, and `INF`/`NaN` as non-finite numbers. All become unreadable. The whitespace-padded numbers agree |
| `t="d"`, every form, in one workbook per date system | "Repaired Records". The decided forms all agree. The undecided ones stay refused, although Excel shows values for some (a `Z` zone, hour 24, more than three fraction digits, the basic form, a time alone) |
| error codes outside the documented eight (`#SPILL!`, `#CALC!`, `#FIELD!`, `#BLOCKED!`, `#UNKNOWN!`, `#BUSY!`, `#CONNECT!`, `#PYTHON!`, `#FOO!`), and `#n/a`, `#N/A ` | "Repaired Records". Excel turns the unknown codes into TEXT, but keeps `#BUSY!` as an error (ISERROR TRUE), and turns `#n/a` and `#N/A ` into #N/A. The file was repaired, so all stay refused |

**One spurious contradiction.** `pos-rless-formulas` C1, a `NOW()`, holds Excel's own recalculation, not the file's cache. A volatile formula's value is never compared.

## The questions asked as formulas

**`--period`: Excel's DAY, MONTH and YEAR round the serial to the nearest SECOND.** So does `TEXT(x,"yyyy-mm-dd")`.
- At 0.7 ms before midnight (`44592.999999991895`), DAY gives the next day and TEXT gives `2022-02-01`, while `TEXT(x,"…hh:mm:ss.000")` gives `23:59:59.999`.
- At 8.6 ms before the end of 9999-12-31 (`2958465.9999999`), DAY, MONTH and YEAR give #NUM! and the date TEXT gives #VALUE!. Past the last date, it is no date.
- HOUR rounds the same way. INT is the plain floor.

PER-1 is amended to round to the second, not the millisecond.

**The product's formatter disagrees.** A code that draws a date and no time (`yyyy-mm-dd`) rounds to the millisecond there: it draws `2022-01-31` at 0.7 ms before midnight, and `9999-12-31` for `2958465.9999999`. Codes with a time already round to the second. This is a formatter defect older than B5; the number-format golden gains these rows (`excel:observed`).

**1904:**
- serial 0 is 1904-01-01 (DAY 1), and 59 is 1904-02-29;
- 1462 is 1908-01-02;
- 40000.99999999994 is 2013-07-08.

A negative serial gives #NUM! from DAY, MONTH and YEAR, and TEXT draws `-1904-01-01`: still undecided for display, and no date for `--period`.

**A strict workbook** (`conformance="strict"`) draws like the compatibility system: 0 is `1900-01-00`, and 60 is `1900-02-29`. [MS-OI29500] 2.1.585 j says otherwise; Excel observed wins (decision 40).

**Overflow (NUM-1):**
- SUM, AVERAGE, STDEV and VAR over 9E+307, 9E+307, 1 all give #NUM!. AVERAGE over an overflowing sum is DECIDED: refuse.
- STDEV and VAR over 1E+200, -1E+200, 0 give #NUM! (the squares overflow), while SUM and AVERAGE give 0.

**IC-51:** `TEXT(x,"0.000000000000000")` shows 15 significant digits, then zeros:
- 1.9999992569016276 shows `1.999999256901630`;
- 1.5204998778130465 shows `1.520499877813050`;
- 0.47950012218695348 shows `0.479500122186953`.

The number-format golden gains the six rows.

**Formulas (`excel-probe-4`, 94 rows):**
- The 35 decided rows all agree, and the one Excel model matches all 35 bit for bit. LibreOffice's doubles, which the golden held, differ by 1 to 4 ulps in 11 of them.
- The 59 undecided rows Excel answered:
  - negative bases to negative fractional powers;
  - PMT at growth past the decided range;
  - 15-digit ties in ROUND and comparison: `100000000000000.5 = 100000000000001` is TRUE;
  - ROUND's digits past -20..42 (d=100 leaves x unchanged; d=-21 gives 0 for 1234.5678 and 1e300 for 1e300);
  - text-to-number (`$1,000`, `(5)`, `18:30`, `7%`).

  The product refuses all 59, which is safe. Teaching the model these is a deferred follow-up, not a B5 item.
