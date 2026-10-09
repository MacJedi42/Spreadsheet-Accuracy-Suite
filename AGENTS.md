# AGENTS.md: how to use the Spreadsheet Accuracy Suite

You are probably here to answer one question: does this spreadsheet tool get the same answers Microsoft Excel does? This file tells you how to answer it with the suite, how to report the result honestly, and how to check the suite itself.

Read [README.md](README.md) for what the suite contains and [EXCEL-BEHAVIOUR.md](EXCEL-BEHAVIOUR.md) for the rules it holds.

## The rules that come before anything else

1. **Excel is the reference. Mac or Windows desktop Excel is the authority.** Most of the world's spreadsheets are Excel files, outside Google Sheets. When the user wants to know whether the suite itself is right, prioritise a check against real Microsoft Excel, on Mac or Windows (see "Checking the suite against Excel" below).
2. **LibreOffice is decent, but not 100% in line with Excel.** It agrees with Excel almost everywhere, and the suite uses it as a cross-check, but EXCEL-BEHAVIOUR.md §4 lists dozens of places it differs: near-zero arithmetic, booleans in SUM, powers, the 1900 epoch, clock rounding, 16-digit display, and more.
   - Never "correct" a golden row to match LibreOffice.
   - If LibreOffice and the suite disagree, the suite's `basis` column says where its answer came from. Only Excel can overrule it.
3. **openpyxl, pandas and other libraries are not references either.** openpyxl rejects files Excel opens, and reads some things Excel repairs.
4. **Never edit a golden to make a tool pass.** A golden changes only on new Excel evidence, cited in the file.
5. **A refusal row is not a failure:** `refuse`, `unreadable:*`, `undecided:*`, `excel:repaired`.
   - It means no source decided the answer, or Excel itself calls the input damaged.
   - A tool that refuses or reports damage there passes.
   - A tool that answers is not proved wrong, but report it separately as "answered an undecided case".
6. **Report what you measured, with evidence:** counts per suite, every failure with its file, cell and expected-versus-got values, and the commands you ran. Do not round "a few failures" down to "passes".

## Know which rows test Excel and which test a tool's own choices

The goldens were built for an internal tool, and a few rows encode that tool's own presentation or policy choices rather than Excel's behaviour. **Exclude them when you judge Excel fidelity:**

| file | rows to exclude | why |
|---|---|---|
| `goldens/number-format/number_format_golden.tsv` | any `basis` containing `tool:`, alone or inside a compound such as `1904:tool:iso` or `settled:X9+tool:iso` (85 rows: 54 `I`, 28 `DI`, 3 `X`) | the internal tool draws built-in date formats 14 and 22 in ISO form (2020-01-31); Excel draws them per locale |
| `goldens/decoding/decoding_golden.tsv` | `basis` starting `b1:`, `b4:` or `r8:`; and kind `T` | the internal tool's own policies: its `--period` refuses serials below 61 (`b1:floor`); it refuses a part that inflates past 512 MiB (`b4:decision-33`, the decompression-bomb shapes); its time limits on 1,000- and 2,000-sheet books (`T` rows, `r8:zip-lookup`) |
| `goldens/decoding/decoding_golden.tsv` | kinds `Q` and `W` | these spell the internal tool's own commands (`aggregate|A|sum|1:5`, `set-cell`); their EXPECTED VALUES are Excel's, so translate the query to the tool under test, or skip them |
| `goldens/formula/formula_golden.tsv` | the `must` column | the internal tool's list of rows it may not refuse; ignore it |

Everything else (`excel:*`, `ecma:*`, `opc:*`, `xml:*`, `lo`, `error:*`, `settled:*`, `undecided:*`) is about Excel, ECMA-376 or the file format.

## Testing a tool, suite by suite

Write a small ADAPTER for the tool under test: a script that asks the tool for a cell's value, a formula's result or a formatted string. Keep the adapter separate from the suite. The suite ships data, not code tied to any one tool.

### A. Formulas: does the tool compute what Excel computes?

**The best test is `excel-evidence/request-3/files/`**: 52 workbooks that Excel itself recalculated and saved.
- Each formula sits in column Z of sheet `S`, beside its inputs, and each cell's cached `<v>` is Excel's own answer.
- `excel-evidence/request-3/manifest.json` maps each formula-golden workbook to its rows: id, cell, formula, inputs, and the golden's expectation and basis, taken from the golden. When they ever differ, `goldens/formula/formula_golden.tsv` is the authority, by row id.
- **The one exception is `lookup-shapes.xlsx`.** It has no manifest entry; `lookup-shapes.json` describes it. Its formulas are on `Sheet1` in column H, driven by defined names on sheet `S`, and Excel's answers are its own cached values.

**The procedure:**
1. Copy a workbook, and remove every formula cell's cached value (`<v>`) and its `t` attribute, so the tool cannot simply read Excel's answer back.
2. Have the tool compute the workbook.
3. Compare each formula cell with Excel's value, read from the ORIGINAL file's `<v>` at full precision, or with `excel-evidence/request-3/excel-cache.jsonl`.

**What counts as a match:**
- **numbers:** the identical double for `excel:observed` rows. Count within-1e-13 matches separately: Excel's PMT is not reproducible past about 11 units of 2^-52;
- **logicals:** the same TRUE/FALSE, as a logical (`t="b"`), not the number 1;
- **errors:** the same error code.

**Then the full golden, `goldens/formula/formula_golden.tsv`** (44,543 rows, including 29,057 real formulas from the corpora). Each `F` row is:

```
F  <id>  <home sheet>  <formula>  <inputs>  <expect>  <basis>  <must>
```

- `inputs` is `Sheet!A1=n:<number>` / `b:TRUE` / `s:<text>` / `e:<#ERR>`, separated by the character U+001F (ASCII unit separator, which no formula can contain); a cell not listed is blank.
- `expect` is `n:<value>`, `b:TRUE|FALSE`, or `refuse`.

Build one small workbook per row, or many rows per workbook in separate cells: inputs at their addresses, the formula at the cell after `@` in the id. Compute it and compare. The file's header documents the rules, the tolerances and every left-out class.

**Pitfalls the golden is built to catch:**
- `=0.1+0.2=0.3` is TRUE in Excel (15-digit equality);
- `=A1-B1` over nearly equal numbers snaps to 0 only at the bare root;
- `=-2^2` is 4;
- `=SUM(range)` skips logicals;
- `=TRUE>1E300` is TRUE;
- integer powers come from repeated squaring;
- `ROUND` works on the 15-digit decimal.

### B. Reading: does the tool read every cell Excel reads?

**`goldens/decoding/`** holds 89 workbooks and `decoding_golden.tsv`.

- **Cell rows (`C`):** open the workbook named by the shape's `X` row, read the cell, and compare with `expect`:
  - `n:<number>`, `s:<text>`, `b:TRUE|FALSE`, `e:<#CODE>` or `blank`;
  - `date:<iso>@<serial>`: the serial is the number, in the workbook's date system;
  - `unreadable:<kind>`: the tool must not give a number, text or blank here; reporting the cell unreadable or damaged passes.
- **File rows (`F`):** `F <shape> refuse:<part>` means the workbook is damaged at that part, and the tool must REFUSE it with an error. Reading it as empty, partial or complete is a failure, the worst kind: a silent wrong answer.
- **Sheet rows (`S`):** the used range and the non-blank count.

**Excel's evidence is in `excel-evidence/request-4/decoding/saved/`.**
- It holds Excel's own saved copy of each shape, its repair log (`*.repair.txt`) where Excel repaired the file, and `NOTES.txt`.
- A shape Excel had to repair is damage: refusing it is right, even where Excel's repaired copy shows a value.
- **Where a row's evidence lives depends on its basis:**
  - `excel:observed` or `excel:repaired`: `saved/<shape>.xlsx`, plus its repair log;
  - `excel:refused`: no file; `NOTES.txt` records the refusal;
  - `excel:asked`: `excel-cache.jsonl`, from one of the `ask-*.xlsx` workbooks, not the shape's own file;
  - `ecma:*`, `opc:*`, `xml:*`, `settled:*` and `undecided:*`: the cited clause or rule. Some of those shapes were never sent to Excel, so they have no saved copy.
- **The fields of each row kind**, from the TSV header:
  - `C <shape> <sheet> <cell> <expect> <shown> <raw> <basis> <must> <rule>`: `shown` is the text a grid should show (`-` means not judged); `raw` is the value as the file spells it, which a tool may show for an unreadable value;
  - `S <shape> <sheet> <fact> <expect> <basis> <must> <rule>`: the fields are `used` (used range), `nonblank` (non-blank count), `view` (whether the sheet holds content), `uncached` and `volatile` (formulas with no readable cache, and volatile formulas), and `clear` (whether a recalculate-on-load flag may be cleared: an internal-tool command, so translate it or skip it).

**Five shapes are built at run time, not shipped:**
- decompression bombs: a 513 MiB string table or styles part;
- 1,000- and 2,000-sheet books;
- a sheet with huge neighbours.

Their `X` rows and parameters are in the TSV; build them if you test resource behaviour.

**`probes/`:** 31 smaller hostile workbooks. Each should read as EXCEL-BEHAVIOUR.md §2 says, or be refused as damage, never as empty or wrong.

### C. Display: does the tool show what Excel shows?

**`goldens/number-format/number_format_golden.tsv`:**

```
C <code index> <files using it> <format code>
R <code index> <value> <expected text> <basis>
```

- Render each `R` row's value through its code and compare the text exactly.
- **Render under the en-US locale,** as the golden was generated (`lo` is LibreOffice in en-US; `lo-en` forced en-US). Any other locale mismatches on every code with a month or weekday name.
- `I`/`D`/`DI` rows hold built-in ids, the 1904 date system, and dates by id. The header explains each.
- `X` rows hold Excel's own TEXT() answers and the rows settled from them.
- Exclude `tool:*` rows (see above).

**The pitfalls:**
- at most 15 significant digits, then zeros;
- clock times rounded to the second, not truncated;
- serial 60 is 1900-02-29;
- in the 1904 system, serial 0 is 1904-01-01;
- `[h]:mm` elapsed time;
- minus signs on values that round to zero.

### D. Real files: `corpora/`

**1. Read every cell.** Read every cell of a sample, or of everything, and compare with an independent reader you write to EXCEL-BEHAVIOUR.md §2. A plain zip-plus-XML reader in a few hundred lines is enough. Three cautions:
- Find parts through relationships.
- Handle cells with no `r`.
- Do not trust a library that skips what it does not understand.

The internal tool's own run read about 16 million cells this way with 0 disagreements.

**2. Recompute against Excel's caches.** In a workbook Excel last saved, each formula's cached `<v>` is Excel's answer. Use it only when ALL of these hold:
- `docProps/app.xml` names Microsoft Excel as the application. openpyxl also writes "Microsoft Excel" there, so check `xl/workbook.xml` too: Excel writes a `calcPr` with a plausible `calcId`, a nonzero Excel build number of 5 or 6 digits. Verify a `calcId` of 0, or any odd value, with an independent recompute before trusting the cache;
- `calcPr` has no `calcMode="manual"`, no `fullCalcOnLoad="1"`, and no `fullPrecision="0"` (precision as displayed);
- the formula is not volatile (NOW, TODAY, RAND, OFFSET, INDIRECT, CELL, INFO) and reads no external link;
- the sheet has no `sheetPr transitionEvaluation="1"` (Lotus rules).

**Do not count these as defects:**
- a cache that disagrees with a recomputation in a file that fails one of the checks above;
- older Excel versions (`calcId` ≤ 152511) computed some powers with x87 double rounding.

**3. Prevalence matters.** Use the rare shapes (README §1) to show a tool handles what real files contain, not just what test files contain.

### E. Writing: does a write leave the file valid and honest?

If the tool writes:
- **Every edited workbook** must open in Excel without a repair prompt.
- **Every part the tool did not mean to change** must be byte-identical, `goldens/fidelity/` included: charts, pivots, a table, VBA and defined names.
- **Every cached value it leaves** must be right, or the workbook must ask for recalculation (`calcPr fullCalcOnLoad="1"`).

To check the caches, strip them from a copy, have Excel (or LibreOffice, as a second opinion) recompute, and compare.

## Scoring and reporting

**For each suite, report:**
- the rows run;
- pass;
- FAIL (a wrong value, a wrong type, or a damaged file read as whole);
- refused a decided row (allowed, but list them);
- answered an undecided row (list them);
- excluded rows, by reason.

**Lead with FAILs.** Give each one's file, cell or row id, the expected value, what the tool gave, and the row's basis, so the user can see how strong the evidence is: `excel:observed` is strongest, then `excel:*`, `ecma:*`, and `lo`.

**Wording:** say "matches Excel on N of M decided rows". Never say "100% accurate" unless every decided row in every suite passed.

## Checking the suite against Excel (when the user wants to verify the suite itself)

**Prioritise real Microsoft Excel, on Mac or Windows.** LibreOffice can screen for candidates, but it cannot confirm a row (rule 2). How this suite asked Excel, which you can repeat:

1. **Hand Excel the suite's own workbooks, never formulas typed by hand.** Typing goes wrong: `=-10^16` is +1E16 in Excel. Use `excel-evidence/request-3/files/` or build workbooks the same way.
2. **Clear the macOS download quarantine,** or Excel opens the files read-only in Protected View and never recalculates: `xattr -dr com.apple.quarantine <folder>`.
3. **For each workbook:** open it, recalculate fully, save as .xlsx, and close it.
   - **On Mac:** AppleScript `calculate full`.
   - **On Windows:** VBA `Application.CalculateFull`, or the COM object from PowerShell.
   - **If Excel offers a repair:** accept it and keep the repair log. A repair is itself an answer: Excel calls that input damaged.
4. **Read the answers from the SAVED file's XML, at full precision:** the `<v>` of each cell.
   - Do not read displayed text or AppleScript values, which are cut to 15 digits.
   - Do not read a LibreOffice re-save, which stores 15 digits.
5. **Record the Excel version and build, the OS, the region and locale, and the decimal and thousands separators.** Text-to-number conversion and date reading depend on them.
6. **Before sharing Excel's saved files, blank anything personal:**
   - `docProps/core.xml` creator and lastModifiedBy;
   - `xl/workbook.xml` `x15ac:absPath`, a local folder path that can include a user name;
   - any paths in repair logs.
7. **Where Excel disagrees with a golden row,** Excel wins. Correct the row, cite the evidence in the file (the request, the workbook and cell), and keep the evidence in `excel-evidence/`.

**Windows and Mac Excel** should agree on everything here. If they ever disagree, record both, and treat the row as platform-dependent rather than picking one.

## Contributing evidence

New rows need a source. Add Excel's evidence under `excel-evidence/request-N/`, with the request as sent, the saved workbooks, any repair logs, and a `SETTLED.md` saying what it settles. Update the golden and its header's basis counts, and say which rows changed and why.
