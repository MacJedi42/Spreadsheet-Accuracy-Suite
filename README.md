# Spreadsheet Accuracy Suite

**A test suite for anything that reads, computes, formats or writes Excel workbooks.** Its question is whether the tool gets the same answer Microsoft Excel does, on real-world files and on the edge cases that break implementations quietly.

This suite was built to prove internal tools: every number such a tool prints or writes must be the number Excel shows, or the tool must say it cannot be sure. **It is being made open source to help make sure [storytold/gridcraft](https://github.com/storytold/gridcraft) is accurate.** Anyone who needs to test and measure their own tools against this proof is welcome to use it.

> **Agents and automated testers: read [AGENTS.md](AGENTS.md) first.** It explains how to run each part against an implementation, how to score the results, and how to check the suite itself against Excel.

## The principle: Excel is the reference

A spreadsheet answer is right when it is what Excel shows. Every expectation in this suite names its source, and the sources rank in this order:

1. **Excel observed.** Desktop Microsoft Excel for Mac 16.113 was asked four times (`excel-evidence/request-1` to `request-4`), and Windows Excel (Microsoft 365, build 16.0.20430) re-ran requests 3 and 4 (`request-5-windows`). It recalculated or opened the suite's own workbooks and saved them, and the saved files are kept here. The value Excel wrote into each cell (`<v>`) is the evidence.
2. **Excel's own caches.** The cached values in real files that Excel last saved, such as the corpus workbooks.
3. **Microsoft's documentation:** ECMA-376 (Office Open XML) and Microsoft's implementation notes for it, [MS-OI29500].
4. **LibreOffice and openpyxl, only where they agree with Excel.** LibreOffice is a good second opinion, but it is not Excel. This suite records dozens of places where it differs (see [EXCEL-BEHAVIOUR.md](EXCEL-BEHAVIOUR.md)).

**Where no source decides, the expected outcome is a refusal** (`refuse`, `unreadable:*`, `undecided:*`). A tool that answers there is not proved wrong, but it is not proved right either. The suite never guesses an answer for Excel.

## What is in the suite, and what each part proves

### 1. `corpora/`: about 32,000 real-world workbooks (3.1 GB)

**What it proves:** a tool survives real files, not just tidy test files, and reads them as Excel does. Its sources are listed in [CORPORA.md](CORPORA.md).

| corpus | files | what it is |
|---|---|---|
| `corpora/spreadsheetbench/` | 5,458 .xlsx, 6 .xlsm, and `dataset.json` | SpreadsheetBench v0.1: 912 real questions from Excel help forums, each with its input and answer workbooks |
| `corpora/fuse/` | 10,702 .xlsx | the .xlsx part of FUSE, a corpus of spreadsheets crawled from the public web |
| `corpora/enron/` | 15,871 .xlsx, 58 .xls | the Enron spreadsheet corpus: business workbooks from the Enron email archive |

**What they are used for:**
- **Reading.** Every cell, decoded by an independent reader, against the tool under test. The internal tool this suite was built for reports about 16 million cells agreeing with 0 defects on such a run; that run's logs are not shipped here, so treat it as the tool's own claim.
- **Excel's own answers.** A workbook Excel last saved, fully calculated, carries Excel's computed value in every formula's cache, so a recomputation can be checked against it.
- **Rare shapes, found by scanning all 32,031 files:**
  - the 1904 date system: 273 files;
  - cells written without their `r` address: 3 files, 9,334 cells;
  - Japanese phonetic guides: 2 files;
  - `<t>` text with spaces at its ends and no `xml:space`: about 7,500 cells in about 100 files;
  - duplicate styled blanks: 3 files.

### 2. `goldens/formula/formula_golden.tsv`: 44,543 formulas with Excel's answers

**What it proves:** a formula engine computes what Excel computes, bit for bit where Excel was observed. It covers:
- addition and subtraction near zero (Excel's "snap" to zero);
- 15-significant-digit equality;
- ROUND at every tie;
- integer and fractional powers;
- PMT;
- logicals and text in arithmetic and comparisons;
- blanks;
- errors;
- percent;
- cross-sheet references.

**The rows:**
- **15,486 curated rows.** Request 3 is Excel's own answer for 15,046 of them, the ones Excel could be handed; basis `excel:observed` (14,942 rows).
- **29,057 corpus rows:** real formulas from the corpora, with their real inputs.

The file's header documents the row format, every rule and its source. The workbooks Excel recalculated for request 3, with Excel's answers cached in them, are in `excel-evidence/request-3/files/`, and `manifest.json` maps every row to its cell.

### 3. `goldens/number-format/number_format_golden.tsv`: about 190,000 display checks

**What it proves:** a formatter draws a number through a format code as Excel shows it. It covers:
- every format code found in the public corpora, plus curated codes (1,638 codes in all);
- the built-in formats;
- both date systems (1900 and 1904);
- the 15-significant-digit rule;
- clock rounding at the second;
- half-ties at the 15th digit;
- minus signs;
- AM/PM;
- elapsed time;
- fractions.

It has 181,897 checks of a value under a code, plus built-in id, date-system and Excel-answer sections. The 882 `X` rows hold Excel's TEXT() answers from request 4, and rows settled from those answers by the rule they showed.

### 4. `goldens/decoding/`: 89 workbooks, each one way a reader can go wrong

**What it proves:** a reader reads every cell Excel reads, where Excel reads it, and refuses what Excel calls damaged. The shapes:
- cells and rows without addresses;
- self-closing rows;
- duplicate cells and out-of-order rows;
- phonetic runs;
- `t="d"` ISO dates;
- the 1904 date system;
- UTF-16, windows-1252 and other declared encodings;
- string tables at unusual part names;
- truncated, checksum-damaged and not-well-formed parts;
- overflow;
- date bucketing at midnight.

`decoding_golden.tsv` lists 787 checks: cell values, whole-file refusals, sheet facts, queries, writes and resource bounds. Excel opened every shape it could (request 4); its saved copies and repair logs are in `excel-evidence/request-4/decoding/saved/`.

### 5. `excel-evidence/`: what Excel itself did

| request | when | what Excel was asked | what it settled |
|---|---|---|---|
| `request-1` | 2026-10-02 | LOOKUP's result extension, the snap to zero, ROUND, small cases, typed into workbooks | the shapes behind the formula golden's first Excel rules |
| `request-2` | 2026-10-02 | the display and date classes where two independent readers disagreed on real files | which reading Excel shows, class by class (`ANALYSIS.md`) |
| `request-3` | 2026-10-02 | recalculate every curated formula-golden row (52 workbooks) and LOOKUP shapes | Excel's exact answer for 15,046 formulas; snapping, powers, equality, ROUND digits, PMT, logical ordering, text comparison (`SETTLED.md`) |
| `request-4` | 2026-10-08 | open 78 decoding shapes and save them, and answer ask workbooks of DAY, MONTH, YEAR, TEXT, SUM, AVERAGE, STDEV and error codes; recalculate 94 more formulas | which file shapes Excel reads, repairs or refuses; text trimming; encodings; rounding at the second; overflow; strict workbooks (`SETTLED.md`) |
| `request-5-windows` | 2026-10-09 | requests 3 and 4 again, on **Windows** Excel: the same original workbooks | Windows matches Mac on 15,045 of 15,046 formulas, all 94 probe formulas, 795 of 797 decoded cells and every damaged-file verdict. One platform difference: `(-8)^(-1/3)` is -0.5000000000000001 on Mac and -0.5 on Windows (`SUMMARY.md`) |

**The run's details:**
- Each request folder keeps the request as sent (`REQUEST.md`), Excel's saved workbooks, its repair logs, and the import of its answers (`excel-cache.jsonl`).
- Requests 1 to 4: Excel for Mac 16.113.3 or 16.113.4 on macOS 27, locale en_NZ. Request 5: Windows Excel 16.0.20430 on Windows 11 Pro, locale en-NZ.
- The person's local folder paths are blanked from the saved files.

### 6. `goldens/fidelity/`: three workbooks with charts, pivots, a table, macros and defined names

A tool that edits a workbook must leave these intact. They come from other open-source projects' test suites; see `goldens/fidelity/PROVENANCE.md`.

### 7. `probes/`: 31 small hostile workbooks

These are the first shapes that exposed reader defects:
- 1904 dates;
- cells without addresses;
- duplicates;
- ISO dates;
- phonetic runs;
- UTF-16;
- NaN and INF;
- invalid UTF-8;
- truncation;
- a misplaced string table;
- a 4,000,000-entry string table (a decompression bomb in miniature);
- 1,000 sheets;
- midnight edges.

Each tests one failure that went silently wrong in a real implementation.

### 8. [EXCEL-BEHAVIOUR.md](EXCEL-BEHAVIOUR.md)

The behaviours this suite established, each with its source. It also lists where LibreOffice differs from Excel.

## How these findings were made

The suite grew out of a code review of an internal spreadsheet tool, in batches, between September and October 2026. Each batch built its test gate first, from Excel's documented behaviour and from LibreOffice where it agreed, and then measured the tool against it.

**Where no reference decided an answer, Excel was asked directly.** Excel opened, recalculated and saved the suite's own workbooks (requests 1 to 4), so each answer is Excel's own saved value, not a hand-typed one. The goldens are generated independently of the tool they measure, and share no code with it. The internal project also cross-checked the decoding golden with a second, independent decoder, which agreed on every row they share; that decoder is not shipped here.

**The corpora are public research datasets** (see [CORPORA.md](CORPORA.md)). No private or customer workbook is in this suite.

## Reading the header comments

The goldens' header comments and the Excel evidence write-ups were written during the internal project's review, and keep some of its references:
- **`docs/review/2026-10-02-excel-*` and `docs/review/2026-10-08-excel-*`** are this repo's `excel-evidence/request-1` to `request-4`.
- **`HANDOFF`, "decision N", "trap N", "gate round N" and "batch B3/B4/B5"** point to the internal project's own design notes, which are not published. The text around each says what was decided and why. The evidence it rests on is in `excel-evidence/`.
- **"the product", "the internal tool" and `<internal tool>`** mean the internal tool the suite was built to prove.

## Licence

The suite's own files, the goldens and documentation, are under this repository's `LICENSE`. The corpora and the fidelity workbooks keep their original terms (see `CORPORA.md` and `goldens/fidelity/PROVENANCE.md`).
