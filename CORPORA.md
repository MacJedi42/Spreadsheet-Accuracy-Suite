# The corpora

The suite's real-world workbooks are three public research corpora, copied as they were obtained, with no file edited. Each keeps its original terms; check them before reusing a file outside testing.

| folder | files | source |
|---|---|---|
| `corpora/spreadsheetbench/` | 5,458 .xlsx, 6 .xlsm, `all_data_912_v0.1/dataset.json` | **SpreadsheetBench v0.1**: Zeyao Ma et al., "SpreadsheetBench: Towards Challenging Real World Spreadsheet Manipulation", NeurIPS 2024 Datasets and Benchmarks track. 912 instructions collected from Excel help forums, each with input and answer workbooks. Repository: https://github.com/RUCKBReasoning/SpreadsheetBench |
| `corpora/fuse/` | 10,702 .xlsx | **FUSE**: Titus Barik, Kevin Lubick, Justin Smith, John Slankas and Emerson Murphy-Hill, "FUSE: A Reproducible, Extendable, Internet-scale Corpus of Spreadsheets", MSR 2015. Spreadsheets crawled from the public web, hosted by the OpenScience tera-PROMISE repository (`FUSE-README.txt`; http://openscience.us/repo/spreadsheet/fuse). Only the .xlsx part is here, under the corpus's own file identifiers. |
| `corpora/enron/` | 15,871 .xlsx, 58 .xls | **The Enron spreadsheet corpus**: Felienne Hermans and Emerson Murphy-Hill, "Enron's Spreadsheets and Related Emails: A Dataset and Analysis", ICSE 2015 (SEIP). Spreadsheets attached to email in the public Enron archive. File names are `<employee>__<id>__<original name>`. |

## Why these three

- **SpreadsheetBench** is what people ask Excel to do: lookups, conditional sums, dates, text and reshaping, in workbooks real users posted.
- **FUSE** is the open web, with every producer and every oddity. It holds:
  - files from tools other than Excel (the cells without `r` addresses are Open XML SDK output);
  - Japanese workbooks with phonetic guides;
  - 240 workbooks in the 1904 date system (across all three corpora, 150 of the 273 name Mac Excel as their producer).
- **Enron** is one company's working spreadsheets: real business models with deep formula chains, shared formulas, defined names and VBA-module listings.

## What the suite measured in them

These figures come from a byte scan of every part of all 32,031 .xlsx workbooks, and a decode of samples.

| shape | files |
|---|---|
| `workbookPr date1904="1"` (the 1904 date system) | 273 (FUSE 240, Enron 33) |
| cells with no `r` address | 3 (FUSE), 9,334 cells |
| phonetic runs (`<rPh>`) with text | 2 (FUSE) |
| `<t>` with whitespace at an end and no `xml:space` | about 100 (about 7,500 cells) |
| duplicate styled blank cells (byte-identical) | 3 (FUSE) |
| a VBA module listed as a sheet (`r:id=""`) | about 400 (Enron) |
| UTF-16 parts | 89 (SpreadsheetBench, custom XML only) |
| `t="d"` ISO date cells, string tables at other names, declared non-UTF encodings | 0 |

**The formula golden's corpus rows** are 29,057 formulas taken from these files with their real inputs: Enron 21,592, FUSE 7,244, SpreadsheetBench 221. Each row's id names its corpus and the first 10 hex digits of the SHA-1 of the file's path relative to that corpus's root.

## Using them with Excel's own answers

A formula's cached value in a file Excel last saved is Excel's answer. AGENTS.md (section D) gives the checks a file must pass before its caches count as evidence:
- saved by Excel;
- calculated automatically and completely;
- no precision-as-displayed;
- no volatile functions;
- no external links;
- no Lotus evaluation.
