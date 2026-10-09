> **Record copy.** The original is `excel-request-3/REQUEST.md`, beside 52 workbooks that `scripts/formula-golden.py --excel-export` writes from the golden (branch `gate/excel-export`, `4114af7`), plus `scripts/excel-lookup-shapes.py`. Not committed: the workbooks and the importer manifest (`excel-request-3-import/manifest.json`, 7 MB), which are regenerable. Answers are read with `--excel-import`.

# What does Microsoft Excel compute? Request 3: let Excel recalculate 52 workbooks

**For:** a session with desktop Microsoft Excel, the same Mac as requests 1 and 2 (Excel for Mac 16.113.3, driven by AppleScript).
**From:** an internal spreadsheet tool's project, review batch B4 (2026-10-02).
**Time:** about 20 to 40 minutes, almost all of it Excel opening and saving files. Nothing is typed by hand.

## Why

The internal tool writes cached values into `.xlsx` files, so it must compute formulas exactly as Excel does, to the last bit, or refuse. Its test suite (the "formula golden") holds about 15,000 small test formulas. Each one's expected answer comes from LibreOffice plus four models of Excel's arithmetic.

Requests 1 and 2 had Excel answer questions one cell at a time. This request hands Excel the test formulas themselves, and the answers will:
- settle the 473 formulas where nothing tells us what Excel does (the tool refuses those today);
- check every other formula against real Excel;
- settle the open questions:
  - does `=(A-B)` snap a tiny result to 0 like `=A-B`?
  - how does TRUE order against numbers?
  - does Excel compute a whole-number power by repeated squaring?
  - what do 2-D LOOKUP ranges read?
  - how is text such as "1,000.5" read as a number?

The `files/` workbooks are generated test data, with no private content.

## What to do

The full instructions are in **`manifest-summary.txt`**, under "WHAT TO DO". In short:

0. **Clear the download quarantine first**, or Excel opens the files read-only (Protected View) and never recalculates or saves them:
   `xattr -dr com.apple.quarantine <this folder>`
1. **Smoke test.** Open `files/division.xlsx` first.
   - If Excel offers to repair it, stop and send back its repair log.
   - If it opens cleanly, carry on.
2. **For EVERY workbook in `files/`**, including `lookup-shapes.xlsx`, one at a time:
   1. open it;
   2. run a **full recalculation** (AppleScript `calculate full`);
   3. **save it as .xlsx under the same name**, in `files/`;
   4. close it.

   Calculation can stay Automatic.
3. **Once:** `sh excel-settings.sh > excel-settings.txt`, with Excel open. It only READS settings: Excel's version, its decimal and thousands separators, its language, and the Mac's region. How Excel turns text into numbers depends on these.
4. **If Excel offers to repair a workbook**, accept, save its repair log beside it as `<name>.repair.txt`, and continue. The `risky-*.xlsx` files hold one deliberately hostile formula each, such as a 400-digit number or parentheses nested 1,000 deep, so a repair there is itself an answer.
5. **Do not type into, sort or fill any cell.** Every answer is read from the saved files, at full precision.

## What to send back

1. The whole `files/` folder, with every workbook saved by Excel.
2. `excel-settings.txt`.
3. Any `*.repair.txt` logs.
4. A short note: the Excel version, and anything surprising (a prompt, a file that would not open or save, a formula Excel rewrote).

That's all. A script on our side (`formula-golden.py --excel-import`) reads every answer back. Thank you.

(`lookup-shapes-libreoffice.txt` is our reference, not something to check. It says what LibreOffice computes for the LOOKUP shapes, for comparison with Excel's answers.)
