# What does Microsoft Excel read? Request 4: open and save 79 workbooks

**For:** a session with desktop Microsoft Excel, the same Mac as requests 1 to 3 (Excel for Mac 16.113.3, driven by AppleScript).
**From:** an internal spreadsheet tool's project, review batch B5 (2026-10-08).
**Time:** about 20 to 30 minutes, almost all of it Excel opening and saving files. Nothing is typed by hand.

Every workbook here is generated test data, with no private content.

## Why

The internal tool reads and writes `.xlsx` files for AI agents. It must read every cell as Excel does, or say that it cannot.

**The decoding part (`decoding/`, 78 workbooks).** Its new test, the "decoding golden", holds one small workbook per shape a reader can get wrong:
- cells and rows written without their address;
- Japanese phonetic guides;
- ISO dates stored as text (`t="d"`);
- the 1904 date system;
- text in UTF-16;
- string tables at unusual paths;
- damaged files.

The file format's standard decides most of these. For the rest, the tool refuses to answer until Excel has. The questions:
- **Cells without an address:** where does Excel put a cell or row written without one? What does it do with two cells at one address, a cell whose address names another row, and rows out of order?
- **ISO dates:** which spellings does it read in a `t="d"` cell, in a 1900 and a 1904 workbook? Microsoft documents only "a limited number", and does not list them.
- **Booleans and escapes:** does it read `true`/`false` in a boolean cell (Microsoft documents only 1 and 0)? An escaped tab?
- **Date functions:** what do DAY, MONTH, YEAR and TEXT give a hair before midnight, on the last day, and at serial 0, in each date system and in a strict workbook?
- **Rounding:** how does it round a lone 5 at the 16th digit?
- **Overflow:** does AVERAGE overflow when SUM does?
- **Unlisted errors and spaces:** what does it keep of an error code it does not list (such as `#SPILL!` stored in a file), and of spaces at the ends of text written without `xml:space="preserve"`?

**The formula part (`formulas/`, 1 workbook).** It holds 94 formulas the formula golden still has questions about. Request 3 answered every other curated formula.

## What to do

**0. Clear the download quarantine first**, or Excel opens the files read-only (Protected View) and never saves them:
   `xattr -dr com.apple.quarantine <this folder>`

### Part 1: `decoding/`

1. **Smoke test.** Open `decoding/files/pos-mixed.xlsx` and save it as in step 2.5 below. If that works, carry on.
2. **For EVERY workbook in `decoding/files/`**, one at a time:
   1. Open it.
   2. If Excel offers to **repair** it, accept, and save the repair log as `decoding/saved/<name>.repair.txt`.
   3. If it **refuses** to open it, or opens it in **Protected View**, write one line in `decoding/saved/NOTES.txt` (`<name>: refused` or `<name>: Protected View`), and go on to the next file.
   4. For the `ask-*.xlsx` files only, run a **full recalculation** (AppleScript `calculate full`).
   5. **Save it as .xlsx under the same name, in `decoding/saved/`** (Save As, Excel Workbook). This is a new file, not saved over the original.
   6. Close it.

   Many of these files are deliberately unusual or damaged, so a repair prompt or a refusal is itself an answer. Record it and move on.

### Part 2: `formulas/`

For `formulas/files/excel-probe-4.xlsx`:
1. open it;
2. run a full recalculation (`calculate full`);
3. **save it as .xlsx under the same name, in `formulas/files/`** (over the original, as in request 3);
4. close it.

Then, once, with Excel open: `sh formulas/excel-settings.sh > formulas/excel-settings.txt`. It only READS settings, as in request 3.

### Throughout

**Do not type into, sort or fill any cell.** Every answer is read from the saved files, at full precision.

## What to send back

1. `decoding/saved/`: every workbook, any `*.repair.txt`, and `NOTES.txt`.
2. `formulas/files/excel-probe-4.xlsx` as saved, and `formulas/excel-settings.txt`.
3. A short note: the Excel version, and anything surprising.

`manifest.json` in each part is our list of what each file asks; nothing in it needs checking.

**Not sent:** 4 shapes damaged at the zip level, which no application opens, so their answer is recorded as a refusal. Also 5 shapes the test builds at run time: decompression bombs and 1,000-sheet workbooks.

**On our side:**
- `scripts/decoding-golden.py --excel-import decoding` and `scripts/formula-golden.py --excel-import formulas` read every answer back.
- Before any returned workbook is kept, the person's name in `docProps/core.xml` and the folder path (`x15ac:absPath`) in `xl/workbook.xml` and in the repair logs are blanked.

Thank you.
