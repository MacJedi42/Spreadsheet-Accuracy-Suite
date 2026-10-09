# Excel request 3: results

Run on 2026-10-02, Excel for Mac 16.113.3, macOS 27.0.1, locale en_NZ. Driven by AppleScript.

## What was done
- Cleared the download quarantine (`xattr -dr com.apple.quarantine`) before opening anything.
- Smoke test: `files/division.xlsx` opened cleanly, with no repair prompt.
- All 52 workbooks in `files/` (51 from the manifest plus `lookup-shapes.xlsx`) were opened, run through `calculate full`, saved as .xlsx under the same name, and closed. Calculation stayed Automatic. No cell was typed into, sorted or filled.
- Display alerts were left on, so repair prompts appeared and the user answered them by hand. Every save succeeded.
- Every workbook in `files/` was re-saved by Excel (none is byte-identical to the original).
- Originals were backed up before the run (outside this folder, not included).

## Repair prompts (5 workbooks)
Excel showed "We found a problem with some content in '<name>'. Do you want us to try to recover as much as we can?" The user accepted (Yes) each time. Each log is in `repair/` and reads, for all five: **"Removed Records: Formula from /xl/worksheets/sheet1.xml part"**.

| Workbook | What Excel did to the saved file |
|---|---|
| `power.xlsx` | Dropped the formula and value of `Z27` (`=1^1E+308`) and `Z30` (`=1E+308*10`); both rows are now empty. The other 33 formulas were kept and recalculated. |
| `risky-hostile-00006.xlsx` | Formula dropped (`=` followed by 1000 nested parentheses). The row is now empty. |
| `risky-out-of-range-00010.xlsx` | Formula dropped (`=IF(<400-digit literal>>0,1,2)`). The row is now empty. |
| `risky-out-of-range-00011.xlsx` | Formula dropped (333-char `0.000...1+1`). The row is now empty. |
| `risky-out-of-range-00012.xlsx` | Formula dropped (323-char `0.000...1*1`). The row is now empty. |

Reading: Excel's file loader refuses a formula it cannot parse or hold, and removes it from the file rather than erroring in the cell. There is no cached value to read for these cells; the answer is "formula refused/removed". The `<name>.repair.txt` notes that were first written to `files/` (if still there) say no log text was captured; the real logs are the ones in `repair/`.

Every other `risky-*` workbook opened with no prompt, so Excel accepted them. That includes:
- `risky-hostile-00007` to `-00012` (nested parentheses at 100, 64, 90 and 250 deep, 200 unary minuses, 4001 tokens), apart from `-00006` above;
- all of `risky-percent-*` (stacked `%%`);
- `risky-out-of-range-00001` to `-00009` (subnormal inputs).

Read their saved formula text and values with `--excel-import` to see how Excel rewrote them.

## Settings (`excel-settings.txt`)
- The script's `osascript` line failed (`use system separators` is not an AppleScript property Excel accepts), so that value was not read.
- Added by hand to the file: Excel evaluates `TEXT(1234.5,"#,##0.00")` as `1,234.50`, and `VALUE("1,000.5")` as `1000.5`. So the decimal separator is "." and the thousands separator is ",".
- Excel's own languages: not set (it follows the system). macOS is `en_NZ`.
- The "Use system separators" checkbox itself was not checked in Preferences.

## Not done
- Nothing has been compared against the internal tool's golden. `formula-golden.py --excel-import` has not been run.
- No surprises beyond the five repairs above were seen, but nobody inspected the saved formula text of the other files, so Excel may have rewritten formulas without anyone noticing.

## Send back
`files/` (all workbooks), `excel-settings.txt`, `repair/`, and this file.
