# What Excel's answers to request 2 settle

**The request:** `../2026-10-02-excel-unexplained-request.md`.
**The answers:** `RESULTS.md`, and `partB.xlsx` (Microsoft Excel for Mac 16.113.3, 2026-10-02, driven by AppleScript). The workbook's author fields and its embedded local path are blanked.
**The question behind it:** which reader is right in each class of the read differential's 6,263 unexplained disagreements (the stream reader against IronCalc's model, over SpreadsheetBench). Cite as "Excel observed (16.113.3, 2026-10-02, request 2)", with the item.

## Verdicts, class by class

| class | cells | Excel observed | right | how the harness can prove it per cell |
|---|---|---|---|---|
| is_error: an array formula over names (57716), `DATEVALUE` (CF_18814), `TEXT(…,"DDD")` (11276) | 2,979 | A5: each stored value survives a full recalc unchanged | stream (the file's cache) | the model's error against the stream's value, where the stream equals the file's own cache, array members included |
| error_cells (per-sheet counts, 11276 and 57716) | 17 | follows from A5 | stream | the sheet's count difference is exactly the cells excused above |
| formula text | 2,947 | A3: AF14 is AE14's text moved one column (S14, AF$9) | stream | the stream equals the file's own `<f>` text, with a shared formula translated by the harness's OWN translator (trap 37), not the product's |
| display: `_xHHHH_` escapes (545-35) | 180 | A4: Excel decodes them (`LEN` 3, `CODE` 30 and 6) | stream | the harness decodes the raw string itself and gets the stream's text |
| display: day 0 under `d` (118-8) | 42 | A2 and B2: `0` on a sheet that shows zeros | stream (but see N1) | Excel's B2 table, as curated rows of the number-format golden |
| display: just below a .5 tie (55427, 32468, 50442, 49613) | 12 | A1 and B1: rounded UP, every one | stream | the stored value's 15-digit decimal is an exact tie, and the stream rounded it half away |
| aggregate | 86 | A6: `MAX` ignores an empty cell; the rest follow from A5 | stream | the column's difference is the excused cells, or an empty cell the model counted as 0 |

**The internal tool is right in every class.** Nothing in the 6,263 is a reader defect in the internal tool, and IronCalc is wrong in each. Getting the number to 0 is harness work: each class needs a per-cell proof the harness makes itself, not an excuse for the whole class.

## Part B, checked against the internal tool

`<internal tool> get-range partB.xlsx Sheet1 A1:N8` (B4 round-4 binary) against Excel's display:

- **B1, all 8 rows match:** 15%, -14,705,084, 1,322,198, -2,252, 1, 3, 1.01 and 29%.
- **A correction to RESULTS.md's caveat:** it says rows 5 and 6 may have been snapped to exactly 0.5 and 2.5. They were not. The saved file holds `<v>0.49999999999999994</v>` and `<v>2.4999999999999996</v>`; the 0.5 and 2.5 came from AppleScript reading only 15 digits. So Excel shows a value one ulp below a .5 tie as rounded UP: it rounds the 15-digit decimal, not the double. That is the internal tool's rule.
- **B2, 56 of the 60 cells match.** That includes day 0 (`0`, `00`, Saturday, `1/0/1900`, `Saturday 0 January 1900`), serial 0.5, the fictitious 29 Feb 1900 (serial 60, a Wednesday), and serial 59 as a Tuesday.
- **The 4 that differ are all in column M**, format `d-mmm-yy`, at serials 0, 0.5, 1 and 61, where the day has one digit. Excel saved it as built-in id 15, and drew it with a two-digit day (`00-Jan-00`, `01-Mar-00`), while the internal tool draws `0-Jan-00` and `1-Mar-00`. A5 shows the same for built-in id 16 (`d-mmm`): CF_18814's E1 shows `01-Jun` in Excel and `1-Jun` in the internal tool. See N2.

## New items for the review log

- **N1 (low, deferred). A sheet's `showZeros="0"` is ignored.**
  - Excel hides every zero on such a sheet (A2: 118-8's Schedule H4 and P28 are blank in Excel). The internal tool prints `0`.
  - The value itself is right, and the grid's rule, "a bare value is a NUMBER", argues for keeping it.
  - The proposal is to say the sheet hides zeros (a sheet-level note) rather than blank the cell. That is the user's call.
- **N2 (low, deferred). Built-in ids 15 and 16 depend on the locale, like 14 and 22 (decision 39).**
  - Excel for Mac in the user's (NZ) locale draws 15 as `dd-mmm-yy` and 16 as `dd-mmm`. ECMA-376's codes, which the internal tool uses, are `d-mmm-yy` and `d-mmm`; en-US Excel shows those.
  - The month is a name, so nothing is ambiguous; only the leading zero differs.
  - The proposal is to keep ECMA's codes and say so in decision 39's note. Id 17 (`mmm-yy`) has no day, so it is not affected.
- **N3 (medium, open, cluster T3). The differential cannot break its own number down.** The fix is the per-cell proofs in the table above:
  - formula text from the file, with an independent shared-formula translator;
  - error cells against the file's cache, array members included;
  - `_xHHHH_` decoding;
  - the 15-digit tie;
  - Excel's B2 table for serials 0 to 61;
  - aggregates that the excused cells and empty cells account for.

  After them, any disagreement left is new and specific.

## For the number-format golden (B3's gate)

B3 left out "a calendar code over serials 0–1 or 60–61", because no LibreOffice rendering is Excel's. B2 now gives Excel's own answers for serials 0, 0.5, 1, 59, 60 and 61 under ten codes. B1 gives eight near-tie displays. Both can be curated rows with basis `excel:observed`; column M's would be built-in id 15 drawn per N2's decision.
