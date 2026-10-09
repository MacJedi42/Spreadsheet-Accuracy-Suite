# Excel results: request 2 (the internal tool batch B4)

**Excel:** Microsoft Excel for Mac 16.113.3
**Run:** 2026-10-02, driven via AppleScript (not by hand).
**Method:** Calculation set to Manual before opening any file. Files were opened as copies (originals untouched) and closed without saving. "Excel shows" is Excel's formatted-text property (`string value`) after autofitting the column. A5's full recalc is `calculate full`. Part B was built under Automatic calculation, then Manual was restored.
**Deliverable:** `partB.xlsx` (Part B workbook), beside this file.

## Summary of verdicts
- A1: Excel rounds all five near-tie cells UP. The internal tool is right, IronCalc is wrong.
- A2: Excel shows a BLANK for day 0 in H4/P28, not "0" and not #VALUE!. Cause is the sheet's `showZeros="0"`. Day 0 under `d` on a normal sheet shows `0`.
- A3: Excel moves AE14's text per cell. The internal tool is right (AF14 = AE14's text shifted one column), IronCalc is wrong.
- A4: Excel DECODES `_xHHHH_`. LEN = 3, not 9.
- A5: stored values survive a full recalc unchanged (no errors).
- A6: an empty formatted cell is ignored by MAX, COUNT and COUNTA.

## A1: numbers just below a .5 tie (stored values untouched)
| file | sheet | cell | format | Excel shows |
|---|---|---|---|---|
| 1_55427_answer.xlsx | KS4 data | K76 | `0%` | **15%** |
| 1_32468_answer.xlsx | SUMMARY | I20 | `#,##0` | **-14,705,084** |
| 1_50442_answer.xlsx | RESULTS 1 | G20 | `#,##0` | **1,322,198** |
| 1_50442_answer.xlsx | CALCS | B12 | `#,##0` | **1,322,198** |
| 3_49613_input.xlsx | RESULTS 1 | C33 | `#,##0` | **-2,252** |

Caveat: AppleScript returns values at about 15 significant digits, so I could not confirm Excel kept all 17 digits on load. The displayed results are what Excel renders.

## A2: day 0 under `d`
| file | sheet | cell | stored | format | Excel shows |
|---|---|---|---|---|---|
| 1_118-8_answer.xlsx | Schedule | H4 | 0 | `d` | **blank** |
| 1_118-8_answer.xlsx | Schedule | P28 | 0 | `d` | **blank** |

Both are formulas (`IFERROR(INDEX(...),"")`) with cached `<v>0</v>`, numFmt `d`. The Schedule sheet's `<sheetView ... showZeros="0">` makes Excel hide zero values. There is no conditional formatting on that sheet. Verified in the XML (`xl/worksheets/sheet4.xml`).
Control: in `partB.xlsx` (default sheet, zeros shown) the number 0 under `d` shows `0`.
So the display depends on the sheet's show-zeros flag, not the `d` format. The internal tool's "0" is right only when zeros are shown.

## A3: formula bar, `1_44017_answer.xlsx`, sheet Data
| cell | Excel's formula bar |
|---|---|
| AD14 (own formula) | `=IF($L14,Q14*(1+SUM($M14:INDEX($M14:$P14,(YEARFRAC($L14,AD$9)*$J14+1)))*(AD$9>=$L14)),)` |
| AE14 | `=IF($L14,R14*(1+SUM($M14:INDEX($M14:$P14,(YEARFRAC($L14,AE$9)*$J14+1)))*(AE$9>=$L14)),)` |
| AF14 | `=IF($L14,S14*(1+SUM($M14:INDEX($M14:$P14,(YEARFRAC($L14,AF$9)*$J14+1)))*(AF$9>=$L14)),)` |
| AG14 | `=IF($L14,T14*(1+SUM($M14:INDEX($M14:$P14,(YEARFRAC($L14,AG$9)*$J14+1)))*(AG$9>=$L14)),)` |
| AD15 | `=IF($L15,Q15*(1+SUM($M15:INDEX($M15:$P15,(YEARFRAC($L15,AD$9)*$J15+1)))*(AD$9>=$L15)),)` |
| AE15 | `=IF($L15,R15*(1+SUM($M15:INDEX($M15:$P15,(YEARFRAC($L15,AE$9)*$J15+1)))*(AE$9>=$L15)),)` |
| AF15 | `=IF($L15,S15*(1+SUM($M15:INDEX($M15:$P15,(YEARFRAC($L15,AF$9)*$J15+1)))*(AF$9>=$L15)),)` |

AF14 matches the internal tool's reading exactly (S14 and AF$9).

## A4: `_xHHHH_` escapes, `1_545-35_input.xlsx`, sheet Data
Y3 and Z3 each have length 3 as read by AppleScript, so Excel decodes the escapes.

| formula | Excel gives |
|---|---|
| `=LEN(Y3)` | 3 |
| `=CODE(MID(Y3,2,1))` | 30 |
| `=LEN(Z3)` | 3 |
| `=CODE(MID(Z3,3,1))` | 6 |

Note: my first attempt, in AA1, returned empty results for unknown reasons. A retry in AZ50 gave the values above.

## A5: stored values vs full recalc
Nothing changed on full recalc.
| file | sheet | cells | on opening = after recalc |
|---|---|---|---|
| 1_CF_18814_answer.xlsx | attendance-Dec | E1 | `01-Jun` (serial 44348), `=DATEVALUE("1"&B1&B2)`, fmt `d-mmm` |
| | | C3, D3 | `01`, `02` (fmt `dd`) |
| 1_11276_answer.xlsx | ATTENDENCE | F3, G3 | `Wed`, `Thu` (`=TEXT(F4,"DDD")`) |
| 1_57716_answer.xlsx | Calendar | B3, C3, D3, H3 | `30`, `31`, `01`, `05` (fmt `dd`) |
| | | I3 | empty string (`""`) (`=IF(N(B3),"",SUMPRODUCT(...))`) |

## A6: `1_334-11_answer.xlsx`, sheet imported Data
C5 is truly empty (General format).
| formula | Excel gives |
|---|---|
| `=MAX(C2:C5)` | -12671.88 (read from a rounded AppleScript value, not exact) |
| `=COUNT(C2:C5)` | 3 |
| `=COUNTA(C2:C5)` | 3 |

## B1: display just below a .5 tie
| row | A formula | A value read | B format | Excel shows |
|---|---|---|---|---|
| 1 | `=0.145` | 0.145 | `0%` | 15% |
| 2 | `=-14705083.5+3*2^-29` | -14705083.5 (15 digits) | `#,##0` | -14,705,084 |
| 3 | `=1322197.5-2^-31` | 1322197.5 (15 digits) | `#,##0` | 1,322,198 |
| 4 | `=-2251.5+2^-40` | -2251.499999999999 | `#,##0` | -2,252 |
| 5 | `=0.5-2^-54` | 0.5 | `0` | 1 |
| 6 | `=2.5-2^-51` | 2.5 | `0` | 3 |
| 7 | `=1.005` | 1.005 | `0.00` | 1.01 |
| 8 | `=0.285` | 0.285 | `0%` | 29% |

Caveat: in rows 5 and 6, Excel's formula arithmetic appears to have snapped the result to exactly 0.5 and 2.5, so those rows test the exact tie, not a value just below it. Rows 2 and 3 may be affected the same way. The A1 corpus cells (stored values loaded from file) are the cleaner evidence.

## B2: day 0 and the 1900 leap-year day
Columns: D raw number, E `d`, F `dd`, G `ddd`, H `dddd`, I `m`, J `mmm`, K `yyyy`, L `m/d/yyyy`, M `d-mmm-yy`, N `dddd d mmmm yyyy`.

| D | E | F | G | H | I | J | K | L | M | N |
|---|---|---|---|---|---|---|---|---|---|---|
| 0 | 0 | 00 | Sat | Saturday | 1 | Jan | 1900 | 1/0/1900 | 00-Jan-00 | Saturday 0 January 1900 |
| 0.5 | 0 | 00 | Sat | Saturday | 1 | Jan | 1900 | 1/0/1900 | 00-Jan-00 | Saturday 0 January 1900 |
| 1 | 1 | 01 | Sun | Sunday | 1 | Jan | 1900 | 1/1/1900 | 01-Jan-00 | Sunday 1 January 1900 |
| 59 | 28 | 28 | Tue | Tuesday | 2 | Feb | 1900 | 2/28/1900 | 28-Feb-00 | Tuesday 28 February 1900 |
| 60 | 29 | 29 | Wed | Wednesday | 2 | Feb | 1900 | 2/29/1900 | 29-Feb-00 | Wednesday 29 February 1900 |
| 61 | 1 | 01 | Thu | Thursday | 3 | Mar | 1900 | 3/1/1900 | 01-Mar-00 | Thursday 1 March 1900 |

Notes: serial 0 is a Saturday, serial 60 is the fictitious 29 Feb 1900, and the weekday for serial 59 is Tuesday (a real calendar would give Wednesday). `m/d/yyyy` for 0 gives `1/0/1900`.

## Surprises
- No prompts on opening, no formulas rewritten, and no values changed under Manual calculation.
- The A2 blank (showZeros) was not anticipated by the request.
- Excel was left running with Manual calculation set.
