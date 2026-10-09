# Request 5: the same evidence, on Windows Excel

**What was run.** Requests 3 and 4 were run again on **Windows** desktop Excel, to find anything where Windows and Mac Excel differ. The run handed Windows Excel the exact original workbooks the Mac sessions received, never files Excel had already saved.

| | |
|---|---|
| Excel | Microsoft 365 Apps for business, Excel 16.0, build 20430 (16.0.20430.20146), 64-bit |
| Windows | Windows 11 Pro 26H2, build 26300, in a virtual machine |
| Region | en-NZ (culture, system locale and formats), New Zealand time zone, English (US) display; the same region as the Mac runs (en_NZ) |
| Date | 2026-10-09 |
| How | Excel driven through its automation interface (COM) from PowerShell, with alerts off |

**What each workbook got:**
- **Request 3's 52 formula workbooks, request 4's 7 ask workbooks and its probe workbook** were opened, fully recalculated (`Application.CalculateFull`) and saved.
- **Request 4's 71 decoding shapes** were opened and saved without recalculation, as on the Mac.
- **A workbook that would not open normally** was retried with Excel's own repair (`CorruptLoad = xlRepairFile`).

**The folders** hold Excel's saved workbooks, their repair logs (`*.repair.txt`), `results.json` (per file: opened, repaired, saved, error) and `excel-cache.jsonl` (each cell's answer, read back from the saved `<v>`). The author fields are blank.

## Results: Windows against Mac

| check | result |
|---|---|
| request 3's 15,046 curated formulas | **15,045 identical bit for bit**; 1 differs (below) |
| request 4's 94 probe formulas | **94 of 94 identical** |
| decoding cell values read on both platforms (797) | **795 identical**; the other 2 are a `NOW()`, which reads the clock, and a file both platforms refused |
| shapes Mac Excel opened cleanly (50) | all 50 open cleanly on Windows |
| shapes Mac Excel repaired (27) | **every one is damaged on Windows too**: 15 repaired, 12 refused |
| the shape Mac Excel refused (a cut-off workbook part) | refused on Windows too |

### The one platform difference

| row | formula | Mac Excel 16.113 | Windows Excel 16.0.20430 |
|---|---|---|---|
| `excel-probe.00146@Z146` | `=A146^B146`, A146 = -8, B146 = -0.3333333333333333 | -0.5000000000000001 | **-0.5** |

A negative base to a negative fractional power is computed differently on the two platforms, a last-bit difference in the power function. The suite's goldens already treat this class as undecided (a refusal). Treat it as **platform-dependent**: a tool may match either, and should not be held to one.

### Repaired against refused

Twelve shapes Mac Excel repaired, after its repair prompt was accepted, were refused on Windows. That is, the files would not open even with Excel's repair mode through automation. The shapes:
- a missing sheet or string-table part, or no string table at all;
- invalid UTF-8 in four places;
- mismatched tags;
- a cut-off sheet, string table or styles part;
- the error-code ask workbook.

This is most likely the automation interface, not Excel: an interactive repair prompt tries more than the `xlRepairFile` load mode does. **Both platforms agree that every one of these files is damaged,** and that is the suite's rule: a tool should refuse them, not read them.
