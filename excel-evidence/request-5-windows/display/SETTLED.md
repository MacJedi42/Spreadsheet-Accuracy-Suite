# Request 5, display: what Windows Excel showed for every number-format row

**Source.** Microsoft 365 Excel 16.0.20430.20146, 64-bit, on Windows 11 Pro in a virtual machine, driven through COM from PowerShell, 2026-10-10.
- **The workbooks:** column A holds the value under the format code (or a built-in id), B the code, and C `=TEXT(A,B)`. Each workbook was opened read-only and fully recalculated.
- **What was read back:** the value Excel read (`Value2`, round-trip), the cell's displayed text (`Range.Text`, in a column 80 characters wide) and `TEXT()`'s result.
- **Region:** en-US for the golden rows, with each run checked by a culture probe. The half-tie pairs were asked in en-NZ; their codes draw the same in both.

The workbooks hold only generated values, no personal data. A workbook holds at most about 250 custom number formats, so the codes were spread over several.

## The files

| file | rows | what |
|---|---|---|
| `number-format-cache.jsonl` | 189,361 | every `R`, `I`, `D` and `DI` row of `goldens/number-format/number_format_golden.tsv`, as it stood before this request. Fields: `row`, `key`, `code`, `value` and `text_en_US` (Excel's `TEXT()`), plus three that appear only when they differ: `cell_en_US` (the cell), `text_en_NZ`/`cell_en_NZ` (the same in en-NZ), and `excel_value` (what Excel read). |
| `number-format-asks.jsonl` | 8,819 | follow-up asks: every built-in id row by `TEXT()` under the id's en-US code; the `tool:en` and `tool:iso` texts; pairs the earlier golden had left out as undecided; the Excel section's settled rows. `why` says what each is for. |
| `half-ties.jsonl` | 7,177 | displays at a 15-digit tie (below). `rule_15digit` and `rule_binary` are two candidate readings, kept for reference; Excel's text is `text_en_NZ`. |
| `r5b-minus-probes.jsonl` | 540 | the minus sign at zero over 54 code shapes and 10 values |
| `dates-probes.jsonl` | 95 | shapes the golden held once or not at all: exponents, fractions, sub-second times |

The golden now takes Excel's `TEXT()` as the expected text of every row Excel answered (basis `excel:asked`). An independent check matched all 195,210 such rows against these files.

## What Excel showed

**Displays at a 15-digit tie.** One rule reproduces all 7,177 pairs, and `TEXT()` agreed with the cell on every one:
1. take the exact binary value; for a `%` code, the exact decimal times 100, not the double product;
2. round it to 15 significant digits, with an exact tie at the 16th digit going toward zero;
3. round that half away from zero at the code's last shown place.

36035847931.395 under `#,##0.00` is `36,035,847,931.40`, and 1234567890123455 under `0` is `1234567890123450`. LibreOffice rounds the binary value instead, and disagrees on 2,261 of these pairs past about 2^41.

**The minus sign at zero.** A negative whose digit placeholders all show zero gets no minus: -0.004 under `0.0` is `0.0`, under `#` the cell is empty, and under `"10x"0` it is `10x0` (a digit in a quoted literal does not count). The exceptions:
- **a section with no digit placeholders** keeps the minus: `\:` gives `-:`, `"abc"` gives `-abc`, and `""` gives a minus alone;
- **a fraction with a whole part** keeps it: `# ?/8` at -0.004 is `-0` and spaces, while `?/?` gives `0/1`;
- **an exponent** never shows zero, and **General** keeps the minus;
- **a second section** never gets a minus;
- **under conditions** it depends on the condition: `[<0]0.0` shows -0.4 as `0.4`, but `[>-1]0.0` shows `-0.4`.

**Other shapes:**
- each `_x` pad is one space;
- a fraction part that prints nothing leaves a space for each character it spans;
- `0.E+00` keeps its point;
- a section of only a colour still draws General, but a colour beside a language tag (`[$-409][Red]`) draws nothing;
- a `*` fill expands to the column in the cell, and `TEXT()` drops it;
- `#?/12` at 99999 cannot be drawn: `TEXT()` gives #VALUE! and the cell shows `####`;
- a code's language tag draws names, AM/PM and digits in that language (`[$-41F]` gives `Ocak`, and `[$-3000401]` draws Arabic-Indic digits).

**In en-NZ:** built-in date and time ids follow the region, and `TEXT()` writes AM/PM as `am`/`pm` while the cell shows `AM`.

**A reader note.** Excel read the `<v>7.2744917013307</v>` it was given as 7.2744917013306996, one ulp from the nearest double. That is 1 of 7,177 values; every other value read back exactly.
