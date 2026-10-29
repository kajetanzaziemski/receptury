# receptury – notes for agents

Google Apps Script (container-bound to a Google Sheet) with one sidebar that loads a saved recipe into the `Receptura` sheet.
Files: `Code.js` (server), `Sidebar.html` (UI, same dropdown + filter pattern as the `wyceny` project), `appsscript.json`.
UI strings and sheet names are Polish; keep them as they are.

## Sheets

### `Zapisane receptury` (source, headers in row 1)
`Receptura` | `Produkt` | `Receptura - skład`

`Receptura - skład` is `[[code;value];[code;value];...]`: one value per component. Fields are split on `;`, or on `,` when an entry has no `;`.
The sidebar lists one dropdown entry per row, labelled `Produkt (Receptura)`. Same label = one entry, the first row wins, and the sidebar warns when those rows hold different `Receptura - skład`.

### `Komponenty recepturowe` (lookup, headers in row 1)
`Kod komponentu` (A) | `Cena PLN/kg(szt)` (B) | `Pozycja menu` (C)

A code from `Receptura - skład` is translated to the `Pozycja menu` of its FIRST row here, and that is what lands in `Receptura!B`. A code with no menu position is loaded as it is and reported in the sidebar warning.

### `Receptura` (destination)
The header block is in `C1:G4`; columns A and B hold the entry list. Do not overwrite anything in rows 1–4, except the two title cells below.

Formulas are shown as the user writes them (Polish locale, `;` separators). `setFormula` in Apps Script needs English names and `,` separators, so `Code.js` stores them that way.

#### Mode 1 – default: masses are the input, content (%) is computed

| | C | D | E | F | G |
|---|---|---|---|---|---|
| 1 | Cena (zł/kg) | Masa (kg) | Zawartość (%) | Koszt (zł) | Udział w koszcie (%) |
| 2 | Produkt: | Łącznie: | Brakuje do 100%: | Łącznie: | Łącznie: |
| 3 | `=F3/D3` | `=SUM(D4:D)` | `=1-SUM(E4:E)` | `=SUM(F4:F)` | `=SUM(G4:G)` |
| 4 | `=ArrayFormula(IF(B4:B<>"";XLOOKUP(B4:B;'Komponenty recepturowe'!C:C;'Komponenty recepturowe'!B:B;"???");))` | (empty, input) | `=ArrayFormula(IF(D4:D<>"";D4:D/D$3;))` | `=ArrayFormula(IF(D4:D<>"";D4:D*C4:C;))` | `=ArrayFormula(IF(F4:F<>"";F4:F/F$3;))` |

#### Mode 2 – total mass: D3 is typed, content (%) is the input, masses are computed
Reached when `D3` is edited to a value (or a formula that does not read `D4:D`). Only these cells differ from mode 1:

| | C | D | E | F | G |
|---|---|---|---|---|---|
| 3 | `=F3/D3` | (typed by the user) | `=1-SUM(E4:E)` | `=SUM(F4:F)` | `=SUM(G4:G)` |
| 4 | (as mode 1) | `=ArrayFormula(IF(E4:E<>"";D3*E4:E;))` | (empty, input) | `=ArrayFormula(IF(D4:D<>"";D4:D*C4:C;))` | `=ArrayFormula(IF(F4:F<>"";F4:F/F$3;))` |

- Cell **B1** = the recipe name (`Receptura`), cell **B2** = the product in parentheses, e.g. `(Kisiel)`. The loader writes them only when the cell is empty or holds just `-`; anything else is left alone.
- Column **B** = component (a `Pozycja menu`), from row 4 down. `B4` carries a data validation (the template for the cells below).
- Columns whose row-4 cell holds an array formula are spilled: never write values into them below row 4, or the spill breaks (`#REF!`). That is `C` always, `E` and `G` in mode 1 and `D` and `G` in mode 2, plus `F` always.
- `C` = price looked up from `Komponenty recepturowe` (menu position in its C → price in its B; `???` when not found).

## What the script does (`Code.js`)

### Loading (`loadRecipe`)
- Appends the components of the chosen recipe below the destination's data range (`getDataRange().getLastRow() + 1`, never above `FIRST_ROW`, currently 5). Nothing is merged with existing entries, so duplicates are allowed on purpose.
- Writes the `Pozycja menu` to column B and the value to column D only. **It does not know about mode 2**: there D is spilled, so loading in mode 2 breaks the spill (open question, see below).
- Fills `B1` (recipe name) and `B2` (`(product)`) if they are empty or `-` (`fillIfBlank_`).
- Never multiplies values (an earlier `Masa` column was dropped from the source).
- Then calls `formatRecipeSheet_`, because writes made by a script do not fire `onEdit`.

### Formatting (`formatRecipeSheet_`, run by the simple `onEdit` trigger and after loading)
Triggered by an edit on `Receptura` that touches column B at row 5 or below. Range = row 5 down to the last row with data in column B, plus one.
- Column B: every cell that has no data validation gets the one from `B4`. Existing validations are not verified or changed.
- Columns B:G: alternating backgrounds, first row white (`#ffffff`), second grey (`#f3f3f3`). No header is touched.
- Colours and columns are the constants `ROW_WHITE`, `ROW_GREY`, `LAST_FORMAT_COL` at the top of `Code.js`.
- Rows left below the range keep whatever background they had (it is not cleared when the list shrinks).

### Switch to mode 2 (`switchToTotalMassMode_`, same `onEdit`)
Triggered by an edit of exactly `D3`. It switches only when D3 now holds something that does not read `D4:D` (a typed value, or a formula without that range). It does nothing when D3 was emptied, when D3 still sums `D4:D`, or when `D4` already reads `E4:E`.
- Sets the mode 2 formulas from `TOTAL_MASS_FORMULAS` (all of rows 3–4 listed above except `D3`, which keeps the user's value).
- Clears `D4:D` and `E4` first: the array formula in `D4` spills over `D5:D`, so old typed masses must go. They are not converted to percentages.
- There is no automatic way back: restore mode 1 by hand (`D3` = `=SUM(D4:D)`, `D4` empty, `E4` = the mode 1 array formula).

## Open questions
- What the loader should write in mode 2 (the source values would belong in `E`, but they may be masses, not fractions).

## Git
Development goes to `main`. Local checkout: `git clone https://github.com/kajetanzaziemski/receptury.git`; set `git config pull.rebase false` (or `true`) to silence the divergent-branches hint. The script can be synced with the Apps Script project using `clasp`.
