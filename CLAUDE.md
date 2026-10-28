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
The header block is in `C1:G4`; columns A and B hold the entry list. Do not overwrite anything in rows 1–4.

| | C | D | E | F | G |
|---|---|---|---|---|---|
| 1 | Cena (zł/kg) | Masa (kg) | Zawartość (%) | Koszt (zł) | Udział w koszcie (%) |
| 2 | Produkt: | Łącznie: | Brakuje do 100%: | Łącznie: | Łącznie: |
| 3 | `=F3/D3` | `=SUM(D4:D)` | `=1-SUM(E4:E)` | `=SUM(F4:F)` | `=SUM(G4:G)` |
| 4 | `=ArrayFormula(IF(B4:B<>"";XLOOKUP(B4:B;'Komponenty recepturowe'!C:C;'Komponenty recepturowe'!B:B;"???");))` | (empty) | `=ArrayFormula(IF(D4:D<>"";D4:D/D$3;))` | `=ArrayFormula(IF(D4:D<>"";D4:D*C4:C;))` | `=ArrayFormula(IF(F4:F<>"";F4:F/F$3;))` |

- Column **B** = component (a `Pozycja menu`), from row 4 down. `B4` carries a data validation.
- Column **D** = the value (mass, kg) of that component, from row 4 down.
- Columns **C, E, F, G** from row 4 down are spilled by the array formulas above: never write values there, or the spill breaks (`#REF!`).
- `C` = price looked up from `Komponenty recepturowe` (menu position in its C → price in its B; `???` when not found).

## What the loader does (`loadRecipe` in `Code.js`)
- Appends the components of the chosen recipe below the destination's data range (`getDataRange().getLastRow() + 1`, never above `FIRST_ROW`, currently 5). Nothing is merged with existing entries, so duplicates are allowed on purpose.
- Writes the `Pozycja menu` to column B and the value to column D only.
- Copies the data validation of `B4` onto every newly written B cell.
- Never multiplies values (an earlier `Masa` column was dropped from the source).

## Git
Development goes to `main`. Local checkout: `git clone https://github.com/kajetanzaziemski/receptury.git`; set `git config pull.rebase false` (or `true`) to silence the divergent-branches hint. The script can be synced with the Apps Script project using `clasp`.
