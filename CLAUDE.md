# receptury – notes for agents

Google Apps Script (container-bound to a Google Sheet) with one sidebar that loads a saved recipe into the `Receptura` sheet.
Files: `Code.js` (server), `Sidebar.html` (UI, same dropdown + filter pattern as the `wyceny` project), `PasteDialog.html` (modal dialog for a paste over D and E), `appsscript.json`.
UI strings and sheet names are Polish; keep them as they are. Every `.html` file must also exist in the Apps Script project (copying only `Code.js` leaves e.g. `PasteDialog` missing).

## Sheets

### `Zapisane receptury` (source, headers in row 1)
`Receptura` | `Produkt` | `Receptura - skład`

`Receptura - skład` is `[[code;value];[code;value];...]`: one value per component. Fields are split on `;`, or on `,` when an entry has no `;`.
The sidebar lists one dropdown entry per row, labelled `Produkt (Receptura)`. Same label = one entry, the first row wins, and the sidebar warns when those rows hold different `Receptura - skład`.

### `Komponenty recepturowe` (lookup, headers in row 1)
`Kod komponentu` (A) | `Cena PLN/kg(szt)` (B) | `Pozycja menu` (C)

A code from `Receptura - skład` is translated to the `Pozycja menu` of its FIRST row here, and that is what lands in `Receptura!B`. A code with no menu position is loaded as it is and reported in the sidebar warning.

### `Receptura` (destination)
The header block is in `C1:G4`; columns A and B hold the entry list. Do not overwrite anything in rows 1–4, except the two title cells below and the cells the state switch rewrites (`D3`, `D4`, `E4`).

Formulas are written with `;` separators (the user's Polish locale) in `Code.js` too. Apps Script's `setFormula` is documented to want `,`, so `setFormula_` writes the `;` form first and, if the cell shows a parse error (`#ERROR!` / `#BŁĄD!`), writes it again with commas. Keep any new formula free of `;` inside string literals, or that fallback breaks it.

#### Rows 1–4 (the same in both states unless noted)

| | C | D | E | F | G |
|---|---|---|---|---|---|
| 1 | Cena (zł/kg) | Masa (kg) | Zawartość (%) | Koszt (zł) | Udział w koszcie (%) |
| 2 | Produkt: | Łącznie: | Brakuje do 100%: | Łącznie: | Łącznie: |
| 3 | `=F3/D3` | **state dependent** | `=1-SUM(E4:E)` | `=SUM(F4:F)` | `=SUM(G4:G)` |
| 4 | `=ArrayFormula(IF(B4:B<>"";XLOOKUP(B4:B;'Komponenty recepturowe'!C:C;'Komponenty recepturowe'!B:B;"???");))` | **state dependent** | **state dependent** | `=ArrayFormula(IF(D4:D<>"";D4:D*C4:C;))` | `=ArrayFormula(IF(F4:F<>"";F4:F/F$3;))` |

#### Two states
The sheet is always in one of them. Each has one typed column and one computed (spilled) column.

| | Absolute mass (default) | Relative content |
|---|---|---|
| Typed | `D4:D` masses | `E4:E` content (%), and `D3` the total mass |
| `D3` | `=SUM(D4:D)` | typed number |
| `D4` | empty (first typed mass) | `=ArrayFormula(IF(E4:E<>"";D3*E4:E;))` |
| `E4` | `=ArrayFormula(IF(D4:D<>"";D4:D/D$3;))` | empty (first typed content) |

Never write values into a spilled column below row 4 yourself, or the spill breaks (`#REF!`). `C`, `F`, `G` are always spilled.

Other cells: **B1** = recipe name, **B2** = the product in parentheses, e.g. `(Kisiel)` (the loader fills them only when empty or `-`). Column **B** = components (a `Pozycja menu`) from row 4 down; `B4` carries the data validation that is the template for the cells below. `C` = price looked up from `Komponenty recepturowe` (menu position in its C → price in its B; `???` when not found).

## What the script does (`Code.js`)

### Triggers: simple vs installable
- `onEdit` is a **simple** trigger. It runs without authorization, so it can edit the spreadsheet but **cannot open dialogs**: `showModalDialog` fails there with "Określone uprawnienia nie są wystarczające do wywołania Ui.showModalDialog" (verified by the user). Do not move UI calls into it.
- `onEditInstalled` is an **installable** edit trigger, created once per spreadsheet from the menu **Receptury → Włącz pytanie przy wklejeniu w D i E** (`installPasteTrigger`; needs the `script.scriptapp` scope). It runs with the user's authorization and may open the dialog. Both triggers fire on every edit; each ignores the other's case.
- Errors in either trigger are not thrown (nobody could see them): `reportError_` writes them to the execution log and shows a toast.

### State switching (`onEdit`)
Editing the **computed** column of a state switches to the other state, in which that column is the typed one. The values of that column are kept; only the edited cells take the typed value; the other column becomes computed from it.

- **Absolute → relative** (`switchToRelative_`), on an edit of `E4:E` or of `D3`:
  - `E4:E` become typed numbers = every mass / the sum of the masses, the edited cells keeping what was typed.
  - `D3` becomes a typed number: the former sum (edit in `E`), or what the user typed (edit of `D3`).
  - `D4:D` is cleared and `D4` gets the mass formula. Masses of the other rows therefore stay the same.
- **Relative → absolute** (`switchToAbsolute_`), on an edit of `D4:D`:
  - `D4:D` become typed numbers = `D3` * content of each row, the edited cells keeping what was typed.
  - `D3` gets `=SUM(D4:D)`, `E4:E` is cleared and `E4` gets the content formula.
- Which state the sheet is in is read from the row-4 formula of the column the edit did NOT touch (`isAbsoluteState_`: `E4` reads `D4:D`; `isRelativeState_`: `D4` reads `E4:E`), because the edit may have overwritten the formula of its own column.
- Clearing cells does not switch (`hasTyped_`). An emptied `D3` in the absolute state gets its `=SUM(D4:D)` back. Editing `D3` in the relative state is just changing the total.
- Multi-column edits are ignored by the state switch (a paste over D and E is handled by `onEditInstalled`).
- Note: after a switch to relative, the contents typed may not add up to 100 %; `E3` ("Brakuje do 100%") shows the rest.

### Paste over D and E together (`onEditInstalled` + `PasteDialog.html`)
Without the installable trigger a two-column paste is not handled. With it:
- The dialog shows the pasted rows (row, component from column B, mass, content), lets the user choose **keep masses (D)** or **keep contents (E)**, edit the values of the kept column, and (when keeping contents) edit the **total mass**. Percent input like `30%` is read as `0.3`.
- `applyPasteChoice` then applies it (`payload.priorAbsolute` is the state before the paste, read from `D3` being the sum):
  - **keep D**: the dialog masses are written to `D`; the sheet ends in the absolute state (contents computed). From the relative state the other rows' masses are total * their content, as in `switchToAbsolute_`.
  - **keep E**: the dialog contents are written to `E`, `D3` = the dialog total; the sheet ends in the relative state. Rows that were not pasted keep their content (from the absolute state: mass / the total the dialog opened with).
- The default total offered is the sum of `D4:D` (from the absolute state) or the current `D3` (from the relative state). The default choice is the typed column of the state before the paste.
- Cancel leaves the pasted values exactly as pasted (and the formulas that the paste overwrote stay overwritten).
- Clearing a selection that spans D and E does not open the dialog.

### Loading (`loadRecipe`)
- Appends the components of the chosen recipe below the destination's data range (`getDataRange().getLastRow() + 1`, never above `FIRST_ROW`, currently 5). Nothing is merged with existing entries, so duplicates are allowed on purpose.
- Writes the `Pozycja menu` to column B and the value to column D (masses). If the sheet is in the relative state it is switched to absolute first (`switchToAbsolute_`, nothing is lost), because `D` is spilled there.
- Fills `B1` (recipe name) and `B2` (`(product)`) if they are empty or `-` (`fillIfBlank_`).
- Never multiplies values (an earlier `Masa` column was dropped from the source).
- Then calls `formatRecipeSheet_`, because writes made by a script do not fire `onEdit`.

### Formatting (`formatRecipeSheet_`, run by the simple `onEdit` trigger and after loading)
Triggered by an edit on `Receptura` that touches column B at row 5 or below. Range = row 5 down to the last row with data in column B, plus one.
- Column B: every cell that has no data validation gets the one from `B4`. Existing validations are not verified or changed.
- Columns B:G: alternating backgrounds, first row white (`#ffffff`), second grey (`#f3f3f3`). No header is touched.
- Colours and columns are the constants `ROW_WHITE`, `ROW_GREY`, `LAST_FORMAT_COL` at the top of `Code.js`.
- Rows left below the range keep whatever background they had (it is not cleared when the list shrinks).

## Git
Development goes to `main`. Local checkout: `git clone https://github.com/kajetanzaziemski/receptury.git`; set `git config pull.rebase false` (or `true`) to silence the divergent-branches hint. The script can be synced with the Apps Script project using `clasp`.
