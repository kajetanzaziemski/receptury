// ---------------------------------------------------------------------------
// RECEPTURY: a sidebar that loads ONE row of the SOURCE_SHEET ("Zapisane receptury") into the RECIPE_SHEET ("Receptura").
// The layout of the sheets is described in CLAUDE.md.
//
// Source sheet (headers in row 1): Receptura | Produkt | Receptura - skład
//   "Receptura - skład" = "[[component code;value];[component code;value];...]"  (one value per component)
// The sidebar offers one dropdown (with a search field) with an entry per row: "Produkt (Receptura)". Entries with the same
// label are shown once; the FIRST row is loaded, and a warning is returned if the rows with that label hold more than one
// distinct "Receptura - skład".
//
// Component codes are translated first: COMPONENTS_SHEET ("Komponenty recepturowe", headers in row 1) maps "Kod komponentu"
// to "Pozycja menu"; the FIRST row with the code gives the menu position, which is what goes into the destination.
// A code that is not found there is loaded as it is and reported in the warning.
//
// Destination: every component is appended below the data range (never below FIRST_ROW), the menu position in column B and
// its value in column D. Nothing is merged with what is already there, so duplicates are possible. B1 gets the Receptura name
// and B2 the Produkt in parentheses, but only when the cell is empty or holds just "-".
//
// Formatting of the destination (formatRecipeSheet_, on every edit of column B and after every load): from FIRST_ROW to the last
// row with data in column B + 1, column B gets the data validation of VALIDATION_CELL (B4) where a cell has none, and B:G get
// alternating row backgrounds (first row white, second grey).
//
// Two formula modes of the destination (see CLAUDE.md):
//   mass input (default)  D4:D typed, E4:E computed, D3 = SUM(D4:D)
//   total-mass mode      D3 typed (switchToTotalMassMode_, on an edit of D3 that does not read D4:D), E4:E typed, D4:D = D3 * E4:E
//   back to mass input   (switchToMassInputMode_, on an edit in D5:D while in total-mass mode)
// ---------------------------------------------------------------------------

const SOURCE_SHEET = 'Zapisane receptury';
const RECIPE_SHEET = 'Receptura';
const COMPONENTS_SHEET = 'Komponenty recepturowe';
const PRODUCT_COLUMN = 'Produkt';
const RECIPE_NAME_COLUMN = 'Receptura';
const RECIPE_COLUMN = 'Receptura - skład';
const CODE_COLUMN = 'Kod komponentu';
const MENU_COLUMN = 'Pozycja menu';
const KEY_COL = 2;          // B
const VALUE_COL = 4;        // D
const CONTENT_COL = 5;      // E
const LAST_FORMAT_COL = 7;  // G
const FIRST_ROW = 5;
const VALIDATION_CELL = 'B4';
const RECIPE_NAME_CELL = 'B1';
const PRODUCT_CELL = 'B2';
const MASS_TOTAL_CELL = 'D3';
const ROW_WHITE = '#ffffff';
const ROW_GREY = '#f3f3f3';

// Formulas of both modes (setFormula wants English names and commas, whatever the locale of the sheet).
const PRICE_FORMULA = '=ArrayFormula(IF(B4:B<>"",XLOOKUP(B4:B,\'Komponenty recepturowe\'!C:C,\'Komponenty recepturowe\'!B:B,"???"),))';
const COMMON_FORMULAS = {
  C3: '=F3/D3',
  E3: '=1-SUM(E4:E)',
  F3: '=SUM(F4:F)',
  G3: '=SUM(G4:G)',
  C4: PRICE_FORMULA,
  F4: '=ArrayFormula(IF(D4:D<>"",D4:D*C4:C,))',
  G4: '=ArrayFormula(IF(F4:F<>"",F4:F/F$3,))'
};
// total-mass mode: D3 is left alone (it holds the value the user typed), D4 is computed, E4:E is the input
const TOTAL_MASS_FORMULAS = Object.assign({}, COMMON_FORMULAS, {
  D4: '=ArrayFormula(IF(E4:E<>"",D3*E4:E,))'
});
// mass input: D4:D is the input (D4 stays empty), E4:E is computed
const MASS_INPUT_FORMULAS = Object.assign({}, COMMON_FORMULAS, {
  D3: '=SUM(D4:D)',
  E4: '=ArrayFormula(IF(D4:D<>"",D4:D/D$3,))'
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Receptury')
    .addItem('Otwórz panel', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('Receptury');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Simple trigger on the destination sheet:
 *  - an edit of D3 that is a value or a formula not reading D4:D switches the sheet to the total-mass formulas;
 *  - an edit in D5:D while in that mode switches back to the mass-input formulas;
 *  - an edit that touches column B at row 5 or below re-applies the formatting.
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getName() !== RECIPE_SHEET) return;

    if (e.range.getA1Notation() === MASS_TOTAL_CELL) {
      switchToTotalMassMode_(sheet);
      return;
    }

    // typed into the mass column below its first row: the masses are the input again
    if (e.range.getLastRow() >= FIRST_ROW && e.range.getColumn() <= VALUE_COL && e.range.getLastColumn() >= VALUE_COL) {
      const single = e.range.getNumRows() === 1 && e.range.getNumColumns() === 1;
      const typed = !single || (e.value !== undefined && e.value !== null && e.value !== '');   // clearing a cell is not typing a value
      if (typed && isTotalMassMode_(sheet)) {
        switchToMassInputMode_(sheet);
        return;
      }
    }

    if (e.range.getLastRow() < FIRST_ROW) return;
    if (e.range.getColumn() > KEY_COL || e.range.getLastColumn() < KEY_COL) return;
    formatRecipeSheet_(sheet);
  } catch (err) {
    // an edit must never fail because of formatting
  }
}

/** Is D4 the total-mass array formula (it reads E4:E)? */
function isTotalMassMode_(sheet) {
  return sheet.getRange('D4').getFormula().replace(/\$/g, '').toUpperCase().indexOf('E4:E') >= 0;
}

/**
 * D3 was edited. If it now holds something that does not read D4:D directly (a typed value or a formula without that range),
 * the total is independent: D4 becomes D3 * E4:E and E4:E turns into an input column. Does nothing when D3 was emptied, still
 * sums D4:D, or the sheet is already in this mode. The old D5:D values are cleared, because the array formula in D4 spills over them.
 */
function switchToTotalMassMode_(sheet) {
  const total = sheet.getRange(MASS_TOTAL_CELL);
  if (String(total.getValue()).trim() === '' && !total.getFormula()) return;
  if (total.getFormula().replace(/\$/g, '').toUpperCase().indexOf('D4:D') >= 0) return;
  if (isTotalMassMode_(sheet)) return;

  const maxRows = sheet.getMaxRows();
  if (maxRows >= 4) sheet.getRange(4, VALUE_COL, maxRows - 3, 1).clearContent();
  sheet.getRange('E4').clearContent();
  Object.keys(TOTAL_MASS_FORMULAS).forEach(a1 => sheet.getRange(a1).setFormula(TOTAL_MASS_FORMULAS[a1]));
}

/**
 * Back to the mass-input formulas: D3 = SUM(D4:D), D4's array formula removed (the cells typed into D5:D stay), E4:E computed again.
 * E4:E is cleared first, because the array formula in E4 spills over what was typed there.
 */
function switchToMassInputMode_(sheet) {
  sheet.getRange('D4').clearContent();
  const maxRows = sheet.getMaxRows();
  if (maxRows >= 4) sheet.getRange(4, CONTENT_COL, maxRows - 3, 1).clearContent();
  Object.keys(MASS_INPUT_FORMULAS).forEach(a1 => sheet.getRange(a1).setFormula(MASS_INPUT_FORMULAS[a1]));
}

/**
 * FIRST_ROW .. (last row with data in column B) + 1: data validation copied from VALIDATION_CELL into the cells of column B that
 * have none (existing validations are not checked), and alternating backgrounds in B:G (first row white, second grey).
 */
function formatRecipeSheet_(sheet) {
  const maxRows = sheet.getMaxRows();
  if (maxRows < FIRST_ROW) return;

  const lastRow = sheet.getLastRow();
  let lastData = FIRST_ROW - 1;
  if (lastRow >= FIRST_ROW) {
    const keys = sheet.getRange(FIRST_ROW, KEY_COL, lastRow - FIRST_ROW + 1, 1).getValues();
    for (let i = keys.length - 1; i >= 0; i--) {
      if (String(keys[i][0]).trim() !== '') { lastData = FIRST_ROW + i; break; }
    }
  }
  const end = Math.min(lastData + 1, maxRows);
  const count = end - FIRST_ROW + 1;

  const keyRange = sheet.getRange(FIRST_ROW, KEY_COL, count, 1);
  const template = sheet.getRange(VALIDATION_CELL).getDataValidation();
  if (template) {
    const rules = keyRange.getDataValidations();
    if (rules.some(r => !r[0])) {
      keyRange.setDataValidations(rules.map(r => [r[0] || template]));
    }
  }

  const colors = [];
  for (let i = 0; i < count; i++) {
    const color = i % 2 === 0 ? ROW_WHITE : ROW_GREY;
    const row = [];
    for (let c = KEY_COL; c <= LAST_FORMAT_COL; c++) row.push(color);
    colors.push(row);
  }
  sheet.getRange(FIRST_ROW, KEY_COL, count, LAST_FORMAT_COL - KEY_COL + 1).setBackgrounds(colors);
}

/** Puts `text` into the cell only when it is empty or holds just "-" (and there is something to put). */
function fillIfBlank_(sheet, a1, text) {
  if (!text) return;
  const cell = sheet.getRange(a1);
  const cur = String(cell.getValue()).trim();
  if (cur === '' || cur === '-') cell.setValue(text);
}

/** A number from a cell (a decimal comma and spaces are fine); NaN when empty or not a number. */
function toNumber_(v) {
  if (typeof v === 'number') return v;
  const t = String(v === null || v === undefined ? '' : v).replace(/[\s ]/g, '').replace(',', '.');
  return t === '' ? NaN : Number(t);
}

/** The source sheet as rows: [{ row (sheet row number), label, recipe, product, text (Receptura - skład) }]; throws a Polish message. */
function readSource_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SOURCE_SHEET);
  if (!sheet) throw new Error('Brak arkusza "' + SOURCE_SHEET + '".');
  const all = sheet.getDataRange().getValues();
  const head = all[0].map(h => String(h).trim());
  const idx = {};
  [PRODUCT_COLUMN, RECIPE_NAME_COLUMN, RECIPE_COLUMN].forEach(h => {
    idx[h] = head.indexOf(h);
    if (idx[h] < 0) throw new Error('Brak kolumny "' + h + '" w arkuszu "' + SOURCE_SHEET + '".');
  });
  const cell = (r, h) => String(r[idx[h]] === null || r[idx[h]] === undefined ? '' : r[idx[h]]).trim();
  const rows = [];
  all.slice(1).forEach((r, i) => {
    const product = cell(r, PRODUCT_COLUMN), recipe = cell(r, RECIPE_NAME_COLUMN);
    if (!product && !recipe) return;
    rows.push({ row: i + 2, label: product + ' (' + recipe + ')', recipe: recipe, product: product, text: cell(r, RECIPE_COLUMN) });
  });
  return rows;
}

/** { lowercased code: "Pozycja menu" } from COMPONENTS_SHEET, the first row of a code winning; throws a Polish message. */
function readMenuPositions_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(COMPONENTS_SHEET);
  if (!sheet) throw new Error('Brak arkusza "' + COMPONENTS_SHEET + '".');
  const all = sheet.getDataRange().getValues();
  const head = all[0].map(h => String(h).trim());
  const codeIdx = head.indexOf(CODE_COLUMN), menuIdx = head.indexOf(MENU_COLUMN);
  if (codeIdx < 0) throw new Error('Brak kolumny "' + CODE_COLUMN + '" w arkuszu "' + COMPONENTS_SHEET + '".');
  if (menuIdx < 0) throw new Error('Brak kolumny "' + MENU_COLUMN + '" w arkuszu "' + COMPONENTS_SHEET + '".');
  const map = {};
  all.slice(1).forEach(r => {
    const code = String(r[codeIdx]).trim().toLowerCase();
    const menu = String(r[menuIdx]).trim();
    if (code && !(code in map)) map[code] = menu;
  });
  return map;
}

/**
 * "[[name;value];[name;value]]" -> [{name, value}]. Inside an entry the fields are separated by semicolons, or by commas when
 * there is no semicolon; the first field is the name, the last one the value (a number when it parses as one, else the text).
 */
function parseRecipe_(text) {
  const raw = String(text === null || text === undefined ? '' : text).trim();
  const out = [];
  const re = /\[([^\[\]]*)\]/g;                  // innermost brackets = one entry each
  let m;
  while ((m = re.exec(raw)) !== null) {
    const fields = m[1].indexOf(';') >= 0 ? m[1].split(';') : m[1].split(',');
    const name = (fields[0] || '').trim();
    if (!name) continue;
    const v = fields.length > 1 ? fields[fields.length - 1].trim() : '';
    const n = toNumber_(v);
    out.push({ name: name, value: v === '' ? '' : (Number.isFinite(n) ? n : v) });
  }
  return out;
}

/** For the sidebar: the dropdown entries, one per distinct label, in sheet order: { entries: [{row, label}] } | { entries: [], error }. Never throws. */
function getEntries() {
  try {
    const seen = {}, entries = [];
    readSource_().forEach(r => {
      if (seen[r.label]) return;
      seen[r.label] = true;
      entries.push({ row: r.row, label: r.label });
    });
    return { entries: entries };
  } catch (e) {
    return { entries: [], error: e.message };
  }
}

/**
 * Loads the source row `row` (the first one of its label) below the data range of the destination sheet.
 * Returns { ok, written, warning } | { error }.
 */
function loadRecipe(row) {
  try {
    const rows = readSource_();
    const chosen = rows.filter(r => r.row === row)[0];
    if (!chosen) return { error: 'Nie znaleziono wybranej pozycji – odśwież listę.' };

    const same = rows.filter(r => r.label === chosen.label);
    const distinct = same.map(r => r.text).filter((t, i, a) => a.indexOf(t) === i);
    const components = parseRecipe_(chosen.text);
    if (!components.length) return { error: 'Brak odczytywalnego składu dla "' + chosen.label + '".' };

    // codes -> menu positions (a code without a position stays as it is)
    const menu = readMenuPositions_();
    const unknown = [];
    components.forEach(c => {
      const found = menu[c.name.toLowerCase()];
      if (found) c.name = found;
      else unknown.push(c.name);
    });

    const warnings = [];
    if (distinct.length > 1) {
      warnings.push('Znaleziono ' + distinct.length + ' różnych pozycji "' + RECIPE_COLUMN + '" dla "' + chosen.label + '" – wczytano pierwszą.');
    }
    if (unknown.length) {
      warnings.push('Brak "' + MENU_COLUMN + '" w arkuszu "' + COMPONENTS_SHEET + '" dla: ' + unknown.join(', ') + ' – wczytano kod komponentu.');
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RECIPE_SHEET);
    if (!sheet) return { error: 'Brak arkusza "' + RECIPE_SHEET + '".' };

    // the first empty row, as the data range tells, but never above FIRST_ROW
    const start = Math.max(sheet.getDataRange().getLastRow() + 1, FIRST_ROW);
    sheet.getRange(start, KEY_COL, components.length, 1).setValues(components.map(c => [c.name]));
    sheet.getRange(start, VALUE_COL, components.length, 1).setValues(components.map(c => [c.value]));

    fillIfBlank_(sheet, RECIPE_NAME_CELL, chosen.recipe);
    fillIfBlank_(sheet, PRODUCT_CELL, chosen.product ? '(' + chosen.product + ')' : '');

    formatRecipeSheet_(sheet);       // a script write does not fire onEdit

    SpreadsheetApp.flush();
    return { ok: true, written: components.length, warning: warnings.join('\n') };
  } catch (e) {
    return { error: e.message };
  }
}
