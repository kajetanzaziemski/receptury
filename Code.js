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
// its mass in column D. Nothing is merged with what is already there, so duplicates are possible. B1 gets the Receptura name
// and B2 the Produkt in parentheses, but only when the cell is empty or holds just "-".
//
// Formatting of the destination (formatRecipeSheet_, on every edit of column B and after every load): from FIRST_ROW to the last
// row with data in column B + 1, column B gets the data validation of VALIDATION_CELL (B4) where a cell has none, and B:G get
// alternating row backgrounds (first row white, second grey).
//
// Two states of the destination (see CLAUDE.md):
//   absolute mass     D4:D (masses) are typed, D3 = SUM(D4:D), E4:E (content %) = D / D$3
//   relative content  E4:E (content %) are typed, D3 (total mass) is typed, D4:D = D3 * E
// Editing the computed column of a state switches to the other state, in which that column is the typed one. The values of that
// column are kept as they were (except the edited cells, which take the typed value) and the other column becomes computed.
// Values pasted over both columns at once need a decision, which a simple trigger cannot ask for: an INSTALLABLE edit trigger
// (onEditInstalled, switched on from the menu) opens a dialog (PasteDialog.html) asking which column to keep.
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
const FIRST_ROW = 5;        // first row the loader writes to
const HEAD_ROW = 4;         // row of the array formulas = first row of the mass / content columns
const VALIDATION_CELL = 'B4';
const RECIPE_NAME_CELL = 'B1';
const PRODUCT_CELL = 'B2';
const MASS_TOTAL_CELL = 'D3';
const ROW_WHITE = '#ffffff';
const ROW_GREY = '#f3f3f3';
const PASTE_TRIGGER_HANDLER = 'onEditInstalled';

// The formulas that differ between the two states (written with ; as the argument separator, see setFormula_).
// Everything else in rows 3–4 (C3, E3, F3, G3, C4, F4, G4) is the same in both states and is not touched.
const TOTAL_SUM_FORMULA = '=SUM(D4:D)';                                      // D3, absolute mass
const CONTENT_FORMULA = '=ArrayFormula(IF(D4:D<>"";D4:D/D$3;))';               // E4, absolute mass
const MASS_FORMULA = '=ArrayFormula(IF(E4:E<>"";D3*E4:E;))';                   // D4, relative content

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Receptury')
    .addItem('Otwórz panel', 'showSidebar')
    .addItem('Włącz pytanie przy wklejeniu w D i E', 'installPasteTrigger')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('Receptury');
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Menu: creates the installable edit trigger that asks what to keep after a paste over D and E (once; asks for authorization). */
function installPasteTrigger() {
  const ui = SpreadsheetApp.getUi();
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === PASTE_TRIGGER_HANDLER);
  if (!exists) {
    ScriptApp.newTrigger(PASTE_TRIGGER_HANDLER).forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  }
  ui.alert(exists
    ? 'Wyzwalacz jest już włączony.'
    : 'Włączono. Wklejenie wartości do kolumn D i E zapyta, którą kolumnę zachować.');
}

/**
 * Simple trigger on the destination sheet:
 *  - an edit of D3 or of D4:D (the computed cells of the relative state) / E4:E (the computed cells of the absolute state)
 *    switches the state;
 *  - an edit that touches column B at row 5 or below re-applies the formatting.
 * The state is read from the row-4 formula of the column the edit did NOT touch, because the edit may have overwritten its own.
 * (An edit over D and E together is left to onEditInstalled.)
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const range = e.range;
    const sheet = range.getSheet();
    if (sheet.getName() !== RECIPE_SHEET) return;

    if (range.getA1Notation() === MASS_TOTAL_CELL) {
      onTotalEdited_(sheet);
      return;
    }

    const oneColumn = range.getColumn() === range.getLastColumn();
    if (oneColumn && range.getRow() >= HEAD_ROW) {
      if (range.getColumn() === VALUE_COL) {
        if (!isAbsoluteState_(sheet) && hasTyped_(range)) switchToAbsolute_(sheet, range);
        return;
      }
      if (range.getColumn() === CONTENT_COL) {
        if (!isRelativeState_(sheet) && hasTyped_(range)) switchToRelative_(sheet, range);
        return;
      }
    }

    if (range.getLastRow() < FIRST_ROW) return;
    if (range.getColumn() > KEY_COL || range.getLastColumn() < KEY_COL) return;
    formatRecipeSheet_(sheet);
  } catch (err) {
    // an edit must never fail because of formatting
  }
}

/**
 * Installable edit trigger (see installPasteTrigger): values pasted over D and E together. Both columns now hold real values, so
 * which one is to be kept has to be asked; the dialog also lets the values (and the total mass) be corrected first.
 */
function onEditInstalled(e) {
  try {
    if (!e || !e.range) return;
    const range = e.range;
    const sheet = range.getSheet();
    if (sheet.getName() !== RECIPE_SHEET) return;
    if (range.getColumn() > VALUE_COL || range.getLastColumn() < CONTENT_COL) return;

    const first = Math.max(range.getRow(), HEAD_ROW);
    const last = range.getLastRow();
    if (last < first) return;
    const n = last - first + 1;
    const d = sheet.getRange(first, VALUE_COL, n, 1).getValues();
    const c = sheet.getRange(first, CONTENT_COL, n, 1).getValues();
    const names = sheet.getRange(first, KEY_COL, n, 1).getValues();
    const rows = d.map((r, i) => ({ row: first + i, name: String(names[i][0]), d: r[0], e: c[i][0] }));
    if (!rows.some(r => r.d !== '' || r.e !== '')) return;             // clearing is not pasting

    const priorAbsolute = reads_(sheet.getRange(MASS_TOTAL_CELL).getFormula(), 'D4:D');
    let defaultTotal;
    if (priorAbsolute) {
      const all = sheet.getRange(HEAD_ROW, VALUE_COL, sheet.getMaxRows() - HEAD_ROW + 1, 1).getValues();
      defaultTotal = all.reduce((s, r) => s + (typeof r[0] === 'number' ? r[0] : 0), 0);
    } else {
      const t = toNumber_(sheet.getRange(MASS_TOTAL_CELL).getValue());
      defaultTotal = Number.isFinite(t) ? t : null;
    }

    const payload = { keep: priorAbsolute ? 'D' : 'E', priorAbsolute: priorAbsolute, defaultTotal: defaultTotal, rows: rows };
    const t = HtmlService.createTemplateFromFile('PasteDialog');
    t.payload = JSON.stringify(payload).replace(/</g, '\\u003c');
    SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(520).setHeight(560), 'Wklejono wartości do kolumn D i E');
  } catch (err) {
    // an edit must never fail because of the dialog
  }
}

/** A value typed in the paste dialog: '' stays '', "30%" -> 0.3, a number with a decimal comma -> number; anything else -> NaN. */
function parseDialogValue_(s) {
  const t = String(s === null || s === undefined ? '' : s).trim();
  if (t === '') return '';
  if (/%$/.test(t)) { const n = toNumber_(t.slice(0, -1)); return Number.isFinite(n) ? n / 100 : NaN; }
  return toNumber_(t);
}

/**
 * Called by the paste dialog. payload = { keep: 'D'|'E', priorAbsolute, defaultTotal, total, rows: [{row, d, e}] } (strings from the dialog).
 *  keep D: the masses (as in the dialog) are typed, the contents are computed from them;
 *  keep E: the contents (as in the dialog) are typed and the total mass is `total`; the masses are computed.
 * Rows outside the pasted ones keep their values in the column that is kept. Returns { ok } | { error }.
 */
function applyPasteChoice(p) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RECIPE_SHEET);
    if (!sheet) return { error: 'Brak arkusza "' + RECIPE_SHEET + '".' };

    const rows = [];
    for (let i = 0; i < p.rows.length; i++) {
      const d = parseDialogValue_(p.rows[i].d), c = parseDialogValue_(p.rows[i].e);
      if (d !== '' && !Number.isFinite(d)) return { error: 'Nieprawidłowa masa w wierszu ' + p.rows[i].row + '.' };
      if (c !== '' && !Number.isFinite(c)) return { error: 'Nieprawidłowa zawartość w wierszu ' + p.rows[i].row + '.' };
      rows.push({ row: p.rows[i].row, d: d, e: c });
    }
    if (!rows.length) return { ok: true };
    const first = rows[0].row, lastPasted = rows[rows.length - 1].row;
    const maxRows = sheet.getMaxRows();

    if (p.keep === 'D') {
      const dRange = sheet.getRange(first, VALUE_COL, rows.length, 1);
      dRange.setValues(rows.map(r => [r.d]));
      if (p.priorAbsolute) {
        sheet.getRange(HEAD_ROW, CONTENT_COL, maxRows - HEAD_ROW + 1, 1).clearContent();
        setFormula_(sheet, 'E4', CONTENT_FORMULA);
      } else {
        switchToAbsolute_(sheet, dRange);     // the other rows: total mass * their typed contents
      }
      return { ok: true };
    }

    const total = toNumber_(p.total);
    if (!Number.isFinite(total)) return { error: 'Podaj masę łączną (liczbę).' };
    const last = Math.max(lastDataRow_(sheet, [KEY_COL, VALUE_COL, CONTENT_COL]), lastPasted);
    const n = last - HEAD_ROW + 1;
    let content;
    if (p.priorAbsolute) {
      // contents of the rows that were not pasted: their mass / the total the dialog started with
      const t0 = Number(p.defaultTotal);
      content = sheet.getRange(HEAD_ROW, VALUE_COL, n, 1).getValues().map(r => (typeof r[0] === 'number' && t0 ? r[0] / t0 : ''));
    } else {
      content = sheet.getRange(HEAD_ROW, CONTENT_COL, n, 1).getValues().map(r => r[0]);
    }
    rows.forEach(r => { content[r.row - HEAD_ROW] = r.e; });

    sheet.getRange(HEAD_ROW, CONTENT_COL, n, 1).setValues(content.map(v => [v]));
    sheet.getRange(MASS_TOTAL_CELL).setValue(total);
    sheet.getRange(HEAD_ROW, VALUE_COL, maxRows - HEAD_ROW + 1, 1).clearContent();
    setFormula_(sheet, 'D4', MASS_FORMULA);
    SpreadsheetApp.flush();
    return { ok: true };
  } catch (err) {
    return { error: err.message };
  }
}

/** Does the formula read the given range (e.g. "D4:D")? */
function reads_(formula, a1) {
  return String(formula).replace(/\$/g, '').toUpperCase().indexOf(a1) >= 0;
}

/** Absolute mass state: E4 is the formula computing the content from the masses. */
function isAbsoluteState_(sheet) { return reads_(sheet.getRange('E4').getFormula(), 'D4:D'); }

/** Relative content state: D4 is the formula computing the masses from the content. */
function isRelativeState_(sheet) { return reads_(sheet.getRange('D4').getFormula(), 'E4:E'); }

/** Did the edit put at least one value into the range? (Clearing cells is not typing.) */
function hasTyped_(range) {
  return range.getValues().some(r => r[0] !== '' && r[0] !== null);
}

/**
 * D3 was edited. In the absolute state D3 is the sum, so a typed value there switches to the relative state (D3 becomes the typed
 * total). An emptied D3 gets its sum back. In the relative state D3 is the typed total and nothing happens.
 */
function onTotalEdited_(sheet) {
  if (!isAbsoluteState_(sheet)) return;
  const total = sheet.getRange(MASS_TOTAL_CELL);
  if (reads_(total.getFormula(), 'D4:D')) return;
  if (String(total.getValue()).trim() === '') { setFormula_(sheet, MASS_TOTAL_CELL, TOTAL_SUM_FORMULA); return; }
  switchToRelative_(sheet, null);
}

/**
 * Writes a formula given with ; as the argument separator. Apps Script is documented to want , there, so when the sheet answers
 * with a parse error the same formula is written again with commas (none of the formulas has a ; inside a string).
 */
function setFormula_(sheet, a1, formula) {
  const cell = sheet.getRange(a1);
  cell.setFormula(formula);
  SpreadsheetApp.flush();
  if (/^#(ERROR|BŁĄD)/i.test(String(cell.getDisplayValue()).trim())) cell.setFormula(formula.replace(/;/g, ','));
}

/** The last row (at least HEAD_ROW) with something in one of the columns. */
function lastDataRow_(sheet, cols) {
  const last = sheet.getLastRow();
  let result = HEAD_ROW;
  if (last < HEAD_ROW) return result;
  cols.forEach(c => {
    const v = sheet.getRange(HEAD_ROW, c, last - HEAD_ROW + 1, 1).getValues();
    for (let i = v.length - 1; i >= 0; i--) {
      if (String(v[i][0]).trim() !== '') { result = Math.max(result, HEAD_ROW + i); break; }
    }
  });
  return result;
}

/** Puts the typed values of the edited range over `arr` (an array indexed from HEAD_ROW). */
function applyTyped_(arr, edited) {
  if (!edited) return;
  edited.getValues().forEach((r, i) => {
    if (r[0] !== '' && r[0] !== null) arr[edited.getRow() + i - HEAD_ROW] = r[0];
  });
}

/**
 * Absolute mass -> relative content. The contents (E4:E) become typed values: the share of every mass in the total of the masses,
 * the edited cells keeping what was typed. The total mass (D3) becomes a typed number too (what it was; or what the user typed
 * in it), and D4:D turns into D3 * E. `edited` = the typed cells of E4:E, or null when D3 was edited.
 */
function switchToRelative_(sheet, edited) {
  const last = Math.max(lastDataRow_(sheet, [KEY_COL, VALUE_COL]), edited ? edited.getLastRow() : 0);
  const n = last - HEAD_ROW + 1;
  const masses = sheet.getRange(HEAD_ROW, VALUE_COL, n, 1).getValues().map(r => r[0]);
  const sum = masses.reduce((s, v) => s + (typeof v === 'number' ? v : 0), 0);
  const content = masses.map(v => (typeof v === 'number' && sum ? v / sum : ''));
  applyTyped_(content, edited);

  if (edited) sheet.getRange(MASS_TOTAL_CELL).setValue(sum);           // the formula goes, the number stays
  sheet.getRange(HEAD_ROW, CONTENT_COL, n, 1).setValues(content.map(v => [v]));
  const maxRows = sheet.getMaxRows();
  sheet.getRange(HEAD_ROW, VALUE_COL, maxRows - HEAD_ROW + 1, 1).clearContent();
  setFormula_(sheet, 'D4', MASS_FORMULA);
}

/**
 * Relative content -> absolute mass. The masses (D4:D) become typed values: total mass * content of each row, the edited cells
 * keeping what was typed. D3 turns into the sum of the masses and E4:E into the content computed from them.
 * `edited` = the typed cells of D4:D.
 */
function switchToAbsolute_(sheet, edited) {
  const last = Math.max(lastDataRow_(sheet, [KEY_COL, CONTENT_COL]), edited ? edited.getLastRow() : 0);
  const n = last - HEAD_ROW + 1;
  const total = toNumber_(sheet.getRange(MASS_TOTAL_CELL).getValue());
  const content = sheet.getRange(HEAD_ROW, CONTENT_COL, n, 1).getValues().map(r => r[0]);
  const masses = content.map(v => (typeof v === 'number' && Number.isFinite(total) ? total * v : ''));
  applyTyped_(masses, edited);

  const maxRows = sheet.getMaxRows();
  sheet.getRange(HEAD_ROW, CONTENT_COL, maxRows - HEAD_ROW + 1, 1).clearContent();
  sheet.getRange(HEAD_ROW, VALUE_COL, n, 1).setValues(masses.map(v => [v]));
  setFormula_(sheet, MASS_TOTAL_CELL, TOTAL_SUM_FORMULA);
  setFormula_(sheet, 'E4', CONTENT_FORMULA);
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
 * Loads the source row `row` (the first one of its label) below the data range of the destination sheet. The loaded values are
 * masses, so a sheet in the relative content state is switched to the absolute mass state first (nothing is lost).
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

    if (isRelativeState_(sheet)) switchToAbsolute_(sheet, null);     // D4:D is spilled there

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
