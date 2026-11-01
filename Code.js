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
// Rows of the destination are NOT fixed numbers: the row that holds the array formulas and the template data validation is
// the one with MARKER ("TECHNICZNY") in column A (see layout_). Everything is relative to it: the total row is the row right above,
// the entries start in the row right below.
//
// Destination: every component is appended below the data range (never above the row under the marker), the menu position in
// column B and its mass in column D. Nothing is merged with what is already there, so duplicates are possible. The cells
// RECIPE_NAME_CELL and PRODUCT_CELL get the Receptura name and the Produkt in parentheses, but only when empty or holding just "-".
//
// Formatting of the destination (formatRecipeSheet_, on every edit of column B and after every load): from the row under the
// marker to the last row with data in column B + 1, column B gets the data validation of the marker row where a cell has none,
// and B:G get alternating row backgrounds (first row white, second grey).
//
// Two states of the destination (see CLAUDE.md), r = the marker row:
//   absolute mass     D{r}:D (masses) are typed, D{r-1} = SUM(D{r}:D), E{r}:E (content %) = D / D${r-1}
//   relative content  E{r}:E (content %) are typed, D{r-1} (total mass) is typed, D{r}:D = D{r-1} * E
// Editing the computed column of a state switches to the other state, in which that column is the typed one. The values of that
// column are kept as they were (except the edited cells, which take the typed value) and the other column becomes computed.
// Values pasted over both columns at once need a decision. A SIMPLE trigger has no permission to open a dialog ("Określone
// uprawnienia nie są wystarczające do wywołania Ui.showModalDialog"), so an INSTALLABLE edit trigger (onEditInstalled, switched
// on once from the menu) opens PasteDialog.html asking which column to keep.
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
const MARKER = 'TECHNICZNY';
const MARKER_SEARCH_ROWS = 200;
const RECIPE_NAME_CELL = 'B1';
const PRODUCT_CELL = 'B2';
const ROW_WHITE = '#ffffff';
const ROW_GREY = '#f3f3f3';
const PASTE_TRIGGER_HANDLER = 'onEditInstalled';

/**
 * Where things are on the destination sheet, from the row with MARKER in column A:
 * { head (the marker row: array formulas, template validation, first row of the mass / content columns),
 *   total (the row above: total mass), first (the first row the loader may write to), totalCell, validationCell,
 *   massCell, contentCell, massRange, contentRange }. Throws a Polish message when there is no marker.
 */
function layout_(sheet) {
  const n = Math.min(sheet.getMaxRows(), MARKER_SEARCH_ROWS);
  const col = sheet.getRange(1, 1, n, 1).getValues();
  const i = col.findIndex(r => String(r[0]).trim().toUpperCase() === MARKER);
  if (i < 0) throw new Error('Nie znaleziono znacznika "' + MARKER + '" w kolumnie A arkusza "' + sheet.getName() + '".');
  const head = i + 1;
  if (head < 2) throw new Error('Znacznik "' + MARKER + '" nie może być w pierwszym wierszu – nad nim jest wiersz sumy.');
  return {
    head: head,
    total: head - 1,
    first: head + 1,
    totalCell: 'D' + (head - 1),
    validationCell: 'B' + head,
    massCell: 'D' + head,
    contentCell: 'E' + head,
    massRange: 'D' + head + ':D',
    contentRange: 'E' + head + ':E'
  };
}

/**
 * The formulas that differ between the two states (written with ; as the argument separator, see setFormula_). Everything else in
 * the header block (price, cost, share) is the same in both states and is not touched.
 */
function formulas_(L) {
  return {
    totalSum: '=SUM(' + L.massRange + ')',                                                                                // total, absolute mass
    content: '=ArrayFormula(IF(' + L.massRange + '<>"";' + L.massRange + '/D$' + L.total + ';))',                         // content cell, absolute mass
    mass: '=ArrayFormula(IF(' + L.contentRange + '<>"";D' + L.total + '*' + L.contentRange + ';))'                         // mass cell, relative content
  };
}

/** Is the installable paste trigger there? false also when the project triggers cannot be read (a simple onOpen may lack the permission). */
function hasPasteTrigger_() {
  try {
    return ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === PASTE_TRIGGER_HANDLER);
  } catch (err) {
    return false;
  }
}

/** (Re)builds the menu; the item that switches the paste trigger on is there only while the trigger is not. */
function buildMenu_(pasteTriggerOn) {
  const menu = SpreadsheetApp.getUi().createMenu('Receptury').addItem('Wczytaj recepturę', 'showSidebar');
  if (!pasteTriggerOn) menu.addItem('Włącz pytanie przy wklejeniu w D i E', 'installPasteTrigger');
  menu.addToUi();
}

function onOpen() {
  buildMenu_(hasPasteTrigger_());
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
  buildMenu_(true);
  ui.alert(exists
    ? 'Wyzwalacz jest już włączony.'
    : 'Włączono. Wklejenie wartości do kolumn D i E zapyta, którą kolumnę zachować.');
}

/** An error in a trigger: written to the execution log and shown as a toast (a trigger has nobody to throw to). */
function reportError_(where, err) {
  console.error(where + ': ' + (err && err.stack ? err.stack : err));
  try {
    SpreadsheetApp.getActive().toast(where + ': ' + (err && err.message ? err.message : err), 'Receptury – błąd', 15);
  } catch (ignored) {
    // no toast available
  }
}

/**
 * Simple trigger on the destination sheet:
 *  - an edit of the total mass or of the mass column (the computed cells of the relative state) / the content column (the computed
 *    cells of the absolute state) switches the state;
 *  - an edit that touches column B under the marker row re-applies the formatting.
 * The state is read from the marker-row formula of the column the edit did NOT touch, because the edit may have overwritten its own.
 * (An edit over D and E together is left to onEditInstalled.)
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const range = e.range;
    const sheet = range.getSheet();
    if (sheet.getName() !== RECIPE_SHEET) return;
    const L = layout_(sheet);

    if (range.getA1Notation() === L.totalCell) {
      onTotalEdited_(sheet, L);
      return;
    }

    const oneColumn = range.getColumn() === range.getLastColumn();
    if (oneColumn && range.getRow() >= L.head) {
      if (range.getColumn() === VALUE_COL) {
        if (!isAbsoluteState_(sheet, L) && hasTyped_(range)) switchToAbsolute_(sheet, L, range);
        return;
      }
      if (range.getColumn() === CONTENT_COL) {
        if (!isRelativeState_(sheet, L) && hasTyped_(range)) switchToRelative_(sheet, L, range);
        return;
      }
    }

    if (range.getLastRow() < L.first) return;
    if (range.getColumn() > KEY_COL || range.getLastColumn() < KEY_COL) return;
    formatRecipeSheet_(sheet, L);
  } catch (err) {
    reportError_('onEdit', err);
  }
}

/**
 * Installable edit trigger (see installPasteTrigger): values pasted over D and E together. Both columns now hold real values, so
 * which one is to be kept has to be asked; the dialog also lets the values (and the total mass) be corrected first.
 * Does nothing when the cells were only cleared.
 */
function onEditInstalled(e) {
  try {
    if (!e || !e.range) return;
    const range = e.range;
    const sheet = range.getSheet();
    if (sheet.getName() !== RECIPE_SHEET) return;
    if (range.getColumn() > VALUE_COL || range.getLastColumn() < CONTENT_COL) return;
    const L = layout_(sheet);

    const first = Math.max(range.getRow(), L.head);
    const last = range.getLastRow();
    if (last < first) return;
    const n = last - first + 1;
    const d = sheet.getRange(first, VALUE_COL, n, 1).getValues();
    const c = sheet.getRange(first, CONTENT_COL, n, 1).getValues();
    const names = sheet.getRange(first, KEY_COL, n, 1).getValues();
    const rows = d.map((r, i) => ({ row: first + i, name: String(names[i][0]), d: r[0], e: c[i][0] }));
    if (!rows.some(r => r.d !== '' || r.e !== '')) return;             // clearing is not pasting

    const priorAbsolute = reads_(sheet.getRange(L.totalCell).getFormula(), L.massRange);
    let defaultTotal;
    if (priorAbsolute) {
      const all = sheet.getRange(L.head, VALUE_COL, sheet.getMaxRows() - L.head + 1, 1).getValues();
      defaultTotal = all.reduce((s, r) => s + (typeof r[0] === 'number' ? r[0] : 0), 0);
    } else {
      const t = toNumber_(sheet.getRange(L.totalCell).getValue());
      defaultTotal = Number.isFinite(t) ? t : null;
    }

    const payload = { keep: priorAbsolute ? 'D' : 'E', priorAbsolute: priorAbsolute, defaultTotal: defaultTotal, rows: rows };
    const t = HtmlService.createTemplateFromFile('PasteDialog');
    t.payload = JSON.stringify(payload).replace(/</g, '\\u003c');
    SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(520).setHeight(560), 'Wklejono wartości do kolumn D i E');
  } catch (err) {
    reportError_('okno wklejania', err);
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
    const L = layout_(sheet);
    const F = formulas_(L);

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
        sheet.getRange(L.head, CONTENT_COL, maxRows - L.head + 1, 1).clearContent();
        setFormula_(sheet, L.contentCell, F.content);
      } else {
        switchToAbsolute_(sheet, L, dRange);     // the other rows: total mass * their typed contents
      }
      return { ok: true };
    }

    const total = toNumber_(p.total);
    if (!Number.isFinite(total)) return { error: 'Podaj masę łączną (liczbę).' };
    const last = Math.max(lastDataRow_(sheet, L, [KEY_COL, VALUE_COL, CONTENT_COL]), lastPasted);
    const n = last - L.head + 1;
    let content;
    if (p.priorAbsolute) {
      // contents of the rows that were not pasted: their mass / the total the dialog started with
      const t0 = Number(p.defaultTotal);
      content = sheet.getRange(L.head, VALUE_COL, n, 1).getValues().map(r => (typeof r[0] === 'number' && t0 ? r[0] / t0 : ''));
    } else {
      content = sheet.getRange(L.head, CONTENT_COL, n, 1).getValues().map(r => r[0]);
    }
    rows.forEach(r => { content[r.row - L.head] = r.e; });

    sheet.getRange(L.head, CONTENT_COL, n, 1).setValues(content.map(v => [v]));
    sheet.getRange(L.totalCell).setValue(total);
    sheet.getRange(L.head, VALUE_COL, maxRows - L.head + 1, 1).clearContent();
    setFormula_(sheet, L.massCell, F.mass);
    SpreadsheetApp.flush();
    return { ok: true };
  } catch (err) {
    return { error: err.message };
  }
}

/** Does the formula read the given range (e.g. "D5:D")? */
function reads_(formula, a1) {
  return String(formula).replace(/\$/g, '').toUpperCase().indexOf(a1) >= 0;
}

/** Absolute mass state: the content cell is the formula computing the content from the masses. */
function isAbsoluteState_(sheet, L) { return reads_(sheet.getRange(L.contentCell).getFormula(), L.massRange); }

/** Relative content state: the mass cell is the formula computing the masses from the content. */
function isRelativeState_(sheet, L) { return reads_(sheet.getRange(L.massCell).getFormula(), L.contentRange); }

/** Did the edit put at least one value into the range? (Clearing cells is not typing.) */
function hasTyped_(range) {
  return range.getValues().some(r => r[0] !== '' && r[0] !== null);
}

/**
 * The total mass was edited. In the absolute state it is the sum, so a typed value there switches to the relative state (it becomes
 * the typed total). An emptied total gets its sum back. In the relative state it is the typed total and nothing happens.
 */
function onTotalEdited_(sheet, L) {
  if (!isAbsoluteState_(sheet, L)) return;
  const total = sheet.getRange(L.totalCell);
  if (reads_(total.getFormula(), L.massRange)) return;
  if (String(total.getValue()).trim() === '') { setFormula_(sheet, L.totalCell, formulas_(L).totalSum); return; }
  switchToRelative_(sheet, L, null);
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

/** The last row (at least the marker row) with something in one of the columns. */
function lastDataRow_(sheet, L, cols) {
  const last = sheet.getLastRow();
  let result = L.head;
  if (last < L.head) return result;
  cols.forEach(c => {
    const v = sheet.getRange(L.head, c, last - L.head + 1, 1).getValues();
    for (let i = v.length - 1; i >= 0; i--) {
      if (String(v[i][0]).trim() !== '') { result = Math.max(result, L.head + i); break; }
    }
  });
  return result;
}

/** Puts the typed values of the edited range over `arr` (an array indexed from the marker row). */
function applyTyped_(arr, L, edited) {
  if (!edited) return;
  edited.getValues().forEach((r, i) => {
    if (r[0] !== '' && r[0] !== null) arr[edited.getRow() + i - L.head] = r[0];
  });
}

/**
 * Absolute mass -> relative content. The contents become typed values: the share of every mass in the total of the masses,
 * the edited cells keeping what was typed. The total mass becomes a typed number too (what it was; or what the user typed
 * in it), and the mass column turns into total * content. `edited` = the typed cells of the content column, or null when the
 * total was edited.
 */
function switchToRelative_(sheet, L, edited) {
  const last = Math.max(lastDataRow_(sheet, L, [KEY_COL, VALUE_COL]), edited ? edited.getLastRow() : 0);
  const n = last - L.head + 1;
  const masses = sheet.getRange(L.head, VALUE_COL, n, 1).getValues().map(r => r[0]);
  const sum = masses.reduce((s, v) => s + (typeof v === 'number' ? v : 0), 0);
  const content = masses.map(v => (typeof v === 'number' && sum ? v / sum : ''));
  applyTyped_(content, L, edited);

  if (edited) sheet.getRange(L.totalCell).setValue(sum);               // the formula goes, the number stays
  sheet.getRange(L.head, CONTENT_COL, n, 1).setValues(content.map(v => [v]));
  const maxRows = sheet.getMaxRows();
  sheet.getRange(L.head, VALUE_COL, maxRows - L.head + 1, 1).clearContent();
  setFormula_(sheet, L.massCell, formulas_(L).mass);
}

/**
 * Relative content -> absolute mass. The masses become typed values: total mass * content of each row, the edited cells
 * keeping what was typed. The total turns into the sum of the masses and the content column into the content computed from them.
 * `edited` = the typed cells of the mass column.
 */
function switchToAbsolute_(sheet, L, edited) {
  const last = Math.max(lastDataRow_(sheet, L, [KEY_COL, CONTENT_COL]), edited ? edited.getLastRow() : 0);
  const n = last - L.head + 1;
  const total = toNumber_(sheet.getRange(L.totalCell).getValue());
  const content = sheet.getRange(L.head, CONTENT_COL, n, 1).getValues().map(r => r[0]);
  const masses = content.map(v => (typeof v === 'number' && Number.isFinite(total) ? total * v : ''));
  applyTyped_(masses, L, edited);

  const maxRows = sheet.getMaxRows();
  sheet.getRange(L.head, CONTENT_COL, maxRows - L.head + 1, 1).clearContent();
  sheet.getRange(L.head, VALUE_COL, n, 1).setValues(masses.map(v => [v]));
  const F = formulas_(L);
  setFormula_(sheet, L.totalCell, F.totalSum);
  setFormula_(sheet, L.contentCell, F.content);
}

/**
 * The row under the marker .. (last row with data in column B) + 1: data validation copied from the marker row's cell of column B
 * into the cells of column B that have none (existing validations are not checked), and alternating backgrounds in B:G
 * (first row white, second grey).
 */
function formatRecipeSheet_(sheet, L) {
  const maxRows = sheet.getMaxRows();
  if (maxRows < L.first) return;

  const lastRow = sheet.getLastRow();
  let lastData = L.first - 1;
  if (lastRow >= L.first) {
    const keys = sheet.getRange(L.first, KEY_COL, lastRow - L.first + 1, 1).getValues();
    for (let i = keys.length - 1; i >= 0; i--) {
      if (String(keys[i][0]).trim() !== '') { lastData = L.first + i; break; }
    }
  }
  const end = Math.min(lastData + 1, maxRows);
  const count = end - L.first + 1;

  const keyRange = sheet.getRange(L.first, KEY_COL, count, 1);
  const template = sheet.getRange(L.validationCell).getDataValidation();
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
  sheet.getRange(L.first, KEY_COL, count, LAST_FORMAT_COL - KEY_COL + 1).setBackgrounds(colors);
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
    const L = layout_(sheet);

    if (isRelativeState_(sheet, L)) switchToAbsolute_(sheet, L, null);     // the mass column is spilled there

    // the first empty row, as the data range tells, but never above the row under the marker
    const start = Math.max(sheet.getDataRange().getLastRow() + 1, L.first);
    sheet.getRange(start, KEY_COL, components.length, 1).setValues(components.map(c => [c.name]));
    sheet.getRange(start, VALUE_COL, components.length, 1).setValues(components.map(c => [c.value]));

    fillIfBlank_(sheet, RECIPE_NAME_CELL, chosen.recipe);
    fillIfBlank_(sheet, PRODUCT_CELL, chosen.product ? '(' + chosen.product + ')' : '');

    formatRecipeSheet_(sheet, L);    // a script write does not fire onEdit

    SpreadsheetApp.flush();
    return { ok: true, written: components.length, warning: warnings.join('\n') };
  } catch (e) {
    return { error: e.message };
  }
}
