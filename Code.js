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
// its value in column D. Nothing is merged with what is already there, so duplicates are possible. The data validation of
// VALIDATION_CELL (B4) is copied onto the new cells of column B.
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
const FIRST_ROW = 5;
const VALIDATION_CELL = 'B4';

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

/** A number from a cell (a decimal comma and spaces are fine); NaN when empty or not a number. */
function toNumber_(v) {
  if (typeof v === 'number') return v;
  const t = String(v === null || v === undefined ? '' : v).replace(/[\s ]/g, '').replace(',', '.');
  return t === '' ? NaN : Number(t);
}

/** The source sheet as rows: [{ row (sheet row number), label, text (Receptura - skład) }]; throws a Polish message. */
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
    rows.push({ row: i + 2, label: product + ' (' + recipe + ')', text: cell(r, RECIPE_COLUMN) });
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
    const keyRange = sheet.getRange(start, KEY_COL, components.length, 1);
    keyRange.setValues(components.map(c => [c.name]));
    sheet.getRange(start, VALUE_COL, components.length, 1).setValues(components.map(c => [c.value]));

    // the same data validation as the template cell (none there = nothing to copy)
    const rule = sheet.getRange(VALIDATION_CELL).getDataValidation();
    if (rule) keyRange.setDataValidation(rule);

    SpreadsheetApp.flush();
    return { ok: true, written: components.length, warning: warnings.join('\n') };
  } catch (e) {
    return { error: e.message };
  }
}
