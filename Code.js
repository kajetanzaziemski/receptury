// ---------------------------------------------------------------------------
// RECEPTURY: a sidebar that loads ONE row of the SOURCE_SHEET ("Zapisane receptury") into the RECIPE_SHEET ("Receptura").
//
// Source sheet (headers in row 1): Produkt | Receptura | Asortyment | Receptura - skład
//   "Receptura - skład" = "[[component;value];[component;value];...]"  (one value per component)
// The user picks an identifier column (Produkt / Receptura / Asortyment), then one of its unique values; the FIRST row
// with that value is loaded. If the rows with that value hold more than one distinct "Receptura - skład", a warning is returned.
//
// Destination: components (keys) sit in column B from row 4 down. The value of each component goes to column D of the
// row whose B matches it. A component that is not in column B is added in the first empty row below the data range.
// ---------------------------------------------------------------------------

const SOURCE_SHEET = 'Zapisane receptury';
const RECIPE_SHEET = 'Receptura';
const ID_COLUMNS = ['Receptura', 'Produkt', 'Asortyment'];   // first = the default in the sidebar
const RECIPE_COLUMN = 'Receptura - skład';
const KEY_COL = 2;          // B
const VALUE_COL = 4;        // D
const FIRST_ROW = 4;

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

/** The source sheet: { values (rows under the header), idx: {header: column index} }; throws a Polish message. */
function readSource_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SOURCE_SHEET);
  if (!sheet) throw new Error('Brak arkusza "' + SOURCE_SHEET + '".');
  const all = sheet.getDataRange().getValues();
  const head = all[0].map(h => String(h).trim());
  const idx = {};
  ID_COLUMNS.concat([RECIPE_COLUMN]).forEach(h => {
    idx[h] = head.indexOf(h);
    if (idx[h] < 0) throw new Error('Brak kolumny "' + h + '" w arkuszu "' + SOURCE_SHEET + '".');
  });
  return { values: all.slice(1), idx: idx };
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
    const n = Number(v.replace(/[\s ]/g, '').replace(',', '.'));
    out.push({ name: name, value: v === '' ? '' : (Number.isFinite(n) ? n : v) });
  }
  return out;
}

/** For the sidebar: the identifier columns offered, the first being the default. */
function getIdColumns() { return ID_COLUMNS; }

/** Unique non-empty values of one identifier column, in sheet order: { values } | { values: [], error }. Never throws. */
function getUniqueValues(column) {
  try {
    if (ID_COLUMNS.indexOf(column) < 0) return { values: [], error: 'Nieznana kolumna: ' + column };
    const src = readSource_();
    const seen = {}, values = [];
    src.values.forEach(r => {
      const v = String(r[src.idx[column]]).trim();
      if (v && !seen[v]) { seen[v] = true; values.push(v); }
    });
    return { values: values };
  } catch (e) {
    return { values: [], error: e.message };
  }
}

/**
 * Loads the first source row whose `column` equals `value` into the destination sheet.
 * Returns { ok, written, added, missingNames, warning? } | { error }.
 */
function loadRecipe(column, value) {
  try {
    if (ID_COLUMNS.indexOf(column) < 0) return { error: 'Nieznana kolumna: ' + column };
    const src = readSource_();
    const matches = src.values.filter(r => String(r[src.idx[column]]).trim() === value);
    if (!matches.length) return { error: 'Nie znaleziono "' + value + '" w kolumnie "' + column + '".' };

    const texts = matches.map(r => String(r[src.idx[RECIPE_COLUMN]]).trim());
    const distinct = texts.filter((t, i) => texts.indexOf(t) === i);
    const components = parseRecipe_(texts[0]);
    if (!components.length) return { error: 'Brak odczytywalnego składu w pierwszym dopasowaniu dla "' + value + '".' };
    const warning = distinct.length > 1
      ? 'Znaleziono ' + distinct.length + ' różnych pozycji "' + RECIPE_COLUMN + '" dla "' + value + '" – wczytano pierwszą.'
      : '';

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RECIPE_SHEET);
    if (!sheet) return { error: 'Brak arkusza "' + RECIPE_SHEET + '".' };

    // rows of the existing components (column B from FIRST_ROW down); the data range tells where the free rows start
    const last = Math.max(sheet.getDataRange().getLastRow(), FIRST_ROW - 1);
    const keys = last >= FIRST_ROW ? sheet.getRange(FIRST_ROW, KEY_COL, last - FIRST_ROW + 1, 1).getValues() : [];
    const rowOf = {};
    keys.forEach((r, i) => {
      const k = String(r[0]).trim().toLowerCase();
      if (k && !(k in rowOf)) rowOf[k] = FIRST_ROW + i;
    });

    let next = last + 1;
    const added = [];
    const writes = [];
    components.forEach(c => {
      const k = c.name.toLowerCase();
      let row = rowOf[k];
      if (row === undefined) {
        row = next++;
        rowOf[k] = row;
        added.push(c.name);
        sheet.getRange(row, KEY_COL).setValue(c.name);
      }
      writes.push([row, c.value]);
    });
    writes.forEach(w => sheet.getRange(w[0], VALUE_COL).setValue(w[1]));
    SpreadsheetApp.flush();
    return { ok: true, written: writes.length, added: added, warning: warning };
  } catch (e) {
    return { error: e.message };
  }
}
