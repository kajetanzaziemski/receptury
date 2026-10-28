// ---------------------------------------------------------------------------
// RECEPTURY: a sidebar that loads ONE row of the SOURCE_SHEET ("Zapisane receptury") into the RECIPE_SHEET ("Receptura").
//
// Source sheet (headers in row 1): Produkt | Receptura | Asortyment | Receptura - skład
//   "Receptura - skład" = "[[component;value];[component;value];...]"  (one value per component)
// The sidebar offers one dropdown with an entry per row: "Produkt (Receptura / Asortyment)". Entries with the same label are
// shown once; the FIRST row is loaded, and a warning is returned if the rows with that label hold more than one distinct
// "Receptura - skład".
//
// Destination: components (keys) sit in column B from row 4 down. The value of each component goes to column D of the
// row whose B matches it. A component that is not in column B is added in the first empty row below the data range.
// ---------------------------------------------------------------------------

const SOURCE_SHEET = 'Zapisane receptury';
const RECIPE_SHEET = 'Receptura';
const PRODUCT_COLUMN = 'Produkt';
const RECIPE_NAME_COLUMN = 'Receptura';
const ASSORT_COLUMN = 'Asortyment';
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

/** The source sheet as rows: [{ row (sheet row number), label, text (Receptura - skład) }]; throws a Polish message. */
function readSource_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SOURCE_SHEET);
  if (!sheet) throw new Error('Brak arkusza "' + SOURCE_SHEET + '".');
  const all = sheet.getDataRange().getValues();
  const head = all[0].map(h => String(h).trim());
  const idx = {};
  [PRODUCT_COLUMN, RECIPE_NAME_COLUMN, ASSORT_COLUMN, RECIPE_COLUMN].forEach(h => {
    idx[h] = head.indexOf(h);
    if (idx[h] < 0) throw new Error('Brak kolumny "' + h + '" w arkuszu "' + SOURCE_SHEET + '".');
  });
  const cell = (r, h) => String(r[idx[h]] === null || r[idx[h]] === undefined ? '' : r[idx[h]]).trim();
  const rows = [];
  all.slice(1).forEach((r, i) => {
    const product = cell(r, PRODUCT_COLUMN), recipe = cell(r, RECIPE_NAME_COLUMN), assort = cell(r, ASSORT_COLUMN);
    if (!product && !recipe && !assort) return;
    rows.push({ row: i + 2, label: product + ' (' + recipe + ' / ' + assort + ')', text: cell(r, RECIPE_COLUMN) });
  });
  return rows;
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
 * Loads the source row `row` (the first one of its label) into the destination sheet.
 * Returns { ok, written, added, warning } | { error }.
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
    const warning = distinct.length > 1
      ? 'Znaleziono ' + distinct.length + ' różnych pozycji "' + RECIPE_COLUMN + '" dla "' + chosen.label + '" – wczytano pierwszą.'
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
      let target = rowOf[k];
      if (target === undefined) {
        target = next++;
        rowOf[k] = target;
        added.push(c.name);
        sheet.getRange(target, KEY_COL).setValue(c.name);
      }
      writes.push([target, c.value]);
    });
    writes.forEach(w => sheet.getRange(w[0], VALUE_COL).setValue(w[1]));
    SpreadsheetApp.flush();
    return { ok: true, written: writes.length, added: added, warning: warning };
  } catch (e) {
    return { error: e.message };
  }
}
