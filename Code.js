// ---------------------------------------------------------------------------
// RECEPTURY: a sidebar that lists the keys of the "Asortyment" column of the ASSORT_SHEET sheet (a copy of the
// "Gotowe asortymenty" sheet from the wyceny project) and loads the matching "Receptura" cells into RECIPE_SHEET.
//
// Source sheet: a row whose column A is "Asortyment" holds the section titles, the next row the element headers,
// data rows follow (A = key). The "Receptura" column (section row, else header row) holds
//   "[[name1,pricePerKg1,price1];[name2,pricePerKg2,price2];...]"
// Destination sheet: one row per component:  Asortyment | Komponent | Cena za kg | Cena
// ---------------------------------------------------------------------------

const ASSORT_SHEET = 'Gotowe asortymenty';
const RECIPE_SHEET = 'Receptura';
const RECIPE_HEADERS = ['Asortyment', 'Komponent', 'Cena za kg', 'Cena'];

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
 * "Receptura" cell -> { components: [{name, pricePerKg, price}], error? }.
 * Fields are separated by commas, or by semicolons when a number has a decimal comma. The pair "[name;price]"
 * is accepted too (pricePerKg = 0). An empty cell gives no components.
 */
function parseRecipe_(text) {
  const raw = String(text === null || text === undefined ? '' : text).trim();
  if (!raw) return { components: [] };

  const num = s => {
    const t = s.trim();
    return t === '' ? 0 : Number(t.replace(/[\s ]/g, '').replace(',', '.'));
  };
  const components = [];
  const re = /\[([^\[\]]*)\]/g;                  // innermost brackets = one entry each
  let m;
  while ((m = re.exec(raw)) !== null) {
    const fields = m[1].indexOf(';') >= 0 ? m[1].split(';') : m[1].split(',');
    const name = (fields[0] || '').trim();
    if (!name) continue;
    if (fields.length !== 2 && fields.length !== 3) {
      return { components: [], error: 'Nieprawidłowy komponent „' + name + '” (oczekiwano: nazwa, cena za kg, cena).' };
    }
    const pricePerKg = fields.length === 3 ? num(fields[1]) : 0;
    const price = num(fields[fields.length - 1]);
    if (!Number.isFinite(pricePerKg) || !Number.isFinite(price)) {
      return { components: [], error: 'Nieprawidłowa cena komponentu „' + name + '”.' };
    }
    components.push({ name: name, pricePerKg: pricePerKg, price: price });
  }
  if (!components.length) return { components: [], error: 'Nie udało się odczytać pola Receptura.' };
  return { components: components };
}

/** The source sheet as { key: {row, components, error} } in sheet order; throws a Polish message on a bad layout. */
function readAssortments_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ASSORT_SHEET);
  if (!sheet) throw new Error('Brak arkusza "' + ASSORT_SHEET + '".');

  const values = sheet.getDataRange().getValues();
  const secIdx = values.findIndex(r => String(r[0]).trim() === 'Asortyment');
  if (secIdx < 0 || secIdx + 1 >= values.length) {
    throw new Error('Nie znaleziono nagłówka "Asortyment" w arkuszu "' + ASSORT_SHEET + '".');
  }
  const sec = values[secIdx], head = values[secIdx + 1];
  const keyIdx = sec.findIndex(v => String(v).trim() === 'Asortyment');
  let recipeIdx = sec.findIndex(v => String(v).trim() === 'Receptura');
  if (recipeIdx < 0) recipeIdx = head.findIndex(v => String(v).trim() === 'Receptura');
  if (recipeIdx < 0) throw new Error('Brak kolumny "Receptura" w arkuszu "' + ASSORT_SHEET + '".');

  const items = [];
  for (let r = secIdx + 2; r < values.length; r++) {
    const key = String(values[r][keyIdx]).trim();
    if (!key) continue;
    const recipe = parseRecipe_(values[r][recipeIdx]);
    items.push({ key: key, row: r + 1, components: recipe.components, error: recipe.error || '' });
  }
  return items;
}

/** For the sidebar: { items: [{key, count, error}] } | { items: [], error }. Never throws. */
function getAssortmentKeys() {
  try {
    return {
      items: readAssortments_().map(i => ({ key: i.key, count: i.components.length, error: i.error }))
    };
  } catch (e) {
    return { items: [], error: e.message };
  }
}

/**
 * Loads the receptury of `keys` into RECIPE_SHEET (created if missing). `append` = add under the existing rows,
 * otherwise the rows under the header are replaced. Checks everything first: on any problem nothing is changed.
 * Returns { ok, rows, loaded, empty } | { error }.
 */
function loadRecipes(keys, append) {
  try {
    if (!keys || !keys.length) return { error: 'Nie wybrano żadnego asortymentu.' };
    const byKey = {};
    readAssortments_().forEach(i => { if (!(i.key in byKey)) byKey[i.key] = i; });

    const problems = [], rows = [], empty = [];
    keys.forEach(k => {
      const item = byKey[k];
      if (!item) { problems.push(k + ': nie znaleziono asortymentu'); return; }
      if (item.error) { problems.push(k + ': ' + item.error); return; }
      if (!item.components.length) { empty.push(k); return; }
      item.components.forEach(c => rows.push([k, c.name, c.pricePerKg, c.price]));
    });
    if (problems.length) return { error: 'Nic nie zmieniono. ' + problems.join(' | ') };
    if (!rows.length) return { error: 'Wybrane asortymenty nie mają receptury.' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(RECIPE_SHEET) || ss.insertSheet(RECIPE_SHEET);
    sheet.getRange(1, 1, 1, RECIPE_HEADERS.length).setValues([RECIPE_HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);

    if (!append) {
      const last = sheet.getLastRow();
      if (last > 1) sheet.getRange(2, 1, last - 1, Math.max(sheet.getLastColumn(), RECIPE_HEADERS.length)).clearContent();
    }
    const start = append ? Math.max(sheet.getLastRow(), 1) + 1 : 2;
    sheet.getRange(start, 1, rows.length, RECIPE_HEADERS.length).setValues(rows);
    SpreadsheetApp.flush();
    return { ok: true, rows: rows.length, loaded: keys.length - empty.length, empty: empty };
  } catch (e) {
    return { error: e.message };
  }
}
