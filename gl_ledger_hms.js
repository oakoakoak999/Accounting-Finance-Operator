const { chromium } = require('playwright');
const XLSX = require('xlsx');
const { google } = require('googleapis');
const path = require('path');
const fs = require('fs');
const os = require('os');

const SHEET_ID  = '157EM10aUaWiD6cWehjD20p1YEzsklNqG_IuF1XDI9YQ';
const SHEET_TAB = 'GL';

function getSheetClient() {
  const tokenPath = path.join(__dirname, '.gsheets-token.json');
  if (!fs.existsSync(tokenPath))
    throw new Error('No .gsheets-token.json — run: node get-sheets-token.mjs');
  const auth = new google.auth.OAuth2(env.GMAIL_CLIENT_ID, env.GMAIL_CLIENT_SECRET, env.GMAIL_REDIRECT_URI);
  auth.setCredentials(JSON.parse(fs.readFileSync(tokenPath, 'utf8')));
  auth.on('tokens', tokens => {
    const existing = JSON.parse(fs.readFileSync(tokenPath, 'utf8'));
    fs.writeFileSync(tokenPath, JSON.stringify({ ...existing, ...tokens }, null, 2), 'utf8');
  });
  return google.sheets({ version: 'v4', auth });
}

// Load credentials from .env
const envPath = path.join(__dirname, '.env');
const env = Object.fromEntries(
  fs.readFileSync(envPath, 'utf8')
    .split('\n')
    .filter(l => l.includes('='))
    .map(l => l.split('=').map(s => s.trim()))
);

const BASE_URL = env.SMARTERP_URL;
const USERNAME = env.SMARTERP_USERNAME;
const PASSWORD = env.SMARTERP_PASSWORD;

// The report's numeric menu_id/action and Studio field record IDs change on every ERP update,
// so everything below is resolved at runtime from these stable names instead of being hardcoded.
const REPORT_MENU_PATH = 'Accounting/Query/General Ledger/GL Ledger all Detail (HMS)';
const REPORT_MENU_NAME = 'GL Ledger all Detail (HMS)';

// Call an Odoo model method over the authenticated web session (JSON-RPC).
async function callKw(page, model, method, args, kwargs = {}) {
  const res = await page.evaluate(async ({ model, method, args, kwargs }) => {
    const r = await fetch('/web/dataset/call_kw', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ jsonrpc: '2.0', method: 'call', params: { model, method, args, kwargs } }),
    });
    return r.json();
  }, { model, method, args, kwargs });
  if (res.error) throw new Error(`call_kw ${model}.${method}: ${res.error.data?.message || res.error.message}`);
  return res.result;
}

// CLI args: node gl_ledger_hms.js <fromMonth> <toMonth> <year>
// e.g. node gl_ledger_hms.js jan may 2026
const MONTH_NAMES = ['January','February','March','April','May','June',
                     'July','August','September','October','November','December'];

function parseMonthArg(arg) {
  const lower = arg.toLowerCase();
  let idx = MONTH_NAMES.findIndex(m => m.toLowerCase().startsWith(lower));
  if (idx === -1) idx = parseInt(arg) - 1;
  if (isNaN(idx) || idx < 0 || idx > 11) throw new Error(`Invalid month: "${arg}"`);
  return idx; // 0-based
}

const _rawArgs = process.argv.slice(2);
const HEADLESS = _rawArgs.some(a => a.toLowerCase() === '--headless');
const _args = _rawArgs.filter(a => a.toLowerCase() !== '--headless');
if (_args.length < 3)
  throw new Error('Usage: node gl_ledger_hms.js <fromMonth> <toMonth> <year> [--headless]\nExample: node gl_ledger_hms.js jan may 2026 --headless');

const FROM_MONTH_IDX = parseMonthArg(_args[0]);
const TO_MONTH_IDX   = parseMonthArg(_args[1]);
const YEAR           = parseInt(_args[2]);
if (isNaN(YEAR)) throw new Error(`Invalid year: "${_args[2]}"`);

const TO_DAY = new Date(YEAR, TO_MONTH_IDX + 1, 0).getDate(); // last day of TO month

const fromParsed    = { month: MONTH_NAMES[FROM_MONTH_IDX], year: String(YEAR), day: '1' };
const toParsed      = { month: MONTH_NAMES[TO_MONTH_IDX],   year: String(YEAR), day: String(TO_DAY) };
const FROM_YYYYMMDD = `${YEAR}-${String(FROM_MONTH_IDX + 1).padStart(2,'0')}-01`;
const TO_YYYYMMDD   = `${YEAR}-${String(TO_MONTH_IDX + 1).padStart(2,'0')}-${String(TO_DAY).padStart(2,'0')}`;

console.log(`Date range: ${FROM_YYYYMMDD} → ${TO_YYYYMMDD} (to-day: ${TO_DAY})`);

(async () => {
  const browser = await chromium.launch({ headless: HEADLESS, channel: 'chrome' });
  const context = await browser.newContext();
  const page    = await context.newPage();

  // Login
  await page.goto(`${BASE_URL}/web/login`, { waitUntil: 'load' });
  await page.fill('#login', USERNAME);
  await page.fill('#password', PASSWORD);
  await page.click('button[type=submit]');
  await page.waitForLoadState('load');
  await page.waitForTimeout(2000);

  // Resolve the report's current menu_id + action by its stable menu path (IDs churn on ERP updates)
  const menus = await callKw(page, 'ir.ui.menu', 'search_read',
    [[['name', '=', REPORT_MENU_NAME]], ['id', 'complete_name', 'action']], { context: { lang: 'en_US' } });
  const menu = menus.find(m => m.complete_name === REPORT_MENU_PATH) || menus.find(m => m.action);
  if (!menu || !menu.action)
    throw new Error(`Could not resolve report menu "${REPORT_MENU_PATH}" (found ${menus.length} candidates)`);
  const actionId = String(menu.action).split(',').pop(); // "ir.actions.act_window,1618" -> "1618"
  console.log(`Resolved report: menu_id=${menu.id} action=${actionId}`);
  await page.goto(`${BASE_URL}/web#menu_id=${menu.id}&action=${actionId}&cids=2`, { waitUntil: 'load', timeout: 60000 });
  await page.waitForTimeout(3000);

  // Resolve the Studio field IDs from the DOM (numeric record IDs change on ERP updates).
  // Both date fields share the ..._amdate_0 suffix; the lower record id is From, the higher is To.
  await page.waitForSelector('[id^="x_x_gl_ledger_all_descr_hms_"][id$="_amdate_0"]', { timeout: 30000 });
  const fieldIds = await page.evaluate(() => {
    const dates = [...document.querySelectorAll('[id^="x_x_gl_ledger_all_descr_hms_"][id$="_amdate_0"]')]
      .map(e => e.id)
      .sort((a, b) => parseInt(a.match(/_(\d+)_amdate_0$/)[1]) - parseInt(b.match(/_(\d+)_amdate_0$/)[1]));
    const bu = document.querySelector('[id^="x_x_gl_ledger_all_descr_hms_"][id$="_rcid_0"]');
    return { fromId: dates[0], toId: dates[1], buId: bu ? bu.id : null };
  });
  if (!fieldIds.fromId || !fieldIds.toId || !fieldIds.buId)
    throw new Error(`Could not resolve GL form fields: ${JSON.stringify(fieldIds)}`);
  console.log(`Resolved fields: from=${fieldIds.fromId} to=${fieldIds.toId} bu=${fieldIds.buId}`);

  // Helper: open datepicker for a field and navigate to the target month/day
  async function pickDate(fieldId, targetMonth, targetYear, day) {
    const months = ['January', 'February', 'March', 'April', 'May', 'June',
                    'July', 'August', 'September', 'October', 'November', 'December'];
    const target = `${targetMonth} ${targetYear}`;
    await page.locator(`#${fieldId}`).click();
    await page.waitForTimeout(800);

    let header = await page.$eval('.o_datetime_picker .o_header_part', el => el.textContent.trim());
    let attempts = 0;
    while (header !== target && attempts < 24) {
      const [curM, curY] = header.split(' ');
      const curIdx = parseInt(curY) * 12 + months.indexOf(curM);
      const tarIdx = parseInt(targetYear) * 12 + months.indexOf(targetMonth);
      await page.click(curIdx < tarIdx ? '.o_datetime_picker .o_next' : '.o_datetime_picker .o_previous');
      await page.waitForTimeout(300);
      header = await page.$eval('.o_datetime_picker .o_header_part', el => el.textContent.trim());
      attempts++;
    }

    await page.click(`.o_datetime_picker .o_date_item_cell:has-text("${day}"):not(.o_out_of_scope)`);
    await page.waitForTimeout(300);
    await page.keyboard.press('Tab');
    await page.waitForTimeout(300);
  }

  // Set From date
  await pickDate(fieldIds.fromId, fromParsed.month, fromParsed.year, fromParsed.day);

  // Set To date
  await pickDate(fieldIds.toId, toParsed.month, toParsed.year, toParsed.day);

  // Set BU -> HMS
  const buField = page.locator(`#${fieldIds.buId}`);
  await buField.click();
  await buField.fill('HMS');
  await page.waitForTimeout(1000);
  await page.click('.o-autocomplete--dropdown-item:first-child');
  await page.waitForTimeout(500);

  const today = new Date().toISOString().split('T')[0]; // YYYY-MM-DD
  const downloadPath = path.join(os.homedir(), 'Downloads', `GL Ledger all Detail (HMS) ${today}.xlsx`);

  // Capture the call_button response to get the dynamic download URL
  const callButtonPromise = page.waitForResponse(r => r.url().includes('call_button'));

  // Click Print
  await page.click('button.btn.btn-primary:has-text("Print")');

  // Extract download URL from the action response
  const callButtonRes = await callButtonPromise;
  const action = (await callButtonRes.json()).result;
  const fileUrl = `${BASE_URL}/${action.url}`;
  console.log('Downloading from:', fileUrl);

  // Fetch the file inside the browser context (inherits session cookies)
  const fileBytes = await page.evaluate(async (url) => {
    const res = await fetch(url);
    const buf = await res.arrayBuffer();
    return Array.from(new Uint8Array(buf));
  }, fileUrl);
  fs.writeFileSync(downloadPath, Buffer.from(fileBytes));
  console.log(`Saved ${fileBytes.length} bytes -> ${downloadPath}`);

  // Apply filters using xlsx — no Excel process needed
  const wb = XLSX.readFile(downloadPath);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

  // Find header row dynamically (look for 'ACCOUNT CODE')
  const headerRowIdx = data.findIndex(row => row && row.some(cell => String(cell ?? '').toUpperCase() === 'ACCOUNT CODE'));
  if (headerRowIdx === -1) throw new Error('Header row not found in XLSX');
  const dataStart = headerRowIdx + 1;
  const lastRow   = data.length;
  if (!ws['!rows']) ws['!rows'] = [];

  // Find column indices from header row
  const headerRow   = data[headerRowIdx];
  const acctDtCol   = headerRow.findIndex(h => String(h ?? '').toUpperCase() === 'ACCOUNTING DT');
  const acctCodeCol = headerRow.findIndex(h => String(h ?? '').toUpperCase() === 'ACCOUNT CODE');
  const acctNameCol = headerRow.findIndex(h => String(h ?? '').toUpperCase() === 'ACCOUNT NAME');
  const deptCol     = headerRow.findIndex(h => String(h ?? '').toUpperCase() === 'DEPT');
  const deptmtCol   = headerRow.findIndex(h => String(h ?? '').toUpperCase() === 'DEPARTMENT');
  const stateCol    = headerRow.findIndex(h => String(h ?? '').toUpperCase() === 'STATE');

  // Convert ACCOUNTING DT column to YYYY-MM-DD
  if (acctDtCol >= 0) {
    for (let i = dataStart; i < lastRow; i++) {
      const cellAddr = XLSX.utils.encode_cell({ r: i, c: acctDtCol });
      const cell = ws[cellAddr];
      if (!cell || cell.v == null) continue;
      let formatted = null;
      if (typeof cell.v === 'number') {
        const p = XLSX.SSF.parse_date_code(cell.v);
        if (p) formatted = `${p.y}-${String(p.m).padStart(2,'0')}-${String(p.d).padStart(2,'0')}`;
      } else if (typeof cell.v === 'string') {
        const m = cell.v.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        if (m) formatted = `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;
      }
      if (formatted) {
        cell.v = formatted; cell.t = 's'; cell.w = formatted; delete cell.z;
        if (data[i]) data[i][acctDtCol] = formatted;
      }
    }
  }

  // Account code remapping rules. First matching rule wins; dept sets must stay
  // disjoint — rules 1 and 3 are inverses of each other, so a dept listed in both
  // would remap differently depending on rule order.
  const REMAP_RULES = [
    {
      // These depts lack IT budget — expenses moved to Software licenses fee.
      depts:     new Set(['204000','204100','204101','204102','204200','204201','204202']),
      fromCode:  '52030100504', // Network IT expenses
      toCode:    '52030100506',
      toName:    'Software licenses fee',
    },
    {
      // These depts lack IT budget — expenses moved to Software licenses fee.
      depts:     new Set(['200020','200021']),
      fromCode:  '52030101107', // Computer system maintenance Contract
      toCode:    '52030100506',
      toName:    'Software licenses fee',
    },
    {
      // IT Security never carries license spend — a Software licenses fee entry
      // here is a miscode and belongs under Network IT expenses.
      depts:     new Set(['203304']), // IT Security
      fromCode:  '52030100506', // Software licenses fee
      toCode:    '52030100504',
      toName:    'Network IT expenses',
    },
  ];

  for (let i = dataStart; i < lastRow; i++) {
    const row      = data[i];
    const deptCode = String(row[deptCol] ?? '').trim().split(/\s+/)[0];
    const code     = String(row[acctCodeCol] ?? '').trim();
    const rule     = REMAP_RULES.find(r => r.depts.has(deptCode) && r.fromCode === code);
    if (rule) {
      data[i][acctCodeCol] = rule.toCode;
      data[i][acctNameCol] = rule.toName;
      const codeCell = XLSX.utils.encode_cell({ r: i, c: acctCodeCol });
      const nameCell = XLSX.utils.encode_cell({ r: i, c: acctNameCol });
      if (ws[codeCell]) { ws[codeCell].v = rule.toCode; ws[codeCell].w = rule.toCode; ws[codeCell].t = 's'; }
      if (ws[nameCell]) { ws[nameCell].v = rule.toName; ws[nameCell].w = rule.toName; ws[nameCell].t = 's'; }
    }
  }

  // Remap DEPT/DEPARTMENT for specific account codes
  const DEPT_REMAP_RULES = [
    {
      depts:    new Set(['206000']),      // Administration Office
      fromCode: '52030100601',           // Cleaning services expenses
      toDept:   '100000',
      toDeptmt: 'MD Office',
    },
  ];

  for (let i = dataStart; i < lastRow; i++) {
    const row      = data[i];
    const deptCode = String(row[deptCol] ?? '').trim().split(/\s+/)[0];
    const code     = String(row[acctCodeCol] ?? '').trim();
    const rule     = DEPT_REMAP_RULES.find(r => r.depts.has(deptCode) && r.fromCode === code);
    if (rule) {
      data[i][deptCol]   = rule.toDept;
      data[i][deptmtCol] = rule.toDeptmt;
      const dCell  = XLSX.utils.encode_cell({ r: i, c: deptCol });
      const dmCell = XLSX.utils.encode_cell({ r: i, c: deptmtCol });
      if (!ws[dCell])  ws[dCell]  = {};
      if (!ws[dmCell]) ws[dmCell] = {};
      ws[dCell].v  = rule.toDept;   ws[dCell].w  = rule.toDept;   ws[dCell].t  = 's';
      ws[dmCell].v = rule.toDeptmt; ws[dmCell].w = rule.toDeptmt; ws[dmCell].t = 's';
    }
  }

  // Fill void DEPT and DEPARTMENT with MD Office (100000)
  for (let i = dataStart; i < lastRow; i++) {
    if (String(data[i][deptCol] ?? '').trim()) continue;
    data[i][deptCol]   = '100000';
    data[i][deptmtCol] = 'MD Office';
    const dCell  = XLSX.utils.encode_cell({ r: i, c: deptCol });
    const dmCell = XLSX.utils.encode_cell({ r: i, c: deptmtCol });
    if (!ws[dCell])  ws[dCell]  = {};
    if (!ws[dmCell]) ws[dmCell] = {};
    ws[dCell].v  = '100000';   ws[dCell].w  = '100000';   ws[dCell].t  = 's';
    ws[dmCell].v = 'MD Office'; ws[dmCell].w = 'MD Office'; ws[dmCell].t = 's';
  }

  for (let i = dataStart; i < lastRow; i++) {
    const row = data[i];
    const accountCode = String(row[acctCodeCol] ?? '');
    const state       = String(row[stateCol] ?? '');
    const acctDt      = String(row[acctDtCol] ?? '');
    const passes = (accountCode.startsWith('4') || accountCode.startsWith('5'))
      && state === 'posted'
      && acctDt >= FROM_YYYYMMDD
      && acctDt <= TO_YYYYMMDD;
    ws['!rows'][i] = { ...(ws['!rows'][i] || {}), hidden: !passes };
  }

  // AutoFilter on header row (1-based Excel row)
  const lastCol = XLSX.utils.encode_col(headerRow.length - 1);
  ws['!autofilter'] = { ref: `A${headerRowIdx + 1}:${lastCol}${lastRow}` };

  XLSX.writeFile(wb, downloadPath);
  console.log('Filters applied via xlsx.');

  // Append filtered rows to Google Sheet
  const sheets = getSheetClient();

  const headerRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_TAB}!A1:Y1`,
  });
  const sheetHeaders = (headerRes.data.values?.[0] || []).map(h => String(h).trim());

  const xlsxHeaders = headerRow.map(h => String(h ?? '').trim());
  const colMap = {};
  xlsxHeaders.forEach((name, i) => {
    const si = sheetHeaders.findIndex(sh => sh.toLowerCase() === name.toLowerCase());
    if (si >= 0) colMap[i] = si;
  });
  console.log('Column mapping:', xlsxHeaders.map((n, i) => colMap[i] !== undefined ? `${n}→${sheetHeaders[colMap[i]]}` : null).filter(Boolean).join(', '));

  const numCols = sheetHeaders.length;
  const rowsToAppend = [];
  for (let i = dataStart; i < lastRow; i++) {
    const row = data[i];
    const accountCode = String(row[acctCodeCol] ?? '');
    const state       = String(row[stateCol] ?? '');
    const acctDt      = String(row[acctDtCol] ?? '');
    if (!((accountCode.startsWith('4') || accountCode.startsWith('5'))
      && state === 'posted'
      && acctDt >= FROM_YYYYMMDD
      && acctDt <= TO_YYYYMMDD)) continue;
    const sheetRow = new Array(numCols).fill('');
    for (const [xi, si] of Object.entries(colMap)) sheetRow[si] = row[xi] ?? '';
    rowsToAppend.push(sheetRow);
  }

  if (rowsToAppend.length === 0) {
    console.log('No rows matched filter — nothing appended to sheet.');
  } else {
    // Find first empty row by scanning column A only (ignores Z:AH formulas)
    const colARes = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_TAB}!A:A`,
    });
    const firstEmptyRow = (colARes.data.values || []).length + 1;
    const lastWriteRow  = firstEmptyRow + rowsToAppend.length - 1;
    console.log(`Writing ${rowsToAppend.length} rows into A${firstEmptyRow}:Y${lastWriteRow}`);

    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_TAB}!A${firstEmptyRow}:Y${lastWriteRow}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: rowsToAppend },
    });
    console.log(`Appended ${rowsToAppend.length} rows to "${SHEET_TAB}" tab.`);
  }

  console.log('Done. GL Ledger HMS report downloaded and filtered.');
  await browser.close();
})();
