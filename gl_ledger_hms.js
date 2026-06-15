const { chromium } = require('playwright');
const XLSX = require('xlsx');
const { google } = require('googleapis');
const path = require('path');
const fs = require('fs');

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

const _args = process.argv.slice(2);
if (_args.length < 3)
  throw new Error('Usage: node gl_ledger_hms.js <fromMonth> <toMonth> <year>\nExample: node gl_ledger_hms.js jan may 2026');

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
  const browser = await chromium.launch({ headless: false, channel: 'chrome' });
  const context = await browser.newContext();
  const page    = await context.newPage();

  // Login
  await page.goto(`${BASE_URL}/web/login`, { waitUntil: 'load' });
  await page.fill('#login', USERNAME);
  await page.fill('#password', PASSWORD);
  await page.click('button[type=submit]');
  await page.waitForLoadState('load');
  await page.waitForTimeout(2000);

  // Go directly to Query -> GL Ledger all Detail (HMS)
  await page.goto(`${BASE_URL}/web#menu_id=1079&action=1430&cids=2`, { waitUntil: 'load', timeout: 60000 });
  await page.waitForTimeout(3000);

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
  await pickDate('x_x_gl_ledger_all_descr_hms_884_amdate_0', fromParsed.month, fromParsed.year, fromParsed.day);

  // Set To date
  await pickDate('x_x_gl_ledger_all_descr_hms_885_amdate_0', toParsed.month, toParsed.year, toParsed.day);

  // Set BU -> HMS
  const buField = page.locator('#x_x_gl_ledger_all_descr_hms_886_rcid_0');
  await buField.click();
  await buField.fill('HMS');
  await page.waitForTimeout(1000);
  await page.click('.o-autocomplete--dropdown-item:first-child');
  await page.waitForTimeout(500);

  const today = new Date().toISOString().split('T')[0]; // YYYY-MM-DD
  const downloadPath = `C:\\Users\\Thanapol.ph\\Downloads\\GL Ledger all Detail (HMS) ${today}.xlsx`;

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

  // Find ACCOUNTING DT column index from header row
  const headerRow  = data[headerRowIdx];
  const acctDtCol  = headerRow.findIndex(h => String(h ?? '').toUpperCase() === 'ACCOUNTING DT');
  const acctCodeCol = headerRow.findIndex(h => String(h ?? '').toUpperCase() === 'ACCOUNT CODE');
  const stateCol   = headerRow.findIndex(h => String(h ?? '').toUpperCase() === 'STATE');

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
