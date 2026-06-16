const { google } = require('googleapis');
const path = require('path');
const fs = require('fs');

const SHEET_ID  = '157EM10aUaWiD6cWehjD20p1YEzsklNqG_IuF1XDI9YQ';
const SHEET_TAB = 'GL';
const CLEAR_RANGE = `${SHEET_TAB}!A2:Y`;

const envPath = path.join(__dirname, '.env');
const env = Object.fromEntries(
  fs.readFileSync(envPath, 'utf8')
    .split('\n')
    .filter(l => l.includes('='))
    .map(l => l.split('=').map(s => s.trim()))
);

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

(async () => {
  const sheets = await getSheetClient();
  await sheets.spreadsheets.values.clear({
    spreadsheetId: SHEET_ID,
    range: CLEAR_RANGE,
  });
  console.log(`Cleared ${CLEAR_RANGE} (header row kept).`);
})();
