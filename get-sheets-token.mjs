import { createServer } from 'http';
import { google } from 'googleapis';
import { readFileSync, writeFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { join, dirname } from 'path';
import open from 'open';

const __dir = dirname(fileURLToPath(import.meta.url));
const env = Object.fromEntries(
  readFileSync(join(__dir, '.env'), 'utf8')
    .split('\n').filter(l => l.includes('='))
    .map(l => l.split('=').map(s => s.trim()))
);

const auth = new google.auth.OAuth2(
  env.GMAIL_CLIENT_ID,
  env.GMAIL_CLIENT_SECRET,
  env.GMAIL_REDIRECT_URI
);

const authUrl = auth.generateAuthUrl({
  access_type: 'offline',
  prompt: 'consent',
  scope: ['https://www.googleapis.com/auth/spreadsheets'],
  login_hint: 'thanapol.ph@princgroup.com',
});

console.log('Opening browser for Sheets authorization...');
open(authUrl);

const server = createServer(async (req, res) => {
  const url = new URL(req.url, 'http://localhost:3000');
  const code = url.searchParams.get('code');
  if (!code) { res.end('No code found.'); return; }

  try {
    const { tokens } = await auth.getToken(code);
    const tokenPath = join(__dir, '.gsheets-token.json');
    writeFileSync(tokenPath, JSON.stringify(tokens, null, 2), 'utf8');
    res.end('Authorization successful! You can close this tab.');
    server.close();
    console.log('Sheets token saved → ' + tokenPath);
  } catch (err) {
    res.end('Error: ' + err.message);
    server.close();
  }
});

server.listen(3000, () => console.log('Waiting for Google callback on http://localhost:3000 ...'));
