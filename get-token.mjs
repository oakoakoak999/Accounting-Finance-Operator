import { createServer } from 'http';
import { google } from 'googleapis';
import { config } from 'dotenv';
import open from 'open';

config();

const oauth2Client = new google.auth.OAuth2(
  process.env.GMAIL_CLIENT_ID,
  process.env.GMAIL_CLIENT_SECRET,
  process.env.GMAIL_REDIRECT_URI
);

const authUrl = oauth2Client.generateAuthUrl({
  access_type: 'offline',
  prompt: 'consent',
  scope: ['https://www.googleapis.com/auth/gmail.readonly'],
});

console.log('Opening browser for authorization...');
open(authUrl);

const server = createServer(async (req, res) => {
  const url = new URL(req.url, 'http://localhost:3000');
  const code = url.searchParams.get('code');
  if (!code) { res.end('No code found.'); return; }

  try {
    const { tokens } = await oauth2Client.getToken(code);
    res.end('Authorization successful! You can close this tab.');
    server.close();
    console.log('\nYour refresh token:');
    console.log(tokens.refresh_token);
    console.log('\nPaste it into .env as GMAIL_REFRESH_TOKEN=');
  } catch (err) {
    res.end('Error: ' + err.message);
    server.close();
  }
});

server.listen(3000, () => console.log('Waiting for Google callback on http://localhost:3000 ...'));
