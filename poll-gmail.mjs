import { google } from 'googleapis';
import { config } from 'dotenv';

config();

const POLL_INTERVAL_MS = 2 * 60 * 1000; // every 2 minutes

const oauth2Client = new google.auth.OAuth2(
  process.env.GMAIL_CLIENT_ID,
  process.env.GMAIL_CLIENT_SECRET,
  process.env.GMAIL_REDIRECT_URI
);

oauth2Client.setCredentials({ refresh_token: process.env.GMAIL_REFRESH_TOKEN });

const gmail = google.gmail({ version: 'v1', auth: oauth2Client });

async function checkMail() {
  const query = `from:${process.env.WATCH_FROM} is:unread`;

  const res = await gmail.users.messages.list({ userId: 'me', q: query, maxResults: 10 });
  const messages = res.data.messages || [];

  if (messages.length === 0) {
    console.log(`[${now()}] No new mail from ${process.env.WATCH_FROM}`);
    return;
  }

  for (const msg of messages) {
    const detail = await gmail.users.messages.get({ userId: 'me', id: msg.id, format: 'full' });
    const headers = detail.data.payload.headers;
    const subject = headers.find(h => h.name === 'Subject')?.value ?? '(no subject)';
    const snippet = detail.data.snippet;

    console.log(`[${now()}] NEW MAIL from ${process.env.WATCH_FROM}`);
    console.log(`  Subject : ${subject}`);
    console.log(`  Snippet : ${snippet}`);

    // TODO: trigger your script/action here

    // Mark as read
    await gmail.users.messages.modify({
      userId: 'me',
      id: msg.id,
      requestBody: { removeLabelIds: ['UNREAD'] },
    });

    console.log(`  Marked as read.`);
  }
}

function now() {
  return new Date().toLocaleString('th-TH', { hour12: false });
}

console.log(`Polling Gmail every ${POLL_INTERVAL_MS / 60000} min for mail from ${process.env.WATCH_FROM}...`);
checkMail();
setInterval(checkMail, POLL_INTERVAL_MS);
