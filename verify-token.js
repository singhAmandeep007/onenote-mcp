// End-to-end check of the stored token against Microsoft Graph.
// Works for BOTH JWT (work/school) and compact (personal account) tokens,
// because it asks Graph directly instead of guessing from the token shape.
// Run:  node verify-token.js
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const tokenFilePath = path.join(__dirname, '.access-token.txt');

if (!fs.existsSync(tokenFilePath)) {
  console.error('No .access-token.txt found. Run: npm run auth');
  process.exit(1);
}

const raw = fs.readFileSync(tokenFilePath, 'utf8');
let token;
try { token = JSON.parse(raw).token; } catch { token = raw.trim(); }
token = token.replace(/^bearer\s+/i, '').trim();

const kind = token.split('.').length === 3 ? 'JWT (work/school)' : 'compact (personal account)';
console.log('Token format:', kind);

async function hit(label, url) {
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  const body = await res.text();
  if (res.ok) {
    console.log(`✅ ${label}: ${res.status}`);
    return JSON.parse(body);
  }
  console.error(`❌ ${label}: ${res.status}`);
  console.error('   ' + body.slice(0, 300));
  return null;
}

(async () => {
  const me = await hit('/me', 'https://graph.microsoft.com/v1.0/me');
  if (me) console.log('   signed in as:', me.userPrincipalName || me.mail || me.displayName);

  const nb = await hit('/me/onenote/notebooks', 'https://graph.microsoft.com/v1.0/me/onenote/notebooks');
  if (nb && Array.isArray(nb.value)) {
    console.log(`   notebooks found: ${nb.value.length}`);
    nb.value.slice(0, 10).forEach(n => console.log('     •', n.displayName));
  }

  if (me && nb) {
    console.log('\nAll good — token works. Restart the Claude app and ask me to list your notebooks.');
  } else {
    console.error('\nToken was rejected. Delete .access-token.txt and re-run: npm run auth');
    process.exit(2);
  }
})();
