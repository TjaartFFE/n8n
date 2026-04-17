#!/usr/bin/env node
// FFE Pakplan Generator — standalone orchestrator
//
// Downloads Pakvolumes + Markpryse from SharePoint, generates per-producer
// Excel pack-plan files, and emails results via Microsoft Graph.
//
// Authentication: Azure AD client credentials (no user login required).
// The registered app needs (Microsoft Graph Application permissions, admin consent):
//   • Sites.Read.All  (or Sites.ReadWrite.All) — to download SharePoint files via Graph
//   • Mail.Send                                — to send email
//
// Usage:
//   node scripts/pakplan/run.mjs
//   PAKPLAN_DRY_RUN=1 node scripts/pakplan/run.mjs   # skip email
//
// Required env vars: AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET
// Optional env vars: PAKPLAN_SEND_AS (default: tjaart@ffesa.co.za), PAKPLAN_DRY_RUN

import https from 'https';
import { generatePackPlans } from './generator.mjs';

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------

// Graph API file download — uses Sites.Read.All (Microsoft Graph), no SharePoint-specific token needed
const GRAPH_SITE = 'ffesa.sharepoint.com:/sites/FFEPublicData:';
const GRAPH_FOLDER = 'Shared Documents/FFE Bemarking/LIVE';

const PAKVOLUMES_URL =
  `https://graph.microsoft.com/v1.0/sites/${GRAPH_SITE}/drive/root:/${GRAPH_FOLDER}/Pakvolumes LIVE 2026.xlsm:/content`;

const MARKPRYSE_URL =
  `https://graph.microsoft.com/v1.0/sites/${GRAPH_SITE}/drive/root:/${GRAPH_FOLDER}/Markpryse LIVE 2026.xlsm:/content`;

const REQUIRED_VARS = ['AZURE_TENANT_ID', 'AZURE_CLIENT_ID', 'AZURE_CLIENT_SECRET'];

// ---------------------------------------------------------------------------
// HTTP helpers (zero npm dependencies — built-in https only)
// ---------------------------------------------------------------------------

/**
 * HTTPS POST — returns response body as string.
 * @param {string} urlStr
 * @param {Record<string,string>} headers
 * @param {string|Buffer} body
 * @returns {Promise<string>}
 */
function httpsPost(urlStr, headers, body) {
  return new Promise((resolve, reject) => {
    const u = new URL(urlStr);
    const bodyBuf = typeof body === 'string' ? Buffer.from(body, 'utf8') : body;
    const options = {
      hostname: u.hostname,
      path: u.pathname + u.search,
      method: 'POST',
      headers: { 'Content-Length': bodyBuf.length, ...headers },
    };
    const req = https.request(options, (res) => {
      const chunks = [];
      res.on('data', (c) => chunks.push(c));
      res.on('end', () => {
        const text = Buffer.concat(chunks).toString('utf8');
        if (res.statusCode >= 400) {
          reject(new Error(`HTTP ${res.statusCode} POST ${urlStr}\n${text.slice(0, 500)}`));
        } else {
          resolve(text);
        }
      });
    });
    req.on('error', reject);
    req.write(bodyBuf);
    req.end();
  });
}

/**
 * HTTPS GET — returns raw response as Buffer. Follows one redirect.
 * @param {string} urlStr
 * @param {Record<string,string>} headers
 * @returns {Promise<Buffer>}
 */
function httpsGet(urlStr, headers) {
  return new Promise((resolve, reject) => {
    https.get(urlStr, { headers }, (res) => {
      if (res.statusCode === 301 || res.statusCode === 302) {
        resolve(httpsGet(res.headers.location, headers));
        res.resume();
        return;
      }
      const chunks = [];
      res.on('data', (c) => chunks.push(c));
      res.on('end', () => {
        const buf = Buffer.concat(chunks);
        if (res.statusCode >= 400) {
          reject(new Error(`HTTP ${res.statusCode} GET ${urlStr}\n${buf.toString('utf8').slice(0, 500)}`));
        } else {
          resolve(buf);
        }
      });
      res.on('error', reject);
    }).on('error', reject);
  });
}

// ---------------------------------------------------------------------------
// Auth
// ---------------------------------------------------------------------------

/**
 * Fetch an OAuth2 client-credentials access token from Azure AD.
 * @param {string} tenantId
 * @param {string} clientId
 * @param {string} clientSecret
 * @param {string} scope  e.g. 'https://graph.microsoft.com/.default'
 * @returns {Promise<string>} access_token
 */
async function fetchToken(tenantId, clientId, clientSecret, scope) {
  const body = new URLSearchParams({
    grant_type: 'client_credentials',
    client_id: clientId,
    client_secret: clientSecret,
    scope,
  }).toString();

  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const response = await httpsPost(
    tokenUrl,
    { 'Content-Type': 'application/x-www-form-urlencoded' },
    body,
  );

  const data = JSON.parse(response);
  if (data.error) {
    throw new Error(`Token fetch failed (${data.error}): ${data.error_description}`);
  }
  return data.access_token;
}

// ---------------------------------------------------------------------------
// Email payload builder (mirrors Code: Build Email node exactly)
// ---------------------------------------------------------------------------

/**
 * @param {Array} outputItems - from generatePackPlans()
 * @param {string} sendAs - mailbox to send from/to
 * @returns {object} Graph API sendMail payload
 */
function buildEmailPayload(outputItems, sendAs) {
  const files    = outputItems.filter((i) => i.json.fileB64 && i.json.fileName);
  const diagItem = outputItems.find((i) => i.json.producerName === '__DIAG__');

  const exportWeek = (files[0] || outputItems[0]).json.exportWeek || 'W??';
  const today      = new Date().toLocaleDateString('af-ZA');

  const attachments = files.map((f) => ({
    '@odata.type': '#microsoft.graph.fileAttachment',
    name: f.json.fileName,
    contentBytes: f.json.fileB64,
    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  }));

  const producerList = files.map((f) => `  - ${f.json.producerName} (${f.json.variety})`).join('\n');
  const diagText     = diagItem ? '\n\n[DIAG]\n' + diagItem.json.diag : '';

  return {
    message: {
      subject: `FFE Pakplanne - ${exportWeek} (geskep op ${today})`,
      body: {
        contentType: 'Text',
        content:
          `Hallo Tjaart,\n\nAangeheg die pakplanne vir ${exportWeek}:\n` +
          producerList + diagText +
          '\n\nGroete,\nFFE Pakplan AI',
      },
      toRecipients: [{ emailAddress: { address: sendAs } }],
      attachments,
    },
    saveToSentItems: true,
  };
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function main() {
  // Validate required env vars
  const missing = REQUIRED_VARS.filter((v) => !process.env[v]);
  if (missing.length) {
    throw new Error(
      `Missing required environment variable(s): ${missing.join(', ')}\n` +
      'Copy .env.local.example → .env.local, fill in the values, then run:\n' +
      '  pnpm exec dotenvx run -f .env.local -- node scripts/pakplan/run.mjs',
    );
  }

  const {
    AZURE_TENANT_ID,
    AZURE_CLIENT_ID,
    AZURE_CLIENT_SECRET,
    PAKPLAN_SEND_AS = 'tjaart@ffesa.co.za',
    PAKPLAN_DRY_RUN,
  } = process.env;

  // 1. Fetch single Graph token (used for both file downloads and email)
  console.log('Fetching auth token...');
  const graphToken = await fetchToken(
    AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET,
    'https://graph.microsoft.com/.default',
  );
  console.log('  Token acquired');

  // 2. Download source Excel files via Graph API (requires Sites.Read.All on Microsoft Graph)
  console.log('Downloading source files from SharePoint...');
  const [pakvolBuf, markpryseBuf] = await Promise.all([
    httpsGet(PAKVOLUMES_URL, { Authorization: `Bearer ${graphToken}` }),
    httpsGet(MARKPRYSE_URL,  { Authorization: `Bearer ${graphToken}` }),
  ]);
  console.log(`  Pakvolumes : ${pakvolBuf.length.toLocaleString()} bytes`);
  console.log(`  Markpryse  : ${markpryseBuf.length.toLocaleString()} bytes`);

  // 3. Generate pack plans
  console.log('Generating pack plans...');
  const outputItems = generatePackPlans(
    pakvolBuf.toString('base64'),
    markpryseBuf.toString('base64'),
  );

  const files    = outputItems.filter((i) => i.json.fileB64 && i.json.fileName);
  const diagItem = outputItems.find((i) => i.json.producerName === '__DIAG__');

  console.log(`  Generated ${files.length} pack plan(s)`);
  if (files.length > 0) {
    for (const f of files) {
      console.log(`    - ${f.json.fileName}  (${f.json.producerName} / ${f.json.variety})`);
    }
  }
  if (diagItem) {
    console.log('\nDiagnostics:');
    console.log(diagItem.json.diag.split('\n').map((l) => '  ' + l).join('\n'));
  }

  // 4. Dry-run exit
  if (PAKPLAN_DRY_RUN) {
    console.log(`\n[DRY RUN] Email skipped. Would send to: ${PAKPLAN_SEND_AS}`);
    return;
  }

  if (files.length === 0) {
    console.warn('\nNo pack plans generated — email not sent.');
    return;
  }

  // 5. Build and send email
  const payload = buildEmailPayload(outputItems, PAKPLAN_SEND_AS);

  console.log(`\nSending email to ${PAKPLAN_SEND_AS}...`);
  // Use /users/{id}/sendMail — /me/sendMail requires delegated auth, not app-only
  await httpsPost(
    `https://graph.microsoft.com/v1.0/users/${PAKPLAN_SEND_AS}/sendMail`,
    {
      Authorization: `Bearer ${graphToken}`,
      'Content-Type': 'application/json',
    },
    JSON.stringify(payload),
  );
  console.log(`  Email sent: "${payload.message.subject}"`);
}

main().catch((err) => {
  console.error('\nError:', err.message);
  process.exit(1);
});
