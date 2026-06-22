import { msalInstance, loginRequest } from "./authConfig";

// Your file's sharing URL (from SharePoint/OneDrive)
const SHARE_URL = "https://mergencompass-my.sharepoint.com/:x:/p/sandeep/IQBxf7rS-YmgRJG5clEQAPwgAcK13A5u8XVoU5QTRZB796c";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

// Encode a sharing URL into the format Graph API's /shares endpoint expects
function encodeShareUrl(url) {
  const base64 = btoa(unescape(encodeURIComponent(url)))
    .replace(/=/g, "")
    .replace(/\//g, "_")
    .replace(/\+/g, "-");
  return "u!" + base64;
}

async function getAccessToken() {
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length === 0) throw new Error("No signed-in account found");

  try {
    const response = await msalInstance.acquireTokenSilent({
      ...loginRequest,
      account: accounts[0],
    });
    return response.accessToken;
  } catch (err) {
    // Silent token failed (expired session etc.) — fall back to popup
    const response = await msalInstance.acquireTokenPopup(loginRequest);
    return response.accessToken;
  }
}

async function graphFetch(path, token) {
  const res = await fetch(`${GRAPH_BASE}${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Graph API error ${res.status}: ${body}`);
  }
  return res.json();
}

// Resolve the sharing link to get the actual driveId + itemId
async function resolveShareLink(token) {
  const encoded = encodeShareUrl(SHARE_URL);
  const data = await graphFetch(`/shares/${encoded}/driveItem`, token);
  return { driveId: data.parentReference.driveId, itemId: data.id };
}

// Read a worksheet's used range as values (handles formulas already evaluated by Excel Online)
async function getSheetValues(driveId, itemId, sheetName, token) {
  const path = `/drives/${driveId}/items/${itemId}/workbook/worksheets('${encodeURIComponent(sheetName)}')/usedRange(valuesOnly=true)`;
  const data = await graphFetch(path, token);
  return data.values || [];
}

// ── MAIN EXPORT: fetch and parse all sheets we need ──────────────────────
export async function fetchLivePilotData() {
  const token = await getAccessToken();
  const { driveId, itemId } = await resolveShareLink(token);

  const [accMapRows, onboardingDDRows, paymentsRows] = await Promise.all([
    getSheetValues(driveId, itemId, "Person name mapped to Acc ID", token),
    getSheetValues(driveId, itemId, "Accounts Trans - Onboarding(DD)", token),
    getSheetValues(driveId, itemId, "Payments Transaction", token),
  ]);

  // Parse Person name mapped to Acc ID (header rows assumed in first 4 rows per original file)
  const accToName = {};
  accMapRows.slice(4).forEach((row) => {
    if (row[1] && row[2]) accToName[String(row[1]).trim()] = String(row[2]).trim();
  });

  // Parse Onboarding (DD) — columns: A txn_id, ... E destination, F authorize, ... I created_at, K user_name, L agent_name
  const onboarding = [];
  onboardingDDRows.slice(4).forEach((row) => {
    if (!row[0] || row[0] === "Transaction ID") return;
    const agent = row[11] ? String(row[11]).trim() : null;
    const user = row[10] ? String(row[10]).trim() : null;
    onboarding.push({
      txn_id: String(row[0]).slice(0, 16),
      authorize: row[5] ? String(row[5]) : "",
      destination: row[4] ? String(row[4]) : "",
      created_at: row[8] ? String(row[8]) : "",
      user_name: user && user !== "#N/A" ? user : "Unknown",
      agent_name: agent && agent !== "#N/A" ? agent : null,
    });
  });

  // Parse Payments — columns: B txn_id, D account(sender id), E destination(receiver id), K amount, I created_at, L sender name, M receiver name, (col for included flag if present)
  const payments = [];
  paymentsRows.slice(3).forEach((row) => {
    if (!row[1] || row[1] === "Transaction ID") return;
    const amtRaw = row[10] ? String(row[10]).trim() : "0";
    const amt = parseFloat(amtRaw.replace("BWP", "").trim()) || 0;
    const sender = row[11] ? String(row[11]).trim() : "";
    const receiver = row[12] ? String(row[12]).trim() : "";
    // Column index 13 (N) reserved for inclusion flag if/when added to the sheet
    const inclusionFlag = row[13] ? String(row[13]).trim().toLowerCase() : null;
    const included = inclusionFlag ? inclusionFlag === "included" : true;

    payments.push({
      txn_id: String(row[1]).slice(0, 16),
      account: row[3] ? String(row[3]) : "",
      destination: row[4] ? String(row[4]) : "",
      amount: Math.round(amt * 100) / 100,
      created_at: row[8] ? String(row[8]) : "",
      sender: sender && sender !== "#N/A" ? sender : "",
      receiver: receiver && receiver !== "#N/A" ? receiver : "",
      included,
    });
  });

  return { accToName, onboarding, payments, fetchedAt: new Date().toISOString() };
}
