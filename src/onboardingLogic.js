// ─── ONBOARDING IDENTITY LOGIC (v2.1 — dash-row based merchant grant) ───────
//
// This replaces the v2 "agent-onboarded = merchant" rule entirely. Agents can
// now onboard BOTH merchants and consumers, so simply checking who's in the
// Authorize column is no longer sufficient. The new rule:
//
//   0. Anyone who is currently an AGENT (per the Agents tab) is excluded from
//      merchant/consumer classification entirely — agents are neither. (The
//      Agents list now exists purely for the Internal Transfers tab.)
//   1. RULE A: If Authorize = Exto Backend ID directly → Destination is a
//      MERCHANT. (Unchanged from v1/v2.)
//   2. RULE B: For anyone onboarded by anyone OTHER than Exto directly, check
//      whether that PERSON (by name, across all their account IDs — a person
//      can have multiple IDs from card + phone wallet activations) has ANY
//      row where they are the Destination with BOTH Owner and Authorize
//      blank/dash. That dash-row represents a separate "merchant status
//      granted" event. If such a row exists → MERCHANT. Otherwise → CONSUMER.
//   3. This is fully dynamic against the live sheet — if a dash-row is added
//      later (e.g. merchant status granted 15 days after initial onboarding),
//      the very next data refresh reclassifies that person as a merchant
//      automatically, with no special-casing needed.

const EXTO_BACKEND_ONBOARD_ID = "rHBFTaCH4Pcm3D5bSMH8nq6iibanr8vkKE";
const EXTO_BACKEND_PAY_ID = "rPezy2QDaNVDJ5LpABXpqAMiWPrricuKwW";
const EXTO_IDS = new Set([EXTO_BACKEND_ONBOARD_ID, EXTO_BACKEND_PAY_ID]);

function normalizeName(name) {
  return (name || "").trim().toUpperCase().replace(/\s+/g, " ");
}

function isDash(val) {
  const s = (val || "").trim();
  return s === "" || s === "—" || s === "-" || s === "0";
}

/**
 * Builds fast lookup structures from the current agents list. Agents are now
 * used ONLY to exclude people from merchant/consumer classification (they're
 * neither) and to drive the Internal Transfers tab — they no longer determine
 * who counts as a merchant. Deliberately includes inactive agents, since an
 * agent's historical activity should always be excluded the same way.
 */
export function buildAgentLookup(agents) {
  const idSet = new Set();
  const nameSet = new Set();
  (agents || []).forEach(agent => {
    (agent.accountIds || []).forEach(id => idSet.add(id));
    nameSet.add(normalizeName(agent.name));
  });
  return { idSet, nameSet };
}

export function isAgentName(name, agentLookup) {
  return agentLookup.nameSet.has(normalizeName(name));
}

export function isAgentId(id, agentLookup) {
  return agentLookup.idSet.has(id);
}

/**
 * Scans ALL raw onboarding rows (before any dedup/filtering) to find every
 * person who has at least one "dash-row" — a row where they're the
 * Destination with both Owner and Authorize blank — which is the signal that
 * merchant status was explicitly granted to them. Keyed by normalized name,
 * since one person's multiple account IDs should all benefit from a dash-row
 * found on any one of them.
 */
function buildDashRowGrantSet(rawRows) {
  const granted = new Set();
  rawRows.forEach(row => {
    if (isDash(row.owner) && isDash(row.authorize) && row.user_name) {
      granted.add(normalizeName(row.user_name));
    }
  });
  return granted;
}

/**
 * Given raw onboarding rows (from Accounts Trans - Onboarding(DD), already
 * filtered to Submission === "tesSUCCESS"), dedupes by unique Owner ID and
 * classifies each unique onboarding event as a merchant or consumer onboarding
 * per the v2.1 rule described above.
 *
 * Row shape expected: { owner, destination, authorize, dlt_close, user_name, agent_name }
 *
 * `agentLookup` comes from buildAgentLookup(agents) — pass the CURRENT agents
 * list each time data is recalculated. Agents are excluded from the resulting
 * events entirely (neither merchant nor consumer), since they're tracked
 * separately via the Internal Transfers tab.
 */
export function dedupeOnboardingEvents(rawRows, agentLookup) {
  // Dash-row grants must be computed from the FULL raw row set (including
  // dash rows themselves, and including rows for people who might be filtered
  // out elsewhere), so this runs before any deduplication.
  const dashRowGrants = buildDashRowGrantSet(rawRows);

  const seenOwners = new Set();
  const events = [];

  rawRows.forEach(row => {
    // Dash rows themselves (Owner AND Authorize both blank) are a status
    // grant, not a real onboarding event — they don't represent one person
    // onboarding another, so they never become an event in their own right.
    if (isDash(row.owner) && isDash(row.authorize)) return;

    const ownerKey = row.owner || row.destination; // fallback if owner missing
    if (seenOwners.has(ownerKey)) return;
    seenOwners.add(ownerKey);

    // Step 0: if the person being onboarded (Destination) is themselves an
    // agent, exclude this row entirely — agents are neither merchants nor
    // consumers. (Whether the ONBOARDER happens to be an agent is irrelevant
    // here — agents onboarding people is completely normal; Rule A/B below
    // determine the outcome for the person being onboarded.)
    if (isAgentId(row.destination, agentLookup) || isAgentName(row.user_name, agentLookup)) return;

    // Rule A: direct Exto onboarding always makes the Destination a merchant.
    const onboardedDirectlyByExto = EXTO_IDS.has((row.authorize || "").trim());

    // Rule B: anyone else needs a dash-row grant (checked by name, across all
    // of that person's account IDs) to count as a merchant.
    const hasDashRowGrant = dashRowGrants.has(normalizeName(row.user_name));

    const isMerchantOnboarding = onboardedDirectlyByExto || hasDashRowGrant;

    events.push({
      owner: ownerKey,
      destination: row.destination,
      authorize: row.authorize,
      dlt_close: row.dlt_close,
      user_name: row.user_name,
      // agent_name here means "who onboarded this consumer" for display
      // purposes on the Consumers/Merchants tabs — only meaningful for
      // consumer events, so it's left null for merchant events.
      agent_name: isMerchantOnboarding ? null : row.agent_name,
      isMerchantOnboarding,
    });
  });

  return events;
}

/**
 * All unique REAL merchants — determined ENTIRELY from onboarding events
 * (never from payments). Agents are already excluded upstream in
 * dedupeOnboardingEvents, so nothing further to filter out here.
 */
export function getAllMerchantNames(onboardingEvents) {
  const names = new Set();
  onboardingEvents.forEach(e => {
    if (e.isMerchantOnboarding) names.add(e.user_name);
  });
  return [...names].sort();
}

/** All consumer names — anyone onboarded who is NOT a merchant (and not an agent,
 *  already excluded upstream). */
export function getAllConsumerNames(onboardingEvents) {
  const names = new Set();
  onboardingEvents.forEach(e => {
    if (!e.isMerchantOnboarding) names.add(e.user_name);
  });
  return names;
}

/** Every consumer, paired with the name of whoever onboarded them (deduped by
 *  Owner). Used by the Consumers tab for a flat, unambiguous list. */
export function getAllConsumersWithMerchant(onboardingEvents) {
  return onboardingEvents
    .filter(e => !e.isMerchantOnboarding)
    .map(e => ({ name: e.user_name, onboardedBy: e.agent_name || "Unknown", dlt_close: e.dlt_close, owner: e.owner }));
}

/** Consumers onboarded by a specific merchant (Owner-deduped). Note: under the
 *  new rule a merchant's onboarding chain isn't guaranteed to flow through
 *  Authorize == that merchant's ID in every case (Rule B can grant merchant
 *  status to someone onboarded by an agent or another merchant), so this
 *  matches on the resolved agent_name exactly as recorded on the event. */
export function getConsumersForMerchant(merchantName, onboardingEvents) {
  return onboardingEvents.filter(e => !e.isMerchantOnboarding && e.agent_name === merchantName);
}

/** Whether a merchant was successfully onboarded at all. */
export function isMerchantOnboarded(merchantName, onboardingEvents) {
  return onboardingEvents.some(e => e.isMerchantOnboarding && e.user_name === merchantName);
}
