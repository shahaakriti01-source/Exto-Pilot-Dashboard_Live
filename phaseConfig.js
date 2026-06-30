// ─── PHASE CONFIGURATION ─────────────────────────────────────────────────────
// All incentive variables are configurable per phase and persist via localStorage
// (this is a standalone deployed app, not a Claude artifact, so we use the browser's
// own storage rather than the artifact-only window.storage API).

// ─── AGENTS (v2 — formerly hardcoded "field coordinators") ──────────────────
// This is now the SINGLE SOURCE OF TRUTH for who counts as an agent (i.e. who,
// when found in the Onboarding sheet's Authorize column, makes the Destination
// person a MERCHANT rather than a consumer of an existing merchant). Every tab
// in the app reads from this list — there is no hardcoded fallback list anymore.
//
// Pre-seeded with Exto's 2 permanent accounts + the 4 field coordinators known
// at the time v2 was built, so v2 starts out behaviorally identical to v1. New
// agents can be added freely from the Agents tab; Exto's 2 entries are
// permanent and cannot be deleted (only the human agents can be removed).
export const DEFAULT_AGENTS = [
  {
    id: "exto-pay",
    name: "Exto Backend ID - Pay",
    accountIds: ["rPezy2QDaNVDJ5LpABXpqAMiWPrricuKwW"],
    dateAdded: "2026-01-01",
    active: true,
    notes: "Exto's own backend account — sends activation/stipend payments. Permanent.",
    permanent: true,
  },
  {
    id: "exto-onboard",
    name: "Exto Backend ID - Onboard",
    accountIds: ["rHBFTaCH4Pcm3D5bSMH8nq6iibanr8vkKE"],
    dateAdded: "2026-01-01",
    active: true,
    notes: "Exto's own backend account — onboards merchants directly. Permanent.",
    permanent: true,
  },
  {
    id: "sandeep-taterway",
    name: "Sandeep Taterway",
    accountIds: ["rh3sNS8MRNn7qNJmJCAkFpKEUbFzdgLTiv", "rBwmS88Gv9J4G6ci7tbzjb8rT6Hs9Qntjt"],
    dateAdded: "2026-01-01",
    active: true,
    notes: "",
    permanent: false,
  },
  {
    id: "segomotso-mckenzie",
    name: "Segomotso Sadie M Mckenzie",
    accountIds: ["rUjEeeV3Wwf7eekei5AnLogpD21fPvTTBP", "rQgZnieFmztfdXpyTvDDcvPuVn73ef36e"],
    dateAdded: "2026-01-01",
    active: true,
    notes: "",
    permanent: false,
  },
  {
    id: "cindy-sibanda",
    name: "Cindy Sean Sibanda",
    accountIds: ["rKuaBPcpQ6Zv8ky5mjYGacssFcZoxeAUGB", "rn2Q3rj7urcCiq2gjB5FPtugEKDG92ce4G"],
    dateAdded: "2026-01-01",
    active: true,
    notes: "",
    permanent: false,
  },
  {
    id: "brian-moeng",
    name: "Brian Ketumile K Moeng",
    accountIds: [
      "rMCgd8TAdf6n3jBU2wziA2LVJHeHMVqhS7",
      "rE1G3LNPeG7FY1wekFJHNPSwNGnqeCVFdS",
      "rUzMmixdx22ycztzbSpHURUh3a6c5F8Akd",
      "r4DGpAZdfYpikFAfEtpmQPPfFEjfkTzrK8",
    ],
    dateAdded: "2026-01-01",
    active: true,
    notes: "",
    permanent: false,
  },
];

export const DEFAULT_INCENTIVE_CONFIG = {
  consumerRate: 13,
  consumerCap: 50,
  activationBonus: 95,
  activationThreshold: 2,
  fixedRate: 1,
  variableRate: 0.5, // percent
  variableCap: 27,
  txnCap: 150,
  milestones: [
    { count: 50, bonus: 100 },
    { count: 100, bonus: 125 },
    { count: 150, bonus: 150 },
  ],
};

// Consumer incentives are NOT currently paid out — every value defaults to 0.
// Same shape as the merchant config above, so the same calculation engine can
// be reused the moment real numbers are entered here.
export const DEFAULT_CONSUMER_INCENTIVE_CONFIG = {
  consumerRate: 0,
  consumerCap: 0,
  activationBonus: 0,
  activationThreshold: 0,
  fixedRate: 0,
  variableRate: 0,
  variableCap: 0,
  txnCap: 0,
  milestones: [],
};

// Dashboard-tab targets, from the Exto Pay / PulaConnect KPI deck.
// Independent of the Incentives-tab config above.
export const DEFAULT_DASHBOARD_TARGETS = {
  0: { merchants: 20, consumers: 200, p2mTransactions: 500, cashOutTransactions: 100 },
  1: { merchants: 30, consumers: 300, p2mTransactions: 800, cashOutTransactions: 100 },
  2: { merchants: 150, consumers: 5000, p2mTransactions: 7500, cashOutTransactions: 100 },
  3: { merchants: 1000, consumers: 50000, p2mTransactions: 50000, cashOutTransactions: 100 },
};

const STORAGE_KEY_INCENTIVE = "exto_incentive_config_by_phase";
const STORAGE_KEY_CONSUMER_INCENTIVE = "exto_consumer_incentive_config_by_phase";
const STORAGE_KEY_TARGETS = "exto_dashboard_targets_by_phase";
const STORAGE_KEY_PAID_TRIGGERS = "exto_paid_triggers";

function safeGet(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    if (raw) return JSON.parse(raw);
  } catch (e) {
    console.error(`Failed to read ${key} from localStorage:`, e);
  }
  return fallback;
}

function safeSet(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch (e) {
    console.error(`Failed to write ${key} to localStorage:`, e);
  }
}

const STORAGE_KEY_AGENTS = "exto_agents_list";

export function loadAgents() {
  const saved = safeGet(STORAGE_KEY_AGENTS, null);
  if (!saved || !Array.isArray(saved) || saved.length === 0) return [...DEFAULT_AGENTS];
  // Make sure Exto's 2 permanent entries always exist, even if a corrupted or
  // very old save somehow lost them.
  const ids = new Set(saved.map(a => a.id));
  const merged = [...saved];
  DEFAULT_AGENTS.filter(a => a.permanent).forEach(a => {
    if (!ids.has(a.id)) merged.unshift(a);
  });
  return merged;
}

export function saveAgents(agents) {
  safeSet(STORAGE_KEY_AGENTS, agents);
}

export function loadIncentiveConfigs() {
  return safeGet(STORAGE_KEY_INCENTIVE, {
    0: { ...DEFAULT_INCENTIVE_CONFIG },
    1: { ...DEFAULT_INCENTIVE_CONFIG },
    2: { ...DEFAULT_INCENTIVE_CONFIG },
    3: { ...DEFAULT_INCENTIVE_CONFIG },
  });
}

export function saveIncentiveConfigs(configs) {
  safeSet(STORAGE_KEY_INCENTIVE, configs);
}

export function loadConsumerIncentiveConfigs() {
  return safeGet(STORAGE_KEY_CONSUMER_INCENTIVE, {
    0: { ...DEFAULT_CONSUMER_INCENTIVE_CONFIG },
    1: { ...DEFAULT_CONSUMER_INCENTIVE_CONFIG },
    2: { ...DEFAULT_CONSUMER_INCENTIVE_CONFIG },
    3: { ...DEFAULT_CONSUMER_INCENTIVE_CONFIG },
  });
}

export function saveConsumerIncentiveConfigs(configs) {
  safeSet(STORAGE_KEY_CONSUMER_INCENTIVE, configs);
}

export function loadDashboardTargets() {
  const saved = safeGet(STORAGE_KEY_TARGETS, null);
  if (!saved) return { ...DEFAULT_DASHBOARD_TARGETS };
  // Merge in any new default fields (e.g. cashOutTransactions) that might be
  // missing from a config saved before this field existed, per phase.
  const merged = {};
  Object.keys(DEFAULT_DASHBOARD_TARGETS).forEach(phase => {
    merged[phase] = { ...DEFAULT_DASHBOARD_TARGETS[phase], ...(saved[phase] || {}) };
  });
  return merged;
}

export function saveDashboardTargets(targets) {
  safeSet(STORAGE_KEY_TARGETS, targets);
}

export function loadPaidTriggers() {
  const arr = safeGet(STORAGE_KEY_PAID_TRIGGERS, []);
  return new Set(arr);
}

export function savePaidTriggers(paidSet) {
  safeSet(STORAGE_KEY_PAID_TRIGGERS, [...paidSet]);
}

const STORAGE_KEY_PAID_BASELINES = "exto_paid_baselines";

// Tracks, per running-total trigger ID (e.g. "Oratile-onb"), the amount that
// was already paid as of the last "Mark as Paid" click. The trigger only ever
// shows the amount ABOVE this baseline as currently due — so onboarding more
// consumers (or doing more transactions) later never silently looks "already
// paid" just because the merchant's trigger ID stayed the same.
export function loadPaidBaselines() {
  return safeGet(STORAGE_KEY_PAID_BASELINES, {});
}

export function savePaidBaselines(baselines) {
  safeSet(STORAGE_KEY_PAID_BASELINES, baselines);
}
