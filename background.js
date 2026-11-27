// SPO Reading View - service worker (2.3.0)
// Layer 1: static declarativeNetRequest rules (Doc.aspx?action=default|edit) - no flicker.
// Layer 2: fallback for Doc.aspx URLs in unexpected formats, once per document per tab.
// Layer 3: share_resolver.js (content script) turns in-place sharing links into Doc.aspx URLs.
importScripts('urlutil.js');

const DEFAULTS = { force_view_default: true, force_view_edit: false, fallback: true, debug: false };
const WINDOW_MS = 30000;
const SP_URLS = ['https://*.sharepoint.com/*', 'https://*.sharepoint.cn/*'];
const MAIN = { urls: SP_URLS, types: ['main_frame'] };

let S = { ...DEFAULTS };
const ready = chrome.storage.local.get(DEFAULTS).then(v => { S = v; });
const log = (...a) => { if (S.debug) console.log('[SPO Reading View]', ...a); };

async function applyRulesets() {
  await ready;
  const ids = ['force_view_default', 'force_view_edit'];
  await chrome.declarativeNetRequest.updateEnabledRulesets({
    enableRulesetIds: ids.filter(i => S[i]),
    disableRulesetIds: ids.filter(i => !S[i])
  });
}
chrome.storage.onChanged.addListener((c, area) => {
  if (area !== 'local') return;
  for (const k of Object.keys(c)) S[k] = c[k].newValue;
  applyRulesets();
});
chrome.runtime.onInstalled.addListener(applyRulesets);

const shareTabs = new Map(), shareDocs = new Map(), shareReqs = new Set();
const recent = (m, k) => { const t = m.get(k); return t !== undefined && Date.now() - t < WINDOW_MS; };

chrome.webRequest.onBeforeRequest.addListener(d => {
  log('1 request', { tabId: d.tabId, url: brief(d.url) });
  if (!isSharingLink(d.url)) return;
  shareReqs.add(d.requestId);
  if (d.tabId >= 0) shareTabs.set(d.tabId, Date.now());
}, MAIN);
chrome.webRequest.onBeforeRedirect.addListener(d => {
  log('2 redirect', { tabId: d.tabId, status: d.statusCode, to: brief(d.redirectUrl) });
  if (!shareReqs.has(d.requestId) && !isSharingLink(d.url)) return;
  const c = classifyDoc(d.redirectUrl);
  if (c) shareDocs.set(c.docKey, Date.now());
}, MAIN);
const finish = d => { shareReqs.delete(d.requestId); };
chrome.webRequest.onCompleted.addListener(finish, MAIN);
chrome.webRequest.onErrorOccurred.addListener(finish, MAIN);

const key = tabId => 'tab_' + tabId;
async function getForced(tabId) { return (await chrome.storage.session.get(key(tabId)))[key(tabId)] || []; }
async function setForced(tabId, v) { await chrome.storage.session.set({ [key(tabId)]: v }); }
chrome.tabs.onRemoved.addListener(tabId => { shareTabs.delete(tabId); chrome.storage.session.remove(key(tabId)); });
chrome.tabs.onReplaced.addListener((added, removed) => {
  if (shareTabs.has(removed)) shareTabs.set(added, shareTabs.get(removed));
});

const inFlight = new Set();
async function check(tabId, url, reason, assumeShared = false) {
  const c = classifyDoc(url);
  if (!c) return;
  const fromSharingLink = assumeShared || recent(shareDocs, c.docKey) || recent(shareTabs, tabId);
  const flightKey = tabId + '|' + c.docKey;
  if (inFlight.has(flightKey)) return;
  inFlight.add(flightKey);
  try {
    await ready;
    log('5 doc url', { tabId, reason, action: c.action, isPpt: c.isPpt, fromSharingLink });
    if (!S.fallback || !S.force_view_default) return;
    const forced = await getForced(tabId);
    const ctx = { fromSharingLink, alreadyForced: forced.includes(c.docKey), overrideEdit: S.force_view_edit };
    if (!shouldForce(url, ctx)) { log('6 no change', ctx); return; }
    await setForced(tabId, [...forced, c.docKey].slice(-20));
    shareTabs.delete(tabId); shareDocs.delete(c.docKey);
    log('6 forcing Reading View', ctx);
    await chrome.tabs.update(tabId, { url: toViewUrl(url) });
  } finally { inFlight.delete(flightKey); }
}
chrome.tabs.onUpdated.addListener((tabId, info, tab) => {
  const url = info.url || (info.status === 'loading' ? tab.url : null);
  if (url) check(tabId, url, 'updated');
});
chrome.runtime.onStartup.addListener(async () => {
  applyRulesets();
  const tabs = await chrome.tabs.query({ url: SP_URLS });
  for (const t of tabs) if (t.url) check(t.id, t.url, 'startup', true);
});
