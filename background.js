// SPO Reading View - service worker
// Rulesets persist their enabled state in chrome.storage.local so the user's
// choice survives browser restarts and extension updates.

const DEFAULTS = { force_view_default: true, force_view_edit: false };

async function applySettings() {
  const s = await chrome.storage.local.get(DEFAULTS);
  const enable = [], disable = [];
  for (const id of Object.keys(DEFAULTS)) (s[id] ? enable : disable).push(id);
  await chrome.declarativeNetRequest.updateEnabledRulesets({
    enableRulesetIds: enable,
    disableRulesetIds: disable
  });
}

chrome.runtime.onInstalled.addListener(applySettings);
chrome.runtime.onStartup.addListener(applySettings);
chrome.storage.onChanged.addListener((changes, area) => {
  if (area === 'local') applySettings();
});
