const DEFAULTS = { force_view_default: true, force_view_edit: false };

document.addEventListener('DOMContentLoaded', async () => {
  const s = await chrome.storage.local.get(DEFAULTS);
  for (const id of Object.keys(DEFAULTS)) {
    const box = document.getElementById(id);
    box.checked = !!s[id];
    box.addEventListener('change', () => chrome.storage.local.set({ [id]: box.checked }));
  }
  document.getElementById('ver').textContent = 'Version ' + chrome.runtime.getManifest().version;
});
