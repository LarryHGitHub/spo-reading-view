// SPO Reading View - sharing-link resolver (content script, document_start)
//
// Sharing links such as https://tenant.sharepoint.com/:x:/s/site/<token> render the Office web
// editor in place and never navigate to Doc.aspx, so URL rewriting cannot apply. This script asks
// SharePoint (same origin, the user's own session) which file the link points to, then navigates
// to the same Doc.aspx URL that OneDrive uses:
//   {siteUrl}/_layouts/15/Doc.aspx?sourcedoc={GUID}&file={name}&action=default&mobileredirect=true&share={token}
// The extension's existing rule then turns action=default into Reading View.
// Nothing is sent anywhere except to the SharePoint host the user is already visiting.
(() => {
  const DEFAULTS = { force_view_default: true, fallback: true, debug: false };
  const RESOLVABLE = /^\/:([wx]):\/[a-z]\//i;          // Word / Excel only (PowerPoint ignores View)
  const OFFICE_EXT = /\.(?:docx?|docm|dotx?|rtf|odt|xlsx?|xlsm|xlsb|csv|ods)$/i;
  const GUARD = 'spoReadingView.resolved:' + location.pathname;

  const m = RESOLVABLE.exec(location.pathname);
  if (!m) return;

  const log = (s, ...a) => { if (s.debug) console.log('[SPO Reading View]', ...a); };

  // base64url "u!" encoding used by the /shares API
  function encodeShare(url) {
    const bytes = new TextEncoder().encode(url);
    let bin = ''; bytes.forEach(b => { bin += String.fromCharCode(b); });
    return 'u!' + btoa(bin).replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
  }

  async function lookup(shareUrl, s) {
    const id = encodeShare(shareUrl);
    const sel = '?$select=name,webUrl,sharepointIds';
    for (const ver of ['v2.0', 'v2.1']) {
      try {
        const r = await fetch(`${location.origin}/_api/${ver}/shares/${id}/driveItem${sel}`,
          { credentials: 'include', headers: { Accept: 'application/json' } });
        log(s, 'resolver', ver, 'HTTP', r.status);
        if (!r.ok) continue;
        const j = await r.json();
        const ids = j.sharepointIds || {};
        if (ids.listItemUniqueId && ids.siteUrl && j.name) return { name: j.name, guid: ids.listItemUniqueId, siteUrl: ids.siteUrl };
        log(s, 'resolver', ver, 'missing fields', Object.keys(j));
      } catch (e) { log(s, 'resolver', ver, 'error', String(e)); }
    }
    return null;
  }

  chrome.storage.local.get(DEFAULTS, async s => {
    if (!s.force_view_default || !s.fallback) return;
    // Do not loop: if we already redirected from this link in this tab (e.g. user pressed Back), stop.
    try { if (sessionStorage.getItem(GUARD)) { log(s, 'resolver: already handled, skip'); return; } } catch {}

    const shareUrl = location.origin + location.pathname;
    const token = location.pathname.split('/').filter(Boolean).pop();
    const info = await lookup(shareUrl, s);
    if (!info) { log(s, 'resolver: could not resolve, leaving page as is'); return; }
    if (!OFFICE_EXT.test(info.name)) { log(s, 'resolver: not a Word/Excel file, skip'); return; }

    const site = info.siteUrl.replace(/\/+$/, '');
    if (new URL(site).origin !== location.origin) { log(s, 'resolver: site on other host, skip'); return; }
    const target = `${site}/_layouts/15/Doc.aspx?sourcedoc=${encodeURIComponent('{' + info.guid.replace(/[{}]/g, '').toUpperCase() + '}')}`
      + `&file=${encodeURIComponent(info.name)}&action=default&mobileredirect=true&share=${encodeURIComponent(token)}`;
    try { sessionStorage.setItem(GUARD, '1'); } catch {}
    log(s, 'resolver: redirecting to Doc.aspx (action=default)');
    location.replace(target);
  });
})();
