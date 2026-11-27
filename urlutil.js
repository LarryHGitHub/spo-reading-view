// Pure URL helpers (no chrome.* APIs) - shared by background.js and unit tests.
const DOC_PATH_RE = /\/_layouts\/15\/(?:doc2?|wopiframe2?)\.aspx$/i;
const SHARE_PATH_RE = /^\/:[a-z]:\/[a-z]\//i;
const PPT_EXT_RE = /\.(?:pptx?|pptm|ppsx?|potx?|odp)$/i;
const SP_HOST_RE = /\.sharepoint\.(?:com|cn)$/i;

function parse(u) { try { return new URL(u); } catch { return null; } }
function isSpHost(url) { return !!url && url.protocol === 'https:' && SP_HOST_RE.test(url.hostname); }
function getParamCI(url, name) {
  name = name.toLowerCase();
  for (const [k, v] of url.searchParams) if (k.toLowerCase() === name) return v;
  return null;
}
function isSharingLink(u) {
  const url = parse(u);
  return !!(isSpHost(url) && SHARE_PATH_RE.test(url.pathname) && !DOC_PATH_RE.test(url.pathname));
}
function classifyDoc(u) {
  const url = parse(u);
  if (!isSpHost(url) || !DOC_PATH_RE.test(url.pathname)) return null;
  const action = (getParamCI(url, 'action') || 'missing').toLowerCase();
  const file = getParamCI(url, 'file') || '';
  const docKey = (url.hostname + '|' + (getParamCI(url, 'sourcedoc') || file || url.pathname)).toLowerCase();
  return { action, docKey, isPpt: PPT_EXT_RE.test(file) };
}
function toViewUrl(u) {
  const url = parse(u);
  const drop = new Set(['action', 'mobileredirect', 'defaultitemopen']);
  for (const k of [...url.searchParams.keys()]) if (drop.has(k.toLowerCase())) url.searchParams.delete(k);
  url.searchParams.append('DefaultItemOpen', '1');
  url.searchParams.append('Action', 'View');
  return url.toString();
}
function shouldForce(u, ctx) {
  const c = classifyDoc(u);
  if (!c || c.isPpt || ctx.alreadyForced) return false;
  if (c.action === 'view' || c.action === 'embedview') return false;
  if (ctx.fromSharingLink) return true;
  if (c.action === 'default' || c.action === 'missing') return true;
  if (c.action === 'edit') return !!ctx.overrideEdit;
  return false;
}
function brief(u) {
  const url = parse(u);
  if (!url) return String(u).slice(0, 40);
  const kind = isSharingLink(u) ? 'SHARE ' + url.pathname.slice(0, 7)
             : DOC_PATH_RE.test(url.pathname) ? 'DOC ' + url.pathname.split('/').pop()
             : 'OTHER ' + url.pathname.split('/').slice(-1)[0];
  const names = [...url.searchParams.keys()].join(',');
  const act = getParamCI(url, 'action');
  return `${kind} | params=[${names}]` + (act !== null ? ` | action=${act}` : '');
}
if (typeof module !== 'undefined') module.exports = { isSharingLink, classifyDoc, toViewUrl, shouldForce, brief };
