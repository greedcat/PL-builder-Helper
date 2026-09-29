// ─────────────────────────────────────────────────────────────
// SHARED CLIPBOARD HELPER
// navigator.clipboard requires a secure context (https/localhost) — pages
// opened directly from disk (file://) silently fail, so fall back to the
// classic textarea + execCommand('copy') trick in that case.
// ─────────────────────────────────────────────────────────────
function copyToClipboard(text) {
  if (navigator.clipboard && navigator.clipboard.writeText && window.isSecureContext) {
    return navigator.clipboard.writeText(text).catch(() => legacyCopy(text));
  }
  return Promise.resolve().then(() => legacyCopy(text));
}

function legacyCopy(text) {
  const ta = document.createElement("textarea");
  ta.value = text;
  ta.style.position = "fixed";
  ta.style.opacity = "0";
  document.body.appendChild(ta);
  ta.focus();
  ta.select();
  try { document.execCommand("copy"); } catch (e) { /* ignore */ }
  document.body.removeChild(ta);
}
