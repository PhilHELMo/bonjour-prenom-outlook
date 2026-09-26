/* Bonjour Prénom — activation automatique Outlook */
const CLOSING_KEY = "bonjourPrenomClosing";
const HTML_MARKER = 'data-bonjour-prenom="1"';
Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
Office.actions.associate("onMessageRecipientsChangedHandler", onMessageRecipientsChangedHandler);
function onNewMessageComposeHandler(event) { tryAutoInsert(event); }
function onMessageRecipientsChangedHandler(event) { tryAutoInsert(event); }
function tryAutoInsert(event) {
  const item = Office.context.mailbox.item;
  const closing = Office.context.roamingSettings.get(CLOSING_KEY) || "Bien à toi";
  item.to.getAsync((toResult) => {
    if (toResult.status !== Office.AsyncResultStatus.Succeeded || !toResult.value || toResult.value.length === 0) { event.completed(); return; }
    const recipient = toResult.value[0];
    const firstName = firstNameFromDisplayName(recipient.displayName, recipient.emailAddress);
    if (!firstName) { event.completed(); return; }
    item.body.getTypeAsync((tr) => {
      if (tr.status !== Office.AsyncResultStatus.Succeeded) { event.completed(); return; }
      const isHtml = tr.value === Office.CoercionType.Html;
      const coercionType = isHtml ? Office.CoercionType.Html : Office.CoercionType.Text;
      item.body.getAsync(coercionType, (br) => {
        if (br.status !== Office.AsyncResultStatus.Succeeded) { event.completed(); return; }
        const body = br.value || "";
        if ((isHtml && body.includes(HTML_MARKER)) ||
            (stripHtml(body).replace(/\s+/g," ").toLocaleLowerCase("fr-FR").includes(`bonjour ${firstName},`.toLocaleLowerCase("fr-FR")) &&
             stripHtml(body).replace(/\s+/g," ").toLocaleLowerCase("fr-FR").includes(closing.toLocaleLowerCase("fr-FR")))) {
          event.completed(); return;
        }
        const content = isHtml
          ? `<div data-bonjour-prenom="1"><p>Bonjour ${escapeHtml(firstName)},</p><p><br></p><p><br></p><p>${escapeHtml(closing)}</p><p><br></p></div>`
          : `Bonjour ${firstName},\r\n\r\n\r\n\r\n${closing}\r\n\r\n`;
        item.body.prependAsync(content, { coercionType }, () => event.completed());
      });
    });
  });
}
function stripHtml(s) { return String(s).replace(/<[^>]*>/g, " "); }

function firstNameFromDisplayName(displayName, emailAddress) {
  let s = (displayName || "").trim();
  if (!s || s.includes("@")) {
    const local = (emailAddress || "").split("@")[0] || "";
    s = local.replace(/[._-]+/g, " ").trim();
  }
  if (s.includes(",")) {
    const parts = s.split(",").map(x => x.trim()).filter(Boolean);
    if (parts.length >= 2) s = parts[1];
  }
  s = s.replace(/^(m\.?|mr\.?|mme\.?|mlle\.?|dr\.?|pr\.?)\s+/i, "");
  const first = s.split(/\s+/).filter(Boolean)[0] || "";
  if (!first) return "";
  if (first === first.toUpperCase() && /[A-ZÀ-ÖØ-Þ]/.test(first)) {
    return first.charAt(0).toUpperCase() + first.slice(1).toLocaleLowerCase("fr-FR");
  }
  return first;
}
function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
}
