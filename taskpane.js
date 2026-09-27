/* Bonjour Prénom — panneau et insertion manuelle */
const CLOSING_KEY = "bonjourPrenomClosing";
Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) return;
  const btn = document.getElementById("insert");
  const select = document.getElementById("closing");
  const settings = Office.context.roamingSettings;
  select.value = settings.get(CLOSING_KEY) || "Bien à toi";
  btn.disabled = false;
  select.addEventListener("change", () => {
    settings.set(CLOSING_KEY, select.value);
    settings.saveAsync((r) => status(r.status === Office.AsyncResultStatus.Succeeded
      ? `Formule enregistrée : ${select.value}` : "Impossible d’enregistrer le choix."));
  });
  btn.addEventListener("click", () => insertEnvelope(select.value));
  Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => {
    const item = Office.context.mailbox.item;
    btn.disabled = !item || !item.body || !item.to;
    status(item ? "" : "Ouvrez un brouillon pour insérer la formule.");
  });
  btn.disabled = !Office.context.mailbox.item;
});
function status(message) { document.getElementById("status").textContent = message; }
function insertEnvelope(closing) {
  const item = Office.context.mailbox.item;
  if (!item || !item.to || !item.body) { status("Ouvrez un brouillon pour insérer la formule."); return; }
  status("Lecture du destinataire…");
  item.to.getAsync((r) => {
    if (r.status !== Office.AsyncResultStatus.Succeeded) { status("Impossible de lire le champ « À » : " + r.error.message); return; }
    if (!r.value || r.value.length === 0) { status("Ajoutez d’abord un destinataire dans le champ « À »."); return; }
    const recipient = r.value[0];
    const firstName = firstNameFromDisplayName(recipient.displayName, recipient.emailAddress);
    if (!firstName) { status("Le prénom n’a pas pu être déterminé automatiquement."); return; }
    item.body.getTypeAsync((tr) => {
      if (tr.status !== Office.AsyncResultStatus.Succeeded) { status("Impossible de déterminer le format du message."); return; }
      const isHtml = tr.value === Office.CoercionType.Html;
      const content = isHtml
        ? `<div data-bonjour-prenom="1"><p>Bonjour ${escapeHtml(firstName)},</p><p><br></p><p><br></p><p>${escapeHtml(closing)}</p><p><br></p></div>`
        : `Bonjour ${firstName},\r\n\r\n\r\n\r\n${closing}\r\n\r\n`;
      item.body.prependAsync(content, { coercionType: isHtml ? Office.CoercionType.Html : Office.CoercionType.Text },
        (wr) => status(wr.status === Office.AsyncResultStatus.Succeeded ? `Salutation et « ${closing} » insérés.` : "Insertion impossible : " + wr.error.message));
    });
  });
}

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
