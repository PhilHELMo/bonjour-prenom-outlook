/* Bonjour Prénom — traitement local dans Outlook */
Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) return;
  const btn = document.getElementById("insert");
  btn.disabled = false;
  btn.addEventListener("click", insertGreeting);
});

function status(message) {
  document.getElementById("status").textContent = message;
}

function firstNameFromDisplayName(displayName, emailAddress) {
  let s = (displayName || "").trim();

  // Si Outlook ne fournit pas de nom exploitable, tentative prudente via l'adresse.
  if (!s || s.includes("@")) {
    const local = (emailAddress || "").split("@")[0] || "";
    s = local.replace(/[._-]+/g, " ").trim();
  }

  // "DUPONT, Jean" -> "Jean"
  if (s.includes(",")) {
    const parts = s.split(",").map(x => x.trim()).filter(Boolean);
    if (parts.length >= 2) s = parts[1];
  }

  // Retire quelques civilités fréquentes.
  s = s.replace(/^(m\.?|mr\.?|mme\.?|mlle\.?|dr\.?|pr\.?)\s+/i, "");

  const first = s.split(/\s+/).filter(Boolean)[0] || "";
  if (!first) return "";

  // Conserve les noms déjà correctement capitalisés; corrige surtout le tout-majuscules.
  if (first === first.toUpperCase() && /[A-ZÀ-ÖØ-Þ]/.test(first)) {
    return first.charAt(0).toUpperCase() + first.slice(1).toLocaleLowerCase("fr-FR");
  }
  return first;
}

function insertGreeting() {
  const item = Office.context.mailbox.item;
  status("Lecture du destinataire…");

  item.to.getAsync((r) => {
    if (r.status !== Office.AsyncResultStatus.Succeeded) {
      status("Impossible de lire le champ « À » : " + r.error.message);
      return;
    }
    if (!r.value || r.value.length === 0) {
      status("Ajoutez d’abord un destinataire dans le champ « À ».");
      return;
    }

    const recipient = r.value[0];
    const firstName = firstNameFromDisplayName(recipient.displayName, recipient.emailAddress);
    if (!firstName) {
      status("Le prénom n’a pas pu être déterminé automatiquement.");
      return;
    }

    item.body.getTypeAsync((typeResult) => {
      if (typeResult.status !== Office.AsyncResultStatus.Succeeded) {
        status("Impossible de déterminer le format du message.");
        return;
      }

      const isHtml = typeResult.value === Office.CoercionType.Html;
      const greeting = isHtml
        ? `<p>Bonjour ${escapeHtml(firstName)},</p>`
        : `Bonjour ${firstName},\r\n\r\n`;

      item.body.prependAsync(
        greeting,
        { coercionType: isHtml ? Office.CoercionType.Html : Office.CoercionType.Text },
        (writeResult) => {
          if (writeResult.status === Office.AsyncResultStatus.Succeeded) {
            status(`Salutation insérée : Bonjour ${firstName},`);
          } else {
            status("Insertion impossible : " + writeResult.error.message);
          }
        }
      );
    });
  });
}

function escapeHtml(s) {
  return s.replace(/[&<>"']/g, c => ({
    "&":"&amp;", "<":"&lt;", ">":"&gt;", '"':"&quot;", "'":"&#39;"
  })[c]);
}
