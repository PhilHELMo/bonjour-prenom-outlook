/* Bonjour Prénom v4 — réponses + nouveaux messages */
const CLOSING_KEY = "bonjourPrenomClosing";
const HTML_MARKER = 'data-bonjour-prenom="1"';

function onNewMessageComposeHandler(event) {
  tryAutoInsert(event);
}

function onMessageRecipientsChangedHandler(event) {
  tryAutoInsert(event);
}

function tryAutoInsert(event) {
  const item = Office.context.mailbox.item;
  const closing = Office.context.roamingSettings.get(CLOSING_KEY) || "Bien à toi";

  item.to.getAsync({ asyncContext: event }, function (toResult) {
    const evt = toResult.asyncContext;

    if (toResult.status !== Office.AsyncResultStatus.Succeeded ||
        !toResult.value || toResult.value.length === 0) {
      evt.completed();
      return;
    }

    const recipient = toResult.value[0];
    const firstName = firstNameFromDisplayName(recipient.displayName, recipient.emailAddress);
    if (!firstName) {
      evt.completed();
      return;
    }

    item.body.getTypeAsync({ asyncContext: evt }, function (typeResult) {
      const evt2 = typeResult.asyncContext;
      if (typeResult.status !== Office.AsyncResultStatus.Succeeded) {
        evt2.completed();
        return;
      }

      const isHtml = typeResult.value === Office.CoercionType.Html;
      const coercionType = isHtml ? Office.CoercionType.Html : Office.CoercionType.Text;

      item.body.getAsync(coercionType, { asyncContext: evt2 }, function (bodyResult) {
        const evt3 = bodyResult.asyncContext;
        if (bodyResult.status !== Office.AsyncResultStatus.Succeeded) {
          evt3.completed();
          return;
        }

        const body = bodyResult.value || "";
        if (alreadyInserted(body, isHtml)) {
          evt3.completed();
          return;
        }

        const content = isHtml
          ? `<div data-bonjour-prenom="1"><p>Bonjour ${escapeHtml(firstName)},</p><p><br></p><p><br></p><p>${escapeHtml(closing)}</p><p><br></p></div>`
          : `Bonjour ${firstName},\r\n\r\n\r\n\r\n${closing}\r\n\r\n`;

        item.body.prependAsync(
          content,
          { coercionType: coercionType, asyncContext: evt3 },
          function (writeResult) {
            writeResult.asyncContext.completed();
          }
        );
      });
    });
  });
}

function alreadyInserted(body, isHtml) {
  if (isHtml && body.includes(HTML_MARKER)) return true;
  const plain = stripHtml(body).replace(/\s+/g, " ").toLocaleLowerCase("fr-FR");
  return plain.includes("bonjour ") &&
         (plain.includes("bien à toi") || plain.includes("bien à vous"));
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

function stripHtml(s) {
  return String(s).replace(/<[^>]*>/g, " ");
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({
    "&":"&amp;", "<":"&lt;", ">":"&gt;", '"':"&quot;", "'":"&#39;"
  })[c]);
}

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
Office.actions.associate("onMessageRecipientsChangedHandler", onMessageRecipientsChangedHandler);
