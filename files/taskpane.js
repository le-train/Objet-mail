/* =============================================================================
   Complement Outlook - Prefixes d'objet
   -----------------------------------------------------------------------------
   POUR PERSONNALISER : modifiez uniquement le tableau PREFIXES ci-dessous.
   - label  : texte affiche sur le bouton
   - tag    : prefixe reellement insere dans l'objet (entre crochets recommande)
   - color  : couleur du texte / de la pastille
   - bg     : couleur de fond de la pastille
   Ajoutez ou retirez des entrees librement. Aucune autre modification requise.
   ============================================================================= */
var PREFIXES = [
  { label: "Projet X",   tag: "[PROJET-X]",   color: "#185FA5", bg: "#E6F1FB" },
  { label: "Urgent",     tag: "[URGENT]",     color: "#A32D2D", bg: "#FCEBEB" },
  { label: "RH",         tag: "[RH]",         color: "#534AB7", bg: "#EEEDFE" },
  { label: "Commercial", tag: "[COMMERCIAL]", color: "#0F6E56", bg: "#E1F5EE" },
  { label: "Interne",    tag: "[INTERNE]",    color: "#5F5E5A", bg: "#F1EFE8" }
];

/* -----------------------------------------------------------------------------
   A partir d'ici, rien n'a besoin d'etre modifie pour un usage standard.
   ----------------------------------------------------------------------------- */

var mailItem = null;

Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    mailItem = Office.context.mailbox.item;
    buildButtons();
    document.getElementById("clearBtn").addEventListener("click", removePrefix);

    // Volet epingle : se mettre a jour quand l'utilisateur change de message.
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, function () {
      mailItem = Office.context.mailbox.item;
      setStatus("");
    });
  }
});

// Retire la barre de notification "Pensez a choisir un prefixe" si elle est presente.
function dismissNudge() {
  if (mailItem && mailItem.notificationMessages) {
    mailItem.notificationMessages.removeAsync("prefixNudge", function () {});
  }
}

// Construit dynamiquement un bouton par prefixe defini ci-dessus.
function buildButtons() {
  var container = document.getElementById("buttons");
  PREFIXES.forEach(function (p) {
    var btn = document.createElement("button");
    btn.className = "btn";
    btn.setAttribute("aria-label", "Insérer le préfixe " + p.tag);

    var name = document.createElement("span");
    name.className = "name";
    name.textContent = p.label;
    name.style.color = p.color;

    var tag = document.createElement("span");
    tag.className = "tag";
    tag.textContent = p.tag;
    tag.style.color = p.color;
    tag.style.background = p.bg;

    btn.appendChild(name);
    btn.appendChild(tag);
    btn.addEventListener("click", function () { applyPrefix(p.tag); });
    container.appendChild(btn);
  });
}

// Construit une expression reguliere qui reconnait n'importe lequel de nos prefixes
// en debut d'objet (pour eviter d'empiler [URGENT] [URGENT] ...).
function buildPrefixRegex() {
  var escaped = PREFIXES.map(function (p) {
    return p.tag.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  });
  // ex : ^\s*(\[URGENT\]|\[RH\])\s*
  return new RegExp("^\\s*(" + escaped.join("|") + ")\\s*", "i");
}

// Retire un eventuel prefixe gere deja present, sans toucher au reste de l'objet.
function stripManagedPrefix(subject) {
  return (subject || "").replace(buildPrefixRegex(), "").replace(/^\s+/, "");
}

// Insere (ou remplace) le prefixe choisi au debut de l'objet.
function applyPrefix(tag) {
  if (!mailItem) { return; }
  mailItem.subject.getAsync(function (res) {
    if (res.status !== Office.AsyncResultStatus.Succeeded) {
      return setStatus("Impossible de lire l'objet.", "err");
    }
    var base = stripManagedPrefix(res.value);
    var newSubject = (tag + " " + base).trim();
    if (newSubject.length > 255) { newSubject = newSubject.substring(0, 255); }

    mailItem.subject.setAsync(newSubject, function (r) {
      if (r.status === Office.AsyncResultStatus.Succeeded) {
        dismissNudge();
        setStatus("Préfixe " + tag + " appliqué.", "ok");
      } else {
        setStatus("L'objet n'a pas pu être mis à jour.", "err");
      }
    });
  });
}

// Retire le prefixe gere de l'objet courant.
function removePrefix() {
  if (!mailItem) { return; }
  mailItem.subject.getAsync(function (res) {
    if (res.status !== Office.AsyncResultStatus.Succeeded) {
      return setStatus("Impossible de lire l'objet.", "err");
    }
    var base = stripManagedPrefix(res.value);
    mailItem.subject.setAsync(base, function (r) {
      if (r.status === Office.AsyncResultStatus.Succeeded) {
        setStatus("Préfixe retiré.", "ok");
      } else {
        setStatus("L'objet n'a pas pu être mis à jour.", "err");
      }
    });
  });
}

function setStatus(msg, kind) {
  var el = document.getElementById("status");
  el.textContent = msg;
  el.className = "status" + (kind ? " " + kind : "");
}
