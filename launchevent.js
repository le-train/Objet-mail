/* =============================================================================
   Complement Outlook - Prefixes d'objet (v2) : activation evenementielle
   -----------------------------------------------------------------------------
   S'execute AUTOMATIQUEMENT a l'ouverture d'un nouveau message OU d'une nouvelle
   reunion, et affiche un bandeau de rappel avec un bouton qui ouvre le volet.

   OPTION : pour appliquer aussi un prefixe par defaut automatiquement, renseignez
   DEFAULT_PREFIX (ex. "[SLO]"). Laissez vide pour ne rien imposer.
   ============================================================================= */
var DEFAULT_PREFIX = ""; // ex. "[SLO]" ; "" = aucun prefixe applique d'office

// --- Message : declenche a chaque nouveau mail en composition ---
function onNewMessageComposeHandler(event) {
  runNudge("msgOpenPane", event);
}

// --- Reunion : declenche a chaque nouvelle reunion/rendez-vous en composition ---
function onNewAppointmentOrganizerHandler(event) {
  runNudge("apptOpenPane", event);
}

// Logique commune : prefixe par defaut eventuel, puis bandeau de rappel.
function runNudge(commandId, event) {
  applyDefaultPrefixIfNeeded(function () {
    showPrefixNudge(commandId, function () {
      event.completed(); // toujours signaler la fin, sinon Outlook bloque le runtime
    });
  });
}

// Affiche le bandeau d'information avec un bouton "Choisir un prefixe".
function showPrefixNudge(commandId, done) {
  var details = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
    message: "Pensez à choisir un préfixe d'objet normalisé.",
    icon: "icon16",
    actions: [
      {
        actionText: "Choisir un préfixe",
        actionType: Office.MailboxEnums.ActionType.ShowTaskPane,
        commandId: commandId,
        contextData: JSON.stringify({ source: "launchevent" })
      }
    ]
  };
  Office.context.mailbox.item.notificationMessages.addAsync("prefixNudge", details, function () {
    if (typeof done === "function") { done(); }
  });
}

// Applique eventuellement un prefixe par defaut, en evitant tout doublon.
function applyDefaultPrefixIfNeeded(done) {
  if (!DEFAULT_PREFIX) { return done(); }
  var item = Office.context.mailbox.item;
  item.subject.getAsync(function (res) {
    if (res.status !== Office.AsyncResultStatus.Succeeded) { return done(); }
    var current = res.value || "";
    if (current.indexOf(DEFAULT_PREFIX) === 0) { return done(); }
    var newSubject = (DEFAULT_PREFIX + " " + current).trim();
    if (newSubject.length > 255) { newSubject = newSubject.substring(0, 255); }
    item.subject.setAsync(newSubject, function () { done(); });
  });
}

// Relie les noms declares dans le manifeste aux fonctions ci-dessus.
if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
  Office.actions.associate("onNewAppointmentOrganizerHandler", onNewAppointmentOrganizerHandler);
}
