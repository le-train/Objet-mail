/* =============================================================================
   Complement Outlook - Prefixes d'objet : activation evenementielle
   -----------------------------------------------------------------------------
   Ce code s'execute AUTOMATIQUEMENT a l'ouverture d'un nouveau message, sans clic.
   Comme Outlook ne permet pas d'ouvrir de force le volet interactif, on affiche
   une barre de notification avec un bouton qui ouvre le volet en un clic.

   OPTION : pour appliquer aussi un prefixe par defaut automatiquement (que
   l'utilisateur pourra changer ensuite), mettez DEFAULT_PREFIX a la valeur voulue,
   par ex. "[INTERNE]". Laissez la chaine vide pour ne rien imposer.
   ============================================================================= */
var DEFAULT_PREFIX = ""; // ex. "[INTERNE]" ; "" = aucun prefixe applique d'office

// Declenche a chaque nouveau message en composition.
function onNewMessageComposeHandler(event) {
  applyDefaultPrefixIfNeeded(function () {
    showPrefixNudge(function () {
      // Toujours signaler la fin du traitement, sinon Outlook bloque le runtime.
      event.completed();
    });
  });
}

// Affiche la barre d'information avec un bouton "Choisir un prefixe".
function showPrefixNudge(done) {
  var details = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
    message: "Pensez à choisir un préfixe d'objet normalisé.",
    icon: "icon16",
    actions: [
      {
        actionText: "Choisir un préfixe",
        actionType: Office.MailboxEnums.ActionType.ShowTaskPane,
        commandId: "openPrefixPane",
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

// Relie le nom declare dans le manifeste a la fonction ci-dessus.
// (Office.onReady n'est PAS execute dans le runtime d'evenement : on associe directement.)
if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
}
