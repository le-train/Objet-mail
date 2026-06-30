# Complément Outlook — Préfixes d'objet

Volet qui s'affiche pendant la rédaction d'un message et permet d'insérer, en un
clic, un préfixe normalisé au début de l'objet (ex. `[PROJET-X]`, `[URGENT]`).
Fonctionne sur Outlook classique (Windows), nouvel Outlook, Outlook sur le web et
Outlook pour Mac. Aucune installation sur les postes : tout est hébergé en ligne et
déployé de façon centralisée.

## Contenu du dossier

- `manifest.xml` — déclaration du complément (bouton de ruban, volet, permissions,
  et activation automatique à la composition).
- `taskpane.html` — interface du volet.
- `taskpane.js` — logique du volet : construit les boutons et écrit le préfixe.
- `commands.html` — page hôte (sans UI) du code d'événement, pour nouvel Outlook/Web/Mac.
- `launchevent.js` — code exécuté automatiquement à l'ouverture d'un nouveau message.
- `assets/` — icônes (à remplacer par les vôtres si besoin).

## Comment se comporte l'ouverture « automatique »

Important : Outlook **n'autorise pas** un complément à ouvrir de force le volet
interactif (avec les boutons) pendant la composition — cette capacité existe dans
Word/Excel mais pas pour les compléments de messagerie. Ce projet obtient l'effet
recherché par deux mécanismes complémentaires :

1. À chaque nouveau message, une **barre de notification** apparaît automatiquement
   (« Pensez à choisir un préfixe ») avec un bouton qui ouvre le volet en un clic.
2. Le volet est **épinglable** : une fois épinglé par l'utilisateur, il reste affiché
   d'un message à l'autre, sans reclic.

Vous pouvez aussi faire appliquer un **préfixe par défaut** automatiquement
(modifiable ensuite) : voir `DEFAULT_PREFIX` en haut de `launchevent.js`.

## 1. Personnaliser la liste des préfixes

Ouvrez `taskpane.js` et modifiez uniquement le tableau `PREFIXES` en haut du fichier.
Chaque entrée a un libellé (`label`), le préfixe inséré (`tag`) et deux couleurs.
Ajoutez ou retirez des lignes librement ; le reste du code s'adapte automatiquement
(y compris la protection contre l'empilement de préfixes type `[URGENT] [URGENT]`).

## 2. Héberger les fichiers en HTTPS

Les fichiers doivent être servis par une URL **HTTPS** (obligatoire). Options simples :
un conteneur de stockage statique (Azure Storage, etc.), un hébergement web interne,
ou tout serveur HTTPS de l'entreprise. Déposez `taskpane.html`, `taskpane.js` et le
dossier `assets/` à la racine du site.

Puis, dans `manifest.xml`, remplacez **toutes** les occurrences de
`https://VOTRE-DOMAINE.example.com` par votre URL réelle, et remplacez le `Id`
(GUID) par un identifiant unique (générateur de GUID en ligne).

## 3. Tester sur un poste (sideload)

- **Nouvel Outlook / Web** : Paramètres → Compléments / « Gérer les compléments » →
  « Mes compléments » → « Ajouter un complément personnalisé » → à partir d'un fichier
  → sélectionnez votre `manifest.xml`.
- **Outlook classique (Windows)** : Fichier → Gérer les compléments (ouvre Outlook sur
  le web) → même procédure « Ajouter un complément personnalisé ».

Ouvrez un nouveau message : un bouton **Préfixes** apparaît dans le ruban. Cliquez‑le,
le volet s'ouvre. Comme le volet est *épinglable* (`SupportsPinning`), vous pouvez
l'épingler pour qu'il reste ouvert d'un message à l'autre.

## 4. Déployer à toute l'entreprise

Dans le **Centre d'administration Microsoft 365** → Paramètres → Applications intégrées
(ou Centre d'administration Exchange → Compléments) → Téléverser un complément
personnalisé → à partir du `manifest.xml` → choisissez les utilisateurs/groupes
concernés. La mise à jour de la liste des préfixes se fait ensuite côté serveur
(remplacer `taskpane.js`), sans retoucher les postes.

## Pour aller plus loin (optionnel)

- **Préfixe par défaut imposé** : renseignez `DEFAULT_PREFIX` dans `launchevent.js`.
- **Imposer un préfixe avant l'envoi** : on peut ajouter une validation sur l'événement
  d'envoi (`OnMessageSend`) pour bloquer l'envoi si aucun préfixe n'est présent. Cela
  demande un événement supplémentaire dans le manifeste et une fonction dédiée.
- **Manifeste unifié (JSON)** : Microsoft propose aussi un format de manifeste plus
  récent. Le `manifest.xml` fourni ici reste pleinement pris en charge.

## Notes techniques importantes

- L'activation automatique exige le jeu de fonctionnalités **Mailbox 1.10** (déjà
  déclaré dans le manifeste). Sur des clients plus anciens, le bouton de ruban et le
  volet fonctionnent quand même ; seule la barre de notification automatique ne
  s'affiche pas.
- Après modification du manifeste d'un complément à activation événementielle déjà
  déployé par l'administrateur, un **nouveau consentement administrateur** est requis
  dans le Centre d'administration ; les utilisateurs sont bloqués tant qu'il n'est pas
  accordé.
- Le code d'événement (`launchevent.js`) ne passe **pas** par `Office.onReady` : la
  liaison se fait via `Office.actions.associate`, déjà en place en bas du fichier.

> Remarque : un complément web nécessite une connexion réseau pour se charger, et n'a
> aucun code installé localement sur le poste.
