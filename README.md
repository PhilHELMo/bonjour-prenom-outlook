Voici le README mis à jour pour la **V4**, en remplaçant la description de l'ancienne version manuelle par le fonctionnement automatique actuel.

# Bonjour Prénom — complément Outlook Web

## Fonction

**Bonjour Prénom** est un mini complément pour Outlook Web qui ajoute automatiquement une salutation personnalisée et une formule de fin lors de la rédaction d'un message.

Le complément utilise le premier destinataire du champ **À**, récupère son nom d'affichage Outlook et tente d'en extraire le prénom.

Par exemple, pour le destinataire **Jean Dupont**, le complément ajoute automatiquement :

```text
Bonjour Jean,


Bien à toi

[signature Outlook]
```

La signature Outlook existante est conservée.

## Fonctionnement automatique

Le complément fonctionne sans devoir cliquer sur le bouton **Bonjour Prénom**.

### Réponse, Répondre à tous et Transférer

Lors de l'ouverture de la fenêtre de rédaction, le complément utilise l'événement Outlook :

`OnNewMessageCompose`

Si le premier destinataire du champ **À** est déjà disponible, son prénom est détecté et la salutation est insérée automatiquement.

### Nouveau message

Lors de la création d'un nouveau message, aucun destinataire n'est encore connu. Le complément attend donc que vous ajoutiez un destinataire.

Lorsque vous saisissez une adresse ou un nom et sélectionnez le contact proposé par Outlook, le complément utilise l'événement :

`OnMessageRecipientsChanged`

Il lit alors le premier destinataire du champ **À**, extrait son prénom et insère automatiquement la salutation.

Par exemple :

1. Cliquez sur **Nouveau message**.
2. Commencez à saisir `Jean Dupont`.
3. Sélectionnez le contact proposé par Outlook.
4. Dès que Jean Dupont apparaît dans le champ **À**, le complément peut ajouter automatiquement `Bonjour Jean,`.

## Formule de fin

La formule de fin par défaut est :

`Bien à toi`

Le panneau **Bonjour Prénom** permet également de sélectionner :

`Bien à vous`

Le choix est enregistré dans les paramètres itinérants Office (`roamingSettings`).

Le bouton du complément reste disponible pour ouvrir le panneau et effectuer une insertion manuelle si nécessaire.

## Prévention des doublons

Les événements Outlook peuvent être déclenchés plusieurs fois, notamment lorsque les destinataires sont modifiés.

Le complément vérifie donc le contenu du message avant l'insertion afin d'éviter d'ajouter plusieurs fois la salutation et la formule de fin.

## Détection du prénom

Le prénom est déduit du `displayName` fourni par Outlook.

Quelques exemples :

- `Jean Dupont` → `Jean`
- `DUPONT, Jean` → `Jean`
- `JEAN DUPONT` → `Jean`

Si Outlook ne fournit pas de nom d'affichage exploitable, le complément tente de déduire un nom à partir de la partie située avant `@` dans l'adresse électronique.

Les noms complexes peuvent nécessiter une adaptation.

Le complément n'interroge volontairement aucun annuaire ni service externe pour rechercher le prénom.

## Confidentialité

Le code fourni n'a aucun backend, aucune base de données et aucun appel réseau applicatif.

Le seul script externe chargé par le complément est **Office.js depuis Microsoft**.

Les fichiers du complément sont hébergés sous forme de fichiers statiques sur GitHub Pages. Le code ne transmet pas à cet hébergement le nom du destinataire, son adresse électronique ou le contenu du message.

Le traitement du nom et la modification du corps du message sont effectués dans le contexte du complément Outlook.

## Permission

Le manifeste demande :

`ReadWriteItem`

Cette permission permet au complément d'accéder aux informations nécessaires concernant l'élément Outlook en cours et d'en modifier le corps.

Le complément ne demande pas :

`ReadWriteMailbox`

## Événements Outlook utilisés

La version V4 utilise deux événements d'activation :

- `OnNewMessageCompose` : ouverture d'un nouveau formulaire de composition, notamment pour les réponses et transferts ;
- `OnMessageRecipientsChanged` : modification des destinataires pendant la rédaction.

La configuration automatique de la V4 nécessite la prise en charge de **Mailbox 1.11**.

## Fichiers

Le complément utilise notamment :

```text
manifest.xml
taskpane.html
taskpane.js
commands.html
autorun.html
autorun.js
icon-16.png
icon-32.png
icon-64.png
icon-80.png
icon-128.png
```

### `manifest.xml`

Déclare le complément, ses permissions, son bouton et les événements d'activation Outlook.

### `taskpane.html` / `taskpane.js`

Gèrent le panneau **Bonjour Prénom**, le choix entre `Bien à toi` et `Bien à vous` et l'insertion manuelle.

### `autorun.html` / `autorun.js`

Gèrent l'activation automatique lors de l'ouverture d'une composition et lors de la modification des destinataires.

### `commands.html`

Page utilisée par Outlook pour les commandes du complément.

## Hébergement

Les fichiers statiques sont actuellement prévus pour être hébergés avec GitHub Pages.

Le `manifest.xml` contient les URL HTTPS permettant à Outlook de charger les fichiers du complément.

Après une modification de `taskpane.js`, `autorun.js` ou d'un autre fichier hébergé, il faut attendre le redéploiement de GitHub Pages avant de tester la nouvelle version dans Outlook.

## Installation

1. Hébergez les fichiers HTML, JavaScript et les icônes sur GitHub Pages.
2. Vérifiez que les fichiers sont accessibles en HTTPS.
3. Dans Outlook Web, ouvrez la gestion des compléments.
4. Allez dans **Mes compléments** puis **Compléments personnalisés**.
5. Choisissez **Ajouter à partir d'un fichier**.
6. Sélectionnez `manifest.xml`.
7. Confirmez l'installation.
8. Actualisez complètement Outlook Web.

Lors d'une mise à jour importante du manifeste, il peut être nécessaire de supprimer la version précédente du complément puis de réimporter le nouveau `manifest.xml`.

## Test

### Tester une réponse

1. Ouvrez un message reçu.
2. Cliquez sur **Répondre**.
3. Ne cliquez pas sur le bouton **Bonjour Prénom**.
4. Vérifiez que la salutation apparaît automatiquement avec le prénom du destinataire.

### Tester un nouveau message

1. Cliquez sur **Nouveau message**.
2. Le complément ne doit encore rien insérer.
3. Saisissez le nom ou l'adresse d'un destinataire.
4. Sélectionnez le destinataire proposé par Outlook.
5. Vérifiez que `Bonjour Prénom,` apparaît automatiquement après l'ajout du destinataire.

## Limites

Le résultat dépend des informations de destinataire fournies par Outlook.

Le complément utilise uniquement le **premier destinataire du champ À** pour déterminer le prénom.

Il ne tente pas de choisir une salutation différente lorsqu'il y a plusieurs destinataires.

La détection automatique des destinataires nécessite la prise en charge des événements Outlook utilisés par le manifeste, notamment **Mailbox 1.11**.

## Version

**V4**

Fonctionnalités principales :

- insertion automatique sur les réponses et compositions ;
- détection du destinataire lors d'un nouveau message ;
- extraction automatique du prénom ;
- choix entre `Bien à toi` et `Bien à vous` ;
- conservation de la signature Outlook existante ;
- prévention des insertions en double ;
- bouton disponible comme solution manuelle.

Vous pouvez remplacer intégralement le contenu de votre `README.md` GitHub par cette version.
