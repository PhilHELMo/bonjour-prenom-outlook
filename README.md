# Bonjour Prénom — mini complément Outlook Web

## Fonction
Dans une réponse ou un nouveau message, le bouton **Bonjour Prénom** ouvre un petit panneau.
Cliquez sur **Insérer « Bonjour Prénom, »** : le complément lit uniquement le premier destinataire
du champ **À**, extrait son prénom à partir du nom d'affichage, puis ajoute `Bonjour Prénom,`
au début du corps.

## Confidentialité
Le code fourni n'a aucun backend, aucune base de données et aucun appel réseau applicatif.
Le seul script externe chargé est **Office.js depuis Microsoft**.
Le site HTTPS qui héberge les fichiers statiques les sert au navigateur ; le code ne lui transmet
pas le nom, l'adresse ou le contenu du message.

## Permission
Le manifeste demande `ReadWriteItem`, nécessaire pour lire les informations du message en cours
et modifier son corps. Il ne demande pas `ReadWriteMailbox`.

## Installation
1. Hébergez le contenu de ce dossier sur une URL HTTPS (par ex. un hébergement statique).
2. Dans `manifest.xml`, remplacez TOUTES les occurrences de `YOUR_HTTPS_DOMAIN`
   par votre domaine, sans `https://` et sans slash final.
   Exemple : `moncompte.github.io/bonjour-prenom`
   ATTENTION : si votre hébergement est dans un sous-dossier, les URL du manifeste doivent
   pointer exactement vers ce sous-dossier.
3. Vérifiez que `taskpane.html`, `taskpane.js`, `commands.html` et les icônes sont accessibles en HTTPS.
4. Dans Outlook Web : Applications / Compléments > Mes compléments > Compléments personnalisés
   > Ajouter à partir d'un fichier, puis choisissez `manifest.xml`.
5. Ouvrez une réponse, cherchez **Bonjour Prénom** dans les applications/actions de rédaction,
   puis épinglez-le si Outlook le permet.

## Limite volontaire
Le prénom est déduit du `displayName` Outlook. Pour `Jean Dupont`, le résultat est `Jean`.
Pour `DUPONT, Jean`, le code tente `Jean`. Les noms complexes peuvent nécessiter une adaptation.
Le complément n'interroge volontairement aucun annuaire ni service externe.

## Variante automatique
Cette version est volontairement déclenchée par un bouton : elle évite l'activation événementielle
et reste simple à auditer. Une version qui insère automatiquement la salutation dès l'ouverture
d'une réponse est possible, mais nécessite un manifeste/événement plus complexe.
