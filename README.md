# Baromètre social KAMA CI — application web

Les collaborateurs remplissent le questionnaire dans leur navigateur. Chaque soumission est
écrite automatiquement dans la feuille **Saisie des réponses** du classeur
`Téléchargements\Barometre_social_KAMA_CI.xlsx` (modifiable dans `config.local.bat`). L'onglet **Résultats** (scores, lecture, graphique)
se recalcule à l'ouverture du fichier dans Excel.

## Lancer

Double-cliquer sur `lancer.bat` (il crée `config.local.bat` à partir de `config.exemple.bat`
au premier lancement ; y définir le chemin du classeur et le code d'accès), puis ouvrir :

- sur ce PC : http://localhost:8080
- depuis un autre poste du réseau : `http://<adresse-IP-de-ce-PC>:8080`
  (autoriser le port 8080 dans le pare-feu Windows si nécessaire)

## Page des résultats (responsables)

http://localhost:8080/resultats — protégée par le code défini dans `config.local.bat`
(`BAROMETRE_ADMIN_CODE`, à changer). Elle affiche le nombre de répondants, le score global,
les scores par thème et le contenu de la feuille « Saisie des réponses », avec un bouton pour
télécharger le fichier Excel.

## Fonctionnement

- Les questions sont lues dans l'onglet **Questionnaire** : modifier une question dans Excel
  puis relancer l'application suffit.
- Le collaborateur ne voit que le questionnaire (service, ancienneté, 24 notes de 1 à 5,
  commentaire facultatif). Aucun nom n'est demandé.
- **Fichier ouvert dans Excel** : Excel verrouille le fichier. Les réponses sont alors mises en
  attente dans `donnees/reponses_en_attente.json` et écrites automatiquement (nouvel essai
  chaque minute) dès que le fichier est fermé. Le collaborateur, lui, voit sa soumission validée.
- Chaque soumission est aussi copiée dans `donnees/journal_reponses.jsonl` (sauvegarde).
- Le fichier est modifié directement au niveau XML : graphique, commentaires et mise en
  forme sont conservés.

## Correction apportée au classeur

Les formules de **Résultats** portaient sur les lignes 3 à 41 de la saisie : le 1er répondant
(ligne 2) était ignoré et la capacité limitée à 40 réponses. À la première écriture,
l'application les étend automatiquement aux lignes 2 à 1000.

## Configuration (`config.local.bat` ou variables d'environnement)

Le classeur Excel, le dossier `donnees/` et `config.local.bat` ne sont jamais envoyés sur Git.

| Variable | Rôle | Défaut |
|---|---|---|
| `BAROMETRE_EXCEL` | Chemin du classeur | `Téléchargements\Barometre_social_KAMA_CI.xlsx` |
| `BAROMETRE_ADMIN_CODE` | Code de la page /resultats | (page désactivée si vide) |
| `BAROMETRE_SERVICES` | Liste déroulante des services, séparés par `;` | saisie libre |
| `PORT` / `HOST` | Port / interface d'écoute | `8080` / `0.0.0.0` |
