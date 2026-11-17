# Catalogue Carrefour Publimax : mise en service

Appli de saisie des dossiers magasin (tablette, installable, hors ligne). Le catalogue et les images sont chiffrés
(`data/*.bin`) : sans la clé, le site ne montre rien. La clé n'est PAS dans ce dépôt.

## 1. Hébergement
Brancher ce dépôt sur l'hébergeur (Netlify : Add new site > Import an existing project > GitHub > `publimax-catalogue`,
aucune commande de build, dossier de publication `.`). Chaque push met le site à jour.

## 2. Accès immédiat (sans script)
Lien d'accès : `https://ADRESSE-DU-SITE/#k=CLÉ`. La personne ouvre le lien sur sa tablette, choisit son code, installe l'appli.
À n'envoyer qu'aux personnes autorisées (qui a le lien a le catalogue).

## 3. Demande d'accès par e-mail (script Google)
1. script.google.com > Nouveau projet, coller `apps-script/Code.gs`.
2. Paramètres du projet > Propriétés du script : ajouter `KEY` = la clé du catalogue.
3. Exécuter une fois la fonction `autoriser` et accepter les autorisations.
4. Déployer > Nouveau déploiement > Application Web : exécuter en tant que « Moi », accès « Tout le monde ». Copier l'URL `/exec`.
5. Mettre cette URL dans `config.json` : `{"apps":"https://script.google.com/macros/s/.../exec"}` et pousser.
Ensuite : la personne saisit son e-mail, tu reçois un mail Autoriser / Refuser, sa tablette se débloque et elle choisit son code.
Après un changement de droits dans le script : refaire un **Nouveau déploiement** (nouvelle URL).

## Dossiers
Les dossiers restent sur la tablette (navigateur). Récap > Exporter (.json), Partager, Imprimer, ou « Envoyer à Publimax »
(Drive > « Catalogue Carrefour - dossiers », si le script est en place).
