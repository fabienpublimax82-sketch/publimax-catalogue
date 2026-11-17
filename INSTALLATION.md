# Catalogue Carrefour Publimax : mise en service

Appli de saisie des dossiers magasin (tablette, installable, hors ligne) : https://fabienpublimax82-sketch.github.io/publimax-catalogue/
Le catalogue et les images sont chiffrés (`data/*.bin`) : sans la clé, le site ne montre rien. La clé n'est PAS dans ce dépôt.

## Accès
Lien d'accès : `https://fabienpublimax82-sketch.github.io/publimax-catalogue/#k=CLÉ`. La personne ouvre le lien sur sa tablette,
choisit son code, installe l'appli. À n'envoyer qu'aux personnes autorisées (qui a le lien a le catalogue).
La demande d'accès par e-mail n'est pas en service dans cette version.

## Envoi vers Google Drive (photos + Sheet de récap)
1. script.google.com > Nouveau projet, coller `apps-script/Code.gs` à la place du contenu.
2. Paramètres du projet (roue dentée) > Propriétés du script : ajouter `KEY` = la clé du catalogue (ce qui suit `#k=` dans le lien).
3. Dans l'éditeur, choisir la fonction `autoriser`, Exécuter, accepter les autorisations. Elle crée le dossier Drive
   « Catalogue Carrefour - surveys » et le classeur « Catalogue Carrefour - récap surveys ».
4. Déployer > Nouveau déploiement > Application Web : exécuter en tant que « Moi », accès « Tout le monde ». Copier l'URL `/exec`.
5. Dans l'appli, écran des dossiers, bouton ⚙ : coller cette URL. (Ou la mettre dans `config.json` : `{"apps":"https://script.google.com/macros/s/.../exec"}`.)

Ensuite : chaque photo part dans Drive > Catalogue Carrefour - surveys > « Magasin date [id] », les lignes retenues dans le classeur
(onglets Lignes et Dossiers, avec les liens des photos). Sans réseau tout attend sur la tablette et part au retour du réseau,
à l'ouverture de l'appli ou au bouton ☁.
Pour modifier le script plus tard : Déployer > Gérer les déploiements > modifier > Nouvelle version (l'URL ne change pas).

## Dossiers
Les dossiers restent aussi sur la tablette. Récap > Exporter (.json), Partager, Imprimer en secours.
