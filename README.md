# Xlsx2Kobo
Un outil pour transformer un fichier Excel (.xlsx) en données XML et téléchargement vers la plateforme KoboToolbox.  
Version de python utilisé: 3.13.5

Après avoir cloner ce dépôt en local, suivez les étapes suivantes

## Créer et activez un environnement virtuel
```bash
python -m venv venv
source venv/Scripts/activate
```

## Installer les packages nécessaires

```bash
pip install -r requirements.txt
```
python xlsx2xml.py data.xlsx

## Préparer le fichier Excel en suivant les règles de mise en forme suivant
1. Le fichier excel doit contenir la feuille ```IDSheet``` (après la feuille contenant les données). Cette feuille contient les constantes `Token` et `ID du formulaire` comme dans l'exemple suivant: 
|ID du formulaire|a7baHDrXP8tRzfBdVeJVPY|
|-----|----|
|Token|0aded612c72ed182e017dca80a17dd71b97b7225|
3. Les entêtes des colonnes dans l'onglet des données doivent être le format xml
5. \_\_version\_\_ et les autres colonnes de métadonnées doivent suivre les colonnes de données. \_\_version\_\_ est la premièer colonne après le contenu.
6. Formatez toutes les cellules en format texte.
7. Les groupes répétées sont répertoriés dans de nouvelles feuilles
8. Les groupes non-répétés sont à mettre dans le premier onglet. Les colonnes contiennent le caractère '::' comme indicateur de groupe.
9. Pas besoin de remplir _id, _uuid, ils seront générés automatiquement. Par contre il est important de remplir la version, autrement il sera automatiquement rempli par ```None```
6. Remplir les _uuid pour mettre à jour des soumissions spécifiques

## Transformer les données sous format xml
```bash
python xlsx2xlm.py data.xlsx
```
```data.xlsx``` étant le nom du fichier

## Télécharger les données sur les serveur KoboToolbox
post.py postera les données dans tempfiles/*.xml vers kc.kobotoolbox.org
Avant de lancer ```post.py```:
- Vérifiez si il y a un fichier ~/.netrc existant. Sinon, créez-le.
- Assurez-vous que les permissions du fichier ~/.netrc sont read/write, avec un accès restreint uniquement au propriétaire. La commande chmod ci-bas fera cela.
- Utilisez l'exemple de fichie .netrc file comme un guide, modifiez  ~/.netrc avec vos identifiants de connexion.
- Si vous n’utilisez pas kc.kobotoolbox.org comme serveur de destination, mettez à jour dans le fichier post.py (lignes 10 à 12) l’hôte du fichier .netrc, l’URL du serveur et le préfixe du nom du fichier journal (log).
- Les login dans le fichier .netrc doivent correspondre au compte propriétaire pour avoir tous les accès requis

```
$ chmod 600 .netrc
$ python post.py
```
