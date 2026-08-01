# Dépannage

## Le bouton Excel ne fait rien

L’installateur doit avoir été exécuté une fois et le fichier `Actualiser_Excel.cmd` doit se trouver dans le même dossier que le classeur.

## Windows affiche un avertissement

Le lanceur est un script PowerShell téléchargé depuis le dépôt. Vérifier que l’URL correspond bien à `Nathan20202/excel-catalogues-auto-update`, puis utiliser le raccourci créé par l’installateur.

## Aucun classeur connecté n’est trouvé

Le dossier sélectionné doit contenir les versions comportant l’onglet `00 - Actualisation`. Le moteur reconnaît la clé technique stockée dans cet onglet.

## Une mise à jour échoue

Consulter `Actualisation.log` dans le dossier des classeurs. Aucune écriture n’est poursuivie pour le jeu de données en erreur. La sauvegarde se trouve dans `_sauvegardes_actualisation`.

## Restaurer une sauvegarde

Fermer Excel, copier la sauvegarde souhaitée depuis `_sauvegardes_actualisation`, puis lui redonner le nom du classeur principal.

