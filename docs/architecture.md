# Architecture et garanties

## Flux

1. GitHub Actions vérifie périodiquement les sources publiques.
2. Les catalogues JSON versionnés sont mis à jour dans `data/`.
3. Le lanceur Windows télécharge `config/workbooks.json` et les jeux de données utiles.
4. Excel est piloté localement par son interface COM.
5. Chaque enregistrement distant est associé à une ligne Excel par son ID.
6. Seules les colonnes publiques sont écrites.
7. Le classeur est recalculé, sauvegardé et journalisé.

## Règles anti-perte

- Une sauvegarde `.xlsx` est créée avant toute écriture.
- Les colonnes listées dans `preserveColumns` ne sont jamais écrites.
- Les colonnes listées dans `formulaColumns` ne sont jamais écrites.
- Une entrée absente du dépôt n’est jamais supprimée du classeur local.
- Une nouveauté reçoit une nouvelle ligne et des valeurs personnelles par défaut.
- Les propriétés techniques commençant par `_` ne sont jamais envoyées vers Excel.
- Le moteur interrompt le traitement si l’ID, l’onglet ou le tableau attendu manque.

## Mise à jour conservative

Un contrôle automatique ne remplace pas une validation éditoriale. Les pages inaccessibles, changements de contenu et nouveautés détectées sont journalisés. Les candidats incomplets restent hors du catalogue principal tant qu’ils ne disposent pas d’assez d’éléments fiables.

## Coût

L’architecture n’utilise que :

- un dépôt GitHub public ;
- GitHub Actions sur les quotas gratuits des dépôts publics ;
- des sources publiques sans clé obligatoire ;
- Windows PowerShell et Excel déjà installés sur l’ordinateur.

