# Excel Catalogues Auto Update

Infrastructure gratuite d’actualisation pour six classeurs Excel personnels :

- films, séries et animés ;
- marques de vêtements ;
- sites européens de cartes Pokémon ;
- codes promo, cashback et comparateurs ;
- Google Cloud Digital Leader.
- guide Tech & High-Tech : meilleurs achats, rapports qualité/prix et configurations PC.

Le classeur Google Cloud comprend trois examens blancs complets de 60 questions et leurs trois corrigés détaillés. Les réponses saisies restent locales et sont protégées lors de chaque actualisation.

## Principe

GitHub héberge uniquement les catalogues publics et exécute les contrôles planifiés. Les données personnelles restent dans les fichiers Excel de l’utilisateur :

- statuts et cases de suivi ;
- notes et favoris ;
- commentaires ;
- réponses aux examens ;
- historique personnel.

Le moteur Windows associe chaque ligne distante à la ligne locale grâce à un identifiant stable. Il met à jour les colonnes publiques, ajoute les nouveautés et ignore explicitement les colonnes personnelles et les formules. Avant chaque opération, il crée une sauvegarde horodatée.

## Installation

1. Télécharger [`Installer-Actualisation-Excel.ps1`](installer/Installer-Actualisation-Excel.ps1).
2. Placer les classeurs dans un même dossier.
3. Exécuter l’installateur avec PowerShell et sélectionner ce dossier.
4. Utiliser ensuite le raccourci **Actualiser mes classeurs Excel** ou le bouton présent dans l’onglet `00 - Actualisation`.

Aucun abonnement, aucune clé API et aucun secret GitHub ne sont nécessaires sur l’ordinateur.

## Cadences

| Catalogue | Actualisation centrale |
|---|---|
| Cinéma, séries et animés | Chaque lundi |
| Sites Pokémon | Chaque mercredi |
| Codes promo et comparateurs | Chaque vendredi |
| Marques de vêtements | Le 3 de chaque mois |
| Google Cloud Digital Leader | Le 5 de chaque mois et détection des changements officiels |
| Tech & High-Tech | Chaque samedi, avec validation éditoriale mensuelle minimum |

Tous les workflows peuvent aussi être lancés manuellement depuis l’onglet **Actions** de GitHub.

## Sources et prudence

Le système privilégie les données publiques, les sites officiels et les sources sans clé payante. Les métadonnées établies peuvent être mises à jour automatiquement. Les nouvelles entrées ambiguës sont placées dans des fichiers de candidats afin d’éviter d’injecter silencieusement une marque, une boutique ou une œuvre mal identifiée.

## Structure

- `data/` : catalogues publics et états de contrôle ;
- `config/workbooks.json` : règles de correspondance et colonnes protégées ;
- `scripts/` : contrôles, validation et rafraîchissement ;
- `installer/` : moteur local Windows et installation ;
- `.github/workflows/` : planification gratuite avec GitHub Actions.

## Confidentialité

Le dépôt ne doit jamais contenir de mot de passe, d’adresse privée, de notes personnelles ou de réponses d’examen. Le moteur local ne téléverse rien : il effectue uniquement des téléchargements HTTPS depuis ce dépôt public.
