# Artisans de confiance

Projet de groupe réalisé à l’EPF (promotion 2028) autour de la gestion d’une entreprise d’artisans, avec **Excel VBA, MySQL et SQL**.

L’objectif est de développer un outil de gestion permettant de centraliser les informations sur les artisans, les fournisseurs et les clients, de suivre les commandes et de disposer d’indicateurs pour piloter l’activité commerciale.

## Fonctionnalités

### Gestion des artisans et des fournisseurs
- Gestion des artisans : ajout, modification et suppression
- Consultation des fiches artisans et fournisseurs
- Suivi des stocks et des seuils critiques
- Calcul du prix moyen par type de travaux
- Consultation des disponibilités des artisans

### Gestion des clients et des commandes
- Gestion des comptes clients et consultation de leur historique
- Consultation des commandes et des factures
- Recherche de clients
- Classement des commandes par statut
- Application d’une réduction de 10 %
- Indicateurs sur les délais et la qualité des commandes
- Calcul d’un indice de confiance

### Suivi de l’activité commerciale
- Bilans des ventes annuelles et par catégorie de travaux
- Suivi des commandes terminées par artisan
- Identification des travaux urgents et en retard
- Statistiques sur les retards et la satisfaction des clients

## Technologies utilisées

- **Excel VBA** : interface de gestion, formulaires et automatisation des traitements
- **MySQL / SQL** : création des tables, gestion des données et requêtes
- **Modélisation de données** : conception des modèles conceptuel et logique de données

## Structure du dépôt

```text
artisans-de-confiance/
├── database/
│   ├── 01_creation_tables.sql
│   ├── 02_requetes_gestion_artisans.sql
│   └── mcd/
│       └── artisans_de_confiance.mcd
├── docs/
│   ├── cahier_de_recette_mysql.pdf
│   ├── recherche_documentaire.pdf
│   ├── structuration_donnees_mcd_mld.pdf
│   ├── sujet_projet_2526.pdf
│   └── images/
└── vba/
    ├── tableau_de_bord_artisans.xlsm
    └── modules/
        ├── gestion_artisan_fournisseur/
        ├── gestion_commandes_clients/
        └── gestion_commerciale/
```

Le dossier `database` contient les scripts SQL et le modèle conceptuel de données. Le dossier `vba` comprend le classeur Excel avec macros ainsi que les fichiers exportés des modules et formulaires VBA. Le dossier `docs` regroupe la documentation du projet.

## Installation et utilisation

### Base de données

1. Créer la base de données MySQL :

   ```sql
   CREATE DATABASE artisans_de_confiance;
   ```

2. Exécuter `database/01_creation_tables.sql` pour créer les tables.
3. Exécuter `database/02_requetes_gestion_artisans.sql` pour charger les données de test et créer les vues prévues par le script.

### Tableau de bord Excel

1. Ouvrir `vba/tableau_de_bord_artisans.xlsm` avec Excel.
2. Activer les macros si nécessaire.
3. Utiliser les formulaires et fonctionnalités du classeur.

Les fichiers VBA exportés permettent également de consulter le code indépendamment du classeur Excel. Les fichiers `.frm` et `.frx` associés doivent rester ensemble.

## Documentation et tests

Le dossier `docs` contient les documents de conception et de recherche du projet. Le **cahier de recette MySQL** présente les scénarios de test et les résultats attendus pour les fonctionnalités testées.

---

Projet académique réalisé en groupe à l’EPF.
