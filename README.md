# Artisans de Confiance

Projet transverse **VBA / MySQL** réalisé en groupe (Groupe 3, EPF, promotion 2028).

L'objectif est de fournir un tableau de bord pour l'entreprise *Artisans de Confiance* : suivi des artisans, des fournisseurs, des clients et des commandes, ainsi que le pilotage de l'activité commerciale.

## Fonctionnalités

Le tableau de bord couvre trois grands domaines :

1. **Gestion des artisans et des fournisseurs** : ajout, modification et suppression d'un artisan, fiche artisan et fiche fournisseur, bilan des stocks et seuils critiques, prix moyen par type de travaux, artisans indisponibles.
2. **Gestion des commandes clients** : comptes clients, fiche client avec historique de commandes et factures, recherche de clients, commandes par statut, réduction de 10 %, indicateurs de qualité des délais et des commandes, indice de confiance.
3. **Gestion commerciale** : bilan des ventes annuel et par catégorie de travaux, commandes terminées par artisan, travaux d'urgence, travaux en retard, statistiques sur les retards et la satisfaction.

## Structure du dépôt

```
artisans-de-confiance/
├── database/
│   ├── 01_creation_tables.sql          # création des tables MySQL (à exécuter en premier)
│   ├── 02_requetes_gestion_artisans.sql # données de test + vues (gestion des artisans)
│   └── mcd/
│       └── artisans_de_confiance.mcd   # modèle conceptuel de données (fichier Looping)
├── vba/
│   ├── tableau_de_bord_artisans.xlsm   # classeur Excel final avec les macros
│   └── modules/                        # code VBA exporté (sauvegarde lisible)
│       ├── gestion_commandes_clients/  # formulaires de la partie 3 (2a à 2h)
│       └── gestion_commerciale/        # formulaire et module de la partie 4 (3a à 3f)
└── docs/
    ├── cahier_de_recette_mysql.pdf
    ├── recherche_documentaire.pdf
    ├── structuration_donnees_mcd_mld.pdf
    ├── sujet_projet_2526.pdf
    └── images/                         # MCD, MLD et page de garde
```

## Utilisation

### Base de données

1. Créer la base : `CREATE DATABASE artisans_de_confiance;`
2. Exécuter `database/01_creation_tables.sql` (le script commence par `USE artisans_de_confiance;`).
3. Exécuter `database/02_requetes_gestion_artisans.sql` pour insérer les données de test et créer les vues.

### Tableau de bord Excel

Ouvrir `vba/tableau_de_bord_artisans.xlsm` et activer les macros.

Les dossiers `vba/modules/` contiennent le même code exporté en fichiers texte, ce qui permet de le lire et de suivre ses modifications dans Git. Chaque fichier `.frm` doit rester à côté de son fichier `.frx`.

## Tests

Le cahier de recette (`docs/cahier_de_recette_mysql.pdf`) décrit les scénarios de test et les résultats attendus pour chaque fonctionnalité.
