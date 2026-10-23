-- 1) GESTION DES ARTISANS ET DES FOURNISSEURS

-- AJOUT DES DONNÉES DE TEST
INSERT INTO CATEGORIE_TRAVAUX (__id_categorie, libelle_categorie_travaux, description_categorie_travaux) VALUES
('CAT001', 'Plomberie', 'Travaux de plomberie et sanitaires'),
('CAT002', 'Électricité', 'Installation et réparation électrique'),
('CAT003', 'Maçonnerie', 'Travaux de maçonnerie et construction');

INSERT INTO ARTISAN (__id_artisan, nom_artisan, coordonnees_artisan, telephone_artisan, email_artisan, disponibilite_artisan, __id_categorie) VALUES
('ART001', 'Martin Dupont', '15 Rue de la République, Paris', '0145236789', 'martin.dupont@email.com', 1, 'CAT001'),
('ART002', 'Sophie Bernard', '28 Avenue Victor Hugo, Lyon', '0478563412', 'sophie.bernard@email.com', 0, 'CAT002'),
('ART003', 'Pierre Moreau', '7 Place du Commerce, Marseille', '0491678534', 'pierre.moreau@email.com', 1, 'CAT003'),
('ART004', 'Marie Laurent', '42 Boulevard Saint-Germain, Lille', '0321458796', 'marie.laurent@email.com', 1, 'CAT001'),
('ART005', 'Jean Petit', '63 Rue de la Paix, Bordeaux', '0555123456', 'jean.petit@email.com', 0, 'CAT002');

INSERT INTO PRODUIT (__id_produit, nom_produit, description_produit, prix_unitaire_produit, seuil_critique_produit) VALUES
('PROD001', 'Tuyau PVC 50mm', 'Tuyau pour évacuation eaux', 15.50, 5),
('PROD002', 'Interrupteur simple', 'Interrupteur unipolaire', 8.90, 10),
('PROD003', 'Sac de ciment 25kg', 'Ciment pour maçonnerie', 12.00, 8),
('PROD004', 'Robinet chromé', 'Robinet de lavabo chromé', 45.00, 3);

INSERT INTO FOURNISSEUR (__id_fournisseur, nom_fournisseur, coordonnees_fournisseur, telephone_fournisseur, email_fournisseur) VALUES
('FOUR001', 'BricoPro', 'Zone Industrielle Nord, Paris', '0145123654', 'contact@bricopro.fr'),
('FOUR002', 'ElecPlus', 'Parc d''Activités Sud, Lyon', '0478965412', 'info@elecplus.com'),
('FOUR003', 'Materiaux Express', 'Rue des Artisans, Marseille', '0491632587', 'client@materiaux-express.fr');

INSERT INTO APPROVISIONNE (__id_fournisseur, __id_produit) VALUES
('FOUR001', 'PROD001'),
('FOUR001', 'PROD004'),
('FOUR002', 'PROD002'),
('FOUR003', 'PROD003');

INSERT INTO POSSEDE (__id_produit, __id_artisan, quantite_stock) VALUES
('PROD001', 'ART001', 3),
('PROD004', 'ART001', 5),
('PROD002', 'ART002', 12),
('PROD003', 'ART003', 2),
('PROD001', 'ART004', 4);

INSERT INTO STATUT_COMMANDE (id_statut_commande, libelle_statut_commande) VALUES
('STAT001', 'En attente'),
('STAT002', 'En cours'),
('STAT003', 'Terminé'),
('STAT004', 'Annulé');

-- a. Vues pour gestion des artisans
CREATE OR REPLACE VIEW vue_formulaire_ajout_artisan AS
SELECT 
    'Formulaire Ajout Artisan' as operation,
    'Identifiant unique' as identifiant_artisan,
    'Nom de famille' as nom_artisan,
    'Adresse complète' as adresse_artisan,
    'Numéro de téléphone' as telephone_artisan,
    'Email professionnel' as email_artisan,
    'Disponible (0/1)' as disponibilite_artisan,
    'Catégorie métier' as categorie_metier;

CREATE OR REPLACE VIEW vue_formulaire_modification_artisan AS
SELECT 
    a.__id_artisan as identifiant,
    a.nom_artisan as nom,
    a.coordonnees_artisan as adresse,
    a.telephone_artisan as telephone,
    a.email_artisan as email,
    a.disponibilite_artisan as disponible,
    a.__id_categorie as code_categorie,
    ct.libelle_categorie_travaux as metier
FROM ARTISAN a
INNER JOIN CATEGORIE_TRAVAUX ct ON a.__id_categorie = ct.__id_categorie;

CREATE OR REPLACE VIEW vue_liste_suppression_artisans AS
SELECT 
    a.__id_artisan as identifiant,
    a.nom_artisan as nom,
    a.telephone_artisan as telephone,
    a.email_artisan as email,
    ct.libelle_categorie_travaux as specialite,
    COUNT(c.__id_commande) as commandes_en_cours,
    CASE 
        WHEN COUNT(c.__id_commande) > 0 THEN 'IMPRIMABLE - Commandes en cours'
        ELSE 'SUPPRIMABLE - Aucune commande'
    END as statut
FROM ARTISAN a
INNER JOIN CATEGORIE_TRAVAUX ct ON a.__id_categorie = ct.__id_categorie
LEFT JOIN COMMANDE c ON a.__id_artisan = c.__id_artisan
GROUP BY a.__id_artisan, a.nom_artisan, a.telephone_artisan, a.email_artisan, ct.libelle_categorie_travaux
ORDER BY commandes_en_cours DESC;

-- b. Affichage fiche artisan avec fournisseurs et stock
CREATE OR REPLACE VIEW vue_fiche_artisan_complete AS
SELECT 
    a.__id_artisan,
    a.nom_artisan,
    a.coordonnees_artisan,
    a.telephone_artisan,
    a.email_artisan,
    a.disponibilite_artisan,
    ct.libelle_categorie_travaux,
    p.__id_produit,
    p.nom_produit,
    p.description_produit,
    ps.quantite_stock,
    f.nom_fournisseur,
    f.telephone_fournisseur
FROM ARTISAN a
INNER JOIN CATEGORIE_TRAVAUX ct ON a.__id_categorie = ct.__id_categorie
LEFT JOIN POSSEDE ps ON a.__id_artisan = ps.__id_artisan
LEFT JOIN PRODUIT p ON ps.__id_produit = p.__id_produit
LEFT JOIN APPROVISIONNE app ON p.__id_produit = app.__id_produit
LEFT JOIN FOURNISSEUR f ON app.__id_fournisseur = f.__id_fournisseur
ORDER BY a.nom_artisan, p.nom_produit;

-- c. Affichage fiche fournisseur avec historique
CREATE OR REPLACE VIEW vue_fiche_fournisseur_historique AS
SELECT 
    f.__id_fournisseur,
    f.nom_fournisseur,
    f.coordonnees_fournisseur,
    f.telephone_fournisseur,
    f.email_fournisseur,
    p.__id_produit,
    p.nom_produit,
    p.description_produit,
    p.prix_unitaire_produit,
    COUNT(DISTINCT a.__id_artisan) as nombre_artisans_approvisionnes
FROM FOURNISSEUR f
INNER JOIN APPROVISIONNE app ON f.__id_fournisseur = app.__id_fournisseur
INNER JOIN PRODUIT p ON app.__id_produit = p.__id_produit
LEFT JOIN POSSEDE ps ON p.__id_produit = ps.__id_produit
LEFT JOIN ARTISAN a ON ps.__id_artisan = a.__id_artisan
GROUP BY f.__id_fournisseur, p.__id_produit
ORDER BY f.nom_fournisseur, p.nom_produit;

-- d. Bilan des stocks critiques
CREATE OR REPLACE VIEW vue_bilan_stocks_critiques AS
SELECT 
    p.__id_produit,
    p.nom_produit,
    p.description_produit,
    p.prix_unitaire_produit,
    p.seuil_critique_produit,
    SUM(ps.quantite_stock) as stock_total,
    CASE 
        WHEN SUM(ps.quantite_stock) <= p.seuil_critique_produit THEN 'CRITIQUE'
        WHEN SUM(ps.quantite_stock) <= p.seuil_critique_produit * 2 THEN 'FAIBLE'
        ELSE 'SUFFISANT'
    END as etat_stock,
    GROUP_CONCAT(DISTINCT a.nom_artisan) as artisans_concernes
FROM PRODUIT p
LEFT JOIN POSSEDE ps ON p.__id_produit = ps.__id_produit
LEFT JOIN ARTISAN a ON ps.__id_artisan = a.__id_artisan
GROUP BY p.__id_produit
HAVING stock_total <= p.seuil_critique_produit * 2
ORDER BY etat_stock, stock_total ASC;

-- e. Prix moyen par type de travaux
CREATE OR REPLACE VIEW vue_prix_moyen_par_travaux AS
SELECT 
    ct.__id_categorie,
    ct.libelle_categorie_travaux,
    COUNT(DISTINCT c.__id_commande) as nombre_commandes,
    ROUND(AVG(c.montant_total_commande), 2) as prix_moyen,
    MIN(c.montant_total_commande) as prix_minimum,
    MAX(c.montant_total_commande) as prix_maximum,
    COUNT(DISTINCT a.__id_artisan) as nombre_artisans
FROM CATEGORIE_TRAVAUX ct
INNER JOIN ARTISAN a ON ct.__id_categorie = a.__id_categorie
INNER JOIN COMMANDE c ON a.__id_artisan = c.__id_artisan
WHERE a.disponibilite_artisan = 1
GROUP BY ct.__id_categorie, ct.libelle_categorie_travaux
ORDER BY prix_moyen DESC;

-- f. Artisans non disponibles
CREATE OR REPLACE VIEW vue_artisans_non_disponibles AS
SELECT 
    a.__id_artisan,
    a.nom_artisan,
    a.telephone_artisan,
    a.email_artisan,
    ct.libelle_categorie_travaux,
    COUNT(c.__id_commande) as nombre_commandes_en_cours,
    GROUP_CONCAT(DISTINCT c.__id_commande) as commandes_affectees
FROM ARTISAN a
INNER JOIN CATEGORIE_TRAVAUX ct ON a.__id_categorie = ct.__id_categorie
INNER JOIN COMMANDE c ON a.__id_artisan = c.__id_artisan
INNER JOIN A_POUR ap ON c.__id_commande = ap.__id_commande
INNER JOIN STATUT_COMMANDE sc ON ap.id_statut_commande = sc.id_statut_commande
WHERE a.disponibilite_artisan = 0 
   OR sc.libelle_statut_commande IN ('En cours', 'En attente', 'Planifié')
GROUP BY a.__id_artisan, a.nom_artisan, a.telephone_artisan, a.email_artisan, ct.libelle_categorie_travaux
ORDER BY nombre_commandes_en_cours DESC;