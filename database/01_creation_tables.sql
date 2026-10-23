-- ----------------------------------------------------------
-- Script MYSQL pour mcd 
-- ----------------------------------------------------------
USE artisans_de_confiance;

-- ----------------------------
-- Table: MOTIF_RETARD
-- ----------------------------
CREATE TABLE MOTIF_RETARD (
  id_motif VARCHAR(50) NOT NULL,
  libelle_motif VARCHAR(50) NOT NULL,
  description_motif VARCHAR(50) NOT NULL,
  CONSTRAINT MOTIF_RETARD_PK PRIMARY KEY (id_motif)
)ENGINE=InnoDB;


-- ----------------------------
-- Table: STATUT_COMMANDE
-- ----------------------------
CREATE TABLE STATUT_COMMANDE (
  id_statut_commande VARCHAR(50) NOT NULL,
  libelle_statut_commande VARCHAR(50) NOT NULL,
  CONSTRAINT STATUT_COMMANDE_PK PRIMARY KEY (id_statut_commande)
)ENGINE=InnoDB;


-- ----------------------------
-- Table: CLIENT
-- ----------------------------
CREATE TABLE CLIENT (
  __id_client VARCHAR(50) NOT NULL,
  nom_client VARCHAR(50) NOT NULL,
  prenom_client VARCHAR(50) NOT NULL,
  adresse_client VARCHAR(50) NOT NULL,
  code_postal_client VARCHAR(50) NOT NULL,
  telephone_client VARCHAR(50) NOT NULL,
  email_client VARCHAR(50) NOT NULL,
  compte_fidelite_client TINYINT(1) NOT NULL,
  CONSTRAINT CLIENT_PK PRIMARY KEY (__id_client)
)ENGINE=InnoDB;


-- ----------------------------
-- Table: PRODUIT
-- ----------------------------
CREATE TABLE PRODUIT (
  __id_produit VARCHAR(50) NOT NULL,
  nom_produit VARCHAR(50) NOT NULL,
  description_produit VARCHAR(50) NOT NULL,
  prix_unitaire_produit FLOAT NOT NULL,
  seuil_critique_produit INT NOT NULL,
  CONSTRAINT PRODUIT_PK PRIMARY KEY (__id_produit)
)ENGINE=InnoDB;


-- ----------------------------
-- Table: FOURNISSEUR
-- ----------------------------
CREATE TABLE FOURNISSEUR (
  __id_fournisseur VARCHAR(50) NOT NULL,
  nom_fournisseur VARCHAR(50) NOT NULL,
  coordonnees_fournisseur VARCHAR(50) NOT NULL,
  telephone_fournisseur VARCHAR(50) NOT NULL,
  email_fournisseur VARCHAR(50) NOT NULL,
  CONSTRAINT FOURNISSEUajout_artisanexempleR_PK PRIMARY KEY (__id_fournisseur)
)ENGINE=InnoDB;


-- ----------------------------
-- Table: CATEGORIE_TRAVAUX
-- ----------------------------
CREATE TABLE CATEGORIE_TRAVAUX (
  __id_categorie VARCHAR(50) NOT NULL,
  libelle_categorie_travaux VARCHAR(50) NOT NULL,
  description_categorie_travaux VARCHAR(50) NOT NULL,
  CONSTRAINT CATEGORIE_TRAVAUX_PK PRIMARY KEY (__id_categorie)
)ENGINE=InnoDB;


-- ----------------------------
-- Table: APPROVISIONNE
-- ----------------------------
CREATE TABLE APPROVISIONNE (
  __id_fournisseur VARCHAR(50) NOT NULL,
  __id_produit VARCHAR(50) NOT NULL,
  CONSTRAINT APPROVISIONNE_PK PRIMARY KEY (__id_fournisseur, __id_produit)
)ENGINE=InnoDB;


-- ----------------------------
-- Table: COMMANDE
-- ----------------------------
CREATE TABLE COMMANDE (
  __id_commande VARCHAR(50) NOT NULL,
  date_creation_commande DATE NOT NULL,
  date_debut_prevue_commande DATE NOT NULL,
  date_fin_prevue_commande DATE NOT NULL,
  date_debut_reelle_commande DATE NOT NULL,
  date_fin_reelle_commande DATE NOT NULL,
  montant_total_commande FLOAT NOT NULL,
  reduction_commande TINYINT(1) NOT NULL,
  __id_client VARCHAR(50) NOT NULL,
  __id_artisan VARCHAR(50) NOT NULL,
  __id_facture VARCHAR(50) NOT NULL,
  __id_categorie VARCHAR(50) NOT NULL,
  __id_avis VARCHAR(50) NOT NULL,
  CONSTRAINT COMMANDE_PK PRIMARY KEY (__id_commande)
)ENGINE=InnoDB;


-- ----------------------------
-- Table: AVIS
-- ----------------------------
CREATE TABLE AVIS (
  __id_avis VARCHAR(50) NOT NULL,
  note_avis FLOAT NOT NULL,
  commentaire_avis VARCHAR(50) NOT NULL,
  date_avis DATE NOT NULL,
  note_finition_avis FLOAT NOT NULL,
  note_ponctualite_avis FLOAT NOT NULL,
  __id_commande VARCHAR(50) NOT NULL,
  __id_artisan VARCHAR(50) NOT NULL,
  CONSTRAINT AVIS_PK PRIMARY KEY (__id_avis)
)ENGINE=InnoDB;


-- ----------------------------
-- Table: A_POUR_MOTIF
-- ----------------------------
CREATE TABLE A_POUR_MOTIF (
  id_motif VARCHAR(50) NOT NULL,
  __id_commande VARCHAR(50) NOT NULL,
  CONSTRAINT A_POUR_MOTIF_PK PRIMARY KEY (id_motif, __id_commande),
  CONSTRAINT A_POUR_MOTIF_id_motif_FK FOREIGN KEY (id_motif) REFERENCES MOTIF_RETARD (id_motif)
)ENGINE=InnoDB;


-- ----------------------------
-- Table: FACTURE
-- ----------------------------
CREATE TABLE FACTURE (
  __id_facture VARCHAR(50) NOT NULL,
  date_emission_facture DATE NOT NULL,
  montant_facture FLOAT NOT NULL,
  mode_paiement_facture VARCHAR(50) NOT NULL,
  statut_paiement_facture VARCHAR(50) NOT NULL,
  __id_commande VARCHAR(50) NOT NULL,
  CONSTRAINT FACTURE_PK PRIMARY KEY (__id_facture)
)ENGINE=InnoDB;


-- ----------------------------
-- Table: Acheter
-- ----------------------------
CREATE TABLE Acheter (
  __id_produit VARCHAR(50) NOT NULL,
  __id_artisan VARCHAR(50) NOT NULL,
  Quantite_produit FLOAT NOT NULL,
  CONSTRAINT Acheter_PK PRIMARY KEY (__id_produit, __id_artisan)
)ENGINE=InnoDB;


-- ----------------------------
-- Table: A_POUR
-- ----------------------------
CREATE TABLE A_POUR (
  id_statut_commande VARCHAR(50) NOT NULL,
  __id_commande VARCHAR(50) NOT NULL,
  CONSTRAINT A_POUR_PK PRIMARY KEY (id_statut_commande, __id_commande),
  CONSTRAINT A_POUR_id_statut_commande_FK FOREIGN KEY (id_statut_commande) REFERENCES STATUT_COMMANDE (id_statut_commande)
)ENGINE=InnoDB;


-- ----------------------------
-- Table: ARTISAN
-- ----------------------------
CREATE TABLE ARTISAN (
  __id_artisan VARCHAR(50) NOT NULL,
  nom_artisan VARCHAR(50) NOT NULL,
  coordonnees_artisan VARCHAR(50) NOT NULL,
  telephone_artisan VARCHAR(50) NOT NULL,
  email_artisan VARCHAR(50) NOT NULL,
  disponibilite_artisan TINYINT(1) NOT NULL,
  __id_categorie VARCHAR(50) NOT NULL,
  CONSTRAINT ARTISAN_PK PRIMARY KEY (__id_artisan)
)ENGINE=InnoDB;


-- ----------------------------
-- Table: POSSEDE
-- ----------------------------
CREATE TABLE POSSEDE (
  __id_produit VARCHAR(50) NOT NULL,
  __id_artisan VARCHAR(50) NOT NULL,
  quantite_stock FLOAT NOT NULL,
  CONSTRAINT POSSEDE_PK PRIMARY KEY (__id_produit, __id_artisan),
  CONSTRAINT POSSEDE___id_artisan_FK FOREIGN KEY (__id_artisan) REFERENCES ARTISAN (__id_artisan)
)ENGINE=InnoDB;


-- ----------------------------
-- Table: COMPTE_CLIENT
-- ----------------------------
CREATE TABLE COMPTE_CLIENT (
  id_compte_client VARCHAR(50) NOT NULL,
  login_compte_client VARCHAR(50) NOT NULL,
  mdp_compte_client VARCHAR(50) NOT NULL,
  date_creation_compte_client DATE NOT NULL,
  __id_client VARCHAR(50) NOT NULL,
  CONSTRAINT COMPTE_CLIENT_PK PRIMARY KEY (id_compte_client)
)ENGINE=InnoDB;


-- ===== FOREIGN KEYS =====

ALTER TABLE APPROVISIONNE
  ADD CONSTRAINT APPROVISIONNE___id_fournisseur_FK FOREIGN KEY (__id_fournisseur)
  REFERENCES FOURNISSEUR  (__id_fournisseur);

ALTER TABLE APPROVISIONNE
  ADD CONSTRAINT APPROVISIONNE___id_produit_FK FOREIGN KEY (__id_produit)
  REFERENCES PRODUIT  (__id_produit);

ALTER TABLE COMMANDE
  ADD CONSTRAINT COMMANDE___id_client_FK FOREIGN KEY (__id_client)
  REFERENCES CLIENT  (__id_client);

ALTER TABLE COMMANDE
  ADD CONSTRAINT COMMANDE___id_artisan_FK FOREIGN KEY (__id_artisan)
  REFERENCES ARTISAN (__id_artisan);

ALTER TABLE COMMANDE
  ADD CONSTRAINT COMMANDE___id_facture_FK FOREIGN KEY (__id_facture)
  REFERENCES FACTURE  (__id_facture);

ALTER TABLE COMMANDE
  ADD CONSTRAINT COMMANDE___id_categorie_FK FOREIGN KEY (__id_categorie)
  REFERENCES CATEGORIE_TRAVAUX  (__id_categorie);

ALTER TABLE COMMANDE
  ADD CONSTRAINT COMMANDE___id_avis_FK FOREIGN KEY (__id_avis)
  REFERENCES AVIS  (__id_avis);

ALTER TABLE AVIS
  ADD CONSTRAINT AVIS___id_commande_FK FOREIGN KEY (__id_commande)
  REFERENCES COMMANDE  (__id_commande);

ALTER TABLE AVIS
  ADD CONSTRAINT AVIS___id_artisan_FK FOREIGN KEY (__id_artisan)
  REFERENCES ARTISAN (__id_artisan);

ALTER TABLE A_POUR_MOTIF
  ADD CONSTRAINT A_POUR_MOTIF___id_commande_FK FOREIGN KEY (__id_commande)
  REFERENCES COMMANDE  (__id_commande);

ALTER TABLE FACTURE
  ADD CONSTRAINT FACTURE___id_commande_FK FOREIGN KEY (__id_commande)
  REFERENCES COMMANDE  (__id_commande);

ALTER TABLE Acheter
  ADD CONSTRAINT Acheter___id_produit_FK FOREIGN KEY (__id_produit)
  REFERENCES PRODUIT  (__id_produit);

ALTER TABLE Acheter
  ADD CONSTRAINT Acheter___id_artisan_FK FOREIGN KEY (__id_artisan)
  REFERENCES ARTISAN (__id_artisan);

ALTER TABLE A_POUR
  ADD CONSTRAINT A_POUR___id_commande_FK FOREIGN KEY (__id_commande)
  REFERENCES COMMANDE  (__id_commande);

ALTER TABLE ARTISAN
  ADD CONSTRAINT ARTISAN___id_categorie_FK FOREIGN KEY (__id_categorie)
  REFERENCES CATEGORIE_TRAVAUX  (__id_categorie);

ALTER TABLE POSSEDE
  ADD CONSTRAINT POSSEDE___id_produit_FK FOREIGN KEY (__id_produit)
  REFERENCES PRODUIT  (__id_produit);

ALTER TABLE COMPTE_CLIENT
  ADD CONSTRAINT COMPTE_CLIENT___id_client_FK FOREIGN KEY (__id_client)
  REFERENCES CLIENT  (__id_client);
