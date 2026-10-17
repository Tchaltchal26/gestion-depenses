# 01 — Architecture technique et stratégie

Ce document répond aux points 1 à 9 de la « première réponse attendue » du cahier des charges.
Il distingue en permanence **ce qui est livré et testé**, **ce qui est livré sans test automatique**
et **ce qui est conçu mais pas encore développé** (voir aussi `05-etat-avancement-et-limites.md`).

---

## 1. Architecture technique recommandée

### 1.1 Principe : une couche d'accès aux données remplaçable

```
 ┌──────────────────────────── Excel (.xlsm) ────────────────────────────┐
 │  Interface (modUI)      feuilles-écrans, boutons, formulaire connexion │
 │        │                                                              │
 │  Règles métier          modEtablissement  modAnnee  modStructure      │
 │        │                modEleves  modInscriptions  modRapports       │
 │        │                                                              │
 │  Services transverses   modSecurite (droits)  modAudit (journal)      │
 │        │                modLicence (+ modRSA, modSHA256)              │
 │        │                modSauvegarde (copies vérifiées, export CSV)  │
 │        ▼                                                              │
 │  Accès aux données      modDAL  ◄── seul module qui lit/écrit les T_* │
 │        │                                                              │
 │  Stockage (prototype)   feuilles T_* masquées (une table par feuille) │
 └────────┼───────────────────────────────────────────────────────────────┘
          ▼  (évolution prévue : même interface DAL_*)
   Base relationnelle (Access/ACE, SQL Server Express, PostgreSQL)
```

| Couche | Rôle | État |
|---|---|---|
| Interface | Écrans sur feuilles nommées, boutons, formulaire de connexion masqué, export PDF | Livrée, **non testée automatiquement** (Excel requis) |
| Règles métier | Validations, cycles de vie, cloisonnement, numérotation | Livrée et **testée** |
| Droits | Rôles, habilitations par établissement, niveaux 0-4 | Livrée et **testée** |
| Audit | Journal en ajout seul chaîné SHA-256, ancre externe | Livrée et **testée** |
| Licences | Vérification RSA-SHA256, empreinte, expiration, révocation | Livrée et **testée** |
| Données | Feuilles `T_*` adressées par identifiant et nom de colonne | Livrée et **testée** |

### 1.2 Choix technologiques

- **Excel de bureau + VBA** : interface, automatisations, documents. Le code métier évite toute
  dépendance Windows (pas de `Scripting.Dictionary`, pas d'API `Declare`) : il fonctionne sous
  Excel Windows, Excel Mac et **LibreOffice** (mode VBA), ce qui permet les tests automatisés.
- **Cryptographie en VBA pur** : SHA-256 (FIPS 180-4) et vérification RSA PKCS#1 v1.5, validées
  contre les vecteurs officiels et contre la bibliothèque Python `cryptography`.
- **Outil éditeur en Python** (`tools/licence_admin/sg_licence.py`) : génération des clés,
  signature des licences, registre commercial SQLite. **La clé privée n'existe que là.**
- **Power Query / Power Pivot** : non utilisés dans le prototype ; prévus pour les imports
  (aperçu, nettoyage) et les tableaux de bord consolidés multi-établissements.

### 1.3 Pourquoi pas un classeur partagé comme base de données

Un classeur Excel n'offre ni transactions, ni verrouillage d'enregistrement, ni contrôle d'accès
au niveau du stockage. Le prototype est donc **mono-poste / mono-utilisateur simultané**.
Pour plusieurs postes en écriture simultanée, la cible est :

| Besoin | Recommandation |
|---|---|
| 1 poste, 1 établissement | Prototype actuel (feuilles `T_*`) |
| 2-5 postes, réseau local | Base Access (.accdb) sur partage réseau via ADO — remplacement de `modDAL` |
| Réseau d'établissements, accès distant | SQL Server Express / PostgreSQL + API ; Excel devient un client |

L'interface `DAL_*` (Insert, Update, Get, Query, NewId) a été conçue pour ce remplacement :
aucun module métier n'accède aux cellules.

### 1.4 Compatibilité

| Environnement | Statut |
|---|---|
| Excel 2016/2019/2021/365 Windows (32 et 64 bits) | Cible principale. Aucun `Declare` : pas de problème 32/64 bits |
| Excel 365 Mac | Métier compatible ; `Scripting.FileSystemObject` absent (empreinte réduite) ; installateur PowerShell indisponible (import manuel) |
| Excel pour le Web | **Incompatible** : VBA ne s'y exécute pas. Consultation seule des feuilles |
| LibreOffice 7.x/24.x | Utilisé pour les tests du code métier. Pas une cible de production |

---

## 2. Modèle de données

Détaillé dans `02-modele-donnees.md` (22 tables, relations, règles). Principes :

- chaque table a un **identifiant technique permanent** en 1re colonne :
  `PREFIXE-INSTALLATION-SEQUENCE` (ex. `ELV-3F2A-000123`), unique même en cas de fusion de postes ;
- les données métier cloisonnées portent un **`EtabID`** ;
- **aucune suppression** : statuts (`ACTIVE`, `TERMINEE`, `ANNULEE`), dates de fin, historiques ;
- toutes les valeurs sont stockées **en texte** (dates ISO `AAAA-MM-JJ`) pour éviter les
  conversions automatiques d'Excel (matricules, zéros de tête, formules injectées).

---

## 3. Modules et dépendances

```
modUtil, modRec, modSHA256, modRSA            (aucune dépendance)
   └─ modSchema ─ modDAL
         └─ modAudit ◄──► modSecurite ◄── modLicence
               └─ modEtablissement ─ modAnnee ─ modStructure
                     └─ modEleves ─ modInscriptions ─ modRapports
                           └─ modSauvegarde ─ modSetup ─ modDemo ─ modUI
```

| Module | Contenu | Prochaines extensions |
|---|---|---|
| Établissements | création, paramètres, numérotations, appartenance | logo, modèles de documents |
| Années | préparation → ouverte → clôturée → réouverture motivée | clôture avec bascule automatique |
| Structure | cycles, niveaux (ordre = passages), filières, classes par année, reconduction | matières, coefficients, salles, groupes |
| Élèves | dossier permanent, doublons, recherche, responsables, liens multi-enfants | photo, documents, autorisations |
| Inscriptions | 1re inscription, réinscription (passage/redoublement déduit), transfert entrant, changement de classe historisé, suspension, réintégration, départ, radiation, fin de cycle, annulation, transfert inter-établissements | — |
| Rapports | indicateurs, effectifs par classe, contrôle qualité, certificat numéroté, liste de classe | bulletins, états financiers |
| À développer | Présences, Notes & examens, Finances, Personnel, Emplois du temps, Stocks, modules complémentaires | voir plan § 9 |

---

## 4. Sécurité et séparation des établissements

### 4.1 Contrôles réalisés dans le code (testés)

- **Point de contrôle unique** `Sec_Exiger(etab, module, niveau)` appelé par toutes les opérations.
- **Habilitations par établissement** : un rôle n'est valable que pour l'établissement
  et la période indiqués.
- **Contrôle d'appartenance** `Etab_ExigerAppartenance` : toute référence (élève, classe, année,
  responsable, inscription) doit appartenir à l'établissement de l'opération ; les tentatives
  sont refusées et **journalisées** (`TENTATIVE_INTER_ETAB`).
- **Recherche, listes, exports** toujours filtrés par `EtabID`.
- **Administrateur technique sans accès aux dossiers élèves** (le propriétaire gère licences,
  comptes et sauvegardes sans voir les données scolaires).
- Mots de passe salés et hachés (SHA-256 itéré), verrouillage après 5 échecs, réinitialisation
  par un administrateur avec changement obligatoire, désactivation des comptes sortants.

### 4.2 Ce que ces contrôles ne garantissent pas

Dans un fichier Excel, **toute personne qui possède le fichier peut lire les feuilles** avec un
autre outil (LibreOffice, extraction du ZIP, script). La protection des feuilles et du projet VBA
n'est **pas** une authentification. Conséquences :

- en mode **indépendant**, un fichier = un établissement : la fuite inter-établissements est
  impossible par construction ;
- en mode **centralisé**, le fichier contient plusieurs établissements : il ne doit être confié
  qu'à des personnes autorisées pour **tous** ces établissements, ou bien il faut passer à une
  base de données serveur où les droits sont appliqués **au stockage** ;
- le chiffrement du fichier par Excel (« Chiffrer avec mot de passe », AES) est recommandé pour
  les postes portables ; il protège le fichier au repos, pas contre un utilisateur légitime.

---

## 5. Licences commerciales

Résumé (détails dans `03-licences.md`) :

1. Le client génère depuis l'application une **demande d'activation** `.sgreq`
   (identifiant d'installation + empreinte du poste).
2. L'éditeur émet avec `sg_licence.py` un **fichier de licence signé** `.sglic`
   (RSA 3072 bits, SHA-256) indiquant titulaire, établissements, modules, type, dates,
   version couverte, nombre d'installations, période de grâce.
3. Le classeur vérifie la signature avec la **clé publique intégrée**, l'empreinte, la
   révocation, la version, la date. Il débloque uniquement les modules achetés.
4. En cas d'expiration, révocation, copie ou horloge incohérente : **mode consultation**.
   Lecture, export et sauvegarde restent possibles. **Aucune donnée n'est supprimée.**

Types gérés : annuelle, mensuelle, perpétuelle (non désactivée par la fin de maintenance),
essai, démonstration, multi-établissements, réseau ; essai intégré de 30 jours sans licence.

## 6. Protections contre la copie non autorisée

| Mesure | Effet | Limite |
|---|---|---|
| Signature RSA de la licence | Impossible de fabriquer ou modifier une licence sans la clé privée | Le code VBA peut être modifié pour ignorer la vérification |
| Empreinte d'installation | Copie sur un autre poste → mode consultation | Copie sur le **même** poste non détectable hors ligne |
| Date maximale vue | Recul d'horloge détecté | Contournable en modifiant les données système |
| Liste de révocation signée | Révocation d'une installation compromise | Doit être importée (hors ligne) |
| Transfert officiel | Nouvelle licence pour le nouveau poste, ancienne révoquée | Procédure manuelle |
| Protection du projet VBA | Dissuade la consultation ordinaire | **Contournable** avec des outils publics |

**Aucune de ces mesures n'est inviolable.** Elles visent l'usage ordinaire et la copie simple.
Une protection plus forte exige une activation en ligne (serveur de licences, comptage des
installations simultanées) et la signature numérique du projet VBA (certificat de code) ;
elles sont conçues (`03-licences.md` § 6) mais non développées.

## 7. Limites techniques d'Excel

| Limite | Conséquence | Réponse |
|---|---|---|
| Pas de transactions | Une erreur au milieu d'une opération pourrait laisser une écriture partielle | Toutes les validations sont faites **avant** la première écriture ; opérations regroupées sous une référence `TxRef` |
| Pas d'accès concurrent fiable | Un seul utilisateur en écriture | Base serveur pour le multiposte |
| Sécurité au niveau du fichier | Lecture possible hors application | Mode indépendant, chiffrement du fichier, base serveur |
| Volume | Performances correctes jusqu'à quelques dizaines de milliers de lignes par table ; recherches linéaires | Archivage par année, migration SQL au-delà |
| VBA non disponible sur le Web / mobile | Pas d'accès distant | Client web futur sur base serveur |
| Journal modifiable par un expert | Chaîne recalculable | Ancre externe, sauvegardes indépendantes, base avec journal protégé |

## 8. Sauvegarde et archivage

- **Sauvegarde vérifiée** (`Sauv_Creer`) : copie horodatée du classeur, **réouverture en lecture
  seule** et recontrôle complet du journal ; l'empreinte de tête doit être identique. Une
  sauvegarde non vérifiée n'est pas comptée comme valide.
- **Clôture annuelle** impossible sans sauvegarde vérifiée du jour.
- **Ancre externe** : l'empreinte du journal est communiquée hors du classeur (impression,
  courriel à la direction) ; elle permet de détecter un journal reconstruit.
- **Export cloisonné CSV UTF-8** par établissement, toujours disponible (même licence expirée) :
  restitution des données au client, migration, archivage dans un format documenté.
- **Règle 3-2-1 recommandée** : 3 copies, 2 supports, 1 hors site ; ne jamais dépendre d'un seul
  ordinateur, d'une seule clé USB ou d'un seul compte cloud.
- **Restauration** : procédure dans `06-guide-installation.md` § 5 (vérifier la copie, puis la
  substituer au fichier courant ; l'opération est journalisée dans la copie restaurée).
- **Années scolaires** : une nouvelle année crée de nouveaux enregistrements ; les années
  clôturées restent consultables dans le même fichier. Archivage d'années anciennes vers un
  fichier séparé : prévu (phase 4).

## 9. Plan de développement

| Phase | Contenu | État |
|---|---|---|
| 1. Cadrage | Questions `04-questions-cadrage.md`, hypothèses par défaut | Hypothèses posées, **réponses attendues** |
| 2. Conception | Architecture, modèle de données, droits, licences, archivage | **Fait** (ce dossier) |
| 3. Prototype | Établissements, utilisateurs, années, structure, élèves, responsables, inscriptions, audit, licences, sauvegarde, contrôle qualité, tableau de bord, certificat | **Fait et testé** (métier) ; interface à valider sous Excel |
| 4a. Pédagogie | Matières/coefficients, évaluations, notes (absence ≠ zéro), compétences, bulletins, délibérations, verrouillage après validation | À faire |
| 4b. Finances | Tarifs par niveau/année, factures, échéanciers, paiements partiels, reçus numérotés, annulations contrôlées, caisse, clôtures | À faire |
| 4c. Vie scolaire | Présences/absences, discipline | À faire |
| 4d. Personnel, EDT, stocks | | À faire |
| 5. Sécurité commerciale | Serveur d'activation en ligne, signature du code, mises à jour signées | Conçu, à faire |
| 6. Validation | Tests sous Excel Windows/Mac, tests de charge, test de restauration | Partiel (métier testé sous LibreOffice) |
| 7. Livraison | Manuels, formation, procédure de mise à jour | Guides initiaux livrés |
