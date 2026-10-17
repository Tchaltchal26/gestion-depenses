# 06 — Guides d'installation, de création d'établissement, d'année scolaire et de maintenance

> Le classeur est construit à partir des sources (`src/vba`). Il n'est pas versionné en binaire :
> c'est ce qui garantit que le code livré est exactement le code relu et testé.

## 1. Installation (Excel Windows)

### 1.1 Préparation éditeur (une seule fois)

```
cd tools/licence_admin
python3 -m pip install cryptography
python3 sg_licence.py cles --dossier ~/SG-cles          # phrase secrète demandée
python3 sg_licence.py integrer-cle --publique ~/SG-cles/cle_publique.pem --vba ../../src/vba/modLicence.bas
```

Conserver `cle_privee.pem` et la phrase secrète **hors ligne, en deux exemplaires** (coffre,
support chiffré). Sa perte empêche d'émettre de nouvelles licences ; elle n'affecte jamais les
données des clients.

Sans clé intégrée, le classeur fonctionne en **essai intégré** (30 jours, 2 établissements,
100 élèves) — c'est le mode utilisé pour la démonstration.

### 1.2 Construction du classeur

1. Excel : *Fichier > Options > Centre de gestion de la confidentialité > Paramètres du Centre >
   Paramètres des macros* → cocher **Accès approuvé au modèle d'objet du projet VBA**.
2. PowerShell :
   ```
   powershell -ExecutionPolicy Bypass -File installer\Installer-SG.ps1 -Destination C:\SG\SG_Application.xlsm -Mode INDEPENDANT
   ```
   Ajouter `-Demo` pour charger le jeu fictif, `-Mode CENTRALISE` pour un groupe scolaire.
3. Saisir l'identifiant, le nom et le mot de passe de l'**administrateur technique**.
4. Décocher l'option du point 1.
5. Recommandé : *Fichier > Informations > Protéger le classeur > Chiffrer avec mot de passe*,
   et protéger le projet VBA (*Outils > Propriétés de VBAProject > Protection*).

### 1.3 Installation manuelle (Mac ou sans PowerShell)

1. Nouveau classeur → *Enregistrer sous* `.xlsm`.
2. Éditeur VBA (Alt+F11 / Option+F11) → *Fichier > Importer* chaque `src/vba/*.bas`
   (sauf `modTests.bas` en production). Les fichiers sont en UTF-8 : sous Windows, l'installateur
   les convertit ; en import manuel les accents des **commentaires** peuvent être altérés, sans
   incidence sur le fonctionnement (le code n'utilise pas d'accents hors tests).
3. Copier le contenu de `src/vba/ThisWorkbook.code` dans le module `ThisWorkbook`.
4. Exécuter la macro `SG_Installer` (mode `INDEPENDANT` par défaut).

## 2. Création d'un établissement

Compte : administrateur central (ou administrateur technique initial).

1. En mode indépendant, un seul établissement est autorisé par installation.
2. Appeler `Etab_Creer` (via l'écran Paramétrage — à venir — ou la fenêtre Exécution) :
   code (2-12 caractères, sans espace, identique à celui de la licence), nom, type
   (`MATERNELLE`, `PRIMAIRE`, `COLLEGE`, `LYCEE`, `LYCEE_TECHNIQUE`, `CENTRE_FORMATION`, `COMPLEXE`),
   statut (`PUBLIC`, `PRIVE`, `CONFESSIONNEL`), pays, ville, devise.
3. Habiliter les utilisateurs : `Sec_CreerUtilisateur` puis `Sec_Habiliter userId, etabId, "SECRETAIRE"`.
4. Ajuster les paramètres : `Param_Set etabId, "FORMAT_MATRICULE", "{ETAB}-{AAAA}-{N5}", "motif"`.

## 3. Configuration des cycles, niveaux et filières

```
cyc = Cycle_Creer(etab, "COL", "College", "NOTES", 3)          ' COMPETENCES / NOTES / MIXTE / MODULAIRE
n6  = Niveau_Creer(etab, cyc, "6E", "Sixieme", 1)              ' l'ORDRE détermine passage / redoublement
n5  = Niveau_Creer(etab, cyc, "5E", "Cinquieme", 2)
fil = Filiere_Creer(etab, "D", "Serie scientifique", "SERIE")  ' SERIE / FILIERE / SPECIALITE / OPTION / SECTION
cls = Classe_Creer(etab, annee, n6, "", "6EA", "6e A", 45, "Salle 12")
```

Un complexe scolaire déclare plusieurs cycles dans le même établissement (maternelle par
compétences, collège sur 20…). Les ordres de niveaux sont uniques et continus sur tout
l'établissement.

## 4. Ouverture et clôture annuelle

1. **Créer** l'année suivante (`Annee_Creer`, statut `PREPARATION`) pendant que l'année courante
   est ouverte ; **reconduire** les classes (`Classe_ReconduireAnnee`) ; ajuster les capacités.
2. **Réinscrire** (`Inscr_Inscrire ..., "REINSCRIPTION"`) : la décision (passage, redoublement,
   orientation) est déduite des niveaux et enregistrée ; la date d'effet est le 1er jour de la
   nouvelle année.
3. En fin d'année : **Sauvegarder** (sauvegarde vérifiée obligatoire), puis
   `Annee_Cloturer etab, annee, "motif"`. Les données de l'année deviennent non modifiables.
4. **Ouvrir** la nouvelle année (`Annee_Ouvrir`).
5. Réouverture exceptionnelle : `Annee_Reouvrir` (niveau administration, motif détaillé ≥ 10
   caractères, aucune autre année ouverte), journalisée.

## 5. Sauvegarde et restauration

- **ACCUEIL → Sauvegarder** : choisir un dossier **hors du poste** (clé, NAS, cloud). Le message
  final indique le fichier et l'**ancre** (empreinte du journal) : la noter ou l'envoyer à la
  direction.
- **Restauration** :
  1. ne pas écraser le fichier courant : le renommer `..._incident_AAAAMMJJ.xlsm` ;
  2. ouvrir la copie choisie, **JOURNAL → Vérifier l'intégrité** en saisissant l'ancre attendue ;
  3. si « Journal intègre », l'enregistrer sous le nom de production ;
  4. les saisies postérieures à la sauvegarde sont à ressaisir (consulter le journal du fichier
     incident pour les retrouver).
- **Test de restauration** : à faire au moins une fois par trimestre sur un poste distinct.

## 6. Renouvellement et transfert de licence

- **Renouvellement** : l'éditeur émet une licence pour la même empreinte ; le client l'installe.
- **Changement d'ordinateur / panne** : le client crée une demande depuis le nouveau poste ;
  l'éditeur exécute `transferer` (mêmes droits, ancienne licence révoquée) ; les données sont
  copiées avec le classeur. Pendant la procédure, l'ancien fichier reste consultable.

## 7. Mise à jour et migration

1. Sauvegarde vérifiée.
2. Construire le nouveau classeur avec la nouvelle version des sources **ou** remplacer les
   modules dans le classeur existant (supprimer puis importer chaque `.bas`). Les feuilles `T_*`
   sont conservées.
3. Exécuter `SG_InitialiserBase` : les tables manquantes sont créées, les colonnes nouvelles
   ajoutées en fin de table (`SG_MigrerSchema`), jamais supprimées ; l'opération est journalisée.
4. Contrôle qualité + vérification du journal.
