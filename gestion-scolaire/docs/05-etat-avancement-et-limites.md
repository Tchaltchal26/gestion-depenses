# 05 — État d'avancement et limites connues

Légende : **Testé** = couvert par la suite automatisée (exécutée sous LibreOffice en mode VBA) ·
**Livré non testé** = code écrit mais dépendant d'Excel, à valider sous Excel · **À faire**.

## 1. Fonctionnalités

| Domaine | Fonctionnalité | État |
|---|---|---|
| Noyau | Couche d'accès aux données par identifiant/colonne, stockage texte, anti-injection de formules | Testé |
| Noyau | Initialisation idempotente, migration de schéma par ajout de colonnes | Testé |
| Noyau | SHA-256, Base64, UTF-8, vérification RSA en VBA pur | Testé (vecteurs officiels + Python) |
| Sécurité | Comptes, hachage salé, politique de mot de passe, verrouillage, réinitialisation, désactivation | Testé |
| Sécurité | Rôles par défaut (10), droits par module 0-4, habilitations par établissement et période | Testé |
| Sécurité | Administrateur technique sans accès aux données élèves | Testé |
| Isolation | Contrôle d'appartenance de toutes les références, recherche et export cloisonnés | Testé |
| Audit | Journal chaîné, ancien/nouvelle valeur, motif, auteur, TxRef, détection d'altération, ancre externe | Testé |
| Établissements | Création, modes indépendant/centralisé, paramètres, numérotations | Testé |
| Années | Création, chevauchement, ouverture unique, clôture avec sauvegarde, verrouillage, réouverture motivée | Testé |
| Structure | Cycles (compétences/notes), niveaux ordonnés, filières, classes par année, reconduction | Testé |
| Élèves | Dossier permanent, matricule paramétrable, validations, doublons (accents/tirets), homonyme motivé | Testé |
| Élèves | Modification journalisée champ par champ, motif obligatoire paramétrable, recherche sans accents | Testé |
| Responsables | Création, liens multiples, rôles légal/financier/urgence/récupération | Testé |
| Inscriptions | 1re inscription, réinscription (passage/redoublement déduits), transfert entrant, capacité | Testé |
| Inscriptions | Changement de classe historisé (date d'effet ≠ date d'opération), suspension, réintégration | Testé |
| Inscriptions | Départ, radiation (niveau validation), fin de cycle, annulation conservée | Testé |
| Inscriptions | Transfert inter-établissements (nouveau dossier lié, historique conservé à l'origine) | Testé |
| Rapports | Indicateurs, effectifs par classe, contrôle qualité (11 contrôles), liste de classe triée | Testé |
| Documents | Certificat de scolarité numéroté et journalisé (données) | Testé |
| Documents | Modèle imprimable + export PDF | Livré non testé |
| Licences | Format signé, vérification, empreinte, expiration/grâce/alertes, perpétuelle, version, horloge, révocation, essai intégré, mode consultation | Testé |
| Licences | Outil éditeur : clés, émission, transfert, révocation, liste signée, registre chaîné | Testé (pytest) |
| Licences | Demande d'activation VBA acceptée par l'outil Python | Testé (contrôle croisé) |
| Sauvegarde | Export CSV cloisonné, contrôle « sauvegarde vérifiée du jour » | Testé |
| Sauvegarde | Copie + réouverture + vérification du journal (`Sauv_Creer`) | Livré non testé (Excel) |
| Interface | Accueil, fiche élève, liste, inscriptions, tableau de bord, qualité, journal, licence, sauvegarde | Livré non testé (Excel) |
| Interface | Formulaire de connexion à mot de passe masqué généré à l'installation | Livré non testé (Excel) |
| Installation | Script PowerShell (conversion d'encodage, import, ThisWorkbook, installation) | Livré non testé (Windows) |
| Démonstration | Complexe scolaire fictif : 4 cycles, 10 classes, 40 élèves, 4 comptes | Testé (exécution LibreOffice) |
| Paramétrage | Écran de paramétrage (établissements, structure, utilisateurs) | À faire — fonctions disponibles en VBA |
| Pédagogie | Matières, coefficients, évaluations, notes, compétences, bulletins, délibérations, examens | À faire |
| Finances | Tarifs, factures, échéanciers, paiements, reçus, caisse, dépenses | À faire |
| Présences, personnel, emplois du temps, stocks, modules complémentaires | | À faire |
| Imports | Excel/CSV avec aperçu, doublons et rapport d'erreurs | À faire |
| Licences | Serveur d'activation en ligne, signature du code, mises à jour signées | Conçu, à faire |

## 2. Limites connues

1. **Sécurité au niveau du fichier** : quiconque possède le classeur peut lire les feuilles avec un
   autre outil. Les droits SG s'appliquent dans l'application uniquement. Mode centralisé ⇒ ne
   confier le fichier qu'à des personnes autorisées pour tous ses établissements, ou passer à une
   base serveur.
2. **Journal d'audit** : il détecte les modifications d'événements, mais un expert peut recalculer
   toute la chaîne ; d'où l'ancre externe. Une modification **directe des tables de données**
   (hors application) n'est pas détectée par le journal lui-même — prévu : contrôle de
   réconciliation « valeur actuelle = dernière valeur journalisée ».
3. **Pas de transactions** : les validations précèdent les écritures, mais une coupure d'Excel en
   pleine opération peut laisser une écriture partielle (détectable par le contrôle qualité).
4. **Mono-utilisateur en écriture** : pas d'accès simultané fiable au même classeur.
5. **Mots de passe** : SHA-256 itéré (1 000 itérations par défaut) faute de PBKDF2/Argon2 natif ;
   suffisant contre la lecture directe, faible contre une attaque hors ligne d'un mot de passe court.
6. **Licences** : contournables par modification du code VBA ; copie sur le même poste non
   détectable hors ligne (voir `03-licences.md` § 7).
7. **Performance** : recherches linéaires ; adapté à quelques milliers d'élèves. Sous LibreOffice,
   la vérification RSA prend ~6 s (beaucoup plus rapide sous Excel) ; elle n'a lieu qu'à
   l'ouverture et à l'installation d'une licence.
8. **Tests** : exécutés sous LibreOffice, qui n'est pas Excel. Différences rencontrées et
   contournées : codes de format horaire (`nn`), paramètres optionnels `String` (remplacés par
   `Variant` + `IsMissing`), fonctions renvoyant `Byte()`. La suite doit être rejouée sous Excel
   (`Installer-SG.ps1 -Tests`).
9. **Mot de passe initial** saisi par `InputBox` à l'installation (affiché en clair à l'écran).
10. **Excel pour le Web et mobiles** : non pris en charge (VBA).
