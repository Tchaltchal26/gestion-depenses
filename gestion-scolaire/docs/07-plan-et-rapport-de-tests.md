# 07 — Plan de tests et rapport de validation

## 1. Organisation

| Suite | Outil | Commande | Portée |
|---|---|---|---|
| Code VBA métier | LibreOffice 24.2 (mode VBA), piloté par UNO | `python3 tools/tests/run_vba_tests.py` | Tous les modules sauf `modUI` |
| Outil éditeur de licences | pytest | `python3 -m pytest tools/tests/test_licence.py -q` | `sg_licence.py` |
| Contrôle croisé | lanceur VBA | automatique en fin de suite VBA | Demande d'activation VBA lue par l'outil Python |
| Sous Excel (à faire) | `Installer-SG.ps1 -Tests` | voir `06` | Même suite `modTests` dans Excel ; tests de licence à alimenter avec les fixtures |

Chaque test VBA repart d'une **base vierge** (tables vidées — fonction réservée au mode test),
crée ses données **fictives** et un administrateur, puis vérifie les résultats attendus.
Les licences de test sont signées par une **clé RSA éphémère** générée à chaque exécution (jamais
enregistrée) ; une seconde clé produit la licence « signée par un autre éditeur ».

## 2. Plan de tests (résultats attendus)

| Test | Couvre | Exemples de vérifications |
|---|---|---|
| Utilitaires | SHA-256, Base64, UTF-8, dates, normalisation, grands entiers | Vecteurs FIPS 180-4 ; 29/02 bissextile ; réduction modulaire = Python |
| CoucheDonnees | DAL | Formule stockée comme texte ; zéros de tête ; ID en double refusé ; insertion invalide sans écriture partielle ; ID non modifiable |
| Securite | Comptes | Mot de passe faible refusé ; empreinte seule stockée ; sel non journalisé ; échecs journalisés ; verrou après 5 échecs ; réinitialisation et changement ; compte désactivé refusé mais conservé |
| Permissions | Rôles | Secrétaire : crée un dossier, ne clôture pas ; enseignant : lecture seule, pas de journal ; admin technique : aucun accès élèves ; retrait de droit et d'habilitation effectifs et journalisés |
| Isolation | Multi-établissements | Recherche, lecture, modification, inscription inter-établissements refusées et journalisées ; export ne contenant que l'établissement |
| Annees | Cycle de vie | Doublon, chevauchement, une seule ouverte, clôture sans sauvegarde refusée, écriture en année clôturée refusée, réouverture motivée |
| Structure | Organisation | Codes uniques, mode d'évaluation, ordre unique, multi-cycles, reconduction sans doublon, classes source intactes |
| Eleves | Dossiers | Matricule au format ; doublon malgré accents ; homonyme motivé ; dates impossibles/futures ; champs protégés ; historique ancienne/nouvelle valeur/motif/auteur ; recherche sans accents ; responsables multi-enfants |
| Inscriptions | Parcours | Double inscription, date hors année, capacité, changement de classe historisé (veille de la date d'effet), radiation réservée, départ (dossier conservé), suspension/réintégration, annulation conservée |
| Reinscription | Pluriannuel | Passage et redoublement déduits ; nouvel enregistrement ; année précédente inchangée |
| Transfert | Inter-établissements | Statut TRANSFERE, nouveau dossier lié, historique non copié, opérations regroupées (TxRef) |
| Audit | Intégrité | Chaîne intègre ; ancre retrouvée/inconnue ; modification d'un événement détectée ; rupture de chaînage détectée |
| Sauvegarde | Continuité | Sauvegarde vérifiée enregistrée avec l'empreinte du journal ; ancre exportable ; export complet |
| Rapports | Pilotage | Indicateurs, effectifs, liste triée, anomalies détectées sans correction automatique, certificats numérotés séquentiellement et journalisés |
| Licences | Commercial | Essai intégré et ses limites ; licence falsifiée, autre clé, autre poste, fichier illisible refusés ; modules débloqués/bloqués ; établissement hors licence en consultation ; alertes, grâce, expiration (consultation, export, sauvegarde maintenus, aucune donnée supprimée) ; copie détectée ; recul d'horloge ; révocation authentique/falsifiée ; perpétuelle après maintenance ; version non couverte |
| Schema | Maintenance | Schéma complet, version enregistrée, réinitialisation sans doublon |
| pytest (23) | Outil éditeur | Falsifications (4 variantes), autre clé, ligne ajoutée, CRLF/BOM, validations, dates, révocations, demande altérée, cycle émission→transfert→révocation→liste, journal du registre altéré détecté, intégration de la clé publique, absence de clé privée dans `src/` |

## 3. Rapport de la dernière exécution

Voir `docs/rapport-tests-vba.txt` (généré par le lanceur) — résumé :

| Suite | Résultat | Environnement |
|---|---|---|
| VBA (16 tests, 235 vérifications) | **235 réussies, 0 échec** | LibreOffice 24.2.7, Linux, 2026-10-17, ~410 s |
| Contrôle croisé demande d'activation | **OK** | idem |
| pytest outil éditeur | **23 réussis, 0 échec** | Python 3.13, cryptography 50 |
| Démonstration (`SG_DemoPourApercu`) | 40 élèves, 296 événements, journal intègre, 0 anomalie | LibreOffice |

Défauts trouvés et corrigés pendant la validation : contrainte d'unicité du numéro de série qui
empêchait le transfert de licence (registre) ; trois incompatibilités LibreOffice (format horaire,
paramètres optionnels, retour `Byte()`) ; deux erreurs dans les tests eux-mêmes.

## 4. Non couvert par les tests automatiques

- `modUI` (écrans, boutons, formulaire, PDF), `Sauv_Creer` (copie + réouverture), installateur
  PowerShell : **à valider manuellement sous Excel Windows** avec la check-list :
  1. installation → connexion masquée ; 2. création d'une fiche, doublon, homonyme ;
  3. inscription, changement de classe, fin d'inscription ; 4. tableau de bord, qualité, journal ;
  5. sauvegarde dans un dossier, vérification de l'ancre ; 6. certificat PDF ;
  7. demande d'activation → licence émise → installation ; 8. modification de la date système
  → statut HORLOGE ; 9. copie du fichier sur un autre poste → AUTRE_INSTALLATION.
- Tests de charge (10 000 élèves) et Excel Mac.
