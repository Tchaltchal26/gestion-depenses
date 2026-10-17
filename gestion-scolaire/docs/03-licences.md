# 03 — Licences commerciales, activation et protection

## 1. Répartition des rôles

| Élément | Où | Contient |
|---|---|---|
| Clé **privée** RSA 3072 bits | Poste de l'éditeur uniquement, fichier `cle_privee.pem` chiffré par phrase secrète | Sert à **signer** |
| Clé **publique** | Constante `CLE_PUBLIQUE_MODULE` de `modLicence.bas` | Sert à **vérifier** ; ne permet pas de signer |
| Registre commercial | `registre_licences.db` (SQLite) chez l'éditeur | Licences, activations, journal chaîné |
| Licence | Fichier `.sglic` remis au client | Droits signés |

L'éditeur gère les licences **sans accéder aux données scolaires** : il ne reçoit qu'une demande
d'activation (identifiant d'installation, empreinte, codes d'établissements, version).

## 2. Format de licence `SGLIC-1`

```
FORMAT=SGLIC-1
LICENCE_ID=LIC-3E1478930A04
SERIE=SG-D647-6D6F-9EED-4AE7
TITULAIRE=Groupe scolaire fictif
ETABLISSEMENTS=ETB-A;ETB-B
EDITION=STANDARD
TYPE=ANNUELLE                 (ANNUELLE, MENSUELLE, PERPETUELLE, ESSAI, DEMONSTRATION, MULTI_ETABLISSEMENTS, RESEAU)
EMISE_LE=2026-09-01
EXPIRE_LE=2027-08-31          (vide pour une perpétuelle)
MAINTENANCE_JUSQU_AU=2027-08-31
VERSION_MAJEURE_MAX=0         (versions couvertes)
INSTALLATIONS_MAX=1
NB_ETABLISSEMENTS_MAX=2
MODULES=ELEVES;INSCR;NOTES    (le noyau ETAB, PARAM, UTIL, ANNEE, AUDIT, LICENCE, SAUVEG, STRUCT est toujours inclus)
EMPREINTE_INSTALLATION=<64 hex> ou *   (* = licence non liée à un poste, ex. essai)
GRACE_JOURS=15
SIGNATURE=<base64 RSASSA-PKCS1-v1_5 / SHA-256 des lignes précédentes jointes par LF>
```

Toute modification d'un seul caractère (module ajouté, date prolongée, type changé, ligne
insérée) invalide la signature : vérifié par les tests Python **et** VBA.

## 3. Activation hors ligne (livrée)

1. Le client : **ACCUEIL → Demande d'activation** → fichier `.sgreq` (avec somme de contrôle).
2. L'éditeur :
   ```
   python3 sg_licence.py --registre registre.db emettre --privee cle_privee.pem \
       --titulaire "Groupe scolaire X" --etablissements "ETB-A;ETB-B" --type ANNUELLE \
       --modules "ELEVES;INSCR" --demande client.sgreq --sortie client.sglic
   ```
3. Le client : **ACCUEIL → Installer une licence** → choix du `.sglic`. Le classeur vérifie
   signature, format, révocation, empreinte, version, dates ; l'activation est journalisée.

## 4. Comportement selon l'état

| Statut | Saisie | Consultation / export / sauvegarde | Cause |
|---|---|---|---|
| `ESSAI_INTEGRE` | Oui (2 établissements, 100 élèves) | Oui | Aucune licence, < 30 jours depuis l'installation |
| `VALIDE` | Oui, modules et établissements de la licence | Oui | — |
| `GRACE` | Oui + alerte | Oui | Expirée depuis ≤ `GRACE_JOURS` |
| `EXPIREE` | **Non** | Oui | Au-delà de la grâce |
| `NON_ACTIVEE` | **Non** | Oui | Essai intégré terminé |
| `AUTRE_INSTALLATION` | **Non** | Oui | Copie sur un autre poste |
| `REVOQUEE` | **Non** | Oui | Liste de révocation importée |
| `VERSION_NON_COUVERTE` | **Non** | Oui | Version majeure > `VERSION_MAJEURE_MAX` |
| `HORLOGE` | **Non** | Oui | Date système antérieure à une date déjà observée |
| `INVALIDE` | Installation refusée | — | Signature ou format incorrect |

Alertes avant expiration à 30, 15 et 7 jours. Une **perpétuelle** reste valide après la fin de
maintenance (seule une alerte est affichée). Les fonctions de **récupération** (comptes,
installation d'une licence, sauvegarde, export) restent toujours disponibles : **l'expiration ou
la révocation ne rend jamais les données inaccessibles**.

## 5. Opérations de l'éditeur (`sg_licence.py`)

| Commande | Usage |
|---|---|
| `cles --dossier D` | Génère la paire de clés (refuse d'écraser une clé existante). Phrase secrète ≥ 12 caractères |
| `integrer-cle --publique D/cle_publique.pem --vba src/vba/modLicence.bas` | Intègre la clé publique avant de construire le classeur |
| `emettre ...` | Émet et enregistre une licence |
| `verifier --publique ... fichier.sglic` | Contrôle une licence |
| `transferer LIC-ID --demande nouveau.sgreq --motif ...` | Changement d'ordinateur / panne : nouvelle licence, mêmes droits, ancienne révoquée |
| `revoquer LIC-ID --motif ... [--statut SUSPENDUE]` | Révocation / suspension motivée |
| `revocations --sortie liste.sgrev` | Liste de révocation signée à distribuer |
| `liste`, `journal` | Rapport commercial ; vérification du journal chaîné du registre |

Renouvellement : émettre une nouvelle licence (nouvelle date) pour la même empreinte ; le client
l'installe, l'ancienne est remplacée (événement `ACTIVATION` avec ancienne et nouvelle référence).

## 6. Conçu mais non développé

- **Serveur d'activation en ligne** (HTTPS, authentification des administrateurs, limitation des
  tentatives, comptage des activations simultanées, liste de révocation en ligne, période de grâce
  hors connexion). L'API prévue : `POST /activations` (demande signée → licence), `GET /revocations`.
- **Signature du projet VBA** par certificat de signature de code et contrôle d'intégrité des modules.
- **Mises à jour signées** (paquet de modules + signature éditeur vérifiée avant import).

## 7. Limites (à communiquer aux clients et aux revendeurs)

- Un utilisateur capable de modifier le code VBA peut supprimer les contrôles de licence.
  La protection du projet VBA par mot de passe se contourne avec des outils publics.
- Hors ligne, deux copies du même fichier **sur le même poste** ne peuvent pas être distinguées.
- L'empreinte repose sur l'identifiant d'installation, le nom du poste et le numéro de volume :
  un changement de disque ou de nom de poste impose un transfert (procédure prévue, non bloquante
  pour les données).
- La protection juridique (contrat de licence, conditions d'utilisation) reste indispensable.
