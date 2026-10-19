# Backend — plateforme de transport (prototype, phase 5)

API NestJS + PostgreSQL/PostGIS : courses par application **et par SMS**, attribution d'un
chauffeur unique, vagues d'offres, expirations, OTP, chiffrement des numéros.

## Ce qui fonctionne (vérifié par 54 tests automatisés sur une vraie base PostgreSQL)

- Connexion par code SMS (OTP) : expiration 5 min, 5 essais max, 3 envois / 15 min, code masqué dans les journaux.
- Inscription chauffeur → validation admin → disponibilité → envoi de position (retrait si batterie < 5 %).
- Recherche de repères tolérante aux fautes et aux alias (ex. « mustashfa » → Hôpital).
- Prix indicatif (base + prix/km, arrondi à 50 FCFA).
- Création de course idempotente, zone desservie, 1 demande active par client, 3 demandes / heure max.
- Attribution : rayon 1,5 → 3 → 5 km, classement distance + taux d'acceptation, vagues (3 chauffeurs app, 1 seul chauffeur SMS), expiration 20 s (app) / 90 s (SMS), refus = vague suivante immédiate, « aucun chauffeur » après délai.
- **Zéro double attribution** sous accès concurrents réels (20 tours × 10 chauffeurs simultanés, double appui, SMS contre application, expiration pendant les acceptations).
- Mode SMS complet : `CLANDO grand marche > hopital`, format compact de l'app (`M1 C lat lng D id`), choix numérotés si lieu ambigu, `DISPO`, `STOP`, `OUI/NON/FIN code`, `ANNULER`, SMS en double ou en retard, signature HMAC du webhook, SMS ≤ 160 caractères.
- Déroulement : arrivé → début avec code client à 4 chiffres → fin ; annulations ; numéros visibles uniquement pendant la course.

## Ce qui est SIMULÉ ou PAS ENCORE FAIT

| Élément | État |
|---|---|
| Envoi réel de SMS | **Simulé** (`SMS_PROVIDER=fake`). Aucune passerelle réelle branchée (passerelle SIM, agrégateur ou opérateur à venir). |
| Notifications vers les applications (Socket.IO / FCM) | **Simulées** (journal). Les applications peuvent interroger `GET /v1/rides/:id` et `GET /v1/driver/offers`. |
| Repères, zone, tarifs | **Fictifs** (`sql/900_seed_dev.sql`) — à remplacer par la collecte terrain. |
| Temps d'arrivée par la route (OSRM) | Non branché : distance à vol d'oiseau. |
| Redis, multi-instances | Non utilisés : un seul serveur, état dans PostgreSQL. |
| Tableau de bord admin, applications Flutter, avis, incidents, exports, documents chauffeurs | Phases suivantes. |
| Console opérateur (lieu introuvable → rappel humain) | Non faite : le client reçoit un message d'aide. |

## Installer et lancer sur votre PC Windows (disque G:)

### 1. Logiciels à installer (une seule fois)

- **Git** : https://git-scm.com/download/win
- **Node.js 20 LTS ou plus récent** : https://nodejs.org
- **Docker Desktop** (pour la base PostgreSQL + PostGIS) : https://www.docker.com/products/docker-desktop
  *Sans Docker : installez PostgreSQL 16 depuis https://www.postgresql.org/download/windows/
  puis l'extension PostGIS via « Stack Builder », et créez l'utilisateur/base `transport` (mot de passe `transport`).*

### 2. Récupérer le code sur G:

Dans **PowerShell** :

```powershell
G:
git clone -b ccr-40a6f0f5-erf0uy https://github.com/tchaltchal26/gestion-depenses.git
cd G:\gestion-depenses\transport-tchad\backend
```

Pour récupérer plus tard les nouvelles versions : `cd G:\gestion-depenses` puis `git pull`.

### 3. Préparer et démarrer

```powershell
npm install
npm run init:env        # crée .env avec des clés aléatoires (développement)
npm run db:up           # démarre PostgreSQL + PostGIS dans Docker (Docker Desktop doit être ouvert)
npm run db:seed-dev     # crée les tables + données fictives de test
npm run build
npm start               # → API démarrée sur http://localhost:3000/v1
```

### 4. Essayer

- Ouvrez **http://localhost:3000/dev** : simulateur de téléphones en mode SMS.
  1. Créez un chauffeur (ex. 66000010, Moussa, Clando), puis envoyez depuis son téléphone : `DISPO grand marche`
  2. Ouvrez un client (ex. 66000001) et envoyez : `CLANDO grand marche > hopital`
  3. Le chauffeur reçoit l'offre : répondez `OUI <code>` puis `FIN <code>`.
  4. Essayez aussi : `TAXI marche` (choix multiple), `ANNULER`, ou laissez expirer l'offre puis cliquez « Faire avancer le répartiteur ».
- Lancez les tests automatisés : `npm test` (utilise une base séparée `transport_test`, voir ci-dessous).

### Lancer les tests sur votre PC

Les tests effacent et recréent la base `transport_test`. Créez-la une fois :

```powershell
docker compose exec db psql -U transport -c "CREATE DATABASE transport_test"
$env:TEST_DATABASE_URL="postgres://transport:transport@localhost:5433/transport_test"
npm test
```

## Problèmes fréquents

**« authentification par mot de passe échouée pour l'utilisateur transport »**
Un autre PostgreSQL (installé directement sur Windows) occupe le port 5432 et reçoit la connexion
à la place de celui de Docker. Depuis cette version, Docker utilise le port **5433** :

```powershell
docker compose down
git pull
```
Puis dans le fichier `.env`, remplacez `localhost:5432` par `localhost:5433`
(ou supprimez `.env` et relancez `npm run init:env`), et reprenez à `npm run db:up`.

**« connection refused »** : la base démarre encore ; attendez 30 secondes et relancez.

## Structure

```
sql/001_schema.sql           schéma (PostGIS, contraintes anti-double attribution)
sql/900_seed_dev.sql         données FICTIVES de développement
src/rides/rides.service.ts   courses, attribution, vagues, machine d'états
src/sms/                     parseur, modèles de SMS, réception, envoi, fournisseur simulé
src/auth/                    OTP, jetons, rôles
src/geo/geo.service.ts       repères, zone desservie, prix indicatif
src/dev/                     simulateur (actif seulement avec DEV_TOOLS=1, interdit en production)
test/                        54 tests (unitaires, parcours app, parcours SMS, concurrence)
```

## Principales routes

| Méthode | Route | Rôle |
|---|---|---|
| POST | `/v1/auth/otp/request` `{phone}` · `/v1/auth/otp/verify` `{phone, code}` | public |
| GET | `/v1/landmarks/search?q=` · `/v1/tariffs/quote?type=&from=&to=` | connecté |
| POST | `/v1/rides` (en-tête `Idempotency-Key`) · GET `/v1/rides/:id` · POST `/v1/rides/:id/cancel` | client |
| POST | `/v1/driver/register` · PUT `/v1/driver/availability` · POST `/v1/driver/location` · GET `/v1/driver/offers` | chauffeur |
| POST | `/v1/offers/:id/accept` · `/decline` · `/v1/rides/:id/arrived` · `/start` `{pin}` · `/complete` | chauffeur |
| POST | `/v1/admin/drivers/:id/approve` · `/v1/admin/users/:id/suspend` | admin |
| POST | `/v1/sms/inbound/:provider` (en-tête `X-Signature` = HMAC-SHA256 du corps) | passerelle SMS |

Créer un administrateur : `npm run admin:create -- 66123456`.
