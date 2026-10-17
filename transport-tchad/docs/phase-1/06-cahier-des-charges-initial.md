# 6. Cahier des charges initial (V1 — pilote)

Priorités : **M** = indispensable au pilote, **S** = souhaitable au pilote, **C** = plus tard.

## 6.1 Contexte et objectifs

Mettre en relation clients et transporteurs (taxi, rakcha, clando) à Abéché, via application
Android avec data et via SMS sans data, avec attribution d'un chauffeur unique par course,
paiement en espèces et administration centralisée.

## 6.2 Acteurs

Client · Chauffeur (taximan, rakchaman, clandoman) · Opérateur de régulation · Administrateur ·
Super-administrateur · Passerelle SMS (système).

## 6.3 Application client

| Id | Exigence | Prio |
|---|---|---|
| CL-01 | Inscription/connexion par numéro de téléphone avec OTP par SMS ; repli : vérification par appel de l'opérateur | M |
| CL-02 | Profil : nom, langue (français / arabe), contact d'urgence facultatif | M |
| CL-03 | Position GPS automatique, avec indicateur de précision ; ajustement manuel sur la carte | M |
| CL-04 | Saisie du départ et de la destination par recherche de repères (tolérante aux fautes, FR/arabe) ou point sur la carte | M |
| CL-05 | Choix taxi / rakcha / clando, avec disponibilité affichée par type | M |
| CL-06 | Prix indicatif selon grille zone→zone ; mention « à confirmer avec le chauffeur » pour rakcha/clando si l'admin le configure | M |
| CL-07 | Affichage des chauffeurs proches (positions approximatives, sans identité) | S |
| CL-08 | Création, suivi d'état et annulation de la demande (motif) | M |
| CL-09 | Fiche du chauffeur attribué : prénom, photo, véhicule, plaque, note, code de prise en charge | M |
| CL-10 | Bouton d'appel du chauffeur | M |
| CL-11 | Suivi du chauffeur sur la carte quand les positions sont disponibles | S |
| CL-12 | Historique des courses | M |
| CL-13 | Évaluation (1–5 étoiles + motif) et signalement d'incident | M |
| CL-14 | Mode dégradé : détection d'absence de data, bascule « Demander par SMS » avec message pré-rempli (GPS + repère) | M |
| CL-15 | Page d'instructions SMS et numéro de l'opérateur, consultable hors ligne | M |
| CL-16 | Cartes de la ville disponibles hors ligne | S |
| CL-17 | Partage de la course à un proche (lien ou SMS) | C |

## 6.4 Application chauffeur

| Id | Exigence | Prio |
|---|---|---|
| CH-01 | Inscription par téléphone + OTP | M |
| CH-02 | Profil professionnel, type de transport, véhicule (plaque, couleur, photo) | M |
| CH-03 | Téléversement des documents (photos compressées) ; liste configurable par type | M |
| CH-04 | Statut de vérification visible ; pas d'offres tant que non validé | M |
| CH-05 | Bouton Disponible / Indisponible ; indisponibilité automatique après inactivité prolongée | M |
| CH-06 | Envoi de position adaptatif (selon état et batterie), en lots si réseau intermittent | M |
| CH-07 | Réception d'offre (son + vibration, écran plein) : départ, destination, distance, prix indicatif, compte à rebours | M |
| CH-08 | Accepter / refuser ; message clair si déjà attribuée | M |
| CH-09 | Appel du client ; ouverture de la navigation vers le départ | M |
| CH-10 | Étapes : arrivé → début (saisie du code client) → fin | M |
| CH-11 | Historique et total des revenus déclarés (espèces) par jour/semaine | M |
| CH-12 | Mode SMS : réception des offres par SMS, réponse pré-remplie `OUI <code>` ; déclaration de disponibilité par SMS | M |
| CH-13 | Gestion de la batterie faible et des notifications retardées (rejet des offres expirées côté app) | M |
| CH-14 | Interface utilisable avec peu de lecture (icônes, couleurs, messages vocaux courts) | S |

## 6.5 Tableau de bord administrateur

| Id | Exigence | Prio |
|---|---|---|
| AD-01 | Liste/recherche/fiche des clients et chauffeurs | M |
| AD-02 | File de vérification des documents (valider/refuser avec motif) | M |
| AD-03 | Suspension / réactivation des comptes, avec motif et durée | M |
| AD-04 | Gestion des types de transport (activer/désactiver) | M |
| AD-05 | Zones desservies (dessin de polygones) et grille tarifaire zone→zone | M |
| AD-06 | Paramètres d'attribution (rayons, délais, tailles de vague) | M |
| AD-07 | Suivi en temps réel des demandes et courses ; création de course pour un client au téléphone (console opérateur) | M |
| AD-08 | Incidents et réclamations : affectation, suivi, clôture | M |
| AD-09 | Journal et coût des SMS (par jour, par type de message), alerte de dépassement | M |
| AD-10 | Comptes administrateurs, rôles et permissions ; double authentification | M |
| AD-11 | Statistiques : courses, taux de service, délais, annulations, demande non servie par zone/heure | M |
| AD-12 | Abonnements des chauffeurs et revenus | S |
| AD-13 | Exports CSV/Excel | M |
| AD-14 | Journal d'audit des actions administrateurs | M |
| AD-15 | Supervision : état des services, passerelle SMS, erreurs | M |
| AD-16 | Gestion des repères (ajout, alias, fusion des doublons) | M |

## 6.6 Mode SMS (serveur)

| Id | Exigence | Prio |
|---|---|---|
| SMS-01 | Réception via webhook signé ; enregistrement de chaque message | M |
| SMS-02 | Parseur : type, départ, destination (repères), format compact de l'app | M |
| SMS-03 | Désambiguïsation par choix numérotés ; repli vers l'opérateur | M |
| SMS-04 | Idempotence et limitation par numéro | M |
| SMS-05 | Offres séquentielles aux chauffeurs SMS avec codes à usage unique et expiration | M |
| SMS-06 | Messages en GSM-7 ≤ 160 caractères ; modèles éditables par l'admin | M |
| SMS-07 | Commandes : `DISPO`, `STOP`, `OUI/NON <code>`, `FIN <code>`, `ANNULER`, `AIDE` | M |
| SMS-08 | Abstraction de fournisseur (SIM, agrégateur, SMPP) | M |
| SMS-09 | USSD | C |

## 6.7 Exigences non fonctionnelles

| Id | Exigence |
|---|---|
| NF-01 | Android 8+ ; fonctionnement correct sur téléphone 2 Go de RAM ; APK < 20 Mo hors cartes |
| NF-02 | Consommation data de l'app client < 2 Mo par course (cartes hors ligne) — à mesurer |
| NF-03 | Disponibilité serveur visée 99,5 % pendant le pilote |
| NF-04 | Attribution : temps serveur < 1 s ; zéro double attribution (test automatisé de concurrence) |
| NF-05 | Langues : français et arabe (RTL) ; SMS en caractères latins |
| NF-06 | Toutes les communications en TLS 1.2+ ; données sensibles chiffrées au repos |
| NF-07 | Sauvegarde quotidienne, restauration testée mensuellement ; RPO 24 h (puis 15 min avec WAL), RTO 4 h |
| NF-08 | Journaux sans numéros de téléphone en clair |
| NF-09 | Accessibilité : contrastes forts (lecture en plein soleil), boutons larges |

## 6.8 Hors périmètre V1

Paiement in-app, iOS, USSD, masquage des numéros, comptes entreprises, multi-villes,
tarification dynamique, chat texte in-app.
