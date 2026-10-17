# 1. Vision produit et hypothèses à vérifier

## 1.1 Vision produit

> **Permettre à n'importe quel habitant d'Abéché, avec ou sans forfait Internet, de trouver
> en quelques minutes un taxi, un rakcha ou un clando fiable et identifié, à un prix connu
> d'avance — et permettre aux transporteurs de trouver plus de clients avec moins de temps
> perdu à vide.**

### Pour qui

| Utilisateur | Problème actuel (hypothèse) | Ce que la plateforme apporte |
|---|---|---|
| Client avec smartphone et data | Doit marcher jusqu'à un axe ou attendre au hasard ; prix négocié à chaque fois | Carte, chauffeurs proches, prix indicatif, suivi, historique |
| Client sans data (forfait épuisé, téléphone simple) | Idem, sans alternative | Demande par **SMS** (puis USSD si accord opérateur), confirmation par SMS, appel classique |
| Taximan / rakchaman / clandoman | Temps à vide, recherche de clients, insécurité | Demandes de courses proches, réputation, historique de revenus |
| Femmes, personnes âgées, visiteurs, ONG | Sécurité, méconnaissance des prix | Chauffeur identifié et vérifié, partage de la course, signalement |
| Entreprises / ONG / administrations (Abéché est une base humanitaire importante) | Transport du personnel local, traçabilité | Comptes entreprise, facturation mensuelle (phase ultérieure) |

### Principes directeurs

1. **SMS d'abord compatible** : chaque parcours essentiel (demander, accepter, confirmer,
   annuler) doit avoir un équivalent SMS. L'application enrichit l'expérience, elle n'est pas
   la seule porte d'entrée.
2. **Repères avant adresses** : à Abéché, on se repère par quartier, marché, mosquée, école,
   rond-point. La base de **repères locaux** est un actif central du projet.
3. **Espèces par défaut** : pas de paiement numérique obligatoire en V1.
4. **Économe en données et en batterie** : application légère, cartes hors ligne de la ville,
   envoi de position adaptatif.
5. **Un seul chauffeur par course, toujours** : l'attribution est atomique côté serveur,
   quel que soit le canal.
6. **Humain en renfort** : un opérateur de régulation (centre d'appel léger) traite les
   demandes ambiguës et les incidents pendant le pilote.

### Ce que le projet n'est pas (en V1)

- Pas un service de livraison, pas de covoiturage interurbain.
- Pas de paiement in-app obligatoire.
- Pas de tarification dynamique (« surge ») : socialement risquée et injustifiable au démarrage.

## 1.2 Hypothèses à vérifier au Tchad

Chaque hypothèse est classée par criticité (**C** = critique, le projet change si elle est
fausse ; **I** = importante ; **S** = secondaire) avec une méthode de vérification.

### Marché et usages

| # | Hypothèse | Crit. | Méthode de vérification |
|---|---|---|---|
| H1 | Il y a à Abéché assez de taxis, rakchas et clandos pour assurer une densité de chauffeurs utile (≥ 50 chauffeurs actifs simultanés dans la zone pilote) | C | Comptage terrain sur 3 jours (matin/midi/soir) sur 5 points clés ; rencontre des associations/syndicats de transporteurs |
| H2 | Les clients sont prêts à demander un transport à distance plutôt que de héler | C | 100 entretiens courts clients (marchés, université, hôpital, gares) |
| H3 | Les chauffeurs acceptent d'être identifiés et vérifiés par une plateforme | C | 30 entretiens chauffeurs + réunion avec leurs représentants |
| H4 | Une part significative des chauffeurs possède un smartphone Android (≥ Android 8) | C | Questionnaire chauffeurs (modèle de téléphone, version Android, opérateur, budget data) |
| H5 | Les prix pratiqués sont relativement stables par trajet type (permettant une grille indicative zone à zone) | I | Relevé de 200 trajets réels avec prix payé |
| H6 | Les rakchas fonctionnent (au moins en partie) à la place et en partage, différemment des clandos (course individuelle) | I | Observation terrain + entretiens ; impacte le modèle de « course » |
| H7 | Les langues d'usage sont l'arabe tchadien et le français ; une part des chauffeurs lit peu ou pas | I | Entretiens ; conditionne SMS (alphabet), icônes, messages vocaux |
| H8 | Les clients acceptent un frais de service ou les chauffeurs un abonnement | I | Test de prix lors des entretiens (« combien paieriez-vous… ») puis test réel pendant le pilote |

### Télécoms et technique

| # | Hypothèse | Crit. | Méthode de vérification |
|---|---|---|---|
| H9 | Les opérateurs actifs sont **Airtel Tchad** et **Moov Africa Tchad**. *Note : « Tigo » n'existe plus au Tchad ; la marque a été remplacée par Moov Africa (groupe Maroc Telecom) en décembre 2020.* | C | Confirmé par la presse ; à reconfirmer auprès des boutiques d'Abéché |
| H10 | La couverture data (3G/4G) à Abéché est disponible mais irrégulière | C | Test terrain avec 2 téléphones (1 SIM Airtel, 1 SIM Moov) sur 20 points, à 3 moments de la journée : débit, latence, coupures |
| H11 | Le SMS fonctionne quasi partout où il y a du réseau voix, avec un délai acceptable (< 1 min la plupart du temps) | C | Envoi de 200 SMS de test entre opérateurs (Airtel→Moov, Moov→Airtel, même réseau), mesure des délais et pertes |
| H12 | Les opérateurs proposent un **numéro court SMS** et/ou un accès **SMPP/API** aux entreprises, et l'**USSD** via contrat | C | Rendez-vous avec les services Entreprises d'Airtel et Moov à N'Djamena ; demande de grille tarifaire écrite |
| H13 | Le prix d'un SMS A2P (application vers personne) se situe entre 10 et 40 FCFA | I | Devis opérateurs et agrégateurs (chiffre actuel = estimation non vérifiée) |
| H14 | Le GPS des téléphones fonctionne sans data (vrai techniquement), mais le premier repérage peut prendre plusieurs minutes sans assistance réseau | I | Test terrain « mode avion + GPS » sur 5 modèles de téléphones courants |
| H15 | La couverture OpenStreetMap d'Abéché est exploitable (rues, quartiers) | I | Inspection OSM ; comparaison terrain ; contributions possibles (communauté OSM / HOT active dans la région) |
| H16 | Les coupures d'électricité rendent la recharge des téléphones des chauffeurs difficile | S | Entretiens chauffeurs ; envisager batteries externes en kit de bienvenue |

### Réglementaire et commercial

| # | Hypothèse | Crit. | Méthode de vérification |
|---|---|---|---|
| H17 | Aucune loi tchadienne spécifique n'encadre encore les plateformes de VTC/mise en relation | C | Consultation d'un avocat d'affaires à N'Djamena ; Ministère des Transports |
| H18 | Les clandos (motos-taxis) sont autorisés à Abéché (il y a eu des restrictions sur les motos dans certaines villes tchadiennes pour raisons de sécurité) | C | Mairie d'Abéché, gouvernorat du Ouaddaï, police ; vérifier arrêtés en vigueur |
| H19 | Les chauffeurs disposent de documents vérifiables (permis, carte grise, autorisation de transport, assurance) | I | Échantillon terrain ; définir le minimum acceptable avec les autorités |
| H20 | Les traitements de données doivent être déclarés à l'**ANSICE** (loi n° 007/PR/2015) | C | Avocat + ANSICE ; voir document 07 |
| H21 | L'attribution d'un numéro court relève de l'**ARCEP Tchad** et/ou des opérateurs | C | ARCEP Tchad + opérateurs |
| H22 | Des API marchandes Airtel Money / Moov Money sont accessibles à une jeune société | S (V1) | Services entreprises des opérateurs ; non bloquant en V1 (espèces) |

## 1.3 Principaux risques

| Risque | Type | Probabilité | Impact | Mitigation |
|---|---|---|---|---|
| Densité insuffisante de chauffeurs (« poule et œuf ») | Commercial | Élevée | Critique | Recruter les chauffeurs **avant** les clients ; zone pilote restreinte ; 3 mois gratuits ; incitations aux premiers inscrits |
| Les chauffeurs contournent la plateforme après la première course (échange de numéros) | Commercial | Élevée | Élevé | Modèle par abonnement plutôt que commission ; valeur continue (flux de clients, réputation) ; ne pas lutter frontalement |
| Coût SMS trop élevé par rapport au prix d'une course (clando ≈ quelques centaines de FCFA) | Économique | Moyenne | Élevé | Limiter à 3–4 SMS par course ; offres séquentielles ; inciter les chauffeurs à passer en data ; voir doc 02 |
| Opérateurs lents ou refus pour numéro court/USSD | Télécom | Moyenne | Élevé | Démarrer avec une passerelle SIM (pilote limité, risque assumé) + appels ; négocier en parallèle |
| Blocage des SIM de la passerelle par l'opérateur (usage commercial non autorisé) | Télécom | Moyenne | Moyen | Usage pilote déclaré à l'opérateur ; volumes faibles ; migrer vers SMPP officiel |
| Interdiction ou restriction des motos-taxis | Réglementaire | Inconnue | Élevé | Vérification préalable (H18) ; architecture multi-types permettant de désactiver un type |
| Sécurité des passagers / agressions | Sécurité | Moyenne | Critique | Vérification des chauffeurs, partage de course, signalement, suspension rapide, lien avec police locale |
| Faible littératie / langue | Adoption | Moyenne | Élevé | Arabe tchadien + français, icônes, messages vocaux, ambassadeurs terrain, opérateur téléphonique |
| Connectivité vers serveurs hébergés hors du Tchad (latence, coupure internationale) | Technique | Moyenne | Moyen | Protocoles tolérants (reprise, idempotence) ; SMS comme canal de secours ; surveillance |
| Fraude : fausses demandes, faux chauffeurs, comptes multiples | Sécurité | Moyenne | Moyen | Vérification OTP, limites par numéro, KYC chauffeurs, réputation, liste noire |
