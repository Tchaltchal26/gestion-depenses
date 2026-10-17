# 4. Budget prévisionnel (FCFA)

> **Avertissement** : aucun devis n'a été obtenu. Les montants ci-dessous sont des
> **estimations** fondées sur les hypothèses listées. Chaque ligne marquée **[DEVIS]** doit
> être remplacée par un devis réel avant décision d'investissement.

## 4.1 Hypothèses générales

| Hypothèse | Valeur retenue | Source / statut |
|---|---|---|
| Taux de change | 1 EUR = 655,957 FCFA (parité fixe) ; 1 USD ≈ 600 FCFA | Parité fixe EUR ; USD à ajuster |
| Développeur senior (backend/lead), Tchad ou Afrique centrale, mensuel | 700 000 – 1 000 000 | Estimation à valider par 3 entretiens de recrutement |
| Développeur Flutter confirmé, mensuel | 500 000 – 700 000 | Idem |
| Développeur intermédiaire / front admin, mensuel | 350 000 – 500 000 | Idem |
| Agent terrain / opérateur de régulation à Abéché, mensuel | 100 000 – 150 000 | Estimation à valider localement |
| Prix moyen d'un SMS A2P sortant | 25 FCFA | **[DEVIS]** opérateurs / agrégateurs |
| SMS sortants par course SMS | ≈ 3,8 | Voir doc 02 §2.5 |
| VPS en Europe | 15 – 60 EUR par serveur et par mois | Tarifs publics OVH/Hetzner/Scaleway (à revérifier) |
| Comptes développeur | Google Play : 25 USD une fois ; Apple : 99 USD/an (non nécessaire en V1) | Tarifs publics |
| Freelances ou prestataires étrangers | Non retenus par défaut (coût 2–4× supérieur) | Option si compétences introuvables localement |

## 4.2 Scénario 1 — Version initiale économique (pilote Abéché)

**Périmètre** : application Android client + chauffeur (Flutter), backend, admin simple,
mode SMS via **passerelle SIM**, opérateur de régulation, zone pilote (centre-ville et
2–3 quartiers), 3 types de transport, paiement en espèces, 150–300 chauffeurs.
**Équipe** : 1 lead backend/admin, 1 dev Flutter, 1 designer à temps partiel, 1 chef de projet
terrain à Abéché. **Durée** : ≈ 6 mois jusqu'au lancement pilote.

| Poste | Min | Max | Détail |
|---|---:|---:|---|
| Étude et conception | 2 000 000 | 3 000 000 | Enquête terrain (H1–H8, H10–H11), UX, cahier des charges complet |
| Développement des applications | 3 500 000 | 5 000 000 | 1 dev Flutter × 6 mois + design |
| Développement du backend | 4 200 000 | 6 000 000 | 1 lead × 6 mois |
| Tableau de bord | 1 500 000 | 2 500 000 | React + Refine, écrans essentiels |
| GPS et cartographie | 1 000 000 | 1 500 000 | Collecte de 500–1 000 repères (2 agents, ≈ 6 semaines), contribution OSM, OSRM |
| Intégration SMS | 500 000 | 1 000 000 | 2 smartphones passerelle, SIM, onduleur, intégration |
| Tests et sécurité | 1 000 000 | 2 000 000 | Tests terrain, revue de sécurité légère |
| Hébergement (pendant le développement) | 200 000 | 400 000 | 6 mois |
| Recrutement et formation des chauffeurs | 1 500 000 | 2 500 000 | Sessions, autocollants, gilets, kit d'accueil |
| Communication et lancement | 2 000 000 | 3 500 000 | Radio locale, affiches, ambassadeurs |
| Frais administratifs | 1 000 000 | 2 500 000 | Création de société, juriste, déclarations ANSICE, comptes **[DEVIS]** |
| **Sous-total** | **18 400 000** | **29 900 000** | |
| Imprévus (10 %) | 1 840 000 | 2 990 000 | |
| **Total initial** | **≈ 20 M** | **≈ 33 M** | |

**Coûts mensuels après lancement** (hypothèse : 3 000 courses/mois dont 1 500 par SMS)

| Poste | Min | Max |
|---|---:|---:|
| Hébergement (2 VPS, sauvegardes, stockage) | 40 000 | 100 000 |
| SMS (courses + OTP) | 150 000 | 300 000 |
| Maintenance et évolutions (dev à temps partiel) | 400 000 | 700 000 |
| Opérateurs de régulation (2 en rotation) | 200 000 | 300 000 |
| Agent terrain / support chauffeurs | 100 000 | 150 000 |
| Communication continue | 100 000 | 300 000 |
| Téléphonie et data de l'équipe | 30 000 | 60 000 |
| **Total mensuel** | **≈ 1,0 M** | **≈ 1,9 M** |

## 4.3 Scénario 2 — Version professionnelle (toute la ville d'Abéché)

**Périmètre** : scénario 1 + **numéro court SMS** sur Airtel et Moov (SMPP), **USSD**,
masquage des numéros si possible, application iOS optionnelle, abonnement chauffeur payable
via Airtel Money / Moov Money, test d'intrusion externe, supervision complète,
haute disponibilité de base. **Équipe** : chef de projet, 2 backend, 2 Flutter, 1 front,
designer, testeur, DevOps à temps partiel. **Durée** : 9–10 mois.

| Poste | Min | Max |
|---|---:|---:|
| Étude et conception | 4 000 000 | 6 000 000 |
| Développement des applications | 10 000 000 | 16 000 000 |
| Développement du backend | 12 000 000 | 18 000 000 |
| Tableau de bord | 4 000 000 | 6 000 000 |
| GPS et cartographie (toute la ville) | 2 000 000 | 4 000 000 |
| Intégration SMS + USSD (dont frais opérateurs **[DEVIS]**) | 5 000 000 | 11 000 000 |
| Tests et sécurité (dont test d'intrusion externe **[DEVIS]**) | 4 000 000 | 7 000 000 |
| Hébergement (pendant le développement) | 500 000 | 1 000 000 |
| Recrutement et formation des chauffeurs | 3 000 000 | 5 000 000 |
| Communication et lancement | 5 000 000 | 10 000 000 |
| Frais administratifs et juridiques | 3 000 000 | 5 000 000 |
| **Sous-total** | **52 500 000** | **89 000 000** |
| Imprévus (10 %) | 5 250 000 | 8 900 000 |
| **Total initial** | **≈ 58 M** | **≈ 98 M** |

**Mensuel** : ≈ **3,5 M à 6,7 M** (hébergement 150–400 k ; SMS 400 k–1 M ; maintenance
2 devs 1,5–2,5 M ; centre de régulation 4 personnes 500–800 k ; terrain 300–500 k ;
marketing 500 k–1 M ; abonnements numéro court/USSD 100–500 k **[DEVIS]**).

## 4.4 Scénario 3 — Version évolutive à grande échelle (plusieurs villes)

**Périmètre** : scénario 2 + N'Djamena et 2–3 autres villes, comptes entreprises/ONG avec
facturation, paiements mobiles intégrés, infrastructure multi-instances (Kubernetes ou
équivalent managé), équipe produit permanente, centre d'appel, bureaux locaux.
**Durée** : 12–18 mois.

| Poste | Min | Max |
|---|---:|---:|
| Produit, conception, études multi-villes | 10 000 000 | 18 000 000 |
| Développement (apps, backend, admin, paiements, B2B) | 60 000 000 | 110 000 000 |
| Cartographie et repères (4 villes) | 8 000 000 | 15 000 000 |
| Télécoms (SMS/USSD/voix, numéros, intégrations) **[DEVIS]** | 12 000 000 | 25 000 000 |
| Sécurité, conformité, audits | 10 000 000 | 18 000 000 |
| Infrastructure (première année) | 6 000 000 | 12 000 000 |
| Recrutement et formation chauffeurs (4 villes) | 12 000 000 | 25 000 000 |
| Marketing et lancement | 20 000 000 | 40 000 000 |
| Juridique, licences, administratif | 6 000 000 | 12 000 000 |
| **Sous-total** | **144 000 000** | **275 000 000** |
| Imprévus (10 %) | 14 400 000 | 27 500 000 |
| **Total initial** | **≈ 160 M** | **≈ 300 M** |

**Mensuel** : ≈ **10 M à 25 M** selon nombre de villes et volume SMS.

## 4.5 Ce qui nécessite un devis réel avant tout engagement

1. Tarifs SMS A2P, numéro court, SMPP et USSD chez **Airtel Tchad** et **Moov Africa Tchad**.
2. Frais éventuels ARCEP Tchad pour le numéro court / services à valeur ajoutée.
3. Honoraires juridiques (statuts, CGU, politique de confidentialité, déclarations ANSICE,
   avis sur la réglementation transport).
4. Salaires réels des profils (3 entretiens par poste).
5. Test d'intrusion externe.
6. Coûts de communication locale (radio, affichage, impression).
7. Conditions d'accès aux API Airtel Money / Moov Money (scénarios 2 et 3).

## 4.6 Modèle économique — comparaison

| Modèle | Avantages | Inconvénients au Tchad | Recommandation |
|---|---|---|---|
| **Commission par course** (ex. 10 %) | Proportionnel à l'usage | Paiement en espèces ⇒ commission impossible à prélever automatiquement ; incite au contournement ; montant faible sur une course clando | Pas en V1 |
| **Abonnement chauffeur** (hebdomadaire) | Simple, prévisible, compatible espèces / mobile money ; pas de friction à chaque course | Le chauffeur doit percevoir la valeur avant de payer | **Recommandé V1** après 2–3 mois gratuits ; montants à tester (ordre de grandeur à valider : clando < rakcha < taxi) |
| **Crédit de courses prépayé** (le chauffeur recharge, chaque course acceptée débite un petit montant) | Lié à la valeur réellement reçue ; fonctionne avec recharge mobile money ou agent | Nécessite un portefeuille interne et la gestion de remboursements | **Recommandé en V2**, en alternative à l'abonnement |
| **Frais de service client** transparents (petit montant fixe) | Revenus indépendants des chauffeurs | Sensibilité au prix ; difficile à collecter en espèces (le chauffeur collecte pour la plateforme) | Éventuellement pour les courses SMS (coût réel) — à tester |
| **Contrats entreprises / ONG** | Panier élevé, paiement par virement, facturation mensuelle | Cycle de vente long ; exigences de sécurité des ONG (certaines interdisent les motos) | **Fort potentiel à Abéché** (présence humanitaire) dès la V2 |
| **Services aux chauffeurs** (assurance, crédit téléphone, entretien, formation) | Fidélisation | Nécessite partenaires | V3 |

**Point mort indicatif (scénario 1)** : avec 1,0–1,9 M FCFA de charges mensuelles et un
abonnement moyen hypothétique de 1 500 FCFA/semaine (≈ 6 000/mois), il faut **170 à 320
chauffeurs payants**. C'est ambitieux pour Abéché : le pilote doit vérifier ce chiffre, ou
le modèle doit combiner abonnement + contrats entreprises.
