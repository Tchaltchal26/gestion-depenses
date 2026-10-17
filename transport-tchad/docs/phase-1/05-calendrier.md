# 5. Calendrier de développement (scénario 1 — pilote Abéché)

Durée totale estimée : **≈ 26 semaines** jusqu'au lancement pilote, puis 8 semaines de pilote
avant décision de généralisation. Les semaines supposent une équipe disponible dès S1 ;
le recrutement peut ajouter 4 à 8 semaines.

| Semaines | Phase | Livrables | Critères de validation |
|---|---|---|---|
| S1–S4 | **Phase 1 bis — Vérifications terrain et légales** (en parallèle de la phase 2) | Rapport d'enquête (H1–H22), tests réseau et SMS, compte rendu des rendez-vous Airtel/Moov, ARCEP, ANSICE, mairie | Hypothèses critiques confirmées ou plan B décidé ; go/no-go écrit |
| S2–S5 | **Phase 2 — Cahier des charges complet** | Spécifications fonctionnelles détaillées, règles tarifaires, textes SMS, politiques (annulation, conservation) | Validation signée par le porteur du projet |
| S4–S7 | **Phase 3 — Architecture, données, API** | Schéma de base, contrat OpenAPI, ADR (décisions d'architecture), environnement CI | Revue technique ; contrat d'API figé en v1 |
| S5–S8 | **Phase 4 — Maquettes** | Maquettes Figma client, chauffeur, admin ; tests papier avec 10 utilisateurs locaux | ≥ 8/10 testeurs réalisent une demande sans aide |
| S5–S8 | Collecte des repères (parallèle) | 500+ repères géolocalisés, alias FR/arabe | Échantillon de 50 repères vérifié sur place |
| S8–S16 | **Phase 5 — Backend + tests automatisés** | API, attribution, machine d'états, module SMS (simulateur), admin API | Couverture de tests ≥ 80 % sur `rides`/`matching`/`sms` ; test de concurrence : 0 double attribution sur 10 000 acceptations simultanées simulées |
| S10–S18 | **Phase 6 — Application client** | APK de test (canal interne Play Store) | Parcours complet sur 3 téléphones d'entrée de gamme ; taille < 20 Mo |
| S12–S20 | **Phase 7 — Application chauffeur** | APK de test | Réception d'offre < 3 s en data ; consommation batterie mesurée sur 8 h |
| S12–S18 | **Phase 8 — Tableau de bord** | Admin web | Validation de chauffeurs, tarifs, zones, suivi des courses, exports |
| S16–S21 | **Phase 9 — Mode SMS réel** | Passerelle SIM en service, parseur entraîné sur vrais messages | ≥ 90 % des SMS de test bien interprétés ; 0 double attribution ; délai médian demande→confirmation mesuré |
| S20–S24 | **Phase 10a — Sécurité et tests terrain** | Revue sécurité, test de charge, bêta fermée (30 chauffeurs, 100 clients) | Aucune faille critique ouverte ; taux de courses servies ≥ 60 % dans la zone pilote |
| S24–S26 | **Phase 10b — Déploiement** | Production, supervision, sauvegardes testées, procédures | Restauration de sauvegarde réussie ; astreinte définie |
| S26–S34 | **Pilote public** | Tableau de bord des indicateurs | Indicateurs cibles (à confirmer) : temps d'attribution médian < 3 min ; ≥ 70 % des demandes servies ; note moyenne ≥ 4/5 ; décision sur le modèle économique |

Jalons de décision (go/no-go) : **fin S4** (faisabilité terrain/légale), **fin S8**
(maquettes et budget confirmés), **fin S24** (bêta), **fin S34** (généralisation).
