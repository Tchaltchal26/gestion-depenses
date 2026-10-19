# Plateforme de transport urbain — Tchad (pilote : Abéché)

Ce dossier contient les livrables du projet, phase par phase. Il est provisoirement
hébergé dans le dépôt `gestion-depenses` ; il est recommandé de le déplacer dans un
dépôt dédié avant la phase 5 (développement backend).

## État d'avancement

| Phase | Contenu | État |
|------:|---------|------|
| 1 | Analyse, hypothèses, risques, architecture, budget, calendrier, cahier des charges initial | **Livré (ce dossier)** — à valider |
| 2 | Cahier des charges complet | Non commencé (dépend des réponses aux questions bloquantes) |
| 3 | Architecture détaillée, modèle de données, contrats d'API | Non commencé (ébauche en phase 1) |
| 4 | Maquettes client / chauffeur / admin | Non commencé |
| 5 | Backend + tests automatisés | **Prototype fonctionnel** : 54 tests passent ; SMS et notifications encore simulés — voir [backend/README.md](backend/README.md) |
| 6 → 10 | Applications, admin, SMS réel, tests terrain, déploiement | Non commencé |

Le backend existe sous forme de prototype testé, mais **rien n'est encore en service réel** :
aucun SMS réel, aucune application mobile, données de repères fictives. Les chiffres de budget
sont des estimations avec hypothèses explicites, pas des devis.

## Documents de la phase 1

1. [Vision produit et hypothèses à vérifier](docs/phase-1/01-vision-et-hypotheses.md)
2. [Mode Internet vs mode SMS : ce qui est réellement possible](docs/phase-1/02-internet-vs-sms.md)
3. [Architecture technique, modèle de données et algorithme d'attribution](docs/phase-1/03-architecture.md)
4. [Budget prévisionnel en FCFA — trois scénarios](docs/phase-1/04-budget.md)
5. [Calendrier de développement](docs/phase-1/05-calendrier.md)
6. [Cahier des charges initial](docs/phase-1/06-cahier-des-charges-initial.md)
7. [Sécurité, conformité et cadre réglementaire](docs/phase-1/07-securite-et-reglementation.md)
8. [Questions bloquantes](docs/phase-1/08-questions-bloquantes.md)
