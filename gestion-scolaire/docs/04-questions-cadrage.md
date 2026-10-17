# 04 — Questions de cadrage et hypothèses retenues

Le développement n'a pas été bloqué : chaque question est accompagnée de l'**hypothèse par défaut**
appliquée dans le prototype. Les questions marquées **[structurant]** conditionnent l'architecture
des phases suivantes et doivent être confirmées avant le module concerné.

| # | Question | Hypothèse appliquée | Impact si différent |
|---|---|---|---|
| 1 | **[structurant]** Pays, système éducatif et réglementation (protection des données, archives, comptabilité) ? | Système francophone générique ; aucune règle officielle codée en dur | Barèmes, mentions, conditions de passage, durées de conservation, plan comptable |
| 2 | Devise ? | `XOF` (modifiable par établissement) | Formats monétaires, arrondis (module Finances) |
| 3 | **[structurant]** Systèmes de notation par cycle (sur 10, 20, 100, lettres, compétences) ? Moyennes pondérées ? Rattrapage ? | Maternelle/primaire par compétences, collège/lycée sur 20 avec coefficients ; note manquante ≠ zéro | Modèle `T_Notes` et règles de calcul |
| 4 | **[structurant]** Versions d'Excel utilisées (Windows/Mac, 32/64 bits, 365 Web) ? | Excel 2016+ Windows de bureau | Mac : pas d'installateur automatique ; Web : VBA impossible |
| 5 | **[structurant]** Nombre d'établissements et mode (indépendant / centralisé) ? | Indépendant par défaut ; centralisé disponible | Centralisé + plusieurs postes ⇒ base serveur |
| 6 | **[structurant]** Nombre d'utilisateurs **simultanés** en écriture ? | 1 | ≥ 2 ⇒ migration de `modDAL` vers Access/SQL |
| 7 | Besoins hors ligne (coupures, zones sans Internet) ? | Fonctionnement 100 % hors ligne, activation par fichier | Activation en ligne possible en complément |
| 8 | Volume d'archives (élèves par an, années à conserver) ? | ≤ 3 000 élèves actifs, 10 ans dans le même fichier | Au-delà : archivage annuel en fichiers séparés ou base SQL |
| 9 | Budget et délais ? | Développement par modules, prototype d'abord | Ordre des phases 4a-4d |
| 10 | **[structurant]** Modules prioritaires après le noyau ? | 1) Finances (frais, paiements, reçus) 2) Notes et bulletins 3) Présences | Ordre du plan |
| 11 | Modèle commercial : annuel, perpétuel, par établissement, par nombre d'élèves ? | Annuelle par établissement + perpétuelle par version majeure | Champs de licence (ex. `MAX_ELEVES`) |
| 12 | Nombre d'installations autorisées par licence ? | 1 poste (champ `INSTALLATIONS_MAX` prêt pour davantage) | Comptage en ligne nécessaire pour > 1 |
| 13 | Serveur de licences souhaité (hébergement, budget, disponibilité) ? | Non : activation hors ligne signée | Développement du serveur (§ 6 de `03-licences.md`) |
| 14 | Modalités de paiement et de renouvellement (revendeurs, délais, grâce) ? | Grâce de 15 jours, alertes 30/15/7 jours | Paramètres de licence |
| 15 | Documents officiels exigés (modèles de bulletins, certificats, PV) et mentions légales ? | Certificat de scolarité générique | Modèles par pays / ministère |
| 16 | Photos des élèves et données sensibles (santé, religion) autorisées ? | Non stockées | Consentements, droits spécifiques, chiffrement |
| 17 | Comptabilité générale réglementaire attendue ou simple suivi de trésorerie ? | Suivi de trésorerie (distinct de la comptabilité générale) | Export vers un logiciel comptable agréé |
| 18 | Langues de l'interface (français, anglais, arabe — écriture de droite à gauche) ? | Français | Textes externalisés, mise en page RTL |
