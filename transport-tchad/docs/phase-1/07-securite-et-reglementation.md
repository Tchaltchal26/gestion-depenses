# 7. Sécurité, conformité et cadre réglementaire

## 7.1 Cadre réglementaire tchadien (état de vérification)

| Domaine | Ce qui a été vérifié (sources publiques, octobre 2026) | Ce qui reste à vérifier | Qui consulter |
|---|---|---|---|
| Protection des données personnelles | **Loi n° 007/PR/2015 du 10 février 2015** sur la protection des données à caractère personnel ; décret d'application **n° 075/PR/2019**. Autorité : **ANSICE** (Agence Nationale de Sécurité Informatique et de Certification Électronique), créée par la **loi n° 006/PR/2015**. Selon les analyses publiques : déclaration préalable des traitements auprès de l'ANSICE, autorisation pour certaines catégories (biométrie, numéros d'identification nationaux, infractions…), sanctions administratives (amendes de l'ordre de 1 à 10 M FCFA selon une source) et pénales | Formalités exactes ; statut des **photos de pièces d'identité et de permis** (numéro d'identification national ⇒ autorisation ?) ; conditions de **transfert hors du Tchad** (hébergement en Europe) ; obligation de DPO et de notification de violation (les sources secondaires se contredisent) | Avocat à N'Djamena ; ANSICE directement |
| Télécommunications | Les deux opérateurs mobiles sont **Airtel Tchad** et **Moov Africa Tchad** (Moov a remplacé Tigo en décembre 2020). Régulateur : **ARCEP Tchad** | Procédure et coût d'attribution d'un **numéro court** ; régime des services à valeur ajoutée (SVA) ; légalité d'une passerelle SIM à usage commercial | ARCEP Tchad ; services Entreprises Airtel et Moov |
| Transport | — (aucune source fiable trouvée) | Existence d'un cadre pour les plateformes de mise en relation ; licences taxi ; statut des **motos-taxis (clandos)** et **rakchas** à Abéché ; port du casque ; assurances obligatoires | Ministère des Transports ; mairie d'Abéché ; gouvernorat du Ouaddaï |
| Société et fiscalité | — | Forme juridique, guichet unique de création d'entreprise, TVA sur les abonnements, obligations comptables | Expert-comptable / guichet unique (ANIE, à confirmer) |
| Paiements mobiles (V2) | — | Statut de marchand / agrégateur ; règles BEAC/COBAC sur la monnaie électronique ; un portefeuille interne peut constituer de l'émission de monnaie électronique ⇒ à éviter sans avis juridique | Airtel Money, Moov Money ; avocat |

**Règle du projet** : aucune donnée réelle d'utilisateur n'est collectée (pilote compris)
avant que les formalités ANSICE aient été clarifiées par écrit.

## 7.2 Mesures de sécurité

| Domaine | Mesures |
|---|---|
| Authentification | OTP à 6 chiffres, valide 5 min, 5 essais max, limite par numéro et par IP ; JWT 15 min + jeton de rafraîchissement rotatif révocable ; 2FA (TOTP) obligatoire pour les admins |
| Autorisations | Contrôle d'accès par rôle **et** par ressource (un chauffeur ne voit que ses courses) ; permissions fines pour les admins ; principe du moindre privilège |
| Communications | HTTPS/WSS uniquement (TLS 1.2+), HSTS ; webhooks SMS signés HMAC + liste d'IP |
| Numéros de téléphone | Chiffrés au repos (chiffrement applicatif) + empreinte HMAC pour la recherche ; révélés seulement aux deux parties d'une course active ; masqués dans l'admin pour les rôles sans besoin |
| Positions GPS | Envoyées seulement quand le chauffeur est disponible ou en course ; visibles par le client **uniquement pendant sa course** ; positions « à proximité » arrondies et anonymes ; historique détaillé purgé après 30 jours (agrégats conservés) |
| Validation | Validation stricte de toutes les entrées (DTO NestJS + class-validator), requêtes paramétrées (ORM), limites de taille |
| Abus | Limitation de débit (Redis), réputation des numéros, liste noire, détection de comptes multiples par appareil |
| Documents chauffeurs | Stockage objet privé chiffré, URL signées de courte durée, accès journalisé |
| Journaux | Pas de numéros, de codes OTP ou de positions précises en clair dans les logs ; identifiants pseudonymes |
| Secrets | Hors dépôt Git, rotation, accès restreint |
| Infrastructure | Pare-feu (seuls 443/22 ouverts, SSH par clé), mises à jour automatiques de sécurité, conteneurs non root, analyse des dépendances en CI |
| Sauvegardes et reprise | Quotidiennes + WAL, chiffrées, autre région ; restauration testée chaque mois ; procédure écrite de reprise (RTO 4 h) |

## 7.3 Politique de conservation (proposition à valider juridiquement)

| Donnée | Durée |
|---|---|
| Comptes actifs | Durée de la relation |
| Compte supprimé à la demande | Effacement/anonymisation sous 30 jours, sauf obligation légale |
| Courses (sans position détaillée) | 3 ans (litiges, statistiques) — à valider |
| Traces GPS détaillées | 30 jours |
| Contenu des SMS | 90 jours, puis seules les métadonnées (date, coût, type) |
| Documents chauffeurs | Durée d'activité + 1 an |
| Journaux techniques | 90 jours |
| Journaux d'audit admin | 3 ans |

## 7.4 Procédure de gestion des incidents de sécurité (résumé)

1. **Détection** (alerte, signalement) → enregistrement dans le registre des incidents.
2. **Qualification** sous 1 h : gravité, données concernées.
3. **Confinement** : révocation des jetons/clés, blocage de comptes, coupure du composant.
4. **Éradication et restauration** depuis une sauvegarde saine si nécessaire.
5. **Notification** : utilisateurs concernés et ANSICE si la loi ou la prudence l'exige
   (obligation exacte à confirmer).
6. **Retour d'expérience** sous 7 jours, actions correctives suivies.

## 7.5 Sécurité des personnes (spécifique au transport)

Vérification des chauffeurs (documents + rencontre physique pendant le pilote), code de
prise en charge, bouton de signalement, suspension immédiate possible par l'opérateur,
procédure avec la police locale à définir, partage de la course (V2).
