# 2. Mode Internet vs mode SMS : ce qui est réellement possible

## 2.1 Le principe de base (à ne jamais oublier)

Une application mobile **ne peut pas** parler directement à d'autres téléphones par le réseau
cellulaire. Sans Internet, les seuls canaux qui existent entre un téléphone et notre serveur
sont ceux que fournit l'opérateur :

| Canal | Fonctionne sans data ? | Fonctionne sur téléphone simple ? | Nécessite |
|---|---|---|---|
| **SMS** | Oui | Oui | Un numéro qui reçoit les SMS et les transmet au serveur (passerelle) |
| **USSD** (`*123#`) | Oui | Oui | Un contrat avec **chaque** opérateur (ou un agrégateur agréé) |
| **Appel vocal** | Oui | Oui | Un numéro + soit un opérateur humain, soit un serveur vocal (IVR) |
| **« Bip » / appel manqué** | Oui | Oui | Une ligne capable de détecter les appels entrants (passerelle SIM ou ligne opérateur) |
| **Internet (data)** | Non | Non | Forfait data |

Tout le « mode sans Internet » repose donc sur une **passerelle** entre le réseau de
l'opérateur et notre serveur. C'est une infrastructure à payer et, pour un service
professionnel, un contrat à signer.

## 2.2 Matrice de faisabilité

| Fonction | Immédiat (sans accord) | Passerelle SMS nécessaire | Accord Airtel/Moov nécessaire | Impossible sans canal |
|---|---|---|---|---|
| Application avec data : carte, demande, suivi, avis | ✅ | | | |
| Notifications push (FCM) | ✅ (avec data) | | | |
| Appel client ↔ chauffeur (appel classique) | ✅ | | | |
| Recevoir une demande de course par SMS | | ✅ | ✅ pour un numéro court ou un volume pro | |
| Envoyer des SMS de confirmation/offre | | ✅ | ✅ pour un nom d'expéditeur (Sender ID) et des volumes | |
| Numéro court mémorisable (ex. 1234) | | | ✅ (+ ARCEP Tchad, à vérifier) | |
| Menus USSD `*xxx#` | | | ✅ avec chaque opérateur | |
| Vérification OTP du numéro | | ✅ | | |
| Masquage des numéros (appel via numéro intermédiaire) | | | ✅ (plateforme voix) | |
| Suivi en direct du chauffeur sur carte, sans data | | | | ❌ Impossible (sauf positions envoyées par SMS, coûteuses et espacées) |
| Position GPS du client sans data | ✅ Le GPS fonctionne sans data **si l'application est installée** ; le client l'envoie ensuite par SMS | | | |
| Demande depuis un téléphone simple sans GPS | Repère texte uniquement (« Grand marché ») | ✅ | | Position GPS ❌ |
| Paiement mobile intégré | | | ✅ API Airtel Money / Moov Money | |

## 2.3 Options de passerelle SMS (comparées)

| Option | Description | Délai | Coût | Fiabilité | Verdict |
|---|---|---|---|---|---|
| **A. Passerelle SIM** | 2 smartphones Android (1 SIM Airtel, 1 SIM Moov) ou un modem GSM multi-SIM, connectés au serveur ; une application passerelle relaie SMS entrants/sortants via HTTPS | Jours | Matériel ≈ 150–300 k FCFA + SMS au tarif grand public | Moyenne : débit limité (quelques SMS/min par SIM), risque de blocage SIM par l'opérateur, dépend de l'électricité et du réseau du local | **Pilote uniquement**, volume faible, à déclarer à l'opérateur |
| **B. Agrégateur international** (Infobip, etc.) | API HTTP ; envoi A2P vers le Tchad | Semaines | 30–70 FCFA/SMS (estimation non vérifiée) | Bonne en envoi ; **réception (numéro entrant local) rarement disponible au Tchad** | Utile pour l'envoi (OTP, confirmations) en complément |
| **C. Contrat direct opérateurs** (SMPP / API + numéro court) | Connexion SMPP avec Airtel et avec Moov, numéro court identique sur les deux réseaux si possible | Mois (négociation, ARCEP) | Frais d'ouverture + tarif négocié par SMS (à obtenir) | Élevée, débit pro, Sender ID | **Cible pour la production** |
| **D. Agrégateur local agréé** (s'il en existe à N'Djamena) | Un intermédiaire déjà connecté aux deux opérateurs | Semaines | Variable | Bonne | À rechercher : peut raccourcir fortement le délai de C |

**Recommandation** : démarrer le pilote avec **A** (envoi + réception) éventuellement doublé de
**B** pour l'envoi, tout en lançant immédiatement la négociation **C/D**. Le code serveur
utilise une interface `SmsProvider` unique pour pouvoir changer de passerelle sans réécrire
la logique métier.

## 2.4 Parcours détaillés sans Internet

### 2.4.1 Comment un client sans Internet contacte le serveur

**Cas 1 — Client avec l'application installée, sans data**

1. L'application détecte l'absence de connexion et propose « Demander par SMS ».
2. Elle lit la position GPS (sans data, le premier repérage peut prendre 30 s à quelques minutes).
3. Elle compose un SMS structuré et compact, par exemple :
   `M1 C 13.8291 20.8324 D 412 R` → *version 1, Clando, position GPS, Destination = repère n° 412, Retour attendu*.
4. L'application **ouvre l'application SMS du téléphone avec le message pré-rempli** ;
   le client appuie sur « Envoyer ».
   *Pourquoi pas d'envoi automatique ? Google Play restreint fortement les permissions
   `SEND_SMS`/`READ_SMS` aux applications SMS par défaut. Le pré-remplissage évite ce blocage.*
5. Le SMS part vers le numéro de la plateforme (au tarif SMS normal pour le client, sauf
   accord de gratuité avec l'opérateur).

**Cas 2 — Téléphone simple ou client sans application**

Le client envoie un SMS en texte libre, avec un format très simple affiché sur les
autocollants des véhicules et les affiches :

```
CLANDO grand marche > hopital
TAXI rond point > universite
RAKCHA mosquee X
```

- Le serveur reconnaît le type par mot-clé (`TAXI`, `RAKCHA`, `CLANDO`, aussi leurs variantes
  et abréviations `T`, `R`, `C`) et cherche les repères dans la **base de repères d'Abéché**
  (recherche approchée tolérante aux fautes, sans accents, translittérations arabes courantes).
- Si un repère est ambigu, le serveur répond :
  `1=Grand marche 2=Marche a betail. Repondez 1 ou 2`.
- Si rien n'est reconnu : `Lieu non trouve. Un operateur va vous appeler.` → la demande
  apparaît dans la console de l'opérateur de régulation.

**Cas 3 — Appel ou « bip »** (pour les personnes qui n'écrivent pas)

Le client appelle (ou bipe) le numéro de la plateforme ; l'opérateur humain crée la demande
pour lui depuis le tableau de bord. Le reste du parcours (SMS de confirmation, appel du
chauffeur) est identique.

### 2.4.2 Comment un chauffeur sans Internet reçoit une demande

1. Le chauffeur se déclare disponible par SMS : `DISPO marche central` (ou depuis son
   application en mode SMS, qui pré-remplit la position GPS). Sa position est celle du repère
   déclaré, datée ; elle « vieillit » et expire après un délai configurable (ex. 30 min).
2. Le serveur lui envoie une offre :
   `Course K7Q4: CLANDO Grand marche > Hopital, 500F. Client a 600m. Repondez OUI K7Q4 (90s)`
3. Le chauffeur répond `OUI K7Q4` (ou `NON K7Q4`). Le code de course est aléatoire et à usage
   unique ; seule une réponse **venant du numéro enregistré du chauffeur** est acceptée.
4. Le serveur attribue la course de façon atomique (voir doc 03) :
   - au chauffeur : `OK K7Q4. Client: Amina 66xxxxxx. Appelez-le.`
   - au client : `Chauffeur: Moussa, moto rouge TD-1234, 4.7★. Il vous appelle. Code: 58`
5. Les autres chauffeurs ayant répondu trop tard reçoivent : `K7Q4 deja attribuee.`
   (envoyé seulement s'ils ont répondu — pour économiser les SMS).
6. Fin de course : le chauffeur envoie `FIN K7Q4` (facultatif ; sinon clôture automatique
   après délai et confirmation lors de la prochaine connexion).

### 2.4.3 Comment les deux parties communiquent

- **Appel vocal classique** dans tous les cas (le plus fiable et le plus naturel localement).
- En V1, le numéro est communiqué uniquement **après** attribution et uniquement aux deux
  parties de la course. Le masquage complet des numéros nécessite une plateforme voix
  opérateur (phase ultérieure).
- Le **code de prise en charge** (2 chiffres) permet au client de vérifier qu'il monte avec le
  bon chauffeur.

## 2.5 Économie du mode SMS

Les montants ci-dessous sont des **hypothèses à remplacer par des devis**.

| Étape | SMS sortants (payés par la plateforme) |
|---|---|
| Offre au chauffeur (séquentielle, 1 à 2 en moyenne) | 1,5 |
| Confirmation chauffeur | 1 |
| Confirmation client | 1 |
| Désambiguïsation (≈ 30 % des cas) | 0,3 |
| **Total moyen** | **≈ 3,8 SMS** |

À 25 FCFA/SMS : **≈ 95 FCFA par course SMS**, à comparer à un prix de course clando de
quelques centaines de FCFA (à relever, H5). Conséquences de conception :

- Offres **séquentielles** (un chauffeur SMS à la fois), jamais de diffusion massive par SMS.
- Pas de SMS de « course refusée par les autres ».
- Messages en **alphabet GSM-7 (latin sans accents)** : 160 caractères par SMS. Un SMS en
  caractères arabes passe en UCS-2 : **70 caractères**, donc souvent 2 SMS facturés.
- Encourager les chauffeurs actifs à passer en data : un petit forfait data quotidien
  coûte probablement moins cher que les SMS d'une journée de courses (à vérifier avec les
  tarifs réels). Une subvention de forfait data aux meilleurs chauffeurs peut être rentable.

## 2.6 Sécurité propre au canal SMS

| Menace | Contre-mesure |
|---|---|
| Fausses demandes / canulars | Limite de demandes par numéro (ex. 3 actives / heure) ; réputation du numéro (annulations, absences) ; liste noire ; première demande d'un nouveau numéro rappelée par l'opérateur si signalements |
| Usurpation du numéro d'expéditeur | Les numéros entrants via opérateur sont fiables dans la majorité des cas, mais pas garantis. Codes de course aléatoires à usage unique ; actions sensibles (acceptation) uniquement depuis le numéro enregistré **et** avec le code ; aucune donnée personnelle envoyée à un numéro non lié à une course |
| Doublons (client qui renvoie le SMS faute de réponse) | Clé d'idempotence : (numéro, empreinte du contenu normalisé, fenêtre de 5 min) → renvoie le statut de la demande existante au lieu d'en créer une nouvelle |
| SMS arrivant en retard | Chaque offre a une date d'expiration ; une réponse après expiration ou après attribution reçoit « déjà attribuée/expirée » et ne modifie rien |
| Double attribution | Transition d'état conditionnelle en base (un seul `UPDATE … WHERE status = 'OFFERED'` peut réussir) — voir doc 03 |
| Rejeu d'une ancienne réponse | Code de course lié à l'offre, invalide après usage ou expiration |
| Inondation de la passerelle | Limitation de débit par numéro et globale ; file d'attente ; alerte au-delà d'un seuil de coût journalier |
| Passerelle SIM compromise | Webhook signé (HMAC) + liste d'IP autorisées ; jeton révocable par appareil |
