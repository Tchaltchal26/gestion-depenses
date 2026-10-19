# Tester SG sur votre PC Windows (disque G:)

## 1. Récupérer le projet sur G:

**Avec Git** (invite de commandes) :
```
G:
git clone -b ccr-2c29e05a-0tgbo6 https://github.com/tchaltchal26/gestion-depenses.git
cd gestion-depenses\gestion-scolaire
```

**Sans Git** : sur GitHub, ouvrir le dépôt `tchaltchal26/gestion-depenses`, choisir la branche
`ccr-2c29e05a-0tgbo6`, *Code > Download ZIP*, puis extraire le ZIP dans `G:\`.

## 2. Autoriser l'import des modules (une seule fois)

Excel : *Fichier > Options > Centre de gestion de la confidentialité > Paramètres du Centre de
gestion de la confidentialité > Paramètres des macros* → cocher
**Accès approuvé au modèle d'objet du projet VBA**. (À décocher après les essais.)

## 3. Lancer la suite de tests dans Excel

Dans PowerShell, depuis `G:\gestion-depenses\gestion-scolaire` :
```
powershell -ExecutionPolicy Bypass -File installer\Installer-SG.ps1 -Tests -Destination G:\SG_Tests.xlsm
```
Résultat attendu (après quelques minutes) : `VERIFICATIONS : 235 | REUSSIES : 235 | ECHECS : 0`.
Le rapport est enregistré dans `G:\SG_Tests.rapport-tests.txt` : **envoyez-le-moi**, surtout s'il
contient des lignes `[ECHEC]`.

## 4. Essayer l'application avec la démonstration

```
powershell -ExecutionPolicy Bypass -File installer\Installer-SG.ps1 -Demo -Destination G:\SG_Demo.xlsm
```
Créez l'administrateur quand Excel le demande, puis essayez les écrans de l'ACCUEIL.
Comptes de démonstration : `dir.demo`, `secr.demo`, `prof.demo`, `audit.demo`
(mot de passe `Demo2026x`).

Check-list à vérifier : connexion masquée · fiche élève (création, doublon, modification) ·
liste · inscription et changement de classe · tableau de bord · contrôle qualité · journal et
vérification d'intégrité · sauvegarde dans un dossier · certificat PDF · demande d'activation.
Notez tout message d'erreur (capture d'écran) et transmettez-le.
