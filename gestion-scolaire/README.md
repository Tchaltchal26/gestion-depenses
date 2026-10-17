# SG — Logiciel universel de gestion scolaire (Excel / VBA)

Prototype fonctionnel (phase 3) d'un logiciel de gestion scolaire commercialisable, avec Excel comme
interface, VBA pour les règles métier, une traçabilité chaînée et un système de licences signées.

> Ce dossier est indépendant de l'application Django « gestion des dépenses » présente à la racine
> du dépôt.

## Contenu

| Dossier | Contenu |
|---|---|
| `src/vba/` | 20 modules VBA (métier, sécurité, audit, licences, interface, tests, démonstration) |
| `tools/licence_admin/sg_licence.py` | Outil **éditeur** : clés, émission, transfert, révocation, registre |
| `tools/tests/` | Lanceur des tests VBA (LibreOffice) et tests Python de l'outil de licences |
| `installer/Installer-SG.ps1` | Construction du classeur `.xlsm` sous Excel Windows |
| `docs/` | Architecture, modèle de données, licences, cadrage, avancement, installation, tests |

## Démarrage rapide

```bash
# Tests de l'outil de licences
python3 -m pytest tools/tests/test_licence.py -q
# Tests du code VBA métier (LibreOffice requis)
python3 tools/tests/run_vba_tests.py
```

Sous Windows : `powershell -ExecutionPolicy Bypass -File installer\Installer-SG.ps1 -Demo`
(voir `docs/06-guide-installation.md`).

## Documentation

1. [Architecture, sécurité, limites, sauvegarde, plan](docs/01-architecture.md)
2. [Modèle de données](docs/02-modele-donnees.md)
3. [Licences et protection](docs/03-licences.md)
4. [Questions de cadrage et hypothèses](docs/04-questions-cadrage.md)
5. [État d'avancement et limites connues](docs/05-etat-avancement-et-limites.md)
6. [Guides d'installation et d'exploitation](docs/06-guide-installation.md)
7. [Plan et rapport de tests](docs/07-plan-et-rapport-de-tests.md)

Toutes les données de démonstration et de test sont fictives.
