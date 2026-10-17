#!/usr/bin/env python3
"""Execute le code VBA de SG dans LibreOffice (mode de compatibilite VBA).

Le code metier (modules .bas) est injecte dans un classeur Calc vierge, les
feuilles de donnees sont construites par modSetup, puis la suite de tests
modTests.SG_ExecuterTests est lancee. Le rapport est affiche et enregistre.

Limites : LibreOffice n'est pas Excel. Les modules d'interface (modUI) et les
appels propres a Windows ne sont pas executes ici ; ils doivent etre valides
dans Excel avec modTests (voir docs/07-plan-et-rapport-de-tests.md).

Usage :
    python3 tools/tests/run_vba_tests.py [--rapport chemin.txt] [--xlsx sortie.xlsx]
"""
from __future__ import annotations

import argparse
import os
import re
import socket
import subprocess
import sys
import tempfile
import time
from pathlib import Path

import uno  # fourni par LibreOffice
from com.sun.star.beans import PropertyValue

RACINE = Path(__file__).resolve().parents[2]
VBA_DIR = RACINE / "src" / "vba"
sys.path.insert(0, str(RACINE / "tools" / "licence_admin"))

# Modules exclus de l'execution LibreOffice (interface Excel uniquement).
EXCLUS = {"modUI.bas", "ThisWorkbook.bas"}


def pv(nom, val):
    p = PropertyValue()
    p.Name = nom
    p.Value = val
    return p


def port_libre() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def demarrer_office():
    port = port_libre()
    profil = tempfile.mkdtemp(prefix="sg-lo-")
    proc = subprocess.Popen(
        [
            "soffice",
            "--headless",
            "--invisible",
            "--norestore",
            "--nologo",
            f"-env:UserInstallation=file://{profil}",
            f"--accept=socket,host=127.0.0.1,port={port};urp;",
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(120):
        try:
            ctx = resolver.resolve(f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext")
            return proc, ctx
        except Exception:
            time.sleep(0.5)
    proc.kill()
    raise RuntimeError("LibreOffice n'a pas demarre")


def code_module(chemin: Path) -> tuple[str, str]:
    texte = chemin.read_text(encoding="utf-8")
    nom = re.search(r'Attribute VB_Name = "([^"]+)"', texte).group(1)
    lignes = [l for l in texte.splitlines() if not l.startswith("Attribute ")]
    return nom, "Option VBASupport 1\n" + "\n".join(lignes) + "\n"


def preparer_fixtures() -> dict[str, str]:
    """Licences de test signees avec une cle RSA ephemere (jamais conservee)."""
    import sg_licence as L

    cle = L.generer_cle(2048)
    pub = cle.public_key()
    autre = L.generer_cle(2048)
    mod_hex = L.module_hex(pub)
    empreinte = "a" * 64

    def lic(**kw):
        base = dict(
            licence_id="LIC-TEST-0001",
            serie="SG-TEST-0001",
            titulaire="Groupe Scolaire Fictif Les Acacias",
            etablissements="ETB-A",
            edition="STANDARD",
            type_licence="ANNUELLE",
            emise_le="2026-09-01",
            expire_le="2027-08-31",
            maintenance_jusqu_au="2027-08-31",
            version_majeure_max="0",
            installations_max="1",
            nb_etablissements_max="1",
            modules="ELEVES;INSCR;STRUCT;ANNEE",
            empreinte_installation=empreinte,
            grace_jours="15",
        )
        base.update(kw)
        return L.emettre_texte(cle, L.Licence(**base))

    valide = lic()
    falsifiee = valide.replace("MODULES=ELEVES;INSCR;STRUCT;ANNEE", "MODULES=ELEVES;INSCR;STRUCT;ANNEE;FINANCES")
    autre_cle = L.emettre_texte(autre, L.Licence(**{**L.Licence.champs_exemple(), "empreinte_installation": empreinte}))
    perpetuelle = lic(licence_id="LIC-TEST-0002", type_licence="PERPETUELLE", expire_le="", maintenance_jusqu_au="2027-01-31")
    perp_v_sup = lic(licence_id="LIC-TEST-0003", type_licence="PERPETUELLE", expire_le="", version_majeure_max="-1")
    toute_install = lic(licence_id="LIC-TEST-0004", empreinte_installation="*", type_licence="ESSAI", expire_le="2026-10-31")
    revoc = L.emettre_revocations_texte(cle, ["LIC-TEST-0001"], "2026-10-01")
    return {
        "MODULUS": mod_hex,
        "EMPREINTE": empreinte,
        "LIC_VALIDE": valide,
        "LIC_FALSIFIEE": falsifiee,
        "LIC_AUTRE_CLE": autre_cle,
        "LIC_PERPETUELLE": perpetuelle,
        "LIC_PERP_VERSION": perp_v_sup,
        "LIC_ESSAI_TOUTE_INSTALL": toute_install,
        "REVOCATIONS": revoc,
    }


def controle_croise_demande(sp, serie: str) -> str:
    """La demande d'activation produite par le VBA doit etre acceptee par l'outil editeur."""
    import sg_licence as L

    script = sp.getScript("vnd.sun.star.script:Standard.modTests.SG_TestDemande?language=Basic&location=document")
    texte = str(script.invoke(("", serie), (), ())[0])
    try:
        champs = L.lire_demande(texte)
        ok = champs.get("ETABLISSEMENTS") == "ETB-A;ETB-B" and len(champs.get("EMPREINTE", "")) == 64
        falsifiee_refusee = False
        try:
            L.lire_demande(texte.replace("ETB-A;ETB-B", "ETB-A;ETB-Z"))
        except L.ErreurLicence:
            falsifiee_refusee = True
        ok = ok and falsifiee_refusee
    except L.ErreurLicence as exc:
        return f"CONTROLE CROISE DEMANDE VBA -> PYTHON : ECHEC ({exc})\n"
    return "CONTROLE CROISE DEMANDE VBA -> PYTHON : " + ("OK" if ok else "ECHEC") + "\n"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--rapport", default=str(RACINE / "docs" / "rapport-tests-vba.txt"))
    ap.add_argument("--xlsx", help="Enregistrer le classeur de test (structure + donnees fictives)")
    ap.add_argument("--appel", default="modTests.SG_ExecuterTests")
    ap.add_argument("--filtre", default="", help="Ne lancer que les tests dont le nom contient ce texte")
    args = ap.parse_args()

    fixtures = preparer_fixtures()
    if os.environ.get("SG_FIXTURES_DUMP"):
        import json

        Path(os.environ["SG_FIXTURES_DUMP"]).write_text(json.dumps(fixtures), encoding="utf-8")
    proc, ctx = demarrer_office()
    try:
        desk = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desk.loadComponentFromURL("private:factory/scalc", "_blank", 0, ())
        libs = doc.BasicLibraries
        if not libs.hasByName("Standard"):
            libs.createLibrary("Standard")
        libs.VBACompatibilityMode = True
        lib = libs.getByName("Standard")
        for f in sorted(VBA_DIR.glob("*.bas")):
            if f.name in EXCLUS:
                continue
            nom, code = code_module(f)
            lib.insertByName(nom, code)
        module, fonction = args.appel.split(".")
        sp = doc.getScriptProvider()
        script = sp.getScript(f"vnd.sun.star.script:Standard.{module}.{fonction}?language=Basic&location=document")
        debut = time.time()
        # Les fixtures sont transmises en parametre (cle=valeur separes par
        # le caractere 30) : rien n'est ecrit dans le classeur.
        serie = "\x1e".join(f"{k}\x1f{v}" for k, v in fixtures.items())
        params = (args.filtre, serie)
        resultat = script.invoke(params, (), ())[0]
        duree = time.time() - debut
        rapport = str(resultat) + f"\nDuree d'execution (LibreOffice) : {duree:.1f} s\n"
        if fonction == "SG_ExecuterTests" and not args.filtre:
            rapport += controle_croise_demande(sp, serie)
        print(rapport)
        Path(args.rapport).write_text(rapport, encoding="utf-8")
        if args.xlsx:
            doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(args.xlsx)), (pv("FilterName", "Calc MS Excel 2007 XML"),))
        doc.close(True)
        return 0 if "ECHECS : 0" in rapport and "ECHEC (" not in rapport and ": ECHEC" not in rapport else 1
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=10)
        except subprocess.TimeoutExpired:
            proc.kill()


if __name__ == "__main__":
    sys.exit(main())
