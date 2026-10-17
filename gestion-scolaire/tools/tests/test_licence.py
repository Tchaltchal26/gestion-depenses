"""Tests de l'outil editeur de licences (sg_licence.py).

Lancer : python3 -m pytest tools/tests/test_licence.py -q
"""
from __future__ import annotations

import hashlib
import re
import sqlite3
import sys
from pathlib import Path

import pytest

RACINE = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(RACINE / "tools" / "licence_admin"))
import sg_licence as L  # noqa: E402


@pytest.fixture(scope="module")
def cle():
    return L.generer_cle(2048)


def licence(**kw) -> L.Licence:
    champs = L.Licence.champs_exemple()
    champs.update(kw)
    return L.Licence(**champs)


def test_emission_et_verification(cle):
    texte = L.emettre_texte(cle, licence())
    champs = L.verifier_texte(cle.public_key(), texte)
    assert champs["FORMAT"] == "SGLIC-1"
    assert champs["TYPE"] == "ANNUELLE"
    assert texte.splitlines()[-1].startswith("SIGNATURE=")


@pytest.mark.parametrize(
    "avant, apres",
    [
        ("MODULES=ELEVES", "MODULES=ELEVES;FINANCES"),
        ("EXPIRE_LE=2027-08-31", "EXPIRE_LE=2099-08-31"),
        ("TYPE=ANNUELLE", "TYPE=PERPETUELLE"),
        ("EMPREINTE_INSTALLATION=*", "EMPREINTE_INSTALLATION=" + "0" * 64),
    ],
)
def test_falsification_detectee(cle, avant, apres):
    texte = L.emettre_texte(cle, licence())
    assert avant in texte
    with pytest.raises(L.ErreurLicence):
        L.verifier_texte(cle.public_key(), texte.replace(avant, apres))


def test_autre_cle_refusee(cle):
    autre = L.generer_cle(2048)
    with pytest.raises(L.ErreurLicence):
        L.verifier_texte(cle.public_key(), L.emettre_texte(autre, licence()))


def test_ligne_ajoutee_refusee(cle):
    texte = L.emettre_texte(cle, licence())
    lignes = texte.splitlines()
    lignes.insert(1, "MODULES=FINANCES")
    with pytest.raises(L.ErreurLicence):
        L.verifier_texte(cle.public_key(), "\n".join(lignes))


def test_fins_de_ligne_windows_acceptees(cle):
    texte = L.emettre_texte(cle, licence()).replace("\n", "\r\n")
    assert L.verifier_texte(cle.public_key(), "﻿" + texte)["LICENCE_ID"] == "LIC-EXEMPLE"


@pytest.mark.parametrize(
    "kw",
    [
        {"type_licence": "INCONNUE"},
        {"expire_le": "", "type_licence": "ANNUELLE"},
        {"modules": "ELEVES;MODULE_FANTOME"},
        {"empreinte_installation": "abc"},
        {"etablissements": "A;B", "nb_etablissements_max": "1"},
        {"titulaire": "Nom\nINJECTION=1"},
        {"expire_le": "2027-02-30"},
    ],
)
def test_validation_avant_signature(cle, kw):
    with pytest.raises((L.ErreurLicence, ValueError)):
        L.emettre_texte(cle, licence(**kw))


def test_perpetuelle_sans_expiration(cle):
    texte = L.emettre_texte(cle, licence(type_licence="PERPETUELLE", expire_le=""))
    assert "EXPIRE_LE=\n" in texte


def test_dates_expiration():
    import datetime as dt

    d = dt.date(2026, 9, 1)
    assert L.date_expiration("ANNUELLE", d, None) == "2027-08-31"
    assert L.date_expiration("MENSUELLE", d, None) == "2026-09-30"
    assert L.date_expiration("PERPETUELLE", d, None) == ""


def test_revocations(cle):
    texte = L.emettre_revocations_texte(cle, ["LIC-B", "LIC-A", "LIC-A"], "2026-10-01")
    champs = L.verifier_texte(cle.public_key(), texte)
    assert champs["LICENCES"] == "LIC-A;LIC-B"


def demande(empreinte="e" * 64) -> str:
    lignes = "\n".join(
        [
            "FORMAT=SGREQ-1",
            "INSTALL_ID=abcd",
            f"EMPREINTE={empreinte}",
            "TITULAIRE=Groupe fictif",
            "ETABLISSEMENTS=ETB-A",
            "VERSION=0.1.0",
            "DATE=2026-10-17",
        ]
    )
    return lignes + "\nCONTROLE=" + hashlib.sha256(lignes.encode()).hexdigest() + "\n"


def test_demande_activation():
    assert L.lire_demande(demande())["EMPREINTE"] == "e" * 64
    with pytest.raises(L.ErreurLicence):
        L.lire_demande(demande().replace("ETB-A", "ETB-Z"))


def test_registre_cycle_complet(tmp_path, cle):
    from cryptography.hazmat.primitives import serialization

    priv = tmp_path / "cle_privee.pem"
    priv.write_bytes(cle.private_bytes(serialization.Encoding.PEM, serialization.PrivateFormat.PKCS8, serialization.NoEncryption()))
    pub = tmp_path / "cle_publique.pem"
    pub.write_bytes(cle.public_key().public_bytes(serialization.Encoding.PEM, serialization.PublicFormat.SubjectPublicKeyInfo))
    reg = str(tmp_path / "registre.db")
    dem1 = tmp_path / "poste1.sgreq"
    dem1.write_text(demande("1" * 64), encoding="utf-8")
    sortie = tmp_path / "client.sglic"
    assert L.main(["--registre", reg, "emettre", "--privee", str(priv), "--titulaire", "Ecole fictive",
                   "--etablissements", "ETB-A", "--type", "ANNUELLE", "--modules", "ELEVES;INSCR",
                   "--demande", str(dem1), "--sortie", str(sortie), "--debut", "2026-09-01"]) == 0
    champs = L.verifier_texte(L.charger_cle_publique(pub), sortie.read_text(encoding="utf-8"))
    assert champs["EMPREINTE_INSTALLATION"] == "1" * 64
    assert champs["EXPIRE_LE"] == "2027-08-31"
    lid = champs["LICENCE_ID"]

    # Transfert vers un nouveau poste (panne materielle) : nouvelle licence, ancienne revoquee.
    dem2 = tmp_path / "poste2.sgreq"
    dem2.write_text(demande("2" * 64), encoding="utf-8")
    sortie2 = tmp_path / "client2.sglic"
    assert L.main(["--registre", reg, "transferer", lid, "--privee", str(priv), "--demande", str(dem2),
                   "--motif", "Disque dur en panne", "--sortie", str(sortie2)]) == 0
    champs2 = L.verifier_texte(L.charger_cle_publique(pub), sortie2.read_text(encoding="utf-8"))
    assert champs2["EMPREINTE_INSTALLATION"] == "2" * 64
    assert champs2["LICENCE_ID"] != lid
    assert champs2["EXPIRE_LE"] == champs["EXPIRE_LE"]  # memes droits

    db = sqlite3.connect(reg)
    assert db.execute("SELECT statut FROM licences WHERE licence_id=?", (lid,)).fetchone()[0] == "REVOQUEE"
    assert db.execute("SELECT remplace FROM licences WHERE licence_id=?", (champs2["LICENCE_ID"],)).fetchone()[0] == lid

    # Liste de revocation signee
    rev = tmp_path / "revocations.sgrev"
    assert L.main(["--registre", reg, "revocations", "--privee", str(priv), "--sortie", str(rev)]) == 0
    assert lid in L.verifier_texte(L.charger_cle_publique(pub), rev.read_text(encoding="utf-8"))["LICENCES"]

    # Journal du registre : integre, puis altere
    assert L.main(["--registre", reg, "journal"]) == 0
    db.execute("UPDATE journal SET details='{}' WHERE id=1")
    db.commit()
    assert L.main(["--registre", reg, "journal"]) == 2


def test_motif_obligatoire(tmp_path):
    reg = L.Registre(tmp_path / "r.db", "test")
    with pytest.raises(L.ErreurLicence):
        reg.changer_statut("LIC-X", "REVOQUEE", " ")


def test_integration_cle_publique(tmp_path, cle):
    source = (RACINE / "src" / "vba" / "modLicence.bas").read_text(encoding="utf-8")
    copie = tmp_path / "modLicence.bas"
    copie.write_text(source, encoding="utf-8")
    L.integrer_cle(L.module_hex(cle.public_key()), copie)
    ligne = next(l for l in copie.read_text(encoding="utf-8").splitlines() if "CLE_PUBLIQUE_MODULE As String" in l)
    assert re.search(r'= "[0-9a-f]{512}" \' <<SG_CLE_PUBLIQUE>>', ligne)
    assert len(ligne) < 1023  # limite de longueur de ligne VBA
    # Le marqueur reste en place : la cle peut etre remplacee a nouveau.
    L.integrer_cle(L.module_hex(L.generer_cle(2048).public_key()), copie)


def test_aucune_cle_privee_dans_le_code_client():
    for f in (RACINE / "src").rglob("*"):
        if f.is_file():
            contenu = f.read_text(encoding="utf-8", errors="ignore")
            assert "PRIVATE KEY" not in contenu, f
