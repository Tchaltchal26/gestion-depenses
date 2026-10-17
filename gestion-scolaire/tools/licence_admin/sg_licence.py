#!/usr/bin/env python3
"""Outil d'administration des licences SG (cote EDITEUR uniquement).

Cet outil detient la cle privee de signature. Il ne doit JAMAIS etre livre aux
clients ni copie dans le classeur Excel. Le classeur ne contient que la cle
publique (module RSA) servant a verifier les licences.

Fonctions :
  - generation de la paire de cles RSA 3072 bits (cle privee chiffree) ;
  - integration de la cle publique dans src/vba/modLicence.bas ;
  - emission de licences signees (annuelle, mensuelle, perpetuelle, essai,
    demonstration, multi-etablissements, reseau) ;
  - lecture des demandes d'activation hors ligne (.sgreq) ;
  - registre commercial SQLite avec journal chaine (SHA-256) ;
  - revocation, transfert d'installation, liste de revocation signee.

Format de licence (texte UTF-8, une cle par ligne, ordre fixe) :
    FORMAT=SGLIC-1
    LICENCE_ID=...
    ...
    SIGNATURE=<base64 RSASSA-PKCS1-v1_5 SHA-256 des lignes precedentes, jointes par LF>
"""
from __future__ import annotations

import argparse
import base64
import dataclasses
import datetime as dt
import getpass
import hashlib
import json
import os
import re
import sqlite3
import sys
import uuid
from pathlib import Path

from cryptography.exceptions import InvalidSignature
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import padding, rsa

FORMAT_LICENCE = "SGLIC-1"
FORMAT_REVOCATION = "SGREV-1"
FORMAT_DEMANDE = "SGREQ-1"
TYPES = {"ANNUELLE", "MENSUELLE", "PERPETUELLE", "ESSAI", "DEMONSTRATION", "MULTI_ETABLISSEMENTS", "RESEAU"}
MODULES_CONNUS = {
    "ETAB", "PARAM", "UTIL", "ANNEE", "STRUCT", "ELEVES", "INSCR", "AUDIT", "SAUVEG",
    "PRESENCES", "NOTES", "EXAMENS", "FINANCES", "PERSONNEL", "EDT", "STOCKS", "DOCUMENTS",
    "BIBLIO", "TRANSPORT", "CANTINE", "DISCIPLINE", "STAGES",
}
MARQUEUR_CLE = "' <<SG_CLE_PUBLIQUE>>"


class ErreurLicence(Exception):
    pass


@dataclasses.dataclass
class Licence:
    licence_id: str
    serie: str
    titulaire: str
    etablissements: str
    edition: str
    type_licence: str
    emise_le: str
    expire_le: str
    maintenance_jusqu_au: str
    version_majeure_max: str
    installations_max: str
    nb_etablissements_max: str
    modules: str
    empreinte_installation: str
    grace_jours: str = "15"

    ORDRE = (
        ("LICENCE_ID", "licence_id"),
        ("SERIE", "serie"),
        ("TITULAIRE", "titulaire"),
        ("ETABLISSEMENTS", "etablissements"),
        ("EDITION", "edition"),
        ("TYPE", "type_licence"),
        ("EMISE_LE", "emise_le"),
        ("EXPIRE_LE", "expire_le"),
        ("MAINTENANCE_JUSQU_AU", "maintenance_jusqu_au"),
        ("VERSION_MAJEURE_MAX", "version_majeure_max"),
        ("INSTALLATIONS_MAX", "installations_max"),
        ("NB_ETABLISSEMENTS_MAX", "nb_etablissements_max"),
        ("MODULES", "modules"),
        ("EMPREINTE_INSTALLATION", "empreinte_installation"),
        ("GRACE_JOURS", "grace_jours"),
    )

    @classmethod
    def champs_exemple(cls) -> dict:
        return dict(
            licence_id="LIC-EXEMPLE", serie="SG-EXEMPLE", titulaire="Etablissement fictif",
            etablissements="ETB-A", edition="STANDARD", type_licence="ANNUELLE",
            emise_le="2026-09-01", expire_le="2027-08-31", maintenance_jusqu_au="2027-08-31",
            version_majeure_max="0", installations_max="1", nb_etablissements_max="1",
            modules="ELEVES", empreinte_installation="*", grace_jours="15",
        )

    def valider(self) -> None:
        for _, attr in self.ORDRE:
            v = getattr(self, attr)
            if not isinstance(v, str) or "\n" in v or "\r" in v:
                raise ErreurLicence(f"Valeur invalide pour {attr}")
        if self.type_licence not in TYPES:
            raise ErreurLicence(f"Type de licence inconnu : {self.type_licence}")
        for d in (self.emise_le, self.expire_le, self.maintenance_jusqu_au):
            if d:
                dt.date.fromisoformat(d)
        if self.type_licence != "PERPETUELLE" and not self.expire_le:
            raise ErreurLicence("Une licence non perpetuelle doit avoir une date d'expiration")
        inconnus = {m for m in self.modules.split(";") if m} - MODULES_CONNUS
        if inconnus:
            raise ErreurLicence(f"Modules inconnus : {sorted(inconnus)}")
        if self.empreinte_installation != "*" and not re.fullmatch(r"[0-9a-f]{64}", self.empreinte_installation):
            raise ErreurLicence("Empreinte d'installation invalide (64 hex ou *)")
        nb_etabs = len([e for e in self.etablissements.split(";") if e])
        if nb_etabs > int(self.nb_etablissements_max):
            raise ErreurLicence("Plus d'etablissements que NB_ETABLISSEMENTS_MAX")
        int(self.version_majeure_max), int(self.installations_max), int(self.grace_jours)

    def lignes(self) -> list[str]:
        return [f"FORMAT={FORMAT_LICENCE}"] + [f"{k}={getattr(self, a)}" for k, a in self.ORDRE]


# ---------------------------------------------------------------------------
# Cryptographie
# ---------------------------------------------------------------------------
def generer_cle(bits: int = 3072) -> rsa.RSAPrivateKey:
    return rsa.generate_private_key(public_exponent=65537, key_size=bits)


def module_hex(pub: rsa.RSAPublicKey) -> str:
    n = pub.public_numbers().n
    return format(n, "x").rjust((pub.key_size + 7) // 8 * 2, "0")


def signer(cle: rsa.RSAPrivateKey, contenu: str) -> str:
    sig = cle.sign(contenu.encode("utf-8"), padding.PKCS1v15(), hashes.SHA256())
    return base64.b64encode(sig).decode("ascii")


def separer(texte: str) -> tuple[str, bytes, dict[str, str]]:
    lignes = [l.strip() for l in texte.replace("\r\n", "\n").replace("\r", "\n").lstrip("﻿").split("\n")]
    lignes = [l for l in lignes if l]
    if not lignes or not lignes[-1].startswith("SIGNATURE="):
        raise ErreurLicence("Signature absente")
    contenu = "\n".join(lignes[:-1])
    sig = base64.b64decode(lignes[-1][len("SIGNATURE="):], validate=True)
    champs = {}
    for l in lignes[:-1]:
        if "=" not in l:
            raise ErreurLicence(f"Ligne invalide : {l}")
        k, v = l.split("=", 1)
        if k in champs:
            raise ErreurLicence(f"Champ en double : {k}")
        champs[k] = v
    return contenu, sig, champs


def verifier_texte(pub: rsa.RSAPublicKey, texte: str) -> dict[str, str]:
    contenu, sig, champs = separer(texte)
    try:
        pub.verify(sig, contenu.encode("utf-8"), padding.PKCS1v15(), hashes.SHA256())
    except InvalidSignature as e:
        raise ErreurLicence("Signature invalide") from e
    return champs


def emettre_texte(cle: rsa.RSAPrivateKey, lic: Licence) -> str:
    lic.valider()
    contenu = "\n".join(lic.lignes())
    return contenu + "\nSIGNATURE=" + signer(cle, contenu) + "\n"


def emettre_revocations_texte(cle: rsa.RSAPrivateKey, ids: list[str], date: str) -> str:
    contenu = "\n".join([f"FORMAT={FORMAT_REVOCATION}", f"EMIS_LE={date}", "LICENCES=" + ";".join(sorted(set(ids)))])
    return contenu + "\nSIGNATURE=" + signer(cle, contenu) + "\n"


def lire_demande(texte: str) -> dict[str, str]:
    champs = {}
    for l in texte.replace("\r", "").split("\n"):
        if "=" in l:
            k, v = l.strip().split("=", 1)
            champs[k] = v
    if champs.get("FORMAT") != FORMAT_DEMANDE:
        raise ErreurLicence("Ce fichier n'est pas une demande d'activation SG")
    controle = champs.pop("CONTROLE", "")
    attendu = hashlib.sha256("\n".join(f"{k}={v}" for k, v in champs.items()).encode("utf-8")).hexdigest()
    if controle != attendu:
        raise ErreurLicence("Demande alteree ou incomplete (controle incorrect)")
    if not re.fullmatch(r"[0-9a-f]{64}", champs.get("EMPREINTE", "")):
        raise ErreurLicence("Empreinte d'installation absente ou invalide")
    return champs


def charger_cle_privee(chemin: Path) -> rsa.RSAPrivateKey:
    data = chemin.read_bytes()
    phrase = os.environ.get("SG_LICENCE_PASSPHRASE")
    if b"ENCRYPTED" in data and phrase is None:
        phrase = getpass.getpass("Phrase secrete de la cle privee : ")
    return serialization.load_pem_private_key(data, password=phrase.encode() if phrase else None)


def charger_cle_publique(chemin: Path) -> rsa.RSAPublicKey:
    return serialization.load_pem_public_key(chemin.read_bytes())


# ---------------------------------------------------------------------------
# Registre commercial (SQLite) avec journal chaine
# ---------------------------------------------------------------------------
SCHEMA = """
CREATE TABLE IF NOT EXISTS licences (
    licence_id TEXT PRIMARY KEY, serie TEXT NOT NULL, titulaire TEXT, etablissements TEXT,
    edition TEXT, type_licence TEXT, emise_le TEXT, expire_le TEXT, maintenance_jusqu_au TEXT,
    version_majeure_max TEXT, installations_max TEXT, nb_etablissements_max TEXT, modules TEXT,
    empreinte_installation TEXT, grace_jours TEXT, statut TEXT NOT NULL, remplace TEXT, texte TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS activations (
    id INTEGER PRIMARY KEY AUTOINCREMENT, licence_id TEXT NOT NULL, empreinte TEXT NOT NULL,
    install_id TEXT, date_activation TEXT NOT NULL, statut TEXT NOT NULL, motif TEXT
);
CREATE TABLE IF NOT EXISTS journal (
    id INTEGER PRIMARY KEY AUTOINCREMENT, horodatage TEXT NOT NULL, operateur TEXT NOT NULL,
    operation TEXT NOT NULL, details TEXT NOT NULL, hash_prec TEXT NOT NULL, hash TEXT NOT NULL
);
"""


class Registre:
    def __init__(self, chemin: Path, operateur: str):
        self.db = sqlite3.connect(chemin)
        self.db.executescript(SCHEMA)
        self.operateur = operateur

    def journaliser(self, operation: str, details: dict) -> None:
        prec = self.db.execute("SELECT hash FROM journal ORDER BY id DESC LIMIT 1").fetchone()
        hprec = prec[0] if prec else "0" * 64
        horo = dt.datetime.now().isoformat(timespec="seconds")
        det = json.dumps(details, ensure_ascii=False, sort_keys=True)
        h = hashlib.sha256(f"{hprec}|{horo}|{self.operateur}|{operation}|{det}".encode()).hexdigest()
        self.db.execute(
            "INSERT INTO journal(horodatage, operateur, operation, details, hash_prec, hash) VALUES (?,?,?,?,?,?)",
            (horo, self.operateur, operation, det, hprec, h),
        )

    def verifier_journal(self) -> tuple[bool, int]:
        hprec = "0" * 64
        n = 0
        for horo, op_, operation, det, hp, h in self.db.execute(
            "SELECT horodatage, operateur, operation, details, hash_prec, hash FROM journal ORDER BY id"
        ):
            calc = hashlib.sha256(f"{hprec}|{horo}|{op_}|{operation}|{det}".encode()).hexdigest()
            if hp != hprec or h != calc:
                return False, n
            hprec = h
            n += 1
        return True, n

    def enregistrer(self, lic: Licence, texte: str, remplace: str | None = None) -> None:
        vals = dataclasses.asdict(lic)
        self.db.execute(
            f"INSERT INTO licences({','.join(vals)}, statut, remplace, texte) VALUES ({','.join('?' * len(vals))},?,?,?)",
            (*vals.values(), "ACTIVE", remplace, texte),
        )
        if lic.empreinte_installation != "*":
            self.db.execute(
                "INSERT INTO activations(licence_id, empreinte, date_activation, statut) VALUES (?,?,?,?)",
                (lic.licence_id, lic.empreinte_installation, dt.date.today().isoformat(), "ACTIVE"),
            )
        self.journaliser("EMISSION", {"licence_id": lic.licence_id, "type": lic.type_licence, "remplace": remplace})
        self.db.commit()

    def activations_actives(self, licence_id: str) -> int:
        return self.db.execute(
            "SELECT COUNT(*) FROM activations WHERE licence_id=? AND statut='ACTIVE'", (licence_id,)
        ).fetchone()[0]

    def licence(self, licence_id: str) -> sqlite3.Row:
        self.db.row_factory = sqlite3.Row
        r = self.db.execute("SELECT * FROM licences WHERE licence_id=?", (licence_id,)).fetchone()
        self.db.row_factory = None
        if r is None:
            raise ErreurLicence(f"Licence inconnue : {licence_id}")
        return r

    def changer_statut(self, licence_id: str, statut: str, motif: str) -> None:
        if not motif.strip():
            raise ErreurLicence("Motif obligatoire")
        self.licence(licence_id)
        self.db.execute("UPDATE licences SET statut=? WHERE licence_id=?", (statut, licence_id))
        if statut == "REVOQUEE":
            self.db.execute(
                "UPDATE activations SET statut='REVOQUEE', motif=? WHERE licence_id=? AND statut='ACTIVE'",
                (motif, licence_id),
            )
        self.journaliser("STATUT", {"licence_id": licence_id, "statut": statut, "motif": motif})
        self.db.commit()

    def revoquees(self) -> list[str]:
        return [r[0] for r in self.db.execute("SELECT licence_id FROM licences WHERE statut='REVOQUEE'")]


def nouvelle_licence_id() -> str:
    return "LIC-" + uuid.uuid4().hex[:12].upper()


def nouvelle_serie() -> str:
    h = uuid.uuid4().hex.upper()
    return f"SG-{h[0:4]}-{h[4:8]}-{h[8:12]}-{h[12:16]}"


def date_expiration(type_licence: str, debut: dt.date, duree_jours: int | None) -> str:
    if type_licence == "PERPETUELLE":
        return ""
    if duree_jours is None:
        duree_jours = {"MENSUELLE": 30, "ESSAI": 30, "DEMONSTRATION": 15}.get(type_licence, 365)
    return (debut + dt.timedelta(days=duree_jours - 1)).isoformat()


def integrer_cle(modulus: str, chemin_vba: Path) -> None:
    texte = chemin_vba.read_text(encoding="utf-8")
    motif = re.compile(r'(Private Const CLE_PUBLIQUE_MODULE As String = ")[0-9a-fA-F]*(" ' + re.escape(MARQUEUR_CLE) + ")")
    if not motif.search(texte):
        raise ErreurLicence("Marqueur de cle publique introuvable dans modLicence.bas")
    # VBA limite une ligne a 1023 caracteres : 3072 bits = 768 hex, cela tient.
    texte = motif.sub(lambda m: m.group(1) + modulus + m.group(2), texte)
    chemin_vba.write_text(texte, encoding="utf-8")


# ---------------------------------------------------------------------------
# Ligne de commande
# ---------------------------------------------------------------------------
def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--registre", default="registre_licences.db")
    ap.add_argument("--operateur", default=os.environ.get("USER", "editeur"))
    sp = ap.add_subparsers(dest="cmd", required=True)

    g = sp.add_parser("cles", help="Generer la paire de cles (une seule fois)")
    g.add_argument("--dossier", required=True)
    g.add_argument("--bits", type=int, default=3072)
    g.add_argument("--sans-phrase", action="store_true", help="Cle privee non chiffree (tests uniquement)")

    i = sp.add_parser("integrer-cle", help="Ecrire la cle publique dans modLicence.bas")
    i.add_argument("--publique", required=True)
    i.add_argument("--vba", required=True)

    e = sp.add_parser("emettre", help="Emettre une licence signee")
    e.add_argument("--privee", required=True)
    e.add_argument("--titulaire", required=True)
    e.add_argument("--etablissements", required=True, help="Codes separes par ;")
    e.add_argument("--type", dest="type_licence", required=True, choices=sorted(TYPES))
    e.add_argument("--edition", default="STANDARD")
    e.add_argument("--duree-jours", type=int)
    e.add_argument("--debut", default=dt.date.today().isoformat())
    e.add_argument("--maintenance-jusqu-au", default="")
    e.add_argument("--version-majeure-max", default="0")
    e.add_argument("--installations-max", default="1")
    e.add_argument("--nb-etablissements-max")
    e.add_argument("--modules", required=True)
    e.add_argument("--demande", help="Fichier .sgreq fourni par le client")
    e.add_argument("--empreinte", help="Empreinte d'installation (64 hex) ou * (non liee)")
    e.add_argument("--grace-jours", default="15")
    e.add_argument("--sortie", required=True)

    v = sp.add_parser("verifier", help="Verifier une licence")
    v.add_argument("--publique", required=True)
    v.add_argument("fichier")

    r = sp.add_parser("revoquer", help="Revoquer ou suspendre une licence")
    r.add_argument("licence_id")
    r.add_argument("--motif", required=True)
    r.add_argument("--statut", default="REVOQUEE", choices=["REVOQUEE", "SUSPENDUE", "ACTIVE"])

    t = sp.add_parser("transferer", help="Transferer une licence vers une nouvelle installation")
    t.add_argument("licence_id")
    t.add_argument("--privee", required=True)
    t.add_argument("--demande", required=True)
    t.add_argument("--motif", required=True)
    t.add_argument("--sortie", required=True)

    lr = sp.add_parser("revocations", help="Produire la liste de revocation signee")
    lr.add_argument("--privee", required=True)
    lr.add_argument("--sortie", required=True)

    sp.add_parser("liste", help="Lister les licences")
    sp.add_parser("journal", help="Verifier l'integrite du journal du registre")

    a = ap.parse_args(argv)

    if a.cmd == "cles":
        d = Path(a.dossier)
        d.mkdir(parents=True, exist_ok=True)
        if (d / "cle_privee.pem").exists():
            raise SystemExit("Une cle privee existe deja dans ce dossier : refus d'ecraser.")
        cle = generer_cle(a.bits)
        if a.sans_phrase:
            chiff = serialization.NoEncryption()
        else:
            phrase = os.environ.get("SG_LICENCE_PASSPHRASE") or getpass.getpass("Nouvelle phrase secrete : ")
            if len(phrase) < 12:
                raise SystemExit("Phrase secrete trop courte (12 caracteres minimum).")
            chiff = serialization.BestAvailableEncryption(phrase.encode())
        (d / "cle_privee.pem").write_bytes(
            cle.private_bytes(serialization.Encoding.PEM, serialization.PrivateFormat.PKCS8, chiff)
        )
        os.chmod(d / "cle_privee.pem", 0o600)
        (d / "cle_publique.pem").write_bytes(
            cle.public_key().public_bytes(serialization.Encoding.PEM, serialization.PublicFormat.SubjectPublicKeyInfo)
        )
        print(f"Cles creees dans {d}. Sauvegardez cle_privee.pem hors ligne, en deux exemplaires.")
        return 0

    if a.cmd == "integrer-cle":
        integrer_cle(module_hex(charger_cle_publique(Path(a.publique))), Path(a.vba))
        print("Cle publique integree.")
        return 0

    if a.cmd == "verifier":
        champs = verifier_texte(charger_cle_publique(Path(a.publique)), Path(a.fichier).read_text(encoding="utf-8"))
        print(json.dumps(champs, ensure_ascii=False, indent=2))
        return 0

    reg = Registre(Path(a.registre), a.operateur)

    if a.cmd == "emettre":
        empreinte = a.empreinte
        if a.demande:
            dem = lire_demande(Path(a.demande).read_text(encoding="utf-8"))
            empreinte = dem["EMPREINTE"]
        if not empreinte:
            raise SystemExit("Fournir --demande ou --empreinte (utiliser * pour une licence non liee).")
        etabs = [x for x in a.etablissements.split(";") if x]
        debut = dt.date.fromisoformat(a.debut)
        lic = Licence(
            licence_id=nouvelle_licence_id(), serie=nouvelle_serie(), titulaire=a.titulaire,
            etablissements=";".join(etabs), edition=a.edition, type_licence=a.type_licence,
            emise_le=debut.isoformat(), expire_le=date_expiration(a.type_licence, debut, a.duree_jours),
            maintenance_jusqu_au=a.maintenance_jusqu_au, version_majeure_max=a.version_majeure_max,
            installations_max=a.installations_max, nb_etablissements_max=a.nb_etablissements_max or str(len(etabs)),
            modules=a.modules, empreinte_installation=empreinte, grace_jours=a.grace_jours,
        )
        texte = emettre_texte(charger_cle_privee(Path(a.privee)), lic)
        reg.enregistrer(lic, texte)
        Path(a.sortie).write_text(texte, encoding="utf-8")
        print(f"Licence {lic.licence_id} ({lic.serie}) ecrite dans {a.sortie}")
        return 0

    if a.cmd == "revoquer":
        reg.changer_statut(a.licence_id, a.statut, a.motif)
        print(f"{a.licence_id} -> {a.statut}")
        return 0

    if a.cmd == "transferer":
        anc = reg.licence(a.licence_id)
        if anc["statut"] != "ACTIVE":
            raise SystemExit("Seule une licence active peut etre transferee.")
        dem = lire_demande(Path(a.demande).read_text(encoding="utf-8"))
        champs = {k: anc[k] for k in Licence.champs_exemple()}
        champs.update(licence_id=nouvelle_licence_id(), empreinte_installation=dem["EMPREINTE"])
        lic = Licence(**champs)
        texte = emettre_texte(charger_cle_privee(Path(a.privee)), lic)
        reg.changer_statut(a.licence_id, "REVOQUEE", f"Transfert vers {lic.licence_id} : {a.motif}")
        reg.enregistrer(lic, texte, remplace=a.licence_id)
        Path(a.sortie).write_text(texte, encoding="utf-8")
        print(f"Nouvelle licence {lic.licence_id} ; ancienne {a.licence_id} revoquee.")
        return 0

    if a.cmd == "revocations":
        texte = emettre_revocations_texte(charger_cle_privee(Path(a.privee)), reg.revoquees(), dt.date.today().isoformat())
        reg.journaliser("LISTE_REVOCATION", {"nb": len(reg.revoquees())})
        reg.db.commit()
        Path(a.sortie).write_text(texte, encoding="utf-8")
        print(f"Liste de revocation ecrite dans {a.sortie}")
        return 0

    if a.cmd == "liste":
        for row in reg.db.execute(
            "SELECT licence_id, serie, titulaire, type_licence, expire_le, statut FROM licences ORDER BY emise_le"
        ):
            print(" | ".join(x or "" for x in row))
        return 0

    if a.cmd == "journal":
        ok, n = reg.verifier_journal()
        print(("Journal integre" if ok else "JOURNAL ALTERE") + f" ({n} evenements verifies)")
        return 0 if ok else 2
    return 1


if __name__ == "__main__":
    try:
        sys.exit(main())
    except ErreurLicence as exc:
        print(f"Erreur : {exc}", file=sys.stderr)
        sys.exit(2)
