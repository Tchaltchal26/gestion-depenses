Attribute VB_Name = "modSchema"
'===============================================================================
' SG - Module modSchema
' Role : definition unique du modele de donnees (tables et colonnes).
'        Chaque table est une feuille technique masquee dont la ligne 1 contient
'        les noms de colonnes. La premiere colonne est TOUJOURS l'identifiant
'        technique unique et permanent. Les donnees sont retrouvees par
'        identifiant et par nom de colonne, jamais par numero de ligne.
' Evolution : toute modification du schema incremente SCHEMA_VERSION et doit
'        etre accompagnee d'une procedure de migration (modSetup).
'===============================================================================
Option Explicit

Public Const SCHEMA_VERSION As Long = 1

Public Function SchemaTables() As Variant
    SchemaTables = Array("T_Systeme", "T_Etablissements", "T_Parametres", "T_Compteurs", _
        "T_Roles", "T_Permissions", "T_Utilisateurs", "T_Habilitations", _
        "T_Annees", "T_Cycles", "T_Niveaux", "T_Filieres", "T_Classes", _
        "T_Eleves", "T_Responsables", "T_LiensResponsables", _
        "T_Inscriptions", "T_Affectations", "T_Transferts", _
        "T_Sauvegardes", "T_Revocations", "T_Audit")
End Function

Public Function SchemaColonnes(ByVal table As String) As String
    Select Case table
        Case "T_Systeme"
            SchemaColonnes = "Cle;Valeur"
        Case "T_Etablissements"
            SchemaColonnes = "EtabID;Code;Nom;TypeEtab;Statut;Pays;Ville;Adresse;Telephone;Email;Devise;Langue;Actif;DateCreation;CreePar"
        Case "T_Parametres"
            SchemaColonnes = "ParamID;EtabID;Cle;Valeur;ModifieLe;ModifiePar"
        Case "T_Compteurs"
            SchemaColonnes = "CompteurID;EtabID;Cle;Valeur"
        Case "T_Roles"
            SchemaColonnes = "RoleID;Code;Libelle;Systeme"
        Case "T_Permissions"
            SchemaColonnes = "PermID;RoleID;Module;Niveau"
        Case "T_Utilisateurs"
            SchemaColonnes = "UserID;Login;NomComplet;Sel;HashMdp;Iterations;Actif;AdminTech;AdminCentral;Echecs;VerrouJusqua;DoitChangerMdp;DateCreation;DernierAcces"
        Case "T_Habilitations"
            SchemaColonnes = "HabID;UserID;EtabID;RoleID;DateDebut;DateFin;Actif"
        Case "T_Annees"
            SchemaColonnes = "AnneeID;EtabID;Code;Libelle;DateDebut;DateFin;Statut;OuverteLe;ClotureeLe;ClotureePar"
        Case "T_Cycles"
            SchemaColonnes = "CycleID;EtabID;Code;Libelle;ModeEvaluation;Ordre;Actif"
        Case "T_Niveaux"
            SchemaColonnes = "NiveauID;EtabID;CycleID;Code;Libelle;Ordre;Actif"
        Case "T_Filieres"
            SchemaColonnes = "FiliereID;EtabID;Code;Libelle;TypeFiliere;Actif"
        Case "T_Classes"
            SchemaColonnes = "ClasseID;EtabID;AnneeID;NiveauID;FiliereID;Code;Libelle;Capacite;Salle;Actif"
        Case "T_Eleves"
            SchemaColonnes = "EleveID;EtabID;Matricule;Nom;Prenoms;Sexe;DateNaissance;LieuNaissance;Nationalite;Adresse;Telephone;EtabPrecedent;EleveOrigineID;Statut;DateStatut;CleDoublon;DateCreation;CreePar;ModifieLe;Version"
        Case "T_Responsables"
            SchemaColonnes = "RespID;EtabID;Nom;Prenoms;Telephone;Email;Adresse;Profession;Actif"
        Case "T_LiensResponsables"
            SchemaColonnes = "LienID;EtabID;EleveID;RespID;Lien;Legal;Financier;Urgence;AutoriseRecuperer;Actif"
        Case "T_Inscriptions"
            SchemaColonnes = "InscriptionID;EtabID;AnneeID;EleveID;ClasseID;TypeInscription;Decision;DateInscription;DateEffet;Statut;DateFin;MotifFin;CreePar;DateOperation"
        Case "T_Affectations"
            SchemaColonnes = "AffectationID;EtabID;InscriptionID;ClasseID;DateDebut;DateFin;Motif;DateOperation;Utilisateur"
        Case "T_Transferts"
            SchemaColonnes = "TransfertID;EtabSource;EleveSource;EtabCible;EleveCible;DateTransfert;Motif;Utilisateur"
        Case "T_Sauvegardes"
            SchemaColonnes = "SauvID;EtabID;Horodatage;Chemin;TypeSauv;NbEvenementsAudit;HashAudit;Statut;Utilisateur"
        Case "T_Revocations"
            SchemaColonnes = "LicenceID;ImporteeLe;ListeEmiseLe"
        Case "T_Audit"
            SchemaColonnes = "EventID;Horodatage;UserID;EtabID;Module;Operation;TableCible;EnregistrementID;Champ;AncienneValeur;NouvelleValeur;Motif;Resultat;TxRef;HashPrec;Hash"
        Case Else
            SG_Raise SG_ERR_INTROUVABLE, "Table inconnue : " & table
    End Select
End Function

' Prefixe des identifiants generes pour chaque table.
Public Function SchemaPrefixe(ByVal table As String) As String
    Select Case table
        Case "T_Etablissements": SchemaPrefixe = "ETB"
        Case "T_Parametres": SchemaPrefixe = "PRM"
        Case "T_Compteurs": SchemaPrefixe = "CPT"
        Case "T_Roles": SchemaPrefixe = "ROL"
        Case "T_Permissions": SchemaPrefixe = "PER"
        Case "T_Utilisateurs": SchemaPrefixe = "USR"
        Case "T_Habilitations": SchemaPrefixe = "HAB"
        Case "T_Annees": SchemaPrefixe = "ANS"
        Case "T_Cycles": SchemaPrefixe = "CYC"
        Case "T_Niveaux": SchemaPrefixe = "NIV"
        Case "T_Filieres": SchemaPrefixe = "FIL"
        Case "T_Classes": SchemaPrefixe = "CLS"
        Case "T_Eleves": SchemaPrefixe = "ELV"
        Case "T_Responsables": SchemaPrefixe = "RSP"
        Case "T_LiensResponsables": SchemaPrefixe = "LNK"
        Case "T_Inscriptions": SchemaPrefixe = "INS"
        Case "T_Affectations": SchemaPrefixe = "AFF"
        Case "T_Transferts": SchemaPrefixe = "TRF"
        Case "T_Sauvegardes": SchemaPrefixe = "SAV"
        Case "T_Audit": SchemaPrefixe = "EVT"
        Case Else: SchemaPrefixe = "ID"
    End Select
End Function

' Tables dont les lignes appartiennent a un etablissement (colonne EtabID).
' Toute lecture metier de ces tables DOIT etre filtree par etablissement.
Public Function SchemaEstCloisonnee(ByVal table As String) As Boolean
    SchemaEstCloisonnee = InList(table, "T_Parametres;T_Compteurs;T_Habilitations;T_Annees;T_Cycles;T_Niveaux;T_Filieres;T_Classes;T_Eleves;T_Responsables;T_LiensResponsables;T_Inscriptions;T_Affectations;T_Sauvegardes")
End Function
