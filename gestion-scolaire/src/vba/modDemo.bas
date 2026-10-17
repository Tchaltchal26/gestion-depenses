Attribute VB_Name = "modDemo"
'===============================================================================
' SG - Module modDemo (jeu de donnees de demonstration)
' Donnees ENTIEREMENT FICTIVES : noms generes par combinaison aleatoire de
' listes generiques ; aucune personne reelle. Toutes les donnees passent par
' les regles metier et sont donc journalisees comme une saisie normale.
' Refuse de s'executer si un etablissement existe deja (protection des
' donnees reelles).
'===============================================================================
Option Explicit

Private Const MDP_DEMO As String = "Demo2026x"

Public Function SG_ChargerDemo() As String
    Dim e As String, a As String, cyc As Variant, niv As Variant, i As Long, j As Long, k As Long
    Dim cycles As Variant, niveaux As Variant, classes As New Collection, idCycle As String, idNiv As String
    Dim noms As Variant, prenomsF As Variant, prenomsM As Variant, r As Variant, el As String, ordre As Long
    Dim annNaiss As Long, sexe As String, resp As String, cls As String, nb As Long, filA As String, filD As String
    Dim ins As String

    If Not Sec_EstConnecte() Then SG_Raise SG_ERR_ACCES, "Se connecter avec le compte administrateur pour charger la demonstration."
    If DAL_Count("T_Etablissements", RecNew()) > 0 Then SG_Raise SG_ERR_ETAT, "Des etablissements existent deja : la demonstration ne se charge que dans un classeur vierge."
    Rnd -1
    Randomize 2026

    e = Etab_Creer(RecFrom("Code=DEMO|Nom=Complexe scolaire fictif Les Flamboyants|TypeEtab=COMPLEXE|Statut=PRIVE|Pays=Pays fictif|Ville=Ville fictive|Telephone=0000000000|Devise=XOF"))
    Sec_Habiliter Sec_UtilisateurCourant(), e, "DIRECTEUR", "2026-01-01"
    Sec_Habiliter Sec_CreerUtilisateur("dir.demo", "Direction (demo)", MDP_DEMO), e, "DIRECTEUR", "2026-01-01"
    Sec_Habiliter Sec_CreerUtilisateur("secr.demo", "Secretariat (demo)", MDP_DEMO), e, "SECRETAIRE", "2026-01-01"
    Sec_Habiliter Sec_CreerUtilisateur("prof.demo", "Enseignant (demo)", MDP_DEMO), e, "ENSEIGNANT", "2026-01-01"
    Sec_Habiliter Sec_CreerUtilisateur("audit.demo", "Auditeur (demo)", MDP_DEMO), e, "AUDITEUR", "2026-01-01"

    a = Annee_Creer(e, "2026-2027", "Annee scolaire 2026-2027", "2026-09-01", "2027-07-31")
    Annee_Ouvrir e, a
    filA = Filiere_Creer(e, "A", "Serie litteraire (A)", "SERIE")
    filD = Filiere_Creer(e, "D", "Serie scientifique (D)", "SERIE")

    ' Code cycle | libelle | mode d'evaluation | niveaux (code:libelle:age a la rentree)
    cycles = Array( _
        "MAT|Maternelle|COMPETENCES|PS:Petite section:3,MS:Moyenne section:4,GS:Grande section:5", _
        "PRI|Primaire|COMPETENCES|CP:Cours preparatoire:6,CE1:Cours elementaire 1:7,CM2:Cours moyen 2:10", _
        "COL|College|NOTES|6E:Sixieme:11,3E:Troisieme:14", _
        "LYC|Lycee|NOTES|1ERE:Premiere:16")
    ordre = 0
    For i = 0 To UBound(cycles)
        cyc = Split(cycles(i), "|")
        idCycle = Cycle_Creer(e, CStr(cyc(0)), CStr(cyc(1)), CStr(cyc(2)), i + 1)
        niveaux = Split(cyc(3), ",")
        For j = 0 To UBound(niveaux)
            niv = Split(niveaux(j), ":")
            ordre = ordre + 1
            idNiv = Niveau_Creer(e, idCycle, CStr(niv(0)), CStr(niv(1)), ordre)
            If cyc(0) = "LYC" Then
                classes.Add Classe_Creer(e, a, idNiv, filA, niv(0) & "A", niv(1) & " A", 6) & "|" & niv(2)
                classes.Add Classe_Creer(e, a, idNiv, filD, niv(0) & "D", niv(1) & " D", 6) & "|" & niv(2)
            Else
                classes.Add Classe_Creer(e, a, idNiv, "", niv(0) & "A", niv(1) & " A", 6) & "|" & niv(2)
            End If
        Next j
    Next i

    noms = Split("ABENA;BOKO;DANSO;EKUE;FOFANA;GNAHO;HOUNKPE;ISSA;KOFFI;LAWSON;MENSAH;NDIAYE;OUEDRAOGO;SANOU;TCHAKOUNTE;YAO;ZINSOU;AGBO", ";")
    prenomsF = Split("Ama;Aicha;Brigitte;Clarisse;Esther;Fanta;Grace;Habiba;Ines;Josiane;Kadi;Laetitia;Mariam;Nadege;Rokia", ";")
    prenomsM = Split("Abdou;Bruno;Cedric;Didier;Emmanuel;Fabrice;Gildas;Hamed;Ismael;Jules;Kader;Landry;Moussa;Noel;Omar", ";")

    For k = 1 To classes.Count
        cls = Split(classes(k), "|")(0)
        annNaiss = 2026 - CLng(Split(classes(k), "|")(1))
        For i = 1 To 4
            r = RecNew()
            If Rnd < 0.5 Then
                sexe = "F": RecSet r, "Prenoms", prenomsF(Int(Rnd * (UBound(prenomsF) + 1)))
            Else
                sexe = "M": RecSet r, "Prenoms", prenomsM(Int(Rnd * (UBound(prenomsM) + 1)))
            End If
            RecSet r, "Nom", noms(Int(Rnd * (UBound(noms) + 1)))
            RecSet r, "Sexe", sexe
            RecSet r, "DateNaissance", annNaiss & "-" & PadLeft(CStr(1 + Int(Rnd * 12)), 2) & "-" & PadLeft(CStr(1 + Int(Rnd * 28)), 2)
            RecSet r, "LieuNaissance", "Ville fictive"
            RecSet r, "Nationalite", "Fictive"
            On Error Resume Next
            el = Eleve_Creer(e, r)
            If Err.Number = SG_ERR_DOUBLON Then
                Err.Clear
                el = Eleve_Creer(e, r, True, "Homonyme genere par la demonstration")
            End If
            On Error GoTo 0
            resp = Resp_Creer(e, RecFrom("Nom=" & RecStr(r, "Nom") & "|Prenoms=Parent|Telephone=00000" & PadLeft(CStr(k * 10 + i), 5) & "|Profession=Fictive"))
            Resp_Lier e, el, resp, IIf(Rnd < 0.5, "Pere", "Mere"), "LEGAL;FINANCIER;URGENCE;RECUPERATION"
            Inscr_Inscrire e, el, a, cls, "PREMIERE", "2026-09-01"
            nb = nb + 1
        Next i
    Next k

    ' Quelques evenements de vie scolaire pour illustrer l'historique.
    ins = DAL_First("T_Inscriptions", RecFrom("EtabID=" & e & "|ClasseID=" & Split(classes(5), "|")(0) & "|Statut=ACTIVE"))
    Inscr_Terminer e, ins, "DEPART", "2026-12-18", "Demenagement de la famille (fictif)"
    ins = DAL_First("T_Inscriptions", RecFrom("EtabID=" & e & "|ClasseID=" & Split(classes(classes.Count - 1), "|")(0) & "|Statut=ACTIVE"))
    Inscr_ChangerClasse e, ins, Split(classes(classes.Count), "|")(0), "2026-11-02", "Changement de serie (fictif)"

    SG_ChargerDemo = nb & " eleves fictifs crees dans " & classes.Count & " classes. Comptes : dir.demo, secr.demo, prof.demo, audit.demo (mot de passe " & MDP_DEMO & ")."
End Function

' Utilise par le lanceur LibreOffice pour produire un classeur d'apercu.
Public Function SG_DemoPourApercu(Optional ByVal filtre As Variant, Optional ByVal fixtures As Variant) As String
    SG_InitialiserBase "INDEPENDANT", 2
    If DAL_Count("T_Utilisateurs", RecNew()) = 0 Then Sec_InitialiserAdministrateur "admin", "Administrateur (demo)", MDP_DEMO
    Sec_Connexion "admin", MDP_DEMO
    Lic_Recharger
    SG_DemoPourApercu = SG_ChargerDemo() & " | Audit : " & Audit_Verifier() & " | Anomalies : " & QC_Anomalies(DAL_First("T_Etablissements", RecFrom("Code=DEMO"))).Count
End Function
