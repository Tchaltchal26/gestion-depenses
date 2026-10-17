Attribute VB_Name = "modAnnee"
'===============================================================================
' SG - Module modAnnee (annees scolaires)
' Cycle de vie : PREPARATION -> OUVERTE -> CLOTUREE (-> OUVERTE par
'   reouverture exceptionnelle, motivee et journalisee).
' Regles :
'   - une nouvelle annee cree de NOUVEAUX enregistrements ; rien n'est ecrase ;
'   - une seule annee OUVERTE a la fois par etablissement ;
'   - aucune ecriture dans une annee CLOTUREE ;
'   - la cloture exige une sauvegarde verifiee du jour.
'===============================================================================
Option Explicit

Public Function Annee_Creer(ByVal etabId As String, ByVal code As String, ByVal libelle As String, _
                            ByVal dateDebut As String, ByVal dateFin As String) As String
    Dim aid As String, r As Variant, ids As Collection, i As Long, a As Variant
    Sec_Exiger etabId, "ANNEE", NIV_ADMIN
    code = Trim$(code)
    If Len(code) = 0 Then SG_Raise SG_ERR_VALIDATION, "Code d'annee obligatoire (ex. 2026-2027)."
    If Not IsISODate(dateDebut) Or Not IsISODate(dateFin) Then SG_Raise SG_ERR_VALIDATION, "Dates d'annee invalides."
    If dateFin <= dateDebut Then SG_Raise SG_ERR_VALIDATION, "La date de fin doit suivre la date de debut."
    If Len(DAL_First("T_Annees", RecFrom("EtabID=" & etabId & "|Code=" & code))) > 0 Then SG_Raise SG_ERR_DOUBLON, "Cette annee existe deja pour l'etablissement."
    Set ids = DAL_Query("T_Annees", RecFrom("EtabID=" & etabId))
    For i = 1 To ids.Count
        a = DAL_Get("T_Annees", ids(i))
        If dateDebut <= RecStr(a, "DateFin") And dateFin >= RecStr(a, "DateDebut") Then
            SG_Raise SG_ERR_VALIDATION, "Les dates chevauchent l'annee " & RecStr(a, "Code") & "."
        End If
    Next i
    aid = DAL_NewId("T_Annees")
    r = RecNew()
    RecSet r, "AnneeID", aid
    RecSet r, "EtabID", etabId
    RecSet r, "Code", code
    RecSet r, "Libelle", IIf(Len(libelle) > 0, libelle, "Annee scolaire " & code)
    RecSet r, "DateDebut", dateDebut
    RecSet r, "DateFin", dateFin
    RecSet r, "Statut", "PREPARATION"
    DAL_Insert "T_Annees", r
    Audit_LogCreation etabId, "ANNEE", "T_Annees", r, ""
    Annee_Creer = aid
End Function

Public Function Annee_Statut(ByVal anneeId As String) As String
    Annee_Statut = DAL_Value("T_Annees", anneeId, "Statut")
End Function

Public Function Annee_Ouverte(ByVal etabId As String) As String
    Annee_Ouverte = DAL_First("T_Annees", RecFrom("EtabID=" & etabId & "|Statut=OUVERTE"))
End Function

Public Sub Annee_Ouvrir(ByVal etabId As String, ByVal anneeId As String)
    Dim ouverte As String
    Sec_Exiger etabId, "ANNEE", NIV_ADMIN
    Etab_ExigerAppartenance "T_Annees", anneeId, etabId
    If Annee_Statut(anneeId) <> "PREPARATION" Then SG_Raise SG_ERR_ETAT, "Seule une annee en preparation peut etre ouverte (utiliser la reouverture pour une annee cloturee)."
    ouverte = Annee_Ouverte(etabId)
    If Len(ouverte) > 0 Then SG_Raise SG_ERR_ETAT, "L'annee " & DAL_Value("T_Annees", ouverte, "Code") & " est encore ouverte : la cloturer d'abord."
    DAL_Update "T_Annees", anneeId, "Statut", "OUVERTE"
    DAL_Update "T_Annees", anneeId, "OuverteLe", NowStamp()
    Audit_Log etabId, "ANNEE", "OUVERTURE", "T_Annees", anneeId, "Statut", "PREPARATION", "OUVERTE", ""
End Sub

Public Sub Annee_Cloturer(ByVal etabId As String, ByVal anneeId As String, ByVal motif As String)
    Sec_Exiger etabId, "ANNEE", NIV_ADMIN
    Etab_ExigerAppartenance "T_Annees", anneeId, etabId
    If Annee_Statut(anneeId) <> "OUVERTE" Then SG_Raise SG_ERR_ETAT, "Seule une annee ouverte peut etre cloturee."
    If Not Sauv_ExisteVerifieeDuJour(etabId) Then SG_Raise SG_ERR_ETAT, "Une sauvegarde verifiee du jour est obligatoire avant la cloture."
    Audit_DebutTx "Cloture annee"
    DAL_Update "T_Annees", anneeId, "Statut", "CLOTUREE"
    DAL_Update "T_Annees", anneeId, "ClotureeLe", NowStamp()
    DAL_Update "T_Annees", anneeId, "ClotureePar", Sec_UtilisateurCourant()
    Audit_Log etabId, "ANNEE", "CLOTURE", "T_Annees", anneeId, "Statut", "OUVERTE", "CLOTUREE", motif
    Audit_FinTx
End Sub

' Reouverture exceptionnelle : niveau administration + motif obligatoire.
Public Sub Annee_Reouvrir(ByVal etabId As String, ByVal anneeId As String, ByVal motif As String)
    Dim ouverte As String
    Sec_Exiger etabId, "ANNEE", NIV_ADMIN
    Etab_ExigerAppartenance "T_Annees", anneeId, etabId
    If Len(Trim$(motif)) < 10 Then SG_Raise SG_ERR_VALIDATION, "Motif detaille obligatoire (10 caracteres minimum) pour une reouverture."
    If Annee_Statut(anneeId) <> "CLOTUREE" Then SG_Raise SG_ERR_ETAT, "Seule une annee cloturee peut etre reouverte."
    ouverte = Annee_Ouverte(etabId)
    If Len(ouverte) > 0 Then SG_Raise SG_ERR_ETAT, "Une autre annee est ouverte : la cloturer d'abord."
    DAL_Update "T_Annees", anneeId, "Statut", "OUVERTE"
    Audit_Log etabId, "ANNEE", "REOUVERTURE", "T_Annees", anneeId, "Statut", "CLOTUREE", "OUVERTE", motif
End Sub

' Controle utilise par tous les modules qui ecrivent dans une annee.
Public Sub Annee_ExigerModifiable(ByVal etabId As String, ByVal anneeId As String)
    Etab_ExigerAppartenance "T_Annees", anneeId, etabId
    If Annee_Statut(anneeId) = "CLOTUREE" Then SG_Raise SG_ERR_ETAT, "L'annee " & DAL_Value("T_Annees", anneeId, "Code") & " est cloturee : aucune modification n'est possible."
End Sub
