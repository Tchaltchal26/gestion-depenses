Attribute VB_Name = "modStructure"
'===============================================================================
' SG - Module modStructure (cycles, niveaux, filieres, classes)
' Chaque etablissement definit sa propre organisation : un complexe scolaire
' peut avoir simultanement maternelle, primaire, college et lycee.
' Les classes sont rattachees a une annee scolaire : une nouvelle annee recoit
' de nouvelles classes (reconduction), les anciennes restent consultables.
'===============================================================================
Option Explicit

Public Const MODES_EVALUATION As String = "COMPETENCES;NOTES;MIXTE;MODULAIRE"
Public Const TYPES_FILIERE As String = "SERIE;FILIERE;SPECIALITE;OPTION;SECTION"

Public Function Cycle_Creer(ByVal etabId As String, ByVal code As String, ByVal libelle As String, _
                            ByVal modeEvaluation As String, ByVal ordre As Long) As String
    Dim id As String, r As Variant
    Sec_Exiger etabId, "STRUCT", NIV_ADMIN
    ControlerCode etabId, "T_Cycles", code
    If Not InList(modeEvaluation, MODES_EVALUATION) Then SG_Raise SG_ERR_VALIDATION, "Mode d'evaluation inconnu : " & modeEvaluation
    id = DAL_NewId("T_Cycles")
    r = RecFrom("CycleID=" & id & "|EtabID=" & etabId & "|Code=" & UCase$(code) & _
                "|ModeEvaluation=" & modeEvaluation & "|Ordre=" & ordre & "|Actif=1")
    RecSet r, "Libelle", libelle
    DAL_Insert "T_Cycles", r
    Audit_LogCreation etabId, "STRUCT", "T_Cycles", r, ""
    Cycle_Creer = id
End Function

Public Function Niveau_Creer(ByVal etabId As String, ByVal cycleId As String, ByVal code As String, _
                             ByVal libelle As String, ByVal ordre As Long) As String
    Dim id As String, r As Variant
    Sec_Exiger etabId, "STRUCT", NIV_ADMIN
    Etab_ExigerAppartenance "T_Cycles", cycleId, etabId
    ControlerCode etabId, "T_Niveaux", code
    If Len(DAL_First("T_Niveaux", RecFrom("EtabID=" & etabId & "|Ordre=" & ordre))) > 0 Then SG_Raise SG_ERR_DOUBLON, "Un niveau porte deja l'ordre " & ordre & " (l'ordre determine les passages)."
    id = DAL_NewId("T_Niveaux")
    r = RecFrom("NiveauID=" & id & "|EtabID=" & etabId & "|CycleID=" & cycleId & "|Code=" & UCase$(code) & _
                "|Ordre=" & ordre & "|Actif=1")
    RecSet r, "Libelle", libelle
    DAL_Insert "T_Niveaux", r
    Audit_LogCreation etabId, "STRUCT", "T_Niveaux", r, ""
    Niveau_Creer = id
End Function

Public Function Filiere_Creer(ByVal etabId As String, ByVal code As String, ByVal libelle As String, ByVal typeFiliere As String) As String
    Dim id As String, r As Variant
    Sec_Exiger etabId, "STRUCT", NIV_ADMIN
    ControlerCode etabId, "T_Filieres", code
    If Not InList(typeFiliere, TYPES_FILIERE) Then SG_Raise SG_ERR_VALIDATION, "Type de filiere inconnu : " & typeFiliere
    id = DAL_NewId("T_Filieres")
    r = RecFrom("FiliereID=" & id & "|EtabID=" & etabId & "|Code=" & UCase$(code) & "|TypeFiliere=" & typeFiliere & "|Actif=1")
    RecSet r, "Libelle", libelle
    DAL_Insert "T_Filieres", r
    Audit_LogCreation etabId, "STRUCT", "T_Filieres", r, ""
    Filiere_Creer = id
End Function

Public Function Classe_Creer(ByVal etabId As String, ByVal anneeId As String, ByVal niveauId As String, _
                             ByVal filiereId As String, ByVal code As String, ByVal libelle As String, _
                             ByVal capacite As Long, Optional ByVal salle As Variant) As String
    ' Valeurs par defaut (IsMissing : forme compatible Excel et LibreOffice).
    If IsMissing(salle) Then salle = ""
    Dim id As String, r As Variant
    Sec_Exiger etabId, "STRUCT", NIV_ADMIN
    Annee_ExigerModifiable etabId, anneeId
    Etab_ExigerAppartenance "T_Niveaux", niveauId, etabId
    If Len(filiereId) > 0 Then Etab_ExigerAppartenance "T_Filieres", filiereId, etabId
    code = UCase$(Trim$(code))
    If Len(code) = 0 Then SG_Raise SG_ERR_VALIDATION, "Code de classe obligatoire."
    If Len(DAL_First("T_Classes", RecFrom("EtabID=" & etabId & "|AnneeID=" & anneeId & "|Code=" & code))) > 0 Then SG_Raise SG_ERR_DOUBLON, "La classe " & code & " existe deja pour cette annee."
    If capacite < 1 Or capacite > 500 Then SG_Raise SG_ERR_VALIDATION, "Capacite invalide."
    id = DAL_NewId("T_Classes")
    r = RecNew()
    RecSet r, "ClasseID", id
    RecSet r, "EtabID", etabId
    RecSet r, "AnneeID", anneeId
    RecSet r, "NiveauID", niveauId
    RecSet r, "FiliereID", filiereId
    RecSet r, "Code", code
    RecSet r, "Libelle", libelle
    RecSet r, "Capacite", CStr(capacite)
    RecSet r, "Salle", salle
    RecSet r, "Actif", "1"
    DAL_Insert "T_Classes", r
    Audit_LogCreation etabId, "STRUCT", "T_Classes", r, ""
    Classe_Creer = id
End Function

' Cree pour l'annee cible une copie des classes de l'annee source
' (memes niveaux, filieres, codes et capacites). Les classes source restent intactes.
Public Function Classe_ReconduireAnnee(ByVal etabId As String, ByVal anneeSource As String, ByVal anneeCible As String) As Long
    Dim ids As Collection, i As Long, c As Variant, n As Long
    Sec_Exiger etabId, "STRUCT", NIV_ADMIN
    Etab_ExigerAppartenance "T_Annees", anneeSource, etabId
    Annee_ExigerModifiable etabId, anneeCible
    Set ids = DAL_Query("T_Classes", RecFrom("EtabID=" & etabId & "|AnneeID=" & anneeSource & "|Actif=1"))
    Audit_DebutTx "Reconduction des classes"
    For i = 1 To ids.Count
        c = DAL_Get("T_Classes", ids(i))
        If Len(DAL_First("T_Classes", RecFrom("EtabID=" & etabId & "|AnneeID=" & anneeCible & "|Code=" & RecStr(c, "Code")))) = 0 Then
            Classe_Creer etabId, anneeCible, RecStr(c, "NiveauID"), RecStr(c, "FiliereID"), RecStr(c, "Code"), _
                         RecStr(c, "Libelle"), CLng(RecStr(c, "Capacite")), RecStr(c, "Salle")
            n = n + 1
        End If
    Next i
    Audit_FinTx
    Classe_ReconduireAnnee = n
End Function

Public Function Classe_Effectif(ByVal classeId As String) As Long
    Classe_Effectif = DAL_Count("T_Inscriptions", RecFrom("ClasseID=" & classeId & "|Statut=ACTIVE"))
End Function

Public Function Niveau_Ordre(ByVal niveauId As String) As Long
    Niveau_Ordre = CLng(DAL_Value("T_Niveaux", niveauId, "Ordre"))
End Function

Private Sub ControlerCode(ByVal etabId As String, ByVal table As String, ByVal code As String)
    code = UCase$(Trim$(code))
    If Len(code) = 0 Or InStr(code, "|") > 0 Or InStr(code, "=") > 0 Then SG_Raise SG_ERR_VALIDATION, "Code invalide."
    If Len(DAL_First(table, RecFrom("EtabID=" & etabId & "|Code=" & code))) > 0 Then SG_Raise SG_ERR_DOUBLON, "Le code " & code & " existe deja."
End Sub
