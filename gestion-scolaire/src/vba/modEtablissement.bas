Attribute VB_Name = "modEtablissement"
'===============================================================================
' SG - Module modEtablissement (etablissements, parametres, numerotations)
' Mode INDEPENDANT : un seul etablissement par classeur.
' Mode CENTRALISE  : plusieurs etablissements, donnees cloisonnees par EtabID.
'===============================================================================
Option Explicit

Public Const TYPES_ETAB As String = "MATERNELLE;PRIMAIRE;COLLEGE;LYCEE;LYCEE_TECHNIQUE;CENTRE_FORMATION;COMPLEXE"
Public Const STATUTS_ETAB As String = "PUBLIC;PRIVE;CONFESSIONNEL"

Public Function Etab_Mode() As String
    Etab_Mode = Sys_Get("MODE", "INDEPENDANT")
End Function

Public Function Etab_Creer(rec As Variant) As String
    Dim code As String, eid As String, r As Variant, nb As Long, maxE As Long
    Sec_Exiger "*", "ETAB", NIV_ADMIN
    code = UCase$(RecStr(rec, "Code"))
    If Len(code) < 2 Or Len(code) > 12 Or InStr(code, ";") > 0 Or InStr(code, " ") > 0 Then SG_Raise SG_ERR_VALIDATION, "Code etablissement invalide (2 a 12 caracteres, sans espace)."
    If Len(RecStr(rec, "Nom")) = 0 Then SG_Raise SG_ERR_VALIDATION, "Le nom de l'etablissement est obligatoire."
    If Not InList(RecStr(rec, "TypeEtab"), TYPES_ETAB) Then SG_Raise SG_ERR_VALIDATION, "Type d'etablissement inconnu : " & RecStr(rec, "TypeEtab")
    If Len(RecStr(rec, "Statut")) > 0 And Not InList(RecStr(rec, "Statut"), STATUTS_ETAB) Then SG_Raise SG_ERR_VALIDATION, "Statut inconnu."
    If Len(DAL_First("T_Etablissements", RecFrom("Code=" & code))) > 0 Then SG_Raise SG_ERR_DOUBLON, "Ce code d'etablissement existe deja."
    nb = DAL_Count("T_Etablissements", RecNew())
    If Etab_Mode() = "INDEPENDANT" And nb >= 1 Then SG_Raise SG_ERR_ETAT, "Mode independant : un seul etablissement par installation."
    maxE = Lic_NbEtabsMax()
    If maxE > 0 And nb >= maxE Then SG_Raise SG_ERR_LICENCE, "Nombre d'etablissements autorise par la licence atteint (" & maxE & ")."
    If Not Lic_EtabAutorise(code) Then SG_Raise SG_ERR_LICENCE, "Le code " & code & " ne figure pas dans la licence."

    eid = DAL_NewId("T_Etablissements")
    r = RecNew()
    RecSet r, "EtabID", eid
    RecSet r, "Code", code
    RecSet r, "Nom", RecStr(rec, "Nom")
    RecSet r, "TypeEtab", RecStr(rec, "TypeEtab")
    RecSet r, "Statut", RecStr(rec, "Statut")
    RecSet r, "Pays", RecStr(rec, "Pays")
    RecSet r, "Ville", RecStr(rec, "Ville")
    RecSet r, "Adresse", RecStr(rec, "Adresse")
    RecSet r, "Telephone", RecStr(rec, "Telephone")
    RecSet r, "Email", RecStr(rec, "Email")
    RecSet r, "Devise", IIf(Len(RecStr(rec, "Devise")) > 0, RecStr(rec, "Devise"), "XOF")
    RecSet r, "Langue", IIf(Len(RecStr(rec, "Langue")) > 0, RecStr(rec, "Langue"), "fr")
    RecSet r, "Actif", "1"
    RecSet r, "DateCreation", NowStamp()
    RecSet r, "CreePar", Sec_UtilisateurCourant()
    DAL_Insert "T_Etablissements", r
    Audit_LogCreation eid, "ETAB", "T_Etablissements", r, ""
    ParametresParDefaut eid, code
    Etab_Creer = eid
End Function

Private Sub ParametresParDefaut(ByVal eid As String, ByVal code As String)
    ' Parametres modifiables sans toucher au code.
    InsererParam eid, "FORMAT_MATRICULE", "{ETAB}-{AAAA}-{N5}"
    InsererParam eid, "NOTE_MAX", "20"
    InsererParam eid, "CAPACITE_STRICTE", "1"
    InsererParam eid, "AGE_MIN_INSCRIPTION", "2"
    InsererParam eid, "MOTIF_OBLIGATOIRE_CHAMPS", "Nom;Prenoms;DateNaissance;Sexe"
End Sub

Private Sub InsererParam(ByVal eid As String, ByVal cle As String, ByVal valeur As String)
    DAL_Insert "T_Parametres", RecFrom("ParamID=" & DAL_NewId("T_Parametres") & "|EtabID=" & eid & "|Cle=" & cle & "|Valeur=" & valeur & "|ModifieLe=" & NowStamp() & "|ModifiePar=" & Sec_UtilisateurCourant())
End Sub

Public Function Param_Get(ByVal etabId As String, ByVal cle As String, Optional ByVal defaut As Variant) As String
    ' Valeurs par defaut (IsMissing : forme compatible Excel et LibreOffice).
    If IsMissing(defaut) Then defaut = ""
    Dim pid As String
    pid = DAL_First("T_Parametres", RecFrom("EtabID=" & etabId & "|Cle=" & cle))
    If Len(pid) = 0 Then Param_Get = defaut Else Param_Get = DAL_Value("T_Parametres", pid, "Valeur")
End Function

Public Sub Param_Set(ByVal etabId As String, ByVal cle As String, ByVal valeur As String, ByVal motif As String)
    Dim pid As String, ancien As String
    Sec_Exiger etabId, "PARAM", NIV_ADMIN
    If Len(Trim$(cle)) = 0 Then SG_Raise SG_ERR_VALIDATION, "Cle de parametre vide."
    pid = DAL_First("T_Parametres", RecFrom("EtabID=" & etabId & "|Cle=" & cle))
    If Len(pid) = 0 Then
        InsererParam etabId, cle, valeur
        pid = DAL_First("T_Parametres", RecFrom("EtabID=" & etabId & "|Cle=" & cle))
    Else
        ancien = DAL_Value("T_Parametres", pid, "Valeur")
        DAL_Update "T_Parametres", pid, "Valeur", valeur
        DAL_Update "T_Parametres", pid, "ModifieLe", NowStamp()
        DAL_Update "T_Parametres", pid, "ModifiePar", Sec_UtilisateurCourant()
    End If
    Audit_Log etabId, "PARAM", "MODIFICATION", "T_Parametres", pid, cle, ancien, valeur, motif
End Sub

' Compteur sequentiel par etablissement (matricules, numeros de documents).
Public Function Compteur_Suivant(ByVal etabId As String, ByVal cle As String) As Long
    Dim cid As String, n As Long
    cid = DAL_First("T_Compteurs", RecFrom("EtabID=" & etabId & "|Cle=" & cle))
    If Len(cid) = 0 Then
        n = 1
        DAL_Insert "T_Compteurs", RecFrom("CompteurID=" & DAL_NewId("T_Compteurs") & "|EtabID=" & etabId & "|Cle=" & cle & "|Valeur=1")
    Else
        n = CLng(DAL_Value("T_Compteurs", cid, "Valeur")) + 1
        DAL_Update "T_Compteurs", cid, "Valeur", CStr(n)
    End If
    Compteur_Suivant = n
End Function

' Applique un format de numerotation : {ETAB} code, {AAAA} annee, {N5} numero sur 5 chiffres.
Public Function Format_Numero(ByVal modele As String, ByVal codeEtab As String, ByVal annee As String, ByVal n As Long) As String
    Dim s As String, k As Long
    s = Replace(modele, "{ETAB}", codeEtab)
    s = Replace(s, "{AAAA}", annee)
    For k = 1 To 9
        s = Replace(s, "{N" & k & "}", PadLeft(CStr(n), k))
    Next k
    s = Replace(s, "{N}", CStr(n))
    Format_Numero = s
End Function

Public Function Etab_Code(ByVal etabId As String) As String
    Etab_Code = DAL_Value("T_Etablissements", etabId, "Code")
End Function

' Liste des etablissements visibles par l'utilisateur courant.
Public Function Etab_Accessibles() As Collection
    Dim tous As Collection, res As New Collection, i As Long
    Set tous = DAL_Query("T_Etablissements", RecNew())
    For i = 1 To tous.Count
        If Sec_AAccesEtab(tous(i)) Then res.Add tous(i)
    Next i
    Set Etab_Accessibles = res
End Function

' Controle d'appartenance : toute reference a un enregistrement cloisonne
' doit appartenir a l'etablissement de l'operation.
Public Sub Etab_ExigerAppartenance(ByVal table As String, ByVal id As String, ByVal etabId As String)
    If Len(id) = 0 Then SG_Raise SG_ERR_VALIDATION, "Reference manquante (" & table & ")."
    If Not DAL_Exists(table, id) Then SG_Raise SG_ERR_INTROUVABLE, "Reference introuvable : " & table & " / " & id
    If DAL_Value(table, id, "EtabID") <> etabId Then
        Audit_Log etabId, "ETAB", "TENTATIVE_INTER_ETAB", table, id, "", "", "", "", "REFUSE"
        SG_Raise SG_ERR_ISOLATION, "Cet enregistrement n'appartient pas a l'etablissement courant."
    End If
End Sub
