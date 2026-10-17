Attribute VB_Name = "modSauvegarde"
'===============================================================================
' SG - Module modSauvegarde (sauvegarde, verification, export, restauration)
' Une sauvegarde n'est consideree fiable qu'apres VERIFICATION : la copie est
' rouverte en lecture seule et son journal d'audit est recontrole ; son
' empreinte de tete doit etre identique a celle du classeur courant.
' La copie contient tout le classeur : en mode CENTRALISE elle contient tous
' les etablissements et exige donc un droit SAUVEG global. Pour remettre a un
' etablissement SES donnees, utiliser l'export cloisonne (Export_EtablissementCSV).
'===============================================================================
Option Explicit

Private Sub ExigerDroitSauvegarde(ByVal etabId As String)
    If Etab_Mode() = "CENTRALISE" Or etabId = "*" Then
        Sec_Exiger "*", "SAUVEG", NIV_ECRITURE
    Else
        If Sec_Niveau("*", "SAUVEG") < NIV_ECRITURE Then Sec_Exiger etabId, "SAUVEG", NIV_ECRITURE
    End If
End Sub

Private Function Separateur() As String
    Separateur = Application.PathSeparator
End Function

' Cree une copie horodatee du classeur puis la verifie. Retourne SauvID.
' (Excel uniquement : SaveCopyAs / Workbooks.Open)
Public Function Sauv_Creer(ByVal etabId As String, ByVal dossier As String, Optional ByVal typeSauv As Variant) As String
    ' Valeurs par defaut (IsMissing : forme compatible Excel et LibreOffice).
    If IsMissing(typeSauv) Then typeSauv = "MANUELLE"
    Dim chemin As String, nb As Long, tete As String, verif As String, wb As Object, sid As String
    ExigerDroitSauvegarde etabId
    If Len(dossier) = 0 Then dossier = ThisWorkbook.Path
    If Right$(dossier, 1) = Separateur() Then dossier = Left$(dossier, Len(dossier) - 1)
    If Len(Dir$(dossier, vbDirectory)) = 0 Then SG_Raise SG_ERR_VALIDATION, "Dossier de sauvegarde introuvable : " & dossier
    ' Journaliser AVANT la copie : l'evenement fait partie de la sauvegarde.
    Audit_Log etabId, "SAUVEG", "SAUVEGARDE_DEBUT", "T_Sauvegardes", "", "", "", dossier, typeSauv
    nb = Audit_NbEvenements()
    tete = Audit_HashTete()
    chemin = dossier & Separateur() & "SG_Sauvegarde_" & CompactStamp() & "_" & RandomHex(4) & ".xlsm"
    ThisWorkbook.SaveCopyAs chemin
    verif = Sauv_VerifierFichier(chemin, tete, nb)
    sid = Sauv_Enregistrer(etabId, chemin, typeSauv, nb, tete, IIf(Left$(verif, 2) = "OK", "VERIFIEE", "ECHEC : " & verif))
    If Left$(verif, 2) <> "OK" Then SG_Raise SG_ERR_ETAT, "La sauvegarde a ete creee mais sa verification a echoue : " & verif
    Sauv_Creer = sid
End Function

' Ouvre une copie en lecture seule et controle son journal.
Public Function Sauv_VerifierFichier(ByVal chemin As String, ByVal teteAttendue As String, ByVal nbAttendu As Long) As String
    Dim wb As Object, ws As Object, r As String, n As Long
    On Error GoTo echec
    Application.EnableEvents = False
    Set wb = Workbooks.Open(Filename:=chemin, ReadOnly:=True, UpdateLinks:=0)
    Application.EnableEvents = True
    Set ws = wb.Worksheets("T_Audit")
    r = Audit_VerifierFeuille(ws, teteAttendue)
    n = DAL_LastRow(ws) - 1
    wb.Close SaveChanges:=False
    If Left$(r, 2) <> "OK" Then Sauv_VerifierFichier = r: Exit Function
    If n < nbAttendu Then Sauv_VerifierFichier = "ALTERE||journal incomplet dans la copie": Exit Function
    Sauv_VerifierFichier = "OK|" & n
    Exit Function
echec:
    Application.EnableEvents = True
    Sauv_VerifierFichier = "ERREUR||" & Err.Description
End Function

Private Function Sauv_Enregistrer(ByVal etabId As String, ByVal chemin As String, ByVal typeSauv As String, _
                                  ByVal nb As Long, ByVal tete As String, ByVal statut As String) As String
    Dim sid As String, r As Variant
    sid = DAL_NewId("T_Sauvegardes")
    r = RecNew()
    RecSet r, "SauvID", sid
    RecSet r, "EtabID", etabId
    RecSet r, "Horodatage", NowStamp()
    RecSet r, "Chemin", chemin
    RecSet r, "TypeSauv", typeSauv
    RecSet r, "NbEvenementsAudit", CStr(nb)
    RecSet r, "HashAudit", tete
    RecSet r, "Statut", statut
    RecSet r, "Utilisateur", Sec_UtilisateurCourant()
    DAL_Insert "T_Sauvegardes", r
    Audit_Log etabId, "SAUVEG", "SAUVEGARDE", "T_Sauvegardes", sid, "Statut", "", statut, chemin, IIf(statut = "VERIFIEE", "OK", "ECHEC")
    Sauv_Enregistrer = sid
End Function

' Simulation reservee au classeur de tests (LibreOffice ne sait pas rouvrir
' une copie comme Excel). Enregistre une sauvegarde VERIFIEE fictive.
Public Function Sauv_SimulerPourTests(ByVal etabId As String) As String
    If Not DAL_ModeTest() Then SG_Raise SG_ERR_ACCES, "Interdit hors mode test"
    ExigerDroitSauvegarde etabId
    Sauv_SimulerPourTests = Sauv_Enregistrer(etabId, "(test)", "TEST", Audit_NbEvenements(), Audit_HashTete(), "VERIFIEE")
End Function

Public Function Sauv_ExisteVerifieeDuJour(ByVal etabId As String) As Boolean
    Dim ids As Collection, i As Long, s As Variant, today As String
    today = TodayISO()
    Set ids = DAL_Query("T_Sauvegardes", RecFrom("Statut=VERIFIEE"))
    For i = 1 To ids.Count
        s = DAL_Get("T_Sauvegardes", ids(i))
        If Left$(RecStr(s, "Horodatage"), 10) = today And (RecStr(s, "EtabID") = etabId Or RecStr(s, "EtabID") = "*") Then
            Sauv_ExisteVerifieeDuJour = True
            Exit Function
        End If
    Next i
    Sauv_ExisteVerifieeDuJour = False
End Function

' Ancre externe : a imprimer ou transmettre hors du classeur (direction,
' courriel, registre papier). Permet de detecter un journal recalcule.
Public Function Sauv_AncreAudit() As String
    Sec_Exiger "*", "AUDIT", NIV_LECTURE
    Sauv_AncreAudit = NowStamp() & " | evenements=" & Audit_NbEvenements() & " | empreinte=" & Audit_HashTete()
    Audit_Log "*", "AUDIT", "EXPORT_ANCRE", "T_Audit", "", "", "", Audit_HashTete(), ""
End Function

'-------------------------------------------------------------------------------
' Export cloisonne (CSV UTF-8) : uniquement les lignes de l'etablissement.
' Toujours disponible, y compris en mode consultation (licence expiree).
'-------------------------------------------------------------------------------
Private Function CsvCellule(ByVal v As String) As String
    ' Neutralise l'injection de formules a l'ouverture dans un tableur.
    If Len(v) > 0 Then
        If InStr(1, "=+-@", Left$(v, 1)) > 0 Then v = "'" & v
    End If
    CsvCellule = """" & Replace(v, """", """""") & """"
End Function

Public Function Export_TableCSV(ByVal etabId As String, ByVal table As String, ByVal chemin As String) As Long
    Dim ws As Object, n As Long, nc As Long, data As Variant, i As Long, c As Long, colEtab As Long
    Dim ligne As String, texte As String, nb As Long, cols As Variant
    If Not SchemaEstCloisonnee(table) And table <> "T_Etablissements" Then SG_Raise SG_ERR_ACCES, "Table non exportable par etablissement : " & table
    If Sec_Niveau(etabId, "SAUVEG") < NIV_LECTURE And Sec_Niveau(etabId, "AUDIT") < NIV_LECTURE Then Sec_Exiger etabId, "SAUVEG", NIV_LECTURE
    Set ws = DAL_Ws(table)
    cols = Split(SchemaColonnes(table), ";")
    nc = UBound(cols) + 1
    n = DAL_LastRow(ws)
    data = ws.Range(ws.Cells(1, 1), ws.Cells(IIf(n < 1, 1, n), nc)).Value
    colEtab = DAL_Col(ws, "EtabID")
    texte = ChrW(&HFEFF)
    For c = 1 To nc
        If c > 1 Then texte = texte & ";"
        texte = texte & CsvCellule(CStr(data(1, c)))
    Next c
    texte = texte & vbCrLf
    For i = 2 To n
        If CStr(data(i, colEtab)) = etabId Then
            ligne = ""
            For c = 1 To nc
                If c > 1 Then ligne = ligne & ";"
                ligne = ligne & CsvCellule(NzStr(data(i, c)))
            Next c
            texte = texte & ligne & vbCrLf
            nb = nb + 1
        End If
    Next i
    EcrireFichierUtf8 chemin, texte
    Audit_Log etabId, "SAUVEG", "EXPORT", table, "", "", "", CStr(nb) & " ligne(s) -> " & chemin, ""
    Export_TableCSV = nb
End Function

Public Function Export_EtablissementCSV(ByVal etabId As String, ByVal dossier As String) As Long
    Dim t As Variant, total As Long, code As String
    code = Etab_Code(etabId)
    For Each t In SchemaTables()
        If SchemaEstCloisonnee(CStr(t)) Or CStr(t) = "T_Etablissements" Then
            total = total + Export_TableCSV(etabId, CStr(t), PathJoin(dossier, code & "_" & t & ".csv"))
        End If
    Next t
    Export_EtablissementCSV = total
End Function

Public Sub EcrireFichierUtf8(ByVal chemin As String, ByVal texte As String)
    Dim f As Integer, b() As Byte, v As Variant, i As Long, n As Long
    v = Utf8Encode(texte)
    n = ByteCount(v)
    If n > 0 Then
        ReDim b(0 To n - 1)
        For i = 0 To n - 1
            b(i) = v(i)
        Next i
    End If
    If Len(Dir$(chemin)) > 0 Then Kill chemin
    f = FreeFile
    Open chemin For Binary Access Write As #f
    If n > 0 Then Put #f, , b
    Close #f
End Sub

Public Function LireFichierUtf8(ByVal chemin As String) As String
    Dim f As Integer, b() As Byte, n As Long
    f = FreeFile
    Open chemin For Binary Access Read As #f
    n = LOF(f)
    If n > 0 Then
        ReDim b(0 To n - 1)
        Get #f, , b
    End If
    Close #f
    If n = 0 Then LireFichierUtf8 = "" Else LireFichierUtf8 = Utf8Decode(b)
End Function
