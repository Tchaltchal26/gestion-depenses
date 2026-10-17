Attribute VB_Name = "modDAL"
'===============================================================================
' SG - Module modDAL (couche d'acces aux donnees)
' Role : SEUL module autorise a lire/ecrire les feuilles de donnees T_*.
'        Les modules metier n'accedent jamais directement aux cellules.
' Principes :
'   - recherche par identifiant et par nom de colonne (jamais d'adresse fixe) ;
'   - toutes les valeurs sont stockees en texte (format "@") afin qu'Excel ne
'     convertisse pas les dates, matricules ou codes ;
'   - aucune fonction de suppression de ligne n'est exposee : les donnees
'     historiques ne sont jamais effacees (statuts, dates de fin, annulations).
' Evolution : cette interface (DAL_*) est le point de remplacement prevu pour
'   un stockage SQL (Access/ACE, SQL Server, PostgreSQL) en mode multiposte.
'===============================================================================
Option Explicit

Private Const XL_UP As Long = -4162
Private mModeTest As Boolean

'-------------------------------------------------------------------------------
' Acces aux feuilles
'-------------------------------------------------------------------------------
Public Function DAL_Ws(ByVal table As String) As Object
    On Error GoTo absente
    Set DAL_Ws = ThisWorkbook.Worksheets(table)
    If DAL_Ws.Name <> table Then GoTo absente
    Exit Function
absente:
    SG_Raise SG_ERR_INTROUVABLE, "Table de donnees absente : " & table & ". Lancer l'initialisation (modSetup)."
End Function

Public Function DAL_TableExiste(ByVal table As String) As Boolean
    Dim ws As Object
    DAL_TableExiste = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = table Then DAL_TableExiste = True: Exit Function
    Next ws
End Function

Public Function DAL_LastRow(ws As Object) As Long
    DAL_LastRow = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
End Function

Private Function NbCols(ws As Object) As Long
    Dim c As Long
    c = 1
    Do While Len(CStr(ws.Cells(1, c + 1).Value)) > 0
        c = c + 1
    Loop
    NbCols = c
End Function

Public Function DAL_Col(ws As Object, ByVal champ As String) As Long
    Dim c As Long, n As Long
    n = NbCols(ws)
    For c = 1 To n
        If StrComp(CStr(ws.Cells(1, c).Value), champ, vbTextCompare) = 0 Then DAL_Col = c: Exit Function
    Next c
    SG_Raise SG_ERR_INTROUVABLE, "Colonne '" & champ & "' absente de " & ws.Name
End Function

' Lit toute la table (en-tetes compris) dans un tableau 2D base 1.
Private Function LireTout(ws As Object, ByRef nLignes As Long, ByRef nCols As Long) As Variant
    nLignes = DAL_LastRow(ws)
    nCols = NbCols(ws)
    If nCols < 2 Then nCols = 2
    LireTout = ws.Range(ws.Cells(1, 1), ws.Cells(nLignes, nCols)).Value
End Function

'-------------------------------------------------------------------------------
' Lecture
'-------------------------------------------------------------------------------
Public Function DAL_FindRow(ByVal table As String, ByVal id As String) As Long
    Dim ws As Object, data As Variant, n As Long, nc As Long, i As Long
    Set ws = DAL_Ws(table)
    data = LireTout(ws, n, nc)
    For i = 2 To n
        If CStr(data(i, 1)) = id Then DAL_FindRow = i: Exit Function
    Next i
    DAL_FindRow = 0
End Function

Public Function DAL_Exists(ByVal table As String, ByVal id As String) As Boolean
    If Len(id) = 0 Then DAL_Exists = False Else DAL_Exists = (DAL_FindRow(table, id) > 0)
End Function

Public Function DAL_Get(ByVal table As String, ByVal id As String) As Variant
    Dim ws As Object, row As Long, c As Long, n As Long, r As Variant
    row = DAL_FindRow(table, id)
    If row = 0 Then SG_Raise SG_ERR_INTROUVABLE, "Enregistrement introuvable : " & table & " / " & id
    Set ws = DAL_Ws(table)
    n = NbCols(ws)
    r = RecNew()
    For c = 1 To n
        RecSet r, CStr(ws.Cells(1, c).Value), NzStr(ws.Cells(row, c).Value)
    Next c
    DAL_Get = r
End Function

Public Function DAL_Value(ByVal table As String, ByVal id As String, ByVal champ As String) As String
    Dim ws As Object, row As Long
    row = DAL_FindRow(table, id)
    If row = 0 Then SG_Raise SG_ERR_INTROUVABLE, "Enregistrement introuvable : " & table & " / " & id
    Set ws = DAL_Ws(table)
    DAL_Value = NzStr(ws.Cells(row, DAL_Col(ws, champ)).Value)
End Function

' Identifiants des lignes dont tous les champs du filtre sont egaux
' (comparaison textuelle insensible a la casse). Filtre vide = toutes.
Public Function DAL_Query(ByVal table As String, filtre As Variant) As Collection
    Dim ws As Object, data As Variant, n As Long, nc As Long, i As Long, k As Long
    Dim res As New Collection, cols() As Long, nf As Long, ok As Boolean
    Set ws = DAL_Ws(table)
    data = LireTout(ws, n, nc)
    nf = RecCount(filtre)
    If nf > 0 Then
        ReDim cols(0 To nf - 1)
        For k = 0 To nf - 1
            cols(k) = DAL_Col(ws, RecKey(filtre, k))
        Next k
    End If
    For i = 2 To n
        ok = (Len(CStr(data(i, 1))) > 0)
        For k = 0 To nf - 1
            If Not ok Then Exit For
            If StrComp(CStr(data(i, cols(k))), CStr(RecValAt(filtre, k)), vbTextCompare) <> 0 Then ok = False
        Next k
        If ok Then res.Add CStr(data(i, 1))
    Next i
    Set DAL_Query = res
End Function

' Comme DAL_Query mais retourne les enregistrements complets (une seule lecture).
Public Function DAL_QueryRecs(ByVal table As String, filtre As Variant) As Collection
    Dim ws As Object, data As Variant, n As Long, nc As Long, i As Long, k As Long, c As Long
    Dim res As New Collection, cols() As Long, nf As Long, ok As Boolean, r As Variant
    Set ws = DAL_Ws(table)
    data = LireTout(ws, n, nc)
    nf = RecCount(filtre)
    If nf > 0 Then
        ReDim cols(0 To nf - 1)
        For k = 0 To nf - 1
            cols(k) = DAL_Col(ws, RecKey(filtre, k))
        Next k
    End If
    For i = 2 To n
        ok = (Len(CStr(data(i, 1))) > 0)
        For k = 0 To nf - 1
            If Not ok Then Exit For
            If StrComp(CStr(data(i, cols(k))), CStr(RecValAt(filtre, k)), vbTextCompare) <> 0 Then ok = False
        Next k
        If ok Then
            r = RecNew()
            For c = 1 To nc
                RecSet r, CStr(data(1, c)), NzStr(data(i, c))
            Next c
            res.Add r
        End If
    Next i
    Set DAL_QueryRecs = res
End Function

Public Function DAL_Count(ByVal table As String, filtre As Variant) As Long
    DAL_Count = DAL_Query(table, filtre).Count
End Function

' Premier identifiant correspondant au filtre, ou "" si aucun.
Public Function DAL_First(ByVal table As String, filtre As Variant) As String
    Dim c As Collection
    Set c = DAL_Query(table, filtre)
    If c.Count = 0 Then DAL_First = "" Else DAL_First = c(1)
End Function

' Lecture brute d'une colonne pour les identifiants donnes (rapports, recherche).
Public Function DAL_Column(ByVal table As String, ByVal champ As String) As Variant
    Dim ws As Object, data As Variant, n As Long, nc As Long, i As Long, c As Long
    Dim out() As String
    Set ws = DAL_Ws(table)
    data = LireTout(ws, n, nc)
    c = DAL_Col(ws, champ)
    If n < 2 Then DAL_Column = Array(): Exit Function
    ReDim out(0 To n - 2, 0 To 1)
    For i = 2 To n
        out(i - 2, 0) = CStr(data(i, 1))
        out(i - 2, 1) = CStr(data(i, c))
    Next i
    DAL_Column = out
End Function

'-------------------------------------------------------------------------------
' Ecriture (bas niveau : la journalisation est faite par la couche metier)
'-------------------------------------------------------------------------------
Public Sub DAL_Insert(ByVal table As String, rec As Variant, Optional ByVal controlerDoublon As Boolean = True)
    Dim ws As Object, row As Long, i As Long, id As String
    Set ws = DAL_Ws(table)
    id = RecStr(rec, CStr(ws.Cells(1, 1).Value))
    If Len(id) = 0 Then SG_Raise SG_ERR_VALIDATION, "Identifiant manquant pour " & table
    If controlerDoublon Then
        If DAL_FindRow(table, id) > 0 Then SG_Raise SG_ERR_DOUBLON, "Identifiant deja utilise : " & id
    End If
    row = DAL_LastRow(ws) + 1
    ' Valider toutes les colonnes avant d'ecrire (ecriture tout ou rien).
    Dim cols() As Long
    ReDim cols(0 To RecCount(rec) - 1)
    For i = 0 To RecCount(rec) - 1
        cols(i) = DAL_Col(ws, RecKey(rec, i))
    Next i
    For i = 0 To RecCount(rec) - 1
        EcrireCellule ws, row, cols(i), RecValAt(rec, i)
    Next i
End Sub

Public Sub DAL_Update(ByVal table As String, ByVal id As String, ByVal champ As String, ByVal valeur As Variant)
    Dim ws As Object, row As Long, c As Long
    Set ws = DAL_Ws(table)
    row = DAL_FindRow(table, id)
    If row = 0 Then SG_Raise SG_ERR_INTROUVABLE, "Enregistrement introuvable : " & table & " / " & id
    c = DAL_Col(ws, champ)
    If c = 1 Then SG_Raise SG_ERR_VALIDATION, "L'identifiant technique n'est pas modifiable"
    EcrireCellule ws, row, c, valeur
End Sub

Private Sub EcrireCellule(ws As Object, ByVal row As Long, ByVal col As Long, ByVal valeur As Variant)
    Dim s As String
    s = SafeField(valeur)
    ' Neutralise l'interpretation en formule (injection) meme si le format
    ' texte a ete retire par erreur.
    If Len(s) > 0 Then
        If InStr(1, "=+@", Left$(s, 1)) > 0 Then s = "'" & s
    End If
    ws.Cells(row, col).NumberFormat = "@"
    ws.Cells(row, col).Value = s
End Sub

'-------------------------------------------------------------------------------
' Table systeme et identifiants
'-------------------------------------------------------------------------------
Public Function Sys_Get(ByVal cle As String, Optional ByVal defaut As Variant) As String
    ' Valeurs par defaut (IsMissing : forme compatible Excel et LibreOffice).
    If IsMissing(defaut) Then defaut = ""
    If DAL_FindRow("T_Systeme", cle) = 0 Then Sys_Get = defaut Else Sys_Get = DAL_Value("T_Systeme", cle, "Valeur")
End Function

Public Sub Sys_Set(ByVal cle As String, ByVal valeur As String)
    If DAL_FindRow("T_Systeme", cle) = 0 Then
        DAL_Insert "T_Systeme", RecFrom("Cle=" & cle & "|Valeur=" & valeur)
    Else
        DAL_Update "T_Systeme", cle, "Valeur", valeur
    End If
End Sub

' Identifiant unique et permanent : PREFIXE-INSTALLATION-SEQUENCE.
' La partie installation evite les collisions lors d'une fusion de donnees
' provenant de plusieurs postes ou etablissements.
Public Function DAL_NewId(ByVal table As String) As String
    Dim prefixe As String, cle As String, n As Long, inst As String
    prefixe = SchemaPrefixe(table)
    cle = "SEQ_" & prefixe
    n = CLng("0" & Sys_Get(cle, "0")) + 1
    Sys_Set cle, CStr(n)
    inst = UCase$(Left$(Sys_Get("INSTALL_ID", "0000"), 4))
    DAL_NewId = prefixe & "-" & inst & "-" & PadLeft(CStr(n), 6)
End Function

'-------------------------------------------------------------------------------
' Reinitialisation : reservee a la suite de tests automatisee (classeur de test)
'-------------------------------------------------------------------------------
Public Sub DAL_ActiverModeTest()
    mModeTest = True
End Sub

Public Function DAL_ModeTest() As Boolean
    DAL_ModeTest = mModeTest
End Function

Public Sub DAL_ViderPourTests(ByVal table As String)
    Dim ws As Object, n As Long
    If Not mModeTest Then SG_Raise SG_ERR_ACCES, "Effacement interdit hors mode test"
    Set ws = DAL_Ws(table)
    n = DAL_LastRow(ws)
    If n >= 2 Then ws.Range(ws.Cells(2, 1), ws.Cells(n, NbCols(ws))).ClearContents
End Sub

' Acces direct pour les tests de falsification uniquement.
Public Sub DAL_FalsifierPourTests(ByVal table As String, ByVal id As String, ByVal champ As String, ByVal valeur As String)
    Dim ws As Object
    If Not mModeTest Then SG_Raise SG_ERR_ACCES, "Interdit hors mode test"
    Set ws = DAL_Ws(table)
    ws.Cells(DAL_FindRow(table, id), DAL_Col(ws, champ)).Value = valeur
End Sub
