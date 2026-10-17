Attribute VB_Name = "modSetup"
'===============================================================================
' SG - Module modSetup (initialisation et migration du schema)
' SG_InitialiserBase : cree les tables manquantes, les parametres systeme et
'   les roles par defaut. Peut etre relancee sans risque (idempotente).
' SG_MigrerSchema    : ajoute les colonnes manquantes A LA FIN des tables
'   existantes. Ne supprime, ne renomme et ne deplace jamais de colonne.
'===============================================================================
Option Explicit

Public Sub SG_InitialiserBase(Optional ByVal mode As Variant, Optional ByVal iterationsMdp As Long = 1000)
    ' Valeurs par defaut (IsMissing : forme compatible Excel et LibreOffice).
    If IsMissing(mode) Then mode = "INDEPENDANT"
    Dim t As Variant
    If Not InList(mode, "INDEPENDANT;CENTRALISE") Then SG_Raise SG_ERR_VALIDATION, "Mode inconnu : " & mode
    For Each t In SchemaTables()
        CreerTable CStr(t)
    Next t
    SG_MigrerSchema
    If Len(Sys_Get("INSTALL_ID", "")) = 0 Then
        Sys_Set "INSTALL_ID", Left$(SHA256Hex(RandomHex(32) & "|" & CompactStamp() & "|" & CStr(Timer) & "|" & Environ$("COMPUTERNAME")), 32)
        Sys_Set "INSTALL_DATE", TodayISO()
        Sys_Set "MODE", mode
        Sys_Set "ITER_MDP", CStr(iterationsMdp)
        Audit_Log "*", "PARAM", "INITIALISATION", "T_Systeme", "INSTALL_ID", "", "", "Mode " & mode & " ; version " & SG_VERSION, ""
    End If
    Sys_Set "SCHEMA_VERSION", CStr(SCHEMA_VERSION)
    Sys_Set "APP_VERSION", SG_VERSION
    Sec_CreerRolesParDefaut
End Sub

Private Sub CreerTable(ByVal table As String)
    Dim ws As Object, cols As Variant, i As Long
    If DAL_TableExiste(table) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = table
    cols = Split(SchemaColonnes(table), ";")
    ws.Columns.NumberFormat = "@"
    For i = 0 To UBound(cols)
        ws.Cells(1, i + 1).Value = cols(i)
    Next i
    ws.Rows(1).Font.Bold = True
End Sub

Public Sub SG_MigrerSchema()
    Dim t As Variant, ws As Object, cols As Variant, i As Long, c As Long, trouve As Boolean, derniere As Long
    For Each t In SchemaTables()
        Set ws = DAL_Ws(CStr(t))
        cols = Split(SchemaColonnes(CStr(t)), ";")
        For i = 0 To UBound(cols)
            trouve = False
            derniere = 0
            c = 1
            Do While Len(CStr(ws.Cells(1, c).Value)) > 0
                If CStr(ws.Cells(1, c).Value) = cols(i) Then trouve = True
                derniere = c
                c = c + 1
            Loop
            If Not trouve Then
                ws.Columns(derniere + 1).NumberFormat = "@"
                ws.Cells(1, derniere + 1).Value = cols(i)
                If DAL_TableExiste("T_Audit") And CStr(t) <> "T_Audit" Then
                    If DAL_LastRow(DAL_Ws("T_Systeme")) >= 1 Then
                        Audit_Log "*", "PARAM", "MIGRATION_SCHEMA", CStr(t), "", CStr(cols(i)), "", "colonne ajoutee", "Schema v" & SCHEMA_VERSION
                    End If
                End If
            End If
        Next i
    Next t
End Sub

' Controle de coherence du schema (utilise par le controle qualite).
Public Function SG_VerifierSchema() As String
    Dim t As Variant, ws As Object, cols As Variant, i As Long, manquants As String
    manquants = ""
    For Each t In SchemaTables()
        If Not DAL_TableExiste(CStr(t)) Then
            manquants = manquants & CStr(t) & " (table) ; "
        Else
            Set ws = DAL_Ws(CStr(t))
            cols = Split(SchemaColonnes(CStr(t)), ";")
            If CStr(ws.Cells(1, 1).Value) <> cols(0) Then manquants = manquants & CStr(t) & " (identifiant) ; "
            For i = 0 To UBound(cols)
                On Error Resume Next
                Err.Clear
                DAL_Col ws, CStr(cols(i))
                If Err.Number <> 0 Then manquants = manquants & CStr(t) & "." & cols(i) & " ; "
                On Error GoTo 0
            Next i
        End If
    Next t
    If Len(manquants) = 0 Then SG_VerifierSchema = "OK" Else SG_VerifierSchema = "INCOMPLET : " & manquants
End Function
