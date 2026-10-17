Attribute VB_Name = "modRec"
'===============================================================================
' SG - Module modRec
' Role : "enregistrement" leger (paires champ/valeur) utilise entre les couches.
'        Remplace Scripting.Dictionary, indisponible hors Windows.
' Forme : Array(tableauDesChamps, tableauDesValeurs)
'===============================================================================
Option Explicit

Public Function RecNew() As Variant
    RecNew = Array(Array(), Array())
End Function

Public Function RecCount(r As Variant) As Long
    RecCount = UBound(r(0)) + 1
End Function

Public Function RecKey(r As Variant, ByVal i As Long) As String
    RecKey = r(0)(i)
End Function

Public Function RecValAt(r As Variant, ByVal i As Long) As Variant
    RecValAt = r(1)(i)
End Function

Private Function RecIndex(r As Variant, ByVal champ As String) As Long
    Dim ks As Variant, i As Long
    ks = r(0)
    For i = 0 To UBound(ks)
        If StrComp(ks(i), champ, vbTextCompare) = 0 Then RecIndex = i: Exit Function
    Next i
    RecIndex = -1
End Function

Public Function RecHas(r As Variant, ByVal champ As String) As Boolean
    RecHas = (RecIndex(r, champ) >= 0)
End Function

Public Sub RecSet(r As Variant, ByVal champ As String, ByVal valeur As Variant)
    Dim ks As Variant, vs As Variant, i As Long, n As Long
    i = RecIndex(r, champ)
    vs = r(1)
    If i >= 0 Then
        vs(i) = valeur
        r(1) = vs
        Exit Sub
    End If
    ks = r(0)
    n = UBound(ks) + 1
    ReDim Preserve ks(0 To n)
    ReDim Preserve vs(0 To n)
    ks(n) = champ
    vs(n) = valeur
    r(0) = ks
    r(1) = vs
End Sub

Public Function RecGet(r As Variant, ByVal champ As String, Optional ByVal defaut As Variant = "") As Variant
    Dim i As Long
    i = RecIndex(r, champ)
    If i < 0 Then RecGet = defaut Else RecGet = r(1)(i)
End Function

Public Function RecStr(r As Variant, ByVal champ As String) As String
    RecStr = Trim$(NzStr(RecGet(r, champ, "")))
End Function

' Construit un enregistrement a partir d'une liste "champ=valeur|champ=valeur".
' Pratique pour les tests et les donnees de demonstration.
Public Function RecFrom(ByVal spec As String) As Variant
    Dim r As Variant, parts As Variant, i As Long, p As Long
    r = RecNew()
    If Len(spec) = 0 Then RecFrom = r: Exit Function
    parts = Split(spec, "|")
    For i = LBound(parts) To UBound(parts)
        p = InStr(1, parts(i), "=")
        If p > 0 Then RecSet r, Left$(parts(i), p - 1), Mid$(parts(i), p + 1)
    Next i
    RecFrom = r
End Function
