Attribute VB_Name = "modUtil"
'===============================================================================
' SG - Logiciel universel de gestion scolaire
' Module  : modUtil
' Role    : fonctions utilitaires sans dependance (erreurs, dates, textes,
'           UTF-8, Base64, hexadecimal).
' Regle   : aucun acces aux feuilles ici. Code compatible Excel (Windows/Mac)
'           et LibreOffice (mode VBA) pour permettre les tests automatises.
'===============================================================================
Option Explicit

' Base des numeros d'erreur metier (equivaut a vbObjectError + 1000).
Public Const SG_ERR_BASE As Long = -2147220504
Public Const SG_ERR_VALIDATION As Long = -2147220503
Public Const SG_ERR_ACCES As Long = -2147220502
Public Const SG_ERR_INTROUVABLE As Long = -2147220501
Public Const SG_ERR_DOUBLON As Long = -2147220500
Public Const SG_ERR_ETAT As Long = -2147220499
Public Const SG_ERR_LICENCE As Long = -2147220498
Public Const SG_ERR_ISOLATION As Long = -2147220497
Public Const SG_ERR_LECTURE_SEULE As Long = -2147220496

Public Const SG_VERSION As String = "0.1.0"
Public Const SG_VERSION_MAJEURE As Long = 0

'-------------------------------------------------------------------------------
' Erreurs
'-------------------------------------------------------------------------------
Public Sub SG_Raise(ByVal numero As Long, ByVal message As String)
    Err.Raise numero, "SG", message
End Sub

Public Function SG_CodeErreur(ByVal numero As Long) As String
    Select Case numero
        Case SG_ERR_VALIDATION: SG_CodeErreur = "VALIDATION"
        Case SG_ERR_ACCES: SG_CodeErreur = "ACCES"
        Case SG_ERR_INTROUVABLE: SG_CodeErreur = "INTROUVABLE"
        Case SG_ERR_DOUBLON: SG_CodeErreur = "DOUBLON"
        Case SG_ERR_ETAT: SG_CodeErreur = "ETAT"
        Case SG_ERR_LICENCE: SG_CodeErreur = "LICENCE"
        Case SG_ERR_ISOLATION: SG_CodeErreur = "ISOLATION"
        Case SG_ERR_LECTURE_SEULE: SG_CodeErreur = "LECTURE_SEULE"
        Case Else: SG_CodeErreur = "SYSTEME"
    End Select
End Function

'-------------------------------------------------------------------------------
' Dates : stockage au format ISO texte (AAAA-MM-JJ) pour eviter toute
' ambiguite de format regional et garantir des tris corrects.
'-------------------------------------------------------------------------------
' Horodatage construit sans chaine de format : les codes de minutes ("nn")
' different entre Excel et LibreOffice.
Public Function StampOf(ByVal d As Date) As String
    StampOf = Year(d) & "-" & PadLeft(CStr(Month(d)), 2) & "-" & PadLeft(CStr(Day(d)), 2) & " " & _
              PadLeft(CStr(Hour(d)), 2) & ":" & PadLeft(CStr(Minute(d)), 2) & ":" & PadLeft(CStr(Second(d)), 2)
End Function

Public Function NowStamp() As String
    NowStamp = StampOf(Now)
End Function

' Horodatage compact AAAAMMJJhhmmss (noms de fichiers, references).
Public Function CompactStamp() As String
    CompactStamp = Replace(Replace(Replace(NowStamp(), "-", ""), ":", ""), " ", "")
End Function

Public Function TodayISO() As String
    TodayISO = Format(Date, "yyyy-mm-dd")
End Function

Public Function DateISO(ByVal d As Date) As String
    DateISO = Format(d, "yyyy-mm-dd")
End Function

' Vrai si s est une date ISO AAAA-MM-JJ valide (calendrier reel).
Public Function IsISODate(ByVal s As String) As Boolean
    Dim y As Long, m As Long, j As Long
    IsISODate = False
    If Len(s) <> 10 Then Exit Function
    If Mid$(s, 5, 1) <> "-" Or Mid$(s, 8, 1) <> "-" Then Exit Function
    If Not IsAllDigits(Left$(s, 4)) Or Not IsAllDigits(Mid$(s, 6, 2)) Or Not IsAllDigits(Right$(s, 2)) Then Exit Function
    y = CLng(Left$(s, 4)): m = CLng(Mid$(s, 6, 2)): j = CLng(Right$(s, 2))
    If y < 1900 Or y > 2200 Or m < 1 Or m > 12 Or j < 1 Then Exit Function
    If j > DaysInMonth(y, m) Then Exit Function
    IsISODate = True
End Function

Public Function ISOToDate(ByVal s As String) As Date
    If Not IsISODate(s) Then SG_Raise SG_ERR_VALIDATION, "Date invalide (format attendu AAAA-MM-JJ) : '" & s & "'"
    ISOToDate = DateSerial(CLng(Left$(s, 4)), CLng(Mid$(s, 6, 2)), CLng(Right$(s, 2)))
End Function

Public Function ISOAddDays(ByVal s As String, ByVal n As Long) As String
    ISOAddDays = DateISO(ISOToDate(s) + n)
End Function

Public Function ISODiffDays(ByVal debut As String, ByVal fin As String) As Long
    ISODiffDays = CLng(ISOToDate(fin) - ISOToDate(debut))
End Function

Private Function DaysInMonth(ByVal y As Long, ByVal m As Long) As Long
    Select Case m
        Case 1, 3, 5, 7, 8, 10, 12: DaysInMonth = 31
        Case 4, 6, 9, 11: DaysInMonth = 30
        Case Else
            If (y Mod 4 = 0 And y Mod 100 <> 0) Or (y Mod 400 = 0) Then DaysInMonth = 29 Else DaysInMonth = 28
    End Select
End Function

'-------------------------------------------------------------------------------
' Textes
'-------------------------------------------------------------------------------
Public Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long, c As String
    If Len(s) = 0 Then IsAllDigits = False: Exit Function
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c < "0" Or c > "9" Then IsAllDigits = False: Exit Function
    Next i
    IsAllDigits = True
End Function

Public Function IsBlank(ByVal v As Variant) As Boolean
    If IsEmpty(v) Or IsNull(v) Then IsBlank = True: Exit Function
    IsBlank = (Len(Trim$(CStr(v))) = 0)
End Function

Public Function NzStr(ByVal v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Then NzStr = "" Else NzStr = CStr(v)
End Function

' Normalisation pour recherche et detection de doublons :
' majuscules, sans accents, espaces/tirets/apostrophes unifies.
Public Function NormalizeText(ByVal s As String) As String
    Dim i As Long, c As String, r As String, code As Long, prevSpace As Boolean
    s = UCase$(Trim$(s))
    r = ""
    prevSpace = False
    For i = 1 To Len(s)
        code = AscW(Mid$(s, i, 1))
        c = StripAccent(code)
        If c = " " Or c = "-" Or c = "'" Or c = "." Or code = 8217 Then
            If Not prevSpace And Len(r) > 0 Then r = r & " "
            prevSpace = True
        Else
            r = r & c
            prevSpace = False
        End If
    Next i
    NormalizeText = RTrim$(r)
End Function

Private Function StripAccent(ByVal code As Long) As String
    Select Case code
        Case 192 To 198, 224 To 230: StripAccent = "A"
        Case 199, 231: StripAccent = "C"
        Case 200 To 203, 232 To 235: StripAccent = "E"
        Case 204 To 207, 236 To 239: StripAccent = "I"
        Case 209, 241: StripAccent = "N"
        Case 210 To 214, 216, 242 To 246, 248: StripAccent = "O"
        Case 217 To 220, 249 To 252: StripAccent = "U"
        Case 221, 253, 255, 376: StripAccent = "Y"
        Case 338, 339: StripAccent = "OE"
        Case Else: StripAccent = UCase$(ChrW(code))
    End Select
End Function

' Remplace les separateurs techniques afin qu'une valeur ne puisse pas
' corrompre une serialisation (journal d'audit, fichiers de licence).
Public Function SafeField(ByVal v As Variant) As String
    Dim s As String
    s = NzStr(v)
    s = Replace(s, Chr$(30), " ")
    s = Replace(s, Chr$(31), " ")
    SafeField = s
End Function

Public Function PadLeft(ByVal s As String, ByVal n As Long, Optional ByVal ch As Variant) As String
    ' Valeurs par defaut (IsMissing : forme compatible Excel et LibreOffice).
    If IsMissing(ch) Then ch = "0"
    If Len(s) >= n Then PadLeft = s Else PadLeft = String$(n - Len(s), ch) & s
End Function

'-------------------------------------------------------------------------------
' UTF-8 (implementation portable, sans ADODB)
'-------------------------------------------------------------------------------
Public Function Utf8Encode(ByVal s As String) As Variant
    Dim out() As Byte, n As Long, i As Long, cp As Long, lo As Long
    If Len(s) = 0 Then
        Utf8Encode = out
        Exit Function
    End If
    ReDim out(0 To Len(s) * 4)
    n = 0
    i = 1
    Do While i <= Len(s)
        cp = AscW(Mid$(s, i, 1))
        If cp < 0 Then cp = cp + 65536
        If cp >= &HD800& And cp <= &HDBFF& And i < Len(s) Then
            lo = AscW(Mid$(s, i + 1, 1))
            If lo < 0 Then lo = lo + 65536
            If lo >= &HDC00& And lo <= &HDFFF& Then
                cp = &H10000 + (cp - &HD800&) * 1024 + (lo - &HDC00&)
                i = i + 1
            End If
        End If
        If cp < &H80& Then
            out(n) = cp: n = n + 1
        ElseIf cp < &H800& Then
            out(n) = &HC0& Or (cp \ 64): out(n + 1) = &H80& Or (cp And 63): n = n + 2
        ElseIf cp < &H10000 Then
            out(n) = &HE0& Or (cp \ 4096): out(n + 1) = &H80& Or ((cp \ 64) And 63)
            out(n + 2) = &H80& Or (cp And 63): n = n + 3
        Else
            out(n) = &HF0& Or (cp \ 262144): out(n + 1) = &H80& Or ((cp \ 4096) And 63)
            out(n + 2) = &H80& Or ((cp \ 64) And 63): out(n + 3) = &H80& Or (cp And 63): n = n + 4
        End If
        i = i + 1
    Loop
    ReDim Preserve out(0 To n - 1)
    Utf8Encode = out
End Function

Public Function Utf8Decode(b As Variant) As String
    Dim i As Long, n As Long, cp As Long, r As String, c As Long
    n = ByteCount(b)
    r = ""
    i = 0
    Do While i < n
        c = b(i)
        If c < &H80& Then
            cp = c: i = i + 1
        ElseIf (c And &HE0&) = &HC0& And i + 1 < n Then
            cp = (c And 31) * 64 + (b(i + 1) And 63): i = i + 2
        ElseIf (c And &HF0&) = &HE0& And i + 2 < n Then
            cp = (c And 15) * 4096 + (b(i + 1) And 63) * 64 + (b(i + 2) And 63): i = i + 3
        ElseIf (c And &HF8&) = &HF0& And i + 3 < n Then
            cp = (c And 7) * 262144 + (b(i + 1) And 63) * 4096 + (b(i + 2) And 63) * 64 + (b(i + 3) And 63): i = i + 4
        Else
            cp = 65533: i = i + 1
        End If
        If cp >= &H10000 Then
            cp = cp - &H10000
            r = r & ChrW(&HD800& + (cp \ 1024)) & ChrW(&HDC00& + (cp And 1023))
        Else
            r = r & ChrW(cp)
        End If
    Loop
    ' Ignore une eventuelle marque d'ordre des octets (BOM).
    If Len(r) > 0 Then
        If AscW(Left$(r, 1)) = &HFEFF& Then r = Mid$(r, 2)
    End If
    Utf8Decode = r
End Function

' Nombre d'octets d'un tableau (0 si non dimensionne).
Public Function ByteCount(b As Variant) As Long
    On Error GoTo vide
    ByteCount = UBound(b) - LBound(b) + 1
    Exit Function
vide:
    ByteCount = 0
End Function

'-------------------------------------------------------------------------------
' Hexadecimal
'-------------------------------------------------------------------------------
Public Function BytesToHex(b As Variant) As String
    Dim i As Long, r As String
    r = ""
    For i = 0 To ByteCount(b) - 1
        r = r & Right$("0" & LCase$(Hex$(b(i))), 2)
    Next i
    BytesToHex = r
End Function

Public Function HexToBytes(ByVal h As String) As Variant
    Dim out() As Byte, i As Long, n As Long
    h = Replace(Replace(Trim$(h), " ", ""), vbLf, "")
    If Len(h) Mod 2 = 1 Then h = "0" & h
    n = Len(h) \ 2
    If n = 0 Then
        HexToBytes = out
        Exit Function
    End If
    ReDim out(0 To n - 1)
    For i = 0 To n - 1
        out(i) = CByte(CLng("&H" & Mid$(h, i * 2 + 1, 2)))
    Next i
    HexToBytes = out
End Function

'-------------------------------------------------------------------------------
' Base64
'-------------------------------------------------------------------------------
Private Function B64Alphabet() As String
    B64Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
End Function

Public Function Base64Encode(b As Variant) As String
    Dim i As Long, n As Long, r As String, a As String, v As Long, rest As Long
    a = B64Alphabet()
    n = ByteCount(b)
    r = ""
    For i = 0 To n - 1 Step 3
        rest = n - i
        v = CLng(b(i)) * 65536
        If rest > 1 Then v = v + CLng(b(i + 1)) * 256
        If rest > 2 Then v = v + CLng(b(i + 2))
        r = r & Mid$(a, (v \ 262144) + 1, 1) & Mid$(a, ((v \ 4096) And 63) + 1, 1)
        If rest > 1 Then r = r & Mid$(a, ((v \ 64) And 63) + 1, 1) Else r = r & "="
        If rest > 2 Then r = r & Mid$(a, (v And 63) + 1, 1) Else r = r & "="
    Next i
    Base64Encode = r
End Function

Public Function Base64Decode(ByVal s As String) As Variant
    Dim out() As Byte, clean As String, i As Long, c As String, p As Long
    Dim v As Long, nb As Long, n As Long, pad As Long, a As String
    a = B64Alphabet()
    clean = ""
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c = "=" Then
            pad = pad + 1
        ElseIf InStr(1, a, c, vbBinaryCompare) > 0 Then
            If pad > 0 Then SG_Raise SG_ERR_VALIDATION, "Base64 invalide"
            clean = clean & c
        ElseIf c <> " " And c <> vbCr And c <> vbLf And c <> vbTab Then
            SG_Raise SG_ERR_VALIDATION, "Base64 invalide"
        End If
    Next i
    If (Len(clean) + pad) Mod 4 <> 0 Then SG_Raise SG_ERR_VALIDATION, "Base64 invalide (longueur)"
    nb = (Len(clean) * 6) \ 8
    If nb = 0 Then
        Base64Decode = out
        Exit Function
    End If
    ReDim out(0 To nb - 1)
    n = 0: v = 0
    Dim bits As Long
    bits = 0
    For i = 1 To Len(clean)
        p = InStr(1, a, Mid$(clean, i, 1), vbBinaryCompare) - 1
        v = (v * 64 + p) And &HFFFFFF
        bits = bits + 6
        If bits >= 8 Then
            bits = bits - 8
            If n < nb Then out(n) = (v \ (2 ^ bits)) And 255
            n = n + 1
        End If
    Next i
    Base64Decode = out
End Function

'-------------------------------------------------------------------------------
' Divers
'-------------------------------------------------------------------------------
Public Function RandomHex(ByVal nChars As Long) As String
    Dim i As Long, r As String
    Randomize
    r = ""
    For i = 1 To nChars
        r = r & LCase$(Hex$(Int(Rnd * 16)))
    Next i
    RandomHex = r
End Function

Public Function JoinCollection(ByVal c As Collection, ByVal sep As String) As String
    Dim i As Long, r As String
    r = ""
    For i = 1 To c.Count
        If i > 1 Then r = r & sep
        r = r & CStr(c(i))
    Next i
    JoinCollection = r
End Function

' Joint un dossier et un nom de fichier avec le separateur du dossier.
Public Function PathJoin(ByVal dossier As String, ByVal nom As String) As String
    Dim sep As String
    If InStr(dossier, "\") > 0 Then sep = "\" Else sep = "/"
    If Right$(dossier, 1) = sep Then PathJoin = dossier & nom Else PathJoin = dossier & sep & nom
End Function

Public Function InList(ByVal value As String, ByVal listSemicolon As String) As Boolean
    Dim parts As Variant, i As Long
    parts = Split(listSemicolon, ";")
    For i = LBound(parts) To UBound(parts)
        If UCase$(Trim$(parts(i))) = UCase$(Trim$(value)) Then InList = True: Exit Function
    Next i
    InList = False
End Function
