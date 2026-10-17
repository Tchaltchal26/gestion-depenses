Attribute VB_Name = "modSHA256"
'===============================================================================
' SG - Module modSHA256
' Role : empreinte SHA-256 (FIPS 180-4) en VBA pur, sans API Windows, afin de
'        fonctionner a l'identique sous Excel Windows, Excel Mac et LibreOffice.
' Usage: chainage du journal d'audit, empreinte d'installation, verification
'        des licences (avec modRSA), hachage des mots de passe.
' Note : les entiers VBA (Long) sont signes 32 bits ; les additions modulo 2^32
'        passent par des Double pour eviter les depassements.
'===============================================================================
Option Explicit

Private K(0 To 63) As Long
Private P2(0 To 31) As Double
Private mInit As Boolean

Private Sub InitTables()
    Dim kk As Variant, i As Long
    kk = Array("428a2f98", "71374491", "b5c0fbcf", "e9b5dba5", "3956c25b", "59f111f1", "923f82a4", "ab1c5ed5", _
               "d807aa98", "12835b01", "243185be", "550c7dc3", "72be5d74", "80deb1fe", "9bdc06a7", "c19bf174", _
               "e49b69c1", "efbe4786", "0fc19dc6", "240ca1cc", "2de92c6f", "4a7484aa", "5cb0a9dc", "76f988da", _
               "983e5152", "a831c66d", "b00327c8", "bf597fc7", "c6e00bf3", "d5a79147", "06ca6351", "14292967", _
               "27b70a85", "2e1b2138", "4d2c6dfc", "53380d13", "650a7354", "766a0abb", "81c2c92e", "92722c85", _
               "a2bfe8a1", "a81a664b", "c24b8b70", "c76c51a3", "d192e819", "d6990624", "f40e3585", "106aa070", _
               "19a4c116", "1e376c08", "2748774c", "34b0bcb5", "391c0cb3", "4ed8aa4a", "5b9cca4f", "682e6ff3", _
               "748f82ee", "78a5636f", "84c87814", "8cc70208", "90befffa", "a4506ceb", "bef9a3f7", "c67178f2")
    For i = 0 To 63
        K(i) = HexToLong(CStr(kk(i)))
    Next i
    P2(0) = 1
    For i = 1 To 31
        P2(i) = P2(i - 1) * 2
    Next i
    mInit = True
End Sub

Private Function HexToLong(ByVal h As String) As Long
    Dim d As Double
    d = CDbl(CLng("&H" & Left$(h, 4))) * 65536# + CDbl(CLng("&H" & Right$(h, 4)))
    HexToLong = ToSigned(d)
End Function

Private Function ToUnsigned(ByVal x As Long) As Double
    If x < 0 Then ToUnsigned = CDbl(x) + 4294967296# Else ToUnsigned = CDbl(x)
End Function

Private Function ToSigned(ByVal d As Double) As Long
    d = d - Int(d / 4294967296#) * 4294967296#
    If d >= 2147483648# Then ToSigned = CLng(d - 4294967296#) Else ToSigned = CLng(d)
End Function

Private Function Add32(ByVal a As Long, ByVal b As Long) As Long
    Add32 = ToSigned(ToUnsigned(a) + ToUnsigned(b))
End Function

' Decalage logique a droite (non signe).
Private Function Shr(ByVal x As Long, ByVal n As Long) As Long
    If n = 0 Then Shr = x: Exit Function
    Shr = ToSigned(Int(ToUnsigned(x) / P2(n)))
End Function

' Decalage a gauche modulo 2^32.
Private Function Shl(ByVal x As Long, ByVal n As Long) As Long
    If n = 0 Then Shl = x: Exit Function
    Dim d As Double
    d = ToUnsigned(x) * P2(n)
    Shl = ToSigned(d - Int(d / 4294967296#) * 4294967296#)
End Function

Private Function Rotr(ByVal x As Long, ByVal n As Long) As Long
    Rotr = Shr(x, n) Or Shl(x, 32 - n)
End Function

Public Function SHA256Bytes(msg As Variant) As Variant
    Dim H(0 To 7) As Long, W(0 To 63) As Long
    Dim a As Long, b As Long, c As Long, d As Long, e As Long, f As Long, g As Long, hh As Long
    Dim t1 As Long, t2 As Long, s0 As Long, s1 As Long
    Dim n As Long, total As Long, blocks As Long, i As Long, j As Long, blk As Long, idx As Long
    Dim padded() As Byte, bitLen As Double, out(0 To 31) As Byte

    If Not mInit Then InitTables
    n = ByteCount(msg)
    total = ((n + 9 + 63) \ 64) * 64
    ReDim padded(0 To total - 1)
    For i = 0 To n - 1
        padded(i) = msg(LBound(msg) + i)
    Next i
    padded(n) = &H80
    bitLen = CDbl(n) * 8#
    For i = 0 To 7
        padded(total - 1 - i) = CByte(bitLen - Int(bitLen / 256#) * 256#)
        bitLen = Int(bitLen / 256#)
    Next i

    H(0) = HexToLong("6a09e667"): H(1) = HexToLong("bb67ae85")
    H(2) = HexToLong("3c6ef372"): H(3) = HexToLong("a54ff53a")
    H(4) = HexToLong("510e527f"): H(5) = HexToLong("9b05688c")
    H(6) = HexToLong("1f83d9ab"): H(7) = HexToLong("5be0cd19")

    blocks = total \ 64
    For blk = 0 To blocks - 1
        For j = 0 To 15
            idx = blk * 64 + j * 4
            W(j) = ToSigned(CDbl(padded(idx)) * 16777216# + CDbl(padded(idx + 1)) * 65536# + _
                            CDbl(padded(idx + 2)) * 256# + CDbl(padded(idx + 3)))
        Next j
        For j = 16 To 63
            s0 = Rotr(W(j - 15), 7) Xor Rotr(W(j - 15), 18) Xor Shr(W(j - 15), 3)
            s1 = Rotr(W(j - 2), 17) Xor Rotr(W(j - 2), 19) Xor Shr(W(j - 2), 10)
            W(j) = ToSigned(ToUnsigned(W(j - 16)) + ToUnsigned(s0) + ToUnsigned(W(j - 7)) + ToUnsigned(s1))
        Next j
        a = H(0): b = H(1): c = H(2): d = H(3): e = H(4): f = H(5): g = H(6): hh = H(7)
        For j = 0 To 63
            s1 = Rotr(e, 6) Xor Rotr(e, 11) Xor Rotr(e, 25)
            t1 = ToSigned(ToUnsigned(hh) + ToUnsigned(s1) + ToUnsigned((e And f) Xor ((Not e) And g)) + _
                          ToUnsigned(K(j)) + ToUnsigned(W(j)))
            s0 = Rotr(a, 2) Xor Rotr(a, 13) Xor Rotr(a, 22)
            t2 = Add32(s0, (a And b) Xor (a And c) Xor (b And c))
            hh = g: g = f: f = e
            e = Add32(d, t1)
            d = c: c = b: b = a
            a = Add32(t1, t2)
        Next j
        H(0) = Add32(H(0), a): H(1) = Add32(H(1), b): H(2) = Add32(H(2), c): H(3) = Add32(H(3), d)
        H(4) = Add32(H(4), e): H(5) = Add32(H(5), f): H(6) = Add32(H(6), g): H(7) = Add32(H(7), hh)
    Next blk

    Dim u As Double
    For i = 0 To 7
        u = ToUnsigned(H(i))
        For j = 3 To 0 Step -1
            out(i * 4 + j) = CByte(u - Int(u / 256#) * 256#)
            u = Int(u / 256#)
        Next j
    Next i
    SHA256Bytes = out
End Function

' Empreinte hexadecimale (minuscules) du texte encode en UTF-8.
Public Function SHA256Hex(ByVal texte As String) As String
    Dim b As Variant, h As Variant
    b = Utf8Encode(texte)
    h = SHA256Bytes(b)
    SHA256Hex = BytesToHex(h)
End Function

' Derivation de mot de passe : SHA-256 itere avec sel.
' Limite documentee : bien plus faible qu'Argon2/PBKDF2 natif ; protege
' uniquement contre la lecture directe des mots de passe dans le classeur.
Public Function PasswordHash(ByVal motDePasse As String, ByVal sel As String, ByVal iterations As Long) As String
    Dim h As String, i As Long
    h = SHA256Hex(sel & ":" & motDePasse)
    For i = 2 To iterations
        h = SHA256Hex(h & sel)
    Next i
    PasswordHash = h
End Function
