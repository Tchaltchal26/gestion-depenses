Attribute VB_Name = "modRSA"
'===============================================================================
' SG - Module modRSA
' Role : VERIFICATION de signatures RSA PKCS#1 v1.5 / SHA-256 en VBA pur.
'        Ce module ne contient et ne manipule AUCUNE cle privee : la signature
'        des licences se fait exclusivement cote editeur (outil sg_licence.py).
' Arithmetique : grands entiers en base 65536 (tableaux de Double, poids
'        faible en premier) ; reduction modulaire par l'algorithme D de Knuth.
'===============================================================================
Option Explicit

Private Const B As Double = 65536#

' Prefixe DigestInfo ASN.1 pour SHA-256 (RFC 8017, section 9.2).
Private Const DIGESTINFO_SHA256 As String = "3031300d060960864801650304020105000420"

'-------------------------------------------------------------------------------
' Conversions
'-------------------------------------------------------------------------------
' Octets gros-boutistes -> limbs (poids faible en premier).
Private Function BytesToLimbs(bytes As Variant) As Variant
    Dim n As Long, nl As Long, i As Long, r() As Double, pos As Long
    n = ByteCount(bytes)
    nl = (n + 1) \ 2
    If nl = 0 Then nl = 1
    ReDim r(0 To nl - 1)
    For i = 0 To n - 1
        pos = n - 1 - i ' indice de l'octet en partant du poids faible
        If (i Mod 2) = 0 Then
            r(i \ 2) = r(i \ 2) + bytes(LBound(bytes) + pos)
        Else
            r(i \ 2) = r(i \ 2) + CDbl(bytes(LBound(bytes) + pos)) * 256#
        End If
    Next i
    BytesToLimbs = r
End Function

Private Function LimbsToBytes(x As Variant, ByVal nBytes As Long) As Variant
    Dim r() As Byte, i As Long, limb As Double, v As Double
    ReDim r(0 To nBytes - 1)
    For i = 0 To nBytes - 1
        limb = 0
        If (i \ 2) <= UBound(x) Then limb = x(i \ 2)
        If (i Mod 2) = 0 Then v = limb - Int(limb / 256#) * 256# Else v = Int(limb / 256#)
        r(nBytes - 1 - i) = CByte(v)
    Next i
    LimbsToBytes = r
End Function

Private Function Trimmed(x As Variant) As Long
    ' Nombre de limbs significatifs (au moins 1).
    Dim i As Long
    For i = UBound(x) To 0 Step -1
        If x(i) <> 0 Then Trimmed = i + 1: Exit Function
    Next i
    Trimmed = 1
End Function

'-------------------------------------------------------------------------------
' Multiplication et reduction
'-------------------------------------------------------------------------------
Private Function MulBig(a As Variant, bb As Variant) As Variant
    Dim na As Long, nb As Long, r() As Double, i As Long, j As Long
    Dim carry As Double, t As Double, ai As Double
    na = Trimmed(a): nb = Trimmed(bb)
    ReDim r(0 To na + nb)
    For i = 0 To na - 1
        ai = a(i)
        If ai <> 0 Then
            carry = 0
            For j = 0 To nb - 1
                t = r(i + j) + ai * bb(j) + carry
                carry = Int(t / B)
                r(i + j) = t - carry * B
            Next j
            j = i + nb
            Do While carry <> 0
                t = r(j) + carry
                carry = Int(t / B)
                r(j) = t - carry * B
                j = j + 1
            Loop
        End If
    Next i
    MulBig = r
End Function

' Reste de u modulo v (Knuth, TAOCP vol. 2, 4.3.1, algorithme D).
Private Function ModBig(uIn As Variant, vIn As Variant) As Variant
    Dim n As Long, m As Long, s As Long, i As Long, j As Long
    Dim v() As Double, u() As Double, r() As Double
    Dim qhat As Double, rhat As Double, p As Double, k As Double, t As Double, borrow As Double
    Dim mult As Double, carry As Double, nu As Long

    n = Trimmed(vIn)
    nu = Trimmed(uIn)
    If nu < n Then
        ReDim r(0 To n - 1)
        For i = 0 To nu - 1: r(i) = uIn(i): Next i
        ModBig = r
        Exit Function
    End If
    If n = 1 Then
        ' Diviseur d'un seul limb : reduction directe.
        ReDim r(0 To 0)
        t = 0
        For i = nu - 1 To 0 Step -1
            t = t * B + uIn(i)
            t = t - Int(t / vIn(0)) * vIn(0)
        Next i
        r(0) = t
        ModBig = r
        Exit Function
    End If

    ' Normalisation : decaler pour que le bit de poids fort de v soit a 1.
    s = 0
    t = vIn(n - 1)
    Do While t < 32768#
        t = t * 2: s = s + 1
    Loop
    mult = 2 ^ s
    ReDim v(0 To n - 1)
    carry = 0
    For i = 0 To n - 1
        t = vIn(i) * mult + carry
        carry = Int(t / B)
        v(i) = t - carry * B
    Next i
    m = nu - n
    ReDim u(0 To nu)
    carry = 0
    For i = 0 To nu - 1
        t = uIn(i) * mult + carry
        carry = Int(t / B)
        u(i) = t - carry * B
    Next i
    u(nu) = carry

    For j = m To 0 Step -1
        t = u(j + n) * B + u(j + n - 1)
        qhat = Int(t / v(n - 1))
        rhat = t - qhat * v(n - 1)
        Do While qhat >= B Or qhat * v(n - 2) > B * rhat + u(j + n - 2)
            qhat = qhat - 1
            rhat = rhat + v(n - 1)
            If rhat >= B Then Exit Do
        Loop
        ' Multiplier et soustraire.
        k = 0: borrow = 0
        For i = 0 To n - 1
            p = qhat * v(i) + k
            k = Int(p / B)
            t = u(i + j) - (p - k * B) - borrow
            If t < 0 Then
                u(i + j) = t + B: borrow = 1
            Else
                u(i + j) = t: borrow = 0
            End If
        Next i
        t = u(j + n) - k - borrow
        If t < 0 Then
            ' Estimation trop grande d'une unite : rajouter v.
            u(j + n) = t + B
            carry = 0
            For i = 0 To n - 1
                t = u(i + j) + v(i) + carry
                carry = Int(t / B)
                u(i + j) = t - carry * B
            Next i
            u(j + n) = u(j + n) + carry - B
        Else
            u(j + n) = t
        End If
    Next j

    ' Denormalisation du reste.
    ReDim r(0 To n - 1)
    carry = 0
    For i = n - 1 To 0 Step -1
        t = u(i) + carry * B
        r(i) = Int(t / mult)
        carry = t - r(i) * mult
    Next i
    ModBig = r
End Function

' base^65537 mod n (exposant public standard).
Private Function PowMod65537(baseVal As Variant, modulus As Variant) As Variant
    Dim x As Variant, i As Long, sq As Variant
    x = ModBig(baseVal, modulus)
    Dim orig As Variant
    orig = x
    For i = 1 To 16
        sq = MulBig(x, x)
        x = ModBig(sq, modulus)
    Next i
    sq = MulBig(x, orig)
    PowMod65537 = ModBig(sq, modulus)
End Function

'-------------------------------------------------------------------------------
' API publique
'-------------------------------------------------------------------------------
' Verifie une signature RSASSA-PKCS1-v1_5 SHA-256 (exposant 65537).
'   message      : octets signes
'   signature    : octets de la signature (taille = taille du module)
'   modulusHex   : module public en hexadecimal
Public Function RSA_VerifySHA256(message As Variant, signature As Variant, ByVal modulusHex As String) As Boolean
    Dim modBytes As Variant, nMod As Variant, sig As Variant, em As Variant
    Dim kLen As Long, emBytes As Variant, expected As String, h As Variant, i As Long, tLen As Long
    Dim hexEm As String

    RSA_VerifySHA256 = False
    modBytes = HexToBytes(modulusHex)
    ' Retirer les zeros de tete eventuels.
    Dim start As Long
    start = 0
    Do While start < ByteCount(modBytes) - 1
        If modBytes(start) <> 0 Then Exit Do
        start = start + 1
    Loop
    kLen = ByteCount(modBytes) - start
    If kLen < 128 Then Exit Function ' refuse les cles < 1024 bits
    If ByteCount(signature) <> kLen Then Exit Function

    nMod = BytesToLimbs(modBytes)
    sig = BytesToLimbs(signature)
    ' La signature doit etre strictement inferieure au module.
    If CompareBig(sig, nMod) >= 0 Then Exit Function

    em = PowMod65537(sig, nMod)
    emBytes = LimbsToBytes(em, kLen)

    h = SHA256Bytes(message)
    tLen = Len(DIGESTINFO_SHA256) \ 2 + 32
    If kLen < tLen + 11 Then Exit Function
    expected = "0001" & String$((kLen - tLen - 3) * 2, "f") & "00" & DIGESTINFO_SHA256 & BytesToHex(h)
    hexEm = BytesToHex(emBytes)
    RSA_VerifySHA256 = (hexEm = expected)
End Function

Private Function CompareBig(a As Variant, bb As Variant) As Long
    Dim na As Long, nb As Long, i As Long
    na = Trimmed(a): nb = Trimmed(bb)
    If na <> nb Then CompareBig = IIf(na > nb, 1, -1): Exit Function
    For i = na - 1 To 0 Step -1
        If a(i) <> bb(i) Then CompareBig = IIf(a(i) > bb(i), 1, -1): Exit Function
    Next i
    CompareBig = 0
End Function

' Expose pour les tests : reste de la division de deux entiers hexadecimaux.
Public Function RSA_TestModHex(ByVal aHex As String, ByVal mHex As String) As String
    Dim ab As Variant, mb As Variant, a As Variant, m As Variant, r As Variant, rb As Variant
    ab = HexToBytes(aHex): mb = HexToBytes(mHex)
    a = BytesToLimbs(ab): m = BytesToLimbs(mb)
    r = ModBig(a, m)
    rb = LimbsToBytes(r, (UBound(r) + 1) * 2)
    RSA_TestModHex = BytesToHex(rb)
End Function
