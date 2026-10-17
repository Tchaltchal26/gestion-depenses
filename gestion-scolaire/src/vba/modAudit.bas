Attribute VB_Name = "modAudit"
'===============================================================================
' SG - Module modAudit (tracabilite integrale)
' Role : journal d'audit en ajout seul, chaine par empreintes SHA-256.
'        Hash(n) = SHA256(Hash(n-1) & separateur & contenu canonique(n))
' Ce que cela garantit : toute modification, insertion ou suppression d'une
'        ligne au milieu du journal est DETECTEE par Audit_Verifier.
' Ce que cela NE garantit PAS : une personne ayant acces au fichier et au code
'        peut recalculer toute la chaine. D'ou l'"ancre" : l'empreinte de tete
'        est exportee regulierement hors du classeur (sauvegardes, impression,
'        envoi a la direction) et comparee lors des verifications.
'        Un journal VBA n'est pas inviolable (voir docs/01-architecture.md).
'===============================================================================
Option Explicit

Private Const HASH_ORIGINE As String = "0000000000000000000000000000000000000000000000000000000000000000"
Private mTxRef As String
Private mTxProfondeur As Long

' Une "transaction" regroupe les evenements d'une meme operation metier sous
' une reference commune (TxRef). Les operations imbriquees reutilisent la
' reference de l'operation englobante.
' Limite : il ne s'agit pas d'une transaction ACID (pas d'annulation
' automatique) ; les controles sont faits AVANT toute ecriture.
Public Function Audit_DebutTx(ByVal libelle As String) As String
    If mTxProfondeur = 0 Then mTxRef = "TX-" & CompactStamp() & "-" & RandomHex(6)
    mTxProfondeur = mTxProfondeur + 1
    Audit_DebutTx = mTxRef
End Function

Public Sub Audit_FinTx()
    If mTxProfondeur > 0 Then mTxProfondeur = mTxProfondeur - 1
    If mTxProfondeur = 0 Then mTxRef = ""
End Sub

Public Sub Audit_ReinitialiserTx()
    mTxProfondeur = 0
    mTxRef = ""
End Sub

Private Function Canonique(rec As Variant) As String
    Dim champs As Variant, i As Long, s As String
    champs = Split("EventID;Horodatage;UserID;EtabID;Module;Operation;TableCible;EnregistrementID;Champ;AncienneValeur;NouvelleValeur;Motif;Resultat;TxRef", ";")
    s = ""
    For i = 0 To UBound(champs)
        If i > 0 Then s = s & Chr$(31)
        s = s & SafeField(RecGet(rec, CStr(champs(i)), ""))
    Next i
    Canonique = s
End Function

Public Function Audit_HashTete() As String
    Dim ws As Object, n As Long
    Set ws = DAL_Ws("T_Audit")
    n = DAL_LastRow(ws)
    If n < 2 Then
        Audit_HashTete = HASH_ORIGINE
    Else
        Audit_HashTete = CStr(ws.Cells(n, DAL_Col(ws, "Hash")).Value)
    End If
End Function

Public Function Audit_NbEvenements() As Long
    Audit_NbEvenements = DAL_LastRow(DAL_Ws("T_Audit")) - 1
End Function

' Enregistre un evenement. Appele par la couche metier APRES verification
' des droits et AVANT/APRES l'ecriture selon le cas (voir chaque module).
Public Function Audit_Log(ByVal etabId As String, ByVal module As String, ByVal operation As String, _
                          ByVal tableCible As String, ByVal enregId As String, ByVal champ As String, _
                          ByVal ancienne As Variant, ByVal nouvelle As Variant, ByVal motif As String, _
                          Optional ByVal resultat As Variant) As String
    ' Valeurs par defaut (IsMissing : forme compatible Excel et LibreOffice).
    If IsMissing(resultat) Then resultat = "OK"
    Dim rec As Variant, prec As String, id As String
    id = DAL_NewId("T_Audit")
    rec = RecNew()
    RecSet rec, "EventID", id
    RecSet rec, "Horodatage", NowStamp()
    RecSet rec, "UserID", Sec_UtilisateurCourant()
    RecSet rec, "EtabID", etabId
    RecSet rec, "Module", module
    RecSet rec, "Operation", operation
    RecSet rec, "TableCible", tableCible
    RecSet rec, "EnregistrementID", enregId
    RecSet rec, "Champ", champ
    RecSet rec, "AncienneValeur", SafeField(ancienne)
    RecSet rec, "NouvelleValeur", SafeField(nouvelle)
    RecSet rec, "Motif", motif
    RecSet rec, "Resultat", resultat
    RecSet rec, "TxRef", mTxRef
    prec = Audit_HashTete()
    RecSet rec, "HashPrec", prec
    RecSet rec, "Hash", SHA256Hex(prec & Chr$(30) & Canonique(rec))
    DAL_Insert "T_Audit", rec, False
    Audit_Log = id
End Function

' Journalise la creation d'un enregistrement complet en un seul evenement :
' NouvelleValeur contient "champ=valeur" pour chaque champ renseigne.
Public Sub Audit_LogCreation(ByVal etabId As String, ByVal module As String, ByVal table As String, rec As Variant, ByVal motif As String)
    Dim i As Long, id As String, s As String
    id = RecStr(rec, Split(SchemaColonnes(table), ";")(0))
    s = ""
    For i = 1 To RecCount(rec) - 1
        If Len(NzStr(RecValAt(rec, i))) > 0 Then
            If Len(s) > 0 Then s = s & " ; "
            s = s & RecKey(rec, i) & "=" & NzStr(RecValAt(rec, i))
        End If
    Next i
    Audit_Log etabId, module, "CREATION", table, id, "", "", s, motif
End Sub

' Verifie toute la chaine du classeur courant.
' Retourne "OK|n" ou "ALTERE|EventID|raison".
Public Function Audit_Verifier(Optional ByVal ancreAttendue As Variant) As String
    ' Valeurs par defaut (IsMissing : forme compatible Excel et LibreOffice).
    If IsMissing(ancreAttendue) Then ancreAttendue = ""
    Audit_Verifier = Audit_VerifierFeuille(DAL_Ws("T_Audit"), ancreAttendue)
End Function

' Verification sur une feuille quelconque (utilisee aussi pour controler une
' copie de sauvegarde avant restauration).
Public Function Audit_VerifierFeuille(ws As Object, Optional ByVal ancreAttendue As Variant) As String
    ' Valeurs par defaut (IsMissing : forme compatible Excel et LibreOffice).
    If IsMissing(ancreAttendue) Then ancreAttendue = ""
    Dim data As Variant, n As Long, nc As Long, i As Long, c As Long
    Dim rec As Variant, prec As String, calc As String
    Dim colHash As Long, colPrec As Long, trouve As Boolean
    n = DAL_LastRow(ws)
    nc = UBound(Split(SchemaColonnes("T_Audit"), ";")) + 1
    If n < 2 Then
        If Len(ancreAttendue) > 0 And ancreAttendue <> HASH_ORIGINE Then
            Audit_VerifierFeuille = "ALTERE||journal vide alors qu'une ancre est attendue"
        Else
            Audit_VerifierFeuille = "OK|0"
        End If
        Exit Function
    End If
    data = ws.Range(ws.Cells(1, 1), ws.Cells(n, nc)).Value
    For c = 1 To nc
        If CStr(data(1, c)) = "Hash" Then colHash = c
        If CStr(data(1, c)) = "HashPrec" Then colPrec = c
    Next c
    If colHash = 0 Or colPrec = 0 Then
        Audit_VerifierFeuille = "ALTERE||structure du journal modifiee"
        Exit Function
    End If
    prec = HASH_ORIGINE
    trouve = (Len(ancreAttendue) = 0)
    For i = 2 To n
        rec = RecNew()
        For c = 1 To nc
            RecSet rec, CStr(data(1, c)), NzStr(data(i, c))
        Next c
        If CStr(data(i, colPrec)) <> prec Then
            Audit_VerifierFeuille = "ALTERE|" & data(i, 1) & "|chainage rompu (ligne inseree ou supprimee)"
            Exit Function
        End If
        calc = SHA256Hex(prec & Chr$(30) & Canonique(rec))
        If calc <> CStr(data(i, colHash)) Then
            Audit_VerifierFeuille = "ALTERE|" & data(i, 1) & "|contenu modifie"
            Exit Function
        End If
        If calc = ancreAttendue Then trouve = True
        prec = calc
    Next i
    If Not trouve Then
        ' L'ancre exportee doit figurer dans la chaine (sinon : chaine recalculee).
        Audit_VerifierFeuille = "ALTERE||ancre externe introuvable : journal reconstruit ou tronque"
        Exit Function
    End If
    Audit_VerifierFeuille = "OK|" & (n - 1)
End Function

' Historique d'un enregistrement (EventID), dans l'ordre chronologique.
' La consultation est elle-meme soumise aux droits du module AUDIT.
Public Function Audit_Historique(ByVal etabId As String, ByVal enregId As String) As Collection
    Sec_Exiger etabId, "AUDIT", NIV_LECTURE
    Set Audit_Historique = DAL_Query("T_Audit", RecFrom("EtabID=" & etabId & "|EnregistrementID=" & enregId))
End Function
