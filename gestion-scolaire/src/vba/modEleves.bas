Attribute VB_Name = "modEleves"
'===============================================================================
' SG - Module modEleves (dossiers eleves/etudiants et responsables)
' Regles :
'   - dossier permanent : identifiant technique (EleveID) et matricule jamais
'     modifies ni reutilises ; aucune suppression (le depart change le statut) ;
'   - chaque modification est journalisee champ par champ (ancienne/nouvelle
'     valeur, auteur, date, motif) ;
'   - detection des doublons (nom + prenoms normalises + date de naissance) ;
'   - toute lecture est cloisonnee par etablissement.
'===============================================================================
Option Explicit

Public Const STATUTS_ELEVE As String = "NOUVEAU;ACTIF;SUSPENDU;PARTI;TRANSFERE;RADIE;DIPLOME"
Private Const CHAMPS_MODIFIABLES As String = "Nom;Prenoms;Sexe;DateNaissance;LieuNaissance;Nationalite;Adresse;Telephone;EtabPrecedent"

Public Function Eleve_CleDoublon(ByVal nom As String, ByVal prenoms As String, ByVal dateNaissance As String) As String
    Eleve_CleDoublon = NormalizeText(nom) & "#" & NormalizeText(prenoms) & "#" & dateNaissance
End Function

Public Function Eleve_Doublons(ByVal etabId As String, ByVal nom As String, ByVal prenoms As String, ByVal dateNaissance As String) As Collection
    Dim f As Variant
    Sec_Exiger etabId, "ELEVES", NIV_LECTURE
    f = RecFrom("EtabID=" & etabId)
    RecSet f, "CleDoublon", Eleve_CleDoublon(nom, prenoms, dateNaissance)
    Set Eleve_Doublons = DAL_Query("T_Eleves", f)
End Function

Private Sub ControlerIdentite(ByVal etabId As String, ByVal nom As String, ByVal prenoms As String, ByVal sexe As String, ByVal dateNaissance As String)
    Dim ageMin As Long
    If Len(nom) = 0 Then SG_Raise SG_ERR_VALIDATION, "Le nom est obligatoire."
    If InStr(nom & prenoms, "|") > 0 Or InStr(nom & prenoms, "#") > 0 Then SG_Raise SG_ERR_VALIDATION, "Caracteres | et # interdits dans les noms."
    If Len(prenoms) = 0 Then SG_Raise SG_ERR_VALIDATION, "Le(s) prenom(s) sont obligatoires."
    If Not InList(sexe, "M;F") Then SG_Raise SG_ERR_VALIDATION, "Sexe : valeur attendue M ou F."
    If Not IsISODate(dateNaissance) Then SG_Raise SG_ERR_VALIDATION, "Date de naissance invalide (AAAA-MM-JJ)."
    If dateNaissance >= TodayISO() Then SG_Raise SG_ERR_VALIDATION, "La date de naissance doit etre dans le passe."
    ageMin = CLng(Param_Get(etabId, "AGE_MIN_INSCRIPTION", "2"))
    If ISODiffDays(dateNaissance, TodayISO()) < ageMin * 365 Then SG_Raise SG_ERR_VALIDATION, "Age inferieur au minimum parametre (" & ageMin & " ans)."
End Sub

' Cree un dossier. rec : Nom, Prenoms, Sexe, DateNaissance, LieuNaissance,
' Nationalite, Adresse, Telephone, EtabPrecedent.
' Si un doublon probable existe, l'operation est refusee sauf confirmation
' explicite (forcer=True) accompagnee d'un motif.
Public Function Eleve_Creer(ByVal etabId As String, rec As Variant, Optional ByVal forcer As Boolean = False, _
                            Optional ByVal motifDoublon As Variant) As String
    ' Valeurs par defaut (IsMissing : forme compatible Excel et LibreOffice).
    If IsMissing(motifDoublon) Then motifDoublon = ""
    Dim nom As String, prenoms As String, dn As String, sexe As String, cle As String, doublons As Collection
    Dim id As String, r As Variant, mat As String, maxE As Long, champs As Variant, i As Long
    Sec_Exiger etabId, "ELEVES", NIV_ECRITURE
    nom = UCase$(RecStr(rec, "Nom"))
    prenoms = RecStr(rec, "Prenoms")
    sexe = UCase$(RecStr(rec, "Sexe"))
    dn = RecStr(rec, "DateNaissance")
    ControlerIdentite etabId, nom, prenoms, sexe, dn
    maxE = Lic_MaxEleves()
    If maxE > 0 Then
        If DAL_Count("T_Eleves", RecNew()) >= maxE Then SG_Raise SG_ERR_LICENCE, "Limite de la version d'essai atteinte (" & maxE & " eleves)."
    End If
    cle = Eleve_CleDoublon(nom, prenoms, dn)
    Dim f As Variant
    f = RecFrom("EtabID=" & etabId)
    RecSet f, "CleDoublon", cle
    Set doublons = DAL_Query("T_Eleves", f)
    If doublons.Count > 0 Then
        If Not forcer Then SG_Raise SG_ERR_DOUBLON, "Doublon probable : dossier " & DAL_Value("T_Eleves", doublons(1), "Matricule") & " (meme nom, prenoms et date de naissance)."
        If Len(Trim$(motifDoublon)) = 0 Then SG_Raise SG_ERR_VALIDATION, "Motif obligatoire pour confirmer un homonyme."
    End If

    id = DAL_NewId("T_Eleves")
    mat = Format_Numero(Param_Get(etabId, "FORMAT_MATRICULE", "{ETAB}-{AAAA}-{N5}"), Etab_Code(etabId), Format(Date, "yyyy"), Compteur_Suivant(etabId, "MATRICULE"))
    If Len(DAL_First("T_Eleves", RecFrom("EtabID=" & etabId & "|Matricule=" & mat))) > 0 Then SG_Raise SG_ERR_DOUBLON, "Matricule deja attribue : verifier le format de numerotation."
    r = RecNew()
    RecSet r, "EleveID", id
    RecSet r, "EtabID", etabId
    RecSet r, "Matricule", mat
    RecSet r, "Nom", nom
    RecSet r, "Prenoms", prenoms
    RecSet r, "Sexe", sexe
    RecSet r, "DateNaissance", dn
    champs = Split("LieuNaissance;Nationalite;Adresse;Telephone;EtabPrecedent;EleveOrigineID", ";")
    For i = 0 To UBound(champs)
        RecSet r, CStr(champs(i)), RecStr(rec, CStr(champs(i)))
    Next i
    RecSet r, "Statut", "NOUVEAU"
    RecSet r, "DateStatut", TodayISO()
    RecSet r, "CleDoublon", cle
    RecSet r, "DateCreation", NowStamp()
    RecSet r, "CreePar", Sec_UtilisateurCourant()
    RecSet r, "ModifieLe", NowStamp()
    RecSet r, "Version", "1"
    DAL_Insert "T_Eleves", r
    Audit_LogCreation etabId, "ELEVES", "T_Eleves", r, IIf(doublons.Count > 0, "Homonyme confirme : " & motifDoublon, "")
    Eleve_Creer = id
End Function

Public Function Eleve_Get(ByVal etabId As String, ByVal eleveId As String) As Variant
    Sec_Exiger etabId, "ELEVES", NIV_LECTURE
    Etab_ExigerAppartenance "T_Eleves", eleveId, etabId
    Eleve_Get = DAL_Get("T_Eleves", eleveId)
End Function

Public Sub Eleve_Modifier(ByVal etabId As String, ByVal eleveId As String, ByVal champ As String, _
                          ByVal valeur As String, ByVal motif As String)
    Dim e As Variant, ancien As String, nom As String, prenoms As String, sexe As String, dn As String
    Sec_Exiger etabId, "ELEVES", NIV_ECRITURE
    Etab_ExigerAppartenance "T_Eleves", eleveId, etabId
    If Not InList(champ, CHAMPS_MODIFIABLES) Then SG_Raise SG_ERR_VALIDATION, "Le champ " & champ & " n'est pas modifiable directement."
    If InList(champ, Param_Get(etabId, "MOTIF_OBLIGATOIRE_CHAMPS", "Nom;Prenoms;DateNaissance;Sexe")) And Len(Trim$(motif)) = 0 Then
        SG_Raise SG_ERR_VALIDATION, "Motif obligatoire pour modifier le champ " & champ & "."
    End If
    e = DAL_Get("T_Eleves", eleveId)
    If champ = "Nom" Or champ = "Sexe" Then valeur = UCase$(Trim$(valeur)) Else valeur = Trim$(valeur)
    ancien = RecStr(e, champ)
    If ancien = valeur Then Exit Sub
    nom = RecStr(e, "Nom"): prenoms = RecStr(e, "Prenoms"): sexe = RecStr(e, "Sexe"): dn = RecStr(e, "DateNaissance")
    Select Case champ
        Case "Nom": nom = valeur
        Case "Prenoms": prenoms = valeur
        Case "Sexe": sexe = valeur
        Case "DateNaissance": dn = valeur
    End Select
    ControlerIdentite etabId, nom, prenoms, sexe, dn
    DAL_Update "T_Eleves", eleveId, champ, valeur
    DAL_Update "T_Eleves", eleveId, "CleDoublon", Eleve_CleDoublon(nom, prenoms, dn)
    DAL_Update "T_Eleves", eleveId, "ModifieLe", NowStamp()
    DAL_Update "T_Eleves", eleveId, "Version", CStr(CLng(RecStr(e, "Version")) + 1)
    Audit_Log etabId, "ELEVES", "MODIFICATION", "T_Eleves", eleveId, champ, ancien, valeur, motif
End Sub

' Changement de statut : reserve aux operations d'inscription (depart,
' transfert, radiation, suspension, reintegration, fin de cycle).
Public Sub Eleve_ChangerStatut(ByVal etabId As String, ByVal eleveId As String, ByVal statut As String, ByVal motif As String)
    Dim ancien As String
    Sec_Exiger etabId, "INSCR", NIV_ECRITURE
    Etab_ExigerAppartenance "T_Eleves", eleveId, etabId
    If Not InList(statut, STATUTS_ELEVE) Then SG_Raise SG_ERR_VALIDATION, "Statut inconnu : " & statut
    ancien = DAL_Value("T_Eleves", eleveId, "Statut")
    If ancien = statut Then Exit Sub
    DAL_Update "T_Eleves", eleveId, "Statut", statut
    DAL_Update "T_Eleves", eleveId, "DateStatut", TodayISO()
    DAL_Update "T_Eleves", eleveId, "ModifieLe", NowStamp()
    Audit_Log etabId, "ELEVES", "CHANGEMENT_STATUT", "T_Eleves", eleveId, "Statut", ancien, statut, motif
End Sub

' Recherche multicritere simple : texte libre (nom, prenoms, matricule),
' insensible aux accents et a la casse, + statut optionnel.
Public Function Eleve_Rechercher(ByVal etabId As String, ByVal texte As String, Optional ByVal statut As Variant) As Collection
    ' Valeurs par defaut (IsMissing : forme compatible Excel et LibreOffice).
    If IsMissing(statut) Then statut = ""
    Dim recs As Collection, res As New Collection, i As Long, e As Variant, cible As String, mots As Variant, k As Long, ok As Boolean
    Dim f As Variant
    Sec_Exiger etabId, "ELEVES", NIV_LECTURE
    f = RecFrom("EtabID=" & etabId)
    If Len(statut) > 0 Then RecSet f, "Statut", statut
    Set recs = DAL_QueryRecs("T_Eleves", f)
    mots = Split(NormalizeText(texte), " ")
    For i = 1 To recs.Count
        e = recs(i)
        cible = NormalizeText(RecStr(e, "Nom") & " " & RecStr(e, "Prenoms") & " " & RecStr(e, "Matricule"))
        ok = True
        For k = LBound(mots) To UBound(mots)
            If Len(mots(k)) > 0 And InStr(1, cible, mots(k), vbBinaryCompare) = 0 Then ok = False: Exit For
        Next k
        If ok Then res.Add RecStr(e, "EleveID")
    Next i
    Set Eleve_Rechercher = res
End Function

'-------------------------------------------------------------------------------
' Responsables (parents, tuteurs, responsables financiers, contacts d'urgence)
'-------------------------------------------------------------------------------
Public Function Resp_Creer(ByVal etabId As String, rec As Variant) As String
    Dim id As String, r As Variant, champs As Variant, i As Long
    Sec_Exiger etabId, "ELEVES", NIV_ECRITURE
    If Len(RecStr(rec, "Nom")) = 0 Then SG_Raise SG_ERR_VALIDATION, "Le nom du responsable est obligatoire."
    If Len(RecStr(rec, "Telephone")) = 0 And Len(RecStr(rec, "Email")) = 0 Then SG_Raise SG_ERR_VALIDATION, "Un telephone ou un courriel est obligatoire."
    id = DAL_NewId("T_Responsables")
    r = RecNew()
    RecSet r, "RespID", id
    RecSet r, "EtabID", etabId
    champs = Split("Nom;Prenoms;Telephone;Email;Adresse;Profession", ";")
    For i = 0 To UBound(champs)
        RecSet r, CStr(champs(i)), RecStr(rec, CStr(champs(i)))
    Next i
    RecSet r, "Nom", UCase$(RecStr(rec, "Nom"))
    RecSet r, "Actif", "1"
    DAL_Insert "T_Responsables", r
    Audit_LogCreation etabId, "ELEVES", "T_Responsables", r, ""
    Resp_Creer = id
End Function

' roles : combinaison de LEGAL;FINANCIER;URGENCE;RECUPERATION
Public Function Resp_Lier(ByVal etabId As String, ByVal eleveId As String, ByVal respId As String, _
                          ByVal lien As String, ByVal roles As String) As String
    Dim id As String, r As Variant
    Sec_Exiger etabId, "ELEVES", NIV_ECRITURE
    Etab_ExigerAppartenance "T_Eleves", eleveId, etabId
    Etab_ExigerAppartenance "T_Responsables", respId, etabId
    If Len(Trim$(lien)) = 0 Then SG_Raise SG_ERR_VALIDATION, "Preciser le lien (pere, mere, tuteur...)."
    If Len(DAL_First("T_LiensResponsables", RecFrom("EtabID=" & etabId & "|EleveID=" & eleveId & "|RespID=" & respId & "|Actif=1"))) > 0 Then SG_Raise SG_ERR_DOUBLON, "Ce responsable est deja lie a cet eleve."
    id = DAL_NewId("T_LiensResponsables")
    r = RecNew()
    RecSet r, "LienID", id
    RecSet r, "EtabID", etabId
    RecSet r, "EleveID", eleveId
    RecSet r, "RespID", respId
    RecSet r, "Lien", lien
    RecSet r, "Legal", IIf(InList("LEGAL", roles), "1", "0")
    RecSet r, "Financier", IIf(InList("FINANCIER", roles), "1", "0")
    RecSet r, "Urgence", IIf(InList("URGENCE", roles), "1", "0")
    RecSet r, "AutoriseRecuperer", IIf(InList("RECUPERATION", roles), "1", "0")
    RecSet r, "Actif", "1"
    DAL_Insert "T_LiensResponsables", r
    Audit_LogCreation etabId, "ELEVES", "T_LiensResponsables", r, ""
    Resp_Lier = id
End Function

' Enfants (EleveID) rattaches a un responsable, dans l'etablissement donne.
Public Function Resp_Enfants(ByVal etabId As String, ByVal respId As String) As Collection
    Dim liens As Collection, res As New Collection, i As Long
    Sec_Exiger etabId, "ELEVES", NIV_LECTURE
    Etab_ExigerAppartenance "T_Responsables", respId, etabId
    Set liens = DAL_Query("T_LiensResponsables", RecFrom("EtabID=" & etabId & "|RespID=" & respId & "|Actif=1"))
    For i = 1 To liens.Count
        res.Add DAL_Value("T_LiensResponsables", liens(i), "EleveID")
    Next i
    Set Resp_Enfants = res
End Function
