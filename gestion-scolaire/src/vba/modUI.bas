Attribute VB_Name = "modUI"
'===============================================================================
' SG - Module modUI (interface Excel)
' Ecrans = feuilles de saisie mises en forme + boutons ; les cellules de
' saisie sont reperees par des NOMS (F_Nom, F_Prenoms...), jamais par adresse
' dans le code metier. Toutes les regles sont dans les modules metier : ce
' module ne fait que lire les saisies, appeler le metier et afficher.
' Specifique Excel de bureau (Windows/Mac) : formes, boites de dialogue,
' UserForm, export PDF. Non execute par les tests LibreOffice.
'===============================================================================
Option Explicit

Private mEtab As String
Private Const MDP_FEUILLES As String = "SG-structure"   ' protection d'interface uniquement (non securitaire)

Private Const C_FOND As Long = 16316664       ' RGB(248,247,248) approx.
Private Const C_PRIMAIRE As Long = 7884319     ' RGB(31,78,120)
Private Const C_SAISIE As Long = 16777215
Private Const C_ALERTE As Long = 13421823      ' RGB(255,204,204)

'-------------------------------------------------------------------------------
' Ouverture / fermeture
'-------------------------------------------------------------------------------
Public Sub SG_AuOuverture()
    On Error GoTo erreur
    If Not DAL_TableExiste("T_Systeme") Then
        MsgBox "Classeur non initialise. Lancer l'installation (SG_Installer).", vbExclamation, "SG"
        Exit Sub
    End If
    ' Classeur de tests (sans interface) : ne rien faire a l'ouverture.
    If Not DAL_TableExiste("ACCUEIL") Then Exit Sub
    ProtegerDonnees
    Lic_Recharger
    ThisWorkbook.Worksheets("ACCUEIL").Activate
    AfficherBandeau
    SG_Connexion
    Exit Sub
erreur:
    UI_Erreur
End Sub

Public Sub SG_AFermeture()
    On Error Resume Next
    Sec_Deconnexion
End Sub

Private Sub ProtegerDonnees()
    Dim t As Variant, ws As Object
    For Each t In SchemaTables()
        Set ws = ThisWorkbook.Worksheets(CStr(t))
        ws.Visible = 2 ' xlSheetVeryHidden
        ' UserInterfaceOnly : le code VBA peut ecrire, l'utilisateur non.
        ws.Protect Password:=MDP_FEUILLES, UserInterfaceOnly:=True
    Next t
End Sub

Private Sub UI_Erreur()
    Dim titre As String
    Audit_ReinitialiserTx
    Select Case Err.Number
        Case SG_ERR_VALIDATION: titre = "Saisie incorrecte"
        Case SG_ERR_ACCES: titre = "Acces refuse"
        Case SG_ERR_ISOLATION: titre = "Acces refuse (autre etablissement)"
        Case SG_ERR_DOUBLON: titre = "Doublon"
        Case SG_ERR_ETAT: titre = "Operation impossible"
        Case SG_ERR_LICENCE, SG_ERR_LECTURE_SEULE: titre = "Licence"
        Case Else: titre = "Erreur technique"
    End Select
    MsgBox Err.Description, IIf(titre = "Erreur technique", vbCritical, vbExclamation), "SG - " & titre
End Sub

'-------------------------------------------------------------------------------
' Connexion et contexte
'-------------------------------------------------------------------------------
Public Sub SG_Connexion()
    Dim login As String, mdp As String
    On Error GoTo erreur
    If FormulaireExiste("frmConnexion") Then
        VBA.UserForms.Add("frmConnexion").Show
    Else
        login = InputBox("Identifiant :", "SG - Connexion")
        If Len(login) = 0 Then Exit Sub
        mdp = InputBox("Mot de passe (formulaire masque non installe) :", "SG - Connexion")
        Sec_Connexion login, mdp
        ApresConnexion
    End If
    Exit Sub
erreur:
    UI_Erreur
End Sub

' Appele par frmConnexion apres une connexion reussie.
Public Sub ApresConnexion()
    Dim etabs As Collection
    Set etabs = Etab_Accessibles()
    If etabs.Count >= 1 Then mEtab = etabs(1) Else mEtab = ""
    MettreAJourListeEtabs etabs
    AfficherBandeau
    If DAL_Value("T_Utilisateurs", Sec_UtilisateurCourant(), "DoitChangerMdp") = "1" Then
        MsgBox "Votre mot de passe est temporaire : merci de le changer (menu Mon compte).", vbInformation, "SG"
    End If
End Sub

Public Sub SG_Deconnexion()
    Sec_Deconnexion
    mEtab = ""
    AfficherBandeau
End Sub

Private Function FormulaireExiste(ByVal nom As String) As Boolean
    Dim c As Object
    On Error GoTo non
    Set c = ThisWorkbook.VBProject.VBComponents(nom)
    FormulaireExiste = True
    Exit Function
non:
    ' Acces au projet VBA refuse : on suppose que le formulaire est installe.
    FormulaireExiste = (Err.Number = 1004)
End Function

Private Sub MettreAJourListeEtabs(etabs As Collection)
    Dim i As Long, liste As String, ws As Object
    Set ws = ThisWorkbook.Worksheets("ACCUEIL")
    liste = ""
    For i = 1 To etabs.Count
        If i > 1 Then liste = liste & ","
        liste = liste & Etab_Code(etabs(i))
    Next i
    With ws.Range("SG_ETAB_CODE")
        .Validation.Delete
        If Len(liste) > 0 Then
            .Validation.Add Type:=3, AlertStyle:=1, Operator:=1, Formula1:=liste
            .Value = Etab_Code(mEtab)
        Else
            .Value = ""
        End If
    End With
End Sub

Public Sub SG_ChoisirEtablissement()
    Dim code As String, id As String
    On Error GoTo erreur
    code = CStr(ThisWorkbook.Worksheets("ACCUEIL").Range("SG_ETAB_CODE").Value)
    id = DAL_First("T_Etablissements", RecFrom("Code=" & code))
    If Len(id) = 0 Or Not Sec_AAccesEtab(id) Then SG_Raise SG_ERR_ACCES, "Etablissement non autorise : " & code
    mEtab = id
    AfficherBandeau
    Exit Sub
erreur:
    UI_Erreur
End Sub

Private Function EtabCourant() As String
    If Len(mEtab) = 0 Then SG_Raise SG_ERR_ACCES, "Se connecter et choisir un etablissement."
    EtabCourant = mEtab
End Function

Private Sub AfficherBandeau()
    Dim ws As Object, e As Variant, txt As String
    Set ws = ThisWorkbook.Worksheets("ACCUEIL")
    e = Lic_Etat()
    If Sec_EstConnecte() Then txt = "Connecte : " & Sec_LoginCourant() Else txt = "Non connecte"
    If Len(mEtab) > 0 Then txt = txt & "   |   " & DAL_Value("T_Etablissements", mEtab, "Nom")
    ws.Range("SG_SESSION").Value = txt
    ws.Range("SG_LICENCE").Value = "Licence : " & RecStr(e, "Statut") & " - " & RecStr(e, "Message") & IIf(Len(RecStr(e, "Alerte")) > 0, "  /  " & RecStr(e, "Alerte"), "")
    ws.Range("SG_LICENCE").Interior.Color = IIf(RecStr(e, "Ecriture") = "1" And Len(RecStr(e, "Alerte")) = 0, C_FOND, C_ALERTE)
End Sub

'-------------------------------------------------------------------------------
' Lecture des cellules nommees
'-------------------------------------------------------------------------------
Private Function Saisie(ByVal nom As String) As String
    Dim v As Variant
    v = ThisWorkbook.Names(nom).RefersToRange.Value
    If IsDate(v) And VarType(v) = vbDate Then
        Saisie = DateISO(CDate(v))
    Else
        Saisie = Trim$(NzStr(v))
    End If
End Function

Private Sub Afficher(ByVal nom As String, ByVal valeur As Variant)
    ThisWorkbook.Names(nom).RefersToRange.Value = valeur
End Sub

Private Function ChampsFiche() As Variant
    ChampsFiche = Array("Matricule", "Nom", "Prenoms", "Sexe", "DateNaissance", "LieuNaissance", "Nationalite", "Adresse", "Telephone", "EtabPrecedent", "Statut")
End Function

Private Function EleveParMatricule(ByVal matricule As String) As String
    EleveParMatricule = DAL_First("T_Eleves", RecFrom("EtabID=" & EtabCourant() & "|Matricule=" & matricule))
    If Len(EleveParMatricule) = 0 Then SG_Raise SG_ERR_INTROUVABLE, "Aucun dossier avec le matricule " & matricule & " dans cet etablissement."
End Function

'-------------------------------------------------------------------------------
' Fiche eleve
'-------------------------------------------------------------------------------
Public Sub SG_FicheNouvelle()
    Dim c As Variant
    For Each c In ChampsFiche()
        Afficher "F_" & c, ""
    Next c
    ThisWorkbook.Worksheets("FICHE").Activate
End Sub

Public Sub SG_FicheEnregistrer()
    Dim r As Variant, c As Variant, id As String, forcer As Boolean, motif As String
    On Error GoTo erreur
    If Len(Saisie("F_Matricule")) > 0 Then SG_Raise SG_ERR_VALIDATION, "Ce dossier existe deja : utiliser 'Modifier un champ'."
    r = RecNew()
    For Each c In ChampsFiche()
        If c <> "Matricule" And c <> "Statut" Then RecSet r, CStr(c), Saisie("F_" & c)
    Next c
    On Error Resume Next
    id = Eleve_Creer(EtabCourant(), r)
    If Err.Number = SG_ERR_DOUBLON Then
        If MsgBox(Err.Description & vbLf & vbLf & "S'agit-il d'un homonyme (personne differente) ?", vbYesNo + vbQuestion, "SG - Doublon probable") = vbYes Then
            Err.Clear
            motif = InputBox("Motif (piece justificative verifiee) :", "SG - Confirmation d'homonyme")
            On Error GoTo erreur
            id = Eleve_Creer(EtabCourant(), r, True, motif)
        Else
            Exit Sub
        End If
    ElseIf Err.Number <> 0 Then
        GoTo erreur
    End If
    On Error GoTo erreur
    ChargerFiche id
    MsgBox "Dossier cree : matricule " & Saisie("F_Matricule"), vbInformation, "SG"
    Exit Sub
erreur:
    UI_Erreur
End Sub

Private Sub ChargerFiche(ByVal eleveId As String)
    Dim e As Variant, c As Variant
    e = Eleve_Get(EtabCourant(), eleveId)
    For Each c In ChampsFiche()
        Afficher "F_" & c, RecStr(e, CStr(c))
    Next c
End Sub

Public Sub SG_FicheCharger()
    Dim m As String
    On Error GoTo erreur
    m = InputBox("Matricule :", "SG - Ouvrir un dossier", Saisie("F_Matricule"))
    If Len(m) = 0 Then Exit Sub
    ChargerFiche EleveParMatricule(m)
    Exit Sub
erreur:
    UI_Erreur
End Sub

Public Sub SG_FicheModifierChamp()
    Dim champ As String, valeur As String, motif As String, id As String
    On Error GoTo erreur
    id = EleveParMatricule(Saisie("F_Matricule"))
    champ = InputBox("Champ a modifier (Nom, Prenoms, Sexe, DateNaissance, LieuNaissance, Nationalite, Adresse, Telephone, EtabPrecedent) :", "SG - Modification")
    If Len(champ) = 0 Then Exit Sub
    valeur = Saisie("F_" & champ)
    If MsgBox("Enregistrer la valeur affichee '" & valeur & "' pour " & champ & " ?", vbYesNo + vbQuestion, "SG") <> vbYes Then Exit Sub
    motif = InputBox("Motif de la modification :", "SG - Modification")
    Eleve_Modifier EtabCourant(), id, champ, valeur, motif
    ChargerFiche id
    Exit Sub
erreur:
    UI_Erreur
End Sub

'-------------------------------------------------------------------------------
' Liste et recherche
'-------------------------------------------------------------------------------
Public Sub SG_ListeRechercher()
    Dim ws As Object, res As Collection, i As Long, e As Variant, ligne As Long, an As String, ins As String
    On Error GoTo erreur
    Set ws = ThisWorkbook.Worksheets("ELEVES")
    Set res = Eleve_Rechercher(EtabCourant(), Saisie("L_Recherche"), Saisie("L_Statut"))
    ws.Range("A8:G5000").ClearContents
    an = Annee_Ouverte(EtabCourant())
    ligne = 8
    For i = 1 To res.Count
        e = DAL_Get("T_Eleves", res(i))
        ws.Cells(ligne, 1).Value = RecStr(e, "Matricule")
        ws.Cells(ligne, 2).Value = RecStr(e, "Nom")
        ws.Cells(ligne, 3).Value = RecStr(e, "Prenoms")
        ws.Cells(ligne, 4).Value = RecStr(e, "Sexe")
        ws.Cells(ligne, 5).Value = RecStr(e, "DateNaissance")
        If Len(an) > 0 Then
            ins = Inscr_Active(EtabCourant(), res(i), an)
            If Len(ins) > 0 Then ws.Cells(ligne, 6).Value = DAL_Value("T_Classes", DAL_Value("T_Inscriptions", ins, "ClasseID"), "Code")
        End If
        ws.Cells(ligne, 7).Value = RecStr(e, "Statut")
        ligne = ligne + 1
    Next i
    ws.Range("SG_LISTE_NB").Value = res.Count & " dossier(s)"
    Exit Sub
erreur:
    UI_Erreur
End Sub

'-------------------------------------------------------------------------------
' Inscriptions
'-------------------------------------------------------------------------------
Private Function ClasseParCode(ByVal code As String, ByVal anneeId As String) As String
    ClasseParCode = DAL_First("T_Classes", RecFrom("EtabID=" & EtabCourant() & "|AnneeID=" & anneeId & "|Code=" & UCase$(code)))
    If Len(ClasseParCode) = 0 Then SG_Raise SG_ERR_INTROUVABLE, "Classe " & code & " introuvable pour l'annee choisie."
End Function

Private Function AnneeParCode(ByVal code As String) As String
    If Len(code) = 0 Then
        AnneeParCode = Annee_Ouverte(EtabCourant())
    Else
        AnneeParCode = DAL_First("T_Annees", RecFrom("EtabID=" & EtabCourant() & "|Code=" & code))
    End If
    If Len(AnneeParCode) = 0 Then SG_Raise SG_ERR_INTROUVABLE, "Annee scolaire introuvable."
End Function

Public Sub SG_Inscrire()
    Dim an As String, id As String
    On Error GoTo erreur
    an = AnneeParCode(Saisie("I_Annee"))
    id = Inscr_Inscrire(EtabCourant(), EleveParMatricule(Saisie("I_Matricule")), an, ClasseParCode(Saisie("I_Classe"), an), _
                        UCase$(Saisie("I_Type")), Saisie("I_DateEffet"))
    MsgBox "Inscription enregistree (" & id & ").", vbInformation, "SG"
    Exit Sub
erreur:
    UI_Erreur
End Sub

Public Sub SG_ChangerClasse()
    Dim an As String, ins As String, motif As String
    On Error GoTo erreur
    an = AnneeParCode(Saisie("I_Annee"))
    ins = Inscr_Active(EtabCourant(), EleveParMatricule(Saisie("I_Matricule")), an)
    If Len(ins) = 0 Then SG_Raise SG_ERR_INTROUVABLE, "Aucune inscription active pour cet eleve et cette annee."
    motif = InputBox("Motif du changement de classe :", "SG")
    Inscr_ChangerClasse EtabCourant(), ins, ClasseParCode(Saisie("I_Classe"), an), Saisie("I_DateEffet"), motif
    MsgBox "Changement de classe enregistre. L'affectation precedente est conservee dans l'historique.", vbInformation, "SG"
    Exit Sub
erreur:
    UI_Erreur
End Sub

Public Sub SG_TerminerInscription()
    Dim an As String, ins As String, typ As String, motif As String
    On Error GoTo erreur
    an = AnneeParCode(Saisie("I_Annee"))
    ins = Inscr_Active(EtabCourant(), EleveParMatricule(Saisie("I_Matricule")), an)
    If Len(ins) = 0 Then SG_Raise SG_ERR_INTROUVABLE, "Aucune inscription active pour cet eleve et cette annee."
    typ = UCase$(InputBox("Type de fin : DEPART, TRANSFERT_SORTANT, RADIATION ou FIN_CYCLE", "SG"))
    If Len(typ) = 0 Then Exit Sub
    motif = InputBox("Motif :", "SG")
    If MsgBox("Confirmer la fin d'inscription (" & typ & ") ? Le dossier est conserve.", vbYesNo + vbExclamation, "SG") <> vbYes Then Exit Sub
    Inscr_Terminer EtabCourant(), ins, typ, Saisie("I_DateEffet"), motif
    MsgBox "Fin d'inscription enregistree.", vbInformation, "SG"
    Exit Sub
erreur:
    UI_Erreur
End Sub

'-------------------------------------------------------------------------------
' Tableau de bord, qualite, journal
'-------------------------------------------------------------------------------
Public Sub SG_TableauBordActualiser()
    Dim ws As Object, an As String, ind As Variant, tableau As Variant, i As Long, j As Long, k As Variant, ligne As Long
    On Error GoTo erreur
    Set ws = ThisWorkbook.Worksheets("TABLEAU_BORD")
    an = AnneeParCode(Saisie("T_Annee"))
    ind = Rapport_Indicateurs(EtabCourant(), an)
    ws.Range("A6:H200").ClearContents
    ws.Range("A6:H200").Interior.Pattern = -4142
    ligne = 6
    For Each k In Split("Inscrits;Premieres;Reinscriptions;TransfertsEntrants;Passages;Redoublements;Departs;TransfertsSortants;Radiations;FinsCycle;Annulees;Filles;Garcons", ";")
        ws.Cells(ligne, 1).Value = CStr(k)
        ws.Cells(ligne, 2).Value = RecGet(ind, CStr(k), 0)
        ligne = ligne + 1
    Next k
    tableau = Rapport_EffectifsClasses(EtabCourant(), an)
    For i = 0 To UBound(tableau, 1)
        For j = 0 To 6
            ws.Cells(6 + i, 4 + j).Value = tableau(i, j)
        Next j
        If i > 0 Then
            If tableau(i, 6) = "SUREFFECTIF" Then ws.Range(ws.Cells(6 + i, 4), ws.Cells(6 + i, 10)).Interior.Color = C_ALERTE
        End If
    Next i
    ws.Range("SG_TB_MAJ").Value = "Calcule le " & NowStamp() & " a partir des donnees sources."
    Exit Sub
erreur:
    UI_Erreur
End Sub

Public Sub SG_QualiteActualiser()
    Dim ws As Object, qc As Collection, i As Long, p As Variant, j As Long
    On Error GoTo erreur
    Set ws = ThisWorkbook.Worksheets("QUALITE")
    Set qc = QC_Anomalies(EtabCourant())
    ws.Range("A6:E5000").ClearContents
    For i = 1 To qc.Count
        p = Split(qc(i), "|")
        For j = 0 To 4
            ws.Cells(5 + i, j + 1).Value = p(j)
        Next j
    Next i
    ws.Range("SG_QC_NB").Value = qc.Count & " anomalie(s) - aucune correction automatique"
    Exit Sub
erreur:
    UI_Erreur
End Sub

Public Sub SG_JournalActualiser()
    Dim ws As Object, src As Object, n As Long, i As Long, ligne As Long, c As Long, colEtab As Long, cols As Variant
    On Error GoTo erreur
    Sec_Exiger EtabCourant(), "AUDIT", NIV_LECTURE
    Set ws = ThisWorkbook.Worksheets("JOURNAL")
    Set src = DAL_Ws("T_Audit")
    ws.Range("A6:M5000").ClearContents
    cols = Array("Horodatage", "UserID", "Module", "Operation", "TableCible", "EnregistrementID", "Champ", "AncienneValeur", "NouvelleValeur", "Motif", "Resultat", "TxRef")
    colEtab = DAL_Col(src, "EtabID")
    n = DAL_LastRow(src)
    ligne = 6
    For i = n To 2 Step -1
        If CStr(src.Cells(i, colEtab).Value) = EtabCourant() Then
            For c = 0 To UBound(cols)
                ws.Cells(ligne, c + 1).Value = src.Cells(i, DAL_Col(src, CStr(cols(c)))).Value
            Next c
            ligne = ligne + 1
            If ligne > 505 Then Exit For
        End If
    Next i
    Audit_Log EtabCourant(), "AUDIT", "CONSULTATION_JOURNAL", "T_Audit", "", "", "", CStr(ligne - 6) & " evenements", ""
    Exit Sub
erreur:
    UI_Erreur
End Sub

Public Sub SG_JournalVerifier()
    Dim r As String, ancre As String
    On Error GoTo erreur
    ' Droit global (administrateurs) ou lecture du journal de l'etablissement
    ' courant (directeur, auditeur) : le bouton est sur l'ecran JOURNAL.
    If Sec_Niveau("*", "AUDIT") < NIV_LECTURE Then Sec_Exiger EtabCourant(), "AUDIT", NIV_LECTURE
    ancre = InputBox("Empreinte d'ancre externe a controler (laisser vide pour verifier la chaine seule) :", "SG - Integrite du journal")
    r = Audit_Verifier(ancre)
    If Left$(r, 2) = "OK" Then
        MsgBox "Journal integre : " & Split(r, "|")(1) & " evenements verifies." & vbLf & "Empreinte de tete : " & Audit_HashTete(), vbInformation, "SG"
    Else
        MsgBox "ALERTE : journal altere." & vbLf & r, vbCritical, "SG"
    End If
    Exit Sub
erreur:
    UI_Erreur
End Sub

'-------------------------------------------------------------------------------
' Licence et sauvegarde
'-------------------------------------------------------------------------------
Public Sub SG_LicenceInstaller()
    Dim f As Variant, st As String
    On Error GoTo erreur
    f = Application.GetOpenFilename("Licence SG (*.sglic),*.sglic", , "Choisir le fichier de licence")
    If VarType(f) = vbBoolean Then Exit Sub
    st = Lic_Installer(LireFichierUtf8(CStr(f)))
    AfficherBandeau
    MsgBox "Licence installee : " & st & vbLf & Lic_MessageCourant(), vbInformation, "SG"
    Exit Sub
erreur:
    UI_Erreur
End Sub

Public Sub SG_LicenceDemande()
    Dim f As Variant, titulaire As String, codes As String
    On Error GoTo erreur
    titulaire = InputBox("Titulaire de la licence (nom juridique) :", "SG - Demande d'activation")
    codes = InputBox("Codes des etablissements (separes par ;) :", "SG - Demande d'activation")
    f = Application.GetSaveAsFilename("demande_activation.sgreq", "Demande SG (*.sgreq),*.sgreq")
    If VarType(f) = vbBoolean Then Exit Sub
    EcrireFichierUtf8 CStr(f), Lic_GenererDemande(titulaire, codes)
    MsgBox "Demande enregistree. Transmettez ce fichier a l'editeur ; il vous renverra un fichier .sglic.", vbInformation, "SG"
    Exit Sub
erreur:
    UI_Erreur
End Sub

Public Sub SG_LicenceRevocations()
    Dim f As Variant, n As Long
    On Error GoTo erreur
    f = Application.GetOpenFilename("Liste de revocation SG (*.sgrev),*.sgrev")
    If VarType(f) = vbBoolean Then Exit Sub
    n = Lic_ImporterRevocations(LireFichierUtf8(CStr(f)))
    Lic_Recharger
    AfficherBandeau
    MsgBox n & " revocation(s) importee(s).", vbInformation, "SG"
    Exit Sub
erreur:
    UI_Erreur
End Sub

Public Sub SG_Sauvegarder()
    Dim d As Object, dossier As String, sid As String
    On Error GoTo erreur
    Set d = Application.FileDialog(4) ' msoFileDialogFolderPicker
    d.Title = "Dossier de sauvegarde (de preference hors de ce poste : cle, NAS, cloud)"
    If d.Show <> -1 Then Exit Sub
    dossier = d.SelectedItems(1)
    sid = Sauv_Creer(IIf(Etab_Mode() = "CENTRALISE", "*", EtabCourant()), dossier)
    MsgBox "Sauvegarde creee et VERIFIEE :" & vbLf & DAL_Value("T_Sauvegardes", sid, "Chemin") & vbLf & vbLf & _
           "Ancre a conserver hors du classeur :" & vbLf & DAL_Value("T_Sauvegardes", sid, "HashAudit"), vbInformation, "SG"
    Exit Sub
erreur:
    UI_Erreur
End Sub

Public Sub SG_ExporterEtablissement()
    Dim d As Object, n As Long
    On Error GoTo erreur
    Set d = Application.FileDialog(4)
    If d.Show <> -1 Then Exit Sub
    n = Export_EtablissementCSV(EtabCourant(), d.SelectedItems(1))
    MsgBox n & " ligne(s) exportee(s) (CSV UTF-8, donnees de cet etablissement uniquement).", vbInformation, "SG"
    Exit Sub
erreur:
    UI_Erreur
End Sub

'-------------------------------------------------------------------------------
' Documents
'-------------------------------------------------------------------------------
Public Sub SG_CertificatScolarite()
    Dim c As Variant, ws As Object, champ As Variant, f As Variant
    On Error GoTo erreur
    c = Doc_CertificatScolarite(EtabCourant(), EleveParMatricule(Saisie("F_Matricule")))
    Set ws = ThisWorkbook.Worksheets("DOC_CERTIFICAT")
    For Each champ In Split("Numero;DateEmission;Etablissement;Ville;Matricule;Nom;Prenoms;DateNaissance;LieuNaissance;Classe;Annee;Signataire", ";")
        Afficher "D_" & champ, RecStr(c, CStr(champ))
    Next champ
    f = Application.GetSaveAsFilename(RecStr(c, "Numero") & ".pdf", "PDF (*.pdf),*.pdf")
    If VarType(f) <> vbBoolean Then ws.ExportAsFixedFormat Type:=0, Filename:=CStr(f), Quality:=0
    ws.Activate
    Exit Sub
erreur:
    UI_Erreur
End Sub

'-------------------------------------------------------------------------------
' Construction de l'interface (appelee une fois par l'installateur)
'-------------------------------------------------------------------------------
Private Function Feuille(ByVal nom As String) As Object
    Dim ws As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nom)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = nom
    End If
    ws.Cells.Clear
    Dim s As Object
    For Each s In ws.Shapes
        s.Delete
    Next s
    ws.Cells.Interior.Color = C_FOND
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 10
    ActiveWindow.DisplayGridlines = False
    Set Feuille = ws
End Function

Private Sub Titre(ws As Object, ByVal texte As String, ByVal sousTitre As String)
    With ws.Range("A1:L1")
        .Merge
        .Value = texte
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = C_PRIMAIRE
        .RowHeight = 34
        .VerticalAlignment = -4108
    End With
    ws.Range("A2").Value = sousTitre
    ws.Range("A2").Font.Italic = True
End Sub

Private Sub Bouton(ws As Object, ByVal cellule As String, ByVal libelle As String, ByVal macro As String, Optional ByVal largeur As Double = 150)
    Dim r As Object, s As Object
    Set r = ws.Range(cellule)
    Set s = ws.Shapes.AddShape(5, r.Left, r.Top, largeur, 24) ' msoShapeRoundedRectangle
    s.Fill.ForeColor.RGB = C_PRIMAIRE
    s.Line.Visible = False
    s.TextFrame2.TextRange.Text = libelle
    s.TextFrame2.TextRange.Font.Size = 10
    s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    s.TextFrame2.VerticalAnchor = 3
    s.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    s.OnAction = "'" & ThisWorkbook.Name & "'!" & macro
End Sub

Private Sub ChampSaisie(ws As Object, ByVal ligne As Long, ByVal libelle As String, ByVal nom As String, Optional ByVal lectureSeule As Boolean = False)
    ws.Cells(ligne, 2).Value = libelle
    ws.Cells(ligne, 2).HorizontalAlignment = -4152
    With ws.Cells(ligne, 3)
        .Interior.Color = IIf(lectureSeule, C_FOND, C_SAISIE)
        .Borders.Color = RGB(200, 200, 200)
        .NumberFormat = "@"
        .Locked = lectureSeule
    End With
    ws.Range(ws.Cells(ligne, 3), ws.Cells(ligne, 5)).Merge
    ThisWorkbook.Names.Add Name:=nom, RefersTo:="='" & ws.Name & "'!" & ws.Cells(ligne, 3).Address
End Sub

Public Sub SG_ConstruireInterface()
    Dim ws As Object, champs As Variant, i As Long
    Application.ScreenUpdating = False

    Set ws = Feuille("DOC_CERTIFICAT")
    ws.Columns("A").ColumnWidth = 4: ws.Columns("B").ColumnWidth = 28: ws.Columns("C:F").ColumnWidth = 18
    ws.Range("B2").Value = "": ThisWorkbook.Names.Add Name:="D_Etablissement", RefersTo:="='DOC_CERTIFICAT'!$B$2"
    ws.Range("B2").Font.Size = 14: ws.Range("B2").Font.Bold = True
    ws.Range("B3").Value = "": ThisWorkbook.Names.Add Name:="D_Ville", RefersTo:="='DOC_CERTIFICAT'!$B$3"
    ws.Range("B6:F6").Merge: ws.Range("B6").Value = "CERTIFICAT DE SCOLARITE": ws.Range("B6").Font.Size = 16
    ws.Range("B6").Font.Bold = True: ws.Range("B6").HorizontalAlignment = -4108
    ws.Range("B7").Value = "N°": ThisWorkbook.Names.Add Name:="D_Numero", RefersTo:="='DOC_CERTIFICAT'!$C$7"
    ws.Range("B9").Value = "Le chef d'etablissement soussigne certifie que l'eleve :"
    champs = Array("Nom", "Prenoms", "DateNaissance", "LieuNaissance", "Matricule", "Classe", "Annee")
    For i = 0 To UBound(champs)
        ws.Cells(11 + i, 2).Value = Choose(i + 1, "Nom", "Prenoms", "Ne(e) le", "A", "Matricule", "Classe", "Annee scolaire")
        ThisWorkbook.Names.Add Name:="D_" & champs(i), RefersTo:="='DOC_CERTIFICAT'!" & ws.Cells(11 + i, 3).Address
        ws.Cells(11 + i, 3).Font.Bold = True
    Next i
    ws.Range("B19").Value = "est regulierement inscrit(e) dans l'etablissement pour l'annee scolaire indiquee."
    ws.Range("B21").Value = "Fait le": ThisWorkbook.Names.Add Name:="D_DateEmission", RefersTo:="='DOC_CERTIFICAT'!$C$21"
    ws.Range("E23").Value = "Signature et cachet"
    ws.Range("E24").Value = "": ThisWorkbook.Names.Add Name:="D_Signataire", RefersTo:="='DOC_CERTIFICAT'!$E$25"
    ws.Range("B30").Value = "Document genere par SG - numero unique enregistre au journal d'audit."
    ws.Range("B30").Font.Size = 8
    ws.PageSetup.PrintArea = "$A$1:$F$31"
    ws.Cells.Interior.Pattern = -4142

    Set ws = Feuille("JOURNAL")
    Titre ws, "Journal d'audit", "500 derniers evenements de l'etablissement courant (lecture seule)."
    For i = 0 To 11
        ws.Cells(5, i + 1).Value = Choose(i + 1, "Horodatage", "Utilisateur", "Module", "Operation", "Table", "Enregistrement", "Champ", "Ancienne valeur", "Nouvelle valeur", "Motif", "Resultat", "Transaction")
    Next i
    ws.Rows(5).Font.Bold = True
    Bouton ws, "A3", "Actualiser", "SG_JournalActualiser"
    Bouton ws, "D3", "Verifier l'integrite", "SG_JournalVerifier"

    Set ws = Feuille("QUALITE")
    Titre ws, "Controle qualite des donnees", "Anomalies detectees ; aucune correction n'est faite automatiquement."
    For i = 0 To 4
        ws.Cells(5, i + 1).Value = Choose(i + 1, "Gravite", "Categorie", "Enregistrement", "Description", "Correction proposee")
    Next i
    ws.Rows(5).Font.Bold = True
    ws.Columns("D:E").ColumnWidth = 60
    ws.Range("G3").Value = "": ThisWorkbook.Names.Add Name:="SG_QC_NB", RefersTo:="='QUALITE'!$G$3"
    Bouton ws, "A3", "Lancer le controle", "SG_QualiteActualiser"

    Set ws = Feuille("TABLEAU_BORD")
    Titre ws, "Tableau de bord", "Indicateurs recalcules a partir des donnees sources."
    ChampSaisie ws, 3, "Annee (vide = ouverte)", "T_Annee"
    ws.Range("A5").Value = "Indicateur": ws.Range("B5").Value = "Valeur"
    ws.Rows(5).Font.Bold = True
    ws.Range("A40").Value = "": ThisWorkbook.Names.Add Name:="SG_TB_MAJ", RefersTo:="='TABLEAU_BORD'!$A$40"
    ws.Columns("A").ColumnWidth = 22
    Bouton ws, "G3", "Actualiser", "SG_TableauBordActualiser"

    Set ws = Feuille("INSCRIPTION")
    Titre ws, "Inscriptions", "Inscription, changement de classe et fin d'inscription. Aucune donnee n'est supprimee."
    ChampSaisie ws, 5, "Matricule", "I_Matricule"
    ChampSaisie ws, 6, "Annee (vide = ouverte)", "I_Annee"
    ChampSaisie ws, 7, "Classe (code)", "I_Classe"
    ChampSaisie ws, 8, "Type (PREMIERE / REINSCRIPTION / TRANSFERT_ENTRANT)", "I_Type"
    ChampSaisie ws, 9, "Date d'effet (AAAA-MM-JJ)", "I_DateEffet"
    ws.Range("C8").Validation.Add Type:=3, AlertStyle:=1, Operator:=1, Formula1:="PREMIERE,REINSCRIPTION,TRANSFERT_ENTRANT"
    ws.Columns("B").ColumnWidth = 45
    Bouton ws, "B12", "Inscrire", "SG_Inscrire"
    Bouton ws, "C12", "Changer de classe", "SG_ChangerClasse"
    Bouton ws, "E12", "Terminer l'inscription", "SG_TerminerInscription"

    Set ws = Feuille("ELEVES")
    Titre ws, "Eleves", "Recherche sans accents ni majuscules (nom, prenoms, matricule)."
    ChampSaisie ws, 3, "Recherche", "L_Recherche"
    ChampSaisie ws, 4, "Statut (optionnel)", "L_Statut"
    ws.Range("C4").Validation.Add Type:=3, AlertStyle:=1, Operator:=1, Formula1:=Replace(STATUTS_ELEVE, ";", ",")
    For i = 0 To 6
        ws.Cells(7, i + 1).Value = Choose(i + 1, "Matricule", "Nom", "Prenoms", "Sexe", "Naissance", "Classe", "Statut")
    Next i
    ws.Rows(7).Font.Bold = True
    ws.Range("I3").Value = "": ThisWorkbook.Names.Add Name:="SG_LISTE_NB", RefersTo:="='ELEVES'!$I$3"
    ws.Columns("A:G").ColumnWidth = 18
    Bouton ws, "G3", "Rechercher", "SG_ListeRechercher"

    Set ws = Feuille("FICHE")
    Titre ws, "Fiche eleve", "Les champs grises sont geres par le systeme. Toute modification est journalisee."
    champs = Array("Matricule", "Nom", "Prenoms", "Sexe", "DateNaissance", "LieuNaissance", "Nationalite", "Adresse", "Telephone", "EtabPrecedent", "Statut")
    For i = 0 To UBound(champs)
        ChampSaisie ws, 4 + i, Choose(i + 1, "Matricule", "Nom", "Prenoms", "Sexe (M/F)", "Date de naissance (AAAA-MM-JJ)", "Lieu de naissance", "Nationalite", "Adresse", "Telephone", "Etablissement precedent", "Statut"), _
                    "F_" & champs(i), (champs(i) = "Matricule" Or champs(i) = "Statut")
    Next i
    ws.Range("C7").Validation.Add Type:=3, AlertStyle:=1, Operator:=1, Formula1:="M,F"
    ws.Columns("B").ColumnWidth = 32
    Bouton ws, "B17", "Nouvelle fiche", "SG_FicheNouvelle"
    Bouton ws, "C17", "Enregistrer", "SG_FicheEnregistrer"
    Bouton ws, "E17", "Ouvrir (matricule)", "SG_FicheCharger"
    Bouton ws, "B19", "Modifier un champ", "SG_FicheModifierChamp"
    Bouton ws, "C19", "Certificat de scolarite", "SG_CertificatScolarite"

    Set ws = Feuille("ACCUEIL")
    Titre ws, "SG - Gestion scolaire", "Version " & SG_VERSION & " - donnees de demonstration fictives uniquement."
    ws.Range("A4:L4").Merge: ThisWorkbook.Names.Add Name:="SG_SESSION", RefersTo:="='ACCUEIL'!$A$4"
    ws.Range("A5:L5").Merge: ThisWorkbook.Names.Add Name:="SG_LICENCE", RefersTo:="='ACCUEIL'!$A$5"
    ws.Range("A7").Value = "Etablissement :"
    ws.Range("B7").Interior.Color = C_SAISIE
    ThisWorkbook.Names.Add Name:="SG_ETAB_CODE", RefersTo:="='ACCUEIL'!$B$7"
    ws.Columns("A:L").ColumnWidth = 14
    Bouton ws, "D7", "Choisir", "SG_ChoisirEtablissement", 90
    Bouton ws, "A10", "Connexion", "SG_Connexion"
    Bouton ws, "C10", "Deconnexion", "SG_Deconnexion"
    Bouton ws, "A12", "Eleves", "modUI.AllerEleves"
    Bouton ws, "C12", "Fiche eleve", "SG_FicheNouvelle"
    Bouton ws, "E12", "Inscriptions", "modUI.AllerInscription"
    Bouton ws, "A14", "Tableau de bord", "modUI.AllerTableauBord"
    Bouton ws, "C14", "Controle qualite", "modUI.AllerQualite"
    Bouton ws, "E14", "Journal d'audit", "modUI.AllerJournal"
    Bouton ws, "A16", "Sauvegarder", "SG_Sauvegarder"
    Bouton ws, "C16", "Exporter (CSV)", "SG_ExporterEtablissement"
    Bouton ws, "A18", "Installer une licence", "SG_LicenceInstaller"
    Bouton ws, "C18", "Demande d'activation", "SG_LicenceDemande"
    Bouton ws, "E18", "Importer revocations", "SG_LicenceRevocations"
    ws.Activate
    Application.ScreenUpdating = True
End Sub

Public Sub AllerEleves(): ThisWorkbook.Worksheets("ELEVES").Activate: End Sub
Public Sub AllerInscription(): ThisWorkbook.Worksheets("INSCRIPTION").Activate: End Sub
Public Sub AllerTableauBord(): ThisWorkbook.Worksheets("TABLEAU_BORD").Activate: End Sub
Public Sub AllerQualite(): ThisWorkbook.Worksheets("QUALITE").Activate: End Sub
Public Sub AllerJournal(): ThisWorkbook.Worksheets("JOURNAL").Activate: End Sub

'-------------------------------------------------------------------------------
' Formulaire de connexion (mot de passe masque), cree par l'installateur.
' Necessite l'option "Acces approuve au modele d'objet du projet VBA".
'-------------------------------------------------------------------------------
Public Sub SG_CreerFormulaireConnexion()
    Dim comp As Object, f As Object, c As Object
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents("frmConnexion")
    On Error GoTo 0
    Set comp = ThisWorkbook.VBProject.VBComponents.Add(3) ' vbext_ct_MSForm
    comp.Name = "frmConnexion"
    comp.Properties("Caption") = "SG - Connexion"
    comp.Properties("Width") = 260
    comp.Properties("Height") = 170
    Set f = comp.Designer
    Set c = f.Controls.Add("Forms.Label.1", "lblLogin"): c.Caption = "Identifiant": c.Left = 12: c.Top = 14: c.Width = 70
    Set c = f.Controls.Add("Forms.TextBox.1", "txtLogin"): c.Left = 90: c.Top = 12: c.Width = 145
    Set c = f.Controls.Add("Forms.Label.1", "lblMdp"): c.Caption = "Mot de passe": c.Left = 12: c.Top = 44: c.Width = 70
    Set c = f.Controls.Add("Forms.TextBox.1", "txtMdp"): c.Left = 90: c.Top = 42: c.Width = 145: c.PasswordChar = "*"
    Set c = f.Controls.Add("Forms.CommandButton.1", "btnOK"): c.Caption = "Se connecter": c.Left = 90: c.Top = 80: c.Width = 80: c.Default = True
    Set c = f.Controls.Add("Forms.CommandButton.1", "btnAnnuler"): c.Caption = "Annuler": c.Left = 175: c.Top = 80: c.Width = 60: c.Cancel = True
    comp.CodeModule.AddFromString _
        "Private Sub btnOK_Click()" & vbLf & _
        "    On Error GoTo erreur" & vbLf & _
        "    Sec_Connexion txtLogin.Text, txtMdp.Text" & vbLf & _
        "    txtMdp.Text = """"" & vbLf & _
        "    Unload Me" & vbLf & _
        "    ApresConnexion" & vbLf & _
        "    Exit Sub" & vbLf & _
        "erreur:" & vbLf & _
        "    txtMdp.Text = """"" & vbLf & _
        "    MsgBox Err.Description, vbExclamation, ""SG - Connexion""" & vbLf & _
        "End Sub" & vbLf & _
        "Private Sub btnAnnuler_Click()" & vbLf & _
        "    Unload Me" & vbLf & _
        "End Sub"
End Sub

'-------------------------------------------------------------------------------
' Installation (appelee par installer/Installer-SG.ps1 ou manuellement)
'-------------------------------------------------------------------------------
' avecDemo : charge la demonstration dans la MEME execution que la connexion de
' l'administrateur. La creation du formulaire par code reinitialise le projet
' VBA a la fin de la macro : la session serait perdue pour un appel separe.
Public Sub SG_Installer(Optional ByVal mode As Variant, Optional ByVal avecDemo As Variant)
    ' Valeurs par defaut (IsMissing : forme compatible Excel et LibreOffice).
    If IsMissing(mode) Then mode = "INDEPENDANT"
    If IsMissing(avecDemo) Then avecDemo = False
    Dim login As String, mdp As String, nom As String
    SG_InitialiserBase mode, 1000
    SG_ConstruireInterface
    On Error Resume Next
    SG_CreerFormulaireConnexion
    If Err.Number <> 0 Then MsgBox "Formulaire de connexion non cree (acces au projet VBA refuse). La connexion utilisera une boite de saisie.", vbExclamation, "SG"
    On Error GoTo 0
    ' Saisie redemandee tant qu'elle est refusee : une faute de frappe ne doit
    ' pas interrompre l'installation (Annuler sur l'identifiant abandonne).
    login = "admin"
    Do While DAL_Count("T_Utilisateurs", RecNew()) = 0
        login = InputBox("Identifiant de l'administrateur technique :", "SG - Installation", login)
        If Len(login) = 0 Then
            MsgBox "Installation interrompue : aucun administrateur cree. Relancer SG_Installer.", vbExclamation, "SG"
            Exit Sub
        End If
        nom = InputBox("Nom complet de l'administrateur :", "SG - Installation", nom)
        mdp = InputBox("Mot de passe initial (8 caracteres, lettres et chiffres). A changer apres installation :", "SG - Installation")
        On Error Resume Next
        Sec_InitialiserAdministrateur login, nom, mdp
        If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, "SG - Installation"
        On Error GoTo 0
    Loop
    If Len(mdp) > 0 Then Sec_Connexion login, mdp
    If CBool(avecDemo) Then
        If Sec_EstConnecte() Then
            MsgBox SG_ChargerDemo(), vbInformation, "SG - Demonstration"
        Else
            MsgBox "Demonstration non chargee : se connecter avec l'administrateur puis lancer SG_ChargerDemo.", vbExclamation, "SG"
        End If
    End If
    ProtegerDonnees
    AfficherBandeau
End Sub
