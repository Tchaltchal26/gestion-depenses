Attribute VB_Name = "modTests"
'===============================================================================
' SG - Module modTests (tests automatises)
' Execution : SG_ExecuterTests("", fixtures) depuis le lanceur LibreOffice
'   (tools/tests/run_vba_tests.py) ou depuis Excel (bouton "Tests" dans un
'   CLASSEUR DE TEST UNIQUEMENT : les tests effacent les donnees).
' Donnees : exclusivement fictives.
'===============================================================================
Option Explicit

Private mOK As Long
Private mKO As Long
Private mRapport As String
Private mTest As String
Private mFx As Variant
Private mDerniere As String
Private Const MDP As String = "Secret2026"
Private Const D_REF As String = "2026-10-17"

'-------------------------------------------------------------------------------
' Cadre de test
'-------------------------------------------------------------------------------
Private Sub Verifier(ByVal condition As Boolean, ByVal libelle As String)
    mDerniere = libelle
    If condition Then
        mOK = mOK + 1
    Else
        mKO = mKO + 1
        mRapport = mRapport & "  [ECHEC] " & mTest & " : " & libelle & vbLf
    End If
End Sub

Private Sub VerifierEgal(ByVal obtenu As Variant, ByVal attendu As Variant, ByVal libelle As String)
    Verifier CStr(obtenu) = CStr(attendu), libelle & " (attendu '" & attendu & "', obtenu '" & obtenu & "')"
End Sub

' A appeler juste apres une instruction executee sous On Error Resume Next.
Private Sub VerifierErreur(ByVal numeroObtenu As Long, ByVal descr As String, ByVal numeroAttendu As Long, ByVal libelle As String)
    mDerniere = libelle
    If numeroObtenu = numeroAttendu Then
        mOK = mOK + 1
    Else
        mKO = mKO + 1
        mRapport = mRapport & "  [ECHEC] " & mTest & " : " & libelle & " (attendu " & SG_CodeErreur(numeroAttendu) & _
                   ", obtenu " & IIf(numeroObtenu = 0, "aucune erreur", SG_CodeErreur(numeroObtenu) & " / " & descr) & ")" & vbLf
    End If
    Audit_ReinitialiserTx
End Sub

Private Function Fx(ByVal cle As String) As String
    Fx = RecStr(mFx, cle)
End Function

Public Function SG_ExecuterTests(Optional ByVal filtre As Variant, Optional ByVal fixtures As Variant) As String
    ' Valeurs par defaut (IsMissing : forme compatible Excel et LibreOffice).
    If IsMissing(filtre) Then filtre = ""
    If IsMissing(fixtures) Then fixtures = ""
    Dim tests As Variant, i As Long, t0 As Single, parts As Variant, p As Variant, kv As Variant
    mOK = 0: mKO = 0: mRapport = ""
    mFx = RecNew()
    If Len(fixtures) > 0 Then
        parts = Split(fixtures, Chr$(30))
        For Each p In parts
            kv = Split(p, Chr$(31))
            If UBound(kv) >= 1 Then RecSet mFx, CStr(kv(0)), CStr(kv(1))
        Next p
    End If
    DAL_ActiverModeTest
    On Error GoTo fatal
    SG_InitialiserBase "CENTRALISE", 2
    On Error GoTo 0
    tests = Array("Utilitaires", "CoucheDonnees", "Securite", "Permissions", "Isolation", "Annees", "Structure", _
                  "Eleves", "Inscriptions", "Reinscription", "Transfert", "Audit", "Sauvegarde", "Rapports", "Licences", "Schema")
    For i = LBound(tests) To UBound(tests)
        If Len(filtre) = 0 Or InStr(1, tests(i), filtre, vbTextCompare) > 0 Then
            mTest = tests(i)
            mDerniere = ""
            t0 = Timer
            On Error GoTo plantage
            Reinitialiser
            Select Case tests(i)
                Case "Utilitaires": T_Utilitaires
                Case "CoucheDonnees": T_CoucheDonnees
                Case "Securite": T_Securite
                Case "Permissions": T_Permissions
                Case "Isolation": T_Isolation
                Case "Annees": T_Annees
                Case "Structure": T_Structure
                Case "Eleves": T_Eleves
                Case "Inscriptions": T_Inscriptions
                Case "Reinscription": T_Reinscription
                Case "Transfert": T_Transfert
                Case "Audit": T_Audit
                Case "Sauvegarde": T_Sauvegarde
                Case "Rapports": T_Rapports
                Case "Licences": T_Licences
                Case "Schema": T_Schema
            End Select
            GoTo suivant
plantage:
            mKO = mKO + 1
            mRapport = mRapport & "  [ECHEC] " & mTest & " : erreur inattendue " & SG_CodeErreur(Err.Number) & " - " & Err.Description & _
                       " (apres la verification : " & IIf(Len(mDerniere) > 0, mDerniere, "aucune") & ")" & vbLf
            Resume suivant
suivant:
            On Error GoTo 0
            Audit_ReinitialiserTx
            mRapport = mRapport & "[TEST] " & tests(i) & " termine (" & Format(Timer - t0, "0.0") & " s)" & vbLf
        End If
    Next i
    SG_ExecuterTests = "RAPPORT DE TESTS SG " & SG_VERSION & " - " & NowStamp() & vbLf & mRapport & _
                       "VERIFICATIONS : " & (mOK + mKO) & " | REUSSIES : " & mOK & " | ECHECS : " & mKO
    Exit Function
fatal:
    SG_ExecuterTests = "ERREUR FATALE A L'INITIALISATION : " & Err.Number & " - " & Err.Description
End Function

' Base vierge + administrateur connecte.
Private Sub Reinitialiser()
    Dim t As Variant
    Sec_Deconnexion
    For Each t In SchemaTables()
        DAL_ViderPourTests CStr(t)
    Next t
    SG_InitialiserBase "CENTRALISE", 2
    Sec_InitialiserAdministrateur "admin", "Administrateur Technique", MDP
    Sec_Connexion "admin", MDP
    Lic_Recharger "", "", D_REF
End Sub

Private Function CreerEtab(ByVal code As String) As String
    CreerEtab = Etab_Creer(RecFrom("Code=" & code & "|Nom=College fictif " & code & "|TypeEtab=COMPLEXE|Statut=PRIVE|Pays=Pays fictif|Ville=Ville fictive"))
End Function

Private Function CreerUtilisateur(ByVal login As String, ByVal etabId As String, ByVal role As String) As String
    Dim uid As String
    uid = Sec_CreerUtilisateur(login, "Utilisateur " & login, MDP)
    If Len(etabId) > 0 Then Sec_Habiliter uid, etabId, role, "2026-01-01"
    CreerUtilisateur = uid
End Function

Private Sub Connecter(ByVal login As String)
    Sec_Connexion login, MDP
End Sub

' Contexte standard : etablissement, annee ouverte, 2 niveaux, 3 classes.
' Retour : Array(etab, annee, niv6, niv5, cls6A, cls6B, cls5A)
Private Function Contexte(ByVal code As String, Optional ByVal capacite As Long = 30) As Variant
    Dim e As String, a As String, cyc As String, n6 As String, n5 As String
    e = CreerEtab(code)
    Sec_Habiliter Sec_UtilisateurCourant(), e, "DIRECTEUR", "2026-01-01"
    a = Annee_Creer(e, "2026-2027", "", "2026-09-01", "2027-07-31")
    Annee_Ouvrir e, a
    cyc = Cycle_Creer(e, "COL", "College", "NOTES", 3)
    n6 = Niveau_Creer(e, cyc, "6E", "Sixieme", 1)
    n5 = Niveau_Creer(e, cyc, "5E", "Cinquieme", 2)
    Contexte = Array(e, a, n6, n5, _
                     Classe_Creer(e, a, n6, "", "6EA", "6e A", capacite), _
                     Classe_Creer(e, a, n6, "", "6EB", "6e B", capacite), _
                     Classe_Creer(e, a, n5, "", "5EA", "5e A", capacite))
End Function

Private Function NouvelEleve(ByVal etabId As String, ByVal nom As String, ByVal prenoms As String, _
                             Optional ByVal dn As Variant, Optional ByVal sexe As Variant) As String
    ' Valeurs par defaut (IsMissing : forme compatible Excel et LibreOffice).
    If IsMissing(dn) Then dn = "2014-03-12"
    If IsMissing(sexe) Then sexe = "F"
    Dim r As Variant
    r = RecNew()
    RecSet r, "Nom", nom
    RecSet r, "Prenoms", prenoms
    RecSet r, "Sexe", sexe
    RecSet r, "DateNaissance", dn
    RecSet r, "LieuNaissance", "Ville fictive"
    NouvelEleve = Eleve_Creer(etabId, r)
End Function

Private Function NbAudit(ByVal operation As String, ByVal enregId As String) As Long
    NbAudit = DAL_Count("T_Audit", RecFrom("Operation=" & operation & "|EnregistrementID=" & enregId))
End Function

'-------------------------------------------------------------------------------
' Tests
'-------------------------------------------------------------------------------
Private Sub T_Utilitaires()
    VerifierEgal SHA256Hex(""), "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855", "SHA-256 vide"
    VerifierEgal SHA256Hex("abc"), "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad", "SHA-256 abc"
    VerifierEgal SHA256Hex("abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq"), "248d6a61d20638b8e5c026930c3e6039a33ce45964ff2167f6ecedd419db06c1", "SHA-256 deux blocs"
    VerifierEgal SHA256Hex("Élève àéèô €"), "e924337cca858fe19c07a06425a0af978e5480e58d6da45482a022b97e387a90", "SHA-256 UTF-8"
    VerifierEgal Base64Encode(Utf8Encode("Élève")), "w4lsw6h2ZQ==", "Base64"
    VerifierEgal Utf8Decode(Base64Decode("w4lsw6h2ZQ==")), "Élève", "Base64 + UTF-8 aller-retour"
    Verifier IsISODate("2024-02-29"), "29 fevrier bissextile valide"
    Verifier Not IsISODate("2026-02-29"), "29 fevrier non bissextile invalide"
    Verifier Not IsISODate("2026-13-01"), "mois 13 invalide"
    Verifier Not IsISODate("12/03/2014"), "format non ISO refuse"
    VerifierEgal ISOAddDays("2026-12-31", 1), "2027-01-01", "ajout de jours"
    VerifierEgal NormalizeText("  Aïcha-Marie  d'Arc "), "AICHA MARIE D ARC", "normalisation"
    VerifierEgal Format_Numero("{ETAB}-{AAAA}-{N5}", "LYC", "2026", 42), "LYC-2026-00042", "format de numero"
    VerifierEgal RSA_TestModHex("0123456789abcdef0123456789abcdef0123456789abcdeffedcba9876543210", "fedcba98765432100f1e2d3c"), "e22216624b8e7d3dca8f5f34", "reduction modulaire grands entiers"
End Sub

Private Sub T_CoucheDonnees()
    Dim id As String, r As Variant, c As Collection
    id = DAL_NewId("T_Responsables")
    Verifier id <> DAL_NewId("T_Responsables"), "identifiants uniques"
    r = RecFrom("RespID=" & id & "|EtabID=X|Telephone=0102030405|Actif=1")
    RecSet r, "Nom", "=1+1"
    RecSet r, "Prenoms", "007"
    DAL_Insert "T_Responsables", r
    VerifierEgal DAL_Value("T_Responsables", id, "Nom"), "=1+1", "une formule est stockee comme texte"
    VerifierEgal DAL_Value("T_Responsables", id, "Prenoms"), "007", "les zeros de tete sont conserves"
    VerifierEgal DAL_Value("T_Responsables", id, "Telephone"), "0102030405", "telephone conserve en texte"
    On Error Resume Next
    DAL_Insert "T_Responsables", r
    VerifierErreur Err.Number, Err.Description, SG_ERR_DOUBLON, "identifiant en double refuse"
    DAL_Insert "T_Responsables", RecFrom("Nom=Sans identifiant")
    VerifierErreur Err.Number, Err.Description, SG_ERR_VALIDATION, "insertion sans identifiant refusee"
    DAL_Insert "T_Responsables", RecFrom("RespID=Z1|ColonneInconnue=1")
    VerifierErreur Err.Number, Err.Description, SG_ERR_INTROUVABLE, "colonne inconnue refusee"
    VerifierEgal DAL_Exists("T_Responsables", "Z1"), False, "insertion invalide : rien n'est ecrit"
    DAL_Update "T_Responsables", id, "RespID", "AUTRE"
    VerifierErreur Err.Number, Err.Description, SG_ERR_VALIDATION, "identifiant technique non modifiable"
    On Error GoTo 0
    Set c = DAL_Query("T_Responsables", RecFrom("EtabID=x"))
    VerifierEgal c.Count, 1, "requete insensible a la casse"
End Sub

Private Sub T_Securite()
    Dim uid As String, i As Long
    On Error Resume Next
    Sec_InitialiserAdministrateur "admin2", "Autre", MDP
    VerifierErreur Err.Number, Err.Description, SG_ERR_ETAT, "second administrateur initial refuse"
    Sec_CreerUtilisateur "faible", "Mot de passe faible", "abc"
    VerifierErreur Err.Number, Err.Description, SG_ERR_VALIDATION, "mot de passe trop court refuse"
    Sec_CreerUtilisateur "faible", "Mot de passe faible", "abcdefghij"
    VerifierErreur Err.Number, Err.Description, SG_ERR_VALIDATION, "mot de passe sans chiffre refuse"
    On Error GoTo 0
    uid = Sec_CreerUtilisateur("Secr1", "Secretaire fictive", MDP)
    Verifier Len(DAL_First("T_Utilisateurs", RecFrom("Login=secr1"))) > 0, "identifiant de connexion en minuscules"
    Verifier InStr(DAL_Value("T_Utilisateurs", uid, "HashMdp"), MDP) = 0 And Len(DAL_Value("T_Utilisateurs", uid, "HashMdp")) = 64, "mot de passe stocke sous forme d'empreinte"
    Verifier DAL_Count("T_Audit", RecFrom("EnregistrementID=" & uid & "|Operation=CREATION")) = 1, "creation de compte journalisee"
    Verifier InStr(DAL_Value("T_Audit", DAL_First("T_Audit", RecFrom("EnregistrementID=" & uid & "|Operation=CREATION")), "NouvelleValeur"), DAL_Value("T_Utilisateurs", uid, "Sel")) = 0, "le sel n'est pas journalise"
    Sec_Deconnexion
    On Error Resume Next
    Sec_Connexion "secr1", "Mauvais2026"
    VerifierErreur Err.Number, Err.Description, SG_ERR_ACCES, "mauvais mot de passe refuse"
    Sec_Connexion "inconnu", MDP
    VerifierErreur Err.Number, Err.Description, SG_ERR_ACCES, "compte inconnu refuse"
    On Error GoTo 0
    Verifier Not Sec_EstConnecte(), "aucune session apres echec"
    Verifier DAL_Count("T_Audit", RecFrom("Operation=CONNEXION|Resultat=ECHEC")) >= 2, "echecs de connexion journalises"
    On Error Resume Next
    For i = 1 To 5
        Sec_Connexion "secr1", "Mauvais2026"
    Next i
    Err.Clear
    Sec_Connexion "secr1", MDP
    VerifierErreur Err.Number, Err.Description, SG_ERR_ACCES, "compte verrouille apres 5 echecs (meme avec le bon mot de passe)"
    On Error GoTo 0
    Verifier DAL_Count("T_Audit", RecFrom("Operation=VERROUILLAGE|EnregistrementID=" & uid)) >= 1, "verrouillage journalise"
    Connecter "admin"
    Sec_ReinitialiserMotDePasse uid, "Temporaire99", "Demande de la secretaire (oubli)"
    Sec_Connexion "secr1", "Temporaire99"
    Verifier Sec_EstConnecte(), "connexion apres reinitialisation par un administrateur"
    VerifierEgal DAL_Value("T_Utilisateurs", uid, "DoitChangerMdp"), "1", "changement de mot de passe exige"
    Sec_ChangerMotDePasse "Temporaire99", "Nouveau2026x"
    Sec_Deconnexion
    Sec_Connexion "secr1", "Nouveau2026x"
    Verifier Sec_EstConnecte(), "connexion avec le nouveau mot de passe"
    On Error Resume Next
    Sec_CreerUtilisateur "pirate", "Creation non autorisee", MDP
    VerifierErreur Err.Number, Err.Description, SG_ERR_ACCES, "un utilisateur ordinaire ne cree pas de compte"
    On Error GoTo 0
    Connecter "admin"
    Sec_DesactiverUtilisateur uid, "Depart de l'etablissement"
    Sec_Deconnexion
    On Error Resume Next
    Sec_Connexion "secr1", "Nouveau2026x"
    VerifierErreur Err.Number, Err.Description, SG_ERR_ACCES, "compte desactive refuse"
    On Error GoTo 0
    Verifier DAL_Exists("T_Utilisateurs", uid), "le compte desactive est conserve (historique)"
End Sub

Private Sub T_Permissions()
    Dim cx As Variant, e As String, el As String, uS As String, uE As String, uT As String, hab As String
    cx = Contexte("ETB-A")
    e = cx(0)
    uS = CreerUtilisateur("secr", e, "SECRETAIRE")
    uE = CreerUtilisateur("prof", e, "ENSEIGNANT")
    uT = CreerUtilisateur("tech", "", "")
    DAL_Update "T_Utilisateurs", uT, "AdminTech", "1"
    Connecter "secr"
    el = NouvelEleve(e, "KOUADIO", "Awa")
    Verifier Len(el) > 0, "la secretaire cree un dossier"
    On Error Resume Next
    Sauv_SimulerPourTests e
    VerifierErreur Err.Number, Err.Description, SG_ERR_ACCES, "mode centralise : la secretaire ne fait pas de sauvegarde globale"
    Annee_Cloturer e, cx(1), "Test"
    VerifierErreur Err.Number, Err.Description, SG_ERR_ACCES, "la secretaire ne cloture pas l'annee"
    Cycle_Creer e, "LYC", "Lycee", "NOTES", 4
    VerifierErreur Err.Number, Err.Description, SG_ERR_ACCES, "la secretaire ne modifie pas la structure"
    On Error GoTo 0
    Verifier DAL_Count("T_Audit", RecFrom("Operation=ACCES_REFUSE|EtabID=" & e)) >= 2, "refus d'acces journalises"
    Connecter "prof"
    VerifierEgal RecStr(Eleve_Get(e, el), "Nom"), "KOUADIO", "l'enseignant consulte un dossier"
    On Error Resume Next
    NouvelEleve e, "TRAORE", "Moussa"
    VerifierErreur Err.Number, Err.Description, SG_ERR_ACCES, "l'enseignant ne cree pas de dossier"
    Eleve_Modifier e, el, "Telephone", "0700000000", ""
    VerifierErreur Err.Number, Err.Description, SG_ERR_ACCES, "l'enseignant ne modifie pas de dossier"
    Audit_Historique e, el
    VerifierErreur Err.Number, Err.Description, SG_ERR_ACCES, "l'enseignant ne consulte pas le journal"
    On Error GoTo 0
    Connecter "tech"
    VerifierEgal Sec_Niveau(e, "ELEVES"), NIV_AUCUN, "l'administrateur technique n'a pas acces aux dossiers eleves"
    On Error Resume Next
    Eleve_Rechercher e, ""
    VerifierErreur Err.Number, Err.Description, SG_ERR_ACCES, "recherche d'eleves refusee a l'administrateur technique"
    On Error GoTo 0
    Connecter "admin"
    Sec_ModifierPermission "ENSEIGNANT", "ELEVES", NIV_AUCUN, "Politique de confidentialite"
    Verifier DAL_Count("T_Audit", RecFrom("Operation=MODIFICATION_DROITS")) = 1, "modification des droits journalisee"
    Connecter "prof"
    On Error Resume Next
    Eleve_Get e, el
    VerifierErreur Err.Number, Err.Description, SG_ERR_ACCES, "droit retire : consultation refusee"
    On Error GoTo 0
    Connecter "admin"
    hab = DAL_First("T_Habilitations", RecFrom("UserID=" & uS & "|EtabID=" & e))
    Sec_RetirerHabilitation hab, "Mutation"
    Connecter "secr"
    Verifier Not Sec_AAccesEtab(e), "habilitation retiree : plus aucun acces a l'etablissement"
    Verifier DAL_Exists("T_Habilitations", hab), "l'habilitation retiree est conservee"
End Sub

Private Sub T_Isolation()
    Dim a As Variant, b As Variant, ea As String, eb As String, elA As String, elB As String, res As Collection
    Dim chemin As String, contenu As String
    a = Contexte("ETB-A")
    b = Contexte("ETB-B")
    ea = a(0): eb = b(0)
    CreerUtilisateur "secra", ea, "SECRETAIRE"
    CreerUtilisateur "multi", ea, "SECRETAIRE"
    Sec_Habiliter DAL_First("T_Utilisateurs", RecFrom("Login=multi")), eb, "SECRETAIRE", "2026-01-01"
    Connecter "secra"
    elA = NouvelEleve(ea, "DIALLO", "Fatou")
    Connecter "multi"
    elB = NouvelEleve(eb, "DIALLO", "Fatou")
    Verifier elA <> elB, "meme identite dans deux etablissements = deux dossiers distincts (pas de faux doublon)"
    Connecter "secra"
    On Error Resume Next
    Eleve_Rechercher eb, ""
    VerifierErreur Err.Number, Err.Description, SG_ERR_ACCES, "recherche dans un autre etablissement refusee"
    Eleve_Get ea, elB
    VerifierErreur Err.Number, Err.Description, SG_ERR_ISOLATION, "lecture d'un dossier d'un autre etablissement via son identifiant refusee"
    Eleve_Modifier ea, elB, "Telephone", "000", ""
    VerifierErreur Err.Number, Err.Description, SG_ERR_ISOLATION, "modification inter-etablissements refusee"
    Inscr_Inscrire ea, elA, a(1), b(4), "PREMIERE"
    VerifierErreur Err.Number, Err.Description, SG_ERR_ISOLATION, "inscription dans une classe d'un autre etablissement refusee"
    Inscr_Inscrire ea, elA, b(1), a(4), "PREMIERE"
    VerifierErreur Err.Number, Err.Description, SG_ERR_ISOLATION, "inscription dans l'annee d'un autre etablissement refusee"
    On Error GoTo 0
    Set res = Eleve_Rechercher(ea, "diallo")
    VerifierEgal res.Count, 1, "la recherche ne retourne que les dossiers de l'etablissement"
    If res.Count = 1 Then VerifierEgal res(1), elA, "dossier de l'etablissement A"
    Verifier DAL_Count("T_Audit", RecFrom("Operation=TENTATIVE_INTER_ETAB")) >= 3, "tentatives inter-etablissements journalisees"
    Connecter "multi"
    VerifierEgal Eleve_Rechercher(ea, "").Count, 1, "utilisateur multi-etablissements : A separe"
    VerifierEgal Eleve_Rechercher(eb, "").Count, 1, "utilisateur multi-etablissements : B separe"
    VerifierEgal Etab_Accessibles().Count, 2, "etablissements accessibles"
    Connecter "admin"
    chemin = DossierTemp() & "sg_test_export_A.csv"
    VerifierEgal Export_TableCSV(ea, "T_Eleves", chemin), 1, "export cloisonne : une ligne"
    contenu = LireFichierUtf8(chemin)
    Verifier InStr(contenu, DAL_Value("T_Eleves", elA, "Matricule")) > 0, "l'export contient le dossier de A"
    Verifier InStr(contenu, DAL_Value("T_Eleves", elB, "Matricule")) = 0, "l'export ne contient aucun dossier de B"
    Verifier DAL_Count("T_Audit", RecFrom("Operation=EXPORT|EtabID=" & ea)) = 1, "export journalise"
End Sub

Private Function DossierTemp() As String
    Dim d As String
    d = Environ$("TEMP")
    If Len(d) = 0 Then d = Environ$("TMPDIR")
    If Len(d) = 0 Then d = "/tmp"
    DossierTemp = PathJoin(d, "")
End Function

Private Sub T_Annees()
    Dim e As String, a1 As String, a2 As String, cyc As String, n As String, cls As String, el As String
    e = CreerEtab("ETB-A")
    Sec_Habiliter Sec_UtilisateurCourant(), e, "DIRECTEUR", "2026-01-01"
    a1 = Annee_Creer(e, "2026-2027", "", "2026-09-01", "2027-07-31")
    VerifierEgal Annee_Statut(a1), "PREPARATION", "nouvelle annee en preparation"
    On Error Resume Next
    Annee_Creer e, "2026-2027", "", "2026-09-01", "2027-07-31"
    VerifierErreur Err.Number, Err.Description, SG_ERR_DOUBLON, "annee en double refusee"
    Annee_Creer e, "2027-X", "", "2027-07-01", "2028-06-30"
    VerifierErreur Err.Number, Err.Description, SG_ERR_VALIDATION, "chevauchement de dates refuse"
    Annee_Creer e, "2027-2028", "", "2028-07-31", "2027-09-01"
    VerifierErreur Err.Number, Err.Description, SG_ERR_VALIDATION, "fin avant debut refusee"
    On Error GoTo 0
    Annee_Ouvrir e, a1
    a2 = Annee_Creer(e, "2027-2028", "", "2027-09-01", "2028-07-31")
    On Error Resume Next
    Annee_Ouvrir e, a2
    VerifierErreur Err.Number, Err.Description, SG_ERR_ETAT, "une seule annee ouverte a la fois"
    Annee_Cloturer e, a1, "Fin d'annee"
    VerifierErreur Err.Number, Err.Description, SG_ERR_ETAT, "cloture refusee sans sauvegarde verifiee"
    On Error GoTo 0
    cyc = Cycle_Creer(e, "PRI", "Primaire", "COMPETENCES", 2)
    n = Niveau_Creer(e, cyc, "CP", "Cours preparatoire", 1)
    cls = Classe_Creer(e, a1, n, "", "CPA", "CP A", 25)
    el = NouvelEleve(e, "BAMBA", "Ines", "2019-05-01")
    Inscr_Inscrire e, el, a1, cls, "PREMIERE", "2026-09-01"
    Sauv_SimulerPourTests e
    Annee_Cloturer e, a1, "Fin d'annee scolaire"
    VerifierEgal Annee_Statut(a1), "CLOTUREE", "annee cloturee"
    VerifierEgal DAL_Count("T_Inscriptions", RecFrom("AnneeID=" & a1)), 1, "les inscriptions de l'annee cloturee sont conservees"
    On Error Resume Next
    Classe_Creer e, a1, n, "", "CPB", "CP B", 25
    VerifierErreur Err.Number, Err.Description, SG_ERR_ETAT, "creation de classe refusee dans une annee cloturee"
    Inscr_ChangerClasse e, DAL_First("T_Inscriptions", RecFrom("AnneeID=" & a1)), cls, "2027-01-10", "x"
    VerifierErreur Err.Number, Err.Description, SG_ERR_ETAT, "modification d'inscription refusee dans une annee cloturee"
    Annee_Reouvrir e, a1, "court"
    VerifierErreur Err.Number, Err.Description, SG_ERR_VALIDATION, "reouverture sans motif detaille refusee"
    On Error GoTo 0
    Annee_Ouvrir e, a2
    On Error Resume Next
    Annee_Reouvrir e, a1, "Correction d'une erreur de saisie signalee par la direction"
    VerifierErreur Err.Number, Err.Description, SG_ERR_ETAT, "reouverture refusee tant qu'une autre annee est ouverte"
    On Error GoTo 0
    Verifier DAL_Count("T_Audit", RecFrom("Operation=CLOTURE|EnregistrementID=" & a1)) = 1, "cloture journalisee"
    Verifier DAL_Count("T_Annees", RecFrom("EtabID=" & e)) = 2, "les deux annees coexistent"
End Sub

Private Sub T_Structure()
    Dim cx As Variant, e As String, a2 As String, n As Long, ids As Collection
    cx = Contexte("ETB-A")
    e = cx(0)
    On Error Resume Next
    Cycle_Creer e, "COL", "Doublon", "NOTES", 9
    VerifierErreur Err.Number, Err.Description, SG_ERR_DOUBLON, "code de cycle en double refuse"
    Cycle_Creer e, "XXX", "Mode inconnu", "LETTRES", 9
    VerifierErreur Err.Number, Err.Description, SG_ERR_VALIDATION, "mode d'evaluation inconnu refuse"
    Classe_Creer e, cx(1), cx(2), "", "6EA", "Doublon", 30
    VerifierErreur Err.Number, Err.Description, SG_ERR_DOUBLON, "classe en double refusee"
    Niveau_Creer e, DAL_First("T_Cycles", RecFrom("EtabID=" & e)), "4E", "Ordre deja pris", 1
    VerifierErreur Err.Number, Err.Description, SG_ERR_DOUBLON, "ordre de niveau unique"
    On Error GoTo 0
    ' Complexe scolaire : un cycle maternelle par competences cohabite avec le college.
    Niveau_Creer e, Cycle_Creer(e, "MAT", "Maternelle", "COMPETENCES", 1), "PS", "Petite section", 0
    Filiere_Creer e, "TI", "Technicien en informatique", "FILIERE"
    VerifierEgal DAL_Count("T_Cycles", RecFrom("EtabID=" & e)), 2, "plusieurs cycles dans un meme etablissement"
    a2 = Annee_Creer(e, "2027-2028", "", "2027-09-01", "2028-07-31")
    n = Classe_ReconduireAnnee(e, cx(1), a2)
    VerifierEgal n, 3, "reconduction : 3 classes creees"
    VerifierEgal DAL_Count("T_Classes", RecFrom("AnneeID=" & cx(1))), 3, "les classes de l'annee source sont intactes"
    Set ids = DAL_Query("T_Classes", RecFrom("AnneeID=" & a2 & "|Code=6EA"))
    Verifier ids.Count = 1, "classe reconduite retrouvee"
    If ids.Count = 1 Then Verifier ids(1) <> cx(4), "nouvel identifiant pour la classe reconduite"
    VerifierEgal Classe_ReconduireAnnee(e, cx(1), a2), 0, "reconduction repetee sans doublon"
End Sub

Private Sub T_Eleves()
    Dim cx As Variant, e As String, el As String, r As Variant, ev As Variant, el2 As String, rsp As String, h As Collection
    cx = Contexte("ETB-A")
    e = cx(0)
    el = NouvelEleve(e, "Koné", "Aïcha Marie")
    ev = DAL_Get("T_Eleves", el)
    VerifierEgal RecStr(ev, "Matricule"), "ETB-A-" & Format(Date, "yyyy") & "-00001", "matricule genere selon le format parametre"
    VerifierEgal RecStr(ev, "Nom"), "KONÉ", "nom en majuscules"
    VerifierEgal RecStr(ev, "Statut"), "NOUVEAU", "statut initial"
    VerifierEgal NbAudit("CREATION", el), 1, "creation journalisee"
    Verifier InStr(DAL_Value("T_Audit", DAL_First("T_Audit", RecFrom("Operation=CREATION|EnregistrementID=" & el)), "NouvelleValeur"), "Prenoms=Aïcha Marie") > 0, "valeurs initiales journalisees"
    On Error Resume Next
    NouvelEleve e, "KONE", "Aicha-Marie"
    VerifierErreur Err.Number, Err.Description, SG_ERR_DOUBLON, "doublon detecte malgre accents et tirets"
    r = RecFrom("Nom=KONE|Prenoms=Aicha Marie|Sexe=F|DateNaissance=2014-03-12")
    Eleve_Creer e, r, True, ""
    VerifierErreur Err.Number, Err.Description, SG_ERR_VALIDATION, "homonyme : motif obligatoire"
    NouvelEleve e, "SANS", "Date", "2014-02-30"
    VerifierErreur Err.Number, Err.Description, SG_ERR_VALIDATION, "date de naissance impossible refusee"
    NouvelEleve e, "FUTUR", "Bebe", "2030-01-01"
    VerifierErreur Err.Number, Err.Description, SG_ERR_VALIDATION, "date de naissance future refusee"
    NouvelEleve e, "", "Sans nom"
    VerifierErreur Err.Number, Err.Description, SG_ERR_VALIDATION, "nom obligatoire"
    NouvelEleve e, "SEXE", "Inconnu", "2014-01-01", "X"
    VerifierErreur Err.Number, Err.Description, SG_ERR_VALIDATION, "sexe invalide refuse"
    On Error GoTo 0
    el2 = Eleve_Creer(e, r, True, "Jumelle homonyme verifiee sur acte de naissance")
    Verifier DAL_Value("T_Eleves", el2, "Matricule") <> DAL_Value("T_Eleves", el, "Matricule"), "homonyme confirme : matricule distinct"
    On Error Resume Next
    Eleve_Modifier e, el, "Nom", "KONE-TRAORE", ""
    VerifierErreur Err.Number, Err.Description, SG_ERR_VALIDATION, "motif obligatoire pour modifier le nom"
    Eleve_Modifier e, el, "Matricule", "X", "tentative"
    VerifierErreur Err.Number, Err.Description, SG_ERR_VALIDATION, "matricule non modifiable"
    Eleve_Modifier e, el, "Statut", "ACTIF", "tentative"
    VerifierErreur Err.Number, Err.Description, SG_ERR_VALIDATION, "statut non modifiable directement"
    On Error GoTo 0
    Eleve_Modifier e, el, "Nom", "KONE-TRAORE", "Jugement suppletif fictif n° 12"
    Eleve_Modifier e, el, "Telephone", "0700000001", ""
    VerifierEgal DAL_Value("T_Eleves", el, "Nom"), "KONE-TRAORE", "nom modifie"
    VerifierEgal DAL_Value("T_Eleves", el, "Version"), "3", "version incrementee"
    Set h = Audit_Historique(e, el)
    VerifierEgal h.Count, 3, "historique : creation + 2 modifications"
    ev = DAL_Get("T_Audit", DAL_First("T_Audit", RecFrom("Operation=MODIFICATION|EnregistrementID=" & el & "|Champ=Nom")))
    VerifierEgal RecStr(ev, "AncienneValeur"), "KONÉ", "ancienne valeur conservee"
    VerifierEgal RecStr(ev, "NouvelleValeur"), "KONE-TRAORE", "nouvelle valeur journalisee"
    VerifierEgal RecStr(ev, "Motif"), "Jugement suppletif fictif n° 12", "motif journalise"
    VerifierEgal RecStr(ev, "UserID"), Sec_UtilisateurCourant(), "auteur journalise"
    VerifierEgal Eleve_Rechercher(e, "aicha traore").Count, 1, "recherche multicritere sans accents"
    VerifierEgal Eleve_Rechercher(e, DAL_Value("T_Eleves", el2, "Matricule")).Count, 1, "recherche par matricule"
    rsp = Resp_Creer(e, RecFrom("Nom=Kone|Prenoms=Ibrahim|Telephone=0700000002|Profession=Commercant"))
    Resp_Lier e, el, rsp, "Pere", "LEGAL;FINANCIER;URGENCE"
    Resp_Lier e, el2, rsp, "Pere", "LEGAL;FINANCIER"
    VerifierEgal Resp_Enfants(e, rsp).Count, 2, "un responsable lie a plusieurs enfants"
    VerifierEgal DAL_Value("T_LiensResponsables", DAL_First("T_LiensResponsables", RecFrom("EleveID=" & el2)), "Urgence"), "0", "roles du responsable distincts par enfant"
    On Error Resume Next
    Resp_Lier e, el, rsp, "Pere", "LEGAL"
    VerifierErreur Err.Number, Err.Description, SG_ERR_DOUBLON, "lien en double refuse"
    On Error GoTo 0
End Sub

Private Sub T_Inscriptions()
    Dim cx As Variant, e As String, el(1 To 4) As String, ins As String, h As Collection, i As Long, a1 As Variant, a2 As Variant
    cx = Contexte("ETB-A", 3)
    e = cx(0)
    For i = 1 To 4
        el(i) = NouvelEleve(e, "ELEVE" & Chr$(64 + i), "Test", "2014-0" & i & "-10")
    Next i
    ins = Inscr_Inscrire(e, el(1), cx(1), cx(4), "PREMIERE", "2026-09-01")
    VerifierEgal DAL_Value("T_Eleves", el(1), "Statut"), "ACTIF", "eleve actif apres inscription"
    VerifierEgal DAL_Count("T_Affectations", RecFrom("InscriptionID=" & ins)), 1, "affectation initiale creee"
    On Error Resume Next
    Inscr_Inscrire e, el(1), cx(1), cx(5), "PREMIERE"
    VerifierErreur Err.Number, Err.Description, SG_ERR_DOUBLON, "double inscription active refusee"
    Inscr_Inscrire e, el(2), cx(1), cx(4), "REINSCRIPTION"
    VerifierErreur Err.Number, Err.Description, SG_ERR_ETAT, "reinscription sans inscription anterieure refusee"
    Inscr_Inscrire e, el(2), cx(1), cx(4), "PREMIERE", "2027-08-15"
    VerifierErreur Err.Number, Err.Description, SG_ERR_VALIDATION, "date d'effet hors annee refusee"
    On Error GoTo 0
    Inscr_Inscrire e, el(2), cx(1), cx(4), "PREMIERE", "2026-09-01"
    Inscr_Inscrire e, el(3), cx(1), cx(4), "TRANSFERT_ENTRANT", "2026-09-15"
    On Error Resume Next
    Inscr_Inscrire e, el(4), cx(1), cx(4), "PREMIERE", "2026-09-20"
    VerifierErreur Err.Number, Err.Description, SG_ERR_ETAT, "capacite de classe respectee"
    Inscr_ChangerClasse e, ins, cx(5), "2026-11-02", ""
    VerifierErreur Err.Number, Err.Description, SG_ERR_VALIDATION, "motif obligatoire pour un changement de classe"
    Inscr_ChangerClasse e, ins, cx(5), "2026-08-01", "Equilibrage"
    VerifierErreur Err.Number, Err.Description, SG_ERR_VALIDATION, "date d'effet anterieure refusee"
    On Error GoTo 0
    Inscr_ChangerClasse e, ins, cx(5), "2026-11-02", "Equilibrage des effectifs"
    Set h = Inscr_Historique(e, ins)
    VerifierEgal h.Count, 2, "historique : deux affectations"
    a1 = DAL_Get("T_Affectations", h(1))
    a2 = DAL_Get("T_Affectations", h(2))
    VerifierEgal RecStr(a1, "ClasseID"), cx(4), "ancienne affectation conservee"
    VerifierEgal RecStr(a1, "DateFin"), "2026-11-01", "ancienne affectation close la veille de la date d'effet"
    VerifierEgal RecStr(a2, "DateDebut"), "2026-11-02", "date d'effet distincte"
    Verifier Left$(RecStr(a2, "DateOperation"), 10) = TodayISO(), "date d'operation enregistree"
    VerifierEgal DAL_Value("T_Inscriptions", ins, "ClasseID"), cx(5), "classe courante mise a jour"
    Inscr_Inscrire e, el(4), cx(1), cx(4), "PREMIERE", "2026-11-03"
    VerifierEgal Classe_Effectif(cx(4)), 3, "place liberee reutilisee"
    ' Fin d'inscription
    CreerUtilisateur "secr", e, "SECRETAIRE"
    Connecter "secr"
    On Error Resume Next
    Inscr_Terminer e, Inscr_Active(e, el(2), cx(1)), "RADIATION", "2027-01-15", "Discipline"
    VerifierErreur Err.Number, Err.Description, SG_ERR_ACCES, "radiation reservee au niveau validation"
    On Error GoTo 0
    Inscr_Terminer e, Inscr_Active(e, el(2), cx(1)), "DEPART", "2027-01-15", "Demenagement de la famille"
    VerifierEgal DAL_Value("T_Eleves", el(2), "Statut"), "PARTI", "statut PARTI"
    Verifier DAL_Exists("T_Eleves", el(2)), "le dossier d'un eleve parti est conserve"
    VerifierEgal DAL_Count("T_Inscriptions", RecFrom("EleveID=" & el(2) & "|Statut=TERMINEE")), 1, "inscription terminee conservee"
    Connecter "admin"
    Inscr_Suspendre e, Inscr_Active(e, el(3), cx(1)), "Mesure conservatoire"
    VerifierEgal DAL_Value("T_Eleves", el(3), "Statut"), "SUSPENDU", "suspension"
    Inscr_Reintegrer e, el(3), "Fin de la mesure"
    VerifierEgal DAL_Value("T_Eleves", el(3), "Statut"), "ACTIF", "reintegration"
    Inscr_Annuler e, Inscr_Active(e, el(4), cx(1)), "Saisie sur le mauvais dossier"
    VerifierEgal DAL_Count("T_Inscriptions", RecFrom("EleveID=" & el(4) & "|Statut=ANNULEE")), 1, "inscription annulee conservee avec statut ANNULEE"
    VerifierEgal DAL_Value("T_Eleves", el(4), "Statut"), "NOUVEAU", "statut eleve apres annulation"
End Sub

Private Sub T_Reinscription()
    Dim cx As Variant, e As String, a2 As String, elP As String, elR As String, insAncien As String, insP As String, insR As String
    Dim c5 As String, c6 As String, avant As String
    cx = Contexte("ETB-A")
    e = cx(0)
    elP = NouvelEleve(e, "PASSAGE", "Paul", "2014-01-01", "M")
    elR = NouvelEleve(e, "REDOUBLE", "Rita", "2014-02-02")
    insAncien = Inscr_Inscrire(e, elP, cx(1), cx(4), "PREMIERE", "2026-09-01")
    Inscr_Inscrire e, elR, cx(1), cx(4), "PREMIERE", "2026-09-01"
    avant = DAL_Value("T_Inscriptions", insAncien, "ClasseID") & "|" & DAL_Value("T_Inscriptions", insAncien, "AnneeID")
    a2 = Annee_Creer(e, "2027-2028", "", "2027-09-01", "2028-07-31")
    Classe_ReconduireAnnee e, cx(1), a2
    c5 = DAL_First("T_Classes", RecFrom("AnneeID=" & a2 & "|Code=5EA"))
    c6 = DAL_First("T_Classes", RecFrom("AnneeID=" & a2 & "|Code=6EA"))
    insP = Inscr_Inscrire(e, elP, a2, c5, "REINSCRIPTION")
    insR = Inscr_Inscrire(e, elR, a2, c6, "REINSCRIPTION")
    VerifierEgal DAL_Value("T_Inscriptions", insP, "Decision"), "PASSAGE", "passage au niveau superieur deduit"
    VerifierEgal DAL_Value("T_Inscriptions", insR, "Decision"), "REDOUBLEMENT", "redoublement deduit"
    VerifierEgal DAL_Value("T_Inscriptions", insP, "DateEffet"), "2027-09-01", "pre-inscription : effet au debut de la nouvelle annee"
    Verifier insP <> insAncien, "nouvel enregistrement pour la nouvelle annee"
    VerifierEgal DAL_Value("T_Inscriptions", insAncien, "ClasseID") & "|" & DAL_Value("T_Inscriptions", insAncien, "AnneeID"), avant, "inscription de l'annee precedente inchangee"
    VerifierEgal DAL_Count("T_Inscriptions", RecFrom("EleveID=" & elP)), 2, "historique pluriannuel"
End Sub

Private Sub T_Transfert()
    Dim a As Variant, b As Variant, el As String, ins As String, nouvel As String, t As String
    a = Contexte("ETB-A")
    b = Contexte("ETB-B")
    el = NouvelEleve(a(0), "MOBILE", "Marc", "2013-06-06", "M")
    ins = Inscr_Inscrire(a(0), el, a(1), a(4), "PREMIERE", "2026-09-01")
    nouvel = Transfert_VersEtablissement(a(0), ins, b(0), "2027-02-01", "Mutation des parents")
    VerifierEgal DAL_Value("T_Eleves", el, "Statut"), "TRANSFERE", "dossier source au statut TRANSFERE"
    VerifierEgal DAL_Value("T_Inscriptions", ins, "Statut"), "TERMINEE", "inscription source terminee"
    VerifierEgal DAL_Value("T_Eleves", nouvel, "EtabID"), b(0), "nouveau dossier dans l'etablissement d'accueil"
    VerifierEgal DAL_Value("T_Eleves", nouvel, "EleveOrigineID"), el, "lien vers le dossier d'origine"
    VerifierEgal DAL_Value("T_Eleves", nouvel, "EtabPrecedent"), "ETB-A", "etablissement precedent renseigne"
    t = DAL_First("T_Transferts", RecFrom("EleveSource=" & el))
    Verifier Len(t) > 0, "transfert enregistre"
    VerifierEgal DAL_Count("T_Inscriptions", RecFrom("EtabID=" & b(0) & "|EleveID=" & nouvel)), 0, "l'historique d'inscription n'est pas copie dans l'autre etablissement"
    Inscr_Inscrire b(0), nouvel, b(1), b(4), "TRANSFERT_ENTRANT", "2027-02-01"
    VerifierEgal DAL_Value("T_Eleves", nouvel, "Statut"), "ACTIF", "inscription dans l'etablissement d'accueil"
    Verifier Len(DAL_Value("T_Audit", DAL_First("T_Audit", RecFrom("Operation=TRANSFERT_SORTANT|EnregistrementID=" & ins)), "TxRef")) > 0, "operations du transfert regroupees sous une reference"
End Sub

Private Sub T_Audit()
    Dim cx As Variant, e As String, el As String, r As String, evt As String, ancre As String
    cx = Contexte("ETB-A")
    e = cx(0)
    el = NouvelEleve(e, "AUDIT", "Alice")
    Eleve_Modifier e, el, "Telephone", "0700000003", ""
    r = Audit_Verifier()
    VerifierEgal Left$(r, 3), "OK|", "chaine d'audit integre"
    ancre = Audit_HashTete()
    VerifierEgal Left$(Audit_Verifier(ancre), 3), "OK|", "ancre externe retrouvee"
    VerifierEgal Left$(Audit_Verifier(String$(64, "f")), 6), "ALTERE", "ancre inconnue detectee (journal reconstruit)"
    evt = DAL_First("T_Audit", RecFrom("Operation=MODIFICATION|EnregistrementID=" & el))
    DAL_FalsifierPourTests "T_Audit", evt, "NouvelleValeur", "0799999999"
    r = Audit_Verifier()
    Verifier Left$(r, 6) = "ALTERE" And InStr(r, evt) > 0, "modification d'un evenement detectee : " & r
    DAL_FalsifierPourTests "T_Audit", evt, "NouvelleValeur", "0700000003"
    VerifierEgal Left$(Audit_Verifier(), 3), "OK|", "chaine retablie apres restauration de la valeur"
    DAL_FalsifierPourTests "T_Audit", evt, "HashPrec", String$(64, "0")
    Verifier Left$(Audit_Verifier(), 6) = "ALTERE", "rupture du chainage detectee"
End Sub

Private Sub T_Sauvegarde()
    Dim cx As Variant, e As String, s As String, ancre As String
    cx = Contexte("ETB-A")
    e = cx(0)
    Verifier Not Sauv_ExisteVerifieeDuJour(e), "aucune sauvegarde au depart"
    s = Sauv_SimulerPourTests(e)
    Verifier Sauv_ExisteVerifieeDuJour(e), "sauvegarde verifiee du jour enregistree"
    VerifierEgal DAL_Value("T_Sauvegardes", s, "HashAudit"), DAL_Value("T_Audit", DAL_First("T_Audit", RecFrom("Operation=SAUVEGARDE|EnregistrementID=" & s)), "HashPrec"), "empreinte du journal memorisee avec la sauvegarde"
    ancre = Sauv_AncreAudit()
    Verifier InStr(ancre, "empreinte=") > 0, "ancre exportable"
    VerifierEgal Export_EtablissementCSV(e, Left$(DossierTemp(), Len(DossierTemp()) - 1)) >= 8, True, "export complet de l'etablissement"
End Sub

Private Sub T_Rapports()
    Dim cx As Variant, e As String, a As String, z As String, b As String, c As String, ind As Variant, tab As Variant
    Dim qc As Collection, i As Long, trouve As Boolean, cert As Variant, cert2 As Variant, lst As Variant, nb As Long
    cx = Contexte("ETB-A")
    e = cx(0): a = cx(1)
    z = NouvelEleve(e, "ZOUMA", "Zoe")
    b = NouvelEleve(e, "ADJOVI", "Bruno", "2014-05-05", "M")
    c = NouvelEleve(e, "MENSAH", "Carla")
    Inscr_Inscrire e, z, a, cx(4), "PREMIERE", "2026-09-01"
    Inscr_Inscrire e, b, a, cx(4), "PREMIERE", "2026-09-01"
    Inscr_Inscrire e, c, a, cx(5), "PREMIERE", "2026-09-01"
    Inscr_Terminer e, Inscr_Active(e, c, a), "DEPART", "2026-12-20", "Demenagement"
    ind = Rapport_Indicateurs(e, a)
    VerifierEgal RecGet(ind, "Inscrits"), 2, "indicateur : inscrits actifs"
    VerifierEgal RecGet(ind, "Premieres"), 3, "indicateur : premieres inscriptions"
    VerifierEgal RecGet(ind, "Departs"), 1, "indicateur : departs"
    VerifierEgal RecGet(ind, "Filles"), 1, "indicateur : filles actives"
    VerifierEgal RecGet(ind, "Garcons"), 1, "indicateur : garcons actifs"
    tab = Rapport_EffectifsClasses(e, a)
    VerifierEgal UBound(tab, 1), 3, "effectifs : une ligne par classe"
    VerifierEgal tab(1, 0) & "=" & tab(1, 3), "6EA=2", "effectif de la classe 6EA"
    lst = Liste_Classe(e, cx(4))
    VerifierEgal lst(1, 2) & "," & lst(2, 2), "ADJOVI,ZOUMA", "liste de classe triee par nom"
    VerifierEgal QC_Anomalies(e).Count, 2, "controle qualite : seuls les responsables manquants sont signales"
    DAL_FalsifierPourTests "T_Eleves", c, "Statut", "ACTIF"
    DAL_FalsifierPourTests "T_Classes", cx(4), "Capacite", "1"
    Set qc = QC_Anomalies(e)
    trouve = False
    For i = 1 To qc.Count
        If InStr(qc(i), "INSCRIPTION|" & c) > 0 Then trouve = True
    Next i
    Verifier trouve, "anomalie : eleve actif sans inscription"
    trouve = False
    For i = 1 To qc.Count
        If InStr(qc(i), "CAPACITE|" & cx(4)) > 0 Then trouve = True
    Next i
    Verifier trouve, "anomalie : sureffectif"
    VerifierEgal DAL_Value("T_Eleves", c, "Statut"), "ACTIF", "le controle qualite ne corrige rien automatiquement"
    cert = Doc_CertificatScolarite(e, z)
    cert2 = Doc_CertificatScolarite(e, b)
    VerifierEgal RecStr(cert, "Numero"), "CS-ETB-A-" & Year(Date) & "-00001", "certificat numerote"
    VerifierEgal RecStr(cert2, "Numero"), "CS-ETB-A-" & Year(Date) & "-00002", "numerotation sequentielle"
    VerifierEgal RecStr(cert, "Classe"), "6e A", "classe sur le certificat"
    VerifierEgal NbAudit("DOCUMENT_EMIS", z), 1, "emission du document journalisee"
    DAL_FalsifierPourTests "T_Eleves", c, "Statut", "PARTI"
    On Error Resume Next
    Doc_CertificatScolarite e, c
    VerifierErreur Err.Number, Err.Description, SG_ERR_ETAT, "pas de certificat pour un eleve non inscrit"
    On Error GoTo 0
End Sub

Private Sub T_Licences()
    Dim cx As Variant, e As String, eb As String, el As String, st As String, mo As String, fp As String, nbAvant As Long, est As Variant
    If Len(Fx("MODULUS")) = 0 Then
        Verifier False, "fixtures de licence absentes (lancer via run_vba_tests.py)"
        Exit Sub
    End If
    mo = Fx("MODULUS")
    fp = Fx("EMPREINTE")
    ' Etat sans licence : essai integre
    VerifierEgal Lic_StatutCourant(), "ESSAI_INTEGRE", "sans licence : essai integre"
    cx = Contexte("ETB-A")
    e = cx(0)
    eb = CreerEtab("ETB-B")
    Sec_Habiliter Sec_UtilisateurCourant(), eb, "DIRECTEUR", "2026-01-01"
    On Error Resume Next
    CreerEtab "ETB-C"
    VerifierErreur Err.Number, Err.Description, SG_ERR_LICENCE, "essai : 2 etablissements maximum"
    Lic_InstallerAvecCle Fx("LIC_FALSIFIEE"), mo, fp, D_REF
    VerifierErreur Err.Number, Err.Description, SG_ERR_LICENCE, "licence falsifiee (module ajoute) refusee"
    Lic_InstallerAvecCle Fx("LIC_AUTRE_CLE"), mo, fp, D_REF
    VerifierErreur Err.Number, Err.Description, SG_ERR_LICENCE, "licence signee par une autre cle refusee"
    Lic_InstallerAvecCle Fx("LIC_VALIDE"), mo, String$(64, "b"), D_REF
    VerifierErreur Err.Number, Err.Description, SG_ERR_LICENCE, "licence d'une autre installation refusee"
    Lic_InstallerAvecCle "texte quelconque", mo, fp, D_REF
    VerifierErreur Err.Number, Err.Description, SG_ERR_LICENCE, "fichier illisible refuse"
    On Error GoTo 0
    VerifierEgal Lic_StatutCourant(), "ESSAI_INTEGRE", "les refus ne modifient pas l'etat"
    Verifier DAL_Count("T_Audit", RecFrom("Module=LICENCE|Operation=ACTIVATION|Resultat=REFUSE")) >= 4, "refus d'activation journalises"

    st = Lic_InstallerAvecCle(Fx("LIC_VALIDE"), mo, fp, D_REF)
    VerifierEgal st, "VALIDE", "licence authentique activee"
    Verifier DAL_Count("T_Audit", RecFrom("Module=LICENCE|Operation=ACTIVATION|Resultat=OK")) = 1, "activation journalisee"
    Verifier Lic_ModuleAutorise("ELEVES") And Lic_ModuleAutorise("INSCR"), "modules achetes debloques"
    Verifier Not Lic_ModuleAutorise("FINANCES"), "module non achete bloque"
    el = NouvelEleve(e, "LICENCE", "Lea")
    Verifier Len(el) > 0, "saisie autorisee pour l'etablissement licencie"
    On Error Resume Next
    NouvelEleve eb, "HORS", "Licence"
    VerifierErreur Err.Number, Err.Description, SG_ERR_LECTURE_SEULE, "etablissement absent de la licence : consultation seule"
    On Error GoTo 0

    ' Alertes et expiration
    Lic_Recharger mo, fp, "2027-08-25"
    est = Lic_Etat()
    VerifierEgal RecStr(est, "JoursRestants"), "6", "jours restants"
    Verifier Len(RecStr(est, "Alerte")) > 0, "alerte avant expiration"
    Lic_Recharger mo, fp, "2027-09-10"
    VerifierEgal Lic_StatutCourant(), "GRACE", "periode de grace apres expiration"
    Verifier Len(NouvelEleve(e, "GRACE", "Gaston", "2014-01-01", "M")) > 0, "saisie possible pendant la grace"
    Lic_Recharger mo, fp, "2027-09-20"
    VerifierEgal Lic_StatutCourant(), "EXPIREE", "licence expiree apres la grace"
    nbAvant = DAL_Count("T_Eleves", RecNew())
    On Error Resume Next
    NouvelEleve e, "EXPIRE", "Eric", "2014-01-01", "M"
    VerifierErreur Err.Number, Err.Description, SG_ERR_LECTURE_SEULE, "expiration : saisie refusee"
    On Error GoTo 0
    VerifierEgal RecStr(Eleve_Get(e, el), "Nom"), "LICENCE", "expiration : les donnees restent consultables"
    VerifierEgal DAL_Count("T_Eleves", RecNew()), nbAvant, "expiration : aucune donnee supprimee"
    Verifier Export_TableCSV(e, "T_Eleves", DossierTemp() & "sg_test_export_expire.csv") >= 1, "expiration : export toujours possible"
    Verifier Len(Sauv_SimulerPourTests(e)) > 0, "expiration : sauvegarde toujours possible"

    ' Copie sur un autre poste
    Lic_Recharger mo, String$(64, "c"), "2027-09-20"
    VerifierEgal Lic_StatutCourant(), "AUTRE_INSTALLATION", "copie detectee (empreinte differente)"
    On Error Resume Next
    NouvelEleve e, "COPIE", "Carl", "2014-01-01", "M"
    VerifierErreur Err.Number, Err.Description, SG_ERR_LECTURE_SEULE, "copie : saisie refusee"
    On Error GoTo 0

    ' Recul d'horloge (date maximale deja vue : 2027-09-20)
    Lic_Recharger mo, fp, "2027-08-01"
    VerifierEgal Lic_StatutCourant(), "HORLOGE", "recul de l'horloge detecte"
    On Error Resume Next
    NouvelEleve e, "HORLOGE", "Hugo", "2014-01-01", "M"
    VerifierErreur Err.Number, Err.Description, SG_ERR_LECTURE_SEULE, "horloge incoherente : saisie refusee"
    On Error GoTo 0
    Lic_Recharger mo, fp, "2027-09-21"
    VerifierEgal Lic_StatutCourant(), "EXPIREE", "horloge corrigee : retour a l'etat reel"

    ' Revocation
    VerifierEgal Lic_ImporterRevocations(Fx("REVOCATIONS"), mo), 1, "liste de revocation authentique importee"
    Lic_Recharger mo, fp, "2027-09-21"
    VerifierEgal Lic_StatutCourant(), "REVOQUEE", "licence revoquee"
    On Error Resume Next
    Lic_ImporterRevocations Replace(Fx("REVOCATIONS"), "LIC-TEST-0001", "LIC-TEST-0009"), mo
    VerifierErreur Err.Number, Err.Description, SG_ERR_LICENCE, "liste de revocation falsifiee refusee"
    On Error GoTo 0
    Verifier DAL_Exists("T_Eleves", el), "revocation : donnees conservees"

    ' Perpetuelle et versions (evaluation pure)
    est = Lic_Evaluer(Fx("LIC_PERPETUELLE"), mo, fp, "2035-01-01", "", "", "")
    VerifierEgal RecStr(est, "Statut"), "VALIDE", "perpetuelle valide apres la fin de maintenance"
    Verifier InStr(RecStr(est, "Alerte"), "Maintenance") > 0, "perpetuelle : alerte de fin de maintenance seulement"
    VerifierEgal RecStr(Lic_Evaluer(Fx("LIC_PERP_VERSION"), mo, fp, D_REF, "", "", ""), "Statut"), "VERSION_NON_COUVERTE", "version majeure non couverte"
    VerifierEgal RecStr(Lic_Evaluer(Fx("LIC_ESSAI_TOUTE_INSTALL"), mo, String$(64, "d"), D_REF, "", "", ""), "Statut"), "VALIDE", "licence d'essai non liee a une installation"
    VerifierEgal RecStr(Lic_Evaluer("", mo, fp, D_REF, "", "", "2026-09-01"), "Statut"), "NON_ACTIVEE", "essai integre termine apres 30 jours"
    VerifierEgal RecStr(Lic_Evaluer("", mo, fp, D_REF, "", "", "2026-10-01"), "Ecriture"), "1", "essai integre en cours"
    Verifier InStr(Lic_GenererDemande("Groupe fictif", "ETB-A"), "EMPREINTE=" & Lic_Empreinte()) > 0, "demande d'activation hors ligne generee"
    Verifier Len(Lic_ClePublique()) = 0 Or Len(Lic_ClePublique()) >= 256, "aucune cle privee : seule la cle publique est integree"
End Sub

Private Sub T_Schema()
    VerifierEgal SG_VerifierSchema(), "OK", "schema complet"
    VerifierEgal Sys_Get("SCHEMA_VERSION"), CStr(SCHEMA_VERSION), "version de schema enregistree"
    SG_InitialiserBase "CENTRALISE", 2
    VerifierEgal DAL_Count("T_Roles", RecNew()), UBound(Sec_RolesParDefaut()) + 1, "initialisation relancee sans doublon"
End Sub

' Demande d'activation pour controle croise par l'outil editeur (Python).
Public Function SG_TestDemande(Optional ByVal filtre As Variant, Optional ByVal fixtures As Variant) As String
    ' Valeurs par defaut (IsMissing : forme compatible Excel et LibreOffice).
    If IsMissing(filtre) Then filtre = ""
    If IsMissing(fixtures) Then fixtures = ""
    DAL_ActiverModeTest
    SG_InitialiserBase "CENTRALISE", 2
    SG_TestDemande = Lic_GenererDemande("Groupe Scolaire Fictif Les Acacias", "ETB-A;ETB-B")
End Function
