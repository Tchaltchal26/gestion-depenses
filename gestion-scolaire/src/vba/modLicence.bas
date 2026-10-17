Attribute VB_Name = "modLicence"
'===============================================================================
' SG - Module modLicence (licences commerciales, cote CLIENT)
' Principe :
'   - la licence est un fichier texte signe (RSA-SHA256) par l'editeur ;
'   - le classeur ne contient QUE la cle publique : il peut verifier une
'     licence, jamais en fabriquer une ;
'   - la licence est liee a une empreinte d'installation (activation hors
'     ligne par echange de fichiers : demande .sgreq -> licence .sglic) ;
'   - expiration, revocation, autre installation ou horloge incoherente
'     => MODE CONSULTATION : les donnees restent lisibles, exportables et
'        sauvegardables. Aucune donnee n'est jamais supprimee ou chiffree.
' Limites (honnetes) : quelqu'un qui modifie le code VBA peut neutraliser ces
'   controles. La protection vise l'usage ordinaire et la copie simple ; voir
'   docs/03-licences.md pour les mesures complementaires (serveur, signature
'   du code, contrat).
'===============================================================================
Option Explicit

' Cle publique de l'editeur (module RSA hexadecimal). Remplacee par
' "sg_licence.py integrer-cle". Vide = produit non configure (essai integre).
Private Const CLE_PUBLIQUE_MODULE As String = "" ' <<SG_CLE_PUBLIQUE>>

Private Const JOURS_ESSAI_INTEGRE As Long = 30
Private Const ESSAI_MAX_ETABS As Long = 2
Private Const ESSAI_MAX_ELEVES As Long = 100
Public Const MODULES_NOYAU As String = "ETAB;PARAM;UTIL;ANNEE;AUDIT;LICENCE;SAUVEG;STRUCT"
' Modules utilisables pour la recuperation meme en mode consultation.
Private Const MODULES_RECUPERATION As String = "UTIL;LICENCE;SAUVEG"

Private mEtat As Variant
Private mCharge As Boolean

'-------------------------------------------------------------------------------
' Lecture du format
'-------------------------------------------------------------------------------
' Retourne un enregistrement des champs + __CONTENU (texte signe) + __SIGNATURE.
Public Function Lic_Parser(ByVal texte As String) As Variant
    Dim lignes As Variant, i As Long, l As String, contenu As String, sig As String, rec As Variant, p As Long
    rec = RecNew()
    texte = Replace(texte, vbCr & vbLf, vbLf)
    texte = Replace(texte, vbCr, vbLf)
    If Len(texte) > 0 Then
        If AscW(Left$(texte, 1)) = &HFEFF& Then texte = Mid$(texte, 2)
    End If
    lignes = Split(texte, vbLf)
    contenu = ""
    sig = ""
    For i = LBound(lignes) To UBound(lignes)
        l = Trim$(Replace(lignes(i), vbTab, ""))
        If Len(l) > 0 Then
            If Len(sig) > 0 Then SG_Raise SG_ERR_LICENCE, "Contenu apres la signature."
            If Left$(l, 10) = "SIGNATURE=" Then
                sig = Mid$(l, 11)
                If Len(sig) = 0 Then SG_Raise SG_ERR_LICENCE, "Signature vide."
            Else
                p = InStr(1, l, "=")
                If p < 2 Then SG_Raise SG_ERR_LICENCE, "Ligne de licence invalide."
                If RecHas(rec, Left$(l, p - 1)) Then SG_Raise SG_ERR_LICENCE, "Champ en double : " & Left$(l, p - 1)
                RecSet rec, Left$(l, p - 1), Mid$(l, p + 1)
                If Len(contenu) > 0 Then contenu = contenu & vbLf
                contenu = contenu & l
            End If
        End If
    Next i
    If Len(sig) = 0 Then SG_Raise SG_ERR_LICENCE, "Signature absente."
    RecSet rec, "__CONTENU", contenu
    RecSet rec, "__SIGNATURE", sig
    Lic_Parser = rec
End Function

Public Function Lic_SignatureValide(ByVal contenu As String, ByVal signatureB64 As String, ByVal modulusHex As String) As Boolean
    Dim m As Variant, s As Variant
    Lic_SignatureValide = False
    If Len(modulusHex) = 0 Then Exit Function
    On Error GoTo invalide
    m = Utf8Encode(contenu)
    s = Base64Decode(signatureB64)
    Lic_SignatureValide = RSA_VerifySHA256(m, s, modulusHex)
    Exit Function
invalide:
    Lic_SignatureValide = False
End Function

'-------------------------------------------------------------------------------
' Evaluation (fonction pure : aucune ecriture)
'-------------------------------------------------------------------------------
Private Function Etat(ByVal statut As String, ByVal message As String, Optional lic As Variant) As Variant
    Dim r As Variant
    If IsMissing(lic) Then r = RecNew() Else r = lic
    RecSet r, "Statut", statut
    RecSet r, "Message", message
    RecSet r, "Ecriture", IIf(InList(statut, "VALIDE;GRACE;ESSAI_INTEGRE"), "1", "0")
    Etat = r
End Function

' texte        : contenu du fichier .sglic ("" = aucune licence installee)
' modulusHex   : cle publique de verification
' empreinte    : empreinte de l'installation courante
' dateRef      : date du jour (AAAA-MM-JJ)
' dateMaxVue   : plus grande date deja observee (detection de recul d'horloge)
' revoquees    : liste "id;id;..." issue d'une liste de revocation signee
' dateInstall  : date de premiere installation (essai integre)
Public Function Lic_Evaluer(ByVal texte As String, ByVal modulusHex As String, ByVal empreinte As String, _
                            ByVal dateRef As String, ByVal dateMaxVue As String, ByVal revoquees As String, _
                            ByVal dateInstall As String) As Variant
    Dim lic As Variant, typ As String, expire As String, jours As Long, grace As Long, vmax As Long
    Dim fp As String, r As Variant

    If Len(dateMaxVue) > 0 Then
        If dateRef < ISOAddDays(dateMaxVue, -1) Then
            Lic_Evaluer = Etat("HORLOGE", "La date de l'ordinateur est anterieure a une date deja utilisee (" & dateMaxVue & "). Corriger la date systeme.")
            Exit Function
        End If
    End If

    If Len(Trim$(texte)) = 0 Then
        If Len(dateInstall) = 0 Then dateInstall = dateRef
        jours = ISODiffDays(dateInstall, dateRef)
        If jours < JOURS_ESSAI_INTEGRE Then
            r = Etat("ESSAI_INTEGRE", "Version d'essai : " & (JOURS_ESSAI_INTEGRE - jours) & " jour(s) restant(s), " & ESSAI_MAX_ETABS & " etablissements et " & ESSAI_MAX_ELEVES & " eleves maximum.")
            RecSet r, "MODULES", "*"
            RecSet r, "ETABLISSEMENTS", "*"
            RecSet r, "NB_ETABLISSEMENTS_MAX", CStr(ESSAI_MAX_ETABS)
            RecSet r, "MAX_ELEVES", CStr(ESSAI_MAX_ELEVES)
            RecSet r, "JoursRestants", CStr(JOURS_ESSAI_INTEGRE - jours)
        Else
            r = Etat("NON_ACTIVEE", "Periode d'essai terminee : installer une licence pour reprendre la saisie.")
        End If
        Lic_Evaluer = r
        Exit Function
    End If

    On Error GoTo malformee
    lic = Lic_Parser(texte)
    On Error GoTo 0

    If Not Lic_SignatureValide(RecStr(lic, "__CONTENU"), RecStr(lic, "__SIGNATURE"), modulusHex) Then
        Lic_Evaluer = Etat("INVALIDE", "Licence non authentique ou modifiee (signature incorrecte).")
        Exit Function
    End If
    If RecStr(lic, "FORMAT") <> "SGLIC-1" Then
        Lic_Evaluer = Etat("INVALIDE", "Format de licence non pris en charge.")
        Exit Function
    End If
    If InList(RecStr(lic, "LICENCE_ID"), revoquees) Then
        Lic_Evaluer = Etat("REVOQUEE", "Cette licence a ete revoquee par l'editeur.", lic)
        Exit Function
    End If
    fp = RecStr(lic, "EMPREINTE_INSTALLATION")
    If fp <> "*" And LCase$(fp) <> LCase$(empreinte) Then
        Lic_Evaluer = Etat("AUTRE_INSTALLATION", "Licence emise pour une autre installation (copie ou changement d'ordinateur). Demander un transfert a l'editeur.", lic)
        Exit Function
    End If
    vmax = CLng(RecStr(lic, "VERSION_MAJEURE_MAX"))
    If SG_VERSION_MAJEURE > vmax Then
        Lic_Evaluer = Etat("VERSION_NON_COUVERTE", "Cette version du logiciel n'est pas couverte par la licence (version majeure max " & vmax & ").", lic)
        Exit Function
    End If

    typ = RecStr(lic, "TYPE")
    expire = RecStr(lic, "EXPIRE_LE")
    grace = CLng("0" & RecStr(lic, "GRACE_JOURS"))
    If typ = "PERPETUELLE" Then
        ' La fin de maintenance ne desactive jamais une licence perpetuelle.
        r = Etat("VALIDE", "Licence perpetuelle valide.", lic)
        If Len(RecStr(lic, "MAINTENANCE_JUSQU_AU")) > 0 And dateRef > RecStr(lic, "MAINTENANCE_JUSQU_AU") Then
            RecSet r, "Alerte", "Maintenance et mises a jour terminees le " & RecStr(lic, "MAINTENANCE_JUSQU_AU") & " (l'utilisation reste autorisee)."
        End If
        Lic_Evaluer = r
        Exit Function
    End If
    If Not IsISODate(expire) Then
        Lic_Evaluer = Etat("INVALIDE", "Date d'expiration absente ou invalide.")
        Exit Function
    End If
    jours = ISODiffDays(dateRef, expire)
    If jours >= 0 Then
        r = Etat("VALIDE", "Licence valide jusqu'au " & expire & ".", lic)
        RecSet r, "JoursRestants", CStr(jours)
        If jours <= 7 Then
            RecSet r, "Alerte", "La licence expire dans " & jours & " jour(s)."
        ElseIf jours <= 15 Then
            RecSet r, "Alerte", "La licence expire dans " & jours & " jours."
        ElseIf jours <= 30 Then
            RecSet r, "Alerte", "La licence expire dans " & jours & " jours."
        End If
    ElseIf -jours <= grace Then
        r = Etat("GRACE", "Licence expiree le " & expire & " : periode de grace (" & (grace + jours) & " jour(s) restant(s)).", lic)
        RecSet r, "Alerte", RecStr(r, "Message")
    Else
        r = Etat("EXPIREE", "Licence expiree le " & expire & ". Renouveler la licence pour reprendre la saisie.", lic)
    End If
    Lic_Evaluer = r
    Exit Function
malformee:
    Lic_Evaluer = Etat("INVALIDE", "Fichier de licence illisible : " & Err.Description)
End Function

'-------------------------------------------------------------------------------
' Etat courant de l'installation
'-------------------------------------------------------------------------------
Public Function Lic_ClePublique() As String
    Lic_ClePublique = CLE_PUBLIQUE_MODULE
End Function

' Empreinte : identifiant d'installation (cree au premier lancement) +
' caracteristiques du poste. Une copie du classeur sur un autre poste
' produit une empreinte differente.
Public Function Lic_Empreinte() As String
    Dim info As String
    info = Sys_Get("INSTALL_ID", "") & "|" & Environ$("COMPUTERNAME") & "|" & Environ$("USERDOMAIN") & "|" & Environ$("HOSTNAME")
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then info = info & "|" & CStr(fso.GetDrive(fso.GetDriveName(ThisWorkbook.Path)).SerialNumber)
    On Error GoTo 0
    Lic_Empreinte = SHA256Hex(info)
End Function

Private Function ListeRevoquees() As String
    Dim c As Collection
    Set c = DAL_Query("T_Revocations", RecNew())
    ListeRevoquees = JoinCollection(c, ";")
End Function

' Evalue et memorise l'etat. Appele a l'ouverture et apres installation.
Public Sub Lic_Recharger(Optional ByVal modulusHex As Variant, Optional ByVal empreinte As Variant, _
                         Optional ByVal dateRef As Variant)
    ' Valeurs par defaut (IsMissing : forme compatible Excel et LibreOffice).
    If IsMissing(modulusHex) Then modulusHex = ""
    If IsMissing(empreinte) Then empreinte = ""
    If IsMissing(dateRef) Then dateRef = ""
    If Len(modulusHex) = 0 Then modulusHex = CLE_PUBLIQUE_MODULE
    If modulusHex <> CLE_PUBLIQUE_MODULE And Not DAL_ModeTest() Then SG_Raise SG_ERR_LICENCE, "Cle de verification non autorisee."
    If Len(empreinte) = 0 Then empreinte = Lic_Empreinte()
    If Len(dateRef) = 0 Then dateRef = TodayISO()
    mEtat = Lic_Evaluer(Sys_Get("LICENCE_TEXTE", ""), modulusHex, empreinte, dateRef, _
                        Sys_Get("DATE_MAX_VUE", ""), ListeRevoquees(), Sys_Get("INSTALL_DATE", dateRef))
    If RecStr(mEtat, "Statut") <> "HORLOGE" Then
        If dateRef > Sys_Get("DATE_MAX_VUE", "") Then Sys_Set "DATE_MAX_VUE", dateRef
    End If
    mCharge = True
End Sub

Public Function Lic_Etat() As Variant
    If Not mCharge Then Lic_Recharger
    Lic_Etat = mEtat
End Function

Public Function Lic_StatutCourant() As String
    Lic_StatutCourant = RecStr(Lic_Etat(), "Statut")
End Function

Public Function Lic_MessageCourant() As String
    Lic_MessageCourant = RecStr(Lic_Etat(), "Message")
End Function

Public Function Lic_ModuleAutorise(ByVal module As String) As Boolean
    Dim e As Variant
    e = Lic_Etat()
    If InList(module, MODULES_NOYAU) Then Lic_ModuleAutorise = True: Exit Function
    Lic_ModuleAutorise = (RecStr(e, "MODULES") = "*" Or InList(module, RecStr(e, "MODULES")))
End Function

Public Function Lic_EcritureAutorisee(ByVal etabId As String, ByVal module As String) As Boolean
    Dim e As Variant, code As String
    Lic_EcritureAutorisee = False
    If InList(module, MODULES_RECUPERATION) Then Lic_EcritureAutorisee = True: Exit Function
    e = Lic_Etat()
    If RecStr(e, "Ecriture") <> "1" Then Exit Function
    If Not Lic_ModuleAutorise(module) Then Exit Function
    If etabId <> "*" And Len(etabId) > 0 And RecStr(e, "ETABLISSEMENTS") <> "*" Then
        If Not DAL_Exists("T_Etablissements", etabId) Then Exit Function
        code = DAL_Value("T_Etablissements", etabId, "Code")
        If Not InList(code, RecStr(e, "ETABLISSEMENTS")) Then Exit Function
    End If
    Lic_EcritureAutorisee = True
End Function

Public Function Lic_NbEtabsMax() As Long
    Dim e As Variant
    e = Lic_Etat()
    If RecStr(e, "Ecriture") <> "1" Then Lic_NbEtabsMax = 0: Exit Function
    Lic_NbEtabsMax = CLng("0" & RecStr(e, "NB_ETABLISSEMENTS_MAX"))
End Function

Public Function Lic_EtabAutorise(ByVal code As String) As Boolean
    Dim e As Variant
    e = Lic_Etat()
    Lic_EtabAutorise = (RecStr(e, "ETABLISSEMENTS") = "*" Or InList(code, RecStr(e, "ETABLISSEMENTS")))
End Function

Public Function Lic_MaxEleves() As Long
    ' 0 = illimite
    Lic_MaxEleves = CLng("0" & RecStr(Lic_Etat(), "MAX_ELEVES"))
End Function

'-------------------------------------------------------------------------------
' Activation, revocation, demande
'-------------------------------------------------------------------------------
Public Function Lic_Installer(ByVal texte As String) As String
    Lic_Installer = Lic_InstallerAvecCle(texte, CLE_PUBLIQUE_MODULE, Lic_Empreinte(), TodayISO())
End Function

' Variante parametree : la cle differente de la cle integree n'est acceptee
' qu'en mode test (classeur de tests automatises).
Public Function Lic_InstallerAvecCle(ByVal texte As String, ByVal modulusHex As String, ByVal empreinte As String, ByVal dateRef As String) As String
    Dim e As Variant, ancien As String
    Sec_Exiger "*", "LICENCE", NIV_ADMIN
    If modulusHex <> CLE_PUBLIQUE_MODULE And Not DAL_ModeTest() Then SG_Raise SG_ERR_LICENCE, "Cle de verification non autorisee."
    e = Lic_Evaluer(texte, modulusHex, empreinte, dateRef, Sys_Get("DATE_MAX_VUE", ""), ListeRevoquees(), Sys_Get("INSTALL_DATE", dateRef))
    If InList(RecStr(e, "Statut"), "INVALIDE;AUTRE_INSTALLATION;REVOQUEE") Then
        Audit_Log "*", "LICENCE", "ACTIVATION", "T_Systeme", "LICENCE_TEXTE", "", "", RecStr(e, "LICENCE_ID"), RecStr(e, "Message"), "REFUSE"
        SG_Raise SG_ERR_LICENCE, "Licence refusee : " & RecStr(e, "Message")
    End If
    ancien = ""
    If Len(Sys_Get("LICENCE_TEXTE", "")) > 0 Then
        On Error Resume Next
        ancien = RecStr(Lic_Parser(Sys_Get("LICENCE_TEXTE", "")), "LICENCE_ID")
        On Error GoTo 0
    End If
    Sys_Set "LICENCE_TEXTE", texte
    Sys_Set "LICENCE_ACTIVEE_LE", NowStamp()
    Audit_Log "*", "LICENCE", "ACTIVATION", "T_Systeme", "LICENCE_TEXTE", "LICENCE_ID", ancien, RecStr(e, "LICENCE_ID"), _
              RecStr(e, "TYPE") & " ; " & RecStr(e, "Statut") & " ; expire " & RecStr(e, "EXPIRE_LE")
    Lic_Recharger modulusHex, empreinte, dateRef
    Lic_InstallerAvecCle = Lic_StatutCourant()
End Function

Public Function Lic_ImporterRevocations(ByVal texte As String, Optional ByVal modulusHex As Variant) As Long
    ' Valeurs par defaut (IsMissing : forme compatible Excel et LibreOffice).
    If IsMissing(modulusHex) Then modulusHex = ""
    Dim rec As Variant, ids As Variant, i As Long, n As Long
    Sec_Exiger "*", "LICENCE", NIV_ADMIN
    If Len(modulusHex) = 0 Then modulusHex = CLE_PUBLIQUE_MODULE
    If modulusHex <> CLE_PUBLIQUE_MODULE And Not DAL_ModeTest() Then SG_Raise SG_ERR_LICENCE, "Cle de verification non autorisee."
    rec = Lic_Parser(texte)
    If RecStr(rec, "FORMAT") <> "SGREV-1" Or Not Lic_SignatureValide(RecStr(rec, "__CONTENU"), RecStr(rec, "__SIGNATURE"), modulusHex) Then
        Audit_Log "*", "LICENCE", "IMPORT_REVOCATIONS", "T_Revocations", "", "", "", "", "Liste non authentique", "REFUSE"
        SG_Raise SG_ERR_LICENCE, "Liste de revocation non authentique."
    End If
    ids = Split(RecStr(rec, "LICENCES"), ";")
    For i = LBound(ids) To UBound(ids)
        If Len(ids(i)) > 0 And Not DAL_Exists("T_Revocations", CStr(ids(i))) Then
            DAL_Insert "T_Revocations", RecFrom("LicenceID=" & ids(i) & "|ImporteeLe=" & NowStamp() & "|ListeEmiseLe=" & RecStr(rec, "EMIS_LE"))
            n = n + 1
        End If
    Next i
    Audit_Log "*", "LICENCE", "IMPORT_REVOCATIONS", "T_Revocations", "", "", "", CStr(n) & " nouvelle(s)", "Liste du " & RecStr(rec, "EMIS_LE")
    mCharge = False
    Lic_ImporterRevocations = n
End Function

' Demande d'activation hors ligne a transmettre a l'editeur (.sgreq).
Public Function Lic_GenererDemande(ByVal titulaire As String, ByVal codesEtabs As String) As String
    Dim lignes As String, controle As String
    lignes = "FORMAT=SGREQ-1" & vbLf & "INSTALL_ID=" & Sys_Get("INSTALL_ID", "") & vbLf & _
             "EMPREINTE=" & Lic_Empreinte() & vbLf & "TITULAIRE=" & SafeField(titulaire) & vbLf & _
             "ETABLISSEMENTS=" & codesEtabs & vbLf & "VERSION=" & SG_VERSION & vbLf & "DATE=" & TodayISO()
    controle = SHA256Hex(lignes)
    Audit_Log "*", "LICENCE", "DEMANDE_ACTIVATION", "", "", "", "", codesEtabs, ""
    Lic_GenererDemande = lignes & vbLf & "CONTROLE=" & controle & vbLf
End Function
