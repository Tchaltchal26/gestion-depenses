Attribute VB_Name = "modSecurite"
'===============================================================================
' SG - Module modSecurite (utilisateurs, roles, droits, session)
' Modele :
'   - comptes individuels (T_Utilisateurs), mots de passe sales et haches ;
'   - roles configurables (T_Roles) et droits par module (T_Permissions) ;
'   - habilitations PAR ETABLISSEMENT (T_Habilitations) : un utilisateur n'a
'     acces qu'aux etablissements pour lesquels il est habilite ;
'   - indicateurs globaux : AdminTech (technique : comptes, licences,
'     sauvegardes, SANS acces aux dossiers eleves) et AdminCentral (reseau :
'     creation d'etablissements et gestion des comptes).
' Niveaux : 0 aucun, 1 lecture, 2 ecriture, 3 validation, 4 administration.
' Limite : dans un classeur Excel, ces controles s'appliquent a l'interface
'   et au code SG ; ils n'empechent pas un utilisateur qui ouvre le fichier
'   avec d'autres outils de lire les feuilles. Voir docs/01-architecture.md.
'===============================================================================
Option Explicit

Public Const NIV_AUCUN As Long = 0
Public Const NIV_LECTURE As Long = 1
Public Const NIV_ECRITURE As Long = 2
Public Const NIV_VALIDATION As Long = 3
Public Const NIV_ADMIN As Long = 4

Public Const MODULES_SG As String = "ETAB;PARAM;UTIL;ANNEE;STRUCT;ELEVES;INSCR;AUDIT;LICENCE;SAUVEG"
Private Const MAX_ECHECS As Long = 5
Private Const MINUTES_VERROU As Long = 15

Private mUserID As String
Private mLogin As String

'-------------------------------------------------------------------------------
' Session
'-------------------------------------------------------------------------------
Public Function Sec_UtilisateurCourant() As String
    If Len(mUserID) = 0 Then Sec_UtilisateurCourant = "SYSTEME" Else Sec_UtilisateurCourant = mUserID
End Function

Public Function Sec_LoginCourant() As String
    Sec_LoginCourant = mLogin
End Function

Public Function Sec_EstConnecte() As Boolean
    Sec_EstConnecte = (Len(mUserID) > 0)
End Function

Public Function Sec_Connexion(ByVal login As String, ByVal motDePasse As String) As Boolean
    Dim uid As String, u As Variant, echecs As Long
    Sec_Deconnexion
    uid = DAL_First("T_Utilisateurs", RecFrom("Login=" & LCase$(Trim$(login))))
    If Len(uid) = 0 Then
        Audit_Log "*", "UTIL", "CONNEXION", "T_Utilisateurs", "", "Login", "", login, "", "ECHEC"
        SG_Raise SG_ERR_ACCES, "Identifiant ou mot de passe incorrect."
    End If
    u = DAL_Get("T_Utilisateurs", uid)
    If RecStr(u, "Actif") <> "1" Then
        Audit_Log "*", "UTIL", "CONNEXION", "T_Utilisateurs", uid, "", "", "", "Compte desactive", "REFUSE"
        SG_Raise SG_ERR_ACCES, "Ce compte est desactive."
    End If
    If Len(RecStr(u, "VerrouJusqua")) > 0 And RecStr(u, "VerrouJusqua") > NowStamp() Then
        Audit_Log "*", "UTIL", "CONNEXION", "T_Utilisateurs", uid, "", "", "", "Compte verrouille", "REFUSE"
        SG_Raise SG_ERR_ACCES, "Compte temporairement verrouille apres plusieurs echecs. Reessayer plus tard."
    End If
    If PasswordHash(motDePasse, RecStr(u, "Sel"), CLng(RecStr(u, "Iterations"))) <> RecStr(u, "HashMdp") Then
        echecs = CLng("0" & RecStr(u, "Echecs")) + 1
        DAL_Update "T_Utilisateurs", uid, "Echecs", CStr(echecs)
        If echecs >= MAX_ECHECS Then
            DAL_Update "T_Utilisateurs", uid, "VerrouJusqua", StampOf(Now + MINUTES_VERROU / 1440#)
            Audit_Log "*", "UTIL", "VERROUILLAGE", "T_Utilisateurs", uid, "Echecs", "", CStr(echecs), "Echecs repetes", "OK"
        End If
        Audit_Log "*", "UTIL", "CONNEXION", "T_Utilisateurs", uid, "", "", "", "Mot de passe incorrect", "ECHEC"
        SG_Raise SG_ERR_ACCES, "Identifiant ou mot de passe incorrect."
    End If
    DAL_Update "T_Utilisateurs", uid, "Echecs", "0"
    DAL_Update "T_Utilisateurs", uid, "VerrouJusqua", ""
    DAL_Update "T_Utilisateurs", uid, "DernierAcces", NowStamp()
    mUserID = uid
    mLogin = RecStr(u, "Login")
    Audit_Log "*", "UTIL", "CONNEXION", "T_Utilisateurs", uid, "", "", "", "", "OK"
    Sec_Connexion = True
End Function

Public Sub Sec_Deconnexion()
    If Len(mUserID) > 0 Then Audit_Log "*", "UTIL", "DECONNEXION", "T_Utilisateurs", mUserID, "", "", "", ""
    mUserID = ""
    mLogin = ""
End Sub

'-------------------------------------------------------------------------------
' Calcul des droits
'-------------------------------------------------------------------------------
Private Function NiveauGlobal(u As Variant, ByVal module As String) As Long
    Dim n As Long
    n = NIV_AUCUN
    If RecStr(u, "AdminTech") = "1" Then
        If InList(module, "UTIL;LICENCE;SAUVEG;PARAM") Then n = NIV_ADMIN
        If module = "AUDIT" Or module = "ETAB" Then n = NIV_LECTURE
    End If
    If RecStr(u, "AdminCentral") = "1" Then
        If InList(module, "ETAB;UTIL") Then n = NIV_ADMIN
        If InList(module, "AUDIT;LICENCE;SAUVEG") And n < NIV_LECTURE Then n = NIV_LECTURE
    End If
    NiveauGlobal = n
End Function

Public Function Sec_NiveauRole(ByVal roleId As String, ByVal module As String) As Long
    Dim pid As String
    pid = DAL_First("T_Permissions", RecFrom("RoleID=" & roleId & "|Module=" & module))
    If Len(pid) = 0 Then Sec_NiveauRole = NIV_AUCUN Else Sec_NiveauRole = CLng("0" & DAL_Value("T_Permissions", pid, "Niveau"))
End Function

' Niveau effectif de l'utilisateur courant pour un etablissement et un module.
' etabId = "*" designe les operations globales (comptes, licences...).
Public Function Sec_Niveau(ByVal etabId As String, ByVal module As String) As Long
    Dim u As Variant, n As Long, habs As Collection, i As Long, h As Variant, nr As Long, today As String
    Sec_Niveau = NIV_AUCUN
    If Len(mUserID) = 0 Then Exit Function
    u = DAL_Get("T_Utilisateurs", mUserID)
    If RecStr(u, "Actif") <> "1" Then Exit Function
    n = NiveauGlobal(u, module)
    If etabId <> "*" And Len(etabId) > 0 Then
        today = TodayISO()
        Set habs = DAL_Query("T_Habilitations", RecFrom("UserID=" & mUserID & "|EtabID=" & etabId & "|Actif=1"))
        For i = 1 To habs.Count
            h = DAL_Get("T_Habilitations", habs(i))
            If RecStr(h, "DateDebut") <= today And (Len(RecStr(h, "DateFin")) = 0 Or RecStr(h, "DateFin") >= today) Then
                nr = Sec_NiveauRole(RecStr(h, "RoleID"), module)
                If nr > n Then n = nr
            End If
        Next i
    End If
    Sec_Niveau = n
End Function

Public Function Sec_AAccesEtab(ByVal etabId As String) As Boolean
    Dim m As Variant
    For Each m In Split(MODULES_SG, ";")
        If Sec_Niveau(etabId, CStr(m)) > NIV_AUCUN Then Sec_AAccesEtab = True: Exit Function
    Next m
    Sec_AAccesEtab = False
End Function

' Point de controle UNIQUE appele par toutes les operations metier.
Public Sub Sec_Exiger(ByVal etabId As String, ByVal module As String, ByVal niveau As Long)
    If Len(mUserID) = 0 Then SG_Raise SG_ERR_ACCES, "Aucun utilisateur connecte."
    If Sec_Niveau(etabId, module) < niveau Then
        Audit_Log etabId, module, "ACCES_REFUSE", "", "", "", "", "niveau " & niveau, "", "REFUSE"
        SG_Raise SG_ERR_ACCES, "Droits insuffisants pour cette operation (" & module & ", niveau " & niveau & ")."
    End If
    If niveau >= NIV_ECRITURE Then
        If Not Lic_EcritureAutorisee(etabId, module) Then
            Audit_Log etabId, module, "ECRITURE_REFUSEE_LICENCE", "", "", "", "", Lic_StatutCourant(), "", "REFUSE"
            SG_Raise SG_ERR_LECTURE_SEULE, "Mode consultation : " & Lic_MessageCourant() & " Les donnees restent consultables et exportables."
        End If
    End If
End Sub

'-------------------------------------------------------------------------------
' Gestion des comptes
'-------------------------------------------------------------------------------
Private Sub ControlerMotDePasse(ByVal mdp As String)
    Dim i As Long, lettre As Boolean, chiffre As Boolean, c As String
    If Len(mdp) < 8 Then SG_Raise SG_ERR_VALIDATION, "Le mot de passe doit contenir au moins 8 caracteres."
    For i = 1 To Len(mdp)
        c = Mid$(mdp, i, 1)
        If c >= "0" And c <= "9" Then
            chiffre = True
        ElseIf UCase$(c) <> LCase$(c) Then
            lettre = True
        End If
    Next i
    If Not (lettre And chiffre) Then SG_Raise SG_ERR_VALIDATION, "Le mot de passe doit contenir des lettres et des chiffres."
End Sub

Private Function CreerCompte(ByVal login As String, ByVal nomComplet As String, ByVal mdp As String, _
                             ByVal adminTech As Boolean, ByVal adminCentral As Boolean) As String
    Dim rec As Variant, uid As String, sel As String, it As Long
    login = LCase$(Trim$(login))
    If Len(login) < 3 Or InStr(login, " ") > 0 Or InStr(login, "|") > 0 Or InStr(login, "=") > 0 Then SG_Raise SG_ERR_VALIDATION, "Identifiant de connexion invalide (3 caracteres minimum, sans espace)."
    If Len(Trim$(nomComplet)) = 0 Then SG_Raise SG_ERR_VALIDATION, "Le nom complet est obligatoire."
    If Len(DAL_First("T_Utilisateurs", RecFrom("Login=" & login))) > 0 Then SG_Raise SG_ERR_DOUBLON, "Cet identifiant existe deja : " & login
    ControlerMotDePasse mdp
    sel = RandomHex(32)
    it = CLng(Sys_Get("ITER_MDP", "1000"))
    uid = DAL_NewId("T_Utilisateurs")
    rec = RecNew()
    RecSet rec, "UserID", uid
    RecSet rec, "Login", login
    RecSet rec, "NomComplet", Trim$(nomComplet)
    RecSet rec, "Sel", sel
    RecSet rec, "HashMdp", PasswordHash(mdp, sel, it)
    RecSet rec, "Iterations", CStr(it)
    RecSet rec, "Actif", "1"
    RecSet rec, "AdminTech", IIf(adminTech, "1", "0")
    RecSet rec, "AdminCentral", IIf(adminCentral, "1", "0")
    RecSet rec, "Echecs", "0"
    RecSet rec, "DoitChangerMdp", "0"
    RecSet rec, "DateCreation", NowStamp()
    DAL_Insert "T_Utilisateurs", rec
    ' Le journal ne contient jamais le sel ni l'empreinte du mot de passe.
    Audit_Log "*", "UTIL", "CREATION", "T_Utilisateurs", uid, "", "", "Login=" & login & " ; AdminTech=" & IIf(adminTech, "1", "0") & " ; AdminCentral=" & IIf(adminCentral, "1", "0"), ""
    CreerCompte = uid
End Function

' Premier lancement uniquement : cree l'administrateur technique initial.
Public Function Sec_InitialiserAdministrateur(ByVal login As String, ByVal nomComplet As String, ByVal mdp As String) As String
    If DAL_Count("T_Utilisateurs", RecNew()) > 0 Then SG_Raise SG_ERR_ETAT, "Des comptes existent deja : initialisation refusee."
    Sec_InitialiserAdministrateur = CreerCompte(login, nomComplet, mdp, True, True)
End Function

Public Function Sec_CreerUtilisateur(ByVal login As String, ByVal nomComplet As String, ByVal mdp As String, _
                                     Optional ByVal adminTech As Boolean = False, Optional ByVal adminCentral As Boolean = False) As String
    Sec_Exiger "*", "UTIL", NIV_ADMIN
    Sec_CreerUtilisateur = CreerCompte(login, nomComplet, mdp, adminTech, adminCentral)
End Function

Public Function Sec_RoleID(ByVal code As String) As String
    Sec_RoleID = DAL_First("T_Roles", RecFrom("Code=" & code))
    If Len(Sec_RoleID) = 0 Then SG_Raise SG_ERR_INTROUVABLE, "Role inconnu : " & code
End Function

Public Function Sec_Habiliter(ByVal userId As String, ByVal etabId As String, ByVal roleCode As String, _
                              Optional ByVal dateDebut As Variant, Optional ByVal dateFin As Variant) As String
    ' Valeurs par defaut (IsMissing : forme compatible Excel et LibreOffice).
    If IsMissing(dateDebut) Then dateDebut = ""
    If IsMissing(dateFin) Then dateFin = ""
    Dim rec As Variant, hid As String
    Sec_Exiger "*", "UTIL", NIV_ADMIN
    If Not DAL_Exists("T_Utilisateurs", userId) Then SG_Raise SG_ERR_INTROUVABLE, "Utilisateur inconnu."
    If Not DAL_Exists("T_Etablissements", etabId) Then SG_Raise SG_ERR_INTROUVABLE, "Etablissement inconnu."
    If Len(dateDebut) = 0 Then dateDebut = TodayISO()
    If Not IsISODate(dateDebut) Then SG_Raise SG_ERR_VALIDATION, "Date de debut invalide."
    If Len(dateFin) > 0 Then
        If Not IsISODate(dateFin) Or dateFin < dateDebut Then SG_Raise SG_ERR_VALIDATION, "Date de fin invalide."
    End If
    hid = DAL_NewId("T_Habilitations")
    rec = RecNew()
    RecSet rec, "HabID", hid
    RecSet rec, "UserID", userId
    RecSet rec, "EtabID", etabId
    RecSet rec, "RoleID", Sec_RoleID(roleCode)
    RecSet rec, "DateDebut", dateDebut
    RecSet rec, "DateFin", dateFin
    RecSet rec, "Actif", "1"
    DAL_Insert "T_Habilitations", rec
    Audit_LogCreation etabId, "UTIL", "T_Habilitations", rec, "Attribution du role " & roleCode
    Sec_Habiliter = hid
End Function

Public Sub Sec_RetirerHabilitation(ByVal habId As String, ByVal motif As String)
    Dim h As Variant
    Sec_Exiger "*", "UTIL", NIV_ADMIN
    If Len(Trim$(motif)) = 0 Then SG_Raise SG_ERR_VALIDATION, "Motif obligatoire."
    h = DAL_Get("T_Habilitations", habId)
    DAL_Update "T_Habilitations", habId, "Actif", "0"
    DAL_Update "T_Habilitations", habId, "DateFin", TodayISO()
    Audit_Log RecStr(h, "EtabID"), "UTIL", "RETRAIT_HABILITATION", "T_Habilitations", habId, "Actif", "1", "0", motif
End Sub

Public Sub Sec_DesactiverUtilisateur(ByVal userId As String, ByVal motif As String)
    Sec_Exiger "*", "UTIL", NIV_ADMIN
    If Len(Trim$(motif)) = 0 Then SG_Raise SG_ERR_VALIDATION, "Motif obligatoire."
    If userId = mUserID Then SG_Raise SG_ERR_ETAT, "Impossible de desactiver son propre compte."
    DAL_Update "T_Utilisateurs", userId, "Actif", "0"
    Audit_Log "*", "UTIL", "DESACTIVATION", "T_Utilisateurs", userId, "Actif", "1", "0", motif
End Sub

Public Sub Sec_ChangerMotDePasse(ByVal ancien As String, ByVal nouveau As String)
    Dim u As Variant, sel As String, it As Long
    If Len(mUserID) = 0 Then SG_Raise SG_ERR_ACCES, "Aucun utilisateur connecte."
    u = DAL_Get("T_Utilisateurs", mUserID)
    If PasswordHash(ancien, RecStr(u, "Sel"), CLng(RecStr(u, "Iterations"))) <> RecStr(u, "HashMdp") Then
        Audit_Log "*", "UTIL", "CHANGEMENT_MDP", "T_Utilisateurs", mUserID, "", "", "", "Ancien mot de passe incorrect", "ECHEC"
        SG_Raise SG_ERR_ACCES, "Ancien mot de passe incorrect."
    End If
    ControlerMotDePasse nouveau
    sel = RandomHex(32)
    it = CLng(Sys_Get("ITER_MDP", "1000"))
    DAL_Update "T_Utilisateurs", mUserID, "Sel", sel
    DAL_Update "T_Utilisateurs", mUserID, "Iterations", CStr(it)
    DAL_Update "T_Utilisateurs", mUserID, "HashMdp", PasswordHash(nouveau, sel, it)
    DAL_Update "T_Utilisateurs", mUserID, "DoitChangerMdp", "0"
    Audit_Log "*", "UTIL", "CHANGEMENT_MDP", "T_Utilisateurs", mUserID, "", "", "", ""
End Sub

' Procedure de recuperation d'acces : un administrateur fixe un mot de passe
' temporaire, que l'utilisateur devra changer a la prochaine connexion.
Public Sub Sec_ReinitialiserMotDePasse(ByVal userId As String, ByVal mdpTemporaire As String, ByVal motif As String)
    Dim sel As String, it As Long
    Sec_Exiger "*", "UTIL", NIV_ADMIN
    If Len(Trim$(motif)) = 0 Then SG_Raise SG_ERR_VALIDATION, "Motif obligatoire."
    ControlerMotDePasse mdpTemporaire
    sel = RandomHex(32)
    it = CLng(Sys_Get("ITER_MDP", "1000"))
    DAL_Update "T_Utilisateurs", userId, "Sel", sel
    DAL_Update "T_Utilisateurs", userId, "Iterations", CStr(it)
    DAL_Update "T_Utilisateurs", userId, "HashMdp", PasswordHash(mdpTemporaire, sel, it)
    DAL_Update "T_Utilisateurs", userId, "DoitChangerMdp", "1"
    DAL_Update "T_Utilisateurs", userId, "Echecs", "0"
    DAL_Update "T_Utilisateurs", userId, "VerrouJusqua", ""
    Audit_Log "*", "UTIL", "REINITIALISATION_MDP", "T_Utilisateurs", userId, "", "", "", motif
End Sub

Public Sub Sec_ModifierPermission(ByVal roleCode As String, ByVal module As String, ByVal niveau As Long, ByVal motif As String)
    Dim rid As String, pid As String, ancien As String, rec As Variant
    Sec_Exiger "*", "UTIL", NIV_ADMIN
    If Len(Trim$(motif)) = 0 Then SG_Raise SG_ERR_VALIDATION, "Motif obligatoire."
    If Not InList(module, MODULES_SG) Then SG_Raise SG_ERR_VALIDATION, "Module inconnu : " & module
    If niveau < NIV_AUCUN Or niveau > NIV_ADMIN Then SG_Raise SG_ERR_VALIDATION, "Niveau invalide."
    rid = Sec_RoleID(roleCode)
    pid = DAL_First("T_Permissions", RecFrom("RoleID=" & rid & "|Module=" & module))
    If Len(pid) = 0 Then
        pid = DAL_NewId("T_Permissions")
        rec = RecFrom("PermID=" & pid & "|RoleID=" & rid & "|Module=" & module & "|Niveau=" & niveau)
        DAL_Insert "T_Permissions", rec
        ancien = "0"
    Else
        ancien = DAL_Value("T_Permissions", pid, "Niveau")
        DAL_Update "T_Permissions", pid, "Niveau", CStr(niveau)
    End If
    Audit_Log "*", "UTIL", "MODIFICATION_DROITS", "T_Permissions", pid, roleCode & "." & module, ancien, CStr(niveau), motif
End Sub

'-------------------------------------------------------------------------------
' Roles par defaut (cree a l'initialisation ; modifiables ensuite)
' Format : CODE|Libelle|MODULE:niveau,MODULE:niveau...
'-------------------------------------------------------------------------------
Public Function Sec_RolesParDefaut() As Variant
    Sec_RolesParDefaut = Array( _
        "DIRECTEUR|Directeur / Chef d'etablissement|ETAB:1,PARAM:4,UTIL:1,ANNEE:4,STRUCT:4,ELEVES:4,INSCR:4,AUDIT:1,LICENCE:1,SAUVEG:2", _
        "SECRETAIRE|Secretaire|PARAM:1,ANNEE:1,STRUCT:1,ELEVES:2,INSCR:2,SAUVEG:2", _
        "COMPTABLE|Comptable|PARAM:1,ANNEE:1,STRUCT:1,ELEVES:1,INSCR:1", _
        "CAISSIER|Caissier|ANNEE:1,ELEVES:1,INSCR:1", _
        "RESP_PEDAGO|Responsable pedagogique|PARAM:1,ANNEE:1,STRUCT:3,ELEVES:1,INSCR:3", _
        "ENSEIGNANT|Enseignant|ANNEE:1,STRUCT:1,ELEVES:1", _
        "FORMATEUR|Formateur|ANNEE:1,STRUCT:1,ELEVES:1", _
        "SURVEILLANT|Surveillant|ANNEE:1,STRUCT:1,ELEVES:1", _
        "GEST_STOCK|Gestionnaire de stock|PARAM:1", _
        "AUDITEUR|Auditeur (lecture seule)|ETAB:1,PARAM:1,UTIL:1,ANNEE:1,STRUCT:1,ELEVES:1,INSCR:1,AUDIT:1,LICENCE:1,SAUVEG:1")
End Function

Public Sub Sec_CreerRolesParDefaut()
    Dim roles As Variant, i As Long, parts As Variant, perms As Variant, j As Long, rid As String, kv As Variant
    roles = Sec_RolesParDefaut()
    For i = LBound(roles) To UBound(roles)
        parts = Split(roles(i), "|")
        If Len(DAL_First("T_Roles", RecFrom("Code=" & parts(0)))) = 0 Then
            rid = DAL_NewId("T_Roles")
            DAL_Insert "T_Roles", RecFrom("RoleID=" & rid & "|Code=" & parts(0) & "|Libelle=" & parts(1) & "|Systeme=1")
            perms = Split(parts(2), ",")
            For j = LBound(perms) To UBound(perms)
                kv = Split(perms(j), ":")
                DAL_Insert "T_Permissions", RecFrom("PermID=" & DAL_NewId("T_Permissions") & "|RoleID=" & rid & "|Module=" & kv(0) & "|Niveau=" & kv(1))
            Next j
            Audit_Log "*", "UTIL", "CREATION_ROLE", "T_Roles", rid, "", "", parts(0) & " : " & parts(2), "Initialisation"
        End If
    Next i
End Sub
