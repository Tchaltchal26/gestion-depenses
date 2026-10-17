Attribute VB_Name = "modInscriptions"
'===============================================================================
' SG - Module modInscriptions
' Une inscription = un eleve, une annee, un etablissement. Elle porte la
' classe COURANTE ; l'historique complet des classes est dans T_Affectations
' (date d'effet distincte de la date d'operation, motif, auteur).
' Operations : premiere inscription, reinscription (passage / redoublement /
'   orientation), transfert entrant, changement de classe, suspension,
'   reintegration, depart, transfert sortant, radiation, fin de cycle,
'   annulation. Aucune operation ne supprime de ligne.
'===============================================================================
Option Explicit

Public Const TYPES_INSCRIPTION As String = "PREMIERE;REINSCRIPTION;TRANSFERT_ENTRANT"
Public Const TYPES_FIN As String = "DEPART;TRANSFERT_SORTANT;RADIATION;FIN_CYCLE"

Public Function Inscr_Active(ByVal etabId As String, ByVal eleveId As String, ByVal anneeId As String) As String
    Inscr_Active = DAL_First("T_Inscriptions", RecFrom("EtabID=" & etabId & "|EleveID=" & eleveId & "|AnneeID=" & anneeId & "|Statut=ACTIVE"))
End Function

Private Sub ControlerClasse(ByVal etabId As String, ByVal anneeId As String, ByVal classeId As String, ByVal forcerCapacite As Boolean)
    Dim cap As Long
    Etab_ExigerAppartenance "T_Classes", classeId, etabId
    If DAL_Value("T_Classes", classeId, "AnneeID") <> anneeId Then SG_Raise SG_ERR_VALIDATION, "La classe n'appartient pas a l'annee choisie."
    If DAL_Value("T_Classes", classeId, "Actif") <> "1" Then SG_Raise SG_ERR_ETAT, "Classe inactive."
    cap = CLng(DAL_Value("T_Classes", classeId, "Capacite"))
    If Classe_Effectif(classeId) >= cap Then
        If Param_Get(etabId, "CAPACITE_STRICTE", "1") = "1" Or Not forcerCapacite Then
            SG_Raise SG_ERR_ETAT, "Capacite de la classe atteinte (" & cap & ")."
        End If
    End If
End Sub

Private Function NouvelleAffectation(ByVal etabId As String, ByVal inscriptionId As String, ByVal classeId As String, _
                                     ByVal dateDebut As String, ByVal motif As String) As String
    Dim id As String, r As Variant
    id = DAL_NewId("T_Affectations")
    r = RecNew()
    RecSet r, "AffectationID", id
    RecSet r, "EtabID", etabId
    RecSet r, "InscriptionID", inscriptionId
    RecSet r, "ClasseID", classeId
    RecSet r, "DateDebut", dateDebut
    RecSet r, "DateFin", ""
    RecSet r, "Motif", motif
    RecSet r, "DateOperation", NowStamp()
    RecSet r, "Utilisateur", Sec_UtilisateurCourant()
    DAL_Insert "T_Affectations", r
    Audit_LogCreation etabId, "INSCR", "T_Affectations", r, motif
    NouvelleAffectation = id
End Function

Private Function DecisionPassage(ByVal etabId As String, ByVal eleveId As String, ByVal nouvelleClasse As String) As String
    ' Compare le niveau de la derniere inscription a celui de la nouvelle classe.
    Dim ids As Collection, i As Long, derniere As String, dDate As String, ins As Variant, oPrec As Long, oNouv As Long
    Set ids = DAL_Query("T_Inscriptions", RecFrom("EtabID=" & etabId & "|EleveID=" & eleveId))
    dDate = ""
    For i = 1 To ids.Count
        ins = DAL_Get("T_Inscriptions", ids(i))
        If RecStr(ins, "Statut") <> "ANNULEE" And RecStr(ins, "DateEffet") > dDate Then
            dDate = RecStr(ins, "DateEffet"): derniere = RecStr(ins, "ClasseID")
        End If
    Next i
    If Len(derniere) = 0 Then DecisionPassage = "": Exit Function
    oPrec = Niveau_Ordre(DAL_Value("T_Classes", derniere, "NiveauID"))
    oNouv = Niveau_Ordre(DAL_Value("T_Classes", nouvelleClasse, "NiveauID"))
    If oNouv = oPrec Then
        DecisionPassage = "REDOUBLEMENT"
    ElseIf oNouv = oPrec + 1 Then
        DecisionPassage = "PASSAGE"
    ElseIf oNouv > oPrec Then
        DecisionPassage = "SAUT_DE_CLASSE"
    Else
        DecisionPassage = "ORIENTATION"
    End If
End Function

' Inscription (premiere, reinscription ou transfert entrant).
Public Function Inscr_Inscrire(ByVal etabId As String, ByVal eleveId As String, ByVal anneeId As String, _
                               ByVal classeId As String, ByVal typeInscription As String, _
                               Optional ByVal dateEffet As Variant, Optional ByVal forcerCapacite As Boolean = False) As String
    ' Valeurs par defaut (IsMissing : forme compatible Excel et LibreOffice).
    If IsMissing(dateEffet) Then dateEffet = ""
    Dim id As String, r As Variant, statutEleve As String, decision As String, a As Variant
    Sec_Exiger etabId, "INSCR", NIV_ECRITURE
    Etab_ExigerAppartenance "T_Eleves", eleveId, etabId
    Annee_ExigerModifiable etabId, anneeId
    If Not InList(typeInscription, TYPES_INSCRIPTION) Then SG_Raise SG_ERR_VALIDATION, "Type d'inscription inconnu : " & typeInscription
    a = DAL_Get("T_Annees", anneeId)
    If Len(dateEffet) = 0 Then dateEffet = TodayISO()
    If Not IsISODate(dateEffet) Then SG_Raise SG_ERR_VALIDATION, "Date d'effet invalide."
    If dateEffet < RecStr(a, "DateDebut") Or dateEffet > RecStr(a, "DateFin") Then
        If dateEffet > RecStr(a, "DateFin") Or RecStr(a, "Statut") = "OUVERTE" Then
            SG_Raise SG_ERR_VALIDATION, "La date d'effet doit etre comprise dans l'annee scolaire."
        End If
        ' Annee en preparation : pre-inscription anticipee, effet au premier jour.
        dateEffet = RecStr(a, "DateDebut")
    End If
    statutEleve = DAL_Value("T_Eleves", eleveId, "Statut")
    If InList(statutEleve, "RADIE;TRANSFERE;DIPLOME") And typeInscription <> "REINSCRIPTION" Then
        SG_Raise SG_ERR_ETAT, "Dossier au statut " & statutEleve & " : utiliser la reinscription ou la reintegration."
    End If
    If statutEleve = "RADIE" Then SG_Raise SG_ERR_ETAT, "Eleve radie : une reintegration autorisee est necessaire."
    If Len(Inscr_Active(etabId, eleveId, anneeId)) > 0 Then SG_Raise SG_ERR_DOUBLON, "L'eleve a deja une inscription active pour cette annee."
    ControlerClasse etabId, anneeId, classeId, forcerCapacite
    If typeInscription = "REINSCRIPTION" Then
        If DAL_Count("T_Inscriptions", RecFrom("EtabID=" & etabId & "|EleveID=" & eleveId)) = 0 Then SG_Raise SG_ERR_ETAT, "Aucune inscription anterieure : utiliser la premiere inscription."
        decision = DecisionPassage(etabId, eleveId, classeId)
    End If

    Audit_DebutTx "Inscription"
    id = DAL_NewId("T_Inscriptions")
    r = RecNew()
    RecSet r, "InscriptionID", id
    RecSet r, "EtabID", etabId
    RecSet r, "AnneeID", anneeId
    RecSet r, "EleveID", eleveId
    RecSet r, "ClasseID", classeId
    RecSet r, "TypeInscription", typeInscription
    RecSet r, "Decision", decision
    RecSet r, "DateInscription", TodayISO()
    RecSet r, "DateEffet", dateEffet
    RecSet r, "Statut", "ACTIVE"
    RecSet r, "CreePar", Sec_UtilisateurCourant()
    RecSet r, "DateOperation", NowStamp()
    DAL_Insert "T_Inscriptions", r
    Audit_LogCreation etabId, "INSCR", "T_Inscriptions", r, ""
    NouvelleAffectation etabId, id, classeId, dateEffet, "Affectation initiale"
    Eleve_ChangerStatut etabId, eleveId, "ACTIF", "Inscription " & DAL_Value("T_Annees", anneeId, "Code")
    Audit_FinTx
    Inscr_Inscrire = id
End Function

Private Function AffectationCourante(ByVal etabId As String, ByVal inscriptionId As String) As String
    AffectationCourante = DAL_First("T_Affectations", RecFrom("EtabID=" & etabId & "|InscriptionID=" & inscriptionId & "|DateFin="))
End Function

Private Function ExigerInscriptionActive(ByVal etabId As String, ByVal inscriptionId As String) As Variant
    Dim ins As Variant
    Etab_ExigerAppartenance "T_Inscriptions", inscriptionId, etabId
    ins = DAL_Get("T_Inscriptions", inscriptionId)
    Annee_ExigerModifiable etabId, RecStr(ins, "AnneeID")
    If RecStr(ins, "Statut") <> "ACTIVE" Then SG_Raise SG_ERR_ETAT, "L'inscription n'est pas active (" & RecStr(ins, "Statut") & ")."
    ExigerInscriptionActive = ins
End Function

' Changement de classe ou de filiere en cours d'annee.
Public Sub Inscr_ChangerClasse(ByVal etabId As String, ByVal inscriptionId As String, ByVal nouvelleClasse As String, _
                               ByVal dateEffet As String, ByVal motif As String, Optional ByVal forcerCapacite As Boolean = False)
    Dim ins As Variant, aff As String, debutAff As String
    Sec_Exiger etabId, "INSCR", NIV_ECRITURE
    ins = ExigerInscriptionActive(etabId, inscriptionId)
    If Len(Trim$(motif)) = 0 Then SG_Raise SG_ERR_VALIDATION, "Motif obligatoire pour un changement de classe."
    If Not IsISODate(dateEffet) Then SG_Raise SG_ERR_VALIDATION, "Date d'effet invalide."
    If nouvelleClasse = RecStr(ins, "ClasseID") Then SG_Raise SG_ERR_VALIDATION, "L'eleve est deja dans cette classe."
    ControlerClasse etabId, RecStr(ins, "AnneeID"), nouvelleClasse, forcerCapacite
    aff = AffectationCourante(etabId, inscriptionId)
    debutAff = DAL_Value("T_Affectations", aff, "DateDebut")
    If dateEffet <= debutAff Then SG_Raise SG_ERR_VALIDATION, "La date d'effet doit suivre le debut de l'affectation actuelle (" & debutAff & ")."
    Audit_DebutTx "Changement de classe"
    DAL_Update "T_Affectations", aff, "DateFin", ISOAddDays(dateEffet, -1)
    Audit_Log etabId, "INSCR", "FIN_AFFECTATION", "T_Affectations", aff, "DateFin", "", ISOAddDays(dateEffet, -1), motif
    NouvelleAffectation etabId, inscriptionId, nouvelleClasse, dateEffet, motif
    DAL_Update "T_Inscriptions", inscriptionId, "ClasseID", nouvelleClasse
    Audit_Log etabId, "INSCR", "CHANGEMENT_CLASSE", "T_Inscriptions", inscriptionId, "ClasseID", RecStr(ins, "ClasseID"), nouvelleClasse, motif
    Audit_FinTx
End Sub

' Fin d'inscription : depart, transfert sortant, radiation, fin de cycle.
Public Sub Inscr_Terminer(ByVal etabId As String, ByVal inscriptionId As String, ByVal typeFin As String, _
                          ByVal dateFin As String, ByVal motif As String)
    Dim ins As Variant, aff As String, statut As String
    Sec_Exiger etabId, "INSCR", NIV_ECRITURE
    If typeFin = "RADIATION" Then Sec_Exiger etabId, "INSCR", NIV_VALIDATION
    ins = ExigerInscriptionActive(etabId, inscriptionId)
    If Not InList(typeFin, TYPES_FIN) Then SG_Raise SG_ERR_VALIDATION, "Type de fin inconnu : " & typeFin
    If Len(Trim$(motif)) = 0 Then SG_Raise SG_ERR_VALIDATION, "Motif obligatoire."
    If Not IsISODate(dateFin) Or dateFin < RecStr(ins, "DateEffet") Then SG_Raise SG_ERR_VALIDATION, "Date de fin invalide."
    Select Case typeFin
        Case "DEPART": statut = "PARTI"
        Case "TRANSFERT_SORTANT": statut = "TRANSFERE"
        Case "RADIATION": statut = "RADIE"
        Case "FIN_CYCLE": statut = "DIPLOME"
    End Select
    Audit_DebutTx "Fin d'inscription"
    DAL_Update "T_Inscriptions", inscriptionId, "Statut", "TERMINEE"
    DAL_Update "T_Inscriptions", inscriptionId, "DateFin", dateFin
    DAL_Update "T_Inscriptions", inscriptionId, "MotifFin", typeFin & " : " & motif
    Audit_Log etabId, "INSCR", typeFin, "T_Inscriptions", inscriptionId, "Statut", "ACTIVE", "TERMINEE", motif
    aff = AffectationCourante(etabId, inscriptionId)
    If Len(aff) > 0 Then
        DAL_Update "T_Affectations", aff, "DateFin", dateFin
        Audit_Log etabId, "INSCR", "FIN_AFFECTATION", "T_Affectations", aff, "DateFin", "", dateFin, motif
    End If
    Eleve_ChangerStatut etabId, RecStr(ins, "EleveID"), statut, motif
    Audit_FinTx
End Sub

Public Sub Inscr_Suspendre(ByVal etabId As String, ByVal inscriptionId As String, ByVal motif As String)
    Dim ins As Variant
    Sec_Exiger etabId, "INSCR", NIV_VALIDATION
    ins = ExigerInscriptionActive(etabId, inscriptionId)
    If Len(Trim$(motif)) = 0 Then SG_Raise SG_ERR_VALIDATION, "Motif obligatoire."
    Eleve_ChangerStatut etabId, RecStr(ins, "EleveID"), "SUSPENDU", motif
End Sub

' Reintegration : apres suspension (inscription active) ou apres radiation /
' depart (nouvelle inscription a creer ensuite).
Public Sub Inscr_Reintegrer(ByVal etabId As String, ByVal eleveId As String, ByVal motif As String)
    Dim statut As String
    Sec_Exiger etabId, "INSCR", NIV_VALIDATION
    Etab_ExigerAppartenance "T_Eleves", eleveId, etabId
    If Len(Trim$(motif)) = 0 Then SG_Raise SG_ERR_VALIDATION, "Motif obligatoire."
    statut = DAL_Value("T_Eleves", eleveId, "Statut")
    Select Case statut
        Case "SUSPENDU": Eleve_ChangerStatut etabId, eleveId, "ACTIF", "Reintegration : " & motif
        Case "RADIE", "PARTI": Eleve_ChangerStatut etabId, eleveId, "NOUVEAU", "Reintegration : " & motif
        Case Else: SG_Raise SG_ERR_ETAT, "Reintegration impossible depuis le statut " & statut & "."
    End Select
End Sub

' Annulation d'une inscription saisie par erreur : conservee avec statut ANNULEE.
Public Sub Inscr_Annuler(ByVal etabId As String, ByVal inscriptionId As String, ByVal motif As String)
    Dim ins As Variant, aff As String
    Sec_Exiger etabId, "INSCR", NIV_VALIDATION
    ins = ExigerInscriptionActive(etabId, inscriptionId)
    If Len(Trim$(motif)) < 5 Then SG_Raise SG_ERR_VALIDATION, "Motif detaille obligatoire pour une annulation."
    Audit_DebutTx "Annulation inscription"
    DAL_Update "T_Inscriptions", inscriptionId, "Statut", "ANNULEE"
    DAL_Update "T_Inscriptions", inscriptionId, "MotifFin", "ANNULATION : " & motif
    Audit_Log etabId, "INSCR", "ANNULATION", "T_Inscriptions", inscriptionId, "Statut", "ACTIVE", "ANNULEE", motif
    aff = AffectationCourante(etabId, inscriptionId)
    If Len(aff) > 0 Then
        DAL_Update "T_Affectations", aff, "DateFin", RecStr(ins, "DateEffet")
        Audit_Log etabId, "INSCR", "FIN_AFFECTATION", "T_Affectations", aff, "DateFin", "", RecStr(ins, "DateEffet"), "Annulation"
    End If
    If DAL_Count("T_Inscriptions", RecFrom("EtabID=" & etabId & "|EleveID=" & RecStr(ins, "EleveID") & "|Statut=ACTIVE")) = 0 Then
        Eleve_ChangerStatut etabId, RecStr(ins, "EleveID"), "NOUVEAU", "Annulation de l'inscription"
    End If
    Audit_FinTx
End Sub

' Historique des affectations d'une inscription (AffectationID, chronologique).
Public Function Inscr_Historique(ByVal etabId As String, ByVal inscriptionId As String) As Collection
    Sec_Exiger etabId, "INSCR", NIV_LECTURE
    Etab_ExigerAppartenance "T_Inscriptions", inscriptionId, etabId
    Set Inscr_Historique = DAL_Query("T_Affectations", RecFrom("EtabID=" & etabId & "|InscriptionID=" & inscriptionId))
End Function

'-------------------------------------------------------------------------------
' Transfert entre deux etablissements geres par la meme installation
' (mode centralise). L'etablissement d'accueil recoit un NOUVEAU dossier
' contenant uniquement l'identite ; l'historique reste dans l'etablissement
' d'origine. Le lien est conserve (EleveOrigineID + T_Transferts).
'-------------------------------------------------------------------------------
Public Function Transfert_VersEtablissement(ByVal etabSource As String, ByVal inscriptionId As String, _
                                            ByVal etabCible As String, ByVal dateTransfert As String, ByVal motif As String) As String
    Dim ins As Variant, e As Variant, r As Variant, nouvelId As String, tid As String
    If etabSource = etabCible Then SG_Raise SG_ERR_VALIDATION, "Etablissements source et cible identiques."
    Sec_Exiger etabSource, "INSCR", NIV_VALIDATION
    Sec_Exiger etabCible, "ELEVES", NIV_ECRITURE
    ins = ExigerInscriptionActive(etabSource, inscriptionId)
    e = DAL_Get("T_Eleves", RecStr(ins, "EleveID"))
    ' Tous les controles avant la premiere ecriture.
    If Len(Trim$(motif)) = 0 Then SG_Raise SG_ERR_VALIDATION, "Motif obligatoire."
    If Not IsISODate(dateTransfert) Or dateTransfert < RecStr(ins, "DateEffet") Then SG_Raise SG_ERR_VALIDATION, "Date de transfert invalide."
    If Not DAL_Exists("T_Etablissements", etabCible) Then SG_Raise SG_ERR_INTROUVABLE, "Etablissement d'accueil inconnu."
    Audit_DebutTx "Transfert inter-etablissements"
    r = RecNew()
    RecSet r, "Nom", RecStr(e, "Nom")
    RecSet r, "Prenoms", RecStr(e, "Prenoms")
    RecSet r, "Sexe", RecStr(e, "Sexe")
    RecSet r, "DateNaissance", RecStr(e, "DateNaissance")
    RecSet r, "LieuNaissance", RecStr(e, "LieuNaissance")
    RecSet r, "Nationalite", RecStr(e, "Nationalite")
    RecSet r, "EtabPrecedent", Etab_Code(etabSource)
    RecSet r, "EleveOrigineID", RecStr(e, "EleveID")
    nouvelId = Eleve_Creer(etabCible, r, True, "Transfert depuis " & Etab_Code(etabSource))
    Inscr_Terminer etabSource, inscriptionId, "TRANSFERT_SORTANT", dateTransfert, "Vers " & Etab_Code(etabCible) & " : " & motif
    tid = DAL_NewId("T_Transferts")
    r = RecNew()
    RecSet r, "TransfertID", tid
    RecSet r, "EtabSource", etabSource
    RecSet r, "EleveSource", RecStr(e, "EleveID")
    RecSet r, "EtabCible", etabCible
    RecSet r, "EleveCible", nouvelId
    RecSet r, "DateTransfert", dateTransfert
    RecSet r, "Motif", motif
    RecSet r, "Utilisateur", Sec_UtilisateurCourant()
    DAL_Insert "T_Transferts", r
    Audit_LogCreation etabSource, "INSCR", "T_Transferts", r, motif
    Audit_Log etabCible, "INSCR", "TRANSFERT_RECU", "T_Transferts", tid, "", "", nouvelId, motif
    Audit_FinTx
    Transfert_VersEtablissement = nouvelId
End Function
