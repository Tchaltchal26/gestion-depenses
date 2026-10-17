Attribute VB_Name = "modRapports"
'===============================================================================
' SG - Module modRapports (indicateurs, controle qualite, documents)
' Tous les indicateurs sont recalcules a partir des tables sources (aucune
' valeur saisie a la main). Le controle qualite SIGNALE les anomalies et
' propose une correction ; il ne corrige jamais automatiquement.
'===============================================================================
Option Explicit

Private Sub Incr(r As Variant, ByVal champ As String)
    RecSet r, champ, CLng(RecGet(r, champ, 0)) + 1
End Sub

Public Function Rapport_Indicateurs(ByVal etabId As String, ByVal anneeId As String) As Variant
    Dim r As Variant, ins As Collection, el As Collection, sexes As New Collection, i As Long, x As Variant, motif As String
    Sec_Exiger etabId, "INSCR", NIV_LECTURE
    Etab_ExigerAppartenance "T_Annees", anneeId, etabId
    r = RecNew()
    For Each x In Split("Inscrits;Premieres;Reinscriptions;TransfertsEntrants;Passages;Redoublements;Departs;TransfertsSortants;Radiations;FinsCycle;Annulees;Filles;Garcons", ";")
        RecSet r, CStr(x), 0
    Next x
    Set el = DAL_QueryRecs("T_Eleves", RecFrom("EtabID=" & etabId))
    For i = 1 To el.Count
        sexes.Add RecStr(el(i), "Sexe"), RecStr(el(i), "EleveID")
    Next i
    Set ins = DAL_QueryRecs("T_Inscriptions", RecFrom("EtabID=" & etabId & "|AnneeID=" & anneeId))
    For i = 1 To ins.Count
        x = ins(i)
        Select Case RecStr(x, "Statut")
            Case "ACTIVE"
                Incr r, "Inscrits"
                If sexes(RecStr(x, "EleveID")) = "F" Then
                    Incr r, "Filles"
                Else
                    Incr r, "Garcons"
                End If
            Case "ANNULEE"
                Incr r, "Annulees"
            Case "TERMINEE"
                motif = RecStr(x, "MotifFin")
                If Left$(motif, 6) = "DEPART" Then Incr r, "Departs"
                If Left$(motif, 17) = "TRANSFERT_SORTANT" Then Incr r, "TransfertsSortants"
                If Left$(motif, 9) = "RADIATION" Then Incr r, "Radiations"
                If Left$(motif, 9) = "FIN_CYCLE" Then Incr r, "FinsCycle"
        End Select
        If RecStr(x, "Statut") <> "ANNULEE" Then
            Select Case RecStr(x, "TypeInscription")
                Case "PREMIERE": Incr r, "Premieres"
                Case "REINSCRIPTION": Incr r, "Reinscriptions"
                Case "TRANSFERT_ENTRANT": Incr r, "TransfertsEntrants"
            End Select
            If RecStr(x, "Decision") = "PASSAGE" Then Incr r, "Passages"
            If RecStr(x, "Decision") = "REDOUBLEMENT" Then Incr r, "Redoublements"
        End If
    Next i
    Rapport_Indicateurs = r
End Function

' Tableau 2D (ligne 0 = en-tetes) : Code, Libelle, Niveau, Effectif, Capacite, Taux, Alerte.
Public Function Rapport_EffectifsClasses(ByVal etabId As String, ByVal anneeId As String) As Variant
    Dim cls As Collection, out() As Variant, i As Long, c As Variant, eff As Long, cap As Long
    Sec_Exiger etabId, "INSCR", NIV_LECTURE
    Etab_ExigerAppartenance "T_Annees", anneeId, etabId
    Set cls = DAL_QueryRecs("T_Classes", RecFrom("EtabID=" & etabId & "|AnneeID=" & anneeId))
    ReDim out(0 To cls.Count, 0 To 6)
    out(0, 0) = "Classe": out(0, 1) = "Libelle": out(0, 2) = "Niveau": out(0, 3) = "Effectif"
    out(0, 4) = "Capacite": out(0, 5) = "Taux": out(0, 6) = "Alerte"
    For i = 1 To cls.Count
        c = cls(i)
        eff = Classe_Effectif(RecStr(c, "ClasseID"))
        cap = CLng("0" & RecStr(c, "Capacite"))
        out(i, 0) = RecStr(c, "Code")
        out(i, 1) = RecStr(c, "Libelle")
        out(i, 2) = DAL_Value("T_Niveaux", RecStr(c, "NiveauID"), "Code")
        out(i, 3) = eff
        out(i, 4) = cap
        If cap > 0 Then out(i, 5) = Round(eff / cap, 3) Else out(i, 5) = 0
        If eff > cap Then
            out(i, 6) = "SUREFFECTIF"
        ElseIf eff = cap Then
            out(i, 6) = "COMPLET"
        Else
            out(i, 6) = ""
        End If
    Next i
    Rapport_EffectifsClasses = out
End Function

'-------------------------------------------------------------------------------
' Controle qualite des donnees
' Chaque anomalie : "GRAVITE|CATEGORIE|ENREGISTREMENT|DESCRIPTION|CORRECTION PROPOSEE"
'-------------------------------------------------------------------------------
Private Sub Anomalie(res As Collection, ByVal gravite As String, ByVal categorie As String, ByVal enreg As String, _
                     ByVal descr As String, ByVal correction As String)
    res.Add gravite & "|" & categorie & "|" & enreg & "|" & descr & "|" & correction
End Sub

Private Function Contient(c As Collection, ByVal cle As String) As Boolean
    Dim v As Variant
    On Error GoTo non
    v = c(cle)
    Contient = True
    Exit Function
non:
    Contient = False
End Function

Public Function QC_Anomalies(ByVal etabId As String) As Collection
    Dim res As New Collection, eleves As Collection, ins As Collection, cls As Collection, affs As Collection
    Dim annees As Collection, i As Long, k As Long, x As Variant, st As String
    Dim idxEleve As New Collection, idxClasse As New Collection, idxAnnee As New Collection, mats As New Collection
    Dim actifsInscrits As New Collection, affOuvertes As New Collection, n As Long, r As String, lic As Variant
    Sec_Exiger etabId, "ELEVES", NIV_LECTURE

    Set annees = DAL_QueryRecs("T_Annees", RecFrom("EtabID=" & etabId))
    For i = 1 To annees.Count
        idxAnnee.Add RecStr(annees(i), "Statut"), RecStr(annees(i), "AnneeID")
    Next i
    Set cls = DAL_QueryRecs("T_Classes", RecFrom("EtabID=" & etabId))
    For i = 1 To cls.Count
        idxClasse.Add RecStr(cls(i), "AnneeID"), RecStr(cls(i), "ClasseID")
    Next i
    Set eleves = DAL_QueryRecs("T_Eleves", RecFrom("EtabID=" & etabId))
    For i = 1 To eleves.Count
        x = eleves(i)
        idxEleve.Add RecStr(x, "Statut"), RecStr(x, "EleveID")
        If Contient(mats, RecStr(x, "Matricule")) Then
            Anomalie res, "CRITIQUE", "DOUBLON", RecStr(x, "EleveID"), "Matricule " & RecStr(x, "Matricule") & " attribue deux fois.", "Verifier les dossiers et attribuer un nouveau matricule par une procedure autorisee."
        Else
            mats.Add 1, RecStr(x, "Matricule")
        End If
        If Not IsISODate(RecStr(x, "DateNaissance")) Then
            Anomalie res, "MOYENNE", "DATE", RecStr(x, "EleveID"), "Date de naissance invalide.", "Corriger la date avec un motif."
        End If
        If Len(DAL_First("T_LiensResponsables", RecFrom("EtabID=" & etabId & "|EleveID=" & RecStr(x, "EleveID") & "|Legal=1|Actif=1"))) = 0 And RecStr(x, "Statut") = "ACTIF" Then
            Anomalie res, "INFO", "RESPONSABLE", RecStr(x, "EleveID"), "Aucun responsable legal enregistre.", "Completer le dossier."
        End If
    Next i

    Set affs = DAL_QueryRecs("T_Affectations", RecFrom("EtabID=" & etabId))
    For i = 1 To affs.Count
        x = affs(i)
        If Len(RecStr(x, "DateFin")) = 0 Then
            If Contient(affOuvertes, RecStr(x, "InscriptionID")) Then
                Anomalie res, "HAUTE", "AFFECTATION", RecStr(x, "InscriptionID"), "Plusieurs affectations ouvertes pour une meme inscription.", "Clore l'affectation erronee par un changement de classe motive."
            Else
                affOuvertes.Add 1, RecStr(x, "InscriptionID")
            End If
        End If
    Next i

    Set ins = DAL_QueryRecs("T_Inscriptions", RecFrom("EtabID=" & etabId))
    For i = 1 To ins.Count
        x = ins(i)
        If Not Contient(idxEleve, RecStr(x, "EleveID")) Then
            Anomalie res, "CRITIQUE", "ORPHELIN", RecStr(x, "InscriptionID"), "Inscription liee a un dossier eleve inexistant.", "Restaurer le dossier depuis une sauvegarde verifiee."
        ElseIf Not Contient(idxClasse, RecStr(x, "ClasseID")) Then
            Anomalie res, "CRITIQUE", "ORPHELIN", RecStr(x, "InscriptionID"), "Inscription liee a une classe inexistante.", "Restaurer la classe ou reaffecter l'eleve."
        ElseIf Not Contient(idxAnnee, RecStr(x, "AnneeID")) Then
            Anomalie res, "CRITIQUE", "ORPHELIN", RecStr(x, "InscriptionID"), "Inscription liee a une annee inexistante.", "Restaurer depuis une sauvegarde verifiee."
        Else
            If idxClasse(RecStr(x, "ClasseID")) <> RecStr(x, "AnneeID") Then
                Anomalie res, "HAUTE", "INCOHERENCE", RecStr(x, "InscriptionID"), "La classe n'appartient pas a l'annee de l'inscription.", "Reaffecter l'eleve dans une classe de la bonne annee."
            End If
            If RecStr(x, "Statut") = "ACTIVE" Then
                st = idxEleve(RecStr(x, "EleveID"))
                If InList(st, "PARTI;RADIE;TRANSFERE;DIPLOME") Then
                    Anomalie res, "HAUTE", "STATUT", RecStr(x, "InscriptionID"), "Inscription active alors que l'eleve est " & st & ".", "Terminer l'inscription ou reintegrer l'eleve."
                End If
                If Not Contient(affOuvertes, RecStr(x, "InscriptionID")) Then
                    Anomalie res, "HAUTE", "AFFECTATION", RecStr(x, "InscriptionID"), "Inscription active sans affectation en cours.", "Creer l'affectation par un changement de classe motive."
                End If
                If InList(idxAnnee(RecStr(x, "AnneeID")), "OUVERTE;PREPARATION") And Not Contient(actifsInscrits, RecStr(x, "EleveID")) Then
                    actifsInscrits.Add 1, RecStr(x, "EleveID")
                End If
            End If
        End If
    Next i
    For i = 1 To eleves.Count
        x = eleves(i)
        If RecStr(x, "Statut") = "ACTIF" And Not Contient(actifsInscrits, RecStr(x, "EleveID")) Then
            Anomalie res, "MOYENNE", "INSCRIPTION", RecStr(x, "EleveID"), "Eleve actif sans inscription dans l'annee en cours.", "Inscrire l'eleve ou enregistrer son depart."
        End If
    Next i
    For i = 1 To cls.Count
        n = Classe_Effectif(RecStr(cls(i), "ClasseID"))
        If n > CLng("0" & RecStr(cls(i), "Capacite")) Then
            Anomalie res, "MOYENNE", "CAPACITE", RecStr(cls(i), "ClasseID"), "Classe " & RecStr(cls(i), "Code") & " en sureffectif (" & n & "/" & RecStr(cls(i), "Capacite") & ").", "Ouvrir une classe ou ajuster la capacite."
        End If
    Next i

    r = Audit_Verifier()
    If Left$(r, 2) <> "OK" Then Anomalie res, "CRITIQUE", "AUDIT", Split(r, "|")(1), "Journal d'audit altere : " & Split(r, "|")(2) & ".", "Ne plus modifier les donnees ; comparer avec la derniere sauvegarde verifiee."
    lic = Lic_Etat()
    If RecStr(lic, "Ecriture") <> "1" Then
        Anomalie res, "HAUTE", "LICENCE", "", RecStr(lic, "Message"), "Contacter l'editeur pour renouveler ou transferer la licence."
    ElseIf Len(RecStr(lic, "Alerte")) > 0 Then
        Anomalie res, "INFO", "LICENCE", "", RecStr(lic, "Alerte"), "Planifier le renouvellement."
    End If
    Set QC_Anomalies = res
End Function

'-------------------------------------------------------------------------------
' Documents
'-------------------------------------------------------------------------------
' Donnees d'un certificat de scolarite numerote (le modele imprimable est
' rempli par modUI). L'emission est journalisee avec son numero.
Public Function Doc_CertificatScolarite(ByVal etabId As String, ByVal eleveId As String) As Variant
    Dim e As Variant, et As Variant, annee As String, ins As String, r As Variant, num As String, i As Variant
    Sec_Exiger etabId, "ELEVES", NIV_LECTURE
    Etab_ExigerAppartenance "T_Eleves", eleveId, etabId
    annee = Annee_Ouverte(etabId)
    If Len(annee) = 0 Then SG_Raise SG_ERR_ETAT, "Aucune annee scolaire ouverte."
    ins = Inscr_Active(etabId, eleveId, annee)
    If Len(ins) = 0 Then SG_Raise SG_ERR_ETAT, "L'eleve n'est pas inscrit pour l'annee en cours : certificat impossible."
    e = DAL_Get("T_Eleves", eleveId)
    et = DAL_Get("T_Etablissements", etabId)
    num = Format_Numero(Param_Get(etabId, "FORMAT_CERTIFICAT", "CS-{ETAB}-{AAAA}-{N5}"), RecStr(et, "Code"), CStr(Year(Date)), Compteur_Suivant(etabId, "CERTIFICAT"))
    r = RecNew()
    RecSet r, "Numero", num
    RecSet r, "DateEmission", TodayISO()
    RecSet r, "Etablissement", RecStr(et, "Nom")
    RecSet r, "Ville", RecStr(et, "Ville")
    For Each i In Split("Matricule;Nom;Prenoms;Sexe;DateNaissance;LieuNaissance", ";")
        RecSet r, CStr(i), RecStr(e, CStr(i))
    Next i
    RecSet r, "Classe", DAL_Value("T_Classes", DAL_Value("T_Inscriptions", ins, "ClasseID"), "Libelle")
    RecSet r, "Annee", DAL_Value("T_Annees", annee, "Code")
    RecSet r, "Signataire", Sec_LoginCourant()
    Audit_Log etabId, "ELEVES", "DOCUMENT_EMIS", "T_Eleves", eleveId, "CertificatScolarite", "", num, ""
    Doc_CertificatScolarite = r
End Function

' Liste d'une classe (eleves inscrits actifs), triee par nom puis prenoms.
' Tableau 2D : ligne 0 = en-tetes (N, Matricule, Nom, Prenoms, Sexe, DateNaissance).
Public Function Liste_Classe(ByVal etabId As String, ByVal classeId As String) As Variant
    Dim ins As Collection, out() As Variant, i As Long, j As Long, k As Long, e As Variant, n As Long, tmp As Variant
    Sec_Exiger etabId, "ELEVES", NIV_LECTURE
    Etab_ExigerAppartenance "T_Classes", classeId, etabId
    Set ins = DAL_QueryRecs("T_Inscriptions", RecFrom("EtabID=" & etabId & "|ClasseID=" & classeId & "|Statut=ACTIVE"))
    n = ins.Count
    ReDim out(0 To n, 0 To 5)
    out(0, 0) = "N": out(0, 1) = "Matricule": out(0, 2) = "Nom": out(0, 3) = "Prenoms": out(0, 4) = "Sexe": out(0, 5) = "Date de naissance"
    For i = 1 To n
        e = DAL_Get("T_Eleves", RecStr(ins(i), "EleveID"))
        out(i, 1) = RecStr(e, "Matricule"): out(i, 2) = RecStr(e, "Nom"): out(i, 3) = RecStr(e, "Prenoms")
        out(i, 4) = RecStr(e, "Sexe"): out(i, 5) = RecStr(e, "DateNaissance")
    Next i
    ' Tri par insertion (effectifs de classe : quelques dizaines de lignes).
    For i = 2 To n
        j = i
        Do While j > 1
            If NormalizeText(out(j - 1, 2) & " " & out(j - 1, 3)) <= NormalizeText(out(j, 2) & " " & out(j, 3)) Then Exit Do
            For k = 1 To 5
                tmp = out(j, k): out(j, k) = out(j - 1, k): out(j - 1, k) = tmp
            Next k
            j = j - 1
        Loop
    Next i
    For i = 1 To n
        out(i, 0) = i
    Next i
    Liste_Classe = out
End Function
