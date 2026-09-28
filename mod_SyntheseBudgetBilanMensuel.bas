Option Explicit

' ------------------------------------------------------------------------
' Budget_Bilan_Mensuel : calcule et affiche le total des entrées, le total
' des dépenses, et la différence entre les deux, pour la période B1:B2.
' ------------------------------------------------------------------------
Public Sub Budget_Bilan_Mensuel()
    Dim totalDepenses As Double
    Dim totalEntrees As Double
    Dim Montant As Double
    Dim nbLigne As Long
    Dim vbComp As Object
    Dim codeMod As Object
    Dim codeStr As String
    Dim debPlageTravail As Range
    Dim reponse As VbMsgBoxResult
    
    ' 1. Récupération des pointeurs vers la feuille et le tableau
    Set wsSynthese = mod_Criteres.GetFeuille(NOM_FEUILLE_SYNTHESE)
    Set wsResultat = mod_Criteres.GetFeuille(NOM_FEUILLE_RESULTAT)
    Set tbl = mod_DonneesTable.GetOperationsValue(NOM_FEUILLE_DONNEES, "TblOperations")
    If tbl Is Nothing Then
        MsgBox "Le tableau ne contient aucune ligne de données.", vbExclamation
        Exit Sub
    End If
    If tbl.DataBodyRange Is Nothing Then
        MsgBox "Le tableau est vide.", vbExclamation
        Exit Sub
    End If
    
    ' 1. Verifier si la feuille de travail existe deja, pour eviter d'ecraser du travail sans prevenir.
    If Not wsResultat Is Nothing Then
        ' La feuille existe deja : on demande confirmation avant de tout reconstruire,
        ' car cela va effacer sa mise en forme actuelle.
        reponse = MsgBox(FR("La feuille '" & NOM_FEUILLE_RESULTAT & "' existe deja." & vbCrLf & _
                            "Voulez-vous la reconstruire enti{e1}rement (sa mise en forme actuelle sera perdue) ?"), _
                            vbYesNo + vbQuestion, "Confirmation de reconstruction")
        If reponse = vbNo Then
            MsgBox FR("Installation annul{e2}e, aucune modification effectu{e2}e."), vbInformation
            Exit Sub
        End If
        
        ' On la rend visible temporairement : impossible de la modifier/supprimer
        ' proprement tant qu'elle est en xlSheetVeryHidden.
        wsResultat.Visible = xlSheetVisible
        wsResultat.Cells.Clear
        Call mod_Display.SupprimerFormesExistantesFN(wsResultat)
        Call mod_Display.SupprimerNomsExistantsFN(wsResultat, NOM_FEUILLE_RESULTAT)
    Else
        ' La feuille n'existe pas encore : on la cree, positionnee en derniere position
        ' pour ne pas perturber l'ordre des onglets existants (Accueil, Synthese...).
        Set wsResultat = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        wsResultat.Name = NOM_FEUILLE_RESULTAT
    End If
    
    ' 4. Contrairement aux autres macros, celle-ci REACTIVE le double-clic
    ' (c'est depuis cet écran de bilan que l'utilisateur double-clique sur
    ' un total pour en voir le détail -- voir ShowDetailForTotal ci-dessous).
    AllowDetailDoubleClick = True
    ' Paramètre de navigation
    RecherOperations = False
    ' Désactive les événements
    Application.EnableEvents = False
        
    ' On masque la feuille Synthese. Cette option est prise pour éviter à l'opérateur de se déplacer en dehors de la feuille crée
    wsSynthese.Visible = xlSheetVeryHidden
    
    ' 2. Charge TOUT le tableau de données (en-têtes incluses en ligne 1)
    tblData = tbl.Range.value
    tblDataLineTotal = UBound(tblData, 1)
    
    ' 3. Récupération des index de colonne dans la base de données
    mod_Display.RecupIndexCol
    
    ' 5. Construction de la zone des boutons
    ' On détermine la position du bouton
    Set zoneBouton = wsResultat.Range(cellSortieDep).Offset(0, 0)
    
    'On définit son titre
    texteBouton = "Sortir"
    nomMacroBouton = "Sortir"
    Call mod_Display.ConstruireBoutons(wsResultat, zoneBouton, texteBouton, nomMacroBouton)

    ' 6. Construction de la zone des instructions
    ' On fournit le message à afficher en remplaçant les carractères accentué par les balise de la fonction FR
    Message = " Bilan Mensuel" & Chr(10) & _
                "Affiche une synth{e2}se des dépenses et des revenus sur la p{e2}riode du:" & Chr(10) & _
                "Mois: " & critMois & " Ann{e2}e: " & critAnnee & Chr(10) & _
                 "{A2} la fin sortez avec le bouton ""Sortir"""
                 
    ' On fournit le titre, la position de la cellule dans laquelle on veut écrire de titre
    Titre = "Instructions"
    Set PositionTitre = wsResultat.Range(cellSortieDep).Offset(2, 0)
    
    ' On détermine la plage de début et de fin de la zone
    Set debutZone = wsResultat.Range(cellSortieDep).Offset(3, 0)
    Set finZone = wsResultat.Range(cellSortieDep).Offset(6, 6)
    
    Call mod_Display.ConstruireZoneTexte(wsResultat, Titre, Message, PositionTitre, debutZone, finZone)
    
    ' 7. Lire les critères saisis par l'utilisateur (B1 à B6).
    mod_Criteres.GetSelectCriteres
    
    ' 8. Nettoyage de la zone de sortie
    Set debPlageTravail = wsResultat.Range(cellSortieDep).Offset(7, 0)
    Call mod_Display.PrepareOutputArea(wsSynthese, , debPlageTravail)

    
    ' Taille maximale du tableau de résultat = nombre total de lignes source
    ReDim tabResultat(1 To 3, 1 To 2)
    totalDepenses = 0
    totalEntrees = 0

    ' 5. On parcourt toute la table et on additionne, ligne par ligne, entrées
    ' et dépenses du mois/année sélectionnés.
    For nbLigne = 2 To tblDataLineTotal
        If mod_DonneesTable.RowMatchesFilter(nbLigne, tblData(nbLigne, colMoisBud), tblData(nbLigne, colAnneeBud), tblData(nbLigne, colMontant), True) Then
            Montant = tblData(nbLigne, colMontant)
            If Montant < 0 Then
                ' Une dépense est additionnée en valeur positive (voir mod_DonneesTable
                ' pour la même logique appliquée au filtrage par seuil).
                totalDepenses = totalDepenses + Abs(Montant)
            ElseIf Montant >= 0 Then
                totalEntrees = totalEntrees + Montant
            End If
        End If
    Next nbLigne

    ' On définit la plage de sortie
    'Set plageSortieEnTetes = wsResultat.Range(cellSortieDep).Resize(9, 1)
    ' On efface l'ancien affichage avant d'écrire le nouveau bilan.
    wsResultat.Range(cellSortieDep).Offset(9, 1) = "Total Entrées"
    wsResultat.Range(cellSortieDep).Offset(10, 1) = "Total Dépenses"
    wsResultat.Range(cellSortieDep).Offset(11, 1) = "Différence (Entrées - Dépenses)"
    wsResultat.Range(cellSortieDep).Offset(9, 2) = totalEntrees
    wsResultat.Range(cellSortieDep).Offset(10, 2) = totalDepenses
    wsResultat.Range(cellSortieDep).Offset(11, 2) = totalEntrees - totalDepenses

    wsResultat.Range(cellSortieDep).Offset(9, 2).NumberFormat = "#,##0.00"
    wsResultat.Range(cellSortieDep).Offset(10, 2).NumberFormat = "#,##0.00"
    wsResultat.Range(cellSortieDep).Offset(11, 2).NumberFormat = "#,##0.00"
    wsResultat.Columns.AutoFit
    wsResultat.rows.AutoFit
    
    ' Réactive les événements
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    'AllowDetailDoubleClick = True
    MsgBox "Bilan calculé : " & Format(totalEntrees - totalDepenses, "#,##0.00"), vbInformation
End Sub

' ------------------------------------------------------------------------
' ShowDetailForTotal : crée une nouvelle feuille contenant le détail complet
' (Positif ou Négatif) de la période B1:B2, triée selon les paramètres
' demandés. Appelée automatiquement par le double-clic défini dans ThisWorkbook
' ------------------------------------------------------------------------
Public Sub ShowDetailForTotal(ByVal showType As String, Optional ByVal sortField As String = "Date", Optional ByVal sortOrder As String = "Croissant")
    Dim wsDetail As Worksheet
    Dim nomFeuille As String
    Dim critAnnee As String
    Dim critMois As String
    Dim ligneAffichage As Long
    Dim nbLigne As Long, idxRes As Long
    Dim sortColumn As Long
    Dim sortOrderValue As Long
    Dim debPlageTravail As Range
    Dim reponse As VbMsgBoxResult
    
    ' 1. Récupération des pointeurs vers la feuille et le tableau
    Set wsSynthese = mod_Criteres.GetFeuille(NOM_FEUILLE_SYNTHESE)
    Set wsResultat = mod_Criteres.GetFeuille(NOM_FEUILLE_RESULTAT)
    Set tbl = mod_DonneesTable.GetOperationsValue(NOM_FEUILLE_DONNEES, "TblOperations")
    If tbl Is Nothing Then
        MsgBox "Le tableau ne contient aucune ligne de données.", vbExclamation
        Exit Sub
    End If
    If tbl.DataBodyRange Is Nothing Then
        MsgBox "Le tableau est vide.", vbExclamation
        Exit Sub
    End If

    nomFeuille = "frm_" & FormaterChaine("Détail_" & showType)
    Set wsDetail = mod_Criteres.GetFeuille(nomFeuille)
    
    ' 1. Verifier si la feuille de travail existe deja, pour eviter d'ecraser du travail sans prevenir.
    If Not wsDetail Is Nothing Then
        ' La feuille existe deja : on demande confirmation avant de tout reconstruire,
        ' car cela va effacer sa mise en forme actuelle.
        reponse = MsgBox(FR("La feuille '" & nomFeuille & "' existe deja." & vbCrLf & _
                            "Voulez-vous la reconstruire enti{e1}rement (sa mise en forme actuelle sera perdue) ?"), _
                            vbYesNo + vbQuestion, "Confirmation de reconstruction")
        If reponse = vbNo Then
            MsgBox FR("Installation annul{e2}e, aucune modification effectu{e2}e."), vbInformation
            Exit Sub
        End If
        
        ' On la rend visible temporairement : impossible de la modifier/supprimer
        ' proprement tant qu'elle est en xlSheetVeryHidden.
        wsDetail.Visible = xlSheetVisible
        wsDetail.Cells.Clear
        Call mod_Display.SupprimerFormesExistantesFN(wsDetail)
        Call mod_Display.SupprimerNomsExistantsFN(wsDetail, nomFeuille)
    Else
        ' La feuille n'existe pas encore : on la cree, positionnee en derniere position
        ' pour ne pas perturber l'ordre des onglets existants (Accueil, Synthese...).
        Set wsDetail = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        wsDetail.Name = nomFeuille
    End If
    
    ' On masque la feuille Synthese. Cette option est prise pour éviter à l'opérateur de se déplacer en dehors de la feuille crée
    wsResultat.Visible = xlSheetVeryHidden
    
    ' 2. Charge TOUT le tableau de données (en-têtes incluses en ligne 1)
    tblData = tbl.Range.value
    tblDataLineTotal = UBound(tblData, 1)
    
    ' Paramètre de navigation
    RecherOperations = False
    
    ' 3. Récupération des index de colonne dans la base de données
    mod_Display.RecupIndexCol

    ' 4. Lire les critères saisis par l'utilisateur (B1 à B6).
    mod_Criteres.GetSelectCriteres

    ' 5. Construction de la zone des boutons
    ' On détermine la position du bouton
    Set zoneBouton = wsDetail.Range(cellSortieDep).Offset(0, 0)
    
    'On définit son titre
    texteBouton = "Sortir"
    nomMacroBouton = "Sortir"
    Call mod_Display.ConstruireBoutons(wsDetail, zoneBouton, texteBouton, nomMacroBouton)

    ' 6. Construction de la zone des instructions
    ' On fournit le message à afficher en remplaçant les carractères accentué par les balise de la fonction FR
    Message = nomFeuille & Chr(10) & _
                "Affiche toutes les op{e2}rations de la p{e2}riode sélectionn{e2}e et de type " & showType & Chr(10) & _
                 "{A2} la fin sortez avec le bouton ""Sortir"""
                 
    ' On fournit le titre, la position de la cellule dans laquelle on veut écrire le titre
    Titre = "Instructions"
    Set PositionTitre = wsDetail.Range(cellSortieDep).Offset(2, 0)
    
    ' On détermine la plage de début et de fin de la zone
    Set debutZone = wsDetail.Range(cellSortieDep).Offset(3, 0)
    Set finZone = wsDetail.Range(cellSortieDep).Offset(5, 5)
    
    Call mod_Display.ConstruireZoneTexte(wsDetail, Titre, Message, PositionTitre, debutZone, finZone)
    
    ' 7. Affiche les entêtes du tableau de sortie
    ' Définit le début de la page de travail
    Set debPlageTravail = wsDetail.Range(cellSortieDep).Offset(6, 0)
    
    MonArray = Array("Date", "Tiers", "Catégorie", "Montant", "Notes", "Budget")
    Call mod_Display.PrepareOutputArea(wsDetail, MonArray, debPlageTravail)
    
    nbColonne = UBound(MonArray) - LBound(MonArray) + 1
    
    ' wsDetail.Range("A1:G1").value = Array("Date", "Libellé", "Catégorie", "Montant", "MoisBudget", "AnnéeBudget", "Budget")

    sortColumn = mod_Display.PosSortieIndex(wsDetail, MonArray, critTriChamps)
    sortOrderValue = mod_Actions.ResolveDetailSortOrder(critTriOrdre)
    
    ' Récupére la position du champs dans l'ARRAY (Application.Match est naturellement insensible à la casse)
    mod_Display.RecupPosArray
    
    ' Récupération des index de colonne dans la feuille de sortie
    Call mod_Display.RecupPosSortieIndex(wsResultat, debPlageTravail)
    
    ' Taille maximale du tableau de résultat = nombre total de lignes source
    ReDim tabResultat(1 To tblDataLineTotal, 1 To nbColonne)

    ' On copie chaque opération correspondant à la période ET au type demandé
    ' (Positif = une "entrée" d'argent, Négatif = une "dépense").
    ligneAffichage = 2
    
    'For Each ligne In tbl.ListRows
    For nbLigne = 2 To tblDataLineTotal
        ' On vérifie qu'on est sur la période désirée
        If mod_DonneesTable.RowMatchesFilter(nbLigne, tblData(nbLigne, colMoisBud), tblData(nbLigne, colAnneeBud), tblData(nbLigne, colMontant), True) Then
        'If mod_DonneesTable.RowMatchesPeriod(ligne, tbl, critAnnee, critMois) Then
            If showType = "Positif" Then
                If tblData(nbLigne, colMontant) >= 0 Then
                    idxRes = idxRes + 1
                    'On écrit la ligne dans le tableau
                    tabResultat(idxRes, posDate) = tblData(nbLigne, colDate)
                    tabResultat(idxRes, posTiers) = tblData(nbLigne, colTiers)
                    tabResultat(idxRes, posCategorie) = tblData(nbLigne, colCategorie)
                    tabResultat(idxRes, posMontant) = tblData(nbLigne, colMontant)
                    tabResultat(idxRes, posNotes) = tblData(nbLigne, colNotes)
                    tabResultat(idxRes, posBudget) = tblData(nbLigne, colBudget)
                    ligneAffichage = ligneAffichage + 1
                End If
            End If
            If showType = "Négatif" Then
                If tblData(nbLigne, colMontant) < 0 Then
                    idxRes = idxRes + 1
                    'On écrit la ligne dans le tableau
                    tabResultat(idxRes, posDate) = tblData(nbLigne, colDate)
                    tabResultat(idxRes, posTiers) = tblData(nbLigne, colTiers)
                    tabResultat(idxRes, posCategorie) = tblData(nbLigne, colCategorie)
                    tabResultat(idxRes, posMontant) = tblData(nbLigne, colMontant)
                    tabResultat(idxRes, posNotes) = tblData(nbLigne, colNotes)
                    tabResultat(idxRes, posBudget) = tblData(nbLigne, colBudget)
                    ligneAffichage = ligneAffichage + 1
                End If
            End If
        End If
    Next nbLigne

    If idxRes > 0 Then

        Application.ScreenUpdating = False
        
        ' Injection directe du tableau mémoire dans la plage d'affichage
        ' On redimentionne la taille de la plage pour pas voir s'afficher des erreur type #N/A dans les cellules en trop
        plageSortieEcriture.Resize(idxRes, nbColonne).value = tabResultat
        
        ' Mise en forme rapide des colonnes
        mod_Display.MiseEnPage wsResultat, idxRes
        
        Application.ScreenUpdating = True
        ' Tri natif Excel du tableau de détail selon le champ/ordre choisis par l'utilisateur.
        plageSortieEcriture.Resize(idxRes, nbColonne).Sort Key1:=plageSortieEcriture.Resize(idxRes, nbColonne).Columns(sortColumn), Order1:=sortOrderValue, Header:=xlNo
    Else
        ' Aucune opération ne correspond : on l'indique clairement plutôt
        ' que de laisser une feuille vide sans explication.
        wsDetail.Range(cellSortieDep).Offset(3, 1).value = "Aucune opération correspondante."
    End If
    
    MsgBox "Feuille créée : " & nomFeuille, vbInformation
End Sub
