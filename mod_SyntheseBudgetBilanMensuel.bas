Option Explicit

' =====================================================================================
' MODULE : mod_SyntheseBudgetBilanMensuel

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
    
    ' PHASE 6 : le double-clic est abandonné (jugé trop discret par l'operateur),
    ' remplace par 2 boutons explicites -- voir plus bas. AllowDetailDoubleClick
    ' n'est donc plus mis a True ici.
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

    ' PHASE 6 : 2 boutons remplacent l'ancien double-clic sur les cellules de
    ' total, pour ouvrir le détail correspondant dans le nouvel écran central
    ' frm_RechercheOperations (prefiltre "DetailTotal").
    Set zoneBouton = wsResultat.Range(cellSortieDep).Offset(0, 2).Resize(1, 2)
    texteBouton = FR("Voir le d{e2}tail des entr{e2}es")
    nomMacroBouton = "VoirDetailEntreesRO"
    Call mod_Display.ConstruireBoutons(wsResultat, zoneBouton, texteBouton, nomMacroBouton)

    Set zoneBouton = wsResultat.Range(cellSortieDep).Offset(0, 4).Resize(1, 2)
    texteBouton = FR("Voir le d{e2}tail des d{e2}penses")
    nomMacroBouton = "VoirDetailDepensesRO"
    Call mod_Display.ConstruireBoutons(wsResultat, zoneBouton, texteBouton, nomMacroBouton)

    ' 6. Construction de la zone des instructions
    ' On fournit le message à afficher en remplaçant les carractères accentué par les balise de la fonction FR
    Message = " Bilan Mensuel" & Chr(10) & _
                "Affiche une synth{e2}se des dépenses et des revenus sur la p{e2}riode du:" & Chr(10) & _
                "Mois: " & critMois & " Ann{e2}e: " & critAnnee & Chr(10) & _
                 "Utilisez les boutons ci-dessus pour voir le d{e2}tail des op{e2}rations. " & _
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
' PHASE 6 : ShowDetailForTotal (et le double-clic qui l'appelait, côté
' ThisWorkbook/mod_Actions) est supprimée. Le détail d'un total s'ouvre
' désormais dans l'écran central frm_RechercheOperations, via les 2 boutons
' ajoutés dans Budget_Bilan_Mensuel ci-dessus (prefiltre "DetailTotal").
' NOTE POUR L'OPERATEUR : la case "trier automatiquement par le champ choisi
' sur Synthese" n'a PAS été reprise -- l'opérateur trie desormais lui-même
' avec les flèches de filtre natives du nouvel écran, comme pour tout le
' reste. Simplification assumée pour éviter une correspondance de colonnes
' fragile entre l'ancien et le nouveau tableau.
' ------------------------------------------------------------------------
Public Sub VoirDetailEntreesRO()
    mod_RechercheOperations.RechercherOperations "DetailTotal", "Positif"
End Sub

Public Sub VoirDetailDepensesRO()
    mod_RechercheOperations.RechercherOperations "DetailTotal", "Negatif"
End Sub
