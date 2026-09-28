Option Explicit

Public Sub Synthese_Export_Care_error()
    Dim dateOperation As Variant
    Dim tabResultat() As Variant
    Dim nbLigne As Long, idxRes As Long
    Dim reponse As VbMsgBoxResult
    Dim debPlageTravail As Range
    
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
    
    ' On masque la feuille Synthese. Cette option est prise pour éviter à l'opérateur de se déplacer en dehors de la feuille crée
    wsSynthese.Visible = xlSheetVeryHidden
    
    ' 2. Charge TOUT le tableau de données (en-têtes incluses en ligne 1)
    tblData = tbl.Range.value
    tblDataLineTotal = UBound(tblData, 1)
    
    ' 3. Récupération des index de colonne dans la base de données
    mod_Display.RecupIndexCol
    
    ' 4. Autorise ou non le double click
    AllowDetailDoubleClick = False
    ' Paramètre de navigation
    RecherOperations = False
    ' Désactive les événements
    Application.EnableEvents = False

    ' 5. Construction de la zone des boutons
    ' On détermine la position du bouton
    Set zoneBouton = wsResultat.Range(cellSortieDep).Offset(0, 0)
    
    'On définit son titre
    texteBouton = "Sortir"
    nomMacroBouton = "Sortir"
    Call mod_Display.ConstruireBoutons(wsResultat, zoneBouton, texteBouton, nomMacroBouton)
    
    ' 6. Construction de la zone des instructions
    ' On fournit le message à afficher en remplaçant les carractères accentué par les balise de la fonction FR
    Message = " Ligne santé en erreur " & Chr(10) & _
              " Liste des données de santée en erreur" & Chr(10) & _
              " Extrait les opérations de sant{e2} mal renseign{e2}e" & Chr(10) & _
              " On se base sur le champ Date_consult et l'on renvoie les enregistrements dont la valeur vaut 01/01/1900"
                 
    ' On fournit le titre, la position de la cellule dans laquelle on veut écrire de titre
    Titre = "Instructions"
    Set PositionTitre = wsResultat.Range(cellSortieDep).Offset(2, 0)
    
    ' On détermine la plage de début et de fin de la zone
    Set debutZone = wsResultat.Range(cellSortieDep).Offset(3, 0)
    Set finZone = wsResultat.Range(cellSortieDep).Offset(6, 5)
    
    Call mod_Display.ConstruireZoneTexte(wsResultat, Titre, Message, PositionTitre, debutZone, finZone)
    
     
    ' 7. Affiche les entêtes du tableau de sortie
    ' Définit le début de la page de travail
    Set debPlageTravail = wsResultat.Range(cellSortieDep).Offset(7, 0)
    
    MonArray = Array("Date", "Tiers", "Montant", "Notes", "Date_consult", "Spe_consult")
    Call mod_Display.PrepareOutputArea(wsResultat, MonArray, debPlageTravail)
  
    nbColonne = UBound(MonArray) - LBound(MonArray) + 1
    
    ' Récupére la position du champs dans l'ARRAY (Application.Match est naturellement insensible à la casse)
    mod_Display.RecupPosArray
    
    ' Récupération des index de colonne dans la feuille de sortie
    Call mod_Display.RecupPosSortieIndex(wsResultat, debPlageTravail)
    
    ' Taille maximale du tableau de résultat = nombre total de lignes source
    ReDim tabResultat(1 To tblDataLineTotal, 1 To nbColonne)
           
    ' 8. Lire les critères saisis par l'utilisateur (B1 à B6).
    mod_Criteres.GetSelectCriteres

    ' 9. Remplissage du tableau de sortie
    For nbLigne = 2 To tblDataLineTotal
        If Not IsEmpty(tblData(nbLigne, colDateConsult)) Then
            dateOperation = tblData(nbLigne, colDateConsult)
            ' 2 correspond à 02/01/1900
            If dateOperation = "02/01/1900" Then
                idxRes = idxRes + 1
                'On écrit la ligne dans le tableau
                tabResultat(idxRes, posDate) = tblData(nbLigne, colDate)
                tabResultat(idxRes, posTiers) = tblData(nbLigne, colTiers)
                tabResultat(idxRes, posMontant) = tblData(nbLigne, colMontant)
                tabResultat(idxRes, posNotes) = tblData(nbLigne, colNotes)
                tabResultat(idxRes, posDateConsult) = tblData(nbLigne, colDateConsult)
                tabResultat(idxRes, posSpeConsult) = tblData(nbLigne, colSpeConsult)
            End If
        End If
    Next nbLigne

    ' 10. Écriture en bloc sur la feuille Synthese
    If idxRes > 0 Then
        Application.ScreenUpdating = False
        
        ' Injection directe du tableau mémoire dans la plage d'affichage
        ' On redimentionne la taille de la plage pour pas voir s'afficher des erreur type #N/A dans les cellules en trop
        plageSortieEcriture.Resize(idxRes, nbColonne).value = tabResultat
        
        ' Mise en forme rapide des colonnes
        mod_Display.MiseEnPage wsResultat, idxRes
         
         ' Réactive les événements
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If

    MsgBox "Traitement terminé : " & idxRes & " dépenses mal renseignée(s) dans le champs Notes.", vbInformation

End Sub
