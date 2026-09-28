Option Explicit
' Fonction permettant le suivi des remboursement de santé
' Elle affiche la liste des opérations  liées à la catégorie "Frais, remb santé"
' Ajoute un champs calculé statuant du résulat du Rbt

Public Sub Synthese_Care()
    Dim tabResultat() As Variant
    Dim tabNeg() As mod_Rapports.TOperation ' Contient les enregistrements Négatif
    Dim countNeg As Long: countNeg = 0 ' Déclaration de la variable countNeg et initialisation de cette dernière à 0
    Dim tabPos() As mod_Rapports.TOperation ' Contient les enregistrements Positif
    Dim countPos As Long: countPos = 0
    Dim nbLigne As Long, nbLigneNeg As Long, nbLignePos As Long, idxRes As Long
    Dim mnt As Double
    Dim countMatches As Long
    Dim rngMontants As Range, debPlageTravail As Range
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
    Message = " Suivi Sant{e2}" & Chr(10) & _
                "Les op{e2}rations list{e2}es sont toutes les op{e2}rations traitant du suivi sant{e2}" & Chr(10) & _
                 "{A2} la fin sortez avec le bouton ""Sortir"""
                 
    ' On fournit le titre, la position de la cellule dans laquelle on veut écrire de titre
    Titre = "Instructions"
    Set PositionTitre = wsResultat.Range(cellSortieDep).Offset(2, 0)
    
    ' On détermine la plage de début et de fin de la zone
    Set debutZone = wsResultat.Range(cellSortieDep).Offset(3, 0)
    Set finZone = wsResultat.Range(cellSortieDep).Offset(5, 6)
    
    Call mod_Display.ConstruireZoneTexte(wsResultat, Titre, Message, PositionTitre, debutZone, finZone)
    
    
    ' 7. Affiche les entêtes du tableau de sortie
    ' Définit le début de la page de travail
    Set debPlageTravail = wsResultat.Range(cellSortieDep).Offset(7, 0)
    
    MonArray = Array("Date", "Tiers", "Montant", "Notes", "Date_consult", "Spe_consult", "Statut")
    Call mod_Display.PrepareOutputArea(wsResultat, MonArray, debPlageTravail)
    
    nbColonne = UBound(MonArray) - LBound(MonArray) + 1
    
    ' Récupére la position du champs dans l'ARRAY (Application.Match est naturellement insensible à la casse)
    Call mod_Display.RecupPosArray
    
    ' Récupération des index de colonne dans la feuille de sortie
    Call mod_Display.RecupPosSortieIndex(wsResultat, debPlageTravail)
    
    ' Taille maximale du tableau de résultat = nombre total de lignes source
    ' ATTENTION on ajoute une dernière colonne pour le statut
    ReDim tabResultat(1 To tblDataLineTotal, 1 To nbColonne + 1)
           
    ' 8. Lire les critères saisis par l'utilisateur (B1 à B6).
    mod_Criteres.GetSelectCriteres

    ' 9. Remplissage des tableaux tabNeg et tabPos (filtrés sur "Frais, remb santé")
    For nbLigne = 2 To tblDataLineTotal
        ' On contrôle que la valeur de la colonne "Catégorie" à pour valeur: "Frais, remb santé"
        If mod_DataStructure.CellText(tblData(nbLigne, colCategorie)) = mod_DataStructure.CellText("Frais, remb santé") Then
        
            ' S'assurer que la valeur de Montant est numérique
            If IsNumeric(tblData(nbLigne, colMontant)) Then
                ' On récupérère le montant
                mnt = CDbl(tblData(nbLigne, colMontant))
            
                If mnt < 0 Then
                    countNeg = countNeg + 1
                    ' On redimentionne le tableau "tabNeg" pour que sa longueur vaille countNeg
                    ReDim Preserve tabNeg(1 To countNeg)
                    tabNeg(countNeg).LigneOrigine = nbLigne ' On stocke le numéro de la ligne d'origine (celle de Données)
                    tabNeg(countNeg).Montant = mnt ' Converti en positif pour la comparaison
                    tabNeg(countNeg).Notes = mod_DataStructure.CellText(tblData(nbLigne, colNotes))
                    tabNeg(countNeg).Utilise = False '
                    
                ElseIf mnt > 0 Then
                    countPos = countPos + 1
                    ReDim Preserve tabPos(1 To countPos)
                    tabPos(countPos).LigneOrigine = nbLigne
                    tabPos(countPos).Montant = mnt
                    tabPos(countPos).Notes = mod_DataStructure.CellText(tblData(nbLigne, colNotes))
                    tabPos(countPos).Utilise = False
                End If
            End If
        End If
    Next nbLigne

    ' 10. Algorithme de rapprochement (Négatif <-> Positif(s) selon le champ Notes)
    If countNeg > 0 Then
    
        ' Maintenant que le tri entre tabNeg et tabPos est réalisé on parcours tabNeg et on vérifie si une opération correspondante exite en positif
        For nbLigneNeg = 1 To countNeg
            ' Étape A : Compter combien d'enregistrements positifs non utilisés partagent la même note
            countMatches = 0
        'trouve = False
        'idxRes = idxRes + 1
        'enrMatch = 0
            ' On vérifie la présence de données dans "countPos"
            If countPos > 0 Then
                ' On parcours "countPos"
                For nbLignePos = 1 To countPos
                    ' On vérifie que la ligne n'a pas été utilisé lors d'un traitement précédent
                    If Not tabPos(nbLignePos).Utilise Then
                        ' On utilise le champ Notes afin de trouver les correspondances
                        If mod_DataStructure.CellText(tabPos(nbLignePos).Notes) = mod_DataStructure.CellText(tabNeg(nbLigneNeg).Notes) Then
                            countMatches = countMatches + 1
                        End If
                    End If
                Next nbLignePos
            End If
            
            ' Étape B : Traitement selon la présence de correspondances
            If countMatches > 0 Then
                ' Match trouvé : statut OK pour le négatif
                tabNeg(nbLigneNeg).Utilise = True
'                        tabNeg(nbLigneNeg).Utilise = True
'                        tabPos(nbLignePos).Utilise = True
                idxRes = idxRes + 1
                ' On affecte à "tabResultat" les données de la ligne négative
                tabResultat(idxRes, posDate) = tblData(tabNeg(nbLigneNeg).LigneOrigine, colDate)
                tabResultat(idxRes, posTiers) = tblData(tabNeg(nbLigneNeg).LigneOrigine, colTiers)
                tabResultat(idxRes, posMontant) = tblData(tabNeg(nbLigneNeg).LigneOrigine, colMontant)
                tabResultat(idxRes, posNotes) = tblData(tabNeg(nbLigneNeg).LigneOrigine, colNotes)
                tabResultat(idxRes, posDateConsult) = tblData(tabNeg(nbLigneNeg).LigneOrigine, colDateConsult)
                tabResultat(idxRes, posSpeConsult) = tblData(tabNeg(nbLigneNeg).LigneOrigine, colSpeConsult)
                tabResultat(idxRes, 7) = "OK"
                
                ' Écriture de TOUS les enregistrements Positifs correspondants (statut OK)
                For nbLignePos = 1 To countPos
                    If Not tabPos(nbLignePos).Utilise Then
                        If mod_DataStructure.CellText(tabPos(nbLignePos).Notes) = mod_DataStructure.CellText(tabNeg(nbLigneNeg).Notes) Then
                            tabPos(nbLignePos).Utilise = True
                            
                            idxRes = idxRes + 1
                            tabResultat(idxRes, posDate) = tblData(tabPos(nbLignePos).LigneOrigine, colDate)
                            tabResultat(idxRes, posTiers) = tblData(tabPos(nbLignePos).LigneOrigine, colTiers)
                            tabResultat(idxRes, posMontant) = tblData(tabPos(nbLignePos).LigneOrigine, colMontant)
                            tabResultat(idxRes, posNotes) = tblData(tabPos(nbLignePos).LigneOrigine, colNotes)
                            tabResultat(idxRes, posDateConsult) = tblData(tabPos(nbLignePos).LigneOrigine, colDateConsult)
                            tabResultat(idxRes, posSpeConsult) = tblData(tabPos(nbLignePos).LigneOrigine, colSpeConsult)
                            tabResultat(idxRes, 7) = "OK"
                        End If
                    End If
                Next nbLignePos
            Else
                ' Aucun match trouvé : statut KO pour l'enregistrement Négatif
                idxRes = idxRes + 1
                tabResultat(idxRes, posDate) = tblData(tabNeg(nbLigneNeg).LigneOrigine, colDate)
                tabResultat(idxRes, posTiers) = tblData(tabNeg(nbLigneNeg).LigneOrigine, colTiers)
                tabResultat(idxRes, posMontant) = tblData(tabNeg(nbLigneNeg).LigneOrigine, colMontant)
                tabResultat(idxRes, posNotes) = tblData(tabNeg(nbLigneNeg).LigneOrigine, colNotes)
                tabResultat(idxRes, posDateConsult) = tblData(tabNeg(nbLigneNeg).LigneOrigine, colDateConsult)
                tabResultat(idxRes, posSpeConsult) = tblData(tabNeg(nbLigneNeg).LigneOrigine, colSpeConsult)
                tabResultat(idxRes, 7) = "KO"
            End If
        Next nbLigneNeg
    End If
    
    ' 11. Marquer KO les enregistrements Positifs restants (sans correspondance négative)
    If countPos > 0 Then
        For nbLignePos = 1 To countPos
            If Not tabPos(nbLignePos).Utilise Then
                idxRes = idxRes + 1
                tabResultat(idxRes, posDate) = tblData(tabPos(nbLignePos).LigneOrigine, colDate)
                tabResultat(idxRes, posTiers) = tblData(tabPos(nbLignePos).LigneOrigine, colTiers)
                tabResultat(idxRes, posMontant) = tblData(tabPos(nbLignePos).LigneOrigine, colMontant)
                tabResultat(idxRes, posNotes) = tblData(tabPos(nbLignePos).LigneOrigine, colNotes)
                tabResultat(idxRes, posDateConsult) = tblData(tabPos(nbLignePos).LigneOrigine, colDateConsult)
                tabResultat(idxRes, posSpeConsult) = tblData(tabPos(nbLignePos).LigneOrigine, colSpeConsult)
                tabResultat(idxRes, 7) = "KO"
            End If
        Next nbLignePos
    End If
    
    ' 12. Écriture en bloc sur la feuille Synthese
    If idxRes > 0 Then
        Application.ScreenUpdating = False
        
        ' Injection directe du tableau mémoire dans la plage d'affichage
        ' On redimentionne la taille de la plage pour pas voir s'afficher des erreur type #N/A dans les cellules en trop
        plageSortieEcriture.Resize(idxRes, nbColonne).value = tabResultat
        
        ' Mise en forme rapide des colonnes
        mod_Display.MiseEnPage wsResultat, idxRes
        
        ' Tri par date croissante (colonne E)
        plageSortieEcriture.Sort _
            Key1:=plageSortieEcriture.Columns(posDateConsult), _
            Order1:=xlDescending, _
            Header:=xlNo
            

        ' Couleur rouge sur les montants dont le statut en K est "KO"
        ' 1. Cibler uniquement la colonne des montants dans la plage de sortie
        Dim plageMontants As Range
        Set plageMontants = plageSortieEcriture.Columns(posMontant)
        plageMontants.FormatConditions.Delete
        With plageMontants.FormatConditions.Add(Type:=xlExpression, Formula1:="=" & plageSortieEcriture.Cells(1, 7).Address(False, True) & "=""KO""")
            .Font.Color = vbRed
            .Font.Bold = True
        End With
        
        ' Activation des boutons de filtre automatique sur l'entête
        'plageSortieEcriture.Resize(1, nbColonne + 1).Columns.AutoFit
        
         ' Réactive les événements
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If

    MsgBox "Traitement terminé : " & idxRes & " dépenses de santé traitées", vbInformation

End Sub

