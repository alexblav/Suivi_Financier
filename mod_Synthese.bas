Attribute VB_Name = "mod_Synthese"
Option Explicit

' Cellule de demarrage de la plage d'ecriture
Public Const cellSortieDep As String = "E1"

' 1. Definition des couleurs (exemple avec un theme bleu)
Public Const coulEntete As Long = &H794E1F     ' Bleu principal (en-tete)
Public Const coulClair1 As Long = &HF8F2EE  ' Nuance 1 (tres claire)
Public Const coulClair2 As Long = &HF8E3DA   ' Nuance 2 (legerement plus soutenue)

Public AllowDetailDoubleClick As Boolean

' Stocke le numero d'index d'une colonne dans le tableau TblOperations
Public colCategorie As Long, colMontant As Long, colNotes As Long, colDate As Long, colLibelle As Long, colDateConsult As Long, colSpeConsult As Long
Public colType As Long, colBudget As Long, colTiers As Long, colMoisBud As Long, colAnneeBud As Long, colCheque As Long
Public colID As Long

' --- Ajout "Suivi Sante" (Phase 1) ---
' Index de colonne (dans TblOperations) des 4 nouveaux champs de suivi sante.
' Remplis par mod_Display.RecupIndexCol, exactement comme les variables ci-dessus.
Public colStatutSante As Long, colFranchise As Long, colSoldeSante As Long, colDepassementHoraires As Long, colCommentaireSante As Long

' Stocke l'index d'une colonne, calcule a partir d'un des champ de l'ARRAY,sur toute la feuille de sortie
' EXEMPLE Si on veut poser les entete a partir de E1 dans synthese,
' Array("Date", "Tiers", "Montant"): posDate=5,posTiers=6,posMontant=7
Public colSortieCategorie As Long, colSortieMontant As Long, colSortieNotes As Long, colSortieDate As Long, colSortieLibelle As Long, colSortieDateConsult As Long, colSortieSpeConsult As Long
Public colSortieType As Long, colSortieBudget As Long, colSortieTiers As Long, colSortieMoisBudget As Long, colSortieAnneeBudget As Long, colSortieCheque As Long
Public colSortieStatutSante As Long, colSortieSoldeSante As Long

Public critAnnee As String, critMois As String, critNbOperations As Long, critMontantMin As Double, critTriChamps As String, critTriOrdre As String

' Stocke la position d'une valeur dans un ARRAY
' EXEMPLE Array("Date", "Tiers", "Montant"): posDate=1,posTiers=2,posMontant=3
' A utiliser avec plageSortieEcriture
Public posCategorie As Long, posMontant As Long, posNotes As Long, posDate As Long, posDateConsult As Long, posSpeConsult As Long
Public posCheque As Long, posMoisBudget As Long, posAnneeBudget As Long, posType As Long, posBudget As Long, posTiers As Long, posMoisBud As Long, posAnneeBud As Long
Public posStatutSante As Long, posSoldeSante As Long

Public wsSynthese As Worksheet
Public tbl As ListObject
Public tblData As Variant
Public tblDataLineTotal As Long
Public MonArray As Variant, nbColonne As Long, tblSortie As Variant
Public plageSortieEnTetes As Range, plageSortieEcriture As Range

    
' Fonction permettant le suivi des remboursement de sante
' Elle affiche la liste des operations  liees a la categorie "Frais, remb sante"
' Ajoute un champs calcule statuant du resulat du Rbt

Public Sub Synthese_Care()
    Dim tabResultat() As Variant
    Dim tabNeg() As mod_Rapports.TOperation ' Contient les enregistrements Negatif
    Dim countNeg As Long: countNeg = 0 ' Declaration de la variable countNeg et initialisation de cette derniere a 0
    Dim tabPos() As mod_Rapports.TOperation ' Contient les enregistrements Positif
    Dim countPos As Long: countPos = 0
    Dim nbLigne As Long, nbLigneNeg As Long, nbLignePos As Long, idxRes As Long
    Dim mnt As Double
    Dim countMatches As Long
    Dim rngMontants As Range

    ' 1. Recuperation des pointeurs vers la feuille et le tableau
    Set wsSynthese = mod_Criteres.GetFeuilleSynthese()
    Set tbl = mod_DonneesTable.GetOperationsTable()
    If tbl Is Nothing Then
        MsgBox "Le tableau ne contient aucune ligne de donn" & ChrW(233) & "es.", vbExclamation
        Exit Sub
    End If
    If tbl.DataBodyRange Is Nothing Then
        MsgBox "Le tableau est vide.", vbExclamation
        Exit Sub
    End If

    ' Charge TOUT le tableau (en-tetes inclus en ligne 1)
    tblData = tbl.Range.value
    tblDataLineTotal = UBound(tblData, 1)
    
    ' 2. Recuperation des index de colonne dans la base
    mod_Display.RecupIndexCol
    
    AllowDetailDoubleClick = False
    
    ' Affiche les entetes du tableau de sortie
    MonArray = Array("Date", "Tiers", "Montant", "Notes", "Date_consult", "Spe_consult", "Statut")
    mod_Display.PrepareOutputArea wsSynthese, MonArray
    nbColonne = UBound(MonArray) - LBound(MonArray) + 1
    
    ' Recupere la position du champs dans l'ARRAY (Application.Match est naturellement insensible a la casse)
    mod_Display.RecupPosArray
    
    ' Recuperation des index de colonne dans la feuille de sortie
    mod_Display.RecupColSortieIndex
    
    ' Taille maximale du tableau de resultat = nombre total de lignes source
    ' ATTENTION on ajoute une derniere colonne pour le statut
    ReDim tabResultat(1 To tblDataLineTotal, 1 To nbColonne + 1)
           
    ' 4. Lire les criteres saisis par l'utilisateur (B1 a B6).
    mod_Criteres.GetSelectCriteres

    ' 5. Remplissage des tableaux tabNeg et tabPos (filtres sur "Frais, remb sante")
    For nbLigne = 2 To tblDataLineTotal
        ' On controle que la valeur de la colonne "Categorie" a pour valeur: "Frais, remb sante"
        If mod_DataStructure.CellText(tblData(nbLigne, colCategorie)) = mod_DataStructure.CellText("Frais, remb sant" & ChrW(233)) Then
        
            ' S'assurer que la valeur de Montant est numerique
            If IsNumeric(tblData(nbLigne, colMontant)) Then
                ' On recuperere le montant
                mnt = CDbl(tblData(nbLigne, colMontant))
            
                If mnt < 0 Then
                    countNeg = countNeg + 1
                    ' On redimentionne le tableau "tabNeg" pour que sa longueur vaille countNeg
                    ReDim Preserve tabNeg(1 To countNeg)
                    tabNeg(countNeg).LigneOrigine = nbLigne ' On stocke le numero de la ligne d'origine (celle de Donnees)
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

    ' 5. Algorithme de rapprochement (Negatif <-> Positif(s) selon le champ Notes)
    If countNeg > 0 Then
    
        ' Maintenant que le tri entre tabNeg et tabPos est realise on parcours tabNeg et on verifie si une operation correspondante exite en positif
        For nbLigneNeg = 1 To countNeg
            ' Etape A : Compter combien d'enregistrements positifs non utilises partagent la meme note
            countMatches = 0
        'trouve = False
        'idxRes = idxRes + 1
        'enrMatch = 0
            ' On verifie la presence de donnees dans "countPos"
            If countPos > 0 Then
                ' On parcours "countPos"
                For nbLignePos = 1 To countPos
                    ' On verifie que la ligne n'a pas ete utilise lors d'un traitement precedent
                    If Not tabPos(nbLignePos).Utilise Then
                        ' On utilise le champ Notes afin de trouver les correspondances
                        If mod_DataStructure.CellText(tabPos(nbLignePos).Notes) = mod_DataStructure.CellText(tabNeg(nbLigneNeg).Notes) Then
                            countMatches = countMatches + 1
                        End If
                    End If
                Next nbLignePos
            End If
            
            ' Etape B : Traitement selon la presence de correspondances
            If countMatches > 0 Then
                ' Match trouve : statut OK pour le negatif
                tabNeg(nbLigneNeg).Utilise = True
'                        tabNeg(nbLigneNeg).Utilise = True
'                        tabPos(nbLignePos).Utilise = True
                idxRes = idxRes + 1
                ' On affecte a "tabResultat" les donnees de la ligne negative
                tabResultat(idxRes, posDate) = tblData(tabNeg(nbLigneNeg).LigneOrigine, colDate)
                tabResultat(idxRes, posTiers) = tblData(tabNeg(nbLigneNeg).LigneOrigine, colTiers)
                tabResultat(idxRes, posMontant) = tblData(tabNeg(nbLigneNeg).LigneOrigine, colMontant)
                tabResultat(idxRes, posNotes) = tblData(tabNeg(nbLigneNeg).LigneOrigine, colNotes)
                tabResultat(idxRes, posDateConsult) = tblData(tabNeg(nbLigneNeg).LigneOrigine, colDateConsult)
                tabResultat(idxRes, posSpeConsult) = tblData(tabNeg(nbLigneNeg).LigneOrigine, colSpeConsult)
                tabResultat(idxRes, 7) = "OK"
                
                ' Ecriture de TOUS les enregistrements Positifs correspondants (statut OK)
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
                ' Aucun match trouve : statut KO pour l'enregistrement Negatif
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
    
    ' 6. Marquer KO les enregistrements Positifs restants (sans correspondance negative)
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
    
    ' 7. Ecriture en bloc sur la feuille Synthese
    If idxRes > 0 Then
        Application.ScreenUpdating = False
        
        ' Injection directe du tableau memoire dans la plage d'affichage
        ' On redimentionne la taille de la plage pour pas voir s'afficher des erreur type #N/A dans les cellules en trop
        plageSortieEcriture.Resize(idxRes, nbColonne).value = tabResultat
        
        ' Mise en forme rapide des colonnes
        mod_Display.MiseEnPage (idxRes)
        
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
        
        ' Activation des boutons de filtre automatique sur l'entete
        'plageSortieEcriture.Resize(1, nbColonne + 1).Columns.AutoFit
        
        Application.ScreenUpdating = True
    End If

    MsgBox "Traitement termin" & ChrW(233) & " : " & idxRes & " d" & ChrW(233) & "penses de sant" & ChrW(233) & " trait" & ChrW(233) & "es", vbInformation

End Sub


' Extrait les operations de sante mal renseignee
' On se base sur le champ Date_consult et l'on renvoie les enregistrements dont la valeur vaut 01/01/1900

Public Sub Export_Care_error()
    Dim dateOperation As Variant
    Dim tabResultat() As Variant
    Dim nbLigne As Long, idxRes As Long
    
    ' 1. Recuperation des pointeurs vers la feuille et le tableau
    Set wsSynthese = mod_Criteres.GetFeuilleSynthese()
    Set tbl = mod_DonneesTable.GetOperationsTable()
    If tbl Is Nothing Then
        MsgBox "Le tableau ne contient aucune ligne de donn" & ChrW(233) & "es.", vbExclamation
        Exit Sub
    End If
    If tbl.DataBodyRange Is Nothing Then
        MsgBox "Le tableau est vide.", vbExclamation
        Exit Sub
    End If

    ' Charge TOUT le tableau (en-tetes inclus en ligne 1)
    tblData = tbl.Range.value
    tblDataLineTotal = UBound(tblData, 1)
    
    ' 2. Recuperation des index de colonne dans la base
    mod_Display.RecupIndexCol
    
    AllowDetailDoubleClick = False
    
    ' Affiche les entetes du tableau de sortie
    MonArray = Array("Date", "Tiers", "Montant", "Notes", "Date_consult", "Spe_consult")
    mod_Display.PrepareOutputArea wsSynthese, MonArray
    nbColonne = UBound(MonArray) - LBound(MonArray) + 1
    
    ' Recupere la position du champs dans l'ARRAY (Application.Match est naturellement insensible a la casse)
    mod_Display.RecupPosArray
    
    ' Recuperation des index de colonne dans la feuille de sortie
    mod_Display.RecupColSortieIndex
    
    ' Taille maximale du tableau de resultat = nombre total de lignes source
    ReDim tabResultat(1 To tblDataLineTotal, 1 To nbColonne)
           
    ' 4. Lire les criteres saisis par l'utilisateur (B1 a B6).
    mod_Criteres.GetSelectCriteres

    ' 5. Remplissage du tableau de sortie
    For nbLigne = 2 To tblDataLineTotal
        If Not IsEmpty(tblData(nbLigne, colDateConsult)) Then
            dateOperation = tblData(nbLigne, colDateConsult)
            ' 2 correspond a 02/01/1900
            If dateOperation = "02/01/1900" Then
                idxRes = idxRes + 1
                'On ecrit la ligne dans le tableau
                tabResultat(idxRes, posDate) = tblData(nbLigne, colDate)
                tabResultat(idxRes, posTiers) = tblData(nbLigne, colTiers)
                tabResultat(idxRes, posMontant) = tblData(nbLigne, colMontant)
                tabResultat(idxRes, posNotes) = tblData(nbLigne, colNotes)
                tabResultat(idxRes, posDateConsult) = tblData(nbLigne, colDateConsult)
                tabResultat(idxRes, posSpeConsult) = tblData(nbLigne, colSpeConsult)
            End If
        End If
    Next nbLigne

    ' 6. Ecriture en bloc sur la feuille Synthese
    If idxRes > 0 Then
        Application.ScreenUpdating = False
        
        ' Injection directe du tableau memoire dans la plage d'affichage
        ' On redimentionne la taille de la plage pour pas voir s'afficher des erreur type #N/A dans les cellules en trop
        plageSortieEcriture.Resize(idxRes, nbColonne).value = tabResultat
        
        ' Mise en forme rapide des colonnes
        mod_Display.MiseEnPage (idxRes)
        
        Application.ScreenUpdating = True
    End If

    MsgBox "Traitement termin" & ChrW(233) & " : " & idxRes & " d" & ChrW(233) & "penses mal renseign" & ChrW(233) & "e(s) dans le champs Notes.", vbInformation

End Sub

' ------------------------------------------------------------------------
' Buget_Mensuel : affiche, dans les colonnes E a J de Synthese, toutes les
' operations de la periode selectionnee (B1 = annee, B2 = mois) dont le
' montant depasse le seuil B4. Les depenses les plus importantes (nombre
' fixe en B3) sont surlignees en rouge.
' ------------------------------------------------------------------------
Public Sub Buget_Mensuel()
    Dim ligneAffichage As Long
    Dim lignesDepenses() As Long
    Dim valeursDepenses() As Double
    Dim nombreDepenses As Long
    Dim nombreSurligne As Long
    Dim nbLigne As Long, idxRes As Long


    ' 1. Recuperation des pointeurs vers la feuille et le tableau
    Set wsSynthese = mod_Criteres.GetFeuilleSynthese()
    Set tbl = mod_DonneesTable.GetOperationsTable()
    If tbl Is Nothing Then
        MsgBox "Le tableau ne contient aucune ligne de donn" & ChrW(233) & "es.", vbExclamation
        Exit Sub
    End If
    If tbl.DataBodyRange Is Nothing Then
        MsgBox "Le tableau est vide.", vbExclamation
        Exit Sub
    End If

    ' Charge TOUT le tableau (en-tetes inclus en ligne 1)
    tblData = tbl.Range.value
    tblDataLineTotal = UBound(tblData, 1)
    
    ' 2. Recuperation des index de colonne dans la base
    mod_Display.RecupIndexCol
    
    AllowDetailDoubleClick = False
    
    ' Affiche les entetes du tableau de sortie
    MonArray = Array("Date", "Tiers", "Cat" & ChrW(233) & "gorie", "Montant", "Notes", "Budget")
    mod_Display.PrepareOutputArea wsSynthese, MonArray
    nbColonne = UBound(MonArray) - LBound(MonArray) + 1
    
    ' Recupere la position du champs dans l'ARRAY (Application.Match est naturellement insensible a la casse)
    mod_Display.RecupPosArray
    
    ' Recuperation des index de colonne dans la feuille de sortie
    mod_Display.RecupColSortieIndex
    
    ' Taille maximale du tableau de resultat = nombre total de lignes source
    ReDim tabResultat(1 To tblDataLineTotal, 1 To nbColonne)
           
    ' 4. Lire les criteres saisis par l'utilisateur (B1 a B6).
    mod_Criteres.GetSelectCriteres

    ' 5. Remplissage du tableau de sortie
    For nbLigne = 2 To tblDataLineTotal
        ' On recupere le montant de la ligne.
        If mod_DonneesTable.RowMatchesFilter(nbLigne, tblData(nbLigne, colMoisBud), tblData(nbLigne, colAnneeBud), tblData(nbLigne, colMontant), True) Then
            idxRes = idxRes + 1
            'On ecrit la ligne dans le tableau
            tabResultat(idxRes, posDate) = tblData(nbLigne, colDate)
            tabResultat(idxRes, posTiers) = tblData(nbLigne, colTiers)
            tabResultat(idxRes, posCategorie) = tblData(nbLigne, colCategorie)
            tabResultat(idxRes, posMontant) = tblData(nbLigne, colMontant)
            tabResultat(idxRes, posNotes) = tblData(nbLigne, colNotes)
            tabResultat(idxRes, posBudget) = tblData(nbLigne, colBudget)
        End If
    Next nbLigne
    
    'Affichage
    If idxRes > 0 Then
        Application.ScreenUpdating = False
        
        ' Injection directe du tableau memoire dans la plage d'affichage
        ' On redimentionne la taille de la plage pour pas voir s'afficher des erreur type #N/A dans les cellules en trop
        plageSortieEcriture.Resize(idxRes, nbColonne).value = tabResultat
        ligneAffichage = 2
        For nbLigne = 1 To idxRes
            ' Une entree (montant positif) est affichee en vert.
            If tabResultat(nbLigne, posMontant) > 0 Then
                plageSortieEcriture.Cells(nbLigne, posMontant).Font.Color = RGB(0, 128, 0)
            End If
            ' Une depense (montant negatif) est memorisee comme candidate au surlignage.
            If tabResultat(nbLigne, posMontant) < 0 Then
                mod_Rapports.AddExpenseForHighlight lignesDepenses, valeursDepenses, nombreDepenses, ligneAffichage, Abs(tabResultat(nbLigne, posMontant))
            End If
            ligneAffichage = ligneAffichage + 1
        Next nbLigne
    
        ' Determiner combien de depenses surligner, puis les surligner.
        nombreSurligne = mod_Rapports.GetTopCount(critNbOperations, nombreDepenses)
        mod_Rapports.HighlightTopRows lignesDepenses, valeursDepenses, nombreDepenses, nombreSurligne
        
        ' Mise en forme rapide des colonnes
        mod_Display.MiseEnPage (idxRes)
        
        Application.ScreenUpdating = True
    End If
    
    MsgBox "Traitement termin" & ChrW(233) & " : " & nombreDepenses & " d" & ChrW(233) & "penses n" & ChrW(233) & "gatives trouv" & ChrW(233) & "es, " & nombreSurligne & " ligne(s) surlign" & ChrW(233) & "e(s).", vbInformation
End Sub

' ------------------------------------------------------------------------
' Budget_Bilan_Mensuel : calcule et affiche le total des entrees, le total
' des depenses, et la difference entre les deux, pour la periode B1:B2.
' ------------------------------------------------------------------------
Public Sub Budget_Bilan_Mensuel()
    Dim totalDepenses As Double
    Dim totalEntrees As Double
    Dim Montant As Double
    Dim nbLigne As Long

    ' 1. Recuperation des pointeurs vers la feuille et le tableau
    Set wsSynthese = mod_Criteres.GetFeuilleSynthese()
    Set tbl = mod_DonneesTable.GetOperationsTable()
    If tbl Is Nothing Then
        MsgBox "Le tableau ne contient aucune ligne de donn" & ChrW(233) & "es.", vbExclamation
        Exit Sub
    End If
    If tbl.DataBodyRange Is Nothing Then
        MsgBox "Le tableau est vide.", vbExclamation
        Exit Sub
    End If

    ' Charge TOUT le tableau (en-tetes inclus en ligne 1)
    tblData = tbl.Range.value
    tblDataLineTotal = UBound(tblData, 1)
    
    ' 2. Recuperation des index de colonne dans la base
    mod_Display.RecupIndexCol
    
    ' 3. Lire les criteres saisis par l'utilisateur (B1 a B6).
    mod_Criteres.GetSelectCriteres
    
    ' 4. Nettoyage de la zone de sortie
    mod_Display.PrepareOutputArea wsSynthese
    
    ' Contrairement aux autres macros, celle-ci REACTIVE le double-clic
    ' (c'est depuis cet ecran de bilan que l'utilisateur double-clique sur
    ' un total pour en voir le detail -- voir ShowDetailForTotal ci-dessous).
    AllowDetailDoubleClick = True
    
    ' Taille maximale du tableau de resultat = nombre total de lignes source
    ReDim tabResultat(1 To 3, 1 To 2)
    totalDepenses = 0
    totalEntrees = 0

    ' 5. On parcourt toute la table et on additionne, ligne par ligne, entrees
    ' et depenses du mois/annee selectionnes.
    For nbLigne = 2 To tblDataLineTotal
    'For Each ligne In tbl.ListRows
        'If mod_DonneesTable.RowMatchesPeriod(ligne, tbl, critAnnee, critMois) Then
        If mod_DonneesTable.RowMatchesFilter(nbLigne, tblData(nbLigne, colMoisBud), tblData(nbLigne, colAnneeBud), tblData(nbLigne, colMontant), True) Then
            Montant = tblData(nbLigne, colMontant)
            'typeOperation = mod_data_structure.CellText(mod_DonneesTable.TableValue(ligne, tbl, "Type"))
            'If typeOperation = "Negatif" Or (typeOperation = "" And Montant < 0) Then
            If Montant < 0 Then
                ' Une depense est additionnee en valeur positive (voir mod_DonneesTable
                ' pour la meme logique appliquee au filtrage par seuil).
                totalDepenses = totalDepenses + Abs(Montant)
            ElseIf Montant >= 0 Then
                totalEntrees = totalEntrees + Montant
            End If
        End If
    Next nbLigne
    'Next ligne

    ' On definit la plage de sortie
    Set plageSortieEnTetes = wsSynthese.Range(cellSortieDep).Resize(3, 2)
    ' On efface l'ancien affichage avant d'ecrire le nouveau bilan.
    plageSortieEnTetes.Cells(1, 1).value = "Total Entr" & ChrW(233) & "es"
    plageSortieEnTetes.Cells(2, 1).value = "Total D" & ChrW(233) & "penses"
    plageSortieEnTetes.Cells(3, 1).value = "Diff" & ChrW(233) & "rence (Entr" & ChrW(233) & "es - D" & ChrW(233) & "penses)"
    plageSortieEnTetes.Cells(1, 2).value = totalEntrees
    plageSortieEnTetes.Cells(2, 2).value = totalDepenses
    plageSortieEnTetes.Cells(3, 2).value = totalEntrees - totalDepenses

    plageSortieEnTetes.Cells(3, 2).NumberFormat = "#,##0.00"
    plageSortieEnTetes.Columns.AutoFit
    plageSortieEnTetes.rows.AutoFit

    'AllowDetailDoubleClick = True
    MsgBox "Bilan calcul" & ChrW(233) & " : " & Format(totalEntrees - totalDepenses, "#,##0.00"), vbInformation
End Sub

' ------------------------------------------------------------------------
' ShowDetailForTotal : cree une nouvelle feuille contenant le detail complet
' (Positif ou Negatif) de la periode B1:B2, triee selon les parametres
' demandes. Appelee automatiquement par le double-clic defini dans le
' module de la feuille Synthese (en interne : "Feuil5").
' ------------------------------------------------------------------------
Public Sub ShowDetailForTotal(ByVal showType As String, Optional ByVal sortField As String = "Date", Optional ByVal sortOrder As String = "Croissant")
    Dim wsSynthese As Worksheet
    Dim tbl As ListObject
    Dim wsDetail As Worksheet
    Dim nomFeuille As String
    Dim critAnnee As String
    Dim critMois As String
    Dim ligneAffichage As Long
    Dim ligne As ListRow
    Dim sortColumn As Long
    Dim sortOrderValue As Long

    Set wsSynthese = mod_Criteres.GetFeuilleSynthese()
    Set tbl = mod_DonneesTable.GetOperationsTable()
    If tbl Is Nothing Then Exit Sub

    critAnnee = mod_data_structure.CellText(wsSynthese.Range("B1").Value2)
    critMois = mod_data_structure.CellText(wsSynthese.Range("B2").Value2)

    ' Le nom de la feuille de detail integre un horodatage pour que chaque
    ' consultation cree une feuille distincte, sans jamais ecraser un detail
    ' precedent (voir le rapport d'audit, Sec.5.4, pour une piste alternative
    ' si tu preferes eviter l'accumulation de feuilles au fil du temps).
    nomFeuille = "D" & ChrW(233) & "tail_" & IIf(showType = "N" & ChrW(233) & "gatif", "D" & ChrW(233) & "penses", "Entr" & ChrW(233) & "es") & "_" & Format(Now, "yyyymmdd_hhnnss")
    Set wsDetail = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
    wsDetail.Name = nomFeuille
    wsDetail.Range("A1:G1").value = Array("Date", "Libell" & ChrW(233), "Cat" & ChrW(233) & "gorie", "Montant", "MoisBudget", "Ann" & ChrW(233) & "eBudget", "Budget")

    sortColumn = ResolveDetailSortColumn(sortField)
    sortOrderValue = ResolveDetailSortOrder(sortOrder)

    ' On copie chaque operation correspondant a la periode ET au type demande
    ' (Positif = une "entree" d'argent, Negatif = une "depense").
    ligneAffichage = 2
    For Each ligne In tbl.ListRows
        If mod_DonneesTable.RowMatchesPeriod(ligne, tbl, critAnnee, critMois) Then
            If mod_data_structure.CellText(mod_DonneesTable.TableValue(ligne, tbl, "Type")) = showType Then
                mod_Ecriture.WriteDetailRow wsDetail, ligneAffichage, ligne, tbl
                ligneAffichage = ligneAffichage + 1
            End If
        End If
    Next ligne

    If ligneAffichage = 2 Then
        ' Aucune operation ne correspond : on l'indique clairement plutot
        ' que de laisser une feuille vide sans explication.
        wsDetail.Range("A2").value = "Aucune op" & ChrW(233) & "ration correspondante."
    Else
        wsDetail.Columns("A:G").AutoFit
        wsDetail.Range("A2:A" & ligneAffichage - 1).NumberFormat = "dd/mm/yyyy"
        wsDetail.Range("D2:D" & ligneAffichage - 1).NumberFormat = "#,##0.00"
        ' Tri natif Excel du tableau de detail selon le champ/ordre choisis par l'utilisateur.
        wsDetail.Range("A1:G" & ligneAffichage - 1).Sort Key1:=wsDetail.Cells(2, sortColumn), Order1:=sortOrderValue, Header:=xlYes
    End If

    MsgBox "Feuille cr" & ChrW(233) & ChrW(233) & "e : " & nomFeuille, vbInformation
End Sub

' ------------------------------------------------------------------------
' ResolveDetailSortColumn : traduit le nom de champ saisi par l'utilisateur
' (via la boite de dialogue InputBox de Feuil5) en numero de colonne (1 a 7)
' dans la feuille de detail. "Date" est utilise par defaut si la saisie est
' vide ou non reconnue, pour garantir un tri chronologique explicite.
' ------------------------------------------------------------------------
Private Function ResolveDetailSortColumn(ByVal sortField As String) As Long
    Select Case LCase$(Trim$(sortField))
        Case "date": ResolveDetailSortColumn = 1
        Case "libell" & ChrW(233), "libelle": ResolveDetailSortColumn = 2
        Case "cat" & ChrW(233) & "gorie", "categorie": ResolveDetailSortColumn = 3
        Case "montant": ResolveDetailSortColumn = 4
        Case "moisbudget": ResolveDetailSortColumn = 5
        Case "ann" & ChrW(233) & "ebudget", "anneebudget": ResolveDetailSortColumn = 6
        Case "budget": ResolveDetailSortColumn = 7
        Case Else: ResolveDetailSortColumn = 1
    End Select
End Function

' ------------------------------------------------------------------------
' ResolveDetailSortOrder : traduit "Croissant"/"Decroissant" (texte saisi
' par l'utilisateur) en constante Excel xlAscending/xlDescending.
' ------------------------------------------------------------------------
Private Function ResolveDetailSortOrder(ByVal sortOrder As String) As Long
    If LCase$(Trim$(sortOrder)) = "d" & ChrW(233) & "croissant" Or LCase$(Trim$(sortOrder)) = "decroissant" Then
        ResolveDetailSortOrder = xlDescending
    Else
        ResolveDetailSortOrder = xlAscending
    End If
End Function

