Attribute VB_Name = "mod_Display"
Option Explicit
' Ce module regroupe les macros qui préparent une ZONE D'AFFICHAGE sur une
' feuille (nettoyage, mise en place des en-têtes) AVANT que d'autres modules
' n'y écrivent des données. Il ne contient volontairement AUCUNE logique de
' calcul ou de filtrage : uniquement de la mise en forme / nettoyage visuel.

' Les macros de ce module ne sont pas visibles dans la liste "Macros" d'Excel
Option Private Module

' ----------------------------------------------------------------------
' PrepareOutputArea : vide une zone de résultats et y place de nouveaux en-têtes
' ----------------------------------------------------------------------
' Paramètres :
'   ws            -> la feuille sur laquelle travailler (ex : Synthese)
'   headerAddress -> l'adresse où écrire les en-têtes (ex : "E1:J1")
'   headers       -> un tableau de textes, ex : Array("Date", "Libellé", ...)
'
' Pourquoi limiter le nettoyage à "E1:R10000" ? Parce que les colonnes A et B
' de la feuille Synthese contiennent les PARAMETRES saisis par l'utilisateur
' (année, mois, seuils...). On ne doit surtout pas les effacer par erreur
' à chaque nouvel affichage : la zone nettoyée commence donc à la colonne E.
Public Sub PrepareOutputArea(ByVal ws As Worksheet, Optional ByVal headers As Variant, Optional ByVal debPlageTravail As Range)
    Dim nbCols As Long

    ' 1. On supprime les données d'entête sur une longuenr de 20 colonnes en partant de cellSortieDep
    If Not debPlageTravail Is Nothing Then
        debPlageTravail.Resize(10000, 20).ClearContents
        debPlageTravail.Resize(10000, 20).FormatConditions.Delete
        debPlageTravail.Resize(10000, 20).Interior.ColorIndex = xlNone
        debPlageTravail.Resize(10000, 20).Font.Color = RGB(0, 0, 0)
        debPlageTravail.Resize(10000, 20).Font.Bold = False
    Else
        ws.Range(cellSortieDep).Resize(10000, 20).ClearContents
        ws.Range(cellSortieDep).Resize(10000, 20).FormatConditions.Delete
        ws.Range(cellSortieDep).Resize(10000, 20).Interior.ColorIndex = xlNone
        ws.Range(cellSortieDep).Resize(10000, 20).Font.Color = RGB(0, 0, 0)
        ws.Range(cellSortieDep).Resize(10000, 20).Font.Bold = False
    End If
    
    ' 2. On désactive les filtre au besoin
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    
    ' 3. On traitÃ© le cas ou les entêtes sont fournit
    If Not IsMissing(headers) And Not debPlageTravail Is Nothing Then
        nbCols = UBound(headers) - LBound(headers) + 1
        ' 1. On définit la plage dynamique d'en-têtes à partir de la cellule de départ
        Set plageSortieEnTetes = debPlageTravail.Resize(1, nbCols)
        
        ' 2. Définit la plage de travail
        ' la méthode offset par de cellSortieDep décale la ligne de 1 et 0 colonne
        ' Range étire la plage à partir de la nouvelle valeur Offset, de 10000 lignes et nbCols
        Set plageSortieEcriture = debPlageTravail.Offset(1, 0).Resize(10000, nbCols)
        
        ' 3. Enfin, on écrit les nouveaux en-têtes de colonnes à l'endroit demandé.
        plageSortieEnTetes.value = headers
    End If
    
End Sub

' Récupération de la position absolu d'une valeur dans un ARRAY
' Récupére la position du champs Entete dans l'ARRAY headers (Application.Match est naturellement insensible à la casse)
Public Sub RecupColSortieIndex(ByVal ws As Worksheet, Optional ByVal posZone As Range)
    If posZone Is Nothing Then
        colSortieDate = ColSortieIndex(ws, MonArray, "Date")
        colSortieTiers = ColSortieIndex(ws, MonArray, "Tiers")
        colSortieType = ColSortieIndex(ws, MonArray, "Type_operation")
        colSortieCategorie = ColSortieIndex(ws, MonArray, "Catégorie")
        colSortieMontant = ColSortieIndex(ws, MonArray, "Montant")
        colSortieCheque = ColSortieIndex(ws, MonArray, "Num_Cheque")
        colSortieNotes = ColSortieIndex(ws, MonArray, "Notes")
        colSortieBudget = ColSortieIndex(ws, MonArray, "Budget")
        colSortieAnneeBudget = ColSortieIndex(ws, MonArray, "AnneeBudget")
        colSortieMoisBudget = ColSortieIndex(ws, MonArray, "MoisBudget")
        colSortieDateConsult = ColSortieIndex(ws, MonArray, "Date_consult")
        colSortieSpeConsult = ColSortieIndex(ws, MonArray, "Spe_Consult")
        colSortieStatutSante = ColSortieIndex(ws, MonArray, "StatutSante")
        colSortieSoldeSante = ColSortieIndex(ws, MonArray, "SoldeSante")
    Else
        colSortieDate = ColSortieIndex(ws, MonArray, "Date", posZone)
        colSortieTiers = ColSortieIndex(ws, MonArray, "Tiers", posZone)
        colSortieType = ColSortieIndex(ws, MonArray, "Type_operation", posZone)
        colSortieCategorie = ColSortieIndex(ws, MonArray, "Catégorie", posZone)
        colSortieMontant = ColSortieIndex(ws, MonArray, "Montant", posZone)
        colSortieCheque = ColSortieIndex(ws, MonArray, "Num_Cheque", posZone)
        colSortieNotes = ColSortieIndex(ws, MonArray, "Notes", posZone)
        colSortieBudget = ColSortieIndex(ws, MonArray, "Budget", posZone)
        colSortieAnneeBudget = ColSortieIndex(ws, MonArray, "AnneeBudget", posZone)
        colSortieMoisBudget = ColSortieIndex(ws, MonArray, "MoisBudget", posZone)
        colSortieDateConsult = ColSortieIndex(ws, MonArray, "Date_consult", posZone)
        colSortieSpeConsult = ColSortieIndex(ws, MonArray, "Spe_Consult", posZone)
        colSortieStatutSante = ColSortieIndex(ws, MonArray, "StatutSante", posZone)
        colSortieSoldeSante = ColSortieIndex(ws, MonArray, "SoldeSante", posZone)
    End If
End Sub

' Fiabilise la récupération de la position absolu d'une colonne dans une feuille en fonction de la valeur de l'entête
Public Function ColSortieIndex(ByVal ws As Worksheet, ByVal headers As Variant, Entete As String, Optional ByVal posZone As Range)
    Dim posEntete As Long
    Dim colStart As Long
    
    ' Par défaut, la fonction renvoie 0 (valeur inutilisable car les colonnes commencent à 1)
    ColSortieIndex = 0
    
    ' 1. Vérification des paramètres d'entrée
    If ws Is Nothing Then Exit Function
    If Not IsArray(headers) Then Exit Function
    If Trim(Entete) = "" Then Exit Function
    
    ' 2. Vérification de la cellule de départ
    If posZone Is Nothing Then
        On Error Resume Next
        colStart = ws.Range(cellSortieDep).Column
        If Err.Number <> 0 Then
            Err.Clear
            Exit Function
        End If
        On Error GoTo 0
    Else
        On Error Resume Next
        colStart = posZone.Column
        If Err.Number <> 0 Then
            Err.Clear
            Exit Function
        End If
        On Error GoTo 0
    End If
    
    ' 3. Recherche (Application.Match est naturellement insensible à la casse)
    posEntete = GetPosArray(Entete, headers)
    
    ' 4. Si l'en-tête n'est pas trouvé, posEntete contient une erreur
    If IsError(posEntete) Then Exit Function
    
    ' 5. Calcul et retour du numéro de colonne
    ColSortieIndex = colStart + CLng(posEntete) - 1
    
End Function

' Récupération de la position absolu d'une valeur dans un ARRAY
' Récupére la position du champs Entete dans l'ARRAY headers (Application.Match est naturellement insensible à la casse)
Public Sub RecupPosArray()
    posDate = GetPosArray("Date", MonArray)
    posTiers = GetPosArray("Tiers", MonArray)
    posType = GetPosArray("Type_operation", MonArray)
    posCategorie = GetPosArray("Catégorie", MonArray)
    posMontant = GetPosArray("Montant", MonArray)
    posCheque = GetPosArray("Num_Cheque", MonArray)
    posNotes = GetPosArray("Notes", MonArray)
    posBudget = GetPosArray("Budget", MonArray)
    posMoisBudget = GetPosArray("MoisBudget", MonArray)
    posAnneeBudget = GetPosArray("AnneeBudget", MonArray)
    posDateConsult = GetPosArray("Date_consult", MonArray)
    posSpeConsult = GetPosArray("Spe_Consult", MonArray)
    posStatutSante = GetPosArray("StatutSante", MonArray)
    posSoldeSante = GetPosArray("SoldeSante", MonArray)
End Sub

' Gére de potentiel erreur dans le nom de la colonne à rechercher dans l'ARRAY
Private Function GetPosArray(ByVal colName As String, ByVal headers As Variant) As Long
    On Error Resume Next
    GetPosArray = Application.Match(colName, headers, 0)
    On Error GoTo 0
    ' Renvoie 0 si la colonne n'existe pas
End Function

' Récupére le numéro d'index d'une colonne dans un tableau par son nom d'entête
Public Sub RecupIndexCol()
    
    colID = GetColumnIndex(tbl, "ID_Transaction")
    colDate = GetColumnIndex(tbl, "Date_Comptable")
    colTiers = GetColumnIndex(tbl, "Tiers")
    colType = GetColumnIndex(tbl, "Type_operation")
    colCategorie = GetColumnIndex(tbl, "Categorie")
    colSousCategorie = GetColumnIndex(tbl, "SousCategorie")   ' PHASE 1/6 : 0 si Phase 1 pas encore installee
    colMontant = GetColumnIndex(tbl, "Montant")
    colCheque = GetColumnIndex(tbl, "Num_Cheque")
    colNotes = GetColumnIndex(tbl, "Notes")
    colBudget = GetColumnIndex(tbl, "Budget")
    colMoisBud = GetColumnIndex(tbl, "MoisBudget")
    colAnneeBud = GetColumnIndex(tbl, "AnneeBudget")
    colDateConsult = GetColumnIndex(tbl, "Date_consult")
    colSpeConsult = GetColumnIndex(tbl, "Spe_consult")
    colStatutSante = GetColumnIndex(tbl, "StatutSante")
    colFranchise = GetColumnIndex(tbl, "Franchise")
    colSoldeSante = GetColumnIndex(tbl, "SoldeSante")
    colDepassementHoraires = GetColumnIndex(tbl, "DepassementHoraires")
    colCommentaireSante = GetColumnIndex(tbl, "CommentaireSante")
    colFranchise = GetColumnIndex(tbl, "Franchise")
    colBeneficiaire = GetColumnIndex(tbl, "Beneficiaire")
End Sub

' Gére de potentiel erreur dans le nom de la colonne à rechercher
Private Function GetColumnIndex(ByVal tbl As ListObject, ByVal colName As String) As Long
    On Error Resume Next
    GetColumnIndex = tbl.ListColumns(colName).index
    On Error GoTo 0
    ' Renvoie 0 si la colonne n'existe pas
End Function

Public Sub MiseEnPage(ws As Worksheet, totalLignes As Long, Optional Filter As Boolean = True)
    Dim i As Long

    ' On doit activer la feuille pour pouvoir modifier certains reglages de la FENETRE
    ' (comme l'affichage du quadrillage), car ces reglages sont attaches a la fenetre
    ' au moment ou une feuille donnee est affichee, et non a la feuille elle-mÃªme.
    ws.Activate
    
    
    ' Supprime le quadrillage (les lignes grises entre les cellules) pour donner
    ' un aspect "formulaire" plutot que "tableur classique", comme demandÃ©.
    ActiveWindow.DisplayGridlines = False
    
    ' Police par defaut de toute la feuille, coherente avec un rendu "formulaire" sobre.
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    
    If colSortieDate <> 0 Then
        plageSortieEcriture.Columns(posDate).NumberFormat = "dd/mm/yyyy"
    End If
    If colSortieBudget <> 0 Then
        plageSortieEcriture.Columns(posBudget).NumberFormat = "mm/yyyy"
    End If
    If colSortieDateConsult <> 0 Then
        plageSortieEcriture.Columns(posDateConsult).NumberFormat = "dd/mm/yyyy"
    End If
    If colSortieMontant <> 0 Then
        plageSortieEcriture.Columns(posMontant).NumberFormat = "#,##0.00"
    End If
    If colSortieSoldeSante <> 0 Then
        plageSortieEcriture.Columns(posSoldeSante).NumberFormat = "#,##0.00"
    End If
    
    If Filter Then
        plageSortieEnTetes.AutoFilter
    End If
    
    plageSortieEcriture.Offset(-1, 0).Columns.AutoFit
    plageSortieEcriture.rows.AutoFit
    
    'Mise en forme du tableau

    ' 1. Formatage de la ligne d'en-tête
    With plageSortieEnTetes
        .Interior.Color = coulEntete
        .Font.Color = RGB(255, 255, 255) ' Texte blanc
        .Font.Bold = True
    End With
    
    ' 2. Alternance sur les lignes de la plage de données
    For i = 1 To totalLignes
        If i Mod 2 <> 0 Then
            plageSortieEcriture.rows(i).Interior.Color = coulClair1
        Else
            plageSortieEcriture.rows(i).Interior.Color = coulClair2
        End If
    Next i

    ' On se replace en A1 et on desactive les barres de titres de lignes/colonnes
    ' (references L1C1 grisees) pour un rendu plus epure. Optionnel, mais renforce
    ' l'effet "formulaire" plutot que "feuille de calcul".
    ws.Range("A1").Select
    ActiveWindow.DisplayHeadings = False
    
End Sub


' =====================================================================================
' Garantie la bonne transcription des carractère accentué dans les affichages
' remplace les balise de la chaine par le carractère Unicode lié à son code
' =====================================================================================
Public Function FR(ByVal texte As String) As String
    Dim r As String
    r = texte
    r = Replace(r, "{e2}", ChrW(233)) ' remplace dans r la chaine "{e2}" par "é"
    r = Replace(r, "{e1}", ChrW(232)) ' è
    r = Replace(r, "{ea}", ChrW(234)) ' ê
    r = Replace(r, "{a2}", ChrW(224)) ' à
    r = Replace(r, "{c2}", ChrW(231)) ' ç
    r = Replace(r, "{o2}", ChrW(244)) ' ô
    r = Replace(r, "{i2}", ChrW(238)) ' î
    r = Replace(r, "{E2}", ChrW(201)) ' É
    r = Replace(r, "{E1}", ChrW(200)) ' È
    r = Replace(r, "{E3}", ChrW(202)) ' Ê
    r = Replace(r, "{A2}", ChrW(192)) ' À
    FR = r
End Function

'Suppression de tous les caractères accentués d'une chaine
' Suppression de tous les espaces
' Positionnement de la première lettre en majuscule
' Utilisé pour le nommage des objet en fonction de leur titre

Public Function FormaterChaine(ByVal texte As String) As String
    Dim i As Long
    Dim accents As String
    Dim sansAccents As String
    Dim res As String
    
    res = texte
    
    ' 1. Liste des caractères accentués et leurs équivalents
    accents = "àáâãäåèéêëìíîïòóôõöùúûüçñýÿÀÁÂÃÄÅÈÉÊËÌÍÎÏÒÓÔÕÖÙÚÛÜÇÑÝ"
    sansAccents = "aaaaaaeeeeiiiiooooouuuucnyyAAAAAAEEEEIIIIOOOOOUUUUCNY"
    
    ' Remplacement des ligatures spécifiques
    res = Replace(res, "œ", "oe")
    res = Replace(res, "Œ", "OE")
    res = Replace(res, "æ", "ae")
    res = Replace(res, "Æ", "AE")
    
    ' Remplacement caractère par caractère
    For i = 1 To Len(accents)
        res = Replace(res, Mid(accents, i, 1), Mid(sansAccents, i, 1))
    Next i
    
    ' 2. Suppression de tous les types d'espaces (espace standard, insecable, tabulation)
    res = Replace(res, " ", "")
    res = Replace(res, Chr(160), "") ' Espace insecable (fréquent dans les exports web/Excel)
    res = Replace(res, vbTab, "")
    
    ' 3. Première lettre en majuscule
    If Len(res) > 0 Then
        res = UCase(Left(res, 1)) & Mid(res, 2)
    End If
    
    FormaterChaine = res
End Function

Public Sub SupprimerFormesExistantesFN(ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
End Sub

Public Sub SupprimerNomsExistantsFN(ws As Worksheet, ByVal nomFeuille As String)
    Dim n As Name
    Dim i As Long
    For i = ThisWorkbook.Names.count To 1 Step -1
        Set n = ThisWorkbook.Names(i)
        On Error Resume Next
        If InStr(1, n.RefersTo, "'" & nomFeuille & "'", vbTextCompare) > 0 Then
            n.Delete
        End If
        On Error GoTo 0
    Next i
End Sub
' =====================================================================================
' SOUS-PROCEDURE - Construction de la zone des boutons (ligne 2) + compteur
' =====================================================================================
Public Sub ConstruireBoutons(ws As Worksheet, zoneBouton As Range, texteBouton As String, nomMacro As String)

    Dim bouton As Button

    ' On definit la position de chaque bouton en s'appuyant sur une PLAGE DE CELLULES
    ' (et non des coordonnees en pixels fixes) : ainsi, si jamais la largeur des
    ' colonnes change, les boutons suivent automatiquement. C'est plus robuste qu'un
    ' positionnement en pixels pour un environnement qui doit rester stable.


    zoneBouton.RowHeight = 22

    ' Ajout d'un bouton de type "Controle de formulaire"
    ' plus leger, plus fiable vis-a-vis du problÃ¨me de DPI/multi-ecrans.
    Set bouton = ws.Buttons.Add( _
        zoneBouton.Left, zoneBouton.Top, _
        zoneBouton.Width, zoneBouton.Height)
    With bouton
        ' IMPORTANT : le texte affiché est déjà son intitulé FINAL,
        ' a savoir sa fonction de reinitialisation (et non plus "Valider ce cas").
        .Caption = FR(texteBouton)
        ' On affecte le nom de la macro
        .OnAction = "Sortir"
        .Name = "btn" & FR(FormaterChaine(texteBouton))
    End With
End Sub


' =====================================================================================
' SOUS-PROCEDURE - Construction de la zone des instructions
' =====================================================================================
Public Sub ConstruireZoneTexte(ws As Worksheet, Titre As String, Message As String, ByVal PositionTitre As Range, ByVal debutZone As Range, ByVal finZone As Range)
    
    ' --- Titre "Instructions" ---
    With PositionTitre
        .value = FR(Titre)
        .Font.Bold = True
        .Font.Size = 11
    End With

    ' --- Bloc de texte des 4 Ã©tapes, sur une plage fusionnee pour ressembler a un
    ' encadre de note plutot qu'a des cellules de tableur classiques. ---
    'Set debutZone = ws.Range(cellSortieDep).Offset(3, 0)
    'Set finZone = ws.Range(cellSortieDep).Offset(6, 6)
    Dim zoneTexte As Range
    Set zoneTexte = ws.Range(debutZone, finZone)

    With zoneTexte
        .Merge
        ' Le texte exact que tu m'as fourni, avec des sauts de ligne internes
        ' (Chr(10) est le caractere "retour a la ligne" a l'interieur d'une cellule).
        .value = FR(Message)
        .WrapText = True                     ' le texte revient a la ligne dans la cellule
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
        .Font.Size = 9
        .Font.Color = RGB(80, 80, 80)
        .Interior.Color = RGB(245, 245, 242) ' fond gris trÃ¨s clair, type "encadre note"
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(210, 210, 205)
        .Locked = True                       ' cellule non modifiable par l'utilisateur final
    End With

    ' Remarque : le texte des instructions ci-dessus a ete legerement adapte par
    ' rapport a l'original du UserForm, car les Ã©tapes 1 a 3 de l'ancien texte
    ' decrivaient le fonctionnement de l'Option A (clic sur une ligne puis liste
    ' partagee puis bouton Valider). Comme on est passe a l'Option B (liste
    ' dÃ©roulante directement sur chaque ligne, statut automatique), le texte des
    ' Ã©tapes a ete adapte pour rester exact vis-a-vis du nouveau fonctionnement.
    ' Dis-moi si tu preferes une autre formulation, c'est une simple modification
    ' de texte, sans impact sur le reste du code.

    ' Ajuste la hauteur des lignes du bloc pour laisser de la place au texte
    ws.rows(LIGNE_DEBUT_TEXTE_INSTRUCTIONS & ":" & LIGNE_FIN_TEXTE_INSTRUCTIONS).RowHeight = 16

End Sub

