Option Explicit

' =====================================================================================
' MODULE : mod_Display

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
'   headers       -> un tableau de textes, ex. : Array("Date", "Libellé", ...)
'
' Pourquoi limiter le nettoyage à "E1:R10000" ? Parce que les colonnes A et B
' de la feuille Synthese contiennent les PARAMETRES saisis par l'utilisateur
' (année, mois, seuils...). On ne doit surtout pas les effacer par erreur
' à chaque nouvel affichage : la zone nettoyée commence donc à la colonne E.
Public Sub PrepareOutputArea(ByVal ws As Worksheet, Optional ByVal headers As Variant, Optional ByVal debPlageTravail As Range)
    Dim nbCols As Long

    ' 1. On supprime les données d'en-tête sur 20 colonnes à partir de cellSortieDep.
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
    
    ' 2. On désactive les filtres au besoin
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    
    ' 3. On traite le cas où les en-têtes sont fournis
    If Not IsMissing(headers) And Not debPlageTravail Is Nothing Then
        nbCols = UBound(headers) - LBound(headers) + 1
        ' 1. On définit la plage dynamique d'en-têtes à partir de la cellule de départ
        Set plageSortieEnTetes = debPlageTravail.Resize(1, nbCols)
        
        ' 2. Définition de la plage de travail
        ' la méthode Offset par rapport à cellSortieDep décale la ligne de 1 et 0 colonne
        ' Range étire la plage à partir de la nouvelle valeur Offset, de 10000 lignes et nbCols
        Set plageSortieEcriture = debPlageTravail.Offset(1, 0).Resize(10000, nbCols)
        
        ' 3. Enfin, on écrit les nouveaux en-têtes de colonnes à l'endroit demandé.
        plageSortieEnTetes.value = headers
    End If
    
End Sub

' Récupération de la position absolue d'une valeur dans un tableau.
' Récupère la position du champ Entete dans le tableau headers (Application.Match est insensible à la casse).
Public Sub RecupPosSortieIndex(ByVal ws As Worksheet, Optional ByVal posZone As Range)
    If posZone Is Nothing Then
        posSortieDate = mod_Display.PosSortieIndex(ws, MonArray, "Date")
        posSortieTiers = mod_Display.PosSortieIndex(ws, MonArray, "Tiers")
        posSortieType = mod_Display.PosSortieIndex(ws, MonArray, "Type_operation")
        posSortieCategorie = mod_Display.PosSortieIndex(ws, MonArray, "Catégorie")
        posSortieMontant = mod_Display.PosSortieIndex(ws, MonArray, "Montant")
        posSortieCheque = mod_Display.PosSortieIndex(ws, MonArray, "Num_Cheque")
        posSortieNotes = mod_Display.PosSortieIndex(ws, MonArray, "Notes")
        posSortieBudget = mod_Display.PosSortieIndex(ws, MonArray, "Budget")
        posSortieAnneeBudget = mod_Display.PosSortieIndex(ws, MonArray, "AnneeBudget")
        posSortieMoisBudget = mod_Display.PosSortieIndex(ws, MonArray, "MoisBudget")
        posSortieDateConsult = mod_Display.PosSortieIndex(ws, MonArray, "Date_consult")
        posSortieSpeConsult = mod_Display.PosSortieIndex(ws, MonArray, "Spe_Consult")
        posSortieStatutSante = mod_Display.PosSortieIndex(ws, MonArray, "StatutSante")
        posSortieSoldeSante = mod_Display.PosSortieIndex(ws, MonArray, "SoldeSante")
        posSortieID = mod_Display.PosSortieIndex(ws, MonArray, "ID_Transaction")
    Else
        posSortieDate = mod_Display.PosSortieIndex(ws, MonArray, "Date", posZone)
        posSortieTiers = mod_Display.PosSortieIndex(ws, MonArray, "Tiers", posZone)
        posSortieType = mod_Display.PosSortieIndex(ws, MonArray, "Type_operation", posZone)
        posSortieCategorie = mod_Display.PosSortieIndex(ws, MonArray, "Catégorie", posZone)
        posSortieMontant = mod_Display.PosSortieIndex(ws, MonArray, "Montant", posZone)
        posSortieCheque = mod_Display.PosSortieIndex(ws, MonArray, "Num_Cheque", posZone)
        posSortieNotes = mod_Display.PosSortieIndex(ws, MonArray, "Notes", posZone)
        posSortieBudget = mod_Display.PosSortieIndex(ws, MonArray, "Budget", posZone)
        posSortieAnneeBudget = mod_Display.PosSortieIndex(ws, MonArray, "AnneeBudget", posZone)
        posSortieMoisBudget = mod_Display.PosSortieIndex(ws, MonArray, "MoisBudget", posZone)
        posSortieDateConsult = mod_Display.PosSortieIndex(ws, MonArray, "Date_consult", posZone)
        posSortieSpeConsult = mod_Display.PosSortieIndex(ws, MonArray, "Spe_Consult", posZone)
        posSortieStatutSante = mod_Display.PosSortieIndex(ws, MonArray, "StatutSante", posZone)
        posSortieSoldeSante = mod_Display.PosSortieIndex(ws, MonArray, "SoldeSante", posZone)
        posSortieID = mod_Display.PosSortieIndex(ws, MonArray, "ID_Transaction", posZone)
    End If
End Sub

' Fiabilise la récupération de la position absolue d'une colonne dans une feuille, à partir de la valeur de l'en-tête.
Public Function PosSortieIndex(ByVal ws As Worksheet, ByVal headers As Variant, Entete As String, Optional ByVal posZone As Range)
    Dim posEntete As Long
    Dim colStart As Long
    
    ' Par défaut, la fonction renvoie 0 (valeur inutilisable car les colonnes commencent à 1)
    PosSortieIndex = 0
    
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
    PosSortieIndex = colStart + CLng(posEntete) - 1
    
End Function

' Récupération de la position absolue d'une valeur dans un tableau.
' Récupère la position du champ Entete dans le tableau headers (Application.Match est insensible à la casse).
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
    posID = GetPosArray("ID_Transaction", MonArray)
    posValider = GetPosArray("Valider", MonArray)
End Sub

' Gère les erreurs potentielles dans le nom de la colonne à rechercher dans l'ARRAY.
Public Function GetPosArray(ByVal colName As String, ByVal headers As Variant) As Long
    On Error Resume Next
    GetPosArray = Application.Match(colName, headers, 0)
    On Error GoTo 0
    ' Renvoie 0 si la colonne n'existe pas
End Function

' Récupère l'index d'une colonne dans un tableau à partir de son en-tête.
' On utilise ListObject pour effectuer la recherche dans "TblOperations".
Public Sub RecupIndexCol()
    
    colID = GetColumnIndex(tbl, "ID_Transaction")
    colDate = GetColumnIndex(tbl, "Date_Comptable")
    colTiers = GetColumnIndex(tbl, "Tiers")
    colType = GetColumnIndex(tbl, "Type_operation")
    colCategorie = GetColumnIndex(tbl, "Categorie")
    ' Ajout : cette ligne manquait ici. Sans elle, colSousCategorie restait
    ' toujours à 0 (sa valeur par défaut), et tous les "If colSousCategorie <> 0"
    ' du chantier Phase 5/6 (mod_RechercheOperations, mod_SuiviSante,
    ' mod_SuiviSanteFormulaire, mod_FormulairesNotes) tombaient silencieusement
    ' dans leur branche "colonne absente" sans plantage, donc sans se faire
    ' remarquer. Repéré en préparant l'ajout du bouton "Revoir la ventilation".
    colSousCategorie = GetColumnIndex(tbl, "SousCategorie")
    colMontant = GetColumnIndex(tbl, "Montant")
    colCheque = GetColumnIndex(tbl, "Num_Cheque")
    colNotes = GetColumnIndex(tbl, "Notes")
    colBudget = GetColumnIndex(tbl, "Budget")
    colMoisBud = GetColumnIndex(tbl, "MoisBudget")
    colAnneeBud = GetColumnIndex(tbl, "AnneeBudget")
    colDecalageManuel = GetColumnIndex(tbl, "DecalageManuel")
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

' Gère les erreurs potentielles dans le nom de la colonne à rechercher.
Public Function GetColumnIndex(ByVal tbl As ListObject, ByVal colName As String) As Long
    On Error Resume Next
    GetColumnIndex = tbl.ListColumns(colName).index
    On Error GoTo 0
    ' Renvoie 0 si la colonne n'existe pas
End Function

Public Sub MiseEnPage(ws As Worksheet, totalLignes As Long, Optional Filter As Boolean = True)
    Dim i As Long

    ' On doit activer la feuille pour pouvoir modifier certains réglages de la FENÊTRE
    ' (comme l'affichage du quadrillage), car ces réglages sont attachés à la fenêtre
    ' au moment où une feuille donnée est affichée, et non à la feuille elle-même.
    ws.Activate
    
    
    ' Supprime le quadrillage (les lignes grises entre les cellules) pour donner
    ' un aspect "formulaire" plutôt que "tableur classique", comme demandé.
    ActiveWindow.DisplayGridlines = False
    
    ' Police par défaut de toute la feuille, cohérente avec un rendu "formulaire" sobre.
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    
    If posSortieDate <> 0 Then
        plageSortieEcriture.Columns(posDate).NumberFormat = "dd/mm/yyyy"
    End If
    If posSortieBudget <> 0 Then
        plageSortieEcriture.Columns(posBudget).NumberFormat = "mm/yyyy"
    End If
    If posSortieDateConsult <> 0 Then
        plageSortieEcriture.Columns(posDateConsult).NumberFormat = "dd/mm/yyyy"
    End If
    If posSortieMontant <> 0 Then
        plageSortieEcriture.Columns(posMontant).NumberFormat = "#,##0.00"
    End If
    If posSortieSoldeSante <> 0 Then
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

    ' On se replace en A1 et on désactive les barres de titres de lignes/colonnes
    ' (références L1C1 grisées) pour un rendu plus épuré. Optionnel, mais renforce
    ' l'effet "formulaire" plutôt que "feuille de calcul".
    ws.Range("A1").Select
    ActiveWindow.DisplayHeadings = False
    
End Sub

' =====================================================================================
' AppliquerCouleurMontant (ajout 08/10/2026, décision opérateur) : RÈGLE UNIQUE de
' couleur des montants pour TOUT le classeur, afin que tous les écrans parlent le
' même langage visuel :
'   - montant positif (entrée d'argent)           -> VERT  (coulMontantPositif)
'   - montant négatif ou nul (dépense)            -> NOIR  (coulMontantNegatif)
'   - montant "en alerte" (enAlerte = Vrai)       -> en GRAS, et en ROUGE
'     (coulMontantAlerte) s'il s'agit d'une dépense. Une entrée d'argent en alerte
'     reste VERTE (la norme est respectée) mais passe en gras.
' C'est l'écran appelant qui décide de ce qu'est une "alerte" (statut santé KO,
' plus grosses dépenses du mois...) : cette fonction ne fait que l'afficher.
' La cellule doit contenir le montant SIGNÉ (négatif = dépense), comme TblOperations.
' Le gras est TOUJOURS repositionné (Vrai ou Faux) : une cellule réutilisée d'un
' affichage précédent ne garde donc jamais un gras périmé.
' =====================================================================================
Public Sub AppliquerCouleurMontant(ByVal cellule As Range, Optional ByVal enAlerte As Boolean = False)
 
    Dim valeurMontant As Double
    valeurMontant = mod_DataStructure.ToDouble(cellule.Value2)
 
    With cellule.Font
        If enAlerte And valeurMontant < 0 Then
            .Color = coulMontantAlerte
        ElseIf valeurMontant > 0 Then
            .Color = coulMontantPositif
        Else
            .Color = coulMontantNegatif
        End If
        .Bold = enAlerte
    End With
 
End Sub

' =====================================================================================
' Garantie la bonne transcription des caractères accentués dans les affichages
' remplace les balises de la chaîne par le caractère Unicode lié à son code
' =====================================================================================
Public Function FR(ByVal texte As String) As String
    Dim r As String
    r = texte
    r = Replace(r, "{e2}", ChrW(233)) ' remplace dans r la chaîne "{e2}" par "é"
    r = Replace(r, "{e1}", ChrW(232)) ' remplace dans r la chaîne "{e1}" par "è"
    r = Replace(r, "{ea}", ChrW(234)) ' remplace dans r la chaîne "{ea}" par "ê"
    r = Replace(r, "{a2}", ChrW(224)) ' remplace dans r la chaîne "{a2}" par "à"
    r = Replace(r, "{c2}", ChrW(231)) ' remplace dans r la chaîne "{c2}" par "ç"
    r = Replace(r, "{o2}", ChrW(244)) ' remplace dans r la chaîne "{o2}" par "ô"
    r = Replace(r, "{i2}", ChrW(238)) ' remplace dans r la chaîne "{i2}" par "î"
    r = Replace(r, "{E2}", ChrW(201)) ' remplace dans r la chaîne "{E2}" par "É"
    r = Replace(r, "{E1}", ChrW(200)) ' remplace dans r la chaîne "{E1}" par "È"
    r = Replace(r, "{E3}", ChrW(202)) ' remplace dans r la chaîne "{E3}" par "Ê"
    r = Replace(r, "{A2}", ChrW(192)) ' remplace dans r la chaîne "{A2}" par "À"
    FR = r
End Function

'Suppression de tous les caractères accentués d'une chaîne.
'Suppression de tous les espaces.
'Mise en majuscule de la première lettre.
'Utilisée pour nommer les objets à partir de leur titre.

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
    
    ' 2. Suppression de tous les types d'espaces (espace standard, insécable, tabulation)
    res = Replace(res, " ", "")
    res = Replace(res, Chr(160), "") ' Espace insécable (fréquent dans les exports web/Excel)
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
' SOUS-PROCÉDURE - Construction de la zone des boutons (ligne 2) et du compteur
' =====================================================================================
Public Sub ConstruireBoutons(ws As Worksheet, zoneBouton As Range, texteBouton As String, nomMacro As String)

    Dim bouton As Button

    ' On définit la position de chaque bouton en s'appuyant sur une PLAGE DE CELLULES
    ' (et non des coordonnées en pixels fixes) : ainsi, si la largeur des
    ' colonnes change, les boutons suivent automatiquement. C'est plus robuste qu'un
    ' positionnement en pixels pour un environnement qui doit rester stable.


    zoneBouton.RowHeight = 22

    ' Ajout d'un bouton de type "Contrôle de formulaire"
    ' plus léger et plus fiable face aux problèmes de DPI et de multi-écrans.
    Set bouton = ws.Buttons.Add( _
        zoneBouton.Left, zoneBouton.Top, _
        zoneBouton.Width, zoneBouton.Height)
    With bouton
        ' IMPORTANT : le texte affiché est déjà son intitulé FINAL,
        ' à savoir sa fonction de réinitialisation (et non plus "Valider ce cas").
        .Caption = FR(texteBouton)
        ' On affecte le nom de la macro
        .OnAction = nomMacro
        .Name = "btn" & FR(FormaterChaine(texteBouton))
    End With
End Sub


' =====================================================================================
' SOUS-PROCÉDURE - Construction de la zone des instructions
' =====================================================================================
Public Sub ConstruireZoneTexte(ws As Worksheet, Titre As String, Message As String, ByVal PositionTitre As Range, ByVal debutZone As Range, ByVal finZone As Range)
    
    ' --- Titre "Instructions" ---
    With PositionTitre
        .value = FR(Titre)
        .Font.Bold = True
        .Font.Size = 11
    End With

    ' --- Bloc de texte des 4 étapes, sur une plage fusionnée pour ressembler à un
    ' encadré de note plutôt qu'à des cellules de tableur classiques. ---
    Dim zoneTexte As Range
    Set zoneTexte = ws.Range(debutZone, finZone)

    With zoneTexte
        .Merge
        ' Le texte exact que tu m'as fourni, avec des sauts de ligne internes
        ' (Chr(10) est le caractère "retour à la ligne" à l'intérieur d'une cellule).
        .value = FR(Message)
        .WrapText = True                     ' le texte revient à la ligne dans la cellule
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
        .Font.Size = 9
        .Font.Color = RGB(80, 80, 80)
        .Interior.Color = RGB(245, 245, 242) ' fond gris très clair, type "encadré note"
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(210, 210, 205)
        .Locked = True                       ' cellule non modifiable par l'utilisateur final
    End With

    ' Ajuste la hauteur des lignes du bloc pour laisser de la place au texte
    zoneTexte.rows.AutoFit

End Sub

