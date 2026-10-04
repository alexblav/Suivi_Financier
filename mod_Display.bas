Option Explicit
' Ce module regroupe les macros qui prparent une ZONE D'AFFICHAGE sur une
' feuille (nettoyage, mise en place des en-ttes) AVANT que d'autres modules
' n'y crivent des donnes. Il ne contient volontairement AUCUNE logique de
' calcul ou de filtrage : uniquement de la mise en forme / nettoyage visuel.

' Les macros de ce module ne sont pas visibles dans la liste "Macros" d'Excel
Option Private Module

' ----------------------------------------------------------------------
' PrepareOutputArea : vide une zone de rsultats et y place de nouveaux en-ttes
' ----------------------------------------------------------------------
' Paramtres :
'   ws            -> la feuille sur laquelle travailler (ex : Synthese)
'   headerAddress -> l'adresse o crire les en-ttes (ex : "E1:J1")
'   headers       -> un tableau de textes, ex : Array("Date", "Libell", ...)
'
' Pourquoi limiter le nettoyage  "E1:R10000" ? Parce que les colonnes A et B
' de la feuille Synthese contiennent les PARAMETRES saisis par l'utilisateur
' (anne, mois, seuils...). On ne doit surtout pas les effacer par erreur
'  chaque nouvel affichage : la zone nettoye commence donc  la colonne E.
Public Sub PrepareOutputArea(ByVal ws As Worksheet, Optional ByVal headers As Variant, Optional ByVal debPlageTravail As Range)
    Dim nbCols As Long

    ' 1. On supprime les donnes d'entte sur une longuenr de 20 colonnes en partant de cellSortieDep
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
    
    ' 2. On dsactive les filtre au besoin
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    
    ' 3. On traite le cas ou les enttes sont fournit
    If Not IsMissing(headers) And Not debPlageTravail Is Nothing Then
        nbCols = UBound(headers) - LBound(headers) + 1
        ' 1. On dfinit la plage dynamique d'en-ttes  partir de la cellule de dpart
        Set plageSortieEnTetes = debPlageTravail.Resize(1, nbCols)
        
        ' 2. Dfinit la plage de travail
        ' la mthode offset par de cellSortieDep dcale la ligne de 1 et 0 colonne
        ' Range tire la plage  partir de la nouvelle valeur Offset, de 10000 lignes et nbCols
        Set plageSortieEcriture = debPlageTravail.Offset(1, 0).Resize(10000, nbCols)
        
        ' 3. Enfin, on crit les nouveaux en-ttes de colonnes  l'endroit demand.
        plageSortieEnTetes.value = headers
    End If
    
End Sub

' Rcupration de la position absolu d'une valeur dans un ARRAY
' Rcupre la position du champs Entete dans l'ARRAY headers (Application.Match est naturellement insensible  la casse)
Public Sub RecupPosSortieIndex(ByVal ws As Worksheet, Optional ByVal posZone As Range)
    If posZone Is Nothing Then
        posSortieDate = mod_Display.PosSortieIndex(ws, MonArray, "Date")
        posSortieTiers = mod_Display.PosSortieIndex(ws, MonArray, "Tiers")
        posSortieType = mod_Display.PosSortieIndex(ws, MonArray, "Type_operation")
        posSortieCategorie = mod_Display.PosSortieIndex(ws, MonArray, "Catgorie")
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
        posSortieCategorie = mod_Display.PosSortieIndex(ws, MonArray, "Catgorie", posZone)
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

' Fiabilise la rcupration de la position absolu d'une colonne dans une feuille en fonction de la valeur de l'entte
Public Function PosSortieIndex(ByVal ws As Worksheet, ByVal headers As Variant, Entete As String, Optional ByVal posZone As Range)
    Dim posEntete As Long
    Dim colStart As Long
    
    ' Par dfaut, la fonction renvoie 0 (valeur inutilisable car les colonnes commencent  1)
    PosSortieIndex = 0
    
    ' 1. Vrification des paramtres d'entre
    If ws Is Nothing Then Exit Function
    If Not IsArray(headers) Then Exit Function
    If Trim(Entete) = "" Then Exit Function
    
    ' 2. Vrification de la cellule de dpart
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
    
    ' 3. Recherche (Application.Match est naturellement insensible  la casse)
    posEntete = GetPosArray(Entete, headers)
    
    ' 4. Si l'en-tte n'est pas trouv, posEntete contient une erreur
    If IsError(posEntete) Then Exit Function
    
    ' 5. Calcul et retour du numro de colonne
    PosSortieIndex = colStart + CLng(posEntete) - 1
    
End Function

' Rcupration de la position absolu d'une valeur dans un ARRAY
' Rcupre la position du champs Entete dans l'ARRAY headers (Application.Match est naturellement insensible  la casse)
Public Sub RecupPosArray()
    posDate = GetPosArray("Date", MonArray)
    posTiers = GetPosArray("Tiers", MonArray)
    posType = GetPosArray("Type_operation", MonArray)
    posCategorie = GetPosArray("Catgorie", MonArray)
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

' Gre de potentiel erreur dans le nom de la colonne  rechercher dans l'ARRAY
Public Function GetPosArray(ByVal colName As String, ByVal headers As Variant) As Long
    On Error Resume Next
    GetPosArray = Application.Match(colName, headers, 0)
    On Error GoTo 0
    ' Renvoie 0 si la colonne n'existe pas
End Function

' Rcupre le numro d'index d'une colonne dans un tableau par son nom d'entte
' On utlise ListObject car on recherhe dans le tableau nomm "TblOperations"
Public Sub RecupIndexCol()
    
    colID = GetColumnIndex(tbl, "ID_Transaction")
    colDate = GetColumnIndex(tbl, "Date_Comptable")
    colTiers = GetColumnIndex(tbl, "Tiers")
    colType = GetColumnIndex(tbl, "Type_operation")
    colCategorie = GetColumnIndex(tbl, "Categorie")
    ' Ajout : cette ligne manquait ici. Sans elle, colSousCategorie restait
    ' toujours  0 (sa valeur par dfaut), et tous les "If colSousCategorie <> 0"
    ' du chantier Phase 5/6 (mod_RechercheOperations, mod_SuiviSante,
    ' mod_SuiviSanteFormulaire, mod_FormulairesNotes) tombaient silencieusement
    ' dans leur branche "colonne absente"  sans plantage, donc sans se faire
    ' remarquer. Repr en prparant l'ajout du bouton "Revoir la ventilation".
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

' Gre de potentiel erreur dans le nom de la colonne  rechercher
Public Function GetColumnIndex(ByVal tbl As ListObject, ByVal colName As String) As Long
    On Error Resume Next
    GetColumnIndex = tbl.ListColumns(colName).index
    On Error GoTo 0
    ' Renvoie 0 si la colonne n'existe pas
End Function

Public Sub MiseEnPage(ws As Worksheet, totalLignes As Long, Optional Filter As Boolean = True)
    Dim i As Long

    ' On doit activer la feuille pour pouvoir modifier certains reglages de la FENETRE
    ' (comme l'affichage du quadrillage), car ces reglages sont attaches a la fenetre
    ' au moment ou une feuille donnee est affichee, et non a la feuille elle-meme.
    ws.Activate
    
    
    ' Supprime le quadrillage (les lignes grises entre les cellules) pour donner
    ' un aspect "formulaire" plutot que "tableur classique", comme demande.
    ActiveWindow.DisplayGridlines = False
    
    ' Police par defaut de toute la feuille, coherente avec un rendu "formulaire" sobre.
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

    ' 1. Formatage de la ligne d'en-tte
    With plageSortieEnTetes
        .Interior.Color = coulEntete
        .Font.Color = RGB(255, 255, 255) ' Texte blanc
        .Font.Bold = True
    End With
    
    ' 2. Alternance sur les lignes de la plage de donnes
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
' Garantie la bonne transcription des carractre accentu dans les affichages
' remplace les balise de la chaine par le carractre Unicode li  son code
' =====================================================================================
Public Function FR(ByVal texte As String) As String
    Dim r As String
    r = texte
    r = Replace(r, "{e2}", ChrW(233)) ' remplace dans r la chaine "{e2}" par ""
    r = Replace(r, "{e1}", ChrW(232)) ' 
    r = Replace(r, "{ea}", ChrW(234)) ' 
    r = Replace(r, "{a2}", ChrW(224)) ' 
    r = Replace(r, "{c2}", ChrW(231)) ' 
    r = Replace(r, "{o2}", ChrW(244)) ' 
    r = Replace(r, "{i2}", ChrW(238)) ' 
    r = Replace(r, "{E2}", ChrW(201)) ' 
    r = Replace(r, "{E1}", ChrW(200)) ' 
    r = Replace(r, "{E3}", ChrW(202)) ' 
    r = Replace(r, "{A2}", ChrW(192)) ' 
    FR = r
End Function

'Suppression de tous les caractres accentus d'une chaine
' Suppression de tous les espaces
' Positionnement de la premire lettre en majuscule
' Utilis pour le nommage des objet en fonction de leur titre

Public Function FormaterChaine(ByVal texte As String) As String
    Dim i As Long
    Dim accents As String
    Dim sansAccents As String
    Dim res As String
    
    res = texte
    
    ' 1. Liste des caractres accentus et leurs quivalents
    accents = ""
    sansAccents = "aaaaaaeeeeiiiiooooouuuucnyyAAAAAAEEEEIIIIOOOOOUUUUCNY"
    
    ' Remplacement des ligatures spcifiques
    res = Replace(res, "", "oe")
    res = Replace(res, "", "OE")
    res = Replace(res, "", "ae")
    res = Replace(res, "", "AE")
    
    ' Remplacement caractre par caractre
    For i = 1 To Len(accents)
        res = Replace(res, Mid(accents, i, 1), Mid(sansAccents, i, 1))
    Next i
    
    ' 2. Suppression de tous les types d'espaces (espace standard, insecable, tabulation)
    res = Replace(res, " ", "")
    res = Replace(res, Chr(160), "") ' Espace insecable (frquent dans les exports web/Excel)
    res = Replace(res, vbTab, "")
    
    ' 3. Premire lettre en majuscule
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
    ' plus leger, plus fiable vis-a-vis du probleme de DPI/multi-ecrans.
    Set bouton = ws.Buttons.Add( _
        zoneBouton.Left, zoneBouton.Top, _
        zoneBouton.Width, zoneBouton.Height)
    With bouton
        ' IMPORTANT : le texte affich est dj son intitul FINAL,
        ' a savoir sa fonction de reinitialisation (et non plus "Valider ce cas").
        .Caption = FR(texteBouton)
        ' On affecte le nom de la macro
        .OnAction = nomMacro
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

    ' --- Bloc de texte des 4 etapes, sur une plage fusionnee pour ressembler a un
    ' encadre de note plutot qu'a des cellules de tableur classiques. ---
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
        .Interior.Color = RGB(245, 245, 242) ' fond gris trs clair, type "encadre note"
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(210, 210, 205)
        .Locked = True                       ' cellule non modifiable par l'utilisateur final
    End With

    ' Ajuste la hauteur des lignes du bloc pour laisser de la place au texte
    zoneTexte.rows.AutoFit

End Sub

