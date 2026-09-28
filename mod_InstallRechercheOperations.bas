Option Explicit

' =====================================================================================
' MODULE : mod_InstallRechercheOperations
'
' ROLE (PHASE 5a) :
'   Construit la mise en page STATIQUE (aucune logique de clic pour l'instant)
'   de la feuille masquee "frm_RechercheOperations" : un ecran de recherche
'   et de correction en masse pour TblOperations, base sur un TABLEAU EXCEL
'   CLASSIQUE avec filtre automatique natif (les fleches de filtre dans
'   l'entete font tout le travail de filtrage croise Date/Tiers/Montant/
'   Catégorie/Notes, sans code personnalise).
'
'   Colonnes du tableau (dans cet ordre) :
'     A - Valider       : l'opérateur y inscrit "Oui" sur les lignes finies
'     B - Date
'     C - Tiers
'     D - Montant
'     E - Catégorie     : liste déroulante (avertissement, pas de blocage :
'                         on peut taper une nouvelle catégorie qui n'existe
'                         pas encore)
'     F - SousCategorie : idem, PHASE 5 (catégories a 2 niveaux)
'     G - Notes         : texte libre, SAUF si la valeur est déjà une clé
'                         santé valide (grisee dans ce cas - voir Phase 5b)
'     H - Ventile        : colonne INFORMATIVE (non modifiable), PHASE 5.
'                         "Oui" si la ligne représente une PART VENTILEE
'                         d'une opération bancaire (elle vient alors de
'                         TblVentilations, pas de TblOperations), ou si
'                         l'opération PARENTE d'une ligne normale a ete
'                         ventilee (Catégorie = "Ventile"). C'est le "tag"
'                         de tracabilite demandé par l'opérateur : "on peut
'                         prevenir un tag indiquant que cette opération fait
'                         partie d'une ventilation, pour information". Une
'                         opération ventilee reste ainsi accessible ICI de 2
'                         facons : via sa ligne parente (Catégorie="Ventile"),
'                         ou directement via chacune de ses parts (une ligne
'                         par sous-catégorie de la ventilation).
'     I - ID_Transaction : colonne technique MASQUEE (ID de l'opération, ou
'                         de l'opération PARENTE pour une part ventilee)
'     J - SourceLigne    : colonne technique MASQUEE, PHASE 5 : "O" (ligne de
'                         TblOperations) ou "V" (part de TblVentilations),
'                         sert a savoir OU ecrire au moment d'appliquer
'     K - LigneVentilation : colonne technique MASQUEE, PHASE 5 : pour une
'                         ligne "V", position de la part DANS TblVentilations
'                         (DataBodyRange). Vide/non utilisée pour une ligne "O".
'
'   La Phase 5b (mod_RechercheOperations) remplira le tableau depuis
'   TblOperations ET TblVentilations (bouton "Rechercher") et appliquera les
'   lignes marquees (bouton "Appliquer les lignes marquees") dans la bonne
'   table source, colonne par colonne, jamais par un tri/decoupage de texte.
'
'   Ajout suite  un test oprateur : un 3e bouton "Revoir la ventilation"
'   permet, sur une ligne dont la colonne Ventil vaut "Oui", de rouvrir le
'   formulaire de ventilation (frm_Ventilation) avec le dtail dj enregistr
'   dans TblVentilations, au lieu d'un formulaire vide (voir
'   mod_RechercheOperations.RevoirVentilationRO et l'explication complte en
'   tte de mod_Ventilation).
'
'  PROPOS DES ACCENTS : tout ce qui s'affiche dans Excel continue  passer par
' la fonction FR() pour rester 100% sr  l'import VBA. Les commentaires que
' j'ajoute  partir de maintenant utilisent de vrais caractres accentus pour
' rester lisibles (convention valide avec l'oprateur) ; les anciens
' commentaires du fichier restent tels quels pour l'instant.
'
' A FAIRE POUR INSTALLER CE MODULE :
'   1. Alt+F11, Fichier > Importer un fichier..., choisir ce fichier .bas
'   2. Ctrl+G : CreerFeuilleRechercheOperations (ATTENTION : reconstruit toute la
'      feuille, y compris ses boutons  c'est le seul moyen de faire apparatre
'      le nouveau bouton "Revoir la ventilation" ; rimporter le fichier seul
'      ne suffit pas, comme dj repr plus tt sur ce chantier)
'   3. Pour revoir la feuille : AfficherFeuilleRecherchePourEdition
'      Pour la remasquer : MasquerFeuilleRechercheApresEdition
' =====================================================================================


Public Const NOM_FEUILLE_RECHERCHE As String = "frm_RechercheOperations"
Public Const NOM_TABLE_RECHERCHE As String = "TblRechercheOperations"

Public Const RO_LIGNE_BOUTONS As Long = 2
Public Const RO_LIGNE_ENTETES As Long = 4

' Position des colonnes DANS LE TABLEAU (1 = première colonne du tableau, A)
Public Const RO_COL_VALIDER As Long = 1
Public Const RO_COL_DATE As Long = 2
Public Const RO_COL_TIERS As Long = 3
Public Const RO_COL_MONTANT As Long = 4
Public Const RO_COL_CATEGORIE As Long = 5
Public Const RO_COL_SOUSCATEGORIE As Long = 6   ' PHASE 5
Public Const RO_COL_NOTES As Long = 7
Public Const RO_COL_VENTILE As Long = 8         ' PHASE 5 (informatif, non modifiable)
Public Const RO_COL_ID As Long = 9
Public Const RO_COL_SOURCE As Long = 10         ' PHASE 5 : "O" ou "V" (technique, masquee)
Public Const RO_COL_LIGNEVEN As Long = 11       ' PHASE 5 : ligne dans TblVentilations si SourceLigne="V" (technique, masquee)


'Private Function FR(ByVal texte As String) As String
'    Dim r As String
'    r = texte
'    r = Replace(r, "{e2}", ChrW(233))
'    r = Replace(r, "{e1}", ChrW(232))
'    r = Replace(r, "{ea}", ChrW(234))
'    r = Replace(r, "{a2}", ChrW(224))
'    r = Replace(r, "{c2}", ChrW(231))
'    r = Replace(r, "{o2}", ChrW(244))
'    r = Replace(r, "{i2}", ChrW(238))
'    r = Replace(r, "{E2}", ChrW(201))
'    FR = r
'End Function


' =====================================================================================
' MACRO D'INSTALLATION
' =====================================================================================
Sub CreerFeuilleRechercheOperations()

    Dim ws As Worksheet
    Dim reponse As VbMsgBoxResult
    Dim tbl As ListObject
    Dim plageDepart As Range

    Set ws = ObtenirFeuilleSansErreurRO(NOM_FEUILLE_RECHERCHE)

    If Not ws Is Nothing Then
        reponse = MsgBox("La feuille '" & NOM_FEUILLE_RECHERCHE & "' existe deja." & vbCrLf & _
                          "Veux-tu la reconstruire entierement (sa mise en forme actuelle sera perdue) ?", _
                          vbYesNo + vbQuestion, "Confirmation de reconstruction")
        If reponse = vbNo Then
            MsgBox "Installation annulee, aucune modification effectuee.", vbInformation
            Exit Sub
        End If
        ws.Visible = xlSheetVisible
        On Error Resume Next
        ws.ListObjects(NOM_TABLE_RECHERCHE).Delete
        On Error GoTo 0
        ws.Cells.Clear
        Call SupprimerFormesExistantesRO(ws)
        Call SupprimerNomsExistantsRO(ws, NOM_FEUILLE_RECHERCHE)
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = NOM_FEUILLE_RECHERCHE
    End If

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10

    ws.Columns("A").ColumnWidth = 10
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 24
    ws.Columns("D").ColumnWidth = 12
    ws.Columns("E").ColumnWidth = 20
    ws.Columns("F").ColumnWidth = 20
    ws.Columns("G").ColumnWidth = 40
    ws.Columns("H").ColumnWidth = 10
    ws.Columns("I").ColumnWidth = 12
    ws.Columns("J").ColumnWidth = 10
    ws.Columns("K").ColumnWidth = 10

    ' --- Boutons ---
    Dim zoneBtn1 As Range, zoneBtn2 As Range
    Dim btn As Button

    Set zoneBtn1 = ws.Range("A" & RO_LIGNE_BOUTONS & ":B" & RO_LIGNE_BOUTONS)
    zoneBtn1.RowHeight = 22
    Set btn = ws.Buttons.Add(zoneBtn1.Left, zoneBtn1.Top, zoneBtn1.Width, zoneBtn1.Height)
    With btn
        .Caption = FR("Rechercher")
        .OnAction = "RechercherOperations"
        .Name = "btnRechercherOperations"
    End With

    Set zoneBtn2 = ws.Range("C" & RO_LIGNE_BOUTONS & ":E" & RO_LIGNE_BOUTONS)
    Set btn = ws.Buttons.Add(zoneBtn2.Left, zoneBtn2.Top, zoneBtn2.Width, zoneBtn2.Height)
    With btn
        .Caption = FR("Appliquer les lignes marqu{e2}es")
        .OnAction = "AppliquerLignesMarquees"
        .Name = "btnAppliquerLignesMarquees"
    End With

    ' Ajout suite  un test oprateur : bouton pour rouvrir le dtail d'une
    ' opration dj ventile (slectionner une ligne du tableau dont la colonne
    ' "Ventil" vaut "Oui", puis cliquer ici)  voir mod_RechercheOperations.RevoirVentilationRO.
    Dim zoneBtn3 As Range
    Set zoneBtn3 = ws.Range("F" & RO_LIGNE_BOUTONS & ":G" & RO_LIGNE_BOUTONS)
    Set btn = ws.Buttons.Add(zoneBtn3.Left, zoneBtn3.Top, zoneBtn3.Width, zoneBtn3.Height)
    With btn
        .Caption = FR("Revoir la ventilation")
        .OnAction = "RevoirVentilationRO"
        .Name = "btnRevoirVentilationRO"
    End With

    ' --- Petit rappel du fonctionnement, au-dessus du tableau ---
    With ws.Range("A" & (RO_LIGNE_ENTETES - 2) & ":G" & (RO_LIGNE_ENTETES - 2))
        .Merge
        .value = FR("Utilise les fl{e2}ches de filtre dans l'en-t{ea}te (comme un filtre Excel classique) pour restreindre la liste. " & _
                    "Corrige Cat{e2}gorie/Notes directement dans les cellules, inscris 'Oui' dans Valider, puis clique sur 'Appliquer les lignes marqu{e2}es'. " & _
                    "Pour revoir le d{e2}tail d'une op{e2}ration ventil{e2}e (colonne Ventil{e2} = Oui) : s{e2}lectionne sa ligne puis clique sur 'Revoir la ventilation'.")
        .Font.Size = 9
        .Font.Color = RGB(80, 80, 80)
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    ws.rows(RO_LIGNE_ENTETES - 2).RowHeight = 28

    ' --- Tableau (headers + 1 ligne vide de depart, indispensable pour créer un ListObject) ---
    ws.Range("A" & RO_LIGNE_ENTETES).value = "Valider"
    ws.Range("B" & RO_LIGNE_ENTETES).value = "Date"
    ws.Range("C" & RO_LIGNE_ENTETES).value = "Tiers"
    ws.Range("D" & RO_LIGNE_ENTETES).value = "Montant"
    ws.Range("E" & RO_LIGNE_ENTETES).value = "Categorie"
    ws.Range("F" & RO_LIGNE_ENTETES).value = "SousCategorie"
    ws.Range("G" & RO_LIGNE_ENTETES).value = "Notes"
    ws.Range("H" & RO_LIGNE_ENTETES).value = "Ventile"
    ws.Range("I" & RO_LIGNE_ENTETES).value = "ID_Transaction"
    ws.Range("J" & RO_LIGNE_ENTETES).value = "SourceLigne"
    ws.Range("K" & RO_LIGNE_ENTETES).value = "LigneVentilation"

    Set plageDepart = ws.Range("A" & RO_LIGNE_ENTETES & ":K" & (RO_LIGNE_ENTETES + 1))
    Set tbl = ws.ListObjects.Add(xlSrcRange, plageDepart, , xlYes)
    tbl.Name = NOM_TABLE_RECHERCHE
    tbl.TableStyle = "TableStyleMedium2"

    ' Colonnes techniques masquees (PHASE 5 : Ventile reste VISIBLE, c'est le
    ' tag informatif demandé par l'opérateur -- seules I/J/K, qui ne servent
    ' qu'au code, sont masquees)
    ws.Columns("I").Hidden = True
    ws.Columns("J").Hidden = True
    ws.Columns("K").Hidden = True

    ws.Visible = xlSheetVeryHidden

    MsgBox "La feuille '" & NOM_FEUILLE_RECHERCHE & "' a ete creee et masquee." & vbCrLf & _
           "Pour la revoir : AfficherFeuilleRecherchePourEdition", vbInformation, "Installation terminee"

End Sub


' =====================================================================================
' FONCTIONS UTILITAIRES
' =====================================================================================
Private Function ObtenirFeuilleSansErreurRO(ByVal nomFeuille As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nomFeuille)
    On Error GoTo 0
    Set ObtenirFeuilleSansErreurRO = ws
End Function

Private Sub SupprimerFormesExistantesRO(ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
End Sub

Private Sub SupprimerNomsExistantsRO(ws As Worksheet, ByVal nomFeuille As String)
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
' OUTILS DEVELOPPEUR
' =====================================================================================
Sub AfficherFeuilleRecherchePourEdition()
    Dim ws As Worksheet
    Set ws = ObtenirFeuilleSansErreurRO(NOM_FEUILLE_RECHERCHE)
    If ws Is Nothing Then
        MsgBox "La feuille '" & NOM_FEUILLE_RECHERCHE & "' n'existe pas encore.", vbExclamation
        Exit Sub
    End If
    ws.Visible = xlSheetVisible
    ws.Activate
    MsgBox "La feuille est maintenant visible. Pense a la remasquer avec" & vbCrLf & _
           "MasquerFeuilleRechercheApresEdition", vbInformation
End Sub

Sub MasquerFeuilleRechercheApresEdition()
    Dim ws As Worksheet
    Set ws = ObtenirFeuilleSansErreurRO(NOM_FEUILLE_RECHERCHE)
    If ws Is Nothing Then
        MsgBox "La feuille '" & NOM_FEUILLE_RECHERCHE & "' n'existe pas.", vbExclamation
        Exit Sub
    End If
    ws.Visible = xlSheetVeryHidden
    MsgBox "La feuille est de nouveau masquee.", vbInformation
End Sub
