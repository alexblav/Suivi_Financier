Attribute VB_Name = "mod_Pointage"
Option Explicit

' =====================================================================================
' MODULE : mod_Pointage                                     (ajout 11/10/2026)
'
' RÔLE : le POINTAGE des opérations. Pointer une opération, c'est indiquer qu'elle a
' déjà été analysée : cela permet de savoir d'un coup d'oeil, à l'import comme dans les
' recherches, ce qui reste à examiner.
'
' PRINCIPES (décidés avec l'opérateur) :
'   - Le pointage est une colonne de TblOperations nommée "Pointe", qui vaut "Oui"
'     (pointée) ou rien (à analyser). Une liste déroulante (Oui / vide) la saisit.
'   - Le pointage appartient à l'OPÉRATION BANCAIRE. Une part de ventilation n'a pas de
'     pointage propre : elle affiche et modifie celui de son opération d'origine
'     (retrouvée par ID_Transaction). Pointer une part pointe donc toute l'opération.
'   - RETIRER un pointage (Oui -> vide) demande TOUJOURS une confirmation, pour éviter
'     une dévalidation par erreur (voir ConfirmerDepointage). Pointer n'en demande pas.
'
' OÙ LE POINTAGE APPARAÎT :
'   - à l'import et en modification d'une opération : frm_ControleCategories ;
'   - dans l'écran de recherche (colonne Pointe, filtrable avec les flèches d'Excel) ;
'   - dans les formulaires qui affichent une opération : frm_Ventilation,
'     frm_RapprochementNotes, frm_GenerationCle et frm_SuiviSante.
'
' COMMENT LES ÉCRANS SONT BRANCHÉS (un seul point d'entrée) :
'   Une modification d'une cellule de pointage est interceptée par l'événement du
'   CLASSEUR (ThisWorkbook.Workbook_SheetChange -> TraiterChangementFeuille). Aucun code
'   n'est donc à coller dans les feuilles. Chaque formulaire déclare sa cellule de
'   pointage à l'affichage (AfficherPointageFormulaire) ; l'écran de recherche est
'   reconnu par son nom de feuille.
'
' INSTALLATION : lancer UNE FOIS la macro AjouterColonnePointage (Ctrl+G), puis
' reconstruire les écrans concernés (voir le guide, section "Le pointage").
' =====================================================================================

' Feuille -> Array(adresse de la cellule, ID_Transaction, écriture directe, ...) :
' formulaires actuellement ouverts et leur cellule de pointage.
Private g_Champs As Object
' ID_Transaction -> Boolean : pointage choisi pour les opérations d'un import EN COURS
' (elles ne sont pas encore dans TblOperations ; voir ControlerCategories).
Private g_PointageImport As Object

Private Const COL_TECH_ID As String = "ID_Transaction"

' =====================================================================================
' INSTALLATION : ajoute la colonne "Pointe" à TblOperations
' =====================================================================================
Public Sub AjouterColonnePointage()

    Dim tblOp As ListObject
    Dim colonne As ListColumn

    Set tblOp = mod_DonneesTable.GetOperationsTable()
    If tblOp Is Nothing Then Exit Sub

    On Error Resume Next
    Set colonne = tblOp.ListColumns(mod_VarGlobales.NOM_COL_POINTE)
    On Error GoTo 0

    If Not colonne Is Nothing Then
        MsgBox "La colonne " & mod_VarGlobales.NOM_COL_POINTE & " existe deja dans " & tblOp.Name & _
               ". Rien n'a ete modifie.", vbInformation, "Pointage"
        Exit Sub
    End If

    ' ListColumns.Add sans argument ajoute la colonne après la dernière du tableau.
    Set colonne = tblOp.ListColumns.Add
    colonne.Name = mod_VarGlobales.NOM_COL_POINTE

    ' Liste déroulante Oui / vide sur toute la colonne (les nouvelles lignes du tableau
    ' héritent de la validation quand Excel étend le tableau).
    If Not colonne.DataBodyRange Is Nothing Then PoserListePointage colonne.DataBodyRange

    MsgBox "Colonne " & mod_VarGlobales.NOM_COL_POINTE & " ajoutee a la fin de " & tblOp.Name & "." & vbCrLf & _
           "Toutes les operations existantes sont a analyser (colonne vide)." & vbCrLf & vbCrLf & _
           "Reconstruisez ensuite les ecrans concernes (voir le guide, section Pointage).", _
           vbInformation, "Pointage"

End Sub


' =====================================================================================
' OUTILS DE BASE
' =====================================================================================

' Vrai si la valeur d'une cellule signifie "pointée" (Oui, insensible à la casse).
Public Function TexteEnPointage(ByVal valeur As Variant) As Boolean
    TexteEnPointage = (LCase$(mod_DataStructure.CellText(valeur)) = LCase$(mod_VarGlobales.POINTE_OUI))
End Function

' Texte à écrire dans une cellule pour un état de pointage : "Oui" ou vide.
Public Function TextePointage(ByVal pointe As Boolean) As String
    If pointe Then TextePointage = mod_VarGlobales.POINTE_OUI Else TextePointage = ""
End Function

' Pose la liste déroulante Oui / vide sur une plage. La valeur vide s'obtient aussi avec
' la touche Suppr (IgnoreBlank). Toute autre valeur est refusée.
Public Sub PoserListePointage(ByVal plage As Range)
    On Error Resume Next
    With plage.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=mod_VarGlobales.POINTE_OUI & ","
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
        .ErrorTitle = mod_Display.FR("Pointage")
        .ErrorMessage = mod_Display.FR("Choisissez Oui (op{e2}ration analys{e2}e) ou laissez la cellule vide.")
    End With
    On Error GoTo 0
End Sub

' Pose, à l'installation d'un formulaire, l'étiquette "Pointé", le champ de saisie (liste
' Oui / vide) et, si nomChamp n'est pas vide, un nom défini sur la cellule du champ (les
' modules retrouvent le champ par ce nom, jamais par son adresse).
Public Sub PoserChampPointage(ByVal ws As Worksheet, ByVal celluleLibelle As Range, _
                              ByVal celluleChamp As Range, Optional ByVal nomChamp As String = "")
    mod_InstallCommun.PoserEtiquette celluleLibelle, mod_Display.FR("Point{e2}")
    mod_InstallCommun.PoserChampSaisie celluleChamp, "@"
    celluleChamp.HorizontalAlignment = xlCenter
    PoserListePointage celluleChamp
    If nomChamp <> "" Then mod_InstallCommun.PoserNom ws, nomChamp, celluleChamp
End Sub

' Message de confirmation avant de RETIRER un pointage. Renvoie True si l'opérateur
' confirme. Le bouton par défaut est "Non" : un Entrée distrait ne dévalide rien.
Public Function ConfirmerDepointage(Optional ByVal nbOperations As Long = 1) As Boolean
    Dim texte As String
    If nbOperations <= 1 Then
        texte = mod_Display.FR("Vous allez retirer le pointage de cette op{e2}ration.") & vbCrLf & vbCrLf & _
                mod_Display.FR("Elle sera de nouveau consid{e2}r{e2}e comme {e2}tant {a2} analyser.")
    Else
        texte = mod_Display.FR("Vous allez retirer le pointage de ") & nbOperations & mod_Display.FR(" op{e2}rations.") & vbCrLf & vbCrLf & _
                mod_Display.FR("Elles seront de nouveau consid{e2}r{e2}es comme {e2}tant {a2} analyser.")
    End If
    texte = texte & vbCrLf & vbCrLf & mod_Display.FR("Confirmer le retrait du pointage ?")
    ConfirmerDepointage = (MsgBox(texte, vbYesNo + vbExclamation + vbDefaultButton2, _
                                  mod_Display.FR("Retirer le pointage")) = vbYes)
End Function


' =====================================================================================
' LECTURE / ÉCRITURE DANS TblOperations (par ID_Transaction)
' =====================================================================================

' Dictionnaire ID_Transaction -> numéro de ligne (1 = première ligne de données).
' Vide (0 élément) si le tableau est vide ; Nothing si le tableau est introuvable.
Private Function IndexLignesParID(ByVal tblOp As ListObject) As Object

    Dim d As Object
    Dim ids As Variant
    Dim i As Long, colId As Long

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set IndexLignesParID = d
    If tblOp Is Nothing Then Exit Function
    If tblOp.DataBodyRange Is Nothing Then Exit Function

    colId = mod_Display.GetColumnIndex(tblOp, COL_TECH_ID)
    If colId = 0 Then Exit Function

    ids = ValeursColonne(tblOp.ListColumns(colId).DataBodyRange)
    For i = 1 To UBound(ids, 1)
        If Not d.Exists(mod_DataStructure.CellText(ids(i, 1))) Then d.Add mod_DataStructure.CellText(ids(i, 1)), i
    Next i

End Function

' Valeurs d'une plage d'UNE colonne, toujours sous forme de tableau à 2 dimensions
' (Range.Value renvoie une simple valeur, et non un tableau, quand il n'y a qu'une cellule).
Private Function ValeursColonne(ByVal plage As Range) As Variant
    Dim v As Variant
    Dim t(1 To 1, 1 To 1) As Variant
    v = plage.Value
    If IsArray(v) Then
        ValeursColonne = v
    Else
        t(1, 1) = v
        ValeursColonne = t
    End If
End Function

' Pointage actuel d'une opération dans TblOperations. "trouve" vaut False si l'opération
' (ou la colonne Pointe) n'existe pas dans le tableau : la valeur renvoyée est alors False.
Public Function LirePointage(ByVal idTransaction As String, ByRef trouve As Boolean) As Boolean

    Dim tblOp As ListObject
    Dim index As Object
    Dim colP As Long

    trouve = False
    If idTransaction = "" Then Exit Function

    Set tblOp = OperationsSansMessage()
    If tblOp Is Nothing Then Exit Function
    colP = mod_Display.GetColumnIndex(tblOp, mod_VarGlobales.NOM_COL_POINTE)
    If colP = 0 Then Exit Function

    Set index = IndexLignesParID(tblOp)
    If Not index.Exists(idTransaction) Then Exit Function

    trouve = True
    LirePointage = TexteEnPointage(tblOp.ListColumns(colP).DataBodyRange.Cells(index(idTransaction), 1).Value)

End Function

' Écrit le pointage d'une opération. Renvoie False (sans message) si l'opération ou la
' colonne Pointe est introuvable.
Public Function EcrirePointage(ByVal idTransaction As String, ByVal pointe As Boolean) As Boolean

    Dim tblOp As ListObject
    Dim index As Object
    Dim colP As Long

    If idTransaction = "" Then Exit Function
    Set tblOp = OperationsSansMessage()
    If tblOp Is Nothing Then Exit Function
    colP = mod_Display.GetColumnIndex(tblOp, mod_VarGlobales.NOM_COL_POINTE)
    If colP = 0 Then Exit Function

    Set index = IndexLignesParID(tblOp)
    If Not index.Exists(idTransaction) Then Exit Function

    tblOp.ListColumns(colP).DataBodyRange.Cells(index(idTransaction), 1).Value = TextePointage(pointe)
    EcrirePointage = True

End Function

' TblOperations sans le message d'erreur de GetOperationsTable (les événements de saisie
' ne doivent jamais afficher de message technique parasite).
Private Function OperationsSansMessage() As ListObject
    On Error Resume Next
    Set OperationsSansMessage = ThisWorkbook.Worksheets("Import_data").ListObjects("TblOperations")
    On Error GoTo 0
End Function

' Vrai si la colonne Pointe est installée dans TblOperations.
Public Function PointageInstalle() As Boolean
    Dim tblOp As ListObject
    Set tblOp = OperationsSansMessage()
    If tblOp Is Nothing Then Exit Function
    PointageInstalle = (mod_Display.GetColumnIndex(tblOp, mod_VarGlobales.NOM_COL_POINTE) <> 0)
End Function


' =====================================================================================
' POINTAGE DES OPÉRATIONS D'UN IMPORT EN COURS
' =====================================================================================
' Pendant l'import, les nouvelles opérations ne sont pas encore dans TblOperations :
' leur pointage est conservé en mémoire par frm_ControleCategories. Cette mémoire
' permet aux formulaires ouverts depuis le contrôle (frm_Ventilation) d'afficher la
' bonne valeur.
Public Sub ViderPointageImport()
    Set g_PointageImport = Nothing
End Sub

Public Sub DefinirPointageImport(ByVal idTransaction As String, ByVal pointe As Boolean)
    If idTransaction = "" Then Exit Sub
    If g_PointageImport Is Nothing Then
        Set g_PointageImport = CreateObject("Scripting.Dictionary")
        g_PointageImport.CompareMode = 1
    End If
    g_PointageImport(idTransaction) = pointe
End Sub

Private Function PointageImport(ByVal idTransaction As String) As Boolean
    If g_PointageImport Is Nothing Then Exit Function
    If g_PointageImport.Exists(idTransaction) Then PointageImport = CBool(g_PointageImport(idTransaction))
End Function


' =====================================================================================
' FORMULAIRES : affichage du pointage d'une opération
' =====================================================================================
' Appelée par un formulaire chaque fois qu'il affiche une opération.
'   champ           : cellule de pointage du formulaire (une seule cellule)
'   idTransaction   : ID_Transaction de l'opération affichée (pour une part de
'                     ventilation : celui de l'opération d'origine)
'   ecritureDirecte : True = un changement est écrit tout de suite dans TblOperations
'                     (formulaires sans bouton "Annuler" pour le pointage) ;
'                     False = le formulaire relit la cellule lui-même à sa validation
'                     (frm_ControleCategories : "Annuler" ne doit rien modifier).
'   valeurAffichee  : (facultatif) valeur à afficher à la place de celle de la table
'                     (frm_ControleCategories, qui garde ses choix en mémoire).
' Si l'opération n'est pas encore dans TblOperations (import en cours) et que
' l'écriture est directe, la cellule est GRISE : elle affiche la valeur choisie à
' l'import mais ne se modifie que dans le contrôle des catégories.
Public Sub AfficherPointageFormulaire(ByVal ws As Worksheet, ByVal champ As Range, ByVal idTransaction As String, _
                                      Optional ByVal ecritureDirecte As Boolean = True, _
                                      Optional ByVal valeurAffichee As Variant)

    Dim pointe As Boolean, trouve As Boolean
    Dim evAvant As Boolean
    Dim cellule As Range

    If Not ChampPresent(champ) Then Exit Sub
    Set cellule = champ.Cells(1, 1)
    pointe = LirePointage(idTransaction, trouve)

    If Not IsMissing(valeurAffichee) Then
        pointe = CBool(valeurAffichee)
    ElseIf Not trouve Then
        pointe = PointageImport(idTransaction)
    End If

    If g_Champs Is Nothing Then
        Set g_Champs = CreateObject("Scripting.Dictionary")
        g_Champs.CompareMode = 1
    End If
    g_Champs(ws.Name) = Array(cellule.Address(False, False), idTransaction, ecritureDirecte)

    evAvant = Application.EnableEvents
    Application.EnableEvents = False
    On Error Resume Next
    cellule.Value = TextePointage(pointe)
    If ecritureDirecte And Not trouve Then
        cellule.Interior.Color = mod_InstallCommun.CoulFondLecture()
    Else
        cellule.Interior.Color = mod_InstallCommun.CoulFondSaisie()
    End If
    On Error GoTo 0
    Application.EnableEvents = evAvant

End Sub

' Vrai si le champ de pointage existe réellement sur la feuille : l'étiquette "Pointé" est
' dans la cellule de gauche et le champ n'est pas pris dans une fusion (cas d'un écran
' ancien, pas encore reconstruit : le pointage n'y est alors ni affiché ni modifiable).
Public Function ChampPresent(ByVal champ As Range) As Boolean
    On Error Resume Next
    If champ Is Nothing Then Exit Function
    If champ.Cells(1, 1).MergeArea.Cells.Count > 1 Then Exit Function
    ChampPresent = (mod_DataStructure.CellText(champ.Cells(1, 1).Offset(0, -1).Value) = mod_Display.FR("Point{e2}"))
End Function

' Variante pour les formulaires dont le champ est un NOM défini (rnPointe, gcPointe,
' ssPointe). Sans effet, et sans erreur, si la feuille n'a pas encore été reconstruite
' avec le champ de pointage.
Public Sub AfficherPointageParNom(ByVal ws As Worksheet, ByVal nomChamp As String, ByVal idTransaction As String)
    Dim champ As Range
    On Error Resume Next
    Set champ = ws.Range(nomChamp)
    On Error GoTo 0
    If champ Is Nothing Then Exit Sub
    AfficherPointageFormulaire ws, champ, idTransaction
End Sub

' ID_Transaction de la ligne ligneOp (1 = première ligne de données) de TblOperations.
Public Function IdDeLigneOperation(ByVal ligneOp As Long) As String
    Dim tblOp As ListObject
    Dim colId As Long
    If ligneOp < 1 Then Exit Function
    Set tblOp = OperationsSansMessage()
    If tblOp Is Nothing Then Exit Function
    If tblOp.DataBodyRange Is Nothing Then Exit Function
    If ligneOp > tblOp.DataBodyRange.Rows.Count Then Exit Function
    colId = mod_Display.GetColumnIndex(tblOp, COL_TECH_ID)
    If colId = 0 Then Exit Function
    IdDeLigneOperation = mod_DataStructure.CellText(tblOp.DataBodyRange.Cells(ligneOp, colId).Value)
End Function


' =====================================================================================
' ÉVÉNEMENT : modification d'une cellule (appelé par ThisWorkbook.Workbook_SheetChange)
' =====================================================================================
Public Sub TraiterChangementFeuille(ByVal Sh As Object, ByVal Target As Range)

    Dim info As Variant

    ' Sortie immédiate et peu coûteuse pour toute modification sans rapport.
    If Sh Is Nothing Then Exit Sub
    If Sh.Name = mod_VarGlobales.NOM_FEUILLE_RECHERCHE Then
        TraiterChangementRecherche Sh, Target
        Exit Sub
    End If

    If g_Champs Is Nothing Then Exit Sub
    If Not g_Champs.Exists(Sh.Name) Then Exit Sub
    info = g_Champs(Sh.Name)
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Address(False, False) <> CStr(info(0)) Then Exit Sub

    TraiterChangementFormulaire Sh, CStr(info(1)), CBool(info(2)), Target

End Sub

' --- Cellule de pointage d'un formulaire ------------------------------------------------
Private Sub TraiterChangementFormulaire(ByVal ws As Worksheet, ByVal idTransaction As String, _
                                        ByVal ecritureDirecte As Boolean, ByVal cellule As Range)

    Dim saisie As String
    Dim nouveau As Boolean, courant As Boolean, trouve As Boolean
    Dim reponse As String

    On Error GoTo Sortie
    saisie = mod_DataStructure.CellText(cellule.Value)

    ' Valeur inattendue (collage qui contourne la liste) : on remet l'état actuel.
    If saisie <> "" And LCase$(saisie) <> LCase$(mod_VarGlobales.POINTE_OUI) Then
        MsgBox mod_Display.FR("Pointage : choisissez Oui ou laissez la cellule vide."), vbExclamation
        RemettreEtatActuel cellule, idTransaction
        Exit Sub
    End If
    nouveau = (saisie <> "")

    courant = LirePointage(idTransaction, trouve)

    If Not trouve Then
        If ecritureDirecte Then
            ' Opération pas encore enregistrée (import en cours) : non modifiable ici.
            MsgBox mod_Display.FR("Cette op{e2}ration n'est pas encore enregistr{e2}e : son pointage se modifie dans le contr{o2}le des cat{e2}gories."), _
                   vbInformation, mod_Display.FR("Pointage")
            RemettreEtatActuel cellule, idTransaction
        Else
            AppliquerValeurCellule cellule, nouveau
        End If
        Exit Sub
    End If

    ' Retrait d'un pointage : confirmation obligatoire.
    If courant And Not nouveau Then
        If Not ConfirmerDepointage() Then
            AppliquerValeurCellule cellule, True
            Exit Sub
        End If
    End If

    If ecritureDirecte Then EcrirePointage idTransaction, nouveau
    AppliquerValeurCellule cellule, nouveau

Sortie:
    If Err.Number <> 0 Then
        reponse = Err.Description
        Application.EnableEvents = True
        MsgBox mod_Display.FR("Erreur lors de la modification du pointage : ") & reponse, vbExclamation
    End If

End Sub

' Remet dans la cellule l'état de pointage en vigueur (table, sinon mémoire d'import).
Private Sub RemettreEtatActuel(ByVal cellule As Range, ByVal idTransaction As String)
    Dim trouve As Boolean, pointe As Boolean
    pointe = LirePointage(idTransaction, trouve)
    If Not trouve Then pointe = PointageImport(idTransaction)
    AppliquerValeurCellule cellule, pointe
End Sub

' Écrit "Oui" ou vide dans la cellule SANS relancer l'événement.
Private Sub AppliquerValeurCellule(ByVal cellule As Range, ByVal pointe As Boolean)
    Dim evAvant As Boolean
    evAvant = Application.EnableEvents
    Application.EnableEvents = False
    On Error Resume Next
    cellule.Value = TextePointage(pointe)
    On Error GoTo 0
    Application.EnableEvents = evAvant
End Sub

' --- Colonne Pointe de l'écran de recherche ---------------------------------------------
' Une modification d'une cellule Pointe s'applique tout de suite à l'opération d'origine
' (TblOperations), puis TOUTES les lignes de l'écran qui concernent cette opération
' (l'opération et ses parts de ventilation) sont mises à jour. Le retrait d'un pointage
' demande confirmation ; en cas de refus, rien ne change.
Private Sub TraiterChangementRecherche(ByVal ws As Worksheet, ByVal Target As Range)

    Dim tblR As ListObject
    Dim colP As Long, colId As Long
    Dim zone As Range, cellule As Range
    Dim tblOp As ListObject, colOpP As Long
    Dim index As Object, aEcrire As Object
    Dim idTexte As String, saisie As String
    Dim ligneOp As Long
    Dim nouveau As Boolean, courant As Boolean
    Dim nbRetraits As Long
    Dim cle As Variant
    Dim ids As Variant, vals As Variant
    Dim i As Long
    Dim evAvant As Boolean
    Dim message As String

    On Error Resume Next
    Set tblR = ws.ListObjects(mod_VarGlobales.NOM_TABLE_RECHERCHE)
    On Error GoTo 0
    If tblR Is Nothing Then Exit Sub
    If tblR.DataBodyRange Is Nothing Then Exit Sub

    colP = mod_Display.GetColumnIndex(tblR, mod_VarGlobales.NOM_COL_POINTE)
    colId = mod_Display.GetColumnIndex(tblR, COL_TECH_ID)
    If colP = 0 Or colId = 0 Then Exit Sub

    Set zone = Intersect(Target, tblR.ListColumns(colP).DataBodyRange)
    If zone Is Nothing Then Exit Sub

    evAvant = Application.EnableEvents
    On Error GoTo Erreur
    Application.EnableEvents = False

    Set tblOp = OperationsSansMessage()
    If tblOp Is Nothing Then GoTo RemiseEtat
    colOpP = mod_Display.GetColumnIndex(tblOp, mod_VarGlobales.NOM_COL_POINTE)
    If colOpP = 0 Then
        MsgBox mod_Display.FR("La colonne Pointe n'est pas install{e2}e dans TblOperations (macro AjouterColonnePointage)."), vbExclamation
        GoTo RemiseEtat
    End If
    Set index = IndexLignesParID(tblOp)

    ' 1) Ce que l'opérateur demande : ID -> nouvel état. Une valeur inattendue annule tout.
    Set aEcrire = CreateObject("Scripting.Dictionary")
    aEcrire.CompareMode = 1
    For Each cellule In zone.Cells
        idTexte = mod_DataStructure.CellText(ws.Cells(cellule.Row, tblR.ListColumns(colId).Range.Column).Value)
        saisie = mod_DataStructure.CellText(cellule.Value)
        If saisie <> "" And LCase$(saisie) <> LCase$(mod_VarGlobales.POINTE_OUI) Then
            MsgBox mod_Display.FR("Pointage : choisissez Oui ou laissez la cellule vide."), vbExclamation
            aEcrire.RemoveAll
            GoTo RemiseEtat
        End If
        If idTexte <> "" Then aEcrire(idTexte) = (saisie <> "")
    Next cellule

    ' 2) Retraits de pointage : une seule confirmation pour l'ensemble.
    nbRetraits = 0
    For Each cle In aEcrire.Keys
        If index.Exists(CStr(cle)) Then
            courant = TexteEnPointage(tblOp.ListColumns(colOpP).DataBodyRange.Cells(index(CStr(cle)), 1).Value)
            If courant And Not CBool(aEcrire(cle)) Then nbRetraits = nbRetraits + 1
        End If
    Next cle
    If nbRetraits > 0 Then
        If Not ConfirmerDepointage(nbRetraits) Then
            aEcrire.RemoveAll       ' refus : l'écran est remis dans l'état de la table
            GoTo RemiseEtat
        End If
    End If

    ' 3) Écriture dans TblOperations.
    For Each cle In aEcrire.Keys
        If index.Exists(CStr(cle)) Then
            tblOp.ListColumns(colOpP).DataBodyRange.Cells(index(CStr(cle)), 1).Value = TextePointage(CBool(aEcrire(cle)))
        End If
    Next cle

RemiseEtat:
    ' 4) Toutes les lignes de l'écran reprennent l'état de TblOperations (opération et
    '    parts de ventilation d'une même opération restent cohérentes).
    If Not tblOp Is Nothing And colOpP <> 0 Then
        ids = ValeursColonne(tblR.ListColumns(colId).DataBodyRange)
        vals = ValeursColonne(tblR.ListColumns(colP).DataBodyRange)
        For i = 1 To UBound(ids, 1)
            idTexte = mod_DataStructure.CellText(ids(i, 1))
            If index.Exists(idTexte) Then
                ligneOp = index(idTexte)
                nouveau = TexteEnPointage(tblOp.ListColumns(colOpP).DataBodyRange.Cells(ligneOp, 1).Value)
                ' Écriture seulement si la cellule diffère (évite de réécrire des milliers de lignes).
                If mod_DataStructure.CellText(vals(i, 1)) <> TextePointage(nouveau) Then
                    tblR.ListColumns(colP).DataBodyRange.Cells(i, 1).Value = TextePointage(nouveau)
                End If
            End If
        Next i
    End If

    Application.EnableEvents = evAvant
    Exit Sub

Erreur:
    message = Err.Description
    Application.EnableEvents = True
    MsgBox mod_Display.FR("Erreur lors de la modification du pointage : ") & message, vbExclamation

End Sub
