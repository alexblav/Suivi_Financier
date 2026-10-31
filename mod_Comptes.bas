Option Explicit

' ============================================================================
'  MODULE : mod_Comptes
'  RÔLE   : Gérer la liste des comptes bancaires (préparation du multi-comptes).
'
'  Le nom du compte n'est présent que dans le fichier CSV de la banque (colonne "Nom du compte").
'  À l'import (mod_ImportOFX), ce nom est cherché dans le tableau TblComptes (feuille Param) :
'  trouvé -> on réutilise sa clé; absent -> une nouvelle ligne est créée avec la clé suivante.
'  La clé est ensuite écrite dans la colonne CleCompte de TblOperations (Import_data).
'
'  Les opérations antérieures à cette évolution n'ont pas de clé : l'opérateur la saisit
'  directement dans Import_data (elles appartiennent toutes au même compte).
'
'  INSTALLATION : lancer UNE FOIS la macro InstallerGestionComptes (Ctrl+G). Elle crée le
'  tableau TblComptes et la colonne CleCompte; elle ne touche à rien si ils existent déjà.
' ============================================================================

Private Const COL_TBL_CLE As String = "CleCompte"
Private Const COL_TBL_NOM As String = "NomCompte"


' ----------------------------------------------------------------------------
' INSTALLATION : tableau TblComptes (Param) + colonne CleCompte (TblOperations)
' ----------------------------------------------------------------------------
Public Sub InstallerGestionComptes()

    Dim wsParam As Worksheet
    Dim tblComptes As ListObject, tblOp As ListObject
    Dim colonne As ListColumn
    Dim colDebut As Long
    Dim msgTbl As String, msgCol As String

    On Error Resume Next
    Set wsParam = ThisWorkbook.Worksheets(mod_Categories.NOM_FEUILLE_PARAM)
    On Error GoTo 0
    If wsParam Is Nothing Then
        MsgBox "La feuille " & mod_Categories.NOM_FEUILLE_PARAM & " est introuvable.", vbExclamation, "Comptes"
        Exit Sub
    End If

    ' --- Tableau TblComptes ---
    On Error Resume Next
    Set tblComptes = wsParam.ListObjects(mod_VarGlobales.NOM_TABLE_COMPTES)
    On Error GoTo 0

    If tblComptes Is Nothing Then
        colDebut = mod_VarGlobales.COMPTES_COL_DEBUT
        wsParam.Range(wsParam.Cells(1, colDebut), wsParam.Cells(1, colDebut + mod_VarGlobales.COMPTES_NB_COL - 1)).value = _
            Array(COL_TBL_CLE, COL_TBL_NOM)
        ' Le nom du compte peut ressembler à un nombre (n° de compte) : format TEXTE.
        wsParam.Range(wsParam.Cells(2, colDebut + 1), wsParam.Cells(2, colDebut + 1)).NumberFormat = "@"
        Set tblComptes = wsParam.ListObjects.Add(xlSrcRange, _
            wsParam.Range(wsParam.Cells(1, colDebut), wsParam.Cells(2, colDebut + mod_VarGlobales.COMPTES_NB_COL - 1)), , xlYes)
        tblComptes.Name = mod_VarGlobales.NOM_TABLE_COMPTES
        msgTbl = "Tableau " & mod_VarGlobales.NOM_TABLE_COMPTES & " cree (feuille " & mod_Categories.NOM_FEUILLE_PARAM & ")."
    Else
        msgTbl = "Tableau " & mod_VarGlobales.NOM_TABLE_COMPTES & " deja present."
    End If

    ' --- Colonne CleCompte dans TblOperations ---
    Set tblOp = mod_DonneesTable.GetOperationsTable()
    If tblOp Is Nothing Then Exit Sub

    On Error Resume Next
    Set colonne = tblOp.ListColumns(mod_VarGlobales.NOM_COL_CLE_COMPTE)
    On Error GoTo 0

    If colonne Is Nothing Then
        Set colonne = tblOp.ListColumns.Add
        colonne.Name = mod_VarGlobales.NOM_COL_CLE_COMPTE
        msgCol = "Colonne " & mod_VarGlobales.NOM_COL_CLE_COMPTE & " ajoutee a la fin de " & tblOp.Name & _
                 " (vide pour les operations existantes : saisissez-y la cle du compte)."
    Else
        msgCol = "Colonne " & mod_VarGlobales.NOM_COL_CLE_COMPTE & " deja presente (liste de validation heritee retiree)."
    End If

    ' Excel recopie sur une colonne ajoutée la validation (liste Oui) de la colonne voisine,
    ' Pointe : on la retire, et on retire aussi celle qui serait déjà héritée si la colonne existait.
    If Not colonne.DataBodyRange Is Nothing Then
        On Error Resume Next
        colonne.DataBodyRange.Validation.Delete
        On Error GoTo 0
    End If

    MsgBox msgTbl & vbCrLf & msgCol, vbInformation, "Comptes"

End Sub


' ----------------------------------------------------------------------------
' Renvoie la clé du compte dont le nom est fourni; le crée dans TblComptes s'il est inconnu.
' Renvoie 0 si TblComptes n'est pas installé (l'appelant avertit l'opérateur).
' ----------------------------------------------------------------------------
Public Function ObtenirOuCreerCleCompte(ByVal nomCompte As String) As Long

    Dim tblComptes As ListObject
    Dim donnees As Variant
    Dim i As Long, cleMax As Long, cleCourante As Long
    Dim nouvelleLigne As ListRow

    nomCompte = Trim$(nomCompte)
    If nomCompte = "" Then Exit Function

    Set tblComptes = TableComptes()
    If tblComptes Is Nothing Then Exit Function

    If Not tblComptes.DataBodyRange Is Nothing Then
        donnees = tblComptes.DataBodyRange.Value2
        For i = 1 To UBound(donnees, 1)
            If IsNumeric(donnees(i, 1)) And Not IsEmpty(donnees(i, 1)) Then
                cleCourante = CLng(donnees(i, 1))
                If cleCourante > cleMax Then cleMax = cleCourante
                If StrComp(Trim$(CStr(donnees(i, 2))), nomCompte, vbTextCompare) = 0 Then
                    ObtenirOuCreerCleCompte = cleCourante
                    Exit Function
                End If
            End If
        Next i
    End If

    ' Compte inconnu : nouvelle ligne. Un tableau vide n'a qu'une ligne vide à remplir.
    If tblComptes.DataBodyRange Is Nothing Then
        Set nouvelleLigne = tblComptes.ListRows.Add
    ElseIf tblComptes.ListRows.count = 1 And IsEmpty(tblComptes.DataBodyRange.Cells(1, 1).Value2) _
           And Trim$(CStr(tblComptes.DataBodyRange.Cells(1, 2).Value2)) = "" Then
        Set nouvelleLigne = tblComptes.ListRows(1)
    Else
        Set nouvelleLigne = tblComptes.ListRows.Add
    End If

    nouvelleLigne.Range.Cells(1, 2).NumberFormat = "@"
    nouvelleLigne.Range.Cells(1, 1).Value2 = cleMax + 1
    nouvelleLigne.Range.Cells(1, 2).Value2 = nomCompte
    ObtenirOuCreerCleCompte = cleMax + 1

End Function


' Tableau TblComptes (Nothing sans message s'il n'est pas installé).
Public Function TableComptes() As ListObject
    On Error Resume Next
    Set TableComptes = ThisWorkbook.Worksheets(mod_Categories.NOM_FEUILLE_PARAM).ListObjects(mod_VarGlobales.NOM_TABLE_COMPTES)
    On Error GoTo 0
End Function
