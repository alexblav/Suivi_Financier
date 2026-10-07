Attribute VB_Name = "mod_DonneesTable"
Option Explicit

' =====================================================================================
' MODULE : mod_DonneesTable

' Ce module regroupe tout ce qui concerne l'ACCÈS BRUT à la table Excel
' "TblOperations" (feuille Données) : la retrouver, lire une colonne
' par son nom, et vérifier si une ligne correspond à une période ou un filtre.
'
' Aucune macro de ce module n'écrit quoi que ce soit à l'écran : il ne fait
' que LIRE des informations et répondre par Vrai/Faux ou par une valeur.

' ----------------------------------------------------------------------
' GetOperationsTable : retrouve le tableau structuré Excel qui contient
' l'historique des opérations bancaires.
' ----------------------------------------------------------------------
' On utilise "On Error Resume Next" pour éviter un plantage si la feuille
' "Données" ou le tableau "TblOperations" a été renommé ou supprimé.
' Dans ce cas, la fonction retourne Nothing, et affiche un message clair
' plutôt que de laisser Excel afficher une erreur technique incompréhensible.
Public Function GetOperationsTable() As ListObject
    On Error Resume Next
    Set GetOperationsTable = ThisWorkbook.Worksheets("Import_data").ListObjects("TblOperations")
    On Error GoTo 0
    If GetOperationsTable Is Nothing Then MsgBox "Table TblOperations introuvable.", vbExclamation
End Function
Public Function GetOperationsValue(nomFeuille As String, nomPlage As String) As ListObject
    On Error Resume Next
    Set GetOperationsValue = ThisWorkbook.Worksheets(nomFeuille).ListObjects(nomPlage)
    On Error GoTo 0
    If GetOperationsValue Is Nothing Then MsgBox "Table " & nomPlage & " introuvable.", vbExclamation
End Function
' ----------------------------------------------------------------------
' RowMatchesPeriod : vérifie si une ligne du tableau correspond à l'année
' et au mois sélectionnés par l'utilisateur (cellules B1 et B2 de Synthese).
' ----------------------------------------------------------------------
' Astuce : si critAnnee ou critMois est une chaîne VIDE (""), le filtre
' correspondant est simplement ignoré ("toutes les années" ou "tous les mois").
Public Function RowMatchesPeriod(ByVal Mois As Long, ByVal Annee As Long) As Boolean
    RowMatchesPeriod = (Annee = CLng(critAnnee) And Mois = CLng(critMois))
End Function

' ----------------------------------------------------------------------
' RowMatchesFilter : vérifie si une ligne correspond à la période DEMANDÉE
' (voir RowMatchesPeriod ci-dessus) ET, si demandé, à un montant minimum.
' ----------------------------------------------------------------------
' Paramètre useMinimum : si False, on ne teste QUE la période (le seuil de
' montant est ignoré). Si True, on applique en plus le filtre de montant.
'
' Point important : pour une dépense (montant négatif, ex. : -300), on compare
' la VALEUR ABSOLUE du montant au seuil. En effet, -300 représente une
' dépense de 300 € : comparer "-300 >= seuil" donnerait un résultat faux
' pour n'importe quel seuil positif, ce qui ne correspond pas à l'intention
' de l'utilisateur ("n'afficher que les dépenses de plus de X euros").
Public Function RowMatchesFilter(ByVal nbLigne As Long, ByVal Mois As Long, ByVal Annee As Long, ByVal Montant As Double, ByVal useMinimum As Boolean) As Boolean
    Dim montantCompare As Double
    Dim typeOperation As String

    ' Étape 1 : le filtre de période est toujours appliqué en premier.
    RowMatchesFilter = RowMatchesPeriod(Mois, Annee)
    If Not RowMatchesFilter Then Exit Function   ' Pas la peine de continuer si la période ne correspond pas.
    If Not useMinimum Then Exit Function          ' Si on ne teste pas le montant, on s'arrête ici avec True.

    ' Étape 2 : uniquement si demandé, on applique le filtre de montant minimum.
    Montant = tblData(nbLigne, colMontant)
    If Montant < 0 Then
        ' Dépense : on compare la valeur absolue (voir explication ci-dessus).
        montantCompare = Abs(Montant)
    Else
        montantCompare = Montant
    End If
    RowMatchesFilter = montantCompare >= critMontantMin
End Function

'' ----------------------------------------------------------------------
'' TableValue : lit la valeur d'UNE colonne (identifiée par son NOM, pas son
'' numéro) pour UNE ligne donnée du tableau structuré.
'' ----------------------------------------------------------------------
'' Pourquoi chercher par nom de colonne plutôt que par numéro fixe (ex. : colonne 4) ?
'' Si un jour une colonne est ajoutée ou déplacée dans le tableau "Données",
'' ce code continuera de fonctionner correctement, car il retrouve toujours
'' "Montant" ou "Date" par leur intitulé, où qu'ils se trouvent désormais.
'Public Function TableValue(ByVal ligne As ListRow, ByVal tbl As ListObject, ByVal columnName As String) As Variant
'    TableValue = ligne.Range.Columns(tbl.ListColumns(columnName).index).Value2
'End Function
