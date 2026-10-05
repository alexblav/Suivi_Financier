Attribute VB_Name = "mod_Rapports"
Option Explicit

' =====================================================================================
' MODULE : mod_Rapports

' Ce module regroupe toute la logique de SELECTION, DE COMPTAGE ET DE TRI
' des opérations : "combien de dépenses correspondent au filtre ?",
' "charge-les toutes en mémoire", "trie-les par montant", "identifie les
' N plus grosses dépenses à surligner"...

' ----------------------------------------------------------------------
' TOperation : Type utilisé par la macro "Synthese_Care"
' Elle permet de stocker les dépenses par type
' ----------------------------------------------------------------------
Public Type TOperation
    LigneOrigine As Long    ' Stocke l'index de la ligne d'origine dans Excel
    Montant As Double       ' Valeur absolue du montant
    Notes As String         ' valeur du champs Notes de l'enregistrement
    Utilise As Boolean      ' Indique si la ligne a déjà été lettrée
End Type

' ----------------------------------------------------------------------
' AddExpenseForHighlight : ajoute une dépense candidate au surlignage dans
' deux tableaux parallèles : les numéros de ligne Excel, et les montants.
' ----------------------------------------------------------------------
' "ReDim Preserve" agrandit un tableau tout en conservant son contenu actuel.
' C'est moins performant qu'un ReDim unique fait à l'avance (voir plus haut),
' mais reste très rapide vu le faible nombre de dépenses par mois habituellement.
Public Sub AddExpenseForHighlight(ByRef rows() As Long, ByRef values() As Double, ByRef count As Long, ByVal rowNumber As Long, ByVal amount As Double)
    count = count + 1
    ReDim Preserve rows(1 To count)
    ReDim Preserve values(1 To count)
    rows(count) = rowNumber
    values(count) = amount
End Sub

' ----------------------------------------------------------------------
' GetTopCount : traduit le nombre demandé par l'utilisateur (cellule B3) en
' un nombre de lignes réellement utilisable.
' ----------------------------------------------------------------------
' Règle : si l'utilisateur demande 0, un nombre négatif, ou un nombre plus
' grand que ce qui est disponible, on affiche/surligne TOUT ce qui est
' disponible plutôt que de générer une erreur ou une liste tronquée à tort.
Public Function GetTopCount(ByVal requestedCount As Long, ByVal availableCount As Long) As Long
    If requestedCount <= 0 Or requestedCount > availableCount Then
        GetTopCount = availableCount
    Else
        GetTopCount = requestedCount
    End If
End Function

' ----------------------------------------------------------------------
' HighlightTopRows : trie les lignes candidates par montant décroissant,
' puis met en rouge le texte des "topCount" premières (les plus grosses
' dépenses du mois).
' ----------------------------------------------------------------------

Public Sub HighlightTopRows(ByRef rows() As Long, ByRef values() As Double, ByVal count As Long, ByVal topCount As Long)
    Dim i As Long
    Dim j As Long
    Dim maxIndex As Long
    Dim temporaryDouble As Double
    Dim temporaryLong As Long

    ' Même principe de tri par sélection que SortOperationsByAmount ci-dessus,
    ' mais appliqué ici à deux tableaux parallèles (rows/values) plutôt qu'à
    ' un tableau d'enregistrements.
    For i = 1 To count - 1
        maxIndex = i
        For j = i + 1 To count
            If values(j) > values(maxIndex) Then maxIndex = j
        Next j
        If maxIndex <> i Then
            temporaryDouble = values(i)
            values(i) = values(maxIndex)
            values(maxIndex) = temporaryDouble
            temporaryLong = rows(i)
            rows(i) = rows(maxIndex)
            rows(maxIndex) = temporaryLong
        End If
    Next i

    ' On applique la couleur rouge uniquement sur les "topCount" premières
    ' lignes désormais triées en tête du tableau.
    For i = 1 To topCount
        plageSortieEcriture.Cells(rows(i) - 1, posMontant).Font.Color = RGB(255, 0, 0)
    Next i
End Sub
