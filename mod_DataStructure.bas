Option Explicit
' Ce module regroupe toutes les fonctions de CONVERSION SECURISEE.
' "Sécurisée" veut dire : si la cellule contient une valeur inattendue
' (vide, texte, erreur Excel comme #N/A...), la fonction ne plante pas
' le programme. Elle renvoie une valeur par défaut (0, "" ou False) à la place.
' C'est le même principe partout : on préfère une valeur neutre à un plantage.
'
' Toutes ces fonctions sont regroupées ici (et non éparpillées dans les
' modules "métier") pour qu'il n'existe qu'UN SEUL endroit à modifier
' si on doit un jour changer la façon dont on lit une cellule Excel.

' Les macros de ce module ne sont pas visibles dans la liste "Macros" d'Excel
Option Private Module

' ----------------------------------------------------------------------
' CellText : transforme n'importe quelle valeur de cellule en texte "propre"
' ----------------------------------------------------------------------
' "Propre" signifie : jamais d'erreur, et sans espaces inutiles au début/fin
' (grâce à Trim$). Utilisée par exemple pour comparer un libellé ou un type
' d'opération ("Positif"/"Négatif") de façon fiable, même si la cellule
' contient des espaces superflus copiés depuis la banque.
Public Function CellText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        CellText = ""
    Else
        ' CStr convertit en texte, Trim$ retire les espaces avant/après.
        CellText = Trim$(CStr(value))
    End If
End Function

' ----------------------------------------------------------------------
' ToDouble : équivalent de ToLong, mais pour les nombres à virgule (Double)
' ----------------------------------------------------------------------
' Utilisée notamment pour lire des montants en euros, qui peuvent avoir
' des décimales (ex : 12.50).
Public Function ToDouble(ByVal value As Variant) As Double
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Or Not IsNumeric(value) Then
        ' Valeur invalide : on retourne 0 pour ne jamais fausser une somme.
        ToDouble = 0
    Else
        ToDouble = CDbl(value)
    End If
End Function

' ----------------------------------------------------------------------
' ToLong : version "simplifiée" de TryGetLong pour les cas où on veut juste
' un nombre, sans avoir besoin de savoir si la conversion a réussi.
' ----------------------------------------------------------------------
' Une valeur invalide devient simplement 0. Pratique pour lire un paramètre
' (comme B1, B2, B3 dans la feuille Synthese) sans avoir à écrire un test
' à chaque fois dans le code appelant.
Public Function ToLong(ByVal value As Variant) As Long
    Dim result As Long
    ' On réutilise TryGetLong : si elle réussit, on récupère sa valeur.
    ' Si elle échoue, "result" reste à 0 (valeur par défaut d'un Long en VBA),
    ' et ToLong renvoie donc 0 automatiquement (valeur par défaut d'une fonction).
    If TryGetLong(value, result) Then ToLong = result
End Function
