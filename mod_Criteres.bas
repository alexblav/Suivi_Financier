Option Explicit
' Ce module regroupe tout ce qui concerne la LECTURE DES CRITERES saisis par
' l'utilisateur dans les cellules B1 à B6 de la feuille "Synthese"
' (année, mois, nombre d'opérations à surligner, seuil minimum, champ de tri,
' ordre de tri) — ainsi que la récupération de la feuille "Synthese" elle-même.

' Les macros de ce module ne sont pas visibles dans la liste "Macros" d'Excel
Option Private Module

' ----------------------------------------------------------------------
' GetFeuilleSynthese : retourne la feuille "Synthese" du classeur.
' ----------------------------------------------------------------------
' Centralise ThisWorkbook.Worksheets("Synthese") en un seul endroit :
' si le nom de cette feuille change un jour, une seule ligne est à corriger.
'Public Function GetFeuilleSynthese() As Worksheet
'    Set GetFeuilleSynthese = ThisWorkbook.Worksheets("Synthese")
'End Function
'
'Public Function GetFeuilleDonnees() As Worksheet
'    Set GetFeuilleSynthese = ThisWorkbook.Worksheets("Import_data")
'End Function

Public Function GetFeuille(nomFeuille As String) As Worksheet
    On Error Resume Next
    Set GetFeuille = ThisWorkbook.Worksheets(nomFeuille)
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------
' GetSelectCriteres : lit les 6 cellules de paramètres et les renvoie
' via des paramètres ByRef (c'est-à-dire que la macro qui appelle cette
' fonction reçoit les valeurs directement dans ses propres variables).
' ----------------------------------------------------------------------
' ws              : la feuille Synthese à lire (généralement issue de GetFeuilleSynthese)
' critAnnee        <- lit B1 : année sélectionnée
' critMois         <- lit B2 : mois sélectionné
' critNbOperations <- lit B3 : nombre d'opérations à mettre en évidence
' critMontantMin   <- lit B4 : seuil minimum de montant
' critTriChamps    <- lit B5 : champ utilisé pour trier le détail
' critTriOrdre     <- lit B6 : ordre de tri du détail (croissant/décroissant)
Public Sub GetSelectCriteres()
    If wsSynthese Is Nothing Then
        Set wsSynthese = mod_Criteres.GetFeuille(mod_VarGlobales.NOM_FEUILLE_SYNTHESE)
    End If
    critAnnee = mod_DataStructure.CellText(wsSynthese.Range("critAnnee").Value2)
    critMois = mod_DataStructure.CellText(wsSynthese.Range("critMois").Value2)
    critNbOperations = mod_DataStructure.ToLong(wsSynthese.Range("critNbLigne").Value2)
    critMontantMin = mod_DataStructure.ToDouble(wsSynthese.Range("critValMin").Value2)
    critTriChamps = mod_DataStructure.CellText(wsSynthese.Range("critTriChamp").Value2)
    critTriOrdre = mod_DataStructure.CellText(wsSynthese.Range("critTri").Value2)
End Sub

' ----------------------------------------------------------------------
' TryGetLong : tente de convertir une valeur de cellule en nombre entier (Long)
' ----------------------------------------------------------------------
' Pourquoi cette fonction existe : convertir directement avec CLng() plante
' avec une erreur si la valeur n'est pas un nombre (ex : cellule vide ou texte).
' Cette fonction vérifie D'ABORD que la conversion est possible, puis convertit.
'
' Le résultat est un Boolean : True si la conversion a réussi, False sinon.
' Cela permet à qui appelle cette fonction de savoir si la valeur récupérée
' (paramètre "result", passé par référence avec ByRef) est fiable ou non,
' plutôt que de devoir deviner ou de risquer un plantage.
Public Function TryGetLong(ByVal value As Variant, ByRef result As Long) As Boolean
    ' On vérifie 4 cas invalides avant toute conversion :
    '   - IsError  : la cellule contient une erreur Excel (#N/A, #REF!, etc.)
    '   - IsNull   : valeur nulle (rare avec Excel, plutôt issu de bases de données)
    '   - IsEmpty  : la cellule est complètement vide
    '   - Not IsNumeric : la valeur n'est pas interprétable comme un nombre
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Or Not IsNumeric(value) Then Exit Function
    ' Si on arrive ici, la valeur est bien numérique : la conversion est sûre.
    result = CLng(value)
    ' On indique que tout s'est bien passé en renvoyant True.
    TryGetLong = True
End Function
