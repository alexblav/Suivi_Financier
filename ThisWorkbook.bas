Option Explicit
' Ce module regroupe les macros qui préparent une ZONE D'AFFICHAGE sur une
' feuille (nettoyage, mise en place des en-têtes) AVANT que d'autres modules
' n'y écrivent des données. Il ne contient volontairement AUCUNE logique de
' calcul ou de filtrage : uniquement de la mise en forme / nettoyage visuel.

' Cet événement est déclenché automatiquement lorsqu'un utilisateur double-clique
' sur une cellule de une feuille du classeur
' Son rôle est de transformer le double-clic en action
Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    If Sh.Name <> mod_VarGlobales.NOM_FEUILLE_RESULTAT And AllowDetailDoubleClick Then Exit Sub
    DoubleClick Sh, Target, Cancel
End Sub

' Est exécuté à chaque changement opéré dans la feuille
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)

    If Sh.Name = mod_VarGlobales.NOM_FEUILLE_RESULTAT And RecherOperations Then

    'On Error GoTo Fin

    Dim tblRecherche As Range
    Dim inter As Range
    Dim zoneSurveillee As Range
    Dim colValiderRech As Long
    
    ' ---- DEBUG 1 : Vérifier où le code cherche les en-têtes ----
    ' Si l'adresse des en-têtes pointe sur vos données (ex: B10:I10) au lieu de vos titres (ex: B9:I9),
    ' alors le problème vient de là !
'    MsgBox "1. Plage données totale : " & plageSortieEcriture.Address & vbCrLf & _
'           "2. Ligne des en-têtes supposée : " & plageSortieEnTetes.Address
    
    If Not Intersect(Target, plageSortieEnTetes) Is Nothing Then Exit Sub
    
    ' 3. Récupération de l'ID de la colonne "Valider"
    'colValiderRech = GetPosArray("Valider", plageSortieEnTetes)
    
'    If IsError(posCategorie) Or IsError(posNotes) Or IsError(colValiderRech) Then
'        MsgBox "ERREUR : Impossible de trouver un des en-têtes dans " & entetes.Address
'        Exit Sub
'    End If

    Set zoneSurveillee = Union(plageSortieEcriture.Columns(posCategorie), plageSortieEcriture.Columns(posNotes))

'    ' ---- DEBUG 2 : Vérifier les adresses ----
'    MsgBox "3. Colonnes surveillées : " & zoneSurveillee.Address & vbCrLf & _
'           "4. Cellule modifiée (Target) : " & Target.Address

    Set inter = Intersect(Target, zoneSurveillee)
    
    If inter Is Nothing Then
        MsgBox "ECHEC INTERSECTION : " & Target.Address & " n'est pas dans " & zoneSurveillee.Address
        Exit Sub
    End If

    ' Si on arrive ici, ça fonctionne !
    ' MsgBox "SUCCES ! Intersection trouvée sur : " & inter.Address

    Application.EnableEvents = False
    Dim c As Range
    For Each c In inter.Cells
    
        Intersect(c.EntireRow, plageSortieEcriture.Columns(posValider)).value = "OUI"
    Next c
    Application.EnableEvents = True

    End If

Fin:
    Application.EnableEvents = True
    
End Sub