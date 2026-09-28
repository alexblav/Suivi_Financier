Option Explicit
' Module stockant les macros d'action (attachée à un bouton ou déclenché sur événements dans la feuille) des formulaire

Public Sub Sortir()
    ActionSortir True
End Sub
Private Sub ActionSortir(Suppr As Boolean)
    Dim ActiveFeuille As String
    Dim ws As Worksheet
    
    ActiveFeuille = ActiveSheet.Name
    
    ' On crée un pointeur vers la feuille à supprimer
    Set ws = ThisWorkbook.Worksheets(ActiveFeuille)
    If ActiveFeuille = NOM_FEUILLE_RESULTAT And RecherOperations Then
        AppliquerLignesMarquees
    End If
    If Suppr Then
        ' Désactiver les messages d'avertissement d'Excel ("Voulez-vous vraiment supprimer...")
        Application.DisplayAlerts = False
        ' On appelle la méthode pour supprimer la feuille pointée
        ws.Delete
        ' On réactive obligatoirement les alertes pour ne pas perturber le reste d'Excel
        Application.DisplayAlerts = True
        
        If ActiveFeuille = mod_VarGlobales.NOM_FEUILLE_DETAIL_POSITIF Or ActiveFeuille = mod_VarGlobales.NOM_FEUILLE_DETAIL_NEGATIF Then
            ' On réaffiche la feuille Resultat
            wsResultat.Visible = xlSheetVisible
            
            ' On se repositionne dessus
            wsResultat.Activate
        Else
            'On réaffiche la feuille Synthese
            wsSynthese.Visible = xlSheetVisible
            
            ' On se repositionne dessus
            wsSynthese.Activate
            
            ' Arrêt complet du programme on est revenu au départ
            ' End réinitialise les variable en mémoire si on l'exécute trop tot l'exécution suivant s'arrête sur erreur
            End
        End If
    Else
        ' On masque la feuille en cours
        ws.Visible = xlSheetVeryHidden
        
        ' On réaffiche la feuille Synthese
        wsSynthese.Visible = xlSheetVisible
        
        ' On se repositionne dessus
        wsSynthese.Activate
    End If

End Sub

' On gére les actions a réaliser au moment d'un double clic sur une feuille
Public Sub DoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    Dim sortField As String
    Dim sortOrder As String
    
    If Sh.Name = NOM_FEUILLE_RESULTAT Then
        Call frm_Resultat_Double_Click(Sh, Target, Cancel)
    End If
End Sub

Private Sub frm_Resultat_Double_Click(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    Dim zoneCible As Range
    Dim zoneEntrees As Range
    Dim zoneDepenses As Range
    
    ' Une macro d'affichage ou de calcul peut temporairement interdire le détail ;
    ' dans ce cas l'événement quitte immédiatement sans modifier la feuille.
    ' Ignore tout double-clic tant qu un bilan n a pas activé le détail.
    If Not AllowDetailDoubleClick Then Exit Sub

    ' Un double-clic sur plusieurs cellules n'est pas interprété comme une
    ' demande de détail afin d'éviter toute ambiguïté sur le type à afficher.
    ' Refuse une sélection multiple qui ne désigne pas un total unique.
    If Target.CountLarge > 1 Then Exit Sub

    ' Seules les cellules "Total Entrées" et "Total Dépenses", qui portent les libellés des totaux, sont
    ' interactives. Les autres cellules conservent le comportement Excel normal.
    ' Restreint l interaction aux deux cellules portant les totaux.

    ' Définir la zone cible : cellSortieDep + 9 lignes vers le bas, 1 col vers la droite, sur 2 lignes de haut et 1 col de large (B10:B11) si cellSortieDep=A1
    Set zoneCible = wsResultat.Range(cellSortieDep).Offset(9, 1).Resize(2, 1)

    ' Vérifier l'intersection entre la cellule cliquée (Target) et la zone cible
    If Intersect(Target, zoneCible) Is Nothing Then Exit Sub
    
    'If Intersect(Target, Target.Range(wsResultat.Range(cellSortieDep).Offset(9, 1), wsResultat.Range(cellSortieDep).Offset(10, 1))) Is Nothing Then Exit Sub

    ' L'événement est annulé pour empêcher Excel d'entrer en mode d'édition
    ' lorsque le double-clic est reconnu comme une commande de détail.
    ' Empêche Excel d entrer en mode édition après le double-clic reconnu.
    Cancel = True

    ' Lire les critères saisis par l'utilisateur (B1 à B6).
    mod_Criteres.GetSelectCriteres

    'sortField = InputBox("Champ de tri : Date, Libellé, Catégorie, Montant, MoisBudget, AnnéeBudget ou Budget.", "Choix du champ de tri", "Date")
    If critTriChamps = "" Then Exit Sub
    If Not IsValidDetailSortField(critTriChamps) Then
        MsgBox "Champ de tri inconnu. Le détail ne sera pas créé.", vbExclamation
        Exit Sub
    End If

    'sortOrder = InputBox("Ordre de tri : Croissant ou Décroissant.", "Choix de l'ordre de tri", "Croissant")
    If critTriOrdre = "" Then Exit Sub
    If LCase$(Trim$(critTriOrdre)) <> "croissant" And LCase$(Trim$(critTriOrdre)) <> "décroissant" And LCase$(Trim$(critTriOrdre)) <> "decroissant" Then
        MsgBox "Ordre de tri inconnu. Le détail ne sera pas créé.", vbExclamation
        Exit Sub
    End If

    ' E1 correspond au total des entrées et E2 au total des dépenses ; le type
    ' transmis pilote ensuite le filtrage réalisé par ShowDetailForTotal.
    ' E1 demande le détail des entrées et E2 celui des dépenses.
    Set zoneEntrees = wsResultat.Range(cellSortieDep).Offset(9, 1)
    Set zoneDepenses = wsResultat.Range(cellSortieDep).Offset(10, 1)
    If Target.Address(False, False) = zoneEntrees.Address(False, False) Then
        mod_SyntheseBudgetBilanMensuel.ShowDetailForTotal "Positif", critTriChamps, critTriOrdre
    ElseIf Target.Address(False, False) = zoneDepenses.Address(False, False) Then
        mod_SyntheseBudgetBilanMensuel.ShowDetailForTotal "Négatif", critTriChamps, critTriOrdre
    End If
End Sub

' Vérifie que le champ demandé correspond à une colonne du détail.
Public Function IsValidDetailSortField(ByVal sortField As String) As Boolean
    Select Case LCase$(Trim$(sortField))
        Case "date", "libellé", "libelle", "catégorie", "categorie", "montant", "moisbudget", "annéebudget", "anneebudget", "budget"
            IsValidDetailSortField = True
    End Select
End Function

' ------------------------------------------------------------------------
' ResolveDetailSortOrder : traduit "Croissant"/"Décroissant" (texte saisi
' par l'utilisateur) en constante Excel xlAscending/xlDescending.
' ------------------------------------------------------------------------
Public Function ResolveDetailSortOrder(ByVal critTriOrdre As String) As Long
    If LCase$(Trim$(critTriOrdre)) = "décroissant" Or LCase$(Trim$(critTriOrdre)) = "decroissant" Then
        ResolveDetailSortOrder = xlDescending
    Else
        ResolveDetailSortOrder = xlAscending
    End If
End Function

Sub AfficherFeuilleNotesPourEdition(ByVal nomFeuille As String)
    Dim ws As Worksheet
    Set ws = ObtenirFeuilleSansErreurFN(nomFeuille)
    If ws Is Nothing Then
        MsgBox "La feuille '" & nomFeuille & "' n'existe pas.", vbExclamation
        Exit Sub
    End If
    ws.Visible = xlSheetVisible
    ws.Activate
    MsgBox "La feuille est maintenant visible. Pense a la remasquer avec" & vbCrLf & _
           "MasquerFeuilleNotesApresEdition " & Chr(34) & nomFeuille & Chr(34), vbInformation
End Sub

Sub MasquerFeuilleNotesApresEdition(ByVal nomFeuille As String)
    Dim ws As Worksheet
    Set ws = ObtenirFeuilleSansErreurFN(nomFeuille)
    If ws Is Nothing Then
        MsgBox "La feuille '" & nomFeuille & "' n'existe pas.", vbExclamation
        Exit Sub
    End If
    ws.Visible = xlSheetVeryHidden
    MsgBox "La feuille est de nouveau masquee.", vbInformation
End Sub

' =====================================================================================
' FONCTIONS UTILITAIRES (creation/suppression feuille, noms, formes)
' =====================================================================================
Private Function ObtenirFeuilleSansErreurFN(ByVal nomFeuille As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nomFeuille)
    On Error GoTo 0
    Set ObtenirFeuilleSansErreurFN = ws
End Function
