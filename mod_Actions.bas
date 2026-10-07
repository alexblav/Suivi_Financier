' =====================================================================================
' MODULE : mod_Actions

Option Explicit
' Module regroupant les macros d'action associées aux boutons ou déclenchées par les événements des feuilles de formulaire.

Public Sub Sortir()
    ActionSortir True
End Sub

Public Sub RchDernierImport()
  mod_RechercheOperations.RechercherOperations "DernierImport"
End Sub

Public Sub RchOperationsDuMois()
  mod_RechercheOperations.RechercherOperations "OperationsDuMois"
End Sub

Public Sub RchErreursSante()
  mod_RechercheOperations.RechercherOperations "ErreursSante"
End Sub

' AJOUT 03/10/2026 (demande opérateur) : recherche "générale", sans aucun
' préfiltre; affiche TOUTES les opérations de TblOperations et de
' TblVentilations. Correspond à l'appel sans argument de
' RechercherOperations, dont le paramètre "prefiltre" vaut "" par défaut.
Public Sub RchGenerale()
  mod_RechercheOperations.RechercherOperations
End Sub

Private Sub ActionSortir(Suppr As Boolean)
    Dim ActiveFeuille As String
    Dim ws As Worksheet
    
    ActiveFeuille = ActiveSheet.Name

    ' On crée un pointeur vers la feuille à supprimer
    Set ws = ThisWorkbook.Worksheets(ActiveFeuille)
    ' PHASE 6 : l'ancien appel à AppliquerLignesMarquees depuis l'écran de
    ' recherche a été supprimé avec mod_SyntheseRechercheOperations. Ce mécanisme
    ' n'est donc plus utilisé.
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
            ' End réinitialise les variables en mémoire. S'il est exécuté trop tôt,
            ' l'exécution suivante s'arrête sur une erreur.
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

' PHASE 6 : DoubleClick / frm_Resultat_Double_Click (appelées depuis
' ThisWorkbook.Workbook_SheetBeforeDoubleClick, désormais supprimé), ainsi que
' leurs fonctions utilitaires IsValidDetailSortField et ResolveDetailSortOrder,
' ont été supprimées. Elles ne servaient qu'à l'ancien mécanisme de double-clic
' sur les totaux du Bilan mensuel, remplacé par les deux boutons explicites
' "VoirDetailEntreesRO" / "VoirDetailDepensesRO" (mod_SyntheseBudgetBilanMensuel.bas),
' qui ouvrent directement frm_RechercheOperations (préfiltre "DetailTotal").
' Le tri automatique par champ n'a pas été repris : l'opérateur trie désormais
' lui-même avec les flèches de filtre natives du nouvel écran.

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
' FONCTIONS UTILITAIRES (création/suppression de feuilles, noms et formes)
' =====================================================================================
Private Function ObtenirFeuilleSansErreurFN(ByVal nomFeuille As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nomFeuille)
    On Error GoTo 0
    Set ObtenirFeuilleSansErreurFN = ws
End Function
