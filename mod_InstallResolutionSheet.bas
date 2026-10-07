Option Explicit

' =====================================================================================
' MODULE : mod_InstallResolutionSheet
'
' RÔLE (phase 1 du chantier "UserForm -> feuille dédiée") :
'   Ce module contient UNIQUEMENT la construction de la mise en page statique de
'   la feuille qui remplacera le UserForm "frmResolutionCategories".
'   Il ne contient ENCORE AUCUNE logique de remplissage des cas ambigus ni de
'   gestion des clics : cela viendra en phase 2 (module mod_ResolutionCategories).
'
'   Pourquoi séparer ainsi ? Pour que tu puisses exécuter cette phase 1 seule,
'   vérifier visuellement que la feuille correspond à la maquette validée,
'   AVANT d'y ajouter le comportement (phase 2). C'est le même principe que
'   "construire les murs avant de brancher l'électricité".
'
' À FAIRE POUR INSTALLER CE MODULE :
'   1. Alt+F11 pour ouvrir l'éditeur VBA.
'   2. Fichier > Importer un fichier... > sélectionner ce fichier .bas.
'   3. Dans la fenêtre Exécution immédiate (Ctrl+G), taper :
'        CreerFeuilleResolutionCategories
'      puis appuyer sur Entrée. La feuille est créée, mise en forme, puis masquée.
' =====================================================================================

' -------------------------------------------------------------------------------------
' CONSTANTES DE MISE EN PAGE
' -------------------------------------------------------------------------------------
' Regrouper ici toutes les positions de cellules évite d'avoir des "nombres magiques"
' dispersés dans le code. Si tu veux déplacer une zone, tu n'as qu'UNE ligne à modifier
' ici au lieu de la chercher partout dans le code.
' Ces constantes seront réutilisées telles quelles dans le module de la phase 2.

Public Const NOM_FEUILLE_RESOLUTION As String = "frm_ResolutionCategories"

' --- Zone des boutons (ligne 2) ---
Public Const LIGNE_BOUTONS As Long = 2

' --- Zone des instructions ---
Public Const LIGNE_TITRE_INSTRUCTIONS As Long = 5
Public Const LIGNE_DEBUT_TEXTE_INSTRUCTIONS As Long = 6
Public Const LIGNE_FIN_TEXTE_INSTRUCTIONS As Long = 9




' =====================================================================================
' MACRO PRINCIPALE D'INSTALLATION
' À exécuter UNE SEULE FOIS (ou de nouveau si tu veux réinitialiser complètement
' la mise en forme de la feuille).
' =====================================================================================
Sub CreerFeuilleResolutionCategories()

    Dim ws As Worksheet
    Dim reponseUtilisateur As VbMsgBoxResult

    ' ---------------------------------------------------------------------------------
    ' ÉTAPE A - Vérifier si la feuille existe déjà, pour éviter d'écraser du travail
    ' sans prévenir. On utilise une fonction utilitaire (définie plus bas) qui tente
    ' de retrouver la feuille par son nom sans provoquer d'erreur si elle n'existe pas.
    ' ---------------------------------------------------------------------------------
    Set ws = ObtenirFeuilleSansErreur(NOM_FEUILLE_RESOLUTION)

    If Not ws Is Nothing Then
        ' La feuille existe déjà : on demande confirmation avant de tout reconstruire,
        ' car sa mise en forme actuelle sera effacée.
        reponseUtilisateur = MsgBox( _
            "La feuille '" & NOM_FEUILLE_RESOLUTION & "' existe deja." & vbCrLf & _
            "Veux-tu la reconstruire entierement (sa mise en forme actuelle sera perdue) ?", _
            vbYesNo + vbQuestion, "Confirmation de reconstruction")

        If reponseUtilisateur = vbNo Then
            MsgBox "Installation annulee, aucune modification effectuee.", vbInformation
            Exit Sub
        End If

        ' On la rend temporairement visible : impossible de la modifier ou de la supprimer
        ' correctement tant qu'elle est en xlSheetVeryHidden.
        ws.Visible = xlSheetVisible
        ws.Cells.Clear             ' on efface le contenu et la mise en forme
        Call SupprimerFormesExistantes(ws)   ' on retire les anciens boutons, s'il y en a
    Else
        ' La feuille n'existe pas encore : on la crée en dernière position
        ' pour ne pas perturber l'ordre des onglets existants (Accueil, Synthese...).
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = NOM_FEUILLE_RESOLUTION
    End If

    ' ---------------------------------------------------------------------------------
    ' ÉTAPE B - Mise en forme générale de la feuille (aspect "formulaire")
    ' ---------------------------------------------------------------------------------
    Call AppliquerMiseEnFormeGenerale(ws)

    ' ---------------------------------------------------------------------------------
    ' ÉTAPE C - Construction de la zone des boutons
    ' ---------------------------------------------------------------------------------
    Call ConstruireZoneBoutons(ws)

    ' ---------------------------------------------------------------------------------
    ' ÉTAPE D - Construction de la zone des instructions
    ' ---------------------------------------------------------------------------------
    Call ConstruireZoneInstructions(ws)

    ' ---------------------------------------------------------------------------------
    ' ÉTAPE E - Construction des en-têtes et mise en forme du tableau des cas
    ' ---------------------------------------------------------------------------------
    Call ConstruireTableauCas(ws)

    ' ---------------------------------------------------------------------------------
    ' ÉTAPE F - Figer les volets pour que les boutons, les instructions et les en-têtes
    ' restent visibles pendant que l'utilisateur fait défiler la liste des cas.
    ' ---------------------------------------------------------------------------------
    Call FigerVoletsSousEntetes(ws)

    ' ---------------------------------------------------------------------------------
    ' ÉTAPE G - Masquer complètement la feuille (invisible pour l'utilisateur final,
    ' elle ne sera affichée que lorsqu'un cas ambigu devra être résolu; ce déclenchement
    ' sera codé en phase 2).
    ' ---------------------------------------------------------------------------------
    ws.Visible = xlSheetVeryHidden

    MsgBox "La feuille '" & NOM_FEUILLE_RESOLUTION & "' a ete creee et mise en forme," & vbCrLf & _
           "puis masquee (xlSheetVeryHidden)." & vbCrLf & vbCrLf & _
           "Pour la revoir a l'ecran et verifier le rendu, execute la macro :" & vbCrLf & _
           "   AfficherFeuilleResolutionPourEdition", _
           vbInformation, "Installation terminee"

End Sub


' =====================================================================================
' SOUS-PROCÉDURE - Mise en forme générale (suppression du quadrillage, police, etc.)
' =====================================================================================
Private Sub AppliquerMiseEnFormeGenerale(ws As Worksheet)

    ' On doit activer la feuille pour modifier certains réglages de la FENÊTRE
    ' (comme l'affichage du quadrillage), car ils sont associés à la fenêtre au
    ' moment où une feuille est affichée, et non à la feuille elle-même.
    ws.Activate

    ' Supprime le quadrillage (les lignes grises entre les cellules) pour donner
    ' un aspect "formulaire" plutôt que "tableur classique", comme demandé.
    ActiveWindow.DisplayGridlines = False

    ' On repart d'une largeur de colonnes cohérente pour toute la zone utile.
    ws.Columns("A").ColumnWidth = 2      ' petite marge à gauche
    ws.Columns("B").ColumnWidth = 8      ' colonne "Statut"
    ws.Columns("C").ColumnWidth = 12     ' colonne "Date"
    ws.Columns("D").ColumnWidth = 12     ' colonne "Montant"
    ws.Columns("E").ColumnWidth = 28     ' colonne "Tiers"
    ws.Columns("F").ColumnWidth = 24     ' colonne "Categorie(s)"
    ws.Columns("G").ColumnWidth = 2      ' petite marge à droite

    ' Police par défaut de toute la feuille, cohérente avec un rendu "formulaire" sobre.
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10

    ' On se replace en A1 et on désactive les en-têtes de lignes et de colonnes
    ' (références L1C1 grisées) pour un rendu plus épuré. Optionnel, mais cela renforce
    ' l'effet "formulaire" plutôt que "feuille de calcul".
    ws.Range("A1").Select
    ActiveWindow.DisplayHeadings = False

End Sub


' =====================================================================================
' SOUS-PROCÉDURE - Construction de la zone des boutons (ligne 2) et du compteur
' =====================================================================================
Private Sub ConstruireZoneBoutons(ws As Worksheet)

    Dim boutonValider As Button
    Dim boutonTerminer As Button
    Dim zoneBoutonValider As Range
    Dim zoneBoutonTerminer As Range

    ' On définit la position de chaque bouton en s'appuyant sur une PLAGE DE CELLULES
    ' (et non sur des coordonnées fixes en pixels) : ainsi, si la largeur des
    ' colonnes change, les boutons suivent automatiquement. C'est plus robuste qu'un
    ' positionnement en pixels pour un environnement qui doit rester stable.
    Set zoneBoutonValider = ws.Range(COL_STATUT & LIGNE_BOUTONS & ":" & COL_DATE & LIGNE_BOUTONS)
    Set zoneBoutonTerminer = ws.Range(COL_MONTANT & LIGNE_BOUTONS & ":" & COL_TIERS & LIGNE_BOUTONS)

    zoneBoutonValider.RowHeight = 22
    zoneBoutonTerminer.RowHeight = 22

    ' Ajout d'un bouton de type "Contrôle de formulaire" (PAS ActiveX), comme convenu :
    ' plus léger et plus fiable face au problème initial de DPI et de multi-écrans.
    Set boutonValider = ws.Buttons.Add( _
        zoneBoutonValider.Left, zoneBoutonValider.Top, _
        zoneBoutonValider.Width, zoneBoutonValider.Height)
    With boutonValider
        ' IMPORTANT : le texte affiché est déjà son intitulé FINAL, validé avec toi,
        ' à savoir sa fonction de réinitialisation (et non plus "Valider ce cas").
        .Caption = "R?initialiser cette ligne"
        ' Le nom de macro ci-dessous n'existe pas encore : il sera défini en phase 2.
        ' Cela ne provoque AUCUNE erreur pour l'instant; une erreur ne surviendrait
        ' que si quelqu'un cliquait sur le bouton avant l'installation de la phase 2.
        .OnAction = "ReinitialiserLigneSelectionnee"
        .Name = "btnReinitialiserLigne"
    End With

    Set boutonTerminer = ws.Buttons.Add( _
        zoneBoutonTerminer.Left, zoneBoutonTerminer.Top, _
        zoneBoutonTerminer.Width, zoneBoutonTerminer.Height)
    With boutonTerminer
        .Caption = "Terminer et appliquer"
        .OnAction = "TerminerEtAppliquerChoix"
        .Name = "btnTerminerResolution"
    End With

    ' Cellule du compteur ("X restant(s) sur Y"), à droite de la barre de boutons.
    ' On lui attribue un NOM DÉFINI (plage nommée) plutôt qu'une référence codée en dur
    ' comme "F2" : le code de la phase 2 est ainsi plus lisible
    ' (Range("CompteurCasRestants") est plus explicite que Range("H2")).
    With ws.Range(COL_TIERS & LIGNE_BOUTONS & ":" & COL_CATEGORIE & LIGNE_BOUTONS)
        .Merge
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .Font.Size = 9
        .Font.Color = RGB(120, 120, 120)   ' gris discret, texte d'information secondaire
        .value = ""   ' rempli dynamiquement en phase 2
    End With
    ws.Names.Add Name:="CompteurCasRestants", RefersTo:=ws.Range(COL_TIERS & LIGNE_BOUTONS)

End Sub


' =====================================================================================
' SOUS-PROCÉDURE - Construction de la zone des instructions
' =====================================================================================
Private Sub ConstruireZoneInstructions(ws As Worksheet)

    ' --- Titre "Instructions" ---
    With ws.Range(COL_STATUT & LIGNE_TITRE_INSTRUCTIONS)
        .value = "Instructions"
        .Font.Bold = True
        .Font.Size = 11
    End With

    ' --- Bloc de texte des quatre étapes, sur une plage fusionnée pour ressembler à un
    ' encadré de note plutôt qu'à des cellules de tableur classiques. ---
    Dim zoneTexte As Range
    Set zoneTexte = ws.Range( _
        COL_STATUT & LIGNE_DEBUT_TEXTE_INSTRUCTIONS & ":" & _
        COL_CATEGORIE & LIGNE_FIN_TEXTE_INSTRUCTIONS)

    With zoneTexte
        .Merge
        ' Le texte exact que tu m'as fourni, avec des sauts de ligne internes
        ' (Chr(10) est le caractère "retour à la ligne" à l'intérieur d'une cellule).
        .value = "Les op?rations list?es n'ont pas pu ?tre cat?goris?es de fa?on automatique." & Chr(10) & _
                 "Il faut donc le faire manuellement. Pour ce faire suivre les ?tapes suivantes :" & Chr(10) & _
                 "1. S?lectionner une cat?gorie dans la liste d?roulante en face de l'op?ration concern?e" & Chr(10) & _
                 "2. Le statut de la ligne passe automatiquement ? "" OK "" une fois la cat?gorie choisie" & Chr(10) & _
                 "3. Recommencer pour chaque op?ration puis cliquer sur ""Terminer et appliquer"""
        .WrapText = True                     ' le texte revient à la ligne dans la cellule
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
        .Font.Size = 9
        .Font.Color = RGB(80, 80, 80)
        .Interior.Color = RGB(245, 245, 242) ' fond gris très clair, type "encadré note"
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(210, 210, 205)
        .Locked = True                       ' cellule non modifiable par l'utilisateur final
    End With

    ' Remarque : le texte des instructions ci-dessus a été légèrement adapté par
    ' rapport à l'original du UserForm, car les étapes 1 à 3 de l'ancien texte
    ' décrivaient le fonctionnement de l'option A (clic sur une ligne, puis liste
    ' partagée et bouton Valider). Comme nous sommes passés à l'option B (liste
    ' déroulante directement sur chaque ligne et statut automatique), le texte a
    ' été adapté pour rester conforme au nouveau fonctionnement.
    ' Dis-moi si tu préfères une autre formulation : il s'agit d'une simple
    ' modification de texte, sans impact sur le reste du code.

    ' Ajuste la hauteur des lignes du bloc pour laisser de la place au texte.
    ws.rows(LIGNE_DEBUT_TEXTE_INSTRUCTIONS & ":" & LIGNE_FIN_TEXTE_INSTRUCTIONS).RowHeight = 16

End Sub


' =====================================================================================
' SOUS-PROCÉDURE - Construction des en-têtes et mise en forme du tableau
' =====================================================================================
Private Sub ConstruireTableauCas(ws As Worksheet)

    ' --- Ligne d'en-têtes (row LIGNE_ENTETES_TABLEAU) ---
    With ws.Range(COL_STATUT & LIGNE_ENTETES_TABLEAU & ":" & COL_CATEGORIE & LIGNE_ENTETES_TABLEAU)
        .Font.Bold = True
        .Font.Size = 9.5
        .Font.Color = RGB(90, 90, 90)
        .Interior.Color = RGB(250, 250, 248)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(200, 200, 195)
        .Borders(xlEdgeBottom).Weight = xlMedium
        .VerticalAlignment = xlCenter
    End With
    ws.Range(COL_STATUT & LIGNE_ENTETES_TABLEAU).value = "Statut"
    ws.Range(COL_DATE & LIGNE_ENTETES_TABLEAU).value = "Date"
    ws.Range(COL_MONTANT & LIGNE_ENTETES_TABLEAU).value = "Montant"
    ws.Range(COL_TIERS & LIGNE_ENTETES_TABLEAU).value = "Tiers"
    ws.Range(COL_CATEGORIE & LIGNE_ENTETES_TABLEAU).value = "Cat?gorie(s)"
    ws.rows(LIGNE_ENTETES_TABLEAU).RowHeight = 20

    ' --- Mise en forme "à blanc" des lignes de données préparées à l'avance ---
    ' On ne remplit PAS encore de données réelles ici (ce sera fait dynamiquement en
    ' phase 2, à chaque ouverture, selon le nombre réel de cas ambigus). On prépare
    ' uniquement l'apparence (bordures légères, alternance de fond, alignement),
    ' afin que la feuille soit déjà présentable dès cette phase 1.
    Dim ligneCourante As Long
    Dim derniereLigne As Long
    derniereLigne = LIGNE_PREMIERE_DONNEE + NB_LIGNES_PREPAREES - 1

    With ws.Range(COL_STATUT & LIGNE_PREMIERE_DONNEE & ":" & COL_CATEGORIE & derniereLigne)
        .Font.Size = 9.5
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(230, 230, 226)   ' liseré très discret entre les lignes
    End With

    ws.Range(COL_STATUT & LIGNE_PREMIERE_DONNEE & ":" & COL_STATUT & derniereLigne).HorizontalAlignment = xlCenter
    ws.Range(COL_MONTANT & LIGNE_PREMIERE_DONNEE & ":" & COL_MONTANT & derniereLigne).HorizontalAlignment = xlRight
    ws.Range(COL_MONTANT & LIGNE_PREMIERE_DONNEE & ":" & COL_MONTANT & derniereLigne).NumberFormat = "#,##0.00 ?"
    ws.Range(COL_DATE & LIGNE_PREMIERE_DONNEE & ":" & COL_DATE & derniereLigne).NumberFormat = "dd/mm/yyyy"

    For ligneCourante = LIGNE_PREMIERE_DONNEE To derniereLigne Step 2
        ws.Range(COL_STATUT & ligneCourante & ":" & COL_CATEGORIE & ligneCourante).Interior.Color = RGB(250, 250, 248)
    Next ligneCourante

End Sub


' =====================================================================================
' SOUS-PROCÉDURE - Figer les volets sous la ligne d'en-têtes du tableau
' =====================================================================================
Private Sub FigerVoletsSousEntetes(ws As Worksheet)

    ws.Activate

    ' On sélectionne la première cellule de données : Excel fige tout ce qui se trouve
    ' AU-DESSUS et À GAUCHE de la cellule active au moment de l'appel à FreezePanes.
    ' On veut figer les lignes 1 à LIGNE_ENTETES_TABLEAU (boutons, instructions et
    ' en-têtes du tableau); on sélectionne donc la ligne juste en dessous.
    ws.Range(COL_STATUT & LIGNE_PREMIERE_DONNEE).Select

    ' On s'assure qu'aucun gel antérieur ne reste actif avant d'en appliquer un nouveau.
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True

    ' On revient en haut de la feuille pour un affichage propre à la prochaine ouverture.
    ws.Range("A1").Select

End Sub


' =====================================================================================
' FONCTION UTILITAIRE - Récupérer une feuille par son nom sans provoquer d'erreur
' si elle n'existe pas. Au lieu de disperser "On Error Resume Next" dans le code,
' on centralise cette mécanique ici, ce qui est plus lisible et plus sûr.
' =====================================================================================
Private Function ObtenirFeuilleSansErreur(nomFeuille As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nomFeuille)
    On Error GoTo 0
    Set ObtenirFeuilleSansErreur = ws
End Function


' =====================================================================================
' SOUS-PROCÉDURE UTILITAIRE - Supprimer les anciens boutons avant une reconstruction
' (évite les doublons si l'installation est relancée).
' =====================================================================================
Private Sub SupprimerFormesExistantes(ws As Worksheet)
    Dim uneForme As Shape
    ' On parcourt la collection à l'envers : supprimer un élément pendant un parcours
    ' vers l'avant peut en faire sauter d'autres. C'est une précaution classique.
    Dim i As Long
    For i = ws.Shapes.count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
End Sub


' =====================================================================================
' OUTIL DÉVELOPPEUR (réservé à toi, inaccessible depuis l'interface utilisateur
' normale) : bascule la visibilité de la feuille pour permettre de la retoucher.
' À exécuter depuis la fenêtre Exécution immédiate (Ctrl+G) en tapant :
'     AfficherFeuilleResolutionPourEdition
' =====================================================================================
Sub AfficherFeuilleResolutionPourEdition()
    Dim ws As Worksheet
    Set ws = ObtenirFeuilleSansErreur(NOM_FEUILLE_RESOLUTION)

    If ws Is Nothing Then
        MsgBox "La feuille '" & NOM_FEUILLE_RESOLUTION & "' n'existe pas encore." & vbCrLf & _
               "Execute d'abord la macro CreerFeuilleResolutionCategories.", vbExclamation
        Exit Sub
    End If

    ws.Visible = xlSheetVisible
    ws.Activate
    MsgBox "La feuille est maintenant visible et modifiable." & vbCrLf & _
           "Pense a la remasquer avec la macro MasquerFeuilleResolutionApresEdition" & vbCrLf & _
           "une fois tes retouches terminees.", vbInformation
End Sub

' Pendant symétrique de la macro ci-dessus : masque de nouveau la feuille après retouche.
Sub MasquerFeuilleResolutionApresEdition()
    Dim ws As Worksheet
    Set ws = ObtenirFeuilleSansErreur(NOM_FEUILLE_RESOLUTION)

    If ws Is Nothing Then
        MsgBox "La feuille '" & NOM_FEUILLE_RESOLUTION & "' n'existe pas.", vbExclamation
        Exit Sub
    End If

    ws.Visible = xlSheetVeryHidden
    MsgBox "La feuille est de nouveau masquee (xlSheetVeryHidden).", vbInformation
End Sub
