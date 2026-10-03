Option Explicit

' =====================================================================================
' MODULE : mod_InstallResolutionSheet
'
' ROLE (PHASE 1 du chantier "UserForm -> feuille dediee") :
'   Ce module ne contient QUE la construction de la mise en page statique de la
'   feuille qui va remplacer le UserForm "frmResolutionCategories".
'   Il ne contient ENCORE AUCUNE logique de remplissage des cas ambigus, ni de
'   gestion des clics : cela viendra dans la Phase 2 (module mod_ResolutionCategories).
'
'   Pourquoi separer ainsi ? Pour que tu puisses executer cette Phase 1 seule,
'   verifier visuellement que la feuille correspond a la maquette validee,
'   AVANT qu'on y ajoute le comportement (Phase 2). C'est le meme principe que
'   "construire les murs avant de brancher l'electricite".
'
' A FAIRE POUR INSTALLER CE MODULE :
'   1. Alt+F11 pour ouvrir l'editeur VBA
'   2. Fichier > Importer un fichier... > selectionner ce fichier .bas
'   3. Dans la fenetre Execution immediate (Ctrl+G), taper :
'        CreerFeuilleResolutionCategories
'      puis appuyer sur Entree. La feuille est creee, mise en forme, puis masquee.
' =====================================================================================

' -------------------------------------------------------------------------------------
' CONSTANTES DE MISE EN PAGE
' -------------------------------------------------------------------------------------
' Regrouper ici toutes les positions de cellules evite d'avoir des "nombres magiques"
' disperses dans le code. Si un jour tu veux deplacer une zone, tu changes UNE seule
' ligne ici plutot que de chercher partout dans le code.
' Ces constantes seront reutilisees telles quelles dans le module de la Phase 2.

Public Const NOM_FEUILLE_RESOLUTION As String = "frm_ResolutionCategories"

' --- Zone des boutons (ligne 2) ---
Public Const LIGNE_BOUTONS As Long = 2

' --- Zone des instructions ---
Public Const LIGNE_TITRE_INSTRUCTIONS As Long = 5
Public Const LIGNE_DEBUT_TEXTE_INSTRUCTIONS As Long = 6
Public Const LIGNE_FIN_TEXTE_INSTRUCTIONS As Long = 9




' =====================================================================================
' MACRO PRINCIPALE D'INSTALLATION
' A executer UNE SEULE FOIS (ou a nouveau si tu veux reinitialiser completement
' la mise en forme de la feuille).
' =====================================================================================
Sub CreerFeuilleResolutionCategories()

    Dim ws As Worksheet
    Dim reponseUtilisateur As VbMsgBoxResult

    ' ---------------------------------------------------------------------------------
    ' ETAPE A - Verifier si la feuille existe deja, pour eviter d'ecraser du travail
    ' sans prevenir. On utilise une fonction utilitaire (definie plus bas) qui essaie
    ' de recuperer la feuille par son nom sans provoquer d'erreur si elle n'existe pas.
    ' ---------------------------------------------------------------------------------
    Set ws = ObtenirFeuilleSansErreur(NOM_FEUILLE_RESOLUTION)

    If Not ws Is Nothing Then
        ' La feuille existe deja : on demande confirmation avant de tout reconstruire,
        ' car cela va effacer sa mise en forme actuelle.
        reponseUtilisateur = MsgBox( _
            "La feuille '" & NOM_FEUILLE_RESOLUTION & "' existe deja." & vbCrLf & _
            "Veux-tu la reconstruire entierement (sa mise en forme actuelle sera perdue) ?", _
            vbYesNo + vbQuestion, "Confirmation de reconstruction")

        If reponseUtilisateur = vbNo Then
            MsgBox "Installation annulee, aucune modification effectuee.", vbInformation
            Exit Sub
        End If

        ' On la rend visible temporairement : impossible de la modifier/supprimer
        ' proprement tant qu'elle est en xlSheetVeryHidden.
        ws.Visible = xlSheetVisible
        ws.Cells.Clear             ' on vide tout le contenu et toute la mise en forme
        Call SupprimerFormesExistantes(ws)   ' on retire les anciens boutons s'il y en a
    Else
        ' La feuille n'existe pas encore : on la cree, positionnee en derniere position
        ' pour ne pas perturber l'ordre des onglets existants (Accueil, Synthese...).
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = NOM_FEUILLE_RESOLUTION
    End If

    ' ---------------------------------------------------------------------------------
    ' ETAPE B - Mise en forme generale de la feuille (aspect "formulaire")
    ' ---------------------------------------------------------------------------------
    Call AppliquerMiseEnFormeGenerale(ws)

    ' ---------------------------------------------------------------------------------
    ' ETAPE C - Construction de la zone des boutons
    ' ---------------------------------------------------------------------------------
    Call ConstruireZoneBoutons(ws)

    ' ---------------------------------------------------------------------------------
    ' ETAPE D - Construction de la zone des instructions
    ' ---------------------------------------------------------------------------------
    Call ConstruireZoneInstructions(ws)

    ' ---------------------------------------------------------------------------------
    ' ETAPE E - Construction des en-tetes et de la mise en forme du tableau des cas
    ' ---------------------------------------------------------------------------------
    Call ConstruireTableauCas(ws)

    ' ---------------------------------------------------------------------------------
    ' ETAPE F - Figer les volets pour que boutons + instructions + en-tetes restent
    ' visibles pendant que l'utilisateur fait defiler la liste des cas.
    ' ---------------------------------------------------------------------------------
    Call FigerVoletsSousEntetes(ws)

    ' ---------------------------------------------------------------------------------
    ' ETAPE G - Masquer completement la feuille (invisible pour l'utilisateur final,
    ' elle ne sera reveillee qu'au moment ou un cas ambigu doit etre resolu -
    ' ce declenchement sera code en Phase 2).
    ' ---------------------------------------------------------------------------------
    ws.Visible = xlSheetVeryHidden

    MsgBox "La feuille '" & NOM_FEUILLE_RESOLUTION & "' a ete creee et mise en forme," & vbCrLf & _
           "puis masquee (xlSheetVeryHidden)." & vbCrLf & vbCrLf & _
           "Pour la revoir a l'ecran et verifier le rendu, execute la macro :" & vbCrLf & _
           "   AfficherFeuilleResolutionPourEdition", _
           vbInformation, "Installation terminee"

End Sub


' =====================================================================================
' SOUS-PROCEDURE - Mise en forme generale (suppression du quadrillage, police, etc.)
' =====================================================================================
Private Sub AppliquerMiseEnFormeGenerale(ws As Worksheet)

    ' On doit activer la feuille pour pouvoir modifier certains reglages de la FENETRE
    ' (comme l'affichage du quadrillage), car ces reglages sont attaches a la fenetre
    ' au moment ou une feuille donnee est affichee, et non a la feuille elle-meme.
    ws.Activate

    ' Supprime le quadrillage (les lignes grises entre les cellules) pour donner
    ' un aspect "formulaire" plutot que "tableur classique", comme demande.
    ActiveWindow.DisplayGridlines = False

    ' On repart d'une largeur de colonnes coherente pour toute la zone utile.
    ws.Columns("A").ColumnWidth = 2      ' petite marge a gauche
    ws.Columns("B").ColumnWidth = 8      ' colonne "Statut"
    ws.Columns("C").ColumnWidth = 12     ' colonne "Date"
    ws.Columns("D").ColumnWidth = 12     ' colonne "Montant"
    ws.Columns("E").ColumnWidth = 28     ' colonne "Tiers"
    ws.Columns("F").ColumnWidth = 24     ' colonne "Categorie(s)"
    ws.Columns("G").ColumnWidth = 2      ' petite marge a droite

    ' Police par defaut de toute la feuille, coherente avec un rendu "formulaire" sobre.
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10

    ' On se replace en A1 et on desactive les barres de titres de lignes/colonnes
    ' (references L1C1 grisees) pour un rendu plus epure. Optionnel, mais renforce
    ' l'effet "formulaire" plutot que "feuille de calcul".
    ws.Range("A1").Select
    ActiveWindow.DisplayHeadings = False

End Sub


' =====================================================================================
' SOUS-PROCEDURE - Construction de la zone des boutons (ligne 2) + compteur
' =====================================================================================
Private Sub ConstruireZoneBoutons(ws As Worksheet)

    Dim boutonValider As Button
    Dim boutonTerminer As Button
    Dim zoneBoutonValider As Range
    Dim zoneBoutonTerminer As Range

    ' On definit la position de chaque bouton en s'appuyant sur une PLAGE DE CELLULES
    ' (et non des coordonnees en pixels fixes) : ainsi, si jamais la largeur des
    ' colonnes change, les boutons suivent automatiquement. C'est plus robuste qu'un
    ' positionnement en pixels pour un environnement qui doit rester stable.
    Set zoneBoutonValider = ws.Range(COL_STATUT & LIGNE_BOUTONS & ":" & COL_DATE & LIGNE_BOUTONS)
    Set zoneBoutonTerminer = ws.Range(COL_MONTANT & LIGNE_BOUTONS & ":" & COL_TIERS & LIGNE_BOUTONS)

    zoneBoutonValider.RowHeight = 22
    zoneBoutonTerminer.RowHeight = 22

    ' Ajout d'un bouton de type "Controle de formulaire" (PAS ActiveX), comme convenu :
    ' plus leger, plus fiable vis-a-vis du probleme de DPI/multi-ecrans initial.
    Set boutonValider = ws.Buttons.Add( _
        zoneBoutonValider.Left, zoneBoutonValider.Top, _
        zoneBoutonValider.Width, zoneBoutonValider.Height)
    With boutonValider
        ' IMPORTANT : le texte affiche est deja son intitule FINAL valide avec toi,
        ' a savoir sa fonction de reinitialisation (et non plus "Valider ce cas").
        .Caption = "R?initialiser cette ligne"
        ' Le nom de macro ci-dessous n'existe pas encore : il sera ecrit en Phase 2.
        ' Cela ne provoque AUCUNE erreur maintenant ; l'erreur n'apparaitrait que si
        ' quelqu'un cliquait sur le bouton avant que la Phase 2 soit installee.
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

    ' Cellule du compteur ("X restant(s) sur Y"), a droite de la barre de boutons.
    ' On lui donne un NOM DEFINI (plage nommee) plutot qu'une reference "F2" en dur :
    ' cela rend le code de la Phase 2 plus lisible (Range("CompteurCasRestants") parle
    ' de lui-meme, contrairement a Range("H2")).
    With ws.Range(COL_TIERS & LIGNE_BOUTONS & ":" & COL_CATEGORIE & LIGNE_BOUTONS)
        .Merge
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .Font.Size = 9
        .Font.Color = RGB(120, 120, 120)   ' gris discret, texte d'information secondaire
        .value = ""   ' rempli dynamiquement en Phase 2
    End With
    ws.Names.Add Name:="CompteurCasRestants", RefersTo:=ws.Range(COL_TIERS & LIGNE_BOUTONS)

End Sub


' =====================================================================================
' SOUS-PROCEDURE - Construction de la zone des instructions
' =====================================================================================
Private Sub ConstruireZoneInstructions(ws As Worksheet)

    ' --- Titre "Instructions" ---
    With ws.Range(COL_STATUT & LIGNE_TITRE_INSTRUCTIONS)
        .value = "Instructions"
        .Font.Bold = True
        .Font.Size = 11
    End With

    ' --- Bloc de texte des 4 etapes, sur une plage fusionnee pour ressembler a un
    ' encadre de note plutot qu'a des cellules de tableur classiques. ---
    Dim zoneTexte As Range
    Set zoneTexte = ws.Range( _
        COL_STATUT & LIGNE_DEBUT_TEXTE_INSTRUCTIONS & ":" & _
        COL_CATEGORIE & LIGNE_FIN_TEXTE_INSTRUCTIONS)

    With zoneTexte
        .Merge
        ' Le texte exact que tu m'as fourni, avec des sauts de ligne internes
        ' (Chr(10) est le caractere "retour a la ligne" a l'interieur d'une cellule).
        .value = "Les op?rations list?es n'ont pas pu ?tre cat?goris?es de fa?on automatique." & Chr(10) & _
                 "Il faut donc le faire manuellement. Pour ce faire suivre les ?tapes suivantes :" & Chr(10) & _
                 "1. S?lectionner une cat?gorie dans la liste d?roulante en face de l'op?ration concern?e" & Chr(10) & _
                 "2. Le statut de la ligne passe automatiquement ? "" OK "" une fois la cat?gorie choisie" & Chr(10) & _
                 "3. Recommencer pour chaque op?ration puis cliquer sur ""Terminer et appliquer"""
        .WrapText = True                     ' le texte revient a la ligne dans la cellule
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
        .Font.Size = 9
        .Font.Color = RGB(80, 80, 80)
        .Interior.Color = RGB(245, 245, 242) ' fond gris tres clair, type "encadre note"
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(210, 210, 205)
        .Locked = True                       ' cellule non modifiable par l'utilisateur final
    End With

    ' Remarque : le texte des instructions ci-dessus a ete legerement adapte par
    ' rapport a l'original du UserForm, car les etapes 1 a 3 de l'ancien texte
    ' decrivaient le fonctionnement de l'Option A (clic sur une ligne puis liste
    ' partagee puis bouton Valider). Comme on est passe a l'Option B (liste
    ' deroulante directement sur chaque ligne, statut automatique), le texte des
    ' etapes a ete adapte pour rester exact vis-a-vis du nouveau fonctionnement.
    ' Dis-moi si tu preferes une autre formulation, c'est une simple modification
    ' de texte, sans impact sur le reste du code.

    ' Ajuste la hauteur des lignes du bloc pour laisser de la place au texte
    ws.rows(LIGNE_DEBUT_TEXTE_INSTRUCTIONS & ":" & LIGNE_FIN_TEXTE_INSTRUCTIONS).RowHeight = 16

End Sub


' =====================================================================================
' SOUS-PROCEDURE - Construction des en-tetes et de la mise en forme du tableau
' =====================================================================================
Private Sub ConstruireTableauCas(ws As Worksheet)

    ' --- Ligne d'en-tetes (row LIGNE_ENTETES_TABLEAU) ---
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

    ' --- Mise en forme "a blanc" des lignes de donnees preparees a l'avance ---
    ' On ne remplit PAS encore de vraies donnees ici (ce sera fait dynamiquement en
    ' Phase 2, a chaque ouverture, selon le nombre reel de cas ambigus). On prepare
    ' seulement l'apparence (bordures legeres, alternance de fond, alignement),
    ' pour que la feuille ait deja fiere allure des cette Phase 1.
    Dim ligneCourante As Long
    Dim derniereLigne As Long
    derniereLigne = LIGNE_PREMIERE_DONNEE + NB_LIGNES_PREPAREES - 1

    With ws.Range(COL_STATUT & LIGNE_PREMIERE_DONNEE & ":" & COL_CATEGORIE & derniereLigne)
        .Font.Size = 9.5
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(230, 230, 226)   ' lisere tres discret entre lignes
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
' SOUS-PROCEDURE - Figer les volets sous la ligne d'en-tetes du tableau
' =====================================================================================
Private Sub FigerVoletsSousEntetes(ws As Worksheet)

    ws.Activate

    ' On selectionne la premiere cellule de donnees : Excel fige tout ce qui est
    ' AU-DESSUS et A GAUCHE de la cellule active au moment de l'appel a FreezePanes.
    ' Ici on veut figer les lignes 1 a LIGNE_ENTETES_TABLEAU (boutons, instructions,
    ' en-tetes du tableau), donc on selectionne la ligne juste en dessous.
    ws.Range(COL_STATUT & LIGNE_PREMIERE_DONNEE).Select

    ' On s'assure qu'aucun figement anterieur ne reste actif avant d'en reappliquer un.
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True

    ' On revient en haut de la feuille pour un affichage propre a la prochaine ouverture.
    ws.Range("A1").Select

End Sub


' =====================================================================================
' FONCTION UTILITAIRE - Recuperer une feuille par son nom sans provoquer d'erreur
' si elle n'existe pas (au lieu d'un "On Error Resume Next" disperse dans le code,
' on centralise cette petite mecanique ici, ce qui est plus lisible et plus sur).
' =====================================================================================
Private Function ObtenirFeuilleSansErreur(nomFeuille As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nomFeuille)
    On Error GoTo 0
    Set ObtenirFeuilleSansErreur = ws
End Function


' =====================================================================================
' SOUS-PROCEDURE UTILITAIRE - Supprimer les anciens boutons avant une reconstruction
' (evite d'avoir des boutons en double si on relance l'installation plusieurs fois)
' =====================================================================================
Private Sub SupprimerFormesExistantes(ws As Worksheet)
    Dim uneForme As Shape
    ' On parcourt a l'envers car supprimer un element d'une collection pendant qu'on
    ' la parcourt "vers l'avant" peut sauter des elements - une precaution classique.
    Dim i As Long
    For i = ws.Shapes.count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
End Sub


' =====================================================================================
' OUTIL DEVELOPPEUR (reserve a toi, jamais accessible depuis l'interface utilisateur
' normale) : bascule la visibilite de la feuille pour pouvoir la retoucher a la main.
' A executer depuis la fenetre Execution immediate (Ctrl+G) en tapant :
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

' Pendant symetrique de la macro ci-dessus : remasque la feuille apres retouche.
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
