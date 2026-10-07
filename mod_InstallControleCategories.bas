Option Explicit

' =====================================================================================
' MODULE : mod_InstallControleCategories
'
' PHASE 2 du chantier « Catégorie / Sous-catégorie / Ventilation » - PARTIE 1/2
'
' RÔLE (à lire en premier, même si vous débutez) :
'   Ce module CONSTRUIT la feuille qui sert de "formulaire" pour contrôler les
'   catégories des opérations importées : mise en page, libellés et boutons.
'   Il ne contient AUCUNE logique de fonctionnement : celle-ci se trouve dans
'   le module mod_ControleCategories (partie 2/2).
'
'   Pourquoi une FEUILLE et pas un UserForm ? Parce que vous avez constaté que les
'   UserForms s'affichent mal sur les postes multi-écrans (problème de DPI). On
'   reprend donc la méthode déjà utilisée pour résoudre les catégories ambiguës :
'   une feuille masquée (xlSheetVeryHidden) que le programme affiche au bon moment.
'
'   Analogie : ce module "construit les murs" de la pièce; le module
'   mod_ControleCategories "branche l'électricité".
'
' INSTALLATION (une seule fois) :
'   1. Alt+F11, puis Fichier > Importer un fichier... : importer CE fichier.
'   2. Importer aussi mod_ControleCategories.bas.
'   3. Ctrl+G (fenêtre Exécution), taper : CreerFeuilleControleCategories, puis Entrée.
'   4. Le message final indique le NOM de la feuille créée. Collez dans son
'      module de code les deux procédures du fichier
'      "CodeBehind_frm_ControleCategories.txt" (explications dans ce fichier).
'
' Ce module ne modifie AUCUNE donnée : il ajoute seulement une feuille masquée.
' =====================================================================================

' --- Nom de la feuille-formulaire -------------------------------------------------------
Public Const CTRL_NOM_FEUILLE As String = "frm_ControleCategories"

' --- Position des éléments (regroupés ici pour n'avoir qu'un endroit à modifier) -------
Public Const CTRL_LIGNE_BOUTONS As Long = 2       ' ligne des boutons

Public Const CTRL_ADR_TITRE As String = "B4"      ' titre
Public Const CTRL_ADR_COMPTEUR As String = "B5"   ' "Opération 3 / 27"

' Informations de l'opération (lecture seule) : étiquette en colonne B, valeur en colonne C
Public Const CTRL_ADR_DATE As String = "C7"
Public Const CTRL_ADR_TIERS As String = "C8"
Public Const CTRL_ADR_LIBELLE As String = "C9"
Public Const CTRL_ADR_MONTANT As String = "C10"
Public Const CTRL_ADR_SOURCE As String = "C11"    ' catégorie envoyée par la banque

' Zones de saisie de l'opérateur
Public Const CTRL_ADR_CAT As String = "C13"       ' catégorie choisie
Public Const CTRL_ADR_SOUS As String = "C14"      ' sous-catégorie choisie

Public Const CTRL_ADR_MESSAGE As String = "B16"   ' message d'aide (zone fusionnee B16:E16)

' Zone technique masquée : liste des sous-catégories de la catégorie choisie.
' La colonne 26 correspond à la colonne Z; elle est masquée, loin à droite de la mise en page.
Public Const CTRL_COL_AIDE As Long = 26


' =====================================================================================
' MACRO PRINCIPALE D'INSTALLATION
' =====================================================================================
Public Sub CreerFeuilleControleCategories()

    Dim ws As Worksheet
    Dim wsPrecedente As Worksheet
    Dim reponse As VbMsgBoxResult

    ' On retient la feuille affichée pour y revenir à la fin (l'opérateur ne doit pas
    ' se retrouver ailleurs après l'installation).
    Set wsPrecedente = ActiveSheet

    Set ws = FeuilleSansErreur(CTRL_NOM_FEUILLE)

    If Not ws Is Nothing Then
        ' La feuille existe déjà : on demande confirmation avant de tout reconstruire.
        reponse = MsgBox(mod_Display.FR("La feuille '") & CTRL_NOM_FEUILLE & mod_Display.FR("' existe d{e2}j{a2}.") & vbCrLf & _
                         mod_Display.FR("Voulez-vous la reconstruire enti{e1}rement (sa mise en forme sera perdue) ?"), _
                         vbYesNo + vbQuestion, mod_Display.FR("Confirmation de reconstruction"))
        If reponse = vbNo Then
            MsgBox mod_Display.FR("Installation annul{e2}e, rien n'a {e2}t{e2} modifi{e2}."), vbInformation
            Exit Sub
        End If

        ' Impossible de modifier une feuille "très masquée" : on la rend d'abord visible.
        ws.Visible = xlSheetVisible
        ws.Cells.UnMerge           ' défusionne les cellules fusionnées
        ws.Cells.Clear             ' efface le contenu et la mise en forme
        SupprimerFormes ws         ' supprime les anciens boutons
    Else
        ' Nouvelle feuille, placée en dernière position pour ne pas perturber les onglets.
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = CTRL_NOM_FEUILLE
    End If

    Application.ScreenUpdating = False

    MettreEnFormeGenerale ws
    ConstruireBoutons ws
    ConstruireTitreEtInformations ws
    ConstruireZoneSaisie ws
    ConstruireInstructions ws

    ' La colonne technique (liste des sous-catégories) doit rester invisible.
    ws.Columns(CTRL_COL_AIDE).Hidden = True

    ' Masquage complet : invisible pour l'opérateur, même via clic droit > Afficher.
    ws.Visible = xlSheetVeryHidden

    On Error Resume Next
    wsPrecedente.Activate
    On Error GoTo 0
    Application.ScreenUpdating = True

    MsgBox mod_Display.FR("La feuille '") & CTRL_NOM_FEUILLE & mod_Display.FR("' a {e2}t{e2} cr{e2}{e2}e puis masqu{e2}e.") & vbCrLf & vbCrLf & _
           mod_Display.FR("PROCHAINE {E2}TAPE (indispensable) : coller le code des {e2}v{e2}nements dans la feuille.") & vbCrLf & _
           "Nom interne (CodeName) de cette feuille : " & ws.CodeName & vbCrLf & vbCrLf & _
           mod_Display.FR("Voir le fichier CodeBehind_frm_ControleCategories.txt."), _
           vbInformation, mod_Display.FR("Installation termin{e2}e")

End Sub


' =====================================================================================
' Mise en forme générale : aspect "formulaire" (sans quadrillage), largeurs et police.
' =====================================================================================
Private Sub MettreEnFormeGenerale(ByVal ws As Worksheet)

    ' Le quadrillage est un réglage de la FENÊTRE, pas de la feuille : il faut donc
    ' activer la feuille pour pouvoir le supprimer.
    ws.Activate
    ActiveWindow.DisplayGridlines = False

    ws.Columns("A").ColumnWidth = 2      ' petite marge à gauche
    ws.Columns("B").ColumnWidth = 26     ' étiquettes
    ws.Columns("C").ColumnWidth = 44     ' valeurs et zones de saisie
    ws.Columns("D").ColumnWidth = 22
    ws.Columns("E").ColumnWidth = 22
    ws.Columns("F").ColumnWidth = 2      ' petite marge à droite

    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10

    ws.Range("A1").Select
    ActiveWindow.DisplayHeadings = False

End Sub


' =====================================================================================
' Boutons (contrôles de formulaire, pas ActiveX : plus fiables avec le DPI).
' =====================================================================================
' Chaque bouton est placé d'après une CELLULE (et non selon des pixels fixes) : si la largeur
' des colonnes change, les boutons suivent. Le nom de macro donné à .OnAction correspond
' à celui de la Public Sub concernée dans mod_ControleCategories.
Private Sub ConstruireBoutons(ByVal ws As Worksheet)

    Dim zone As Range

    ws.rows(CTRL_LIGNE_BOUTONS).RowHeight = 26

    ' --- Bouton "Précédent" (cellule B2) ---
    Set zone = ws.Cells(CTRL_LIGNE_BOUTONS, 2)
    AjouterBouton ws, zone.Left, zone.Top, zone.Width, zone.Height, _
                  "< " & mod_Display.FR("Pr{e2}c{e2}dent"), "ControleOperationPrecedente", "btnCtrlPrecedent"

    ' --- Bouton "Suivant" (moitié gauche de la cellule C2) ---
    Set zone = ws.Cells(CTRL_LIGNE_BOUTONS, 3)
    AjouterBouton ws, zone.Left, zone.Top, 130, zone.Height, _
                  "Suivant >", "ControleOperationSuivante", "btnCtrlSuivant"

    ' --- Bouton "Terminer et continuer" (cellule D2) ---
    Set zone = ws.Cells(CTRL_LIGNE_BOUTONS, 4)
    AjouterBouton ws, zone.Left, zone.Top, zone.Width, zone.Height, _
                  "Terminer et continuer", "ControleTerminer", "btnCtrlTerminer"

    ' --- Bouton "Annuler" (cellule E2) ---
    Set zone = ws.Cells(CTRL_LIGNE_BOUTONS, 5)
    AjouterBouton ws, zone.Left, zone.Top, zone.Width, zone.Height, _
                  "Annuler l'import", "ControleAnnuler", "btnCtrlAnnuler"

End Sub

' Crée UN bouton. Regroupe les quatre lignes répétitives pour garder le code lisible.
Private Sub AjouterBouton(ByVal ws As Worksheet, ByVal gauche As Double, ByVal haut As Double, _
                          ByVal largeur As Double, ByVal hauteur As Double, _
                          ByVal legende As String, ByVal nomMacro As String, ByVal nomBouton As String)
    Dim btn As Button
    Set btn = ws.Buttons.Add(gauche, haut, largeur, hauteur)
    btn.Caption = legende
    btn.OnAction = nomMacro    ' macro exécutée au clic
    btn.Name = nomBouton
End Sub


' =====================================================================================
' Titre, compteur et bloc "informations de l'opération" (lecture seule)
' =====================================================================================
Private Sub ConstruireTitreEtInformations(ByVal ws As Worksheet)

    ' --- Titre ---
    With ws.Range("B4:E4")
        .Merge
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(60, 60, 60)
    End With
    ws.Range(CTRL_ADR_TITRE).value = mod_Display.FR("Contr{o2}le des cat{e2}gories des op{e2}rations import{e2}es")

    ' --- Compteur "Operation X / N" (rempli par le programme) ---
    With ws.Range("B5:E5")
        .Merge
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(31, 78, 121)
    End With

    ' --- Étiquettes (colonne B) ---
    EcrireEtiquette ws, "B7", "Date"
    EcrireEtiquette ws, "B8", "Tiers"
    EcrireEtiquette ws, "B9", mod_Display.FR("Libell{e2} / Notes")
    EcrireEtiquette ws, "B10", "Montant"
    EcrireEtiquette ws, "B11", mod_Display.FR("Cat{e2}gorie source (banque)")

    ' --- Valeurs (colonne C) : format TEXTE pour éviter toute réinterprétation ---------
    ' NOTE (ajout du 03/10/2026) : Date (C7), Montant (C10) et Categorie source (C11)
    ' restent TOUJOURS en lecture seule. Tiers (C8) et Notes (C9) font exception : depuis
    ' le 03/10/2026, l'opérateur peut les modifier (sauf pour une opération de santé).
    ' Leur couleur de fond n'est toutefois PAS définie ici à l'installation; elle est
    ' recalculée à chaque affichage par mod_ControleCategories.AfficherOperation
    ' (gris = verrouillé, sans couleur = modifiable), car elle dépend de la sous-catégorie
    ' de l'opération en cours, qui change à chaque affichage. Il est donc normal que ces
    ' cellules n'aient pas de couleur particulière à l'installation.
    With ws.Range("C7:C11")
        .NumberFormat = "@"
        .WrapText = True
        .VerticalAlignment = xlCenter
        .Font.Size = 10
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(225, 225, 220)
        .Borders(xlInsideHorizontal).Color = RGB(225, 225, 220)
    End With
    ws.rows(9).RowHeight = 42           ' le libellé peut être long : on lui laisse trois lignes
    ws.Range(CTRL_ADR_MONTANT).Font.Bold = True
    ws.Range(CTRL_ADR_MONTANT).HorizontalAlignment = xlLeft

    ' --- Bouton "Ventiler" (phase 4) ---
    ' Placé à côté du champ Montant (D10:E10), il ouvre la feuille frm_Ventilation (voir
    ' mod_InstallVentilation); la logique se trouve dans mod_ControleCategories.ControleVentiler.
    ws.rows(10).RowHeight = 22
    Dim zoneBoutonVentiler As Range
    Set zoneBoutonVentiler = ws.Range("D10:E10")
    AjouterBouton ws, zoneBoutonVentiler.Left, zoneBoutonVentiler.Top, zoneBoutonVentiler.Width, zoneBoutonVentiler.Height, _
                  mod_Display.FR("Ventiler..."), "ControleVentiler", "btnCtrlVentiler"

End Sub

' Écrit une étiquette grise (colonne B).
Private Sub EcrireEtiquette(ByVal ws As Worksheet, ByVal adresse As String, ByVal texte As String)
    With ws.Range(adresse)
        .value = texte
        .Font.Bold = True
        .Font.Color = RGB(110, 110, 110)
        .VerticalAlignment = xlCenter
    End With
End Sub


' =====================================================================================
' Zone de saisie : Catégorie et Sous-catégorie (cases jaune pâle, encadrées)
' =====================================================================================
' Les listes déroulantes (validation de données) ne sont PAS définies ici : elles
' dépendent du tableau de correspondance et sont donc créées par le programme à
' chaque ouverture du formulaire (voir mod_ControleCategories).
Private Sub ConstruireZoneSaisie(ByVal ws As Worksheet)

    EcrireEtiquette ws, "B13", mod_Display.FR("Cat{e2}gorie")
    EcrireEtiquette ws, "B14", mod_Display.FR("Sous-cat{e2}gorie")
    ws.Range("B13:B14").Font.Color = RGB(31, 78, 121)

    With ws.Range("C13:C14")
        .NumberFormat = "@"
        .Interior.Color = RGB(255, 250, 225)      ' jaune pâle = "à vous de saisir"
        .Font.Size = 11
        .Font.Bold = True
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 185, 120)
    End With
    ws.rows("13:14").RowHeight = 22

    ' --- Bouton "+" pour créer une nouvelle catégorie/sous-catégorie (phase 3) ---
    ' Placé à côté du champ Categorie (D13:E13), il ouvre la feuille frm_NouvelleCategorie
    ' (voir mod_InstallNouvelleCategorie); la logique se trouve dans mod_ControleCategories.
    Dim zoneBoutonNouvelle As Range
    Set zoneBoutonNouvelle = ws.Range("D13:E13")
    AjouterBouton ws, zoneBoutonNouvelle.Left, zoneBoutonNouvelle.Top, zoneBoutonNouvelle.Width, zoneBoutonNouvelle.Height, _
                  "+ " & mod_Display.FR("Nouvelle cat{e2}gorie"), "ControleNouvelleCategorie", "btnCtrlNouvelleCategorie"

    ' --- Message d'aide (varie selon l'opération affichée) ---
    With ws.Range("B16:E16")
        .Merge
        .WrapText = True
        .VerticalAlignment = xlTop
        .Font.Size = 9.5
        .Font.Italic = True
        .Font.Color = RGB(90, 90, 90)
    End With
    ws.rows(16).RowHeight = 30

End Sub


' =====================================================================================
' Bloc d'instructions (texte fixe)
' =====================================================================================
Private Sub ConstruireInstructions(ByVal ws As Worksheet)

    ws.Range("B18").value = "Instructions"
    ws.Range("B18").Font.Bold = True
    ws.Range("B18").Font.Size = 11

    With ws.Range("B19:E23")
        .Merge
        .value = mod_Display.FR("1. V{e2}rifiez les informations de l'op{e2}ration (en haut).") & Chr(10) & _
                 mod_Display.FR("2. Modifiez si besoin la Cat{e2}gorie, puis la Sous-cat{e2}gorie (listes d{e2}roulantes).") & Chr(10) & _
                 mod_Display.FR("3. Passez {a2} l'op{e2}ration suivante avec 'Suivant' (ou revenez avec 'Pr{e2}c{e2}dent').") & Chr(10) & _
                 mod_Display.FR("4. Quand tout est contr{o2}l{e2}, cliquez sur 'Terminer et continuer'.") & Chr(10) & _
                 mod_Display.FR("Vos choix ne sont pris en compte qu'apr{e1}s 'Terminer et continuer'.")
        .WrapText = True
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
        .Font.Size = 9
        .Font.Color = RGB(80, 80, 80)
        .Interior.Color = RGB(245, 245, 242)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(210, 210, 205)
    End With
    ws.rows("19:23").RowHeight = 16

End Sub


' =====================================================================================
' OUTILS DÉVELOPPEUR (Ctrl+G) : afficher ou masquer la feuille pour la retoucher.
' =====================================================================================
Public Sub AfficherFeuilleControlePourEdition()
    Dim ws As Worksheet
    Set ws = FeuilleSansErreur(CTRL_NOM_FEUILLE)
    If ws Is Nothing Then
        MsgBox mod_Display.FR("La feuille n'existe pas encore. Ex{e2}cutez CreerFeuilleControleCategories."), vbExclamation
        Exit Sub
    End If
    ws.Visible = xlSheetVisible
    ws.Activate
    MsgBox mod_Display.FR("Feuille visible. Remasquez-la ensuite avec : MasquerFeuilleControleApresEdition"), vbInformation
End Sub

Public Sub MasquerFeuilleControleApresEdition()
    Dim ws As Worksheet
    Set ws = FeuilleSansErreur(CTRL_NOM_FEUILLE)
    If ws Is Nothing Then Exit Sub
    ws.Visible = xlSheetVeryHidden
    MsgBox mod_Display.FR("Feuille de nouveau masqu{e2}e."), vbInformation
End Sub


' =====================================================================================
' OUTILS INTERNES
' =====================================================================================

' Retrouve une feuille par son nom SANS provoquer d'erreur si elle n'existe pas.
Private Function FeuilleSansErreur(ByVal nomFeuille As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nomFeuille)
    On Error GoTo 0
    Set FeuilleSansErreur = ws
End Function

' Supprime tous les boutons (évite les doublons si l'installation est relancée).
' On parcourt à l'envers : supprimer en avançant ferait sauter des éléments.
Private Sub SupprimerFormes(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
End Sub