Option Explicit

' =====================================================================================
' MODULE : mod_InstallNouvelleCategorie
'
' PHASE 3 du chantier "Categorie / Sous-categorie / Ventilation" - PARTIE 1/2
'
' ROLE (a lire en premier, meme si vous debutez) :
'   Ce module CONSTRUIT la feuille-formulaire qui permet de creer une nouvelle
'   categorie (ou une nouvelle sous-categorie d'une categorie existante).
'   Elle s'ouvre par-dessus le formulaire de controle des categories (Phase 2), quand
'   l'operateur clique sur le bouton "+ Nouvelle categorie".
'   Ce module ne contient AUCUNE logique : elle se trouve dans mod_NouvelleCategorie
'   (partie 2/2), exactement comme pour la Phase 2.
'
'   Deux "listes modifiables" (menus deroulants qui acceptent aussi une saisie libre,
'   avec un simple avertissement) :
'     - Categorie      : liste des categories existantes, ou saisie d'une nouvelle.
'     - Sous-categorie : liste des sous-categories DEJA CONNUES de la categorie
'                        choisie, ou saisie d'une nouvelle.
'
' INSTALLATION (une seule fois) :
'   1. Alt+F11, Fichier > Importer un fichier... : importer CE fichier.
'   2. Importer aussi mod_NouvelleCategorie.bas.
'   3. Ctrl+G, taper :  CreerFeuilleNouvelleCategorie  puis Entree.
'   4. Coller dans le module de code de la feuille creee les 2 procedures du fichier
'      "CodeBehind_frm_NouvelleCategorie.txt" (memes explications qu'en Phase 2).
'   5. Reimporter mod_InstallControleCategories.bas et mod_ControleCategories.bas
'      (mis a jour pour cette Phase 3 : bouton "+" ajoute, voir la note fournie).
'
' Ce module ne modifie AUCUNE donnee : il ajoute seulement une feuille cachee.
' =====================================================================================

Public Const NC_NOM_FEUILLE As String = "frm_NouvelleCategorie"

Public Const NC_LIGNE_BOUTONS As Long = 2

Public Const NC_ADR_TITRE As String = "B4"
Public Const NC_ADR_CAT As String = "C6"
Public Const NC_ADR_SOUS As String = "C9"
Public Const NC_ADR_MESSAGE As String = "B11"

' Zone technique cachee : sous-categories de la categorie choisie (colonne Z).
Public Const NC_COL_AIDE As Long = 26
Public Const NC_LIGNE_AIDE_MAX As Long = 300


' =====================================================================================
' MACRO PRINCIPALE D'INSTALLATION
' =====================================================================================
Public Sub CreerFeuilleNouvelleCategorie()

    Dim ws As Worksheet
    Dim wsPrecedente As Worksheet
    Dim reponse As VbMsgBoxResult

    Set wsPrecedente = ActiveSheet

    Set ws = FeuilleSansErreur(NC_NOM_FEUILLE)

    If Not ws Is Nothing Then
        reponse = MsgBox(TF("La feuille '") & NC_NOM_FEUILLE & TF("' existe d{e2}j{a2}.") & vbCrLf & _
                         TF("Voulez-vous la reconstruire enti{e1}rement (sa mise en forme sera perdue) ?"), _
                         vbYesNo + vbQuestion, TF("Confirmation de reconstruction"))
        If reponse = vbNo Then
            MsgBox TF("Installation annul{e2}e, rien n'a {e2}t{e2} modifi{e2}."), vbInformation
            Exit Sub
        End If

        ws.Visible = xlSheetVisible
        ws.Cells.UnMerge
        ws.Cells.Clear
        SupprimerFormes ws
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = NC_NOM_FEUILLE
    End If

    Application.ScreenUpdating = False

    MettreEnFormeGenerale ws
    ConstruireBoutons ws
    ConstruireChamps ws

    ws.Columns(NC_COL_AIDE).Hidden = True

    ws.Visible = xlSheetVeryHidden

    On Error Resume Next
    wsPrecedente.Activate
    On Error GoTo 0
    Application.ScreenUpdating = True

    MsgBox TF("La feuille '") & NC_NOM_FEUILLE & TF("' a {e2}t{e2} cr{e2}{e2}e puis masqu{e2}e.") & vbCrLf & vbCrLf & _
           TF("PROCHAINE {E2}TAPE (indispensable) : coller le code des {e2}v{e2}nements dans la feuille.") & vbCrLf & _
           "Nom interne (CodeName) de cette feuille : " & ws.CodeName & vbCrLf & vbCrLf & _
           TF("Voir le fichier CodeBehind_frm_NouvelleCategorie.txt."), _
           vbInformation, TF("Installation termin{e2}e")

End Sub


' =====================================================================================
' Mise en forme generale (memes principes qu'en Phase 2, pour un rendu coherent)
' =====================================================================================
Private Sub MettreEnFormeGenerale(ByVal ws As Worksheet)

    ws.Activate
    ActiveWindow.DisplayGridlines = False

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 26
    ws.Columns("C").ColumnWidth = 44
    ws.Columns("D").ColumnWidth = 22
    ws.Columns("E").ColumnWidth = 22
    ws.Columns("F").ColumnWidth = 2

    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10

    ws.Range("A1").Select
    ActiveWindow.DisplayHeadings = False

End Sub


' =====================================================================================
' Boutons "Valider" et "Annuler"
' =====================================================================================
Private Sub ConstruireBoutons(ByVal ws As Worksheet)

    Dim zone As Range

    ws.rows(NC_LIGNE_BOUTONS).RowHeight = 26

    Set zone = ws.Cells(NC_LIGNE_BOUTONS, 2)
    AjouterBouton ws, zone.Left, zone.Top, zone.Width, zone.Height, "Valider", "NcValider", "btnNcValider"

    Set zone = ws.Cells(NC_LIGNE_BOUTONS, 3)
    AjouterBouton ws, zone.Left, zone.Top, zone.Width, zone.Height, "Annuler", "NcAnnuler", "btnNcAnnuler"

End Sub

Private Sub AjouterBouton(ByVal ws As Worksheet, ByVal gauche As Double, ByVal haut As Double, _
                          ByVal largeur As Double, ByVal hauteur As Double, _
                          ByVal legende As String, ByVal nomMacro As String, ByVal nomBouton As String)
    Dim btn As Button
    Set btn = ws.Buttons.Add(gauche, haut, largeur, hauteur)
    btn.Caption = legende
    btn.OnAction = nomMacro
    btn.Name = nomBouton
End Sub


' =====================================================================================
' Titre, champs de saisie et message d'aide
' =====================================================================================
Private Sub ConstruireChamps(ByVal ws As Worksheet)

    ' --- Titre ---
    With ws.Range("B4:E4")
        .Merge
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(60, 60, 60)
    End With
    ws.Range(NC_ADR_TITRE).value = TF("Nouvelle cat{e2}gorie / sous-cat{e2}gorie")

    ' --- Categorie ---
    With ws.Range("B6")
        .value = TF("Cat{e2}gorie")
        .Font.Bold = True
        .Font.Color = RGB(31, 78, 121)
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("C6")
        .NumberFormat = "@"
        .Interior.Color = RGB(255, 250, 225)
        .Font.Size = 11
        .Font.Bold = True
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 185, 120)
    End With
    ws.rows(6).RowHeight = 22

    With ws.Range("B7:E7")
        .Merge
        .value = TF("Choisissez une cat{e2}gorie existante dans la liste, ou tapez-en une nouvelle.")
        .Font.Size = 8.5
        .Font.Italic = True
        .Font.Color = RGB(120, 120, 120)
    End With

    ' --- Sous-categorie ---
    With ws.Range("B9")
        .value = TF("Sous-cat{e2}gorie")
        .Font.Bold = True
        .Font.Color = RGB(31, 78, 121)
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("C9")
        .NumberFormat = "@"
        .Interior.Color = RGB(255, 250, 225)
        .Font.Size = 11
        .Font.Bold = True
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 185, 120)
    End With
    ws.rows(9).RowHeight = 22

    With ws.Range("B10:E10")
        .Merge
        .value = TF("Sous-cat{e2}gories d{e2}j{a2} connues de la cat{e2}gorie ci-dessus, ou tapez-en une nouvelle.") & vbCrLf & _
                 TF("Vous pouvez aussi laisser ce champ vide.")
        .Font.Size = 8.5
        .Font.Italic = True
        .Font.Color = RGB(120, 120, 120)
        .WrapText = True
    End With
    ws.rows(10).RowHeight = 26

    ' --- Message (erreurs de validation) ---
    With ws.Range("B11:E11")
        .Merge
        .WrapText = True
        .VerticalAlignment = xlTop
        .Font.Size = 9
        .Font.Color = RGB(192, 80, 0)
    End With
    ws.rows(11).RowHeight = 26

End Sub


' =====================================================================================
' OUTILS DEVELOPPEUR (Ctrl+G)
' =====================================================================================
Public Sub AfficherFeuilleNouvelleCategoriePourEdition()
    Dim ws As Worksheet
    Set ws = FeuilleSansErreur(NC_NOM_FEUILLE)
    If ws Is Nothing Then
        MsgBox TF("La feuille n'existe pas encore. Ex{e2}cutez CreerFeuilleNouvelleCategorie."), vbExclamation
        Exit Sub
    End If
    ws.Visible = xlSheetVisible
    ws.Activate
    MsgBox TF("Feuille visible. Remasquez-la avec : MasquerFeuilleNouvelleCategorieApresEdition"), vbInformation
End Sub

Public Sub MasquerFeuilleNouvelleCategorieApresEdition()
    Dim ws As Worksheet
    Set ws = FeuilleSansErreur(NC_NOM_FEUILLE)
    If ws Is Nothing Then Exit Sub
    ws.Visible = xlSheetVeryHidden
    MsgBox TF("Feuille de nouveau masqu{e2}e."), vbInformation
End Sub


' =====================================================================================
' OUTILS INTERNES
' =====================================================================================

Private Function FeuilleSansErreur(ByVal nomFeuille As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nomFeuille)
    On Error GoTo 0
    Set FeuilleSansErreur = ws
End Function

Private Sub SupprimerFormes(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
End Sub

' Traduit des codes en lettres accentuees (le fichier reste 100% ASCII).
Private Function TF(ByVal texte As String) As String
    Dim r As String
    r = texte
    r = Replace(r, "{e2}", ChrW(233))
    r = Replace(r, "{e1}", ChrW(232))
    r = Replace(r, "{a2}", ChrW(224))
    r = Replace(r, "{E2}", ChrW(201))
    TF = r
End Function
