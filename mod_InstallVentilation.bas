Option Explicit

' =====================================================================================
' MODULE : mod_InstallVentilation
'
' PHASE 4 du chantier "Categorie / Sous-categorie / Ventilation" - PARTIE 1/2
' (Version 2 : nouvelle ergonomie, a la demande de l'operateur apres tests)
'
' ROLE (a lire en premier, meme si vous debutez) :
'   Ce module met en place les DEUX briques necessaires a la ventilation d'une
'   operation entre plusieurs sous-categories :
'
'     1. La feuille de DONNEES "Ventilations", qui contient le tableau
'        "TblVentilations" : une ligne par sous-categorie ventilee, reliee a
'        l'operation d'origine par son ID_Transaction. Cette feuille reste
'        simplement CACHEE (pas tres cachee) : comme Param ou Import_data, vous
'        pouvez l'afficher si besoin pour verifier son contenu.
'
'     2. La feuille-FORMULAIRE "frm_Ventilation".
'
'   ERGONOMIE (changee suite a vos remarques de test) : la feuille est coupee en
'   deux zones bien distinctes :
'     - EN HAUT (lignes 16 a 26) : un TABLEAU D'AFFICHAGE des lignes deja ajoutees
'       a cette ventilation (categorie, sous-categorie, montant), avec un petit
'       bouton "Editer" en face de chacune.
'     - EN BAS (lignes 28 et suivantes) : un PETIT FORMULAIRE DE SAISIE, TOUJOURS LE
'       MEME, utilise pour AJOUTER une nouvelle ligne ou MODIFIER une ligne existante
'       (via son bouton Editer). Il reprend exactement le principe du formulaire de
'       controle des categories (Phase 2) : un champ Categorie avec un bouton "+"
'       pour en creer une nouvelle, et un champ Sous-categorie dont la liste se
'       filtre automatiquement selon la categorie choisie.
'
'   Ce module ne contient AUCUNE logique de calcul ou de validation : elle se trouve
'   dans mod_Ventilation (partie 2/2).
'
' INSTALLATION :
'   1. Alt+F11, Fichier > Importer un fichier... : importer CE fichier.
'   2. Importer aussi mod_Ventilation.bas.
'   3. Ctrl+G (fenetre Execution), taper :  PreparerPhase4Ventilation  puis Entree.
'      (cree TblVentilations, la met a jour si elle existe deja, ET reconstruit
'      frm_Ventilation)
'   4. Le message de fin vous donne le CodeName de frm_Ventilation. Collez-y le
'      contenu de CodeBehind_frm_Ventilation.txt (INCHANGE depuis la version
'      precedente : si vous l'aviez deja fait, rien a refaire).
'   5. IMPORTANT, deja rencontre : apres toute mise a jour de mod_InstallControleCategories,
'      il faut RELANCER CreerFeuilleControleCategories (Oui a la reconstruction) pour
'      que les boutons apparaissent reellement sur la feuille. Le simple import du
'      fichier .bas ne suffit pas : c'est l'execution de la macro d'installation qui
'      construit la feuille.
' =====================================================================================

Public Const VEN_NOM_FEUILLE_DONNEES As String = "Ventilations"
Public Const VEN_NOM_TABLE As String = "TblVentilations"

Public Const VEN_NOM_FEUILLE As String = "frm_Ventilation"

Public Const VEN_LIGNE_BOUTONS As Long = 2

Public Const VEN_ADR_TITRE As String = "B4"
Public Const VEN_ADR_DATE As String = "C6"
Public Const VEN_ADR_TIERS As String = "C7"
Public Const VEN_ADR_LIBELLE As String = "C8"
Public Const VEN_ADR_MONTANT As String = "C9"
Public Const VEN_ADR_CATACTUELLE As String = "C10"

Public Const VEN_ADR_MONTANT_A_VENTILER As String = "C12"
Public Const VEN_ADR_TOTAL_SAISI As String = "C13"
Public Const VEN_ADR_RESTE As String = "C14"

' --- Tableau D'AFFICHAGE des lignes deja ajoutees (lecture seule pour l'operateur :
' on ne modifie une ligne qu'en passant par le formulaire de saisie du bas, via le
' bouton "Editer") ---
Public Const VEN_LIGNE_GRILLE_ENTETE As Long = 16
Public Const VEN_LIGNE_GRILLE_DEBUT As Long = 17
Public Const VEN_NB_LIGNES As Long = 10
Public Const VEN_LIGNE_GRILLE_FIN As Long = 26     ' = VEN_LIGNE_GRILLE_DEBUT + VEN_NB_LIGNES - 1

Public Const VEN_COL_CAT As Long = 2      ' colonne B
Public Const VEN_COL_SOUS As Long = 3     ' colonne C
Public Const VEN_COL_MONTANT As Long = 4  ' colonne D
Public Const VEN_COL_NOTES As Long = 5    ' colonne E (ajout 01/10/2026, point 3)
Public Const VEN_COL_EDITER As Long = 6   ' colonne F : bouton "Editer" de chaque ligne (decalee de E a F)

' --- Formulaire de SAISIE (une seule ligne a la fois : ajout ou edition) ---
Public Const VEN_LIGNE_SAISIE_TITRE As Long = 28
Public Const VEN_ADR_SAISIE_CAT As String = "C29"
Public Const VEN_ADR_SAISIE_SOUS As String = "C30"
Public Const VEN_ADR_SAISIE_MONTANT As String = "C31"
' Ajout 01/10/2026 (point 3) : commentaire libre, disponible pour CHAQUE ligne.
' IMPORTANT : volontairement une cellule SIMPLE, NON FUSIONNEE (meme principe que
' Categorie/Sous-categorie/Montant juste au-dessus) -- une premiere version fusionnait
' C32:F32, ce qui provoquait une erreur 1004 "Cette action ne peut pas etre appliquee a
' une cellule fusionnee" des que le code essayait d'ecrire dans VEN_ADR_SAISIE_NOTES
' (qui ne pointe pas forcement vers le coin haut-gauche d'une fusion). Une cellule simple
' elimine ce risque, au prix d'une largeur de saisie un peu plus etroite (compensee par
' un retour a la ligne automatique, voir ConstruireFormulaireSaisie).
Public Const VEN_ADR_SAISIE_NOTES As String = "C32"
Public Const VEN_ADR_SAISIE_MESSAGE As String = "B33"     ' decalee de B32 a B33
Public Const VEN_LIGNE_BOUTON_AJOUTER As Long = 35        ' decalee de 34 a 35

' Zone technique cachee : sous-categories de la categorie choisie DANS LE FORMULAIRE
' DE SAISIE (une seule liste a gerer maintenant, comme dans les formulaires precedents,
' au lieu d'une liste par ligne de grille).
Public Const VEN_COL_AIDE As Long = 26         ' colonne Z
Public Const VEN_LIGNE_AIDE_MAX As Long = 300

' Ajout 01/10/2026 (point 4 : annuler une ventilation) : 2 colonnes techniques ajoutees
' a TblOperations pour memoriser la categorie/sous-categorie d'AVANT la toute premiere
' ventilation d'une operation, afin de pouvoir les restaurer si la ventilation est
' supprimee plus tard (voir mod_ControleCategories.ControleVentiler et
' mod_RechercheOperations.RevoirVentilationRO).
Public Const NOM_COL_CAT_AVANT_VENTILATION As String = "CategorieAvantVentilation"
Public Const NOM_COL_SOUS_AVANT_VENTILATION As String = "SousCategorieAvantVentilation"


' =====================================================================================
' MACRO D'ENSEMBLE : cree/met a jour la table de stockage PUIS le formulaire
' =====================================================================================
Public Sub PreparerPhase4Ventilation()
    PreparerTableVentilations
    AjouterColonnesAnnulationVentilation
    CreerFeuilleVentilation
End Sub


' =====================================================================================
' ETAPE 1 : feuille de donnees "Ventilations" + tableau "TblVentilations"
' =====================================================================================
' Sans danger a relancer : les colonnes deja presentes ne sont jamais touchees, seules
' les colonnes manquantes sont ajoutees (utile si vous avez installe une version
' anterieure de ce module).
Public Sub PreparerTableVentilations()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim colonnesSante As Variant
    Dim i As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(VEN_NOM_FEUILLE_DONNEES)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = VEN_NOM_FEUILLE_DONNEES
    End If

    On Error Resume Next
    Set tbl = ws.ListObjects(VEN_NOM_TABLE)
    On Error GoTo 0

    If tbl Is Nothing Then
        ws.Range("A1:E1").NumberFormat = "@"
        ws.Range("A1:E1").value = Array("ID_Transaction", "Categorie", "SousCategorie", "Montant", "DateVentilation")
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:E1"), , xlYes)
        tbl.Name = VEN_NOM_TABLE
        ' Pas de mise en forme ici : juste apres sa creation, le tableau n'a AUCUNE
        ' ligne de donnees (DataBodyRange est vide), lui appliquer un format plante.
        ' Chaque ligne recoit deja son propre format au moment ou elle est ecrite
        ' (voir mod_Ventilation.AjouterLigneVentilation).
    End If

    ' --- Colonnes de suivi sante (memes noms que dans TblOperations) -------------------
    colonnesSante = Array("Notes", "Date_consult", "Spe_Consult", "StatutSante", "SoldeSante", _
                          "DepassementHoraires", "CommentaireSante", "Franchise", "Beneficiaire")

    For i = LBound(colonnesSante) To UBound(colonnesSante)
        If Not ColonneExisteDansTable(tbl, CStr(colonnesSante(i))) Then
            tbl.ListColumns.Add.Name = CStr(colonnesSante(i))
        End If
    Next i

    ws.Visible = xlSheetHidden

End Sub

Private Function ColonneExisteDansTable(ByVal tbl As ListObject, ByVal nom As String) As Boolean
    Dim lc As ListColumn
    On Error Resume Next
    Set lc = tbl.ListColumns(nom)
    On Error GoTo 0
    ColonneExisteDansTable = Not (lc Is Nothing)
End Function

' Ajout 01/10/2026 (point 4 : annuler une ventilation) -----------------------------------
' Ajoute a TblOperations (PAS TblVentilations) les 2 colonnes techniques qui memorisent la
' categorie/sous-categorie d'avant la toute premiere ventilation d'une operation. Sans
' danger a relancer : ne touche a rien si les colonnes existent deja. Toujours ajoutees A
' LA FIN du tableau (jamais au milieu), pour ne decaler aucune colonne existante -- meme
' principe que mod_Categories.AjouterColonneSousCategorie.
Public Sub AjouterColonnesAnnulationVentilation()

    Dim tblOps As ListObject

    Set tblOps = mod_DonneesTable.GetOperationsTable()
    If tblOps Is Nothing Then Exit Sub

    If Not ColonneExisteDansTable(tblOps, NOM_COL_CAT_AVANT_VENTILATION) Then
        tblOps.ListColumns.Add.Name = NOM_COL_CAT_AVANT_VENTILATION
    End If
    If Not ColonneExisteDansTable(tblOps, NOM_COL_SOUS_AVANT_VENTILATION) Then
        tblOps.ListColumns.Add.Name = NOM_COL_SOUS_AVANT_VENTILATION
    End If

End Sub


' =====================================================================================
' ETAPE 2 : feuille-formulaire "frm_Ventilation"
' =====================================================================================
Public Sub CreerFeuilleVentilation()

    Dim ws As Worksheet
    Dim wsPrecedente As Worksheet
    Dim reponse As VbMsgBoxResult

    Set wsPrecedente = ActiveSheet

    Set ws = FeuilleSansErreur(VEN_NOM_FEUILLE)

    If Not ws Is Nothing Then
        reponse = MsgBox(mod_Display.FR("La feuille '") & VEN_NOM_FEUILLE & mod_Display.FR("' existe d{e2}j{a2}.") & vbCrLf & _
                         mod_Display.FR("Voulez-vous la reconstruire enti{e1}rement (sa mise en forme sera perdue) ?"), _
                         vbYesNo + vbQuestion, mod_Display.FR("Confirmation de reconstruction"))
        If reponse = vbNo Then
            MsgBox mod_Display.FR("Installation annul{e2}e, rien n'a {e2}t{e2} modifi{e2}."), vbInformation
            Exit Sub
        End If
        ws.Visible = xlSheetVisible
        ws.Cells.UnMerge
        ws.Cells.Clear
        SupprimerFormes ws
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = VEN_NOM_FEUILLE
    End If

    Application.ScreenUpdating = False

    MettreEnFormeGenerale ws
    ConstruireBoutonsGlobaux ws
    ConstruireEntete ws
    ConstruireTotaux ws
    ConstruireGrilleAffichage ws
    ConstruireFormulaireSaisie ws

    ws.Columns(VEN_COL_AIDE).Hidden = True
    ws.Visible = xlSheetVeryHidden

    On Error Resume Next
    wsPrecedente.Activate
    On Error GoTo 0
    Application.ScreenUpdating = True

    MsgBox mod_Display.FR("La feuille '") & VEN_NOM_FEUILLE & mod_Display.FR("' a {e2}t{e2} cr{e2}{e2}e puis masqu{e2}e.") & vbCrLf & vbCrLf & _
           mod_Display.FR("Si ce n'est pas deja fait : collez le code des {e2}v{e2}nements dans la feuille") & vbCrLf & _
           "(CodeName : " & ws.CodeName & mod_Display.FR(") -- voir CodeBehind_frm_Ventilation.txt."), _
           vbInformation, mod_Display.FR("Installation termin{e2}e")

End Sub


' =====================================================================================
' Mise en forme generale
' =====================================================================================
Private Sub MettreEnFormeGenerale(ByVal ws As Worksheet)

    ws.Activate
    ActiveWindow.DisplayGridlines = False

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 26
    ws.Columns("C").ColumnWidth = 30
    ws.Columns("D").ColumnWidth = 16
    ws.Columns("E").ColumnWidth = 30
    ws.Columns("F").ColumnWidth = 16
    ws.Columns("G").ColumnWidth = 2

    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10

    ws.Range("A1").Select
    ActiveWindow.DisplayHeadings = False

End Sub


' =====================================================================================
' Boutons globaux : "Terminer" (finalise TOUTE la ventilation) et "Annuler"
' =====================================================================================
Private Sub ConstruireBoutonsGlobaux(ByVal ws As Worksheet)

    Dim zone As Range

    ws.rows(VEN_LIGNE_BOUTONS).RowHeight = 26

    Set zone = ws.Cells(VEN_LIGNE_BOUTONS, 2)
    AjouterBouton ws, zone.Left, zone.Top, zone.Width, zone.Height, _
                  mod_Display.FR("Terminer la ventilation"), "VenTerminer", "btnVenTerminer"

    Set zone = ws.Cells(VEN_LIGNE_BOUTONS, 3)
    AjouterBouton ws, zone.Left, zone.Top, zone.Width, zone.Height, "Annuler", "VenAnnuler", "btnVenAnnuler"

    ' Ajout 01/10/2026 (point 4) : bouton de suppression d'une ventilation EXISTANTE.
    ' Masque par defaut a la construction : c'est mod_Ventilation.OuvrirVentilation qui le
    ' rend visible, et seulement quand la ventilation ouverte existait deja AVANT l'ouverture
    ' du formulaire (rien a supprimer pour une ventilation toute neuve). Un bouton (objet
    ' Button/Shape) flotte AU-DESSUS des cellules : lui donner la largeur de 3 cellules ne
    ' les fusionne pas, aucun risque d'erreur "cellule fusionnee" ici.
    Set zone = ws.Range(ws.Cells(VEN_LIGNE_BOUTONS, 4), ws.Cells(VEN_LIGNE_BOUTONS, 6))
    AjouterBouton ws, zone.Left, zone.Top, zone.Width, zone.Height, _
                  mod_Display.FR("Supprimer cette ventilation"), "VenSupprimerVentilation", "btnVenSupprimerVentilation"
    ws.Shapes("btnVenSupprimerVentilation").Visible = False

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
' En-tete : informations (lecture seule) de l'operation a ventiler -- INCHANGE
' =====================================================================================
Private Sub ConstruireEntete(ByVal ws As Worksheet)

    With ws.Range("B4:F4")
        .Merge
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(60, 60, 60)
    End With
    ws.Range(VEN_ADR_TITRE).value = mod_Display.FR("Ventilation de l'op{e2}ration")

    EcrireEtiquette ws, "B6", "Date"
    EcrireEtiquette ws, "B7", "Tiers"
    EcrireEtiquette ws, "B8", mod_Display.FR("Libell{e2} / Notes")
    EcrireEtiquette ws, "B9", "Montant"
    EcrireEtiquette ws, "B10", mod_Display.FR("Cat{e2}gorie actuelle")

    With ws.Range("C6:C10")
        .NumberFormat = "@"
        .WrapText = True
        .VerticalAlignment = xlCenter
        .Font.Size = 10
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(225, 225, 220)
        .Borders(xlInsideHorizontal).Color = RGB(225, 225, 220)
    End With
    ws.Range(VEN_ADR_MONTANT).Font.Bold = True

End Sub

Private Sub EcrireEtiquette(ByVal ws As Worksheet, ByVal adresse As String, ByVal texte As String)
    With ws.Range(adresse)
        .value = texte
        .Font.Bold = True
        .Font.Color = RGB(110, 110, 110)
        .VerticalAlignment = xlCenter
    End With
End Sub


' =====================================================================================
' Bloc des totaux -- INCHANGE
' =====================================================================================
Private Sub ConstruireTotaux(ByVal ws As Worksheet)

    EcrireEtiquette ws, "B12", mod_Display.FR("Montant {a2} ventiler")
    EcrireEtiquette ws, "B13", "Total saisi"
    EcrireEtiquette ws, "B14", mod_Display.FR("Reste {a2} ventiler")

    With ws.Range("C12:C14")
        ' Format numerique avec le symbole Euro en suffixe (voir mod_Ventilation pour
        ' le detail de la construction de cette chaine).
        .NumberFormat = "#,##0.00" & Chr(34) & " " & ChrW(8364) & Chr(34)
        .Font.Bold = True
        .VerticalAlignment = xlCenter
    End With
    ws.Range(VEN_ADR_RESTE).Font.Size = 12

End Sub


' =====================================================================================
' Tableau D'AFFICHAGE des lignes deja ajoutees (lignes 16 a 26)
' =====================================================================================
' Ces cellules ne sont PAS destinees a etre tapees directement par l'operateur : elles
' sont remplies par le programme (mod_Ventilation) au fur et a mesure des ajouts, et se
' modifient uniquement via le formulaire de saisie du bas (bouton "Editer"). Un fond
' gris leger les distingue visuellement des zones de saisie (jaune pale).
Private Sub ConstruireGrilleAffichage(ByVal ws As Worksheet)

    Dim ligne As Long

    With ws.Cells(VEN_LIGNE_GRILLE_ENTETE, VEN_COL_CAT)
        .value = mod_Display.FR("Cat{e2}gorie")
    End With
    With ws.Cells(VEN_LIGNE_GRILLE_ENTETE, VEN_COL_SOUS)
        .value = mod_Display.FR("Sous-cat{e2}gorie")
    End With
    With ws.Cells(VEN_LIGNE_GRILLE_ENTETE, VEN_COL_MONTANT)
        .value = "Montant"
    End With
    With ws.Cells(VEN_LIGNE_GRILLE_ENTETE, VEN_COL_NOTES)
        .value = "Notes"
    End With
    With ws.Range(ws.Cells(VEN_LIGNE_GRILLE_ENTETE, VEN_COL_CAT), ws.Cells(VEN_LIGNE_GRILLE_ENTETE, VEN_COL_EDITER))
        .Font.Bold = True
        .Font.Color = RGB(31, 78, 121)
        .Interior.Color = RGB(240, 240, 235)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With

    For ligne = VEN_LIGNE_GRILLE_DEBUT To VEN_LIGNE_GRILLE_FIN
        With ws.Range(ws.Cells(ligne, VEN_COL_CAT), ws.Cells(ligne, VEN_COL_NOTES))
            .Interior.Color = RGB(242, 242, 240)          ' gris tres pale = affichage, pas saisie
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(220, 220, 215)
        End With
        ws.Cells(ligne, VEN_COL_CAT).NumberFormat = "@"
        ws.Cells(ligne, VEN_COL_SOUS).NumberFormat = "@"
        ws.Cells(ligne, VEN_COL_MONTANT).NumberFormat = "#,##0.00"
        ws.Cells(ligne, VEN_COL_NOTES).NumberFormat = "@"
        ws.rows(ligne).RowHeight = 18

        ' Bouton "Editer" de cette ligne. Toutes les lignes ont leur bouton des la
        ' construction (que la ligne soit remplie ou non) : cliquer sur une ligne vide
        ' affiche simplement un message, voir mod_Ventilation.VenEditerLigne.
        AjouterBoutonEditer ws, ligne
    Next ligne

End Sub

Private Sub AjouterBoutonEditer(ByVal ws As Worksheet, ByVal ligne As Long)
    Dim zone As Range
    Dim indice As Long
    indice = ligne - VEN_LIGNE_GRILLE_DEBUT + 1
    Set zone = ws.Cells(ligne, VEN_COL_EDITER)
    AjouterBouton ws, zone.Left, zone.Top, zone.Width, zone.Height, _
                  mod_Display.FR("{E2}diter"), "VenEditerLigne", "btnVenEditerLigne" & indice
End Sub


' =====================================================================================
' Formulaire de SAISIE (ajout ou edition d'UNE ligne a la fois)
' =====================================================================================
Private Sub ConstruireFormulaireSaisie(ByVal ws As Worksheet)

    ' --- Titre de la zone ---
    With ws.Range("B" & VEN_LIGNE_SAISIE_TITRE & ":F" & VEN_LIGNE_SAISIE_TITRE)
        .Merge
        .value = mod_Display.FR("Ajouter ou modifier une ligne")
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = RGB(60, 60, 60)
    End With
    ws.rows(VEN_LIGNE_SAISIE_TITRE).RowHeight = 20

    ' --- Categorie + bouton "+" (meme principe que le formulaire de controle) ---
    EcrireEtiquette ws, "B29", mod_Display.FR("Cat{e2}gorie")
    ws.Range("B29").Font.Color = RGB(31, 78, 121)
    With ws.Range(VEN_ADR_SAISIE_CAT)
        .NumberFormat = "@"
        .Interior.Color = RGB(255, 250, 225)
        .Font.Size = 11
        .Font.Bold = True
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 185, 120)
    End With
    ws.rows(29).RowHeight = 22
    Dim zoneBoutonNouvelle As Range
    Set zoneBoutonNouvelle = ws.Range("D29:E29")
    AjouterBouton ws, zoneBoutonNouvelle.Left, zoneBoutonNouvelle.Top, zoneBoutonNouvelle.Width, zoneBoutonNouvelle.Height, _
                  "+ " & mod_Display.FR("Nouvelle cat{e2}gorie"), "VenNouvelleCategorie", "btnVenNouvelleCategorie"

    ' --- Sous-categorie (liste dependante de la Categorie ci-dessus) ---
    EcrireEtiquette ws, "B30", mod_Display.FR("Sous-cat{e2}gorie")
    ws.Range("B30").Font.Color = RGB(31, 78, 121)
    With ws.Range(VEN_ADR_SAISIE_SOUS)
        .NumberFormat = "@"
        .Interior.Color = RGB(255, 250, 225)
        .Font.Size = 11
        .Font.Bold = True
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 185, 120)
    End With
    ws.rows(30).RowHeight = 22

    ' --- Montant ---
    EcrireEtiquette ws, "B31", "Montant"
    ws.Range("B31").Font.Color = RGB(31, 78, 121)
    With ws.Range(VEN_ADR_SAISIE_MONTANT)
        .NumberFormat = "#,##0.00"
        .Interior.Color = RGB(255, 250, 225)
        .Font.Size = 11
        .Font.Bold = True
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 185, 120)
    End With
    ws.rows(31).RowHeight = 22

    ' --- Notes (commentaire libre, ajout 01/10/2026, point 3) ---------------------------
    ' Volontairement une cellule SIMPLE, NON FUSIONNEE (comme Categorie/Sous-categorie/
    ' Montant ci-dessus) : voir le commentaire sur la constante VEN_ADR_SAISIE_NOTES, qui
    ' explique l'erreur 1004 provoquee par une premiere version fusionnee. La largeur un
    ' peu juste d'une seule cellule est compensee par le retour a la ligne automatique
    ' (WrapText) et une rangee plus haute.
    EcrireEtiquette ws, "B32", "Notes"
    ws.Range("B32").Font.Color = RGB(31, 78, 121)
    With ws.Range(VEN_ADR_SAISIE_NOTES)
        .NumberFormat = "@"
        .Interior.Color = RGB(255, 250, 225)
        .Font.Size = 10
        .VerticalAlignment = xlTop
        .WrapText = True
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 185, 120)
    End With
    ws.rows(32).RowHeight = 30

    ' --- Message (erreurs de saisie de cette ligne) ---
    With ws.Range(VEN_ADR_SAISIE_MESSAGE & ":F33")
        .Merge
        .WrapText = True
        .VerticalAlignment = xlTop
        .Font.Size = 9
        .Font.Color = RGB(192, 80, 0)
    End With
    ws.rows(33).RowHeight = 26

    ' --- Boutons "Ajouter la ligne" et "Effacer la saisie" ---
    ws.rows(VEN_LIGNE_BOUTON_AJOUTER).RowHeight = 24
    Dim zoneAjouter As Range, zoneEffacer As Range
    Set zoneAjouter = ws.Range("B" & VEN_LIGNE_BOUTON_AJOUTER)
    AjouterBouton ws, zoneAjouter.Left, zoneAjouter.Top, 150, zoneAjouter.Height, _
                  mod_Display.FR("Ajouter la ligne"), "VenAjouterLigne", "btnVenAjouterLigne"
    Set zoneEffacer = ws.Range("D" & VEN_LIGNE_BOUTON_AJOUTER)
    AjouterBouton ws, zoneEffacer.Left, zoneEffacer.Top, 150, zoneEffacer.Height, _
                  mod_Display.FR("Effacer la saisie"), "VenEffacerSaisie", "btnVenEffacerSaisie"

End Sub


' =====================================================================================
' OUTILS DEVELOPPEUR (Ctrl+G)
' =====================================================================================
Public Sub AfficherFeuilleVentilationPourEdition()
    Dim ws As Worksheet
    Set ws = FeuilleSansErreur(VEN_NOM_FEUILLE)
    If ws Is Nothing Then
        MsgBox mod_Display.FR("La feuille n'existe pas encore. Ex{e2}cutez PreparerPhase4Ventilation."), vbExclamation
        Exit Sub
    End If
    ws.Visible = xlSheetVisible
    ws.Activate
    MsgBox mod_Display.FR("Feuille visible. Remasquez-la avec : MasquerFeuilleVentilationApresEdition") & vbCrLf & vbCrLf & _
           mod_Display.FR("Rappel : les listes d{e2}roulantes ne sont pos{e2}es qu'{a2} l'ouverture normale du formulaire") & _
           mod_Display.FR(" (bouton Ventiler..., ou TesterVentilation) -- elles n'apparaissent pas si vous affichez juste la feuille ainsi."), _
           vbInformation
End Sub

Public Sub MasquerFeuilleVentilationApresEdition()
    Dim ws As Worksheet
    Set ws = FeuilleSansErreur(VEN_NOM_FEUILLE)
    If ws Is Nothing Then Exit Sub
    ws.Visible = xlSheetVeryHidden
    MsgBox mod_Display.FR("Feuille de nouveau masqu{e2}e."), vbInformation
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
