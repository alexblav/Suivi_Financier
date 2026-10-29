Option Explicit

' =====================================================================================
' MODULE : mod_InstallCommun
'
' RÔLE (à lire en premier) :
'   Boîte à outils PARTAGÉE par tous les modules mod_Install* qui construisent une
'   feuille-formulaire (frm_...). Tout ce qui doit être identique d'un formulaire à
'   l'autre est défini ICI, et nulle part ailleurs : couleurs, hauteurs de lignes,
'   boutons, étiquettes, champs en lecture seule / modifiables, compteur, bloc
'   "Instructions". Pour changer l'aspect de TOUS les formulaires, on modifie donc ce
'   seul fichier, puis on relance les macros de création de chaque feuille.
'
' CHARTE GRAPHIQUE (demande de l'opérateur du 09/10/2026) :
'   - Étiquette de champ ................................ RGB(31, 73, 125)
'   - Cellule NON modifiable par l'opérateur (fond) ..... RGB(240, 240, 240)
'   - Cellule MODIFIABLE par l'opérateur (fond) ......... RGB(255, 255, 235)
'   - Dans le bloc Instructions :
'       nom de bouton ............ rouge RGB(255, 0, 0), gras
'       nom d'étiquette de champ . vert RGB(0, 176, 80), gras
'       nom de feuille/formulaire  RGB(22, 54, 92)
'       mot ATTENTION ............ rouge, gras, en majuscules
'   - Quadrillage et en-têtes de lignes/colonnes masqués.
'   - Lignes de 15 points (sauf : lignes de boutons, bloc Instructions, et quelques
'     champs de texte multi-lignes, qui en ont 27).
'   - Boutons : 19,5 points de haut au maximum (0,7 cm), tous de la même hauteur et
'     alignés sur une même ligne.
'
' STRUCTURE D'EN-TÊTE COMMUNE À TOUS LES FORMULAIRES (de haut en bas) :
'     ligne 1 ............ marge (fine)
'     ligne(s) de boutons  (1 ou 2 lignes)
'     fine ligne vide
'     ligne du compteur ... (uniquement si le formulaire a un compteur)
'     fine ligne vide ..... (uniquement si le formulaire a un compteur)
'     ligne Instructions .. MASQUÉE (hauteur 0) : le texte est dans frm_Nomform_aide
'     fine ligne vide
'     ... corps du formulaire
'
' SYNTHESE (regle du 17/10/2026) : tout formulaire affiche masque Synthese (MasquerQuadrillage ->
' MasquerSynthese) ; toute sortie la reaffiche (AfficherSynthese, ActiverFeuille, et le filet de securite
' de mod_Aide.MasquerFormulaire quand plus aucun formulaire n'est affiche).
'
' AIDE : le bloc Instructions de chaque formulaire frm_Nomform est écrit dans la feuille
' très masquée frm_Nomform_aide (voir EcrireInstructions). Le bouton "Aide" du formulaire
' (AjouterBoutonAide, macro AfficherAide de mod_Aide) la montre ; son bouton "Sortir"
' (macro SortirAide) ramène au formulaire.
'
' MISE EN FORME DU TEXTE DES INSTRUCTIONS : le texte source est écrit avec de petites
' balises, interprétées par EcrireInstructions :
'     [b:Valider]   -> nom de bouton   (rouge, gras)
'     [c:Montant]   -> nom de champ    (vert, gras)
'     [f:Synthese]  -> nom de feuille ou de formulaire (RGB(22, 54, 92))
'     ATTENTION     -> repéré automatiquement (rouge, gras)
'   Les marqueurs d'accents de mod_Display.FR ({e2}, {a2}...) sont acceptés aussi.
'
' À PROPOS DES ACCENTS : les textes affichés passent par mod_Display.FR(), comme dans
' le reste du classeur. Les commentaires sont encodés en UTF-8.
' =====================================================================================

' --- Dimensions communes (en points) ------------------------------------------------
Public Const FRM_H_MARGE As Double = 6          ' ligne 1 (au-dessus des boutons)
Public Const FRM_H_BOUTONS As Double = 21       ' ligne portant des boutons
Public Const FRM_H_SEP As Double = 4            ' "fine ligne vide" de séparation
Public Const FRM_H_LIGNE As Double = 15         ' hauteur normale d'une ligne
Public Const FRM_H_MULTI As Double = 27         ' champ de texte sur 2 lignes (exception)
Public Const FRM_BTN_H As Double = 19.5         ' hauteur des boutons de la ligne de boutons
Public Const FRM_BTN_ECART As Double = 2        ' espace entre deux boutons
Public Const FRM_BTN_L As Double = 82.5         ' largeur standard d'un bouton
Public Const FRM_BTN_PLUS As Double = 26        ' largeur d'un bouton "+"
Public Const FRM_BTN_AIDE_L As Double = 50      ' largeur du bouton "Aide"

' --- Sens attendu d'un montant (PoserChampMontant / PoserValidationMontant) ------------
Public Const MONTANT_LIBRE As Long = 0           ' tout nombre (négatif = dépense)
Public Const MONTANT_POSITIF As Long = 1         ' strictement supérieur à 0
Public Const MONTANT_POSITIF_OU_NUL As Long = 2  ' supérieur ou égal à 0

' --- Types de balises du bloc Instructions ---------------------------------------------
Private Const T_BOUTON As Long = 1
Private Const T_CHAMP As Long = 2
Private Const T_FEUILLE As Long = 3
Private Const T_ALERTE As Long = 4

' Résultat de l'analyse des balises (mémoire de travail de EcrireInstructions).
Private m_Debut() As Long
Private m_Long() As Long
Private m_Type() As Long
Private m_Nb As Long


' =====================================================================================
' COULEURS (des fonctions, car RGB() n'est pas autorisé dans une constante VBA)
' =====================================================================================
Public Function CoulEtiquette() As Long
    CoulEtiquette = RGB(31, 73, 125)
End Function

Public Function CoulFondLecture() As Long
    CoulFondLecture = RGB(240, 240, 240)
End Function

Public Function CoulFondSaisie() As Long
    CoulFondSaisie = RGB(255, 255, 235)
End Function

Public Function CoulNomBouton() As Long
    CoulNomBouton = RGB(255, 0, 0)
End Function

Public Function CoulNomChamp() As Long
    CoulNomChamp = RGB(0, 176, 80)
End Function

Public Function CoulNomFeuille() As Long
    CoulNomFeuille = RGB(22, 54, 92)
End Function

Public Function CoulBordureLecture() As Long
    CoulBordureLecture = RGB(200, 200, 200)
End Function

Public Function CoulBordureSaisie() As Long
    CoulBordureSaisie = RGB(200, 190, 140)
End Function


' =====================================================================================
' INTITULÉS DE BOUTONS (un même intitulé pour une même fonction, sur tous les formulaires)
' =====================================================================================
Public Function CapValider() As String
    CapValider = "Valider"
End Function

Public Function CapAnnuler() As String
    CapAnnuler = "Annuler"
End Function

Public Function CapSortir() As String
    CapSortir = "Sortir"
End Function

' Revient au formulaire precedent sans rien modifier.
Public Function CapRetour() As String
    CapRetour = "Retour"
End Function

Public Function CapPasser() As String
    CapPasser = "Passer"
End Function

' Valide l'ensemble du travail du formulaire puis le referme.
Public Function CapEnregistrer() As String
    CapEnregistrer = "Enregistrer"
End Function

' Ajoute une valeur à la liste du champ voisin (ou ouvre la création d'une valeur).
Public Function CapAjouter() As String
    CapAjouter = "+"
End Function

' Crayon (U+2712) suivi du sélecteur d'emoji (U+FE0F), comme posé à la main par l'opérateur.
Public Function CapEditer() As String
    CapEditer = ChrW(&H2712) & ChrW(&HFE0F)
End Function

Public Function CapPrecedent() As String
    CapPrecedent = "< " & mod_Display.FR("Pr{e2}c{e2}dent")
End Function

Public Function CapSuivant() As String
    CapSuivant = mod_Display.FR("Suivant >")
End Function

Public Function CapEffacer() As String
    CapEffacer = mod_Display.FR("Effacer")
End Function


' =====================================================================================
' COMPTEUR : un seul format pour tous les formulaires, "Opération: x/y"
' =====================================================================================
Public Function TexteCompteur(ByVal x As Long, ByVal y As Long) As String
    TexteCompteur = mod_Display.FR("Op{e2}ration: ") & x & "/" & y
End Function


' =====================================================================================
' PRÉPARATION DE LA FEUILLE
' =====================================================================================

' Renvoie la feuille prête à être (re)construite, ou Nothing si l'opérateur refuse la
' reconstruction d'une feuille existante.
'   supprimerNoms    : supprime aussi les noms définis qui pointent sur la feuille
'   supprimerTableau : supprime aussi les tableaux Excel (ListObjects) de la feuille
Public Function PreparerFeuille(ByVal nomFeuille As String, _
                                Optional ByVal supprimerNoms As Boolean = False, _
                                Optional ByVal supprimerTableau As Boolean = False) As Worksheet

    Dim ws As Worksheet
    Dim reponse As VbMsgBoxResult
    Dim i As Long

    Set ws = TrouverFeuille(nomFeuille)

    If Not ws Is Nothing Then
        reponse = MsgBox(mod_Display.FR("La feuille '") & nomFeuille & mod_Display.FR("' existe d{e2}j{a2}.") & vbCrLf & _
                         mod_Display.FR("Voulez-vous la reconstruire enti{e1}rement (sa mise en forme sera perdue) ?"), _
                         vbYesNo + vbQuestion, mod_Display.FR("Confirmation de reconstruction"))
        If reponse = vbNo Then
            MsgBox mod_Display.FR("Installation annul{e2}e, rien n'a {e2}t{e2} modifi{e2}."), vbInformation
            Exit Function
        End If

        ' Impossible de modifier une feuille "très masquée" : on la rend d'abord visible.
        ws.Visible = xlSheetVisible
        ws.Activate
        On Error Resume Next
        ActiveWindow.FreezePanes = False
        ActiveWindow.Split = False
        On Error GoTo 0

        If supprimerTableau Then
            For i = ws.ListObjects.Count To 1 Step -1
                ws.ListObjects(i).Delete
            Next i
        End If

        ws.Cells.UnMerge
        ws.Cells.Clear
        On Error Resume Next
        ws.Cells.Validation.Delete
        ws.Cells.ClearComments
        On Error GoTo 0
        SupprimerFormes ws
        If supprimerNoms Then SupprimerNomsDeLaFeuille ws

        ' Cells.Clear efface contenus et formats, mais PAS les hauteurs de lignes ni les
        ' largeurs de colonnes (ni les colonnes masquées) : on les remet à zéro.
        ws.Rows.UseStandardHeight = True
        ws.Columns.UseStandardWidth = True
        ws.Columns.Hidden = False
    Else
        ' Nouvelle feuille, placée en dernière position pour ne pas perturber les onglets.
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nomFeuille
    End If

    Set PreparerFeuille = ws

End Function

' Aspect "formulaire" : sans quadrillage ni en-têtes, police Calibri 10, lignes de 15
' points (jusqu'à derniereLigne), largeurs de colonnes données dans l'ordre à partir de la
' colonne A.
Public Sub MettreEnForme(ByVal ws As Worksheet, ByVal largeurs As Variant, Optional ByVal derniereLigne As Long = 40)

    Dim i As Long

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False

    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    ws.Cells.VerticalAlignment = xlCenter
    ' (Worksheet.StandardHeight est en lecture seule : on pose donc la hauteur normale
    ' explicitement sur les lignes du formulaire, jusqu'a derniereLigne.)
    ws.Range(ws.Rows(1), ws.Rows(derniereLigne)).RowHeight = FRM_H_LIGNE

    For i = LBound(largeurs) To UBound(largeurs)
        ws.Columns(i - LBound(largeurs) + 1).ColumnWidth = largeurs(i)
    Next i

    ws.Range("A1").Select

End Sub

' A appeler juste apres ws.Activate quand le programme AFFICHE un formulaire : le
' quadrillage et les en-tetes de lignes/colonnes sont des reglages de la FENETRE (et non
' de la feuille), qu'Excel peut remettre a leur valeur par defaut. On les re-masque donc
' a chaque affichage.
'
' REGLE UNIQUE SUR SYNTHESE (decision operateur du 17/10/2026) : tout formulaire (feuille frm_...)
' affiche MASQUE Synthese (MasquerSynthese), pour ne pas inciter l'operateur a quitter le formulaire
' en revenant sur Synthese. Cette procedure etant appelee a CHAQUE affichage d'un formulaire,
' c'est ici que la regle est appliquee. Toute sortie vers Synthese passe par AfficherSynthese, et tout
' retour vers la feuille d'origine d'un formulaire modal par ActiverFeuille : une feuille tres masquee
' ne peut pas etre activee directement (erreur 1004).
Public Sub MasquerQuadrillage()
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    On Error GoTo 0
    MasquerSynthese
End Sub

' Masque Synthese (tres masquee) si la feuille active est un formulaire (frm_...). Sans effet sinon,
' pour ne jamais masquer l'ecran central quand aucun formulaire n'est affiche.
Public Sub MasquerSynthese()
    Dim nomFeuille As Variant
    Dim wsSyn As Worksheet
    On Error Resume Next
    If Left$(ActiveSheet.Name, 4) <> "frm_" Then Exit Sub
    ' Menu ET Synthese (28/10/2026) : les deux ecrans centraux sont masques pendant un formulaire.
    For Each nomFeuille In Array(mod_VarGlobales.NOM_FEUILLE_MENU, mod_VarGlobales.NOM_FEUILLE_SYNTHESE)
        Set wsSyn = Nothing
        Set wsSyn = ThisWorkbook.Worksheets(CStr(nomFeuille))
        If Not wsSyn Is Nothing Then
            If wsSyn.Visible = xlSheetVisible Then wsSyn.Visible = xlSheetVeryHidden
        End If
    Next nomFeuille
    On Error GoTo 0
End Sub

' Sortie vers Synthese : la rend visible puis l'active (Activate echoue sur une feuille masquee).
' Menu ET Synthese redeviennent visibles ; c'est le Menu qui est active.
Public Sub AfficherSynthese()
    Dim wsSyn As Worksheet, wsMenu As Worksheet
    Set wsMenu = TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_MENU)
    Set wsSyn = TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_SYNTHESE)
    On Error Resume Next
    If Not wsSyn Is Nothing Then wsSyn.Visible = xlSheetVisible
    If Not wsMenu Is Nothing Then
        wsMenu.Visible = xlSheetVisible
        wsMenu.Activate
    End If
    On Error GoTo 0
End Sub

' Retour vers la feuille d'origine d'un formulaire modal (ancienne feuille active). Si c'est Synthese,
' elle est rendue visible d'abord. Une feuille introuvable (renommee, supprimee) est ignoree.
Public Sub ActiverFeuille(ByVal nomFeuille As String)
    Dim ws As Worksheet
    If StrComp(nomFeuille, mod_VarGlobales.NOM_FEUILLE_MENU, vbTextCompare) = 0 Then
        AfficherSynthese
        Exit Sub
    End If
    ' Retour sur la Synthese : les deux ecrans centraux sont rendus visibles, puis la Synthese est activee.
    If StrComp(nomFeuille, mod_VarGlobales.NOM_FEUILLE_SYNTHESE, vbTextCompare) = 0 Then
        AfficherSynthese
        On Error Resume Next
        TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_SYNTHESE).Activate
        On Error GoTo 0
        Exit Sub
    End If
    Set ws = TrouverFeuille(nomFeuille)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Activate
    On Error GoTo 0
End Sub

' Vrai si un formulaire (frm_..., hors feuilles d'aide) AUTRE que wsExclu est actuellement affiche.
Public Function AutreFormulaireVisible(ByVal wsExclu As Worksheet) As Boolean
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        If sh.Name <> wsExclu.Name And Left$(sh.Name, 4) = "frm_" And Right$(sh.Name, 5) <> "_aide" Then
            If sh.Visible = xlSheetVisible Then AutreFormulaireVisible = True: Exit Function
        End If
    Next sh
End Function

' Dernière étape : on masque complètement la feuille (invisible pour l'opérateur, même
' via clic droit > Afficher) et on revient sur la feuille qui était affichée avant.
Public Sub TerminerFeuille(ByVal ws As Worksheet, ByVal wsPrecedente As Worksheet)
    Dim wsSyn As Worksheet
    ws.Activate
    ws.Range("A1").Select
    ' Garde-fou : Synthese (écran central) doit rester visible pendant une installation,
    ' sinon il ne reste aucune feuille vers laquelle revenir.
    Set wsSyn = TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_MENU)
    If Not wsSyn Is Nothing Then
        If wsSyn.Visible <> xlSheetVisible Then wsSyn.Visible = xlSheetVisible
    End If
    Set wsSyn = TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_SYNTHESE)
    If Not wsSyn Is Nothing Then
        If wsSyn.Visible <> xlSheetVisible Then wsSyn.Visible = xlSheetVisible
    End If
    ws.Visible = xlSheetVeryHidden
    On Error Resume Next
    wsPrecedente.Activate
    On Error GoTo 0
End Sub


' =====================================================================================
' EN-TÊTE COMMUN : positions des lignes
' =====================================================================================
' nbLignesBoutons : 1 ou 2.   avecCompteur : le formulaire a-t-il un compteur ?
' Ces fonctions donnent les numéros de ligne ; les modules mod_Install* en déduisent
' leurs constantes (et vérifient leur cohérence avec PremiereLigneCorps).

Public Function LigneCompteur(ByVal nbLignesBoutons As Long, ByVal avecCompteur As Boolean) As Long
    If avecCompteur Then LigneCompteur = 3 + nbLignesBoutons
End Function

Public Function LigneInstructions(ByVal nbLignesBoutons As Long, ByVal avecCompteur As Boolean) As Long
    If avecCompteur Then
        LigneInstructions = 5 + nbLignesBoutons
    Else
        LigneInstructions = 3 + nbLignesBoutons
    End If
End Function

Public Function PremiereLigneCorps(ByVal nbLignesBoutons As Long, ByVal avecCompteur As Boolean) As Long
    PremiereLigneCorps = LigneInstructions(nbLignesBoutons, avecCompteur) + 2
End Function

' Pose les hauteurs de l'en-tête (marge, boutons, fines lignes vides, compteur). La
' hauteur de la ligne Instructions est posée par EcrireInstructions.
Public Sub PoserHauteursEntete(ByVal ws As Worksheet, ByVal nbLignesBoutons As Long, ByVal avecCompteur As Boolean)

    Dim i As Long

    ws.Rows(1).RowHeight = FRM_H_MARGE
    For i = 2 To 1 + nbLignesBoutons
        ws.Rows(i).RowHeight = FRM_H_BOUTONS
    Next i
    ws.Rows(2 + nbLignesBoutons).RowHeight = FRM_H_SEP

    If avecCompteur Then
        ws.Rows(LigneCompteur(nbLignesBoutons, True)).RowHeight = FRM_H_LIGNE
        ws.Rows(LigneCompteur(nbLignesBoutons, True) + 1).RowHeight = FRM_H_SEP
    End If

    ws.Rows(LigneInstructions(nbLignesBoutons, avecCompteur) + 1).RowHeight = FRM_H_SEP

End Sub

' Ligne du compteur : texte fusionné sur la largeur du formulaire, bleu des étiquettes,
' en gras. Le texte lui-même est écrit par le programme (voir TexteCompteur).
Public Sub PoserCompteur(ByVal ws As Worksheet, ByVal ligne As Long, ByVal colDeb As Long, ByVal colFin As Long)
    With ws.Range(ws.Cells(ligne, colDeb), ws.Cells(ligne, colFin))
        .Merge
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = CoulEtiquette()
        .Interior.Pattern = xlNone
    End With
End Sub


' =====================================================================================
' BOUTONS (contrôles de formulaire, pas ActiveX : plus fiables avec le DPI)
' =====================================================================================

' Pose un bouton sur une ligne de boutons. "gauche" est le bord gauche du bouton ; il
' est mis à jour pour que le bouton suivant se place juste à droite. Tous les boutons
' d'une ligne ont la même hauteur et le même bord haut (centrés dans la ligne).
Public Sub AjouterBoutonEntete(ByVal ws As Worksheet, ByVal ligne As Long, ByRef gauche As Double, _
                               ByVal legende As String, ByVal nomMacro As String, _
                               ByVal nomBouton As String, ByVal largeur As Double)

    Dim btn As Button
    Dim haut As Double

    haut = ws.Rows(ligne).Top + (ws.Rows(ligne).RowHeight - FRM_BTN_H) / 2
    Set btn = ws.Buttons.Add(gauche, haut, largeur, FRM_BTN_H)
    btn.Caption = legende
    btn.OnAction = nomMacro
    btn.Name = nomBouton

    gauche = gauche + largeur + FRM_BTN_ECART

End Sub

' Pose un bouton DANS une cellule du corps du formulaire (par exemple le "+" à côté d'un
' champ, ou le bouton "Éditer" d'une ligne de tableau). Il épouse la hauteur de la ligne.
Public Sub AjouterBoutonCellule(ByVal ws As Worksheet, ByVal cellule As Range, ByVal legende As String, _
                                ByVal nomMacro As String, ByVal nomBouton As String, _
                                ByVal largeur As Double, Optional ByVal decalage As Double = 2)

    Dim btn As Button

    Set btn = ws.Buttons.Add(cellule.Left + decalage, cellule.Top + 0.5, largeur, cellule.Height - 1)
    btn.Caption = legende
    btn.OnAction = nomMacro
    btn.Name = nomBouton

End Sub


' =====================================================================================
' ÉTIQUETTES ET CHAMPS
' =====================================================================================

' Étiquette de champ : gras, RGB(31, 73, 125), sans fond.
Public Sub PoserEtiquette(ByVal plage As Range, ByVal texte As String)
    If plage.Cells.Count > 1 Then plage.Merge
    With plage
        .Value = texte
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = CoulEtiquette()
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
End Sub

' Champ NON modifiable par l'opérateur (valeur affichée par le programme) : fond gris.
'   formatNombre : format de cellule ("@" pour du texte, par exemple)
'   multiLigne   : texte sur 2 lignes (retour automatique à la ligne, hauteur 27)
Public Sub PoserChampLecture(ByVal plage As Range, Optional ByVal formatNombre As String = "", _
                             Optional ByVal multiLigne As Boolean = False)
    If plage.Cells.Count > 1 Then plage.Merge
    On Error Resume Next
    plage.Validation.Delete   ' un champ repasse en lecture seule ne garde pas son garde-fou de saisie
    On Error GoTo 0
    With plage
        If formatNombre <> "" Then .NumberFormat = formatNombre
        .Interior.Color = CoulFondLecture()
        .Locked = True
        .Font.Size = 10
        .HorizontalAlignment = xlLeft
        .WrapText = multiLigne
        If multiLigne Then
            .VerticalAlignment = xlTop
        Else
            .VerticalAlignment = xlCenter
        End If
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = CoulBordureLecture()
    End With
    If multiLigne Then plage.Rows(1).RowHeight = FRM_H_MULTI
End Sub

' Champ MODIFIABLE par l'opérateur (liste ou saisie libre) : fond jaune pâle.
Public Sub PoserChampSaisie(ByVal plage As Range, Optional ByVal formatNombre As String = "", _
                            Optional ByVal multiLigne As Boolean = False)
    If plage.Cells.Count > 1 Then plage.Merge
    With plage
        If formatNombre <> "" Then .NumberFormat = formatNombre
        .Interior.Color = CoulFondSaisie()
        .Locked = False
        .Font.Size = 10
        .HorizontalAlignment = xlLeft
        .WrapText = multiLigne
        If multiLigne Then
            .VerticalAlignment = xlTop
        Else
            .VerticalAlignment = xlCenter
        End If
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = CoulBordureSaisie()
    End With
    If multiLigne Then plage.Rows(1).RowHeight = FRM_H_MULTI
End Sub

' -------------------------------------------------------------------------------------
' CHAMPS DATE ET MONTANT AVEC GARDE-FOUS (validation de donnees Excel)
' -------------------------------------------------------------------------------------
' La validation affiche une bulle d'aide (format attendu) quand l'operateur clique sur la
' cellule, et REFUSE toute valeur incorrecte des la saisie, avec un message explicite.
' Elle ne protege pas du copier-coller : les controles des boutons relisent donc aussi la
' saisie avec mod_DataStructure.TryLireDate / TryLireMontant.
Public Sub PoserChampDate(ByVal plage As Range, Optional ByVal multiLigne As Boolean = False)
    PoserChampSaisie plage, "dd/mm/yyyy", multiLigne
    PoserValidationDate plage
End Sub

Public Sub PoserChampMontant(ByVal plage As Range, Optional ByVal formatNombre As String = "#,##0.00", _
                             Optional ByVal sens As Long = MONTANT_LIBRE)
    PoserChampSaisie plage, formatNombre
    PoserValidationMontant plage, sens
End Sub

' Validation seule (sans toucher a la mise en forme) : utilisable pour reposer le garde-fou
' a chaque ouverture d'un formulaire, comme les listes deroulantes.
Public Sub PoserValidationDate(ByVal plage As Range)
    Dim cible As Range
    Set cible = plage.Cells(1, 1).MergeArea
    On Error GoTo Echec
    With cible.Validation
        .Delete
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:=CStr(mod_DataStructure.DATE_MIN_SAISIE), Formula2:=CStr(mod_DataStructure.DATE_MAX_SAISIE)
        .IgnoreBlank = True
        .ShowInput = True
        .InputTitle = "Date"
        .InputMessage = "Format : JJ/MM/AAAA (exemple : 15/03/2027)."
        .ShowError = True
        .ErrorTitle = mod_Display.FR("Date invalide")
        .ErrorMessage = mod_Display.FR("Cette date n'est pas valide. Format attendu : JJ/MM/AAAA (exemple : 15/03/2027), entre 2000 et 2100.")
    End With
    Exit Sub
Echec:
    ' La validation est un confort : son echec ne doit jamais bloquer l'ouverture du formulaire.
End Sub

Public Sub PoserValidationMontant(ByVal plage As Range, Optional ByVal sens As Long = MONTANT_LIBRE)
    Dim cible As Range
    Dim aide As String
    Set cible = plage.Cells(1, 1).MergeArea
    On Error GoTo Echec
    With cible.Validation
        .Delete
        Select Case sens
            Case MONTANT_POSITIF
                .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:="0"
                aide = mod_Display.FR("Nombre positif, sans symbole ni unit{e2} (exemple : 98,43).")
            Case MONTANT_POSITIF_OU_NUL
                .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreaterEqual, Formula1:="0"
                aide = mod_Display.FR("Nombre positif ou nul, sans symbole ni unit{e2} (exemple : 98,43).")
            Case Else
                .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="-1000000000", Formula2:="1000000000"
                aide = mod_Display.FR("Nombre, n{e2}gatif pour une d{e2}pense (exemple : -98,43).")
        End Select
        .IgnoreBlank = True
        .ShowInput = True
        .InputTitle = "Montant"
        .InputMessage = aide
        .ShowError = True
        .ErrorTitle = mod_Display.FR("Valeur invalide")
        .ErrorMessage = aide & mod_Display.FR(" Les lettres, symboles et dates sont refus{e2}s.")
    End With
    Exit Sub
Echec:
End Sub

' Titre d'un bloc du formulaire (par exemple "Opération concernée") : texte fusionné,
' gras, bleu des étiquettes, souligné d'un filet.
Public Sub PoserTitreBloc(ByVal plage As Range, ByVal texte As String)
    plage.Merge
    With plage
        .Value = texte
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = CoulEtiquette()
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeBottom).Color = CoulEtiquette()
    End With
End Sub

' Ligne de message du programme (erreurs de saisie, aide contextuelle) : texte fusionné,
' retour à la ligne, taille 9. Sa couleur est choisie par le programme selon le cas.
Public Sub PoserMessage(ByVal plage As Range)
    plage.Merge
    With plage
        .WrapText = True
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
        .Font.Size = 9
        .Font.Color = RGB(90, 90, 90)
    End With
    plage.Rows(1).RowHeight = FRM_H_MULTI
End Sub

' En-tête (ligne de titres de colonnes) d'un tableau : étiquettes de champ, filet dessous.
Public Sub PoserEnteteTableau(ByVal plage As Range)
    With plage
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = CoulEtiquette()
        .Interior.Pattern = xlNone
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeBottom).Color = CoulEtiquette()
    End With
End Sub

' Cellules d'un tableau affiché par le programme (non modifiables) : fond gris.
Public Sub PoserCorpsTableauLecture(ByVal plage As Range)
    With plage
        .Interior.Color = CoulFondLecture()
        .Locked = True
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = CoulBordureLecture()
    End With
End Sub


' =====================================================================================
' BLOC "INSTRUCTIONS"
' =====================================================================================
' AIDE DÉPORTÉE (demande opérateur du 14/10/2026) : le bloc Instructions n'est plus
' affiché sur le formulaire de travail, pour libérer de la place. Son texte est écrit dans
' une feuille d'aide "frm_Nomform_aide" (très masquée), que le bouton "Aide" du formulaire
' affiche (le formulaire reste ouvert) et que le bouton "Sortir" de l'aide referme (voir mod_Aide, à l'exécution).
'
' Cette procédure garde la signature d'origine, pour ne pas toucher aux appels de chaque
' mod_Install* : les numéros de colonnes ne servent plus (l'aide a sa propre mise en page),
' et "ligne" reste la ligne réservée aux instructions. Elle n'est PAS supprimée, mais
' masquée (hauteur 0) : les formulaires remontent ainsi visuellement d'autant, sans que
' les numéros de ligne du corps du formulaire (codés en dur dans mod_Install* et dans les
' modules d'exécution) changent.
Public Sub EcrireInstructions(ByVal ws As Worksheet, ByVal ligne As Long, _
                              ByVal colEtiqDeb As Long, ByVal colEtiqFin As Long, _
                              ByVal colTexteDeb As Long, ByVal colTexteFin As Long, _
                              ByVal texteSource As String)

    CreerFeuilleAide ws, texteSource

    ws.Rows(ligne).RowHeight = 0
    ws.Rows(ligne).Hidden = True

End Sub

' Nom de la feuille d'aide d'un formulaire : frm_Nomform -> frm_Nomform_aide.
Public Function NomFeuilleAide(ByVal nomForm As String) As String
    NomFeuilleAide = nomForm & "_aide"
End Function

' Bouton "Aide" d'un formulaire de travail : à poser sur la ligne de boutons, après les
' autres boutons ("gauche" est mis à jour comme pour AjouterBoutonEntete).
Public Sub AjouterBoutonAide(ByVal ws As Worksheet, ByVal ligne As Long, ByRef gauche As Double)
    AjouterBoutonEntete ws, ligne, gauche, "Aide", "AfficherAide", "btnAide", FRM_BTN_AIDE_L
End Sub

' (Re)construit, sans demander de confirmation, la feuille "frm_Nomform_aide" : bouton
' Sortir en haut, titre, introduction, puis la liste des champs et la liste des boutons
' (une cellule par élément). Le contenu vient de mod_InstallAide ; à défaut, le texte des
' instructions du formulaire est repris tel quel sous un titre par défaut.
Private Sub CreerFeuilleAide(ByVal wsForm As Worksheet, ByVal texteSource As String)

    Dim wsAide As Worksheet
    Dim nom As String
    Dim gauche As Double
    Dim titre As String, intro As String, champs As String, boutons As String
    Dim ligne As Long

    If Not mod_InstallAide.ChargerAide(wsForm.Name, titre, intro, champs, boutons) Then
        titre = "Aide - " & wsForm.Name
        intro = texteSource
        champs = ""
        boutons = ""
    End If

    nom = NomFeuilleAide(wsForm.Name)
    Set wsAide = TrouverFeuille(nom)

    If wsAide Is Nothing Then
        Set wsAide = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsAide.Name = nom
    Else
        wsAide.Visible = xlSheetVisible
        wsAide.Activate
        wsAide.Cells.UnMerge
        wsAide.Cells.Clear
        SupprimerFormes wsAide
        wsAide.Rows.UseStandardHeight = True
        wsAide.Columns.UseStandardWidth = True
    End If

    ' A marge ; B noms (étiquette "Instructions", champs, boutons) ; C textes.
    MettreEnForme wsAide, Array(2, 32, 105), 20

    wsAide.Rows(1).RowHeight = FRM_H_MARGE
    wsAide.Rows(2).RowHeight = FRM_H_BOUTONS
    wsAide.Rows(3).RowHeight = FRM_H_SEP

    ' Ligne 4 : titre
    With wsAide.Range("B4:C4")
        .Merge
        .NumberFormat = "@"
        .Value = mod_Display.FR(titre)
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = CoulEtiquette()
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeBottom).Color = CoulEtiquette()
    End With
    wsAide.Rows(4).RowHeight = 24
    wsAide.Rows(5).RowHeight = FRM_H_SEP

    ' Ligne 6 : introduction ; ligne 7 : séparation
    EcrireBlocInstructions wsAide, 6, 2, 2, 3, 3, intro
    wsAide.Rows(7).RowHeight = FRM_H_SEP

    ligne = 8
    If Len(champs) > 0 Then ligne = EcrireListeAide(wsAide, ligne, "Champs", champs, False)
    If Len(boutons) > 0 Then ligne = EcrireListeAide(wsAide, ligne, "Boutons", boutons, True)

    gauche = wsAide.Cells(1, 2).Left
    AjouterBoutonEntete wsAide, 2, gauche, CapSortir(), "SortirAide", "btnSortirAide", FRM_BTN_L

    wsAide.Range("A1").Select
    wsForm.Activate                      ' la feuille d'aide ne doit pas rester active pour être masquée
    wsAide.Visible = xlSheetVeryHidden

End Sub

' Écrit une liste "Nom|Description" (un élément par ligne de texte, séparés par Chr(10)) à
' partir de "ligne" : un titre de bloc, puis une ligne par élément (nom dans la colonne B,
' description dans la colonne C), puis une fine ligne vide. Renvoie la ligne suivante.
' Les noms de champs sont en vert, ceux de boutons en rouge (comme dans les instructions).
Private Function EcrireListeAide(ByVal ws As Worksheet, ByVal ligne As Long, ByVal titreBloc As String, _
                                 ByVal liste As String, ByVal sontDesBoutons As Boolean) As Long

    Dim elements() As String
    Dim i As Long
    Dim pos As Long
    Dim nomElt As String, desc As String
    Dim brutDesc As String
    Dim hNom As Double, hDesc As Double

    PoserTitreBloc ws.Range(ws.Cells(ligne, 2), ws.Cells(ligne, 3)), titreBloc
    ws.Rows(ligne).RowHeight = FRM_H_LIGNE
    ligne = ligne + 1

    elements = Split(liste, Chr(10))
    For i = LBound(elements) To UBound(elements)
        pos = InStr(elements(i), "|")
        If pos > 0 Then
            nomElt = Left$(elements(i), pos - 1)
            desc = Mid$(elements(i), pos + 1)

            With ws.Cells(ligne, 2)
                .NumberFormat = "@"
                .Value = mod_Display.FR(nomElt)
                .Font.Name = "Calibri"
                .Font.Size = 10
                .Font.Bold = True
                If sontDesBoutons Then .Font.Color = CoulNomBouton() Else .Font.Color = CoulNomChamp()
                .Interior.Color = CoulFondLecture()
                .WrapText = True
                .VerticalAlignment = xlTop
                .HorizontalAlignment = xlLeft
                .Locked = True
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                .Borders.Color = CoulBordureLecture()
            End With

            ws.Cells(ligne, 3).NumberFormat = "@"
            brutDesc = EcrireCelluleTexte(ws.Cells(ligne, 3), desc)

            hNom = HauteurTexte(ws, mod_Display.FR(nomElt), ws.Cells(ligne, 2).Width, 10, True)
            hDesc = HauteurTexte(ws, brutDesc, ws.Cells(ligne, 3).Width)
            If hNom > hDesc Then ws.Rows(ligne).RowHeight = hNom Else ws.Rows(ligne).RowHeight = hDesc
            ligne = ligne + 1
        End If
    Next i

    ws.Rows(ligne).RowHeight = FRM_H_SEP
    EcrireListeAide = ligne + 1

End Function

' Écrit, sur UNE ligne : l'étiquette "Instructions" (cellules colEtiqDeb à colEtiqFin,
' fusionnées) puis le texte (cellules colTexteDeb à colTexteFin, fusionnées), avec les
' mises en couleur décrites en tête de module. La hauteur de la ligne est calculée pour
' que tout le texte soit visible.
Private Sub EcrireBlocInstructions(ByVal ws As Worksheet, ByVal ligne As Long, _
                                   ByVal colEtiqDeb As Long, ByVal colEtiqFin As Long, _
                                   ByVal colTexteDeb As Long, ByVal colTexteFin As Long, _
                                   ByVal texteSource As String)

    Dim zoneEtiq As Range
    Dim zoneTexte As Range
    Dim brut As String

    Set zoneEtiq = ws.Range(ws.Cells(ligne, colEtiqDeb), ws.Cells(ligne, colEtiqFin))
    PoserEtiquette zoneEtiq, "Instructions"
    zoneEtiq.VerticalAlignment = xlTop

    Set zoneTexte = ws.Range(ws.Cells(ligne, colTexteDeb), ws.Cells(ligne, colTexteFin))
    zoneTexte.Merge

    brut = EcrireCelluleTexte(zoneTexte, texteSource)

    ws.Rows(ligne).RowHeight = HauteurTexte(ws, brut, zoneTexte.Width)

End Sub

' Écrit un texte (avec balises, voir en tête de module) dans une cellule ou une plage
' fusionnée : fond gris, bordure fine, texte en 9 points avec les mises en couleur
' partielles. Renvoie le texte brut écrit (sans balises).
Private Function EcrireCelluleTexte(ByVal zoneTexte As Range, ByVal texteSource As String) As String

    Dim brut As String
    Dim k As Long

    brut = AnalyserBalises(mod_Display.FR(texteSource))

    With zoneTexte
        .Value = brut
        .WrapText = True
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
        .Font.Name = "Calibri"
        .Font.Size = 9
        .Font.Bold = False
        .Font.Color = RGB(60, 60, 60)
        .Interior.Color = CoulFondLecture()
        .Locked = True
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = CoulBordureLecture()
    End With

    ' Mises en couleur partielles (après la mise en forme de l'ensemble du texte).
    For k = 1 To m_Nb
        If m_Long(k) > 0 Then
            With zoneTexte.Cells(1, 1).Characters(m_Debut(k), m_Long(k)).Font
                Select Case m_Type(k)
                    Case T_BOUTON
                        .Bold = True
                        .Color = CoulNomBouton()
                    Case T_CHAMP
                        .Bold = True
                        .Color = CoulNomChamp()
                    Case T_FEUILLE
                        .Color = CoulNomFeuille()
                    Case T_ALERTE
                        .Bold = True
                        .Color = RGB(255, 0, 0)
                End Select
            End With
        End If
    Next k

    EcrireCelluleTexte = brut

End Function

' Retire les balises [b:...], [c:...], [f:...] du texte et mémorise, pour chacune, la
' position et la longueur du texte concerné. Repère aussi les mots ATTENTION.
' Renvoie le texte brut (sans balises).
Private Function AnalyserBalises(ByVal source As String) As String

    Dim i As Long, n As Long, fin As Long, pos As Long
    Dim brut As String, contenu As String, code As String
    Dim typ As Long
    Dim traite As Boolean

    m_Nb = 0
    n = Len(source)
    i = 1

    Do While i <= n
        traite = False
        If Mid$(source, i, 1) = "[" And i + 2 <= n Then
            If Mid$(source, i + 2, 1) = ":" Then
                fin = InStr(i, source, "]")
                If fin > i + 2 Then
                    code = Mid$(source, i + 1, 1)
                    typ = 0
                    Select Case code
                        Case "b": typ = T_BOUTON
                        Case "c": typ = T_CHAMP
                        Case "f": typ = T_FEUILLE
                    End Select
                    If typ <> 0 Then
                        contenu = Mid$(source, i + 3, fin - i - 3)
                        MemoriserPlage Len(brut) + 1, Len(contenu), typ
                        brut = brut & contenu
                        i = fin + 1
                        traite = True
                    End If
                End If
            End If
        End If
        If Not traite Then
            brut = brut & Mid$(source, i, 1)
            i = i + 1
        End If
    Loop

    pos = InStr(1, brut, "ATTENTION", vbBinaryCompare)
    Do While pos > 0
        MemoriserPlage pos, Len("ATTENTION"), T_ALERTE
        pos = InStr(pos + Len("ATTENTION"), brut, "ATTENTION", vbBinaryCompare)
    Loop

    AnalyserBalises = brut

End Function

Private Sub MemoriserPlage(ByVal debut As Long, ByVal longueur As Long, ByVal typ As Long)
    m_Nb = m_Nb + 1
    ReDim Preserve m_Debut(1 To m_Nb)
    ReDim Preserve m_Long(1 To m_Nb)
    ReDim Preserve m_Type(1 To m_Nb)
    m_Debut(m_Nb) = debut
    m_Long(m_Nb) = longueur
    m_Type(m_Nb) = typ
End Sub

' Hauteur de ligne nécessaire pour afficher "texte" sur "largeurPts" points de large.
' Excel ne sait pas ajuster automatiquement la hauteur d'une cellule FUSIONNÉE : on
' fait donc ajuster une cellule temporaire (loin à droite et en bas de la feuille) dont
' la colonne a la même largeur, puis on lit la hauteur obtenue et on efface le tout.
Private Function HauteurTexte(ByVal ws As Worksheet, ByVal texte As String, ByVal largeurPts As Double, _
                              Optional ByVal taille As Double = 9, Optional ByVal gras As Boolean = False) As Double

    Const COL_TEMP As Long = 200
    Const LIGNE_TEMP As Long = 2000

    Dim cel As Range
    Dim w As Double
    Dim h As Double
    Dim cible As Double
    Dim passe As Long

    Set cel = ws.Cells(LIGNE_TEMP, COL_TEMP)
    cel.Clear
    cel.WrapText = True
    cel.Font.Name = "Calibri"
    cel.Font.Size = taille
    cel.Font.Bold = gras
    cel.NumberFormat = "@"
    cel.Value = texte

    ' 4 % de marge : les mots en gras sont un peu plus larges que le texte courant.
    cible = largeurPts * 0.96

    ws.Columns(COL_TEMP).ColumnWidth = 20
    For passe = 1 To 2
        w = ws.Columns(COL_TEMP).Width
        If w > 0 Then ws.Columns(COL_TEMP).ColumnWidth = ws.Columns(COL_TEMP).ColumnWidth * cible / w
    Next passe

    ws.Rows(LIGNE_TEMP).AutoFit
    h = ws.Rows(LIGNE_TEMP).RowHeight

    cel.Clear
    ws.Rows(LIGNE_TEMP).UseStandardHeight = True
    ws.Columns(COL_TEMP).UseStandardWidth = True

    h = h + 3
    If h < FRM_H_LIGNE Then h = FRM_H_LIGNE
    If h > 409 Then h = 409
    HauteurTexte = h

End Function


' =====================================================================================
' LISTES DÉROULANTES
' =====================================================================================
' Pose une liste déroulante (validation de données) alimentée par une plage NOMMÉE
' (Beneficiaires, Specialites, Praticiens...). Contrairement aux anciennes versions,
' une anomalie n'est plus avalée en silence : si la plage nommée n'existe pas ou si
' Excel refuse la validation, un message l'explique (sauf si silencieux = Vrai).
' Renvoie Vrai si la liste a été posée.
Public Function PoserListeDeroulante(ByVal rng As Range, ByVal nomPlage As String, _
                                     Optional ByVal silencieux As Boolean = False) As Boolean

    Dim existe As Boolean
    Dim cible As Range

    On Error Resume Next
    existe = (Len(ThisWorkbook.Names(nomPlage).RefersTo) > 0)
    On Error GoTo 0

    If Not existe Then
        If Not silencieux Then
            MsgBox mod_Display.FR("La plage nomm{e2}e '") & nomPlage & mod_Display.FR("' est introuvable dans le classeur : ") & _
                   mod_Display.FR("la liste d{e2}roulante de ") & rng.Address(False, False) & mod_Display.FR(" n'a pas {e2}t{e2} pos{e2}e."), vbExclamation
        End If
        Exit Function
    End If

    Set cible = rng.Cells(1, 1).MergeArea   ' MergeArea exige une seule cellule (erreur 1004 sinon)
    On Error GoTo Echec
    cible.Validation.Delete
    cible.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & nomPlage
    cible.Validation.IgnoreBlank = True
    cible.Validation.InCellDropdown = True
    PoserListeDeroulante = True
    Exit Function

Echec:
    If Not silencieux Then
        MsgBox mod_Display.FR("Liste d{e2}roulante '") & nomPlage & mod_Display.FR("' refus{e2}e par Excel sur ") & _
               rng.Address(False, False) & " : " & Err.Description, vbExclamation
    End If

End Function


' =====================================================================================
' NOMS DÉFINIS
' =====================================================================================
' Associe un nom à une cellule de la feuille (le recrée s'il existe déjà).
Public Sub PoserNom(ByVal ws As Worksheet, ByVal nom As String, ByVal cellule As Range)
    On Error Resume Next
    ws.Names(nom).Delete
    ThisWorkbook.Names(nom).Delete
    On Error GoTo 0
    ws.Names.Add Name:=nom, RefersTo:=cellule
End Sub

' Nom défini pour une liste déroulante dont la longueur varie : la plage commence en
' (ligneDeb, col) et s'arrête après la dernière valeur écrite (sans trou) avant ligneMax.
Public Sub PoserNomListeDynamique(ByVal ws As Worksheet, ByVal nom As String, ByVal col As Long, _
                                  ByVal ligneDeb As Long, ByVal ligneMax As Long)
    Dim debut As String, plage As String
    debut = ws.Cells(ligneDeb, col).Address(True, True)
    plage = ws.Range(ws.Cells(ligneDeb, col), ws.Cells(ligneMax, col)).Address(True, True)
    On Error Resume Next
    ThisWorkbook.Names(nom).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nom, RefersTo:="=OFFSET('" & ws.Name & "'!" & debut & _
        ",0,0,MAX(1,COUNTA('" & ws.Name & "'!" & plage & ")),1)"
End Sub


' =====================================================================================
' OUTILS INTERNES ET OUTILS DÉVELOPPEUR
' =====================================================================================
Public Function TrouverFeuille(ByVal nomFeuille As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nomFeuille)
    On Error GoTo 0
    Set TrouverFeuille = ws
End Function

' Supprime tous les boutons et formes (évite les doublons si l'installation est relancée).
' On parcourt à l'envers : supprimer en avançant ferait sauter des éléments.
Private Sub SupprimerFormes(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
End Sub

' Supprime les noms définis qui pointent sur la feuille (le nom de la feuille apparaît
' dans la référence avec ou sans apostrophes, selon qu'il contient des espaces).
Private Sub SupprimerNomsDeLaFeuille(ByVal ws As Worksheet)
    Dim i As Long
    Dim ref As String
    For i = ThisWorkbook.Names.Count To 1 Step -1
        ref = ""
        On Error Resume Next
        ref = ThisWorkbook.Names(i).RefersTo
        If InStr(1, ref, ws.Name & "!", vbTextCompare) > 0 Or InStr(1, ref, ws.Name & "'!", vbTextCompare) > 0 Then
            ThisWorkbook.Names(i).Delete
        End If
        On Error GoTo 0
    Next i
End Sub

' Rend la feuille visible pour la retoucher à la main.
Public Sub AfficherPourEdition(ByVal nomFeuille As String, ByVal macroCreation As String, ByVal macroMasquage As String)
    Dim ws As Worksheet
    Set ws = TrouverFeuille(nomFeuille)
    If ws Is Nothing Then
        MsgBox mod_Display.FR("La feuille '") & nomFeuille & mod_Display.FR("' n'existe pas encore.") & vbCrLf & _
               mod_Display.FR("Ex{e2}cutez d'abord la macro ") & macroCreation & ".", vbExclamation
        Exit Sub
    End If
    ws.Visible = xlSheetVisible
    ws.Activate
    MsgBox mod_Display.FR("Feuille visible. Remasquez-la ensuite avec : ") & macroMasquage, vbInformation
End Sub

' Remasque la feuille après retouche.
Public Sub MasquerApresEdition(ByVal nomFeuille As String)
    Dim ws As Worksheet
    Set ws = TrouverFeuille(nomFeuille)
    If ws Is Nothing Then Exit Sub
    ws.Visible = xlSheetVeryHidden
    MsgBox mod_Display.FR("Feuille de nouveau masqu{e2}e."), vbInformation
End Sub
