Attribute VB_Name = "mod_InstallEvolution"
Option Explicit

' =====================================================================================
' MODULE : mod_InstallEvolution
'
' ÉVOLUTION DES DÉPENSES DANS LE TEMPS (construction). La logique est dans mod_Evolution.
'
' Une seule macro à lancer : InstallerEvolution. Elle crée :
'   1. la feuille frm_Evolution (masquée) : paramètres, listes de catégories et de
'      sous-catégories, catégories exclues, tableau et graphique, boutons ;
'   2. le bouton "Évolution des dépenses" sur la feuille Synthese (sous le bouton
'      "Analyse budgétaire", déplaçable à la main).
' Cet écran est AUTONOME : il ne dépend pas de frm_Analyse (seulement du chargement des
' données de mod_AnalyseDonnees). Aucun code n'est à coller dans la feuille.
'
' INSTALLATION (une seule fois) :
'   1. Alt+F11, Fichier > Importer : mod_VarGlobales.bas, mod_InstallCommun.bas,
'      mod_AnalyseDonnees.bas, mod_Analyse.bas, mod_InstallAnalyse.bas, mod_Evolution.bas
'      et CE fichier.
'   2. Ctrl+G, taper : InstallerEvolution, puis Entrée (et InstallerAnalyse : l'ancien
'      tableau d'évolution disparaît de frm_Analyse).
'   3. Sur Synthese, cliquer sur le bouton "Évolution des dépenses".
'
' RÉINSTALLATION possible (la feuille est reconstruite : les paramètres reviennent aux
' valeurs par défaut, y compris la liste des catégories exclues).
' =====================================================================================

Private Const LIGNE_BOUTONS As Long = 2
Private Const LIGNE_INSTRUCTIONS As Long = 4
Private Const SYN_LIGNE_BOUTON As Long = 24

Public Sub InstallerEvolution()

    Dim ws As Worksheet
    Dim wsPrecedente As Worksheet

    Set wsPrecedente = ActiveSheet

    Set ws = mod_InstallCommun.PreparerFeuille(mod_VarGlobales.NOM_FEUILLE_EVOLUTION, True, True)
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    ' A marge ; B libellés ; C à O : champs, puis colonnes de temps du tableau (jusqu'à 48).
    mod_InstallCommun.MettreEnForme ws, Array(2, 34, 14, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 11), 100

    mod_InstallCommun.PoserHauteursEntete ws, 1, False
    mod_InstallCommun.EcrireInstructions ws, LIGNE_INSTRUCTIONS, 2, 2, 3, 15, TexteInstructions()

    ConstruireListesTechniques ws
    ConstruireParametres ws
    mod_Evolution.FormaterTableauEvolution ws, 12, "mensuelle"
    ConstruireGraphique ws

    ' Colonnes techniques (listes déroulantes) : masquées.
    ws.Range(ws.Columns(mod_Evolution.EV_COL_LISTE_PERIODE), ws.Columns(mod_Evolution.EV_COL_LISTE_SOUS)).EntireColumn.Hidden = True

    ConstruireBoutons ws        ' en dernier : les boutons se placent d'après les hauteurs de lignes
    ConstruireBoutonSynthese

    mod_InstallCommun.TerminerFeuille ws, wsPrecedente
    Application.ScreenUpdating = True

    MsgBox mod_Display.FR("L'{e2}volution des d{e2}penses est install{e2}e.") & vbCrLf & vbCrLf & _
           mod_Display.FR("Cliquez sur le bouton '{E2}volution des d{e2}penses' de la feuille Synthese."), _
           vbInformation, mod_Display.FR("Installation termin{e2}e")

End Sub


Private Function TexteInstructions() As String
    TexteInstructions = _
        "[b:Actualiser] recalcule le tableau et le graphique apr{e1}s tout changement. [c:P{e2}riode] va du mois en cours " & _
        "{a2} tout l'historique (les derniers mois incluent le mois en cours). [c:Pas de temps] regroupe les colonnes par " & _
        "mois, trimestre ou ann{e2}e (Automatique choisit selon la dur{e2}e). Les listes [c:Cat{e2}gories] et [c:Sous-cat{e2}gories] " & _
        "limitent l'analyse ; vides, toutes les cat{e2}gories sont prises, sauf celles de [c:Cat{e2}gories exclues] (qui ne " & _
        "s'appliquent plus d{e1}s qu'un choix est fait). Avec une seule cat{e2}gorie, le d{e2}tail passe aux sous-cat{e2}gories ; " & _
        "[c:D{e2}tail par] permet de l'imposer. [b:R{e2}initialiser les s{e2}lections] vide les deux listes. " & _
        "Les montants sont des d{e2}penses nettes des remboursements, par mois de budget."
End Function


' =====================================================================================
' Listes techniques (colonnes masquées) et noms
' =====================================================================================
Private Sub ConstruireListesTechniques(ByVal ws As Worksheet)

    EcrireListeFixe ws, mod_Evolution.EV_COL_LISTE_PERIODE, mod_Evolution.LibsPeriodes()
    EcrireListeFixe ws, mod_Evolution.EV_COL_LISTE_PAS, mod_Evolution.LibsPas()
    EcrireListeFixe ws, mod_Evolution.EV_COL_LISTE_GRAPH, mod_Evolution.LibsGraph()
    EcrireListeFixe ws, mod_Evolution.EV_COL_LISTE_DETAIL, mod_Evolution.LibsDetail()

    mod_InstallCommun.PoserNomListeDynamique ws, "ListePeriodeEvolution", mod_Evolution.EV_COL_LISTE_PERIODE, 2, 20
    mod_InstallCommun.PoserNomListeDynamique ws, "ListePasEvolution", mod_Evolution.EV_COL_LISTE_PAS, 2, 20
    mod_InstallCommun.PoserNomListeDynamique ws, "ListeGraphEvolution", mod_Evolution.EV_COL_LISTE_GRAPH, 2, 20
    mod_InstallCommun.PoserNomListeDynamique ws, "ListeDetailEvolution", mod_Evolution.EV_COL_LISTE_DETAIL, 2, 20

    ' Listes dont la longueur varie (remplies par mod_Evolution à chaque actualisation).
    mod_InstallCommun.PoserNomListeDynamique ws, "ListeCategoriesEvolution", mod_Evolution.EV_COL_LISTE_CAT, 2, mod_Evolution.EV_MAX_CAT
    mod_InstallCommun.PoserNomListeDynamique ws, "ListeSousCategoriesEvolution", mod_Evolution.EV_COL_LISTE_SOUS, 2, mod_Evolution.EV_MAX_SOUS

End Sub

Private Sub EcrireListeFixe(ByVal ws As Worksheet, ByVal col As Long, ByVal liste As Variant)
    Dim i As Long
    For i = LBound(liste) To UBound(liste)
        ws.Cells(2 + i - LBound(liste), col).Value = liste(i)
    Next i
End Sub


' =====================================================================================
' Paramètres, sélections et catégories exclues
' =====================================================================================
Private Sub ConstruireParametres(ByVal ws As Worksheet)

    Dim l As Long, k As Long, c As Long
    Dim lstPas As Variant, lstGraph As Variant, lstDetail As Variant

    lstPas = mod_Evolution.LibsPas()
    lstGraph = mod_Evolution.LibsGraph()
    lstDetail = mod_Evolution.LibsDetail()

    mod_InstallCommun.PoserTitreBloc ws.Range("B6:E6"), mod_Display.FR("Param{e1}tres")

    mod_InstallCommun.PoserEtiquette ws.Range("B7"), mod_Display.FR("P{e2}riode")
    mod_InstallCommun.PoserEtiquette ws.Range("B8"), "Pas de temps"
    mod_InstallCommun.PoserEtiquette ws.Range("B9"), mod_Display.FR("Graphique")
    mod_InstallCommun.PoserEtiquette ws.Range("B10"), mod_Display.FR("D{e2}tail par")

    For l = 7 To 10
        mod_InstallCommun.PoserChampSaisie ws.Range(ws.Cells(l, 3), ws.Cells(l, 5))
    Next l

    ws.Range(mod_Evolution.EV_ADR_PERIODE).Value = "12 derniers mois"
    ws.Range(mod_Evolution.EV_ADR_PAS).Value = lstPas(0)
    ws.Range(mod_Evolution.EV_ADR_GRAPH).Value = lstGraph(0)
    ws.Range(mod_Evolution.EV_ADR_DETAIL).Value = lstDetail(0)

    mod_InstallCommun.PoserListeDeroulante ws.Range(mod_Evolution.EV_ADR_PERIODE), "ListePeriodeEvolution"
    mod_InstallCommun.PoserListeDeroulante ws.Range(mod_Evolution.EV_ADR_PAS), "ListePasEvolution"
    mod_InstallCommun.PoserListeDeroulante ws.Range(mod_Evolution.EV_ADR_GRAPH), "ListeGraphEvolution"
    mod_InstallCommun.PoserListeDeroulante ws.Range(mod_Evolution.EV_ADR_DETAIL), "ListeDetailEvolution"

    ' Catégories à afficher (G:I) et sous-catégories à afficher (K:O) : 6 emplacements chacun.
    mod_InstallCommun.PoserTitreBloc ws.Range("G6:I6"), mod_Display.FR("Cat{e2}gories (toutes si vide)")
    mod_InstallCommun.PoserTitreBloc ws.Range("K6:O6"), mod_Display.FR("Sous-cat{e2}gories (toutes si vide)")
    For l = mod_Evolution.EV_LIGNE_SEL_DEB To mod_Evolution.EV_LIGNE_SEL_FIN
        mod_InstallCommun.PoserChampSaisie ws.Range(ws.Cells(l, mod_Evolution.EV_COL_CAT), ws.Cells(l, mod_Evolution.EV_COL_CAT + 2))
        mod_InstallCommun.PoserListeDeroulante ws.Cells(l, mod_Evolution.EV_COL_CAT), "ListeCategoriesEvolution"
        mod_InstallCommun.PoserChampSaisie ws.Range(ws.Cells(l, mod_Evolution.EV_COL_SOUS), ws.Cells(l, mod_Evolution.EV_COL_SOUS + 4))
        mod_InstallCommun.PoserListeDeroulante ws.Cells(l, mod_Evolution.EV_COL_SOUS), "ListeSousCategoriesEvolution"
    Next l

    ' Catégories exclues : une ligne de 4 emplacements (3 colonnes chacun). Sans effet dès
    ' qu'une catégorie ou une sous-catégorie est choisie ci-dessus.
    mod_InstallCommun.PoserEtiquette ws.Cells(mod_Evolution.EV_LIGNE_EXCL, 2), mod_Display.FR("Cat{e2}gories exclues")
    For k = 0 To mod_Evolution.EV_NB_EXCL - 1
        c = mod_Evolution.EV_COL_EXCL + k * mod_Evolution.EV_LARG_EXCL
        mod_InstallCommun.PoserChampSaisie ws.Range(ws.Cells(mod_Evolution.EV_LIGNE_EXCL, c), ws.Cells(mod_Evolution.EV_LIGNE_EXCL, c + mod_Evolution.EV_LARG_EXCL - 1))
        mod_InstallCommun.PoserListeDeroulante ws.Cells(mod_Evolution.EV_LIGNE_EXCL, c), "ListeCategoriesEvolution"
    Next k
    ws.Cells(mod_Evolution.EV_LIGNE_EXCL, mod_Evolution.EV_COL_EXCL).Value = "Virements internes"
    ws.Cells(mod_Evolution.EV_LIGNE_EXCL, mod_Evolution.EV_COL_EXCL + mod_Evolution.EV_LARG_EXCL).Value = "Revenus"

    ' Ligne de résumé (écrite par le programme)
    mod_InstallCommun.PoserChampLecture ws.Range("B16:O16")
    ws.Range("B16").Font.Bold = True
    ws.Range("B16").Font.Color = mod_InstallCommun.CoulEtiquette()

End Sub


' =====================================================================================
' Graphique (ses séries sont posées par mod_Evolution à chaque actualisation)
' =====================================================================================
Private Sub ConstruireGraphique(ByVal ws As Worksheet)

    Dim co As ChartObject
    Dim r As Range

    Set r = ws.Range(ws.Cells(mod_Evolution.EV_LIGNE_TOTAL + 2, 2), ws.Cells(mod_Evolution.EV_LIGNE_TOTAL + 30, 15))
    Set co = ws.ChartObjects.Add(r.Left, r.Top, r.Width, r.Height)
    co.Name = mod_Evolution.EV_GRAPH_NOM
    co.Chart.ChartType = xlColumnStacked
    co.Chart.HasTitle = True
    co.Chart.ChartTitle.Text = mod_Display.FR("{E2}volution des d{e2}penses")

End Sub


' =====================================================================================
' Boutons
' =====================================================================================
Private Sub ConstruireBoutons(ByVal ws As Worksheet)

    Dim gauche As Double

    gauche = ws.Cells(1, 2).Left

    mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, "Actualiser", _
        "ActualiserEvolution", "btnActualiserEvolution", 80
    mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, mod_Display.FR("R{e2}initialiser les s{e2}lections"), _
        "ReinitialiserSelectionsEvolution", "btnReinitialiserEvolution", 160
    mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, mod_InstallCommun.CapSortir(), _
        "SortirEvolution", "btnSortirEvolution", mod_InstallCommun.FRM_BTN_L
    mod_InstallCommun.AjouterBoutonAide ws, LIGNE_BOUTONS, gauche

End Sub

' Bouton d'accès sur la feuille Synthese : sous le bouton "Analyse budgétaire" s'il existe
' (déplaçable à la main ensuite).
Private Sub ConstruireBoutonSynthese()

    Dim ws As Worksheet
    Dim btn As Button
    Dim gauche As Double, haut As Double

    Set ws = mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_MENU)
    If ws Is Nothing Then
        MsgBox mod_Display.FR("La feuille Synthese est introuvable : le bouton d'acc{e1}s n'a pas {e2}t{e2} cr{e2}{e2}."), vbExclamation
        Exit Sub
    End If

    gauche = ws.Cells(SYN_LIGNE_BOUTON, 1).Left + 12
    haut = ws.Cells(SYN_LIGNE_BOUTON, 1).Top
    On Error Resume Next
    ws.Shapes("btnEvolutionDepenses").Delete
    gauche = ws.Shapes("btnAnalyseBudgetaire").Left
    haut = ws.Shapes("btnAnalyseBudgetaire").Top + 24
    On Error GoTo 0

    Set btn = ws.Buttons.Add(gauche, haut, 152, 17.2)
    btn.Caption = mod_Display.FR("{E2}volution des d{e2}penses")
    btn.OnAction = "AfficherEvolution"
    btn.Name = "btnEvolutionDepenses"

End Sub


' =====================================================================================
' OUTILS DÉVELOPPEUR (Ctrl+G)
' =====================================================================================
Public Sub AfficherFeuilleEvolutionPourEdition()
    mod_InstallCommun.AfficherPourEdition mod_VarGlobales.NOM_FEUILLE_EVOLUTION, "InstallerEvolution", "MasquerFeuilleEvolutionApresEdition"
End Sub

Public Sub MasquerFeuilleEvolutionApresEdition()
    mod_InstallCommun.MasquerApresEdition mod_VarGlobales.NOM_FEUILLE_EVOLUTION
End Sub
