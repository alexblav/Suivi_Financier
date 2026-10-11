Attribute VB_Name = "mod_InstallAnalyse"
Option Explicit

' =====================================================================================
' MODULE : mod_InstallAnalyse
'
' ANALYSE BUDGÉTAIRE - V1 (construction). La logique est dans mod_Analyse.
'
' Une seule macro à lancer : InstallerAnalyse. Elle crée :
'   1. la feuille frm_Analyse (masquée) : paramètres, tableau de répartition, tableau
'      d'évolution sur 12 mois, deux graphiques, boutons ;
'   2. le bouton "Analyse budgétaire" sur la feuille Synthese (sous le bloc prévisionnel,
'      déplaçable à la main).
' Aucun code n'est à coller dans la feuille (pas d'événements) : tout passe par des boutons.
'
' INSTALLATION (une seule fois) :
'   1. Alt+F11, Fichier > Importer : mod_VarGlobales.bas, mod_Analyse.bas, CE fichier,
'      mod_RechercheOperations.bas et mod_InstallRechercheOperations.bas (mis à jour).
'   2. Ctrl+G, taper : InstallerAnalyse, puis Entrée.
'   3. Sur Synthese, cliquer sur le bouton "Analyse budgétaire".
'
' RÉINSTALLATION possible (la feuille est reconstruite : les paramètres reviennent aux
' valeurs par défaut, y compris la liste des catégories exclues).
' =====================================================================================

Private Const LIGNE_BOUTONS As Long = 2
Private Const LIGNE_INSTRUCTIONS As Long = 4
Private Const SYN_LIGNE_BOUTON As Long = 22

Public Sub InstallerAnalyse()

    Dim ws As Worksheet
    Dim wsPrecedente As Worksheet

    Set wsPrecedente = ActiveSheet

    Set ws = mod_InstallCommun.PreparerFeuille(mod_VarGlobales.NOM_FEUILLE_ANALYSE, True, True)
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    ' A marge ; B libellés ; C à N : 12 mois de l'évolution ; O : total. Le tableau de
    ' répartition occupe B:E, son graphique G:O.
    mod_InstallCommun.MettreEnForme ws, Array(2, 34, 14, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 11), 100

    mod_InstallCommun.PoserHauteursEntete ws, 1, False
    mod_InstallCommun.EcrireInstructions ws, LIGNE_INSTRUCTIONS, 2, 2, 3, 15, TexteInstructions()

    ConstruireListesTechniques ws
    ConstruireParametres ws
    ConstruireTableauRepartition ws
    ConstruireTableauEvolution ws
    ConstruireGraphiques ws

    ' Colonnes techniques (listes déroulantes, données de l'anneau) : masquées.
    ws.Range(ws.Columns(mod_Analyse.AN_COL_LISTE_CAT), ws.Columns(mod_Analyse.AN_COL_LISTE_ANNEE)).EntireColumn.Hidden = True

    ConstruireBoutons ws        ' en dernier : les boutons se placent d'après les hauteurs de lignes
    ConstruireBoutonSynthese

    mod_InstallCommun.TerminerFeuille ws, wsPrecedente
    Application.ScreenUpdating = True

    MsgBox mod_Display.FR("L'analyse budg{e2}taire est install{e2}e.") & vbCrLf & vbCrLf & _
           mod_Display.FR("Cliquez sur le bouton 'Analyse budg{e2}taire' de la feuille Synthese."), _
           vbInformation, mod_Display.FR("Installation termin{e2}e")

End Sub


Private Function TexteInstructions() As String
    TexteInstructions = _
        "[b:Actualiser] recalcule l'analyse apr{e1}s un changement de [c:P{e2}riode], [c:Ann{e2}e], [c:Mois], [c:Cat{e2}gorie] " & _
        "ou de type de graphique. Cliquez une ligne du tableau puis [b:D{e2}tailler la s{e2}lection] pour voir les " & _
        "sous-cat{e2}gories de cette cat{e2}gorie ; [b:Toutes les cat{e2}gories] revient {a2} la vue g{e2}n{e2}rale. " & _
        "[b:Voir les op{e2}rations] ouvre [f:frm_RechercheOperations] sur les op{e2}rations de la ligne s{e2}lectionn{e2}e " & _
        "(ou de tout l'affichage si aucune ligne n'est s{e2}lectionn{e2}e) ; au retour, l'analyse est recalcul{e2}e. " & _
        "Les montants sont des d{e2}penses nettes des remboursements, par mois de budget. " & _
        "Les cat{e2}gories de la liste [c:Cat{e2}gories exclues] ne sont jamais compt{e2}es."
End Function


' =====================================================================================
' Listes techniques (colonnes masquées) et noms
' =====================================================================================
Private Sub ConstruireListesTechniques(ByVal ws As Worksheet)

    Dim i As Long
    Dim nomF As String

    nomF = "'" & ws.Name & "'!"

    ws.Cells(2, mod_Analyse.AN_COL_LISTE_PERIODE).Value = mod_Analyse.LibPeriodeMois()
    ws.Cells(3, mod_Analyse.AN_COL_LISTE_PERIODE).Value = mod_Analyse.LibPeriodeAnnee()
    ws.Cells(4, mod_Analyse.AN_COL_LISTE_PERIODE).Value = mod_Analyse.LibPeriode12()

    ws.Cells(2, mod_Analyse.AN_COL_LISTE_GRAPH).Value = mod_Analyse.LibBarres()
    ws.Cells(3, mod_Analyse.AN_COL_LISTE_GRAPH).Value = mod_Analyse.LibAnneau()

    ws.Cells(2, mod_Analyse.AN_COL_LISTE_EVO).Value = mod_Analyse.LibEvoBarres()
    ws.Cells(3, mod_Analyse.AN_COL_LISTE_EVO).Value = mod_Analyse.LibEvoCourbes()

    For i = 1 To 12
        ws.Cells(1 + i, mod_Analyse.AN_COL_LISTE_MOIS).Value = i
    Next i

    ws.Cells(2, mod_Analyse.AN_COL_LISTE_CAT).Value = mod_Analyse.LibToutes()
    ws.Cells(2, mod_Analyse.AN_COL_LISTE_ANNEE).Value = Year(Date)

    PoserNomListe "ListePeriodeAnalyse", ws, mod_Analyse.AN_COL_LISTE_PERIODE, 2, 4
    PoserNomListe "ListeGraphAnalyse", ws, mod_Analyse.AN_COL_LISTE_GRAPH, 2, 3
    PoserNomListe "ListeEvoAnalyse", ws, mod_Analyse.AN_COL_LISTE_EVO, 2, 3
    PoserNomListe "ListeMoisAnalyse", ws, mod_Analyse.AN_COL_LISTE_MOIS, 2, 13

    ' Listes dont la longueur varie (remplies par mod_Analyse à chaque actualisation).
    PoserNomDynamique "ListeCategoriesAnalyse", ws, mod_Analyse.AN_COL_LISTE_CAT, 300
    PoserNomDynamique "ListeAnneesAnalyse", ws, mod_Analyse.AN_COL_LISTE_ANNEE, 60

End Sub

Private Sub PoserNomListe(ByVal nom As String, ByVal ws As Worksheet, ByVal col As Long, ByVal ligneDeb As Long, ByVal ligneFin As Long)
    On Error Resume Next
    ThisWorkbook.Names(nom).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nom, RefersTo:="='" & ws.Name & "'!" & _
        ws.Range(ws.Cells(ligneDeb, col), ws.Cells(ligneFin, col)).Address(True, True)
End Sub

' Plage qui s'adapte au nombre de valeurs écrites à partir de la ligne 2.
Private Sub PoserNomDynamique(ByVal nom As String, ByVal ws As Worksheet, ByVal col As Long, ByVal ligneMax As Long)
    Dim debut As String, plage As String
    debut = ws.Cells(2, col).Address(True, True)
    plage = ws.Range(ws.Cells(2, col), ws.Cells(ligneMax, col)).Address(True, True)
    On Error Resume Next
    ThisWorkbook.Names(nom).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nom, RefersTo:="=OFFSET('" & ws.Name & "'!" & debut & ",0,0,MAX(1,COUNTA('" & ws.Name & "'!" & plage & ")),1)"
End Sub


' =====================================================================================
' Paramètres et catégories exclues
' =====================================================================================
Private Sub ConstruireParametres(ByVal ws As Worksheet)

    Dim l As Long

    mod_InstallCommun.PoserTitreBloc ws.Range("B6:E6"), mod_Display.FR("Param{e1}tres")

    mod_InstallCommun.PoserEtiquette ws.Range("B7"), mod_Display.FR("P{e2}riode")
    mod_InstallCommun.PoserEtiquette ws.Range("B8"), mod_Display.FR("Ann{e2}e de budget")
    mod_InstallCommun.PoserEtiquette ws.Range("B9"), "Mois de budget"
    mod_InstallCommun.PoserEtiquette ws.Range("B10"), mod_Display.FR("Cat{e2}gorie {a2} d{e2}tailler")
    mod_InstallCommun.PoserEtiquette ws.Range("B11"), mod_Display.FR("Graphique de r{e2}partition")
    mod_InstallCommun.PoserEtiquette ws.Range("B12"), mod_Display.FR("Graphique d'{e2}volution")

    For l = 7 To 12
        mod_InstallCommun.PoserChampSaisie ws.Range(ws.Cells(l, 3), ws.Cells(l, 5))
    Next l

    ws.Range(mod_Analyse.AN_ADR_PERIODE).Value = mod_Analyse.LibPeriodeMois()
    ws.Range(mod_Analyse.AN_ADR_ANNEE).Value = Year(Date)
    ws.Range(mod_Analyse.AN_ADR_MOIS).Value = Month(Date)
    ws.Range(mod_Analyse.AN_ADR_CATEGORIE).Value = mod_Analyse.LibToutes()
    ws.Range(mod_Analyse.AN_ADR_GRAPH).Value = mod_Analyse.LibBarres()
    ws.Range(mod_Analyse.AN_ADR_GRAPH_EVO).Value = mod_Analyse.LibEvoBarres()

    mod_InstallCommun.PoserListeDeroulante ws.Range(mod_Analyse.AN_ADR_PERIODE), "ListePeriodeAnalyse"
    mod_InstallCommun.PoserListeDeroulante ws.Range(mod_Analyse.AN_ADR_ANNEE), "ListeAnneesAnalyse"
    mod_InstallCommun.PoserListeDeroulante ws.Range(mod_Analyse.AN_ADR_MOIS), "ListeMoisAnalyse"
    mod_InstallCommun.PoserListeDeroulante ws.Range(mod_Analyse.AN_ADR_CATEGORIE), "ListeCategoriesAnalyse"
    mod_InstallCommun.PoserListeDeroulante ws.Range(mod_Analyse.AN_ADR_GRAPH), "ListeGraphAnalyse"
    mod_InstallCommun.PoserListeDeroulante ws.Range(mod_Analyse.AN_ADR_GRAPH_EVO), "ListeEvoAnalyse"

    ' Catégories exclues : colonnes G à I, lignes 7 à 12 (6 emplacements, modifiables).
    mod_InstallCommun.PoserTitreBloc ws.Range("G6:I6"), mod_Display.FR("Cat{e2}gories exclues")
    For l = mod_Analyse.AN_LIGNE_EXCL_DEB To mod_Analyse.AN_LIGNE_EXCL_FIN
        mod_InstallCommun.PoserChampSaisie ws.Range(ws.Cells(l, mod_Analyse.AN_COL_EXCL), ws.Cells(l, mod_Analyse.AN_COL_EXCL + 2))
    Next l
    ws.Cells(mod_Analyse.AN_LIGNE_EXCL_DEB, mod_Analyse.AN_COL_EXCL).Value = "Virements internes"
    ws.Cells(mod_Analyse.AN_LIGNE_EXCL_DEB + 1, mod_Analyse.AN_COL_EXCL).Value = "Revenus"

    ' Ligne de résumé (écrite par le programme)
    mod_InstallCommun.PoserChampLecture ws.Range("B14:O14")
    ws.Range("B14").Font.Bold = True
    ws.Range("B14").Font.Color = mod_InstallCommun.CoulEtiquette()

End Sub


' =====================================================================================
' Tableau de répartition (B:E)
' =====================================================================================
Private Sub ConstruireTableauRepartition(ByVal ws As Worksheet)

    Dim lEnt As Long, lDeb As Long, lFin As Long, lTot As Long

    lEnt = mod_Analyse.AN_LIGNE_ENTETE
    lDeb = lEnt + 1
    lFin = lEnt + mod_Analyse.AN_NB_LIGNES
    lTot = mod_Analyse.AN_LIGNE_TOTAL

    ws.Cells(lEnt, 2).Value = mod_Display.FR("Cat{e2}gorie / sous-cat{e2}gorie")
    ws.Cells(lEnt, 3).Value = mod_Display.FR("D{e2}penses (")& ChrW(8364) & ")"
    ws.Cells(lEnt, 4).Value = "%"
    ws.Cells(lEnt, 5).Value = mod_Display.FR("Nb op.")
    mod_InstallCommun.PoserEnteteTableau ws.Range(ws.Cells(lEnt, 2), ws.Cells(lEnt, 5))
    ws.Range(ws.Cells(lEnt, 3), ws.Cells(lEnt, 5)).HorizontalAlignment = xlRight

    mod_InstallCommun.PoserCorpsTableauLecture ws.Range(ws.Cells(lDeb, 2), ws.Cells(lTot, 5))
    ws.Range(ws.Cells(lDeb, 3), ws.Cells(lTot, 3)).NumberFormat = "#,##0.00"
    ws.Range(ws.Cells(lDeb, 4), ws.Cells(lTot, 4)).NumberFormat = "0.0%"
    ws.Range(ws.Cells(lDeb, 5), ws.Cells(lTot, 5)).NumberFormat = "0"
    ws.Range(ws.Cells(lTot, 2), ws.Cells(lTot, 5)).Font.Bold = True

End Sub


' =====================================================================================
' Tableau d'évolution (B:O)
' =====================================================================================
Private Sub ConstruireTableauEvolution(ByVal ws As Worksheet)

    Dim lEnt As Long, lDeb As Long, lFin As Long, lTot As Long

    lEnt = mod_Analyse.AN_LIGNE_ENTETE_EVO
    lDeb = lEnt + 1
    lFin = lEnt + mod_Analyse.AN_NB_LIGNES_EVO
    lTot = mod_Analyse.AN_LIGNE_TOTAL_EVO

    mod_InstallCommun.PoserTitreBloc ws.Range(ws.Cells(mod_Analyse.AN_LIGNE_TITRE_EVO, 2), ws.Cells(mod_Analyse.AN_LIGNE_TITRE_EVO, 15)), _
        mod_Display.FR("{E2}volution mensuelle des d{e2}penses (12 mois, en ") & ChrW(8364) & ")"

    mod_InstallCommun.PoserEnteteTableau ws.Range(ws.Cells(lEnt, 2), ws.Cells(lEnt, 15))
    ws.Range(ws.Cells(lEnt, 3), ws.Cells(lEnt, 15)).HorizontalAlignment = xlRight

    mod_InstallCommun.PoserCorpsTableauLecture ws.Range(ws.Cells(lDeb, 2), ws.Cells(lTot, 15))
    ws.Range(ws.Cells(lDeb, 3), ws.Cells(lTot, 15)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(lTot, 2), ws.Cells(lTot, 15)).Font.Bold = True
    ws.Range(ws.Cells(lDeb, 15), ws.Cells(lTot, 15)).Font.Bold = True

End Sub


' =====================================================================================
' Graphiques (leurs séries sont posées par mod_Analyse à chaque actualisation)
' =====================================================================================
Private Sub ConstruireGraphiques(ByVal ws As Worksheet)

    Dim co As ChartObject
    Dim r As Range

    ' Répartition : à droite du tableau, de G16 à O47
    Set r = ws.Range(ws.Cells(mod_Analyse.AN_LIGNE_ENTETE, 7), ws.Cells(mod_Analyse.AN_LIGNE_TOTAL, 15))
    Set co = ws.ChartObjects.Add(r.Left, r.Top, r.Width, r.Height)
    co.Name = mod_Analyse.AN_GRAPH_REPARTITION
    co.Chart.ChartType = xlBarClustered
    co.Chart.PlotVisibleOnly = False      ' les données de l'anneau sont dans des colonnes masquées
    co.Chart.HasTitle = True
    co.Chart.ChartTitle.Text = mod_Display.FR("R{e2}partition des d{e2}penses")

    ' Évolution : sous le tableau d'évolution
    Set r = ws.Range(ws.Cells(mod_Analyse.AN_LIGNE_TOTAL_EVO + 2, 2), ws.Cells(mod_Analyse.AN_LIGNE_TOTAL_EVO + 30, 15))
    Set co = ws.ChartObjects.Add(r.Left, r.Top, r.Width, r.Height)
    co.Name = mod_Analyse.AN_GRAPH_EVOLUTION
    co.Chart.ChartType = xlColumnStacked
    co.Chart.HasTitle = True
    co.Chart.ChartTitle.Text = mod_Display.FR("{E2}volution mensuelle des d{e2}penses")

End Sub


' =====================================================================================
' Boutons
' =====================================================================================
Private Sub ConstruireBoutons(ByVal ws As Worksheet)

    Dim gauche As Double

    gauche = ws.Cells(1, 2).Left

    mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, "Actualiser", _
        "ActualiserAnalyse", "btnActualiserAnalyse", 80
    mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, mod_Display.FR("D{e2}tailler la s{e2}lection"), _
        "DetaillerSelectionAnalyse", "btnDetaillerAnalyse", 135
    mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, mod_Display.FR("Toutes les cat{e2}gories"), _
        "ToutesCategoriesAnalyse", "btnToutesCategoriesAnalyse", 125
    mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, mod_Display.FR("Voir les op{e2}rations"), _
        "VoirOperationsAnalyse", "btnVoirOperationsAnalyse", 120
    mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, mod_InstallCommun.CapSortir(), _
        "SortirAnalyse", "btnSortirAnalyse", mod_InstallCommun.FRM_BTN_L

End Sub

' Bouton d'accès sur la feuille Synthese (déplaçable à la main ensuite).
Private Sub ConstruireBoutonSynthese()

    Dim ws As Worksheet
    Dim btn As Button

    Set ws = mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_SYNTHESE)
    If ws Is Nothing Then
        MsgBox mod_Display.FR("La feuille Synthese est introuvable : le bouton d'acc{e1}s n'a pas {e2}t{e2} cr{e2}{e2}."), vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    ws.Shapes("btnAnalyseBudgetaire").Delete
    On Error GoTo 0

    Set btn = ws.Buttons.Add(ws.Cells(SYN_LIGNE_BOUTON, 1).Left + 12, ws.Cells(SYN_LIGNE_BOUTON, 1).Top, 152, 28)
    btn.Caption = mod_Display.FR("Analyse budg{e2}taire")
    btn.OnAction = "AfficherAnalyse"
    btn.Name = "btnAnalyseBudgetaire"

End Sub


' =====================================================================================
' OUTILS DÉVELOPPEUR (Ctrl+G)
' =====================================================================================
Public Sub AfficherFeuilleAnalysePourEdition()
    mod_InstallCommun.AfficherPourEdition mod_VarGlobales.NOM_FEUILLE_ANALYSE, "InstallerAnalyse", "MasquerFeuilleAnalyseApresEdition"
End Sub

Public Sub MasquerFeuilleAnalyseApresEdition()
    mod_InstallCommun.MasquerApresEdition mod_VarGlobales.NOM_FEUILLE_ANALYSE
End Sub
