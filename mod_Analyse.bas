Attribute VB_Name = "mod_Analyse"
Option Explicit

' =====================================================================================
' MODULE : mod_Analyse
'
' ANALYSE BUDGÉTAIRE - V1 (logique). La construction de la feuille est dans
' mod_InstallAnalyse (macro InstallerAnalyse, à lancer une seule fois).
'
' CE QUE FAIT L'ÉCRAN frm_Analyse :
'   1. RÉPARTITION des dépenses d'une période par catégorie. On peut "détailler" une
'      catégorie pour voir ses sous-catégories. Tableau + graphique (barres ou anneau).
'   2. LIEN avec l'écran de recherche : "Voir les opérations" ouvre
'      frm_RechercheOperations (préfiltre "Analyse") sur les opérations qui forment le
'      chiffre sélectionné ; on y retrouve toutes les fonctions habituelles
'      (recatégoriser, ventiler, décaler...). Au retour, l'analyse est recalculée.
'
' L'ÉVOLUTION dans le temps (par mois, trimestre ou année) n'est plus ici : elle a son
' propre écran autonome (frm_Evolution, voir mod_Evolution). Les deux écrans partagent
' seulement le chargement des données (mod_AnalyseDonnees).
'
' RÈGLES DE CALCUL (V1) :
'   - Source : voir mod_AnalyseDonnees (opérations, ventilations remplacées par leurs parts).
'   - Période : mois de budget (donc les décalages de budget sont respectés).
'   - Une dépense est un montant négatif. Pour chaque catégorie on fait le NET de la
'     période (dépenses + remboursements). Seules les catégories dont le net est une
'     dépense sont affichées : les revenus n'apparaissent donc pas.
'   - Les catégories de la liste "exclues" (bloc de l'écran) ne sont jamais comptées
'     (par défaut : Virements internes, Revenus).
'
' POINTS D'ENTRÉE (boutons) : AfficherAnalyse (feuille Synthese), ActualiserAnalyse,
'   DetaillerSelectionAnalyse, ToutesCategoriesAnalyse, VoirOperationsAnalyse, SortirAnalyse.
' Appelées par mod_RechercheOperations : PreparerFiltreAnalyse, LigneRetenueAnalyse,
'   DecrireFiltreAnalyse, RetourDepuisRecherche.
'
' À PROPOS DES ACCENTS : textes affichés via mod_Display.FR(); commentaires en UTF-8.
' =====================================================================================

' --- Disposition de la feuille (partagée avec mod_InstallAnalyse) --------------------
Public Const AN_ADR_PERIODE As String = "C7"
Public Const AN_ADR_ANNEE As String = "C8"
Public Const AN_ADR_MOIS As String = "C9"
Public Const AN_ADR_CATEGORIE As String = "C10"
Public Const AN_ADR_GRAPH As String = "C11"
Public Const AN_LIGNE_EXCL_DEB As Long = 7      ' exclusions : colonne G, lignes 7 à 12
Public Const AN_LIGNE_EXCL_FIN As Long = 12
Public Const AN_COL_EXCL As Long = 7
Public Const AN_ADR_MESSAGE As String = "B14"
Public Const AN_LIGNE_ENTETE As Long = 16       ' en-tête du tableau de répartition
Public Const AN_NB_LIGNES As Long = 30          ' lignes de données du tableau
Public Const AN_LIGNE_TOTAL As Long = 47        ' AN_LIGNE_ENTETE + AN_NB_LIGNES + 1
Public Const AN_NB_MOIS As Long = 12         ' longueur de la période "12 derniers mois"
Public Const AN_NB_TOP_ANNEAU As Long = 8

' --- Colonnes techniques masquées (listes déroulantes, données de l'anneau) ----------
Public Const AN_COL_LISTE_CAT As Long = 27      ' AA
Public Const AN_COL_ANNEAU_LIB As Long = 29     ' AC
Public Const AN_COL_ANNEAU_VAL As Long = 30     ' AD
Public Const AN_COL_LISTE_PERIODE As Long = 32  ' AF
Public Const AN_COL_LISTE_GRAPH As Long = 33    ' AG
Public Const AN_COL_LISTE_MOIS As Long = 35     ' AI
Public Const AN_COL_LISTE_ANNEE As Long = 36    ' AJ

' --- Noms de graphiques ---------------------------------------------------------------
Public Const AN_GRAPH_REPARTITION As String = "chRepartition"

' --- Filtre transmis à l'écran de recherche (voir PreparerFiltreAnalyse) -------------
Private m_fCat As String
Private m_fSous As String
Private m_fDeb As Long
Private m_fFin As Long
Private m_fExcl As Object


' =====================================================================================
' LIBELLÉS (fonctions : FR() ne peut pas servir dans une constante)
' =====================================================================================
Public Function LibPeriodeMois() As String
    LibPeriodeMois = "Mois"
End Function

Public Function LibPeriodeAnnee() As String
    LibPeriodeAnnee = mod_Display.FR("Ann{e2}e")
End Function

Public Function LibPeriode12() As String
    LibPeriode12 = "12 derniers mois"
End Function

Public Function LibToutes() As String
    LibToutes = mod_Display.FR("(toutes)")
End Function

Public Function LibBarres() As String
    LibBarres = "Barres"
End Function

Public Function LibAnneau() As String
    LibAnneau = "Anneau"
End Function

Private Function LibAutres() As String
    LibAutres = mod_AnalyseDonnees.LibAutres()
End Function

Private Function LibSansCategorie() As String
    LibSansCategorie = mod_AnalyseDonnees.LibSansCategorie()
End Function

Private Function LibSansSousCategorie() As String
    LibSansSousCategorie = mod_AnalyseDonnees.LibSansSousCategorie()
End Function

Private Function CategorieVentile() As String
    CategorieVentile = mod_AnalyseDonnees.CategorieVentile()
End Function


' =====================================================================================
' OUTILS
' =====================================================================================
Private Function FeuilleAnalyse() As Worksheet
    Set FeuilleAnalyse = mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_ANALYSE)
End Function

' Indice d'un mois : annee * 12 + (mois - 1). Permet de comparer et de compter des mois.
Private Function IndiceMois(ByVal annee As Long, ByVal mois As Long) As Long
    IndiceMois = mod_AnalyseDonnees.IndiceMois(annee, mois)
End Function

Private Function TexteMois(ByVal idx As Long) As String
    TexteMois = mod_AnalyseDonnees.TexteMois(idx)
End Function

Private Function IndiceDeLigne(ByVal annee As Long, ByVal mois As Long, ByVal dateOp As Variant) As Long
    IndiceDeLigne = mod_AnalyseDonnees.IndiceDeLigne(annee, mois, dateOp)
End Function

' Liste des catégories exclues, lue sur la feuille d'analyse (Dictionary, insensible à la casse).
Private Function LireExclusions() As Object
    Dim ws As Worksheet
    Dim d As Object
    Dim l As Long
    Dim t As String

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set ws = FeuilleAnalyse()
    If Not ws Is Nothing Then
        For l = AN_LIGNE_EXCL_DEB To AN_LIGNE_EXCL_FIN
            t = mod_DataStructure.CellText(ws.Cells(l, AN_COL_EXCL).Value2)
            If t <> "" Then d(t) = 1
        Next l
    End If
    Set LireExclusions = d
End Function

' Charge les opérations et parts en mémoire (voir mod_AnalyseDonnees).
Private Function ChargerLignes() As Boolean
    ChargerLignes = mod_AnalyseDonnees.ChargerLignes()
End Function

' Lit les paramètres de l'écran. deb/fin : période de la répartition (indices de mois);
' catFiltre : "" = toutes les catégories.
Private Function LireParametres(ByRef deb As Long, ByRef fin As Long, ByRef catFiltre As String) As Boolean

    Dim ws As Worksheet
    Dim annee As Long, mois As Long
    Dim per As String

    Set ws = FeuilleAnalyse()
    If ws Is Nothing Then Exit Function

    annee = mod_DataStructure.ToLong(ws.Range(AN_ADR_ANNEE).Value2)
    mois = mod_DataStructure.ToLong(ws.Range(AN_ADR_MOIS).Value2)
    If annee < 1900 Or mois < 1 Or mois > 12 Then
        MsgBox mod_Display.FR("Choisissez une ann{e2}e et un mois valides."), vbExclamation
        Exit Function
    End If

    per = mod_DataStructure.CellText(ws.Range(AN_ADR_PERIODE).Value2)
    If StrComp(per, LibPeriodeAnnee(), vbTextCompare) = 0 Then
        deb = IndiceMois(annee, 1)
        fin = IndiceMois(annee, 12)
    ElseIf StrComp(per, LibPeriode12(), vbTextCompare) = 0 Then
        fin = IndiceMois(annee, mois)
        deb = fin - (AN_NB_MOIS - 1)
    Else
        deb = IndiceMois(annee, mois)
        fin = deb
    End If

    catFiltre = mod_DataStructure.CellText(ws.Range(AN_ADR_CATEGORIE).Value2)
    If StrComp(catFiltre, LibToutes(), vbTextCompare) = 0 Then catFiltre = ""

    LireParametres = True

End Function


' =====================================================================================
' AGRÉGATION DE LA RÉPARTITION
' =====================================================================================
' Net par catégorie (catFiltre = "") ou par sous-catégorie de catFiltre, sur [deb, fin].
' En sortie : uniquement les clés dont le net est une DÉPENSE, montants en positif,
' triées de la plus grosse à la plus petite.
Private Sub AgregerRepartition(ByVal deb As Long, ByVal fin As Long, ByVal catFiltre As String, _
                               ByVal excl As Object, ByRef cles() As String, ByRef vals() As Double, _
                               ByRef nbs() As Long, ByRef nb As Long)

    Dim d As Object
    Dim i As Long, p As Long, cap As Long
    Dim k As String
    Dim retenue As Boolean

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    cap = 64
    ReDim cles(1 To cap)
    ReDim vals(1 To cap)
    ReDim nbs(1 To cap)
    nb = 0

    For i = 1 To mod_AnalyseDonnees.LNb
        If mod_AnalyseDonnees.LIdx(i) >= deb And mod_AnalyseDonnees.LIdx(i) <= fin Then
            retenue = False
            If catFiltre = "" Then
                If Not excl.Exists(mod_AnalyseDonnees.LCat(i)) Then
                    k = mod_AnalyseDonnees.LCat(i)
                    If k = "" Then k = LibSansCategorie()
                    retenue = True
                End If
            Else
                If StrComp(mod_AnalyseDonnees.LCat(i), catFiltre, vbTextCompare) = 0 Then
                    k = mod_AnalyseDonnees.LSous(i)
                    If k = "" Then k = LibSansSousCategorie()
                    retenue = True
                End If
            End If

            If retenue Then
                If d.Exists(k) Then
                    p = d(k)
                Else
                    nb = nb + 1
                    If nb > cap Then
                        cap = cap * 2
                        ReDim Preserve cles(1 To cap)
                        ReDim Preserve vals(1 To cap)
                        ReDim Preserve nbs(1 To cap)
                    End If
                    d(k) = nb
                    p = nb
                    cles(p) = k
                End If
                vals(p) = vals(p) + mod_AnalyseDonnees.LMont(i)
                nbs(p) = nbs(p) + 1
            End If
        End If
    Next i

    ' On ne garde que les dépenses nettes, en positif.
    Dim j As Long
    j = 0
    For i = 1 To nb
        If vals(i) < -0.005 Then
            j = j + 1
            cles(j) = cles(i)
            vals(j) = -vals(i)
            nbs(j) = nbs(i)
        End If
    Next i
    nb = j

    TrierDecroissant cles, vals, nbs, nb

End Sub

' Tri par insertion (quelques dizaines de lignes au plus).
Private Sub TrierDecroissant(ByRef cles() As String, ByRef vals() As Double, ByRef nbs() As Long, ByVal nb As Long)
    Dim i As Long, j As Long
    Dim kc As String, kv As Double, kn As Long

    For i = 2 To nb
        kc = cles(i)
        kv = vals(i)
        kn = nbs(i)
        j = i - 1
        Do While j >= 1
            If vals(j) >= kv Then Exit Do
            cles(j + 1) = cles(j)
            vals(j + 1) = vals(j)
            nbs(j + 1) = nbs(j)
            j = j - 1
        Loop
        cles(j + 1) = kc
        vals(j + 1) = kv
        nbs(j + 1) = kn
    Next i
End Sub


' =====================================================================================
' RAFRAÎCHISSEMENT DE L'ÉCRAN
' =====================================================================================
Public Sub ActualiserAnalyse()

    On Error GoTo Erreur
    Application.ScreenUpdating = False
    RafraichirCoeur
    Application.ScreenUpdating = True
    Exit Sub

Erreur:
    Application.ScreenUpdating = True
    MsgBox mod_Display.FR("Erreur pendant le calcul de l'analyse.") & vbLf & _
           "Erreur " & Err.Number & " : " & Err.Description, vbExclamation, mod_Display.FR("Analyse budg{e2}taire")

End Sub

Private Sub RafraichirCoeur()

    Dim ws As Worksheet
    Dim deb As Long, fin As Long
    Dim catFiltre As String
    Dim excl As Object
    Dim cles() As String, vals() As Double, nbs() As Long, nb As Long
    Dim total As Double, totalNb As Long
    Dim i As Long, r As Long, nbAff As Long
    Dim tabDonnees() As Variant
    Dim reste As Double, resteNb As Long
    Dim scope As String

    Set ws = FeuilleAnalyse()
    If ws Is Nothing Then
        MsgBox mod_Display.FR("La feuille d'analyse n'existe pas : lancez d'abord InstallerAnalyse."), vbExclamation
        Exit Sub
    End If

    If Not LireParametres(deb, fin, catFiltre) Then Exit Sub
    If Not ChargerLignes() Then Exit Sub

    MettreAJourListes ws
    Set excl = LireExclusions()

    ' ---------------- 1) Répartition ----------------
    AgregerRepartition deb, fin, catFiltre, excl, cles, vals, nbs, nb
    For i = 1 To nb
        total = total + vals(i)
        totalNb = totalNb + nbs(i)
    Next i

    ws.Range(ws.Cells(AN_LIGNE_ENTETE + 1, 2), ws.Cells(AN_LIGNE_TOTAL, 5)).ClearContents

    nbAff = nb
    If nb > AN_NB_LIGNES Then nbAff = AN_NB_LIGNES
    If nbAff > 0 Then
        ReDim tabDonnees(1 To nbAff, 1 To 4)
        For i = 1 To nbAff
            If nb > AN_NB_LIGNES And i = AN_NB_LIGNES Then
                ' Dernière ligne : regroupe tout le reste.
                reste = 0
                resteNb = 0
                Dim j As Long
                For j = AN_NB_LIGNES To nb
                    reste = reste + vals(j)
                    resteNb = resteNb + nbs(j)
                Next j
                tabDonnees(i, 1) = LibAutres()
                tabDonnees(i, 2) = reste
                tabDonnees(i, 3) = IIf(total > 0, reste / total, 0)
                tabDonnees(i, 4) = resteNb
            Else
                tabDonnees(i, 1) = cles(i)
                tabDonnees(i, 2) = vals(i)
                tabDonnees(i, 3) = IIf(total > 0, vals(i) / total, 0)
                tabDonnees(i, 4) = nbs(i)
            End If
        Next i
        ws.Range(ws.Cells(AN_LIGNE_ENTETE + 1, 2), ws.Cells(AN_LIGNE_ENTETE + nbAff, 5)).Value = tabDonnees
    End If

    ws.Cells(AN_LIGNE_TOTAL, 2).Value = "Total"
    ws.Cells(AN_LIGNE_TOTAL, 3).Value = total
    ws.Cells(AN_LIGNE_TOTAL, 4).Value = IIf(total > 0, 1, 0)
    ws.Cells(AN_LIGNE_TOTAL, 5).Value = totalNb

    If catFiltre = "" Then
        scope = mod_Display.FR("par cat{e2}gorie")
    Else
        scope = mod_Display.FR("de la cat{e2}gorie ") & catFiltre & mod_Display.FR(", par sous-cat{e2}gorie")
    End If

    ws.Range(AN_ADR_MESSAGE).Value = mod_Display.FR("D{e2}penses nettes ") & scope & " -- " & _
        DecrirePeriode(deb, fin) & " : " & Format(total, "#,##0.00") & " " & ChrW(8364) & _
        " (" & nb & " ligne(s), " & totalNb & mod_Display.FR(" op{e2}ration(s)/part(s)).")

    MettreAJourGraphRepartition ws, cles, vals, nb, nbAff, scope, deb, fin

    On Error Resume Next
    ws.Range(AN_ADR_PERIODE).Select     ' (sans effet si la feuille n'est pas affichée)
    On Error GoTo 0

End Sub

Private Function DecrirePeriode(ByVal deb As Long, ByVal fin As Long) As String
    If deb = fin Then
        DecrirePeriode = mod_Display.FR("mois de budget ") & TexteMois(deb)
    Else
        DecrirePeriode = mod_Display.FR("mois de budget ") & TexteMois(deb) & mod_Display.FR(" {a2} ") & TexteMois(fin)
    End If
End Function

' Listes déroulantes dynamiques : catégories présentes dans les données, années utilisées.
Private Sub MettreAJourListes(ByVal ws As Worksheet)

    Dim d As Object
    Dim liste() As String
    Dim i As Long, j As Long, n As Long
    Dim t As String, kc As String
    Dim a As Long

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    ReDim liste(1 To mod_AnalyseDonnees.LNb + 1)
    For i = 1 To mod_AnalyseDonnees.LNb
        If mod_AnalyseDonnees.LCat(i) <> "" Then
            If Not d.Exists(mod_AnalyseDonnees.LCat(i)) Then
                d(mod_AnalyseDonnees.LCat(i)) = 1
                n = n + 1
                liste(n) = mod_AnalyseDonnees.LCat(i)
            End If
        End If
    Next i
    ' Tri alphabétique par insertion
    For i = 2 To n
        kc = liste(i)
        j = i - 1
        Do While j >= 1
            If StrComp(liste(j), kc, vbTextCompare) <= 0 Then Exit Do
            liste(j + 1) = liste(j)
            j = j - 1
        Loop
        liste(j + 1) = kc
    Next i

    ws.Range(ws.Cells(2, AN_COL_LISTE_CAT), ws.Cells(300, AN_COL_LISTE_CAT)).ClearContents
    ws.Cells(2, AN_COL_LISTE_CAT).Value = LibToutes()
    For i = 1 To n
        If i + 2 <= 300 Then ws.Cells(i + 2, AN_COL_LISTE_CAT).Value = liste(i)
    Next i

    ws.Range(ws.Cells(2, AN_COL_LISTE_ANNEE), ws.Cells(60, AN_COL_LISTE_ANNEE)).ClearContents
    j = 2
    For a = mod_AnalyseDonnees.LAnneeMin To mod_AnalyseDonnees.LAnneeMax
        ws.Cells(j, AN_COL_LISTE_ANNEE).Value = a
        j = j + 1
    Next a

End Sub

Private Sub MettreAJourGraphRepartition(ByVal ws As Worksheet, ByRef cles() As String, ByRef vals() As Double, _
                                        ByVal nb As Long, ByVal nbAff As Long, ByVal scope As String, _
                                        ByVal deb As Long, ByVal fin As Long)

    Dim co As ChartObject
    Dim s As Series
    Dim anneau As Boolean
    Dim i As Long, nbA As Long
    Dim reste As Double

    Set co = Nothing
    On Error Resume Next
    Set co = ws.ChartObjects(AN_GRAPH_REPARTITION)
    On Error GoTo 0
    If co Is Nothing Then Exit Sub

    anneau = (StrComp(mod_DataStructure.CellText(ws.Range(AN_ADR_GRAPH).Value2), LibAnneau(), vbTextCompare) = 0)

    Do While co.Chart.SeriesCollection.Count > 0
        co.Chart.SeriesCollection(1).Delete
    Loop

    If nb = 0 Then
        co.Chart.HasTitle = True
        co.Chart.ChartTitle.Text = mod_Display.FR("Aucune d{e2}pense sur la p{e2}riode")
        Exit Sub
    End If

    If anneau Then
        ' Données de l'anneau : les AN_NB_TOP_ANNEAU plus grosses lignes + "Autres".
        ws.Range(ws.Cells(2, AN_COL_ANNEAU_LIB), ws.Cells(20, AN_COL_ANNEAU_VAL)).ClearContents
        nbA = nb
        If nbA > AN_NB_TOP_ANNEAU Then nbA = AN_NB_TOP_ANNEAU
        For i = 1 To nbA
            ws.Cells(1 + i, AN_COL_ANNEAU_LIB).Value = cles(i)
            ws.Cells(1 + i, AN_COL_ANNEAU_VAL).Value = vals(i)
        Next i
        If nb > AN_NB_TOP_ANNEAU Then
            For i = AN_NB_TOP_ANNEAU + 1 To nb
                reste = reste + vals(i)
            Next i
            nbA = nbA + 1
            ws.Cells(1 + nbA, AN_COL_ANNEAU_LIB).Value = LibAutres()
            ws.Cells(1 + nbA, AN_COL_ANNEAU_VAL).Value = reste
        End If
        co.Chart.ChartType = xlDoughnut
        Set s = co.Chart.SeriesCollection.NewSeries
        s.Values = ws.Range(ws.Cells(2, AN_COL_ANNEAU_VAL), ws.Cells(1 + nbA, AN_COL_ANNEAU_VAL))
        s.XValues = ws.Range(ws.Cells(2, AN_COL_ANNEAU_LIB), ws.Cells(1 + nbA, AN_COL_ANNEAU_LIB))
        s.Name = mod_Display.FR("D{e2}penses")
        co.Chart.HasLegend = True
        co.Chart.Legend.Position = xlLegendPositionRight
        s.HasDataLabels = True
        On Error Resume Next
        s.DataLabels.ShowPercentage = True
        s.DataLabels.ShowValue = False
        s.DataLabels.NumberFormat = "0%"
        On Error GoTo 0
    Else
        co.Chart.ChartType = xlBarClustered
        Set s = co.Chart.SeriesCollection.NewSeries
        s.Values = ws.Range(ws.Cells(AN_LIGNE_ENTETE + 1, 3), ws.Cells(AN_LIGNE_ENTETE + nbAff, 3))
        s.XValues = ws.Range(ws.Cells(AN_LIGNE_ENTETE + 1, 2), ws.Cells(AN_LIGNE_ENTETE + nbAff, 2))
        s.Name = mod_Display.FR("D{e2}penses")
        co.Chart.HasLegend = False
        s.HasDataLabels = True
        On Error Resume Next
        s.DataLabels.NumberFormat = "#,##0"
        co.Chart.Axes(xlCategory).ReversePlotOrder = True    ' la plus grosse dépense en haut
        On Error GoTo 0
    End If

    co.Chart.HasTitle = True
    co.Chart.ChartTitle.Text = mod_Display.FR("R{e2}partition des d{e2}penses ") & scope & " -- " & DecrirePeriode(deb, fin)

End Sub

' =====================================================================================
' BOUTONS
' =====================================================================================

' Ouverture depuis la feuille Synthese : reprend l'année et le mois choisis sur Synthese.
Public Sub AfficherAnalyse()

    Dim ws As Worksheet, wsSyn As Worksheet

    Set ws = FeuilleAnalyse()
    If ws Is Nothing Then
        MsgBox mod_Display.FR("La feuille d'analyse n'existe pas encore : lancez d'abord InstallerAnalyse (Ctrl+G)."), vbExclamation
        Exit Sub
    End If

    Set wsSyn = mod_Criteres.GetFeuille(mod_VarGlobales.NOM_FEUILLE_SYNTHESE)
    If Not wsSyn Is Nothing Then
        On Error Resume Next
        If mod_DataStructure.ToLong(wsSyn.Range("critAnnee").Value2) >= 1900 Then ws.Range(AN_ADR_ANNEE).Value = mod_DataStructure.ToLong(wsSyn.Range("critAnnee").Value2)
        If mod_DataStructure.ToLong(wsSyn.Range("critMois").Value2) >= 1 Then ws.Range(AN_ADR_MOIS).Value = mod_DataStructure.ToLong(wsSyn.Range("critMois").Value2)
        On Error GoTo 0
    End If

    ws.Visible = xlSheetVisible
    ws.Activate
    mod_InstallCommun.MasquerQuadrillage
    If Not wsSyn Is Nothing Then wsSyn.Visible = xlSheetVeryHidden

    ActualiserAnalyse

End Sub

Public Sub SortirAnalyse()
    Dim ws As Worksheet, wsSyn As Worksheet
    Set ws = FeuilleAnalyse()
    Set wsSyn = mod_Criteres.GetFeuille(mod_VarGlobales.NOM_FEUILLE_SYNTHESE)
    If Not wsSyn Is Nothing Then
        wsSyn.Visible = xlSheetVisible
        wsSyn.Activate
    End If
    If Not ws Is Nothing Then ws.Visible = xlSheetVeryHidden
End Sub

' Retour de l'écran de recherche : on réaffiche l'analyse et on la recalcule (la
' recatégorisation ou le décalage d'une opération change les chiffres).
Public Sub RetourDepuisRecherche()
    Dim ws As Worksheet
    Set ws = FeuilleAnalyse()
    If ws Is Nothing Then Exit Sub
    ws.Visible = xlSheetVisible
    ws.Activate
    mod_InstallCommun.MasquerQuadrillage
    ActualiserAnalyse
End Sub

Public Sub ToutesCategoriesAnalyse()
    Dim ws As Worksheet
    Set ws = FeuilleAnalyse()
    If ws Is Nothing Then Exit Sub
    ws.Range(AN_ADR_CATEGORIE).Value = LibToutes()
    ActualiserAnalyse
End Sub

' Ligne sélectionnée dans le tableau de répartition : renvoie son libellé, ou "".
Private Function LibelleSelectionne() As String
    Dim ws As Worksheet
    Dim ligne As Long, col As Long
    Dim t As String

    Set ws = FeuilleAnalyse()
    If ws Is Nothing Then Exit Function
    If ActiveSheet.Name <> ws.Name Then Exit Function

    ligne = ActiveCell.Row
    col = ActiveCell.Column
    If ligne <= AN_LIGNE_ENTETE Or ligne > AN_LIGNE_ENTETE + AN_NB_LIGNES Then Exit Function
    If col < 2 Or col > 5 Then Exit Function

    t = mod_DataStructure.CellText(ws.Cells(ligne, 2).Value2)
    LibelleSelectionne = t
End Function

' Bouton "Détailler la sélection" : la catégorie de la ligne sélectionnée devient le
' filtre ; le tableau et les graphiques passent alors aux sous-catégories.
Public Sub DetaillerSelectionAnalyse()

    Dim ws As Worksheet
    Dim lib As String

    Set ws = FeuilleAnalyse()
    If ws Is Nothing Then Exit Sub

    If mod_DataStructure.CellText(ws.Range(AN_ADR_CATEGORIE).Value2) <> "" And _
       StrComp(mod_DataStructure.CellText(ws.Range(AN_ADR_CATEGORIE).Value2), LibToutes(), vbTextCompare) <> 0 Then
        MsgBox mod_Display.FR("Une cat{e2}gorie est d{e2}j{a2} d{e2}taill{e2}e. Utilisez 'Voir les op{e2}rations' pour une sous-cat{e2}gorie, ou 'Toutes les cat{e2}gories' pour revenir."), vbInformation
        Exit Sub
    End If

    lib = LibelleSelectionne()
    If lib = "" Then
        MsgBox mod_Display.FR("Cliquez d'abord sur une ligne du tableau de r{e2}partition."), vbInformation
        Exit Sub
    End If
    If lib = LibAutres() Or lib = LibSansCategorie() Then
        MsgBox mod_Display.FR("Cette ligne regroupe plusieurs cat{e2}gories : choisissez une cat{e2}gorie pr{e2}cise (ou filtrez dans le champ Cat{e2}gorie)."), vbInformation
        Exit Sub
    End If

    ws.Range(AN_ADR_CATEGORIE).Value = lib
    ActualiserAnalyse

End Sub

' Bouton "Voir les opérations" : ouvre l'écran de recherche sur les opérations qui
' forment le chiffre sélectionné (ou, sans sélection, sur tout le périmètre affiché).
Public Sub VoirOperationsAnalyse()

    Dim ws As Worksheet
    Dim deb As Long, fin As Long
    Dim catFiltre As String, cat As String, sous As String, lib As String
    Dim excl As Object, k As Variant
    Dim listeExcl As String

    Set ws = FeuilleAnalyse()
    If ws Is Nothing Then Exit Sub
    If Not LireParametres(deb, fin, catFiltre) Then Exit Sub

    cat = catFiltre
    sous = ""
    lib = LibelleSelectionne()

    If lib <> "" Then
        If lib = LibAutres() Or lib = LibSansCategorie() Or lib = LibSansSousCategorie() Then
            MsgBox mod_Display.FR("Cette ligne regroupe plusieurs valeurs ou une valeur vide : s{e2}lectionnez une ligne pr{e2}cise."), vbInformation
            Exit Sub
        End If
        If catFiltre = "" Then
            cat = lib
        Else
            sous = lib
        End If
    End If

    Set excl = LireExclusions()
    For Each k In excl.Keys
        If listeExcl <> "" Then listeExcl = listeExcl & ";"
        listeExcl = listeExcl & CStr(k)
    Next k

    mod_RechercheOperations.RechercherOperations "Analyse", cat & "|" & sous & "|" & deb & "|" & fin & "|" & listeExcl

    ' L'écran de recherche prend la place de l'analyse (même principe que les autres formulaires).
    If ActiveSheet.Name = mod_VarGlobales.NOM_FEUILLE_RECHERCHE Then ws.Visible = xlSheetVeryHidden

End Sub


' =====================================================================================
' FILTRE POUR L'ÉCRAN DE RECHERCHE (préfiltre "Analyse")
' =====================================================================================
' param : "categorie|sousCategorie|indiceDebut|indiceFin|exclue1;exclue2"
Public Sub PreparerFiltreAnalyse(ByVal param As String)

    Dim p() As String
    Dim e() As String
    Dim i As Long

    m_fCat = ""
    m_fSous = ""
    m_fDeb = 0
    m_fFin = 0
    Set m_fExcl = CreateObject("Scripting.Dictionary")
    m_fExcl.CompareMode = 1

    p = Split(param, "|")
    If UBound(p) >= 3 Then
        m_fCat = p(0)
        m_fSous = p(1)
        m_fDeb = CLng(p(2))
        m_fFin = CLng(p(3))
    End If
    If UBound(p) >= 4 Then
        If p(4) <> "" Then
            e = Split(p(4), ";")
            For i = LBound(e) To UBound(e)
                m_fExcl(e(i)) = 1
            Next i
        End If
    End If

End Sub

' Vrai si la ligne (opération ou part) fait partie du filtre préparé ci-dessus.
Public Function LigneRetenueAnalyse(ByVal cat As String, ByVal sous As String, _
                                    ByVal annee As Long, ByVal mois As Long, ByVal dateOp As Variant) As Boolean

    Dim idx As Long

    If m_fExcl Is Nothing Then Exit Function

    idx = IndiceDeLigne(annee, mois, dateOp)
    If idx = 0 Then Exit Function
    If idx < m_fDeb Or idx > m_fFin Then Exit Function

    ' L'opération ventilée elle-même n'est jamais comptée : ses parts le sont.
    If StrComp(cat, CategorieVentile(), vbTextCompare) = 0 Then Exit Function

    If m_fCat = "" Then
        If m_fExcl.Exists(cat) Then Exit Function
    Else
        If StrComp(cat, m_fCat, vbTextCompare) <> 0 Then Exit Function
    End If

    If m_fSous <> "" Then
        If StrComp(sous, m_fSous, vbTextCompare) <> 0 Then Exit Function
    End If

    LigneRetenueAnalyse = True

End Function

' Phrase affichée au-dessus du tableau de l'écran de recherche.
Public Function DecrireFiltreAnalyse() As String

    Dim t As String

    t = mod_Display.FR("Analyse budg{e2}taire -- ")
    If m_fCat = "" Then
        t = t & mod_Display.FR("toutes les cat{e2}gories")
    Else
        t = t & mod_Display.FR("cat{e2}gorie ") & m_fCat
    End If
    If m_fSous <> "" Then t = t & mod_Display.FR(", sous-cat{e2}gorie ") & m_fSous
    t = t & " -- " & DecrirePeriode(m_fDeb, m_fFin) & "."
    DecrireFiltreAnalyse = t

End Function
