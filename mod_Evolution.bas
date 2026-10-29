Attribute VB_Name = "mod_Evolution"
Option Explicit

' =====================================================================================
' MODULE : mod_Evolution
'
' ÉVOLUTION DES DÉPENSES DANS LE TEMPS (logique). La construction de la feuille est dans
' mod_InstallEvolution (macro InstallerEvolution, à lancer une seule fois).
'
' ÉCRAN AUTONOME : frm_Evolution s'ouvre depuis son propre bouton de la feuille Synthese
' (AfficherEvolution). Il ne dépend d'aucun autre écran : il partage seulement le
' chargement des données avec frm_Analyse (mod_AnalyseDonnees).
'
' CE QUE FAIT L'ÉCRAN :
'   Un tableau (séries en lignes, temps en colonnes) et un graphique (barres empilées,
'   courbes ou barres groupées) des dépenses nettes, au choix de l'opérateur :
'     - PÉRIODE : mois en cours, mois dernier, 3 / 6 / 12 / 24 / 36 derniers mois (le mois
'       en cours est compris), ou tout l'historique ;
'     - PAS DE TEMPS : mois, trimestre, année, ou Automatique (mois si 36 colonnes au
'       plus, sinon trimestre, sinon année) ;
'     - CATÉGORIES : 0 à 6 catégories (vide = toutes, hors catégories exclues) ;
'     - SOUS-CATÉGORIES : 0 à 6 couples "Catégorie > Sous-catégorie" (vide = toutes) ;
'     - DÉTAIL PAR : catégorie ou sous-catégorie. Automatique = sous-catégorie dès qu'une
'       seule catégorie, ou au moins une sous-catégorie, est choisie.
'   Les sélections de catégories et de sous-catégories se CUMULENT (ET). Les catégories
'   exclues ne s'appliquent que si rien n'est sélectionné (un choix explicite l'emporte).
'
' RÈGLES DE CALCUL : identiques à frm_Analyse (mois de budget, ventilations remplacées par
' leurs parts, net des remboursements). Seules les séries dont le total est une dépense
' sont gardées ; les 8 plus grosses sont affichées, le reste est regroupé dans "Autres".
' Au plus 48 colonnes : au-delà, les 48 plus récentes sont affichées (avec un message).
'
' À PROPOS DES ACCENTS : textes affichés via mod_Display.FR(); commentaires en UTF-8.
' =====================================================================================

' --- Disposition de la feuille (partagée avec mod_InstallEvolution) -------------------
Public Const EV_ADR_PERIODE As String = "C7"
Public Const EV_ADR_PAS As String = "C8"
Public Const EV_ADR_GRAPH As String = "C9"
Public Const EV_ADR_DETAIL As String = "C10"
Public Const EV_COL_CAT As Long = 7             ' catégories : G (fusion G:I), lignes 7 à 12
Public Const EV_COL_SOUS As Long = 11           ' sous-catégories : K (fusion K:O), lignes 7 à 12
Public Const EV_LIGNE_SEL_DEB As Long = 7
Public Const EV_LIGNE_SEL_FIN As Long = 12
Public Const EV_LIGNE_EXCL As Long = 14         ' catégories exclues : ligne 14, 4 emplacements de 3 colonnes
Public Const EV_NB_EXCL As Long = 4
Public Const EV_COL_EXCL As Long = 3            ' premier emplacement : C
Public Const EV_LARG_EXCL As Long = 3
Public Const EV_ADR_MESSAGE As String = "B16"
Public Const EV_LIGNE_TITRE As Long = 18
Public Const EV_LIGNE_ENTETE As Long = 19
Public Const EV_NB_LIGNES As Long = 9           ' 8 séries + "Autres"
Public Const EV_LIGNE_TOTAL As Long = 29        ' EV_LIGNE_ENTETE + EV_NB_LIGNES + 1
Public Const EV_NB_TOP As Long = 8
Public Const EV_NB_PAS_MAX As Long = 48         ' colonnes de temps au maximum
Public Const EV_NB_PAS_AUTO As Long = 36        ' plafond du pas "Automatique"

' --- Colonnes techniques masquées (listes déroulantes) ---------------------------------
Public Const EV_COL_LISTE_PERIODE As Long = 55  ' BC
Public Const EV_COL_LISTE_PAS As Long = 56
Public Const EV_COL_LISTE_GRAPH As Long = 57
Public Const EV_COL_LISTE_DETAIL As Long = 58
Public Const EV_COL_LISTE_CAT As Long = 59
Public Const EV_COL_LISTE_SOUS As Long = 60
Public Const EV_MAX_CAT As Long = 300           ' dernière ligne utilisable de la liste des catégories
Public Const EV_MAX_SOUS As Long = 1500         ' idem pour les couples catégorie > sous-catégorie

Public Const EV_GRAPH_NOM As String = "chEvolution"

Private Const SEP_PAIRE As String = " > "


' =====================================================================================
' LIBELLÉS DES LISTES (fonctions : FR() ne peut pas servir dans une constante)
' =====================================================================================
Public Function LibsPeriodes() As Variant
    LibsPeriodes = Array("Mois en cours", "Mois dernier", "3 derniers mois", "6 derniers mois", _
                         "12 derniers mois", "24 derniers mois", "36 derniers mois", "Tous")
End Function

Public Function LibsPas() As Variant
    LibsPas = Array("Automatique", "Mois", "Trimestre", mod_Display.FR("Ann{e2}e"))
End Function

Public Function LibsGraph() As Variant
    LibsGraph = Array(mod_Display.FR("Barres empil{e2}es"), "Courbes", mod_Display.FR("Barres group{e2}es"))
End Function

Public Function LibsDetail() As Variant
    LibsDetail = Array("Automatique", mod_Display.FR("Cat{e2}gorie"), mod_Display.FR("Sous-cat{e2}gorie"))
End Function

' Rang (à partir de 1) d'un libellé dans une des listes ci-dessus ; 0 si absent.
Private Function IndexLib(ByVal liste As Variant, ByVal lib As String) As Long
    Dim i As Long
    For i = LBound(liste) To UBound(liste)
        If StrComp(CStr(liste(i)), lib, vbTextCompare) = 0 Then
            IndexLib = i - LBound(liste) + 1
            Exit Function
        End If
    Next i
End Function

' Libellé d'un couple catégorie / sous-catégorie, tel que proposé dans la liste des sous-catégories.
Private Function PaireCle(ByVal cat As String, ByVal sous As String) As String
    PaireCle = cat & SEP_PAIRE & sous
End Function


' =====================================================================================
' OUTILS
' =====================================================================================
Private Function FeuilleEvolution() As Worksheet
    Set FeuilleEvolution = mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_EVOLUTION)
End Function

Private Function Texte(ByVal ws As Worksheet, ByVal adresse As String) As String
    Texte = Trim$(mod_DataStructure.CellText(ws.Range(adresse).Value2))
End Function

' Nombre de colonnes de temps entre deb et fin (indices de mois) pour un pas de 1, 3 ou 12 mois.
Private Function NbPas(ByVal deb As Long, ByVal fin As Long, ByVal pas As Long) As Long
    NbPas = fin \ pas - deb \ pas + 1
End Function

' Sélection (catégories ou sous-catégories) lue sur la feuille : Dictionary insensible à la casse.
Private Function LireSelection(ByVal ws As Worksheet, ByVal col As Long) As Object
    Dim d As Object
    Dim l As Long
    Dim t As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    For l = EV_LIGNE_SEL_DEB To EV_LIGNE_SEL_FIN
        t = Trim$(mod_DataStructure.CellText(ws.Cells(l, col).Value2))
        If t <> "" Then d(t) = 1
    Next l
    Set LireSelection = d
End Function

Private Function LireExclusions(ByVal ws As Worksheet) As Object
    Dim d As Object
    Dim k As Long
    Dim t As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    For k = 0 To EV_NB_EXCL - 1
        t = Trim$(mod_DataStructure.CellText(ws.Cells(EV_LIGNE_EXCL, EV_COL_EXCL + k * EV_LARG_EXCL).Value2))
        If t <> "" Then d(t) = 1
    Next k
    Set LireExclusions = d
End Function

Private Function DecrirePeriode(ByVal deb As Long, ByVal fin As Long) As String
    If deb = fin Then
        DecrirePeriode = mod_Display.FR("mois de budget ") & mod_AnalyseDonnees.TexteMois(deb)
    Else
        DecrirePeriode = mod_Display.FR("mois de budget ") & mod_AnalyseDonnees.TexteMois(deb) & _
                         mod_Display.FR(" {a2} ") & mod_AnalyseDonnees.TexteMois(fin)
    End If
End Function

Private Function NomDuPas(ByVal pas As Long) As String
    Select Case pas
        Case 3: NomDuPas = "trimestre"
        Case 12: NomDuPas = mod_Display.FR("ann{e2}e")
        Case Else: NomDuPas = "mois"
    End Select
End Function

Private Function AdjectifDuPas(ByVal pas As Long) As String
    Select Case pas
        Case 3: AdjectifDuPas = "trimestrielle"
        Case 12: AdjectifDuPas = "annuelle"
        Case Else: AdjectifDuPas = "mensuelle"
    End Select
End Function


' =====================================================================================
' PÉRIODE ET PAS DE TEMPS
' =====================================================================================
' Période choisie -> indices de mois deb et fin. Faux (avec message) si rien d'exploitable.
Private Function CalculerPeriode(ByVal lib As String, ByRef deb As Long, ByRef fin As Long) As Boolean

    Dim maintenant As Long
    Dim k As Long, i As Long
    Dim nbMois As Variant

    nbMois = Array(1, 1, 3, 6, 12, 24, 36)
    maintenant = mod_AnalyseDonnees.IndiceMois(Year(Date), Month(Date))
    k = IndexLib(LibsPeriodes(), lib)

    Select Case k
        Case 1
            deb = maintenant
            fin = maintenant
        Case 2
            deb = maintenant - 1
            fin = deb
        Case 3 To 7
            fin = maintenant
            deb = maintenant - nbMois(k - 1) + 1
        Case 8
            deb = 0
            fin = 0
            For i = 1 To mod_AnalyseDonnees.LNb
                If deb = 0 Or mod_AnalyseDonnees.LIdx(i) < deb Then deb = mod_AnalyseDonnees.LIdx(i)
                If mod_AnalyseDonnees.LIdx(i) > fin Then fin = mod_AnalyseDonnees.LIdx(i)
            Next i
        Case Else
            MsgBox mod_Display.FR("Choisissez une p{e2}riode dans la liste."), vbExclamation
            Exit Function
    End Select

    CalculerPeriode = (deb > 0 And fin >= deb)

End Function

' Pas de temps en mois : 1, 3 ou 12.
Private Function CalculerPas(ByVal lib As String, ByVal deb As Long, ByVal fin As Long) As Long
    Select Case IndexLib(LibsPas(), lib)
        Case 2
            CalculerPas = 1
        Case 3
            CalculerPas = 3
        Case 4
            CalculerPas = 12
        Case Else
            If NbPas(deb, fin, 1) <= EV_NB_PAS_AUTO Then
                CalculerPas = 1
            ElseIf NbPas(deb, fin, 3) <= EV_NB_PAS_AUTO Then
                CalculerPas = 3
            Else
                CalculerPas = 12
            End If
    End Select
End Function


' =====================================================================================
' AGRÉGATION
' =====================================================================================
' mat(colonne, série) : dépenses nettes (positives) de chaque série pour chaque colonne de
' temps ; les séries sont triées par total décroissant et seules les dépenses sont gardées.
Private Sub Agreger(ByVal deb As Long, ByVal fin As Long, ByVal pas As Long, ByVal nbP As Long, _
                    ByVal cats As Object, ByVal subs As Object, ByVal excl As Object, _
                    ByVal modeSous As Boolean, _
                    ByRef cles() As String, ByRef mat() As Double, ByRef nb As Long)

    Dim d As Object
    Dim i As Long, p As Long, cap As Long, c As Long, j As Long
    Dim nbGardes As Long, debPas As Long, idx As Long
    Dim k As String, cat As String, sous As String
    Dim retenue As Boolean, filtreLibre As Boolean
    Dim cl() As String, mt() As Double
    Dim totaux() As Double
    Dim ordre() As Long
    Dim tmp As Long
    Dim tot As Double

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    cap = 64
    ReDim cl(1 To cap)
    ReDim mt(1 To nbP, 1 To cap)
    nb = 0
    debPas = deb \ pas
    filtreLibre = (cats.Count = 0 And subs.Count = 0)

    For i = 1 To mod_AnalyseDonnees.LNb
        idx = mod_AnalyseDonnees.LIdx(i)
        If idx >= deb And idx <= fin Then

            cat = mod_AnalyseDonnees.LCat(i)
            sous = mod_AnalyseDonnees.LSous(i)
            If cat = "" Then cat = mod_AnalyseDonnees.LibSansCategorie()
            If sous = "" Then sous = mod_AnalyseDonnees.LibSansSousCategorie()

            retenue = True
            If cats.Count > 0 Then
                If Not cats.Exists(cat) Then retenue = False
            End If
            If retenue And subs.Count > 0 Then
                If Not subs.Exists(PaireCle(cat, sous)) Then retenue = False
            End If
            If retenue And filtreLibre Then
                If excl.Exists(cat) Then retenue = False
            End If

            If retenue Then
                If modeSous Then
                    If cats.Count = 1 Then
                        k = sous
                    Else
                        k = PaireCle(cat, sous)
                    End If
                Else
                    k = cat
                End If

                If d.Exists(k) Then
                    p = d(k)
                Else
                    nb = nb + 1
                    If nb > cap Then
                        cap = cap * 2
                        ReDim Preserve cl(1 To cap)
                        ReDim Preserve mt(1 To nbP, 1 To cap)
                    End If
                    d(k) = nb
                    p = nb
                    cl(p) = k
                End If
                c = idx \ pas - debPas + 1
                mt(c, p) = mt(c, p) - mod_AnalyseDonnees.LMont(i)     ' dépense = positif
            End If

        End If
    Next i

    ' Totaux et tri des séries (on ne garde que celles dont le total est une dépense).
    ReDim totaux(1 To IIf(nb > 0, nb, 1))
    ReDim ordre(1 To IIf(nb > 0, nb, 1))
    For p = 1 To nb
        tot = 0
        For c = 1 To nbP
            tot = tot + mt(c, p)
        Next c
        totaux(p) = tot
        ordre(p) = p
    Next p

    For i = 2 To nb
        tmp = ordre(i)
        j = i - 1
        Do While j >= 1
            If totaux(ordre(j)) >= totaux(tmp) Then Exit Do
            ordre(j + 1) = ordre(j)
            j = j - 1
        Loop
        ordre(j + 1) = tmp
    Next i

    nbGardes = 0
    For i = 1 To nb
        If totaux(ordre(i)) > 0.005 Then nbGardes = nbGardes + 1
    Next i

    If nbGardes = 0 Then
        nb = 0
        Exit Sub
    End If

    ReDim cles(1 To nbGardes)
    ReDim mat(1 To nbP, 1 To nbGardes)
    For i = 1 To nbGardes
        cles(i) = cl(ordre(i))
        For c = 1 To nbP
            mat(c, i) = mt(c, ordre(i))
        Next c
    Next i
    nb = nbGardes

End Sub


' =====================================================================================
' LISTES DÉROULANTES (catégories et couples catégorie > sous-catégorie présents)
' =====================================================================================
Private Sub MettreAJourListes(ByVal ws As Worksheet)

    Dim dC As Object, dP As Object
    Dim lc() As String, lp() As String
    Dim nc As Long, np As Long, i As Long
    Dim cat As String, sous As String, k As String

    Set dC = CreateObject("Scripting.Dictionary")
    dC.CompareMode = 1
    Set dP = CreateObject("Scripting.Dictionary")
    dP.CompareMode = 1
    ReDim lc(1 To mod_AnalyseDonnees.LNb + 1)
    ReDim lp(1 To mod_AnalyseDonnees.LNb + 1)

    For i = 1 To mod_AnalyseDonnees.LNb
        cat = mod_AnalyseDonnees.LCat(i)
        If cat <> "" Then
            If Not dC.Exists(cat) Then
                dC(cat) = 1
                nc = nc + 1
                lc(nc) = cat
            End If
            sous = mod_AnalyseDonnees.LSous(i)
            If sous = "" Then sous = mod_AnalyseDonnees.LibSansSousCategorie()
            k = PaireCle(cat, sous)
            If Not dP.Exists(k) Then
                dP(k) = 1
                np = np + 1
                lp(np) = k
            End If
        End If
    Next i

    TrierTexte lc, nc
    TrierTexte lp, np

    EcrireListe ws, EV_COL_LISTE_CAT, lc, nc, EV_MAX_CAT
    EcrireListe ws, EV_COL_LISTE_SOUS, lp, np, EV_MAX_SOUS

End Sub

Private Sub EcrireListe(ByVal ws As Worksheet, ByVal col As Long, ByRef liste() As String, _
                        ByVal n As Long, ByVal ligneMax As Long)
    Dim i As Long
    Dim valeurs() As Variant
    ws.Range(ws.Cells(2, col), ws.Cells(ligneMax, col)).ClearContents
    If n > ligneMax - 1 Then n = ligneMax - 1
    If n < 1 Then Exit Sub
    ReDim valeurs(1 To n, 1 To 1)
    For i = 1 To n
        valeurs(i, 1) = liste(i)
    Next i
    ws.Range(ws.Cells(2, col), ws.Cells(1 + n, col)).Value = valeurs
End Sub

' Tri alphabétique par insertion des n premiers éléments.
Private Sub TrierTexte(ByRef t() As String, ByVal n As Long)
    Dim i As Long, j As Long
    Dim kc As String
    For i = 2 To n
        kc = t(i)
        j = i - 1
        Do While j >= 1
            If StrComp(t(j), kc, vbTextCompare) <= 0 Then Exit Do
            t(j + 1) = t(j)
            j = j - 1
        Loop
        t(j + 1) = kc
    Next i
End Sub


' =====================================================================================
' TABLEAU ET GRAPHIQUE
' =====================================================================================
' Met en forme le tableau pour nbCol colonnes de temps (cadre gris, en-tête, totaux). La
' zone est d'abord remise à zéro jusqu'à la largeur maximale : un tableau plus étroit que
' le précédent ne laisse donc pas de colonnes grises vides. Appelée aussi par l'installation.
Public Sub FormaterTableauEvolution(ByVal ws As Worksheet, ByVal nbCol As Long, ByVal adjectif As String)

    Dim zone As Range
    Dim derCol As Long, derTitre As Long

    derCol = 3 + nbCol                    ' colonne "Total"
    derTitre = derCol
    If derTitre < 15 Then derTitre = 15   ' le titre couvre au moins B:O

    Set zone = ws.Range(ws.Cells(EV_LIGNE_TITRE, 2), ws.Cells(EV_LIGNE_TOTAL, 3 + EV_NB_PAS_MAX))
    zone.UnMerge
    zone.Clear
    zone.Font.Name = "Calibri"
    zone.Font.Size = 10
    zone.VerticalAlignment = xlCenter

    mod_InstallCommun.PoserTitreBloc ws.Range(ws.Cells(EV_LIGNE_TITRE, 2), ws.Cells(EV_LIGNE_TITRE, derTitre)), _
        mod_Display.FR("{E2}volution ") & adjectif & mod_Display.FR(" des d{e2}penses (en ") & ChrW(8364) & ")"

    mod_InstallCommun.PoserEnteteTableau ws.Range(ws.Cells(EV_LIGNE_ENTETE, 2), ws.Cells(EV_LIGNE_ENTETE, derCol))
    ws.Range(ws.Cells(EV_LIGNE_ENTETE, 3), ws.Cells(EV_LIGNE_ENTETE, derCol)).HorizontalAlignment = xlRight

    mod_InstallCommun.PoserCorpsTableauLecture ws.Range(ws.Cells(EV_LIGNE_ENTETE + 1, 2), ws.Cells(EV_LIGNE_TOTAL, derCol))
    ws.Range(ws.Cells(EV_LIGNE_ENTETE + 1, 3), ws.Cells(EV_LIGNE_TOTAL, derCol)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(EV_LIGNE_TOTAL, 2), ws.Cells(EV_LIGNE_TOTAL, derCol)).Font.Bold = True
    ws.Range(ws.Cells(EV_LIGNE_ENTETE + 1, derCol), ws.Cells(EV_LIGNE_TOTAL, derCol)).Font.Bold = True

End Sub

' Écrit l'en-tête (colonnes de temps), les séries et les totaux. Renvoie le nombre de lignes
' de séries écrites (nbAff) et le total général.
Private Sub EcrireTableau(ByVal ws As Worksheet, ByVal deb As Long, ByVal pas As Long, ByVal nbP As Long, _
                          ByVal modeSous As Boolean, ByRef cles() As String, ByRef mat() As Double, _
                          ByVal nb As Long, ByRef nbAff As Long, ByRef totalGeneral As Double)

    Dim i As Long, c As Long, nbTop As Long
    Dim idxDebut As Long
    Dim totLigne As Double
    Dim totCol() As Double
    Dim reste() As Double
    Dim tabDonnees() As Variant

    FormaterTableauEvolution ws, nbP, AdjectifDuPas(pas)

    ' En-têtes : libellé de la série, une colonne par mois / trimestre / année, "Total".
    ws.Cells(EV_LIGNE_ENTETE, 2).Value = IIf(modeSous, mod_Display.FR("Sous-cat{e2}gorie"), mod_Display.FR("Cat{e2}gorie"))
    If pas = 1 Then
        ws.Range(ws.Cells(EV_LIGNE_ENTETE, 3), ws.Cells(EV_LIGNE_ENTETE, 2 + nbP)).NumberFormat = "mmm-yy"
    Else
        ws.Range(ws.Cells(EV_LIGNE_ENTETE, 3), ws.Cells(EV_LIGNE_ENTETE, 2 + nbP)).NumberFormat = "@"
    End If
    For c = 1 To nbP
        idxDebut = (deb \ pas + c - 1) * pas
        Select Case pas
            Case 1
                ws.Cells(EV_LIGNE_ENTETE, 2 + c).Value = DateSerial(mod_AnalyseDonnees.AnneeDeIndice(idxDebut), _
                                                                    mod_AnalyseDonnees.MoisDeIndice(idxDebut), 1)
            Case 3
                ws.Cells(EV_LIGNE_ENTETE, 2 + c).Value = "T" & ((mod_AnalyseDonnees.MoisDeIndice(idxDebut) + 2) \ 3) & _
                                                         "-" & mod_AnalyseDonnees.AnneeDeIndice(idxDebut)
            Case Else
                ws.Cells(EV_LIGNE_ENTETE, 2 + c).Value = CStr(mod_AnalyseDonnees.AnneeDeIndice(idxDebut))
        End Select
    Next c
    ws.Cells(EV_LIGNE_ENTETE, 3 + nbP).Value = "Total"

    nbAff = 0
    totalGeneral = 0
    If nb = 0 Then Exit Sub

    ' Les EV_NB_TOP premières séries, puis "Autres".
    nbTop = nb
    If nbTop > EV_NB_TOP Then nbTop = EV_NB_TOP
    nbAff = nbTop
    If nb > nbTop Then nbAff = nbTop + 1

    ReDim tabDonnees(1 To nbAff, 1 To nbP + 2)
    ReDim totCol(1 To nbP)
    ReDim reste(1 To nbP)

    For i = 1 To nbTop
        tabDonnees(i, 1) = cles(i)
        totLigne = 0
        For c = 1 To nbP
            tabDonnees(i, 1 + c) = mat(c, i)
            totLigne = totLigne + mat(c, i)
            totCol(c) = totCol(c) + mat(c, i)
        Next c
        tabDonnees(i, nbP + 2) = totLigne
        totalGeneral = totalGeneral + totLigne
    Next i

    If nb > nbTop Then
        For i = nbTop + 1 To nb
            For c = 1 To nbP
                reste(c) = reste(c) + mat(c, i)
            Next c
        Next i
        tabDonnees(nbAff, 1) = mod_AnalyseDonnees.LibAutres()
        totLigne = 0
        For c = 1 To nbP
            tabDonnees(nbAff, 1 + c) = reste(c)
            totLigne = totLigne + reste(c)
            totCol(c) = totCol(c) + reste(c)
        Next c
        tabDonnees(nbAff, nbP + 2) = totLigne
        totalGeneral = totalGeneral + totLigne
    End If

    ws.Range(ws.Cells(EV_LIGNE_ENTETE + 1, 2), ws.Cells(EV_LIGNE_ENTETE + nbAff, 3 + nbP)).Value = tabDonnees

    ws.Cells(EV_LIGNE_TOTAL, 2).Value = "Total"
    For c = 1 To nbP
        ws.Cells(EV_LIGNE_TOTAL, 2 + c).Value = totCol(c)
    Next c
    ws.Cells(EV_LIGNE_TOTAL, 3 + nbP).Value = totalGeneral

End Sub

Private Sub MettreAJourGraphique(ByVal ws As Worksheet, ByVal nbP As Long, ByVal nbAff As Long, _
                                 ByVal graphe As Long, ByVal titre As String)

    Dim co As ChartObject
    Dim s As Series
    Dim i As Long, r As Long
    Dim typeGraphe As Long

    Set co = Nothing
    On Error Resume Next
    Set co = ws.ChartObjects(EV_GRAPH_NOM)
    On Error GoTo 0
    If co Is Nothing Then Exit Sub

    Do While co.Chart.SeriesCollection.Count > 0
        co.Chart.SeriesCollection(1).Delete
    Loop

    co.Chart.HasTitle = True
    If nbAff = 0 Then
        co.Chart.ChartTitle.Text = mod_Display.FR("Aucune d{e2}pense sur la p{e2}riode")
        Exit Sub
    End If

    Select Case graphe
        Case 2: typeGraphe = xlLineMarkers
        Case 3: typeGraphe = xlColumnClustered
        Case Else: typeGraphe = xlColumnStacked
    End Select
    co.Chart.ChartType = typeGraphe

    For i = 1 To nbAff
        r = EV_LIGNE_ENTETE + i
        Set s = co.Chart.SeriesCollection.NewSeries
        s.Name = CStr(ws.Cells(r, 2).Value)
        s.Values = ws.Range(ws.Cells(r, 3), ws.Cells(r, 2 + nbP))
        s.XValues = ws.Range(ws.Cells(EV_LIGNE_ENTETE, 3), ws.Cells(EV_LIGNE_ENTETE, 2 + nbP))
        s.ChartType = typeGraphe
    Next i

    On Error Resume Next
    co.Chart.Axes(xlCategory).CategoryType = xlCategoryScale   ' une colonne par période, sans échelle de dates
    co.Chart.Axes(xlValue).TickLabels.NumberFormat = "#,##0"
    On Error GoTo 0

    co.Chart.HasLegend = True
    co.Chart.Legend.Position = xlLegendPositionBottom
    co.Chart.ChartTitle.Text = titre

End Sub


' =====================================================================================
' RAFRAÎCHISSEMENT DE L'ÉCRAN
' =====================================================================================
Public Sub ActualiserEvolution()

    On Error GoTo Erreur
    Application.ScreenUpdating = False
    RafraichirCoeur
    Application.ScreenUpdating = True
    Exit Sub

Erreur:
    Application.ScreenUpdating = True
    MsgBox mod_Display.FR("Erreur pendant le calcul de l'{e2}volution.") & vbLf & _
           "Erreur " & Err.Number & " : " & Err.Description, vbExclamation, mod_Display.FR("{E2}volution des d{e2}penses")

End Sub

Private Sub RafraichirCoeur()

    Dim ws As Worksheet
    Dim deb As Long, fin As Long, pas As Long, nbP As Long
    Dim cats As Object, subs As Object, excl As Object
    Dim modeSous As Boolean, tronque As Boolean
    Dim cles() As String, mat() As Double, nb As Long
    Dim nbAff As Long
    Dim totalGeneral As Double
    Dim scope As String, titre As String, msg As String
    Dim col As Variant

    Set ws = FeuilleEvolution()
    If ws Is Nothing Then
        MsgBox mod_Display.FR("La feuille d'{e2}volution n'existe pas : lancez d'abord InstallerEvolution."), vbExclamation
        Exit Sub
    End If

    If Not mod_AnalyseDonnees.ChargerLignes() Then Exit Sub
    MettreAJourListes ws

    If Not CalculerPeriode(Texte(ws, EV_ADR_PERIODE), deb, fin) Then Exit Sub
    pas = CalculerPas(Texte(ws, EV_ADR_PAS), deb, fin)
    nbP = NbPas(deb, fin, pas)
    If nbP > EV_NB_PAS_MAX Then
        ' Trop de colonnes : on garde les plus récentes.
        deb = (fin \ pas - EV_NB_PAS_MAX + 1) * pas
        nbP = EV_NB_PAS_MAX
        tronque = True
    End If

    Set cats = LireSelection(ws, EV_COL_CAT)
    Set subs = LireSelection(ws, EV_COL_SOUS)
    Set excl = LireExclusions(ws)

    Select Case IndexLib(LibsDetail(), Texte(ws, EV_ADR_DETAIL))
        Case 2
            modeSous = False
        Case 3
            modeSous = True
        Case Else
            modeSous = (cats.Count = 1 Or subs.Count > 0)
    End Select

    Agreger deb, fin, pas, nbP, cats, subs, excl, modeSous, cles, mat, nb
    EcrireTableau ws, deb, pas, nbP, modeSous, cles, mat, nb, nbAff, totalGeneral

    ' Phrase de portée : sert au message et au titre du graphique.
    If modeSous Then
        If cats.Count = 1 Then
            For Each col In cats.Keys
                scope = mod_Display.FR("de la cat{e2}gorie ") & CStr(col) & mod_Display.FR(", par sous-cat{e2}gorie")
            Next col
        Else
            scope = mod_Display.FR("par sous-cat{e2}gorie")
        End If
    Else
        scope = mod_Display.FR("par cat{e2}gorie")
    End If

    msg = mod_Display.FR("D{e2}penses nettes ") & scope & " -- " & DecrirePeriode(deb, fin) & _
          " (par " & NomDuPas(pas) & ") : " & Format(totalGeneral, "#,##0.00") & " " & ChrW(8364) & "."
    If tronque Then
        msg = msg & mod_Display.FR(" Affichage limit{e2} aux ") & EV_NB_PAS_MAX & mod_Display.FR(" colonnes les plus r{e2}centes : choisissez un pas plus large pour tout voir.")
    End If
    ws.Range(EV_ADR_MESSAGE).Value = msg

    titre = mod_Display.FR("{E2}volution des d{e2}penses ") & scope & " -- " & DecrirePeriode(deb, fin)
    MettreAJourGraphique ws, nbP, nbAff, IndexLib(LibsGraph(), Texte(ws, EV_ADR_GRAPH)), titre

    On Error Resume Next
    ws.Range(EV_ADR_PERIODE).Select     ' (sans effet si la feuille n'est pas affichée)
    On Error GoTo 0

End Sub


' =====================================================================================
' BOUTONS
' =====================================================================================

' Ouverture depuis la feuille Synthese. L'écran est autonome : aucun paramètre n'est repris
' d'un autre écran.
Public Sub AfficherEvolution()

    Dim ws As Worksheet, wsSyn As Worksheet

    Set ws = FeuilleEvolution()
    If ws Is Nothing Then
        MsgBox mod_Display.FR("La feuille d'{e2}volution n'existe pas encore : lancez d'abord InstallerEvolution (Ctrl+G)."), vbExclamation
        Exit Sub
    End If

    Set wsSyn = mod_Criteres.GetFeuille(mod_VarGlobales.NOM_FEUILLE_MENU)

    ws.Visible = xlSheetVisible
    ws.Activate
    mod_InstallCommun.MasquerQuadrillage      ' masque aussi Synthese (regle commune des formulaires)

    ActualiserEvolution

End Sub

Public Sub SortirEvolution()
    Dim ws As Worksheet, wsSyn As Worksheet
    Set ws = FeuilleEvolution()
    mod_InstallCommun.AfficherSynthese
    If Not ws Is Nothing Then mod_Aide.MasquerFormulaire ws
End Sub

' Bouton "Réinitialiser les sélections" : vide les listes de catégories et de sous-catégories.
Public Sub ReinitialiserSelectionsEvolution()

    Dim ws As Worksheet
    Dim l As Long

    Set ws = FeuilleEvolution()
    If ws Is Nothing Then Exit Sub

    For l = EV_LIGNE_SEL_DEB To EV_LIGNE_SEL_FIN
        ws.Cells(l, EV_COL_CAT).MergeArea.ClearContents
        ws.Cells(l, EV_COL_SOUS).MergeArea.ClearContents
    Next l

    ActualiserEvolution

End Sub
