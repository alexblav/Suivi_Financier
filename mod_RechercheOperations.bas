Option Explicit

' =====================================================================================
' MODULE : mod_RechercheOperations
'
' RÔLE (phase 5b) : logique de la feuille frm_RechercheOperations construite
' en phase 5a (mod_InstallRechercheOperations).
'
'   RechercherOperations   : recharge le tableau de recherche depuis TblOperations
'                            ET, PHASE 5, depuis TblVentilations (filtre ensuite
'                            avec les flèches natives d'Excel dans l'en-tête du
'                            tableau).
'   AppliquerLignesMarquees : pour chaque ligne du tableau de recherche ou
'                            "Valider" = "Oui", écrit Catégorie/SousCategorie et
'                            Notes dans la table SOURCE de la ligne (TblOperations
'                            ou TblVentilations -- voir colonne technique
'                            SourceLigne), puis efface la marque.
'
'   Verrouillage "souple" de Notes : il n'y a pas de vraie protection de
'   feuille (déjà abandonnée ailleurs dans ce chantier à cause d'une erreur
'   1004). La cellule Notes d'une ligne dont la clé est déjà valide est
'   simplement grisée (indication visuelle), et AppliquerLignesMarquees
'   IGNORE tout changement sur cette colonne pour cette ligne, quel que soit
'   ce qui y est écrit à l'écran.
'
' PHASE 5 -- CE QUI CHANGE PAR RAPPORT À LA VERSION D'ORIGINE :
'   (a) BUG CORRIGÉ : la liste déroulante de la colonne Catégorie utilisait
'       Formula1:=Join(listeCategories, ","). Cela échoue dès que : (1) une
'       catégorie contient une virgule (ex. : "Alimentation, supermarché"),
'       qu'Excel scinde alors en plusieurs "fausses" catégories dans la liste;
'       (2) le texte assemblé dépasse 255 caractères, limite vite atteinte avec
'       environ 80 catégories (limite d'Excel pour un Formula1 codé en dur).
'       La correction reprend la technique utilisée dans le reste du chantier
'       (mod_Categories, mod_ControleCategories...) : une PLAGE NOMMÉE dynamique
'       (OFFSET/COUNTA) pointant vers une colonne de VRAIES cellules, et non une
'       liste de texte concaténée.
'   (b) NOUVELLE colonne SousCategorie (même traitement que Catégorie).
'   (c) NOUVELLE colonne Ventile (tag informatif) + intégration des lignes de
'       TblVentilations : une opération ventilée reste désormais accessible
'       ICI de deux façons, comme demandé par l'opérateur : via sa ligne
'       PARENTE (Catégorie = "Ventile"), ou directement via chacune de ses
'       parts (une ligne par sous-catégorie de la ventilation, marquée
'       Ventile = "Oui").
'
' À PROPOS DES ACCENTS : les textes affichés passent par mod_Display.FR(); les
' commentaires du fichier sont encodés en UTF-8.
'
' COMMENT TESTER : Ctrl+G, taper RechercherOperations, puis Entrée.
' =====================================================================================

' --- Noms techniques utilisés par ce module, DUPLIQUÉS ICI EN DUR (comme dans
'     les autres modules du chantier), afin qu'il puisse compiler même si
'     mod_InstallVentilation n'a pas encore été importé. ----------------------
Private Const VEN_FEUILLE As String = "Ventilations"
Private Const VEN_TABLE As String = "TblVentilations"
Private Const CATEGORIE_VENTILE As String = "Ventil" ' + « é » accentué, voir CategorieVentileRO()

' Mémoire du préfiltre/paramètre actuellement affichés : sert uniquement à ce que
' RevoirVentilationRO (plus bas) puisse recharger l'écran dans le MÊME contexte
' après une modification, plutôt que de retomber en recherche libre.
Private g_ROPrefiltreActif As String
Private g_ROParamActif As Variant

' Sous-catégorie "santé" : depuis le 07/10/2026 (préfiltre "SuiviSante"), ce module
' la teste directement, toujours via la constante centralisée
' mod_VarGlobales.SOUS_CATEGORIE_SANTE (jamais un texte recopié en dur ici).

Private Function CategorieVentileRO() As String
    CategorieVentileRO = CATEGORIE_VENTILE & ChrW(233)   ' "Ventile"
End Function

' =====================================================================================
' PrefiltreActifRO (ajout 07/10/2026) : renvoie le préfiltre actuellement affiché
' ("", "DernierImport", "SuiviSante"...).
' La variable g_ROPrefiltreActif reste Private : personne ne doit pouvoir la
' MODIFIER depuis un autre module. Cette fonction en donne seulement une LECTURE,
' utilisée par ThisWorkbook.Workbook_SheetBeforeDoubleClick pour n'activer le
' double-clic sur SousCategorie qu'en mode suivi santé.
' Si le projet VBA a été réinitialisé entre-temps (bouton "Sortir" qui exécute
' End, erreur non gérée...), les variables sont effacées et la fonction renvoie "" :
' le double-clic reprend alors simplement son comportement normal d'Excel.
' =====================================================================================
Public Function PrefiltreActifRO() As String
    PrefiltreActifRO = g_ROPrefiltreActif
End Function

' =====================================================================================
' HELPERS PHASE 5 : accès à TblVentilations, dupliqués comme dans les autres
' modules du chantier (mod_SuiviSante, mod_FormulairesNotes, ...).
' =====================================================================================
Private Function ObtenirTableVentilationsRO() As ListObject
    Dim ws As Worksheet
    Dim t As ListObject
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(VEN_FEUILLE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    Set t = ws.ListObjects(VEN_TABLE)
    On Error GoTo 0
    Set ObtenirTableVentilationsRO = t
End Function

Private Function IndexColRO(ByVal t As ListObject, ByVal nomColonne As String) As Long
    On Error Resume Next
    IndexColRO = t.ListColumns(nomColonne).index
    On Error GoTo 0
End Function

' =====================================================================================
' Mémorise les filtres de colonne actifs (flèches natives d'Excel) sur tblRecherche
' afin de pouvoir les restaurer après avoir vidé et réécrit le tableau. Sinon,
' l'opérateur perd son filtre à chaque recherche/actualisation (constat du 01/10/2026 :
' après "Revoir la ventilation", le filtre "Oui" de la colonne "Ventile" semblait
' disparaître, car Excel conservait les anciennes positions des lignes visibles,
' qui ne correspondaient plus à rien après la réécriture du tableau).
' =====================================================================================
Private Function MemoriserFiltresRO(ByVal tbl As ListObject) As Variant
    ' Renvoie un tableau à deux dimensions (4 lignes : Field/Operator/Criteria1/Criteria2,
    ' une colonne par filtre actif trouvé), ou Empty si aucun filtre n'est actif.
    Dim nbFiltres As Long
    Dim resultatFiltres() As Variant
    Dim f As Long
    Dim filtreCol As Filter

    nbFiltres = 0
    If tbl.ShowAutoFilter Then
        For f = 1 To tbl.ListColumns.count
            On Error Resume Next
            Set filtreCol = tbl.AutoFilter.Filters(f)
            On Error GoTo 0
            If Not filtreCol Is Nothing Then
                If filtreCol.On Then
                    nbFiltres = nbFiltres + 1
                    ReDim Preserve resultatFiltres(1 To 4, 1 To nbFiltres)
                    resultatFiltres(1, nbFiltres) = f
                    resultatFiltres(2, nbFiltres) = filtreCol.Operator
                    resultatFiltres(3, nbFiltres) = filtreCol.Criteria1
                    ' Criteria2 ne s'applique qu'à certains filtres (ex. : "entre telle
                    ' et telle date"); il est absent dans les autres cas, d'où On Error.
                    resultatFiltres(4, nbFiltres) = Empty
                    On Error Resume Next
                    resultatFiltres(4, nbFiltres) = filtreCol.Criteria2
                    On Error GoTo 0
                End If
            End If
            Set filtreCol = Nothing
        Next f
    End If

    If nbFiltres = 0 Then
        MemoriserFiltresRO = Empty
    Else
        MemoriserFiltresRO = resultatFiltres
    End If
End Function

' Réapplique les filtres précédemment mémorisés par MemoriserFiltresRO, une fois le
' tableau reconstruit avec les nouvelles données.
Private Sub RestaurerFiltresRO(ByVal tbl As ListObject, ByVal filtresSauvegardes As Variant)
    Dim f As Long
    Dim critere1 As Variant
    Dim operateur As Long

    If IsEmpty(filtresSauvegardes) Then Exit Sub

    On Error Resume Next
    For f = 1 To UBound(filtresSauvegardes, 2)
        critere1 = filtresSauvegardes(3, f)
        operateur = filtresSauvegardes(2, f)

        If operateur = xlFilterValues Then
            If Not IsArray(critere1) Then critere1 = Array(critere1)
        End If

        If operateur = 0 Then
            ' Filtre "simple" (ex. : "non vide", critere1 = "<>") : Excel ne renvoie
            ' aucun Operator exploitable (0 n'est pas une valeur valide de
            ' XlAutoFilterOperator). On l'omet donc lors de la réapplication, sinon
            ' Excel refuse silencieusement le filtre (constat du 01/10/2026).
            tbl.Range.AutoFilter Field:=filtresSauvegardes(1, f), Criteria1:=critere1
        ElseIf IsEmpty(filtresSauvegardes(4, f)) Then
            tbl.Range.AutoFilter Field:=filtresSauvegardes(1, f), _
                                  Criteria1:=critere1, _
                                  Operator:=operateur
        Else
            tbl.Range.AutoFilter Field:=filtresSauvegardes(1, f), _
                                  Criteria1:=critere1, _
                                  Operator:=operateur, _
                                  Criteria2:=filtresSauvegardes(4, f)
        End If
    Next f
    On Error GoTo 0
End Sub

' =====================================================================================
' RechercherOperations : recharge le tableau de recherche depuis TblOperations
' ET TblVentilations.
'
' PHASE 6 (refonte de l'écran central, après discussion avec l'opérateur) :
'   Ajout d'un paramètre facultatif "prefiltre", qui remplace les anciens
'   écrans "Synthese_*" (désormais supprimés) :
'     ""                : recherche libre, comportement D'ORIGINE inchangé
'     "DernierImport"   : uniquement les opérations du DERNIER import (liste
'                         mémorisée par mod_DernierImport.MemoriserDernierImport)
'     "OperationsDuMois" : uniquement les opérations du mois et de l'année choisis
'                         sur Synthese (remplace mod_SyntheseBugetMensuel)
'     "ErreursSante"    : uniquement les lignes dont la consultation santé est en
'                         erreur (remplace mod_SyntheseExportCareError)
'     "DetailTotal"     : détail d'un total du bilan mensuel; "param" vaut alors
'                         "Positif" ou "Negatif" (remplace
'                         mod_SyntheseBudgetBilanMensuel.ShowDetailForTotal)
'     "SuiviSante"      : uniquement les lignes de sous-catégorie « Frais, remb
'                         santé », parts ventilées comprises (remplace l'ancien
'                         écran Mod_SyntheseCare.Synthese_Care, refonte du 07/10/2026)
'   Dans tous les cas, le tableau reste le MÊME (mêmes colonnes, même bouton
'   Appliquer); seules les LIGNES chargées et les COLONNES visibles changent
'   (voir mod_InstallRechercheOperations.DefinirColonnesVisibles).
' =====================================================================================
Public Sub RechercherOperations(Optional ByVal prefiltre As String = "", Optional ByVal param As Variant)

    Dim ws As Worksheet
    Dim tblRecherche As ListObject
    Dim n As Long, i As Long, k As Long, ligneEcran As Long

    Dim resultat() As Variant
    Dim cleVerrouillee() As Boolean
    Dim notesTexte As String, seg0 As String
    Dim catTexte As String, sousTexte As String

    Dim valeursVues As Object
    Dim listeValeurs() As String
    Dim nbValeurs As Long

    ' --- PHASE 5 : accès à TblVentilations (facultatif) ---
    Dim tblVen As ListObject
    Dim tblDataVen As Variant
    Dim nbLignesVen As Long
    Dim venDisponible As Boolean
    Dim colVenID As Long, colVenCat As Long, colVenSousCat As Long, colVenMontant As Long, colVenNotes As Long
    ' Ajout 07/10/2026 (préfiltre "SuiviSante") : colonnes de suivi santé de
    ' TblVentilations. Elles valent 0 si la colonne n'existe pas (installation
    ' ancienne) : la valeur correspondante reste alors simplement vide à l'écran.
    Dim colVenStatut As Long, colVenSolde As Long, colVenDateConsult As Long, colVenSpeConsult As Long

    ' --- PHASE 6 : préparation spécifique au préfiltre demandé ---
    Dim inclureVentilations As Boolean
    Dim listeDernierImport As Object
    Dim showTypeDetail As String
    Dim indicesOp() As Long, nbIndicesOp As Long
    Dim indicesVen() As Long, nbIndicesVen As Long
    Dim nbTotal As Long

    ' Mémorise le préfiltre actif : RevoirVentilationRO le réutilise pour
    ' recharger l'écran dans le MÊME contexte après une modification.
    g_ROPrefiltreActif = prefiltre
    g_ROParamActif = param

    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_RECHERCHE)
    Set tblRecherche = ws.ListObjects(mod_VarGlobales.NOM_TABLE_RECHERCHE)

    Set tbl = mod_DonneesTable.GetOperationsTable()
    If tbl Is Nothing Then
        MsgBox mod_Display.FR("Le tableau TblOperations est introuvable."), vbExclamation
        Exit Sub
    End If
    If tbl.DataBodyRange Is Nothing Then
        MsgBox mod_Display.FR("TblOperations ne contient aucune ligne."), vbInformation
        Exit Sub
    End If

    mod_Display.RecupIndexCol
    tblData = tbl.DataBodyRange.value
    n = UBound(tblData, 1)

    ' --- PHASE 6 : préparation spécifique au préfiltre -------------------
    showTypeDetail = ""
    Select Case prefiltre
        Case "DernierImport"
            Set listeDernierImport = ChargerListeDernierImportRO()
            If listeDernierImport Is Nothing Then
                MsgBox mod_Display.FR("Aucun import n'a encore {e2}t{e2} enregistr{e2}."), vbInformation
                Exit Sub
            End If
        Case "OperationsDuMois"
            mod_Criteres.GetSelectCriteres
        Case "DetailTotal"
            showTypeDetail = CStr(param)
            mod_Criteres.GetSelectCriteres
    End Select

    ' Les lignes de TblVentilations n'ont pas de mois/année budgétaire ni de
    ' colonnes de suivi santé qui leur soient propres : elles ne sont pertinentes
    ' qu'en recherche libre ou pour le dernier import (comme avant cette refonte).
    ' Ajout 07/10/2026 : le préfiltre "SuiviSante" inclut lui aussi les parts
    ' ventilées (décision opérateur : pour le suivi santé, une part ventilée est
    ' une opération comme une autre).
    inclureVentilations = (prefiltre = "" Or prefiltre = "DernierImport" Or prefiltre = "SuiviSante")

    ' --- PHASE 5 : chargement de TblVentilations, si pertinent ---
    venDisponible = False
    If inclureVentilations Then
        Set tblVen = ObtenirTableVentilationsRO()
        If Not tblVen Is Nothing Then
            If Not tblVen.DataBodyRange Is Nothing Then
                colVenID = IndexColRO(tblVen, "ID_Transaction")
                colVenCat = IndexColRO(tblVen, "Categorie")
                colVenSousCat = IndexColRO(tblVen, "SousCategorie")
                colVenMontant = IndexColRO(tblVen, "Montant")
                colVenNotes = IndexColRO(tblVen, "Notes")
                ' Ajout 07/10/2026 : colonnes santé, FACULTATIVES (0 si absentes).
                ' Elles ne sont donc volontairement PAS ajoutées au test
                ' "toutes présentes" juste en dessous : leur absence ne doit pas
                ' empêcher d'afficher les parts ventilées.
                colVenStatut = IndexColRO(tblVen, "StatutSante")
                colVenSolde = IndexColRO(tblVen, "SoldeSante")
                colVenDateConsult = IndexColRO(tblVen, "Date_consult")
                colVenSpeConsult = IndexColRO(tblVen, "Spe_Consult")
                If colVenID <> 0 And colVenCat <> 0 And colVenSousCat <> 0 And colVenMontant <> 0 And colVenNotes <> 0 Then
                    tblDataVen = tblVen.DataBodyRange.value
                    nbLignesVen = UBound(tblDataVen, 1)
                    venDisponible = True
                End If
            End If
        End If
    End If

    ws.Visible = xlSheetVisible
    ws.Activate

    ' Fige l'affichage sous la ligne d'en-tête du tableau (RO_LIGNE_ENTETES),
    ' pour que les intitulés restent visibles pendant le défilement des résultats
    ' (demande opérateur du 01/10/2026).
    ' CORRECTIF du 05/10/2026 (constat opérateur) : on replace D'ABORD la fenêtre
    ' en haut à gauche (ScrollRow/ScrollColumn = 1) AVANT de figer les volets.
    ' Sinon, si la feuille avait déjà défilé, même légèrement, le gel se faisait
    ' depuis la position actuelle et non depuis la ligne 1, donnant l'impression
    ' que les premières lignes étaient masquées. Le problème passait inaperçu avec
    ' RO_LIGNE_ENTETES=5 (écart trop petit), mais devenait visible avec RO_LIGNE_ENTETES=7.
    ' CORRECTIF du 07/10/2026 (constat opérateur, intermittent) : on force aussi
    ' Split à False avant de regeler. FreezePanes et Split sont deux états
    ' distincts dans Excel ; si un fractionnement (Split) restait actif en
    ' mémoire sur la fenêtre (par exemple si elle n'avait pas fini de se
    ' redessiner juste après ws.Activate), remettre FreezePanes à True pouvait
    ' geler les volets à l'ANCIENNE position du split plutôt qu'à la cellule
    ' sélectionnée juste après : la barre de séparation restait visible, mais
    ' l'en-tête n'était pas forcément dans la zone figée. D'où le caractère
    ' intermittent du symptôme (ça dépendait de l'état de la fenêtre à l'instant T).
    ActiveWindow.FreezePanes = False
    ActiveWindow.Split = False
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    ws.Range("A" & (mod_InstallRechercheOperations.RO_LIGNE_ENTETES + 1)).Select
    ActiveWindow.FreezePanes = True

    ' Avant de vider et reconstruire le tableau, on mémorise les filtres actifs
    ' (flèches natives d'Excel) afin de les rétablir à l'identique après l'écriture
    ' des nouvelles données (voir MemoriserFiltresRO / RestaurerFiltresRO plus haut).
    Dim filtresSauvegardes As Variant
    filtresSauvegardes = MemoriserFiltresRO(tblRecherche)

    On Error Resume Next
    If tblRecherche.ShowAutoFilter Then
        tblRecherche.AutoFilter.ShowAllData
    End If
    On Error GoTo 0

    ' --- On vide le tableau de recherche (seul l'en-tête est conservé) ---
    If Not tblRecherche.DataBodyRange Is Nothing Then
        tblRecherche.DataBodyRange.Delete
        With tblRecherche.Sort
            .SortFields.Clear
            .SortFields.Add2 Key:=tblRecherche.ListColumns(2).Range, _
                     SortOn:=xlSortOnValues, _
                     Order:=xlDescending
            .Header = xlYes
            .Apply
        End With
    End If

    ' --- PHASE 6 : on repère d'abord les lignes retenues par le préfiltre
    ' dans deux tableaux d'indices, redimensionnés au fil de l'eau, avant de
    ' dimensionner resultat à la bonne taille. C'est plus simple et plus sûr
    ' que ReDim Preserve sur la première dimension d'un tableau 2D, impossible
    ' en VBA. ----------------------------------------------------------------
    nbIndicesOp = 0
    For i = 1 To n
        If LigneOpRetenuePourPrefiltre(i, prefiltre, showTypeDetail, listeDernierImport) Then
            nbIndicesOp = nbIndicesOp + 1
            ReDim Preserve indicesOp(1 To nbIndicesOp)
            indicesOp(nbIndicesOp) = i
        End If
    Next i

    nbIndicesVen = 0
    If venDisponible Then
        For i = 1 To nbLignesVen
            If prefiltre = "" Then
                nbIndicesVen = nbIndicesVen + 1
                ReDim Preserve indicesVen(1 To nbIndicesVen)
                indicesVen(nbIndicesVen) = i
            ElseIf prefiltre = "DernierImport" Then
                If listeDernierImport.Exists(mod_DataStructure.CellText(tblDataVen(i, colVenID))) Then
                    nbIndicesVen = nbIndicesVen + 1
                    ReDim Preserve indicesVen(1 To nbIndicesVen)
                    indicesVen(nbIndicesVen) = i
                End If
            ElseIf prefiltre = "SuiviSante" Then
                ' Ajout 07/10/2026 : seules les parts affectées à la sous-catégorie
                ' santé (même test que pour TblOperations, même constante centralisée).
                If mod_DataStructure.CellText(tblDataVen(i, colVenSousCat)) = mod_VarGlobales.SOUS_CATEGORIE_SANTE Then
                    nbIndicesVen = nbIndicesVen + 1
                    ReDim Preserve indicesVen(1 To nbIndicesVen)
                    indicesVen(nbIndicesVen) = i
                End If            
            End If
        Next i
    End If

    nbTotal = nbIndicesOp + nbIndicesVen

    If nbTotal = 0 Then
        MsgBox mod_Display.FR("Aucune op{e2}ration {a2} afficher pour ce filtre."), vbInformation
        mod_InstallRechercheOperations.DefinirColonnesVisibles ws, prefiltre
        Exit Sub
    End If

    ReDim resultat(1 To nbTotal, 1 To 16)
    ReDim cleVerrouillee(1 To nbTotal)

    Set valeursVues = CreateObject("Scripting.Dictionary")
    valeursVues.CompareMode = 1
    nbValeurs = 0

    ligneEcran = 0

    ' =====================================================================
    ' 1) Lignes de TblOperations retenues par le préfiltre.
    ' =====================================================================
    For k = 1 To nbIndicesOp
        i = indicesOp(k)
        ligneEcran = ligneEcran + 1

        resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_VALIDER) = ""
        resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_DATE) = tblData(i, colDate)
        resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_TIERS) = mod_DataStructure.CellText(tblData(i, colTiers))
        resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_MONTANT) = tblData(i, colMontant)

        catTexte = mod_DataStructure.CellText(tblData(i, colCategorie))
        resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_CATEGORIE) = catTexte

        sousTexte = ""
        If colSousCategorie <> 0 Then sousTexte = mod_DataStructure.CellText(tblData(i, colSousCategorie))
        resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_SOUSCATEGORIE) = sousTexte

        notesTexte = mod_DataStructure.CellText(tblData(i, colNotes))
        resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_NOTES) = notesTexte

        resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_VENTILE) = _
            IIf(catTexte = CategorieVentileRO(), "Oui", "")

        resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_ID) = mod_DataStructure.CellText(tblData(i, colID))
        resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_SOURCE) = "O"
        resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_LIGNEVEN) = ""

        ' PHASE 6 : colonnes issues des anciens écrans "Synthese_*", toujours
        ' renseignées (même si masquées pour ce préfiltre). C'est sans conséquence
        ' et évite de relire TblOperations si l'opérateur change les colonnes
        ' visibles sans relancer la recherche.
        If colBudget <> 0 Then resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_BUDGET) = tblData(i, colBudget)
        If colStatutSante <> 0 Then resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_STATUTSANTE) = mod_DataStructure.CellText(tblData(i, colStatutSante))
        If colSoldeSante <> 0 Then resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_SOLDESANTE) = tblData(i, colSoldeSante)
        If colDateConsult <> 0 Then resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_DATECONSULT) = tblData(i, colDateConsult)
        If colSpeConsult <> 0 Then resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_SPECONSULT) = mod_DataStructure.CellText(tblData(i, colSpeConsult))

        seg0 = mod_ImportOFX.SegmentTexte(notesTexte, ";", 0)
        cleVerrouillee(ligneEcran) = mod_ImportOFX.EstDateValide(seg0)

        AjouterValeurUnique catTexte, valeursVues, listeValeurs, nbValeurs
        AjouterValeurUnique sousTexte, valeursVues, listeValeurs, nbValeurs
    Next k

    ' =====================================================================
    ' 2) Lignes de TblVentilations retenues (recherche libre ou dernier import
    '    uniquement; voir "inclureVentilations" plus haut).
    ' =====================================================================
    If nbIndicesVen > 0 Then
        ' Index ID_Transaction -> ligne TblOperations, pour retrouver le Tiers
        ' de l'opération parente sans reparcourir tblData pour chaque part ventilée.
        Dim indexParIDPourTiers As Object
        Set indexParIDPourTiers = CreateObject("Scripting.Dictionary")
        indexParIDPourTiers.CompareMode = 1
        For k = 1 To n
            indexParIDPourTiers(mod_DataStructure.CellText(tblData(k, colID))) = k
        Next k

        Dim idParentVen As String, tiersParentVen As String
        For k = 1 To nbIndicesVen
            i = indicesVen(k)
            ligneEcran = ligneEcran + 1

            idParentVen = mod_DataStructure.CellText(tblDataVen(i, colVenID))
            tiersParentVen = ""
            If indexParIDPourTiers.Exists(idParentVen) Then
                tiersParentVen = mod_DataStructure.CellText(tblData(indexParIDPourTiers(idParentVen), colTiers))
            End If

            resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_VALIDER) = ""
            ' Une part ventilée n'a pas de date propre : la date reste vide pour
            ' les autres préfiltres (convention d'origine). Ajout 07/10/2026 (demande
            ' opérateur) : pour le préfiltre "SuiviSante" UNIQUEMENT, on reprend la
            ' date de l'opération PARENTE, pour qu'une dépense ventilée se lise et se
            ' trie comme une dépense ordinaire. indexParIDPourTiers (construit juste
            ' au-dessus pour le Tiers) donne déjà la ligne parente : on le réutilise.
            resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_DATE) = Empty
            If prefiltre = "SuiviSante" Then
                If indexParIDPourTiers.Exists(idParentVen) Then
                    resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_DATE) = _
                        tblData(indexParIDPourTiers(idParentVen), colDate)
                End If
            End If
            resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_TIERS) = tiersParentVen & mod_Display.FR(" [ventilation]")
            resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_MONTANT) = tblDataVen(i, colVenMontant)

            catTexte = mod_DataStructure.CellText(tblDataVen(i, colVenCat))
            sousTexte = mod_DataStructure.CellText(tblDataVen(i, colVenSousCat))
            resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_CATEGORIE) = catTexte
            resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_SOUSCATEGORIE) = sousTexte

            notesTexte = mod_DataStructure.CellText(tblDataVen(i, colVenNotes))
            resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_NOTES) = notesTexte
            resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_VENTILE) = "Oui"

            resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_ID) = idParentVen
            resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_SOURCE) = "V"
            resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_LIGNEVEN) = i

            ' Ajout 07/10/2026 : en mode suivi santé, on affiche aussi les colonnes
            ' santé de la part ventilée, lues dans TblVentilations (qui porte les
            ' mêmes colonnes que TblOperations depuis l'extension de la Phase 4).
            ' Limité volontairement au préfiltre "SuiviSante" : les autres
            ' préfiltres gardent leur comportement actuel.
            If prefiltre = "SuiviSante" Then
                If colVenStatut <> 0 Then resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_STATUTSANTE) = mod_DataStructure.CellText(tblDataVen(i, colVenStatut))
                If colVenSolde <> 0 Then resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_SOLDESANTE) = tblDataVen(i, colVenSolde)
                If colVenDateConsult <> 0 Then resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_DATECONSULT) = tblDataVen(i, colVenDateConsult)
                If colVenSpeConsult <> 0 Then resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_SPECONSULT) = mod_DataStructure.CellText(tblDataVen(i, colVenSpeConsult))
            End If

            seg0 = mod_ImportOFX.SegmentTexte(notesTexte, ";", 0)
            cleVerrouillee(ligneEcran) = mod_ImportOFX.EstDateValide(seg0)

            AjouterValeurUnique catTexte, valeursVues, listeValeurs, nbValeurs
            AjouterValeurUnique sousTexte, valeursVues, listeValeurs, nbValeurs
        Next k
    End If

    ' --- Redimensionner le tableau, puis écrire les données en bloc ---
    tblRecherche.Resize tblRecherche.HeaderRowRange.Resize(nbTotal + 1, 16)
    tblRecherche.ListColumns("ID_Transaction").DataBodyRange.NumberFormat = "@"   ' <- ajout : AVANT l'écriture, sinon Excel convertit les longs ID en nombre
    tblRecherche.DataBodyRange.value = resultat

    tblRecherche.ListColumns("Date").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblRecherche.ListColumns("Montant").DataBodyRange.NumberFormat = "#,##0.00"
    tblRecherche.ListColumns("Budget").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblRecherche.ListColumns("Date_consult").DataBodyRange.NumberFormat = "dd/mm/yyyy"

    ' --- PHASE 5 (correction du bug) : liste déroulante Catégorie/SousCategorie
    '     via une VRAIE plage nommée (technique OFFSET/COUNTA), jamais via du texte
    '     concaténé avec des virgules; voir l'explication en tête de fichier.
    EcrireListeTechniqueRO ws, listeValeurs, nbValeurs

    With tblRecherche.ListColumns("Categorie").DataBodyRange.Validation
        .Delete
        If nbValeurs > 0 Then
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Formula1:="=ListeValeursRechercheOperations"
        End If
    End With
    With tblRecherche.ListColumns("SousCategorie").DataBodyRange.Validation
        .Delete
        If nbValeurs > 0 Then
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Formula1:="=ListeValeursRechercheOperations"
        End If
    End With

    ' --- Griser les cellules Notes déjà verrouillées (clé santé valide) et colorer
    ' en vert les montants positifs (PHASE 6 : demande généralisée à TOUS les
    ' préfiltres, et non plus au seul ancien écran "Budget mensuel"). --------------
    For i = 1 To nbTotal
        If cleVerrouillee(i) Then
            tblRecherche.ListColumns("Notes").DataBodyRange.Cells(i).Interior.Color = RGB(240, 240, 240)
        Else
            tblRecherche.ListColumns("Notes").DataBodyRange.Cells(i).Interior.ColorIndex = xlColorIndexNone
        End If

        tblRecherche.ListColumns("Montant").DataBodyRange.Cells(i).Font.Color = RGB(0, 0, 0)

        ' Ajout 07/10/2026 : on retire aussi le GRAS, posé sur les montants "KO"
        ' par le préfiltre "SuiviSante" (bloc suivant). Sans cette ligne, le gras
        ' pourrait rester visible si l'opérateur relance ensuite un autre préfiltre.
        tblRecherche.ListColumns("Montant").DataBodyRange.Cells(i).Font.Bold = False

        If mod_DataStructure.ToDouble(tblRecherche.ListColumns("Montant").DataBodyRange.Cells(i).value) > 0 Then
            tblRecherche.ListColumns("Montant").DataBodyRange.Cells(i).Font.Color = RGB(0, 128, 0)
        End If
    Next i

    ' --- Ajout 07/10/2026 (préfiltre "SuiviSante", demande opérateur) : reprise de
    ' l'atout visuel de l'ancien écran Synthese_Care. Le montant des lignes dont
    ' StatutSante vaut "KO" passe en ROUGE et GRAS. Ce bloc est placé APRÈS la
    ' boucle ci-dessus (qui remet tout en noir puis colore en vert les montants
    ' positifs) pour que le rouge l'emporte aussi sur un remboursement "KO".
    ' Le statut est relu dans le tableau mémoire "resultat", écrit dans le MÊME
    ' ordre que les lignes de l'écran : la ligne i de "resultat" correspond donc
    ' bien à la ligne i du tableau affiché.
    If prefiltre = "SuiviSante" Then
        For i = 1 To nbTotal
            If mod_DataStructure.CellText(resultat(i, mod_InstallRechercheOperations.RO_COL_STATUTSANTE)) = "KO" Then
                With tblRecherche.ListColumns("Montant").DataBodyRange.Cells(i).Font
                    .Color = RGB(255, 0, 0)
                    .Bold = True
                End With
            End If
        Next i
    End If

    ' --- PHASE 6 : surlignage des N plus grosses dépenses, UNIQUEMENT pour le
    ' préfiltre "OperationsDuMois" (comportement de l'ancien mod_SyntheseBugetMensuel;
    ' l'opérateur a choisi de ne PAS l'étendre aux autres écrans). Le code est repris
    ' ici avec une numérotation propre au tableau de recherche, plutôt que d'utiliser
    ' mod_Rapports.HighlightTopRows tel quel : cette procédure suppose une plage avec
    ' une ligne d'en-tête (convention des anciens écrans "Synthese_*"), contrairement
    ' à tblRecherche.DataBodyRange, qui n'en contient pas.
    If prefiltre = "OperationsDuMois" Then
        SurlignerTopDepensesRO tblRecherche, nbTotal
    End If

    MsgBox nbTotal & mod_Display.FR(" op{e2}ration(s)/part(s) charg{e2}e(s) (dont ") & nbIndicesVen & _
           mod_Display.FR(" ligne(s) de ventilation). Utilise les fl{e2}ches de filtre dans l'en-t{ea}te pour restreindre la liste."), _
           vbInformation

    mod_InstallRechercheOperations.DefinirColonnesVisibles ws, prefiltre

    ' --- Ajout 02/10/2026 : phrase de rappel du filtre actif, ecrite juste
    ' au-dessus du tableau (ligne RO_LIGNE_FILTRE) -- l'operateur sait ainsi en
    ' permanence sur quel sous-ensemble d'operations il travaille, meme apres
    ' avoir quitte puis rouvert cet ecran, ou apres un "Revoir la ventilation"
    ' qui relance cette meme procedure avec le meme prefiltre (g_ROPrefiltreActif).
    ws.Range("A" & mod_InstallRechercheOperations.RO_LIGNE_FILTRE).value = _
        DecrireFiltreActifRO(prefiltre, param, nbTotal)

    ' On remet en place le(s) filtre(s) que l'operateur avait poses avant cette
    ' recherche (voir RestaurerFiltresRO plus haut) -- sans ca, un filtre de colonne
    ' actif avant un "Revoir la ventilation" semblait disparaitre apres coup.
    RestaurerFiltresRO tblRecherche, filtresSauvegardes

End Sub

' =====================================================================================
' DecrireFiltreActifRO (ajout du 02/10/2026) : construit la phrase affichée à la
' ligne RO_LIGNE_FILTRE, juste au-dessus du tableau de recherche, afin que
' l'opérateur sache en permanence sur quel sous-ensemble d'opérations il travaille.
' Appelée UNIQUEMENT par RechercherOperations, après le remplissage du tableau
' (nbTotal est donc déjà connu).
'
' "critMois"/"critAnnee" sont des variables Public déclarées dans
' mod_VarGlobales et renseignées par mod_Criteres.GetSelectCriteres (déjà appelée
' plus haut dans RechercherOperations pour les préfiltres concernés). On les relit
' ici sans les recalculer, comme le fait mod_SyntheseBudgetBilanMensuel ailleurs
' dans ce classeur.
' =====================================================================================
Private Function DecrireFiltreActifRO(ByVal prefiltre As String, ByVal param As Variant, ByVal nbTotal As Long) As String

    Dim texte As String

    Select Case prefiltre

        Case ""
            texte = FR("Recherche globale (aucun filtre de mois ni de montant) -- ") & nbTotal & FR(" ligne(s) affich{e2}e(s).")

        Case "DernierImport"
            texte = FR("Op{e2}rations du dernier import -- ") & nbTotal & FR(" ligne(s) affich{e2}e(s).")

        Case "OperationsDuMois"
            texte = FR("Op{e2}rations du mois budg{e2}taire ") & Format(critMois, "00") & "/" & critAnnee & _
                    FR(" -- ") & nbTotal & FR(" op{e2}ration(s).")

        Case "ErreursSante"
            texte = FR("Op{e2}rations avec une date de consultation invalide -- ") & nbTotal & FR(" op{e2}ration(s).")

        Case "DetailTotal"
            If CStr(param) = "Positif" Then
                texte = FR("D{e2}tail du total POSITIF du mois budg{e2}taire ") & Format(critMois, "00") & "/" & critAnnee
            ElseIf CStr(param) = "Negatif" Then
                texte = FR("D{e2}tail du total N{e2}GATIF du mois budg{e2}taire ") & Format(critMois, "00") & "/" & critAnnee
            Else
                texte = FR("D{e2}tail du total du mois budg{e2}taire ") & Format(critMois, "00") & "/" & critAnnee
            End If
            texte = texte & FR(" -- ") & nbTotal & FR(" op{e2}ration(s).")

        Case "SuiviSante"
            ' Ajout 07/10/2026. La phrase rappelle aussi le double-clic de recalcul,
            ' car le texte d'aide général de l'écran ne peut pas être modifié sans
            ' réinstaller la feuille (voir la réponse du 07/10 sur le décalage entre
            ' le dépôt et le classeur).
            texte = FR("Suivi sant{e2} (parts ventil{e2}es comprises) -- ") & nbTotal & _
                    FR(" ligne(s). Double-clic sur SousCategorie = recalcul des statuts.")
        Case Else
            ' Sécurité : un préfiltre non prévu ici ne doit pas faire échouer
            ' l'affichage; on reste simplement discret.
            texte = nbTotal & FR(" ligne(s) affich{e2}e(s).")

    End Select

    ' Remarque : chaque fragment ci-dessus est deja passe par FR() au moment ou
    ' il est concatene -- "texte" est donc deja la chaine finale, decodee.
    DecrireFiltreActifRO = texte

End Function

' Indique si la ligne "i" de TblOperations (DataBodyRange, 1 = première ligne)
' doit être incluse pour le préfiltre demandé. Toute la logique de filtrage est
' centralisée ici afin d'éviter de la répéter lorsqu'un nouveau préfiltre sera ajouté.
Private Function LigneOpRetenuePourPrefiltre(ByVal i As Long, ByVal prefiltre As String, _
                                              ByVal showTypeDetail As String, ByRef listeDernierImport As Object) As Boolean

    Dim Montant As Double

    Select Case prefiltre

        Case ""
            LigneOpRetenuePourPrefiltre = True

        Case "DernierImport"
            LigneOpRetenuePourPrefiltre = listeDernierImport.Exists(mod_DataStructure.CellText(tblData(i, colID)))

        Case "OperationsDuMois"
            LigneOpRetenuePourPrefiltre = mod_DonneesTable.RowMatchesFilter(i, tblData(i, colMoisBud), tblData(i, colAnneeBud), tblData(i, colMontant), True)

        Case "ErreursSante"
            If colDateConsult <> 0 Then
                LigneOpRetenuePourPrefiltre = DateConsultInvalideRO(tblData(i, colDateConsult))
            End If

        Case "DetailTotal"
            If mod_DonneesTable.RowMatchesFilter(i, tblData(i, colMoisBud), tblData(i, colAnneeBud), tblData(i, colMontant), True) Then
                Montant = mod_DataStructure.ToDouble(tblData(i, colMontant))
                If showTypeDetail = "Positif" Then
                    LigneOpRetenuePourPrefiltre = (Montant >= 0)
                ElseIf showTypeDetail = "Negatif" Then
                    LigneOpRetenuePourPrefiltre = (Montant < 0)
                End If
            End If
        Case "SuiviSante"
            ' Ajout 07/10/2026 (remplace l'ancien Synthese_Care). Le test porte sur
            ' SousCategorie : depuis la Phase 1, la colonne Categorie contient la
            ' catégorie PARENTE ("Santé, prévoyance"), c'était justement la panne
            ' de l'ancien écran.
            If colSousCategorie <> 0 Then
                LigneOpRetenuePourPrefiltre = _
                    (mod_DataStructure.CellText(tblData(i, colSousCategorie)) = mod_VarGlobales.SOUS_CATEGORIE_SANTE)
            End If
        Case Else
            LigneOpRetenuePourPrefiltre = True

    End Select

End Function

' Reproduit le test de l'ancien écran "Erreurs santé" : une consultation est en
' erreur lorsque sa date vaut le 02/01/1900 (numéro de série Excel = 2), valeur
' sentinelle écrite par mod_ImportOFX lorsque le segment Notes ne peut pas être
' converti en date (voir DateSerial(1900,1,2) dans mod_ImportOFX.ImporterOperationsOFX).
' AMÉLIORATION PAR RAPPORT À L'ANCIEN ÉCRAN (à signaler à l'opérateur) :
' l'ancien code comparait le TEXTE "02/01/1900" à la valeur de la cellule, ce qui
' n'est fiable qu'avec certains paramètres régionaux. Ici, on compare le NUMÉRO DE
' SÉRIE (une date Excel est toujours un nombre), ce qui est fiable quels que soient
' le format d'affichage et la langue du poste.
Private Function DateConsultInvalideRO(ByVal valeurCellule As Variant) As Boolean
    If IsEmpty(valeurCellule) Then Exit Function
    On Error Resume Next
    DateConsultInvalideRO = (CLng(valeurCellule) = 2)
    On Error GoTo 0
End Function

' Charge la liste des ID_Transaction du dernier import (feuille technique
' TechDernierImport, alimentée par mod_DernierImport.MemoriserDernierImport,
' appelée automatiquement à la fin de l'import; voir mod_ImportOFX). Renvoie
' Nothing si aucun import n'a été enregistré.
Private Function ChargerListeDernierImportRO() As Object
    Dim wsTech As Worksheet
    Dim nb As Long, i As Long
    Dim d As Object

    On Error Resume Next
    Set wsTech = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_TECH)
    On Error GoTo 0
    If wsTech Is Nothing Then Exit Function

    nb = 0
    On Error Resume Next
    nb = CLng(wsTech.Range("B2").value)
    On Error GoTo 0
    If nb <= 0 Then Exit Function

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    ' La liste commence en ligne 4 (lignes 1-3 = DateImport/NbOperations/en-tête
    ' "ID_Transaction") -- voir mod_DernierImport.MemoriserDernierImport.
    For i = 1 To nb
        d(mod_DataStructure.CellText(wsTech.Cells(3 + i, 1).value)) = True
    Next i

    Set ChargerListeDernierImportRO = d
End Function

' Surligne en rouge les "critNbOperations" dépenses les plus importantes parmi
' les lignes actuellement chargées (même principe que l'ancien mod_Rapports.HighlightTopRows,
' mais adapté ici à tblRecherche.DataBodyRange, qui ne comporte pas d'en-tête,
' contrairement aux anciens écrans "Synthese_*".
Private Sub SurlignerTopDepensesRO(ByVal tblRecherche As ListObject, ByVal nbLignes As Long)

    Dim indicesDepenses() As Long, montantsDepenses() As Double, nombreDepenses As Long
    Dim nombreSurligne As Long
    Dim k As Long, jj As Long, kk As Long
    Dim montantLigne As Double
    Dim maxIdx As Long, tmpD As Double, tmpL As Long

    nombreDepenses = 0
    For k = 1 To nbLignes
        montantLigne = mod_DataStructure.ToDouble(tblRecherche.ListColumns("Montant").DataBodyRange.Cells(k).value)
        If montantLigne < 0 Then
            nombreDepenses = nombreDepenses + 1
            ReDim Preserve indicesDepenses(1 To nombreDepenses)
            ReDim Preserve montantsDepenses(1 To nombreDepenses)
            indicesDepenses(nombreDepenses) = k
            montantsDepenses(nombreDepenses) = Abs(montantLigne)
        End If
    Next k

    If nombreDepenses = 0 Then Exit Sub

    nombreSurligne = mod_Rapports.GetTopCount(critNbOperations, nombreDepenses)

    ' Tri par sélection, par montants décroissants (même principe que
    ' mod_Rapports.HighlightTopRows).
    For jj = 1 To nombreDepenses - 1
        maxIdx = jj
        For kk = jj + 1 To nombreDepenses
            If montantsDepenses(kk) > montantsDepenses(maxIdx) Then maxIdx = kk
        Next kk
        If maxIdx <> jj Then
            tmpD = montantsDepenses(jj): montantsDepenses(jj) = montantsDepenses(maxIdx): montantsDepenses(maxIdx) = tmpD
            tmpL = indicesDepenses(jj): indicesDepenses(jj) = indicesDepenses(maxIdx): indicesDepenses(maxIdx) = tmpL
        End If
    Next jj

    For jj = 1 To nombreSurligne
        tblRecherche.ListColumns("Montant").DataBodyRange.Cells(indicesDepenses(jj)).Font.Color = RGB(255, 0, 0)
    Next jj

End Sub

' Petit utilitaire : ajoute "valeur" à listeValeurs()/nbValeurs si elle n'est pas vide
' et pas déjà présente (comparaison insensible à la casse; voir .CompareMode, défini
' par l'appelant sur le dictionnaire utilisé pour mémoriser les valeurs déjà vues).
Private Sub AjouterValeurUnique(ByVal valeur As String, ByRef vues As Object, ByRef listeValeurs() As String, ByRef nbValeurs As Long)
    If valeur = "" Then Exit Sub
    If vues.Exists(valeur) Then Exit Sub
    vues.Add valeur, True
    nbValeurs = nbValeurs + 1
    ReDim Preserve listeValeurs(1 To nbValeurs)
    listeValeurs(nbValeurs) = valeur
End Sub

' Écrit la liste unique (Catégorie et SousCategorie confondues) dans une colonne
' technique HORS du tableau structuré (au-delà de sa dernière colonne), au format
' Texte, puis (re)définit une plage NOMMÉE DYNAMIQUE (OFFSET/COUNTA) qui la référence.
' Même principe que mod_Categories.RafraichirListesCategories, appliqué localement
' afin de ne pas dépendre de la présence de mod_Categories.
Private Sub EcrireListeTechniqueRO(ByVal ws As Worksheet, ByRef listeValeurs() As String, ByVal nbValeurs As Long)

    ' PHASE 6 : déplacée de "N" à "R" : la colonne N est désormais une vraie
    ' colonne du tableau (SoldeSante); la zone technique doit rester au-delà
    ' de sa dernière colonne (P).
    Const COL_TECHNIQUE As String = "R"
    Dim i As Long

    ws.Range(COL_TECHNIQUE & "1:" & COL_TECHNIQUE & "1000").ClearContents
    ws.Range(COL_TECHNIQUE & "1:" & COL_TECHNIQUE & "1000").NumberFormat = "@"
    ws.Columns(COL_TECHNIQUE).Hidden = True

    ws.Range(COL_TECHNIQUE & "1").value = mod_Display.FR("Liste technique (Cat{e2}gorie+SousCat{e2}gorie) - ne pas modifier")

    If nbValeurs = 0 Then Exit Sub

    For i = 1 To nbValeurs
        ws.Range(COL_TECHNIQUE & (1 + i)).value = listeValeurs(i)
    Next i

    On Error Resume Next
    ThisWorkbook.Names("ListeValeursRechercheOperations").Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add Name:="ListeValeursRechercheOperations", _
        RefersTo:="=OFFSET(" & ws.Name & "!$" & COL_TECHNIQUE & "$2,0,0," & _
                  "MAX(1,COUNTA(" & ws.Name & "!$" & COL_TECHNIQUE & "$2:$" & COL_TECHNIQUE & "$1000)),1)"

End Sub


' =====================================================================================
' AppliquerLignesMarquees : applique Catégorie/SousCategorie/Notes pour les lignes
' marquées "Oui", en écrivant dans la table source de chacune (colonne technique
' SourceLigne : "O" = TblOperations, "V" = TblVentilations).
' =====================================================================================
Public Sub AppliquerLignesMarquees()

    Dim ws As Worksheet
    Dim tblRecherche As ListObject
    Dim n As Long, nTbl As Long, i As Long
    Dim nbAppliquees As Long
    Dim indexParID As Object
    Dim donnees As Variant

    Dim tblVen As ListObject
    Dim venDisponible As Boolean

    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_RECHERCHE)
    Set tblRecherche = ws.ListObjects(mod_VarGlobales.NOM_TABLE_RECHERCHE)

    If tblRecherche.DataBodyRange Is Nothing Then
        MsgBox mod_Display.FR("Aucune ligne {a2} traiter. Utilise d'abord 'Rechercher'."), vbInformation
        Exit Sub
    End If

    Set tbl = mod_DonneesTable.GetOperationsTable()
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    mod_Display.RecupIndexCol
    tblData = tbl.DataBodyRange.value
    nTbl = UBound(tblData, 1)

    Set indexParID = CreateObject("Scripting.Dictionary")
    indexParID.CompareMode = 1
    For i = 1 To nTbl
        indexParID(mod_DataStructure.CellText(tblData(i, colID))) = i
    Next i

    venDisponible = False
    Set tblVen = ObtenirTableVentilationsRO()
    If Not tblVen Is Nothing Then
        If Not tblVen.DataBodyRange Is Nothing Then venDisponible = True
    End If

    donnees = tblRecherche.DataBodyRange.value
    n = UBound(donnees, 1)

    nbAppliquees = 0

    For i = 1 To n
        Dim marque As String
        marque = Trim(CStr(donnees(i, mod_InstallRechercheOperations.RO_COL_VALIDER)))

        If LCase(marque) = "oui" Then
            Dim sourceLigne As String
            sourceLigne = Trim(CStr(donnees(i, mod_InstallRechercheOperations.RO_COL_SOURCE)))

            ' Ajout du 02/10/2026 (après discussion avec l'opérateur, option 2) : Categorie
            ' et SousCategorie ne sont plus gérées par ce mécanisme; elles se modifient
            ' désormais par double-clic (voir EditerCategorieRO et RevoirVentilationRO),
            ' qui écrivent immédiatement dans la table source, sans passer par "Valider"
            ' ni par ce bouton. Seul Notes reste traité ici par lot, comme auparavant.
            Dim notesValeur As String
            notesValeur = CStr(donnees(i, mod_InstallRechercheOperations.RO_COL_NOTES))

            If sourceLigne = "V" Then
                ' --- PHASE 5 : ligne de TblVentilations, identifiée directement
                '     par sa position (colonne technique LigneVentilation) --------
                If venDisponible Then
                    Dim ligneVenCible As Long
                    ligneVenCible = CLng(donnees(i, mod_InstallRechercheOperations.RO_COL_LIGNEVEN))

                    Dim notesOrigineVen As String, seg0Ven As String, verrouilleeVen As Boolean
                    notesOrigineVen = mod_DataStructure.CellText(tblVen.ListColumns("Notes").DataBodyRange.rows(ligneVenCible).value)
                    seg0Ven = mod_ImportOFX.SegmentTexte(notesOrigineVen, ";", 0)
                    verrouilleeVen = mod_ImportOFX.EstDateValide(seg0Ven)
                    If Not verrouilleeVen Then
                        tblVen.ListColumns("Notes").DataBodyRange.rows(ligneVenCible).value = notesValeur
                    End If

                    nbAppliquees = nbAppliquees + 1
                    tblRecherche.ListColumns("Valider").DataBodyRange.Cells(i).value = ""
                End If
            Else
                ' --- Ligne de TblOperations (Notes uniquement; voir le commentaire ci-dessus) ---
                Dim idCible As String
                idCible = Trim(CStr(donnees(i, mod_InstallRechercheOperations.RO_COL_ID)))

                If indexParID.Exists(idCible) Then
                    Dim ligneCible As Long
                    ligneCible = indexParID(idCible)

                    Dim notesOriginale As String, seg0 As String, verrouillee As Boolean
                    notesOriginale = mod_DataStructure.CellText(tblData(ligneCible, colNotes))
                    seg0 = mod_ImportOFX.SegmentTexte(notesOriginale, ";", 0)
                    verrouillee = mod_ImportOFX.EstDateValide(seg0)

                    If Not verrouillee Then
                        tbl.ListColumns("Notes").DataBodyRange.rows(ligneCible).value = notesValeur
                    End If

                    nbAppliquees = nbAppliquees + 1
                    tblRecherche.ListColumns("Valider").DataBodyRange.Cells(i).value = ""
                End If
            End If
        End If
    Next i

    If nbAppliquees > 0 Then
        mod_FormulairesNotes.VerifierNotesSante
        mod_SuiviSante.CalculerSuiviSante AfficherResume:=False
    End If

    MsgBox nbAppliquees & mod_Display.FR(" ligne(s) appliqu{e2}e(s)."), vbInformation

End Sub

' =====================================================================================
' RevoirVentilationRO (ajout à la suite d'un test opérateur)
' =====================================================================================
' Rouvre le formulaire de ventilation pour la ligne actuellement sélectionnée dans
' le tableau de recherche (même principe que ReinitialiserLigneSelectionnee dans
' mod_ResolutionCategories : lecture d'ActiveCell.Row, sans bouton par ligne).
'
' La ligne sélectionnée peut être :
'   - une opération parente déjà ventilée (Categorie = "Ventile");
'   - l'une de ses PARTS (colonne Ventile = "Oui", Source = "V");
'   - ou une opération qui n'est PAS ENCORE ventilée.
' Dans les deux premiers cas, RO_COL_ID contient déjà l'ID_Transaction de l'opération
' bancaire parente (voir RechercherOperations, qui renseigne cette valeur pour tous les
' types de lignes) : c'est cet identifiant qu'attend mod_Ventilation.OuvrirVentilation.
'
' Ajout du 02/10/2026 : cette procédure gère désormais deux cas, selon que la ligne
' est déjà ventilée ou non (voir "etaitDejaVentilee" plus bas) :
'   - déjà ventilée : on rouvre le détail existant (comportement d'origine);
'   - pas encore ventilée : on démarre une NOUVELLE ventilation (formulaire vide).
'     Si l'opérateur la valide, la catégorie de l'opération passe à "Ventile"
'     (voir MarquerOperationVentileeRO), comme dans mod_ControleCategories.ControleVentiler
'     lors d'un import. Auparavant, un double-clic sur une cellule Ventile vide affichait
'     simplement "rien à revoir", sans proposer d'autre action.
'
' LIMITE CONNUE : TblVentilations ne conserve pas le libellé brut de l'opération
' bancaire (il n'existait qu'en mémoire pendant l'import; voir
' mod_ControleCategories.ControleVentiler). Lors d'une ouverture depuis cet écran,
' bien après l'import, ce libellé n'est donc plus disponible : le formulaire l'affiche
' vide plutôt que d'inventer une valeur. Les autres données ne sont pas affectées :
' la date, le Tiers, le montant et la catégorie/sous-catégorie restent exacts, car ils
' sont relus directement dans TblOperations.
Public Sub RevoirVentilationRO()

    Dim ws As Worksheet
    Dim tblRecherche As ListObject
    Dim ligneSelection As Long, ligneRelative As Long
    Dim categorieValeur As String, ventileValeur As String
    Dim idTransaction As String
    Dim ok As Boolean
    Dim etaitDejaVentilee As Boolean

    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_RECHERCHE)
    Set tblRecherche = ws.ListObjects(mod_VarGlobales.NOM_TABLE_RECHERCHE)

    If tblRecherche.DataBodyRange Is Nothing Then
        MsgBox mod_Display.FR("Aucune ligne {a2} traiter. Utilise d'abord 'Rechercher'."), vbInformation
        Exit Sub
    End If

    ligneSelection = ActiveCell.Row
    ligneRelative = ligneSelection - tblRecherche.DataBodyRange.Row + 1

    If ligneRelative < 1 Or ligneRelative > tblRecherche.DataBodyRange.rows.count Then
        MsgBox mod_Display.FR("S{e2}lectionne d'abord une ligne d'op{e2}ration dans le tableau, puis clique sur ce bouton."), vbExclamation
        Exit Sub
    End If

    categorieValeur = Trim(CStr(tblRecherche.DataBodyRange.Cells(ligneRelative, mod_InstallRechercheOperations.RO_COL_CATEGORIE).value))
    ventileValeur = Trim(CStr(tblRecherche.DataBodyRange.Cells(ligneRelative, mod_InstallRechercheOperations.RO_COL_VENTILE).value))
    etaitDejaVentilee = (LCase(ventileValeur) = "oui" Or categorieValeur = CategorieVentileRO())

    idTransaction = Trim(CStr(tblRecherche.DataBodyRange.Cells(ligneRelative, mod_InstallRechercheOperations.RO_COL_ID).value))
    If idTransaction = "" Then
        MsgBox mod_Display.FR("Impossible de retrouver l'ID_Transaction de cette ligne."), vbExclamation
        Exit Sub
    End If

    ' --- On relit l'en-tête (date/Tiers/montant/catégorie actuelle) directement sur
    ' l'opération parente dans TblOperations : TblVentilations n'a pas ces colonnes,
    ' dont OuvrirVentilation a besoin pour l'affichage en lecture seule. ------------
    Dim tblOp As ListObject
    Dim donneesOp As Variant
    Dim i As Long, ligneParent As Long
    Dim dateOp As Variant, tiersOp As String
    Dim montantOp As Double, catActuelleOp As String, sousActuelleOp As String

    Set tblOp = mod_DonneesTable.GetOperationsTable()
    If tblOp Is Nothing Or tblOp.DataBodyRange Is Nothing Then
        MsgBox mod_Display.FR("Le tableau TblOperations est introuvable."), vbExclamation
        Exit Sub
    End If

    mod_Display.RecupIndexCol
    donneesOp = tblOp.DataBodyRange.value

    ligneParent = 0
    For i = 1 To UBound(donneesOp, 1)
        If mod_DataStructure.CellText(donneesOp(i, colID)) = idTransaction Then
            ligneParent = i
            Exit For
        End If
    Next i

    If ligneParent = 0 Then
        MsgBox mod_Display.FR("Op{e2}ration parente introuvable dans TblOperations (ID '") & idTransaction & "').", vbExclamation
        Exit Sub
    End If

    dateOp = donneesOp(ligneParent, colDate)
    tiersOp = mod_DataStructure.CellText(donneesOp(ligneParent, colTiers))
    montantOp = mod_DataStructure.ToDouble(donneesOp(ligneParent, colMontant))
    catActuelleOp = mod_DataStructure.CellText(donneesOp(ligneParent, colCategorie))
    sousActuelleOp = ""
    If colSousCategorie <> 0 Then sousActuelleOp = mod_DataStructure.CellText(donneesOp(ligneParent, colSousCategorie))

    Dim ventilationSupprimee As Boolean
    ok = mod_Ventilation.OuvrirVentilation(idTransaction, dateOp, tiersOp, "", montantOp, catActuelleOp, sousActuelleOp, ventilationSupprimee)

    ws.Activate

    If ok Then
        If Not etaitDejaVentilee Then
            ' Ajout 02/10/2026 (demarrer une ventilation depuis la recherche) : cette
            ' operation ne l'etait pas encore avant l'ouverture du formulaire -- on
            ' bascule sa categorie sur "Ventile" et on memorise l'ancienne, exactement
            ' comme le fait mod_ControleCategories.ControleVentiler a l'import.
            MarquerOperationVentileeRO tblOp, ligneParent, catActuelleOp, sousActuelleOp
        End If
        ' PHASE 6 : on recharge l'ecran dans le MEME contexte qu'avant (meme
        ' prefiltre/param), plutot que de retomber en recherche libre.
        RechercherOperations g_ROPrefiltreActif, g_ROParamActif
    ElseIf ventilationSupprimee Then
        ' Ajout 01/10/2026 (point 4 : annuler une ventilation) : l'operateur a
        ' supprime la ventilation existante (bouton "Supprimer cette ventilation" de
        ' frm_Ventilation) -- on restaure la categorie/sous-categorie d'avant la toute
        ' premiere ventilation de cette operation (voir RestaurerCategorieAvantVentilation
        ' plus bas), puis on recharge l'ecran dans le meme contexte qu'avant, exactement
        ' comme pour une ventilation validee normalement.
        RestaurerCategorieAvantVentilation tblOp, ligneParent
        RechercherOperations g_ROPrefiltreActif, g_ROParamActif
    End If

End Sub

' Ajout du 01/10/2026 (point 4 : annuler une ventilation depuis cet écran) ----------------
' Relit les colonnes techniques CategorieAvantVentilation / SousCategorieAvantVentilation
' (ajoutées par mod_InstallVentilation.AjouterColonnesAnnulationVentilation) de la ligne
' ligneParent dans TblOperations et les recopie dans Categorie / SousCategorie. Il s'agit
' de la catégorie de l'opération AVANT sa première ventilation (mémorisée par
' mod_ControleCategories.ControleVentiler à l'import). On vide ensuite ces deux colonnes :
' il n'y a plus rien à restaurer tant qu'une nouvelle ventilation n'a pas été créée.
' ligneParent : numéro de ligne DANS LE TABLEAU (1 = première ligne de données), tel que
' calculé plus haut par RevoirVentilationRO; ce n'est PAS un numéro de ligne de la feuille.
Private Sub RestaurerCategorieAvantVentilation(ByVal tblOp As ListObject, ByVal ligneParent As Long)

    Dim colCatAvant As Long, colSousAvant As Long
    Dim catAvant As String, sousAvant As String

    colCatAvant = 0
    colSousAvant = 0
    On Error Resume Next
    colCatAvant = tblOp.ListColumns(mod_InstallVentilation.NOM_COL_CAT_AVANT_VENTILATION).index
    colSousAvant = tblOp.ListColumns(mod_InstallVentilation.NOM_COL_SOUS_AVANT_VENTILATION).index
    On Error GoTo 0

    If colCatAvant = 0 Or colSousAvant = 0 Then
        MsgBox mod_Display.FR("Les colonnes 'CategorieAvantVentilation' / 'SousCategorieAvantVentilation' sont introuvables dans TblOperations.") & vbCrLf & _
               mod_Display.FR("La cat{e2}gorie d'origine n'a pas pu {ea}tre restaur{e2}e automatiquement."), vbExclamation
        Exit Sub
    End If

    catAvant = mod_DataStructure.CellText(tblOp.DataBodyRange.Cells(ligneParent, colCatAvant).value)
    sousAvant = mod_DataStructure.CellText(tblOp.DataBodyRange.Cells(ligneParent, colSousAvant).value)

    tblOp.DataBodyRange.Cells(ligneParent, colCategorie).value = catAvant
    If colSousCategorie <> 0 Then tblOp.DataBodyRange.Cells(ligneParent, colSousCategorie).value = sousAvant

    ' On vide les colonnes techniques : l'opération n'est plus ventilée; il n'y a donc
    ' plus de catégorie antérieure à conserver.
    tblOp.DataBodyRange.Cells(ligneParent, colCatAvant).value = ""
    tblOp.DataBodyRange.Cells(ligneParent, colSousAvant).value = ""

End Sub

' Ajout du 02/10/2026 (démarrer une ventilation depuis l'écran de recherche, en plus
' de l'écran de contrôle à l'import) : bascule l'opération vers la catégorie "Ventile",
' après avoir mémorisé son ancienne catégorie/sous-catégorie dans les colonnes
' CategorieAvantVentilation / SousCategorieAvantVentilation de TblOperations (ajoutées
' par mod_InstallVentilation), afin de les restaurer si la ventilation est supprimée
' plus tard (voir RestaurerCategorieAvantVentilation ci-dessus). Même principe que
' mod_ControleCategories.ControleVentiler à l'import.
' ligneParent : numéro de ligne DANS LE TABLEAU (1 = première ligne de données), tel que
' calculé plus haut par RevoirVentilationRO; ce n'est PAS un numéro de ligne de la feuille.
Private Sub MarquerOperationVentileeRO(ByVal tblOp As ListObject, ByVal ligneParent As Long, _
                                       ByVal ancienneCategorie As String, ByVal ancienneSousCategorie As String)

    Dim colCatAvant As Long, colSousAvant As Long

    colCatAvant = 0
    colSousAvant = 0
    On Error Resume Next
    colCatAvant = tblOp.ListColumns(mod_InstallVentilation.NOM_COL_CAT_AVANT_VENTILATION).index
    colSousAvant = tblOp.ListColumns(mod_InstallVentilation.NOM_COL_SOUS_AVANT_VENTILATION).index
    On Error GoTo 0

    If colCatAvant <> 0 And colSousAvant <> 0 Then
        tblOp.DataBodyRange.Cells(ligneParent, colCatAvant).value = ancienneCategorie
        tblOp.DataBodyRange.Cells(ligneParent, colSousAvant).value = ancienneSousCategorie
    Else
        MsgBox mod_Display.FR("Les colonnes 'CategorieAvantVentilation' / 'SousCategorieAvantVentilation' sont introuvables dans TblOperations.") & vbCrLf & _
               mod_Display.FR("La cat{e2}gorie d'origine ne pourra pas {ea}tre restaur{e2}e automatiquement si cette ventilation est supprim{e2}e plus tard."), vbExclamation
    End If

    tblOp.DataBodyRange.Cells(ligneParent, colCategorie).value = CategorieVentileRO()
    If colSousCategorie <> 0 Then tblOp.DataBodyRange.Cells(ligneParent, colSousCategorie).value = ""

    ' La catégorie "Ventile" doit exister dans le référentiel pour que les listes
    ' déroulantes (et un futur contrôle à l'import) l'acceptent. Même appel que
    ' mod_ControleCategories.ControleVentiler.
    mod_Categories.AjouterCategoriePersonnalisee CategorieVentileRO(), ""
    mod_Categories.RafraichirListesCategories

    mod_SuiviSante.CalculerSuiviSante AfficherResume:=False

End Sub

' =====================================================================================
' EditerCategorieRO (ajout du 02/10/2026)
' =====================================================================================
' Déclenchée par un double-clic sur une cellule de la colonne Categorie (voir
' ThisWorkbook.Workbook_SheetBeforeDoubleClick), cette procédure réutilise le MÊME
' formulaire que celui du contrôle des catégories à l'import
' (mod_ControleCategories.ControlerCategories), en mode "une seule opération"
' (paramètre uneSeuleOperation). Elle bénéficie ainsi des fonctionnalités existantes
' (liste déroulante de sous-catégorie dépendante, bouton "+" pour créer une catégorie, etc.)
' sans dupliquer le code.
'
' Ne s'applique PAS à une ligne déjà ventilée (Ventile = "Oui" ou Categorie =
' "Ventile") : sa catégorie se modifie depuis le détail de la ventilation (double-clic
' sur Ventile), et non ici. Un message d'information l'explique afin d'éviter de laisser
' croire à un bogue.
Public Sub EditerCategorieRO()

    Dim ws As Worksheet
    Dim tblRecherche As ListObject
    Dim ligneSelection As Long, ligneRelative As Long
    Dim categorieValeur As String, ventileValeur As String
    Dim idTransaction As String

    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_RECHERCHE)
    Set tblRecherche = ws.ListObjects(mod_VarGlobales.NOM_TABLE_RECHERCHE)

    If tblRecherche.DataBodyRange Is Nothing Then Exit Sub

    ligneSelection = ActiveCell.Row
    ligneRelative = ligneSelection - tblRecherche.DataBodyRange.Row + 1
    If ligneRelative < 1 Or ligneRelative > tblRecherche.DataBodyRange.rows.count Then Exit Sub

    categorieValeur = Trim(CStr(tblRecherche.DataBodyRange.Cells(ligneRelative, mod_InstallRechercheOperations.RO_COL_CATEGORIE).value))
    ventileValeur = Trim(CStr(tblRecherche.DataBodyRange.Cells(ligneRelative, mod_InstallRechercheOperations.RO_COL_VENTILE).value))

    If LCase(ventileValeur) = "oui" Or categorieValeur = CategorieVentileRO() Then
        MsgBox mod_Display.FR("Cette op{e2}ration est ventil{e2}e : sa cat{e2}gorie se g{e2}re depuis le d{e2}tail de la ventilation (double-clique sur la colonne Ventile), pas ici."), vbInformation
        Exit Sub
    End If

    idTransaction = Trim(CStr(tblRecherche.DataBodyRange.Cells(ligneRelative, mod_InstallRechercheOperations.RO_COL_ID).value))
    If idTransaction = "" Then
        MsgBox mod_Display.FR("Impossible de retrouver l'ID_Transaction de cette ligne."), vbExclamation
        Exit Sub
    End If

    Dim tblOp As ListObject
    Dim donneesOp As Variant
    Dim i As Long, ligneParent As Long
    Dim categorieActuelleTbl As String, sousCategorieActuelleTbl As String

    Set tblOp = mod_DonneesTable.GetOperationsTable()
    If tblOp Is Nothing Or tblOp.DataBodyRange Is Nothing Then
        MsgBox mod_Display.FR("Le tableau TblOperations est introuvable."), vbExclamation
        Exit Sub
    End If

    mod_Display.RecupIndexCol
    donneesOp = tblOp.DataBodyRange.value

    ligneParent = 0
    For i = 1 To UBound(donneesOp, 1)
        If mod_DataStructure.CellText(donneesOp(i, colID)) = idTransaction Then
            ligneParent = i
            Exit For
        End If
    Next i

    If ligneParent = 0 Then
        MsgBox mod_Display.FR("Op{e2}ration introuvable dans TblOperations (ID '") & idTransaction & "').", vbExclamation
        Exit Sub
    End If

    categorieActuelleTbl = mod_DataStructure.CellText(donneesOp(ligneParent, colCategorie))
    sousCategorieActuelleTbl = ""
    If colSousCategorie <> 0 Then sousCategorieActuelleTbl = mod_DataStructure.CellText(donneesOp(ligneParent, colSousCategorie))

    ' --- Formulaire de contrôle des catégories, en mode "une seule opération" :
    ' CTRL_OP_CATSOURCE est volontairement laissé vide (pas de source bancaire ici; voir
    ' ControlerCategories). Ce sont categorieActuelleUnique/sousCategorieActuelleUnique
    ' ci-dessous qui préremplissent le formulaire avec la catégorie actuelle. ------------
    Dim ops(1 To 1, 1 To mod_ControleCategories.CTRL_OP_NBCOL) As Variant
    Dim catFinale() As String, sousFinale() As String
    Dim catAvantVen() As String, sousAvantVen() As String
    ' AJOUT du 03/10/2026 : ControlerCategories exige maintenant ces deux tableaux de
    ' sortie supplémentaires (Tiers/Notes modifiables à l'import; voir ce module). Ils ne
    ' sont pas utilisés ici (TiersNotesEditables les garde figés en mode "une seule
    ' opération"; voir mod_ControleCategories); on les déclare uniquement pour l'appel.
    Dim tiersFinaleInutilise() As String, libelleFinaleInutilise() As String
    Dim ok As Boolean

    ops(1, mod_ControleCategories.CTRL_OP_DATE) = donneesOp(ligneParent, colDate)
    ops(1, mod_ControleCategories.CTRL_OP_TIERS) = mod_DataStructure.CellText(donneesOp(ligneParent, colTiers))
    ops(1, mod_ControleCategories.CTRL_OP_LIBELLE) = ""
    ops(1, mod_ControleCategories.CTRL_OP_MONTANT) = mod_DataStructure.ToDouble(donneesOp(ligneParent, colMontant))
    ops(1, mod_ControleCategories.CTRL_OP_CATSOURCE) = ""
    ops(1, mod_ControleCategories.CTRL_OP_ID) = idTransaction

    ok = mod_ControleCategories.ControlerCategories(ops, 1, catFinale, sousFinale, catAvantVen, sousAvantVen, _
                                                     tiersFinaleInutilise, libelleFinaleInutilise, _
                                                     uneSeuleOperation:=True, _
                                                     categorieActuelleUnique:=categorieActuelleTbl, _
                                                     sousCategorieActuelleUnique:=sousCategorieActuelleTbl)

    ws.Activate

    If ok Then
        tblOp.DataBodyRange.Cells(ligneParent, colCategorie).value = catFinale(1)
        If colSousCategorie <> 0 Then tblOp.DataBodyRange.Cells(ligneParent, colSousCategorie).value = sousFinale(1)
        mod_SuiviSante.CalculerSuiviSante AfficherResume:=False
        RechercherOperations g_ROPrefiltreActif, g_ROParamActif
    End If

End Sub

' =====================================================================================
' RecalculerSuiviSanteRO (ajout 07/10/2026, demande opérateur) : relance le calcul
' officiel des statuts santé, puis recharge l'écran dans le MÊME contexte, pour que
' l'opérateur voie immédiatement les StatutSante/SoldeSante à jour (par exemple
' après avoir saisi une Franchise).
' RÉUTILISATION : mod_SuiviSante.CalculerSuiviSante est appelée telle quelle,
' sans aucune modification.
' APPEL : par un double-clic sur une cellule SousCategorie, UNIQUEMENT quand
' l'écran affiche le préfiltre "SuiviSante" (voir ThisWorkbook).
' RAPPEL (règle inchangée de mod_SuiviSante) : un groupe déjà "OK", ou "KO" avec
' DepassementHoraires = Vrai, n'est PAS recalculé. Ce double-clic ne remet donc
' jamais en cause un statut déjà accepté par l'opérateur.
' =====================================================================================
Public Sub RecalculerSuiviSanteRO()
 
    Dim reponse As VbMsgBoxResult
 
    ' Garde-fou contre un double-clic accidentel : le calcul ÉCRIT dans
    ' TblOperations et TblVentilations (StatutSante, SoldeSante, DepassementHoraires).
    reponse = MsgBox(FR("Recalculer les statuts de suivi sant{e2} de toutes les op{e2}rations ?"), _
                     vbYesNo + vbQuestion, FR("Suivi sant{e2}"))
    If reponse <> vbYes Then Exit Sub
 
    ' True = afficher le résumé habituel du calcul (nombre de groupes recalculés).
    mod_SuiviSante.CalculerSuiviSante True
 
    ' CalculerSuiviSante modifie les variables globales tbl et colXxx : on recharge
    ' l'écran, qui les réinitialise proprement, avec le même préfiltre qu'avant.
    RechercherOperations g_ROPrefiltreActif, g_ROParamActif
 
End Sub

' PHASE 6 : retourne sur la feuille "Resultat" (bilan mensuel) si l'écran a été
' ouvert depuis le détail d'un total, sinon sur Synthèse comme avant.
Public Sub SortirRechercheOperations()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_RECHERCHE)
    On Error Resume Next
    If g_ROPrefiltreActif = "DetailTotal" Then
        ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_RESULTAT).Activate
    Else
        Sheets("Synthese").Activate
    End If
    On Error GoTo 0
    ws.Visible = xlSheetVeryHidden
End Sub


' =====================================================================================
' AJOUT du 03/10/2026 (demande opérateur) : FONCTIONS PARTAGÉES PAR LES DEUX BOUTONS
' "DÉCALAGE DE BUDGET" DE L'ÉCRAN DE RECHERCHE (voir plus bas
' AjouterDecalageDepuisRO et DecalerBudgetOperationRO).
' =====================================================================================

' Relit la ligne SÉLECTIONNÉE dans les résultats de recherche et renvoie son
' ID_Transaction, Tiers, Categorie et SousCategorie tels qu'affichés à l'écran.
' Reprend le mécanisme de RevoirVentilationRO (ActiveCell.Row -> ligne relative ->
' lecture directe dans TblRechercheOperations). Renvoie une chaîne vide pour
' idTransaction si aucune ligne valide n'est sélectionnée et affiche alors le message
' d'erreur; l'appelant n'a qu'à tester idTransaction = "".
Private Sub LireSelectionRO(ByRef idTransaction As String, ByRef tiersSel As String, _
                             ByRef categorieSel As String, ByRef sousCategorieSel As String)

    Dim ws As Worksheet
    Dim tblRecherche As ListObject
    Dim ligneSelection As Long, ligneRelative As Long

    idTransaction = ""

    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_RECHERCHE)
    Set tblRecherche = ws.ListObjects(mod_VarGlobales.NOM_TABLE_RECHERCHE)

    If tblRecherche.DataBodyRange Is Nothing Then
        MsgBox mod_Display.FR("Aucune ligne {a2} traiter. Utilise d'abord 'Rechercher'."), vbInformation
        Exit Sub
    End If

    ligneSelection = ActiveCell.Row
    ligneRelative = ligneSelection - tblRecherche.DataBodyRange.Row + 1

    If ligneRelative < 1 Or ligneRelative > tblRecherche.DataBodyRange.rows.count Then
        MsgBox mod_Display.FR("S{e2}lectionne d'abord une ligne d'op{e2}ration dans le tableau, puis clique sur ce bouton."), vbExclamation
        Exit Sub
    End If

    idTransaction = Trim(CStr(tblRecherche.DataBodyRange.Cells(ligneRelative, mod_InstallRechercheOperations.RO_COL_ID).value))
    tiersSel = Trim(CStr(tblRecherche.DataBodyRange.Cells(ligneRelative, mod_InstallRechercheOperations.RO_COL_TIERS).value))
    categorieSel = Trim(CStr(tblRecherche.DataBodyRange.Cells(ligneRelative, mod_InstallRechercheOperations.RO_COL_CATEGORIE).value))
    sousCategorieSel = Trim(CStr(tblRecherche.DataBodyRange.Cells(ligneRelative, mod_InstallRechercheOperations.RO_COL_SOUSCATEGORIE).value))

    If idTransaction = "" Then
        MsgBox mod_Display.FR("Impossible de retrouver l'ID_Transaction de cette ligne."), vbExclamation
    End If

End Sub

' Demande à l'opérateur un nombre entier de mois de décalage (+ ou -) et le valide
' (même méthode que dans mod_ControleCategories.bas : IsNumeric et contrôle des limites).
' Renvoie True si une valeur valide a été saisie (décalage transmis par référence),
' False si l'opérateur a annulé ou saisi une valeur non numérique.
Private Function SaisirDecalageRO(ByVal titreBoite As String, ByRef decalage As Long) As Boolean

    Dim saisie As String

    SaisirDecalageRO = False

    saisie = InputBox(mod_Display.FR("Nombre de mois {a2} d{e2}caler (par exemple 1, ou -1) :"), titreBoite)
    If Trim(saisie) = "" Then Exit Function

    If Not IsNumeric(saisie) Then
        MsgBox mod_Display.FR("Veuillez saisir un nombre entier (par exemple 1 ou -1)."), vbExclamation
        Exit Function
    End If

    decalage = CLng(saisie)
    If decalage < -24 Or decalage > 24 Then
        MsgBox mod_Display.FR("Le d{e2}calage doit {ea}tre compris entre -24 et 24 mois."), vbExclamation
        Exit Function
    End If

    SaisirDecalageRO = True

End Function


' =====================================================================================
' BOUTON "Ajouter un décalage" : ajoute une RÈGLE GÉNÉRALE dans TblDecalagesBudget
' (feuille Param) à partir du Tiers/Categorie/SousCategorie de la ligne sélectionnée
' dans les résultats de recherche. Cette règle s'appliquera à TOUTE opération, déjà
' importée ou future, correspondant aux mêmes critères, et pas seulement à la ligne
' sélectionnée. Pour ne décaler QUE cette opération, voir le bouton "Decaler le budget
' de cette operation" (DecalerBudgetOperationRO), juste après.
' =====================================================================================
Public Sub AjouterDecalageDepuisRO()

    Dim idTransaction As String, tiersSel As String, categorieSel As String, sousCategorieSel As String
    Dim decalage As Long
    Dim reponse As VbMsgBoxResult

    LireSelectionRO idTransaction, tiersSel, categorieSel, sousCategorieSel
    If idTransaction = "" Then Exit Sub

    reponse = MsgBox(mod_Display.FR("Ajouter une r{e2}gle de d{e2}calage pour :") & vbCrLf & vbCrLf & _
                      "Tiers : " & tiersSel & vbCrLf & _
                      "Cat" & ChrW(233) & "gorie : " & categorieSel & vbCrLf & _
                      "Sous-cat" & ChrW(233) & "gorie : " & sousCategorieSel & vbCrLf & vbCrLf & _
                      mod_Display.FR("Cette r{e2}gle s'appliquera {a2} TOUTES les op{e2}rations correspondant {a2} ces crit{ea}res."), _
                      vbOKCancel + vbQuestion, mod_Display.FR("Ajouter un d{e2}calage"))
    If reponse <> vbOK Then Exit Sub

    If Not SaisirDecalageRO(mod_Display.FR("Ajouter un d{e2}calage"), decalage) Then Exit Sub

    If mod_DecalagesBudget.AjouterRegleDecalage(tiersSel, categorieSel, sousCategorieSel, decalage) Then
        MsgBox mod_Display.FR("R{e2}gle ajout{e2}e {a2} TblDecalagesBudget (feuille Param) :") & vbCrLf & _
               "Tiers=" & tiersSel & " / Cat" & ChrW(233) & "gorie=" & categorieSel & " / Sous-cat" & ChrW(233) & "gorie=" & sousCategorieSel & _
               " / D" & ChrW(233) & "calage=" & decalage & vbCrLf & vbCrLf & _
               mod_Display.FR("Elle s'appliquera au prochain calcul de budget (import ou d{e2}calage manuel)."), vbInformation
    End If

End Sub


' =====================================================================================
' BOUTON "Decaler le budget de cette operation" : force le décalage de budget de LA
' SEULE opération sélectionnée, sans toucher aux autres, via
' mod_DecalagesBudget.AppliquerDecalageManuel. Actualise ensuite l'écran dans le même
' contexte qu'avant (même principe que RevoirVentilationRO plus haut).
' =====================================================================================
Public Sub DecalerBudgetOperationRO()

    Dim idTransaction As String, tiersSel As String, categorieSel As String, sousCategorieSel As String
    Dim decalage As Long

    LireSelectionRO idTransaction, tiersSel, categorieSel, sousCategorieSel
    If idTransaction = "" Then Exit Sub

    If Not SaisirDecalageRO(mod_Display.FR("D{e2}caler le budget de cette op{e2}ration"), decalage) Then Exit Sub

    If mod_DecalagesBudget.AppliquerDecalageManuel(idTransaction, decalage) Then
        MsgBox mod_Display.FR("D{e2}calage de ") & decalage & mod_Display.FR(" mois appliqu{e2} {a2} cette op{e2}ration uniquement.") & vbCrLf & _
               mod_Display.FR("Les autres op{e2}rations de m{ea}me Tiers/Cat{e2}gorie ne sont pas affect{e2}es."), vbInformation
        RechercherOperations g_ROPrefiltreActif, g_ROParamActif
    End If

End Sub


