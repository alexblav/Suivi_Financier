Option Explicit

' =====================================================================================
' MODULE : mod_RechercheOperations
'
' ROLE (PHASE 5b) : logique de la feuille frm_RechercheOperations construite
' en Phase 5a (mod_InstallRechercheOperations).
'
'   RechercherOperations   : recharge le tableau de recherche depuis TblOperations
'                            ET, PHASE 5, depuis TblVentilations (filtre ensuite
'                            avec les fleches natives Excel dans l'entete du
'                            tableau).
'   AppliquerLignesMarquees : pour chaque ligne du tableau de recherche ou
'                            "Valider" = "Oui", ecrit Catégorie/SousCategorie et
'                            Notes dans la table SOURCE de la ligne (TblOperations
'                            ou TblVentilations -- voir colonne technique
'                            SourceLigne), puis efface la marque.
'
'   Verrouillage "souple" de Notes : il n'y a pas de vraie protection de
'   feuille (déjà abandonnee ailleurs dans ce chantier a cause d'une erreur
'   1004). La cellule Notes d'une ligne dont la clé est déjà valide est
'   simplement grisee (indication visuelle), et AppliquerLignesMarquees
'   IGNORE tout changement sur cette colonne pour cette ligne, quel que soit
'   ce qui y est ecrit a l'ecran.
'
' PHASE 5 -- CE QUI CHANGE PAR RAPPORT A LA VERSION D'ORIGINE :
'   (a) BUG CORRIGE : la liste déroulante de la colonne Catégorie utilisait
'       Formula1:=Join(listeCategories, ",") -- cela casse des que : (1) une
'       catégorie contient elle-même une virgule (ex: "Alimentation,
'       supermarche"), Excel la scinde alors en plusieurs "fausses"
'       catégories dans la liste ; (2) le texte assemble depasse 255
'       caracteres, ce qui arrive vite avec ~80 catégories (limite d'Excel
'       pour un Formula1 en dur). La correction reutilise la même technique
'       que le reste du chantier (mod_Categories, mod_ControleCategories...) :
'       une PLAGE NOMMEE dynamique (OFFSET/COUNTA) pointant sur une colonne
'       technique de VRAIES cellules, jamais une liste text-jointe.
'   (b) NOUVELLE colonne SousCategorie (même traitement que Catégorie).
'   (c) NOUVELLE colonne Ventile (tag informatif) + intégration des lignes de
'       TblVentilations : une opération ventilee reste desormais accessible
'       ICI de 2 facons, comme demandé par l'opérateur : via sa ligne
'       PARENTE (Catégorie = "Ventile"), ou directement via chacune de ses
'       parts (une ligne par sous-catégorie de la ventilation, marquee
'       Ventile = "Oui").
'
' A PROPOS DES ACCENTS : fichier 100% ASCII, textes accentues via mod_Display.FR().
'
' COMMENT TESTER : Ctrl+G, taper RechercherOperations, Entree.
' =====================================================================================

' --- Noms techniques utilises par ce module, DUPLIQUES ICI EN DUR (comme dans
'     les autres modules de ce chantier) pour que ce module continue a
'     compiler même si mod_InstallVentilation n'a pas encore ete importe. ---
Private Const VEN_FEUILLE As String = "Ventilations"
Private Const VEN_TABLE As String = "TblVentilations"
Private Const CATEGORIE_VENTILE As String = "Ventil" ' + e accentue, voir CategorieVentileRO()

' Mémoire du prefiltre/param actuellement affichés : sert uniquement à ce que
' RevoirVentilationRO (plus bas) puisse recharger l'écran dans le MÊME contexte
' après une modification, plutôt que de retomber en recherche libre.
Private g_ROPrefiltreActif As String
Private g_ROParamActif As Variant

' Nom de la sous-catégorie "santé" -- reprise ici uniquement pour référence
' dans les commentaires, ce module ne teste jamais directement cette valeur.

Private Function CategorieVentileRO() As String
    CategorieVentileRO = CATEGORIE_VENTILE & ChrW(233)   ' "Ventile"
End Function


' =====================================================================================
' HELPERS PHASE 5 : accès a TblVentilations, dupliques comme dans les autres
' modules de ce chantier (mod_SuiviSante, mod_FormulairesNotes, ...).
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
' Memorise les filtres de colonne actifs (fleches natives Excel) sur tblRecherche,
' pour pouvoir les restaurer apres avoir vide et reecrit le tableau -- sinon
' l'operateur perd son filtre a chaque recherche/rafraichissement (constat du
' 01/10/2026 : apres "Revoir la ventilation", la colonne "Ventile" filtree sur "Oui"
' semblait se vider, car Excel garde le souvenir des anciennes POSITIONS de lignes
' visibles, qui ne correspondent plus a rien une fois le tableau reecrit).
' =====================================================================================
Private Function MemoriserFiltresRO(ByVal tbl As ListObject) As Variant
    ' Renvoie un tableau a 2 dimensions (4 lignes : Field/Operator/Criteria1/Criteria2,
    ' une colonne par filtre actif trouve), ou Empty si aucun filtre n'est actif.
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
                    ' Criteria2 ne s'applique qu'a certains types de filtres (ex :
                    ' "entre telle et telle date") -- absent sinon, d'ou le On Error.
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

' Reapplique les filtres precedemment memorises par MemoriserFiltresRO, une fois le
' tableau reconstruit avec les nouvelles donnees.
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
            ' Filtre "simple" (ex : "non vide", critere1 = "<>") : Excel ne renvoie
            ' alors aucun Operator exploitable (0 n'est pas une valeur valide de
            ' XlAutoFilterOperator) -- on l'omet simplement a la reapplication,
            ' sinon Excel refuse silencieusement le filtre (constat du 01/10/2026).
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
' ET TblVentilations
'
' PHASE 6 (refonte "ecran central", apres discussion avec l'operateur) :
'   Ajout d'un parametre optionnel "prefiltre", qui remplace les anciens
'   ecrans "Synthese_*" (desormais supprimes) :
'     ""               : recherche libre, comportement D'ORIGINE inchange
'     "DernierImport"   : uniquement les operations du DERNIER import (liste
'                         memorisee par mod_DernierImport.MemoriserDernierImport)
'     "OperationsDuMois" : uniquement les operations du mois/annee choisis sur
'                         Synthese (remplace mod_SyntheseBugetMensuel)
'     "ErreursSante"    : uniquement les lignes dont la consultation sante est
'                         en erreur (remplace mod_SyntheseExportCareError)
'     "DetailTotal"     : detail d'un total du bilan mensuel ; "param" vaut
'                         alors "Positif" ou "Negatif" (remplace
'                         mod_SyntheseBudgetBilanMensuel.ShowDetailForTotal)
'   Dans tous les cas le tableau reste le MEME (memes colonnes, meme bouton
'   Appliquer) : seules les LIGNES chargees et les COLONNES visibles changent
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

    ' --- PHASE 5 : accès a TblVentilations (facultatif) ---
    Dim tblVen As ListObject
    Dim tblDataVen As Variant
    Dim nbLignesVen As Long
    Dim venDisponible As Boolean
    Dim colVenID As Long, colVenCat As Long, colVenSousCat As Long, colVenMontant As Long, colVenNotes As Long

    ' --- PHASE 6 : preparation propre au prefiltre demande ---
    Dim inclureVentilations As Boolean
    Dim listeDernierImport As Object
    Dim showTypeDetail As String
    Dim indicesOp() As Long, nbIndicesOp As Long
    Dim indicesVen() As Long, nbIndicesVen As Long
    Dim nbTotal As Long

    ' Mémorise le prefiltre actif : reutilise par RevoirVentilationRO pour
    ' recharger l'ecran dans le MÊME contexte apres une modification.
    g_ROPrefiltreActif = prefiltre
    g_ROParamActif = param

    Set ws = ThisWorkbook.Worksheets(mod_InstallRechercheOperations.NOM_FEUILLE_RECHERCHE)
    Set tblRecherche = ws.ListObjects(mod_InstallRechercheOperations.NOM_TABLE_RECHERCHE)

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

    ' --- PHASE 6 : preparation specifique au prefiltre --------------------
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

    ' Les lignes de TblVentilations n'ont pas leur propre mois/annee budgetaire
    ' ni leurs propres colonnes sante : elles ne font sens qu'en recherche libre
    ' ou pour le dernier import (comme avant cette refonte).
    inclureVentilations = (prefiltre = "" Or prefiltre = "DernierImport")

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

    ' Fige l'affichage sous la ligne d'en-tete du tableau (RO_LIGNE_ENTETES),
    ' pour que les intitules de colonnes restent visibles en faisant defiler
    ' une longue liste de resultats -- demande operateur du 01/10/2026.
    ActiveWindow.FreezePanes = False
    ws.Range("A" & (mod_InstallRechercheOperations.RO_LIGNE_ENTETES + 1)).Select
    ActiveWindow.FreezePanes = True

    ' Avant de vider/reconstruire le tableau : on memorise d'abord le(s) filtre(s)
    ' de colonne actif(s) (fleches natives Excel), pour pouvoir les remettre a
    ' l'identique une fois les nouvelles donnees ecrites (voir MemoriserFiltresRO /
    ' RestaurerFiltresRO plus haut).
    Dim filtresSauvegardes As Variant
    filtresSauvegardes = MemoriserFiltresRO(tblRecherche)

    On Error Resume Next
    If tblRecherche.ShowAutoFilter Then
        tblRecherche.AutoFilter.ShowAllData
    End If
    On Error GoTo 0

    ' --- On vide le tableau de recherche (ne garde que l'entete) ---
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

    ' --- PHASE 6 : on repere D'ABORD les lignes retenues par le prefiltre
    ' (deux tableaux d'INDICES, redimensionnes au fur et a mesure), avant de
    ' dimensionner "resultat" exactement a la bonne taille -- plus simple et
    ' plus sur qu'un ReDim Preserve sur un tableau a 2 dimensions (impossible
    ' en VBA sur la 1re dimension). ---
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
    ' 1) Lignes de TblOperations retenues par le prefiltre
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

        ' PHASE 6 : colonnes issues des anciens ecrans "Synthese_*", toujours
        ' remplies (meme si masquees pour ce prefiltre) -- inoffensif, et evite
        ' de relire TblOperations une seconde fois si l'operateur change de
        ' colonnes visibles sans relancer une recherche.
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
    ' 2) Lignes de TblVentilations retenues (recherche libre / dernier import
    '    uniquement -- voir "inclureVentilations" plus haut)
    ' =====================================================================
    If nbIndicesVen > 0 Then
        ' Index ID_Transaction -> ligne TblOperations, pour retrouver le Tiers
        ' parent sans reparcourir tblData pour chaque part ventilee.
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
            resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_DATE) = Empty   ' pas de date propre a une part ventilee
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

            seg0 = mod_ImportOFX.SegmentTexte(notesTexte, ";", 0)
            cleVerrouillee(ligneEcran) = mod_ImportOFX.EstDateValide(seg0)

            AjouterValeurUnique catTexte, valeursVues, listeValeurs, nbValeurs
            AjouterValeurUnique sousTexte, valeursVues, listeValeurs, nbValeurs
        Next k
    End If

    ' --- Redimensionner le tableau puis ecrire en bloc ---
    tblRecherche.Resize tblRecherche.HeaderRowRange.Resize(nbTotal + 1, 16)
    tblRecherche.ListColumns("ID_Transaction").DataBodyRange.NumberFormat = "@"   ' <- ajout : AVANT l'écriture, sinon Excel convertit les longs ID en nombre
    tblRecherche.DataBodyRange.value = resultat

    tblRecherche.ListColumns("Date").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblRecherche.ListColumns("Montant").DataBodyRange.NumberFormat = "#,##0.00"
    tblRecherche.ListColumns("Budget").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblRecherche.ListColumns("Date_consult").DataBodyRange.NumberFormat = "dd/mm/yyyy"

    ' --- PHASE 5 (correction du bug) : liste déroulante Catégorie/SousCategorie
    '     via une VRAIE plage nommee (technique OFFSET/COUNTA), jamais via un
    '     texte joint par des virgules -- voir l'explication en tete de fichier.
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

    ' --- Griser les cellules Notes déjà verrouillées (clé santé valide), et
    ' colorer en vert les montants positifs (PHASE 6 : demande généralisée a
    ' TOUS les prefiltres, plus seulement l'ancien ecran "Budget mensuel") ---
    For i = 1 To nbTotal
        If cleVerrouillee(i) Then
            tblRecherche.ListColumns("Notes").DataBodyRange.Cells(i).Interior.Color = RGB(240, 240, 240)
        Else
            tblRecherche.ListColumns("Notes").DataBodyRange.Cells(i).Interior.ColorIndex = xlColorIndexNone
        End If

        tblRecherche.ListColumns("Montant").DataBodyRange.Cells(i).Font.Color = RGB(0, 0, 0)
        If mod_DataStructure.ToDouble(tblRecherche.ListColumns("Montant").DataBodyRange.Cells(i).value) > 0 Then
            tblRecherche.ListColumns("Montant").DataBodyRange.Cells(i).Font.Color = RGB(0, 128, 0)
        End If
    Next i

    ' --- PHASE 6 : surlignage des N plus grosses dépenses, UNIQUEMENT pour le
    ' prefiltre "OperationsDuMois" (comportement de l'ancien mod_SyntheseBugetMensuel,
    ' l'opérateur a choisi de ne PAS le généraliser aux autres écrans). Repris ici
    ' avec une numérotation propre au tableau de recherche plutôt qu'en réutilisant
    ' mod_Rapports.HighlightTopRows telle quelle : cette dernière suppose une
    ' plage avec une ligne d'en-tête (convention des anciens écrans "Synthese_*"),
    ' ce qui ne correspond pas à tblRecherche.DataBodyRange (pas d'en-tête dedans).
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
' DecrireFiltreActifRO (ajout 02/10/2026) : construit la phrase affichee a la
' ligne RO_LIGNE_FILTRE, juste au-dessus du tableau de recherche, pour que
' l'operateur sache en permanence sur quel sous-ensemble d'operations il
' travaille. Appelee UNIQUEMENT par RechercherOperations, juste apres que ce
' dernier ait fini de remplir le tableau (nbTotal est donc deja connu).
'
' "critMois"/"critAnnee" sont des variables Public declarees dans
' mod_VarGlobales et remplies par mod_Criteres.GetSelectCriteres (deja appelee
' plus haut dans RechercherOperations pour les prefiltres qui en ont besoin) :
' on les relit ici telles quelles, sans les recalculer, exactement comme le
' fait deja mod_SyntheseBudgetBilanMensuel ailleurs dans ce classeur.
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

        Case Else
            ' Securite : un prefiltre pas encore prevu ici ne doit jamais faire
            ' planter l'affichage, juste rester discret.
            texte = nbTotal & FR(" ligne(s) affich{e2}e(s).")

    End Select

    ' Remarque : chaque fragment ci-dessus est deja passe par FR() au moment ou
    ' il est concatene -- "texte" est donc deja la chaine finale, decodee.
    DecrireFiltreActifRO = texte

End Function

' Indique si la ligne "i" de TblOperations (repère DataBodyRange, 1 = 1re
' ligne) doit être incluse pour le prefiltre demandé. Toute la logique de
' filtrage par prefiltre est centralisée ici, pour ne pas la répéter à chaque
' fois qu'un nouveau prefiltre sera ajouté plus tard.
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

        Case Else
            LigneOpRetenuePourPrefiltre = True

    End Select

End Function

' Reproduit le test de l'ancien écran "Erreurs santé" : une consultation est en
' erreur quand sa date vaut le 02/01/1900 (numéro de série Excel = 2), la
' valeur-sentinelle écrite par mod_ImportOFX quand le segment de Notes n'a pas
' pu être transformé en date exploitable (voir DateSerial(1900,1,2) dans
' mod_ImportOFX.ImporterOperationsOFX).
' AMÉLIORATION PAR RAPPORT A L'ANCIEN ÉCRAN (à signaler à l'opérateur) :
' l'ancien code comparait un TEXTE ("02/01/1900") à la valeur de la cellule,
' ce qui ne fonctionne de manière fiable qu'avec certains réglages régionaux.
' Ici on compare le NUMÉRO DE SÉRIE (une date Excel n'est jamais qu'un nombre),
' fiable quels que soient le format d'affichage ou la langue du poste.
Private Function DateConsultInvalideRO(ByVal valeurCellule As Variant) As Boolean
    If IsEmpty(valeurCellule) Then Exit Function
    On Error Resume Next
    DateConsultInvalideRO = (CLng(valeurCellule) = 2)
    On Error GoTo 0
End Function

' Charge la liste des ID_Transaction du dernier import (feuille technique
' TechDernierImport, alimentée par mod_DernierImport.MemoriserDernierImport,
' appelée automatiquement en fin d'import -- voir mod_ImportOFX). Renvoie
' Nothing si aucun import n'a encore été enregistré.
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

' Surligne en rouge les "critNbOperations" plus grosses dépenses parmi les
' lignes actuellement chargées (même idée que l'ancien mod_Rapports.HighlightTopRows,
' mais recodée ici pour travailler directement sur tblRecherche.DataBodyRange,
' qui n'a pas de ligne d'en-tête contrairement à la convention des anciens
' écrans "Synthese_*").
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

    ' Tri par sélection, montants décroissants (même principe que
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

' Petit utilitaire : ajoute "valeur" a listeValeurs()/nbValeurs si non vide et
' pas déjà présente (compare sans tenir compte de la casse : voir .CompareMode
' pose par l'appelant sur le Dictionary utilisé comme "déjà vu").
Private Sub AjouterValeurUnique(ByVal valeur As String, ByRef vues As Object, ByRef listeValeurs() As String, ByRef nbValeurs As Long)
    If valeur = "" Then Exit Sub
    If vues.Exists(valeur) Then Exit Sub
    vues.Add valeur, True
    nbValeurs = nbValeurs + 1
    ReDim Preserve listeValeurs(1 To nbValeurs)
    listeValeurs(nbValeurs) = valeur
End Sub

' Ecrit la liste unique (Catégorie + SousCategorie confondues) dans une
' colonne technique HORS du tableau structure (au-dela de sa dernière
' colonne, ligne par ligne, format Texte force), puis (re)definit une plage
' nommee DYNAMIQUE (OFFSET/COUNTA) pointant dessus. Même principe que
' mod_Categories.RafraichirListesCategories, applique ici localement pour ne
' pas dependre de l'existence de mod_Categories.
Private Sub EcrireListeTechniqueRO(ByVal ws As Worksheet, ByRef listeValeurs() As String, ByVal nbValeurs As Long)

    ' PHASE 6 : deplacee de "N" a "R" -- la colonne N du tableau est desormais
    ' une VRAIE colonne (SoldeSante), la zone technique doit rester au-dela de
    ' la derniere colonne du tableau (P).
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
' AppliquerLignesMarquees : applique Catégorie/SousCategorie/Notes des lignes
' marquees "Oui", en ecrivant dans la table SOURCE de chaque ligne (colonne
' technique SourceLigne : "O" = TblOperations, "V" = TblVentilations).
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

    Set ws = ThisWorkbook.Worksheets(mod_InstallRechercheOperations.NOM_FEUILLE_RECHERCHE)
    Set tblRecherche = ws.ListObjects(mod_InstallRechercheOperations.NOM_TABLE_RECHERCHE)

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

            ' Ajout 02/10/2026 (apres discussion avec l'operateur, option 2) : Categorie
            ' et SousCategorie ne sont plus geres par ce mecanisme -- ils se modifient
            ' desormais par double-clic (voir EditerCategorieRO et RevoirVentilationRO),
            ' qui ecrivent immediatement dans la table source, sans passer par "Valider"
            ' ni par ce bouton. Seul Notes reste gere ici, en lot, comme avant.
            Dim notesValeur As String
            notesValeur = CStr(donnees(i, mod_InstallRechercheOperations.RO_COL_NOTES))

            If sourceLigne = "V" Then
                ' --- PHASE 5 : ligne de TblVentilations, identifiee directement
                '     par sa position (colonne technique LigneVentilation) ---
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
                ' --- Ligne de TblOperations (Notes uniquement, voir commentaire ci-dessus) ---
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
' RevoirVentilationRO (ajout suite a un test operateur)
' =====================================================================================
' Rouvre le formulaire de ventilation pour la ligne actuellement selectionnee dans
' le tableau de recherche (meme principe que ReinitialiserLigneSelectionnee dans
' mod_ResolutionCategories : on lit ActiveCell.Row, pas de bouton par ligne).
'
' La ligne selectionnee peut etre :
'   - une operation PARENTE deja ventilee (Categorie = "Ventile"), ou
'   - l'une de ses PARTS (colonne Ventile = "Oui", Source = "V"),
'   - ou une operation qui n'est PAS ENCORE ventilee.
' Dans les 2 premiers cas, RO_COL_ID contient deja l'ID_Transaction de l'operation
' bancaire PARENTE (voir RechercherOperations, qui l'ecrit ainsi pour tous les types
' de lignes) : c'est cet identifiant qu'attend mod_Ventilation.OuvrirVentilation.
'
' Ajout 02/10/2026 : cette procedure gere maintenant 2 cas, selon que la ligne est
' deja ventilee ou non (voir "etaitDejaVentilee" plus bas) :
'   - deja ventilee  : comportement d'origine, on rouvre le detail existant ;
'   - pas encore ventilee : on demarre une NOUVELLE ventilation (formulaire vide),
'     et si l'operateur va jusqu'au bout, on bascule la categorie de l'operation sur
'     "Ventile" (voir MarquerOperationVentileeRO), exactement comme le fait
'     mod_ControleCategories.ControleVentiler au moment d'un import. Avant ce
'     changement, double-cliquer sur une cellule Ventile vide affichait simplement
'     "rien a revoir" sans rien proposer d'autre.
'
' LIMITE CONNUE : TblVentilations ne conserve pas le libelle brut de l'operation
' bancaire (il n'existait que le temps de l'import, dans un tableau en memoire -
' voir mod_ControleCategories.ControleVentiler). En ouvrant depuis cet ecran,
' bien apres l'import, ce libelle n'est donc jamais disponible : le formulaire
' l'affichera vide plutot que d'inventer une valeur. Rien d'autre n'est affecte :
' date, tiers, montant et categorie/sous-categorie actuelles restent exacts,
' relus directement dans TblOperations.
Public Sub RevoirVentilationRO()

    Dim ws As Worksheet
    Dim tblRecherche As ListObject
    Dim ligneSelection As Long, ligneRelative As Long
    Dim categorieValeur As String, ventileValeur As String
    Dim idTransaction As String
    Dim ok As Boolean
    Dim etaitDejaVentilee As Boolean

    Set ws = ThisWorkbook.Worksheets(mod_InstallRechercheOperations.NOM_FEUILLE_RECHERCHE)
    Set tblRecherche = ws.ListObjects(mod_InstallRechercheOperations.NOM_TABLE_RECHERCHE)

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

    ' --- On relit l'entete (date/tiers/montant/categorie actuelle) directement sur
    ' l'operation PARENTE dans TblOperations : TblVentilations n'a pas ces colonnes,
    ' et OuvrirVentilation en a besoin pour l'affichage en lecture seule. ---
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

' Ajout 01/10/2026 (point 4 : annuler une ventilation, depuis cet ecran) ----------------
' Relit les colonnes techniques CategorieAvantVentilation / SousCategorieAvantVentilation
' (ajoutees par mod_InstallVentilation.AjouterColonnesAnnulationVentilation) de la ligne
' ligneParent dans TblOperations et les recopie dans Categorie / SousCategorie, puisque
' c'est la categorie que l'operation avait AVANT sa toute premiere ventilation (memorisee
' par mod_ControleCategories.ControleVentiler au moment de l'import). On vide ensuite ces
' 2 colonnes techniques : il n'y a plus rien a restaurer tant qu'une nouvelle ventilation
' n'est pas recreee pour cette operation.
' ligneParent : numero de ligne DANS LE TABLEAU (1 = premiere ligne de donnees), au sens
' ou RevoirVentilationRO le calcule plus haut -- PAS un numero de ligne de la feuille.
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

    ' On vide les colonnes techniques : l'operation n'est plus ventilee, il n'y a donc
    ' plus de "categorie d'avant" a conserver pour elle.
    tblOp.DataBodyRange.Cells(ligneParent, colCatAvant).value = ""
    tblOp.DataBodyRange.Cells(ligneParent, colSousAvant).value = ""

End Sub

' Ajout 02/10/2026 (demarrer une ventilation depuis l'ecran de recherche, en plus de
' l'ecran de controle a l'import) : bascule l'operation sur la categorie "Ventile",
' apres avoir memorise son ancienne categorie/sous-categorie dans les colonnes
' CategorieAvantVentilation / SousCategorieAvantVentilation de TblOperations (ajoutees
' par mod_InstallVentilation), pour pouvoir les restaurer si cette ventilation est
' supprimee plus tard (voir RestaurerCategorieAvantVentilation ci-dessus). Meme principe
' que mod_ControleCategories.ControleVentiler au moment d'un import.
' ligneParent : numero de ligne DANS LE TABLEAU (1 = premiere ligne de donnees), au sens
' ou RevoirVentilationRO le calcule plus haut -- PAS un numero de ligne de la feuille.
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

    ' La categorie "Ventile" doit exister dans le referentiel pour que les listes
    ' deroulantes (et un futur controle a l'import) l'acceptent -- meme appel que
    ' mod_ControleCategories.ControleVentiler.
    mod_Categories.AjouterCategoriePersonnalisee CategorieVentileRO(), ""
    mod_Categories.RafraichirListesCategories

    mod_SuiviSante.CalculerSuiviSante AfficherResume:=False

End Sub

' =====================================================================================
' EditerCategorieRO (ajout 02/10/2026)
' =====================================================================================
' Declenchee par un double-clic sur une cellule de la colonne Categorie (voir
' ThisWorkbook.Workbook_SheetBeforeDoubleClick) : reutilise le MEME formulaire que le
' controle des categories a l'import (mod_ControleCategories.ControlerCategories), en
' mode "une seule operation" (son parametre uneSeuleOperation), pour profiter de toutes
' ses fonctionnalites existantes (liste deroulante Sous-categorie dependante, bouton "+"
' de creation d'une nouvelle categorie, etc.) sans dupliquer ce code.
'
' Ne s'applique PAS a une ligne deja ventilee (Ventile = "Oui", ou Categorie =
' "Ventile") : sa categorie se gere depuis le detail de la ventilation (double-clic sur
' Ventile), pas ici -- un message d'information l'explique dans ce cas, pour ne pas
' laisser croire a un bug.
Public Sub EditerCategorieRO()

    Dim ws As Worksheet
    Dim tblRecherche As ListObject
    Dim ligneSelection As Long, ligneRelative As Long
    Dim categorieValeur As String, ventileValeur As String
    Dim idTransaction As String

    Set ws = ThisWorkbook.Worksheets(mod_InstallRechercheOperations.NOM_FEUILLE_RECHERCHE)
    Set tblRecherche = ws.ListObjects(mod_InstallRechercheOperations.NOM_TABLE_RECHERCHE)

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

    ' --- Formulaire de controle des categories, en mode "une seule operation" :
    ' CTRL_OP_CATSOURCE est laisse vide expres (pas de "source banque" ici, voir
    ' ControlerCategories) ; c'est categorieActuelleUnique/sousCategorieActuelleUnique
    ' ci-dessous qui prerempliront le formulaire avec la VRAIE categorie actuelle. ---
    Dim ops(1 To 1, 1 To mod_ControleCategories.CTRL_OP_NBCOL) As Variant
    Dim catFinale() As String, sousFinale() As String
    Dim catAvantVen() As String, sousAvantVen() As String
    ' AJOUT 03/10/2026 : ControlerCategories exige maintenant ces 2 tableaux de sortie
    ' supplementaires (Tiers/Notes modifiables a l'import, voir ce module). Ici on ne s'en
    ' sert pas (TiersNotesEditables les garde figes en mode "une seule operation", voir
    ' mod_ControleCategories) : on les declare juste pour pouvoir passer l'appel.
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

' PHASE 6 : retourne sur la feuille "Resultat" (bilan mensuel) si l'écran a été
' ouvert depuis le détail d'un total, sinon sur Synthèse comme avant.
Public Sub SortirRechercheOperations()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mod_InstallRechercheOperations.NOM_FEUILLE_RECHERCHE)
    On Error Resume Next
    If g_ROPrefiltreActif = "DetailTotal" Then
        ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_RESULTAT).Activate
    Else
        Sheets("Synthese").Activate
    End If
    On Error GoTo 0
    ws.Visible = xlSheetVeryHidden
End Sub
