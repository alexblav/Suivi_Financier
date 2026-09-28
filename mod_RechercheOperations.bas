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
' A PROPOS DES ACCENTS : fichier 100% ASCII, textes accentues via FR().
'
' COMMENT TESTER : Ctrl+G, taper RechercherOperations, Entree.
' =====================================================================================

' --- Noms techniques utilises par ce module, DUPLIQUES ICI EN DUR (comme dans
'     les autres modules de ce chantier) pour que ce module continue a
'     compiler même si mod_InstallVentilation n'a pas encore ete importe. ---
Private Const VEN_FEUILLE As String = "Ventilations"
Private Const VEN_TABLE As String = "TblVentilations"
Private Const CATEGORIE_VENTILE As String = "Ventil" ' + e accentue, voir CategorieVentileRO()

' Nom de la sous-catégorie "santé" -- reprise ici uniquement pour référence
' dans les commentaires, ce module ne teste jamais directement cette valeur.

Private Function CategorieVentileRO() As String
    CategorieVentileRO = CATEGORIE_VENTILE & ChrW(233)   ' "Ventile"
End Function

' Petite fonction miroir de FR(), gardee car déjà utilisée dans ce module a
' l'origine (voir commentaire d'origine sur les accents en tete de fichier) :
'Private Function FR(ByVal texte As String) As String
'    Dim r As String
'    r = texte
'    r = Replace(r, "{e2}", ChrW(233)) ' remplace dans r la chaine "{e2}" par "e" accentue
'    r = Replace(r, "{e1}", ChrW(232))
'    r = Replace(r, "{ea}", ChrW(234))
'    r = Replace(r, "{a2}", ChrW(224))
'    r = Replace(r, "{c2}", ChrW(231))
'    r = Replace(r, "{o2}", ChrW(244))
'    r = Replace(r, "{i2}", ChrW(238))
'    r = Replace(r, "{E2}", ChrW(201))
'    FR = r
'End Function


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
' RechercherOperations : recharge le tableau de recherche depuis TblOperations
' ET TblVentilations
' =====================================================================================
Public Sub RechercherOperations()

    Dim ws As Worksheet
    Dim tblRecherche As ListObject
    Dim n As Long, i As Long, ligneEcran As Long

    Dim resultat() As Variant
    Dim cleVerrouillee() As Boolean
    Dim notesTexte As String, seg0 As String

    Dim valeursVues As Object
    Dim listeValeurs() As String
    Dim nbValeurs As Long

    ' --- PHASE 5 : accès a TblVentilations (facultatif) ---
    Dim tblVen As ListObject
    Dim tblDataVen As Variant
    Dim nbLignesVen As Long
    Dim venDisponible As Boolean
    Dim colVenID As Long, colVenCat As Long, colVenSousCat As Long, colVenMontant As Long, colVenNotes As Long

    Set ws = ThisWorkbook.Worksheets(mod_InstallRechercheOperations.NOM_FEUILLE_RECHERCHE)
    Set tblRecherche = ws.ListObjects(mod_InstallRechercheOperations.NOM_TABLE_RECHERCHE)

    Set tbl = mod_DonneesTable.GetOperationsTable()
    If tbl Is Nothing Then
        MsgBox FR("Le tableau TblOperations est introuvable."), vbExclamation
        Exit Sub
    End If
    If tbl.DataBodyRange Is Nothing Then
        MsgBox FR("TblOperations ne contient aucune ligne."), vbInformation
        Exit Sub
    End If

    mod_Display.RecupIndexCol
    tblData = tbl.DataBodyRange.value
    n = UBound(tblData, 1)

    ' --- PHASE 5 : chargement de TblVentilations, si disponible ---
    venDisponible = False
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

    ws.Visible = xlSheetVisible
    ws.Activate

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

    Dim nbTotal As Long
    nbTotal = n + IIf(venDisponible, nbLignesVen, 0)

    If nbTotal = 0 Then
        MsgBox FR("Aucune op{e2}ration {a2} afficher."), vbInformation
        Exit Sub
    End If

    ReDim resultat(1 To nbTotal, 1 To 11)
    ReDim cleVerrouillee(1 To nbTotal)

    Set valeursVues = CreateObject("Scripting.Dictionary")
    valeursVues.CompareMode = 1
    nbValeurs = 0

    ' --- Petit utilitaire local (via GoSub serait plus lourd ici) : on ecrit
    '     directement en boucle, voir plus bas AjouterValeurUnique ---

    ligneEcran = 0

    ' =====================================================================
    ' 1) Lignes de TblOperations (comportement d'origine, inchange, plus la
    '    colonne SousCategorie et le tag Ventile)
    ' =====================================================================
    For i = 1 To n
        ligneEcran = ligneEcran + 1

        resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_VALIDER) = ""
        resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_DATE) = tblData(i, colDate)
        resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_TIERS) = mod_DataStructure.CellText(tblData(i, colTiers))
        resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_MONTANT) = tblData(i, colMontant)

        Dim catTexte As String, sousTexte As String
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

        seg0 = mod_ImportOFX.SegmentTexte(notesTexte, ";", 0)
        cleVerrouillee(ligneEcran) = mod_ImportOFX.EstDateValide(seg0)

        AjouterValeurUnique catTexte, valeursVues, listeValeurs, nbValeurs
        AjouterValeurUnique sousTexte, valeursVues, listeValeurs, nbValeurs
    Next i

    ' =====================================================================
    ' 2) PHASE 5 : lignes de TblVentilations (une ligne par part ventilee).
    '    Le Tiers est repris de l'opération PARENTE (TblVentilations n'a pas
    '    sa propre colonne Tiers -- une part ventilee n'est pas une opération
    '    bancaire indépendante).
    ' =====================================================================
    If venDisponible Then
        ' Index ID_Transaction -> ligne TblOperations, pour retrouver le Tiers
        ' parent sans reparcourir tblData pour chaque part ventilee.
        Dim indexParIDPourTiers As Object
        Dim k As Long
        Set indexParIDPourTiers = CreateObject("Scripting.Dictionary")
        indexParIDPourTiers.CompareMode = 1
        For k = 1 To n
            indexParIDPourTiers(mod_DataStructure.CellText(tblData(k, colID))) = k
        Next k

        Dim idParentVen As String, tiersParentVen As String
        For i = 1 To nbLignesVen
            ligneEcran = ligneEcran + 1

            idParentVen = mod_DataStructure.CellText(tblDataVen(i, colVenID))
            tiersParentVen = ""
            If indexParIDPourTiers.Exists(idParentVen) Then
                tiersParentVen = mod_DataStructure.CellText(tblData(indexParIDPourTiers(idParentVen), colTiers))
            End If

            resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_VALIDER) = ""
            resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_DATE) = Empty   ' pas de date propre a une part ventilee
            resultat(ligneEcran, mod_InstallRechercheOperations.RO_COL_TIERS) = tiersParentVen & FR(" [ventilation]")
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
        Next i
    End If

    ' --- Redimensionner le tableau puis ecrire en bloc ---
    tblRecherche.Resize tblRecherche.HeaderRowRange.Resize(nbTotal + 1, 11)
    tblRecherche.DataBodyRange.value = resultat

    tblRecherche.ListColumns("Date").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblRecherche.ListColumns("Montant").DataBodyRange.NumberFormat = "#,##0.00"

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

    ' --- Griser les cellules Notes déjà verrouillées (clé santé valide) ---
    For i = 1 To nbTotal
        If cleVerrouillee(i) Then
            tblRecherche.ListColumns("Notes").DataBodyRange.Cells(i).Interior.Color = RGB(240, 240, 240)
        Else
            tblRecherche.ListColumns("Notes").DataBodyRange.Cells(i).Interior.Color = RGB(255, 255, 255)
        End If
    Next i

    MsgBox nbTotal & FR(" op{e2}ration(s)/part(s) charg{e2}e(s) (dont ") & IIf(venDisponible, nbLignesVen, 0) & _
           FR(" ligne(s) de ventilation). Utilise les fl{e2}ches de filtre dans l'en-t{ea}te pour restreindre la liste."), _
           vbInformation

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

    Const COL_TECHNIQUE As String = "N"   ' au-dela de la colonne K (dernière colonne du tableau)
    Dim i As Long

    ws.Range(COL_TECHNIQUE & "1:" & COL_TECHNIQUE & "1000").ClearContents
    ws.Range(COL_TECHNIQUE & "1:" & COL_TECHNIQUE & "1000").NumberFormat = "@"
    ws.Columns(COL_TECHNIQUE).Hidden = True

    ws.Range(COL_TECHNIQUE & "1").value = FR("Liste technique (Cat{e2}gorie+SousCat{e2}gorie) - ne pas modifier")

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
        MsgBox FR("Aucune ligne {a2} traiter. Utilise d'abord 'Rechercher'."), vbInformation
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

            Dim categorieValeur As String, sousCategorieValeur As String, notesValeur As String
            categorieValeur = Trim(CStr(donnees(i, mod_InstallRechercheOperations.RO_COL_CATEGORIE)))
            sousCategorieValeur = Trim(CStr(donnees(i, mod_InstallRechercheOperations.RO_COL_SOUSCATEGORIE)))
            notesValeur = CStr(donnees(i, mod_InstallRechercheOperations.RO_COL_NOTES))

            If sourceLigne = "V" Then
                ' --- PHASE 5 : ligne de TblVentilations, identifiee directement
                '     par sa position (colonne technique LigneVentilation) ---
                If venDisponible Then
                    Dim ligneVenCible As Long
                    ligneVenCible = CLng(donnees(i, mod_InstallRechercheOperations.RO_COL_LIGNEVEN))

                    tblVen.ListColumns("Categorie").DataBodyRange.rows(ligneVenCible).value = categorieValeur
                    tblVen.ListColumns("SousCategorie").DataBodyRange.rows(ligneVenCible).value = sousCategorieValeur

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
                ' --- Ligne de TblOperations (comportement d'origine, plus SousCategorie) ---
                Dim idCible As String
                idCible = Trim(CStr(donnees(i, mod_InstallRechercheOperations.RO_COL_ID)))

                If indexParID.Exists(idCible) Then
                    Dim ligneCible As Long
                    ligneCible = indexParID(idCible)

                    tbl.ListColumns("Categorie").DataBodyRange.rows(ligneCible).value = categorieValeur
                    If colSousCategorie <> 0 Then
                        tbl.ListColumns("SousCategorie").DataBodyRange.rows(ligneCible).value = sousCategorieValeur
                    End If

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

    MsgBox nbAppliquees & FR(" ligne(s) appliqu{e2}e(s)."), vbInformation

End Sub

Public Sub SortirRechercheOperations()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mod_InstallRechercheOperations.NOM_FEUILLE_RECHERCHE)
    Sheets("Synthese").Activate
    ws.Visible = xlSheetVeryHidden
End Sub
