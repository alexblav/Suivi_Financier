Option Explicit

' =====================================================================================
' MODULE : mod_RechercheOperations
'
' ROLE (PHASE 5b) : logique de la feuille frm_RechercheOperations construite
' en Phase 5a (mod_InstallRechercheOperations).
'
'   RechercherOperations   : recharge TOUTE TblOperations dans le tableau de
'                            recherche (filtre ensuite avec les fleches
'                            natives Excel dans l'entete du tableau).
'   AppliquerLignesMarquees : pour chaque ligne du tableau de recherche ou
'                            "Valider" = "Oui", ecrit Categorie et Notes
'                            dans TblOperations (Notes uniquement si elle
'                            n'etait pas deja une cle sante valide au
'                            depart), puis efface la marque.
'
'   Verrouillage "souple" de Notes : il n'y a pas de vraie protection de
'   feuille (deja abandonnee ailleurs dans ce chantier a cause d'une erreur
'   1004). La cellule Notes d'une ligne dont la cle est deja valide est
'   simplement grisee (indication visuelle), et AppliquerLignesMarquees
'   IGNORE tout changement sur cette colonne pour cette ligne, quel que soit
'   ce qui y est ecrit a l'ecran.
'
' A PROPOS DES ACCENTS : fichier 100% ASCII, textes accentues via FR().
'
' COMMENT TESTER : Ctrl+G, taper RechercherOperations, Entree.
' =====================================================================================

' =====================================================================================
' RechercherOperations : recharge le tableau de recherche depuis TblOperations
' =====================================================================================
Public Sub RechercherOperations()

    Dim reponse As VbMsgBoxResult
    Dim debPlageTravail As Range
    Dim nbLigne As Long, idxRes As Long
    'Dim ws As Worksheet
    'Dim tblRecherche As ListObject
    Dim n As Long, i As Long

    Dim resultat() As Variant
    Dim cleVerrouillee() As Boolean
    Dim notesTexte As String, seg0 As String, catTexte As String

    Dim categoriesVues As Object
    Dim listeCategories() As String
    Dim nbCategories As Long

    ' 1. Récupération des pointeurs vers la feuille et le tableau
    Set wsSynthese = mod_Criteres.GetFeuille(mod_VarGlobales.NOM_FEUILLE_SYNTHESE)
    Set wsResultat = mod_Criteres.GetFeuille(mod_VarGlobales.NOM_FEUILLE_RESULTAT)
    Set tbl = mod_DonneesTable.GetOperationsValue(mod_VarGlobales.NOM_FEUILLE_DONNEES, "TblOperations")
    If tbl Is Nothing Then
        MsgBox "Le tableau ne contient aucune ligne de données.", vbExclamation
        Exit Sub
    End If
    If tbl.DataBodyRange Is Nothing Then
        MsgBox "Le tableau est vide.", vbExclamation
        Exit Sub
    End If
    
    ' 1. Verifier si la feuille de travail existe deja, pour eviter d'ecraser du travail sans prevenir.
    If Not wsResultat Is Nothing Then
        ' La feuille existe deja : on demande confirmation avant de tout reconstruire,
        ' car cela va effacer sa mise en forme actuelle.
        reponse = MsgBox(FR("La feuille '" & mod_VarGlobales.NOM_FEUILLE_RESULTAT & "' existe deja." & vbCrLf & _
                            "Voulez-vous la reconstruire enti{e1}rement (sa mise en forme actuelle sera perdue) ?"), _
                            vbYesNo + vbQuestion, "Confirmation de reconstruction")
        If reponse = vbNo Then
            MsgBox FR("Installation annul{e2}e, aucune modification effectu{e2}e."), vbInformation
            Exit Sub
        End If
        
        ' On la rend visible temporairement : impossible de la modifier/supprimer
        ' proprement tant qu'elle est en xlSheetVeryHidden.
        wsResultat.Visible = xlSheetVisible
        wsResultat.Cells.Clear
        Call mod_Display.SupprimerFormesExistantesFN(wsResultat)
        Call mod_Display.SupprimerNomsExistantsFN(wsResultat, mod_VarGlobales.NOM_FEUILLE_RESULTAT)
    Else
        ' La feuille n'existe pas encore : on la cree, positionnee en derniere position
        ' pour ne pas perturber l'ordre des onglets existants (Accueil, Synthese...).
        Set wsResultat = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        wsResultat.Name = mod_VarGlobales.NOM_FEUILLE_RESULTAT
    End If
        
    ' On masque la feuille Synthese. Cette option est prise pour éviter à l'opérateur de se déplacer en dehors de la feuille crée
    wsSynthese.Visible = xlSheetVeryHidden
    
    ' 2. Charge TOUT le tableau de données (en-têtes incluses en ligne 1)
    tblData = tbl.Range.value
    tblDataLineTotal = UBound(tblData, 1)
    
    ' 3. Récupération des index de colonne dans la base de données
    mod_Display.RecupIndexCol
    
    ' 4. Autorise ou non le double click
    AllowDetailDoubleClick = False
    ' Paramètre de navigation
    RecherOperations = True
    ' Désactive les événements
    Application.EnableEvents = False
    
    ' 5. Construction de la zone des boutons
    ' On détermine la position du bouton
    Set zoneBouton = wsResultat.Range(cellSortieDep).Offset(0, 1)
    'On définit son titre
    texteBouton = "Sortir"
    nomMacroBouton = "Sortir"
    Call mod_Display.ConstruireBoutons(wsResultat, zoneBouton, texteBouton, nomMacroBouton)

    ' On détermine la plage de début et de fin de la zone
    Set debutZone = wsResultat.Range(cellSortieDep).Offset(0, 2)
    Set finZone = wsResultat.Range(cellSortieDep).Offset(0, 3)
    Set zoneBouton = wsResultat.Range(debutZone, finZone)
    
    'On définit son titre
    texteBouton = FR("Appliquer les lignes marqu{e2}es")
    nomMacroBouton = "AppliquerLignesMarquees"
    Call mod_Display.ConstruireBoutons(wsResultat, zoneBouton, texteBouton, nomMacroBouton)
    
    ' On détermine la plage de début et de fin de la zone
    Set debutZone = wsResultat.Range(cellSortieDep).Offset(0, 4)
    Set finZone = wsResultat.Range(cellSortieDep).Offset(0, 5)
    Set zoneBouton = wsResultat.Range(debutZone, finZone)
    
'    'On définit son titre
'    texteBouton = FR("Rechercher")
'    nomMacroBouton = "RechercherOperations"
'    Call mod_Display.ConstruireBoutons(wsResultat, zoneBouton, texteBouton, nomMacroBouton)
'
    ' 6. Construction de la zone des instructions
    ' On fournit le message à afficher en remplaçant les carractères accentué par les balise de la fonction FR
    Message = " Recherche Opérations" & Chr(10) & _
                "Utilisez les fl{e2}ches de filtre dans l'en-t{ea}te (comme un filtre Excel classique) pour restreindre la liste. " & _
                    "Corrigez Cat{e2}gorie/Notes directement dans les cellules, inscrivez 'Oui' dans Valider, puis cliquez sur 'Appliquer les lignes marqu{e2}es'."
                 
    ' On fournit le titre, la position de la cellule dans laquelle on veut écrire de titre
    Titre = "Instructions"
    Set PositionTitre = wsResultat.Range(cellSortieDep).Offset(2, 1)
    
    ' On détermine la plage de début et de fin de la zone
    Set debutZone = wsResultat.Range(cellSortieDep).Offset(3, 1)
    Set finZone = wsResultat.Range(cellSortieDep).Offset(6, 6)
    
    Call mod_Display.ConstruireZoneTexte(wsResultat, Titre, Message, PositionTitre, debutZone, finZone)
    
    ' 7. Affiche les entêtes du tableau de sortie
    ' Définit le début de la page de travail
    Set debPlageTravail = wsResultat.Range(cellSortieDep).Offset(8, 1)
    
    MonArray = Array("Valider", "Date", "Tiers", "Montant", "Catégorie", "Notes", "ID_Transaction")
    Call mod_Display.PrepareOutputArea(wsResultat, MonArray, debPlageTravail)
    
    nbColonne = UBound(MonArray) + 1 - LBound(MonArray) + 1
    
    ' Récupére la position du champs dans l'ARRAY (Application.Match est naturellement insensible à la casse)
    mod_Display.RecupPosArray
    
    ' Récupération des index de colonne dans la feuille de sortie
    Call mod_Display.RecupPosSortieIndex(wsResultat, debPlageTravail)
    
    ' Taille maximale du tableau de résultat = nombre total de lignes source
    ReDim tabResultat(1 To tblDataLineTotal, 1 To nbColonne)
           
    ' 4. Lire les critères saisis par l'utilisateur (B1 à B6).
    mod_Criteres.GetSelectCriteres

    ' Déclaration du dictionnaire des catégories
    Set categoriesVues = CreateObject("Scripting.Dictionary")
    nbCategories = 0
    ReDim cleVerrouillee(1 To tblDataLineTotal)
    
    ' 5. Remplissage du tableau de sortie
    For nbLigne = 2 To tblDataLineTotal
        ' On récupére la totalité des lignes
        idxRes = idxRes + 1
        'On écrit la ligne dans le tableau
        tabResultat(idxRes, 1) = ""
        tabResultat(idxRes, posDate) = tblData(nbLigne, colDate)
        tabResultat(idxRes, posTiers) = tblData(nbLigne, colTiers)
        tabResultat(idxRes, posCategorie) = tblData(nbLigne, colCategorie)
        tabResultat(idxRes, posMontant) = tblData(nbLigne, colMontant)
        tabResultat(idxRes, posNotes) = tblData(nbLigne, colNotes)
        tabResultat(idxRes, posID) = tblData(nbLigne, colID)
        
        ' On contrôle que le champ Notes ne contient pas une clé valide de santé
        notesTexte = mod_DataStructure.CellText(tblData(nbLigne, colNotes))
        seg0 = mod_ImportOFX.SegmentTexte(notesTexte, ";", 0)
        If IsNumeric(seg0) Then
            cleVerrouillee(nbLigne) = mod_ImportOFX.EstDateValide(seg0)
        End If
        
        catTexte = Replace(tblData(nbLigne, colCategorie), ",", ";")
        If catTexte <> "" Then
            If Not categoriesVues.Exists(catTexte) Then
                categoriesVues.Add catTexte, True
                nbCategories = nbCategories + 1
                ReDim Preserve listeCategories(1 To nbCategories)
                listeCategories(nbCategories) = catTexte
            End If
        End If
    Next nbLigne

    ' 6. Affichage
    If idxRes > 0 Then
        Application.ScreenUpdating = False
        
        ' Injection directe du tableau mémoire dans la plage d'affichage
        ' On redimentionne la taille de la plage pour pas voir s'afficher des erreur type #N/A dans les cellules en trop
        plageSortieEcriture.Resize(idxRes, nbColonne).value = tabResultat
        
        ' On nomme la plage de sartie
        plageSortieEcriture.Resize(idxRes, nbColonne).Name = mod_VarGlobales.NOM_TABLE_RECHERCHE
        
        ' --- Liste deroulante Categorie : avertissement, pas de blocage (on peut
        '     taper une categorie qui n'existe pas encore) ---
        'With tblRecherche.ListColumns("Categorie").DataBodyRange.Validation
        Dim plage As Range
        Set plage = plageSortieEcriture.Resize(idxRes, nbColonne)
        With plageSortieEcriture.Resize(idxRes, nbColonne).Columns(posCategorie).Validation
            .Delete
            If nbCategories > 0 Then
                .Add Type:=xlValidateList, _
                AlertStyle:=xlValidAlertWarning, _
                Formula1:=Join(listeCategories, ",")
            End If
        End With
        With plageSortieEcriture.Resize(idxRes, nbColonne).Columns(1).Validation
            ' 1. Toujours effacer les validations existantes avant d'en ajouter une nouvelle
            .Delete
    
            ' 2. Ajouter la liste déroulante OUI,NON
            .Add Type:=xlValidateList, _
                AlertStyle:=xlValidAlertStop, _
                Formula1:="OUI,NON"
            ' 3. Configurer l'affichage et autoriser la valeur vide par défaut
            .IgnoreBlank = True          ' Permet de laisser la cellule vide
            .InCellDropdown = True       ' Affiche la flèche de sélection dans la cellule
            .ShowInput = True            ' Affiche le message de saisie si configuré
            .ShowError = True            ' Affiche l'alerte en cas d'erreur de saisie
        End With
            
        ' Mise en forme rapide des colonnes
        mod_Display.MiseEnPage wsResultat, idxRes
        
        ' --- Griser les cellules Notes deja verrouillees (cle sante valide) ---
        For i = 1 To tblDataLineTotal
            If cleVerrouillee(i) Then
                'tblRecherche.ListColumns("Notes").DataBodyRange.Cells(i).Interior.Color = RGB(240, 240, 240)
                plageSortieEcriture.Resize(idxRes, nbColonne).Columns(posNotes).Cells(i - 1).Interior.Color = coulBloc
            End If
        Next i

        
        ' Colonne technique masquee
        wsResultat.Columns(posSortieID).Hidden = True
        wsResultat.Columns(1).ColumnWidth = 1
        
         ' Réactive les événements
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If
    MsgBox idxRes & FR(" op{e2}ration(s) charg{e2}e(s). Utilise les fl{e2}ches de filtre dans l'en-t{ea}te pour restreindre la liste."), vbInformation

End Sub


' =====================================================================================
' AppliquerLignesMarquees : applique Categorie/Notes des lignes marquees "Oui"
' =====================================================================================
Public Sub AppliquerLignesMarquees()

    ' --- Variables LOCALES uniquement : aucune ne porte le nom d'une variable globale
    '     de mod_VarGlobales (sinon elle la "masquerait" silencieusement) ---
    Dim plageTable As Range           ' zone des donnees de TblOperations (sans l'entete)
    Dim tabEcran As Variant           ' copie memoire de la zone affichee a l'ecran
    Dim nbLignesEcran As Long         ' nombre de lignes de cette copie memoire
    Dim indexParID As Object          ' dictionnaire : ID_Transaction -> ligne dans tblData
    Dim idTexte As String             ' un ID lu dans tblData (construction du dictionnaire)
    Dim idCible As String             ' l'ID de la ligne marquee, lu a l'ecran
    Dim i As Long                     ' ligne courante DANS la zone ecran (1 = 1ere operation)
    Dim ligneTbl As Long              ' ligne courante DANS tblData (1 = entete, 2 = 1ere operation)
    Dim ligneCible As Long            ' ligne de la marquee dans tblData (meme repere que ligneTbl)
    Dim categorieValeur As String
    Dim notesOrigine As String, seg0 As String
    Dim verrouillee As Boolean
    Dim nbAppliquees As Long

    ' 1. Garde-fous : les variables globales existent-elles encore ?
    '    (elles sont vides si aucune recherche n'a ete lancée, ou après un "Reset" de VBA)
    If tbl Is Nothing Then
        Set tbl = mod_DonneesTable.GetOperationsValue(mod_VarGlobales.NOM_FEUILLE_DONNEES, "TblOperations")
        If tbl Is Nothing Then
            MsgBox "Le tableau ne contient aucune ligne de données.", vbExclamation
            Exit Sub
        End If
        If tbl.DataBodyRange Is Nothing Then
            MsgBox "Le tableau est vide.", vbExclamation
            Exit Sub
        End If
    End If
    If plageSortieEcriture Is Nothing Then
        MsgBox FR("Aucune recherche n'est affich{e2}e. Cliquez d'abord sur 'Rechercher'."), vbExclamation
        Exit Sub
    End If
    
    ' Si une erreur survient, on passe par GestionErreur pour REACTIVER les evenements
    ' (sans cela Excel resterait "sourd" aux modifications de cellules)
    On Error GoTo GestionErreur
    
    ' 2. Relecture FRAICHE de la table (meme lecture que RechercherOperations : entete en ligne 1)
    '    On NE se fie PAS au contenu actuel de la globale tblData : d'autres macros
    '    (VerifierNotesSante...) la rechargent SANS la ligne d'entete, ce qui decalerait
    '    toutes les lignes d'un cran. La relire ici la rend fiable et a jour.
    tblData = tbl.Range.value
    tblDataLineTotal = UBound(tblData, 1)
    
    Set plageTable = tbl.DataBodyRange
    
    ' 3. Recuperation des index de colonne dans la base de donnees
    mod_Display.RecupIndexCol
    
    ' 4. Lecture de la zone ecran EN BLOC : 1 seul echange avec la feuille, ensuite tout
    '    se passe en memoire. Les lignes masquees par un filtre sont bien incluses.
    tabEcran = plageSortieEcriture.value
    nbLignesEcran = UBound(tabEcran, 1)
    
    ' 5. Index ID_Transaction -> ligne de tblData, construit UNE SEULE FOIS
    '    (plutot qu'un balayage complet de la table pour chaque ligne marquee).
    '    Les lignes SANS ID sont exclues : sinon toutes les lignes vides partageraient la
    '    meme cle "" et une ligne marquee sans ID serait appliquee sur la DERNIERE ligne vide.
    Set indexParID = CreateObject("Scripting.Dictionary")
    For ligneTbl = 2 To tblDataLineTotal
        idTexte = mod_DataStructure.CellText(tblData(ligneTbl, colID))
        If idTexte <> "" Then indexParID(idTexte) = ligneTbl
    Next ligneTbl
    

    ' 6. Application des lignes marquees
    '    Evenements coupes : nos ecritures ne doivent pas declencher Workbook_SheetChange
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    nbAppliquees = 0

    For i = 1 To nbLignesEcran
        ' Une ligne est marquee si sa colonne "Valider" contient OUI (majuscules ou non)
        If LCase$(mod_DataStructure.CellText(tabEcran(i, posValider))) = "oui" Then

            ' On retrouve la ligne d'origine grace a son ID_Transaction
            idCible = mod_DataStructure.CellText(tabEcran(i, posID))

            If indexParID.Exists(idCible) Then
                ligneCible = indexParID(idCible)

                ' Categorie : toujours appliquee
                ' (ligneCible - 1 car plageTable n'a pas de ligne d'entete, contrairement a tblData)
                categorieValeur = mod_DataStructure.CellText(tabEcran(i, posCategorie))
                plageTable.Cells(ligneCible - 1, colCategorie).value = categorieValeur

                ' Notes : verrouillage sur la valeur ORIGINALE (celle de TblOperations avant
                ' cette application), pas sur ce qui est affiche a l'ecran. C'est le meme test
                ' que celui de RechercherOperations qui grise la cellule.
                notesOrigine = mod_DataStructure.CellText(tblData(ligneCible, colNotes))
                seg0 = mod_ImportOFX.SegmentTexte(notesOrigine, ";", 0)
                verrouillee = mod_ImportOFX.EstDateValide(seg0)

                If Not verrouillee Then
                    plageTable.Cells(ligneCible - 1, colNotes).value = CStr(tabEcran(i, posNotes))
                End If

                ' On efface la marque : IMPORTANT de le faire AVANT les recalculs plus bas,
                ' sinon Workbook_SheetDeactivate reproposerait d'appliquer ces memes lignes
                plageSortieEcriture.Cells(i, posValider).value = ""

                nbAppliquees = nbAppliquees + 1
            End If
        End If
    Next i

    ' On remet les evenements AVANT les recalculs : les feuilles de rapprochement ouvertes par
    ' VerifierNotesSante ont besoin des evenements pour fonctionner
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ' 7. Recalcul du suivi sante (uniquement si quelque chose a ete applique)
    If nbAppliquees > 0 Then
        mod_FormulairesNotes.VerifierNotesSante
        mod_SuiviSante.CalculerSuiviSante AfficherResume:=False
        MsgBox nbAppliquees & FR(" ligne(s) appliqu{e2}e(s)."), vbInformation
    End If

    Exit Sub

GestionErreur:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox FR("Erreur pendant l'application des lignes marqu{e2}es :") & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical
End Sub

