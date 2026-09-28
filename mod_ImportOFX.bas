Attribute VB_Name = "mod_ImportOFX"
Option Explicit

' ============================================================================
'  MODULE : mod_ImportOFX
'  ROLE   : Importer les opérations bancaires depuis un fichier OFX ET un
'           fichier CSV complementaire (mêmes opérations, exportes ensemble),
'           les recouper pour récupérer la Catégorie, calculer les colonnes
'           budgetaires, puis ecrire le tout dans la feuille "data_import"
'           sous forme de Tableau Excel trie et filtrable.
'
'  En cas de catégorie ambigue (plusieurs valeurs possibles pour une même
'  opération), le formulaire frmResolutionCategories s'ouvre pour laisser
'  l'opérateur choisir -- sans qu'il ait besoin d'acceder au tableau.
'
'  A appeler depuis le bouton de la feuille "Accueil".
' ============================================================================

' --- Position des colonnes dans la zone de sortie (pour eviter les "magic numbers") ---
Private Const COL_ID As Integer = 1
Private Const COL_DATE_COMPTABLE As Integer = 2
Private Const COL_DATE_OPERATION As Integer = 3
Private Const COL_TYPE As Integer = 4
Private Const COL_TIERS As Integer = 5
Private Const COL_LIBELLE As Integer = 6
Private Const COL_MONTANT As Integer = 7
Private Const COL_NUM_CHEQUE As Integer = 8
Private Const COL_CATEGORIE As Integer = 9
Private Const COL_BUDGET As Integer = 10
Private Const COL_MOIS_BUDGET As Integer = 11
Private Const COL_ANNEE_BUDGET As Integer = 12
Private Const COL_DATE_CONSULT As Integer = 13
Private Const COL_SPE_CONSULT As Integer = 14
Private Const COL_COM_SANTE As String = 18
Private Const NB_COLONNES As Integer = 20

Private Const NOM_TABLE As String = "TblOperations"

' --- Variables PUBLIQUES partagees avec le formulaire frmResolutionCategories ---
' (rempli par ce module avant l'ouverture du formulaire, relu après sa fermeture)
Public g_NbCasAmbigus As Long
Public g_CasTexte() As String       ' texte affiche dans la liste (Date | Montant | Tiers)
Public g_CasCandidats() As String   ' catégories candidates, separees par ";"
Public g_CasChoix() As String       ' rempli par le formulaire : catégorie choisie (ou "" si non traité)


' ----------------------------------------------------------------------------
' MACRO PRINCIPALE - a assigner au bouton de la feuille Accueil
' ----------------------------------------------------------------------------
Public Sub ImporterOperationsOFX()

    Dim cheminOFX As String, cheminCSV As String
    Dim xmlDoc As Object, listeTransactions As Object, noeud As Object
    Dim wsDonnees As Worksheet
    Dim dictFITID As Object, dictCategoriesCSV As Object
    Dim resultats() As Variant, arrFinal() As Variant
    Dim lignesAmbigues() As Long, texteAmbigus() As String, candidatsAmbigus() As String
    Dim nbAmbigus As Long
    Dim fitid As String
    Dim i As Long, k As Long
    Dim nbLues As Long, nbDoublons As Long, nbAjoutees As Long, nbSansCategorie As Long

    ' --- ETAPE 1 : Selection des deux fichiers a importer ---------------------
    cheminOFX = ChoisirFichier("Fichiers OFX (*.ofx), *.ofx", "Selectionner l'export bancaire OFX")
    If cheminOFX = "" Then Exit Sub

    cheminCSV = ChoisirFichier("Fichiers CSV (*.csv), *.csv", _
        "Selectionner le fichier CSV complementaire (memes operations)")
    If cheminCSV = "" Then Exit Sub

    ' --- ETAPE 2 : Chargement du fichier OFX comme document XML ---------------
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.validateOnParse = False
    If Not xmlDoc.Load(cheminOFX) Then
        MsgBox "Le fichier OFX n'a pas pu etre lu." & vbCrLf & xmlDoc.parseError.reason, vbCritical
        Exit Sub
    End If
    Set listeTransactions = xmlDoc.getElementsByTagName("STMTTRN")
    nbLues = listeTransactions.Length
    If nbLues = 0 Then
        MsgBox "Aucune transaction trouvee dans le fichier OFX.", vbExclamation
        Exit Sub
    End If

    ' --- ETAPE 3 : Chargement du CSV -> dictionnaire des catégories par clé ---
    ' Clé = Date + Montant + Tiers. Une même clé peut avoir plusieurs catégories
    ' candidates (cas de plusieurs opérations identiques le même jour).
    Set dictCategoriesCSV = ChargerCategoriesCSV(cheminCSV)
    If dictCategoriesCSV Is Nothing Then Exit Sub  ' l'erreur a déjà ete affichee

    ' --- ETAPE 4 : Preparer la feuille data_import et les FITID déjà connus ---
    Set wsDonnees = ThisWorkbook.Worksheets("Import_data")
    Set dictFITID = CreateObject("Scripting.Dictionary")

    Dim derniereLigne As Long
    derniereLigne = wsDonnees.Cells(wsDonnees.rows.count, "A").End(xlUp).Row
    If derniereLigne >= 2 Then
        Dim plageFITID As Variant
        plageFITID = wsDonnees.Range("A2:A" & derniereLigne).Value2
        For i = 1 To UBound(plageFITID, 1)
            If Not dictFITID.Exists(CStr(plageFITID(i, 1))) Then
                dictFITID.Add CStr(plageFITID(i, 1)), True
            End If
        Next i
    End If

    ' --- ETAPE 5 : Parcourir les transactions OFX -----------------------------
    ' A ce stade, on ne remplit QUE les colonnes 1 a 9 (ID .. Catégorie).
    ' Les colonnes calculees 10 a 14 (Budget, MoisBudget, ...) dependent de la
    ' Catégorie : elles sont calculees plus loin, une fois les cas ambigus resolus.
    ReDim resultats(1 To nbLues, 1 To NB_COLONNES)
    ReDim lignesAmbigues(1 To nbLues)
    ReDim texteAmbigus(1 To nbLues)
    ReDim candidatsAmbigus(1 To nbLues)

    Dim dateComptable As Date, dateOperation As Variant, tiers As String
    Dim libelleDetail As String, Montant As Double, numCheque As Variant
    Dim categorie As String, cle As String, estAmbigu As Boolean, texteCandidats As String

    For i = 0 To listeTransactions.Length - 1
        Set noeud = listeTransactions.Item(i)
        fitid = LireValeurNoeud(noeud, "FITID")

        If fitid <> "" And Not dictFITID.Exists(fitid) Then

            ' --- champs bruts issus de l'OFX ---
            dateComptable = ConvertirDateOFX(LireValeurNoeud(noeud, "DTPOSTED"))
            dateOperation = ConvertirDateOFXVariant(LireValeurNoeud(noeud, "DTUSER"))
            tiers = LireValeurNoeud(noeud, "NAME")
            libelleDetail = LireValeurNoeud(noeud, "MEMO")
            Montant = ConvertirMontantOFX(LireValeurNoeud(noeud, "TRNAMT"))
            numCheque = NettoyerNumCheque(LireValeurNoeud(noeud, "CHECKNUM"))

            ' --- rapprochement avec le CSV pour récupérer la Catégorie ---
            categorie = ""
            estAmbigu = False
            texteCandidats = ""
            cle = CleComposite(dateComptable, Montant, tiers)
            If dictCategoriesCSV.Exists(cle) Then
                categorie = CategorieCommuneOuVide(dictCategoriesCSV(cle), estAmbigu, texteCandidats)
            Else
                nbSansCategorie = nbSansCategorie + 1
            End If

            ' --- ecriture de la ligne en memoire (colonnes 1 a 9 uniquement) ---
            nbAjoutees = nbAjoutees + 1
            resultats(nbAjoutees, COL_ID) = fitid
            resultats(nbAjoutees, COL_DATE_COMPTABLE) = dateComptable
            resultats(nbAjoutees, COL_DATE_OPERATION) = dateOperation
            resultats(nbAjoutees, COL_TYPE) = LireValeurNoeud(noeud, "TRNTYPE")
            resultats(nbAjoutees, COL_TIERS) = tiers
            resultats(nbAjoutees, COL_LIBELLE) = libelleDetail
            resultats(nbAjoutees, COL_MONTANT) = Montant
            resultats(nbAjoutees, COL_NUM_CHEQUE) = numCheque
            resultats(nbAjoutees, COL_CATEGORIE) = categorie

            If estAmbigu Then
                nbAmbigus = nbAmbigus + 1
                lignesAmbigues(nbAmbigus) = nbAjoutees
                texteAmbigus(nbAmbigus) = Format(dateComptable, "dd/mm/yyyy") & "  |  " & _
                                           Format(Montant, "0.00") & " EUR  |  " & tiers
                candidatsAmbigus(nbAmbigus) = texteCandidats
            End If

            dictFITID.Add fitid, True
        Else
            nbDoublons = nbDoublons + 1
        End If
    Next i

    ' --- ETAPE 6 : Resolution des catégories ambigues via le formulaire -------
    If nbAmbigus > 0 Then
        g_NbCasAmbigus = nbAmbigus
        ReDim g_CasTexte(1 To nbAmbigus)
        ReDim g_CasCandidats(1 To nbAmbigus)
        ReDim g_CasChoix(1 To nbAmbigus)
        For i = 1 To nbAmbigus
            g_CasTexte(i) = texteAmbigus(i)
            g_CasCandidats(i) = candidatsAmbigus(i)
            g_CasChoix(i) = ""
        Next i

        'frmResolutionCategories.Show vbModal   ' bloque l'execution jusqu'a fermeture
        AfficherFeuilleResolutionEtAttendre    ' bloque l'execution jusqu'a fermeture (feuille dediee)
        
        ' On applique les choix faits par l'opérateur dans résultats()
        For i = 1 To nbAmbigus
            If g_CasChoix(i) <> "" Then
                resultats(lignesAmbigues(i), COL_CATEGORIE) = g_CasChoix(i)
            End If
        Next i
    End If

    ' --- ETAPE 6bis (PHASE 5) : controle opérateur des catégories/sous-catégories ---
    ' Objectif demandé par l'opérateur : avant de finaliser l'import, lui
    ' demander si les opérations ont ete correctement categorisees par le
    ' rapprochement CSV (étapes precedentes). S'il repond "non" (ou si le
    ' formulaire est simplement lance), on lui présente les opérations
    ' NOUVELLEMENT importees une par une, avec desormais un choix a 2 niveaux
    ' (Catégorie / SousCategorie) au lieu du seul champ Catégorie d'origine.
    '
    ' mod_ControleCategories.ControlerCategories gere lui-même la question
    ' "tout est-il correct ?" et, le cas echeant, l'affichage feuille par
    ' feuille : ce module n'a qu'a lui fournir les opérations candidates et
    ' récupérer les résultats.
    '
    ' IMPORTANT SUR L'ANNULATION : si l'opérateur annule le controle (bouton
    ' "Annuler" du formulaire), ControlerCategories renvoie False. A CE STADE,
    ' RIEN N'A ENCORE ETE ECRIT dans la feuille Import_data (l'ecriture ne se
    ' fait qu'a l'ETAPE 8, plus loin) : on peut donc interrompre tout l'import
    ' proprement avec un simple Exit Sub, sans laisser le classeur dans un
    ' état intermediaire incoherent.
    If nbAjoutees > 0 Then
        Dim opsControle() As Variant
        Dim catFinaleCtrl() As String, sousFinaleCtrl() As String
        Dim kCtrl As Long

        ReDim opsControle(1 To nbAjoutees, 1 To mod_ControleCategories.CTRL_OP_NBCOL)
        For kCtrl = 1 To nbAjoutees
            opsControle(kCtrl, mod_ControleCategories.CTRL_OP_DATE) = resultats(kCtrl, COL_DATE_COMPTABLE)
            opsControle(kCtrl, mod_ControleCategories.CTRL_OP_TIERS) = resultats(kCtrl, COL_TIERS)
            opsControle(kCtrl, mod_ControleCategories.CTRL_OP_LIBELLE) = resultats(kCtrl, COL_LIBELLE)
            opsControle(kCtrl, mod_ControleCategories.CTRL_OP_MONTANT) = resultats(kCtrl, COL_MONTANT)
            opsControle(kCtrl, mod_ControleCategories.CTRL_OP_CATSOURCE) = resultats(kCtrl, COL_CATEGORIE)
            opsControle(kCtrl, mod_ControleCategories.CTRL_OP_ID) = resultats(kCtrl, COL_ID)
        Next kCtrl

        If Not mod_ControleCategories.ControlerCategories(opsControle, nbAjoutees, catFinaleCtrl, sousFinaleCtrl) Then
            MsgBox "Import annule : aucune ligne n'a ete ecrite dans Import_data.", vbInformation, "Import interrompu"
            Exit Sub
        End If

        ' On applique les choix opérateur : résultats(i, COL_CATEGORIE) recoit
        ' desormais la CATEGORIE PARENTE (ex: "Santé, prevoyance"), la
        ' sous-catégorie (ex: "Frais, remb santé") est memorisee a part dans
        ' sousFinaleCtrl() pour être ecrite plus loin (ETAPE 8bis), une fois
        ' que la colonne SousCategorie de TblOperations est atteignable par
        ' son NOM (elle n'existe pas dans le tableau fixe résultats()/arrFinal()).
        For kCtrl = 1 To nbAjoutees
            resultats(kCtrl, COL_CATEGORIE) = catFinaleCtrl(kCtrl)
        Next kCtrl
    End If

    ' --- ETAPE 7 : Calcul des colonnes derivees (Budget, MoisBudget, ...) -----
    ' Fait maintenant, une fois toutes les Catégorie definitives connues.
    Dim budget As Date, dateConsult As Variant, speConsult As Variant
    Dim categorieFinale As String, tiersFinal As String, libelleFinal As String
    Dim segment0 As String

    For i = 1 To nbAjoutees
        categorieFinale = CStr(resultats(i, COL_CATEGORIE))
        tiersFinal = CStr(resultats(i, COL_TIERS))
        libelleFinal = CStr(resultats(i, COL_LIBELLE))

        budget = CalculerBudget(CDate(resultats(i, COL_DATE_COMPTABLE)), tiersFinal, categorieFinale)
        resultats(i, COL_BUDGET) = budget
        resultats(i, COL_MOIS_BUDGET) = Month(budget)
        resultats(i, COL_ANNEE_BUDGET) = Year(budget)

        ' PHASE 5 : ce test portait auparavant sur "categorieFinale" (l'ancienne
        ' colonne Catégorie faisait alors office de plus fine granularite).
        ' Catégorie contient maintenant la catégorie PARENTE (ex: "Santé,
        ' prevoyance", commune a plusieurs sous-catégories) : le test doit
        ' donc porter sur la SOUS-catégorie choisie par l'opérateur a l'étape
        ' 6bis ci-dessus.
        If sousFinaleCtrl(i) = "Frais, remb sant" & Chr(233) Then   ' "Frais, remb santé"
            segment0 = SegmentTexte(libelleFinal, ";", 0)
            If EstDateValide(segment0) Then
                dateConsult = DateSerial(CInt(Left(segment0, 4)), CInt(Mid(segment0, 5, 2)), CInt(Right(segment0, 2)))
            Else
                dateConsult = DateSerial(1900, 1, 2)  ' 02/01/1900 = valeur "impossible a transformer"
            End If
            speConsult = SegmentTexte(libelleFinal, ";", 1)
            On Error Resume Next
            specialiteValide = (Application.WorksheetFunction.CountIf(Range("Specialites"), speConsult) > 0)
            On Error GoTo 0
            
            If Not specialiteValide Then
                resultats(i, COL_COM_SANTE) = speConsult
            Else
                resultats(i, COL_SPE_CONSULT) = speConsult
            End If
        Else
            dateConsult = Empty
            speConsult = Empty
        End If
        resultats(i, COL_DATE_CONSULT) = dateConsult
    Next i

    ' --- ETAPE 8 : Ecriture des nouvelles lignes dans la feuille --------------
    Dim premiereLigneEcriture As Long
    Dim listeIDsImportes() As String   ' memorise les ID_Transaction ajoutes par CET import, AVANT le tri de l'étape 9 (qui va melanger leur position)
    If nbAjoutees > 0 Then

        ReDim arrFinal(1 To nbAjoutees, 1 To NB_COLONNES)
        ReDim listeIDsImportes(1 To nbAjoutees)
        For k = 1 To nbAjoutees
            For i = 1 To NB_COLONNES
                arrFinal(k, i) = resultats(k, i)
            Next i
            listeIDsImportes(k) = CStr(resultats(k, COL_ID))
        Next k

        If wsDonnees.Range("A1").value = "" Then
            wsDonnees.Range(wsDonnees.Cells(1, 1), wsDonnees.Cells(1, NB_COLONNES)).value = Array( _
                "ID_Transaction", "Date_Comptable", "Date_Operation", "Type_Operation", _
                "Tiers", "Notes", "Montant", "Num_Cheque", "Categorie", _
                "Budget", "MoisBudget", "AnneeBudget", "Date_consult", "Spe_Consult")
            derniereLigne = 1
        End If

        premiereLigneEcriture = derniereLigne + 1
        wsDonnees.Range( _
            wsDonnees.Cells(premiereLigneEcriture, 1), _
            wsDonnees.Cells(premiereLigneEcriture + nbAjoutees - 1, NB_COLONNES) _
        ).Value2 = arrFinal

        ' Mise en forme des colonnes dates
        wsDonnees.Range(wsDonnees.Cells(premiereLigneEcriture, COL_DATE_COMPTABLE), _
                         wsDonnees.Cells(premiereLigneEcriture + nbAjoutees - 1, COL_DATE_COMPTABLE)).NumberFormat = "dd/mm/yyyy"
        wsDonnees.Range(wsDonnees.Cells(premiereLigneEcriture, COL_DATE_OPERATION), _
                         wsDonnees.Cells(premiereLigneEcriture + nbAjoutees - 1, COL_DATE_OPERATION)).NumberFormat = "dd/mm/yyyy"
        wsDonnees.Range(wsDonnees.Cells(premiereLigneEcriture, COL_BUDGET), _
                         wsDonnees.Cells(premiereLigneEcriture + nbAjoutees - 1, COL_BUDGET)).NumberFormat = "mm/yyyy"
        wsDonnees.Range(wsDonnees.Cells(premiereLigneEcriture, COL_DATE_CONSULT), _
                         wsDonnees.Cells(premiereLigneEcriture + nbAjoutees - 1, COL_DATE_CONSULT)).NumberFormat = "dd/mm/yyyy"
    End If

    ' --- ETAPE 8bis (PHASE 5) : ecriture de la SousCategorie pour les lignes --
    ' juste importees. La colonne "SousCategorie" n'existe PAS dans le
    ' tableau fixe résultats()/arrFinal() (NB_COLONNES=20, defini avant que
    ' cette colonne n'existe) : on la retrouve donc par son NOM directement
    ' sur la feuille, comme on le fait déjà partout ailleurs dans ce chantier
    ' pour ne jamais dependre d'un numéro de colonne fige.
    '
    ' IMPORTANT SUR L'ORDRE : cette étape doit imperativement s'executer ICI,
    ' AVANT l'ETAPE 9 (tri de TblOperations) : premiereLigneEcriture designe
    ' pour l'instant des lignes ENCORE CONTIGUES et dans l'ordre d'arrivee.
    ' Après le tri de l'ETAPE 9, cette correspondance n'existerait plus.
    If nbAjoutees > 0 Then
        Dim colSousCategorieFeuille As Long
        On Error Resume Next
        colSousCategorieFeuille = 0
        colSousCategorieFeuille = wsDonnees.ListObjects(NOM_TABLE).ListColumns("SousCategorie").index
        On Error GoTo 0

        If colSousCategorieFeuille = 0 Then
            ' TblOperations n'existe pas encore en tant que Tableau Excel a ce
            ' stade precis (cas du tout premier import, avant que l'ETAPE 9 ne
            ' le créé) : on cherche alors directement dans la ligne d'entetes.
            Dim colBalayage As Long
            Dim derniereColEntete As Long
            derniereColEntete = wsDonnees.Cells(1, wsDonnees.Columns.count).End(xlToLeft).Column
            For colBalayage = 1 To derniereColEntete
                If Trim(CStr(wsDonnees.Cells(1, colBalayage).value)) = "SousCategorie" Then
                    colSousCategorieFeuille = colBalayage
                    Exit For
                End If
            Next colBalayage
        End If

        If colSousCategorieFeuille = 0 Then
            ' Phase 1 (mod_Categories.PreparerPhase1Categories) pas encore
            ' installee chez l'opérateur : on previent, mais on NE BLOQUE PAS
            ' l'import pour autant (les opérations sont déjà ecrites a l'étape
            ' precedente ; les bloquer maintenant laisserait le classeur dans
            ' un état incoherent). La sous-catégorie pourra être completee
            ' plus tard, une fois la colonne ajoutee, via RechercherOperations.
            MsgBox "La colonne 'SousCategorie' est introuvable dans TblOperations." & vbCrLf & _
                   "L'import continue, mais la sous-categorie n'a pas ete enregistree pour " & _
                   "les operations qui viennent d'etre importees." & vbCrLf & _
                   "Installez la Phase 1 (mod_Categories.PreparerPhase1Categories) puis " & _
                   "corrigez ces lignes via RechercherOperations.", vbExclamation, "Colonne manquante"
        Else
            For kCtrl = 1 To nbAjoutees
                wsDonnees.Cells(premiereLigneEcriture + kCtrl - 1, colSousCategorieFeuille).value = sousFinaleCtrl(kCtrl)
            Next kCtrl
        End If
    End If

    ' --- ETAPE 9 : Convertir/redimensionner en Tableau Excel, trier, filtrer --
    Dim tbl As ListObject
    Dim wsDerniereLigneFinale As Long, plageComplete As Range
    Dim wsDerniereColonneFinale As Long

    ' IMPORTANT : on NE PEUT PAS utiliser NB_COLONNES (fige a 14, le nombre de
    ' colonnes fournies par l'import OFX/CSV) pour dimensionner le tableau ici.
    ' TblOperations compte maintenant plus de colonnes (StatutSante, Franchise,
    ' SoldeSante, DepassementHoraires, CommentaireSante, Bénéficiaire, ajoutees
    ' manuellement pour le suivi santé). Si on redimensionnait la table a
    ' NB_COLONNES=14, ces colonnes supplementaires seraient retirees de la
    ' definition du tableau structure (tbl.ListColumns("StatutSante") cesserait
    ' de fonctionner), même si les valeurs restaient physiquement dans la
    ' feuille. On calcule donc ici la largeur REELLE utilisée en ligne 1.
    wsDerniereColonneFinale = wsDonnees.Cells(1, wsDonnees.Columns.count).End(xlToLeft).Column
    If wsDerniereColonneFinale < NB_COLONNES Then wsDerniereColonneFinale = NB_COLONNES

    wsDerniereLigneFinale = wsDonnees.Cells(wsDonnees.rows.count, 1).End(xlUp).Row
    If wsDerniereLigneFinale >= 2 Then
        Set plageComplete = wsDonnees.Range( _
            wsDonnees.Cells(1, 1), wsDonnees.Cells(wsDerniereLigneFinale, wsDerniereColonneFinale))

        If wsDonnees.ListObjects.count = 0 Then
            Set tbl = wsDonnees.ListObjects.Add(xlSrcRange, plageComplete, , xlYes)
            tbl.Name = NOM_TABLE
        Else
            Set tbl = wsDonnees.ListObjects(1)
            tbl.Resize plageComplete
        End If
        wsDonnees.Activate   ' garantit que la feuille est active avant AutoFit (evite l'erreur 1004)
        wsDonnees.Range(wsDonnees.Cells(1, 1), wsDonnees.Cells(wsDerniereLigneFinale, wsDerniereColonneFinale)).Columns.AutoFit
        tbl.ShowAutoFilter = True
        
        With tbl.Sort
            .SortFields.Clear
            .SortFields.Add Key:=tbl.ListColumns(COL_DATE_COMPTABLE).Range, Order:=xlDescending
            .Header = xlYes
            .Apply
        End With
    End If

    ' --- ETAPE 9bis-0 : Vérification/reparation du decoupage du champ Notes --
    ' Sur TOUTE la table (pas seulement les lignes de cet import), pour
    ' detecter aussi les cas ou une valeur a change a la source depuis un
    ' import precedent (le rapprochement par FITID ne les detecterait pas).
    mod_FormulairesNotes.VerifierNotesSante

    ' --- ETAPE 9bis : Suivi santé (calcul automatique + formulaire opérateur) --
    ' On recalcule d'abord StatutSante/SoldeSante sur TOUTE la table (pas
    ' seulement les lignes de cet import : ca permet aussi de retraiter les
    ' anciens cas "KO" en attente, par exemple si un remboursement tarde a
    ' arriver). Puis on propose le formulaire pour tous les cas qui ont
    ' besoin d'une action de l'opérateur (Bénéficiaire, Tiers, Franchise,
    ' Depassement d'honoraires).
    mod_SuiviSante.CalculerSuiviSante AfficherResume:=False
    mod_SuiviSanteFormulaire.TraiterCasSuiviSante

    ' --- ETAPE 9ter : Memoriser et afficher les opérations nouvellement importees --
    If nbAjoutees > 0 Then
        mod_DernierImport.MemoriserDernierImport listeIDsImportes, nbAjoutees
        mod_DernierImport.AfficherDernierImportSurSynthese
    End If

    ' --- ETAPE 10 : Rapport final -----------------------------------------------
    MsgBox "Import termine." & vbCrLf & vbCrLf & _
           "Transactions OFX lues : " & nbLues & vbCrLf & _
           "Deja presentes (ignorees) : " & nbDoublons & vbCrLf & _
           "Nouvelles lignes ajoutees : " & nbAjoutees & vbCrLf & _
           "  dont sans correspondance CSV (categorie vide) : " & nbSansCategorie & vbCrLf & _
           "  dont categorie ambigue traitee via le formulaire : " & nbAmbigus, _
           vbInformation, "Import OFX + CSV"

End Sub


' ----------------------------------------------------------------------------
' FONCTIONS UTILITAIRES PRIVEES
' ----------------------------------------------------------------------------

' Ouvre la boite de dialogue de selection de fichier. Renvoie "" si annule.
Private Function ChoisirFichier(filtre As String, Titre As String) As String
    Dim chemin As String
    chemin = Application.GetOpenFilename(filtre, , Titre)
    If chemin = "False" Then
        ChoisirFichier = ""
    Else
        ChoisirFichier = chemin
    End If
End Function

' Lit le texte d'une balise enfant (ex: <TRNAMT>) a l'interieur d'un noeud OFX.
Private Function LireValeurNoeud(parentNoeud As Object, nomBalise As String) As String
    Dim collectionEnfants As Object
    Set collectionEnfants = parentNoeud.getElementsByTagName(nomBalise)
    If collectionEnfants.Length = 0 Then
        LireValeurNoeud = ""
    Else
        LireValeurNoeud = Trim(collectionEnfants.Item(0).Text)
    End If
End Function

' Convertit une date OFX "AAAAMMJJHHMMSS.mmm" en date Excel.
Private Function ConvertirDateOFX(chaineDate As String) As Date
    If Len(chaineDate) < 8 Then
        ConvertirDateOFX = 0
        Exit Function
    End If
    ConvertirDateOFX = DateSerial( _
        CInt(Mid(chaineDate, 1, 4)), CInt(Mid(chaineDate, 5, 2)), CInt(Mid(chaineDate, 7, 2)))
End Function

' Même conversion mais renvoie Empty (cellule vide) si la date est absente
' (utilisée pour Date_Operation, souvent non renseignee).
Private Function ConvertirDateOFXVariant(chaineDate As String) As Variant
    If Len(chaineDate) < 8 Then
        ConvertirDateOFXVariant = Empty
    Else
        ConvertirDateOFXVariant = ConvertirDateOFX(chaineDate)
    End If
End Function

' Montant OFX : toujours ecrit avec un POINT decimal -> Val() l'interprete
' correctement quels que soient les parametres regionaux Windows.
Private Function ConvertirMontantOFX(chaineMontant As String) As Double
    ConvertirMontantOFX = Val(chaineMontant)
End Function

' Montant CSV : ecrit avec une VIRGULE decimale (format francais) -> on la
' remplace par un point avant d'utiliser Val().
Private Function ConvertirMontantCSV(chaineMontant As String) As Double
    ConvertirMontantCSV = Val(Replace(Trim(chaineMontant), ",", "."))
End Function

' Date CSV au format texte "JJ/MM/AAAA".
Private Function ConvertirDateCSV(chaineDate As String) As Date
    Dim parties() As String
    parties = Split(Trim(chaineDate), "/")
    If UBound(parties) = 2 Then
        ConvertirDateCSV = DateSerial(CInt(parties(2)), CInt(parties(1)), CInt(parties(0)))
    End If
End Function

' Numéro de cheque : on ne garde que les valeurs numeriques, le reste devient vide.
Private Function NettoyerNumCheque(brut As String) As Variant
    If brut = "" Or Not IsNumeric(brut) Then
        NettoyerNumCheque = Empty
    Else
        NettoyerNumCheque = brut
    End If
End Function

' Construit la clé de rapprochement OFX <-> CSV : Date + Montant + Tiers.
Private Function CleComposite(dateOp As Date, montantOp As Double, tiersOp As String) As String
    CleComposite = Format(dateOp, "yyyymmdd") & "|" & Format(montantOp, "0.00") & "|" & UCase(Trim(tiersOp))
End Function

' Extrait un segment d'un texte separe par un caractere donne (index base 0).
' Renvoie "" si le segment demandé n'existe pas.
Public Function SegmentTexte(texte As String, separateur As String, index As Integer) As String
    Dim parties() As String
    parties = Split(texte, separateur)
    If index >= LBound(parties) And index <= UBound(parties) Then
        SegmentTexte = Trim(parties(index))
    Else
        SegmentTexte = ""
    End If
End Function

' Teste si un texte représente une date valide et non absurde.
Public Function EstDateValide(texte As String) As Boolean
    If Trim(texte) = "" Then
        EstDateValide = False
    ElseIf Len(texte) = 8 And IsNumeric(texte) Then ' 1. Vérification : la chaine doit faire 8 caracteres et être numerique
        
        Dim Annee As Integer
        Dim Mois As Integer
        Dim jour As Integer
        
        ' Extraction des parties
        Annee = CInt(Left(texte, 4))
        Mois = CInt(Mid(texte, 5, 2))
        jour = CInt(Right(texte, 2))
        
        ' 2. Vérification subsidiaire : les mois et jours sont-ils coherents ?
        If Mois >= 1 And Mois <= 12 And jour >= 1 And jour <= 31 And Annee >= 1900 Then
            EstDateValide = True
        Else
            EstDateValide = False
        End If
        
    Else
        ' Cas de figure ou le format AAAAMMJJ n'est pas respecte
        EstDateValide = False
    End If
End Function


' Calcule la colonne Budget selon la règle metier :
' - Si Tiers = "DRFIP OCCITANIE ET HTE" ET Catégorie = "Salaire/Revenus d'activite"
'   -> mois suivant celui de Date_Comptable
' - Sinon -> mois de Date_Comptable
' DateSerial() gere seul le changement d'année (mois 13 -> janvier année+1).
Private Function CalculerBudget(dateComptable As Date, tiers As String, categorie As String) As Date
    If Trim(UCase(tiers)) = "DRFIP OCCITANIE ET HTE" And categorie = "Salaire/Revenus d'activit" & Chr(233) Then
        CalculerBudget = DateSerial(Year(dateComptable), Month(dateComptable) + 1, 1)
    Else
        CalculerBudget = DateSerial(Year(dateComptable), Month(dateComptable), 1)
    End If
End Function

' Recherche l'index d'une colonne par son nom d'entete (recherche exacte).
Private Function TrouverColonne(entetes() As String, nom As String) As Integer
    Dim i As Integer
    For i = LBound(entetes) To UBound(entetes)
        If Trim(entetes(i)) = nom Then
            TrouverColonne = i
            Exit Function
        End If
    Next i
    TrouverColonne = -1
End Function

' Lit le fichier CSV (UTF-16, tabulations) et construit un dictionnaire :
'   clé "Date|Montant|Tiers" -> Collection des catégories trouvees pour cette clé.
' Une clé peut avoir plusieurs catégories si plusieurs opérations identiques
' existent le même jour (voir CategorieCommuneOuVide pour la resolution).
Private Function ChargerCategoriesCSV(chemin As String) As Object
    Dim flux As Object, contenuBrut As String
    Dim lignes() As String, entetes() As String, champs() As String
    Dim colDate As Integer, colLibelle As Integer, colCategorie As Integer, colMontant As Integer
    Dim i As Long
    Dim dict As Object
    Dim cle As String
    Dim dateOp As Date, montantOp As Double, tiersOp As String, categorieOp As String

    Set dict = CreateObject("Scripting.Dictionary")

    Set flux = CreateObject("ADODB.Stream")
    flux.Type = 2   ' 2 = flux texte
    flux.Charset = "utf-16"   ' le CSV est encode en UTF-16 malgre son extension .csv
    flux.Open
    flux.LoadFromFile chemin
    contenuBrut = flux.ReadText
    flux.Close

    ' Retire un éventuel caractere BOM invisible en tout debut de fichier
    If Len(contenuBrut) > 0 Then
        If AscW(Left(contenuBrut, 1)) = 65279 Then contenuBrut = Mid(contenuBrut, 2)
    End If

    lignes = Split(contenuBrut, vbCrLf)
    entetes = Split(lignes(0), vbTab)

    colDate = TrouverColonne(entetes, "Date")
    colLibelle = TrouverColonne(entetes, "Libell" & Chr(233))       ' "Libelle"
    colCategorie = TrouverColonne(entetes, "Cat" & Chr(233) & "gorie")  ' "Catégorie"
    colMontant = TrouverColonne(entetes, "Montant")

    If colDate = -1 Or colLibelle = -1 Or colCategorie = -1 Or colMontant = -1 Then
        MsgBox "Le fichier CSV n'a pas la structure attendue (colonnes Date, Libell" & Chr(233) & _
               ", Cat" & Chr(233) & "gorie, Montant introuvables).", vbCritical
        Set ChargerCategoriesCSV = Nothing
        Exit Function
    End If

    For i = 1 To UBound(lignes)
        If Trim(lignes(i)) <> "" Then
            champs = Split(lignes(i), vbTab)
            If UBound(champs) >= colMontant Then
                dateOp = ConvertirDateCSV(champs(colDate))
                montantOp = ConvertirMontantCSV(champs(colMontant))
                tiersOp = Trim(champs(colLibelle))
                categorieOp = Trim(champs(colCategorie))

                cle = CleComposite(dateOp, montantOp, tiersOp)

                If Not dict.Exists(cle) Then
                    dict.Add cle, New Collection
                End If
                dict(cle).Add categorieOp
            End If
        End If
    Next i

    Set ChargerCategoriesCSV = dict
End Function

' A partir d'une Collection de catégories candidates pour une même clé :
' - si toutes identiques -> renvoie cette valeur, estAmbigu = False
' - si différentes -> renvoie "", estAmbigu = True, et la liste des valeurs
'   distinctes (separees par ";") dans texteCandidats, prete pour le formulaire
'   de resolution.
Private Function CategorieCommuneOuVide(candidats As Collection, ByRef estAmbigu As Boolean, ByRef texteCandidats As String) As String
    Dim c As Variant
    Dim premiereValeur As String
    Dim toutesIdentiques As Boolean
    Dim distincts As Object
    Set distincts = CreateObject("Scripting.Dictionary")

    premiereValeur = CStr(candidats(1))
    toutesIdentiques = True

    For Each c In candidats
        If Not distincts.Exists(CStr(c)) Then distincts.Add CStr(c), True
        If CStr(c) <> premiereValeur Then toutesIdentiques = False
    Next c

    If toutesIdentiques Then
        estAmbigu = False
        CategorieCommuneOuVide = premiereValeur
        texteCandidats = ""
    Else
        estAmbigu = True
        CategorieCommuneOuVide = ""
        texteCandidats = Join(distincts.Keys, ";")
    End If
End Function
