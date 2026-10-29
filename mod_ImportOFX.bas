Option Explicit

' ============================================================================
'  MODULE : mod_ImportOFX
'  RÔLE   : Importer les opérations bancaires depuis un fichier OFX (et,
'           facultativement, un fichier CSV complémentaire), récupérer la
'           catégorie source, calculer les colonnes budgétaires, puis tout
'           écrire dans la feuille "data_import" sous la forme d'un tableau
'           Excel trié et filtrable.
'
'  Catégorie source (changement du 17/10/2026) : la banque exporte le MEMO OFX
'  sous la forme "<Catégorie> - <Notes>". La catégorie est reconnue contre la
'  colonne CategorieSource de TblCategories (ExtraireCategorieMemo) et retirée
'  de Notes. Un MEMO non reconnu est laissé tel quel dans Notes (analyse manuelle).
'  Le CSV n'est qu'un repli facultatif pour ces MEMO non reconnus.
'
'  En cas de catégorie ambiguë via le CSV (plusieurs valeurs possibles pour une
'  même opération), le formulaire frmResolutionCategories s'ouvre pour laisser
'  l'opérateur choisir, sans qu'il ait besoin d'accéder au tableau.
'
'  À appeler depuis le bouton de la feuille "Accueil".
' ============================================================================

' --- Position des colonnes dans la zone de sortie (pour éviter les "magic numbers") ---
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

' --- Variables PUBLIQUES partagées avec le formulaire frmResolutionCategories ---
' (rempli par ce module avant l'ouverture du formulaire, relu après sa fermeture)
Public g_NbCasAmbigus As Long
Public g_CasTexte() As String       ' texte affiché dans la liste (Date | Montant | Tiers)
Public g_CasCandidats() As String   ' catégories candidates, séparées par ";"
Public g_CasChoix() As String       ' rempli par le formulaire : catégorie choisie (ou "" si non traité)


' ----------------------------------------------------------------------------
' MACRO PRINCIPALE - à assigner au bouton de la feuille Accueil
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

    ' --- ÉTAPE 1 : sélection des deux fichiers à importer ---------------------
    cheminOFX = ChoisirFichier("Fichiers OFX (*.ofx), *.ofx", "Selectionner l'export bancaire OFX")
    If cheminOFX = "" Then Exit Sub

    ' CHANGEMENT du 17/10/2026 : le CSV est désormais FACULTATIF. La catégorie source est lue
    ' dans le MEMO de l'OFX ("<Catégorie> - <Notes>", voir ExtraireCategorieMemo); le CSV ne sert
    ' plus que de repli pour les opérations dont le MEMO n'est pas reconnu. Annuler la fenêtre
    ' de sélection du CSV revient à l'ignorer (l'import continue).
    cheminCSV = ""
    If MsgBox("Joindre aussi le fichier CSV ?" & vbCrLf & vbCrLf & _
              "Facultatif : la categorie est lue dans l'OFX. Le CSV ne sert que de secours " & _
              "pour les operations dont la categorie n'est pas reconnue dans l'OFX.", _
              vbYesNo + vbDefaultButton2 + vbQuestion, "Fichier CSV complementaire") = vbYes Then
        cheminCSV = ChoisirFichier("Fichiers CSV (*.csv), *.csv", _
            "Selectionner le fichier CSV complementaire (memes operations) - Annuler = ignorer")
    End If

    ' --- ÉTAPE 2 : chargement du fichier OFX comme document XML ----------------
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

    ' --- ÉTAPE 3 : chargement du CSV -> dictionnaire des catégories par clé ---
    ' Clé = Date + Montant + Tiers. Une même clé peut avoir plusieurs catégories
    ' candidates (cas de plusieurs opérations identiques le même jour).
    ' Sans CSV : dictionnaire vide (Exists renvoie toujours False, aucun repli possible).
    If cheminCSV <> "" Then
        Set dictCategoriesCSV = ChargerCategoriesCSV(cheminCSV)
        If dictCategoriesCSV Is Nothing Then Exit Sub  ' l'erreur a déjà été affichée
    Else
        Set dictCategoriesCSV = CreateObject("Scripting.Dictionary")
    End If

    ' Catégories source connues (colonne CategorieSource de TblCategories) : elles servent à
    ' reconnaître sans ambiguïté le préfixe "<catégorie> - " du MEMO OFX.
    Dim sourcesConnues() As String, nbSources As Long
    sourcesConnues = mod_Categories.ObtenirCategoriesSource()
    On Error Resume Next
    nbSources = UBound(sourcesConnues) - LBound(sourcesConnues) + 1   ' erreur si tableau non alloué
    If Err.Number <> 0 Then nbSources = 0
    On Error GoTo 0
    Dim nbMemoReconnus As Long, nbMemoNonReconnus As Long, nbReperCSV As Long, nbDivergencesCSV As Long
    Dim catMemo As String, notesMemo As String
    Dim catCSV As String, estAmbiguCSV As Boolean, texteCandidatsCSV As String

    ' --- ÉTAPE 4 : préparer la feuille data_import et les FITID déjà connus ---
    Set wsDonnees = ThisWorkbook.Worksheets("Import_data")
    Set dictFITID = CreateObject("Scripting.Dictionary")

    Dim derniereLigne As Long
    derniereLigne = wsDonnees.Cells(wsDonnees.rows.count, "A").End(xlUp).Row

    If derniereLigne >= 2 Then
        Dim plageFITID As Variant
        plageFITID = wsDonnees.Range("A2:A" & derniereLigne).Value2

        If derniereLigne = 2 Then
            ' Cas particulier : une seule ligne de données. Value2 renvoie alors
            ' directement la valeur (pas un tableau 2D); UBound provoquerait une erreur.
            If Not dictFITID.Exists(CStr(plageFITID)) Then
                dictFITID.Add CStr(plageFITID), True
            End If
        Else
            For i = 1 To UBound(plageFITID, 1)
                If Not dictFITID.Exists(CStr(plageFITID(i, 1))) Then
                    dictFITID.Add CStr(plageFITID(i, 1)), True
                End If
            Next i
        End If
    End If
    
    ' --- ÉTAPE 5 : parcourir les transactions OFX ------------------------------
    ' À ce stade, on ne remplit QUE les colonnes 1 à 9 (ID .. Catégorie).
    ' Les colonnes calculées 10 à 14 (Budget, MoisBudget, ...) dépendent de la
    ' catégorie : elles seront calculées plus loin, une fois les ambiguïtés résolues.
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

            ' --- Catégorie source : d'abord le MEMO de l'OFX, puis (repli) le CSV ---
            ' CHANGEMENT du 17/10/2026 : la banque exporte le MEMO sous la forme
            ' "<Catégorie> - <Notes>" (ou la catégorie seule si Notes est vide). Le préfixe est
            ' reconnu contre les catégories source connues (TblCategories); la catégorie va dans
            ' COL_CATEGORIE et seul le reste va dans Notes (sinon la clé santé "AAAAMMJJ;..."
            ' ne serait plus en tête du champ). Chaque opération ayant son FITID, il n'y a plus
            ' de cas ambigu avec cette méthode.
            categorie = ""
            estAmbigu = False
            texteCandidats = ""
            cle = CleComposite(dateComptable, Montant, tiers)

            If ExtraireCategorieMemo(libelleDetail, sourcesConnues, nbSources, catMemo, notesMemo) Then
                nbMemoReconnus = nbMemoReconnus + 1
                categorie = catMemo
                libelleDetail = notesMemo
                ' Contrôle croisé facultatif : le MEMO fait foi, on compte seulement les écarts.
                If dictCategoriesCSV.Exists(cle) Then
                    catCSV = CategorieCommuneOuVide(dictCategoriesCSV(cle), estAmbiguCSV, texteCandidatsCSV)
                    If Not estAmbiguCSV And catCSV <> "" Then
                        If StrComp(catCSV, categorie, vbTextCompare) <> 0 Then nbDivergencesCSV = nbDivergencesCSV + 1
                    End If
                End If
            ElseIf dictCategoriesCSV.Exists(cle) Then
                ' MEMO non reconnu : repli sur le CSV (comportement historique, avec cas ambigus).
                nbMemoNonReconnus = nbMemoNonReconnus + 1
                nbReperCSV = nbReperCSV + 1
                categorie = CategorieCommuneOuVide(dictCategoriesCSV(cle), estAmbigu, texteCandidats)
                libelleDetail = RetirerPrefixeCategorieMemo(libelleDetail, categorie, texteCandidats)
            Else
                ' Ni MEMO reconnu, ni CSV : catégorie vide (opération à ranger à la main) et MEMO
                ' conservé TEL QUEL dans Notes pour permettre une analyse manuelle.
                nbMemoNonReconnus = nbMemoNonReconnus + 1
                nbSansCategorie = nbSansCategorie + 1
            End If

            ' --- Écriture de la ligne en mémoire (colonnes 1 à 9 uniquement) ---
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

    ' --- ÉTAPE 6 : résolution des catégories ambiguës via le formulaire -------
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

        'frmResolutionCategories.Show vbModal   ' bloque l'exécution jusqu'à la fermeture
        AfficherFeuilleResolutionEtAttendre    ' bloque l'exécution jusqu'à la fermeture (feuille dédiée)
        
        ' On applique les choix faits par l'opérateur dans résultats()
        For i = 1 To nbAmbigus
            If g_CasChoix(i) <> "" Then
                resultats(lignesAmbigues(i), COL_CATEGORIE) = g_CasChoix(i)
            End If
        Next i
    End If

    ' --- ÉTAPE 6bis (PHASE 5) : contrôle opérateur des catégories/sous-catégories ---
    ' Objectif demandé par l'opérateur : avant de finaliser l'import, lui
    ' demander si les opérations ont été correctement catégorisées par le
    ' rapprochement CSV (étapes précédentes). S'il répond "non" (ou si le
    ' formulaire est simplement lancé), on lui présente les opérations
    ' nouvellement importées une par une, avec désormais un choix à deux niveaux
    ' (Catégorie / SousCategorie) au lieu du seul champ Catégorie d'origine.
    '
    ' mod_ControleCategories.ControlerCategories gère lui-même la question
    ' "tout est-il correct ?" et, le cas échéant, l'affichage feuille par
    ' feuille : ce module n'a qu'à lui fournir les opérations candidates et
    ' récupérer les résultats.
    '
    ' IMPORTANT SUR L'ANNULATION : si l'opérateur annule le contrôle (bouton
    ' "Annuler" du formulaire), ControlerCategories renvoie False. À CE STADE,
    ' RIEN N'A ENCORE ÉTÉ ÉCRIT dans la feuille Import_data (l'écriture ne se
    ' fait qu'à l'ÉTAPE 8, plus loin) : on peut donc interrompre tout l'import
    ' proprement avec un simple Exit Sub, sans laisser le classeur dans un
    ' état intermédiaire incohérent.
    If nbAjoutees > 0 Then
        Dim opsControle() As Variant
        Dim catFinaleCtrl() As String, sousFinaleCtrl() As String
        Dim catAvantVentilCtrl() As String, sousAvantVentilCtrl() As String
        ' AJOUT du 03/10/2026 (demande opérateur) : l'opérateur peut désormais corriger
        ' le Tiers et les Notes directement sur l'écran de contrôle des catégories (sauf
        ' pour les opérations de santé; voir mod_ControleCategories.AfficherOperation).
        ' Ces deux tableaux reçoivent les valeurs FINALES (modifiées ou non) renvoyées
        ' par ControlerCategories, selon le même principe que catFinaleCtrl/sousFinaleCtrl.
        Dim tiersFinaleCtrl() As String, libelleFinaleCtrl() As String
        ' AJOUT du 11/10/2026 (pointage, voir mod_Pointage) : pointage choisi par l'operateur
        ' pour chaque operation (tableau de Boolean, Vrai = deja analysee). Il est ecrit dans la
        ' colonne Pointe a l'ETAPE 8quater. Reste vide (Empty) si le controle n'a rien renvoye.
        Dim pointeFinaleCtrl As Variant
        Dim kCtrl As Long

        ReDim opsControle(1 To nbAjoutees, 1 To mod_ControleCategories.CTRL_OP_NBCOL)
        For kCtrl = 1 To nbAjoutees
            opsControle(kCtrl, mod_ControleCategories.CTRL_OP_DATE) = resultats(kCtrl, COL_DATE_COMPTABLE)
            opsControle(kCtrl, mod_ControleCategories.CTRL_OP_TIERS) = resultats(kCtrl, COL_TIERS)
            opsControle(kCtrl, mod_ControleCategories.CTRL_OP_LIBELLE) = resultats(kCtrl, COL_LIBELLE)
            opsControle(kCtrl, mod_ControleCategories.CTRL_OP_MONTANT) = resultats(kCtrl, COL_MONTANT)
            opsControle(kCtrl, mod_ControleCategories.CTRL_OP_CATSOURCE) = resultats(kCtrl, COL_CATEGORIE)
            opsControle(kCtrl, mod_ControleCategories.CTRL_OP_ID) = resultats(kCtrl, COL_ID)
            opsControle(kCtrl, mod_ControleCategories.CTRL_OP_CHEQUE) = resultats(kCtrl, COL_NUM_CHEQUE)
        Next kCtrl

        mod_Echeancier.InitialiserRapprochementImport   ' remet à zéro les rapprochements d'échéances de cet import
        If Not mod_ControleCategories.ControlerCategories(opsControle, nbAjoutees, catFinaleCtrl, sousFinaleCtrl, _
                                                            catAvantVentilCtrl, sousAvantVentilCtrl, _
                                                            tiersFinaleCtrl, libelleFinaleCtrl, _
                                                            pointeFinale:=pointeFinaleCtrl) Then
            MsgBox "Import annule : aucune ligne n'a ete ecrite dans Import_data.", vbInformation, "Import interrompu"
            Exit Sub
        End If

        ' On applique les choix de l'opérateur : résultats(i, COL_CATEGORIE) reçoit
        ' désormais la catégorie PARENTE (ex. : "Santé, prévoyance"); la sous-catégorie
        ' (ex. : "Frais, remb santé") est mémorisée séparément dans sousFinaleCtrl()
        ' pour être écrite plus loin (ÉTAPE 8bis), lorsque la colonne SousCategorie
        ' de TblOperations sera accessible par son nom (elle n'existe pas dans le
        ' tableau fixe résultats()/arrFinal()). COL_TIERS/COL_LIBELLE reçoivent de la
        ' même façon la valeur FINALE, modifiée ou non par l'opérateur. Le code suivant
        ' (calcul du budget, lecture de la date de consultation santé...) lit déjà ces
        ' colonnes; aucune autre modification n'est nécessaire.
        For kCtrl = 1 To nbAjoutees
            resultats(kCtrl, COL_CATEGORIE) = catFinaleCtrl(kCtrl)
            resultats(kCtrl, COL_TIERS) = tiersFinaleCtrl(kCtrl)
            resultats(kCtrl, COL_LIBELLE) = libelleFinaleCtrl(kCtrl)
        Next kCtrl

        ' --- ÉTAPE 6ter (ajout du 18/10/2026) : échéancier AVANT l'écriture ---
        ' Les catégories sont définitives (que l'opérateur ait répondu OUI ou NON au contrôle) :
        ' on repère les opérations qui répondent aux critères d'une échéance, on les affiche et
        ' l'opérateur confirme le rapprochement. Rien n'est encore supprimé de l'échéancier : les
        ' échéances validées sont retirées après l'écriture (mod_Echeancier.TraiterApresImport).
        Dim eID() As Variant, eDt() As Variant, eTiers() As Variant, eMont() As Variant
        Dim eCat() As Variant, eSous() As Variant, eChq() As Variant
        ReDim eID(1 To nbAjoutees): ReDim eDt(1 To nbAjoutees): ReDim eTiers(1 To nbAjoutees): ReDim eMont(1 To nbAjoutees)
        ReDim eCat(1 To nbAjoutees): ReDim eSous(1 To nbAjoutees): ReDim eChq(1 To nbAjoutees)
        For kCtrl = 1 To nbAjoutees
            eID(kCtrl) = resultats(kCtrl, COL_ID)
            eDt(kCtrl) = resultats(kCtrl, COL_DATE_COMPTABLE)
            eTiers(kCtrl) = resultats(kCtrl, COL_TIERS)
            eMont(kCtrl) = resultats(kCtrl, COL_MONTANT)
            eCat(kCtrl) = resultats(kCtrl, COL_CATEGORIE)
            eSous(kCtrl) = sousFinaleCtrl(kCtrl)
            eChq(kCtrl) = resultats(kCtrl, COL_NUM_CHEQUE)
        Next kCtrl
        mod_Echeancier.PreRapprocherAvantEcriture eID, eDt, eTiers, eMont, eCat, eSous, eChq, nbAjoutees
    End If

    ' --- ÉTAPE 7 : calcul des colonnes dérivées (Budget, MoisBudget, ...) ------
    ' Effectué maintenant, une fois toutes les catégories définitives connues.
    Dim budget As Date, dateConsult As Variant, speConsult As Variant
    Dim categorieFinale As String, tiersFinal As String, libelleFinal As String
    Dim segment0 As String
    Dim specialiteValide As Boolean

    For i = 1 To nbAjoutees
        categorieFinale = CStr(resultats(i, COL_CATEGORIE))
        tiersFinal = CStr(resultats(i, COL_TIERS))
        libelleFinal = CStr(resultats(i, COL_LIBELLE))

        budget = CalculerBudget(CDate(resultats(i, COL_DATE_COMPTABLE)), tiersFinal, categorieFinale, sousFinaleCtrl(i))
        resultats(i, COL_BUDGET) = budget
        resultats(i, COL_MOIS_BUDGET) = Month(budget)
        resultats(i, COL_ANNEE_BUDGET) = Year(budget)

        ' PHASE 5 : ce test portait auparavant sur "categorieFinale" (l'ancienne
        ' colonne Catégorie correspondait alors à la granularité la plus fine).
        ' Catégorie contient maintenant la catégorie PARENTE (ex. : "Santé,
        ' prévoyance", commune à plusieurs sous-catégories); le test doit donc
        ' porter sur la SOUS-catégorie choisie par l'opérateur à l'étape 6bis.
        ' MISE À JOUR du 03/10/2026 (regroupement des constantes globales, demande opérateur) :
        ' le texte "Frais, remb santé" était codé en dur ici et dans
        ' mod_ControleCategories (règle de verrouillage Tiers/Notes). Il a été déplacé
        ' dans mod_VarGlobales.SOUS_CATEGORIE_SANTE afin de n'avoir qu'un seul endroit
        ' à modifier si ce libellé change.
        If sousFinaleCtrl(i) = mod_VarGlobales.SOUS_CATEGORIE_SANTE Then   ' "Frais, remb santé"
            ' CORRECTIF du 12/10/2026 : certaines notes arrivent préfixées du nom de la
            ' sous-catégorie ("Frais, remb santé - 20260807;Ophtalmologue;Dédé;98,43").
            ' Le segment 0 n'est alors pas une date et la ligne tombe à tort dans
            ' frm_RapprochementNotes. On retire ce préfixe, et on réécrit la note nettoyée
            ' (c'est resultats(i, COL_LIBELLE) qui est enregistré dans la colonne Notes).
            libelleFinal = RetirerPrefixeSanteNotes(libelleFinal)
            resultats(i, COL_LIBELLE) = libelleFinal
            segment0 = SegmentTexte(libelleFinal, ";", 0)
            If EstDateValide(segment0) Then
                dateConsult = DateSerial(CInt(Left(segment0, 4)), CInt(Mid(segment0, 5, 2)), CInt(Right(segment0, 2)))
            Else
                dateConsult = DateSerial(1900, 1, 2)  ' 02/01/1900 = valeur impossible à convertir
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

    ' --- ÉTAPE 8 : écriture des nouvelles lignes dans la feuille ---------------
    Dim premiereLigneEcriture As Long
    Dim listeIDsImportes() As String   ' mémorise les ID_Transaction ajoutés par CET import AVANT le tri de l'étape 9, qui mélangera leur ordre
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
        ' Forcer la colonne ID_Transaction au format TEXTE avant d'écrire les
        ' valeurs, sinon Excel convertit les longs FITID numériques en nombres,
        ' avec une perte de précision au-delà de 15 à 17 chiffres.
        wsDonnees.Range(wsDonnees.Cells(premiereLigneEcriture, COL_ID), _
                         wsDonnees.Cells(premiereLigneEcriture + nbAjoutees - 1, COL_ID)).NumberFormat = "@"

        wsDonnees.Range( _
            wsDonnees.Cells(premiereLigneEcriture, 1), _
            wsDonnees.Cells(premiereLigneEcriture + nbAjoutees - 1, NB_COLONNES) _
        ).Value2 = arrFinal

        ' Mise en forme des colonnes de dates.
        wsDonnees.Range(wsDonnees.Cells(premiereLigneEcriture, COL_DATE_COMPTABLE), _
                         wsDonnees.Cells(premiereLigneEcriture + nbAjoutees - 1, COL_DATE_COMPTABLE)).NumberFormat = "dd/mm/yyyy"
        wsDonnees.Range(wsDonnees.Cells(premiereLigneEcriture, COL_DATE_OPERATION), _
                         wsDonnees.Cells(premiereLigneEcriture + nbAjoutees - 1, COL_DATE_OPERATION)).NumberFormat = "dd/mm/yyyy"
        wsDonnees.Range(wsDonnees.Cells(premiereLigneEcriture, COL_BUDGET), _
                         wsDonnees.Cells(premiereLigneEcriture + nbAjoutees - 1, COL_BUDGET)).NumberFormat = "mm/yyyy"
        wsDonnees.Range(wsDonnees.Cells(premiereLigneEcriture, COL_DATE_CONSULT), _
                         wsDonnees.Cells(premiereLigneEcriture + nbAjoutees - 1, COL_DATE_CONSULT)).NumberFormat = "dd/mm/yyyy"
    End If

    ' --- ÉTAPE 8bis (PHASE 5) : écriture de SousCategorie pour les lignes juste
    ' importées. Cette colonne n'existe PAS dans le tableau fixe résultats()/arrFinal()
    ' (NB_COLONNES=20, défini avant son ajout); on la retrouve donc directement
    ' sur la feuille par son nom, comme ailleurs dans ce chantier, afin de ne
    ' jamais dépendre d'un numéro de colonne fixe.
    '
    ' IMPORTANT SUR L'ORDRE : cette étape doit impérativement s'exécuter ICI,
    ' AVANT l'ÉTAPE 9 (tri de TblOperations) : premiereLigneEcriture désigne
    ' encore des lignes CONTIGUËS, dans leur ordre d'arrivée.
    ' Après le tri de l'ÉTAPE 9, cette correspondance n'existerait plus.
    If nbAjoutees > 0 Then
        Dim colSousCategorieFeuille As Long
        On Error Resume Next
        colSousCategorieFeuille = 0
        colSousCategorieFeuille = wsDonnees.ListObjects(NOM_TABLE).ListColumns("SousCategorie").index
        On Error GoTo 0

        If colSousCategorieFeuille = 0 Then
            ' TblOperations n'existe pas encore comme tableau Excel à ce stade
            ' précis (premier import, avant sa création à l'ÉTAPE 9) : on cherche
            ' alors directement dans la ligne d'en-têtes.
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
            ' La phase 1 (mod_Categories.PreparerPhase1Categories) n'est pas encore
            ' installée : on avertit l'opérateur, mais on NE BLOQUE PAS l'import
            ' (les opérations sont déjà écrites à l'étape précédente; le bloquer
            ' maintenant laisserait le classeur dans un état incohérent). La
            ' sous-catégorie pourra être complétée plus tard, après l'ajout de la
            ' colonne, via RechercherOperations.
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

    ' --- ÉTAPE 8ter (ajout du 01/10/2026, point 4 : annuler une ventilation) ---
    ' Écrit, pour les lignes importées, la catégorie/sous-catégorie d'AVANT LA
    ' VENTILATION renvoyée par ControlerCategories (catAvantVentilCtrl/
    ' sousAvantVentilCtrl) dans les colonnes techniques CategorieAvantVentilation /
    ' SousCategorieAvantVentilation de TblOperations (ajoutées par
    ' mod_InstallVentilation.AjouterColonnesAnnulationVentilation). Ces champs
    ' restent vides si l'opération n'a pas été ventilée pendant cet import (valeur
    ' par défaut ""). Comme pour SousCategorie, la colonne est retrouvée par son
    ' nom, jamais par un numéro fixe. Si les colonnes ne sont pas encore installées
    ' (phase d'annulation de ventilation), on avertit l'opérateur sans bloquer
    ' l'import, comme pour SousCategorie.
    If nbAjoutees > 0 Then
        Dim colCatAvantVenFeuille As Long, colSousAvantVenFeuille As Long
        On Error Resume Next
        colCatAvantVenFeuille = 0
        colSousAvantVenFeuille = 0
        colCatAvantVenFeuille = wsDonnees.ListObjects(NOM_TABLE).ListColumns(mod_InstallVentilation.NOM_COL_CAT_AVANT_VENTILATION).index
        colSousAvantVenFeuille = wsDonnees.ListObjects(NOM_TABLE).ListColumns(mod_InstallVentilation.NOM_COL_SOUS_AVANT_VENTILATION).index
        On Error GoTo 0

        If colCatAvantVenFeuille = 0 Or colSousAvantVenFeuille = 0 Then
            MsgBox "Les colonnes 'CategorieAvantVentilation' / 'SousCategorieAvantVentilation' sont introuvables dans TblOperations." & vbCrLf & _
                   "L'import continue, mais l'annulation d'une ventilation ne pourra pas restaurer " & _
                   "la categorie d'origine pour les operations qui viennent d'etre importees." & vbCrLf & _
                   "Executez PreparerPhase4Ventilation (mod_InstallVentilation) pour les ajouter.", vbExclamation, "Colonnes manquantes"
        Else
            For kCtrl = 1 To nbAjoutees
                wsDonnees.Cells(premiereLigneEcriture + kCtrl - 1, colCatAvantVenFeuille).value = catAvantVentilCtrl(kCtrl)
                wsDonnees.Cells(premiereLigneEcriture + kCtrl - 1, colSousAvantVenFeuille).value = sousAvantVentilCtrl(kCtrl)
            Next kCtrl
        End If
    End If

    ' --- ÉTAPE 8quater (ajout du 11/10/2026, pointage des operations) ---
    ' Ecrit "Oui" dans la colonne Pointe des lignes que l'operateur a pointees pendant le
    ' controle des categories. Les autres restent vides (= a analyser). Comme pour
    ' SousCategorie, cette etape doit rester AVANT le tri de l'ETAPE 9 (lignes contigues).
    ' Si la colonne n'est pas installee, on previent sans bloquer l'import.
    If nbAjoutees > 0 Then
        If IsArray(pointeFinaleCtrl) Then
            Dim nbPointeesImport As Long, colPointeFeuille As Long
            For kCtrl = 1 To nbAjoutees
                If pointeFinaleCtrl(kCtrl) Then nbPointeesImport = nbPointeesImport + 1
            Next kCtrl
            If nbPointeesImport > 0 Then
                On Error Resume Next
                colPointeFeuille = 0
                colPointeFeuille = wsDonnees.ListObjects(NOM_TABLE).ListColumns(mod_VarGlobales.NOM_COL_POINTE).index
                On Error GoTo 0
                If colPointeFeuille = 0 Then
                    MsgBox "La colonne 'Pointe' est introuvable dans TblOperations." & vbCrLf & _
                           "L'import continue, mais le pointage n'a pas ete enregistre pour " & _
                           nbPointeesImport & " operation(s)." & vbCrLf & _
                           "Executez AjouterColonnePointage (mod_Pointage) puis pointez-les via RechercherOperations.", _
                           vbExclamation, "Colonne manquante"
                Else
                    For kCtrl = 1 To nbAjoutees
                        If pointeFinaleCtrl(kCtrl) Then
                            wsDonnees.Cells(premiereLigneEcriture + kCtrl - 1, colPointeFeuille).value = mod_VarGlobales.POINTE_OUI
                        End If
                    Next kCtrl
                End If
            End If
        End If
    End If

    ' --- ÉTAPE 9 : conversion/redimensionnement en tableau Excel, tri et filtre --
    Dim tbl As ListObject
    Dim wsDerniereLigneFinale As Long, plageComplete As Range
    Dim wsDerniereColonneFinale As Long

    ' IMPORTANT : on NE PEUT PAS utiliser NB_COLONNES (figé à 14, nombre de
    ' colonnes fournies par l'import OFX/CSV) pour dimensionner le tableau ici.
    ' TblOperations compte maintenant plus de colonnes (StatutSante, Franchise,
    ' SoldeSante, DepassementHoraires, CommentaireSante, Bénéficiaire, ajoutées
    ' manuellement pour le suivi santé). Si on redimensionnait la table à
    ' NB_COLONNES=14, ces colonnes supplémentaires seraient retirées de la
    ' définition du tableau structuré (tbl.ListColumns("StatutSante") cesserait
    ' de fonctionner), même si les valeurs restaient physiquement dans la
    ' feuille. On calcule donc ici la largeur RÉELLE utilisée en ligne 1.
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
        wsDonnees.Activate   ' garantit que la feuille est active avant AutoFit (évite l'erreur 1004)
        wsDonnees.Range(wsDonnees.Cells(1, 1), wsDonnees.Cells(wsDerniereLigneFinale, wsDerniereColonneFinale)).Columns.AutoFit
        tbl.ShowAutoFilter = True
        
        With tbl.Sort
            .SortFields.Clear
            .SortFields.Add Key:=tbl.ListColumns(COL_DATE_COMPTABLE).Range, Order:=xlDescending
            .Header = xlYes
            .Apply
        End With
    End If

    ' --- ÉTAPE 9ter-0 (ajout du 11/10/2026, sécurité) : suppression des fichiers sources ---
    ' Les données sont désormais dans le classeur (protégé par passphrase) : l'export bancaire
    ' OFX et son CSV complémentaire ne doivent pas rester en clair sur le disque. Ils sont
    ' écrasés puis supprimés. Un import abandonné ou en erreur (Exit Sub plus haut) ne supprime
    ' RIEN : l'opérateur peut recommencer avec les mêmes fichiers.
    Dim msgSuppr As String
    msgSuppr = SupprimerFichierSensible(cheminOFX)
    If cheminCSV <> "" Then msgSuppr = msgSuppr & SupprimerFichierSensible(cheminCSV)

    ' --- ÉTAPE 9bis-0 : mémorisation du dernier import (toujours effectuée) ---
    ' Écriture rapide, sans affichage, dans la feuille technique très masquée
    ' "TechDernierImport". Elle alimente le bouton "Dernier import" de la feuille
    ' Synthese (mod_Actions.RchDernierImport), que l'opérateur peut utiliser à tout moment.
    If nbAjoutees > 0 Then
        mod_DernierImport.MemoriserDernierImport listeIDsImportes, nbAjoutees
    End If

    ' --- ÉTAPE 9bis-1 (ajout du 09/10/2026) : échéancier ---------------------------
    ' Complète les occurrences des règles de récurrence puis propose le rapprochement
    ' des opérations qui viennent d'être importées avec les échéances à venir (voir
    ' mod_Echeancier). Une échéance rapprochée et validée disparaît de la liste. Le
    ' solde prévisionnel de la feuille Synthese est fait de formules : il se met à jour
    ' tout seul. Sans effet tant que l'échéancier n'est pas installé.
    If nbAjoutees > 0 Then
        mod_Echeancier.TraiterApresImport listeIDsImportes, nbAjoutees
    Else
        mod_Echeancier.MettreAJourEcheancier
    End If

    ' --- ÉTAPE 9bis : suivi santé (désormais À LA DEMANDE, et non automatique) ---
    ' AJOUT du 03/10/2026 (demande opérateur) : l'import ne doit plus enchaîner
    ' systématiquement le traitement de santé (potentiellement long, avec un
    ' formulaire à remplir). L'opérateur peut le lancer plus tard via le bouton
    ' "Traitement des donnees de sante" de la feuille Synthese
    ' (mod_FormulairesNotes.RetraiterSuiviSante). Ici, on lui DEMANDE simplement
    ' s'il souhaite le lancer immédiatement, en deux questions distinctes
    ' (il peut répondre oui à la première sans répondre oui à la seconde, mais
    ' pas l'inverse; voir plus bas) :
    '   1. Vérification/réparation du découpage du champ Notes (rapprochement
    '      des clés Notes avec les consultations via frm_RapprochementNotes).
    '   2. Calcul du suivi santé (StatutSante/SoldeSante) et formulaire des cas
    '      restant à compléter par l'opérateur (Beneficiaire, Tiers, Franchise,
    '      Depassement d'honoraires). Cette étape nécessite l'étape 1; sinon,
    '      des dates de consultation non résolues fausseraient le calcul. On ne
    '      pose donc la deuxième question que si l'opérateur a répondu oui à la première.
    Dim repSante1 As VbMsgBoxResult, repSante2 As VbMsgBoxResult
    Dim santeEtape1Faite As Boolean

    repSante1 = MsgBox("Voulez-vous traiter/revoir les catégories des opérations affectées à Santé, Prévoyance ?", _
                        vbYesNo + vbQuestion, "Suivi santé")
    If repSante1 = vbYes Then
        mod_FormulairesNotes.VerifierNotesSante
        santeEtape1Faite = True
    End If

    If santeEtape1Faite Then
        repSante2 = MsgBox("Voulez-vous réaliser le rapprochement des opérations de santé ?", _
                            vbYesNo + vbQuestion, "Suivi santé")
        If repSante2 = vbYes Then
            ' TraiterCasSuiviSante relance elle-même CalculerSuiviSante en premier
            ' (voir mod_SuiviSanteFormulaire.bas) : inutile de l'appeler une seconde fois ici.
            mod_SuiviSanteFormulaire.TraiterCasSuiviSante
        End If
    End If

    ' --- ÉTAPE 9ter : affichage des opérations importées (désormais À LA DEMANDE) ---
    ' AJOUT du 03/10/2026 (demande opérateur) : même logique que ci-dessus; on ne
    ' bascule plus automatiquement vers l'écran de recherche, on le propose.
    Dim repAfficher As VbMsgBoxResult
    Dim unEcranOuvertSurDemande As Boolean   ' Vrai si la réponse à la question 2 ou 3 ci-dessus est Oui
    unEcranOuvertSurDemande = (santeEtape1Faite And repSante2 = vbYes)

    If nbAjoutees > 0 Then
        repAfficher = MsgBox("Voulez-vous afficher les opérations importées ?", _
                              vbYesNo + vbQuestion, "Import")
        If repAfficher = vbYes Then
            mod_RechercheOperations.RechercherOperations "DernierImport"
            unEcranOuvertSurDemande = True
        End If
    End If

    ' --- ÉTAPE 10 : rapport final -----------------------------------------------
    MsgBox "Import termine." & vbCrLf & vbCrLf & _
           "Transactions OFX lues : " & nbLues & vbCrLf & _
           "Deja presentes (ignorees) : " & nbDoublons & vbCrLf & _
           "Nouvelles lignes ajoutees : " & nbAjoutees & vbCrLf & _
           "  dont categorie reconnue dans le MEMO OFX : " & nbMemoReconnus & vbCrLf & _
           "  dont MEMO non reconnu (laisse tel quel dans Notes) : " & nbMemoNonReconnus & vbCrLf & _
           "     - categorie reprise du CSV : " & nbReperCSV & vbCrLf & _
           "     - sans categorie (a ranger) : " & nbSansCategorie & vbCrLf & _
           "  dont categorie ambigue (CSV) traitee via le formulaire : " & nbAmbigus & vbCrLf & _
           IIf(cheminCSV <> "", "  ecarts OFX/CSV (le MEMO fait foi) : " & nbDivergencesCSV, "  CSV non utilise") & msgSuppr, _
           vbInformation, "Import OFX"

    ' --- ÉTAPE 11 : retour à Synthese (demande opérateur du 03/10/2026) -----------
    ' Si l'opérateur a demandé à voir un écran particulier (recherche ou suivi santé),
    ' on le laisse affiché : inutile de le quitter aussitôt. Sinon (toutes les questions
    ' ont reçu une réponse négative ou aucune ligne n'a été ajoutée), on revient
    ' explicitement à Synthese plutôt que de laisser l'interface sur la dernière feuille
    ' active par hasard.
    If Not unEcranOuvertSurDemande Then
        If mod_VarGlobales.wsSynthese Is Nothing Then
            Set mod_VarGlobales.wsSynthese = mod_Criteres.GetFeuille(mod_VarGlobales.NOM_FEUILLE_MENU)
        End If
        mod_InstallCommun.AfficherSynthese      ' Synthese est masquee pendant les formulaires du controle
    End If

End Sub


' ----------------------------------------------------------------------------
' FONCTIONS UTILITAIRES PRIVÉES
' ----------------------------------------------------------------------------

' Écrase le contenu d'un fichier avec des zéros puis le supprime (aucune copie dans la corbeille).
' Renvoie "" si tout s'est bien passé, sinon une ligne d'avertissement à ajouter au rapport
' (fichier verrouillé, en lecture seule sur un partage, etc.) : l'opérateur doit alors le supprimer lui-même.
' NB : sur SSD, l'écrasement n'offre pas de garantie absolue (répartition de l'usure) ; la
' protection la plus sûre reste le chiffrement du disque (BitLocker).
Public Function SupprimerFichierSensible(ByVal chemin As String) As String
    Dim f As Integer, taille As Long, reste As Long, bloc As Long
    Dim zeros() As Byte

    On Error GoTo Echec
    If Len(Dir$(chemin)) = 0 Then Exit Function
    SetAttr chemin, vbNormal
    taille = FileLen(chemin)
    If taille > 0 Then
        ReDim zeros(0 To 65535)                    ' tableau d'octets initialisé à 0
        f = FreeFile
        Open chemin For Binary Access Write As #f
        reste = taille
        Do While reste > 0
            bloc = IIf(reste > 65536, 65536, reste)
            If bloc < 65536 Then ReDim zeros(0 To bloc - 1)
            Put #f, , zeros
            reste = reste - bloc
        Loop
        Close #f
    End If
    Kill chemin
    Exit Function

Echec:
    On Error Resume Next
    Close #f
    SupprimerFichierSensible = vbCrLf & vbCrLf & "ATTENTION : le fichier source n'a pas pu etre supprime, supprimez-le vous-meme :" & _
                               vbCrLf & chemin
End Function

' Ouvre la boîte de dialogue de sélection de fichier. Renvoie "" si l'opération est annulée.
Public Function ChoisirFichier(filtre As String, Titre As String) As String
    Dim chemin As String
    chemin = Application.GetOpenFilename(filtre, , Titre)
    If chemin = "False" Then
        ChoisirFichier = ""
    Else
        ChoisirFichier = chemin
    End If
End Function

' Lit le texte d'une balise enfant (ex. : <TRNAMT>) à l'intérieur d'un nœud OFX.
Public Function LireValeurNoeud(parentNoeud As Object, nomBalise As String) As String
    Dim collectionEnfants As Object
    Set collectionEnfants = parentNoeud.getElementsByTagName(nomBalise)
    If collectionEnfants.Length = 0 Then
        LireValeurNoeud = ""
    Else
        LireValeurNoeud = Trim(collectionEnfants.Item(0).Text)
    End If
End Function

' Convertit une date OFX "AAAAMMJJHHMMSS.mmm" en date Excel.
Public Function ConvertirDateOFX(chaineDate As String) As Date
    If Len(chaineDate) < 8 Then
        ConvertirDateOFX = 0
        Exit Function
    End If
    ConvertirDateOFX = DateSerial( _
        CInt(Mid(chaineDate, 1, 4)), CInt(Mid(chaineDate, 5, 2)), CInt(Mid(chaineDate, 7, 2)))
End Function

' Même conversion mais renvoie Empty (cellule vide) si la date est absente
' (utilisée pour Date_Operation, souvent non renseignée).
Public Function ConvertirDateOFXVariant(chaineDate As String) As Variant
    If Len(chaineDate) < 8 Then
        ConvertirDateOFXVariant = Empty
    Else
        ConvertirDateOFXVariant = ConvertirDateOFX(chaineDate)
    End If
End Function

' Montant OFX : toujours écrit avec un POINT décimal -> Val() l'interprète
' correctement, quels que soient les paramètres régionaux de Windows.
Private Function ConvertirMontantOFX(chaineMontant As String) As Double
    ConvertirMontantOFX = Val(chaineMontant)
End Function

' Montant CSV : écrit avec une VIRGULE décimale (format français) -> on la
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

' Numéro de chèque : on ne conserve que les valeurs numériques; les autres deviennent vides.
Public Function NettoyerNumCheque(brut As String) As Variant
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

' Extrait un segment d'un texte séparé par un caractère donné (index de base 0).
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

' Reconnaît, au début du MEMO OFX, une catégorie source connue (colonne CategorieSource de
' TblCategories) sous la forme "<catégorie> - <notes>" ou "<catégorie>" seule. Renvoie True
' et alimente categorie (graphie du tableau) et notes (le reste, "" si MEMO = catégorie seule).
' Renvoie False si aucune catégorie connue ne correspond : categorie/notes sont alors vidés et
' l'appelant doit conserver le MEMO tel quel. En cas de plusieurs correspondances, la plus
' longue l'emporte (ex. "Frais, remb santé" plutôt qu'un éventuel préfixe plus court).
Public Function ExtraireCategorieMemo(ByVal memo As String, ByRef sources() As String, ByVal nbSources As Long, _
                                       ByRef categorie As String, ByRef notes As String) As Boolean
    Dim t As String, k As Long, src As String, meilleure As String, lg As Long

    categorie = ""
    notes = ""
    t = Trim$(memo)
    If t = "" Or nbSources = 0 Then Exit Function

    For k = LBound(sources) To UBound(sources)
        src = sources(k)
        lg = Len(src)
        If lg > Len(meilleure) Then
            If StrComp(t, src, vbTextCompare) = 0 Then
                meilleure = src
            ElseIf Len(t) >= lg + 3 Then
                If StrComp(Left$(t, lg + 3), src & " - ", vbTextCompare) = 0 Then meilleure = src
            End If
        End If
    Next k

    If meilleure = "" Then Exit Function

    categorie = meilleure
    If Len(t) > Len(meilleure) Then notes = Trim$(Mid$(t, Len(meilleure) + 4))
    ExtraireCategorieMemo = True
End Function

' Repli CSV : retire du MEMO OFX le préfixe "<catégorie source> - " quand la catégorie vient du
' CSV (categorie) ou, si le rapprochement est ambigu, de l'une des candidates
' (candidatsAmbigus, séparées par ";"). Un MEMO égal à la catégorie seule devient "".
' Sans effet si aucune catégorie ne correspond au début du MEMO.
Private Function RetirerPrefixeCategorieMemo(ByVal memo As String, ByVal categorie As String, ByVal candidatsAmbigus As String) As String
    Dim t As String, liste() As String, k As Long, cat As String
    t = Trim$(memo)
    RetirerPrefixeCategorieMemo = t
    If t = "" Then Exit Function

    If categorie <> "" Then
        ReDim liste(0 To 0)
        liste(0) = categorie
    ElseIf candidatsAmbigus <> "" Then
        liste = Split(candidatsAmbigus, ";")
    Else
        Exit Function
    End If

    For k = LBound(liste) To UBound(liste)
        cat = Trim$(liste(k))
        If cat <> "" Then
            If StrComp(t, cat, vbTextCompare) = 0 Then
                RetirerPrefixeCategorieMemo = ""
                Exit Function
            ElseIf StrComp(Left$(t, Len(cat) + 3), cat & " - ", vbTextCompare) = 0 Then
                RetirerPrefixeCategorieMemo = Trim$(Mid$(t, Len(cat) + 4))
                Exit Function
            End If
        End If
    Next k
End Function

' Retire d'une note santé un préfixe "<sous-catégorie> - " placé devant la clé
' ("Frais, remb santé - 20260807;..." -> "20260807;..."). Sans effet si la note ne
' commence pas par ce préfixe suivi d'une clé commençant par une date AAAAMMJJ.
Public Function RetirerPrefixeSanteNotes(ByVal texte As String) As String
    Dim t As String, reste As String
    t = Trim(texte)
    RetirerPrefixeSanteNotes = t
    If LCase$(Left$(t, Len(mod_VarGlobales.SOUS_CATEGORIE_SANTE))) <> LCase$(mod_VarGlobales.SOUS_CATEGORIE_SANTE) Then Exit Function
    reste = LTrim$(Mid$(t, Len(mod_VarGlobales.SOUS_CATEGORIE_SANTE) + 1))
    If Left$(reste, 1) <> "-" Then Exit Function
    reste = Trim$(Mid$(reste, 2))
    If EstDateValide(SegmentTexte(reste, ";", 0)) Then RetirerPrefixeSanteNotes = reste
End Function

' Teste si un texte représente une date valide et non absurde.
Public Function EstDateValide(texte As String) As Boolean
    If Trim(texte) = "" Then
        EstDateValide = False
    ElseIf Len(texte) = 8 And IsNumeric(texte) Then ' 1. Vérification : la chaîne doit contenir 8 caractères numériques
        
        Dim Annee As Integer
        Dim Mois As Integer
        Dim jour As Integer
        
        ' Extraction des parties
        Annee = CInt(Left(texte, 4))
        Mois = CInt(Mid(texte, 5, 2))
        jour = CInt(Right(texte, 2))
        
        ' 2. Vérification supplémentaire : les mois et les jours sont-ils cohérents ?
        If Mois >= 1 And Mois <= 12 And jour >= 1 And jour <= 31 And Annee >= 1900 Then
            EstDateValide = True
        Else
            EstDateValide = False
        End If
        
    Else
        ' Cas où le format AAAAMMJJ n'est pas respecté.
        EstDateValide = False
    End If
End Function


' Calcule la colonne Budget selon la règle métier :
' MISE À JOUR du 03/10/2026 (demande opérateur) : la règle n'est plus codée en dur
' ici (seul le cas DRFIP OCCITANIE ET HTE / Salaire bénéficiait d'un décalage).
' Elle est désormais lue dans le tableau "TblDecalagesBudget" (feuille Param,
' voir mod_DecalagesBudget.bas) : Tiers/Catégorie/SousCatégorie y sont comparés
' à des critères modifiables par l'opérateur sans toucher au code, et le nombre
' de mois à décaler (+1, -1, 0...) est appliqué ici. Une opération qui ne
' correspond à AUCUNE ligne du tableau garde le comportement d'origine (aucun
' décalage, budget = mois de l'opération).
' MISE À JOUR bis du 03/10/2026 (demande opérateur) : fonction rendue PUBLIQUE, et le
' calcul de la date elle-même (DateSerial) est déplacé dans
' mod_DecalagesBudget.CalculerDateBudget. Objectif : garantir qu'il n'existe
' qu'UNE SEULE fonction, dans tout le projet, qui transforme une date comptable
' + un décalage en date de budget - que ce soit ici, à l'import, ou plus tard
' pour le décalage manuel d'une opération précise (voir
' mod_DecalagesBudget.AppliquerDecalageManuel). Le calcul reste identique à
' l'ancienne version : AUCUNE régression attendue sur l'import.
Public Function CalculerBudget(dateComptable As Date, tiers As String, categorie As String, sousCategorie As String) As Date
    Dim decalage As Long
    decalage = mod_DecalagesBudget.ObtenirDecalageBudget(tiers, categorie, sousCategorie)
    CalculerBudget = mod_DecalagesBudget.CalculerDateBudget(dateComptable, decalage)
End Function

' Recherche l'index d'une colonne à partir de son nom d'en-tête (recherche exacte).
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
'   clé "Date|Montant|Tiers" -> Collection des catégories trouvées pour cette clé.
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
    flux.Charset = "utf-16"   ' le CSV est encodé en UTF-16 malgré son extension .csv
    flux.Open
    flux.LoadFromFile chemin
    contenuBrut = flux.ReadText
    flux.Close

    ' Retire un éventuel caractère BOM invisible au tout début du fichier.
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

' À partir d'une collection de catégories candidates pour une même clé :
' - si toutes identiques -> renvoie cette valeur, estAmbigu = False
' - si différentes -> renvoie "", estAmbigu = True, et la liste des valeurs
'   distinctes (séparées par ";") dans texteCandidats, prêtes pour le formulaire
'   de résolution.
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
