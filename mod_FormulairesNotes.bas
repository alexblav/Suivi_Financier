Option Explicit

' =====================================================================================
' MODULE : mod_FormulairesNotes
'
' RÔLE (phase 4b du chantier "Suivi Santé") :
'   Logique complète des deux feuilles construites en phase 4a
'   (mod_InstallFormulairesNotes) : frm_RapprochementNotes et frm_GenerationCle.
'
'   Point d'entrée unique : VerifierNotesSante(). Appelée automatiquement par
'   mod_ImportOFX juste avant CalculerSuiviSante, mais peut aussi être
'   relancée manuellement (Ctrl+G) : VerifierNotesSante
'
'   DÉROULEMENT :
'     1. On parcourt TOUTE TblOperations (lignes "Frais, remb santé").
'     2. Pour chaque ligne dont Date_consult vaut encore la "sentinelle"
'        (02/01/1900, voir mod_ImportOFX), on ouvre frm_RapprochementNotes :
'        l'opérateur filtre en cascade (Date -> Spécialité -> Bénéficiaire ->
'        Montant) parmi les clés Notes déjà valides ailleurs, et valide.
'     3. Si aucune correspondance : bouton "Pas de correspondance" ->
'        frm_GenerationCle, où l'opérateur crée une nouvelle clé.
'     4. Dans les deux cas, la clé est écrite dans Notes pour cette ligne, et
'        Date_consult/Spe_Consult/Bénéficiaire sont recalculés immédiatement
'        (mêmes formules que mod_ImportOFX, étape 7).
'
'   Même verrou modal que les autres formulaires du classeur : la variable
'   Public g_SaisieEnCours (déjà déclarée dans mod_ResolutionCategories) est
'   RÉUTILISÉE telle quelle.
'
'   PARTICULARITÉ : comme ce classeur évite les UserForms (problème de DPI déjà
'   rencontré) et que le code-behind de la feuille (Worksheet_Change) nécessite
'   de connaître son "nom de code" VBA, non garanti à l'avance ici, la cascade
'   de filtres est gérée par SONDAGE : la boucle d'attente modale
'   (Do While g_SaisieEnCours : DoEvents : Loop) vérifie à chaque passage si
'   Date/Spécialité/Bénéficiaire ont changé depuis le précédent, puis met à jour
'   les listes suivantes si nécessaire. Cette méthode est fiable et ne nécessite
'   aucun code-behind.
'
' À PROPOS DES ACCENTS : les textes affichés sont construits via la fonction FR()
' (recopiée ici à l'identique, comme dans les autres modules du chantier). Les
' commentaires du fichier sont encodés en UTF-8.
'
' COMMENT TESTER :
'   Ctrl+G, taper VerifierNotesSante, Entrée.
' =====================================================================================


' --- Une clé "Notes" valide, décodée, utilisée pour construire les listes de
'     filtres en cascade de frm_RapprochementNotes. ---
Private Type TCandidatCle
    DateTexte As String     ' au format jj/mm/aaaa, pour l'affichage et le filtre
    specialite As String    ' segment 1 de la clé
    beneficiaire As String  ' segment 2 de la clé
    Montant As String       ' segment 3 de la clé, tel quel (texte)
    CleBrute As String      ' la valeur complète du champ Notes
End Type

' Mémorise le bouton cliqué ("Valider" ou "PasDeCorrespondance" pour
' frm_RapprochementNotes; "Valider" pour frm_GenerationCle), afin que le code
' ayant lancé la boucle d'attente sache quoi faire lorsqu'elle se termine.
Private derniereAction As String

' =====================================================================================
' MACRO PRINCIPALE - point d'entrée de la phase 4b
' =====================================================================================
' =====================================================================================
' RetraiterSuiviSante : enchaîne les trois étapes du suivi santé SANS réimporter
' de fichier (utile après une correction manuelle dans Import_data, par exemple
' si une catégorie est modifiée à la main parce que la source a changé et que
' le rapprochement par FITID ne l'aurait pas détecté).
' Ctrl+G : RetraiterSuiviSante
' =====================================================================================
Public Sub RetraiterSuiviSante()
    VerifierNotesSante
    ' mod_SuiviSante.CalculerSuiviSante est lancée par la macro suivante
    mod_SuiviSanteFormulaire.TraiterCasSuiviSante
End Sub


Public Sub VerifierNotesSante()

    ' PHASE 6 : on traite maintenant DEUX sources de lignes "santé" en
    ' attente de clé : les lignes de TblOperations (SousCategorie) ET les
    ' lignes de TblVentilations (SousCategorie également, colonne ajoutée en
    ' phase 4). Chaque liste est traitée séparément (numérotation continue
    ' dans le compteur affiché à l'opérateur), mais le mécanisme est
    ' rigoureusement identique dans les 2 cas.
    '
    ' MODIFIÉ le 07/10/2026 (bouton "Changer la catégorie") : la liste des lignes
    ' de TblVentilations n'est plus calculée une seule fois au départ. Elle est
    ' établie APRÈS le traitement des lignes de TblOperations, puis RELUE chaque
    ' fois qu'une ventilation est réécrite ou supprimée depuis cet écran. Raison :
    ' réécrire une ventilation (mod_Ventilation.VenTerminer) supprime ses lignes de
    ' TblVentilations puis les recrée À LA FIN du tableau. Toutes les lignes situées
    ' après se décalent donc d'un cran ou plus : une liste de numéros de lignes
    ' calculée avant la réécriture désignerait ensuite les MAUVAISES lignes (le même
    ' type de piège que la corruption d'ID_Transaction corrigée le 01/10).
    ' Effet secondaire utile : une ligne santé créée en VENTILANT une opération de
    ' TblOperations depuis ce même écran est proposée tout de suite, et non plus
    ' seulement à la relance suivante.
    Dim donneesInitiales As Variant
    Dim nbLignesTable As Long
    Dim nbLigne As Long
    Dim totalATraiter As Long
    Dim numeroEnCours As Long

    Dim pendingOp() As Long
    Dim nbPendingOp As Long

    ' --- Lignes de TblVentilations (voir ListerVentilationsEnAttente) ---
    Dim pendingVen() As Long         ' numéros de ligne dans TblVentilations
    Dim clesVen() As String          ' identifiant stable de chaque ligne (voir plus bas)
    Dim nbPendingVen As Long
    Dim traitesVen As Object         ' lignes déjà présentées pendant CETTE vérification
    Dim idVenModifie As String       ' ID_Transaction de la ventilation réécrite, ou ""
    Dim relancerVen As Boolean
    Dim k As Long

    Set tbl = mod_DonneesTable.GetOperationsTable()
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    mod_Display.RecupIndexCol

    If colSousCategorie = 0 Then
        ' Phase 1 (catégories à deux niveaux) pas encore installée : on ne peut
        ' pas savoir quelles lignes sont "Frais, remb santé". On sort sans
        ' bloquer (VerifierNotesSante est appelée automatiquement à chaque
        ' import : elle ne doit jamais faire échouer l'import).
        Exit Sub
    End If

    donneesInitiales = tbl.DataBodyRange.value
    nbLignesTable = UBound(donneesInitiales, 1)

    ' --- ÉTAPE 1a : lignes de TblOperations en attente d'une clé -------------
    nbPendingOp = 0
    For nbLigne = 1 To nbLignesTable
        If mod_DataStructure.CellText(donneesInitiales(nbLigne, colSousCategorie)) = "Frais, remb sant" & ChrW(233) Then
            If EstLigneSentinelle(donneesInitiales(nbLigne, colDateConsult)) Then
                nbPendingOp = nbPendingOp + 1
                ReDim Preserve pendingOp(1 To nbPendingOp)
                pendingOp(nbPendingOp) = nbLigne
            End If
        End If
    Next nbLigne

    ' --- ÉTAPE 1b : premier comptage des lignes de TblVentilations en attente,
    '     UNIQUEMENT pour annoncer un total de départ à l'opérateur ("1 sur N").
    '     La liste réellement traitée est recalculée à l'ÉTAPE 3. TblVentilations
    '     est FACULTATIVE : si elle n'existe pas encore, nbPendingVen reste à 0. --
    Set traitesVen = CreateObject("Scripting.Dictionary")
    ListerVentilationsEnAttente traitesVen, "", pendingVen, clesVen, nbPendingVen

    totalATraiter = nbPendingOp + nbPendingVen
    If totalATraiter = 0 Then Exit Sub

    ' --- ÉTAPE 2 : lignes de TblOperations ------------------------------------
    ' Ici, les numéros de ligne restent valables jusqu'au bout : changer la
    ' catégorie d'une opération (ou la ventiler) ne déplace ni ne supprime aucune
    ' ligne de TblOperations.
    numeroEnCours = 0
    For nbLigne = 1 To nbPendingOp
        numeroEnCours = numeroEnCours + 1
        ' CORRECTIF du 05/10/2026 (constat opérateur) : on transmet désormais
        ' directement donneesInitiales (tableau déjà chargé ci-dessus) en paramètre,
        ' au lieu de laisser AfficherRapprochementPourLigne lire la variable globale
        ' PARTAGÉE "tblData", qui peut être obsolète ou vide lorsque cette macro est
        ' lancée seule. Même type de problème que pour wsSynthese, corrigé plus tôt
        ' cette semaine : une variable globale partagée entre plusieurs écrans n'est
        ' pas nécessairement initialisée au bon moment.
        AfficherRapprochementPourLigne pendingOp(nbLigne), "O", numeroEnCours, totalATraiter, donneesInitiales
    Next nbLigne

    ' --- ÉTAPE 3 : lignes de TblVentilations, avec relecture après toute
    '     réécriture d'une ventilation (voir l'explication en tête de procédure).
    '
    '     Pour ne pas reproposer, dans la MÊME vérification, une ligne déjà vue
    '     (par exemple "Passer"), on la mémorise dans traitesVen. On ne peut pas
    '     utiliser son numéro de ligne (il peut changer) : on utilise une clé
    '     "ID_Transaction|rang", où le rang est la position de la ligne parmi les
    '     lignes de la MÊME opération (1re part, 2e part...). Ce rang ne bouge pas
    '     quand une AUTRE ventilation est réécrite.
    '     Quand c'est la ventilation de la ligne en cours qui est réécrite, ses
    '     clés sont oubliées (ses lignes santé ont été remises à zéro par la
    '     réécriture et doivent être reproposées), et ses lignes passent EN TÊTE de
    '     la nouvelle liste : l'opérateur revient ainsi sur la même opération,
    '     conformément à son choix du 07/10/2026. ---------------------------------
    idVenModifie = ""
    Do
        relancerVen = False
        ListerVentilationsEnAttente traitesVen, idVenModifie, pendingVen, clesVen, nbPendingVen
        totalATraiter = numeroEnCours + nbPendingVen

        For k = 1 To nbPendingVen
            numeroEnCours = numeroEnCours + 1
            idVenModifie = AfficherRapprochementPourLigne(pendingVen(k), "V", numeroEnCours, totalATraiter, donneesInitiales)
            If idVenModifie <> "" Then
                ' Ventilation réécrite ou supprimée : les numéros de ligne restants
                ' de pendingVen ne sont plus fiables. On arrête ce passage et on
                ' relit TblVentilations. La ligne en cours n'est pas comptée comme
                ' terminée (le compteur affiché reprend au même numéro).
                numeroEnCours = numeroEnCours - 1
                OublierClesVentilation traitesVen, idVenModifie
                relancerVen = True
                Exit For
            End If
            traitesVen(clesVen(k)) = True
        Next k
    Loop While relancerVen

End Sub

' =====================================================================================
' ListerVentilationsEnAttente (ajout du 07/10/2026)
' =====================================================================================
' Relit TblVentilations et renvoie la liste des lignes "Frais, remb santé" encore en
' attente de clé (Date_consult = sentinelle), en ignorant celles déjà présentées
' pendant la vérification en cours (dictionnaire traitesVen).
'   traitesVen    : clés "ID_Transaction|rang" des lignes déjà présentées.
'   idPrioritaire : si non vide, les lignes de CETTE opération sont placées en tête
'                   de liste (retour sur la même opération après une réécriture).
'   pendingVen    : en sortie, numéros de ligne dans TblVentilations (1 = 1re ligne
'                   de données).
'   clesVen       : en sortie, clé "ID_Transaction|rang" de chaque ligne retenue.
'   nbPendingVen  : en sortie, nombre de lignes retenues (0 si TblVentilations est
'                   absente, vide, ou si ses colonnes nécessaires manquent).
' Cette procédure reprend exactement les critères de l'ancienne ÉTAPE 1b de
' VerifierNotesSante ; seuls la clé stable et l'ordre prioritaire sont nouveaux.
Private Sub ListerVentilationsEnAttente(ByVal traitesVen As Object, ByVal idPrioritaire As String, _
                                        ByRef pendingVen() As Long, ByRef clesVen() As String, _
                                        ByRef nbPendingVen As Long)

    Dim tblVen As ListObject
    Dim donneesVen As Variant
    Dim nbLignesVen As Long
    Dim colVenID As Long, colVenSousCat As Long, colVenDateConsult As Long
    Dim rangParOperation As Object
    Dim cleLigne() As String
    Dim idLigne As String
    Dim r As Long, passe As Long
    Dim estPrioritaire As Boolean

    nbPendingVen = 0

    Set tblVen = ObtenirTableVentilationsFN()
    If tblVen Is Nothing Then Exit Sub
    If tblVen.DataBodyRange Is Nothing Then Exit Sub

    colVenID = IndexColSiExisteFN(tblVen, "ID_Transaction")
    colVenSousCat = IndexColSiExisteFN(tblVen, "SousCategorie")
    colVenDateConsult = IndexColSiExisteFN(tblVen, "Date_consult")
    If colVenID = 0 Or colVenSousCat = 0 Or colVenDateConsult = 0 Then Exit Sub

    donneesVen = tblVen.DataBodyRange.value
    nbLignesVen = UBound(donneesVen, 1)

    ' --- Calcul de la clé stable de chaque ligne : "ID_Transaction|rang". Le
    '     dictionnaire compte, opération par opération, combien de lignes ont déjà
    '     été rencontrées (une clé absente vaut "vide", et vide + 1 = 1). ---------
    Set rangParOperation = CreateObject("Scripting.Dictionary")
    ReDim cleLigne(1 To nbLignesVen)
    For r = 1 To nbLignesVen
        idLigne = mod_DataStructure.CellText(donneesVen(r, colVenID))
        rangParOperation(idLigne) = rangParOperation(idLigne) + 1
        cleLigne(r) = idLigne & "|" & rangParOperation(idLigne)
    Next r

    ' --- Deux passages : 1) lignes de l'opération prioritaire, 2) toutes les
    '     autres. Sans opération prioritaire, le 1er passage ne retient rien. -----
    For passe = 1 To 2
        For r = 1 To nbLignesVen
            idLigne = mod_DataStructure.CellText(donneesVen(r, colVenID))
            estPrioritaire = (idPrioritaire <> "" And idLigne = idPrioritaire)
            If (passe = 1) = estPrioritaire Then
                If mod_DataStructure.CellText(donneesVen(r, colVenSousCat)) = mod_VarGlobales.SOUS_CATEGORIE_SANTE Then
                    If EstLigneSentinelle(donneesVen(r, colVenDateConsult)) Then
                        If Not traitesVen.Exists(cleLigne(r)) Then
                            nbPendingVen = nbPendingVen + 1
                            ReDim Preserve pendingVen(1 To nbPendingVen)
                            ReDim Preserve clesVen(1 To nbPendingVen)
                            pendingVen(nbPendingVen) = r
                            clesVen(nbPendingVen) = cleLigne(r)
                        End If
                    End If
                End If
            End If
        Next r
    Next passe

End Sub

' Retire du dictionnaire traitesVen toutes les clés d'une opération donnée
' ("ID_Transaction|1", "ID_Transaction|2"...). Utilisée après la réécriture d'une
' ventilation : ses lignes ont été recréées, il faut pouvoir les reproposer.
' (ajout du 07/10/2026)
Private Sub OublierClesVentilation(ByVal traitesVen As Object, ByVal idTransaction As String)
    Dim cle As Variant
    Dim aRetirer As New Collection
    ' On ne supprime pas une clé PENDANT qu'on parcourt le dictionnaire (cela peut
    ' perturber le parcours) : on les note d'abord, puis on les retire ensuite.
    For Each cle In traitesVen.Keys
        If Left(CStr(cle), Len(idTransaction) + 1) = idTransaction & "|" Then aRetirer.Add CStr(cle)
    Next cle
    For Each cle In aRetirer
        traitesVen.Remove cle
    Next cle
End Sub

' Vrai si Date_consult contient encore la valeur sentinelle (02/01/1900),
' définie par mod_ImportOFX lorsque le découpage de Notes échoue, ou si le champ
' est complètement vide ou non renseigné.
Private Function EstLigneSentinelle(ByVal dateConsultVal As Variant) As Boolean
    EstLigneSentinelle = True
    If IsDate(dateConsultVal) Then
        If CDate(dateConsultVal) <> DateSerial(1900, 1, 2) Then EstLigneSentinelle = False
    End If
End Function

' =====================================================================================
' HELPERS PHASE 6 : accès à TblVentilations, dupliqués ici en local (comme dans
' mod_SuiviSante) afin que ce module puisse compiler même si
' mod_InstallVentilation n'a pas encore été importé.
' =====================================================================================
Private Function ObtenirTableVentilationsFN() As ListObject
    Dim ws As Worksheet
    Dim t As ListObject
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Ventilations")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    Set t = ws.ListObjects("TblVentilations")
    On Error GoTo 0
    Set ObtenirTableVentilationsFN = t
End Function

Private Function IndexColSiExisteFN(ByVal t As ListObject, ByVal nomColonne As String) As Long
    On Error Resume Next
    IndexColSiExisteFN = t.ListColumns(nomColonne).index
    On Error GoTo 0
End Function

' Retrouve, pour une ligne de TblVentilations, la ligne parente correspondante
' dans TblOperations (via ID_Transaction), afin de récupérer les informations
' que TblVentilations ne stocke pas (Date, Tiers, Num_Cheque : une ligne de
' ventilation ne représente qu'une partie d'une opération bancaire et n'a pas
' de date ni de tiers propres). Renvoie False si la ligne parente est introuvable
' (cas qui ne devrait pas se produire, mais on reste prudents).

' (cas qui ne devrait pas se produire, mais on reste prudents).
' MODIFIÉ le 07/10/2026 : nouveau paramètre de sortie FACULTATIF ligneParentTrouvee
' (numéro de la ligne parente dans TblOperations, 0 si introuvable), utilisé par
' ChangerCategorieVentilationSante. Les appels existants, qui ne le précisent pas,
' fonctionnent exactement comme avant.
Private Function TrouverContexteParentVentilation(ByVal ligneVen As Long, ByRef tblVen As ListObject, _
                                                   ByRef donneesVen As Variant, _
                                                   ByRef dateAff As Variant, ByRef tiersAff As String, _
                                                   ByRef chequeAff As Variant, _
                                                   Optional ByRef ligneParentTrouvee As Long = 0) As Boolean
    Dim colVenID As Long
    Dim idCible As String
    Dim tblOp As ListObject
    Dim donneesOp As Variant
    Dim n As Long, i As Long

    TrouverContexteParentVentilation = False
    dateAff = ""
    tiersAff = ""
    chequeAff = ""
    ligneParentTrouvee = 0     ' ajout du 07/10/2026

    colVenID = IndexColSiExisteFN(tblVen, "ID_Transaction")
    If colVenID = 0 Then Exit Function
    idCible = mod_DataStructure.CellText(donneesVen(ligneVen, colVenID))
    If idCible = "" Then Exit Function

    Set tblOp = mod_DonneesTable.GetOperationsTable()
    If tblOp Is Nothing Then Exit Function
    If tblOp.DataBodyRange Is Nothing Then Exit Function

    donneesOp = tblOp.DataBodyRange.value
    n = UBound(donneesOp, 1)
    For i = 1 To n
        If mod_DataStructure.CellText(donneesOp(i, colID)) = idCible Then
            dateAff = donneesOp(i, colDate)
            tiersAff = mod_DataStructure.CellText(donneesOp(i, colTiers))
            chequeAff = donneesOp(i, colCheque)
            ligneParentTrouvee = i     ' ajout du 07/10/2026
            TrouverContexteParentVentilation = True
            Exit Function
        End If
    Next i
End Function




' =====================================================================================
' AfficherRapprochementPourLigne : ouvre frm_RapprochementNotes pour UNE ligne
' (position dans DataBodyRange), attend la decision de l'opérateur, puis
' applique le résultat.
'
' MODIFIÉ le 07/10/2026 : devient une FONCTION (au lieu d'une procédure Sub) afin
' de pouvoir prévenir VerifierNotesSante qu'une ventilation a été réécrite ou
' supprimée par le bouton "Changer la catégorie". Valeur renvoyée :
'   - "" (chaîne vide) dans tous les cas habituels ;
'   - l'ID_Transaction de l'opération dont la ventilation a été réécrite ou
'     supprimée : VerifierNotesSante doit alors relire TblVentilations, car les
'     lignes de ce tableau ont changé de position.
' Les appels existants qui ignorent ce résultat restent valables tels quels.
' =====================================================================================
Private Function AfficherRapprochementPourLigne(ByVal ligneDepense As Long, ByVal Source As String, ByVal numero As Long, ByVal total As Long, ByRef donneesOp As Variant) As String

    Dim ws As Worksheet
    Dim candidats() As TCandidatCle
    Dim nbCandidats As Long
    Dim dernierDate As String, dernierSpecialite As String, dernierBeneficiaire As String
    Dim montantChoisi As String

    ' --- PHASE 6 : données de contexte de l'opération, lues dans la bonne
    '     source (O = TblOperations, V = TblVentilations). Pour une ligne de
    '     ventilation, Date/Tiers/NumCheque n'existent pas dans TblVentilations
    '     elle-même : on les récupère sur la ligne parente de TblOperations,
    '     via ID_Transaction (voir TrouverContexteParentVentilation). -----------
    ' CORRECTIF du 05/10/2026 : donneesOp est désormais reçu en paramètre (tableau
    ' TblOperations déjà chargé par l'appelant), au lieu d'être lu dans la variable
    ' globale partagée tblData (voir le commentaire au point d'appel).
    Dim dateCtx As Variant, tiersCtx As String, chequeCtx As Variant, notesCtx As String, montantCtx As Double
    Dim catCtx As String, sousCtx As String      ' ajout 09/10/2026 : categorie / sous-categorie affichees
    Dim colVenCat As Long, colVenSous As Long
    Dim tblVen As ListObject
    Dim donneesVen As Variant
    Dim colVenNotes As Long, colVenMontant As Long

    If Source = "O" Then
        dateCtx = donneesOp(ligneDepense, colDate)
        tiersCtx = mod_DataStructure.CellText(donneesOp(ligneDepense, colTiers))
        chequeCtx = donneesOp(ligneDepense, colCheque)
        notesCtx = mod_DataStructure.CellText(donneesOp(ligneDepense, colNotes))
        montantCtx = mod_DataStructure.ToDouble(donneesOp(ligneDepense, colMontant))
        catCtx = mod_DataStructure.CellText(donneesOp(ligneDepense, colCategorie))
        If colSousCategorie <> 0 Then sousCtx = mod_DataStructure.CellText(donneesOp(ligneDepense, colSousCategorie))
    Else
        Set tblVen = ObtenirTableVentilationsFN()
        donneesVen = tblVen.DataBodyRange.value
        colVenNotes = IndexColSiExisteFN(tblVen, "Notes")
        colVenMontant = IndexColSiExisteFN(tblVen, "Montant")
        notesCtx = mod_DataStructure.CellText(donneesVen(ligneDepense, colVenNotes))
        montantCtx = mod_DataStructure.ToDouble(donneesVen(ligneDepense, colVenMontant))
        colVenCat = IndexColSiExisteFN(tblVen, "Categorie")
        colVenSous = IndexColSiExisteFN(tblVen, "SousCategorie")
        If colVenCat <> 0 Then catCtx = mod_DataStructure.CellText(donneesVen(ligneDepense, colVenCat))
        If colVenSous <> 0 Then sousCtx = mod_DataStructure.CellText(donneesVen(ligneDepense, colVenSous))
        If Not TrouverContexteParentVentilation(ligneDepense, tblVen, donneesVen, dateCtx, tiersCtx, chequeCtx) Then
            dateCtx = ""
            tiersCtx = "(operation parente introuvable)"
            chequeCtx = ""
        End If
    End If

    Set ws = ThisWorkbook.Worksheets(mod_InstallFormulairesNotes.NOM_FEUILLE_RAPPROCHEMENT)

    ' On affiche la feuille EN PREMIER, avant toute écriture de cellule ou
    ' redéfinition de plage nommée : définir Name.RefersTo lorsque la feuille est
    ' encore xlSheetVeryHidden provoque une erreur 1004.
    ws.Visible = xlSheetVisible
    ws.Activate
    mod_InstallCommun.MasquerQuadrillage   ' quadrillage et en-tetes toujours masques (09/10/2026)

    ' Reconstruit la liste des clés valides À CHAQUE APPEL (et non une seule fois
    ' pour VerifierNotesSante) : une clé générée via frm_GenerationCle pour une
    ' ligne devient ainsi immédiatement disponible pour les suivantes de la même
    ' session, y compris d'une source à l'autre. Une clé validée dans TblOperations
    ' devient candidate pour TblVentilations, et inversement.
    ConstruireListeCandidats candidats, nbCandidats

    ' --- Contexte de l'opération (lecture seule), affiché que la ligne provienne
    '     de TblOperations ou de TblVentilations. En phase 6, tiersCtx précise
    '     "(ventilation)" afin que l'opérateur sache immédiatement d'où vient
    '     la ligne qu'il traite. -------------------------------------------------
    ws.Range("rnDateOp").value = dateCtx
    ws.Range("rnNumCheque").value = chequeCtx
    If Source = "V" Then
        ws.Range("rnTiersOp").value = tiersCtx & FR(" [ligne de ventilation]")
    Else
        ws.Range("rnTiersOp").value = tiersCtx
    End If
    ws.Range("rnNotes").value = notesCtx
    ws.Range("rnCategorie").value = catCtx
    ws.Range("rnSousCategorie").value = sousCtx
    ws.Range("rnMontantOp").value = montantCtx
    ws.Range("rnLigneEnCours").value = ligneDepense
    ws.Range("rnCompteurCas").value = mod_InstallCommun.TexteCompteur(numero, total)   ' format commun "Operation: x/y" (09/10/2026)

    ' --- Remise à zéro des filtres ---
    ' Le format Texte est également forcé ici (pas seulement sur la colonne
    ' technique) : sinon, la valeur sélectionnée dans la liste déroulante pourrait
    ' être reconvertie en date au moment où elle est écrite dans la cellule.
    ws.Range("rnDate").NumberFormat = "@"
    ws.Range("rnSpecialite").NumberFormat = "@"
    ws.Range("rnBeneficiaire").NumberFormat = "@"
    ws.Range("rnMontant").NumberFormat = "@"

    ws.Range("rnDate").value = ""
    ws.Range("rnSpecialite").value = ""
    ws.Range("rnBeneficiaire").value = ""
    ws.Range("rnMontant").value = ""
    ws.Range("rnCleTrouvee").value = ""

    EcrireListeEtRedefinirNom ws, "rnListeDates", mod_InstallFormulairesNotes.RN_COL_LISTE_DATES, ListeDatesDistinctes(candidats, nbCandidats)
    EcrireListeEtRedefinirNom ws, "rnListeSpecialites", mod_InstallFormulairesNotes.RN_COL_LISTE_SPECIALITES, VideListe()
    EcrireListeEtRedefinirNom ws, "rnListeBeneficiaires", mod_InstallFormulairesNotes.RN_COL_LISTE_BENEFICIAIRES, VideListe()
    EcrireListeEtRedefinirNom ws, "rnListeMontants", mod_InstallFormulairesNotes.RN_COL_LISTE_MONTANTS, VideListe()
    EcrireListeEtRedefinirNom ws, "rnListeCles", mod_InstallFormulairesNotes.RN_COL_LISTE_CLES, VideListe()

    AppliquerListeFN ws.Range("rnDate"), "rnListeDates"
    AppliquerListeFN ws.Range("rnSpecialite"), "rnListeSpecialites"
    AppliquerListeFN ws.Range("rnBeneficiaire"), "rnListeBeneficiaires"
    AppliquerListeFN ws.Range("rnMontant"), "rnListeMontants"

    dernierDate = ""
    dernierSpecialite = ""
    dernierBeneficiaire = ""
    derniereAction = ""

    ' --- Boucle d'attente modale + sondage (cascade des 4 filtres) ---
    g_SaisieEnCours = True
    Do While g_SaisieEnCours
        DoEvents

        If CStr(ws.Range("rnDate").value) <> dernierDate Then
            dernierDate = CStr(ws.Range("rnDate").value)
            dernierSpecialite = ""
            dernierBeneficiaire = ""
            ws.Range("rnSpecialite").value = ""
            ws.Range("rnBeneficiaire").value = ""
            ws.Range("rnMontant").value = ""
            ws.Range("rnCleTrouvee").value = ""
            EcrireListeEtRedefinirNom ws, "rnListeSpecialites", mod_InstallFormulairesNotes.RN_COL_LISTE_SPECIALITES, ListeSpecialitesPour(candidats, nbCandidats, dernierDate)
            EcrireListeEtRedefinirNom ws, "rnListeBeneficiaires", mod_InstallFormulairesNotes.RN_COL_LISTE_BENEFICIAIRES, VideListe()
            EcrireListeEtRedefinirNom ws, "rnListeMontants", mod_InstallFormulairesNotes.RN_COL_LISTE_MONTANTS, VideListe()
        End If

        If CStr(ws.Range("rnSpecialite").value) <> dernierSpecialite Then
            dernierSpecialite = CStr(ws.Range("rnSpecialite").value)
            dernierBeneficiaire = ""
            ws.Range("rnBeneficiaire").value = ""
            ws.Range("rnMontant").value = ""
            ws.Range("rnCleTrouvee").value = ""
            EcrireListeEtRedefinirNom ws, "rnListeBeneficiaires", mod_InstallFormulairesNotes.RN_COL_LISTE_BENEFICIAIRES, ListeBeneficiairesPour(candidats, nbCandidats, dernierDate, dernierSpecialite)
            EcrireListeEtRedefinirNom ws, "rnListeMontants", mod_InstallFormulairesNotes.RN_COL_LISTE_MONTANTS, VideListe()
        End If

        If CStr(ws.Range("rnBeneficiaire").value) <> dernierBeneficiaire Then
            dernierBeneficiaire = CStr(ws.Range("rnBeneficiaire").value)
            ws.Range("rnMontant").value = ""
            ws.Range("rnCleTrouvee").value = ""
            EcrireListeMontantsEtCles ws, candidats, nbCandidats, dernierDate, dernierSpecialite, dernierBeneficiaire
        End If

        montantChoisi = Trim(CStr(ws.Range("rnMontant").value))
        If montantChoisi <> "" Then
            ws.Range("rnCleTrouvee").value = TrouverCle(candidats, nbCandidats, dernierDate, dernierSpecialite, dernierBeneficiaire, montantChoisi)
        Else
            ws.Range("rnCleTrouvee").value = ""
        End If
    Loop

    Sheets("Synthese").Activate
    ws.Visible = xlSheetVeryHidden

    Select Case derniereAction
        Case "Valider"
            AppliquerNouvelleCle ligneDepense, Source, CStr(ws.Range("rnCleTrouvee").value)
        Case "PasDeCorrespondance"
            ' "Retour" dans frm_GenerationCle (09/10/2026) : on revient sur la MEME ligne.
            If AfficherGenerationPourLigne(ligneDepense, Source, donneesOp) Then
                AfficherRapprochementPourLigne = AfficherRapprochementPourLigne(ligneDepense, Source, numero, total, donneesOp)
            End If
        Case "Passer"
            ' Aucune action volontaire : la ligne reste inchangée (KO,
            ' Date_consult toujours sentinelle) et sera reproposée au prochain passage.
        Case "ChangerCategorie"
            ' AJOUT du 07/10/2026 : l'opération a peut-être été mal catégorisée à
            ' l'import (ce n'est pas une dépense de santé). Règle validée par
            ' l'opérateur : si elle reste "Frais, remb santé" après le formulaire,
            ' ou s'il annule, on REVIENT SUR LA MÊME LIGNE (nouvel appel de cette
            ' même fonction, avec les mêmes paramètres).
            If Source = "O" Then
                If Not ChangerCategorieOperationSante(ligneDepense) Then
                    AfficherRapprochementPourLigne = AfficherRapprochementPourLigne(ligneDepense, Source, numero, total, donneesOp)
                End If
            Else
                ' Ligne de ventilation : c'est toute la ventilation de l'opération
                ' parente qui est rouverte (option 1 validée par l'opérateur).
                AfficherRapprochementPourLigne = ChangerCategorieVentilationSante(ligneDepense)
                If AfficherRapprochementPourLigne = "" Then
                    ' Annulation : rien n'a bougé dans TblVentilations, le numéro de
                    ' ligne est donc toujours valable -> retour sur la même ligne.
                    AfficherRapprochementPourLigne = AfficherRapprochementPourLigne(ligneDepense, Source, numero, total, donneesOp)
                End If
                ' Sinon, la ventilation a été réécrite ou supprimée : on renvoie son
                ' ID_Transaction à VerifierNotesSante, qui relit TblVentilations et
                ' repropose en tête de liste les lignes santé de cette opération.
            End If
        Case "Sortir"
            End
    End Select

End Function


' =====================================================================================
' BOUTONS de frm_RapprochementNotes
' =====================================================================================
' Ces actions modifient derniereAction afin d'indiquer le choix de l'opérateur.
Public Sub ValiderRapprochementNotes()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mod_InstallFormulairesNotes.NOM_FEUILLE_RAPPROCHEMENT)

    If Trim(CStr(ws.Range("rnCleTrouvee").value)) = "" Then
        MsgBox FR("Aucune cl{e2} trouv{e2}e avec ces filtres. Affine ta recherche, ou clique sur 'Pas de correspondance'."), vbExclamation
        Exit Sub
    End If

    derniereAction = "Valider"
    g_SaisieEnCours = False
End Sub

Public Sub PasDeCorrespondanceNotes()
    derniereAction = "PasDeCorrespondance"
    g_SaisieEnCours = False
End Sub

' "Passer" : aucune modification. La ligne reste "KO" (StatutSante) et
' Date_consult conserve la valeur sentinelle; la ligne sera donc reproposée
' au prochain appel de VerifierNotesSante ou au prochain import.
Public Sub PasserRapprochementNotes()
    derniereAction = "Passer"
    g_SaisieEnCours = False
End Sub

' "Changer la catégorie" (ajout du 07/10/2026) : ferme l'écran de rapprochement et
' laisse AfficherRapprochementPourLigne ouvrir le formulaire de catégorie adapté
' (opération normale ou ligne de ventilation). Voir la section
' "CHANGER LA CATÉGORIE" plus bas pour le détail.
Public Sub ChangerCategorieRapprochementNotes()
    derniereAction = "ChangerCategorie"
    g_SaisieEnCours = False
End Sub


' =====================================================================================
' CHANGER LA CATÉGORIE DEPUIS frm_RapprochementNotes (ajout du 07/10/2026)
' =====================================================================================
' Demande opérateur : une opération mal catégorisée (ou oubliée) lors de l'import en
' "Frais, remb santé" doit pouvoir être recatégorisée directement depuis l'écran de
' rapprochement, afin qu'elle SORTE de l'analyse santé.
'
' Pourquoi cela suffit : tout le moteur santé (VerifierNotesSante, CalculerSuiviSante,
' TraiterCasSuiviSante) sélectionne ses lignes UNIQUEMENT d'après la colonne
' SousCategorie. Une fois celle-ci changée, la ligne n'est plus jamais reprise.
'
' Décisions de l'opérateur (07/10/2026) :
'   - Ligne de TblOperations : formulaire de contrôle des catégories en mode
'     "une seule opération" (fonction partagée avec l'écran de recherche).
'   - Ligne de TblVentilations (option 1) : on rouvre la ventilation complète de
'     l'opération parente, comme "Revoir la ventilation" depuis l'écran de recherche.
'   - Les colonnes de suivi santé d'une opération qui sort de l'analyse sont vidées.
'   - Si la ligne reste en santé, ou si l'opérateur annule : retour sur la même ligne.
' =====================================================================================

' Ligne de TblOperations. Renvoie True si l'opération est SORTIE de l'analyse santé
' (nouvelle sous-catégorie différente de "Frais, remb santé", y compris si elle a été
' ventilée), False si elle y est restée ou si l'opérateur a annulé.
Private Function ChangerCategorieOperationSante(ByVal ligneOp As Long) As Boolean

    Dim tblOp As ListObject
    Dim nouvelleSousCategorie As String

    ChangerCategorieOperationSante = False

    Set tblOp = mod_DonneesTable.GetOperationsTable()
    If tblOp Is Nothing Then Exit Function

    ' Formulaire de contrôle des catégories + écriture du résultat (fonction
    ' partagée avec mod_RechercheOperations.EditerCategorieRO). False = annulation.
    ' Le nettoyage des colonnes de suivi santé (décision opérateur) est fait par
    ' cette même fonction partagée : il n'y a donc rien à vider ici.
    If Not mod_ControleCategories.ModifierCategorieOperation(tblOp, ligneOp) Then Exit Function

    ' Relecture de la sous-catégorie RÉELLEMENT enregistrée (et non d'une copie en
    ' mémoire, qui pourrait être périmée).
    mod_Display.RecupIndexCol
    nouvelleSousCategorie = mod_DataStructure.CellText(tblOp.DataBodyRange.Cells(ligneOp, colSousCategorie).value)

    ' True uniquement si l'opération est bien SORTIE de l'analyse santé.
    ChangerCategorieOperationSante = (nouvelleSousCategorie <> mod_VarGlobales.SOUS_CATEGORIE_SANTE)

End Function

' Ligne de TblVentilations (option 1 validée par l'opérateur) : rouvre la ventilation
' COMPLÈTE de l'opération parente, avec le même formulaire et les mêmes règles que
' "Revoir la ventilation" depuis l'écran de recherche (mod_Ventilation.OuvrirVentilation).
'
' Renvoie l'ID_Transaction de l'opération parente si sa ventilation a été RÉÉCRITE
' (bouton "Terminer") ou SUPPRIMÉE (bouton "Supprimer cette ventilation"), ou une
' chaîne vide si l'opérateur a annulé (rien n'a changé).
'
' À SAVOIR (signalé dans les instructions de l'écran) :
'   - Réécrire une ventilation recrée TOUTES ses lignes : celles qui restent en
'     "Frais, remb santé" repartent à zéro (KO / date sentinelle), même si elles
'     avaient déjà été rapprochées. Limite connue de mod_Ventilation (voir son
'     en-tête), non modifiée ici.
'   - Une ligne qui n'est plus en santé est recréée SANS colonnes de suivi santé
'     (voir mod_Ventilation.AjouterLigneVentilation) : le nettoyage demandé est donc
'     automatique pour une ligne de ventilation, rien de plus à faire ici.
'   - En cas de suppression, l'opération parente reprend sa catégorie d'avant
'     ventilation (même traitement que depuis l'écran de recherche).
Private Function ChangerCategorieVentilationSante(ByVal ligneVen As Long) As String

    Dim tblVen As ListObject
    Dim donneesVen As Variant
    Dim tblOp As ListObject
    Dim idTransaction As String
    Dim dateParent As Variant, tiersParent As String, chequeParent As Variant
    Dim ligneParent As Long
    Dim montantParent As Double
    Dim categorieParent As String, sousCategorieParent As String
    Dim ventilationSupprimee As Boolean

    ChangerCategorieVentilationSante = ""

    Set tblVen = ObtenirTableVentilationsFN()
    If tblVen Is Nothing Then Exit Function
    If tblVen.DataBodyRange Is Nothing Then Exit Function
    donneesVen = tblVen.DataBodyRange.value

    ' Index des colonnes de TblOperations (colID, colDate...), nécessaires à
    ' TrouverContexteParentVentilation et aux lectures ci-dessous.
    mod_Display.RecupIndexCol

    ' Fonction existante réutilisée : elle retrouve l'opération parente via
    ' ID_Transaction (et renvoie maintenant aussi son numéro de ligne).
    If Not TrouverContexteParentVentilation(ligneVen, tblVen, donneesVen, dateParent, tiersParent, chequeParent, ligneParent) Then
        MsgBox FR("Op{e2}ration parente introuvable dans TblOperations : impossible de rouvrir cette ventilation."), vbExclamation
        Exit Function
    End If

    idTransaction = mod_DataStructure.CellText(donneesVen(ligneVen, IndexColSiExisteFN(tblVen, "ID_Transaction")))

    Set tblOp = mod_DonneesTable.GetOperationsTable()
    montantParent = mod_DataStructure.ToDouble(tblOp.DataBodyRange.Cells(ligneParent, colMontant).value)
    categorieParent = mod_DataStructure.CellText(tblOp.DataBodyRange.Cells(ligneParent, colCategorie).value)
    sousCategorieParent = ""
    If colSousCategorie <> 0 Then sousCategorieParent = mod_DataStructure.CellText(tblOp.DataBodyRange.Cells(ligneParent, colSousCategorie).value)

    ' Même appel que mod_RechercheOperations.RevoirVentilationRO.
    If Not mod_Ventilation.OuvrirVentilation(idTransaction, dateParent, tiersParent, "", montantParent, _
                                             categorieParent, sousCategorieParent, ventilationSupprimee) Then
        If Not ventilationSupprimee Then Exit Function    ' annulation : rien n'a changé
        ' Ventilation supprimée : l'opération parente reprend sa catégorie d'avant
        ' ventilation (procédure existante, rendue publique pour l'occasion).
        mod_RechercheOperations.RestaurerCategorieAvantVentilation tblOp, ligneParent
    End If

    ChangerCategorieVentilationSante = idTransaction

End Function

' "Sortir" : aucune modification. La ligne reste "KO" (StatutSante) et
' Date_consult conserve la valeur sentinelle; la ligne sera donc reproposée
' au prochain appel de VerifierNotesSante ou au prochain import.
Public Sub SortirRapprochementNotes()
    derniereAction = "Sortir"
    g_SaisieEnCours = False
End Sub

Public Sub PasserGenerationCle()
    derniereAction = "Passer"
    g_SaisieEnCours = False
End Sub

' Bouton "Retour" (ajout 09/10/2026) : revient sur frm_RapprochementNotes pour la MEME
' operation, sans rien modifier (erreur de clic sur "Pas de correspondance").
Public Sub RetourGenerationCle()
    derniereAction = "Retour"
    g_SaisieEnCours = False
End Sub


' =====================================================================================
' AfficherGenerationPourLigne : ouvre frm_GenerationCle pour UNE ligne.
' Renvoie Vrai si l'operateur a clique sur "Retour" (09/10/2026) : l'appelant doit alors
' reafficher frm_RapprochementNotes pour la meme ligne.
' =====================================================================================
Private Function AfficherGenerationPourLigne(ByVal ligneDepense As Long, ByVal Source As String, ByRef donneesOp As Variant) As Boolean

    Dim ws As Worksheet

    ' --- PHASE 6 : mêmes lectures généralisées O/V que dans
    '     AfficherRapprochementPourLigne (voir ses commentaires). -------------
    ' CORRECTIF du 05/10/2026 : donneesOp est reçu en paramètre. Voir
    ' AfficherRapprochementPourLigne pour le détail du problème corrigé.
    Dim dateCtx As Variant, tiersCtx As String, chequeCtx As Variant, notesCtx As String, montantCtx As Double
    Dim tblVen As ListObject
    Dim donneesVen As Variant
    Dim colVenNotes As Long, colVenMontant As Long

    If Source = "O" Then
        dateCtx = donneesOp(ligneDepense, colDate)
        tiersCtx = mod_DataStructure.CellText(donneesOp(ligneDepense, colTiers))
        chequeCtx = donneesOp(ligneDepense, colCheque)
        notesCtx = mod_DataStructure.CellText(donneesOp(ligneDepense, colNotes))
        montantCtx = mod_DataStructure.ToDouble(donneesOp(ligneDepense, colMontant))
    Else
        Set tblVen = ObtenirTableVentilationsFN()
        donneesVen = tblVen.DataBodyRange.value
        colVenNotes = IndexColSiExisteFN(tblVen, "Notes")
        colVenMontant = IndexColSiExisteFN(tblVen, "Montant")
        notesCtx = mod_DataStructure.CellText(donneesVen(ligneDepense, colVenNotes))
        montantCtx = mod_DataStructure.ToDouble(donneesVen(ligneDepense, colVenMontant))
        If Not TrouverContexteParentVentilation(ligneDepense, tblVen, donneesVen, dateCtx, tiersCtx, chequeCtx) Then
            dateCtx = ""
            tiersCtx = "(operation parente introuvable)"
            chequeCtx = ""
        End If
    End If

    Set ws = ThisWorkbook.Worksheets(mod_InstallFormulairesNotes.NOM_FEUILLE_GENERATION)

    ws.Range("gcDateOp").value = dateCtx
    If Source = "V" Then
        ws.Range("gcTiersOp").value = tiersCtx & FR(" [ligne de ventilation]")
    Else
        ws.Range("gcTiersOp").value = tiersCtx
    End If
    ws.Range("gcNotes").value = notesCtx
    ws.Range("gcNumCheque").value = chequeCtx
    ws.Range("gcLigneEnCours").value = ligneDepense

    ws.Range("gcDateConsult").value = ""
    ws.Range("gcSpecialite").value = ""
    ws.Range("gcBeneficiaire").value = ""
    ws.Range("gcMontant").value = Abs(montantCtx)
    ws.Range("gcCleGeneree").value = ""

    ws.Visible = xlSheetVisible
    ws.Activate
    mod_InstallCommun.MasquerQuadrillage   ' quadrillage et en-tetes toujours masques (09/10/2026)

    ' Listes deroulantes reposees a chaque ouverture (09/10/2026) : ne dependent plus
    ' d'une validation posee une fois pour toutes a l'installation.
    mod_InstallCommun.PoserListeDeroulante ws.Range("gcSpecialite"), "Specialites"
    mod_InstallCommun.PoserListeDeroulante ws.Range("gcBeneficiaire"), "Beneficiaires"

    derniereAction = ""
    g_SaisieEnCours = True
    Do While g_SaisieEnCours
        DoEvents
    Loop

    ws.Visible = xlSheetVeryHidden

    ' Si "Passer" a été cliqué, derniereAction = "Passer" et aucun des blocs
    ' ci-dessous n'est exécuté : rien n'est écrit, la ligne reste inchangée et
    ' sera reproposée au prochain passage.
    If derniereAction = "Valider" Then
        If Trim(CStr(ws.Range("gcCleGeneree").value)) <> "" Then
            AppliquerNouvelleCle ligneDepense, Source, CStr(ws.Range("gcCleGeneree").value)
        End If
    End If

    AfficherGenerationPourLigne = (derniereAction = "Retour")

End Function


' =====================================================================================
' BOUTONS de frm_GenerationCle
' =====================================================================================
Public Sub GenererCleNotes()

    Dim ws As Worksheet
    Dim dateConsult As Variant
    Dim specialite As String, beneficiaire As String, montantTexte As String
    Dim cle As String

    Set ws = ThisWorkbook.Worksheets(mod_InstallFormulairesNotes.NOM_FEUILLE_GENERATION)

    dateConsult = ws.Range("gcDateConsult").value
    specialite = Trim(CStr(ws.Range("gcSpecialite").value))
    beneficiaire = Trim(CStr(ws.Range("gcBeneficiaire").value))

    If Not IsDate(dateConsult) Then
        MsgBox FR("Merci de saisir une Date_consult valide (JJ/MM/AAAA)."), vbExclamation
        Exit Sub
    End If
    If specialite = "" Then
        MsgBox FR("Merci de choisir une sp{e2}cialit{e2}."), vbExclamation
        Exit Sub
    End If
    If beneficiaire = "" Then
        MsgBox FR("Merci de choisir un b{e2}n{e2}ficiaire."), vbExclamation
        Exit Sub
    End If

    montantTexte = mod_DataStructure.CellText(ws.Range("gcMontant").value)

    cle = Format(CDate(dateConsult), "yyyymmdd") & ";" & specialite & ";" & beneficiaire & ";" & montantTexte

    ws.Range("gcCleGeneree").value = cle
    ws.Range("gcCleGeneree").Copy   ' copie dans le presse-papiers (comme Ctrl+C sur la cellule)

End Sub

Public Sub ValiderGenerationCle()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mod_InstallFormulairesNotes.NOM_FEUILLE_GENERATION)

    If Trim(CStr(ws.Range("gcCleGeneree").value)) = "" Then
        MsgBox FR("Clique d'abord sur 'G{e2}n{e2}rer'."), vbExclamation
        Exit Sub
    End If

    derniereAction = "Valider"
    g_SaisieEnCours = False

End Sub


' =====================================================================================
' BOUTONS "+" de frm_GenerationCle (spécialité / bénéficiaire).
' Même principe que AjouterBeneficiaire/AjouterTiersPraticien dans
' mod_SuiviSanteFormulaire; logique dupliquée ici avec le nom de feuille en
' paramètre afin de conserver l'autonomie entre les deux modules.
' =====================================================================================
Public Sub AjouterSpecialite()
    AjouterValeurDansListeFN "Specialites", mod_InstallFormulairesNotes.NOM_FEUILLE_GENERATION, "gcSpecialite"
End Sub

Public Sub AjouterBeneficiaireDepuisGeneration()
    AjouterValeurDansListeFN "Beneficiaires", mod_InstallFormulairesNotes.NOM_FEUILLE_GENERATION, "gcBeneficiaire"
End Sub

Private Sub AjouterValeurDansListeFN(ByVal nomPlage As String, ByVal nomFeuille As String, ByVal nomCelluleCible As String)

    Dim nouvelleValeur As String
    Dim rngListe As Range
    Dim celluleVide As Range
    Dim c As Range
    Dim ws As Worksheet

    nouvelleValeur = Trim(InputBox(FR("Nouvelle valeur {a2} ajouter {a2} la liste '") & nomPlage & "' :", FR("Ajouter une valeur")))
    If nouvelleValeur = "" Then Exit Sub

    On Error Resume Next
    Set rngListe = ThisWorkbook.Names(nomPlage).RefersToRange
    On Error GoTo 0

    If rngListe Is Nothing Then
        MsgBox FR("La plage nomm{e2}e '") & nomPlage & FR("' est introuvable. V{e2}rifie qu'elle existe bien dans Param."), vbCritical
        Exit Sub
    End If

    If Application.WorksheetFunction.CountIf(rngListe, nouvelleValeur) > 0 Then
        MsgBox FR("Cette valeur existe d{e2}j{a2} dans la liste."), vbInformation
    Else
        Set celluleVide = Nothing
        For Each c In rngListe.Cells
            If Trim(CStr(c.value)) = "" Then
                Set celluleVide = c
                Exit For
            End If
        Next c

        If Not celluleVide Is Nothing Then
            celluleVide.value = nouvelleValeur
        Else
            Dim nouvelleCellule As Range
            Set nouvelleCellule = rngListe.Cells(rngListe.Cells.count + 1)
            nouvelleCellule.value = nouvelleValeur
            ThisWorkbook.Names(nomPlage).RefersTo = "=" & rngListe.Resize(rngListe.rows.count + 1).Address(External:=True)
        End If
    End If

    Set ws = ThisWorkbook.Worksheets(nomFeuille)
    ws.Range(nomCelluleCible).value = nouvelleValeur

End Sub


' =====================================================================================
' AppliquerNouvelleCle : écrit la clé dans Notes pour une ligne, puis
' recalcule Date_consult / Spe_Consult / Bénéficiaire pour CETTE ligne
' uniquement (même logique que dans mod_ImportOFX, étape 7).
' =====================================================================================
Private Sub AppliquerNouvelleCle(ByVal ligneDepense As Long, ByVal Source As String, ByVal cle As String)

    Dim seg0 As String, seg1 As String, seg2 As String
    Dim tblVen As ListObject

    seg0 = mod_ImportOFX.SegmentTexte(cle, ";", 0)
    seg1 = mod_ImportOFX.SegmentTexte(cle, ";", 1)
    seg2 = mod_ImportOFX.SegmentTexte(cle, ";", 2)

    If Source = "O" Then
        ' --- Écriture dans TblOperations (comportement d'origine, inchangé) ---
        Set tbl = mod_DonneesTable.GetOperationsTable()
        If tbl Is Nothing Then Exit Sub

        tbl.ListColumns("Notes").DataBodyRange.rows(ligneDepense).value = cle

        If mod_ImportOFX.EstDateValide(seg0) Then
            tbl.ListColumns("Date_consult").DataBodyRange.rows(ligneDepense).value = _
                DateSerial(CInt(Left(seg0, 4)), CInt(Mid(seg0, 5, 2)), CInt(Right(seg0, 2)))
        End If

        tbl.ListColumns("Spe_Consult").DataBodyRange.rows(ligneDepense).value = seg1

        If seg1 = "non reprise historique" Then
            tbl.ListColumns("StatutSante").DataBodyRange.rows(ligneDepense).value = "OK"
        End If

        If seg2 <> "" Then
            tbl.ListColumns("Beneficiaire").DataBodyRange.rows(ligneDepense).value = seg2
        End If
    Else
        ' --- PHASE 6 : écriture dans TblVentilations. Mêmes colonnes, mêmes
        '     règles, mais sur l'AUTRE tableau : une ligne de ventilation
        '     "Frais, remb santé" possède exactement les mêmes colonnes de
        '     colonnes de suivi que TblOperations (voir
        '     mod_InstallVentilation.PreparerTableVentilations), ce qui permet
        '     de dupliquer cette logique sans rien inventer de nouveau. --------
        Set tblVen = ObtenirTableVentilationsFN()
        If tblVen Is Nothing Then Exit Sub

        tblVen.ListColumns("Notes").DataBodyRange.rows(ligneDepense).value = cle

        If mod_ImportOFX.EstDateValide(seg0) Then
            tblVen.ListColumns("Date_consult").DataBodyRange.rows(ligneDepense).value = _
                DateSerial(CInt(Left(seg0, 4)), CInt(Mid(seg0, 5, 2)), CInt(Right(seg0, 2)))
        End If

        tblVen.ListColumns("Spe_Consult").DataBodyRange.rows(ligneDepense).value = seg1

        If seg1 = "non reprise historique" Then
            tblVen.ListColumns("StatutSante").DataBodyRange.rows(ligneDepense).value = "OK"
        End If

        If seg2 <> "" Then
            tblVen.ListColumns("Beneficiaire").DataBodyRange.rows(ligneDepense).value = seg2
        End If
    End If

End Sub


' =====================================================================================
' ConstruireListeCandidats : construit la liste de toutes les clés Notes
' DÉJÀ VALIDÉES (une seule fois par clé distincte) parmi les lignes de santé
' de TblOperations.
' =====================================================================================
Private Sub ConstruireListeCandidats(ByRef candidats() As TCandidatCle, ByRef nbCandidats As Long)

    Dim n As Long, i As Long
    Dim notesTexte As String, seg0 As String, seg1 As String, seg2 As String, seg3 As String
    Dim vues As Object

    nbCandidats = 0

    Set tbl = mod_DonneesTable.GetOperationsTable()
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    mod_Display.RecupIndexCol

    tblData = tbl.DataBodyRange.value
    n = UBound(tblData, 1)

    Set vues = CreateObject("Scripting.Dictionary")

    ' PHASE 6 : la colonne testée est désormais SousCategorie (et non plus
    ' Categorie, qui contient maintenant la catégorie parente).
    If colSousCategorie <> 0 Then
        For i = 1 To n
            If mod_DataStructure.CellText(tblData(i, colSousCategorie)) = "Frais, remb sant" & ChrW(233) Then
                notesTexte = mod_DataStructure.CellText(tblData(i, colNotes))
                AjouterCandidatSiValide notesTexte, vues, candidats, nbCandidats
            End If
        Next i
    End If

    ' --- PHASE 6 : on ajoute aussi les clés DÉJÀ VALIDÉES présentes dans
    '     TblVentilations, afin qu'une clé générée sur une ligne de ventilation
    '     puisse être sélectionnée pour une ligne de TblOperations, et inversement. ---
    Dim tblVen As ListObject
    Dim donneesVen As Variant
    Dim nVen As Long
    Dim colVenSousCat As Long, colVenNotes As Long

    Set tblVen = ObtenirTableVentilationsFN()
    If Not tblVen Is Nothing Then
        If Not tblVen.DataBodyRange Is Nothing Then
            colVenSousCat = IndexColSiExisteFN(tblVen, "SousCategorie")
            colVenNotes = IndexColSiExisteFN(tblVen, "Notes")
            If colVenSousCat <> 0 And colVenNotes <> 0 Then
                donneesVen = tblVen.DataBodyRange.value
                nVen = UBound(donneesVen, 1)
                For i = 1 To nVen
                    If mod_DataStructure.CellText(donneesVen(i, colVenSousCat)) = "Frais, remb sant" & ChrW(233) Then
                        notesTexte = mod_DataStructure.CellText(donneesVen(i, colVenNotes))
                        AjouterCandidatSiValide notesTexte, vues, candidats, nbCandidats
                    End If
                Next i
            End If
        End If
    End If

End Sub

' Petit utilitaire commun aux deux boucles ci-dessus (TblOperations et
' TblVentilations) : découpe une valeur de Notes et l'ajoute à candidats()
' si son segment 0 (date) est valide et si cette valeur n'a pas déjà été vue.
Private Sub AjouterCandidatSiValide(ByVal notesTexte As String, ByRef vues As Object, _
                                    ByRef candidats() As TCandidatCle, ByRef nbCandidats As Long)
    Dim seg0 As String, seg1 As String, seg2 As String, seg3 As String

    seg0 = mod_ImportOFX.SegmentTexte(notesTexte, ";", 0)

    If mod_ImportOFX.EstDateValide(seg0) Then
        If Not vues.Exists(notesTexte) Then
            vues.Add notesTexte, True

            seg1 = mod_ImportOFX.SegmentTexte(notesTexte, ";", 1)
            seg2 = mod_ImportOFX.SegmentTexte(notesTexte, ";", 2)
            seg3 = mod_ImportOFX.SegmentTexte(notesTexte, ";", 3)

            nbCandidats = nbCandidats + 1
            ReDim Preserve candidats(1 To nbCandidats)
            With candidats(nbCandidats)
                .DateTexte = Mid(seg0, 7, 2) & "/" & Mid(seg0, 5, 2) & "/" & Left(seg0, 4)
                .specialite = seg1
                .beneficiaire = seg2
                .Montant = seg3
                .CleBrute = notesTexte
            End With
        End If
    End If
End Sub


' =====================================================================================
' FONCTIONS DE FILTRAGE (construction des listes pour chaque niveau de la cascade)
' =====================================================================================
Private Function ListeDatesDistinctes(ByRef candidats() As TCandidatCle, ByVal nbCandidats As Long) As String()
    Dim resultat() As String
    Dim nb As Long, i As Long
    nb = 0
    For i = 1 To nbCandidats
        If Not ValeurDejaDansListe(resultat, nb, candidats(i).DateTexte) Then
            nb = nb + 1
            ReDim Preserve resultat(1 To nb)
            resultat(nb) = candidats(i).DateTexte
        End If
    Next i
    TrierDatesDecroissant resultat, nb
    ListeDatesDistinctes = resultat
End Function

' Tri à bulles simple (la liste est toujours courte : elle contient les dates de
' consultation distinctes, et non toutes les lignes de TblOperations). La date
' la plus récente apparaît en premier afin de limiter le défilement dans la liste.
Private Sub TrierDatesDecroissant(ByRef dates() As String, ByVal nb As Long)
    Dim i As Long, j As Long
    Dim tmp As String
    For i = 1 To nb - 1
        For j = 1 To nb - i
            If CDate(dates(j)) < CDate(dates(j + 1)) Then
                tmp = dates(j)
                dates(j) = dates(j + 1)
                dates(j + 1) = tmp
            End If
        Next j
    Next i
End Sub

Private Function ListeSpecialitesPour(ByRef candidats() As TCandidatCle, ByVal nbCandidats As Long, ByVal dateFiltre As String) As String()
    Dim resultat() As String
    Dim nb As Long, i As Long
    nb = 0
    If dateFiltre <> "" Then
        For i = 1 To nbCandidats
            If candidats(i).DateTexte = dateFiltre Then
                If Not ValeurDejaDansListe(resultat, nb, candidats(i).specialite) Then
                    nb = nb + 1
                    ReDim Preserve resultat(1 To nb)
                    resultat(nb) = candidats(i).specialite
                End If
            End If
        Next i
    End If
    ListeSpecialitesPour = resultat
End Function

Private Function ListeBeneficiairesPour(ByRef candidats() As TCandidatCle, ByVal nbCandidats As Long, ByVal dateFiltre As String, ByVal specialiteFiltre As String) As String()
    Dim resultat() As String
    Dim nb As Long, i As Long
    nb = 0
    If dateFiltre <> "" And specialiteFiltre <> "" Then
        For i = 1 To nbCandidats
            If candidats(i).DateTexte = dateFiltre And candidats(i).specialite = specialiteFiltre Then
                If Not ValeurDejaDansListe(resultat, nb, candidats(i).beneficiaire) Then
                    nb = nb + 1
                    ReDim Preserve resultat(1 To nb)
                    resultat(nb) = candidats(i).beneficiaire
                End If
            End If
        Next i
    End If
    ListeBeneficiairesPour = resultat
End Function

' Écrit en une seule fois les deux listes parallèles Montant/Clé (une ligne = un candidat).
Private Sub EcrireListeMontantsEtCles(ws As Worksheet, ByRef candidats() As TCandidatCle, ByVal nbCandidats As Long, _
                                       ByVal dateFiltre As String, ByVal specialiteFiltre As String, ByVal beneficiaireFiltre As String)
    Dim montants() As String
    Dim cles() As String
    Dim nb As Long, i As Long

    nb = 0
    If dateFiltre <> "" And specialiteFiltre <> "" And beneficiaireFiltre <> "" Then
        For i = 1 To nbCandidats
            If candidats(i).DateTexte = dateFiltre And candidats(i).specialite = specialiteFiltre _
               And candidats(i).beneficiaire = beneficiaireFiltre Then
                nb = nb + 1
                ReDim Preserve montants(1 To nb)
                ReDim Preserve cles(1 To nb)
                montants(nb) = candidats(i).Montant
                cles(nb) = candidats(i).CleBrute
            End If
        Next i
    End If

    EcrireListeEtRedefinirNom ws, "rnListeMontants", mod_InstallFormulairesNotes.RN_COL_LISTE_MONTANTS, montants
    EcrireListeEtRedefinirNom ws, "rnListeCles", mod_InstallFormulairesNotes.RN_COL_LISTE_CLES, cles
End Sub

' Parcourt la totalité du tableau "candidats()" à la recherche de la clé saisie
Private Function TrouverCle(ByRef candidats() As TCandidatCle, ByVal nbCandidats As Long, _
                             ByVal dateFiltre As String, ByVal specialiteFiltre As String, _
                             ByVal beneficiaireFiltre As String, ByVal montantFiltre As String) As String
    Dim i As Long
    For i = 1 To nbCandidats
        If candidats(i).DateTexte = dateFiltre And candidats(i).specialite = specialiteFiltre _
           And candidats(i).beneficiaire = beneficiaireFiltre And candidats(i).Montant = montantFiltre Then
            TrouverCle = candidats(i).CleBrute
            Exit Function
        End If
    Next i
    TrouverCle = ""
End Function

' Petit utilitaire : la valeur figure-t-elle déjà dans le tableau resultat(1 To nb) ?
Private Function ValeurDejaDansListe(ByRef resultat() As String, ByVal nb As Long, ByVal valeur As String) As Boolean
    Dim j As Long
    For j = 1 To nb
        If resultat(j) = valeur Then
            ValeurDejaDansListe = True
            Exit Function
        End If
    Next j
    ValeurDejaDansListe = False
End Function

' Tableau de chaînes vide (zéro élément), pour réinitialiser une liste.
Private Function VideListe() As String()
    Dim vide() As String
    VideListe = vide
End Function


' =====================================================================================
' EcrireListeEtRedefinirNom : écrit un tableau de valeurs à partir de la ligne 2
' d'une colonne technique, puis redéfinit le nom pour qu'il pointe exactement
' vers les cellules utilisées (même principe que l'extension des plages
' Beneficiaires/Praticiens dans mod_SuiviSanteFormulaire).
' =====================================================================================
Private Sub EcrireListeEtRedefinirNom(ws As Worksheet, ByVal nomListe As String, ByVal colonne As String, ByRef valeurs() As String)

    Dim i As Long
    Dim nb As Long

    On Error Resume Next
    nb = UBound(valeurs) - LBound(valeurs) + 1
    On Error GoTo 0

    ' IMPORTANT : on force la colonne au format TEXTE avant toute écriture.
    ' Sinon, une valeur ressemblant à une date (ex. : "01/12/1900") est
    ' automatiquement convertie en date par Excel dès qu'elle est affectée
    ' à une cellule, comme lors d'une saisie manuelle. Cette conversion fausse
    ' ensuite les comparaisons strictes de texte utilisées dans ce module
    ' (ListeSpecialitesPour, etc.), car CStr() ne restitue pas nécessairement
    ' le texte d'origine à l'identique.
    ws.Range(colonne & "2:" & colonne & "1000").NumberFormat = "@"
    ws.Range(colonne & "2:" & colonne & "1000").ClearContents

    If nb < 1 Then
        ws.Names(nomListe).RefersTo = "=" & ws.Range(colonne & "2").Address(External:=True)
        Exit Sub
    End If

    For i = 1 To nb
        ws.Range(colonne & (1 + i)).value = valeurs(LBound(valeurs) + i - 1)
    Next i

    ws.Names(nomListe).RefersTo = "=" & ws.Range(colonne & "2:" & colonne & (1 + nb)).Address(External:=True)

End Sub

' Applique une liste déroulante de validation qui pointe vers un nom défini.
Private Sub AppliquerListeFN(ByVal rng As Range, ByVal nomPlage As String)
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & nomPlage
    On Error GoTo 0
End Sub


