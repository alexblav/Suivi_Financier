Option Explicit

' =====================================================================================
' MODULE : mod_SuiviSanteFormulaire
'
' RÔLE (PHASE 2b du chantier "Suivi Santé") :
'   Contient toute la logique du formulaire opérateur construit en phase 2a
'   (feuille frm_SuiviSante, module mod_InstallSuiviSanteSheet) :
'     - repérer les dépenses de santé qui nécessitent une action de l'opérateur
'     - les présenter une par une dans le formulaire
'     - écrire les réponses de l'opérateur dans TblOperations
'     - relancer le calcul automatique (mod_SuiviSante) après chaque cas valide
'
'   Ce module reprend le même "verrou modal" que la résolution des catégories
'   ambiguës : la variable Public g_SaisieEnCours (déjà déclarée dans
'   mod_ResolutionCategories) est réutilisée telle quelle, sans être redéclarée ici.
'
' QUEL CAS EST PROPOSE A L'OPERATEUR ?
'   Une dépense de santé est proposée si, après le dernier calcul automatique :
'     - une ligne de dépense existe pour ce groupe (même "Notes")
'     - StatutSante = "KO"
'     - DepassementHoraires = FAUX
'   (Ce sont exactement les groupes que mod_SuiviSante.CalculerSuiviSante
'   retraiterait au prochain import : on anticipe simplement cette liste pour
'   pouvoir la proposer MAINTENANT à l'opérateur plutôt que d'attendre.)
'
' COMMENT TESTER (avant l'intégration finale à l'import, prévue en phase 3) :
'   Dans la fenêtre Exécution immédiate (Ctrl+G), taper :
'        TraiterCasSuiviSante
'   puis Entrée. La feuille s'affiche avec le premier cas à traiter (s'il y en a).
'
' À PROPOS DES ACCENTS : voir la note dans mod_InstallSuiviSanteSheet. Même
' principe ici : fichier 100 % ASCII, textes accentués construits via la
' fonction mod_Display.FR() (recopiée ici à l'identique pour garder ce module autonome).
'
' IMPORTANT SUR L'ORDRE DES DECLARATIONS DANS CE FICHIER :
'   En VBA, toutes les declarations de niveau module (Type, variables Private/
'   Public en dehors d'une Sub/Function) DOIVENT se trouver AVANT la première
'   Sub ou Function du module. C'est pourquoi le type TCasSuiviSante et les
'   variables tabCas/nbCas/indexCasCourant sont regroupés tout en haut, avant
'   même la fonction FR (cette erreur a déjà été rencontrée dans le module
'   précédent : ne pas la reproduire ici).
' =====================================================================================


' --- Une "fiche" complète pour un cas à traiter : tout ce dont le formulaire
'     a besoin pour afficher une dépense et les remboursements associés. ---
Private Type TCasSuiviSante
    ligneDepense As Long        ' position de la ligne de dépense dans sa table source (DataBodyRange, 1 = première ligne)
    Source As String            ' PHASE 6 : "O" (TblOperations) ou "V" (TblVentilations)
    parentLigneOp As Long       ' PHASE 6 : si Source="V", ligne de l'opération parente dans TblOperations
                                 ' (via ID_Transaction), 0 si introuvable. Sert à corriger le Tiers, qui
                                 ' n'existe que sur l'opération bancaire, jamais sur une ligne de ventilation.
    dateDepense As Variant
    tiersDepense As String
    montantDepense As Double    ' toujours en valeur absolue (positive)
    Remb1Date As Variant
    Remb1Montant As Double
    Remb2Date As Variant
    Remb2Montant As Double
    NbRemb As Long               ' 0, 1 ou 2
    beneficiaireActuel As String
    franchiseActuelle As Double
    commentaireActuel As String
    ' --- Champs de contexte supplémentaires (lecture seule dans le formulaire),
    '     ajoutés pour aider l'opérateur à mieux identifier l'opération --------
    numChequeCtx As String
    dateConsultCtx As Variant
    speConsultCtx As String
    notesCtx As String
End Type

' État de la session en cours, conservé en mémoire pendant le passage en revue
' des cas (rempli par ChargerListeCasSuiviSante, consommé par
' AfficherCasCourant / CasSuivantSuiviSante / ValiderCasSuiviSante).
Private tabCas() As TCasSuiviSante
Private nbCas As Long
Private indexCasCourant As Long

' Même précaution que dans mod_SuiviSante : on compare toujours à "Frais, remb
' santé", construit via ChrW(233), jamais avec un caractère accentué saisi en dur.
Private Function CategorieSanteFormulaire() As String
    CategorieSanteFormulaire = "Frais, remb sant" & ChrW(233)
End Function


' =====================================================================================
' MACRO PRINCIPALE - point d'entree de la Phase 2b
' =====================================================================================
Public Sub TraiterCasSuiviSante()

    Dim ws As Worksheet

    ' On recalcule d'abord StatutSante/SoldeSante à partir des données actuelles
    ' (phase 1), pour ne jamais dépendre du fait que l'opérateur ait pensé à
    ' lancer CalculerSuiviSante avant d'ouvrir le formulaire. Résumé masqué
    ' (AfficherResume:=False) pour ne pas interrompre l'ouverture du formulaire
    ' avec un MsgBox intermédiaire.
    mod_SuiviSante.CalculerSuiviSante AfficherResume:=False

    Call ChargerListeCasSuiviSante

    If nbCas = 0 Then
        MsgBox mod_Display.FR("Aucune depense de sant{e2} n'a besoin d'{ea}tre trait{e2}e pour le moment."), _
               vbInformation, mod_Display.FR("Suivi sant{e2}")
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets(mod_InstallSuiviSanteSheet.NOM_FEUILLE_SUIVI_SANTE)

    ' Le solde est une formule Excel simple qui se recalcule dès que
    ' l'opérateur change la Franchise : pas besoin de VBA pour cela.
    ws.Range("ssSolde").Formula = "=ssMontant-ssRemb1Montant-ssRemb2Montant-ssFranchise"

    ' NB : on ne protège plus la feuille (la protection provoquait une erreur
    ' 1004 lors des changements dynamiques de Locked, pour un gain purement
    ' visuel). Les champs non pertinents restent modifiables à l'écran, mais
    ' ValiderCasSuiviSante contrôle de toute façon la saisie avant d'accepter
    ' un cas : la protection n'était pas le véritable garde-fou.

    ws.Visible = xlSheetVisible
    ws.Activate
    mod_InstallCommun.MasquerQuadrillage   ' quadrillage et en-tetes toujours masques (09/10/2026)

    ' Listes deroulantes reposees a chaque ouverture (09/10/2026).
    mod_InstallCommun.PoserListeDeroulante ws.Range("ssBeneficiaire"), "Beneficiaires"
    mod_InstallCommun.PoserListeDeroulante ws.Range("ssTiersCorrige"), "Praticiens"

    indexCasCourant = 1
    Call AfficherCasCourant(ws)

    ' Même verrou que pour la résolution des catégories : on bloque ici tant que
    ' l'opérateur n'a pas terminé (utile dès la phase 3, quand cette macro sera
    ' appelée depuis ImporterOperationsOFX et devra suspendre l'import).
    g_SaisieEnCours = True
    Do While g_SaisieEnCours
        DoEvents
    Loop

End Sub


' =====================================================================================
' ChargerListeCasSuiviSante : repère toutes les dépenses de santé qui doivent
' être proposées à l'opérateur, et remplit tabCas / nbCas.
' =====================================================================================
Private Sub ChargerListeCasSuiviSante()

    ' IMPORTANT : "tbl" et "tblData" ne sont PAS redeclares ici avec Dim.
    ' Ce sont les variables publiques déclarées dans mod_Synthese, déjà
    ' utilisées par mod_Display.RecupIndexCol pour calculer les colXxx.
    Dim nbLignesTable As Long
    Dim nbLigne As Long, j As Long

    Dim notesVues As Object   ' Scripting.Dictionary : évite de retraiter deux fois le même groupe
    Dim cleNotes As String

    ' --- PHASE 6 : jeu de tableaux memoire pour TblVentilations, facultatif ---
    Dim tblVen As ListObject
    Dim donneesVen As Variant
    Dim nbLignesVen As Long
    Dim venDisponible As Boolean
    Dim colVenSousCat As Long, colVenMontant As Long, colVenNotes As Long, colVenDate As Long
    Dim colVenStatut As Long, colVenHonoraire As Long, colVenBenef As Long, colVenFranchise As Long
    Dim colVenCommentaire As Long, colVenDateConsult As Long, colVenSpeConsult As Long

    nbCas = 0
    Erase tabCas

    Set tbl = mod_DonneesTable.GetOperationsTable()
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    mod_Display.RecupIndexCol
    If colSousCategorie = 0 Then Exit Sub   ' Phase 1 pas installee : rien a proposer

    tblData = tbl.DataBodyRange.value
    nbLignesTable = UBound(tblData, 1)

    ' --- PHASE 6 : chargement de TblVentilations, si disponible et complète ---
    venDisponible = False
    Set tblVen = ObtenirTableVentilationsSSF()
    If Not tblVen Is Nothing Then
        If Not tblVen.DataBodyRange Is Nothing Then
            colVenSousCat = IndexColVentilationSSF(tblVen, "SousCategorie")
            colVenMontant = IndexColVentilationSSF(tblVen, "Montant")
            colVenNotes = IndexColVentilationSSF(tblVen, "Notes")
            colVenStatut = IndexColVentilationSSF(tblVen, "StatutSante")
            colVenHonoraire = IndexColVentilationSSF(tblVen, "DepassementHoraires")
            colVenBenef = IndexColVentilationSSF(tblVen, "Beneficiaire")
            colVenFranchise = IndexColVentilationSSF(tblVen, "Franchise")
            colVenCommentaire = IndexColVentilationSSF(tblVen, "CommentaireSante")
            colVenDateConsult = IndexColVentilationSSF(tblVen, "Date_consult")
            colVenSpeConsult = IndexColVentilationSSF(tblVen, "Spe_Consult")
            If colVenSousCat <> 0 And colVenMontant <> 0 And colVenNotes <> 0 And colVenStatut <> 0 And _
               colVenHonoraire <> 0 And colVenBenef <> 0 And colVenFranchise <> 0 And colVenCommentaire <> 0 And _
               colVenDateConsult <> 0 And colVenSpeConsult <> 0 Then
                donneesVen = tblVen.DataBodyRange.value
                nbLignesVen = UBound(donneesVen, 1)
                venDisponible = True
            End If
        End If
    End If

    Set notesVues = CreateObject("Scripting.Dictionary")

    ' =====================================================================
    ' Boucle externe : un tour par ligne "Frais, remb santé" de TblOperations
    ' rencontrée, plus un tour par ligne de TblVentilations si elle n'a pas
    ' déjà été vue via une ligne de TblOperations du même groupe (Notes).
    ' =====================================================================
    For nbLigne = 1 To nbLignesTable
        If mod_DataStructure.CellText(tblData(nbLigne, colSousCategorie)) = CategorieSanteFormulaire() Then
            cleNotes = mod_DataStructure.CellText(tblData(nbLigne, colNotes))
            If Not notesVues.Exists(cleNotes) Then
                notesVues.Add cleNotes, True
                TraiterGroupeSiKO cleNotes, tblData, nbLignesTable, _
                    venDisponible, donneesVen, nbLignesVen, tblVen, _
                    colVenSousCat, colVenMontant, colVenNotes, colVenStatut, colVenHonoraire, _
                    colVenBenef, colVenFranchise, colVenCommentaire, colVenDateConsult, colVenSpeConsult
            End If
        End If
    Next nbLigne

    If venDisponible Then
        For nbLigne = 1 To nbLignesVen
            If mod_DataStructure.CellText(donneesVen(nbLigne, colVenSousCat)) = CategorieSanteFormulaire() Then
                cleNotes = mod_DataStructure.CellText(donneesVen(nbLigne, colVenNotes))
                If Not notesVues.Exists(cleNotes) Then
                    notesVues.Add cleNotes, True
                    TraiterGroupeSiKO cleNotes, tblData, nbLignesTable, _
                        venDisponible, donneesVen, nbLignesVen, tblVen, _
                        colVenSousCat, colVenMontant, colVenNotes, colVenStatut, colVenHonoraire, _
                        colVenBenef, colVenFranchise, colVenCommentaire, colVenDateConsult, colVenSpeConsult
                End If
            End If
        Next nbLigne
    End If

End Sub


' =====================================================================================
' TraiterGroupeSiKO : reconstitue le groupe complet (depense + jusqu'a 2
' remboursements) qui partage la valeur "cleNotes", en parcourant TblOperations
' PUIS TblVentilations, et ajoute une fiche a tabCas() si le groupe est bien
' KO/Honoraire=Faux. Extrait de ChargerListeCasSuiviSante (Phase 6) pour
' pouvoir être appelé depuis les deux boucles externes (TblOperations et
' TblVentilations) sans dupliquer cette logique.
' =====================================================================================
Private Sub TraiterGroupeSiKO(ByVal cleNotes As String, ByRef tblDataLocal As Variant, ByVal nbLignesTable As Long, _
        ByVal venDisponible As Boolean, ByRef donneesVen As Variant, ByVal nbLignesVen As Long, ByRef tblVen As ListObject, _
        ByVal colVenSousCat As Long, ByVal colVenMontant As Long, ByVal colVenNotes As Long, ByVal colVenStatut As Long, _
        ByVal colVenHonoraire As Long, ByVal colVenBenef As Long, ByVal colVenFranchise As Long, _
        ByVal colVenCommentaire As Long, ByVal colVenDateConsult As Long, ByVal colVenSpeConsult As Long)

    Dim j As Long
    Dim ligneDep As Long, sourceDep As String
    Dim r1d As Variant, r1m As Double, r2d As Variant, r2m As Double, nbR As Long

    ligneDep = 0
    sourceDep = "O"
    nbR = 0
    r1m = 0
    r2m = 0

    ' --- Membres du groupe cote TblOperations ---
    For j = 1 To nbLignesTable
        If mod_DataStructure.CellText(tblDataLocal(j, colSousCategorie)) = CategorieSanteFormulaire() Then
            If mod_DataStructure.CellText(tblDataLocal(j, colNotes)) = cleNotes Then
                If IsNumeric(tblDataLocal(j, colMontant)) Then
                    If mod_DataStructure.ToDouble(tblDataLocal(j, colMontant)) < 0 Then
                        If ligneDep = 0 Then
                            ligneDep = j
                            sourceDep = "O"
                        End If
                    Else
                        nbR = nbR + 1
                        If nbR = 1 Then
                            r1d = tblDataLocal(j, colDate)
                            r1m = mod_DataStructure.ToDouble(tblDataLocal(j, colMontant))
                        ElseIf nbR = 2 Then
                            r2d = tblDataLocal(j, colDate)
                            r2m = mod_DataStructure.ToDouble(tblDataLocal(j, colMontant))
                        End If
                    End If
                End If
            End If
        End If
    Next j

    ' --- Membres du groupe cote TblVentilations (PHASE 6) ---
    If venDisponible Then
        For j = 1 To nbLignesVen
            If mod_DataStructure.CellText(donneesVen(j, colVenSousCat)) = CategorieSanteFormulaire() Then
                If mod_DataStructure.CellText(donneesVen(j, colVenNotes)) = cleNotes Then
                    If IsNumeric(donneesVen(j, colVenMontant)) Then
                        If mod_DataStructure.ToDouble(donneesVen(j, colVenMontant)) < 0 Then
                            If ligneDep = 0 Then
                                ligneDep = j
                                sourceDep = "V"
                            End If
                        Else
                            ' Une ligne de ventilation positive (remboursement ventile)
                            ' n'a pas de date propre : on ne l'utilise pas pour
                            ' l'affichage Remb1/Remb2 (cas non prévu par la règle
                            ' métier actuelle : un remboursement de mutuelle n'est
                            ' jamais ventilé), mais on ne plante pas pour autant.
                        End If
                    End If
                End If
            End If
        Next j
    End If

    If ligneDep = 0 Then Exit Sub

    Dim statutGroupe As String
    Dim honoraireGroupe As Boolean
    Dim beneficiaireActuel As String, franchiseActuelle As Double, commentaireActuel As String
    Dim tiersDepense As String, dateDepense As Variant, numChequeCtx As String
    Dim dateConsultCtx As Variant, speConsultCtx As String, notesCtx As String
    Dim montantDepense As Double
    Dim parentLigneOp As Long

    If sourceDep = "O" Then
        statutGroupe = mod_DataStructure.CellText(tblDataLocal(ligneDep, colStatutSante))
        honoraireGroupe = (tblDataLocal(ligneDep, colDepassementHoraires) = True)
    Else
        statutGroupe = mod_DataStructure.CellText(donneesVen(ligneDep, colVenStatut))
        honoraireGroupe = (donneesVen(ligneDep, colVenHonoraire) = True)
    End If

    If statutGroupe <> "KO" Or honoraireGroupe Then Exit Sub

    parentLigneOp = 0
    If sourceDep = "O" Then
        dateDepense = tblDataLocal(ligneDep, colDate)
        tiersDepense = mod_DataStructure.CellText(tblDataLocal(ligneDep, colTiers))
        montantDepense = Abs(mod_DataStructure.ToDouble(tblDataLocal(ligneDep, colMontant)))
        beneficiaireActuel = mod_DataStructure.CellText(tblDataLocal(ligneDep, colBeneficiaire))
        franchiseActuelle = 0
        If IsNumeric(tblDataLocal(ligneDep, colFranchise)) Then
            franchiseActuelle = mod_DataStructure.ToDouble(tblDataLocal(ligneDep, colFranchise))
        End If
        commentaireActuel = mod_DataStructure.CellText(tblDataLocal(ligneDep, colCommentaireSante))
        numChequeCtx = mod_DataStructure.CellText(tblDataLocal(ligneDep, colCheque))
        dateConsultCtx = tblDataLocal(ligneDep, colDateConsult)
        speConsultCtx = mod_DataStructure.CellText(tblDataLocal(ligneDep, colSpeConsult))
        notesCtx = mod_DataStructure.CellText(tblDataLocal(ligneDep, colNotes))
    Else
        montantDepense = Abs(mod_DataStructure.ToDouble(donneesVen(ligneDep, colVenMontant)))
        beneficiaireActuel = mod_DataStructure.CellText(donneesVen(ligneDep, colVenBenef))
        franchiseActuelle = 0
        If IsNumeric(donneesVen(ligneDep, colVenFranchise)) Then
            franchiseActuelle = mod_DataStructure.ToDouble(donneesVen(ligneDep, colVenFranchise))
        End If
        commentaireActuel = mod_DataStructure.CellText(donneesVen(ligneDep, colVenCommentaire))
        dateConsultCtx = donneesVen(ligneDep, colVenDateConsult)
        speConsultCtx = mod_DataStructure.CellText(donneesVen(ligneDep, colVenSpeConsult))
        notesCtx = mod_DataStructure.CellText(donneesVen(ligneDep, colVenNotes))
        numChequeCtx = ""
        ' Date et Tiers n'existent pas dans TblVentilations : on va les
        ' chercher sur l'opération parente (TblOperations), via ID_Transaction.
        If Not TrouverParentVentilationSSF(tblVen, donneesVen, ligneDep, parentLigneOp, dateDepense, tiersDepense) Then
            dateDepense = ""
            tiersDepense = "(operation parente introuvable)"
        End If
    End If

    nbCas = nbCas + 1
    ReDim Preserve tabCas(1 To nbCas)
    With tabCas(nbCas)
        .ligneDepense = ligneDep
        .Source = sourceDep
        .parentLigneOp = parentLigneOp
        .dateDepense = dateDepense
        .tiersDepense = tiersDepense
        .montantDepense = montantDepense
        .Remb1Date = r1d
        .Remb1Montant = r1m
        .Remb2Date = r2d
        .Remb2Montant = r2m
        .NbRemb = nbR
        .beneficiaireActuel = beneficiaireActuel
        .franchiseActuelle = franchiseActuelle
        .commentaireActuel = commentaireActuel
        .numChequeCtx = numChequeCtx
        .dateConsultCtx = dateConsultCtx
        .speConsultCtx = speConsultCtx
        .notesCtx = notesCtx
    End With

End Sub


' =====================================================================================
' HELPERS PHASE 6 : accès à TblVentilations, dupliqués ici en local (comme dans
' mod_SuiviSante et mod_FormulairesNotes) pour que ce module puisse compiler
' même si mod_InstallVentilation n'a pas encore été importé.
' =====================================================================================
Private Function ObtenirTableVentilationsSSF() As ListObject
    Dim ws As Worksheet
    Dim t As ListObject
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Ventilations")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    Set t = ws.ListObjects("TblVentilations")
    On Error GoTo 0
    Set ObtenirTableVentilationsSSF = t
End Function

Private Function IndexColVentilationSSF(ByVal t As ListObject, ByVal nomColonne As String) As Long
    On Error Resume Next
    IndexColVentilationSSF = t.ListColumns(nomColonne).index
    On Error GoTo 0
End Function

' Retrouve, pour une ligne de TblVentilations, la ligne PARENTE dans
' TblOperations (via ID_Transaction) : renvoie son numéro de ligne (pour
' pouvoir y corriger le Tiers plus tard), ainsi que sa date et son Tiers pour
' l'affichage. Renvoie False si la ligne est introuvable.
Private Function TrouverParentVentilationSSF(ByRef tblVen As ListObject, ByRef donneesVen As Variant, _
        ByVal ligneVen As Long, ByRef parentLigneOp As Long, ByRef dateAff As Variant, ByRef tiersAff As String) As Boolean

    Dim colVenID As Long
    Dim idCible As String
    Dim tblOp As ListObject
    Dim donneesOp As Variant
    Dim n As Long, i As Long

    TrouverParentVentilationSSF = False
    parentLigneOp = 0

    colVenID = IndexColVentilationSSF(tblVen, "ID_Transaction")
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
            parentLigneOp = i
            dateAff = donneesOp(i, colDate)
            tiersAff = mod_DataStructure.CellText(donneesOp(i, colTiers))
            TrouverParentVentilationSSF = True
            Exit Function
        End If
    Next i
End Function


' =====================================================================================
' AfficherCasCourant : remplit le formulaire avec le cas numéro indexCasCourant
' =====================================================================================
Private Sub AfficherCasCourant(ws As Worksheet)

    Dim cas As TCasSuiviSante
    cas = tabCas(indexCasCourant)

    ws.Range("ssDate").value = cas.dateDepense
    ws.Range("ssTiersImporte").value = cas.tiersDepense
    ws.Range("ssMontant").value = cas.montantDepense

    ws.Range("ssRemb1Date").value = cas.Remb1Date
    ws.Range("ssRemb1Montant").value = cas.Remb1Montant

    If cas.NbRemb >= 2 Then
        ws.Range("ssRemb2Date").value = cas.Remb2Date
        ws.Range("ssRemb2Montant").value = cas.Remb2Montant
    Else
        ws.Range("ssRemb2Date").value = ""
        ws.Range("ssRemb2Montant").value = 0   ' 0 et non vide : nécessaire pour que la formule du solde reste juste
    End If

    ws.Range("ssBeneficiaire").value = cas.beneficiaireActuel
    If (Application.WorksheetFunction.CountIf(Range("Praticiens"), cas.tiersDepense) > 0) Then
        ws.Range("ssTiersCorrige").value = cas.tiersDepense
    Else
        ws.Range("ssTiersCorrige").value = ""
    End If
    ws.Range("ssFranchise").value = cas.franchiseActuelle
    ws.Range("ssDepassement").value = ""
    ws.Range("ssCommentaire").value = cas.commentaireActuel

    ws.Range("ssLigneEnCours").value = cas.ligneDepense   ' aide au diagnostic si besoin, non utilisée par le code

    ws.Range("ssNumCheque").value = cas.numChequeCtx
    ws.Range("ssDateConsult").value = cas.dateConsultCtx
    ws.Range("ssSpeConsult").value = cas.speConsultCtx
    ws.Range("ssNotes").value = cas.notesCtx

    Call AppliquerEtatChampsConditionnels(ws, cas)
    Call MettreAJourCompteur(ws)

End Sub


' =====================================================================================
' AppliquerEtatChampsConditionnels : grise ou déverrouille les champs "Tiers
' corrigé" et "Dépassement d'honoraires" selon leur pertinence pour le cas affiché.
' =====================================================================================
Private Sub AppliquerEtatChampsConditionnels(ws As Worksheet, cas As TCasSuiviSante)

    ' On ne joue plus que sur la couleur de fond (gris = non pertinent, jaune
    ' = saisie attendue). La feuille n'est pas protégée et ces champs restent
    ' modifiables même lorsqu'ils sont grisés, mais ValiderCasSuiviSante ignore
    ' leur contenu lorsqu'ils ne sont pas pertinents pour le cas en cours.

    With ws.Range("ssTiersCorrige")
        If EstTiersValide(cas.tiersDepense) Then
            .Interior.Color = mod_InstallCommun.CoulFondLecture()
        Else
            .Interior.Color = mod_InstallCommun.CoulFondSaisie()
        End If
    End With
End Sub


' =====================================================================================
' EstTiersValide : le Tiers de la dépense est-il déjà une valeur de la plage
' nommée "Praticiens" ?
' =====================================================================================
' On compte le nombre d'occurrences d'un Tiers dans la plage de données nommée "Praticiens".
' 1. Application.WorksheetFunction permet d'appeler les fonctions natives d'Excel depuis VBA.
' 2. .CountIf(...) est le nom anglophone de la fonction Excel NB.SI. Elle prend deux arguments : CountIf(Plage, Critère).

Private Function EstTiersValide(ByVal tiers As String) As Boolean
    On Error Resume Next
    EstTiersValide = (Application.WorksheetFunction.CountIf(Range("Praticiens"), tiers) > 0)
    On Error GoTo 0
End Function


' =====================================================================================
' BOUTONS "+" : ajouter une nouvelle valeur a la liste Beneficiaires ou
' Praticiens, directement depuis le formulaire, sans repasser par Excel/Param.
' =====================================================================================
Public Sub AjouterBeneficiaire()
    AjouterValeurDansListe "Beneficiaires", "ssBeneficiaire"
End Sub

Public Sub AjouterTiersPraticien()
    AjouterValeurDansListe "Praticiens", "ssTiersCorrige"
End Sub

' AjouterValeurDansListe : coeur commun aux 2 boutons "+".
'   nomPlage        -> nom de la plage nommee a completer ("Beneficiaires" ou "Praticiens")
'   nomCelluleCible -> nom de la cellule du formulaire dans laquelle on
'                      selectionne directement la valeur ajoutee ("ssBeneficiaire" ou "ssTiersCorrige")
'
' Hypothese : la plage nommee est une simple liste verticale sur une seule
' colonne (comme Tri_champs), pas un tableau dynamique en "eclaboussure"
' (ANCHORARRAY) comme Mois/Annees. Si ce n'est pas le cas, dis-le : le
' fonctionnement devra être adapté.
Private Sub AjouterValeurDansListe(ByVal nomPlage As String, ByVal nomCelluleCible As String)

    Dim nouvelleValeur As String
    Dim rngListe As Range
    Dim celluleVide As Range
    Dim c As Range
    Dim ws As Worksheet

    nouvelleValeur = Trim(InputBox(mod_Display.FR("Nouvelle valeur {a2} ajouter {a2} la liste '") & nomPlage & "' :", _
                                    mod_Display.FR("Ajouter une valeur")))
    If nouvelleValeur = "" Then Exit Sub   ' annule (Echap, ou rien saisi) : on ne fait rien

    On Error Resume Next
    Set rngListe = ThisWorkbook.Names(nomPlage).RefersToRange
    On Error GoTo 0

    If rngListe Is Nothing Then
        MsgBox mod_Display.FR("La plage nomm{e2}e '") & nomPlage & mod_Display.FR("' est introuvable. V{e2}rifie qu'elle existe bien dans Param."), _
               vbCritical
        Exit Sub
    End If

    If Application.WorksheetFunction.CountIf(rngListe, nouvelleValeur) > 0 Then
        MsgBox mod_Display.FR("Cette valeur existe d{e2}j{a2} dans la liste."), vbInformation
    Else
        ' On cherche d'abord une cellule VIDE dans la plage actuelle...
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
            ' ...sinon la plage est déjà pleine : on l'agrandit d'une ligne
            ' vers le bas, puis on redéfinit le nom pour qu'il couvre cette
            ' nouvelle ligne (les listes déroulantes qui utilisent ce nom,
            ' par exemple "=Beneficiaires", suivront automatiquement).
            Dim nouvelleCellule As Range
            Set nouvelleCellule = rngListe.Cells(rngListe.Cells.count + 1)
            nouvelleCellule.value = nouvelleValeur
            ThisWorkbook.Names(nomPlage).RefersTo = _
                "=" & rngListe.Resize(rngListe.rows.count + 1).Address(External:=True)
        End If
    End If

    ' On selectionne directement la valeur ajoutee dans le champ du
    ' formulaire : l'opérateur n'a pas besoin de rouvrir la liste déroulante.
    Set ws = ThisWorkbook.Worksheets(mod_InstallSuiviSanteSheet.NOM_FEUILLE_SUIVI_SANTE)
    ws.Range(nomCelluleCible).value = nouvelleValeur

End Sub


' =====================================================================================
' MettreAJourCompteur : actualise le texte "Operation: x/y"
' =====================================================================================
Private Sub MettreAJourCompteur(ws As Worksheet)
    ' Format commun a tous les formulaires : "Operation: x/y" (09/10/2026), x = cas en cours.
    ws.Range("CompteurCasSante").value = mod_InstallCommun.TexteCompteur(indexCasCourant, nbCas)
End Sub


' =====================================================================================
' BOUTON "Cas suivant" : passe au cas suivant SANS enregistrer de reponse
' =====================================================================================
Public Sub CasSuivantSuiviSante()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mod_InstallSuiviSanteSheet.NOM_FEUILLE_SUIVI_SANTE)

    If indexCasCourant >= nbCas Then
        Call TerminerSessionSuiviSante(ws, mod_Display.FR("Tous les cas ont {e2}t{e2} parcourus."))
        Exit Sub
    End If

    indexCasCourant = indexCasCourant + 1
    Call AfficherCasCourant(ws)

End Sub

Public Sub SortirSuiviSante()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mod_InstallSuiviSanteSheet.NOM_FEUILLE_SUIVI_SANTE)
    ' On se repositionne sur la feuille Synthèse
    Sheets("Synthese").Activate
    ' On masque la feuille en cours
    ws.Visible = xlSheetVeryHidden
    ' Arrêt complet du programme
    End
End Sub

' =====================================================================================
' BOUTON "Valider ce cas" : controle la saisie, l'ecrit dans TblOperations,
' relance le calcul automatique, puis passe au cas suivant.
' =====================================================================================
Public Sub ValiderCasSuiviSante()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mod_InstallSuiviSanteSheet.NOM_FEUILLE_SUIVI_SANTE)

    Dim cas As TCasSuiviSante
    cas = tabCas(indexCasCourant)

    Dim beneficiaire As String
    Dim tiersCorrige As String
    Dim commentaire As String
    Dim franchiseTexte As String
    Dim franchise As Double
    Dim depassementTexte As String
    Dim tiersValideActuellement As Boolean
    Dim depassementRequis As Boolean

    ' --- Controles avant d'accepter la saisie (mieux vaut prevenir tout de
    '     suite que d'ecrire une information incomplete dans TblOperations) ---

    beneficiaire = Trim(CStr(ws.Range("ssBeneficiaire").value))
    If beneficiaire = "" Then
        MsgBox mod_Display.FR("Merci de choisir un B{e2}n{e2}ficiaire avant de valider ce cas."), vbExclamation
        Exit Sub
    End If

    tiersValideActuellement = EstTiersValide(cas.tiersDepense)
    tiersCorrige = Trim(CStr(ws.Range("ssTiersCorrige").value))
    If Not tiersValideActuellement And tiersCorrige = "" Then
        MsgBox mod_Display.FR("Le Tiers import{e2} n'est pas reconnu : merci de choisir une valeur dans 'Tiers corrig{e2}'."), vbExclamation
        Exit Sub
    End If

    franchiseTexte = Trim(CStr(ws.Range("ssFranchise").value))
    If franchiseTexte = "" Then
        franchise = 0
    ElseIf IsNumeric(franchiseTexte) Then
        franchise = mod_DataStructure.ToDouble(ws.Range("ssFranchise").value)
    Else
        MsgBox mod_Display.FR("La Franchise doit {ea}tre un nombre (0 si aucune)."), vbExclamation
        Exit Sub
    End If

'    depassementRequis = (cas.NbRemb >= 2)

'    If depassementRequis And depassementTexte = "" Then
'        MsgBox mod_Display.FR("Merci d'indiquer s'il s'agit d'un d{e2}passement d'honoraires."), vbExclamation
'        Exit Sub
'    End If

    commentaire = CStr(ws.Range("ssCommentaire").value)

    ' --- Écriture limitée à la ligne de cette dépense et aux seules colonnes
    '     concernées (même principe qu'en phase 1 : aucune réécriture globale
    '     du tableau ni de sa mise en forme). En phase 6, la table cible dépend
    '     de cas.Source : TblOperations (comportement d'origine) ou
    '     TblVentilations. Le Tiers n'existe que sur l'opération bancaire :
    '     une correction pour une ligne de ventilation s'écrit donc toujours
    '     dans la ligne parente de TblOperations (cas.ParentLigneOp). ---
    Dim tblVen As ListObject

    depassementTexte = Trim(CStr(ws.Range("ssDepassement").value))

    If cas.Source = "O" Then
        Set tbl = mod_DonneesTable.GetOperationsTable()

        tbl.ListColumns("Beneficiaire").DataBodyRange.rows(cas.ligneDepense).value = beneficiaire
        tbl.ListColumns("Franchise").DataBodyRange.rows(cas.ligneDepense).value = franchise
        tbl.ListColumns("CommentaireSante").DataBodyRange.rows(cas.ligneDepense).value = commentaire

        If Not tiersValideActuellement Then
            tbl.ListColumns("Tiers").DataBodyRange.rows(cas.ligneDepense).value = tiersCorrige
        End If

        If depassementTexte = "Oui" Then
            tbl.ListColumns("DepassementHoraires").DataBodyRange.rows(cas.ligneDepense).value = True
        Else
            tbl.ListColumns("DepassementHoraires").DataBodyRange.rows(cas.ligneDepense).value = False
        End If
    Else
        Set tblVen = ObtenirTableVentilationsSSF()
        If tblVen Is Nothing Then Exit Sub

        tblVen.ListColumns("Beneficiaire").DataBodyRange.rows(cas.ligneDepense).value = beneficiaire
        tblVen.ListColumns("Franchise").DataBodyRange.rows(cas.ligneDepense).value = franchise
        tblVen.ListColumns("CommentaireSante").DataBodyRange.rows(cas.ligneDepense).value = commentaire

        If Not tiersValideActuellement And cas.parentLigneOp <> 0 Then
            Set tbl = mod_DonneesTable.GetOperationsTable()
            tbl.ListColumns("Tiers").DataBodyRange.rows(cas.parentLigneOp).value = tiersCorrige
        End If

        If depassementTexte = "Oui" Then
            tblVen.ListColumns("DepassementHoraires").DataBodyRange.rows(cas.ligneDepense).value = True
        Else
            tblVen.ListColumns("DepassementHoraires").DataBodyRange.rows(cas.ligneDepense).value = False
        End If
    End If

    ' --- On relance le calcul automatique de la phase 1 : StatutSante et
    '     SoldeSante seront immédiatement recalculés avec la Franchise et
    '     l'Honoraire qui viennent d'être renseignés. On réutilise la logique
    '     de mod_SuiviSante plutôt que de la recopier ici. Le résumé habituel
    '     (MsgBox) est désactivé pour ne pas interrompre la revue des cas. ---
    mod_SuiviSante.CalculerSuiviSante AfficherResume:=False

    ' --- Cas suivant ---
    If indexCasCourant >= nbCas Then
        Call TerminerSessionSuiviSante(ws, mod_Display.FR("Tous les cas ont {e2}t{e2} trait{e2}s."))
    Else
        indexCasCourant = indexCasCourant + 1
        Call AfficherCasCourant(ws)
    End If

End Sub


' =====================================================================================
' TerminerSessionSuiviSante : ferme proprement la session (feuille masquée et
' verrou libéré pour que TraiterCasSuiviSante puisse rendre la main).
' =====================================================================================
Private Sub TerminerSessionSuiviSante(ws As Worksheet, ByVal Message As String)
    ws.Visible = xlSheetVeryHidden
    g_SaisieEnCours = False
    MsgBox Message, vbInformation, mod_Display.FR("Suivi sant{e2}")
End Sub


' =====================================================================================
' OUTIL DE DIAGNOSTIC TEMPORAIRE (à supprimer une fois le suivi santé stabilisé)
' =====================================================================================
' Affiche dans la fenêtre Exécution immédiate (Ctrl+G) les index de colonnes
' calculés par RecupIndexCol, puis le détail de chaque ligne "Frais, remb
' santé" (Notes, Montant, StatutSante, Honoraire). Si le problème réapparaît
' après la correction ci-dessus, ce rapport aidera à déterminer immédiatement
' si un index de colonne est incorrect ou si les données elles-mêmes sont en
' cause, sans avoir à deviner.
'
' À UTILISER : Ctrl+G, taper DiagnostiquerSuiviSante, Entrée. Le rapport
' s'affiche dans la même fenêtre (Exécution immédiate).
Public Sub DiagnostiquerSuiviSante()

    Dim tblData As Variant
    Dim nbLignesTable As Long
    Dim nbLigne As Long
    Dim n As Long

    Set tbl = mod_DonneesTable.GetOperationsTable()
    If tbl Is Nothing Then
        Debug.Print "TblOperations introuvable."
        Exit Sub
    End If

    mod_Display.RecupIndexCol

    Debug.Print "--- Index de colonnes ---"
    Debug.Print "Categorie=" & colCategorie & "  SousCategorie=" & colSousCategorie & "  Notes=" & colNotes & _
                "  Montant=" & colMontant & "  Date=" & colDate & "  Tiers=" & colTiers
    Debug.Print "StatutSante=" & colStatutSante & "  Franchise=" & colFranchise & _
                "  SoldeSante=" & colSoldeSante & "  DepassementHoraires=" & colDepassementHoraires & _
                "  Beneficiaire=" & colBeneficiaire & "  CommentaireSante=" & colCommentaireSante
    Debug.Print "(Un 0 signifie que la colonne n'a pas ete trouvee dans TblOperations)"
    Debug.Print "(PHASE 6 : le test utilise desormais SousCategorie, pas Categorie)"
    Debug.Print ""

    If tbl.DataBodyRange Is Nothing Then
        Debug.Print "TblOperations ne contient aucune ligne."
        Exit Sub
    End If

    If colSousCategorie = 0 Then
        Debug.Print "SousCategorie introuvable : Phase 1 (categories a 2 niveaux) pas encore installee."
        Exit Sub
    End If

    tblData = tbl.DataBodyRange.value
    nbLignesTable = UBound(tblData, 1)

    Debug.Print "--- Lignes 'Frais, remb sant" & ChrW(233) & "' (TblOperations) ---"
    n = 0
    For nbLigne = 1 To nbLignesTable
        If mod_DataStructure.CellText(tblData(nbLigne, colSousCategorie)) = "Frais, remb sant" & ChrW(233) Then
            n = n + 1
            Debug.Print n & ") Notes=[" & mod_DataStructure.CellText(tblData(nbLigne, colNotes)) & "]" & _
                        "  Montant=" & tblData(nbLigne, colMontant) & _
                        "  StatutSante=[" & mod_DataStructure.CellText(tblData(nbLigne, colStatutSante)) & "]" & _
                        "  Honoraire=" & tblData(nbLigne, colDepassementHoraires)
        End If
    Next nbLigne
    Debug.Print "--- Fin (" & n & " ligne(s) trouvee(s)) ---"

    ' --- PHASE 6 : même rapport, côté TblVentilations, si disponible ---
    Dim tblVen As ListObject
    Dim donneesVen As Variant
    Dim colVenSousCat As Long, colVenNotes As Long, colVenMontant As Long, colVenStatut As Long, colVenHonoraire As Long
    Dim nVen As Long, i As Long

    Set tblVen = ObtenirTableVentilationsSSF()
    If tblVen Is Nothing Then
        Debug.Print ""
        Debug.Print "(TblVentilations : feuille/tableau non installe -- Phase 4 pas encore faite)"
    ElseIf tblVen.DataBodyRange Is Nothing Then
        Debug.Print ""
        Debug.Print "(TblVentilations : installe, mais aucune ligne pour l'instant)"
    Else
        colVenSousCat = IndexColVentilationSSF(tblVen, "SousCategorie")
        colVenNotes = IndexColVentilationSSF(tblVen, "Notes")
        colVenMontant = IndexColVentilationSSF(tblVen, "Montant")
        colVenStatut = IndexColVentilationSSF(tblVen, "StatutSante")
        colVenHonoraire = IndexColVentilationSSF(tblVen, "DepassementHoraires")
        Debug.Print ""
        If colVenSousCat = 0 Then
            Debug.Print "(TblVentilations : colonne SousCategorie introuvable)"
        Else
            donneesVen = tblVen.DataBodyRange.value
            nVen = UBound(donneesVen, 1)
            Debug.Print "--- Lignes 'Frais, remb sant" & ChrW(233) & "' (TblVentilations) ---"
            n = 0
            For i = 1 To nVen
                If mod_DataStructure.CellText(donneesVen(i, colVenSousCat)) = "Frais, remb sant" & ChrW(233) Then
                    n = n + 1
                    Debug.Print n & ") Notes=[" & mod_DataStructure.CellText(donneesVen(i, colVenNotes)) & "]" & _
                                "  Montant=" & donneesVen(i, colVenMontant) & _
                                "  StatutSante=[" & mod_DataStructure.CellText(donneesVen(i, colVenStatut)) & "]" & _
                                "  Honoraire=" & donneesVen(i, colVenHonoraire)
                End If
            Next i
            Debug.Print "--- Fin (" & n & " ligne(s) trouvee(s)) ---"
        End If
    End If

    MsgBox "Rapport ecrit dans la fenetre Execution immediate (Ctrl+G)." & vbCrLf & _
           "Fais defiler vers le haut si besoin pour tout voir.", vbInformation

End Sub
