Option Explicit

' =====================================================================================
' MODULE : mod_FormulairesNotes
'
' ROLE (PHASE 4b du chantier "Suivi SantÃ©") :
'   Logique complÃ¨te des 2 feuilles construites en Phase 4a
'   (mod_InstallFormulairesNotes) : frm_RapprochementNotes et frm_GenerationCle.
'
'   Point d'entree unique : VerifierNotesSante(). Appelee automatiquement par
'   mod_ImportOFX juste avant CalculerSuiviSante, mais peut aussi Ãªtre
'   relancee manuellement (Ctrl+G) : VerifierNotesSante
'
'   DEROULE :
'     1. On parcourt TOUTE TblOperations (lignes "Frais, remb santÃ©").
'     2. Pour chaque ligne dont Date_consult vaut encore la "sentinelle"
'        (02/01/1900, voir mod_ImportOFX), on ouvre frm_RapprochementNotes :
'        l'opÃ©rateur filtre en cascade (Date -> Specialite -> BÃ©nÃ©ficiaire ->
'        Montant) parmi les clÃ©s Notes dÃ©jÃ  valides ailleurs, et valide.
'     3. Si aucune correspondance : bouton "Pas de correspondance" ->
'        frm_GenerationCle, ou l'opÃ©rateur crÃ©Ã© une nouvelle clÃ©.
'     4. Dans les 2 cas, la clÃ© est ecrite dans Notes pour cette ligne, et
'        Date_consult/Spe_Consult/BÃ©nÃ©ficiaire sont recalcules immediatement
'        (mÃªmes formules que mod_ImportOFX, Ã©tape 7).
'
'   MÃªme verrou modal que les autres formulaires du classeur : la variable
'   Public g_SaisieEnCours (dÃ©jÃ  declaree dans mod_ResolutionCategories) est
'   REUTILISEE telle quelle.
'
'   PARTICULARITE : comme ce classeur evite les UserForm (bug DPI dÃ©jÃ 
'   rencontre) et que le code-behind de feuille (Worksheet_Change) demandÃ© de
'   connaitre le "nom de code" VBA de la feuille (non garanti a l'avance ici),
'   la cascade de filtres est geree par SONDAGE : la boucle d'attente modale
'   (Do While g_SaisieEnCours : DoEvents : Loop) vÃ©rifie a chaque passage si
'   Date/Specialite/BÃ©nÃ©ficiaire ont change depuis le dernier passage, et
'   met a jour les listes suivantes le cas echeant. C'est fiable et ne
'   necessite aucun code-behind.
'
' A PROPOS DES ACCENTS : fichier 100% ASCII, textes accentues construits via
' la fonction FR() (recopiee ici a l'identique, comme dans les autres
' modules du chantier).
'
' COMMENT TESTER :
'   Ctrl+G, taper VerifierNotesSante, Entree.
' =====================================================================================


' --- Une clÃ© "Notes" valide, decodee, utilisÃ©e pour construire les listes de
'     filtres en cascade de frm_RapprochementNotes. ---
Private Type TCandidatCle
    DateTexte As String     ' au format jj/mm/aaaa, pour affichage/filtre
    specialite As String    ' segment 1 de la clÃ©
    beneficiaire As String  ' segment 2 de la clÃ©
    Montant As String       ' segment 3 de la clÃ©, tel quel (texte)
    CleBrute As String      ' la valeur complÃ¨te du champ Notes
End Type

' Memorise quel bouton a ete clique ("Valider" ou "PasDeCorrespondance" pour
' frm_RapprochementNotes ; "Valider" pour frm_GenerationCle), pour que le code
' qui a lance la boucle d'attente sache quoi faire une fois qu'elle se termine.
Private derniereAction As String


'Private Function FR(ByVal texte As String) As String
'    Dim r As String
'    r = texte
'    r = Replace(r, "{e2}", ChrW(233))
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
' MACRO PRINCIPALE - point d'entree de la Phase 4b
' =====================================================================================
' =====================================================================================
' RetraiterSuiviSante : enchaine les 3 Ã©tapes du suivi santÃ© SANS reimporter
' de fichier (utile aprÃ¨s une correction manuelle dans Import_data, par
' exemple si tu changes une CatÃ©gorie a la main parce que la source a change
' et que le rapprochement par FITID ne l'aurait pas detecte).
' Ctrl+G : RetraiterSuiviSante
' =====================================================================================
Public Sub RetraiterSuiviSante()
    VerifierNotesSante
    ' mod_SuiviSante.CalculerSuiviSante est lancé par la macro suivante
    mod_SuiviSanteFormulaire.TraiterCasSuiviSante
End Sub


Public Sub VerifierNotesSante()

    ' PHASE 6 : on traitÃ© maintenant DEUX sources de lignes "santÃ©" en
    ' attente de clÃ© : les lignes de TblOperations (SousCategorie) ET les
    ' lignes de TblVentilations (SousCategorie Ã©galement, colonne ajoutee en
    ' Phase 4). Chaque liste est traitÃ©e separement (numerotation continue
    ' dans le compteur affiche a l'opÃ©rateur), mais le mecanisme est
    ' rigoureusement identique dans les 2 cas.
    Dim donneesInitiales As Variant
    Dim nbLignesTable As Long
    Dim nbLigne As Long
    Dim totalATraiter As Long
    Dim numeroEnCours As Long

    Dim pendingOp() As Long
    Dim nbPendingOp As Long

    Dim tblVen As ListObject
    Dim donneesVen As Variant
    Dim nbLignesVen As Long
    Dim pendingVen() As Long
    Dim nbPendingVen As Long
    Dim colVenSousCat As Long, colVenDateConsult As Long
    Dim venDisponible As Boolean

    Set tbl = mod_DonneesTable.GetOperationsTable()
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    mod_Display.RecupIndexCol

    If colSousCategorie = 0 Then
        ' Phase 1 (catÃ©gories a 2 niveaux) pas encore installee : on ne peut
        ' pas savoir quelles lignes sont "Frais, remb santÃ©". On sort sans
        ' bloquer (VerifierNotesSante est appelee automatiquement a chaque
        ' import : elle ne doit jamais faire planter l'import).
        Exit Sub
    End If

    donneesInitiales = tbl.DataBodyRange.value
    nbLignesTable = UBound(donneesInitiales, 1)

    ' --- ETAPE 1a : lignes de TblOperations en attente de clÃ© ---------------
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

    ' --- ETAPE 1b : lignes de TblVentilations en attente de clÃ© (PHASE 6) ---
    ' TblVentilations est FACULTATIVE : si la Phase 4 n'a pas ete installee,
    ' ou si ses colonnes SousCategorie/Date_consult n'existent pas encore, on
    ' ignore simplement cette partie (aucune erreur).
    nbPendingVen = 0
    venDisponible = False
    Set tblVen = ObtenirTableVentilationsFN()
    If Not tblVen Is Nothing Then
        If Not tblVen.DataBodyRange Is Nothing Then
            colVenSousCat = IndexColSiExisteFN(tblVen, "SousCategorie")
            colVenDateConsult = IndexColSiExisteFN(tblVen, "Date_consult")
            If colVenSousCat <> 0 And colVenDateConsult <> 0 Then
                venDisponible = True
                donneesVen = tblVen.DataBodyRange.value
                nbLignesVen = UBound(donneesVen, 1)
                For nbLigne = 1 To nbLignesVen
                    If mod_DataStructure.CellText(donneesVen(nbLigne, colVenSousCat)) = "Frais, remb sant" & ChrW(233) Then
                        If EstLigneSentinelle(donneesVen(nbLigne, colVenDateConsult)) Then
                            nbPendingVen = nbPendingVen + 1
                            ReDim Preserve pendingVen(1 To nbPendingVen)
                            pendingVen(nbPendingVen) = nbLigne
                        End If
                    End If
                Next nbLigne
            End If
        End If
    End If

    totalATraiter = nbPendingOp + nbPendingVen
    If totalATraiter = 0 Then Exit Sub

    ' --- ETAPE 2 : on propose chaque ligne, TblOperations d'abord puis
    '     TblVentilations (ordre choisi arbitrairement, sans consÃ©quence :
    '     seul le nombre total affiche a l'opÃ©rateur compte) ------------------
    numeroEnCours = 0
    For nbLigne = 1 To nbPendingOp
        numeroEnCours = numeroEnCours + 1
        AfficherRapprochementPourLigne pendingOp(nbLigne), "O", numeroEnCours, totalATraiter
    Next nbLigne

    If venDisponible Then
        For nbLigne = 1 To nbPendingVen
            numeroEnCours = numeroEnCours + 1
            AfficherRapprochementPourLigne pendingVen(nbLigne), "V", numeroEnCours, totalATraiter
        Next nbLigne
    End If

End Sub

' Vrai si la valeur de Date_consult est encore la "sentinelle" (02/01/1900)
' posee par mod_ImportOFX quand le decoupage de Notes Ã©choue, ou si le champ
' est carrement vide/non renseigne.
Private Function EstLigneSentinelle(ByVal dateConsultVal As Variant) As Boolean
    EstLigneSentinelle = True
    If IsDate(dateConsultVal) Then
        If CDate(dateConsultVal) <> DateSerial(1900, 1, 2) Then EstLigneSentinelle = False
    End If
End Function

' =====================================================================================
' HELPERS PHASE 6 : accÃ¨s a TblVentilations, dupliques ICI en local (comme dans
' mod_SuiviSante) pour que ce module continue a compiler mÃªme si
' mod_InstallVentilation n'a pas encore ete importe.
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

' Retrouve, pour une ligne de TblVentilations, la ligne PARENTE correspondante
' dans TblOperations (via ID_Transaction), pour rÃ©cupÃ©rer a l'affichage des
' informations que TblVentilations ne stocke pas elle-mÃªme (Date, Tiers,
' Num_Cheque : une ligne de ventilation ne reprÃ©sente qu'une PARTIE d'une
' opÃ©ration bancaire, elle n'a pas sa propre date ni son propre tiers).
' Renvoie False si la ligne parente n'a pas ete retrouvee (ne devrait
' normalement jamais arriver, mais on reste defensif).
Private Function TrouverContexteParentVentilation(ByVal ligneVen As Long, ByRef tblVen As ListObject, _
                                                   ByRef donneesVen As Variant, _
                                                   ByRef dateAff As Variant, ByRef tiersAff As String, _
                                                   ByRef chequeAff As Variant) As Boolean
    Dim colVenID As Long
    Dim idCible As String
    Dim tblOp As ListObject
    Dim donneesOp As Variant
    Dim n As Long, i As Long

    TrouverContexteParentVentilation = False
    dateAff = ""
    tiersAff = ""
    chequeAff = ""

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
            TrouverContexteParentVentilation = True
            Exit Function
        End If
    Next i
End Function




' =====================================================================================
' AfficherRapprochementPourLigne : ouvre frm_RapprochementNotes pour UNE ligne
' (position dans DataBodyRange), attend la decision de l'opÃ©rateur, puis
' applique le rÃ©sultat.
' =====================================================================================
Private Sub AfficherRapprochementPourLigne(ByVal ligneDepense As Long, ByVal Source As String, ByVal numero As Long, ByVal total As Long)

    Dim ws As Worksheet
    Dim candidats() As TCandidatCle
    Dim nbCandidats As Long
    Dim dernierDate As String, dernierSpecialite As String, dernierBeneficiaire As String
    Dim montantChoisi As String

    ' --- PHASE 6 : donnÃ©es de contexte de l'opÃ©ration, lues dans la bonne
    '     source (O = TblOperations, V = TblVentilations). Pour une ligne de
    '     ventilation, Date/Tiers/NumCheque n'existent pas dans TblVentilations
    '     elle-mÃªme : on va les chercher sur la ligne PARENTE de
    '     TblOperations via ID_Transaction (voir TrouverContexteParentVentilation). ---
    Dim dateCtx As Variant, tiersCtx As String, chequeCtx As Variant, notesCtx As String, montantCtx As Double
    Dim tblVen As ListObject
    Dim donneesVen As Variant
    Dim colVenNotes As Long, colVenMontant As Long

    If Source = "O" Then
        dateCtx = tblData(ligneDepense, colDate)
        tiersCtx = mod_DataStructure.CellText(tblData(ligneDepense, colTiers))
        chequeCtx = tblData(ligneDepense, colCheque)
        notesCtx = mod_DataStructure.CellText(tblData(ligneDepense, colNotes))
        montantCtx = mod_DataStructure.ToDouble(tblData(ligneDepense, colMontant))
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

    Set ws = ThisWorkbook.Worksheets(mod_InstallFormulairesNotes.NOM_FEUILLE_RAPPROCHEMENT)

    ' On affiche la feuille EN PREMIER, avant toute ecriture de cellule ou
    ' redefinition de plage nommee : redefinir Name.RefersTo pendant que la
    ' feuille est encore xlSheetVeryHidden provoque une erreur 1004.
    ws.Visible = xlSheetVisible
    ws.Activate

    ' Reconstruit la liste des clÃ©s valides A CHAQUE APPEL (pas une seule
    ' fois pour tout VerifierNotesSante) : ainsi, une clÃ© gÃ©nÃ©rÃ©e via
    ' frm_GenerationCle pour une ligne devient immediatement disponible comme
    ' candidat pour les lignes suivantes de la mÃªme session -- y compris
    ' d'une source a l'autre (une clÃ© validee sur une ligne de TblOperations
    ' devient un candidat pour une ligne de TblVentilations, et vice versa).
    ConstruireListeCandidats candidats, nbCandidats

    ' --- Contexte de l'opÃ©ration (lecture seule), affiche que la ligne vienne
    '     de TblOperations ou de TblVentilations (PHASE 6 : tiersCtx precise
    '     "(ventilation)" pour que l'opÃ©rateur sache tout de suite d'ou vient
    '     la ligne qu'il traitÃ©) ---
    ws.Range("rnDateOp").value = dateCtx
    ws.Range("rnNumCheque").value = chequeCtx
    If Source = "V" Then
        ws.Range("rnTiersOp").value = tiersCtx & FR(" [ligne de ventilation]")
    Else
        ws.Range("rnTiersOp").value = tiersCtx
    End If
    ws.Range("rnNotes").value = notesCtx
    ws.Range("rnMontantOp").value = montantCtx
    ws.Range("rnLigneEnCours").value = ligneDepense
    ws.Range("rnCompteurCas").value = numero & FR(" sur ") & total

    ' --- Remise a zero des filtres ---
    ' Format Texte force ici aussi (pas seulement sur la colonne technique) :
    ' sans ca, la valeur selectionnee dans la liste dÃ©roulante pourrait Ãªtre
    ' reconvertie en date au moment ou elle atterrit dans la cellule.
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
            AfficherGenerationPourLigne ligneDepense, Source
        Case "Passer"
            ' Volontairement rien a faire : la ligne reste inchangee (KO,
            ' Date_consult toujours sentinelle), elle sera repropose au
            ' prochain passage.
        Case "Sortir"
            End
    End Select

End Sub


' =====================================================================================
' BOUTONS de frm_RapprochementNotes
' =====================================================================================
' Ces actions modifie la valeur de variable
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
' Date_consult reste a la valeur sentinelle : elle sera donc automatiquement
' reproposee au prochain VerifierNotesSante (ou au prochain import).
Public Sub PasserRapprochementNotes()
    derniereAction = "Passer"
    g_SaisieEnCours = False
End Sub

' "Sortir" : aucune modification. La ligne reste "KO" (StatutSante) et
' Date_consult reste a la valeur sentinelle : elle sera donc automatiquement
' reproposee au prochain VerifierNotesSante (ou au prochain import).
Public Sub SortirRapprochementNotes()
    derniereAction = "Sortir"
    g_SaisieEnCours = False
End Sub

Public Sub PasserGenerationCle()
    derniereAction = "Passer"
    g_SaisieEnCours = False
End Sub


' =====================================================================================
' AfficherGenerationPourLigne : ouvre frm_GenerationCle pour UNE ligne
' =====================================================================================
Private Sub AfficherGenerationPourLigne(ByVal ligneDepense As Long, ByVal Source As String)

    Dim ws As Worksheet

    ' --- PHASE 6 : mÃªmes lectures gÃ©nÃ©ralisÃ©es O/V que dans
    '     AfficherRapprochementPourLigne (voir ses commentaires) ---
    Dim dateCtx As Variant, tiersCtx As String, chequeCtx As Variant, notesCtx As String, montantCtx As Double
    Dim tblVen As ListObject
    Dim donneesVen As Variant
    Dim colVenNotes As Long, colVenMontant As Long

    If Source = "O" Then
        dateCtx = tblData(ligneDepense, colDate)
        tiersCtx = mod_DataStructure.CellText(tblData(ligneDepense, colTiers))
        chequeCtx = tblData(ligneDepense, colCheque)
        notesCtx = mod_DataStructure.CellText(tblData(ligneDepense, colNotes))
        montantCtx = mod_DataStructure.ToDouble(tblData(ligneDepense, colMontant))
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

    derniereAction = ""
    g_SaisieEnCours = True
    Do While g_SaisieEnCours
        DoEvents
    Loop

    ws.Visible = xlSheetVeryHidden

    ' Si "Passer" a ete clique, derniereAction = "Passer" et on ne rentre
    ' dans aucun des blocs ci-dessous : rien n'est ecrit, la ligne reste
    ' inchangee et sera reproposee au prochain passage.
    If derniereAction = "Valider" Then
        If Trim(CStr(ws.Range("gcCleGeneree").value)) <> "" Then
            AppliquerNouvelleCle ligneDepense, Source, CStr(ws.Range("gcCleGeneree").value)
        End If
    End If

End Sub


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
    ws.Range("gcCleGeneree").Copy   ' copie dans le presse-papier (comme un Ctrl+C manuel sur la cellule)

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
' BOUTONS "+" de frm_GenerationCle (Specialite / BÃ©nÃ©ficiaire)
' MÃªme principe que AjouterBeneficiaire/AjouterTiersPraticien dans
' mod_SuiviSanteFormulaire, duplique ici (avec le nom de feuille en
' parametre) pour rester autonome de ce module a l'autre.
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
' AppliquerNouvelleCle : ecrit la clÃ© dans Notes pour une ligne, puis
' recalcule Date_consult / Spe_Consult / BÃ©nÃ©ficiaire pour CETTE ligne
' uniquement (mÃªme logique que mod_ImportOFX, Ã©tape 7).
' =====================================================================================
Private Sub AppliquerNouvelleCle(ByVal ligneDepense As Long, ByVal Source As String, ByVal cle As String)

    Dim seg0 As String, seg1 As String, seg2 As String
    Dim tblVen As ListObject

    seg0 = mod_ImportOFX.SegmentTexte(cle, ";", 0)
    seg1 = mod_ImportOFX.SegmentTexte(cle, ";", 1)
    seg2 = mod_ImportOFX.SegmentTexte(cle, ";", 2)

    If Source = "O" Then
        ' --- Ecriture dans TblOperations (comportement d'origine, inchange) ---
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
        ' --- PHASE 6 : ecriture dans TblVentilations. MÃªmes colonnes, mÃªmes
        '     rÃ¨gles, mais sur l'AUTRE tableau : une ligne de ventilation
        '     "Frais, remb santÃ©" possÃ¨de exactement les mÃªmes colonnes de
        '     suivi que TblOperations (voir mod_InstallVentilation.
        '     PreparerTableVentilations), c'est ce qui rend cette duplication
        '     de logique possible sans rien inventer de nouveau. ---
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
' ConstruireListeCandidats : construit la liste de toutes les clÃ©s Notes
' DEJA VALIDES (une seule fois par clÃ© distincte) parmi les lignes santÃ© de
' TblOperations.
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

    ' PHASE 6 : la colonne testee est desormais SousCategorie (et non plus
    ' CatÃ©gorie, qui contient maintenant la catÃ©gorie PARENTE).
    If colSousCategorie <> 0 Then
        For i = 1 To n
            If mod_DataStructure.CellText(tblData(i, colSousCategorie)) = "Frais, remb sant" & ChrW(233) Then
                notesTexte = mod_DataStructure.CellText(tblData(i, colNotes))
                AjouterCandidatSiValide notesTexte, vues, candidats, nbCandidats
            End If
        Next i
    End If

    ' --- PHASE 6 : on ajoute aussi les clÃ©s DEJA VALIDES presentes dans
    '     TblVentilations, pour qu'une clÃ© gÃ©nÃ©rÃ©e sur une ligne de
    '     ventilation devienne selectionnable pour une ligne de
    '     TblOperations (et inversement) -----------------------------------
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

' Petit utilitaire commun aux 2 boucles ci-dessus (TblOperations et
' TblVentilations) : decoupe une valeur de Notes, et l'ajoute a candidats()
' si elle a un segment 0 (date) valide et n'a pas dÃ©jÃ  ete vue.
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
' FONCTIONS DE FILTRAGE (construisent les listes pour chaque niveau de la cascade)
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

' Tri a bulles simple (liste toujours courte : le nombre de dates de
' consultation distinctes, pas le nombre de lignes de TblOperations) : la
' date la plus recente en premier, pour limiter le defilement dans la liste.
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

' Ecrit en une fois les 2 listes paralleles Montant/ClÃ© (mÃªme ligne = mÃªme candidat)
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

' Percours la totalité du tableau "candidats()" à la recherche de la clé saisie
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

' Petit utilitaire : la valeur est-elle dÃ©jÃ  dans le tableau rÃ©sultat(1 To nb) ?
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

' Tableau de chaines vide (0 Ã©lÃ©ment), pour reinitialiser une liste
Private Function VideListe() As String()
    Dim vide() As String
    VideListe = vide
End Function


' =====================================================================================
' EcrireListeEtRedefinirNom : ecrit un tableau de valeurs a partir de la ligne
' 2 d'une colonne technique, puis redefinit le nom pour qu'il pointe
' exactement sur les cellules utilisÃ©es (mÃªme principe que la croissance des
' plages Beneficiaires/Praticiens dans mod_SuiviSanteFormulaire).
' =====================================================================================
Private Sub EcrireListeEtRedefinirNom(ws As Worksheet, ByVal nomListe As String, ByVal colonne As String, ByRef valeurs() As String)

    Dim i As Long
    Dim nb As Long

    On Error Resume Next
    nb = UBound(valeurs) - LBound(valeurs) + 1
    On Error GoTo 0

    ' IMPORTANT : on force la colonne en format TEXTE avant d'ecrire quoi que
    ' ce soit. Sans ca, une valeur qui "ressemble" a une date (ex: "01/12/1900")
    ' est automatiquement convertie par Excel en vraie date des qu'on
    ' l'assigne a une cellule, exactement comme si on la tapait a la main.
    ' Cette conversion casse ensuite silencieusement les comparaisons de texte
    ' strict utilisÃ©es partout dans ce module (ListeSpecialitesPour, etc.),
    ' puisque CStr() d'une date convertie ne redonne pas forcement exactement
    ' le mÃªme texte que celui d'origine.
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

' Applique une liste dÃ©roulante de validation pointant sur un nom defini
Private Sub AppliquerListeFN(ByVal rng As Range, ByVal nomPlage As String)
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & nomPlage
    On Error GoTo 0
End Sub


