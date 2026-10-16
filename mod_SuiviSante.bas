Option Explicit

' ============================================================================
'  MODULE : mod_SuiviSante
'  RÔLE   : Suivi persistant des opérations de la sous-catégorie "Frais, remb
'           santé" dans TblOperations, et également dans TblVentilations depuis
'           la phase 6 (lignes de ventilation affectées à cette même sous-catégorie).
'
'  HISTORIQUE DES PHASES :
'   - Phase 1 (version d'origine) : calcul automatique de StatutSante et
'     SoldeSante, en isolant les lignes de TblOperations dont la COLONNE
'     "Catégorie" valait "Frais, remb santé" (à l'époque, il n'existait pas
'     encore de sous-catégorie : Catégorie était la granularité la plus fine).
'   - Phase 6 (ce fichier) : deux changements structurels demandés par
'     l'opérateur lors de la mise en place des catégories à deux niveaux
'     (Catégorie / SousCategorie) et de la ventilation :
'       (a) le test bascule sur la colonne "SousCategorie" (la colonne
'           "Catégorie" contient désormais la catégorie PARENTE, par exemple
'           "Santé, prévoyance", commune à plusieurs sous-catégories) ;
'       (b) les lignes de TblVentilations dont la SousCategorie vaut aussi
'           "Frais, remb santé" entrent dans la MÊME MÉCANIQUE qu'une ligne
'           normale de TblOperations : même regroupement par clé "Notes",
'           même calcul de solde, même écriture de StatutSante/SoldeSante,
'           mais chacune dans SA PROPRE table (TblOperations ou
'           TblVentilations), puisque ce sont deux tableaux Excel distincts.
'           conséquence directe de la décision de l'opérateur :
'           "les opérations ventilées entrent dans la mécanique principale
'           ... à terme elles doivent être considérées comme une opération
'           comme une autre."
'
'  PRINCIPE DE REGROUPEMENT (inchangé depuis la phase 1) : un "groupe" est
'  l'ensemble des lignes qui partagent exactement la même valeur dans la
'  colonne "Notes" -- qu'elles viennent de TblOperations ou de
'  TblVentilations. Un groupe contient en général une ligne de dépense
'  (montant négatif) et une ou plusieurs lignes de remboursement (montant
'  positif). Depuis la phase 6, la ligne de dépense peut venir soit de
'  TblOperations (opération classique), soit de TblVentilations (partie
'  d'une opération ventilée). Le remboursement, lui, reste presque toujours
'  une ligne bancaire classique de TblOperations (un remboursement de
'  mutuelle n'est jamais ventile).
'
'  RÈGLE DE NON-RETRAITEMENT : un groupe n'est PAS recalculé si sa ligne de
'  dépense a StatutSante = "OK" (son "OK" est alors propagé à toutes ses lignes).
'  Tous les autres groupes sont recalculés, y compris "KO" avec
'  DepassementHoraires = VRAI (correctif du 16/10/2026) : ce drapeau sert
'  seulement à ne plus soumettre le cas à l'opérateur (mod_SuiviSanteFormulaire),
'  il ne doit pas empêcher un groupe soldé ensuite de passer à "OK".
'
'  IMPORTANT : ce module ne réécrit JAMAIS un tableau en entier et ne
'  réécrit même pas la colonne Franchise (saisie réservée à l'opérateur). Il
'  ne touche qu'aux colonnes qu'il a le droit de modifier (StatutSante,
'  SoldeSante, DepassementHoraires), chacune via tbl.ListColumns("NomColonne"),
'  et cela SÉPARÉMENT pour TblOperations et TblVentilations.
'
'  TblVentilations EST FACULTATIVE : si la feuille "Ventilations" ou le
'  tableau "TblVentilations" n'existent pas encore (phase 4 non installée
'  chez l'opérateur), ou si ses colonnes de suivi santé n'ont pas encore été
'  ajoutées, ce module continue de fonctionner normalement sur TblOperations
'  seule. Aucune erreur n'est levée : la partie Ventilations est simplement
'  ignorée.
' ============================================================================

' --- Petit type privé utilisé pour charger en mémoire les opérations de santé
'     des deux tables avant de les regrouper. Le champ "Source" indique dans
'     quel tableau la ligne a été trouvée ("O" = TblOperations, "V" =
'     TblVentilations) et détermine quel tableau mémoire (OP ou VEN) utiliser
'     au moment de réécrire les résultats. ---

' --- Noms de la feuille et du tableau de ventilation dupliqués ici en dur
'     (plutôt que de référencer mod_InstallVentilation.VEN_NOM_xxx) pour que
'     ce module puisse compiler même si mod_InstallVentilation n'a pas encore
'     été importé (phase 4 facultative pour ce module). -----------------------
Private Const VEN_FEUILLE As String = "Ventilations"
Private Const VEN_TABLE As String = "TblVentilations"

Private Type TOperationSante
    Source As String        ' "O" (TblOperations) ou "V" (TblVentilations)
    LigneOrigine As Long    ' numéro de ligne dans le tableau mémoire de sa source, pas dans Excel
    Montant As Double       ' montant tel quel (négatif pour une dépense, positif pour un remboursement)
    Notes As String         ' valeur du champ Notes : c'est la clé de regroupement
    Traite As Boolean       ' cette ligne a-t-elle déjà été rattachée à un groupe traité ?
End Type

' Nom exact de la sous-catégorie à surveiller. ChrW(233) construit le « é »
' de « santé » au lieu de saisir directement le caractère accentué, ce qui
' garantit une comparaison correcte quel que soit l'encodage du fichier .bas.
Private Function SousCategorieSante() As String
    SousCategorieSante = "Frais, remb sant" & ChrW(233)
End Function


' ----------------------------------------------------------------------------
' ValeursColonne : lit une colonne du tableau et renvoie TOUJOURS un tableau
' 2D (1 To n, 1 To 1), même si la table ne contient qu'une seule ligne.
' ----------------------------------------------------------------------------
' Piege classique VBA : Range.Value renvoie un tableau 2D quand la plage
' contient plusieurs cellules, mais renvoie une simple valeur (pas un
' tableau) quand la plage ne contient qu'UNE seule cellule.
Private Function ValeursColonne(ByVal plageColonne As Range) As Variant
    Dim v As Variant
    v = plageColonne.value
    If Not IsArray(v) Then
        Dim tmp(1 To 1, 1 To 1) As Variant
        tmp(1, 1) = v
        ValeursColonne = tmp
    Else
        ValeursColonne = v
    End If
End Function

' ----------------------------------------------------------------------------
' ObtenirTableVentilationsSiExiste : renvoie le ListObject TblVentilations,
' ou Nothing si la feuille ou le tableau n'existe pas encore chez l'opérateur
' (phase 4 non installée). Ne lève jamais d'erreur.
' ----------------------------------------------------------------------------
Private Function ObtenirTableVentilationsSiExiste() As ListObject
    Dim ws As Worksheet
    Dim t As ListObject
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(VEN_FEUILLE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    Set t = ws.ListObjects(VEN_TABLE)
    On Error GoTo 0
    Set ObtenirTableVentilationsSiExiste = t
End Function

' Recherche l'index d'une colonne par son nom, renvoie 0 si absente (jamais d'erreur).
Private Function IndexColSiExiste(ByVal t As ListObject, ByVal nomColonne As String) As Long
    On Error Resume Next
    IndexColSiExiste = t.ListColumns(nomColonne).index
    On Error GoTo 0
End Function


' ----------------------------------------------------------------------------
' MACRO PRINCIPALE
' ----------------------------------------------------------------------------
' Appelee automatiquement depuis ImporterOperationsOFX, ou manuellement
' (Ctrl+G, CalculerSuiviSante) pour retraiter toute la table, par exemple
' après une saisie manuelle de Franchise.
Public Sub CalculerSuiviSante(Optional ByVal AfficherResume As Boolean = True)

    Dim tblData As Variant
    Dim nbLignesTable As Long
    Dim nbLigne As Long

    Dim arrStatut As Variant, arrSolde As Variant, arrHonoraire As Variant, arrFranchise As Variant
    Dim origStatut As Variant, origSolde As Variant, origHonoraire As Variant
    Dim origStatutVen As Variant, origSoldeVen As Variant, origHonoraireVen As Variant

    ' --- Jeux de tableaux memoire pour TblVentilations (Phase 6) : restent
    '     vides (Empty) si TblVentilations n'existe pas ou n'a pas encore ses
    '     colonnes de suivi santé --------------------------------------------
    Dim tblVen As ListObject
    Dim tblDataVen As Variant
    Dim nbLignesVen As Long
    Dim arrStatutVen As Variant, arrSoldeVen As Variant, arrHonoraireVen As Variant, arrFranchiseVen As Variant
    Dim venDisponible As Boolean
    Dim colVenSousCat As Long, colVenMontant As Long, colVenNotes As Long
    Dim colVenStatut As Long, colVenSolde As Long, colVenHonoraire As Long, colVenFranchise As Long

    Dim tabOps() As TOperationSante
    Dim nbOps As Long

    Dim i As Long
    Dim nbGroupesTraites As Long

    ' --- ÉTAPE 1 : récupération de TblOperations et des index de colonnes ---
    Set tbl = mod_DonneesTable.GetOperationsTable()
    If tbl Is Nothing Then
        MsgBox "Le tableau TblOperations est introuvable.", vbExclamation
        Exit Sub
    End If
    If tbl.DataBodyRange Is Nothing Then
        MsgBox "Le tableau TblOperations ne contient aucune ligne.", vbExclamation
        Exit Sub
    End If

    ' RecupIndexCol (mod_Display) remplit les variables publiques colXxx en
    ' recherchant chaque colonne par son nom d'en-tête, jamais par un numéro fixe.
    mod_Display.RecupIndexCol

    ' PHASE 6 : la colonne testée est désormais SousCategorie (et non plus
    ' Catégorie). Si elle n'existe pas encore (phase 1 non installée), on
    ' prévient clairement au lieu de comparer une colonne qui n'a pas le sens
    ' attendu.
    If colSousCategorie = 0 Then
        MsgBox "La colonne 'SousCategorie' est introuvable dans TblOperations." & vbCrLf & _
               "Installez d'abord la gestion des categories a 2 niveaux (Phase 1, " & _
               "mod_Categories.PreparerPhase1Categories) avant de lancer ce calcul.", _
               vbCritical, "Colonne manquante"
        Exit Sub
    End If

    If colStatutSante = 0 Or colFranchise = 0 Or colSoldeSante = 0 Or colDepassementHoraires = 0 Or colCommentaireSante = 0 Then
        MsgBox "Les colonnes de suivi sante (StatutSante, Franchise, SoldeSante, " & _
               "DepassementHoraires, CommentaireSante) sont introuvables " & _
               "dans TblOperations." & vbCrLf & _
               "Verifie qu'elles ont bien ete ajoutees a la table avant de relancer ce calcul.", _
               vbCritical, "Colonnes manquantes"
        Exit Sub
    End If

    ' --- ÉTAPE 2 : lecture en mémoire de TblOperations ---
    tblData = tbl.DataBodyRange.value
    nbLignesTable = UBound(tblData, 1)

    arrStatut = ValeursColonne(tbl.ListColumns("StatutSante").DataBodyRange)
    arrSolde = ValeursColonne(tbl.ListColumns("SoldeSante").DataBodyRange)
    arrHonoraire = ValeursColonne(tbl.ListColumns("DepassementHoraires").DataBodyRange)
    arrFranchise = ValeursColonne(tbl.ListColumns("Franchise").DataBodyRange)   ' jamais écrite, lue seulement
    origStatut = arrStatut          ' copies de départ : seules les cellules modifiées seront écrites
    origSolde = arrSolde
    origHonoraire = arrHonoraire

    ' --- ÉTAPE 2bis (PHASE 6) : lecture en mémoire de TblVentilations, si disponible ---
    venDisponible = False
    Set tblVen = ObtenirTableVentilationsSiExiste()
    If Not tblVen Is Nothing Then
        If Not tblVen.DataBodyRange Is Nothing Then
            colVenSousCat = IndexColSiExiste(tblVen, "SousCategorie")
            colVenMontant = IndexColSiExiste(tblVen, "Montant")
            colVenNotes = IndexColSiExiste(tblVen, "Notes")
            colVenStatut = IndexColSiExiste(tblVen, "StatutSante")
            colVenSolde = IndexColSiExiste(tblVen, "SoldeSante")
            colVenHonoraire = IndexColSiExiste(tblVen, "DepassementHoraires")
            colVenFranchise = IndexColSiExiste(tblVen, "Franchise")

            If colVenSousCat <> 0 And colVenMontant <> 0 And colVenNotes <> 0 And _
               colVenStatut <> 0 And colVenSolde <> 0 And colVenHonoraire <> 0 And colVenFranchise <> 0 Then
                tblDataVen = tblVen.DataBodyRange.value
                nbLignesVen = UBound(tblDataVen, 1)
                arrStatutVen = ValeursColonne(tblVen.ListColumns("StatutSante").DataBodyRange)
                arrSoldeVen = ValeursColonne(tblVen.ListColumns("SoldeSante").DataBodyRange)
                arrHonoraireVen = ValeursColonne(tblVen.ListColumns("DepassementHoraires").DataBodyRange)
                arrFranchiseVen = ValeursColonne(tblVen.ListColumns("Franchise").DataBodyRange)
                origStatutVen = arrStatutVen
                origSoldeVen = arrSoldeVen
                origHonoraireVen = arrHonoraireVen
                venDisponible = True
            End If
            ' Si une colonne de suivi manque encore dans TblVentilations, on ne
            ' bloque pas : on continue sans cette table (venDisponible reste False).
        End If
    End If

    ' --- ÉTAPE 3 : on isole toutes les lignes "Frais, remb santé" des deux
    '     tables dans un seul tableau mémoire combiné -------------------------
    nbOps = 0

    For nbLigne = 1 To nbLignesTable
        If mod_DataStructure.CellText(tblData(nbLigne, colSousCategorie)) = SousCategorieSante() Then
            If IsNumeric(tblData(nbLigne, colMontant)) Then
                nbOps = nbOps + 1
                ReDim Preserve tabOps(1 To nbOps)
                tabOps(nbOps).Source = "O"
                tabOps(nbOps).LigneOrigine = nbLigne
                tabOps(nbOps).Montant = mod_DataStructure.ToDouble(tblData(nbLigne, colMontant))
                tabOps(nbOps).Notes = mod_DataStructure.CellText(tblData(nbLigne, colNotes))
                tabOps(nbOps).Traite = False
            End If
        End If
    Next nbLigne

    If venDisponible Then
        For nbLigne = 1 To nbLignesVen
            If mod_DataStructure.CellText(tblDataVen(nbLigne, colVenSousCat)) = SousCategorieSante() Then
                If IsNumeric(tblDataVen(nbLigne, colVenMontant)) Then
                    nbOps = nbOps + 1
                    ReDim Preserve tabOps(1 To nbOps)
                    tabOps(nbOps).Source = "V"
                    tabOps(nbOps).LigneOrigine = nbLigne
                    tabOps(nbOps).Montant = mod_DataStructure.ToDouble(tblDataVen(nbLigne, colVenMontant))
                    tabOps(nbOps).Notes = mod_DataStructure.CellText(tblDataVen(nbLigne, colVenNotes))
                    tabOps(nbOps).Traite = False
                End If
            End If
        Next nbLigne
    End If

    If nbOps = 0 Then
        MsgBox "Aucune operation '" & SousCategorieSante() & "' trouvee (ni dans TblOperations, " & _
               "ni dans TblVentilations).", vbInformation
        Exit Sub
    End If

    ' --- ÉTAPE 4 : on traite chaque groupe (chaque valeur de Notes) une seule
    '     fois. Le regroupement se fait sur l'ensemble combiné : deux lignes
    '     venant de tables différentes mais partageant la même clé Notes sont
    '     bien traitées comme un seul groupe. ----------------------------------
    nbGroupesTraites = 0
    For i = 1 To nbOps
        If Not tabOps(i).Traite Then
            TraiterGroupeSante tabOps, nbOps, i, _
                arrStatut, arrSolde, arrHonoraire, arrFranchise, _
                arrStatutVen, arrSoldeVen, arrHonoraireVen, arrFranchiseVen, _
                nbGroupesTraites
        End If
    Next i

    ' --- ÉTAPE 5 : réécriture uniquement des colonnes concernées, table par
    '     table. On n'écrit jamais un tableau en entier : seules les colonnes
    '     de suivi sont réécrites, chacune indépendamment. ----------------------
    ' CORRECTIF du 16/10/2026 : l'ancienne écriture "colonne entière" (.Value = tableau de
    ' 4 691 lignes) n'était pas conservée par Excel chez l'opérateur (363 statuts calculés,
    ' 3 présents ensuite). On écrit maintenant CELLULE PAR CELLULE, et SEULEMENT les valeurs
    ' qui changent, événements désactivés ; la première cellule refusée est signalée.
    Dim evAvant As Boolean
    Dim erreurEcriture As String
    evAvant = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    EcrireColonneParDifference tbl.ListColumns("StatutSante").DataBodyRange, arrStatut, origStatut, erreurEcriture
    EcrireColonneParDifference tbl.ListColumns("SoldeSante").DataBodyRange, arrSolde, origSolde, erreurEcriture
    EcrireColonneParDifference tbl.ListColumns("DepassementHoraires").DataBodyRange, arrHonoraire, origHonoraire, erreurEcriture

    If venDisponible Then
        EcrireColonneParDifference tblVen.ListColumns("StatutSante").DataBodyRange, arrStatutVen, origStatutVen, erreurEcriture
        EcrireColonneParDifference tblVen.ListColumns("SoldeSante").DataBodyRange, arrSoldeVen, origSoldeVen, erreurEcriture
        EcrireColonneParDifference tblVen.ListColumns("DepassementHoraires").DataBodyRange, arrHonoraireVen, origHonoraireVen, erreurEcriture
    End If
    ' Format fixe de SoldeSante (correctif du 16/10/2026) : une valeur de l'ordre de 1E-15 écrite
    ' dans une cellule "Standard" la fait passer en format scientifique, et elle le garde ensuite.
    tbl.ListColumns("SoldeSante").DataBodyRange.NumberFormat = "#,##0.00"
    If venDisponible Then tblVen.ListColumns("SoldeSante").DataBodyRange.NumberFormat = "#,##0.00"
    Application.ScreenUpdating = True
    Application.EnableEvents = evAvant

    ' --- Contrôle de l'écriture (ajout du 16/10/2026) : on relit la colonne StatutSante dans
    '     la feuille et on la compare à ce qui devait être écrit. Un écart signale une écriture
    '     qui n'a pas "pris" (événement, protection, tableau différent...) au lieu de le laisser
    '     passer inaperçu. ---
    Dim nbAttendus As Long, nbRelus As Long, r As Long
    Dim relu As Variant
    For r = 1 To UBound(arrStatut, 1)
        If mod_DataStructure.CellText(arrStatut(r, 1)) <> "" Then nbAttendus = nbAttendus + 1
    Next r
    relu = ValeursColonne(tbl.ListColumns("StatutSante").DataBodyRange)
    For r = 1 To UBound(relu, 1)
        If mod_DataStructure.CellText(relu(r, 1)) <> "" Then nbRelus = nbRelus + 1
    Next r
    If nbRelus <> nbAttendus Then
        MsgBox "Anomalie : " & nbAttendus & " statut(s) devaient etre ecrits dans StatutSante de TblOperations, " & _
               "mais " & nbRelus & " seulement sont presents apres l'ecriture." & vbCrLf & _
               "Table : " & tbl.Parent.Name & "!" & tbl.Name & " (" & tbl.DataBodyRange.Rows.Count & " lignes)." & _
               IIf(erreurEcriture <> "", vbCrLf & "Premiere cellule refusee : " & erreurEcriture, ""), _
               vbExclamation, "Suivi sante"
    End If

    If AfficherResume Then
        MsgBox nbGroupesTraites & " groupe(s) de depenses sante recalcule(s) sur " & _
               nbOps & " ligne(s) '" & SousCategorieSante() & "' analysee(s) " & _
               "(TblOperations" & IIf(venDisponible, " + TblVentilations", "") & ")." & vbCrLf & _
               nbRelus & " statut(s) renseigne(s) dans StatutSante.", _
               vbInformation, "Suivi sante"
    End If

End Sub


' ----------------------------------------------------------------------------
' EcrireColonneParDifference : écrit dans la colonne "plage" uniquement les valeurs de
' "nouv" qui diffèrent de "orig" (même forme : tableaux 2D à une colonne). Si Excel refuse
' une cellule, ou ne conserve pas la valeur écrite, la PREMIÈRE anomalie est décrite dans
' premiereErreur (adresse et cause) ; le traitement continue.
' ----------------------------------------------------------------------------
Private Sub EcrireColonneParDifference(ByVal plage As Range, ByRef nouv As Variant, _
                                       ByRef orig As Variant, ByRef premiereErreur As String)
    Dim i As Long
    Dim attendu As String
    For i = 1 To UBound(nouv, 1)
        If CStr(nouv(i, 1)) <> CStr(orig(i, 1)) Then
            attendu = CStr(nouv(i, 1))
            On Error Resume Next
            Err.Clear
            plage.Cells(i, 1).Value = nouv(i, 1)
            If Err.Number <> 0 Then
                If premiereErreur = "" Then premiereErreur = plage.Parent.Name & "!" & plage.Cells(i, 1).Address(False, False) & " : " & Err.Description
            ElseIf CStr(plage.Cells(i, 1).Value) <> attendu Then
                If premiereErreur = "" Then premiereErreur = plage.Parent.Name & "!" & plage.Cells(i, 1).Address(False, False) & " : ecrit [" & attendu & "], relu [" & CStr(plage.Cells(i, 1).Value) & "]"
            End If
            On Error GoTo 0
        End If
    Next i
End Sub


' ----------------------------------------------------------------------------
' TraiterGroupeSante : traite un groupe (toutes les lignes qui partagent la
' même valeur de "Notes" que la ligne tabOps(indexDepart)), en lisant et écrivant
' dans le tableau mémoire correspondant à la source de chaque ligne.
' ----------------------------------------------------------------------------
' Paramètres :
'   tabOps                        -> toutes les opérations santé (2 tables), en mémoire
'   nbOps                         -> nombre de lignes dans tabOps
'   indexDepart                   -> position dans tabOps d'une ligne du groupe à traiter
'   arrStatut/arrSolde/arrHonoraire/arrFranchise (ByRef)      -> colonnes de TblOperations
'   arrStatutVen/arrSoldeVen/arrHonoraireVen/arrFranchiseVen (ByRef) -> colonnes de TblVentilations
'   nbGroupesTraites (ByRef)      -> compteur global, incrémenté si ce groupe a été recalculé
Private Sub TraiterGroupeSante(ByRef tabOps() As TOperationSante, ByVal nbOps As Long, _
                                ByVal indexDepart As Long, _
                                ByRef arrStatut As Variant, ByRef arrSolde As Variant, _
                                ByRef arrHonoraire As Variant, ByRef arrFranchise As Variant, _
                                ByRef arrStatutVen As Variant, ByRef arrSoldeVen As Variant, _
                                ByRef arrHonoraireVen As Variant, ByRef arrFranchiseVen As Variant, _
                                ByRef nbGroupesTraites As Long)

    Dim notesRef As String
    Dim j As Long, k As Long
    Dim indexDepense As Long        ' position dans tabOps (pas dans une table) de la première dépense du groupe, 0 si aucune
    Dim montantDepenses As Double
    Dim montantRemb As Double
    Dim nbLignesGroupe As Long
    Dim indicesGroupe() As Long      ' positions dans tabOps de toutes les lignes du groupe
    Dim statutActuel As String
    Dim honoraireActuel As Boolean
    Dim nouveauStatut As String
    Dim montantFranchise As Double
    Dim solde As Double

    notesRef = tabOps(indexDepart).Notes

    ' --- Sous-étape A : on repère toutes les lignes du groupe, quelle que soit
    '     leur source, et on les marque "Traité" tout de suite -----------------
    nbLignesGroupe = 0
    indexDepense = 0
    montantDepenses = 0
    montantRemb = 0

    For j = 1 To nbOps
        If (Not tabOps(j).Traite) And tabOps(j).Notes = notesRef Then
            tabOps(j).Traite = True

            nbLignesGroupe = nbLignesGroupe + 1
            ReDim Preserve indicesGroupe(1 To nbLignesGroupe)
            indicesGroupe(nbLignesGroupe) = j

            If tabOps(j).Montant < 0 Then
                If indexDepense = 0 Then indexDepense = j
                montantDepenses = montantDepenses + Abs(tabOps(j).Montant)
            Else
                montantRemb = montantRemb + tabOps(j).Montant
            End If
        End If
    Next j

    ' --- Sous-étape B : faut-il retraiter ce groupe ? Le verrou se lit dans le
    '     tableau mémoire correspondant à la source de la dépense. -------------
    If indexDepense <> 0 Then
        If tabOps(indexDepense).Source = "O" Then
            statutActuel = mod_DataStructure.CellText(arrStatut(tabOps(indexDepense).LigneOrigine, 1))
            honoraireActuel = (arrHonoraire(tabOps(indexDepense).LigneOrigine, 1) = True)
        Else
            statutActuel = mod_DataStructure.CellText(arrStatutVen(tabOps(indexDepense).LigneOrigine, 1))
            honoraireActuel = (arrHonoraireVen(tabOps(indexDepense).LigneOrigine, 1) = True)
        End If

        ' CORRECTIF du 16/10/2026 : un groupe dont la dépense est déjà "OK" (calculé, ou posé
        ' à la main : clé "non reprise historique") n'est pas recalculé, mais son "OK" est
        ' propagé à TOUTES ses lignes. Avant, les remboursements rattachés plus tard restaient
        ' vides ou "KO" alors que le groupe était soldé.
        If statutActuel = "OK" Then
            For k = 1 To nbLignesGroupe
                j = indicesGroupe(k)
                If tabOps(j).Source = "O" Then
                    arrStatut(tabOps(j).LigneOrigine, 1) = "OK"
                Else
                    arrStatutVen(tabOps(j).LigneOrigine, 1) = "OK"
                End If
            Next k
            ' Un solde déjà écrit avec un résidu de calcul (ex. 3,55E-15) est ramené au centime.
            If tabOps(indexDepense).Source = "O" Then
                If IsNumeric(arrSolde(tabOps(indexDepense).LigneOrigine, 1)) And Not IsEmpty(arrSolde(tabOps(indexDepense).LigneOrigine, 1)) Then
                    arrSolde(tabOps(indexDepense).LigneOrigine, 1) = Round(CDbl(arrSolde(tabOps(indexDepense).LigneOrigine, 1)), 2)
                End If
            Else
                If IsNumeric(arrSoldeVen(tabOps(indexDepense).LigneOrigine, 1)) And Not IsEmpty(arrSoldeVen(tabOps(indexDepense).LigneOrigine, 1)) Then
                    arrSoldeVen(tabOps(indexDepense).LigneOrigine, 1) = Round(CDbl(arrSoldeVen(tabOps(indexDepense).LigneOrigine, 1)), 2)
                End If
            End If
            Exit Sub
        End If
        ' Un groupe "KO" dont le dépassement d'honoraires a été validé n'est plus soumis à
        ' l'opérateur (voir mod_SuiviSanteFormulaire), mais il est RECALCULÉ : si un
        ' remboursement ultérieur solde le groupe, il doit passer à "OK" (correctif du 16/10/2026).
        ' La variable honoraireActuel n'est donc plus utilisée pour sortir.
    End If

    ' --- Sous-étape C : calcul du solde (uniquement si une ligne de dépense
    '     existe), franchise lue dans le tableau mémoire de sa source -----------
    If indexDepense <> 0 Then
        montantFranchise = 0
        If tabOps(indexDepense).Source = "O" Then
            If IsNumeric(arrFranchise(tabOps(indexDepense).LigneOrigine, 1)) Then
                montantFranchise = mod_DataStructure.ToDouble(arrFranchise(tabOps(indexDepense).LigneOrigine, 1))
            End If
        Else
            If IsNumeric(arrFranchiseVen(tabOps(indexDepense).LigneOrigine, 1)) Then
                montantFranchise = mod_DataStructure.ToDouble(arrFranchiseVen(tabOps(indexDepense).LigneOrigine, 1))
            End If
        End If

        ' Arrondi au centime (correctif du 16/10/2026) : en virgule flottante, 98,43 - 85,43 - 13
        ' donne 3,55E-15 et non 0. Le statut tolérait déjà ce résidu, mais la valeur écrite dans
        ' SoldeSante s'affichait en notation scientifique.
        solde = Round(montantDepenses - montantRemb - montantFranchise, 2)

        If tabOps(indexDepense).Source = "O" Then
            arrSolde(tabOps(indexDepense).LigneOrigine, 1) = solde
            If IsEmpty(arrHonoraire(tabOps(indexDepense).LigneOrigine, 1)) Then
                arrHonoraire(tabOps(indexDepense).LigneOrigine, 1) = False
            End If
        Else
            arrSoldeVen(tabOps(indexDepense).LigneOrigine, 1) = solde
            If IsEmpty(arrHonoraireVen(tabOps(indexDepense).LigneOrigine, 1)) Then
                arrHonoraireVen(tabOps(indexDepense).LigneOrigine, 1) = False
            End If
        End If

        ' Tolérance de 0,005 (un demi-centime) pour l'égalité à zéro, comme en phase 1.
        nouveauStatut = IIf(Abs(solde) < 0.005, "OK", "KO")
    Else
        ' Groupe composé uniquement de remboursements, sans dépense
        ' correspondante : impossible de calculer un solde. Convention
        ' inchangée depuis la phase 1 : ces lignes orphelines restent "KO".
        nouveauStatut = "KO"
    End If

    ' --- Sous-étape D : écriture du statut sur toutes les lignes du groupe,
    '     chacune dans le tableau mémoire de sa source -------------------------
    For k = 1 To nbLignesGroupe
        j = indicesGroupe(k)
        If tabOps(j).Source = "O" Then
            arrStatut(tabOps(j).LigneOrigine, 1) = nouveauStatut
        Else
            arrStatutVen(tabOps(j).LigneOrigine, 1) = nouveauStatut
        End If
    Next k

    nbGroupesTraites = nbGroupesTraites + 1

End Sub
