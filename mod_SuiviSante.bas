Option Explicit

' ============================================================================
'  MODULE : mod_SuiviSante
'  ROLE   : Suivi persistant des opÃ©rations de la sous-catÃ©gorie "Frais, remb
'           santÃ©" dans TblOperations -- ET, depuis la PHASE 6, Ã©galement dans
'           TblVentilations (les lignes de ventilation affectees a cette mÃªme
'           sous-catÃ©gorie).
'
'  HISTORIQUE DES PHASES :
'   - Phase 1 (version d'origine) : calcul automatique de StatutSante et
'     SoldeSante, en isolant les lignes de TblOperations dont la COLONNE
'     "CatÃ©gorie" valait "Frais, remb santÃ©" (a l'epoque, il n'existait pas
'     encore de sous-catÃ©gorie : CatÃ©gorie ETAIT la plus fine granularite).
'   - Phase 6 (ce fichier) : DEUX changements structurels, demandes par
'     l'opÃ©rateur lors de la mise en place de la gestion des catÃ©gories a
'     2 niveaux (CatÃ©gorie / SousCategorie) et de la ventilation :
'       (a) le test bascule sur la colonne "SousCategorie" (la colonne
'           "CatÃ©gorie" contient desormais la catÃ©gorie PARENTE, par exemple
'           "SantÃ©, prevoyance", commune a plusieurs sous-catÃ©gories) ;
'       (b) les lignes de TblVentilations dont la SousCategorie vaut aussi
'           "Frais, remb santÃ©" entrent DANS LA MEME MECANIQUE qu'une ligne
'           normale de TblOperations : mÃªme regroupement par clÃ© "Notes",
'           mÃªme calcul de solde, mÃªme ecriture de StatutSante/SoldeSante --
'           mais chacune dans SA PROPRE table (TblOperations ou
'           TblVentilations), puisque ce sont deux tableaux Excel distincts.
'           C'est la consÃ©quence directe de la decision de l'opÃ©rateur :
'           "les opÃ©rations ventilees entrent dans la mecanique principale
'           ... a terme elles doivent Ãªtre considÃ©rÃ©es comme une opÃ©ration
'           comme une autre."
'
'  PRINCIPE DE REGROUPEMENT (inchange depuis la Phase 1) : un "groupe" est
'  l'ensemble des lignes qui partagent EXACTEMENT la mÃªme valeur dans la
'  colonne "Notes" -- qu'elles viennent de TblOperations ou de
'  TblVentilations. Un groupe contient en general : une ligne de depense
'  (montant negatif) et une ou plusieurs lignes de remboursement (montant
'  positif). Avec la Phase 6, la ligne de DEPENSE peut desormais venir soit
'  de TblOperations (opÃ©ration classique), soit de TblVentilations (partie
'  d'une opÃ©ration ventilee) -- le remboursement, lui, reste presque toujours
'  une ligne bancaire classique de TblOperations (un remboursement de
'  mutuelle n'est jamais ventile).
'
'  REGLE DE NON-RETRAITEMENT (pour ne pas re-poser les mÃªmes questions a
'  chaque import) : un groupe n'est PAS retraitÃ© si sa ligne de depense a :
'     - StatutSante = "OK"                                   -> rien a faire
'     - StatutSante = "KO" ET DepassementHoraires = VRAI      -> rien a faire
'  Dans tous les autres cas (premiÃ¨re fois, ou "KO" avec Honoraire = FAUX),
'  on recalcule.
'
'  IMPORTANT : ce module ne rÃ©Ã©crit JAMAIS un tableau en entier, et ne
'  rÃ©Ã©crit mÃªme pas la colonne Franchise (saisie reservee a l'opÃ©rateur). Il
'  ne touche qu'aux colonnes qu'il a le droit de modifier (StatutSante,
'  SoldeSante, DepassementHoraires), chacune via tbl.ListColumns("NomColonne"),
'  et ce SEPAREMENT pour TblOperations et pour TblVentilations.
'
'  TblVentilations EST FACULTATIVE : si la feuille "Ventilations" ou le
'  tableau "TblVentilations" n'existent pas encore (Phase 4 pas installee
'  chez l'opÃ©rateur), ou si ses colonnes de suivi santÃ© n'ont pas encore ete
'  ajoutees, ce module continue de fonctionner NORMALEMENT sur TblOperations
'  seule -- aucune erreur n'est levee, la partie Ventilations est simplement
'  ignoree.
' ============================================================================

' --- Petit "type" prive, utilisÃ© pour charger en memoire les opÃ©rations de
'     santÃ© des DEUX tables avant de les regrouper. Le champ "Source" indique
'     dans quel tableau la ligne a ete trouvee ("O" = TblOperations,
'     "V" = TblVentilations) : c'est lui qui dit, au moment de reecrire les
'     rÃ©sultats, quel jeu de tableaux memoire (OP ou VEN) utiliser. ---
Private Type TOperationSante
    Source As String        ' "O" (TblOperations) ou "V" (TblVentilations)
    LigneOrigine As Long    ' numÃ©ro de ligne DANS LE TABLEAU MEMOIRE DE SA SOURCE (tblData ou tblDataVen), PAS le numÃ©ro de ligne Excel
    Montant As Double       ' montant tel quel (negatif pour une depense, positif pour un remboursement)
    Notes As String         ' valeur du champ Notes : c'est la clÃ© de regroupement
    Traite As Boolean       ' cette ligne a-t-elle dÃ©jÃ  ete rattachee a un groupe traitÃ© ?
End Type

' Nom exact de la sous-catÃ©gorie a surveiller. On utilisÃ© ChrW(233) pour le
' "e" accentue de "santÃ©" au lieu de taper directement le caractere accentue :
' cela garantit que la comparaison fonctionnera correctement quel que soit
' l'encodage utilisÃ© au moment de l'import du fichier .bas.
Private Function SousCategorieSante() As String
    SousCategorieSante = "Frais, remb sant" & ChrW(233)
End Function

' --- Noms de la feuille et du tableau de ventilation, DUPLIQUES ICI en dur
'     (plutot que de referencer mod_InstallVentilation.VEN_NOM_xxx) pour que
'     ce module continue a COMPILER mÃªme si mod_InstallVentilation n'a pas
'     encore ete importe (Phase 4 optionnelle vis-a-vis de ce module) --------
Private Const VEN_FEUILLE As String = "Ventilations"
Private Const VEN_TABLE As String = "TblVentilations"

' ----------------------------------------------------------------------------
' ValeursColonne : lit une colonne du tableau et renvoie TOUJOURS un tableau
' 2D (1 To n, 1 To 1), mÃªme si la table ne contient qu'une seule ligne.
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
' ou Nothing si la feuille/le tableau n'existe pas encore chez l'opÃ©rateur
' (Phase 4 pas installee). Ne leve jamais d'erreur.
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
' (Ctrl+G, CalculerSuiviSante) pour retraiter toute la table (par exemple
' aprÃ¨s une saisie manuelle de Franchise).
Public Sub CalculerSuiviSante(Optional ByVal AfficherResume As Boolean = True)

    Dim tblData As Variant
    Dim nbLignesTable As Long
    Dim nbLigne As Long

    Dim arrStatut As Variant, arrSolde As Variant, arrHonoraire As Variant, arrFranchise As Variant

    ' --- Jeux de tableaux memoire pour TblVentilations (Phase 6) : restent
    '     vides (Empty) si TblVentilations n'existe pas ou n'a pas encore ses
    '     colonnes de suivi santÃ© -------------------------------------------
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

    ' --- ETAPE 1 : rÃ©cupÃ©ration de TblOperations et des index de colonnes ---
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
    ' recherchant chaque colonne PAR SON NOM D'ENTETE (jamais un numÃ©ro fixe).
    mod_Display.RecupIndexCol

    ' PHASE 6 : la colonne testee est desormais SousCategorie (et non plus
    ' CatÃ©gorie). Si elle n'existe pas encore (Phase 1 pas installee), on
    ' previent clairement au lieu de comparer une colonne qui n'a pas le sens
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

    ' --- ETAPE 2 : lecture memoire de TblOperations ---
    tblData = tbl.DataBodyRange.value
    nbLignesTable = UBound(tblData, 1)

    arrStatut = ValeursColonne(tbl.ListColumns("StatutSante").DataBodyRange)
    arrSolde = ValeursColonne(tbl.ListColumns("SoldeSante").DataBodyRange)
    arrHonoraire = ValeursColonne(tbl.ListColumns("DepassementHoraires").DataBodyRange)
    arrFranchise = ValeursColonne(tbl.ListColumns("Franchise").DataBodyRange)   ' jamais ecrite, lue seulement

    ' --- ETAPE 2bis (PHASE 6) : lecture memoire de TblVentilations, si dispo ---
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
                venDisponible = True
            End If
            ' Si une colonne de suivi manque encore dans TblVentilations, on ne
            ' bloque pas : on continue simplement sans elle (venDisponible reste False).
        End If
    End If

    ' --- ETAPE 3 : on isole toutes les lignes "Frais, remb santÃ©", des DEUX
    '     tables, dans un seul tableau memoire combine ------------------------
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

    ' --- ETAPE 4 : on traitÃ© chaque groupe (= chaque valeur de Notes) une
    '     seule fois. Le regroupement se fait SUR L'ENSEMBLE COMBINE : deux
    '     lignes venant de tables diffÃ©rentes mais partageant la mÃªme clÃ©
    '     Notes sont bien traitÃ©es comme UN SEUL groupe. -----------------------
    nbGroupesTraites = 0
    For i = 1 To nbOps
        If Not tabOps(i).Traite Then
            TraiterGroupeSante tabOps, nbOps, i, _
                arrStatut, arrSolde, arrHonoraire, arrFranchise, _
                arrStatutVen, arrSoldeVen, arrHonoraireVen, arrFranchiseVen, _
                nbGroupesTraites
        End If
    Next i

    ' --- ETAPE 5 : rÃ©Ã©criture UNIQUEMENT des colonnes concernÃ©es, table par
    '     table. On n'ecrit jamais un tableau en entier : seules les colonnes
    '     de suivi sont reecrites, chacune independamment. ---------------------
    Application.ScreenUpdating = False
    tbl.ListColumns("StatutSante").DataBodyRange.value = arrStatut
    tbl.ListColumns("SoldeSante").DataBodyRange.value = arrSolde
    tbl.ListColumns("DepassementHoraires").DataBodyRange.value = arrHonoraire

    If venDisponible Then
        tblVen.ListColumns("StatutSante").DataBodyRange.value = arrStatutVen
        tblVen.ListColumns("SoldeSante").DataBodyRange.value = arrSoldeVen
        tblVen.ListColumns("DepassementHoraires").DataBodyRange.value = arrHonoraireVen
    End If
    Application.ScreenUpdating = True

    If AfficherResume Then
        MsgBox nbGroupesTraites & " groupe(s) de depenses sante recalcule(s) sur " & _
               nbOps & " ligne(s) '" & SousCategorieSante() & "' analysee(s) " & _
               "(TblOperations" & IIf(venDisponible, " + TblVentilations", "") & ").", _
               vbInformation, "Suivi sante"
    End If

End Sub


' ----------------------------------------------------------------------------
' TraiterGroupeSante : traitÃ© UN groupe (toutes les lignes qui partagent la
' mÃªme valeur de "Notes" que la ligne tabOps(indexDepart)), en lisant/ecrivant
' dans le jeu de tableaux memoire correspondant a la SOURCE de chaque ligne.
' ----------------------------------------------------------------------------
' Parametres :
'   tabOps                        -> toutes les opÃ©rations santÃ© (2 tables), en memoire
'   nbOps                         -> nombre de lignes dans tabOps
'   indexDepart                   -> position, dans tabOps, d'une ligne du groupe a traiter
'   arrStatut/arrSolde/arrHonoraire/arrFranchise (ByRef)      -> colonnes de TblOperations
'   arrStatutVen/arrSoldeVen/arrHonoraireVen/arrFranchiseVen (ByRef) -> colonnes de TblVentilations
'   nbGroupesTraites (ByRef)      -> compteur global, incremente si ce groupe a ete recalcule
Private Sub TraiterGroupeSante(ByRef tabOps() As TOperationSante, ByVal nbOps As Long, _
                                ByVal indexDepart As Long, _
                                ByRef arrStatut As Variant, ByRef arrSolde As Variant, _
                                ByRef arrHonoraire As Variant, ByRef arrFranchise As Variant, _
                                ByRef arrStatutVen As Variant, ByRef arrSoldeVen As Variant, _
                                ByRef arrHonoraireVen As Variant, ByRef arrFranchiseVen As Variant, _
                                ByRef nbGroupesTraites As Long)

    Dim notesRef As String
    Dim j As Long, k As Long
    Dim indexDepense As Long        ' position DANS TABOPS (pas dans une table) de la 1ere depense du groupe, 0 si aucune
    Dim montantDepenses As Double
    Dim montantRemb As Double
    Dim nbLignesGroupe As Long
    Dim indicesGroupe() As Long      ' positions DANS TABOPS de toutes les lignes du groupe
    Dim statutActuel As String
    Dim honoraireActuel As Boolean
    Dim nouveauStatut As String
    Dim montantFranchise As Double
    Dim solde As Double

    notesRef = tabOps(indexDepart).Notes

    ' --- Sous-Ã©tape A : on repÃ¨re TOUTES les lignes du groupe (peu importe
    '     leur source) et on les marque "TraitÃ©" tout de suite --------------
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

    ' --- Sous-Ã©tape B : faut-il seulement RETRAITER ce groupe ? (le verrou se
    '     lit dans le jeu de tableaux memoire de la source de la depense) ------
    If indexDepense <> 0 Then
        If tabOps(indexDepense).Source = "O" Then
            statutActuel = mod_DataStructure.CellText(arrStatut(tabOps(indexDepense).LigneOrigine, 1))
            honoraireActuel = (arrHonoraire(tabOps(indexDepense).LigneOrigine, 1) = True)
        Else
            statutActuel = mod_DataStructure.CellText(arrStatutVen(tabOps(indexDepense).LigneOrigine, 1))
            honoraireActuel = (arrHonoraireVen(tabOps(indexDepense).LigneOrigine, 1) = True)
        End If

        If statutActuel = "OK" Then Exit Sub
        If statutActuel = "KO" And honoraireActuel Then Exit Sub
    End If

    ' --- Sous-Ã©tape C : calcul du solde (uniquement si une ligne de depense
    '     existe), franchise lue dans le jeu de tableaux memoire de sa source --
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

        solde = montantDepenses - montantRemb - montantFranchise

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

        ' Tolerance de 0,005 (un demi-centime) pour l'Ã©galitÃ© a zero, comme en Phase 1.
        nouveauStatut = IIf(Abs(solde) < 0.005, "OK", "KO")
    Else
        ' Groupe compose uniquement de remboursements, sans depense
        ' correspondante : impossible de calculer un solde. Convention
        ' inchangee depuis la Phase 1 : ces lignes orphelines restent "KO".
        nouveauStatut = "KO"
    End If

    ' --- Sous-Ã©tape D : ecriture du statut sur TOUTES les lignes du groupe,
    '     chacune dans le jeu de tableaux memoire de SA source ---------------
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
