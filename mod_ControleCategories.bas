Option Explicit

' =====================================================================================
' MODULE : mod_ControleCategories
'
' PHASE 2 du chantier « Catégorie / Sous-catégorie / Ventilation » - PARTIE 2/2
'
' RÔLE (à lire en premier, même si vous débutez) :
'   Ce module contient toute la logique du contrôle des catégories des opérations
'   importées. Il s'appuie sur :
'     - le tableau de correspondance TblCategories (Phase 1, feuille Param),
'     - la feuille-formulaire frm_ControleCategories (mod_InstallControleCategories).
'
'   DÉROULEMENT (fonction ControlerCategories) :
'     1. Pour chaque opération, on cherche sa catégorie SOURCE (celle de la banque)
'        dans TblCategories -> on en déduit les catégories et sous-catégories proposées.
'     2. On demande à l'opérateur :
'          "Les catégories ont-elles été correctement renseignées dans la source ?"
'            OUI -> on garde les propositions telles quelles (voir remarque ci-dessous).
'            NON -> le formulaire présente les opérations une par une.
'            ANNULER -> on abandonne, rien n'est applique.
'     3. Dans le formulaire, l'opérateur peut changer la catégorie et la
'        sous-catégorie de chaque opération avec Précédent / Suivant.
'     4. "Terminer et continuer" renvoie les valeurs finales à l'appelant.
'
'   REMARQUE IMPORTANTE (choix de conception à valider) :
'     Même si l'opérateur répond OUI, les opérations dont la catégorie source est
'     VIDE ou ABSENTE du tableau de correspondance sont tout de même présentées
'     dans le formulaire : sans cela, on ne saurait pas dans quelle catégorie les ranger.
'     Ce comportement peut être désactivé en passant CTRL_FORCER_NON_RANGEES à False (plus bas).
'
'   CE MODULE N'ÉCRIT RIEN dans TblOperations. Il RENVOIE des résultats; c'est
'   l'appelant (plus tard : l'import, phase 5) qui décidera de les écrire.
'   Pour le tester sans risque, utilisez la macro TesterControleCategories.
'
' À PROPOS DES ACCENTS : les textes affichés sont construits par la fonction mod_Display.FR();
' les commentaires, eux, sont encodés en UTF-8.
' =====================================================================================

' --- Positions des colonnes du tableau "ops" que l'appelant doit fournir ---------------
' ops(i, CTRL_OP_DATE)      = date de l'opération (Date ou nombre série Excel)
' ops(i, CTRL_OP_TIERS)     = tiers (texte)
' ops(i, CTRL_OP_LIBELLE)   = libellé / notes (texte)
' ops(i, CTRL_OP_MONTANT)   = montant (nombre, négatif pour une dépense)
' ops(i, CTRL_OP_CATSOURCE) = catégorie envoyée par la banque (texte, éventuellement vide)
' ops(i, CTRL_OP_ID)        = ID_Transaction de l'opération (texte). NOUVEAU (phase 4) :
'                             nécessaire pour relier une ventilation (TblVentilations) à
'                             l'opération d'origine dans TblOperations.
Public Const CTRL_OP_DATE As Long = 1
Public Const CTRL_OP_TIERS As Long = 2
Public Const CTRL_OP_LIBELLE As Long = 3
Public Const CTRL_OP_MONTANT As Long = 4
Public Const CTRL_OP_CATSOURCE As Long = 5
Public Const CTRL_OP_ID As Long = 6
Public Const CTRL_OP_NBCOL As Long = 6

' --- Réglage : présenter aussi les opérations "non rangées" après un OUI ? ------------
Private Const CTRL_FORCER_NON_RANGEES As Boolean = True

' --- Nom (au niveau du classeur) de la liste des sous-categories du formulaire ---------
Private Const CTRL_NOM_LISTE_SOUS As String = "ListeSousCatControle"
Private Const CTRL_LIGNE_AIDE_MAX As Long = 300   ' dernière ligne de la zone technique

' --- État du formulaire ------------------------------------------------------------------
' Vrai tant que le formulaire est ouvert. Public, car lu par le verrou de la feuille.
Public g_CtrlEnCours As Boolean

' --- Verrou d'ACTIVATION, distinct de g_CtrlEnCours (Phase 3) -----------------------
' g_CtrlEnCours pilote la boucle d'attente (DoEvents) : elle NE DOIT PAS repasser à
' False tant que le formulaire de contrôle n'est pas terminé.
' g_CtrlVerrouActif pilote uniquement le réflexe qui ramène de force l'opérateur sur
' la feuille (Worksheet_Deactivate). Sans cette distinction, ouvrir le formulaire de
' création de catégorie (bouton "+") serait impossible : dès qu'on quitterait la feuille
' de contrôle pour afficher celle de création, Worksheet_Deactivate nous y ramènerait
' aussitôt de force, empêchant le second formulaire de s'afficher.
' ControleNouvelleCategorie désactive ce verrou juste le temps d'ouvrir le formulaire de
' création, puis le réactive dès qu'il se referme.
Public g_CtrlVerrouActif As Boolean

Private g_CtrlValide As Boolean          ' Vrai si l'opérateur a validé (Terminer)
Private g_CtrlNomFeuillePrec As String   ' feuille active avant l'ouverture (pour y revenir)

Private g_CtrlOps As Variant             ' copie des opérations fournies
Private g_CtrlNbOps As Long
Private g_CtrlCat() As String            ' catégorie courante de chaque opération
Private g_CtrlSous() As String           ' sous-catégorie courante de chaque opération
' Ajout 01/10/2026 (point 4 : annuler une ventilation) : catégorie/sous-catégorie
' telles qu'elles étaient avant la toute première ventilation de l'opération (vides tant
' que l'opération n'a jamais été ventilée). Mémorisées par ControleVentiler avant
' d'écraser g_CtrlCat/g_CtrlSous par "Ventile", puis renvoyées à l'appelant (import) via
' ControlerCategories pour être écrites dans les deux colonnes techniques de TblOperations
' (CategorieAvantVentilation / SousCategorieAvantVentilation; voir mod_InstallVentilation).
Private g_CtrlCatAvantVen() As String
Private g_CtrlSousAvantVen() As String
' AJOUT 03/10/2026 (demande opérateur : corriger Tiers/Notes à l'import) :
' valeurs COURANTES du Tiers et des Notes de chaque opération. Initialisées avec la
' valeur d'origine (fournie dans "ops"), puis mises à jour par
' EnregistrerOperationAffichee chaque fois que l'opérateur modifie la cellule
' correspondante, SAUF pour une opération de santé (sous-catégorie
' mod_VarGlobales.SOUS_CATEGORIE_SANTE), où toute saisie est ignorée. Voir
' AfficherOperation et EnregistrerOperationAffichee pour le détail de cette règle.
Private g_CtrlTiers() As String
Private g_CtrlLibelle() As String
Private g_CtrlVu() As Boolean            ' opération déjà affichée / enregistrée
Private g_CtrlARanger() As Boolean       ' correspondance introuvable (à ranger à la main)

Private g_CtrlIndices() As Long          ' opérations présentées : position -> numéro d'opération
Private g_CtrlNbAffiches As Long
Private g_CtrlPos As Long                ' position affichée actuellement (1..g_CtrlNbAffiches)

' --- Ajout 02/10/2026 (édition directe depuis l'écran de recherche) -------------------
' Vrai uniquement quand ControlerCategories a été appelée avec uneSeuleOperation:=True
' (voir mod_RechercheOperations.EditerCategorieRO). Dans ce mode : on saute la question
' "Oui/Non/Annuler" (sans objet pour une seule opération déjà catégorisée, hors import)
' et on affiche directement l'opération avec sa catégorie/sous-catégorie ACTUELLE
' préremplie (et non une proposition issue du tableau de correspondance). Le message
' d'aide du formulaire est adapté en conséquence (voir AfficherOperation).
Private g_CtrlModeUnique As Boolean

' --- Référentiel lu dans TblCategories (rechargé à chaque appel) -----------------------
Private g_CtrlMap As Object              ' source (minuscules) -> "categorie" & Tab & "sous-categorie"
Private g_CtrlPaires As Object           ' "categorie" & Tab & "sous-categorie" -> sous-categorie
Private g_CtrlCats As Object             ' catégorie -> catégorie (forme canonique)
Private g_CtrlRefCat() As String
Private g_CtrlRefSous() As String
Private g_CtrlRefN As Long


' =====================================================================================
' FONCTION PRINCIPALE
' =====================================================================================
' Paramètres :
'   ops        : tableau à deux dimensions (1 To nbOps, 1 To CTRL_OP_NBCOL), voir plus haut
'                (transmis par valeur : l'original n'est jamais modifié)
'   nbOps      : nombre d'opérations à contrôler
'   catFinale  : en sortie, catégorie retenue pour chaque opération, indice 1..nbOps
'   sousFinale : en sortie, sous-catégorie retenue pour chaque opération
'   catAvantVentilation, sousAvantVentilation : en sortie (ajout du 01/10/2026, point 4),
'                catégories/sous-catégories d'AVANT LA VENTILATION pour chaque opération
'                (chaîne vide si l'opération n'a jamais été ventilée). À écrire par
'                l'appelant dans les colonnes CategorieAvantVentilation /
'                SousCategorieAvantVentilation de TblOperations afin de restaurer
'                la catégorie d'origine si la ventilation est supprimée plus tard.
'   tiersFinale, libelleFinale : en sortie (ajout du 03/10/2026, demande opérateur) :
'                Tiers et Notes retenus pour chaque opération : valeur d'origine
'                (fournie dans ops) si l'opérateur n'y a pas touché, ou sa correction
'                sinon. EXCEPTION : pour une opération dont la sous-catégorie est
'                mod_VarGlobales.SOUS_CATEGORIE_SANTE, ces 2 champs restent TOUJOURS
'                figés sur leur valeur d'origine, même si l'opérateur saisit autre
'                chose dans la cellule (le champ Notes encode une date de consultation
'                lue ensuite par mod_ImportOFX; voir AfficherOperation). Pour la même
'                raison, en mode "une seule opération" (uneSeuleOperation:=True; voir
'                plus bas, pour l'édition directe depuis l'écran de recherche), ces deux
'                champs restent également figés : cette correction est réservée à l'import.
'   uneSeuleOperation, categorieActuelleUnique, sousCategorieActuelleUnique :
'                (ajout du 02/10/2026, édition directe depuis l'écran de recherche) :
'                paramètres facultatifs, NON utilisés par l'import (valeurs par défaut
'                False/"" : comportement inchangé pour tout appel qui ne les précise pas).
'                Quand uneSeuleOperation:=True (nbOps doit alors valoir 1), la question
'                "Oui/Non/Annuler" est ignorée et le formulaire est prérempli avec
'                categorieActuelleUnique / sousCategorieActuelleUnique (catégorie actuelle
'                de l'opération), plutôt qu'avec une correspondance de CTRL_OP_CATSOURCE.
' Renvoie True si l'opérateur a validé (ou s'il n'y avait rien à contrôler),
'         False s'il a annulé ou si un prérequis manque (un message est alors affiché).
Public Function ControlerCategories(ByVal ops As Variant, ByVal nbOps As Long, _
                                    ByRef catFinale() As String, ByRef sousFinale() As String, _
                                    ByRef catAvantVentilation() As String, ByRef sousAvantVentilation() As String, _
                                    ByRef tiersFinale() As String, ByRef libelleFinale() As String, _
                                    Optional ByVal uneSeuleOperation As Boolean = False, _
                                    Optional ByVal categorieActuelleUnique As String = "", _
                                    Optional ByVal sousCategorieActuelleUnique As String = "") As Boolean

    Dim ws As Worksheet
    Dim i As Long, nbARanger As Long
    Dim Source As String
    Dim morceaux() As String
    Dim reponse As VbMsgBoxResult

    On Error GoTo Erreur

    g_CtrlModeUnique = uneSeuleOperation

    ControlerCategories = False
    If nbOps <= 0 Then
        ControlerCategories = True     ' rien à contrôler : on laisse continuer
        Exit Function
    End If

    ' --- Prérequis 1 : la feuille-formulaire doit exister -------------------------------
    Set ws = FeuilleSansErreur(CTRL_NOM_FEUILLE)
    If ws Is Nothing Then
        MsgBox mod_Display.FR("La feuille '") & CTRL_NOM_FEUILLE & mod_Display.FR("' est introuvable.") & vbCrLf & _
               mod_Display.FR("Ex{e2}cutez d'abord la macro CreerFeuilleControleCategories."), vbExclamation
        Exit Function
    End If

    ' --- Prérequis 2 : le tableau de correspondance doit exister et contenir des lignes ---
    If Not ChargerReferentiel() Then
        MsgBox mod_Display.FR("Le tableau de correspondance TblCategories est introuvable ou vide.") & vbCrLf & _
               mod_Display.FR("Ex{e2}cutez d'abord la macro PreparerPhase1Categories."), vbExclamation
        Exit Function
    End If

    ' --- On travaille sur des COPIES : ops n'est jamais modifié ---------------------------
    g_CtrlOps = ops
    g_CtrlNbOps = nbOps
    ReDim g_CtrlCat(1 To nbOps)
    ReDim g_CtrlSous(1 To nbOps)
    ReDim g_CtrlCatAvantVen(1 To nbOps)
    ReDim g_CtrlSousAvantVen(1 To nbOps)
    ReDim g_CtrlVu(1 To nbOps)
    ReDim g_CtrlARanger(1 To nbOps)
    ReDim g_CtrlTiers(1 To nbOps)
    ReDim g_CtrlLibelle(1 To nbOps)

    ' --- ÉTAPE 1 : application automatique de la correspondance --------------------------
    For i = 1 To nbOps
        Source = mod_DataStructure.CellText(g_CtrlOps(i, CTRL_OP_CATSOURCE))
        If Source <> "" And g_CtrlMap.Exists(Source) Then
            ' Item = "categorie" & Tab & "sous-categorie" : on le redécoupe.
            morceaux = Split(CStr(g_CtrlMap(Source)), vbTab)
            g_CtrlCat(i) = morceaux(0)
            g_CtrlSous(i) = morceaux(1)
        Else
            g_CtrlCat(i) = ""
            g_CtrlSous(i) = ""
            g_CtrlARanger(i) = True
            nbARanger = nbARanger + 1
        End If
        ' AJOUT 03/10/2026 : valeur de départ = valeur fournie par l'appelant. Elle ne
        ' change que si l'opérateur la modifie ET que cela est autorisé (voir
        ' EnregistrerOperationAffichee).
        g_CtrlTiers(i) = mod_DataStructure.CellText(g_CtrlOps(i, CTRL_OP_TIERS))
        g_CtrlLibelle(i) = mod_DataStructure.CellText(g_CtrlOps(i, CTRL_OP_LIBELLE))
    Next i

    ' Ajout du 02/10/2026 : en mode "édition directe" (uneSeuleOperation:=True), on ignore
    ' le résultat de l'ÉTAPE 1 (il n'y a pas de source bancaire à faire correspondre;
    ' CTRL_OP_CATSOURCE n'est même pas renseigné dans ce mode). On préremplit le formulaire
    ' avec la catégorie actuelle de l'opération et on la marque "pas à ranger"
    ' (g_CtrlARanger = False), puisqu'elle est déjà valide. Voir AfficherOperation.
    If g_CtrlModeUnique Then
        g_CtrlCat(1) = categorieActuelleUnique
        g_CtrlSous(1) = sousCategorieActuelleUnique
        g_CtrlARanger(1) = False
        ' Ajout 09/10/2026 : le champ "Categorie source" affiche alors la categorie et la
        ' sous-categorie d'origine ("Categorie / Sous-categorie"), plutot que de rester vide.
        If sousCategorieActuelleUnique <> "" Then
            g_CtrlOps(1, CTRL_OP_CATSOURCE) = categorieActuelleUnique & " / " & sousCategorieActuelleUnique
        Else
            g_CtrlOps(1, CTRL_OP_CATSOURCE) = categorieActuelleUnique
        End If
    End If

    ' --- ÉTAPE 2 : la question à l'opérateur -----------------------------------------------
    ' En mode "edition directe", cette question n'a pas de sens (il n'y a qu'UNE
    ' opération déjà catégorisée, hors import) : on passe directement à la présentation
    ' de l'opération (même résultat qu'une réponse "NON" ici).
    If g_CtrlModeUnique Then
        reponse = vbNo
    Else
        reponse = MsgBox(nbOps & mod_Display.FR(" op{e2}ration(s) import{e2}e(s).") & vbCrLf & vbCrLf & _
                         mod_Display.FR("Les cat{e2}gories de ces op{e2}rations ont-elles {e2}t{e2} correctement renseign{e2}es dans la source (la banque) ?") & _
                         vbCrLf & vbCrLf & _
                         mod_Display.FR("OUI : la correspondance est appliqu{e2}e automatiquement.") & vbCrLf & _
                         mod_Display.FR("NON : les op{e2}rations vous sont pr{e2}sent{e2}es une par une pour les v{e2}rifier.") & vbCrLf & _
                         mod_Display.FR("ANNULER : abandon, rien n'est appliqu{e2}."), _
                         vbYesNoCancel + vbQuestion, mod_Display.FR("Contr{o2}le des cat{e2}gories"))
    End If

    If reponse = vbCancel Then Exit Function      ' False : abandon

    ' --- Quelles opérations présenter dans le formulaire ? ---------------------------------
    ReDim g_CtrlIndices(1 To nbOps)
    g_CtrlNbAffiches = 0

    If reponse = vbNo Then
        ' Cas 1 - "NON" (ou mode "édition directe") : toutes les opérations sont présentées.
        For i = 1 To nbOps
            g_CtrlNbAffiches = g_CtrlNbAffiches + 1
            g_CtrlIndices(g_CtrlNbAffiches) = i
        Next i
    Else
        ' "OUI" : on n'affiche que les opérations impossibles à ranger automatiquement.
        If CTRL_FORCER_NON_RANGEES And nbARanger > 0 Then
            MsgBox nbARanger & mod_Display.FR(" op{e2}ration(s) ont une cat{e2}gorie source vide ou absente du tableau de correspondance.") & _
                   vbCrLf & mod_Display.FR("Elles vont vous {e2}tre pr{e2}sent{e2}es pour {e2}tre rang{e2}es."), _
                   vbInformation, mod_Display.FR("Contr{o2}le des cat{e2}gories")
            For i = 1 To nbOps
                If g_CtrlARanger(i) Then
                    g_CtrlNbAffiches = g_CtrlNbAffiches + 1
                    g_CtrlIndices(g_CtrlNbAffiches) = i
                End If
            Next i
        End If
    End If

    ' --- ÉTAPE 3 : formulaire (uniquement s'il y a quelque chose à présenter) ---------------
    If g_CtrlNbAffiches = 0 Then
        g_CtrlValide = True
    Else
        OuvrirFormulaireEtAttendre ws
    End If

    ' --- ÉTAPE 4 : restitution des résultats -------------------------------------------------
    ControlerCategories = g_CtrlValide
    If g_CtrlValide Then
        ReDim catFinale(1 To nbOps)
        ReDim sousFinale(1 To nbOps)
        ReDim catAvantVentilation(1 To nbOps)
        ReDim sousAvantVentilation(1 To nbOps)
        ReDim tiersFinale(1 To nbOps)
        ReDim libelleFinale(1 To nbOps)
        For i = 1 To nbOps
            catFinale(i) = g_CtrlCat(i)
            sousFinale(i) = g_CtrlSous(i)
            catAvantVentilation(i) = g_CtrlCatAvantVen(i)
            sousAvantVentilation(i) = g_CtrlSousAvantVen(i)
            tiersFinale(i) = g_CtrlTiers(i)
            libelleFinale(i) = g_CtrlLibelle(i)
        Next i
    End If
    Exit Function

Erreur:
    ' Erreur imprévue : on remet Excel en état normal avant d'informer l'opérateur.
    Application.EnableEvents = True
    g_CtrlEnCours = False
    MsgBox mod_Display.FR("Erreur inattendue dans le contr{o2}le des cat{e2}gories :") & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical
    On Error Resume Next
    FermerFormulaire ws
    ControlerCategories = False

End Function

' =====================================================================================
' ModifierCategorieOperation (ajout du 07/10/2026)
' =====================================================================================
' RÔLE : ouvrir le formulaire de contrôle des catégories pour UNE SEULE opération de
' TblOperations (mode "une seule opération" de ControlerCategories, avec la catégorie
' actuelle préremplie), puis écrire le résultat validé dans TblOperations.
'
' POURQUOI CETTE FONCTION : ce traitement existait déjà, mais il était écrit
' directement dans mod_RechercheOperations.EditerCategorieRO (double-clic sur une
' catégorie dans l'écran de recherche). Le nouveau bouton "Changer la catégorie" de
' frm_RapprochementNotes a besoin EXACTEMENT du même traitement. Plutôt que de le
' recopier, on l'a sorti ici dans une fonction partagée, appelée par les deux écrans :
' un correctif futur ne sera ainsi à faire qu'à un seul endroit.
'
' Paramètres :
'   tblOp   : le tableau TblOperations (déjà obtenu par l'appelant).
'   ligneOp : numéro de ligne DANS LE TABLEAU (1 = première ligne de données),
'             et non un numéro de ligne de la feuille Excel.
' Renvoie True si l'opérateur a validé (Categorie/SousCategorie sont alors déjà
' écrites dans TblOperations), False s'il a annulé (rien n'a été modifié).
'
' Ce que la fonction fait EN PLUS de l'écriture (ajout du 08/10/2026) : si la
' nouvelle sous-catégorie n'est pas "Frais, remb santé", elle vide les colonnes de
' suivi santé de la ligne (voir NettoyerColonnesSante, juste après cette fonction).
'
' Ce que la fonction NE fait PAS, volontairement (cela dépend de l'écran appelant) :
' recalculer le suivi santé, rafraîchir un affichage, réactiver une feuille.
'
' NOUVEAUTÉ par rapport à l'ancien code d'EditerCategorieRO : si l'opérateur a utilisé
' le bouton "Ventiler" du formulaire, la catégorie d'AVANT la ventilation (renvoyée
' par ControlerCategories) est désormais mémorisée dans les colonnes
' CategorieAvantVentilation / SousCategorieAvantVentilation, exactement comme le fait
' déjà l'import (mod_ImportOFX). Sans cela, une suppression ultérieure de cette
' ventilation ne pourrait pas restaurer la catégorie d'origine de l'opération.
' =====================================================================================
Public Function ModifierCategorieOperation(ByVal tblOp As ListObject, ByVal ligneOp As Long) As Boolean

    Dim donneesLigne As Variant
    Dim ops(1 To 1, 1 To CTRL_OP_NBCOL) As Variant
    Dim catFinale() As String, sousFinale() As String
    Dim catAvantVen() As String, sousAvantVen() As String
    ' Tiers/Notes : non modifiables en mode "une seule opération" (voir
    ' TiersNotesEditables). Déclarés uniquement parce que ControlerCategories les exige.
    Dim tiersFinaleInutilise() As String, libelleFinaleInutilise() As String
    Dim categorieActuelle As String, sousCategorieActuelle As String
    Dim colCatAvant As Long, colSousAvant As Long

    ModifierCategorieOperation = False
    If tblOp Is Nothing Then Exit Function
    If tblOp.DataBodyRange Is Nothing Then Exit Function
    If ligneOp < 1 Or ligneOp > tblOp.DataBodyRange.rows.count Then Exit Function

    ' Index des colonnes (colDate, colTiers...) retrouvés par leur nom d'en-tête.
    ' ATTENTION : mod_Display.RecupIndexCol lit la variable GLOBALE "tbl"
    ' (mod_VarGlobales) sans l'initialiser elle-même. On la renseigne donc d'abord
    ' avec le tableau reçu : sinon, si aucun écran ne l'a fait avant, tous les index
    ' vaudraient 0 sans aucun message (même type de dépendance cachée que celle de
    ' wsSynthese, corrigée le 01/10/2026). Il s'agit du même tableau TblOperations :
    ' ce n'est donc pas un changement pour le reste du classeur.
    Set tbl = tblOp
    mod_Display.RecupIndexCol

    ' Lecture de la SEULE ligne concernée (tableau d'une ligne sur N colonnes).
    donneesLigne = tblOp.DataBodyRange.rows(ligneOp).value

    categorieActuelle = mod_DataStructure.CellText(donneesLigne(1, colCategorie))
    sousCategorieActuelle = ""
    If colSousCategorie <> 0 Then sousCategorieActuelle = mod_DataStructure.CellText(donneesLigne(1, colSousCategorie))

    ' Préparation de l'opération au format attendu par ControlerCategories.
    ' CTRL_OP_CATSOURCE reste vide : il n'y a pas de catégorie bancaire source hors import.
    ops(1, CTRL_OP_DATE) = donneesLigne(1, colDate)
    ops(1, CTRL_OP_TIERS) = mod_DataStructure.CellText(donneesLigne(1, colTiers))
    ops(1, CTRL_OP_LIBELLE) = ""
    ops(1, CTRL_OP_MONTANT) = mod_DataStructure.ToDouble(donneesLigne(1, colMontant))
    ops(1, CTRL_OP_CATSOURCE) = ""
    ops(1, CTRL_OP_ID) = mod_DataStructure.CellText(donneesLigne(1, colID))

    If Not ControlerCategories(ops, 1, catFinale, sousFinale, catAvantVen, sousAvantVen, _
                               tiersFinaleInutilise, libelleFinaleInutilise, _
                               uneSeuleOperation:=True, _
                               categorieActuelleUnique:=categorieActuelle, _
                               sousCategorieActuelleUnique:=sousCategorieActuelle) Then
        Exit Function    ' annulation : rien n'est écrit
    End If

    ' --- Écriture du résultat validé ---
    tblOp.DataBodyRange.Cells(ligneOp, colCategorie).value = catFinale(1)
    If colSousCategorie <> 0 Then tblOp.DataBodyRange.Cells(ligneOp, colSousCategorie).value = sousFinale(1)

    ' Catégorie d'avant ventilation : ControlerCategories ne la renseigne QUE si
    ' l'opérateur a cliqué sur "Ventiler" (chaîne vide sinon : on ne touche alors à rien).
    If catAvantVen(1) <> "" Or sousAvantVen(1) <> "" Then
        colCatAvant = mod_Display.GetColumnIndex(tblOp, mod_InstallVentilation.NOM_COL_CAT_AVANT_VENTILATION)
        colSousAvant = mod_Display.GetColumnIndex(tblOp, mod_InstallVentilation.NOM_COL_SOUS_AVANT_VENTILATION)
        If colCatAvant <> 0 And colSousAvant <> 0 Then
            tblOp.DataBodyRange.Cells(ligneOp, colCatAvant).value = catAvantVen(1)
            tblOp.DataBodyRange.Cells(ligneOp, colSousAvant).value = sousAvantVen(1)
        End If
    End If

    ' AJOUT du 08/10/2026 (décision opérateur) : nettoyage des colonnes de suivi
    ' santé. Si l'opération N'EST PLUS (ou n'est pas) en "Frais, remb santé" après
    ' validation, ses colonnes santé n'ont plus de sens : on les vide. C'est fait ICI,
    ' dans la fonction partagée, pour que les DEUX écrans appliquent la même règle :
    '   - l'écran de recherche (double-clic sur Categorie ou SousCategorie) ;
    '   - l'écran de rapprochement (bouton "Changer la catégorie").
    ' Sans ce nettoyage, CalculerSuiviSante (qui ne réécrit QUE les lignes santé)
    ' laissait un ancien "KO" sur la ligne, compté ensuite à tort dans le résumé du
    ' préfiltre "SuiviSante" et affiché en rouge gras.
    ' La règle porte sur la NOUVELLE sous-catégorie (et non sur "était-elle santé
    ' avant ?") : rouvrir puis valider une ligne qui garde un "KO" périmé d'une
    ' ancienne modification la corrige donc aussi. Pour une ligne qui n'a jamais été
    ' de santé, ces colonnes sont déjà vides : le nettoyage ne change rien.
    If sousFinale(1) <> mod_VarGlobales.SOUS_CATEGORIE_SANTE Then
        NettoyerColonnesSante tblOp, ligneOp
    End If

    ModifierCategorieOperation = True

End Function

' Vide, pour UNE ligne de TblOperations, toutes les colonnes de suivi santé listées
' dans mod_VarGlobales.COLONNES_SUIVI_SANTE (Notes exclue, voir ce module).
' Une colonne absente du tableau est simplement ignorée (GetColumnIndex renvoie 0).
'   ligneOp : numéro de ligne DANS LE TABLEAU (1 = première ligne de données).
' (ajout du 08/10/2026 ; utilisée uniquement par ModifierCategorieOperation)
Private Sub NettoyerColonnesSante(ByVal tblOp As ListObject, ByVal ligneOp As Long)

    Dim nomsColonnes() As String
    Dim i As Long
    Dim indexColonne As Long

    nomsColonnes = Split(mod_VarGlobales.COLONNES_SUIVI_SANTE, ";")
    For i = LBound(nomsColonnes) To UBound(nomsColonnes)
        indexColonne = mod_Display.GetColumnIndex(tblOp, nomsColonnes(i))
        If indexColonne <> 0 Then tblOp.DataBodyRange.Cells(ligneOp, indexColonne).ClearContents
    Next i

End Sub


' =====================================================================================
' OUVERTURE DU FORMULAIRE ET ATTENTE ("modal" : le code reste bloqué ici)
' =====================================================================================
Private Sub OuvrirFormulaireEtAttendre(ByVal ws As Worksheet)

    ' On retient la feuille actuelle pour y revenir à la fermeture.
    g_CtrlNomFeuillePrec = ActiveSheet.Name

    ' On met à jour la liste des catégories (phase 1) : elle alimente le menu déroulant.
    mod_Categories.RafraichirListesCategories

    g_CtrlValide = False
    g_CtrlPos = 1
    AfficherOperation ws           ' on remplit AVANT d'afficher (pas de flash à l'écran)

    ws.Visible = xlSheetVisible
    ws.Activate
    mod_InstallCommun.MasquerQuadrillage   ' quadrillage et en-tetes toujours masques (09/10/2026)
    ws.Range(CTRL_ADR_CAT).Select

    ' --- Verrou "modal" : la boucle ne se termine que lorsque g_CtrlEnCours repasse à
    ' False (bouton Terminer ou Annuler). DoEvents laisse Excel traiter les clics.
    '
    ' CORRECTIF (constaté par l'opérateur le 2026-09-25) : sans cette ligne, les listes
    ' déroulantes du formulaire restaient inertes (aucune réaction au changement de
    ' catégorie), car Application.EnableEvents était à False à ce stade, sans qu'aucune
    ' erreur ne le signale. Toute la boucle qui suit repose entièrement sur
    ' Worksheet_Change (voir CtrlTraiterChangement) : on GARANTIT donc ici que les
    ' événements sont actifs, plutôt que de faire confiance à l'état laissé par le reste
    ' du classeur (dont l'origine exacte du problème n'a pas été formellement identifiée).
    Application.EnableEvents = True
    g_CtrlEnCours = True
    g_CtrlVerrouActif = True
    Do While g_CtrlEnCours
        DoEvents
    Loop

End Sub


' =====================================================================================
' TIERS / NOTES MODIFIABLES ? (ajout du 03/10/2026, demande opérateur)
' =====================================================================================
' La règle est centralisée ici afin qu'AfficherOperation (affichage et couleur) et
' EnregistrerOperationAffichee (lecture de la saisie) appliquent TOUJOURS la même
' condition. Si elle était copiée aux deux endroits, un correctif ultérieur appliqué
' à un seul d'entre eux pourrait les désynchroniser (cela s'est déjà produit dans ce
' chantier avec le message de validation de Categorie; voir PoserValidationCategorie).
' i : indice de l'opération (1..nbOps, PAS la position affichée g_CtrlPos).
Private Function TiersNotesEditables(ByVal i As Long) As Boolean
    ' Pas encore active pour l'édition directe depuis l'écran de recherche (une seule
    ' opération) : cette possibilité reste pour l'instant réservée à l'import.
    If g_CtrlModeUnique Then
        TiersNotesEditables = False
        Exit Function
    End If
    ' Jamais pour une opération de santé : Notes y encode une date de consultation
    ' lue ensuite par mod_ImportOFX (segment avant le premier ";").
    TiersNotesEditables = (g_CtrlSous(i) <> mod_VarGlobales.SOUS_CATEGORIE_SANTE)
End Function


' =====================================================================================
' AFFICHAGE DE L'OPÉRATION COURANTE (position g_CtrlPos)
' =====================================================================================
Private Sub AfficherOperation(ByVal ws As Worksheet)

    Dim i As Long
    Dim Source As String
    Dim evenementsAvant As Boolean
    Dim numErr As Long, descErr As String

    i = g_CtrlIndices(g_CtrlPos)
    Source = mod_DataStructure.CellText(g_CtrlOps(i, CTRL_OP_CATSOURCE))

    ' On désactive les événements pendant NOS écritures : sinon, écrire dans la cellule
    ' Categorie déclencherait Worksheet_Change comme si l'opérateur l'avait saisie.
    evenementsAvant = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Sortie

    ' --- Compteur et informations (lecture seule) ---
    ws.Range(CTRL_ADR_COMPTEUR).value = mod_InstallCommun.TexteCompteur(g_CtrlPos, g_CtrlNbAffiches)   ' format commun "Operation: x/y" (09/10/2026)
    ws.Range(CTRL_ADR_DATE).value = FormaterDate(g_CtrlOps(i, CTRL_OP_DATE))
    ws.Range(CTRL_ADR_MONTANT).value = Format(mod_DataStructure.ToDouble(g_CtrlOps(i, CTRL_OP_MONTANT)), "#,##0.00") & " " & ChrW(8364)
    ws.Range(CTRL_ADR_SOURCE).value = Source
    ' Libelle du champ adapte au mode (09/10/2026) : categorie bancaire a l'import, categorie
    ' d'origine en edition directe depuis l'ecran de recherche.
    If g_CtrlModeUnique Then
        ws.Cells(ws.Range(CTRL_ADR_SOURCE).Row, 2).value = mod_Display.FR("Cat{e2}gorie d'origine")
    Else
        ws.Cells(ws.Range(CTRL_ADR_SOURCE).Row, 2).value = mod_Display.FR("Cat{e2}gorie source (banque)")
    End If

    ' --- Tiers / Notes : AJOUT du 03/10/2026 (demande opérateur) -----------------------
    ' Ces deux champs sont désormais MODIFIABLES par l'opérateur, sauf dans les deux cas
    ' où une saisie libre risquerait d'altérer les données (voir TiersNotesEditables) :
    ' opération de santé (Notes encode une date de consultation lue par mod_ImportOFX)
    ' ou édition directe depuis l'écran de recherche (pas encore prévue pour ce cas).
    ' On affiche TOUJOURS g_CtrlTiers/g_CtrlLibelle (la valeur COURANTE, éventuellement
    ' déjà corrigée par l'opérateur sur une opération précédente), jamais g_CtrlOps
    ' (la valeur d'ORIGINE, figée), afin que la correction soit conservée avec
    ' Précédent/Suivant.
    ws.Range(CTRL_ADR_TIERS).value = g_CtrlTiers(i)
    ws.Range(CTRL_ADR_LIBELLE).value = g_CtrlLibelle(i)
    If TiersNotesEditables(i) Then
        ' Modifiable : fond jaune pâle, comme tous les champs modifiables des formulaires
        ' (charte commune, voir mod_InstallCommun.CoulFondSaisie - modifié le 09/10/2026).
        ws.Range(CTRL_ADR_TIERS & ":" & CTRL_ADR_LIBELLE).Interior.Color = mod_InstallCommun.CoulFondSaisie()
    Else
        ' Verrouillé : gris, comme toutes les cellules non modifiables des formulaires
        ' (charte commune, voir mod_InstallCommun.CoulFondLecture), notamment la colonne
        ' Notes verrouillée de l'écran de recherche.
        ws.Range(CTRL_ADR_TIERS & ":" & CTRL_ADR_LIBELLE).Interior.Color = mod_InstallCommun.CoulFondLecture()
    End If

    ' --- Zones de saisie : valeurs actuelles et listes déroulantes ---
    ws.Range(CTRL_ADR_CAT).value = g_CtrlCat(i)
    ws.Range(CTRL_ADR_SOUS).value = g_CtrlSous(i)

    PoserValidationCategorie ws.Range(CTRL_ADR_CAT)

    RemplirListeSousCategories ws, g_CtrlCat(i)

    ' --- Message d'aide adapté ---
    ' Ajout du 02/10/2026 : en mode "édition directe" (voir g_CtrlModeUnique), ni le message
    ' "à ranger" ni celui de "correspondance" n'ont de sens (il n'y a pas de catégorie
    ' source bancaire ici) : on affiche un message neutre dédié.
    If g_CtrlModeUnique Then
        ws.Range(CTRL_ADR_MESSAGE).value = mod_Display.FR("Cat{e2}gorie et sous-cat{e2}gorie actuelles de cette op{e2}ration. Vous pouvez les modifier ci-dessous.")
        ws.Range(CTRL_ADR_MESSAGE).Font.Color = RGB(90, 90, 90)
    ElseIf g_CtrlARanger(i) Then
        ws.Range(CTRL_ADR_MESSAGE).value = mod_Display.FR("Cat{e2}gorie source '") & Source & _
            mod_Display.FR("' absente du tableau de correspondance (ou vide) : choisissez la cat{e2}gorie {a2} affecter.")
        ws.Range(CTRL_ADR_MESSAGE).Font.Color = RGB(192, 80, 0)
    Else
        ws.Range(CTRL_ADR_MESSAGE).value = mod_Display.FR("Cat{e2}gorie propos{e2}e par la correspondance (source : '") & Source & _
            mod_Display.FR("'). Vous pouvez la modifier.")
        ws.Range(CTRL_ADR_MESSAGE).Font.Color = RGB(90, 90, 90)
    End If

Sortie:
    ' Ce bloc s'exécute TOUJOURS (fin normale ou erreur) : on réactive les événements.
    numErr = Err.Number
    descErr = Err.Description
    Application.EnableEvents = evenementsAvant
    If numErr <> 0 Then Err.Raise numErr, "AfficherOperation", descErr

End Sub


' =====================================================================================
' LISTE DÉROULANTE DES SOUS-CATÉGORIES (dépendante de la catégorie choisie)
' =====================================================================================
' Principe : on écrit dans la colonne technique masquée (Z) les sous-catégories de la
' catégorie choisie, puis on associe à la cellule Sous-categorie un menu pointant vers
' cette plage (via un nom dynamique). Une plage de cellules ne peut pas être mal
' interprétée, contrairement à un texte séparé par des virgules (voir phase 1).
Private Sub RemplirListeSousCategories(ByVal ws As Worksheet, ByVal categorie As String)

    Dim liste() As String
    Dim dejaVu As Object
    Dim nb As Long, r As Long
    Dim Sortie() As Variant
    Dim colLettre As String

    ' Nettoyage de l'ancienne liste.
    ws.Range(ws.Cells(2, CTRL_COL_AIDE), ws.Cells(CTRL_LIGNE_AIDE_MAX, CTRL_COL_AIDE)).ClearContents

    ' Recherche des sous-catégories de cette catégorie dans le référentiel.
    Set dejaVu = CreateObject("Scripting.Dictionary")
    dejaVu.CompareMode = 1        ' insensible à la casse (à régler AVANT le premier ajout)
    If g_CtrlRefN > 0 Then ReDim liste(1 To g_CtrlRefN)

    If categorie <> "" Then
        For r = 1 To g_CtrlRefN
            If StrComp(g_CtrlRefCat(r), categorie, vbTextCompare) = 0 And g_CtrlRefSous(r) <> "" Then
                If Not dejaVu.Exists(g_CtrlRefSous(r)) Then
                    dejaVu.Add g_CtrlRefSous(r), True
                    nb = nb + 1
                    liste(nb) = g_CtrlRefSous(r)
                End If
            End If
        Next r
    End If

    If nb > 0 Then
        TrierTextes liste, nb
        ReDim Sortie(1 To nb, 1 To 1)
        For r = 1 To nb
            Sortie(r, 1) = liste(r)
        Next r
        With ws.Cells(2, CTRL_COL_AIDE).Resize(nb, 1)
            .NumberFormat = "@"
            .Value2 = Sortie
        End With
    End If

    ' Nom dynamique : s'étend automatiquement au nombre de valeurs écrites.
    colLettre = Chr$(64 + CTRL_COL_AIDE)       ' 26 -> "Z"
    ThisWorkbook.Names.Add Name:=CTRL_NOM_LISTE_SOUS, _
        RefersTo:="=OFFSET(" & CTRL_NOM_FEUILLE & "!$" & colLettre & "$2,0,0," & _
                  "MAX(1,COUNTA(" & CTRL_NOM_FEUILLE & "!$" & colLettre & "$2:$" & colLettre & "$" & CTRL_LIGNE_AIDE_MAX & ")),1)"

    With ws.Range(CTRL_ADR_SOUS).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & CTRL_NOM_LISTE_SOUS
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
        .ErrorTitle = mod_Display.FR("Sous-cat{e2}gorie inconnue")
        .ErrorMessage = mod_Display.FR("Cette sous-cat{e2}gorie n'existe pas pour la cat{e2}gorie choisie.") & vbCrLf & _
                        mod_Display.FR("Choisissez-en une dans la liste, laissez le champ vide, ou utilisez le bouton") & _
                        " '+' " & mod_Display.FR("pour en cr{e2}er une nouvelle.")
    End With

End Sub

' Pose la validation "Categorie" (liste stricte et message d'erreur explicite) sur une
' cellule donnée. La validation est regroupée ici en une fonction unique : le message
' était recopié à trois endroits (AfficherOperation, ControleNouvelleCategorie,
' ControleVentiler) et avait fini par se désynchroniser (deux copies sur trois n'avaient
' plus de message, et la troisième annonçait encore une "phase ultérieure", devenue la
' phase 3 et déjà livrée). Cette fonction évite que cela ne se reproduise.
Private Sub PoserValidationCategorie(ByVal cellule As Range)
    With cellule.Validation
        .Delete
        ' "ListeCategories" : nom créé en phase 1 (RafraichirListesCategories).
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=ListeCategories"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
        .ErrorTitle = mod_Display.FR("Cat{e2}gorie inconnue")
        .ErrorMessage = mod_Display.FR("Choisissez une cat{e2}gorie dans la liste, ou laissez le champ vide.") & vbCrLf & _
                        mod_Display.FR("Pour cr{e2}er une nouvelle cat{e2}gorie, utilisez le bouton '+'.")
    End With
End Sub


' =====================================================================================
' ÉVÉNEMENTS DE LA FEUILLE (appelés par le code-behind de frm_ControleCategories)
' =====================================================================================
' Les événements d'une feuille doivent obligatoirement être écrits dans le module de CETTE
' feuille (voir CodeBehind_frm_ControleCategories.txt). Ils ne font qu'appeler les deux
' procédures ci-dessous : toute la logique reste ici, dans un module standard.

' Appelée à chaque modification d'une cellule de la feuille.
' Si la CATEGORIE change, l'ancienne sous-catégorie n'est plus valable : on l'efface et
' on recharge la liste des sous-catégories de la nouvelle catégorie.
Public Sub CtrlTraiterChangement(ByVal ws As Worksheet, ByVal Target As Range)

    Dim numErr As Long

    If Not g_CtrlEnCours Then Exit Sub                       ' formulaire fermé : on ignore
    If Target.Cells.count > 1 Then Exit Sub                  ' collage multiple : on ignore
    If Target.Address(False, False) <> CTRL_ADR_CAT Then Exit Sub   ' autre cellule : on ignore

    On Error GoTo Sortie
    Application.EnableEvents = False        ' évite que nos écritures relancent l'événement
    ws.Range(CTRL_ADR_SOUS).ClearContents
    RemplirListeSousCategories ws, mod_DataStructure.CellText(Target.value)

Sortie:
    numErr = Err.Number
    Application.EnableEvents = True
    If numErr <> 0 Then MsgBox mod_Display.FR("Erreur lors du changement de cat{e2}gorie : ") & Err.Description, vbExclamation

End Sub

' Appelée quand l'opérateur essaie de quitter la feuille (changement d'onglet...).
' Tant que le formulaire est ouvert, on le ramène dessus : c'est le "verrou modal".
Public Sub CtrlVerrouiller(ByVal ws As Worksheet)
    If g_CtrlVerrouActif Then ws.Activate
End Sub


' =====================================================================================
' ENREGISTREMENT DE L'OPÉRATION AFFICHÉE (contrôle de la saisie)
' =====================================================================================
' Lit les deux cellules de saisie, vérifie leur cohérence avec le référentiel
' et les mémorise. Renvoie False (avec un message) si la saisie est invalide :
' l'opérateur reste alors sur l'opération courante.
Private Function EnregistrerOperationAffichee(ByVal ws As Worksheet) As Boolean

    Dim i As Long
    Dim cat As String, sous As String

    i = g_CtrlIndices(g_CtrlPos)
    cat = mod_DataStructure.CellText(ws.Range(CTRL_ADR_CAT).value)
    sous = mod_DataStructure.CellText(ws.Range(CTRL_ADR_SOUS).value)

    ' Une sous-catégorie sans catégorie n'a pas de sens.
    If cat = "" And sous <> "" Then
        MsgBox mod_Display.FR("Une sous-cat{e2}gorie ne peut pas {e2}tre saisie sans cat{e2}gorie."), vbExclamation
        Exit Function
    End If

    If cat <> "" Then
        ' La catégorie doit exister dans le référentiel (un collage peut contourner le menu).
        If Not g_CtrlCats.Exists(cat) Then
            MsgBox mod_Display.FR("La cat{e2}gorie '") & cat & mod_Display.FR("' n'existe pas dans le tableau de correspondance."), vbExclamation
            Exit Function
        End If
        cat = CStr(g_CtrlCats(cat))         ' forme canonique (majuscules correctes)

        If sous <> "" Then
            If Not g_CtrlPaires.Exists(cat & vbTab & sous) Then
                MsgBox mod_Display.FR("La sous-cat{e2}gorie '") & sous & mod_Display.FR("' n'existe pas dans la cat{e2}gorie '") & cat & "'.", vbExclamation
                Exit Function
            End If
            sous = CStr(g_CtrlPaires(cat & vbTab & sous))
        End If
    End If

    g_CtrlCat(i) = cat
    g_CtrlSous(i) = sous

    ' AJOUT du 03/10/2026 (demande opérateur) : on ne lit les cellules Tiers/Notes que si
    ' cela est autorisé pour la sous-catégorie retenue juste au-dessus (TiersNotesEditables
    ' doit donc être appelée APRÈS g_CtrlSous(i) = sous, pas avant). Sinon, on ignore ce qui
    ' a pu être saisi : g_CtrlTiers/g_CtrlLibelle conservent leur valeur précédente (valeur
    ' d'origine ou correction faite plus tôt, lorsque la sous-catégorie le permettait).
    If TiersNotesEditables(i) Then
        g_CtrlTiers(i) = mod_DataStructure.CellText(ws.Range(CTRL_ADR_TIERS).value)
        g_CtrlLibelle(i) = mod_DataStructure.CellText(ws.Range(CTRL_ADR_LIBELLE).value)
    End If

    g_CtrlVu(i) = True
    EnregistrerOperationAffichee = True

End Function


' =====================================================================================
' ACTIONS DES BOUTONS (macros appelees par les boutons de la feuille)
' =====================================================================================

' --- Bouton "Précédent" ---
Public Sub ControleOperationPrecedente()
    Dim ws As Worksheet
    If Not g_CtrlEnCours Then Exit Sub
    On Error GoTo Erreur
    Set ws = ThisWorkbook.Worksheets(CTRL_NOM_FEUILLE)
    If Not EnregistrerOperationAffichee(ws) Then Exit Sub
    If g_CtrlPos > 1 Then
        g_CtrlPos = g_CtrlPos - 1
        AfficherOperation ws
    End If
    Exit Sub
Erreur:
    SignalerErreur "ControleOperationPrecedente"
End Sub

' --- Bouton "Suivant" ---
Public Sub ControleOperationSuivante()
    Dim ws As Worksheet
    If Not g_CtrlEnCours Then Exit Sub
    On Error GoTo Erreur
    Set ws = ThisWorkbook.Worksheets(CTRL_NOM_FEUILLE)
    If Not EnregistrerOperationAffichee(ws) Then Exit Sub
    If g_CtrlPos < g_CtrlNbAffiches Then
        g_CtrlPos = g_CtrlPos + 1
        AfficherOperation ws
    Else
        MsgBox mod_Display.FR("C'est la derni{e1}re op{e2}ration. Cliquez sur 'Enregistrer' pour valider."), vbInformation
    End If
    Exit Sub
Erreur:
    SignalerErreur "ControleOperationSuivante"
End Sub

' --- Bouton "Terminer et continuer" ---
Public Sub ControleTerminer()

    Dim ws As Worksheet
    Dim p As Long, i As Long
    Dim nbSans As Long, premiere As Long, nbNonVues As Long
    Dim reponse As VbMsgBoxResult

    If Not g_CtrlEnCours Then Exit Sub
    On Error GoTo Erreur
    Set ws = ThisWorkbook.Worksheets(CTRL_NOM_FEUILLE)

    ' 1. On enregistre d'abord l'opération actuellement affichée.
    If Not EnregistrerOperationAffichee(ws) Then Exit Sub

    ' 2. Bilan : combien d'opérations présentées sont encore SANS catégorie ?
    '    Et combien n'ont jamais été affichées ?
    For p = 1 To g_CtrlNbAffiches
        i = g_CtrlIndices(p)
        If g_CtrlCat(i) = "" Then
            nbSans = nbSans + 1
            If premiere = 0 Then premiere = p
        End If
        If Not g_CtrlVu(i) Then nbNonVues = nbNonVues + 1
    Next p

    ' 3. Aucune opération ne peut rester sans catégorie : on y ramène l'opérateur.
    ' MISE À JOUR du 03/10/2026 (demande opérateur) : la règle de gestion ne change pas.
    ' On ne peut pas terminer tant qu'il reste des opérations sans catégorie; ce message
    ' reste donc bloquant dans tous les cas (Exit Sub à la fin, que l'on clique sur OK
    ' ou Annuler). L'opérateur peut toutefois choisir la suite :
    '   - OK      : aller directement à la première opération à corriger
    '               (comportement d'origine, inchangé).
    '   - Annuler : rester sur l'écran actuel pour réfléchir ou corriger à son rythme.
    If nbSans > 0 Then
        reponse = MsgBox(nbSans & mod_Display.FR(" op{e2}ration(s) n'ont pas de cat{e2}gorie.") & vbCrLf & _
                          mod_Display.FR("Choisissez une cat{e2}gorie pour chacune avant de terminer."), _
                          vbOKCancel + vbExclamation)
        If reponse = vbOK Then
            g_CtrlPos = premiere
            AfficherOperation ws
        End If
        Exit Sub
    End If

    ' 4. Les opérations jamais affichées gardent la catégorie proposée. On demande confirmation.
    If nbNonVues > 0 Then
        reponse = MsgBox(nbNonVues & mod_Display.FR(" op{e2}ration(s) n'ont pas {e2}t{e2} affich{e2}es.") & vbCrLf & _
                         mod_Display.FR("Elles garderont la cat{e2}gorie propos{e2}e par la correspondance.") & vbCrLf & vbCrLf & _
                         "Terminer maintenant ?", vbYesNo + vbQuestion, mod_Display.FR("Confirmation"))
        If reponse = vbNo Then Exit Sub
    End If

    g_CtrlValide = True
    FermerFormulaire ws
    Exit Sub

Erreur:
    SignalerErreur "ControleTerminer"

End Sub

' --- Bouton "Annuler l'import" ---
Public Sub ControleAnnuler()

    Dim ws As Worksheet
    Dim reponse As VbMsgBoxResult

    If Not g_CtrlEnCours Then Exit Sub
    On Error GoTo Erreur

    reponse = MsgBox(mod_Display.FR("Annuler le contr{o2}le ? Aucune de vos modifications ne sera appliqu{e2}e."), _
                     vbYesNo + vbExclamation, mod_Display.FR("Confirmation"))
    If reponse = vbNo Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(CTRL_NOM_FEUILLE)
    g_CtrlValide = False
    FermerFormulaire ws
    Exit Sub

Erreur:
    SignalerErreur "ControleAnnuler"

End Sub


' --- Bouton "+ Nouvelle catégorie" (phase 3) ---
' Ouvre le formulaire de création (frm_NouvelleCategorie), prérempli avec la
' catégorie/sous-catégorie actuellement affichée. Si l'opérateur valide, la nouvelle
' paire est appliquée à l'opération en cours et le référentiel est rechargé afin que
' la suite du formulaire (Suivant, Terminer...) la reconnaisse immédiatement.
Public Sub ControleNouvelleCategorie()

    Dim ws As Worksheet
    Dim catActuelle As String, sousActuelle As String
    Dim catRes As String, sousRes As String
    Dim ok As Boolean

    If Not g_CtrlEnCours Then Exit Sub
    On Error GoTo Erreur
    Set ws = ThisWorkbook.Worksheets(CTRL_NOM_FEUILLE)

    catActuelle = mod_DataStructure.CellText(ws.Range(CTRL_ADR_CAT).value)
    sousActuelle = mod_DataStructure.CellText(ws.Range(CTRL_ADR_SOUS).value)

    ' On suspend le verrou d'ACTIVATION (pas g_CtrlEnCours; voir sa déclaration en
    ' tête de module) : sinon Worksheet_Deactivate ramènerait l'opérateur de force sur
    ' cette feuille dès qu'on tenterait d'afficher celle de création.
    g_CtrlVerrouActif = False

    ok = mod_NouvelleCategorie.OuvrirNouvelleCategorie(catActuelle, sousActuelle, catRes, sousRes)

    g_CtrlVerrouActif = True
    ws.Activate                          ' au cas où le focus ne serait pas déjà revenu ici

    If ok Then
        ' Le tableau TblCategories vient de changer (phase 1) : on recharge le
        ' référentiel en mémoire ET la liste déroulante des catégories.
        ChargerReferentiel
        mod_Categories.RafraichirListesCategories

        Application.EnableEvents = False
        PoserValidationCategorie ws.Range(CTRL_ADR_CAT)
        ws.Range(CTRL_ADR_CAT).value = catRes
        Application.EnableEvents = True

        ' RemplirListeSousCategories définit aussi la liste déroulante de la sous-catégorie;
        ' on l'appelle donc APRÈS avoir écrit cette valeur, pour éviter qu'elle soit effacée
        ' par un événement Worksheet_Change entre-temps.
        Application.EnableEvents = False
        ws.Range(CTRL_ADR_SOUS).value = sousRes
        RemplirListeSousCategories ws, catRes
        Application.EnableEvents = True
    End If
    Exit Sub

Erreur:
    g_CtrlVerrouActif = True
    Application.EnableEvents = True
    SignalerErreur "ControleNouvelleCategorie"

End Sub


' --- Bouton "Ventiler" (phase 4) ---
' Ouvre le formulaire de ventilation (frm_Ventilation) pour l'opération actuellement
' affichée. Si l'opérateur valide une ventilation complète (somme exactement égale au
' montant de l'opération), celle-ci est marquée "Ventile" dans le formulaire de contrôle.
' Comme un changement de catégorie normal, elle ne sera écrite dans TblOperations qu'en
' phase 5 (import). Les lignes de ventilation, elles, sont déjà enregistrées dans
' TblVentilations dès la validation (voir mod_Ventilation). Cette nouvelle table n'étant
' utilisée par aucune autre partie du classeur, son écriture immédiate ne présente pas
' de risque pour les données existantes.
Public Sub ControleVentiler()

    Dim ws As Worksheet
    Dim i As Long
    Dim idTransaction As String
    Dim ok As Boolean
    Dim ventilationSupprimee As Boolean

    If Not g_CtrlEnCours Then Exit Sub
    On Error GoTo Erreur
    Set ws = ThisWorkbook.Worksheets(CTRL_NOM_FEUILLE)

    i = g_CtrlIndices(g_CtrlPos)
    idTransaction = mod_DataStructure.CellText(g_CtrlOps(i, CTRL_OP_ID))

    If idTransaction = "" Then
        MsgBox mod_Display.FR("Cette op{e2}ration n'a pas d'identifiant (ID_Transaction) : impossible de la ventiler."), vbExclamation
        Exit Sub
    End If

    ' Ajout du 01/10/2026 (point 4 : annuler une ventilation) : on mémorise la catégorie/
    ' sous-catégorie ACTUELLE avant de risquer de l'écraser par "Ventile", mais UNE SEULE
    ' FOIS par opération. Si l'opérateur revient sur une opération DÉJÀ ventilée (bouton
    ' "Précédent" dans ce même import, puis de nouveau "Ventiler"), g_CtrlCat(i) contient
    ' déjà "Ventile" : il ne faut pas écraser la vraie valeur d'origine, mémorisée la
    ' première fois, par "Ventile", sinon elle serait définitivement perdue.
    If g_CtrlCat(i) <> CategorieVentile Then
        g_CtrlCatAvantVen(i) = g_CtrlCat(i)
        g_CtrlSousAvantVen(i) = g_CtrlSous(i)
    End If

    g_CtrlVerrouActif = False

    ok = mod_Ventilation.OuvrirVentilation(idTransaction, g_CtrlOps(i, CTRL_OP_DATE), _
            mod_DataStructure.CellText(g_CtrlOps(i, CTRL_OP_TIERS)), _
            mod_DataStructure.CellText(g_CtrlOps(i, CTRL_OP_LIBELLE)), _
            mod_DataStructure.ToDouble(g_CtrlOps(i, CTRL_OP_MONTANT)), _
            g_CtrlCat(i), g_CtrlSous(i), ventilationSupprimee)

    g_CtrlVerrouActif = True
    ws.Activate

    If ok Then
        ' La catégorie "Ventile" doit exister dans le référentiel pour qu'EnregistrerOperationAffichee
        ' (Suivant/Précédent/Terminer) l'accepte.
        mod_Categories.AjouterCategoriePersonnalisee CategorieVentile, ""
        ChargerReferentiel
        mod_Categories.RafraichirListesCategories

        g_CtrlCat(i) = CategorieVentile
        g_CtrlSous(i) = ""
        g_CtrlVu(i) = True

        Application.EnableEvents = False
        PoserValidationCategorie ws.Range(CTRL_ADR_CAT)
        ws.Range(CTRL_ADR_CAT).value = CategorieVentile
        ws.Range(CTRL_ADR_SOUS).value = ""
        RemplirListeSousCategories ws, CategorieVentile
        ws.Range(CTRL_ADR_MESSAGE).value = mod_Display.FR("Op{e2}ration ventil{e2}e : la cat{e2}gorie '") & CategorieVentile & _
                                            mod_Display.FR("' a {e2}t{e2} affect{e2}e. Le d{e2}tail est enregistr{e2} dans TblVentilations.")
        ws.Range(CTRL_ADR_MESSAGE).Font.Color = RGB(31, 120, 60)
        Application.EnableEvents = True

    ElseIf ventilationSupprimee Then
        ' Ajout du 01/10/2026 (point 4) : l'opérateur a supprimé une ventilation déjà
        ' existante (bouton "Supprimer cette ventilation" de frm_Ventilation), même au
        ' milieu de l'import (bouton "Précédent", puis "Ventiler"). On restaure la
        ' catégorie/sous-catégorie mémorisée avant la première ventilation, puis on
        ' l'efface : il n'y aura rien à restaurer tant qu'une nouvelle ventilation
        ' n'aura pas été créée pour cette opération.
        g_CtrlCat(i) = g_CtrlCatAvantVen(i)
        g_CtrlSous(i) = g_CtrlSousAvantVen(i)
        g_CtrlCatAvantVen(i) = ""
        g_CtrlSousAvantVen(i) = ""
        g_CtrlVu(i) = True

        Application.EnableEvents = False
        PoserValidationCategorie ws.Range(CTRL_ADR_CAT)
        ws.Range(CTRL_ADR_CAT).value = g_CtrlCat(i)
        RemplirListeSousCategories ws, g_CtrlCat(i)
        ws.Range(CTRL_ADR_SOUS).value = g_CtrlSous(i)
        ws.Range(CTRL_ADR_MESSAGE).value = mod_Display.FR("Ventilation supprim{e2}e : l'op{e2}ration a repris sa cat{e2}gorie d'origine.")
        ws.Range(CTRL_ADR_MESSAGE).Font.Color = RGB(31, 120, 60)
        Application.EnableEvents = True
    End If
    Exit Sub

Erreur:
    g_CtrlVerrouActif = True
    Application.EnableEvents = True
    SignalerErreur "ControleVentiler"

End Sub

' Catégorie réservée qui marque une opération ventilée (voir TblVentilations pour le
' détail des lignes). Orthographe : UN SEUL "l" (corrigée avec l'opérateur).
Private Function CategorieVentile() As String
    CategorieVentile = "Ventil" & ChrW(233)
End Function


' =====================================================================================
' FERMETURE DU FORMULAIRE
' =====================================================================================
' IMPORTANT : g_CtrlEnCours doit passer à False AVANT de masquer la feuille, sinon le
' verrou (CtrlVerrouiller) ramènerait immédiatement l'opérateur dessus.
Private Sub FermerFormulaire(ByVal ws As Worksheet)

    If ws Is Nothing Then Exit Sub

    g_CtrlEnCours = False                     ' libère la boucle d'attente
    g_CtrlVerrouActif = False

    ws.Range(ws.Cells(2, CTRL_COL_AIDE), ws.Cells(CTRL_LIGNE_AIDE_MAX, CTRL_COL_AIDE)).ClearContents
    ws.Visible = xlSheetVeryHidden

    On Error Resume Next                      ' la feuille d'origine a pu être renommée/supprimée
    ThisWorkbook.Worksheets(g_CtrlNomFeuillePrec).Activate
    On Error GoTo 0

End Sub

' Message d'erreur commun aux boutons (le formulaire reste ouvert).
Private Sub SignalerErreur(ByVal nomProcedure As String)
    Application.EnableEvents = True
    MsgBox mod_Display.FR("Erreur dans ") & nomProcedure & " : " & Err.Number & " - " & Err.Description, vbCritical
End Sub


' =====================================================================================
' LECTURE DU RÉFÉRENTIEL (tableau TblCategories, feuille Param)
' =====================================================================================
' Remplit g_CtrlMap, g_CtrlPaires, g_CtrlCats et g_CtrlRefCat/g_CtrlRefSous.
' Renvoie False si le tableau est introuvable ou vide.
' PUBLIC (depuis la phase 3) : ControleNouvelleCategorie le recharge après la création
' d'une catégorie, afin que le reste du formulaire (Suivant, Terminer...) la reconnaisse
' immédiatement, sans attendre la réouverture complète du formulaire de contrôle.
Public Function ChargerReferentiel() As Boolean

    Dim wsParam As Worksheet
    Dim tblCat As ListObject
    Dim src As Variant, cat As Variant, sous As Variant
    Dim r As Long
    Dim s As String, c As String, u As String

    Set wsParam = FeuilleSansErreur(NOM_FEUILLE_PARAM)
    If wsParam Is Nothing Then Exit Function

    On Error Resume Next
    Set tblCat = wsParam.ListObjects(NOM_TABLE_CATEGORIES)
    On Error GoTo 0
    If tblCat Is Nothing Then Exit Function
    If tblCat.DataBodyRange Is Nothing Then Exit Function

    src = LireColonne(tblCat.ListColumns("CategorieSource").DataBodyRange)
    cat = LireColonne(tblCat.ListColumns("Categorie").DataBodyRange)
    sous = LireColonne(tblCat.ListColumns("SousCategorie").DataBodyRange)

    Set g_CtrlMap = CreerDictionnaire()
    Set g_CtrlPaires = CreerDictionnaire()
    Set g_CtrlCats = CreerDictionnaire()
    ReDim g_CtrlRefCat(1 To UBound(cat, 1))
    ReDim g_CtrlRefSous(1 To UBound(cat, 1))
    g_CtrlRefN = 0

    For r = 1 To UBound(cat, 1)
        s = mod_DataStructure.CellText(src(r, 1))
        c = mod_DataStructure.CellText(cat(r, 1))
        u = mod_DataStructure.CellText(sous(r, 1))

        If c <> "" Then                                   ' une ligne sans catégorie est ignorée
            g_CtrlRefN = g_CtrlRefN + 1
            g_CtrlRefCat(g_CtrlRefN) = c
            g_CtrlRefSous(g_CtrlRefN) = u

            If Not g_CtrlCats.Exists(c) Then g_CtrlCats.Add c, c
            If u <> "" Then
                If Not g_CtrlPaires.Exists(c & vbTab & u) Then g_CtrlPaires.Add c & vbTab & u, u
            End If
            ' Une source n'a qu'un rangement : en cas de doublon, la première ligne l'emporte.
            If s <> "" Then
                If Not g_CtrlMap.Exists(s) Then g_CtrlMap.Add s, c & vbTab & u
            End If
        End If
    Next r

    ChargerReferentiel = (g_CtrlRefN > 0)

End Function


' =====================================================================================
' MACRO DE TEST (sans risque : n'écrit RIEN dans vos données)
' =====================================================================================
' Prend les premières opérations de TblOperations (les plus récentes si la table est
' triée par date décroissante, comme après un import), simule leur import, ouvre le
' formulaire, puis affiche les choix de l'opérateur.
' Ctrl+G, taper : TesterControleCategories
Public Sub TesterControleCategories()

    Dim tblOps As ListObject
    Dim rgDate As Range, rgTiers As Range, rgNotes As Range, rgMontant As Range, rgCat As Range, rgId As Range
    Dim vDate As Variant, vTiers As Variant, vNotes As Variant, vMontant As Variant, vCat As Variant, vId As Variant
    Dim ops() As Variant
    Dim indices() As Long
    Dim catF() As String, sousF() As String
    Dim catAvantF() As String, sousAvantF() As String
    ' AJOUT du 03/10/2026 : nouveaux paramètres de sortie de ControlerCategories (Tiers/Notes
    ' modifiables à l'import); ce test ne les affiche pas non plus, même logique que
    ' catAvantF/sousAvantF ci-dessus.
    Dim tiersF() As String, libelleF() As String
    Dim saisie As String
    Dim nbDemande As Long, nbTrouves As Long, nbLignesLues As Long
    Dim i As Long, k As Long
    Dim recap As String, ligne As String
    Dim reponse As VbMsgBoxResult

    Set tblOps = mod_DonneesTable.GetOperationsTable()
    If tblOps Is Nothing Then Exit Sub
    If tblOps.DataBodyRange Is Nothing Then Exit Sub

    saisie = InputBox(mod_Display.FR("Combien d'op{e2}rations voulez-vous tester (1 {a2} 30) ?"), "Test du formulaire", "8")
    If saisie = "" Then Exit Sub
    If Not IsNumeric(saisie) Then
        MsgBox mod_Display.FR("Veuillez saisir un nombre."), vbExclamation
        Exit Sub
    End If
    nbDemande = CLng(saisie)
    If nbDemande < 1 Then nbDemande = 1
    If nbDemande > 30 Then nbDemande = 30

    ' On ne lit que les 300 premières lignes de la table : largement suffisant pour un test.
    nbLignesLues = tblOps.ListRows.count
    If nbLignesLues > 300 Then nbLignesLues = 300

    Set rgDate = plageColonne(tblOps, "Date_Comptable", nbLignesLues)
    Set rgTiers = plageColonne(tblOps, "Tiers", nbLignesLues)
    Set rgNotes = plageColonne(tblOps, "Notes", nbLignesLues)
    Set rgMontant = plageColonne(tblOps, "Montant", nbLignesLues)
    Set rgCat = plageColonne(tblOps, "Categorie", nbLignesLues)
    Set rgId = plageColonne(tblOps, "ID_Transaction", nbLignesLues)
    If rgDate Is Nothing Or rgTiers Is Nothing Or rgNotes Is Nothing Or rgMontant Is Nothing Or rgCat Is Nothing Or rgId Is Nothing Then Exit Sub

    vDate = LireColonne(rgDate)
    vTiers = LireColonne(rgTiers)
    vNotes = LireColonne(rgNotes)
    vMontant = LireColonne(rgMontant)
    vCat = LireColonne(rgCat)
    vId = LireColonne(rgId)

    ' Étape 1 : on repère les lignes qui ont une catégorie (pour tester la correspondance).
    ReDim indices(1 To nbDemande)
    For i = 1 To UBound(vCat, 1)
        If nbTrouves < nbDemande Then
            If mod_DataStructure.CellText(vCat(i, 1)) <> "" Then
                nbTrouves = nbTrouves + 1
                indices(nbTrouves) = i
            End If
        End If
    Next i
    If nbTrouves = 0 Then
        MsgBox mod_Display.FR("Aucune op{e2}ration avec cat{e2}gorie trouv{e2}e parmi les premi{e1}res lignes."), vbInformation
        Exit Sub
    End If

    ' Étape 2 : on construit le tableau "ops" exactement comme le fera l'import.
    ReDim ops(1 To nbTrouves, 1 To CTRL_OP_NBCOL)
    For k = 1 To nbTrouves
        i = indices(k)
        ops(k, CTRL_OP_DATE) = vDate(i, 1)
        ops(k, CTRL_OP_TIERS) = vTiers(i, 1)
        ops(k, CTRL_OP_LIBELLE) = vNotes(i, 1)
        ops(k, CTRL_OP_MONTANT) = vMontant(i, 1)
        ops(k, CTRL_OP_CATSOURCE) = vCat(i, 1)
        ' Préfixe "TEST-" volontaire : si vous testez le bouton "Ventiler", la ligne
        ' créée dans TblVentilations sera immédiatement identifiable comme donnée de test
        ' (facile à filtrer et supprimer) et ne sera jamais confondue avec un véritable
        ' ID_Transaction de TblOperations.
        ops(k, CTRL_OP_ID) = "TEST-" & mod_DataStructure.CellText(vId(i, 1))
    Next k

    ' Option pour tester le cas d'une "catégorie source inconnue".
    reponse = MsgBox(mod_Display.FR("Simuler une cat{e2}gorie source INCONNUE sur la 1{e1}re op{e2}ration (pour tester ce cas) ?"), _
                     vbYesNo + vbQuestion, "Test du formulaire")
    If reponse = vbYes Then ops(1, CTRL_OP_CATSOURCE) = mod_Display.FR("Cat{e2}gorie source inconnue (test)")

    ' --- Appel du contrôle : c'est EXACTEMENT ce que fera l'import en phase 5 ---
    ' (catAvantF/sousAvantF : paramètres de sortie ajoutés le 01/10/2026, point 4;
    ' tiersF/libelleF : ajoutés le 03/10/2026. Ce test ne les affiche pas, mais doit les
    ' fournir comme le ferait n'importe quel appelant.)
    If Not ControlerCategories(ops, nbTrouves, catF, sousF, catAvantF, sousAvantF, tiersF, libelleF) Then
        MsgBox mod_Display.FR("Test termin{e2} : contr{o2}le annul{e2} ou impossible. Rien n'a {e2}t{e2} modifi{e2}."), vbInformation
        Exit Sub
    End If

    ' --- Récapitulatif (affiché à l'écran ET dans la fenêtre Exécution, Ctrl+G) ---
    For k = 1 To nbTrouves
        ligne = FormaterDate(ops(k, CTRL_OP_DATE)) & " | " & _
                Format(mod_DataStructure.ToDouble(ops(k, CTRL_OP_MONTANT)), "0.00") & " | " & _
                mod_DataStructure.CellText(ops(k, CTRL_OP_CATSOURCE)) & "  ->  " & catF(k)
        If sousF(k) <> "" Then ligne = ligne & " / " & sousF(k)
        Debug.Print ligne
        If k <= 12 Then recap = recap & ligne & vbCrLf
    Next k
    If nbTrouves > 12 Then recap = recap & "..." & vbCrLf

    MsgBox mod_Display.FR("R{e2}sultat du test (RIEN n'a {e2}t{e2} modifi{e2} dans vos donn{e2}es) :") & vbCrLf & vbCrLf & recap, _
           vbInformation, "Test du formulaire"

End Sub


' =====================================================================================
' OUTILS INTERNES
' =====================================================================================

' Dictionnaire insensible à la casse (réglé avant tout ajout, sinon Excel refuse).
Private Function CreerDictionnaire() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set CreerDictionnaire = d
End Function

Private Function FeuilleSansErreur(ByVal nomFeuille As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nomFeuille)
    On Error GoTo 0
    Set FeuilleSansErreur = ws
End Function

' Renvoie les nb premières lignes d'une colonne de la table (ou Nothing si elle est introuvable).
Private Function plageColonne(ByVal t As ListObject, ByVal nom As String, ByVal nb As Long) As Range
    Dim lc As ListColumn
    On Error Resume Next
    Set lc = t.ListColumns(nom)
    On Error GoTo 0
    If lc Is Nothing Then
        MsgBox mod_Display.FR("Colonne introuvable dans TblOperations : ") & nom, vbExclamation
    Else
        Set plageColonne = lc.DataBodyRange.Resize(nb, 1)
    End If
End Function

' Lit une plage en mémoire et renvoie TOUJOURS un tableau à deux dimensions.
' (Piège VBA : pour UNE seule cellule, .Value2 renvoie une valeur simple, pas un tableau.)
Private Function LireColonne(ByVal plage As Range) As Variant
    Dim t() As Variant
    If plage.Cells.count = 1 Then
        ReDim t(1 To 1, 1 To 1)
        t(1, 1) = plage.Value2
        LireColonne = t
    Else
        LireColonne = plage.Value2
    End If
End Function

' Tri alphabétique (insensible à la casse) des n premiers éléments d'un tableau de textes.
Private Sub TrierTextes(ByRef t() As String, ByVal n As Long)
    Dim i As Long, j As Long
    Dim cle As String
    For i = 2 To n
        cle = t(i)
        j = i - 1
        Do While j >= 1
            If StrComp(t(j), cle, vbTextCompare) <= 0 Then Exit Do
            t(j + 1) = t(j)
            j = j - 1
        Loop
        t(j + 1) = cle
    Next i
End Sub

' Convertit une date en texte "jj/mm/aaaa", qu'elle soit de type Date ou un numéro de série Excel.
Private Function FormaterDate(ByVal v As Variant) As String
    If IsEmpty(v) Then
        FormaterDate = ""
    ElseIf IsDate(v) Then
        FormaterDate = Format(CDate(v), "dd/mm/yyyy")
    ElseIf IsNumeric(v) Then
        If CDbl(v) > 0 Then
            FormaterDate = Format(CDate(CDbl(v)), "dd/mm/yyyy")
        End If
    Else
        FormaterDate = mod_DataStructure.CellText(v)
    End If
End Function