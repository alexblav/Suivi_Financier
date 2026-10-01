Option Explicit

' =====================================================================================
' MODULE : mod_ControleCategories
'
' PHASE 2 du chantier "Categorie / Sous-categorie / Ventilation" - PARTIE 2/2
'
' ROLE (a lire en premier, meme si vous debutez) :
'   Ce module contient toute la LOGIQUE du controle des categories des operations
'   importees. Il s'appuie sur :
'     - le tableau de correspondance TblCategories (Phase 1, feuille Param),
'     - la feuille-formulaire frm_ControleCategories (mod_InstallControleCategories).
'
'   DEROULEMENT (fonction ControlerCategories) :
'     1. Pour chaque operation, on cherche sa categorie SOURCE (celle de la banque)
'        dans TblCategories -> on en deduit Categorie + Sous-categorie proposees.
'     2. On demande a l'operateur :
'          "Les categories ont-elles ete correctement renseignees dans la source ?"
'            OUI -> on garde les propositions telles quelles (voir remarque ci-dessous).
'            NON -> le formulaire presente les operations une par une.
'            ANNULER -> on abandonne, rien n'est applique.
'     3. Dans le formulaire, l'operateur peut changer la Categorie et la
'        Sous-categorie de chaque operation, avec Precedent / Suivant.
'     4. "Terminer et continuer" renvoie les valeurs finales a l'appelant.
'
'   REMARQUE IMPORTANTE (choix de conception a valider) :
'     Meme si l'operateur repond OUI, les operations dont la categorie source est
'     VIDE ou ABSENTE du tableau de correspondance sont quand meme presentees
'     dans le formulaire : sans cela, on ne saurait pas dans quelle categorie les ranger.
'     Ce comportement se coupe en passant CTRL_FORCER_NON_RANGEES a False (plus bas).
'
'   CE MODULE N'ECRIT RIEN dans TblOperations. Il RENVOIE des resultats ; c'est
'   l'appelant (plus tard : l'import, Phase 5) qui decidera de les ecrire.
'   Pour le tester sans risque, utilisez la macro TesterControleCategories.
'
' A PROPOS DES ACCENTS : fichier 100% ASCII, accents fabriques par la fonction mod_Display.FR().
' =====================================================================================

' --- Positions des colonnes du tableau "ops" que l'appelant doit fournir ---------------
' ops(i, CTRL_OP_DATE)      = date de l'operation (Date ou nombre serie Excel)
' ops(i, CTRL_OP_TIERS)     = tiers (texte)
' ops(i, CTRL_OP_LIBELLE)   = libelle / notes (texte)
' ops(i, CTRL_OP_MONTANT)   = montant (nombre, negatif pour une depense)
' ops(i, CTRL_OP_CATSOURCE) = categorie envoyee par la banque (texte, peut etre vide)
' ops(i, CTRL_OP_ID)        = ID_Transaction de l'operation (texte). NOUVEAU (Phase 4) :
'                             necessaire pour relier une ventilation (TblVentilations) a
'                             l'operation d'origine dans TblOperations.
Public Const CTRL_OP_DATE As Long = 1
Public Const CTRL_OP_TIERS As Long = 2
Public Const CTRL_OP_LIBELLE As Long = 3
Public Const CTRL_OP_MONTANT As Long = 4
Public Const CTRL_OP_CATSOURCE As Long = 5
Public Const CTRL_OP_ID As Long = 6
Public Const CTRL_OP_NBCOL As Long = 6

' --- Reglage : presenter aussi les operations "non rangees" apres un OUI ? -------------
Private Const CTRL_FORCER_NON_RANGEES As Boolean = True

' --- Nom (au niveau du classeur) de la liste des sous-categories du formulaire ---------
Private Const CTRL_NOM_LISTE_SOUS As String = "ListeSousCatControle"
Private Const CTRL_LIGNE_AIDE_MAX As Long = 300   ' derniere ligne de la zone technique

' --- Etat du formulaire ------------------------------------------------------------------
' Vrai tant que le formulaire est ouvert. Public car lu par le verrou de la feuille.
Public g_CtrlEnCours As Boolean

' --- Verrou d'ACTIVATION, distinct de g_CtrlEnCours (Phase 3) -----------------------
' g_CtrlEnCours pilote la boucle d'attente (DoEvents) : elle NE DOIT PAS repasser a
' False tant que le formulaire de controle n'est pas termine.
' g_CtrlVerrouActif pilote uniquement le "reflexe" qui ramene de force l'operateur sur
' la feuille (Worksheet_Deactivate). Sans cette distinction, ouvrir le formulaire de
' creation de categorie (bouton "+") serait impossible : des qu'on quitterait la feuille
' de controle pour afficher celui de creation, Worksheet_Deactivate nous y ramenerait
' aussitot de force, empechant le second formulaire de s'afficher.
' ControleNouvelleCategorie desactive ce verrou juste le temps d'ouvrir le formulaire de
' creation, puis le reactive des qu'il se referme.
Public g_CtrlVerrouActif As Boolean

Private g_CtrlValide As Boolean          ' Vrai si l'operateur a valide (Terminer)
Private g_CtrlNomFeuillePrec As String   ' feuille active avant l'ouverture (pour y revenir)

Private g_CtrlOps As Variant             ' copie des operations fournies
Private g_CtrlNbOps As Long
Private g_CtrlCat() As String            ' categorie courante de chaque operation
Private g_CtrlSous() As String           ' sous-categorie courante de chaque operation
Private g_CtrlVu() As Boolean            ' operation deja affichee / enregistree
Private g_CtrlARanger() As Boolean       ' correspondance introuvable (a ranger a la main)

Private g_CtrlIndices() As Long          ' operations presentees : position -> numero d'operation
Private g_CtrlNbAffiches As Long
Private g_CtrlPos As Long                ' position affichee actuellement (1..g_CtrlNbAffiches)

' --- Referentiel lu dans TblCategories (recharge a chaque appel) -----------------------
Private g_CtrlMap As Object              ' source (minuscules) -> "categorie" & Tab & "sous-categorie"
Private g_CtrlPaires As Object           ' "categorie" & Tab & "sous-categorie" -> sous-categorie
Private g_CtrlCats As Object             ' categorie -> categorie (forme canonique)
Private g_CtrlRefCat() As String
Private g_CtrlRefSous() As String
Private g_CtrlRefN As Long


' =====================================================================================
' FONCTION PRINCIPALE
' =====================================================================================
' Parametres :
'   ops        : tableau a 2 dimensions (1 To nbOps, 1 To CTRL_OP_NBCOL), voir plus haut
'                (transmis par valeur : l'original n'est jamais modifie)
'   nbOps      : nombre d'operations a controler
'   catFinale  : (en sortie) categorie retenue pour chaque operation, indice 1..nbOps
'   sousFinale : (en sortie) sous-categorie retenue pour chaque operation
' Renvoie True si l'operateur a valide (ou s'il n'y avait rien a controler),
'         False s'il a annule ou si un pre-requis manque (un message est alors affiche).
Public Function ControlerCategories(ByVal ops As Variant, ByVal nbOps As Long, _
                                    ByRef catFinale() As String, ByRef sousFinale() As String) As Boolean

    Dim ws As Worksheet
    Dim i As Long, nbARanger As Long
    Dim Source As String
    Dim morceaux() As String
    Dim reponse As VbMsgBoxResult

    On Error GoTo Erreur

    ControlerCategories = False
    If nbOps <= 0 Then
        ControlerCategories = True     ' rien a controler : on laisse continuer
        Exit Function
    End If

    ' --- Pre-requis 1 : la feuille-formulaire doit exister -------------------------------
    Set ws = FeuilleSansErreur(CTRL_NOM_FEUILLE)
    If ws Is Nothing Then
        MsgBox mod_Display.FR("La feuille '") & CTRL_NOM_FEUILLE & mod_Display.FR("' est introuvable.") & vbCrLf & _
               mod_Display.FR("Ex{e2}cutez d'abord la macro CreerFeuilleControleCategories."), vbExclamation
        Exit Function
    End If

    ' --- Pre-requis 2 : le tableau de correspondance doit exister et contenir des lignes ---
    If Not ChargerReferentiel() Then
        MsgBox mod_Display.FR("Le tableau de correspondance TblCategories est introuvable ou vide.") & vbCrLf & _
               mod_Display.FR("Ex{e2}cutez d'abord la macro PreparerPhase1Categories."), vbExclamation
        Exit Function
    End If

    ' --- On travaille sur des COPIES : ops n'est jamais modifie ---------------------------
    g_CtrlOps = ops
    g_CtrlNbOps = nbOps
    ReDim g_CtrlCat(1 To nbOps)
    ReDim g_CtrlSous(1 To nbOps)
    ReDim g_CtrlVu(1 To nbOps)
    ReDim g_CtrlARanger(1 To nbOps)

    ' --- ETAPE 1 : application automatique de la correspondance --------------------------
    For i = 1 To nbOps
        Source = mod_DataStructure.CellText(g_CtrlOps(i, CTRL_OP_CATSOURCE))
        If Source <> "" And g_CtrlMap.Exists(Source) Then
            ' Item = "categorie" & Tab & "sous-categorie" : on le redecoupe.
            morceaux = Split(CStr(g_CtrlMap(Source)), vbTab)
            g_CtrlCat(i) = morceaux(0)
            g_CtrlSous(i) = morceaux(1)
        Else
            g_CtrlCat(i) = ""
            g_CtrlSous(i) = ""
            g_CtrlARanger(i) = True
            nbARanger = nbARanger + 1
        End If
    Next i

    ' --- ETAPE 2 : la question a l'operateur ------------------------------------------------
    reponse = MsgBox(nbOps & mod_Display.FR(" op{e2}ration(s) import{e2}e(s).") & vbCrLf & vbCrLf & _
                     mod_Display.FR("Les cat{e2}gories de ces op{e2}rations ont-elles {e2}t{e2} correctement renseign{e2}es dans la source (la banque) ?") & _
                     vbCrLf & vbCrLf & _
                     mod_Display.FR("OUI : la correspondance est appliqu{e2}e automatiquement.") & vbCrLf & _
                     mod_Display.FR("NON : les op{e2}rations vous sont pr{e2}sent{e2}es une par une pour les v{e2}rifier.") & vbCrLf & _
                     mod_Display.FR("ANNULER : abandon, rien n'est appliqu{e2}."), _
                     vbYesNoCancel + vbQuestion, mod_Display.FR("Contr{o1}le des cat{e2}gories"))

    If reponse = vbCancel Then Exit Function      ' False : abandon

    ' --- Quelles operations presenter dans le formulaire ? ---------------------------------
    ReDim g_CtrlIndices(1 To nbOps)
    g_CtrlNbAffiches = 0

    If reponse = vbNo Then
        ' Lecture 1 - "NON" : toutes les operations sont presentees.
        For i = 1 To nbOps
            g_CtrlNbAffiches = g_CtrlNbAffiches + 1
            g_CtrlIndices(g_CtrlNbAffiches) = i
        Next i
    Else
        ' "OUI" : on n'affiche que les operations impossibles a ranger automatiquement.
        If CTRL_FORCER_NON_RANGEES And nbARanger > 0 Then
            MsgBox nbARanger & mod_Display.FR(" op{e2}ration(s) ont une cat{e2}gorie source vide ou absente du tableau de correspondance.") & _
                   vbCrLf & mod_Display.FR("Elles vont vous {e2}tre pr{e2}sent{e2}es pour {e2}tre rang{e2}es."), _
                   vbInformation, mod_Display.FR("Contr{o1}le des cat{e2}gories")
            For i = 1 To nbOps
                If g_CtrlARanger(i) Then
                    g_CtrlNbAffiches = g_CtrlNbAffiches + 1
                    g_CtrlIndices(g_CtrlNbAffiches) = i
                End If
            Next i
        End If
    End If

    ' --- ETAPE 3 : le formulaire (uniquement s'il y a quelque chose a presenter) ------------
    If g_CtrlNbAffiches = 0 Then
        g_CtrlValide = True
    Else
        OuvrirFormulaireEtAttendre ws
    End If

    ' --- ETAPE 4 : restitution des resultats -------------------------------------------------
    ControlerCategories = g_CtrlValide
    If g_CtrlValide Then
        ReDim catFinale(1 To nbOps)
        ReDim sousFinale(1 To nbOps)
        For i = 1 To nbOps
            catFinale(i) = g_CtrlCat(i)
            sousFinale(i) = g_CtrlSous(i)
        Next i
    End If
    Exit Function

Erreur:
    ' Erreur imprevue : on remet Excel en etat normal avant d'informer l'operateur.
    Application.EnableEvents = True
    g_CtrlEnCours = False
    MsgBox mod_Display.FR("Erreur inattendue dans le contr{o1}le des cat{e2}gories :") & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical
    On Error Resume Next
    FermerFormulaire ws
    ControlerCategories = False

End Function


' =====================================================================================
' OUVERTURE DU FORMULAIRE ET ATTENTE ("modal" : le code reste bloque ici)
' =====================================================================================
Private Sub OuvrirFormulaireEtAttendre(ByVal ws As Worksheet)

    ' On retient la feuille actuelle pour y revenir a la fermeture.
    g_CtrlNomFeuillePrec = ActiveSheet.Name

    ' On met a jour la liste des categories (Phase 1) : elle alimente le menu deroulant.
    mod_Categories.RafraichirListesCategories

    g_CtrlValide = False
    g_CtrlPos = 1
    AfficherOperation ws           ' on remplit AVANT d'afficher (pas de flash a l'ecran)

    ws.Visible = xlSheetVisible
    ws.Activate
    ws.Range(CTRL_ADR_CAT).Select

    ' --- Verrou "modal" : la boucle ne se termine que lorsque g_CtrlEnCours repasse a
    ' False (bouton Terminer ou Annuler). DoEvents laisse Excel traiter les clics.
    '
    ' CORRECTIF (constate par l'operateur le 2026-09-25) : sans cette ligne, les listes
    ' deroulantes du formulaire restaient inertes (aucune reaction au changement de
    ' categorie), car Application.EnableEvents se retrouvait a Faux a ce stade -- sans
    ' qu'aucune erreur ne le signale. Toute la boucle qui suit repose entierement sur
    ' Worksheet_Change (voir CtrlTraiterChangement) : on GARANTIT donc ici que les
    ' evenements sont actifs, plutot que de faire confiance a l'etat laisse par le reste
    ' du classeur (dont l'origine exacte du probleme n'a pas ete formellement identifiee).
    Application.EnableEvents = True
    g_CtrlEnCours = True
    g_CtrlVerrouActif = True
    Do While g_CtrlEnCours
        DoEvents
    Loop

End Sub


' =====================================================================================
' AFFICHAGE DE L'OPERATION COURANTE (position g_CtrlPos)
' =====================================================================================
Private Sub AfficherOperation(ByVal ws As Worksheet)

    Dim i As Long
    Dim Source As String
    Dim evenementsAvant As Boolean
    Dim numErr As Long, descErr As String

    i = g_CtrlIndices(g_CtrlPos)
    Source = mod_DataStructure.CellText(g_CtrlOps(i, CTRL_OP_CATSOURCE))

    ' On coupe les evenements pendant NOS ecritures : sans cela, ecrire dans la cellule
    ' Categorie declencherait Worksheet_Change comme si l'operateur l'avait saisie.
    evenementsAvant = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Sortie

    ' --- Compteur et informations (lecture seule) ---
    ws.Range(CTRL_ADR_COMPTEUR).value = mod_Display.FR("Op{e2}ration ") & g_CtrlPos & " / " & g_CtrlNbAffiches
    ws.Range(CTRL_ADR_DATE).value = FormaterDate(g_CtrlOps(i, CTRL_OP_DATE))
    ws.Range(CTRL_ADR_TIERS).value = mod_DataStructure.CellText(g_CtrlOps(i, CTRL_OP_TIERS))
    ws.Range(CTRL_ADR_LIBELLE).value = mod_DataStructure.CellText(g_CtrlOps(i, CTRL_OP_LIBELLE))
    ws.Range(CTRL_ADR_MONTANT).value = Format(mod_DataStructure.ToDouble(g_CtrlOps(i, CTRL_OP_MONTANT)), "#,##0.00") & " " & ChrW(8364)
    ws.Range(CTRL_ADR_SOURCE).value = Source

    ' --- Zones de saisie : valeurs actuelles + listes deroulantes ---
    ws.Range(CTRL_ADR_CAT).value = g_CtrlCat(i)
    ws.Range(CTRL_ADR_SOUS).value = g_CtrlSous(i)

    PoserValidationCategorie ws.Range(CTRL_ADR_CAT)

    RemplirListeSousCategories ws, g_CtrlCat(i)

    ' --- Message d'aide adapte ---
    If g_CtrlARanger(i) Then
        ws.Range(CTRL_ADR_MESSAGE).value = mod_Display.FR("Cat{e2}gorie source '") & Source & _
            mod_Display.FR("' absente du tableau de correspondance (ou vide) : choisissez la cat{e2}gorie {a2} affecter.")
        ws.Range(CTRL_ADR_MESSAGE).Font.Color = RGB(192, 80, 0)
    Else
        ws.Range(CTRL_ADR_MESSAGE).value = mod_Display.FR("Cat{e2}gorie propos{e2}e par la correspondance (source : '") & Source & _
            mod_Display.FR("'). Vous pouvez la modifier.")
        ws.Range(CTRL_ADR_MESSAGE).Font.Color = RGB(90, 90, 90)
    End If

Sortie:
    ' Ce bloc s'execute TOUJOURS (fin normale ou erreur) : on rend les evenements.
    numErr = Err.Number
    descErr = Err.Description
    Application.EnableEvents = evenementsAvant
    If numErr <> 0 Then Err.Raise numErr, "AfficherOperation", descErr

End Sub


' =====================================================================================
' LISTE DEROULANTE DES SOUS-CATEGORIES (dependante de la categorie choisie)
' =====================================================================================
' Principe : on ecrit dans la colonne technique cachee (Z) les sous-categories de la
' categorie choisie, puis on donne a la cellule Sous-categorie un menu qui pointe vers
' cette plage (via un nom dynamique). Une plage de cellules ne peut pas etre mal
' interpretee, contrairement a un texte separe par des virgules (voir Phase 1).
Private Sub RemplirListeSousCategories(ByVal ws As Worksheet, ByVal categorie As String)

    Dim liste() As String
    Dim dejaVu As Object
    Dim nb As Long, r As Long
    Dim Sortie() As Variant
    Dim colLettre As String

    ' Nettoyage de l'ancienne liste
    ws.Range(ws.Cells(2, CTRL_COL_AIDE), ws.Cells(CTRL_LIGNE_AIDE_MAX, CTRL_COL_AIDE)).ClearContents

    ' Recherche des sous-categories de cette categorie dans le referentiel
    Set dejaVu = CreateObject("Scripting.Dictionary")
    dejaVu.CompareMode = 1        ' insensible a la casse (a regler AVANT le 1er ajout)
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

    ' Nom dynamique : s'etend automatiquement au nombre de valeurs ecrites.
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
                        " '+ " & mod_Display.FR("Nouvelle cat{e2}gorie") & "' " & mod_Display.FR("pour en cr{e2}er une nouvelle.")
    End With

End Sub

' Pose la validation "Categorie" (liste stricte + message d'erreur explicite) sur une
' cellule donnee. Regroupee ici en une seule fonction : ce message etait recopie a
' 3 endroits differents (AfficherOperation, ControleNouvelleCategorie, ControleVentiler)
' et avait fini par se desynchroniser (2 des 3 copies n'avaient plus de message du tout,
' et la 3e annoncait encore une "phase ulterieure" devenue entre-temps la Phase 3, deja
' livree). Passer par cette fonction unique evite que cela ne se reproduise.
Private Sub PoserValidationCategorie(ByVal cellule As Range)
    With cellule.Validation
        .Delete
        ' "ListeCategories" : nom cree en Phase 1 (RafraichirListesCategories).
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=ListeCategories"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
        .ErrorTitle = mod_Display.FR("Cat{e2}gorie inconnue")
        .ErrorMessage = mod_Display.FR("Choisissez une cat{e2}gorie dans la liste, ou laissez le champ vide.") & vbCrLf & _
                        mod_Display.FR("Pour cr{e2}er une nouvelle cat{e2}gorie, utilisez le bouton '+ Nouvelle cat{e2}gorie'.")
    End With
End Sub


' =====================================================================================
' EVENEMENTS DE LA FEUILLE (appeles par le code-behind de frm_ControleCategories)
' =====================================================================================
' Les evenements d'une feuille doivent obligatoirement etre ecrits dans le module de CETTE
' feuille (voir CodeBehind_frm_ControleCategories.txt). Ils ne font qu'appeler les deux
' procedures ci-dessous : toute la logique reste ici, dans un module normal.

' Appelee a chaque modification d'une cellule de la feuille.
' Si la CATEGORIE change : l'ancienne sous-categorie n'est plus valable -> on l'efface et
' on recharge la liste des sous-categories de la nouvelle categorie.
Public Sub CtrlTraiterChangement(ByVal ws As Worksheet, ByVal Target As Range)

    Dim numErr As Long

    If Not g_CtrlEnCours Then Exit Sub                       ' formulaire ferme : on ignore
    If Target.Cells.count > 1 Then Exit Sub                  ' collage multiple : on ignore
    If Target.Address(False, False) <> CTRL_ADR_CAT Then Exit Sub   ' autre cellule : on ignore

    On Error GoTo Sortie
    Application.EnableEvents = False        ' evite que nos ecritures relancent l'evenement
    ws.Range(CTRL_ADR_SOUS).ClearContents
    RemplirListeSousCategories ws, mod_DataStructure.CellText(Target.value)

Sortie:
    numErr = Err.Number
    Application.EnableEvents = True
    If numErr <> 0 Then MsgBox mod_Display.FR("Erreur lors du changement de cat{e2}gorie : ") & Err.Description, vbExclamation

End Sub

' Appelee quand l'operateur essaie de quitter la feuille (changement d'onglet...).
' Tant que le formulaire est ouvert, on le ramene dessus : c'est le "verrou modal".
Public Sub CtrlVerrouiller(ByVal ws As Worksheet)
    If g_CtrlVerrouActif Then ws.Activate
End Sub


' =====================================================================================
' ENREGISTREMENT DE L'OPERATION AFFICHEE (controle de la saisie)
' =====================================================================================
' Lit les deux cellules de saisie, verifie qu'elles sont coherentes avec le referentiel
' et les memorise. Renvoie False (avec un message) si la saisie est invalide : l'operateur
' reste alors sur l'operation courante.
Private Function EnregistrerOperationAffichee(ByVal ws As Worksheet) As Boolean

    Dim i As Long
    Dim cat As String, sous As String

    i = g_CtrlIndices(g_CtrlPos)
    cat = mod_DataStructure.CellText(ws.Range(CTRL_ADR_CAT).value)
    sous = mod_DataStructure.CellText(ws.Range(CTRL_ADR_SOUS).value)

    ' Une sous-categorie sans categorie n'a pas de sens.
    If cat = "" And sous <> "" Then
        MsgBox mod_Display.FR("Une sous-cat{e2}gorie ne peut pas {e2}tre saisie sans cat{e2}gorie."), vbExclamation
        Exit Function
    End If

    If cat <> "" Then
        ' La categorie doit exister dans le referentiel (un collage peut contourner le menu).
        If Not g_CtrlCats.Exists(cat) Then
            MsgBox mod_Display.FR("La cat{e2}gorie '") & cat & mod_Display.FR("' n'existe pas dans le tableau de correspondance."), vbExclamation
            Exit Function
        End If
        cat = CStr(g_CtrlCats(cat))         ' forme canonique (bonnes majuscules)

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
    g_CtrlVu(i) = True
    EnregistrerOperationAffichee = True

End Function


' =====================================================================================
' ACTIONS DES BOUTONS (macros appelees par les boutons de la feuille)
' =====================================================================================

' --- Bouton "Precedent" ---
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
        MsgBox mod_Display.FR("C'est la derni{e1}re op{e2}ration. Cliquez sur 'Terminer et continuer' pour valider."), vbInformation
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

    ' 1. On enregistre d'abord l'operation actuellement affichee.
    If Not EnregistrerOperationAffichee(ws) Then Exit Sub

    ' 2. Bilan : combien d'operations presentees sont encore SANS categorie ?
    '    Et combien n'ont jamais ete affichees ?
    For p = 1 To g_CtrlNbAffiches
        i = g_CtrlIndices(p)
        If g_CtrlCat(i) = "" Then
            nbSans = nbSans + 1
            If premiere = 0 Then premiere = p
        End If
        If Not g_CtrlVu(i) Then nbNonVues = nbNonVues + 1
    Next p

    ' 3. Aucune operation ne peut rester sans categorie : on renvoie l'operateur dessus.
    If nbSans > 0 Then
        MsgBox nbSans & mod_Display.FR(" op{e2}ration(s) n'ont pas de cat{e2}gorie.") & vbCrLf & _
               mod_Display.FR("Choisissez une cat{e2}gorie pour chacune avant de terminer."), vbExclamation
        g_CtrlPos = premiere
        AfficherOperation ws
        Exit Sub
    End If

    ' 4. Operations jamais affichees : elles gardent la categorie proposee. On demande confirmation.
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

    reponse = MsgBox(mod_Display.FR("Annuler le contr{o1}le ? Aucune de vos modifications ne sera appliqu{e2}e."), _
                     vbYesNo + vbExclamation, mod_Display.FR("Confirmation"))
    If reponse = vbNo Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(CTRL_NOM_FEUILLE)
    g_CtrlValide = False
    FermerFormulaire ws
    Exit Sub

Erreur:
    SignalerErreur "ControleAnnuler"

End Sub


' --- Bouton "+ Nouvelle categorie" (Phase 3) ---
' Ouvre le formulaire de creation (frm_NouvelleCategorie), pre-rempli avec la
' categorie/sous-categorie actuellement affichees. Si l'operateur valide, la nouvelle
' paire est appliquee a l'operation en cours, et le referentiel est recharge pour que
' la suite du formulaire (Suivant, Terminer...) la reconnaisse immediatement.
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

    ' On suspend le verrou d'ACTIVATION (pas g_CtrlEnCours, voir sa declaration en
    ' tete de module) : sans cela, Worksheet_Deactivate ramenerait de force l'operateur
    ' sur cette feuille des qu'on tenterait d'afficher celle de creation.
    g_CtrlVerrouActif = False

    ok = mod_NouvelleCategorie.OuvrirNouvelleCategorie(catActuelle, sousActuelle, catRes, sousRes)

    g_CtrlVerrouActif = True
    ws.Activate                          ' au cas ou le focus ne serait pas deja revenu ici

    If ok Then
        ' Le tableau TblCategories vient de changer (Phase 1) : on recharge le
        ' referentiel en memoire ET la liste deroulante des categories.
        ChargerReferentiel
        mod_Categories.RafraichirListesCategories

        Application.EnableEvents = False
        PoserValidationCategorie ws.Range(CTRL_ADR_CAT)
        ws.Range(CTRL_ADR_CAT).value = catRes
        Application.EnableEvents = True

        ' RemplirListeSousCategories pose aussi la liste deroulante de la sous-categorie ;
        ' on l'appelle donc APRES avoir ecrit la sous-categorie, pour ne pas la voir
        ' effacee par un declenchement de Worksheet_Change entre-temps.
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


' --- Bouton "Ventiler" (Phase 4) ---
' Ouvre le formulaire de ventilation (frm_Ventilation) pour l'operation actuellement
' affichee. Si l'operateur valide une ventilation complete (somme exactement egale au
' montant de l'operation), l'operation est marquee "Ventile" dans le formulaire de
' controle -- exactement comme un changement de categorie normal, elle ne sera ecrite
' dans TblOperations qu'a la Phase 5 (import). Les LIGNES de ventilation, elles, sont
' deja enregistrees dans TblVentilations des la validation (voir mod_Ventilation) : cette
' table est nouvelle et n'est utilisee par aucune autre partie du classeur, l'ecrire
' immediatement ne presente donc aucun risque pour vos donnees existantes.
Public Sub ControleVentiler()

    Dim ws As Worksheet
    Dim i As Long
    Dim idTransaction As String
    Dim ok As Boolean

    If Not g_CtrlEnCours Then Exit Sub
    On Error GoTo Erreur
    Set ws = ThisWorkbook.Worksheets(CTRL_NOM_FEUILLE)

    i = g_CtrlIndices(g_CtrlPos)
    idTransaction = mod_DataStructure.CellText(g_CtrlOps(i, CTRL_OP_ID))

    If idTransaction = "" Then
        MsgBox mod_Display.FR("Cette op{e2}ration n'a pas d'identifiant (ID_Transaction) : impossible de la ventiler."), vbExclamation
        Exit Sub
    End If

    g_CtrlVerrouActif = False

    ok = mod_Ventilation.OuvrirVentilation(idTransaction, g_CtrlOps(i, CTRL_OP_DATE), _
            mod_DataStructure.CellText(g_CtrlOps(i, CTRL_OP_TIERS)), _
            mod_DataStructure.CellText(g_CtrlOps(i, CTRL_OP_LIBELLE)), _
            mod_DataStructure.ToDouble(g_CtrlOps(i, CTRL_OP_MONTANT)), _
            g_CtrlCat(i), g_CtrlSous(i))

    g_CtrlVerrouActif = True
    ws.Activate

    If ok Then
        ' La categorie "Ventile" doit exister dans le referentiel pour que
        ' EnregistrerOperationAffichee (Suivant/Precedent/Terminer) l'accepte.
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
    End If
    Exit Sub

Erreur:
    g_CtrlVerrouActif = True
    Application.EnableEvents = True
    SignalerErreur "ControleVentiler"

End Sub

' Categorie reservee qui marque une operation ventilee (voir la table TblVentilations
' pour le detail des lignes). Orthographe : UN SEUL "l" (corrige avec l'operateur).
Private Function CategorieVentile() As String
    CategorieVentile = "Ventil" & ChrW(233)
End Function


' =====================================================================================
' FERMETURE DU FORMULAIRE
' =====================================================================================
' IMPORTANT : g_CtrlEnCours doit passer a False AVANT de masquer la feuille, sinon le
' verrou (CtrlVerrouiller) ramenerait immediatement l'operateur dessus.
Private Sub FermerFormulaire(ByVal ws As Worksheet)

    If ws Is Nothing Then Exit Sub

    g_CtrlEnCours = False                     ' libere la boucle d'attente
    g_CtrlVerrouActif = False

    ws.Range(ws.Cells(2, CTRL_COL_AIDE), ws.Cells(CTRL_LIGNE_AIDE_MAX, CTRL_COL_AIDE)).ClearContents
    ws.Visible = xlSheetVeryHidden

    On Error Resume Next                      ' la feuille d'origine a pu etre renommee/supprimee
    ThisWorkbook.Worksheets(g_CtrlNomFeuillePrec).Activate
    On Error GoTo 0

End Sub

' Message d'erreur commun aux boutons (le formulaire reste ouvert).
Private Sub SignalerErreur(ByVal nomProcedure As String)
    Application.EnableEvents = True
    MsgBox mod_Display.FR("Erreur dans ") & nomProcedure & " : " & Err.Number & " - " & Err.Description, vbCritical
End Sub


' =====================================================================================
' LECTURE DU REFERENTIEL (tableau TblCategories, feuille Param)
' =====================================================================================
' Remplit g_CtrlMap, g_CtrlPaires, g_CtrlCats et g_CtrlRefCat/g_CtrlRefSous.
' Renvoie False si le tableau est introuvable ou vide.
' PUBLIC (depuis la Phase 3) : ControleNouvelleCategorie la rappelle apres la creation
' d'une categorie, pour que le reste du formulaire (Suivant, Terminer...) la reconnaisse
' aussitot sans attendre une reouverture complete du formulaire de controle.
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

        If c <> "" Then                                   ' une ligne sans categorie est ignoree
            g_CtrlRefN = g_CtrlRefN + 1
            g_CtrlRefCat(g_CtrlRefN) = c
            g_CtrlRefSous(g_CtrlRefN) = u

            If Not g_CtrlCats.Exists(c) Then g_CtrlCats.Add c, c
            If u <> "" Then
                If Not g_CtrlPaires.Exists(c & vbTab & u) Then g_CtrlPaires.Add c & vbTab & u, u
            End If
            ' Une source n'a qu'un rangement : en cas de doublon, la 1re ligne l'emporte.
            If s <> "" Then
                If Not g_CtrlMap.Exists(s) Then g_CtrlMap.Add s, c & vbTab & u
            End If
        End If
    Next r

    ChargerReferentiel = (g_CtrlRefN > 0)

End Function


' =====================================================================================
' MACRO DE TEST (sans risque : n'ecrit RIEN dans vos donnees)
' =====================================================================================
' Prend les operations les plus hautes de TblOperations (les plus recentes si la table est
' triee par date decroissante, comme apres un import), fait comme si elles venaient d'etre
' importees, ouvre le formulaire, puis affiche ce que l'operateur a choisi.
' Ctrl+G, taper :  TesterControleCategories
Public Sub TesterControleCategories()

    Dim tblOps As ListObject
    Dim rgDate As Range, rgTiers As Range, rgNotes As Range, rgMontant As Range, rgCat As Range, rgId As Range
    Dim vDate As Variant, vTiers As Variant, vNotes As Variant, vMontant As Variant, vCat As Variant, vId As Variant
    Dim ops() As Variant
    Dim indices() As Long
    Dim catF() As String, sousF() As String
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

    ' On ne lit que les 300 premieres lignes de la table : largement suffisant pour un test.
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

    ' Passe 1 : on repere les lignes qui ont une categorie (pour exercer la correspondance).
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

    ' Passe 2 : on construit le tableau "ops" exactement comme l'import le fera plus tard.
    ReDim ops(1 To nbTrouves, 1 To CTRL_OP_NBCOL)
    For k = 1 To nbTrouves
        i = indices(k)
        ops(k, CTRL_OP_DATE) = vDate(i, 1)
        ops(k, CTRL_OP_TIERS) = vTiers(i, 1)
        ops(k, CTRL_OP_LIBELLE) = vNotes(i, 1)
        ops(k, CTRL_OP_MONTANT) = vMontant(i, 1)
        ops(k, CTRL_OP_CATSOURCE) = vCat(i, 1)
        ' PREFIXE "TEST-" volontaire : si vous testez le bouton "Ventiler", la ligne
        ' creee dans TblVentilations sera ainsi immediatement reconnaissable comme une
        ' donnee de test (facile a filtrer et supprimer), jamais confondue avec un vrai
        ' ID_Transaction de TblOperations.
        ops(k, CTRL_OP_ID) = "TEST-" & mod_DataStructure.CellText(vId(i, 1))
    Next k

    ' Option pour tester le cas "categorie source inconnue".
    reponse = MsgBox(mod_Display.FR("Simuler une cat{e2}gorie source INCONNUE sur la 1{e1}re op{e2}ration (pour tester ce cas) ?"), _
                     vbYesNo + vbQuestion, "Test du formulaire")
    If reponse = vbYes Then ops(1, CTRL_OP_CATSOURCE) = mod_Display.FR("Cat{e2}gorie source inconnue (test)")

    ' --- Appel du controle : c'est EXACTEMENT ce que fera l'import en Phase 5 ---
    If Not ControlerCategories(ops, nbTrouves, catF, sousF) Then
        MsgBox mod_Display.FR("Test termin{e2} : contr{o1}le annul{e2} ou impossible. Rien n'a {e2}t{e2} modifi{e2}."), vbInformation
        Exit Sub
    End If

    ' --- Recapitulatif (affiche a l'ecran ET dans la fenetre Execution, Ctrl+G) ---
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

' Dictionnaire insensible a la casse (regle avant tout ajout, sinon Excel refuse).
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

' Renvoie les nb premieres lignes d'une colonne de la table (ou Nothing si introuvable).
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

' Lit une plage en memoire et renvoie TOUJOURS un tableau a 2 dimensions.
' (Piege VBA : pour UNE seule cellule, .Value2 renvoie une valeur simple, pas un tableau.)
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

' Tri alphabetique (insensible a la casse) des n premiers elements d'un tableau de textes.
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

' Met une date en texte "jj/mm/aaaa", qu'elle soit de type Date ou nombre serie Excel.
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

