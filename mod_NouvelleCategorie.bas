Option Explicit

' =====================================================================================
' MODULE : mod_NouvelleCategorie
'
' PHASE 3 du chantier « Catégorie / Sous-catégorie / Ventilation » - PARTIE 2/2
'
' RÔLE (à lire en premier, même si vous débutez) :
'   Ce module contient toute la logique du formulaire de création d'une catégorie ou
'   d'une sous-catégorie. Il s'appuie sur :
'     - la feuille-formulaire frm_NouvelleCategorie (mod_InstallNouvelleCategorie),
'     - le tableau TblCategories, via les fonctions de mod_Categories (Phase 1).
'
'   FONCTIONNEMENT "PAR-DESSUS" UN AUTRE FORMULAIRE :
'   Ce formulaire est conçu pour s'ouvrir À PARTIR d'un autre formulaire-feuille déjà
'   ouvert (le contrôle des catégories, phase 2), via son bouton "+ Nouvelle catégorie".
'   La fonction OuvrirNouvelleCategorie() ci-dessous gère ce cas : elle attend la
'   fermeture de CE formulaire (Valider ou Annuler) et renvoie le résultat à l'appelant,
'   sans jamais toucher aux données elles-mêmes : c'est l'appelant qui décide quoi faire
'   du résultat (ici : mod_ControleCategories.ControleNouvelleCategorie).
'
'   Le formulaire peut aussi être testé seul (voir TesterNouvelleCategorie), sans passer
'   par le formulaire de contrôle.
'
' À PROPOS DES ACCENTS : les textes affichés sont construits par mod_Display.FR(); les
' commentaires sont encodés en UTF-8.
' =====================================================================================

' --- État du formulaire ------------------------------------------------------------------
Public g_NcEnCours As Boolean        ' Vrai tant que le formulaire est ouvert (boucle d'attente)
Public g_NcVerrouActif As Boolean    ' Vrai tant que Worksheet_Deactivate doit réactiver la feuille

Private g_NcValide As Boolean        ' Vrai si l'opérateur a cliqué sur "Valider"
Private g_NcCatResultat As String
Private g_NcSousResultat As String
Private g_NcNomFeuillePrec As String


' =====================================================================================
' FONCTION PRINCIPALE
' =====================================================================================
' Paramètres :
'   catInitiale, sousInitiale : valeurs proposées au départ (éventuellement vides).
'   catResultat, sousResultat : en sortie, uniquement si la fonction renvoie True,
'                                catégorie et sous-catégorie validées par l'opérateur
'                                et déjà enregistrées dans TblCategories.
' Renvoie True si l'opérateur a cliqué sur "Valider", False s'il a annulé.
Public Function OuvrirNouvelleCategorie(ByVal catInitiale As String, ByVal sousInitiale As String, _
                                        ByRef catResultat As String, ByRef sousResultat As String) As Boolean

    Dim ws As Worksheet

    On Error GoTo Erreur

    Set ws = FeuilleSansErreur(NC_NOM_FEUILLE)
    If ws Is Nothing Then
        MsgBox mod_Display.FR("La feuille '") & NC_NOM_FEUILLE & mod_Display.FR("' est introuvable.") & vbCrLf & _
               mod_Display.FR("Ex{e2}cutez d'abord la macro CreerFeuilleNouvelleCategorie."), vbExclamation
        Exit Function
    End If

    g_NcNomFeuillePrec = ActiveSheet.Name

    g_NcValide = False
    RemplirChamps ws, catInitiale, sousInitiale

    ws.Visible = xlSheetVisible
    ws.Activate
    ws.Range(NC_ADR_CAT).Select

    ' CORRECTIF PRÉVENTIF (voir mod_ControleCategories.OuvrirFormulaireEtAttendre pour
    ' le détail) : ce formulaire repose lui aussi sur Worksheet_Change pendant sa boucle
    ' d'attente; on garantit donc que les événements sont actifs à ce stade.
    Application.EnableEvents = True
    g_NcEnCours = True
    g_NcVerrouActif = True
    Do While g_NcEnCours
        DoEvents
    Loop

    ' À ce stade, NcValider ou NcAnnuler a déjà masqué la feuille (voir FermerFeuilleNc).
    ' On revient sur la feuille qui était active avant l'ouverture.
    On Error Resume Next
    ThisWorkbook.Worksheets(g_NcNomFeuillePrec).Activate
    On Error GoTo 0

    OuvrirNouvelleCategorie = g_NcValide
    If g_NcValide Then
        catResultat = g_NcCatResultat
        sousResultat = g_NcSousResultat
    End If
    Exit Function

Erreur:
    Application.EnableEvents = True
    g_NcEnCours = False
    g_NcVerrouActif = False
    MsgBox mod_Display.FR("Erreur inattendue dans la cr{e2}ation de cat{e2}gorie :") & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical
    OuvrirNouvelleCategorie = False

End Function


' =====================================================================================
' Remplit les champs avec les valeurs de départ et configure les deux listes déroulantes.
' =====================================================================================
Private Sub RemplirChamps(ByVal ws As Worksheet, ByVal catInitiale As String, ByVal sousInitiale As String)

    Dim evenementsAvant As Boolean
    Dim numErr As Long

    evenementsAvant = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Sortie

    ws.Range(NC_ADR_CAT).value = catInitiale
    ws.Range(NC_ADR_SOUS).value = sousInitiale
    ws.Range(NC_ADR_MESSAGE).value = ""

    ' Liste "Categorie" : le nom "ListeCategories" est créé en phase 1
    ' (mod_Categories.RafraichirListesCategories).
    '
    ' PAS D'ALERTE ICI (ni Warning ni Stop) : saisir une valeur HORS liste est
    ' précisément le but de cet écran (créer une catégorie). Interrompre l'opérateur
    ' CHAQUE fois qu'il saisit un nouveau nom serait contre-productif, d'autant que le
    ' clic sur "Valider" déclenche déjà un contrôle plus pertinent (garde-fou
    ' orthographique : "Vouliez-vous dire...?" si le texte ressemble à une catégorie
    ' existante). Une alerte supplémentaire obligeait l'opérateur à confirmer DEUX FOIS
    ' (clic sur Oui dans l'alerte Excel, puis nouveau clic sur Valider); ce comportement
    ' a été corrigé à la suite des essais.
    ' Le menu déroulant (InCellDropdown) reste actif : une catégorie existante peut
    ' toujours être choisie en un clic; seule l'alerte de saisie libre est désactivée.
    With ws.Range(NC_ADR_CAT).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Formula1:="=ListeCategories"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = False    ' AlertStyle n'a alors aucun effet : c'est voulu (voir l'explication ci-dessus).
    End With

    RemplirListeSousCatNc ws, catInitiale

Sortie:
    numErr = Err.Number
    Application.EnableEvents = evenementsAvant
    If numErr <> 0 Then Err.Raise numErr, "RemplirChamps", Err.Description

End Sub


' =====================================================================================
' Liste déroulante des sous-catégories DÉJÀ CONNUES de la catégorie sélectionnée
' =====================================================================================
' Même principe qu'en phase 2 (mod_ControleCategories.RemplirListeSousCategories) :
' les valeurs sont écrites dans une colonne technique masquée et le menu pointe vers
' cette plage via un nom dynamique. AlertStyle Warning : une NOUVELLE sous-catégorie
' peut être saisie librement.
Private Sub RemplirListeSousCatNc(ByVal ws As Worksheet, ByVal categorie As String)

    Dim liste() As String
    Dim Sortie() As Variant
    Dim nb As Long, r As Long
    Dim colLettre As String

    ws.Range(ws.Cells(2, NC_COL_AIDE), ws.Cells(NC_LIGNE_AIDE_MAX, NC_COL_AIDE)).ClearContents

    liste = mod_Categories.ObtenirSousCategories(categorie)

    If EstTableauAlloue(liste) Then
        nb = UBound(liste) - LBound(liste) + 1
        ReDim Sortie(1 To nb, 1 To 1)
        For r = 1 To nb
            Sortie(r, 1) = liste(LBound(liste) + r - 1)
        Next r
        With ws.Cells(2, NC_COL_AIDE).Resize(nb, 1)
            .NumberFormat = "@"
            .Value2 = Sortie
        End With
    End If

    colLettre = Chr$(64 + NC_COL_AIDE)     ' 26 -> "Z"
    ThisWorkbook.Names.Add Name:="ListeSousCatNouvelle", _
        RefersTo:="=OFFSET(" & NC_NOM_FEUILLE & "!$" & colLettre & "$2,0,0," & _
                  "MAX(1,COUNTA(" & NC_NOM_FEUILLE & "!$" & colLettre & "$2:$" & colLettre & "$" & NC_LIGNE_AIDE_MAX & ")),1)"

    ' Même principe que pour Categorie ci-dessus : aucune alerte de saisie ici;
    ' le menu déroulant reste disponible et le garde-fou orthographique intelligent
    ' (Valider) reste le seul contrôle. Il n'est donc pas nécessaire de confirmer deux fois.
    With ws.Range(NC_ADR_SOUS).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Formula1:="=ListeSousCatNouvelle"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = False
    End With

End Sub


' =====================================================================================
' ÉVÉNEMENTS DE LA FEUILLE (appelés par le code-behind de frm_NouvelleCategorie)
' =====================================================================================
' Comme en phase 2, les événements d'une feuille doivent se trouver dans le module de CETTE
' feuille (voir CodeBehind_frm_NouvelleCategorie.txt). Ils ne font qu'appeler les deux
' procédures ci-dessous : toute la logique reste dans ce module standard.

' Si la CATEGORIE change, l'ancienne sous-catégorie n'est plus pertinente : on l'efface
' et on recharge la liste des sous-catégories de la nouvelle catégorie.
Public Sub NcTraiterChangement(ByVal ws As Worksheet, ByVal Target As Range)

    Dim numErr As Long

    If Not g_NcEnCours Then Exit Sub
    If Target.Cells.count > 1 Then Exit Sub
    If Target.Address(False, False) <> NC_ADR_CAT Then Exit Sub

    On Error GoTo Sortie
    Application.EnableEvents = False
    ws.Range(NC_ADR_SOUS).ClearContents
    RemplirListeSousCatNc ws, mod_DataStructure.CellText(Target.value)

Sortie:
    numErr = Err.Number
    Application.EnableEvents = True
    If numErr <> 0 Then MsgBox mod_Display.FR("Erreur lors du changement de cat{e2}gorie : ") & Err.Description, vbExclamation

End Sub

' Verrou "modal" : tant que le formulaire est ouvert, on empêche l'opérateur de le quitter
' en changeant d'onglet (Worksheet_Deactivate le ramène de force sur cette feuille).
Public Sub NcVerrouiller(ByVal ws As Worksheet)
    If g_NcVerrouActif Then ws.Activate
End Sub


' =====================================================================================
' ACTIONS DES BOUTONS
' =====================================================================================

' --- Bouton "Valider" ---
Public Sub NcValider()

    Dim ws As Worksheet
    Dim cat As String, sous As String
    Dim listeCat() As String, listeSous() As String
    Dim procheCat As String, procheSous As String
    Dim reponse As VbMsgBoxResult

    If Not g_NcEnCours Then Exit Sub
    On Error GoTo Erreur
    Set ws = ThisWorkbook.Worksheets(NC_NOM_FEUILLE)

    cat = mod_DataStructure.CellText(ws.Range(NC_ADR_CAT).value)
    sous = mod_DataStructure.CellText(ws.Range(NC_ADR_SOUS).value)

    If cat = "" Then
        ws.Range(NC_ADR_MESSAGE).value = mod_Display.FR("La cat{e2}gorie est obligatoire.")
        Exit Sub
    End If

    ' --- Garde-fou orthographique (niveau 2) : catégorie -----------------------------
    ' Le niveau 1 (accents, espaces, majuscules) est déjà géré silencieusement par
    ' AjouterCategoriePersonnalisee. Ici, on détecte une simple FAUTE DE FRAPPE
    ' (jusqu'à deux lettres de différence, texte d'au moins quatre lettres) et on propose
    ' la catégorie existante la plus proche avant de créer un doublon.
    listeCat = mod_Categories.ObtenirCategories()
    If mod_Categories.TrouverCorrespondanceProche(cat, listeCat, 2, 4, procheCat) Then
        reponse = MsgBox(mod_Display.FR("Vous avez saisi '") & cat & "'." & vbCrLf & _
                         mod_Display.FR("Une cat{e2}gorie tr{e1}s proche existe d{e2}j{a2} : '") & procheCat & "'." & vbCrLf & vbCrLf & _
                         mod_Display.FR("OUI : utiliser '") & procheCat & mod_Display.FR("' (recommand{e2}).") & vbCrLf & _
                         mod_Display.FR("NON : cr{e2}er quand m{ea}me '") & cat & mod_Display.FR("' comme nouvelle cat{e2}gorie."), _
                         vbYesNo + vbQuestion, mod_Display.FR("Cat{e2}gorie proche existante"))
        If reponse = vbYes Then cat = procheCat
    End If

    ' --- Garde-fou orthographique (niveau 2) : sous-catégorie -------------------------
    ' On compare uniquement aux sous-catégories DÉJÀ CONNUES DE LA CATÉGORIE choisie
    ' ci-dessus, et non à toutes celles du classeur, qui pourraient appartenir à une autre catégorie.
    If sous <> "" Then
        listeSous = mod_Categories.ObtenirSousCategories(cat)
        If mod_Categories.TrouverCorrespondanceProche(sous, listeSous, 2, 4, procheSous) Then
            reponse = MsgBox(mod_Display.FR("Vous avez saisi '") & sous & "'." & vbCrLf & _
                             mod_Display.FR("Une sous-cat{e2}gorie tr{e1}s proche existe d{e2}j{a2} dans '") & cat & mod_Display.FR("' : '") & procheSous & "'." & vbCrLf & vbCrLf & _
                             mod_Display.FR("OUI : utiliser '") & procheSous & mod_Display.FR("' (recommand{e2}).") & vbCrLf & _
                             mod_Display.FR("NON : cr{e2}er quand m{ea}me '") & sous & mod_Display.FR("' comme nouvelle sous-cat{e2}gorie."), _
                             vbYesNo + vbQuestion, mod_Display.FR("Sous-cat{e2}gorie proche existante"))
            If reponse = vbYes Then sous = procheSous
        End If
    End If

    ' AjouterCategoriePersonnalisee enregistre la paire dans TblCategories si elle n'y
    ' figure pas déjà et renvoie, par ByRef, la graphie EXACTE à utiliser ensuite
    ' (par exemple, si "sante" existe déjà sous la forme "Sante").
    If Not mod_Categories.AjouterCategoriePersonnalisee(cat, sous) Then
        ws.Range(NC_ADR_MESSAGE).value = mod_Display.FR("Le tableau de correspondance (feuille Param) est introuvable.") & _
                                          " " & mod_Display.FR("Ex{e2}cutez d'abord PreparerPhase1Categories.")
        Exit Sub
    End If

    g_NcCatResultat = cat
    g_NcSousResultat = sous
    g_NcValide = True
    FermerFeuilleNc ws
    Exit Sub

Erreur:
    Application.EnableEvents = True
    MsgBox mod_Display.FR("Erreur dans NcValider : ") & Err.Number & " - " & Err.Description, vbCritical

End Sub

' --- Bouton "Annuler" ---
Public Sub NcAnnuler()

    Dim ws As Worksheet
    Dim reponse As VbMsgBoxResult

    If Not g_NcEnCours Then Exit Sub
    On Error GoTo Erreur

    reponse = MsgBox(mod_Display.FR("Annuler la cr{e2}ation ? Rien ne sera enregistr{e2}."), vbYesNo + vbQuestion, mod_Display.FR("Confirmation"))
    If reponse = vbNo Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(NC_NOM_FEUILLE)
    g_NcValide = False
    FermerFeuilleNc ws
    Exit Sub

Erreur:
    Application.EnableEvents = True
    MsgBox mod_Display.FR("Erreur dans NcAnnuler : ") & Err.Number & " - " & Err.Description, vbCritical

End Sub


' =====================================================================================
' FERMETURE DU FORMULAIRE
' =====================================================================================
' IMPORTANT : g_NcEnCours doit passer à False AVANT de masquer la feuille, sinon le
' verrou (NcVerrouiller) ramènerait immédiatement l'opérateur sur cette feuille.
Private Sub FermerFeuilleNc(ByVal ws As Worksheet)

    g_NcEnCours = False
    g_NcVerrouActif = False

    ws.Range(ws.Cells(2, NC_COL_AIDE), ws.Cells(NC_LIGNE_AIDE_MAX, NC_COL_AIDE)).ClearContents
    ws.Visible = xlSheetVeryHidden

End Sub


' =====================================================================================
' MACRO DE TEST (sans risque : n'écrit que dans TblCategories, jamais dans les opérations)
' =====================================================================================
' Ouvre le formulaire seul, prérempli avec des exemples modifiables, puis affiche
' le résultat. Ctrl+G, taper : TesterNouvelleCategorie
Public Sub TesterNouvelleCategorie()

    Dim catRes As String, sousRes As String

    If mod_NouvelleCategorie.OuvrirNouvelleCategorie("", "", catRes, sousRes) Then
        MsgBox mod_Display.FR("Cat{e2}gorie enregistr{e2}e dans TblCategories :") & vbCrLf & vbCrLf & _
               mod_Display.FR("Cat{e2}gorie : ") & catRes & vbCrLf & _
               mod_Display.FR("Sous-cat{e2}gorie : ") & IIf(sousRes = "", "(" & mod_Display.FR("aucune") & ")", sousRes), _
               vbInformation, "Test"
    Else
        MsgBox mod_Display.FR("Test annul{e2}. Rien n'a {e2}t{e2} enregistr{e2}."), vbInformation, "Test"
    End If

End Sub


' =====================================================================================
' OUTILS INTERNES
' =====================================================================================

Private Function FeuilleSansErreur(ByVal nomFeuille As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nomFeuille)
    On Error GoTo 0
    Set FeuilleSansErreur = ws
End Function

' Vérifie si un tableau dynamique a été alloué (UBound provoque une erreur sinon; c'est
' l'usage standard en VBA pour distinguer un tableau vide d'un tableau jamais rempli).
Private Function EstTableauAlloue(ByRef arr As Variant) As Boolean
    Dim n As Long
    On Error Resume Next
    n = UBound(arr)
    EstTableauAlloue = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

