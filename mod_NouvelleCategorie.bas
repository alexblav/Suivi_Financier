Attribute VB_Name = "mod_NouvelleCategorie"
Option Explicit

' =====================================================================================
' MODULE : mod_NouvelleCategorie
'
' PHASE 3 du chantier "Categorie / Sous-categorie / Ventilation" - PARTIE 2/2
'
' ROLE (a lire en premier, meme si vous debutez) :
'   Ce module contient toute la LOGIQUE du formulaire de creation d'une categorie ou
'   d'une sous-categorie. Il s'appuie sur :
'     - la feuille-formulaire frm_NouvelleCategorie (mod_InstallNouvelleCategorie),
'     - le tableau TblCategories, via les fonctions de mod_Categories (Phase 1).
'
'   FONCTIONNEMENT "PAR-DESSUS" UN AUTRE FORMULAIRE :
'   Ce formulaire est concu pour s'ouvrir A PARTIR d'un autre formulaire-feuille deja
'   ouvert (le controle des categories, Phase 2), via son bouton "+ Nouvelle categorie".
'   La fonction OuvrirNouvelleCategorie() ci-dessous gere ce cas : elle attend la
'   fermeture de CE formulaire (Valider ou Annuler) et renvoie le resultat a l'appelant,
'   sans jamais toucher aux donnees elle-meme -- c'est l'appelant qui decide quoi faire
'   du resultat (ici : mod_ControleCategories.ControleNouvelleCategorie).
'
'   Le formulaire peut aussi etre teste seul (voir TesterNouvelleCategorie), sans passer
'   par le formulaire de controle.
'
' A PROPOS DES ACCENTS : fichier 100% ASCII, accents fabriques par la fonction TF().
' =====================================================================================

' --- Etat du formulaire ------------------------------------------------------------------
Public g_NcEnCours As Boolean        ' Vrai tant que le formulaire est ouvert (boucle d'attente)
Public g_NcVerrouActif As Boolean    ' Vrai tant que Worksheet_Deactivate doit reactiver la feuille

Private g_NcValide As Boolean        ' Vrai si l'operateur a clique sur "Valider"
Private g_NcCatResultat As String
Private g_NcSousResultat As String
Private g_NcNomFeuillePrec As String


' =====================================================================================
' FONCTION PRINCIPALE
' =====================================================================================
' Parametres :
'   catInitiale, sousInitiale : valeurs proposees au depart (peuvent etre vides).
'   catResultat, sousResultat : (en sortie, uniquement si la fonction renvoie True)
'                                la categorie et la sous-categorie validees par
'                                l'operateur, DEJA enregistrees dans TblCategories.
' Renvoie True si l'operateur a clique sur "Valider", False s'il a annule.
Public Function OuvrirNouvelleCategorie(ByVal catInitiale As String, ByVal sousInitiale As String, _
                                        ByRef catResultat As String, ByRef sousResultat As String) As Boolean

    Dim ws As Worksheet

    On Error GoTo Erreur

    Set ws = FeuilleSansErreur(NC_NOM_FEUILLE)
    If ws Is Nothing Then
        MsgBox TF("La feuille '") & NC_NOM_FEUILLE & TF("' est introuvable.") & vbCrLf & _
               TF("Ex{e2}cutez d'abord la macro CreerFeuilleNouvelleCategorie."), vbExclamation
        Exit Function
    End If

    g_NcNomFeuillePrec = ActiveSheet.Name

    g_NcValide = False
    RemplirChamps ws, catInitiale, sousInitiale

    ws.Visible = xlSheetVisible
    ws.Activate
    ws.Range(NC_ADR_CAT).Select

    ' CORRECTIF PREVENTIF (voir mod_ControleCategories.OuvrirFormulaireEtAttendre pour le
    ' detail) : ce formulaire repose lui aussi entierement sur Worksheet_Change pendant
    ' sa boucle d'attente ; on garantit donc que les evenements sont actifs a ce stade.
    Application.EnableEvents = True
    g_NcEnCours = True
    g_NcVerrouActif = True
    Do While g_NcEnCours
        DoEvents
    Loop

    ' A ce stade, NcValider ou NcAnnuler a deja masque la feuille (voir FermerFeuilleNc).
    ' On revient sur la feuille qui etait active avant l'ouverture.
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
    MsgBox TF("Erreur inattendue dans la cr{e2}ation de cat{e2}gorie :") & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical
    OuvrirNouvelleCategorie = False

End Function


' =====================================================================================
' Remplit les champs avec les valeurs de depart, et pose les 2 listes deroulantes
' =====================================================================================
Private Sub RemplirChamps(ByVal ws As Worksheet, ByVal catInitiale As String, ByVal sousInitiale As String)

    Dim evenementsAvant As Boolean
    Dim numErr As Long

    evenementsAvant = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Sortie

    ws.Range(NC_ADR_CAT).Value = catInitiale
    ws.Range(NC_ADR_SOUS).Value = sousInitiale
    ws.Range(NC_ADR_MESSAGE).Value = ""

    ' Liste "Categorie" : le nom "ListeCategories" est cree en Phase 1
    ' (mod_Categories.RafraichirListesCategories).
    '
    ' PAS D'ALERTE ICI (ni Warning, ni Stop) : saisir une valeur HORS liste est le but
    ' meme de cet ecran (creer une nouvelle categorie), donc interrompre l'operateur a
    ' CHAQUE fois qu'il tape un nom nouveau serait n'importe quoi -- sans compter que le
    ' clic sur "Valider" declenche deja un controle plus intelligent (garde-fou
    ' orthographique : "Vouliez-vous dire...?" si le texte ressemble beaucoup a une
    ' categorie existante). Mettre une alerte ICI EN PLUS obligeait l'operateur a
    ' confirmer DEUX FOIS la meme chose (Oui sur l'alerte Excel, puis reclic sur
    ' Valider) -- retour d'un operateur apres test, corrige ainsi plutot qu'en
    ' chainant les deux confirmations.
    ' Le menu deroulant (InCellDropdown) reste actif : on peut toujours choisir une
    ' categorie existante en un clic, seule l'ALERTE de saisie libre est retiree.
    With ws.Range(NC_ADR_CAT).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Formula1:="=ListeCategories"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = False    ' AlertStyle ci-dessus n'a alors aucun effet : c'est voulu, voir l'explication au-dessus.
    End With

    RemplirListeSousCatNc ws, catInitiale

Sortie:
    numErr = Err.Number
    Application.EnableEvents = evenementsAvant
    If numErr <> 0 Then Err.Raise numErr, "RemplirChamps", Err.Description

End Sub


' =====================================================================================
' Liste deroulante des sous-categories DEJA CONNUES de la categorie donnee
' =====================================================================================
' Meme principe qu'en Phase 2 (mod_ControleCategories.RemplirListeSousCategories) :
' les valeurs sont ecrites dans une colonne technique cachee, et le menu pointe vers
' cette plage via un nom dynamique. AlertStyle Warning : une sous-categorie NOUVELLE
' peut etre tapee librement.
Private Sub RemplirListeSousCatNc(ByVal ws As Worksheet, ByVal categorie As String)

    Dim liste() As String
    Dim sortie() As Variant
    Dim nb As Long, r As Long
    Dim colLettre As String

    ws.Range(ws.Cells(2, NC_COL_AIDE), ws.Cells(NC_LIGNE_AIDE_MAX, NC_COL_AIDE)).ClearContents

    liste = mod_Categories.ObtenirSousCategories(categorie)

    If EstTableauAlloue(liste) Then
        nb = UBound(liste) - LBound(liste) + 1
        ReDim sortie(1 To nb, 1 To 1)
        For r = 1 To nb
            sortie(r, 1) = liste(LBound(liste) + r - 1)
        Next r
        With ws.Cells(2, NC_COL_AIDE).Resize(nb, 1)
            .NumberFormat = "@"
            .Value2 = sortie
        End With
    End If

    colLettre = Chr$(64 + NC_COL_AIDE)     ' 26 -> "Z"
    ThisWorkbook.Names.Add Name:="ListeSousCatNouvelle", _
        RefersTo:="=OFFSET(" & NC_NOM_FEUILLE & "!$" & colLettre & "$2,0,0," & _
                  "MAX(1,COUNTA(" & NC_NOM_FEUILLE & "!$" & colLettre & "$2:$" & colLettre & "$" & NC_LIGNE_AIDE_MAX & ")),1)"

    ' Meme raisonnement que pour Categorie ci-dessus : pas d'alerte de saisie ici, le
    ' menu deroulant reste disponible, et le garde-fou orthographique intelligent
    ' (Valider) reste le seul controle -- plus besoin de confirmer deux fois.
    With ws.Range(NC_ADR_SOUS).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Formula1:="=ListeSousCatNouvelle"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = False
    End With

End Sub


' =====================================================================================
' EVENEMENTS DE LA FEUILLE (appeles par le code-behind de frm_NouvelleCategorie)
' =====================================================================================
' Comme en Phase 2, les evenements d'une feuille doivent vivre dans le module de CETTE
' feuille (voir CodeBehind_frm_NouvelleCategorie.txt). Ils ne font qu'appeler ces deux
' procedures : toute la logique reste ici, dans un module normal.

' Si la CATEGORIE change, l'ancienne sous-categorie n'a plus de sens : on l'efface et
' on recharge la liste des sous-categories de la nouvelle categorie.
Public Sub NcTraiterChangement(ByVal ws As Worksheet, ByVal Target As Range)

    Dim numErr As Long

    If Not g_NcEnCours Then Exit Sub
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Address(False, False) <> NC_ADR_CAT Then Exit Sub

    On Error GoTo Sortie
    Application.EnableEvents = False
    ws.Range(NC_ADR_SOUS).ClearContents
    RemplirListeSousCatNc ws, mod_DataStructure.CellText(Target.Value)

Sortie:
    numErr = Err.Number
    Application.EnableEvents = True
    If numErr <> 0 Then MsgBox TF("Erreur lors du changement de cat{e2}gorie : ") & Err.Description, vbExclamation

End Sub

' Verrou "modal" : tant que le formulaire est ouvert, on empeche l'operateur de le quitter
' en changeant d'onglet (Worksheet_Deactivate le ramene de force dessus).
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

    cat = mod_DataStructure.CellText(ws.Range(NC_ADR_CAT).Value)
    sous = mod_DataStructure.CellText(ws.Range(NC_ADR_SOUS).Value)

    If cat = "" Then
        ws.Range(NC_ADR_MESSAGE).Value = TF("La cat{e2}gorie est obligatoire.")
        Exit Sub
    End If

    ' --- Garde-fou orthographique (Niveau 2) : Categorie -----------------------------
    ' Le Niveau 1 (accents, espaces, majuscules) est deja gere silencieusement par
    ' AjouterCategoriePersonnalisee plus bas. Ici, on detecte une simple FAUTE DE FRAPPE
    ' (jusqu'a 2 lettres de difference, texte d'au moins 4 lettres) et on propose la
    ' categorie existante la plus proche avant de creer un doublon.
    listeCat = mod_Categories.ObtenirCategories()
    If mod_Categories.TrouverCorrespondanceProche(cat, listeCat, 2, 4, procheCat) Then
        reponse = MsgBox(TF("Vous avez saisi '") & cat & "'." & vbCrLf & _
                         TF("Une cat{e2}gorie tr{e1}s proche existe d{e2}j{a2} : '") & procheCat & "'." & vbCrLf & vbCrLf & _
                         TF("OUI : utiliser '") & procheCat & TF("' (recommand{e2}).") & vbCrLf & _
                         TF("NON : cr{e2}er quand m{ea}me '") & cat & TF("' comme nouvelle cat{e2}gorie."), _
                         vbYesNo + vbQuestion, TF("Cat{e2}gorie proche existante"))
        If reponse = vbYes Then cat = procheCat
    End If

    ' --- Garde-fou orthographique (Niveau 2) : Sous-categorie -------------------------
    ' On compare uniquement aux sous-categories DEJA CONNUES DE LA CATEGORIE choisie
    ' ci-dessus (et non a toutes les sous-categories du classeur, qui n'auraient pas de
    ' sens dans une autre categorie).
    If sous <> "" Then
        listeSous = mod_Categories.ObtenirSousCategories(cat)
        If mod_Categories.TrouverCorrespondanceProche(sous, listeSous, 2, 4, procheSous) Then
            reponse = MsgBox(TF("Vous avez saisi '") & sous & "'." & vbCrLf & _
                             TF("Une sous-cat{e2}gorie tr{e1}s proche existe d{e2}j{a2} dans '") & cat & TF("' : '") & procheSous & "'." & vbCrLf & vbCrLf & _
                             TF("OUI : utiliser '") & procheSous & TF("' (recommand{e2}).") & vbCrLf & _
                             TF("NON : cr{e2}er quand m{ea}me '") & sous & TF("' comme nouvelle sous-cat{e2}gorie."), _
                             vbYesNo + vbQuestion, TF("Sous-cat{e2}gorie proche existante"))
            If reponse = vbYes Then sous = procheSous
        End If
    End If

    ' AjouterCategoriePersonnalisee enregistre la paire dans TblCategories (si elle n'y
    ' est pas deja) et renvoie, par ByRef, la forme EXACTE a utiliser ensuite (par
    ' exemple si "sante" existait deja sous la forme "Sante").
    If Not mod_Categories.AjouterCategoriePersonnalisee(cat, sous) Then
        ws.Range(NC_ADR_MESSAGE).Value = TF("Le tableau de correspondance (feuille Param) est introuvable.") & _
                                          " " & TF("Ex{e2}cutez d'abord PreparerPhase1Categories.")
        Exit Sub
    End If

    g_NcCatResultat = cat
    g_NcSousResultat = sous
    g_NcValide = True
    FermerFeuilleNc ws
    Exit Sub

Erreur:
    Application.EnableEvents = True
    MsgBox TF("Erreur dans NcValider : ") & Err.Number & " - " & Err.Description, vbCritical

End Sub

' --- Bouton "Annuler" ---
Public Sub NcAnnuler()

    Dim ws As Worksheet
    Dim reponse As VbMsgBoxResult

    If Not g_NcEnCours Then Exit Sub
    On Error GoTo Erreur

    reponse = MsgBox(TF("Annuler la cr{e2}ation ? Rien ne sera enregistr{e2}."), vbYesNo + vbQuestion, TF("Confirmation"))
    If reponse = vbNo Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(NC_NOM_FEUILLE)
    g_NcValide = False
    FermerFeuilleNc ws
    Exit Sub

Erreur:
    Application.EnableEvents = True
    MsgBox TF("Erreur dans NcAnnuler : ") & Err.Number & " - " & Err.Description, vbCritical

End Sub


' =====================================================================================
' FERMETURE DU FORMULAIRE
' =====================================================================================
' IMPORTANT : g_NcEnCours doit passer a False AVANT de masquer la feuille, sinon le
' verrou (NcVerrouiller) ramenerait immediatement l'operateur dessus.
Private Sub FermerFeuilleNc(ByVal ws As Worksheet)

    g_NcEnCours = False
    g_NcVerrouActif = False

    ws.Range(ws.Cells(2, NC_COL_AIDE), ws.Cells(NC_LIGNE_AIDE_MAX, NC_COL_AIDE)).ClearContents
    ws.Visible = xlSheetVeryHidden

End Sub


' =====================================================================================
' MACRO DE TEST (sans risque : n'ecrit que dans TblCategories, jamais dans les operations)
' =====================================================================================
' Ouvre le formulaire seul, pre-rempli avec des valeurs d'exemple modifiables, et affiche
' le resultat. Ctrl+G, taper :  TesterNouvelleCategorie
Public Sub TesterNouvelleCategorie()

    Dim catRes As String, sousRes As String

    If mod_NouvelleCategorie.OuvrirNouvelleCategorie("", "", catRes, sousRes) Then
        MsgBox TF("Cat{e2}gorie enregistr{e2}e dans TblCategories :") & vbCrLf & vbCrLf & _
               TF("Cat{e2}gorie : ") & catRes & vbCrLf & _
               TF("Sous-cat{e2}gorie : ") & IIf(sousRes = "", "(" & TF("aucune") & ")", sousRes), _
               vbInformation, "Test"
    Else
        MsgBox TF("Test annul{e2}. Rien n'a {e2}t{e2} enregistr{e2}."), vbInformation, "Test"
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

' Teste si un tableau dynamique a ete alloue (UBound plante sinon : c'est l'idiome VBA
' standard pour distinguer "tableau vide" de "tableau jamais rempli").
Private Function EstTableauAlloue(ByRef arr As Variant) As Boolean
    Dim n As Long
    On Error Resume Next
    n = UBound(arr)
    EstTableauAlloue = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

' Traduit des codes en lettres accentuees (le fichier reste 100% ASCII).
Private Function TF(ByVal texte As String) As String
    Dim r As String
    r = texte
    r = Replace(r, "{e2}", ChrW(233))
    TF = r
End Function
