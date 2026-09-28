Attribute VB_Name = "mod_Ventilation"
Option Explicit

' =====================================================================================
' MODULE : mod_Ventilation
'
' PHASE 4 du chantier "Categorie / Sous-categorie / Ventilation" - PARTIE 2/2
' (Version 2 : nouvelle ergonomie, a la demande de l'operateur apres tests)
'
' ROLE (a lire en premier, meme si vous debutez) :
'   Ce module contient toute la LOGIQUE du formulaire de ventilation.
'
'   FONCTIONNEMENT D'ENSEMBLE :
'   Les lignes ajoutees a la ventilation sont gardees EN MEMOIRE pendant que le
'   formulaire est ouvert (tableaux g_VenLigneCat / g_VenLigneSous / g_VenLigneMontant),
'   et simplement AFFICHEES dans le tableau du haut (lignes 17 a 26 de la feuille) a
'   chaque changement. La feuille elle-meme ne sert que de VITRINE pour ce tableau : ce
'   ne sont plus ses cellules qui sont lues a la validation finale, contrairement a la
'   version precedente.
'
'   Le formulaire de SAISIE (bas de la feuille) sert a la fois a AJOUTER une nouvelle
'   ligne et a MODIFIER une ligne existante :
'     - g_VenIndexEdition = 0  -> le bouton "Ajouter la ligne" cree une ligne de plus.
'     - g_VenIndexEdition = N  -> le bouton "Ajouter la ligne" REMPLACE la ligne N
'       (celle qu'on est en train de corriger, chargee via son bouton "Editer").
'
'   "Editer" une ligne la RETIRE IMMEDIATEMENT du tableau et la place dans le
'   formulaire de saisie. Consequence volontaire et pratique : si l'operateur clique
'   sur "Editer" puis change d'avis sans cliquer sur "Ajouter la ligne" (par exemple en
'   cliquant sur "Editer" d'une AUTRE ligne, ou sur "Terminer"), la ligne reste
'   supprimee. C'est le seul moyen de supprimer une ligne pour l'instant : simple, et
'   explicitement signale a l'operateur par un message a chaque fois que cela se produit.
'
' A PROPOS DES ACCENTS : fichier 100% ASCII, accents fabriques par la fonction TF().
' =====================================================================================

' --- Etat du formulaire ------------------------------------------------------------------
Public g_VenEnCours As Boolean        ' Vrai tant que le formulaire est ouvert
Public g_VenVerrouActif As Boolean    ' Vrai tant que Worksheet_Deactivate doit reactiver la feuille

Private g_VenValide As Boolean
Private g_VenNomFeuillePrec As String
Private g_VenIdTransaction As String
Private g_VenMontantOperation As Double     ' montant SIGNE de l'operation (tel que dans TblOperations)

' --- Lignes de la ventilation EN COURS (en memoire, tant que le formulaire est ouvert) ---
' --- Lignes de la ventilation EN COURS (en memoire, tant que le formulaire est ouvert) ---
' Taille fixee a 10 : DOIT rester egale a VEN_NB_LIGNES (mod_InstallVentilation). VBA
' n'autorise pas de dimensionner un tableau de module avec une constante d'un autre
' module, d'ou ce nombre repete ici en dur -- si vous changez un jour VEN_NB_LIGNES,
' changez aussi ce "10" a 3 endroits juste en dessous.
Private g_VenNbLignes As Long
Private g_VenLigneCat(1 To 10) As String
Private g_VenLigneSous(1 To 10) As String
Private g_VenLigneMontant(1 To 10) As Double

' --- Index (1..10) de la ligne en cours de modification ; 0 = aucune (mode "ajout") ---
Private g_VenIndexEdition As Long


' =====================================================================================
' FONCTION PRINCIPALE
' =====================================================================================
' Parametres :
'   idTransaction        : ID_Transaction de l'operation (cle vers TblOperations).
'   dateOp               : date de l'operation (Date ou nombre serie Excel).
'   tiers, libelle        : texte affiche en lecture seule.
'   montantOp            : montant SIGNE de l'operation (negatif = depense).
'   categorieActuelle, sousCategorieActuelle : affiches en lecture seule (contexte),
'                          et utilises pour pre-remplir la 1ere saisie.
' Renvoie True si l'operateur a valide une ventilation complete (les lignes ont deja
' ete enregistrees dans TblVentilations), False s'il a annule.
Public Function OuvrirVentilation(ByVal idTransaction As String, ByVal dateOp As Variant, _
                                  ByVal tiers As String, ByVal libelle As String, _
                                  ByVal montantOp As Double, _
                                  ByVal categorieActuelle As String, ByVal sousCategorieActuelle As String) As Boolean

    Dim ws As Worksheet
    Dim evAvant As Boolean

    On Error GoTo Erreur

    Set ws = FeuilleSansErreur(VEN_NOM_FEUILLE)
    If ws Is Nothing Then
        MsgBox TF("La feuille '") & VEN_NOM_FEUILLE & TF("' est introuvable.") & vbCrLf & _
               TF("Ex{e2}cutez d'abord la macro PreparerPhase4Ventilation."), vbExclamation
        Exit Function
    End If

    g_VenIdTransaction = idTransaction
    g_VenMontantOperation = montantOp
    g_VenNomFeuillePrec = ActiveSheet.Name
    g_VenValide = False
    g_VenNbLignes = 0
    g_VenIndexEdition = 0

    RemplirEntete ws, dateOp, tiers, libelle, montantOp, categorieActuelle, sousCategorieActuelle
    RafraichirAffichageLignes ws

    ' Petite aide : la saisie demarre avec la categorie actuelle de l'operation,
    ' puisqu'une ventilation reste tres souvent dans la meme categorie generale.
    evAvant = Application.EnableEvents
    Application.EnableEvents = False
    ws.Range(VEN_ADR_SAISIE_CAT).Value = ""
    ws.Range(VEN_ADR_SAISIE_SOUS).Value = ""
    ws.Range(VEN_ADR_SAISIE_MONTANT).Value = ""
    ws.Range(VEN_ADR_SAISIE_MESSAGE).Value = ""
    PoserValidationSaisieCategorie ws
    If categorieActuelle <> "" And categorieActuelle <> CategorieVentilePublique() Then
        ws.Range(VEN_ADR_SAISIE_CAT).Value = categorieActuelle
        RemplirListeSousCatSaisie ws, categorieActuelle
    Else
        RemplirListeSousCatSaisie ws, ""
    End If
    Application.EnableEvents = evAvant

    RecalculerTotaux ws

    ws.Visible = xlSheetVisible
    ws.Activate
    ws.Range(VEN_ADR_SAISIE_CAT).Select

    ' Voir mod_ControleCategories.OuvrirFormulaireEtAttendre pour le detail de ce
    ' correctif : la boucle qui suit repose entierement sur Worksheet_Change, on
    ' garantit donc que les evenements sont actifs avant de l'entamer.
    Application.EnableEvents = True
    g_VenEnCours = True
    g_VenVerrouActif = True
    Do While g_VenEnCours
        DoEvents
    Loop

    On Error Resume Next
    ThisWorkbook.Worksheets(g_VenNomFeuillePrec).Activate
    On Error GoTo 0

    OuvrirVentilation = g_VenValide
    Exit Function

Erreur:
    Application.EnableEvents = True
    g_VenEnCours = False
    g_VenVerrouActif = False
    MsgBox TF("Erreur inattendue dans la ventilation :") & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical
    OuvrirVentilation = False

End Function


' =====================================================================================
' Remplit l'entete (lecture seule) et le montant a ventiler
' =====================================================================================
Private Sub RemplirEntete(ByVal ws As Worksheet, ByVal dateOp As Variant, ByVal tiers As String, _
                          ByVal libelle As String, ByVal montantOp As Double, _
                          ByVal categorieActuelle As String, ByVal sousCategorieActuelle As String)

    Dim evenementsAvant As Boolean
    Dim numErr As Long
    Dim catAffichee As String

    evenementsAvant = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Sortie

    ws.Range(VEN_ADR_DATE).Value = FormaterDate(dateOp)
    ws.Range(VEN_ADR_TIERS).Value = tiers
    ws.Range(VEN_ADR_LIBELLE).Value = libelle
    ws.Range(VEN_ADR_MONTANT).Value = Format(montantOp, "#,##0.00") & " " & ChrW(8364)

    catAffichee = categorieActuelle
    If sousCategorieActuelle <> "" Then catAffichee = catAffichee & " / " & sousCategorieActuelle
    If catAffichee = "" Then catAffichee = "(" & TF("aucune") & ")"
    ws.Range(VEN_ADR_CATACTUELLE).Value = catAffichee

    ws.Range(VEN_ADR_MONTANT_A_VENTILER).Value = Abs(montantOp)

Sortie:
    numErr = Err.Number
    Application.EnableEvents = evenementsAvant
    If numErr <> 0 Then Err.Raise numErr, "RemplirEntete", Err.Description

End Sub


' =====================================================================================
' AFFICHAGE du tableau des lignes deja ajoutees (relit g_VenLigneCat/Sous/Montant)
' =====================================================================================
Private Sub RafraichirAffichageLignes(ByVal ws As Worksheet)

    Dim evAvant As Boolean
    Dim i As Long, ligne As Long

    evAvant = Application.EnableEvents
    Application.EnableEvents = False

    ' On efface TOUTES les lignes d'affichage d'abord (une ligne editee, donc retiree,
    ' ne doit pas laisser une ancienne valeur trainer en bas du tableau).
    ws.Range(ws.Cells(VEN_LIGNE_GRILLE_DEBUT, VEN_COL_CAT), ws.Cells(VEN_LIGNE_GRILLE_FIN, VEN_COL_MONTANT)).ClearContents

    For i = 1 To g_VenNbLignes
        ligne = VEN_LIGNE_GRILLE_DEBUT + i - 1
        ws.Cells(ligne, VEN_COL_CAT).Value = g_VenLigneCat(i)
        ws.Cells(ligne, VEN_COL_SOUS).Value = g_VenLigneSous(i)
        ws.Cells(ligne, VEN_COL_MONTANT).Value = g_VenLigneMontant(i)
    Next i

    Application.EnableEvents = evAvant

End Sub


' =====================================================================================
' Recalcule les totaux a partir de la liste EN MEMOIRE (plus fiable que relire des
' cellules : les lignes ne sont de toute facon plus stockees dans les cellules).
' =====================================================================================
Private Sub RecalculerTotaux(ByVal ws As Worksheet)

    Dim i As Long
    Dim total As Double
    Dim aVentiler As Double
    Dim reste As Double

    For i = 1 To g_VenNbLignes
        total = total + g_VenLigneMontant(i)
    Next i

    aVentiler = mod_DataStructure.ToDouble(ws.Range(VEN_ADR_MONTANT_A_VENTILER).Value2)
    reste = aVentiler - total

    ws.Range(VEN_ADR_TOTAL_SAISI).Value = total
    ws.Range(VEN_ADR_RESTE).Value = reste

    ' Arrondi a 2 decimales avant comparaison (piege des calculs en virgule flottante,
    ' deja documente dans ce projet).
    If Abs(Round(reste, 2)) < 0.005 Then
        ws.Range(VEN_ADR_RESTE).Font.Color = RGB(31, 120, 60)      ' vert : le compte y est
    Else
        ws.Range(VEN_ADR_RESTE).Font.Color = RGB(192, 40, 40)      ' rouge : ecart restant
    End If

End Sub


' =====================================================================================
' Validation "Categorie" du formulaire de SAISIE (liste stricte + message explicite)
' =====================================================================================
Private Sub PoserValidationSaisieCategorie(ByVal ws As Worksheet)
    With ws.Range(VEN_ADR_SAISIE_CAT).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=ListeCategories"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ErrorTitle = TF("Cat{e2}gorie inconnue")
        .ErrorMessage = TF("Choisissez une cat{e2}gorie dans la liste, ou laissez le champ vide.") & vbCrLf & _
                        TF("Pour cr{e2}er une nouvelle cat{e2}gorie, utilisez le bouton '+ Nouvelle cat{e2}gorie'.")
    End With
End Sub

' Liste deroulante des sous-categories DEJA CONNUES de la categorie choisie dans le
' formulaire de saisie. Meme principe que dans le formulaire de controle (Phase 2) :
' une seule liste a gerer maintenant, ecrite dans la colonne technique cachee (Z).
Private Sub RemplirListeSousCatSaisie(ByVal ws As Worksheet, ByVal categorie As String)

    Dim liste() As String
    Dim sortie() As Variant
    Dim nb As Long, r As Long

    ws.Range(ws.Cells(2, VEN_COL_AIDE), ws.Cells(VEN_LIGNE_AIDE_MAX, VEN_COL_AIDE)).ClearContents

    liste = mod_Categories.ObtenirSousCategories(categorie)

    If EstTableauAlloue(liste) Then
        nb = UBound(liste) - LBound(liste) + 1
        ReDim sortie(1 To nb, 1 To 1)
        For r = 1 To nb
            sortie(r, 1) = liste(LBound(liste) + r - 1)
        Next r
        With ws.Cells(2, VEN_COL_AIDE).Resize(nb, 1)
            .NumberFormat = "@"
            .Value2 = sortie
        End With
    End If

    ThisWorkbook.Names.Add Name:="ListeSousCatVentSaisie", _
        RefersTo:="=OFFSET(" & VEN_NOM_FEUILLE & "!$Z$2,0,0," & _
                  "MAX(1,COUNTA(" & VEN_NOM_FEUILLE & "!$Z$2:$Z$" & VEN_LIGNE_AIDE_MAX & ")),1)"

    With ws.Range(VEN_ADR_SAISIE_SOUS).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=ListeSousCatVentSaisie"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ErrorTitle = TF("Sous-cat{e2}gorie inconnue")
        .ErrorMessage = TF("Cette sous-cat{e2}gorie n'existe pas pour la cat{e2}gorie choisie.") & vbCrLf & _
                        TF("Choisissez-en une dans la liste, laissez le champ vide, ou utilisez le bouton") & _
                        " '+ " & TF("Nouvelle cat{e2}gorie") & "' " & TF("pour en cr{e2}er une nouvelle.")
    End With

End Sub


' =====================================================================================
' EVENEMENTS DE LA FEUILLE (appeles par le code-behind de frm_Ventilation)
' =====================================================================================

' Seule la Categorie DE LA SAISIE declenche une action : on efface la Sous-categorie
' et on recharge sa liste. Le tableau d'affichage (lignes 17-26) n'a plus besoin d'etre
' surveille : ce n'est plus lui qui est directement modifie par l'operateur.
Public Sub VenTraiterChangement(ByVal ws As Worksheet, ByVal Target As Range)

    Dim numErr As Long

    If Not g_VenEnCours Then Exit Sub
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Address(False, False) <> VEN_ADR_SAISIE_CAT Then Exit Sub

    On Error GoTo Sortie
    Application.EnableEvents = False
    ws.Range(VEN_ADR_SAISIE_SOUS).ClearContents
    RemplirListeSousCatSaisie ws, mod_DataStructure.CellText(Target.Value)

Sortie:
    numErr = Err.Number
    Application.EnableEvents = True
    If numErr <> 0 Then MsgBox TF("Erreur lors du changement de cat{e2}gorie : ") & Err.Description, vbExclamation

End Sub

Public Sub VenVerrouiller(ByVal ws As Worksheet)
    If g_VenVerrouActif Then ws.Activate
End Sub


' =====================================================================================
' ACTIONS DES BOUTONS
' =====================================================================================

' --- Bouton "+ Nouvelle categorie" (a cote du champ Categorie de la saisie) ---
Public Sub VenNouvelleCategorie()

    Dim ws As Worksheet
    Dim catActuelle As String, sousActuelle As String
    Dim catRes As String, sousRes As String
    Dim ok As Boolean

    If Not g_VenEnCours Then Exit Sub
    On Error GoTo Erreur
    Set ws = ThisWorkbook.Worksheets(VEN_NOM_FEUILLE)

    catActuelle = mod_DataStructure.CellText(ws.Range(VEN_ADR_SAISIE_CAT).Value)
    sousActuelle = mod_DataStructure.CellText(ws.Range(VEN_ADR_SAISIE_SOUS).Value)

    ' On suspend le verrou d'ACTIVATION (pas g_VenEnCours) le temps d'ouvrir le
    ' formulaire de creation par-dessus -- meme principe qu'en Phase 2/mod_ControleCategories.
    g_VenVerrouActif = False
    ok = mod_NouvelleCategorie.OuvrirNouvelleCategorie(catActuelle, sousActuelle, catRes, sousRes)
    g_VenVerrouActif = True
    ws.Activate

    If ok Then
        mod_Categories.RafraichirListesCategories
        Application.EnableEvents = False
        PoserValidationSaisieCategorie ws
        ws.Range(VEN_ADR_SAISIE_CAT).Value = catRes
        ws.Range(VEN_ADR_SAISIE_SOUS).Value = sousRes
        RemplirListeSousCatSaisie ws, catRes
        Application.EnableEvents = True
    End If
    Exit Sub

Erreur:
    g_VenVerrouActif = True
    Application.EnableEvents = True
    MsgBox TF("Erreur dans VenNouvelleCategorie : ") & Err.Number & " - " & Err.Description, vbCritical

End Sub

' --- Bouton "Ajouter la ligne" ---
Public Sub VenAjouterLigne()

    Dim ws As Worksheet
    Dim cat As String, sous As String
    Dim montant As Variant
    Dim listeCat() As String, listeSous() As String
    Dim indiceCible As Long

    If Not g_VenEnCours Then Exit Sub
    On Error GoTo Erreur
    Set ws = ThisWorkbook.Worksheets(VEN_NOM_FEUILLE)

    cat = mod_DataStructure.CellText(ws.Range(VEN_ADR_SAISIE_CAT).Value)
    sous = mod_DataStructure.CellText(ws.Range(VEN_ADR_SAISIE_SOUS).Value)
    montant = ws.Range(VEN_ADR_SAISIE_MONTANT).Value2

    If cat = "" Then
        ws.Range(VEN_ADR_SAISIE_MESSAGE).Value = TF("La cat{e2}gorie est obligatoire.")
        Exit Sub
    End If
    If Not IsNumeric(montant) Or CDbl(montant) <= 0 Then
        ws.Range(VEN_ADR_SAISIE_MESSAGE).Value = TF("Le montant doit {ea}tre un nombre positif.")
        Exit Sub
    End If

    listeCat = mod_Categories.ObtenirCategories()
    If Not TrouverExact(cat, listeCat) Then
        ws.Range(VEN_ADR_SAISIE_MESSAGE).Value = TF("La cat{e2}gorie '") & cat & TF("' n'existe pas.")
        Exit Sub
    End If
    cat = mod_Categories.FormeCanonique(cat, listeCat)

    If sous <> "" Then
        listeSous = mod_Categories.ObtenirSousCategories(cat)
        If Not TrouverExact(sous, listeSous) Then
            ws.Range(VEN_ADR_SAISIE_MESSAGE).Value = TF("La sous-cat{e2}gorie '") & sous & _
                                                      TF("' n'existe pas pour '") & cat & "'."
            Exit Sub
        End If
        sous = mod_Categories.FormeCanonique(sous, listeSous)
    End If

    If g_VenIndexEdition = 0 Then
        ' Mode AJOUT : une ligne de plus.
        If g_VenNbLignes >= VEN_NB_LIGNES Then
            ws.Range(VEN_ADR_SAISIE_MESSAGE).Value = TF("Limite de ") & VEN_NB_LIGNES & TF(" lignes atteinte.")
            Exit Sub
        End If
        g_VenNbLignes = g_VenNbLignes + 1
        indiceCible = g_VenNbLignes
    Else
        ' Mode EDITION : on remplace la ligne retiree tout a l'heure. Comme elle a deja
        ' ete retiree de la liste (voir VenEditerLigne), on l'ajoute simplement a la fin.
        g_VenNbLignes = g_VenNbLignes + 1
        indiceCible = g_VenNbLignes
    End If

    g_VenLigneCat(indiceCible) = cat
    g_VenLigneSous(indiceCible) = sous
    g_VenLigneMontant(indiceCible) = CDbl(montant)

    g_VenIndexEdition = 0
    RafraichirAffichageLignes ws
    RecalculerTotaux ws
    ViderChampsSaisie ws
    Exit Sub

Erreur:
    Application.EnableEvents = True
    MsgBox TF("Erreur dans VenAjouterLigne : ") & Err.Number & " - " & Err.Description, vbCritical

End Sub

' --- Bouton "Effacer la saisie" ---
Public Sub VenEffacerSaisie()
    Dim ws As Worksheet
    If Not g_VenEnCours Then Exit Sub
    On Error GoTo Erreur
    Set ws = ThisWorkbook.Worksheets(VEN_NOM_FEUILLE)
    g_VenIndexEdition = 0
    ViderChampsSaisie ws
    Exit Sub
Erreur:
    Application.EnableEvents = True
    MsgBox TF("Erreur dans VenEffacerSaisie : ") & Err.Number & " - " & Err.Description, vbCritical
End Sub

Private Sub ViderChampsSaisie(ByVal ws As Worksheet)
    Dim evAvant As Boolean
    evAvant = Application.EnableEvents
    Application.EnableEvents = False
    ws.Range(VEN_ADR_SAISIE_CAT).ClearContents
    ws.Range(VEN_ADR_SAISIE_SOUS).ClearContents
    ws.Range(VEN_ADR_SAISIE_MONTANT).ClearContents
    ws.Range(VEN_ADR_SAISIE_MESSAGE).Value = ""
    RemplirListeSousCatSaisie ws, ""
    Application.EnableEvents = evAvant
    ws.Range(VEN_ADR_SAISIE_CAT).Select
End Sub

' --- Bouton "Editer" d'une ligne (un par ligne du tableau d'affichage) ---
' Application.Caller donne le NOM du bouton clique ("btnVenEditerLigneN") : on en
' extrait N pour savoir quelle ligne charger. C'est la technique standard en VBA pour
' plusieurs boutons qui declenchent la meme macro.
Public Sub VenEditerLigne()

    Dim ws As Worksheet
    Dim nomBouton As String
    Dim indice As Long
    Dim i As Long
    Dim evAvant As Boolean

    If Not g_VenEnCours Then Exit Sub
    On Error GoTo Erreur

    nomBouton = CStr(Application.Caller)
    indice = CLng(Mid$(nomBouton, Len("btnVenEditerLigne") + 1))

    Set ws = ThisWorkbook.Worksheets(VEN_NOM_FEUILLE)

    If indice > g_VenNbLignes Then
        MsgBox TF("Aucune ligne {a2} cette position."), vbInformation
        Exit Sub
    End If

    ' Si une autre ligne etait deja en cours de modification (retiree, jamais
    ' re-ajoutee), le signaler clairement : elle reste perdue.
    If g_VenIndexEdition <> 0 Then
        MsgBox TF("La ligne pr{e2}c{e2}demment retir{e2}e pour modification n'a pas {e2}t{e2} valid{e2}e : elle reste supprim{e2}e."), vbInformation
    End If

    ' On charge la ligne dans le formulaire de saisie...
    evAvant = Application.EnableEvents
    Application.EnableEvents = False
    PoserValidationSaisieCategorie ws
    ws.Range(VEN_ADR_SAISIE_CAT).Value = g_VenLigneCat(indice)
    RemplirListeSousCatSaisie ws, g_VenLigneCat(indice)
    ws.Range(VEN_ADR_SAISIE_SOUS).Value = g_VenLigneSous(indice)
    ws.Range(VEN_ADR_SAISIE_MONTANT).Value = g_VenLigneMontant(indice)
    ws.Range(VEN_ADR_SAISIE_MESSAGE).Value = TF("Ligne retir{e2}e du tableau pour modification.") & _
        TF(" Cliquez sur 'Ajouter la ligne' pour la remettre (avec vos changements), sinon elle restera supprim{e2}e.")
    Application.EnableEvents = evAvant

    ' ...et on la retire IMMEDIATEMENT de la liste (voir l'explication en tete de module).
    For i = indice To g_VenNbLignes - 1
        g_VenLigneCat(i) = g_VenLigneCat(i + 1)
        g_VenLigneSous(i) = g_VenLigneSous(i + 1)
        g_VenLigneMontant(i) = g_VenLigneMontant(i + 1)
    Next i
    g_VenNbLignes = g_VenNbLignes - 1
    g_VenIndexEdition = indice

    RafraichirAffichageLignes ws
    RecalculerTotaux ws
    ws.Range(VEN_ADR_SAISIE_CAT).Select
    Exit Sub

Erreur:
    Application.EnableEvents = True
    MsgBox TF("Erreur dans VenEditerLigne : ") & Err.Number & " - " & Err.Description, vbCritical

End Sub

' --- Bouton "Terminer la ventilation" (global) ---
Public Sub VenTerminer()

    Dim ws As Worksheet
    Dim i As Long
    Dim sommeSaisie As Double, aVentiler As Double

    If Not g_VenEnCours Then Exit Sub
    On Error GoTo Erreur
    Set ws = ThisWorkbook.Worksheets(VEN_NOM_FEUILLE)

    If g_VenIndexEdition <> 0 Then
        ws.Range(VEN_ADR_SAISIE_MESSAGE).Value = _
            TF("Une ligne est en cours de modification : cliquez sur 'Ajouter la ligne' pour la valider, ou sur 'Effacer la saisie' pour l'abandonner, avant de terminer.")
        Exit Sub
    End If

    If g_VenNbLignes = 0 Then
        ws.Range(VEN_ADR_SAISIE_MESSAGE).Value = TF("Aucune ligne ajout{e2}e. Renseignez au moins une cat{e2}gorie et un montant.")
        Exit Sub
    End If

    For i = 1 To g_VenNbLignes
        sommeSaisie = sommeSaisie + g_VenLigneMontant(i)
    Next i
    aVentiler = Abs(g_VenMontantOperation)

    If Abs(Round(sommeSaisie, 2) - Round(aVentiler, 2)) >= 0.005 Then
        ws.Range(VEN_ADR_SAISIE_MESSAGE).Value = _
            TF("Le total saisi (") & Format(sommeSaisie, "#,##0.00") & TF(" {e2}uros) ne correspond pas au montant de l'op{e2}ration (") & _
            Format(aVentiler, "#,##0.00") & TF(" {e2}uros).") & vbCrLf & _
            TF("La somme des lignes doit {ea}tre EXACTEMENT {e2}gale. Corrigez avant de terminer.")
        Exit Sub
    End If

    For i = 1 To g_VenNbLignes
        AjouterLigneVentilation g_VenIdTransaction, g_VenLigneCat(i), g_VenLigneSous(i), g_VenLigneMontant(i)
    Next i

    g_VenValide = True
    FermerFeuilleVen ws
    Exit Sub

Erreur:
    Application.EnableEvents = True
    MsgBox TF("Erreur dans VenTerminer : ") & Err.Number & " - " & Err.Description, vbCritical

End Sub

' --- Bouton "Annuler" (global) ---
Public Sub VenAnnuler()

    Dim ws As Worksheet
    Dim reponse As VbMsgBoxResult

    If Not g_VenEnCours Then Exit Sub
    On Error GoTo Erreur

    reponse = MsgBox(TF("Annuler toute la ventilation ? Rien ne sera enregistr{e2}."), vbYesNo + vbQuestion, TF("Confirmation"))
    If reponse = vbNo Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(VEN_NOM_FEUILLE)
    g_VenValide = False
    FermerFeuilleVen ws
    Exit Sub

Erreur:
    Application.EnableEvents = True
    MsgBox TF("Erreur dans VenAnnuler : ") & Err.Number & " - " & Err.Description, vbCritical

End Sub


' =====================================================================================
' FERMETURE DU FORMULAIRE
' =====================================================================================
Private Sub FermerFeuilleVen(ByVal ws As Worksheet)

    g_VenEnCours = False
    g_VenVerrouActif = False

    ws.Range(ws.Cells(2, VEN_COL_AIDE), ws.Cells(VEN_LIGNE_AIDE_MAX, VEN_COL_AIDE)).ClearContents
    ws.Visible = xlSheetVeryHidden

End Sub


' =====================================================================================
' ENREGISTREMENT D'UNE LIGNE DANS TblVentilations
' =====================================================================================
' Ecrit une ligne dans TblVentilations. Les colonnes sont retrouvees PAR LEUR NOM
' (jamais par un numero fixe) : plus robuste si l'ordre des colonnes change un jour.
'
' Les colonnes de suivi sante (Notes, Date_consult, StatutSante...) ne sont remplies
' QUE si cette ligne est elle-meme categorisee "Frais, remb sante" : c'est la seule
' sous-categorie suivie par le moteur sante existant (mod_SuiviSante). Une ligne
' sante est initialisee exactement comme le fait mod_ImportOFX pour une operation
' fraichement importee dont le libelle n'a pas pu etre decode automatiquement
' (StatutSante = "KO", Date_consult = la "sentinelle" 02/01/1900) : elle sera ainsi
' proposee au rapprochement (frm_RapprochementNotes / frm_GenerationCle) des que
' mod_FormulairesNotes sera adapte pour reconnaitre aussi les lignes de
' TblVentilations (prochaine etape de ce chantier, pas encore faite).
Private Sub AjouterLigneVentilation(ByVal idTransaction As String, ByVal categorie As String, _
                                    ByVal sousCategorie As String, ByVal montant As Double)

    Dim wsData As Worksheet
    Dim tbl As ListObject
    Dim ligne As Long
    Dim estSante As Boolean

    Set wsData = ThisWorkbook.Worksheets(VEN_NOM_FEUILLE_DONNEES)
    Set tbl = wsData.ListObjects(VEN_NOM_TABLE)

    ligne = tbl.ListRows.Add.Range.Row

    EcrireCelluleTexte wsData, tbl, ligne, "ID_Transaction", idTransaction
    EcrireCelluleTexte wsData, tbl, ligne, "Categorie", categorie
    EcrireCelluleTexte wsData, tbl, ligne, "SousCategorie", sousCategorie
    EcrireCelluleNombre wsData, tbl, ligne, "Montant", montant, "#,##0.00"
    EcrireCelluleDate wsData, tbl, ligne, "DateVentilation", Now, "dd/mm/yyyy hh:mm"

    estSante = (mod_Categories.NormaliserTexte(sousCategorie) = mod_Categories.NormaliserTexte(SousCategorieSanteReference()))

    If estSante Then
        EcrireCelluleTexte wsData, tbl, ligne, "Notes", ""
        EcrireCelluleDate wsData, tbl, ligne, "Date_consult", DateSerial(1900, 1, 2), "dd/mm/yyyy"
        EcrireCelluleTexte wsData, tbl, ligne, "Spe_Consult", ""
        EcrireCelluleTexte wsData, tbl, ligne, "StatutSante", "KO"
        EcrireCelluleTexte wsData, tbl, ligne, "CommentaireSante", ""
        EcrireCelluleTexte wsData, tbl, ligne, "Beneficiaire", ""
        ' SoldeSante, DepassementHoraires, Franchise restent VIDES : ce sont des
        ' colonnes que le moteur sante calcule ou que l'operateur saisit lui-meme
        ' plus tard, jamais a la creation de la ligne (meme principe qu'a l'import).
    End If

End Sub

Private Sub EcrireCelluleTexte(ByVal ws As Worksheet, ByVal tbl As ListObject, ByVal ligne As Long, _
                               ByVal nomColonne As String, ByVal valeur As String)
    Dim col As Long
    col = tbl.ListColumns(nomColonne).Index
    ws.Cells(ligne, col).NumberFormat = "@"
    ws.Cells(ligne, col).Value = valeur
End Sub

Private Sub EcrireCelluleNombre(ByVal ws As Worksheet, ByVal tbl As ListObject, ByVal ligne As Long, _
                                ByVal nomColonne As String, ByVal valeur As Double, ByVal formatNombre As String)
    Dim col As Long
    col = tbl.ListColumns(nomColonne).Index
    ws.Cells(ligne, col).NumberFormat = formatNombre
    ws.Cells(ligne, col).Value = valeur
End Sub

Private Sub EcrireCelluleDate(ByVal ws As Worksheet, ByVal tbl As ListObject, ByVal ligne As Long, _
                              ByVal nomColonne As String, ByVal valeur As Date, ByVal formatDate As String)
    Dim col As Long
    col = tbl.ListColumns(nomColonne).Index
    ws.Cells(ligne, col).NumberFormat = formatDate
    ws.Cells(ligne, col).Value = valeur
End Sub

Private Function SousCategorieSanteReference() As String
    SousCategorieSanteReference = "Frais, remb sant" & ChrW(233)
End Function


' =====================================================================================
' MACRO DE TEST
' =====================================================================================
' Ouvre le formulaire seul, sur une operation FICTIVE (aucune lecture ni ecriture dans
' TblOperations). L'ID_Transaction utilise est prefixe "TEST-" : toute ligne enregistree
' dans TblVentilations lors de ce test est donc immediatement reconnaissable, et peut
' etre supprimee sans risque (filtrez la colonne ID_Transaction sur "commence par TEST-").
' Ctrl+G, taper :  TesterVentilation
Public Sub TesterVentilation()

    Dim ok As Boolean
    Dim idTest As String

    idTest = "TEST-" & Format(Now, "yyyymmddhhnnss")

    ok = OuvrirVentilation(idTest, Date, TF("Tiers de test"), TF("Op{e2}ration fictive pour tester le formulaire"), _
                           -100, TF("Sant{e2}, pr{e2}voyance"), "")

    If ok Then
        MsgBox TF("Ventilation enregistr{e2}e dans TblVentilations sous l'identifiant '") & idTest & "'." & vbCrLf & _
               TF("Vous pouvez la retrouver (et la supprimer) sur la feuille '") & VEN_NOM_FEUILLE_DONNEES & "'.", _
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

' Teste si un tableau dynamique a ete alloue (meme idiome qu'ailleurs dans le projet).
Private Function EstTableauAlloue(ByRef arr As Variant) As Boolean
    Dim n As Long
    On Error Resume Next
    n = UBound(arr)
    EstTableauAlloue = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

' Cherche "texte" (accents/espaces/majuscules ignores) dans "liste".
Private Function TrouverExact(ByVal texte As String, ByRef liste() As String) As Boolean
    Dim i As Long
    Dim texteNorm As String
    On Error GoTo PasTrouve
    texteNorm = mod_Categories.NormaliserTexte(texte)
    For i = LBound(liste) To UBound(liste)
        If mod_Categories.NormaliserTexte(liste(i)) = texteNorm Then
            TrouverExact = True
            Exit Function
        End If
    Next i
PasTrouve:
End Function

' Met une date en texte "jj/mm/aaaa", qu'elle soit de type Date ou nombre serie Excel.
Private Function FormaterDate(ByVal v As Variant) As String
    If IsEmpty(v) Then
        FormaterDate = ""
    ElseIf IsDate(v) Then
        FormaterDate = Format(CDate(v), "dd/mm/yyyy")
    ElseIf IsNumeric(v) Then
        If CDbl(v) > 0 Then FormaterDate = Format(CDate(CDbl(v)), "dd/mm/yyyy")
    Else
        FormaterDate = mod_DataStructure.CellText(v)
    End If
End Function

' Meme categorie reservee que dans mod_ControleCategories (duplique volontairement ici :
' ce module doit pouvoir fonctionner seul, y compris pour la macro de test).
Private Function CategorieVentilePublique() As String
    CategorieVentilePublique = "Ventil" & ChrW(233)
End Function

' Traduit des codes en lettres accentuees (le fichier reste 100% ASCII).
Private Function TF(ByVal texte As String) As String
    Dim r As String
    r = texte
    r = Replace(r, "{e2}", ChrW(233))
    r = Replace(r, "{e1}", ChrW(232))
    r = Replace(r, "{ea}", ChrW(234))
    r = Replace(r, "{a2}", ChrW(224))
    TF = r
End Function
