Option Explicit

' =====================================================================================
' MODULE : mod_Ventilation
'
' PHASE 4 du chantier "Catégorie / Sous-catégorie / Ventilation" - PARTIE 2/2
' (Version 2 : nouvelle ergonomie, à la demande de l'opérateur après les tests)
'
' RÔLE (à lire en premier, même si vous débutez) :
'   Ce module contient toute la logique du formulaire de ventilation.
'
'   FONCTIONNEMENT D'ENSEMBLE :
'   Les lignes ajoutées à la ventilation sont gardées EN MÉMOIRE pendant que le
'   formulaire est ouvert (tableaux g_VenLigneCat / g_VenLigneSous / g_VenLigneMontant),
'   et simplement AFFICHÉES dans le tableau du haut (lignes 17 à 26 de la feuille) à
'   chaque changement. La feuille elle-même ne sert que de vitrine pour ce tableau : ce
'   ne sont plus ses cellules qui sont lues lors de la validation finale, contrairement
'   à la version précédente.
'
'   Le formulaire de SAISIE (bas de la feuille) sert a la fois a AJOUTER une nouvelle
'   ligne et à MODIFIER une ligne existante :
'     - g_VenIndexEdition = 0  -> le bouton "Ajouter la ligne" crée une ligne de plus.
'     - g_VenIndexEdition = N  -> le bouton "Ajouter la ligne" REMPLACE la ligne N
'       (celle qu'on corrige, chargée via son bouton "Éditer").
'
'   "Éditer" une ligne la RETIRE IMMÉDIATEMENT du tableau et la place dans le
'   formulaire de saisie. Conséquence volontaire : si l'opérateur clique sur
'   "Éditer" puis change d'avis sans cliquer sur "Ajouter la ligne" (par exemple en
'   cliquant sur "Éditer" une AUTRE ligne ou sur "Terminer"), la ligne reste
'   supprimée. C'est le seul moyen de supprimer une ligne pour l'instant; un message
'   en informe explicitement l'opérateur à chaque fois.
'
'   RÉOUVERTURE D'UNE VENTILATION DÉJÀ VALIDÉE (ajout après un test opérateur) :
'   OuvrirVentilation() vérifie désormais si TblVentilations contient déjà des lignes
'   pour l'ID_Transaction demandé (fonction ChargerLignesExistantes, plus bas). Si oui,
'   elles sont chargées dans g_VenLigneCat/Sous/Montant AVANT l'affichage : l'opérateur
'   retrouve ses données au lieu d'un formulaire vide, que ce soit en revenant en
'   arrière PENDANT le même import (bouton "Précédent" de frm_ControleCategories) ou
'   en rouvrant plus tard depuis frm_RechercheOperations.
'   En conséquence, VenTerminer() ne se contente plus d'ajouter des lignes à la fin :
'   il supprime d'abord les anciennes lignes de cet ID_Transaction
'   (SupprimerLignesExistantes), puis réécrit la liste actuelle pour éviter les doublons.
'   ATTENTION - LIMITE CONNUE : cette suppression/réécriture réinitialise TOUJOURS les
'   colonnes de suivi santé (StatutSante, Notes, etc.) d'une ligne "Frais, remb santé"
'   à leur état initial (KO/sentinelle), même si l'opérateur l'avait déjà traitée dans
'   le suivi santé. Modifier une ventilation déjà rapprochée côté santé fait donc
'   perdre ce rapprochement, qu'il faudra refaire. Cette limite est signalée à
'   l'opérateur, mais pas encore traitée plus finement (à discuter si elle devient
'   gênante en pratique).
'
'   Ajout du 01/10/2026 (points 3, 4 et 6 du suivi des besoins) :
'     - Chaque ligne de la ventilation a désormais un champ "Notes" libre, similaire
'       au champ Notes de TblOperations. Il est écrit pour toutes les lignes, et
'       mod_FormulairesNotes.VerifierNotesSante le remplace automatiquement par la clé
'       de rapprochement si la ligne est une dépense de santé non traitée.
'     - Le montant est désormais écrit avec son signe dans TblVentilations.Montant
'       (négatif pour une dépense, positif pour un remboursement), comme dans
'       TblOperations.Montant. La saisie et l'affichage en mémoire restent positifs :
'       seules l'écriture finale (AjouterLigneVentilation) et la relecture
'       (ChargerLignesExistantes) appliquent le signe ou la valeur absolue.
'     - Un nouveau bouton "Supprimer cette ventilation", visible uniquement si elle
'       existait déjà, permet d'effacer complètement les lignes de TblVentilations
'       pour cette opération. L'appelant (mod_ControleCategories pendant un import ou
'       RevoirVentilationRO depuis la recherche) est averti via le nouveau paramètre
'       de sortie de OuvrirVentilation et doit restaurer l'ancienne catégorie et
'       sous-catégorie de l'opération (colonnes CategorieAvantVentilation et
'       SousCategorieAvantVentilation, ajoutées à TblOperations par mod_InstallVentilation).
'
' À PROPOS DES ACCENTS : tout ce qui s'affiche dans Excel (messages et valeurs de
' cellules) continue à passer par mod_Display.FR() pour rester fiable à l'import
' VBA. Les commentaires ajoutés à partir de maintenant utilisent de vrais caractères
' accentués pour rester lisibles (convention validée avec l'opérateur). Les anciens
' commentaires du fichier restent tels quels pour l'instant; le nettoyage général
' est volontairement reporté à après la mise en production.
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
Private g_VenLigneNotes(1 To 10) As String     ' Ajout 01/10/2026 : commentaire libre de la ligne

' --- Index (1..10) de la ligne en cours de modification ; 0 = aucune (mode "ajout") ---
Private g_VenIndexEdition As Long

' --- Ajout 01/10/2026 (possibilite d'annuler une ventilation) ------------------------
' g_VenEtaitDejaVentilee : vrai si ChargerLignesExistantes a trouve au moins une ligne
' a l'ouverture du formulaire -- sert a n'afficher le bouton "Supprimer cette
' ventilation" que quand il y a effectivement quelque chose a supprimer.
' g_VenSupprimee : vrai si l'operateur a confirme la suppression via ce bouton ; lu par
' OuvrirVentilation juste avant de rendre la main a l'appelant (ControleVentiler ou
' RevoirVentilationRO), qui doit alors restaurer l'ancienne categorie de l'operation.
Private g_VenEtaitDejaVentilee As Boolean
Private g_VenSupprimee As Boolean


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
' Ajout 01/10/2026 (possibilite d'annuler une ventilation) : nouveau parametre de
' SORTIE optionnel "ventilationSupprimee". L'appelant (ControleVentiler pendant un
' import, ou RevoirVentilationRO depuis la recherche) doit le lire juste apres l'appel :
' s'il revient a True, la ventilation existante a ete entierement supprimee (bouton
' "Supprimer cette ventilation") et l'appelant doit alors restaurer l'ancienne
' categorie/sous-categorie de l'operation (voir les colonnes CategorieAvantVentilation /
' SousCategorieAvantVentilation ajoutees a TblOperations par mod_InstallVentilation).
' Reste optionnel pour ne pas casser TesterVentilation, qui ne s'en sert pas.
Public Function OuvrirVentilation(ByVal idTransaction As String, ByVal dateOp As Variant, _
                                  ByVal tiers As String, ByVal libelle As String, _
                                  ByVal montantOp As Double, _
                                  ByVal categorieActuelle As String, ByVal sousCategorieActuelle As String, _
                                  Optional ByRef ventilationSupprimee As Boolean) As Boolean

    Dim ws As Worksheet
    Dim evAvant As Boolean

    ventilationSupprimee = False
    On Error GoTo Erreur

    Set ws = FeuilleSansErreur(VEN_NOM_FEUILLE)
    If ws Is Nothing Then
        MsgBox mod_Display.FR("La feuille '") & VEN_NOM_FEUILLE & mod_Display.FR("' est introuvable.") & vbCrLf & _
               mod_Display.FR("Ex{e2}cutez d'abord la macro PreparerPhase4Ventilation."), vbExclamation
        Exit Function
    End If

    g_VenIdTransaction = idTransaction
    g_VenMontantOperation = montantOp
    g_VenNomFeuillePrec = ActiveSheet.Name
    g_VenValide = False
    g_VenSupprimee = False
    g_VenNbLignes = 0
    g_VenIndexEdition = 0

    ' Si cette opération a déjà été ventilée (TblVentilations contient déjà des
    ' lignes pour cet ID_Transaction), on les recharge AVANT d'afficher la feuille,
    ' afin que l'opérateur retrouve son détail au lieu d'un formulaire vide.
    ' Voir l'explication complète en tête de module.
    ChargerLignesExistantes idTransaction

    ' Ajout 01/10/2026 : le bouton "Supprimer cette ventilation" n'a de sens que s'il y
    ' a deja quelque chose a supprimer -- on le masque donc pour une toute nouvelle
    ' ventilation (ChargerLignesExistantes n'a alors rien trouve).
    ' IMPORTANT (bug corrige le 02/10/2026) : "On Error GoTo 0" ci-dessous REMET A ZERO
    ' toute gestion d'erreur active, y compris le "On Error GoTo Erreur" pose tout en
    ' haut de cette fonction -- il ne la "restaure" PAS. Utiliser "On Error GoTo 0" ici
    ' aurait donc desactive silencieusement le filet de securite pour tout le reste de
    ' OuvrirVentilation (saisie des listes deroulantes, etc.), ce qui a ete reellement
    ' constate par l'operateur. On reactive donc explicitement "Erreur" par son nom.
    g_VenEtaitDejaVentilee = (g_VenNbLignes > 0)
    On Error Resume Next
    ws.Shapes("btnVenSupprimerVentilation").Visible = g_VenEtaitDejaVentilee
    On Error GoTo Erreur

    RemplirEntete ws, dateOp, tiers, libelle, montantOp, categorieActuelle, sousCategorieActuelle
    RafraichirAffichageLignes ws

    ' Petite aide : la saisie demarre avec la categorie actuelle de l'operation,
    ' puisqu'une ventilation reste tres souvent dans la meme categorie generale.
    evAvant = Application.EnableEvents
    Application.EnableEvents = False
    ws.Range(VEN_ADR_SAISIE_CAT).value = ""
    ws.Range(VEN_ADR_SAISIE_SOUS).value = ""
    ws.Range(VEN_ADR_SAISIE_MONTANT).value = ""
    ws.Range(VEN_ADR_SAISIE_MESSAGE).value = ""
    PoserValidationSaisieCategorie ws
    If categorieActuelle <> "" And categorieActuelle <> CategorieVentilePublique() Then
        ws.Range(VEN_ADR_SAISIE_CAT).value = categorieActuelle
        RemplirListeSousCatSaisie ws, categorieActuelle
    Else
        RemplirListeSousCatSaisie ws, ""
    End If
    Application.EnableEvents = evAvant

    RecalculerTotaux ws

    ws.Visible = xlSheetVisible
    ws.Activate
    mod_InstallCommun.MasquerQuadrillage   ' quadrillage et en-tetes toujours masques (09/10/2026)
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

    ventilationSupprimee = g_VenSupprimee
    OuvrirVentilation = g_VenValide
    Exit Function

Erreur:
    Application.EnableEvents = True
    g_VenEnCours = False
    g_VenVerrouActif = False
    MsgBox mod_Display.FR("Erreur inattendue dans la ventilation :") & vbCrLf & _
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

    ws.Range(VEN_ADR_DATE).value = FormaterDate(dateOp)
    ws.Range(VEN_ADR_TIERS).value = tiers
    ws.Range(VEN_ADR_LIBELLE).value = libelle
    ws.Range(VEN_ADR_MONTANT).value = Format(montantOp, "#,##0.00") & " " & ChrW(8364)

    catAffichee = categorieActuelle
    If sousCategorieActuelle <> "" Then catAffichee = catAffichee & " / " & sousCategorieActuelle
    If catAffichee = "" Then catAffichee = "(" & mod_Display.FR("aucune") & ")"
    ws.Range(VEN_ADR_CATACTUELLE).value = catAffichee

    ws.Range(VEN_ADR_MONTANT_A_VENTILER).value = Abs(montantOp)

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
    ws.Range(ws.Cells(VEN_LIGNE_GRILLE_DEBUT, VEN_COL_CAT), ws.Cells(VEN_LIGNE_GRILLE_FIN, VEN_COL_NOTES)).ClearContents

    For i = 1 To g_VenNbLignes
        ligne = VEN_LIGNE_GRILLE_DEBUT + i - 1
        ws.Cells(ligne, VEN_COL_CAT).value = g_VenLigneCat(i)
        ws.Cells(ligne, VEN_COL_SOUS).value = g_VenLigneSous(i)
        ws.Cells(ligne, VEN_COL_MONTANT).value = g_VenLigneMontant(i)
        ws.Cells(ligne, VEN_COL_NOTES).value = g_VenLigneNotes(i)
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

    ws.Range(VEN_ADR_TOTAL_SAISI).value = total
    ws.Range(VEN_ADR_RESTE).value = reste

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
        .ErrorTitle = mod_Display.FR("Cat{e2}gorie inconnue")
        .ErrorMessage = mod_Display.FR("Choisissez une cat{e2}gorie dans la liste, ou laissez le champ vide.") & vbCrLf & _
                        mod_Display.FR("Pour cr{e2}er une nouvelle cat{e2}gorie, utilisez le bouton '+'.")
    End With
End Sub

' Liste deroulante des sous-categories DEJA CONNUES de la categorie choisie dans le
' formulaire de saisie. Meme principe que dans le formulaire de controle (Phase 2) :
' une seule liste a gerer maintenant, ecrite dans la colonne technique cachee (Z).
Private Sub RemplirListeSousCatSaisie(ByVal ws As Worksheet, ByVal categorie As String)

    Dim liste() As String
    Dim Sortie() As Variant
    Dim nb As Long, r As Long

    ws.Range(ws.Cells(2, VEN_COL_AIDE), ws.Cells(VEN_LIGNE_AIDE_MAX, VEN_COL_AIDE)).ClearContents

    liste = mod_Categories.ObtenirSousCategories(categorie)

    If EstTableauAlloue(liste) Then
        nb = UBound(liste) - LBound(liste) + 1
        ReDim Sortie(1 To nb, 1 To 1)
        For r = 1 To nb
            Sortie(r, 1) = liste(LBound(liste) + r - 1)
        Next r
        With ws.Cells(2, VEN_COL_AIDE).Resize(nb, 1)
            .NumberFormat = "@"
            .Value2 = Sortie
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
        .ErrorTitle = mod_Display.FR("Sous-cat{e2}gorie inconnue")
        .ErrorMessage = mod_Display.FR("Cette sous-cat{e2}gorie n'existe pas pour la cat{e2}gorie choisie.") & vbCrLf & _
                        mod_Display.FR("Choisissez-en une dans la liste, laissez le champ vide, ou utilisez le bouton") & _
                        " '+' " & mod_Display.FR("pour en cr{e2}er une nouvelle.")
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
    If Target.Cells.count > 1 Then Exit Sub
    If Target.Address(False, False) <> VEN_ADR_SAISIE_CAT Then Exit Sub

    On Error GoTo Sortie
    Application.EnableEvents = False
    ws.Range(VEN_ADR_SAISIE_SOUS).ClearContents
    RemplirListeSousCatSaisie ws, mod_DataStructure.CellText(Target.value)

Sortie:
    numErr = Err.Number
    Application.EnableEvents = True
    If numErr <> 0 Then MsgBox mod_Display.FR("Erreur lors du changement de cat{e2}gorie : ") & Err.Description, vbExclamation

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

    catActuelle = mod_DataStructure.CellText(ws.Range(VEN_ADR_SAISIE_CAT).value)
    sousActuelle = mod_DataStructure.CellText(ws.Range(VEN_ADR_SAISIE_SOUS).value)

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
        ws.Range(VEN_ADR_SAISIE_CAT).value = catRes
        ws.Range(VEN_ADR_SAISIE_SOUS).value = sousRes
        RemplirListeSousCatSaisie ws, catRes
        Application.EnableEvents = True
    End If
    Exit Sub

Erreur:
    g_VenVerrouActif = True
    Application.EnableEvents = True
    MsgBox mod_Display.FR("Erreur dans VenNouvelleCategorie : ") & Err.Number & " - " & Err.Description, vbCritical

End Sub

' --- Bouton "Ajouter la ligne" ---
Public Sub VenAjouterLigne()

    Dim ws As Worksheet
    Dim cat As String, sous As String
    Dim notes As String
    Dim Montant As Variant
    Dim listeCat() As String, listeSous() As String
    Dim indiceCible As Long

    If Not g_VenEnCours Then Exit Sub
    On Error GoTo Erreur
    Set ws = ThisWorkbook.Worksheets(VEN_NOM_FEUILLE)

    cat = mod_DataStructure.CellText(ws.Range(VEN_ADR_SAISIE_CAT).value)
    sous = mod_DataStructure.CellText(ws.Range(VEN_ADR_SAISIE_SOUS).value)
    Montant = ws.Range(VEN_ADR_SAISIE_MONTANT).Value2
    ' Ajout 01/10/2026 (champ Notes) : commentaire libre, facultatif, disponible pour
    ' CHAQUE ligne de la ventilation (pas seulement les lignes "sante").
    notes = mod_DataStructure.CellText(ws.Range(VEN_ADR_SAISIE_NOTES).value)

    If cat = "" Then
        ws.Range(VEN_ADR_SAISIE_MESSAGE).value = mod_Display.FR("La cat{e2}gorie est obligatoire.")
        Exit Sub
    End If
    If Not IsNumeric(Montant) Or CDbl(Montant) <= 0 Then
        ws.Range(VEN_ADR_SAISIE_MESSAGE).value = mod_Display.FR("Le montant doit {ea}tre un nombre positif.")
        Exit Sub
    End If

    listeCat = mod_Categories.ObtenirCategories()
    If Not TrouverExact(cat, listeCat) Then
        ws.Range(VEN_ADR_SAISIE_MESSAGE).value = mod_Display.FR("La cat{e2}gorie '") & cat & mod_Display.FR("' n'existe pas.")
        Exit Sub
    End If
    cat = mod_Categories.FormeCanonique(cat, listeCat)

    If sous <> "" Then
        listeSous = mod_Categories.ObtenirSousCategories(cat)
        If Not TrouverExact(sous, listeSous) Then
            ws.Range(VEN_ADR_SAISIE_MESSAGE).value = mod_Display.FR("La sous-cat{e2}gorie '") & sous & _
                                                      mod_Display.FR("' n'existe pas pour '") & cat & "'."
            Exit Sub
        End If
        sous = mod_Categories.FormeCanonique(sous, listeSous)
    End If

    If g_VenIndexEdition = 0 Then
        ' Mode AJOUT : une ligne de plus.
        If g_VenNbLignes >= VEN_NB_LIGNES Then
            ws.Range(VEN_ADR_SAISIE_MESSAGE).value = mod_Display.FR("Limite de ") & VEN_NB_LIGNES & mod_Display.FR(" lignes atteinte.")
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
    g_VenLigneMontant(indiceCible) = CDbl(Montant)
    g_VenLigneNotes(indiceCible) = notes

    g_VenIndexEdition = 0
    RafraichirAffichageLignes ws
    RecalculerTotaux ws
    ViderChampsSaisie ws
    Exit Sub

Erreur:
    Application.EnableEvents = True
    MsgBox mod_Display.FR("Erreur dans VenAjouterLigne : ") & Err.Number & " - " & Err.Description, vbCritical

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
    MsgBox mod_Display.FR("Erreur dans VenEffacerSaisie : ") & Err.Number & " - " & Err.Description, vbCritical
End Sub

Private Sub ViderChampsSaisie(ByVal ws As Worksheet)
    Dim evAvant As Boolean
    evAvant = Application.EnableEvents
    Application.EnableEvents = False
    ws.Range(VEN_ADR_SAISIE_CAT).ClearContents
    ws.Range(VEN_ADR_SAISIE_SOUS).ClearContents
    ws.Range(VEN_ADR_SAISIE_MONTANT).ClearContents
    ws.Range(VEN_ADR_SAISIE_NOTES).ClearContents
    ws.Range(VEN_ADR_SAISIE_MESSAGE).value = ""
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
        MsgBox mod_Display.FR("Aucune ligne {a2} cette position."), vbInformation
        Exit Sub
    End If

    ' Si une autre ligne etait deja en cours de modification (retiree, jamais
    ' re-ajoutee), le signaler clairement : elle reste perdue.
    If g_VenIndexEdition <> 0 Then
        MsgBox mod_Display.FR("La ligne pr{e2}c{e2}demment retir{e2}e pour modification n'a pas {e2}t{e2} valid{e2}e : elle reste supprim{e2}e."), vbInformation
    End If

    ' On charge la ligne dans le formulaire de saisie...
    evAvant = Application.EnableEvents
    Application.EnableEvents = False
    PoserValidationSaisieCategorie ws
    ws.Range(VEN_ADR_SAISIE_CAT).value = g_VenLigneCat(indice)
    RemplirListeSousCatSaisie ws, g_VenLigneCat(indice)
    ws.Range(VEN_ADR_SAISIE_SOUS).value = g_VenLigneSous(indice)
    ws.Range(VEN_ADR_SAISIE_MONTANT).value = g_VenLigneMontant(indice)
    ws.Range(VEN_ADR_SAISIE_NOTES).value = g_VenLigneNotes(indice)
    ws.Range(VEN_ADR_SAISIE_MESSAGE).value = mod_Display.FR("Ligne retir{e2}e du tableau pour modification.") & _
        mod_Display.FR(" Cliquez sur 'Ajouter la ligne' pour la remettre (avec vos changements), sinon elle restera supprim{e2}e.")
    Application.EnableEvents = evAvant

    ' ...et on la retire IMMEDIATEMENT de la liste (voir l'explication en tete de module).
    For i = indice To g_VenNbLignes - 1
        g_VenLigneCat(i) = g_VenLigneCat(i + 1)
        g_VenLigneSous(i) = g_VenLigneSous(i + 1)
        g_VenLigneMontant(i) = g_VenLigneMontant(i + 1)
        g_VenLigneNotes(i) = g_VenLigneNotes(i + 1)
    Next i
    g_VenNbLignes = g_VenNbLignes - 1
    g_VenIndexEdition = indice

    RafraichirAffichageLignes ws
    RecalculerTotaux ws
    ws.Range(VEN_ADR_SAISIE_CAT).Select
    Exit Sub

Erreur:
    Application.EnableEvents = True
    MsgBox mod_Display.FR("Erreur dans VenEditerLigne : ") & Err.Number & " - " & Err.Description, vbCritical

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
        ws.Range(VEN_ADR_SAISIE_MESSAGE).value = _
            mod_Display.FR("Une ligne est en cours de modification : cliquez sur 'Ajouter la ligne' pour la valider, ou sur 'Effacer la saisie' pour l'abandonner, avant de terminer.")
        Exit Sub
    End If

    If g_VenNbLignes = 0 Then
        ws.Range(VEN_ADR_SAISIE_MESSAGE).value = mod_Display.FR("Aucune ligne ajout{e2}e. Renseignez au moins une cat{e2}gorie et un montant.")
        Exit Sub
    End If

    For i = 1 To g_VenNbLignes
        sommeSaisie = sommeSaisie + g_VenLigneMontant(i)
    Next i
    aVentiler = Abs(g_VenMontantOperation)

    If Abs(Round(sommeSaisie, 2) - Round(aVentiler, 2)) >= 0.005 Then
        ws.Range(VEN_ADR_SAISIE_MESSAGE).value = _
            mod_Display.FR("Le total saisi (") & Format(sommeSaisie, "#,##0.00") & mod_Display.FR(" {e2}uros) ne correspond pas au montant de l'op{e2}ration (") & _
            Format(aVentiler, "#,##0.00") & mod_Display.FR(" {e2}uros).") & vbCrLf & _
            mod_Display.FR("La somme des lignes doit {ea}tre EXACTEMENT {e2}gale. Corrigez avant de terminer.")
        Exit Sub
    End If

    ' On repart d'une TblVentilations "propre" pour cet ID_Transaction avant de
    ' réécrire la liste actuelle : sinon, rouvrir une ventilation existante pour
    ' la corriger ajouterait des lignes en double (voir l'explication complète
    ' en tête de module, y compris sa limite connue côté suivi santé).
    SupprimerLignesExistantes g_VenIdTransaction

    For i = 1 To g_VenNbLignes
        AjouterLigneVentilation g_VenIdTransaction, g_VenLigneCat(i), g_VenLigneSous(i), g_VenLigneMontant(i), _
                                g_VenLigneNotes(i), g_VenMontantOperation
    Next i

    g_VenValide = True
    FermerFeuilleVen ws
    Exit Sub

Erreur:
    Application.EnableEvents = True
    MsgBox mod_Display.FR("Erreur dans VenTerminer : ") & Err.Number & " - " & Err.Description, vbCritical

End Sub

' --- Bouton "Annuler" (global) ---
Public Sub VenAnnuler()

    Dim ws As Worksheet
    Dim reponse As VbMsgBoxResult

    If Not g_VenEnCours Then Exit Sub
    On Error GoTo Erreur

    reponse = MsgBox(mod_Display.FR("Annuler toute la ventilation ? Rien ne sera enregistr{e2}."), vbYesNo + vbQuestion, mod_Display.FR("Confirmation"))
    If reponse = vbNo Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(VEN_NOM_FEUILLE)
    g_VenValide = False
    FermerFeuilleVen ws
    Exit Sub

Erreur:
    Application.EnableEvents = True
    MsgBox mod_Display.FR("Erreur dans VenAnnuler : ") & Err.Number & " - " & Err.Description, vbCritical

End Sub

' --- Bouton "Supprimer cette ventilation" (global - ajout 01/10/2026, point 4) ------
' Ne s'affiche (voir OuvrirVentilation) que si cette ventilation existait DEJA avant
' l'ouverture du formulaire (g_VenEtaitDejaVentilee) : il n'y a sinon rien a supprimer.
' Contrairement a "Annuler" (qui abandonne la SAISIE en cours sans rien changer a
' TblVentilations), ce bouton supprime definitivement les lignes deja enregistrees pour
' cette operation, et previent l'appelant (via le parametre de sortie de
' OuvrirVentilation) qu'il doit restaurer l'ancienne categorie de l'operation.
Public Sub VenSupprimerVentilation()

    Dim ws As Worksheet
    Dim reponse As VbMsgBoxResult
    Dim avertissementSante As String

    If Not g_VenEnCours Then Exit Sub
    If Not g_VenEtaitDejaVentilee Then Exit Sub   ' garde-fou : le bouton est normalement masque dans ce cas
    On Error GoTo Erreur

    avertissementSante = ""
    If VentilationContientLigneSanteDejaTraitee(g_VenIdTransaction) Then
        avertissementSante = vbCrLf & vbCrLf & _
            mod_Display.FR("ATTENTION : au moins une de ses lignes est une d{e2}pense de sant{e2} d{e2}j{a2} trait{e2}e dans le suivi sant{e2}.") & _
            " " & mod_Display.FR("Ce rapprochement sera {e2}galement perdu, d{e2}finitivement.")
    End If

    reponse = MsgBox(mod_Display.FR("Supprimer compl{e2}tement cette ventilation ?") & vbCrLf & _
                      mod_Display.FR("Toutes ses lignes seront effac{e2}es, et l'op{e2}ration reprendra sa cat{e2}gorie d'origine.") & _
                      avertissementSante, _
                      vbYesNo + vbExclamation, mod_Display.FR("Confirmation de suppression"))
    If reponse = vbNo Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(VEN_NOM_FEUILLE)
    SupprimerLignesExistantes g_VenIdTransaction

    g_VenValide = False
    g_VenSupprimee = True
    FermerFeuilleVen ws
    Exit Sub

Erreur:
    Application.EnableEvents = True
    MsgBox mod_Display.FR("Erreur dans VenSupprimerVentilation : ") & Err.Number & " - " & Err.Description, vbCritical

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
' Montant (ajout 01/10/2026, point 6) : le PARAMETRE "Montant" recu ici reste POSITIF
' (c'est ce que l'operateur saisit et ce qui est garde en memoire pendant l'edition,
' voir VenAjouterLigne). On applique le signe de l'operation PARENTE (montantOperationSigne)
' uniquement au moment de l'ECRITURE dans TblVentilations, pour que la colonne Montant y
' soit SIGNEE exactement comme TblOperations.Montant (negatif = depense, positif =
' remboursement) : c'est ce qu'attendent deja mod_SuiviSante et la mise en forme
' vert/rouge de frm_RechercheOperations.
'
' Notes (ajout 01/10/2026, point 3) : le commentaire libre saisi par l'operateur est
' desormais ecrit pour TOUTE ligne, quelle que soit sa categorie (ce n'est plus reserve
' aux lignes de sante). Si la ligne est elle-meme "Frais, remb sante" et pas encore
' rapprochee, mod_FormulairesNotes.VerifierNotesSante la detectera et remplacera alors
' cette valeur par la cle technique de rapprochement, exactement comme il le fait deja
' pour une ligne de TblOperations -- ce module prend deja en charge TblVentilations
' (voir mod_FormulairesNotes.bas, PHASE 6), il n'y a donc rien d'autre a faire ici.
'
' Les AUTRES colonnes de suivi sante (Date_consult, StatutSante...) restent remplies
' QUE si cette ligne est elle-meme categorisee "Frais, remb sante" : c'est la seule
' sous-categorie suivie par le moteur sante existant (mod_SuiviSante). Une ligne
' sante est initialisee exactement comme le fait mod_ImportOFX pour une operation
' fraichement importee dont le libelle n'a pas pu etre decode automatiquement
' (StatutSante = "KO", Date_consult = la "sentinelle" 02/01/1900).
Private Sub AjouterLigneVentilation(ByVal idTransaction As String, ByVal categorie As String, _
                                    ByVal sousCategorie As String, ByVal Montant As Double, _
                                    ByVal notes As String, ByVal montantOperationSigne As Double)

    Dim wsData As Worksheet
    Dim tbl As ListObject
    Dim ligne As Long
    Dim estSante As Boolean
    Dim montantSigne As Double

    Set wsData = ThisWorkbook.Worksheets(VEN_NOM_FEUILLE_DONNEES)
    Set tbl = wsData.ListObjects(VEN_NOM_TABLE)

    ligne = tbl.ListRows.Add.Range.Row

    montantSigne = Montant * Sgn(montantOperationSigne)

    EcrireCelluleTexte wsData, tbl, ligne, "ID_Transaction", idTransaction
    EcrireCelluleTexte wsData, tbl, ligne, "Categorie", categorie
    EcrireCelluleTexte wsData, tbl, ligne, "SousCategorie", sousCategorie
    EcrireCelluleNombre wsData, tbl, ligne, "Montant", montantSigne, "#,##0.00"
    EcrireCelluleDate wsData, tbl, ligne, "DateVentilation", Now, "dd/mm/yyyy hh:mm"
    EcrireCelluleTexte wsData, tbl, ligne, "Notes", notes

    estSante = (mod_Categories.NormaliserTexte(sousCategorie) = mod_Categories.NormaliserTexte(SousCategorieSanteReference()))

    If estSante Then
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

' =====================================================================================
' RÉOUVERTURE D'UNE VENTILATION EXISTANTE (ajout après un test opérateur)
' =====================================================================================
' Relit TblVentilations et charge dans g_VenLigneCat/Sous/Montant toutes les lignes
' déjà enregistrées pour cet ID_Transaction, dans leur ordre d'apparition dans le
' tableau. Ne fait rien (g_VenNbLignes reste à 0) si aucune ligne n'est trouvée :
' c'est le cas normal d'une première ventilation, qui doit démarrer à vide.
Private Sub ChargerLignesExistantes(ByVal idTransaction As String)

    Dim wsData As Worksheet
    Dim tbl As ListObject
    Dim colID As Long, colCat As Long, colSous As Long, colMontant As Long, colNotes As Long
    Dim r As Long
    Dim idLigne As String

    Set wsData = FeuilleSansErreur(VEN_NOM_FEUILLE_DONNEES)
    If wsData Is Nothing Then Exit Sub          ' Phase 4 pas encore installée : rien à charger

    On Error Resume Next
    Set tbl = wsData.ListObjects(VEN_NOM_TABLE)
    On Error GoTo 0
    If tbl Is Nothing Then Exit Sub
    If tbl.ListRows.count = 0 Then Exit Sub

    colID = tbl.ListColumns("ID_Transaction").index
    colCat = tbl.ListColumns("Categorie").index
    colSous = tbl.ListColumns("SousCategorie").index
    colMontant = tbl.ListColumns("Montant").index
    colNotes = 0
    On Error Resume Next
    colNotes = tbl.ListColumns("Notes").index
    On Error GoTo 0

    For r = 1 To tbl.ListRows.count
        idLigne = mod_DataStructure.CellText(tbl.DataBodyRange.Cells(r, colID).value)
        If idLigne = idTransaction Then
            If g_VenNbLignes >= VEN_NB_LIGNES Then Exit For   ' garde-fou, ne devrait jamais arriver
            g_VenNbLignes = g_VenNbLignes + 1
            g_VenLigneCat(g_VenNbLignes) = mod_DataStructure.CellText(tbl.DataBodyRange.Cells(r, colCat).value)
            g_VenLigneSous(g_VenNbLignes) = mod_DataStructure.CellText(tbl.DataBodyRange.Cells(r, colSous).value)
            ' Ajout 01/10/2026 (point 6, montant signe) : TblVentilations stocke
            ' desormais le montant avec le signe de l'operation d'origine (voir
            ' AjouterLigneVentilation) -- on reprend ici la valeur ABSOLUE en memoire,
            ' puisque le formulaire de saisie continue de n'afficher/demander que des
            ' montants positifs (voir VenAjouterLigne).
            g_VenLigneMontant(g_VenNbLignes) = Abs(mod_DataStructure.ToDouble(tbl.DataBodyRange.Cells(r, colMontant).value))
            If colNotes <> 0 Then
                g_VenLigneNotes(g_VenNbLignes) = mod_DataStructure.CellText(tbl.DataBodyRange.Cells(r, colNotes).value)
            Else
                g_VenLigneNotes(g_VenNbLignes) = ""
            End If
        End If
    Next r

End Sub

' Vrai si au moins une ligne de TblVentilations pour cet ID_Transaction est une depense
' de sante dont le rapprochement a deja ete fait (Date_consult n'est plus la sentinelle
' 02/01/1900). Sert uniquement a renforcer le message d'avertissement affiche avant la
' suppression d'une ventilation (VenSupprimerVentilation) : ce rapprochement serait
' perdu en meme temps que la ligne.
Private Function VentilationContientLigneSanteDejaTraitee(ByVal idTransaction As String) As Boolean

    Dim wsData As Worksheet
    Dim tbl As ListObject
    Dim colID As Long, colSous As Long, colDateConsult As Long
    Dim r As Long
    Dim idLigne As String, sousLigne As String
    Dim dateConsultVal As Variant

    VentilationContientLigneSanteDejaTraitee = False

    Set wsData = FeuilleSansErreur(VEN_NOM_FEUILLE_DONNEES)
    If wsData Is Nothing Then Exit Function

    On Error Resume Next
    Set tbl = wsData.ListObjects(VEN_NOM_TABLE)
    On Error GoTo 0
    If tbl Is Nothing Then Exit Function
    If tbl.ListRows.count = 0 Then Exit Function

    colID = tbl.ListColumns("ID_Transaction").index
    colSous = 0
    colDateConsult = 0
    On Error Resume Next
    colSous = tbl.ListColumns("SousCategorie").index
    colDateConsult = tbl.ListColumns("Date_consult").index
    On Error GoTo 0
    If colSous = 0 Or colDateConsult = 0 Then Exit Function

    For r = 1 To tbl.ListRows.count
        idLigne = mod_DataStructure.CellText(tbl.DataBodyRange.Cells(r, colID).value)
        If idLigne = idTransaction Then
            sousLigne = mod_DataStructure.CellText(tbl.DataBodyRange.Cells(r, colSous).value)
            If mod_Categories.NormaliserTexte(sousLigne) = mod_Categories.NormaliserTexte(SousCategorieSanteReference()) Then
                dateConsultVal = tbl.DataBodyRange.Cells(r, colDateConsult).value
                If IsDate(dateConsultVal) Then
                    If CDate(dateConsultVal) <> DateSerial(1900, 1, 2) Then
                        VentilationContientLigneSanteDejaTraitee = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next r

End Function

' Supprime de TblVentilations toutes les lignes déjà enregistrées pour cet
' ID_Transaction. Appelée par VenTerminer juste avant de réécrire la liste actuelle,
' afin que "modifier une ventilation" ne crée pas de doublons. La boucle parcourt
' les lignes à l'ENVERS (de la dernière à la première) : c'est une règle de base
' en VBA lorsqu'on supprime les lignes d'un tableau dans une boucle; sinon, les
' numéros des lignes restantes se décalent et certaines lignes sont ignorées.
Private Sub SupprimerLignesExistantes(ByVal idTransaction As String)

    Dim wsData As Worksheet
    Dim tbl As ListObject
    Dim colID As Long
    Dim r As Long
    Dim idLigne As String

    Set wsData = FeuilleSansErreur(VEN_NOM_FEUILLE_DONNEES)
    If wsData Is Nothing Then Exit Sub

    On Error Resume Next
    Set tbl = wsData.ListObjects(VEN_NOM_TABLE)
    On Error GoTo 0
    If tbl Is Nothing Then Exit Sub
    If tbl.ListRows.count = 0 Then Exit Sub

    colID = tbl.ListColumns("ID_Transaction").index

    ' Ajout 01/10/2026 (bug operateur) : si l'operation revisitee est la SEULE
    ' ventilation presente dans TblVentilations, la boucle ci-dessous vide
    ' temporairement tout le tableau. Excel declenche alors une alerte native
    ' ("Voulez-vous supprimer la ligne entiere ?") qui met la macro EN PAUSE en
    ' plein milieu de la suppression -- et c'est cette pause qui corrompait les
    ' donnees au redemarrage (ID_Transaction/SourceLigne vides, #N/A affiche
    ' dans LigneVentilation en relancant la recherche). On desactive ces alertes
    ' le temps de la suppression, exactement comme deja fait ailleurs dans ce
    ' classeur pour la meme raison (voir mod_Actions.bas, ActionSortir, qui
    ' supprime une feuille). On restaure ensuite l'etat d'origine plutot que de
    ' forcer True, au cas ou l'appelant les avait deja lui-meme desactivees.
    Dim alertesAvant As Boolean
    alertesAvant = Application.DisplayAlerts
    Application.DisplayAlerts = False

    For r = tbl.ListRows.count To 1 Step -1
        idLigne = mod_DataStructure.CellText(tbl.DataBodyRange.Cells(r, colID).value)
        If idLigne = idTransaction Then
            tbl.ListRows(r).Delete
        End If
    Next r

    Application.DisplayAlerts = alertesAvant

End Sub

Private Sub EcrireCelluleTexte(ByVal ws As Worksheet, ByVal tbl As ListObject, ByVal ligne As Long, _
                               ByVal nomColonne As String, ByVal valeur As String)
    Dim col As Long
    col = tbl.ListColumns(nomColonne).index
    ws.Cells(ligne, col).NumberFormat = "@"
    ws.Cells(ligne, col).value = valeur
End Sub

Private Sub EcrireCelluleNombre(ByVal ws As Worksheet, ByVal tbl As ListObject, ByVal ligne As Long, _
                                ByVal nomColonne As String, ByVal valeur As Double, ByVal formatNombre As String)
    Dim col As Long
    col = tbl.ListColumns(nomColonne).index
    ws.Cells(ligne, col).NumberFormat = formatNombre
    ws.Cells(ligne, col).value = valeur
End Sub

Private Sub EcrireCelluleDate(ByVal ws As Worksheet, ByVal tbl As ListObject, ByVal ligne As Long, _
                              ByVal nomColonne As String, ByVal valeur As Date, ByVal formatDate As String)
    Dim col As Long
    col = tbl.ListColumns(nomColonne).index
    ws.Cells(ligne, col).NumberFormat = formatDate
    ws.Cells(ligne, col).value = valeur
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

    ok = OuvrirVentilation(idTest, Date, mod_Display.FR("Tiers de test"), mod_Display.FR("Op{e2}ration fictive pour tester le formulaire"), _
                           -100, mod_Display.FR("Sant{e2}, pr{e2}voyance"), "")

    If ok Then
        MsgBox mod_Display.FR("Ventilation enregistr{e2}e dans TblVentilations sous l'identifiant '") & idTest & "'." & vbCrLf & _
               mod_Display.FR("Vous pouvez la retrouver (et la supprimer) sur la feuille '") & VEN_NOM_FEUILLE_DONNEES & "'.", _
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



