Option Explicit

' =====================================================================================
' MODULE : mod_InstallRechercheOperations
'
' ROLE (PHASE 5a) :
'   Construit la mise en page STATIQUE (aucune logique de clic pour l'instant)
'   de la feuille masquee "frm_RechercheOperations" : un ecran de recherche
'   et de correction en masse pour TblOperations, base sur un TABLEAU EXCEL
'   CLASSIQUE avec filtre automatique natif (les fleches de filtre dans
'   l'entete font tout le travail de filtrage croise Date/Tiers/Montant/
'   Catégorie/Notes, sans code personnalise).
'
'   Colonnes du tableau (dans cet ordre) :
'     A - Valider       : l'opérateur y inscrit "Oui" sur les lignes finies
'     B - Date
'     C - Tiers
'     D - Montant
'     E - Catégorie     : liste déroulante (avertissement, pas de blocage :
'                         on peut taper une nouvelle catégorie qui n'existe
'                         pas encore)
'     F - SousCategorie : idem, PHASE 5 (catégories a 2 niveaux)
'     G - Notes         : texte libre, SAUF si la valeur est déjà une clé
'                         santé valide (grisee dans ce cas - voir Phase 5b)
'     H - Ventile        : colonne INFORMATIVE (non modifiable), PHASE 5.
'                         "Oui" si la ligne représente une PART VENTILEE
'                         d'une opération bancaire (elle vient alors de
'                         TblVentilations, pas de TblOperations), ou si
'                         l'opération PARENTE d'une ligne normale a ete
'                         ventilee (Catégorie = "Ventile"). C'est le "tag"
'                         de tracabilite demandé par l'opérateur : "on peut
'                         prevenir un tag indiquant que cette opération fait
'                         partie d'une ventilation, pour information". Une
'                         opération ventilee reste ainsi accessible ICI de 2
'                         facons : via sa ligne parente (Catégorie="Ventile"),
'                         ou directement via chacune de ses parts (une ligne
'                         par sous-catégorie de la ventilation).
'     I - ID_Transaction : colonne technique MASQUEE (ID de l'opération, ou
'                         de l'opération PARENTE pour une part ventilee)
'     J - SourceLigne    : colonne technique MASQUEE, PHASE 5 : "O" (ligne de
'                         TblOperations) ou "V" (part de TblVentilations),
'                         sert a savoir OU ecrire au moment d'appliquer
'     K - LigneVentilation : colonne technique MASQUEE, PHASE 5 : pour une
'                         ligne "V", position de la part DANS TblVentilations
'                         (DataBodyRange). Vide/non utilisée pour une ligne "O".
'
'   La Phase 5b (mod_RechercheOperations) remplira le tableau depuis
'   TblOperations ET TblVentilations (bouton "Rechercher") et appliquera les
'   lignes marquees (bouton "Appliquer les lignes marquees") dans la bonne
'   table source, colonne par colonne, jamais par un tri/decoupage de texte.
'
'   Ajout suite à un test opérateur : un 3e bouton "Revoir la ventilation"
'   permet, sur une ligne dont la colonne Ventilé vaut "Oui", de rouvrir le
'   formulaire de ventilation (frm_Ventilation) avec le détail déjà enregistré
'   dans TblVentilations, au lieu d'un formulaire vide (voir
'   mod_RechercheOperations.RevoirVentilationRO et l'explication complète en
'   tête de mod_Ventilation).
'
' =====================================================================================
' REFONTE "ECRAN CENTRAL" (PHASE 6, apres discussion avec l'operateur) :
'   Ce tableau absorbe desormais les anciens ecrans "Synthese_*" en lecture
'   seule (Budget mensuel, Erreurs sante, Dernier import, Detail d'un total),
'   qui n'offraient aucune correction. 5 colonnes REELLES supplementaires ont
'   ete ajoutees a la table pour cela :
'
'     L - Budget         : la date de "mois budgetaire" de l'operation (deja
'                         une vraie colonne de TblOperations, voir mod_ImportOFX)
'     M - StatutSante     : statut du suivi sante (deja une vraie colonne)
'     N - SoldeSante      : solde du suivi sante (deja une vraie colonne)
'     O - Date_consult    : date de consultation extraite des Notes sante
'     P - Spe_consult     : specialite extraite des Notes sante
'
'   Ces 5 colonnes ne sont PAS toujours utiles (par exemple Budget n'a pas de
'   sens en recherche libre) : mod_RechercheOperations les affiche/masque
'   dynamiquement selon le "prefiltre" demande, via DefinirColonnesVisibles()
'   ci-dessous. Elles restent neanmoins TOUJOURS PRESENTES dans le tableau
'   (juste masquees) : c'est le moyen le plus simple et le plus fiable de
'   garder un seul ListObject a colonnes fixes plutot que de le reconstruire
'   a chaque appel.
'
' À PROPOS DES ACCENTS : tout ce qui s'affiche dans Excel continue à passer par
' la fonction FR() pour rester 100% sûr à l'import VBA. Les commentaires que
' j'ajoute à partir de maintenant utilisent de vrais caractères accentués pour
' rester lisibles (convention validée avec l'opérateur) ; les anciens
' commentaires du fichier restent tels quels pour l'instant.
'
' A FAIRE POUR INSTALLER CE MODULE :
'   1. Alt+F11, Fichier > Importer un fichier..., choisir ce fichier .bas
'   2. Ctrl+G : CreerFeuilleRechercheOperations (ATTENTION : reconstruit toute la
'      feuille, y compris ses boutons – c'est le seul moyen de faire apparaître
'      le nouveau bouton "Revoir la ventilation" ; réimporter le fichier seul
'      ne suffit pas, comme déjà repéré plus tôt sur ce chantier)
'   3. Pour revoir la feuille : AfficherFeuilleRecherchePourEdition
'      Pour la remasquer : MasquerFeuilleRechercheApresEdition
'
' MISE A JOUR 01/10/2026 (ergonomie, retours operateur) :
'   - 4e bouton "Sortir" ajoute (appelle mod_RechercheOperations.
'     SortirRechercheOperations, qui existait deja mais n'etait relie a rien).
'   - Bouton "Rechercher" renomme "Recherche globale" (plus explicite : il
'     relance une recherche SANS filtre, par opposition aux recherches
'     filtrees lancees depuis d'autres ecrans).
'   - La zone de commentaire explicatif, qui etait ecrite par erreur sur la
'     MEME ligne que les boutons (donc invisible, cachee dessous), a sa
'     propre ligne maintenant (RO_LIGNE_ENTETES passe de 4 a 5).
'   - Un double-clic sur "Oui" dans la colonne Ventile ouvre desormais
'     directement le detail de la ventilation (voir ThisWorkbook.bas,
'     Workbook_SheetBeforeDoubleClick), en plus du bouton "Revoir la
'     ventilation" qui reste disponible.
'
' MISE A JOUR 02/10/2026 (apres discussion avec l'operateur) :
'   - Le bouton "Revoir la ventilation" ci-dessus est SUPPRIME (devenu inutile) :
'     double-cliquer sur une cellule de la colonne Ventile fait desormais TOUT le
'     travail, que la ligne soit deja ventilee (revoir le detail) ou non (demarrer
'     une nouvelle ventilation) -- voir mod_RechercheOperations.RevoirVentilationRO.
'   - Nouveau : double-cliquer sur une cellule de la colonne Categorie (hors ligne
'     ventilee) ouvre le MEME formulaire que le controle des categories a l'import,
'     en mode "une seule operation" -- voir mod_RechercheOperations.EditerCategorieRO
'     et mod_ControleCategories.ControlerCategories (parametre uneSeuleOperation).
'   - Le bouton "Appliquer les lignes marquees" (et la colonne Valider) ne gere plus
'     que la colonne Notes : Categorie et SousCategorie se modifient desormais par
'     double-clic (ci-dessus), qui ecrit immediatement, sans "Valider" ni ce bouton.
'   - Nouvelle ligne RO_LIGNE_FILTRE (voir plus bas) : phrase recalculee a chaque
'     recherche, qui rappelle a l'operateur sur quel sous-ensemble d'operations il
'     travaille (recherche globale, dernier import, mois precis...).
' =====================================================================================

Public Const RO_LIGNE_BOUTONS As Long = 2
' RO_LIGNE_ENTETES = 5 (et non 4) : la ligne 3, laissee libre entre les
' boutons (ligne 2) et l'entete du tableau, accueille desormais la zone de
' commentaire explicatif ci-dessous. Avant ce changement, ce commentaire
' etait ecrit sur la MEME ligne que les boutons (RO_LIGNE_ENTETES - 2 = 2) :
' invisible, cache sous les boutons eux-memes (constat operateur du 01/10/2026).
Public Const RO_LIGNE_ENTETES As Long = 5

' RO_LIGNE_FILTRE = 4 (ajout 02/10/2026) : ligne restee vide entre le texte d'aide
' (RO_LIGNE_ENTETES - 2 = 3) et l'entete du tableau (RO_LIGNE_ENTETES = 5). Accueille
' desormais une phrase courte, recalculee a chaque recherche (voir
' mod_RechercheOperations.DecrireFiltreActifRO), qui dit a l'operateur sur quel
' sous-ensemble d'operations il travaille actuellement (recherche globale, dernier
' import, mois precis, etc.).
Public Const RO_LIGNE_FILTRE As Long = RO_LIGNE_ENTETES - 1

' Position des colonnes DANS LE TABLEAU (1 = première colonne du tableau, A)
Public Const RO_COL_VALIDER As Long = 1
Public Const RO_COL_DATE As Long = 2
Public Const RO_COL_TIERS As Long = 3
Public Const RO_COL_MONTANT As Long = 4
Public Const RO_COL_CATEGORIE As Long = 5
Public Const RO_COL_SOUSCATEGORIE As Long = 6   ' PHASE 5
Public Const RO_COL_NOTES As Long = 7
Public Const RO_COL_VENTILE As Long = 8         ' PHASE 5 (informatif, non modifiable)
Public Const RO_COL_ID As Long = 9
Public Const RO_COL_SOURCE As Long = 10         ' PHASE 5 : "O" ou "V" (technique, masquee)
Public Const RO_COL_LIGNEVEN As Long = 11       ' PHASE 5 : ligne dans TblVentilations si SourceLigne="V" (technique, masquee)

' PHASE 6 : colonnes issues des anciens ecrans "Synthese_*", visibles ou non
' selon le prefiltre demande (voir DefinirColonnesVisibles plus bas)
Public Const RO_COL_BUDGET As Long = 12
Public Const RO_COL_STATUTSANTE As Long = 13
Public Const RO_COL_SOLDESANTE As Long = 14
Public Const RO_COL_DATECONSULT As Long = 15
Public Const RO_COL_SPECONSULT As Long = 16

' =====================================================================================
' MACRO D'INSTALLATION
' =====================================================================================
Sub CreerFeuilleRechercheOperations()

    Dim ws As Worksheet
    Dim reponse As VbMsgBoxResult
    Dim tbl As ListObject
    Dim plageDepart As Range

    Set ws = ObtenirFeuilleSansErreurRO(NOM_FEUILLE_RECHERCHE)

    If Not ws Is Nothing Then
        reponse = MsgBox("La feuille '" & NOM_FEUILLE_RECHERCHE & "' existe deja." & vbCrLf & _
                          "Veux-tu la reconstruire entierement (sa mise en forme actuelle sera perdue) ?", _
                          vbYesNo + vbQuestion, "Confirmation de reconstruction")
        If reponse = vbNo Then
            MsgBox "Installation annulee, aucune modification effectuee.", vbInformation
            Exit Sub
        End If
        ws.Visible = xlSheetVisible
        On Error Resume Next
        ws.ListObjects(NOM_TABLE_RECHERCHE).Delete
        On Error GoTo 0
        ws.Cells.Clear
        Call SupprimerFormesExistantesRO(ws)
        Call SupprimerNomsExistantsRO(ws, NOM_FEUILLE_RECHERCHE)
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = NOM_FEUILLE_RECHERCHE
    End If

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10

    ws.Columns("A").ColumnWidth = 10
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 24
    ws.Columns("D").ColumnWidth = 12
    ws.Columns("E").ColumnWidth = 20
    ws.Columns("F").ColumnWidth = 20
    ws.Columns("G").ColumnWidth = 40
    ws.Columns("H").ColumnWidth = 10
    ws.Columns("I").ColumnWidth = 12
    ws.Columns("J").ColumnWidth = 10
    ws.Columns("K").ColumnWidth = 10
    ws.Columns("L").ColumnWidth = 12   ' Budget
    ws.Columns("M").ColumnWidth = 14   ' StatutSante
    ws.Columns("N").ColumnWidth = 12   ' SoldeSante
    ws.Columns("O").ColumnWidth = 14   ' Date_consult
    ws.Columns("P").ColumnWidth = 16   ' Spe_consult

    ' --- Boutons ---
    Dim zoneBtn1 As Range, zoneBtn2 As Range
    Dim btn As Button

    Set zoneBtn1 = ws.Range("A" & RO_LIGNE_BOUTONS & ":B" & RO_LIGNE_BOUTONS)
    zoneBtn1.RowHeight = 22
    Set btn = ws.Buttons.Add(zoneBtn1.Left, zoneBtn1.Top, zoneBtn1.Width, zoneBtn1.Height)
    With btn
        .Caption = FR("Recherche globale")
        .OnAction = "RechercherOperations"
        .Name = "btnRechercherOperations"
    End With

    Set zoneBtn2 = ws.Range("C" & RO_LIGNE_BOUTONS & ":E" & RO_LIGNE_BOUTONS)
    Set btn = ws.Buttons.Add(zoneBtn2.Left, zoneBtn2.Top, zoneBtn2.Width, zoneBtn2.Height)
    With btn
        .Caption = FR("Appliquer les lignes marqu{e2}es")
        .OnAction = "AppliquerLignesMarquees"
        .Name = "btnAppliquerLignesMarquees"
    End With

    ' Ajout 02/10/2026 : le bouton "Revoir la ventilation" qui etait ici est supprime,
    ' devenu inutile -- le double-clic sur la colonne Ventile (voir ThisWorkbook.bas)
    ' couvre desormais les 2 cas (revoir une ventilation existante, ou en demarrer une
    ' nouvelle), voir mod_RechercheOperations.RevoirVentilationRO. Le bouton "Sortir"
    ' reprend sa place (F:G) pour ne pas laisser un espace vide.
    Dim zoneBtn4 As Range
    Set zoneBtn4 = ws.Range("F" & RO_LIGNE_BOUTONS & ":G" & RO_LIGNE_BOUTONS)
    Set btn = ws.Buttons.Add(zoneBtn4.Left, zoneBtn4.Top, zoneBtn4.Width, zoneBtn4.Height)
    With btn
        .Caption = FR("Sortir")
        .OnAction = "SortirRechercheOperations"
        .Name = "btnSortirRechercheOperations"
    End With

    ' AJOUT 03/10/2026 (demande operateur) : 2 nouveaux boutons "decalage de budget",
    ' places a droite des 3 boutons existants (colonnes H a M de la ligne des boutons,
    ' inoccupees jusqu'ici). Factorises dans AjouterBoutonsDecalageRO ci-dessous (voir
    ' ce Sub pour le detail), pour pouvoir aussi les ajouter sans tout reconstruire si
    ' la feuille existe deja (c'est d'ailleurs ce Sub qui les cree ici).
    AjouterBoutonsDecalageRO ws

    ' --- Petit rappel du fonctionnement, au-dessus du tableau ---
    ' Refonte 02/10/2026 (apres discussion avec l'operateur) : mode operatoire en
    ' liste a puces (plus lisible qu'un paragraphe), avec les noms de colonnes mis en
    ' evidence (gras + vert) via SurlignerMotsRO ci-dessous -- aucun precedent de ce
    ' genre de mise en forme (Characters) dans ce classeur, a verifier visuellement a
    ' l'import.
    Dim texteAide As String
    texteAide = "- " & FR("Utilise les fl{e2}ches de filtre dans l'en-t{ea}te de chaque colonne pour restreindre la liste affich{e2}e.") & Chr(10) & _
                "- " & FR("Corrige directement la cellule Notes, inscris 'Oui' dans Valider, puis clique sur 'Appliquer les lignes marqu{e2}es' pour l'enregistrer (possible sur plusieurs lignes {a2} la fois).") & Chr(10) & _
                "- " & FR("Double-clique sur une cellule Categorie pour l'{e2}diter, elle et sa sous-cat{e2}gorie, via le formulaire habituel de contr{o2}le des cat{e2}gories (elles ne se modifient plus via Valider/Appliquer).") & Chr(10) & _
                "- " & FR("Double-clique sur une cellule vide de la colonne Ventile pour d{e2}marrer une nouvelle ventilation sur cette op{e2}ration, ou sur une cellule 'Oui' pour revoir ou supprimer une ventilation d{e2}j{a2} enregistr{e2}e.") & Chr(10) & _
                "- " & FR("'Recherche globale' recharge toutes les op{e2}rations sans filtre. 'Sortir' referme cet {e2}cran.")

    With ws.Range("A" & (RO_LIGNE_ENTETES - 2) & ":G" & (RO_LIGNE_ENTETES - 2))
        .Merge
        .value = texteAide
        .Font.Size = 9
        .Font.Color = RGB(80, 80, 80)
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    SurlignerMotsRO ws.Range("A" & (RO_LIGNE_ENTETES - 2)), texteAide, _
                    Array("SousCategorie", "Categorie", "Notes", "Valider", "Ventile")
    ws.rows(RO_LIGNE_ENTETES - 2).RowHeight = 56

    ' --- Zone "filtre actif" (ajout 02/10/2026) : phrase courte, recalculee par
    ' mod_RechercheOperations.RechercherOperations a chaque recherche (voir
    ' DecrireFiltreActifRO), qui rappelle sur quel sous-ensemble d'operations
    ' l'operateur travaille actuellement. Vide au tout premier affichage de
    ' l'ecran (avant la toute premiere recherche).
    With ws.Range("A" & RO_LIGNE_FILTRE & ":G" & RO_LIGNE_FILTRE)
        .Merge
        .value = ""
        .Font.Size = 9
        .Font.Italic = True
        .Font.Color = RGB(31, 78, 121)
        .Interior.Color = RGB(237, 243, 250)
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    ws.rows(RO_LIGNE_FILTRE).RowHeight = 16

    ' --- Tableau (headers + 1 ligne vide de depart, indispensable pour créer un ListObject) ---
    ws.Range("A" & RO_LIGNE_ENTETES).value = "Valider"
    ws.Range("B" & RO_LIGNE_ENTETES).value = "Date"
    ws.Range("C" & RO_LIGNE_ENTETES).value = "Tiers"
    ws.Range("D" & RO_LIGNE_ENTETES).value = "Montant"
    ws.Range("E" & RO_LIGNE_ENTETES).value = "Categorie"
    ws.Range("F" & RO_LIGNE_ENTETES).value = "SousCategorie"
    ws.Range("G" & RO_LIGNE_ENTETES).value = "Notes"
    ws.Range("H" & RO_LIGNE_ENTETES).value = "Ventile"
    ws.Range("I" & RO_LIGNE_ENTETES).value = "ID_Transaction"
    ws.Range("J" & RO_LIGNE_ENTETES).value = "SourceLigne"
    ws.Range("K" & RO_LIGNE_ENTETES).value = "LigneVentilation"
    ws.Range("L" & RO_LIGNE_ENTETES).value = "Budget"
    ws.Range("M" & RO_LIGNE_ENTETES).value = "StatutSante"
    ws.Range("N" & RO_LIGNE_ENTETES).value = "SoldeSante"
    ws.Range("O" & RO_LIGNE_ENTETES).value = "Date_consult"
    ws.Range("P" & RO_LIGNE_ENTETES).value = "Spe_consult"

    Set plageDepart = ws.Range("A" & RO_LIGNE_ENTETES & ":P" & (RO_LIGNE_ENTETES + 1))
    Set tbl = ws.ListObjects.Add(xlSrcRange, plageDepart, , xlYes)
    tbl.Name = NOM_TABLE_RECHERCHE
    tbl.TableStyle = "TableStyleMedium2"

    ' Colonnes techniques masquees (PHASE 5 : Ventile reste VISIBLE, c'est le
    ' tag informatif demandé par l'opérateur -- seules I/J/K, qui ne servent
    ' qu'au code, sont masquees)
    ws.Columns("I").Hidden = True
    ws.Columns("J").Hidden = True
    ws.Columns("K").Hidden = True

    ' PHASE 6 : les 5 colonnes issues des anciens ecrans "Synthese_*" sont
    ' masquees par defaut a l'installation ; mod_RechercheOperations les
    ' affiche/masque ensuite dynamiquement a chaque appel, selon le prefiltre
    ' demande (voir DefinirColonnesVisibles ci-dessous).
    ws.Columns("L").Hidden = True
    ws.Columns("M").Hidden = True
    ws.Columns("N").Hidden = True
    ws.Columns("O").Hidden = True
    ws.Columns("P").Hidden = True

    ws.Visible = xlSheetVeryHidden

    MsgBox "La feuille '" & NOM_FEUILLE_RECHERCHE & "' a ete creee et masquee." & vbCrLf & _
           "Pour la revoir : AfficherFeuilleRecherchePourEdition", vbInformation, "Installation terminee"

End Sub


' =====================================================================================
' AJOUT 03/10/2026 (demande operateur) : ajoute les 2 boutons "decalage de budget" sur
' la feuille de recherche SI ELLE EXISTE DEJA, sans la reconstruire entierement (donc
' sans perdre les lignes/filtres en cours). Appelee automatiquement a la fin de
' CreerFeuilleRechercheOperations ci-dessus (nouvelle installation complete), et peut
' aussi etre relancee seule, par Ctrl+G, sur une feuille deja en place :
'   AjouterBoutonsDecalageRO ThisWorkbook.Worksheets("frm_RechercheOperations")
' Idempotente : si les boutons existent deja (meme nom), ne fait rien.
' =====================================================================================
Public Sub AjouterBoutonsDecalageRO(ByVal ws As Worksheet)

    Dim forme As Shape
    Dim btn As Button
    Dim zoneBtn5 As Range, zoneBtn6 As Range

    ' On verifie si l'un des 2 boutons existe deja (meme principe que les colonnes/
    ' tableaux installes ailleurs dans le projet : ne jamais recreer en double).
    On Error Resume Next
    Set forme = ws.Shapes("btnAjouterDecalageRO")
    On Error GoTo 0
    If Not forme Is Nothing Then Exit Sub

    ws.Range(RO_LIGNE_BOUTONS & ":" & RO_LIGNE_BOUTONS).RowHeight = 22

    Set zoneBtn5 = ws.Range("H" & RO_LIGNE_BOUTONS & ":J" & RO_LIGNE_BOUTONS)
    Set btn = ws.Buttons.Add(zoneBtn5.Left, zoneBtn5.Top, zoneBtn5.Width, zoneBtn5.Height)
    With btn
        .Caption = FR("Ajouter un d{e2}calage")
        .OnAction = "AjouterDecalageDepuisRO"
        .Name = "btnAjouterDecalageRO"
    End With

    Set zoneBtn6 = ws.Range("K" & RO_LIGNE_BOUTONS & ":M" & RO_LIGNE_BOUTONS)
    Set btn = ws.Buttons.Add(zoneBtn6.Left, zoneBtn6.Top, zoneBtn6.Width, zoneBtn6.Height)
    With btn
        .Caption = FR("D{e2}caler cette op{e2}ration")
        .OnAction = "DecalerBudgetOperationRO"
        .Name = "btnDecalerBudgetOperationRO"
    End With

End Sub


' =====================================================================================
' SurlignerMotsRO (ajout 02/10/2026) : met en GRAS + VERT, DANS UNE CELLULE DEJA
' REMPLIE, chaque occurrence exacte (respect de la casse) de chacun des mots de
' la liste "mots". Sert a faire ressortir les noms de colonnes (Categorie,
' SousCategorie, Notes, Valider, Ventile) dans le texte d'aide au-dessus du
' tableau de recherche, pour que l'operateur les repere en un coup d'oeil.
'
' ATTENTION (technique nouvelle dans ce classeur, aucun autre module n'utilise
' Range.Characters) : la recherche se fait avec vbBinaryCompare (respect de la
' casse), PAS vbTextCompare, pour ne jamais accrocher un mot "generique" du
' texte qui ressemblerait a un nom de colonne mais s'ecrirait differemment
' (exemple : "controle des categories", en minuscules et au pluriel dans une
' phrase normale, ne doit pas etre colore comme le nom de colonne "Categorie").
' "texteComplet" DOIT etre exactement la chaine deja ecrite dans "cellule"
' (apres toute substitution mod_Display.FR, {tag} compris) : Characters()
' raisonne en position de caractere dans le texte final affiche, pas dans un
' texte source avec des {tag}.
'
' Remarque sur l'ordre des mots : "SousCategorie" contient "Categorie". Si
' "SousCategorie" est traite APRES "Categorie" dans la liste, cela ne pose
' aucun probleme : le passage sur "SousCategorie" recolore alors l'ensemble du
' mot (y compris la partie deja coloree par "Categorie"), le resultat final
' est donc correct quel que soit l'ordre choisi.
' =====================================================================================
Private Sub SurlignerMotsRO(ByVal cellule As Range, ByVal texteComplet As String, ByVal mots As Variant)

    Dim m As Variant
    Dim motTexte As String
    Dim position As Long

    For Each m In mots
        motTexte = CStr(m)
        If Len(motTexte) > 0 Then
            position = 1
            Do
                position = InStr(position, texteComplet, motTexte, vbBinaryCompare)
                If position = 0 Then Exit Do
                With cellule.Characters(position, Len(motTexte)).Font
                    .Bold = True
                    .Color = RGB(0, 128, 0)
                End With
                position = position + Len(motTexte)
            Loop
        End If
    Next m

End Sub


' =====================================================================================
' DefinirColonnesVisibles (PHASE 6) : affiche/masque les 5 colonnes issues des
' anciens ecrans "Synthese_*" selon le prefiltre demande par l'operateur.
' Appelee par mod_RechercheOperations.RechercherOperations a chaque ouverture.
' Les colonnes A a H (Valider...Ventile) et I/J/K (techniques) ne sont JAMAIS
' concernees ici : elles restent gerees comme avant (I/J/K toujours masquees).
' =====================================================================================
Public Sub DefinirColonnesVisibles(ByVal ws As Worksheet, ByVal prefiltre As String)

    ' On repart d'un etat neutre : les 5 colonnes "Synthese_*" masquees.
    ws.Columns("L").Hidden = True   ' Budget
    ws.Columns("M").Hidden = True   ' StatutSante
    ws.Columns("N").Hidden = True   ' SoldeSante
    ws.Columns("O").Hidden = True   ' Date_consult
    ws.Columns("P").Hidden = True   ' Spe_consult

    Select Case prefiltre
        Case "DernierImport"
            ws.Columns("M").Hidden = False
            ws.Columns("N").Hidden = False
        Case "OperationsDuMois", "DetailTotal"
            ws.Columns("L").Hidden = False
        Case "ErreursSante"
            ws.Columns("O").Hidden = False
            ws.Columns("P").Hidden = False
        ' Case "" (recherche libre) : aucune des 5 colonnes n'a de sens generique,
        ' on les laisse toutes masquees (etat neutre defini plus haut).
    End Select

End Sub


' =====================================================================================
' FONCTIONS UTILITAIRES
' =====================================================================================
Private Function ObtenirFeuilleSansErreurRO(ByVal nomFeuille As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nomFeuille)
    On Error GoTo 0
    Set ObtenirFeuilleSansErreurRO = ws
End Function

Private Sub SupprimerFormesExistantesRO(ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
End Sub

Private Sub SupprimerNomsExistantsRO(ws As Worksheet, ByVal nomFeuille As String)
    Dim n As Name
    Dim i As Long
    For i = ThisWorkbook.Names.count To 1 Step -1
        Set n = ThisWorkbook.Names(i)
        On Error Resume Next
        If InStr(1, n.RefersTo, "'" & nomFeuille & "'", vbTextCompare) > 0 Then
            n.Delete
        End If
        On Error GoTo 0
    Next i
End Sub


' =====================================================================================
' OUTILS DEVELOPPEUR
' =====================================================================================
Sub AfficherFeuilleRecherchePourEdition()
    Dim ws As Worksheet
    Set ws = ObtenirFeuilleSansErreurRO(NOM_FEUILLE_RECHERCHE)
    If ws Is Nothing Then
        MsgBox "La feuille '" & mod_VarGlobales.NOM_FEUILLE_RECHERCHE & "' n'existe pas encore.", vbExclamation
        Exit Sub
    End If
    ws.Visible = xlSheetVisible
    ws.Activate
    MsgBox "La feuille est maintenant visible. Pense a la remasquer avec" & vbCrLf & _
           "MasquerFeuilleRechercheApresEdition", vbInformation
End Sub

Sub MasquerFeuilleRechercheApresEdition()
    Dim ws As Worksheet
    Set ws = ObtenirFeuilleSansErreurRO(mod_VarGlobales.NOM_FEUILLE_RECHERCHE)
    If ws Is Nothing Then
        MsgBox "La feuille '" & mod_VarGlobales.NOM_FEUILLE_RECHERCHE & "' n'existe pas.", vbExclamation
        Exit Sub
    End If
    ws.Visible = xlSheetVeryHidden
    MsgBox "La feuille est de nouveau masquee.", vbInformation
End Sub


