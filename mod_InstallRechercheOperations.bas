Option Explicit

' =====================================================================================
' MODULE : mod_InstallRechercheOperations
'
' ROLE (PHASE 5a) :
'   Construit la mise en page de la feuille masquée "frm_RechercheOperations" : un écran
'   de recherche et de correction en masse pour TblOperations, basé sur un TABLEAU EXCEL
'   CLASSIQUE avec filtre automatique natif (les flèches de filtre dans l'en-tête font
'   tout le travail de filtrage croisé Date/Tiers/Montant/Catégorie/Notes, sans code
'   personnalisé).
'
'   Colonnes du tableau (dans cet ordre) :
'     A - Valider       : l'opérateur y inscrit "Oui" sur les lignes finies
'     B - Date
'     C - Tiers
'     D - Montant
'     E - Catégorie     : liste déroulante (avertissement, pas de blocage :
'                         on peut taper une nouvelle catégorie qui n'existe
'                         pas encore). Modifiable par double-clic (voir ci-dessous).
'     F - SousCategorie : idem, PHASE 5 (catégories à 2 niveaux)
'     G - Notes         : texte libre, SAUF si la valeur est déjà une clé
'                         santé valide (grisée dans ce cas - voir Phase 5b)
'     H - Ventile        : colonne INFORMATIVE (non modifiable), PHASE 5.
'                         "Oui" si la ligne représente une PART VENTILÉE
'                         d'une opération bancaire (elle vient alors de
'                         TblVentilations, pas de TblOperations), ou si
'                         l'opération PARENTE d'une ligne normale a été
'                         ventilée (Catégorie = "Ventile").
'     I - ID_Transaction : colonne technique MASQUÉE (ID de l'opération, ou
'                         de l'opération PARENTE pour une part ventilée)
'     J - SourceLigne    : colonne technique MASQUÉE, PHASE 5 : "O" (ligne de
'                         TblOperations) ou "V" (part de TblVentilations),
'                         sert à savoir où écrire au moment d'appliquer
'     K - LigneVentilation : colonne technique MASQUÉE, PHASE 5 : pour une
'                         ligne "V", position de la part DANS TblVentilations
'                         (DataBodyRange). Vide/non utilisée pour une ligne "O".
'   PHASE 6 (refonte "écran central") : 5 colonnes RÉELLES supplémentaires, affichées
'   ou masquées dynamiquement selon le "préfiltre" demandé (DefinirColonnesVisibles) :
'     L - Budget / M - StatutSante / N - SoldeSante / O - Date_consult / P - Spe_consult
'
'   Ce module ne construit QUE ce qui est fixe : l'en-tête commun (boutons, compteur,
'   Instructions) et la ligne d'en-têtes du tableau. Les lignes du tableau elles-mêmes
'   sont remplies, mises en forme et colorées à chaque ouverture par
'   mod_RechercheOperations.RechercherOperations (gris = non modifiable, jaune = modifiable).
'
' REFONTE DU 09/10/2026 (harmonisation des formulaires) :
'   - Construite avec les outils COMMUNS de mod_InstallCommun (boutons, compteur, bloc
'     Instructions, étiquettes). Le texte d'aide en puces (avec surlignage des noms de
'     colonnes) est remplacé par le bloc Instructions commun, placé SOUS la ligne de
'     boutons ; la fonction SurlignerMotsRO, devenue inutile, est supprimée.
'   - La ligne RO_LIGNE_FILTRE porte désormais le compteur au format commun
'     "Opération: x/y" (cellule RO_ADR_COMPTEUR) et, à sa droite, la phrase qui rappelle
'     le filtre actif (cellule RO_ADR_FILTRE).
'   - Le tableau n'a plus de style Excel (bandes bleues) : ses couleurs suivent la charte
'     commune. L'en-tête du tableau passe de la ligne 7 à la ligne 8 (RO_LIGNE_ENTETES).
'   - AjouterBoutonsDecalageRO est supprimée : les 2 boutons "décalage" font partie de la
'     construction normale.
'   Les largeurs des colonnes et des boutons sont celles relevées sur la feuille réelle
'   (export frm_RechercheOperations_structure.txt).
'
' À FAIRE POUR INSTALLER CE MODULE :
'   1. Alt+F11, Fichier > Importer un fichier... : mod_InstallCommun.bas, puis ce fichier
'      et mod_RechercheOperations.bas (mis à jour).
'   2. Ctrl+G : CreerFeuilleRechercheOperations (reconstruit toute la feuille, y compris
'      ses boutons; réimporter le fichier seul ne suffit pas).
'   3. Pour revoir la feuille : AfficherFeuilleRecherchePourEdition
'      Pour la remasquer : MasquerFeuilleRechercheApresEdition
' =====================================================================================

' En-tête commun (voir mod_InstallCommun) : 1 ligne de boutons + compteur.
'   ligne 1 marge / 2 boutons / 3 fine ligne / 4 compteur et filtre actif / 5 fine ligne /
'   6 Instructions / 7 fine ligne / 8 en-têtes du tableau / 9 et suivantes : données
Public Const RO_LIGNE_BOUTONS As Long = 2
Public Const RO_LIGNE_FILTRE As Long = 4
Public Const RO_ADR_COMPTEUR As String = "A4"     ' "Opération: 192/195" (fusion A4:C4)
Public Const RO_ADR_FILTRE As String = "D4"       ' phrase du filtre actif (fusion D4:H4)
Private Const RO_LIGNE_INSTRUCTIONS As Long = 6
Public Const RO_LIGNE_ENTETES As Long = 8

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
Public Const RO_COL_SOURCE As Long = 10         ' PHASE 5 : "O" ou "V" (technique, masquée)
Public Const RO_COL_LIGNEVEN As Long = 11       ' PHASE 5 : ligne dans TblVentilations si SourceLigne="V" (technique, masquée)

' PHASE 6 : colonnes issues des anciens écrans "Synthese_*", visibles ou non
' selon le préfiltre demandé (voir DefinirColonnesVisibles plus bas)
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
    Dim wsPrecedente As Worksheet

    Set wsPrecedente = ActiveSheet

    Set ws = mod_InstallCommun.PreparerFeuille(NOM_FEUILLE_RECHERCHE, True, True)
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    ' Largeurs relevées sur la feuille réelle (A passe de 9 à 11 pour loger l'étiquette
    ' "Instructions"). Colonnes I à K et L à P : techniques ou affichées selon le
    ' préfiltre (masquées plus bas); Q : libre; R : liste technique des catégories.
    mod_InstallCommun.MettreEnForme ws, Array(11, 9.07, 28.27, 8.93, 16.8, 24.67, 86.8, 7.87, _
                                              12, 10, 10, 12, 14, 12, 14, 16, 10.13, 28), RO_LIGNE_ENTETES

    RO_ConstruireEntete ws
    RO_ConstruireEnteteTableau ws
    RO_ConstruireBoutons ws     ' en dernier : les boutons se placent d'après les hauteurs de lignes

    ' Colonnes techniques masquées (PHASE 5 : Ventile reste VISIBLE, c'est le tag
    ' informatif demandé par l'opérateur -- seules I/J/K, qui ne servent qu'au code,
    ' sont masquées).
    ws.Range("I:K").EntireColumn.Hidden = True

    ' PHASE 6 : les 5 colonnes issues des anciens écrans "Synthese_*" sont masquées par
    ' défaut à l'installation ; mod_RechercheOperations les affiche/masque ensuite
    ' dynamiquement à chaque appel, selon le préfiltre demandé (DefinirColonnesVisibles).
    ws.Range("L:P").EntireColumn.Hidden = True

    ' Liste technique des catégories (colonne R), alimentée à chaque recherche.
    ws.Columns("R").Hidden = True

    mod_InstallCommun.TerminerFeuille ws, wsPrecedente
    Application.ScreenUpdating = True

    MsgBox "La feuille '" & NOM_FEUILLE_RECHERCHE & mod_Display.FR("' a {e2}t{e2} cr{e2}{e2}e et masqu{e2}e.") & vbCrLf & _
           "Pour la revoir : AfficherFeuilleRecherchePourEdition", vbInformation, mod_Display.FR("Installation termin{e2}e")

End Sub


' =====================================================================================
' En-tête commun : hauteurs, compteur / filtre actif et bloc Instructions
' =====================================================================================
Private Sub RO_ConstruireEntete(ByVal ws As Worksheet)

    If mod_InstallCommun.PremiereLigneCorps(1, True) <> RO_LIGNE_ENTETES Or _
       mod_InstallCommun.LigneInstructions(1, True) <> RO_LIGNE_INSTRUCTIONS Or _
       mod_InstallCommun.LigneCompteur(1, True) <> RO_LIGNE_FILTRE Then
        MsgBox "mod_InstallRechercheOperations : constantes de lignes incoherentes avec mod_InstallCommun.", vbCritical
    End If

    mod_InstallCommun.PoserHauteursEntete ws, 1, True

    ' Compteur "Opération: x/y" (A4:C4), rempli par mod_RechercheOperations.RechercherOperations
    mod_InstallCommun.PoserCompteur ws, RO_LIGNE_FILTRE, 1, 3

    ' Phrase du filtre actif (D4:H4), recalculée à chaque recherche (voir
    ' mod_RechercheOperations.DecrireFiltreActifRO) : elle rappelle à l'opérateur sur quel
    ' sous-ensemble d'opérations il travaille. Vide avant la toute première recherche.
    With ws.Range("D" & RO_LIGNE_FILTRE & ":H" & RO_LIGNE_FILTRE)
        .Merge
        .Value = ""
        .Font.Size = 9
        .Font.Italic = True
        .Font.Bold = False
        .Font.Color = mod_InstallCommun.CoulEtiquette()
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    ' Étiquette en colonne A, texte sur B:H (les colonnes I à P sont masquées).
    mod_InstallCommun.EcrireInstructions ws, RO_LIGNE_INSTRUCTIONS, 1, 1, 2, 8, RO_TexteInstructions()

End Sub

Private Function RO_TexteInstructions() As String

    Dim t As String

    t = "Cet {e2}cran permet de rechercher et de corriger les op{e2}rations enregistr{e2}es, ainsi que leurs lignes de ventilation. "
    t = t & "Le tableau est charg{e2} {a2} l'ouverture (recherche globale, dernier import, mois budg{e2}taire, suivi sant{e2}...) ; "
    t = t & "la ligne du dessus rappelle le filtre actif et le nombre de lignes affich{e2}es."
    t = t & Chr(10) & "Utilisez les fl{e1}ches de filtre de l'en-t{ea}te de chaque colonne pour restreindre la liste affich{e2}e. "
    t = t & "Corrigez directement la cellule [c:Notes], inscrivez Oui dans [c:Valider], puis cliquez sur [b:Appliquer les lignes marqu{e2}es] "
    t = t & "(possible sur plusieurs lignes {a2} la fois)."
    t = t & Chr(10) & "Double-cliquez sur une cellule [c:Categorie] pour l'{e2}diter, elle et sa [c:SousCategorie], via le formulaire [f:frm_ControleCategories] "
    t = t & "(elles ne se modifient plus avec [c:Valider]). Double-cliquez sur une cellule vide de la colonne [c:Ventile] pour d{e2}marrer une nouvelle "
    t = t & "ventilation sur cette op{e2}ration ([f:frm_Ventilation]), ou sur une cellule Oui pour revoir ou supprimer une ventilation d{e2}j{a2} enregistr{e2}e."
    t = t & Chr(10) & "Boutons :"
    t = t & Chr(10) & "- [b:Recherche globale] : recharge toutes les op{e2}rations sans filtre."
    t = t & Chr(10) & "- [b:Appliquer les lignes marqu{e2}es] : pour chaque ligne dont [c:Valider] contient Oui, enregistre le champ [c:Notes] en contr{o2}lant la saisie."
    t = t & Chr(10) & "- [b:Sortir] : referme cet {e2}cran et revient sur [f:Synthese]."
    t = t & Chr(10) & "- [b:Ajouter un d{e2}calage] : ajoute un d{e2}calage de budget r{e2}current au syst{e2}me."
    t = t & Chr(10) & "- [b:D{e2}caler cette op{e2}ration] : d{e2}calage ponctuel d'une op{e2}ration ; le calcul repart toujours de la date de l'op{e2}ration "
    t = t & "(pour revenir au mois d'origine, indiquez 0)."
    t = t & Chr(10) & "- [b:Rendre r{e2}currente] : ouvre le formulaire d'{e2}ch{e2}ance pr{e2}rempli avec l'op{e2}ration s{e2}lectionn{e2}e (tiers, cat{e2}gorie, montant, "
    t = t & "p{e2}riodicit{e2} mensuelle, prochaine date {a2} venir). Vous pouvez tout modifier avant de valider ; l'op{e2}ration elle-m{ea}me n'est pas modifi{e2}e."

    RO_TexteInstructions = t

End Function


' =====================================================================================
' En-tête du tableau (seule partie fixe du tableau)
' =====================================================================================
' Le tableau est vidé et reconstruit à chaque ouverture par mod_RechercheOperations;
' seule sa ligne d'en-têtes (et une première ligne vide, indispensable pour créer un
' ListObject) est posée ici.
Private Function RO_ConstruireEnteteTableau(ByVal ws As Worksheet) As ListObject

    Dim tbl As ListObject
    Dim noms As Variant
    Dim i As Long

    noms = Array("Valider", "Date", "Tiers", "Montant", "Categorie", "SousCategorie", "Notes", "Ventile", _
                 "ID_Transaction", "SourceLigne", "LigneVentilation", "Budget", "StatutSante", "SoldeSante", _
                 "Date_consult", "Spe_consult")

    For i = 0 To UBound(noms)
        ws.Cells(RO_LIGNE_ENTETES, i + 1).Value = noms(i)
    Next i

    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(RO_LIGNE_ENTETES, 1), ws.Cells(RO_LIGNE_ENTETES + 1, 16)), , xlYes)
    tbl.Name = NOM_TABLE_RECHERCHE
    tbl.TableStyle = ""           ' pas de style Excel : les couleurs suivent la charte commune

    mod_InstallCommun.PoserEnteteTableau tbl.HeaderRowRange
    ws.Rows(RO_LIGNE_ENTETES).RowHeight = mod_InstallCommun.FRM_H_LIGNE

    Set RO_ConstruireEnteteTableau = tbl

End Function


' =====================================================================================
' Boutons
' =====================================================================================
' Largeurs relevées sur la feuille réelle (export du 09/10/2026), hauteur et alignement
' communs (voir mod_InstallCommun.AjouterBoutonEntete).
Private Sub RO_ConstruireBoutons(ByVal ws As Worksheet)

    Dim gauche As Double

    gauche = ws.Cells(1, 1).Left

    mod_InstallCommun.AjouterBoutonEntete ws, RO_LIGNE_BOUTONS, gauche, mod_Display.FR("Recherche globale"), _
        "RechercherOperations", "btnRechercherOperations", 108
    mod_InstallCommun.AjouterBoutonEntete ws, RO_LIGNE_BOUTONS, gauche, mod_Display.FR("Appliquer les lignes marqu{e2}es"), _
        "AppliquerLignesMarquees", "btnAppliquerLignesMarquees", 154.5
    mod_InstallCommun.AjouterBoutonEntete ws, RO_LIGNE_BOUTONS, gauche, mod_InstallCommun.CapSortir(), _
        "SortirRechercheOperations", "btnSortirRechercheOperations", mod_InstallCommun.FRM_BTN_L
    mod_InstallCommun.AjouterBoutonEntete ws, RO_LIGNE_BOUTONS, gauche, mod_Display.FR("Ajouter un d{e2}calage"), _
        "AjouterDecalageDepuisRO", "btnAjouterDecalageRO", 104.2
    mod_InstallCommun.AjouterBoutonEntete ws, RO_LIGNE_BOUTONS, gauche, mod_Display.FR("D{e2}caler cette op{e2}ration"), _
        "DecalerBudgetOperationRO", "btnDecalerBudgetOperationRO", 117
    ' Ajout 09/10/2026 (échéancier) : déclare l'opération sélectionnée comme récurrente.
    mod_InstallCommun.AjouterBoutonEntete ws, RO_LIGNE_BOUTONS, gauche, mod_Display.FR("Rendre r{e2}currente"), _
        "RendreRecurrenteRO", "btnRendreRecurrenteRO", 105

End Sub


' =====================================================================================
' DefinirColonnesVisibles (PHASE 6) : affiche/masque les 5 colonnes issues des
' anciens écrans "Synthese_*" selon le préfiltre demandé par l'opérateur.
' Appelée par mod_RechercheOperations.RechercherOperations à chaque ouverture.
' Les colonnes A à H (Valider...Ventile) et I/J/K (techniques) ne sont JAMAIS
' concernées ici : elles restent gérées comme avant (I/J/K toujours masquées).
' =====================================================================================
Public Sub DefinirColonnesVisibles(ByVal ws As Worksheet, ByVal prefiltre As String)

    ' On repart d'un état neutre : les 5 colonnes "Synthese_*" masquées.
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
        Case "SuiviSante"
            ' Ajout 07/10/2026 (refonte de Synthese_Care) : on affiche les 4
            ' colonnes de suivi santé. L'ancien écran affichait Date_consult,
            ' Spe_consult et un statut ; on y ajoute SoldeSante, déjà calculé par
            ' mod_SuiviSante.CalculerSuiviSante.
            ws.Columns("M").Hidden = False   ' StatutSante
            ws.Columns("N").Hidden = False   ' SoldeSante
            ws.Columns("O").Hidden = False   ' Date_consult
            ws.Columns("P").Hidden = False   ' Spe_consult
        ' Case "" (recherche libre) : aucune des 5 colonnes n'a de sens générique,
        ' on les laisse toutes masquées (état neutre défini plus haut).
    End Select

End Sub


' =====================================================================================
' OUTILS DEVELOPPEUR
' =====================================================================================
Sub AfficherFeuilleRecherchePourEdition()
    mod_InstallCommun.AfficherPourEdition NOM_FEUILLE_RECHERCHE, "CreerFeuilleRechercheOperations", "MasquerFeuilleRechercheApresEdition"
End Sub

Sub MasquerFeuilleRechercheApresEdition()
    mod_InstallCommun.MasquerApresEdition NOM_FEUILLE_RECHERCHE
End Sub
