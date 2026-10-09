Option Explicit

' =====================================================================================
' MODULE : mod_InstallVentilation
'
' PHASE 4 du chantier « Catégorie / Sous-catégorie / Ventilation » - PARTIE 1/2
' (Version 2 : nouvelle ergonomie, à la demande de l'opérateur après les tests)
'
' RÔLE (à lire en premier, même si vous débutez) :
'   Ce module met en place les DEUX éléments nécessaires à la ventilation d'une
'   opération entre plusieurs sous-catégories :
'
'     1. La feuille de DONNEES "Ventilations", qui contient le tableau
'        "TblVentilations" : une ligne par sous-catégorie ventilée, reliée à
'        l'opération d'origine par son ID_Transaction. Cette feuille reste
'        simplement MASQUÉE (pas très masquée) : comme Param ou Import_data, vous
'        pouvez l'afficher si besoin pour vérifier son contenu.
'
'     2. La feuille-FORMULAIRE "frm_Ventilation".
'
'   ERGONOMIE (modifiée à la suite de vos remarques) : la feuille est divisée en
'   deux zones bien distinctes :
'     - EN HAUT (lignes 16 à 26) : un TABLEAU D'AFFICHAGE des lignes déjà ajoutées
'       à cette ventilation (catégorie, sous-catégorie, montant), avec un bouton
'       "Éditer" en face de chaque ligne.
'     - EN BAS (lignes 28 et suivantes) : un PETIT FORMULAIRE DE SAISIE, TOUJOURS
'       IDENTIQUE, pour AJOUTER une ligne ou MODIFIER une ligne existante (via son
'       bouton "Éditer"). Il reprend le principe du formulaire de contrôle des
'       catégories (phase 2) : un champ Categorie avec un bouton "+" pour en créer
'       une nouvelle, et un champ Sous-categorie dont la liste dépend de la catégorie choisie.
'
'   Ce module ne contient AUCUNE logique de calcul ou de validation : celle-ci se
'   trouve dans mod_Ventilation (partie 2/2).
'
' REFONTE DU 09/10/2026 (harmonisation des formulaires) :
'   La feuille frm_Ventilation est construite avec les outils COMMUNS de
'   mod_InstallCommun (couleurs, boutons, étiquettes, champs, bloc Instructions).
'   Le titre a été remplacé par le bloc Instructions commun, sous la ligne de boutons.
'   Les lignes du tableau et de la zone de saisie gardent les mêmes numéros qu'avant
'   (constantes ci-dessous inchangées). Les boutons "Éditer" (crayon) gardent la taille
'   relevée sur la feuille réelle (export frm_Ventilation_structure.txt).
'
' INSTALLATION :
'   1. Alt+F11, Fichier > Importer un fichier... : importer mod_InstallCommun.bas, puis CE fichier.
'   2. Importer aussi mod_Ventilation.bas.
'   3. Ctrl+G (fenêtre Exécution), taper : PreparerPhase4Ventilation, puis Entrée.
'      (crée TblVentilations, la met à jour si elle existe déjà et reconstruit
'      frm_Ventilation)
'   4. Le message de fin vous donne le CodeName de frm_Ventilation. Collez-y le
'      contenu de CodeBehind_frm_Ventilation.txt (INCHANGÉ depuis la version
'      précédente : si vous l'avez déjà collé, inutile de recommencer).
'   5. IMPORTANT, problème déjà rencontré : après toute mise à jour de
'      mod_InstallControleCategories, il faut RELANCER CreerFeuilleControleCategories
'      (répondre Oui à la reconstruction) pour que les boutons apparaissent réellement
'      sur la feuille. Importer le fichier .bas ne suffit pas : il faut exécuter la
'      macro d'installation pour reconstruire la feuille.
' =====================================================================================

Public Const VEN_NOM_FEUILLE_DONNEES As String = "Ventilations"
Public Const VEN_NOM_TABLE As String = "TblVentilations"

Public Const VEN_NOM_FEUILLE As String = "frm_Ventilation"

' En-tête commun (voir mod_InstallCommun) : 1 ligne de boutons, pas de compteur.
'   ligne 1 marge / 2 boutons / 3 fine ligne / 4 Instructions / 5 fine ligne /
'   6 et suivantes : corps du formulaire
Public Const VEN_LIGNE_BOUTONS As Long = 2
Private Const VEN_LIGNE_INSTRUCTIONS As Long = 4
Private Const VEN_LIGNE_CORPS As Long = 6

Public Const VEN_ADR_DATE As String = "C6"
Public Const VEN_ADR_TIERS As String = "C7"
Public Const VEN_ADR_LIBELLE As String = "C8"
Public Const VEN_ADR_MONTANT As String = "C9"
Public Const VEN_ADR_CATACTUELLE As String = "C10"

Public Const VEN_ADR_MONTANT_A_VENTILER As String = "C12"
Public Const VEN_ADR_TOTAL_SAISI As String = "C13"
Public Const VEN_ADR_RESTE As String = "C14"

' --- Tableau D'AFFICHAGE des lignes déjà ajoutées (lecture seule pour l'opérateur :
' une ligne se modifie uniquement via le formulaire de saisie du bas et son bouton
' "Éditer") ---
Public Const VEN_LIGNE_GRILLE_ENTETE As Long = 16
Public Const VEN_LIGNE_GRILLE_DEBUT As Long = 17
Public Const VEN_NB_LIGNES As Long = 10
Public Const VEN_LIGNE_GRILLE_FIN As Long = 26     ' = VEN_LIGNE_GRILLE_DEBUT + VEN_NB_LIGNES - 1

Public Const VEN_COL_CAT As Long = 2      ' colonne B
Public Const VEN_COL_SOUS As Long = 3     ' colonne C
Public Const VEN_COL_MONTANT As Long = 4  ' colonne D
Public Const VEN_COL_NOTES As Long = 5    ' colonne E (ajout 01/10/2026, point 3)
Public Const VEN_COL_EDITER As Long = 6   ' colonne F : bouton "Éditer" de chaque ligne (déplacé de E à F)

' --- Formulaire de SAISIE (une seule ligne à la fois : ajout ou édition) ---
Public Const VEN_LIGNE_SAISIE_TITRE As Long = 28
Public Const VEN_ADR_SAISIE_CAT As String = "C29"
Public Const VEN_ADR_SAISIE_SOUS As String = "C30"
Public Const VEN_ADR_SAISIE_MONTANT As String = "C31"
' Ajout du 01/10/2026 (point 3) : commentaire libre, disponible pour CHAQUE ligne.
' IMPORTANT : cellule volontairement SIMPLE, NON FUSIONNÉE (même principe que
' Categorie/Sous-categorie/Montant ci-dessus). Une première version fusionnait C32:F32,
' ce qui provoquait une erreur 1004 ("Cette action ne peut pas être appliquée à une
' cellule fusionnée") lorsque le code écrivait dans VEN_ADR_SAISIE_NOTES, qui ne pointe
' pas nécessairement vers le coin supérieur gauche de la fusion. Une cellule simple
' élimine ce risque; sa largeur réduite est compensée par le retour automatique à la
' ligne (voir ConstruireFormulaireSaisie).
Public Const VEN_ADR_SAISIE_NOTES As String = "C32"
Public Const VEN_ADR_SAISIE_MESSAGE As String = "B33"     ' déplacée de B32 à B33
Public Const VEN_LIGNE_BOUTON_AJOUTER As Long = 35        ' déplacée de 34 à 35

' Zone technique masquée : sous-catégories de la catégorie choisie DANS LE FORMULAIRE
' DE SAISIE (une seule liste à gérer, comme dans les formulaires précédents, au lieu
' d'une liste par ligne de la grille).
Public Const VEN_COL_AIDE As Long = 26         ' colonne Z
Public Const VEN_LIGNE_AIDE_MAX As Long = 300

' Ajout du 01/10/2026 (point 4 : annuler une ventilation) : deux colonnes techniques
' ajoutées à TblOperations pour mémoriser la catégorie/sous-catégorie d'AVANT la
' première ventilation d'une opération, afin de pouvoir les restaurer si elle est
' supprimée plus tard (voir mod_ControleCategories.ControleVentiler et
' mod_RechercheOperations.RevoirVentilationRO).
Public Const NOM_COL_CAT_AVANT_VENTILATION As String = "CategorieAvantVentilation"
Public Const NOM_COL_SOUS_AVANT_VENTILATION As String = "SousCategorieAvantVentilation"


' =====================================================================================
' MACRO D'ENSEMBLE : crée ou met à jour la table de stockage, puis le formulaire.
' =====================================================================================
Public Sub PreparerPhase4Ventilation()
    PreparerTableVentilations
    AjouterColonnesAnnulationVentilation
    CreerFeuilleVentilation
End Sub


' =====================================================================================
' ÉTAPE 1 : feuille de données "Ventilations" et tableau "TblVentilations"
' =====================================================================================
' Sans danger à relancer : les colonnes déjà présentes ne sont jamais modifiées;
' seules les colonnes manquantes sont ajoutées (utile si une version antérieure
' de ce module a été installée).
Public Sub PreparerTableVentilations()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim colonnesSante As Variant
    Dim i As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(VEN_NOM_FEUILLE_DONNEES)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = VEN_NOM_FEUILLE_DONNEES
    End If

    On Error Resume Next
    Set tbl = ws.ListObjects(VEN_NOM_TABLE)
    On Error GoTo 0

    If tbl Is Nothing Then
        ws.Range("A1:E1").NumberFormat = "@"
        ws.Range("A1:E1").value = Array("ID_Transaction", "Categorie", "SousCategorie", "Montant", "DateVentilation")
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:E1"), , xlYes)
        tbl.Name = VEN_NOM_TABLE
        ' Pas de mise en forme ici : juste après sa création, le tableau n'a AUCUNE
        ' ligne de données (DataBodyRange est vide); lui appliquer un format provoquerait
        ' une erreur. Chaque ligne reçoit son format au moment où elle est écrite
        ' (voir mod_Ventilation.AjouterLigneVentilation).
    End If

    ' --- Colonnes de suivi santé (mêmes noms que dans TblOperations) -------------------
    colonnesSante = Array("Notes", "Date_consult", "Spe_Consult", "StatutSante", "SoldeSante", _
                          "DepassementHoraires", "CommentaireSante", "Franchise", "Beneficiaire")

    For i = LBound(colonnesSante) To UBound(colonnesSante)
        If Not ColonneExisteDansTable(tbl, CStr(colonnesSante(i))) Then
            tbl.ListColumns.Add.Name = CStr(colonnesSante(i))
        End If
    Next i

    ws.Visible = xlSheetHidden

End Sub

Private Function ColonneExisteDansTable(ByVal tbl As ListObject, ByVal nom As String) As Boolean
    Dim lc As ListColumn
    On Error Resume Next
    Set lc = tbl.ListColumns(nom)
    On Error GoTo 0
    ColonneExisteDansTable = Not (lc Is Nothing)
End Function

' Ajout du 01/10/2026 (point 4 : annuler une ventilation) -------------------------------
' Ajoute à TblOperations (PAS à TblVentilations) les deux colonnes techniques qui
' mémorisent la catégorie/sous-catégorie d'avant la première ventilation d'une opération.
' Sans danger à relancer : ne fait rien si les colonnes existent déjà. Elles sont toujours
' ajoutées À LA FIN du tableau (jamais au milieu), afin de ne déplacer aucune colonne
' existante; même principe que mod_Categories.AjouterColonneSousCategorie.
Public Sub AjouterColonnesAnnulationVentilation()

    Dim tblOps As ListObject

    Set tblOps = mod_DonneesTable.GetOperationsTable()
    If tblOps Is Nothing Then Exit Sub

    If Not ColonneExisteDansTable(tblOps, NOM_COL_CAT_AVANT_VENTILATION) Then
        tblOps.ListColumns.Add.Name = NOM_COL_CAT_AVANT_VENTILATION
    End If
    If Not ColonneExisteDansTable(tblOps, NOM_COL_SOUS_AVANT_VENTILATION) Then
        tblOps.ListColumns.Add.Name = NOM_COL_SOUS_AVANT_VENTILATION
    End If

End Sub

' =====================================================================================
' ÉTAPE 2 : feuille-formulaire "frm_Ventilation"
' =====================================================================================
Public Sub CreerFeuilleVentilation()

    Dim ws As Worksheet
    Dim wsPrecedente As Worksheet

    Set wsPrecedente = ActiveSheet

    Set ws = mod_InstallCommun.PreparerFeuille(VEN_NOM_FEUILLE)
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    ' Colonnes : marge / étiquettes / sous-catégorie / montant / notes / bouton "Éditer" / marge
    mod_InstallCommun.MettreEnForme ws, Array(2, 26, 30, 16, 30, 2.4, 2)

    ConstruireEntete ws
    ConstruireInformations ws
    ConstruireTotaux ws
    ConstruireGrilleAffichage ws
    ConstruireFormulaireSaisie ws
    ConstruireBoutons ws        ' en dernier : les boutons se placent d'après les hauteurs de lignes

    ws.Columns(VEN_COL_AIDE).Hidden = True

    mod_InstallCommun.TerminerFeuille ws, wsPrecedente
    Application.ScreenUpdating = True

    MsgBox mod_Display.FR("La feuille '") & VEN_NOM_FEUILLE & mod_Display.FR("' a {e2}t{e2} cr{e2}{e2}e puis masqu{e2}e.") & vbCrLf & vbCrLf & _
           mod_Display.FR("Si ce n'est pas deja fait : collez le code des {e2}v{e2}nements dans la feuille") & vbCrLf & _
           "(CodeName : " & ws.CodeName & mod_Display.FR(") -- voir CodeBehind_frm_Ventilation.txt."), _
           vbInformation, mod_Display.FR("Installation termin{e2}e")

End Sub


' =====================================================================================
' En-tête commun : hauteurs et bloc Instructions
' =====================================================================================
Private Sub ConstruireEntete(ByVal ws As Worksheet)

    If mod_InstallCommun.PremiereLigneCorps(1, False) <> VEN_LIGNE_CORPS Or _
       mod_InstallCommun.LigneInstructions(1, False) <> VEN_LIGNE_INSTRUCTIONS Then
        MsgBox "mod_InstallVentilation : constantes de lignes incoherentes avec mod_InstallCommun.", vbCritical
    End If

    mod_InstallCommun.PoserHauteursEntete ws, 1, False
    mod_InstallCommun.EcrireInstructions ws, VEN_LIGNE_INSTRUCTIONS, 2, 2, 3, 6, TexteInstructions()

End Sub

Private Function TexteInstructions() As String

    Dim t As String

    t = "Ce formulaire permet de ventiler une op{e2}ration bancaire entre plusieurs cat{e2}gories et sous-cat{e2}gories."
    t = t & Chr(10) & "Le haut du formulaire rappelle l'op{e2}ration ([c:Date], [c:Tiers], [c:Libell{e2} / Notes], [c:Montant], "
    t = t & "[c:Cat{e2}gorie actuelle]) et le calcul : [c:Montant {a2} ventiler], [c:Total saisi] et [c:Reste {a2} ventiler]. "
    t = t & "Le tableau liste les lignes d{e2}j{a2} ajout{e2}es. Pour en ajouter une, renseignez [c:Cat{e2}gorie], [c:Sous-cat{e2}gorie], "
    t = t & "[c:Montant] (et [c:Notes] si besoin) dans la zone de saisie, puis cliquez sur [b:Ajouter la ligne]. "
    t = t & "La ventilation ne peut {ea}tre enregistr{e2}e que lorsque le [c:Reste {a2} ventiler] est exactement nul."
    t = t & Chr(10) & "Boutons :"
    t = t & Chr(10) & "- [b:Enregistrer] : enregistre toutes les lignes et referme le formulaire."
    t = t & Chr(10) & "- [b:Annuler] : abandonne la saisie en cours sans rien enregistrer."
    t = t & Chr(10) & "- [b:Supprimer la ventilation] : visible seulement si l'op{e2}ration est d{e2}j{a2} ventil{e2}e ; efface toutes ses lignes "
    t = t & "et rend {a2} l'op{e2}ration sa cat{e2}gorie d'origine."
    t = t & Chr(10) & "- [b:" & mod_InstallCommun.CapEditer() & "] (en face de chaque ligne du tableau) : retire la ligne du tableau et la recopie dans la zone de saisie pour la corriger."
    t = t & Chr(10) & "- [b:+] : cr{e2}e une nouvelle cat{e2}gorie ou sous-cat{e2}gorie (ouvre [f:frm_NouvelleCategorie])."
    t = t & Chr(10) & "- [b:Ajouter la ligne] : ajoute la ligne saisie au tableau (ou remplace la ligne en cours de correction)."
    t = t & Chr(10) & "- [b:Effacer la saisie] : vide la zone de saisie."
    t = t & Chr(10) & "ATTENTION : une ligne retir{e2}e du tableau avec [b:" & mod_InstallCommun.CapEditer() & "] est perdue si vous ne cliquez pas ensuite sur [b:Ajouter la ligne]."

    TexteInstructions = t

End Function


' =====================================================================================
' Informations de l'opération à ventiler (lecture seule)
' =====================================================================================
' Les valeurs sont écrites dans la cellule de gauche de chaque zone (C6, C7...), la
' zone C:E étant fusionnée pour laisser la place aux textes longs.
Private Sub ConstruireInformations(ByVal ws As Worksheet)

    mod_InstallCommun.PoserEtiquette ws.Range("B6"), "Date"
    mod_InstallCommun.PoserEtiquette ws.Range("B7"), "Tiers"
    mod_InstallCommun.PoserEtiquette ws.Range("B8"), mod_Display.FR("Libell{e2} / Notes")
    mod_InstallCommun.PoserEtiquette ws.Range("B9"), "Montant"
    mod_InstallCommun.PoserEtiquette ws.Range("B10"), mod_Display.FR("Cat{e2}gorie actuelle")

    mod_InstallCommun.PoserChampLecture ws.Range("C6:E6"), "@"
    mod_InstallCommun.PoserChampLecture ws.Range("C7:E7"), "@"
    mod_InstallCommun.PoserChampLecture ws.Range("C8:E8"), "@", True
    mod_InstallCommun.PoserChampLecture ws.Range("C9:E9"), "@"
    mod_InstallCommun.PoserChampLecture ws.Range("C10:E10"), "@"
    ws.Range(VEN_ADR_MONTANT).Font.Bold = True

    ws.Rows(11).RowHeight = mod_InstallCommun.FRM_H_SEP

End Sub


' =====================================================================================
' Bloc des totaux (lecture seule, recalculés par le programme)
' =====================================================================================
Private Sub ConstruireTotaux(ByVal ws As Worksheet)

    Dim formatEuro As String

    ' Format numérique avec le symbole euro en suffixe (voir mod_Ventilation pour
    ' le détail de la construction de cette chaîne).
    formatEuro = "#,##0.00" & Chr(34) & " " & ChrW(8364) & Chr(34)

    mod_InstallCommun.PoserEtiquette ws.Range("B12"), mod_Display.FR("Montant {a2} ventiler")
    mod_InstallCommun.PoserEtiquette ws.Range("B13"), "Total saisi"
    mod_InstallCommun.PoserEtiquette ws.Range("B14"), mod_Display.FR("Reste {a2} ventiler")

    mod_InstallCommun.PoserChampLecture ws.Range(VEN_ADR_MONTANT_A_VENTILER), formatEuro
    mod_InstallCommun.PoserChampLecture ws.Range(VEN_ADR_TOTAL_SAISI), formatEuro
    mod_InstallCommun.PoserChampLecture ws.Range(VEN_ADR_RESTE), formatEuro
    ws.Range("C12:C14").Font.Bold = True

    ws.Rows(15).RowHeight = mod_InstallCommun.FRM_H_SEP

End Sub


' =====================================================================================
' Tableau D'AFFICHAGE des lignes déjà ajoutées (lignes 16 à 26).
' =====================================================================================
' Ces cellules ne sont PAS destinées à être modifiées directement par l'opérateur : elles
' sont remplies par le programme (mod_Ventilation) au fil des ajouts et se modifient
' uniquement via le formulaire de saisie du bas (bouton "Éditer"). Fond gris (non
' modifiable), comme tout champ en lecture seule.
Private Sub ConstruireGrilleAffichage(ByVal ws As Worksheet)

    Dim ligne As Long

    ws.Cells(VEN_LIGNE_GRILLE_ENTETE, VEN_COL_CAT).Value = mod_Display.FR("Cat{e2}gorie")
    ws.Cells(VEN_LIGNE_GRILLE_ENTETE, VEN_COL_SOUS).Value = mod_Display.FR("Sous-cat{e2}gorie")
    ws.Cells(VEN_LIGNE_GRILLE_ENTETE, VEN_COL_MONTANT).Value = "Montant"
    ws.Cells(VEN_LIGNE_GRILLE_ENTETE, VEN_COL_NOTES).Value = "Notes"
    mod_InstallCommun.PoserEnteteTableau ws.Range(ws.Cells(VEN_LIGNE_GRILLE_ENTETE, VEN_COL_CAT), _
                                                  ws.Cells(VEN_LIGNE_GRILLE_ENTETE, VEN_COL_EDITER))

    For ligne = VEN_LIGNE_GRILLE_DEBUT To VEN_LIGNE_GRILLE_FIN
        mod_InstallCommun.PoserCorpsTableauLecture ws.Range(ws.Cells(ligne, VEN_COL_CAT), ws.Cells(ligne, VEN_COL_NOTES))
        ws.Cells(ligne, VEN_COL_CAT).NumberFormat = "@"
        ws.Cells(ligne, VEN_COL_SOUS).NumberFormat = "@"
        ws.Cells(ligne, VEN_COL_MONTANT).NumberFormat = "#,##0.00"
        ws.Cells(ligne, VEN_COL_NOTES).NumberFormat = "@"
    Next ligne

    ws.Rows(VEN_LIGNE_GRILLE_FIN + 1).RowHeight = mod_InstallCommun.FRM_H_SEP

End Sub


' =====================================================================================
' Formulaire de SAISIE (ajout ou édition d'UNE ligne à la fois)
' =====================================================================================
Private Sub ConstruireFormulaireSaisie(ByVal ws As Worksheet)

    mod_InstallCommun.PoserTitreBloc ws.Range("B" & VEN_LIGNE_SAISIE_TITRE & ":F" & VEN_LIGNE_SAISIE_TITRE), _
                                     mod_Display.FR("Ajouter ou modifier une ligne")

    ' --- Catégorie (le bouton "+" est posé avec les autres boutons) ---
    mod_InstallCommun.PoserEtiquette ws.Range("B29"), mod_Display.FR("Cat{e2}gorie")
    mod_InstallCommun.PoserChampSaisie ws.Range(VEN_ADR_SAISIE_CAT), "@"

    ' --- Sous-catégorie (liste dépendante de la catégorie ci-dessus) ---
    mod_InstallCommun.PoserEtiquette ws.Range("B30"), mod_Display.FR("Sous-cat{e2}gorie")
    mod_InstallCommun.PoserChampSaisie ws.Range(VEN_ADR_SAISIE_SOUS), "@"

    ' --- Montant ---
    mod_InstallCommun.PoserEtiquette ws.Range("B31"), "Montant"
    mod_InstallCommun.PoserChampSaisie ws.Range(VEN_ADR_SAISIE_MONTANT), "#,##0.00"

    ' --- Notes (commentaire libre, ajout du 01/10/2026, point 3) ------------------------
    ' Cellule volontairement SIMPLE, NON FUSIONNÉE (comme Categorie/Sous-categorie/
    ' Montant ci-dessus). Voir le commentaire de VEN_ADR_SAISIE_NOTES, qui explique
    ' l'erreur 1004 provoquée par la première version fusionnée. La largeur réduite
    ' de cette cellule est compensée par le retour automatique à la ligne et par une
    ' hauteur de ligne de 27.
    mod_InstallCommun.PoserEtiquette ws.Range("B32"), "Notes"
    mod_InstallCommun.PoserChampSaisie ws.Range(VEN_ADR_SAISIE_NOTES), "@", True

    ' --- Message (erreurs de saisie de cette ligne) ---
    mod_InstallCommun.PoserMessage ws.Range(VEN_ADR_SAISIE_MESSAGE & ":F33")
    ws.Range(VEN_ADR_SAISIE_MESSAGE).Font.Color = RGB(192, 80, 0)

    ws.Rows(34).RowHeight = mod_InstallCommun.FRM_H_SEP
    ws.Rows(VEN_LIGNE_BOUTON_AJOUTER).RowHeight = mod_InstallCommun.FRM_H_BOUTONS

End Sub


' =====================================================================================
' Boutons
' =====================================================================================
Private Sub ConstruireBoutons(ByVal ws As Worksheet)

    Dim gauche As Double
    Dim ligne As Long

    ' --- Boutons globaux du haut : "Enregistrer" (finalise TOUTE la ventilation),
    '     "Annuler" et "Supprimer la ventilation" ---
    gauche = ws.Cells(1, 2).Left
    mod_InstallCommun.AjouterBoutonEntete ws, VEN_LIGNE_BOUTONS, gauche, mod_InstallCommun.CapEnregistrer(), _
        "VenTerminer", "btnVenTerminer", mod_InstallCommun.FRM_BTN_L
    mod_InstallCommun.AjouterBoutonEntete ws, VEN_LIGNE_BOUTONS, gauche, mod_InstallCommun.CapAnnuler(), _
        "VenAnnuler", "btnVenAnnuler", mod_InstallCommun.FRM_BTN_L

    ' Ajout du 01/10/2026 (point 4) : bouton de suppression d'une ventilation EXISTANTE.
    ' Masqué par défaut à la construction : mod_Ventilation.OuvrirVentilation le rend
    ' visible uniquement si la ventilation existait déjà AVANT l'ouverture du formulaire
    ' (rien à supprimer pour une nouvelle ventilation).
    mod_InstallCommun.AjouterBoutonEntete ws, VEN_LIGNE_BOUTONS, gauche, mod_Display.FR("Supprimer la ventilation"), _
        "VenSupprimerVentilation", "btnVenSupprimerVentilation", 140
    ws.Shapes("btnVenSupprimerVentilation").Visible = False

    ' --- Bouton "Éditer" de chaque ligne du tableau ---
    ' Toutes les lignes ont leur bouton dès la construction, qu'elles soient remplies ou
    ' non : cliquer sur une ligne vide affiche simplement un message (voir
    ' mod_Ventilation.VenEditerLigne). Le nom "btnVenEditerLigneN" est relu par le code.
    For ligne = VEN_LIGNE_GRILLE_DEBUT To VEN_LIGNE_GRILLE_FIN
        mod_InstallCommun.AjouterBoutonCellule ws, ws.Cells(ligne, VEN_COL_EDITER), mod_InstallCommun.CapEditer(), _
            "VenEditerLigne", "btnVenEditerLigne" & (ligne - VEN_LIGNE_GRILLE_DEBUT + 1), 13.4, 1
    Next ligne

    ' --- Bouton "+" à côté du champ Catégorie de la zone de saisie ---
    mod_InstallCommun.AjouterBoutonCellule ws, ws.Range("D29"), mod_InstallCommun.CapAjouter(), _
        "VenNouvelleCategorie", "btnVenNouvelleCategorie", mod_InstallCommun.FRM_BTN_PLUS

    ' --- Boutons "Ajouter la ligne" et "Effacer la saisie" (ligne du bas) ---
    gauche = ws.Cells(1, 2).Left
    mod_InstallCommun.AjouterBoutonEntete ws, VEN_LIGNE_BOUTON_AJOUTER, gauche, mod_Display.FR("Ajouter la ligne"), _
        "VenAjouterLigne", "btnVenAjouterLigne", 100
    mod_InstallCommun.AjouterBoutonEntete ws, VEN_LIGNE_BOUTON_AJOUTER, gauche, mod_Display.FR("Effacer la saisie"), _
        "VenEffacerSaisie", "btnVenEffacerSaisie", 100

End Sub


' =====================================================================================
' OUTILS DÉVELOPPEUR (Ctrl+G)
' =====================================================================================
Public Sub AfficherFeuilleVentilationPourEdition()
    mod_InstallCommun.AfficherPourEdition VEN_NOM_FEUILLE, "PreparerPhase4Ventilation", "MasquerFeuilleVentilationApresEdition"
    MsgBox mod_Display.FR("Rappel : les listes d{e2}roulantes ne sont pos{e2}es qu'{a2} l'ouverture normale du formulaire") & _
           mod_Display.FR(" (bouton Ventiler, ou TesterVentilation) -- elles n'apparaissent pas si vous affichez juste la feuille ainsi."), _
           vbInformation
End Sub

Public Sub MasquerFeuilleVentilationApresEdition()
    mod_InstallCommun.MasquerApresEdition VEN_NOM_FEUILLE
End Sub
