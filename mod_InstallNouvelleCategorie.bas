Option Explicit

' =====================================================================================
' MODULE : mod_InstallNouvelleCategorie
'
' PHASE 3 du chantier « Catégorie / Sous-catégorie / Ventilation » - PARTIE 1/2
'
' RÔLE (à lire en premier, même si vous débutez) :
'   Ce module CONSTRUIT la feuille-formulaire qui permet de créer une nouvelle
'   catégorie (ou une nouvelle sous-catégorie d'une catégorie existante).
'   Elle s'ouvre par-dessus un autre formulaire (contrôle des catégories,
'   ventilation ou saisie d'une échéance), quand l'opérateur clique sur le bouton "+".
'   Ce module ne contient AUCUNE logique : elle se trouve dans mod_NouvelleCategorie
'   (partie 2/2), exactement comme pour la Phase 2.
'
'   Deux "listes modifiables" (menus déroulants qui acceptent aussi une saisie libre,
'   avec un simple avertissement) :
'     - Categorie      : liste des catégories existantes, ou saisie d'une nouvelle.
'     - Sous-categorie : liste des sous-catégories DÉJÀ CONNUES de la catégorie
'                        choisie, ou saisie d'une nouvelle.
'
' REFONTE DU 09/10/2026 (harmonisation des formulaires) :
'   Construite avec les outils COMMUNS de mod_InstallCommun. Le titre et les deux
'   lignes d'aide en italique ont été remplacés par le bloc Instructions commun
'   (sous la ligne de boutons). Largeurs de colonnes et de boutons : export
'   frm_NouvelleCategorie_structure.txt.
'
' INSTALLATION (une seule fois, ou après toute modification de ce fichier) :
'   1. Alt+F11, Fichier > Importer un fichier... : importer mod_InstallCommun.bas,
'      puis CE fichier, puis mod_NouvelleCategorie.bas.
'   2. Ctrl+G, taper : CreerFeuilleNouvelleCategorie, puis Entrée.
'   3. Si la feuille vient d'être créée : coller dans son module de code les deux
'      procédures du fichier "CodeBehind_frm_NouvelleCategorie.txt".
'
' Ce module ne modifie AUCUNE donnée : il ajoute seulement une feuille masquée.
' =====================================================================================

Public Const NC_NOM_FEUILLE As String = "frm_NouvelleCategorie"

' En-tête commun (voir mod_InstallCommun) : 1 ligne de boutons, pas de compteur.
'   ligne 1 marge / 2 boutons / 3 fine ligne / 4 Instructions / 5 fine ligne /
'   6 et suivantes : corps du formulaire
Public Const NC_LIGNE_BOUTONS As Long = 2
Private Const NC_LIGNE_INSTRUCTIONS As Long = 4
Private Const NC_LIGNE_CORPS As Long = 6

Public Const NC_ADR_CAT As String = "C6"
Public Const NC_ADR_SOUS As String = "C7"
Public Const NC_ADR_MESSAGE As String = "B9"      ' zone fusionnée B9:E9

' Zone technique masquée : sous-catégories de la catégorie choisie (colonne Z).
Public Const NC_COL_AIDE As Long = 26
Public Const NC_LIGNE_AIDE_MAX As Long = 300


' =====================================================================================
' MACRO PRINCIPALE D'INSTALLATION
' =====================================================================================
Public Sub CreerFeuilleNouvelleCategorie()

    Dim ws As Worksheet
    Dim wsPrecedente As Worksheet

    Set wsPrecedente = ActiveSheet

    Set ws = mod_InstallCommun.PreparerFeuille(NC_NOM_FEUILLE)
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    mod_InstallCommun.MettreEnForme ws, Array(2, 26, 44, 22, 22, 2)

    ConstruireEntete ws
    ConstruireChamps ws
    ConstruireBoutons ws        ' en dernier : les boutons se placent d'après les hauteurs de lignes

    ws.Columns(NC_COL_AIDE).Hidden = True

    mod_InstallCommun.TerminerFeuille ws, wsPrecedente
    Application.ScreenUpdating = True

    MsgBox mod_Display.FR("La feuille '") & NC_NOM_FEUILLE & mod_Display.FR("' a {e2}t{e2} cr{e2}{e2}e puis masqu{e2}e.") & vbCrLf & vbCrLf & _
           mod_Display.FR("PROCHAINE {E2}TAPE (indispensable, si la feuille vient d'{ea}tre cr{e2}{e2}e) : coller le code des {e2}v{e2}nements dans la feuille.") & vbCrLf & _
           "Nom interne (CodeName) de cette feuille : " & ws.CodeName & vbCrLf & vbCrLf & _
           mod_Display.FR("Voir le fichier CodeBehind_frm_NouvelleCategorie.txt."), _
           vbInformation, mod_Display.FR("Installation termin{e2}e")

End Sub


' =====================================================================================
' En-tête commun : hauteurs et bloc Instructions
' =====================================================================================
Private Sub ConstruireEntete(ByVal ws As Worksheet)

    If mod_InstallCommun.PremiereLigneCorps(1, False) <> NC_LIGNE_CORPS Or _
       mod_InstallCommun.LigneInstructions(1, False) <> NC_LIGNE_INSTRUCTIONS Then
        MsgBox "mod_InstallNouvelleCategorie : constantes de lignes incoherentes avec mod_InstallCommun.", vbCritical
    End If

    mod_InstallCommun.PoserHauteursEntete ws, 1, False
    mod_InstallCommun.EcrireInstructions ws, NC_LIGNE_INSTRUCTIONS, 2, 2, 3, 5, TexteInstructions()

End Sub

Private Function TexteInstructions() As String

    Dim t As String

    t = "Ce formulaire permet de cr{e2}er une nouvelle cat{e2}gorie, ou une nouvelle sous-cat{e2}gorie d'une cat{e2}gorie existante."
    t = t & Chr(10) & "Il s'ouvre avec le bouton [b:+] de [f:frm_ControleCategories], de [f:frm_Ventilation] ou de [f:frm_Echeance]."
    t = t & Chr(10) & "Choisissez une [c:Cat{e2}gorie] existante dans la liste, ou tapez-en une nouvelle. Choisissez ensuite une "
    t = t & "[c:Sous-cat{e2}gorie] d{e2}j{a2} connue de cette cat{e2}gorie, ou tapez-en une nouvelle ; ce champ peut aussi rester vide."
    t = t & Chr(10) & "Boutons :"
    t = t & Chr(10) & "- [b:Valider] : enregistre la cat{e2}gorie et la sous-cat{e2}gorie saisies, puis referme le formulaire."
    t = t & Chr(10) & "- [b:Annuler] : referme le formulaire sans rien cr{e2}er."

    TexteInstructions = t

End Function


' =====================================================================================
' Champs de saisie et message d'aide
' =====================================================================================
' Les listes déroulantes (validation de données) sont posées par le programme à
' chaque ouverture (voir mod_NouvelleCategorie), pas ici.
Private Sub ConstruireChamps(ByVal ws As Worksheet)

    mod_InstallCommun.PoserEtiquette ws.Range("B6"), mod_Display.FR("Cat{e2}gorie")
    mod_InstallCommun.PoserChampSaisie ws.Range(NC_ADR_CAT), "@"

    mod_InstallCommun.PoserEtiquette ws.Range("B7"), mod_Display.FR("Sous-cat{e2}gorie")
    mod_InstallCommun.PoserChampSaisie ws.Range(NC_ADR_SOUS), "@"

    ws.Rows(8).RowHeight = mod_InstallCommun.FRM_H_SEP

    ' Message (erreurs de validation) : sa couleur est choisie par le programme.
    mod_InstallCommun.PoserMessage ws.Range("B9:E9")
    ws.Range(NC_ADR_MESSAGE).Font.Color = RGB(192, 80, 0)

End Sub


' =====================================================================================
' Boutons "Valider" et "Annuler"
' =====================================================================================
Private Sub ConstruireBoutons(ByVal ws As Worksheet)

    Dim gauche As Double

    gauche = ws.Cells(1, 2).Left

    mod_InstallCommun.AjouterBoutonEntete ws, NC_LIGNE_BOUTONS, gauche, mod_InstallCommun.CapValider(), _
        "NcValider", "btnNcValider", mod_InstallCommun.FRM_BTN_L
    mod_InstallCommun.AjouterBoutonEntete ws, NC_LIGNE_BOUTONS, gauche, mod_InstallCommun.CapAnnuler(), _
        "NcAnnuler", "btnNcAnnuler", mod_InstallCommun.FRM_BTN_L

End Sub


' =====================================================================================
' OUTILS DÉVELOPPEUR (Ctrl+G)
' =====================================================================================
Public Sub AfficherFeuilleNouvelleCategoriePourEdition()
    mod_InstallCommun.AfficherPourEdition NC_NOM_FEUILLE, "CreerFeuilleNouvelleCategorie", "MasquerFeuilleNouvelleCategorieApresEdition"
End Sub

Public Sub MasquerFeuilleNouvelleCategorieApresEdition()
    mod_InstallCommun.MasquerApresEdition NC_NOM_FEUILLE
End Sub
