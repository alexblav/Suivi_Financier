Option Explicit

' =====================================================================================
' MODULE : mod_InstallControleCategories
'
' PHASE 2 du chantier « Catégorie / Sous-catégorie / Ventilation » - PARTIE 1/2
'
' RÔLE (à lire en premier, même si vous débutez) :
'   Ce module CONSTRUIT la feuille qui sert de "formulaire" pour contrôler les
'   catégories des opérations importées : mise en page, libellés et boutons.
'   Il ne contient AUCUNE logique de fonctionnement : celle-ci se trouve dans
'   le module mod_ControleCategories (partie 2/2).
'
'   Pourquoi une FEUILLE et pas un UserForm ? Parce que les UserForms s'affichent mal
'   sur les postes multi-écrans (problème de DPI). On utilise donc une feuille masquée
'   (xlSheetVeryHidden) que le programme affiche au bon moment.
'
'   Analogie : ce module "construit les murs" de la pièce; le module
'   mod_ControleCategories "branche l'électricité".
'
' REFONTE DU 09/10/2026 (harmonisation des formulaires) :
'   La feuille est désormais construite avec les outils COMMUNS de mod_InstallCommun
'   (couleurs, boutons, étiquettes, champs, compteur, bloc Instructions), comme tous
'   les autres formulaires. Les boutons du haut gardent les largeurs relevées sur la
'   feuille réelle (export frm_ControleCategories_structure.txt). Le titre et le
'   bloc d'instructions du bas ont été remplacés par le bloc Instructions commun,
'   placé sous la ligne de boutons. Les adresses des cellules ont donc changé :
'   toutes sont regroupées dans les constantes ci-dessous.
'
' INSTALLATION (une seule fois, ou après toute modification de ce fichier) :
'   1. Alt+F11, puis Fichier > Importer un fichier... : importer mod_InstallCommun.bas
'      puis CE fichier (et mod_ControleCategories.bas si ce n'est pas déjà fait).
'   2. Ctrl+G (fenêtre Exécution), taper : CreerFeuilleControleCategories, puis Entrée
'      (répondre Oui à la reconstruction si la feuille existe déjà).
'   3. Le message final indique le NOM de la feuille créée. Collez dans son
'      module de code les deux procédures du fichier
'      "CodeBehind_frm_ControleCategories.txt" (inchangées; à refaire seulement si la
'      feuille a été supprimée puis recréée).
'
' Ce module ne modifie AUCUNE donnée : il ajoute seulement une feuille masquée.
' =====================================================================================

' --- Nom de la feuille-formulaire -------------------------------------------------------
Public Const CTRL_NOM_FEUILLE As String = "frm_ControleCategories"

' --- Position des éléments (regroupés ici pour n'avoir qu'un endroit à modifier) -------
' En-tête commun (voir mod_InstallCommun) : 1 ligne de boutons + compteur.
'   ligne 1 marge / 2 boutons / 3 fine ligne / 4 compteur / 5 fine ligne /
'   6 Instructions / 7 fine ligne / 8 et suivantes : corps du formulaire
Public Const CTRL_LIGNE_BOUTONS As Long = 2
Public Const CTRL_ADR_COMPTEUR As String = "B4"   ' "Opération: 3/27"
Private Const CTRL_LIGNE_INSTRUCTIONS As Long = 6
Private Const CTRL_LIGNE_CORPS As Long = 8

' Informations de l'opération : étiquette en colonne B, valeur en colonne C
Public Const CTRL_ADR_DATE As String = "C8"
Public Const CTRL_ADR_TIERS As String = "C9"      ' modifiable (sauf opération de santé)
Public Const CTRL_ADR_LIBELLE As String = "C10"   ' modifiable (sauf opération de santé)
Public Const CTRL_ADR_MONTANT As String = "C11"
Public Const CTRL_ADR_SOURCE As String = "C12"    ' catégorie envoyée par la banque

' Zones de saisie de l'opérateur
Public Const CTRL_ADR_CAT As String = "C14"       ' catégorie choisie
Public Const CTRL_ADR_SOUS As String = "C15"      ' sous-catégorie choisie

Public Const CTRL_ADR_MESSAGE As String = "B17"   ' message d'aide (zone fusionnée B17:E17)

' Zone technique masquée : liste des sous-catégories de la catégorie choisie.
' La colonne 26 correspond à la colonne Z; elle est masquée, loin à droite de la mise en page.
Public Const CTRL_COL_AIDE As Long = 26


' =====================================================================================
' MACRO PRINCIPALE D'INSTALLATION
' =====================================================================================
Public Sub CreerFeuilleControleCategories()

    Dim ws As Worksheet
    Dim wsPrecedente As Worksheet

    ' On retient la feuille affichée pour y revenir à la fin (l'opérateur ne doit pas
    ' se retrouver ailleurs après l'installation).
    Set wsPrecedente = ActiveSheet

    Set ws = mod_InstallCommun.PreparerFeuille(CTRL_NOM_FEUILLE)
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    ' Colonnes : marge / étiquettes / valeurs / zone des boutons "Ventiler" et "+" / marge
    mod_InstallCommun.MettreEnForme ws, Array(2, 26, 44, 22, 22, 2)

    ConstruireEntete ws
    ConstruireInformations ws
    ConstruireZoneSaisie ws
    ConstruireBoutons ws        ' en dernier : les boutons se placent d'après les hauteurs de lignes

    ' La colonne technique (liste des sous-catégories) doit rester invisible.
    ws.Columns(CTRL_COL_AIDE).Hidden = True

    ' Masquage complet : invisible pour l'opérateur, même via clic droit > Afficher.
    mod_InstallCommun.TerminerFeuille ws, wsPrecedente
    Application.ScreenUpdating = True

    MsgBox mod_Display.FR("La feuille '") & CTRL_NOM_FEUILLE & mod_Display.FR("' a {e2}t{e2} cr{e2}{e2}e puis masqu{e2}e.") & vbCrLf & vbCrLf & _
           mod_Display.FR("PROCHAINE {E2}TAPE (indispensable, si la feuille vient d'{ea}tre cr{e2}{e2}e) : coller le code des {e2}v{e2}nements dans la feuille.") & vbCrLf & _
           "Nom interne (CodeName) de cette feuille : " & ws.CodeName & vbCrLf & vbCrLf & _
           mod_Display.FR("Voir le fichier CodeBehind_frm_ControleCategories.txt."), _
           vbInformation, mod_Display.FR("Installation termin{e2}e")

End Sub


' =====================================================================================
' En-tête commun : hauteurs, compteur et bloc Instructions
' =====================================================================================
Private Sub ConstruireEntete(ByVal ws As Worksheet)

    ' Garde-fou : les constantes de ce module doivent correspondre à l'en-tête commun.
    If mod_InstallCommun.PremiereLigneCorps(1, True) <> CTRL_LIGNE_CORPS Or _
       mod_InstallCommun.LigneInstructions(1, True) <> CTRL_LIGNE_INSTRUCTIONS Then
        MsgBox "mod_InstallControleCategories : constantes de lignes incoherentes avec mod_InstallCommun.", vbCritical
    End If

    mod_InstallCommun.PoserHauteursEntete ws, 1, True
    mod_InstallCommun.PoserCompteur ws, mod_InstallCommun.LigneCompteur(1, True), 2, 5

    mod_InstallCommun.EcrireInstructions ws, CTRL_LIGNE_INSTRUCTIONS, 2, 2, 3, 5, TexteInstructions()

End Sub

Private Function TexteInstructions() As String

    Dim t As String

    t = "Ce formulaire permet de contr{o2}ler la cat{e2}gorie et la sous-cat{e2}gorie des op{e2}rations qui viennent d'{ea}tre import{e2}es."
    t = t & Chr(10) & "Les op{e2}rations sont pr{e2}sent{e2}es une par une : v{e2}rifiez les informations affich{e2}es, corrigez si besoin "
    t = t & "[c:Cat{e2}gorie] puis [c:Sous-cat{e2}gorie] (listes d{e2}roulantes), puis passez {a2} l'op{e2}ration suivante. "
    t = t & "[c:Tiers] et [c:Libell{e2} / Notes] sont modifiables, sauf pour une op{e2}ration de sant{e2}."
    t = t & Chr(10) & "Si des op{e2}rations n'ont pas pu {ea}tre reli{e2}es {a2} une sous-cat{e2}gorie interne, vous serez automatiquement "
    t = t & "repositionn{e2} sur la premi{e1}re d'entre elles."
    t = t & Chr(10) & "Boutons :"
    t = t & Chr(10) & "- [b:< Pr{e2}c{e2}dent] et [b:Suivant >] : naviguent entre les op{e2}rations (les choix d{e2}j{a2} faits sont conserv{e2}s)."
    t = t & Chr(10) & "- [b:Ventiler] : r{e2}partit le [c:Montant] de l'op{e2}ration entre plusieurs sous-cat{e2}gories (ouvre [f:frm_Ventilation])."
    t = t & Chr(10) & "- [b:+] : cr{e2}e une nouvelle cat{e2}gorie ou sous-cat{e2}gorie (ouvre [f:frm_NouvelleCategorie])."
    t = t & Chr(10) & "- [b:Enregistrer] : applique tous les choix faits et poursuit l'import."
    t = t & Chr(10) & "- [b:Annuler] : abandonne l'import en cours."
    t = t & Chr(10) & "ATTENTION : vos choix ne sont pris en compte qu'apr{e1}s un clic sur [b:Enregistrer]."

    TexteInstructions = t

End Function


' =====================================================================================
' Informations de l'opération (lecture seule, sauf Tiers et Libellé / Notes)
' =====================================================================================
' NOTE : Tiers (C9) et Libellé / Notes (C10) sont modifiables, sauf pour une opération de
' santé. Leur fond est donc posé ici en "modifiable" (jaune pâle) puis recalculé à chaque
' affichage par mod_ControleCategories.AfficherOperation (gris = verrouillé).
Private Sub ConstruireInformations(ByVal ws As Worksheet)

    mod_InstallCommun.PoserEtiquette ws.Range("B8"), "Date"
    mod_InstallCommun.PoserEtiquette ws.Range("B9"), "Tiers"
    mod_InstallCommun.PoserEtiquette ws.Range("B10"), mod_Display.FR("Libell{e2} / Notes")
    mod_InstallCommun.PoserEtiquette ws.Range("B11"), "Montant"
    mod_InstallCommun.PoserEtiquette ws.Range("B12"), mod_Display.FR("Cat{e2}gorie source (banque)")

    ' Valeurs en format TEXTE pour éviter toute réinterprétation par Excel.
    mod_InstallCommun.PoserChampLecture ws.Range(CTRL_ADR_DATE), "@"
    mod_InstallCommun.PoserChampSaisie ws.Range(CTRL_ADR_TIERS), "@"
    mod_InstallCommun.PoserChampSaisie ws.Range(CTRL_ADR_LIBELLE), "@", True
    mod_InstallCommun.PoserChampLecture ws.Range(CTRL_ADR_MONTANT), "@"
    ws.Range(CTRL_ADR_MONTANT).Font.Bold = True
    mod_InstallCommun.PoserChampLecture ws.Range(CTRL_ADR_SOURCE), "@"

End Sub


' =====================================================================================
' Zone de saisie : Catégorie et Sous-catégorie, puis message d'aide
' =====================================================================================
' Les listes déroulantes (validation de données) ne sont PAS définies ici : elles
' dépendent du tableau de correspondance et sont donc créées par le programme à
' chaque ouverture du formulaire (voir mod_ControleCategories).
Private Sub ConstruireZoneSaisie(ByVal ws As Worksheet)

    ws.Rows(13).RowHeight = mod_InstallCommun.FRM_H_SEP

    mod_InstallCommun.PoserEtiquette ws.Range("B14"), mod_Display.FR("Cat{e2}gorie")
    mod_InstallCommun.PoserEtiquette ws.Range("B15"), mod_Display.FR("Sous-cat{e2}gorie")
    mod_InstallCommun.PoserChampSaisie ws.Range(CTRL_ADR_CAT), "@"
    mod_InstallCommun.PoserChampSaisie ws.Range(CTRL_ADR_SOUS), "@"

    ws.Rows(16).RowHeight = mod_InstallCommun.FRM_H_SEP

    ' Message d'aide (varie selon l'opération affichée; sa couleur est choisie par le programme).
    mod_InstallCommun.PoserMessage ws.Range("B17:E17")

End Sub


' =====================================================================================
' Boutons
' =====================================================================================
' Largeurs relevées sur la feuille réelle (export du 09/10/2026), hauteur et alignement
' communs (voir mod_InstallCommun.AjouterBoutonEntete). Le nom de macro donné à
' .OnAction correspond à celui de la Public Sub concernée dans mod_ControleCategories.
Private Sub ConstruireBoutons(ByVal ws As Worksheet)

    Dim gauche As Double

    gauche = ws.Cells(1, 2).Left

    mod_InstallCommun.AjouterBoutonEntete ws, CTRL_LIGNE_BOUTONS, gauche, mod_InstallCommun.CapPrecedent(), _
        "ControleOperationPrecedente", "btnCtrlPrecedent", mod_InstallCommun.FRM_BTN_L
    mod_InstallCommun.AjouterBoutonEntete ws, CTRL_LIGNE_BOUTONS, gauche, mod_InstallCommun.CapSuivant(), _
        "ControleOperationSuivante", "btnCtrlSuivant", mod_InstallCommun.FRM_BTN_L
    mod_InstallCommun.AjouterBoutonEntete ws, CTRL_LIGNE_BOUTONS, gauche, mod_InstallCommun.CapEnregistrer(), _
        "ControleTerminer", "btnCtrlTerminer", mod_InstallCommun.FRM_BTN_L
    mod_InstallCommun.AjouterBoutonEntete ws, CTRL_LIGNE_BOUTONS, gauche, mod_InstallCommun.CapAnnuler(), _
        "ControleAnnuler", "btnCtrlAnnuler", mod_InstallCommun.FRM_BTN_L

    ' --- Bouton "Ventiler" (phase 4), à côté du champ Montant ---
    ' Il ouvre la feuille frm_Ventilation (voir mod_InstallVentilation); la logique se
    ' trouve dans mod_ControleCategories.ControleVentiler.
    mod_InstallCommun.AjouterBoutonCellule ws, ws.Range("D11"), "Ventiler", _
        "ControleVentiler", "btnCtrlVentiler", mod_InstallCommun.FRM_BTN_L

    ' --- Bouton "+" (phase 3), à côté du champ Catégorie ---
    ' Il ouvre la feuille frm_NouvelleCategorie (voir mod_InstallNouvelleCategorie).
    mod_InstallCommun.AjouterBoutonCellule ws, ws.Range("D14"), mod_InstallCommun.CapAjouter(), _
        "ControleNouvelleCategorie", "btnCtrlNouvelleCategorie", mod_InstallCommun.FRM_BTN_PLUS

End Sub


' =====================================================================================
' OUTILS DÉVELOPPEUR (Ctrl+G) : afficher ou masquer la feuille pour la retoucher.
' =====================================================================================
Public Sub AfficherFeuilleControlePourEdition()
    mod_InstallCommun.AfficherPourEdition CTRL_NOM_FEUILLE, "CreerFeuilleControleCategories", "MasquerFeuilleControleApresEdition"
End Sub

Public Sub MasquerFeuilleControleApresEdition()
    mod_InstallCommun.MasquerApresEdition CTRL_NOM_FEUILLE
End Sub
