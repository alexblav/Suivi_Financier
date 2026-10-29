Attribute VB_Name = "mod_InstallAffectationFacture"
Option Explicit

' =====================================================================================
' MODULE : mod_InstallAffectationFacture
'
' ROLE (ajout 29/10/2026) :
'   Construit la feuille masquee frm_AffectationFacture, ouverte par le bouton "Affecter une
'   facture" de la feuille Synthese (tableau "Factures non liees") pour relier une operation
'   bancaire (Eau, Electricite - gaz - chauffage) a la facture qui l'a reglee.
'   Construite avec les outils COMMUNS de mod_InstallCommun, sur le modele de
'   frm_RapprochementNotes. Le fonctionnement est dans mod_AffectationFacture, qui retrouve
'   les champs par leurs NOMS DEFINIS (afgType, afgFacture...) et non par leurs adresses.
'
'   Corps du formulaire :
'     - "Operation a affecter" : date, tiers, categorie, montant (lecture seule) ;
'     - "Recherche de la facture (filtres progressifs)" : 1. Type de consommation,
'       2. Date de la facture (la plus recente en premier) ;
'     - "Facture retenue" : rappel de la facture choisie (lecture seule).
'   Boutons : Nouvelle facture, Valider, Annuler, Aide.
'
' A FAIRE POUR INSTALLER CE MODULE :
'   1. Alt+F11, Fichier > Importer un fichier... : mod_InstallAffectationFacture.bas,
'      mod_AffectationFacture.bas, puis les versions a jour de mod_Factures.bas,
'      mod_Synthese.bas, mod_InstallAide.bas et mod_InstallConsommations.bas.
'   2. Dans la fenetre Execution immediate (Ctrl+G), lancer :
'        CreerFeuilleAffectationFacture
'      (cree aussi la feuille d'aide frm_AffectationFacture_aide, texte dans mod_InstallAide).
'
' FICHIER 100 % ASCII : les textes accentues passent par mod_Display.FR().
' =====================================================================================

Public Const NOM_FEUILLE_AFFECTATION As String = "frm_AffectationFacture"

' Colonnes du formulaire (trois paires libelle / valeur, comme frm_RapprochementNotes).
Private Const AF_COL_LIBELLE_1 As String = "B"
Private Const AF_COL_VALEUR_1 As String = "C"
Private Const AF_COL_LIBELLE_2 As String = "D"
Private Const AF_COL_VALEUR_2 As String = "E"
Private Const AF_COL_LIBELLE_3 As String = "F"
Private Const AF_COL_VALEUR_3 As String = "G"

' En-tete commun (voir mod_InstallCommun) : 1 ligne de boutons, pas de compteur.
'   ligne 1 marge / 2 boutons / 3 fine ligne / 4 Instructions (masquee) / 5 fine ligne / 6 et suivantes : corps
Private Const AF_LIGNE_BOUTONS As Long = 2
Private Const AF_LIGNE_INSTRUCTIONS As Long = 4
Private Const AF_LIGNE_CORPS As Long = 6
Private Const AF_LIGNE_TITRE_OPERATION As Long = 6
Private Const AF_LIGNE_OPERATION As Long = 7            ' Date operation / Tiers
Private Const AF_LIGNE_CATEGORIE As Long = 8            ' Categorie / Montant
Private Const AF_LIGNE_TITRE_RECHERCHE As Long = 10
Private Const AF_LIGNE_TYPE As Long = 11
Private Const AF_LIGNE_FACTURE As Long = 12
Private Const AF_LIGNE_RETENUE As Long = 14

' Colonnes techniques masquees : listes proposees par les deux filtres, reecrites par mod_AffectationFacture.
Public Const AF_COL_LISTE_TYPES As String = "L"
Public Const AF_COL_LISTE_FACTURES As String = "M"


' =====================================================================================
' MACRO D'INSTALLATION
' =====================================================================================
Public Sub CreerFeuilleAffectationFacture()

    Dim ws As Worksheet
    Dim wsPrecedente As Worksheet

    Set wsPrecedente = ActiveSheet

    Set ws = mod_InstallCommun.PreparerFeuille(NOM_FEUILLE_AFFECTATION, True)
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    ' Colonnes A a M : marge / B a G (formulaire) / marge / I a K libres / L et M listes techniques
    mod_InstallCommun.MettreEnForme ws, Array(2, 21, 22.33, 14, 25.87, 14, 16, 2, 10, 10, 10, 12, 60)

    AF_ConstruireEntete ws
    AF_ConstruireBlocOperation ws
    AF_ConstruireBlocRecherche ws
    AF_ConstruireZoneTechnique ws
    AF_ConstruireBoutons ws     ' en dernier : les boutons se placent d'apres les hauteurs de lignes

    mod_InstallCommun.TerminerFeuille ws, wsPrecedente
    Application.ScreenUpdating = True

    MsgBox "La feuille '" & NOM_FEUILLE_AFFECTATION & mod_Display.FR("' a {e2}t{e2} cr{e2}{e2}e et masqu{e2}e."), _
           vbInformation, mod_Display.FR("Installation termin{e2}e")

End Sub

Private Sub AF_ConstruireEntete(ByVal ws As Worksheet)

    If mod_InstallCommun.PremiereLigneCorps(1, False) <> AF_LIGNE_CORPS Or _
       mod_InstallCommun.LigneInstructions(1, False) <> AF_LIGNE_INSTRUCTIONS Then
        MsgBox "mod_InstallAffectationFacture : constantes de lignes AF incoherentes avec mod_InstallCommun.", vbCritical
    End If

    mod_InstallCommun.PoserHauteursEntete ws, 1, False
    mod_InstallCommun.EcrireInstructions ws, AF_LIGNE_INSTRUCTIONS, 2, 2, 3, 7, AF_TexteInstructions()

End Sub

' Texte repris tel quel si mod_InstallAide n'a pas d'aide pour ce formulaire (il en a une : Def_AffectationFacture).
Private Function AF_TexteInstructions() As String

    Dim t As String

    t = "Ce formulaire relie l'op{e2}ration bancaire affich{e2}e {a2} la facture qui l'a r{e2}gl{e2}e. "
    t = t & "Choisissez le [c:Type de consommation], puis la facture dans [c:Date de la facture] (la plus r{e2}cente en premier), puis cliquez sur [b:Valider]."
    t = t & Chr(10) & "Plusieurs op{e2}rations peuvent {ea}tre li{e2}es {a2} une m{ea}me facture (r{e2}glement en plusieurs {e2}ch{e2}ances)."
    t = t & Chr(10) & "- [b:Nouvelle facture] : aucune facture ne convient, ouvre la saisie d'une facture d{e2}j{a2} li{e2}e {a2} l'op{e2}ration."
    t = t & Chr(10) & "- [b:Annuler] : ferme le formulaire sans rien modifier."

    AF_TexteInstructions = t

End Function

' Bloc "Operation a affecter" (lecture seule)
'   7 : Date operation | Tiers (valeur fusionnee de E a G)
'   8 : Categorie (valeur fusionnee de C a E) | Montant
Private Sub AF_ConstruireBlocOperation(ByVal ws As Worksheet)

    mod_InstallCommun.PoserTitreBloc ws.Range(AF_COL_LIBELLE_1 & AF_LIGNE_TITRE_OPERATION & ":" & AF_COL_VALEUR_3 & AF_LIGNE_TITRE_OPERATION), _
                                     mod_Display.FR("Op{e2}ration {a2} affecter")

    mod_InstallCommun.PoserEtiquette ws.Range(AF_COL_LIBELLE_1 & AF_LIGNE_OPERATION), mod_Display.FR("Date op{e2}ration")
    mod_InstallCommun.PoserEtiquette ws.Range(AF_COL_LIBELLE_2 & AF_LIGNE_OPERATION), "Tiers"
    mod_InstallCommun.PoserChampLecture ws.Range(AF_COL_VALEUR_1 & AF_LIGNE_OPERATION), "dd/mm/yyyy"
    mod_InstallCommun.PoserChampLecture ws.Range(AF_COL_VALEUR_2 & AF_LIGNE_OPERATION & ":" & AF_COL_VALEUR_3 & AF_LIGNE_OPERATION), "@"
    mod_InstallCommun.PoserNom ws, "afgDateOp", ws.Range(AF_COL_VALEUR_1 & AF_LIGNE_OPERATION)
    mod_InstallCommun.PoserNom ws, "afgTiersOp", ws.Range(AF_COL_VALEUR_2 & AF_LIGNE_OPERATION)

    mod_InstallCommun.PoserEtiquette ws.Range(AF_COL_LIBELLE_1 & AF_LIGNE_CATEGORIE), mod_Display.FR("Cat{e2}gorie")
    mod_InstallCommun.PoserEtiquette ws.Range(AF_COL_LIBELLE_3 & AF_LIGNE_CATEGORIE), "Montant"
    mod_InstallCommun.PoserChampLecture ws.Range(AF_COL_VALEUR_1 & AF_LIGNE_CATEGORIE & ":" & AF_COL_VALEUR_2 & AF_LIGNE_CATEGORIE), "@"
    mod_InstallCommun.PoserChampLecture ws.Range(AF_COL_VALEUR_3 & AF_LIGNE_CATEGORIE), "#,##0.00 " & ChrW(8364)
    mod_InstallCommun.PoserNom ws, "afgCategorie", ws.Range(AF_COL_VALEUR_1 & AF_LIGNE_CATEGORIE)
    mod_InstallCommun.PoserNom ws, "afgMontantOp", ws.Range(AF_COL_VALEUR_3 & AF_LIGNE_CATEGORIE)

    ws.Rows(AF_LIGNE_CATEGORIE + 1).RowHeight = mod_InstallCommun.FRM_H_SEP

End Sub

' Bloc "Recherche de la facture" (deux filtres progressifs) et "Facture retenue".
Private Sub AF_ConstruireBlocRecherche(ByVal ws As Worksheet)

    mod_InstallCommun.PoserTitreBloc ws.Range(AF_COL_LIBELLE_1 & AF_LIGNE_TITRE_RECHERCHE & ":" & AF_COL_VALEUR_3 & AF_LIGNE_TITRE_RECHERCHE), _
                                     mod_Display.FR("Recherche de la facture (filtres progressifs)")

    AF_PoserFiltre ws, AF_LIGNE_TYPE, "Type de consommation", "afgType"
    AF_PoserFiltre ws, AF_LIGNE_FACTURE, "Date de la facture", "afgFacture"

    ws.Rows(AF_LIGNE_FACTURE + 1).RowHeight = mod_InstallCommun.FRM_H_SEP

    mod_InstallCommun.PoserEtiquette ws.Range(AF_COL_LIBELLE_1 & AF_LIGNE_RETENUE), mod_Display.FR("Facture retenue")
    mod_InstallCommun.PoserChampLecture ws.Range(AF_COL_VALEUR_1 & AF_LIGNE_RETENUE & ":" & AF_COL_VALEUR_3 & AF_LIGNE_RETENUE), "@"
    ws.Range(AF_COL_VALEUR_1 & AF_LIGNE_RETENUE).Font.Bold = True
    mod_InstallCommun.PoserNom ws, "afgFactureRetenue", ws.Range(AF_COL_VALEUR_1 & AF_LIGNE_RETENUE)

End Sub

' Un filtre : etiquette en B, zone de saisie fusionnee de C a F (liste), nom sur la 1re cellule.
' Format Texte ("@") : une date choisie dans la liste ne doit pas etre reconvertie par Excel.
Private Sub AF_PoserFiltre(ByVal ws As Worksheet, ByVal ligne As Long, ByVal libelle As String, ByVal nomCellule As String)
    mod_InstallCommun.PoserEtiquette ws.Range(AF_COL_LIBELLE_1 & ligne), libelle
    mod_InstallCommun.PoserChampSaisie ws.Range(AF_COL_VALEUR_1 & ligne & ":" & AF_COL_LIBELLE_3 & ligne), "@"
    mod_InstallCommun.PoserNom ws, nomCellule, ws.Range(AF_COL_VALEUR_1 & ligne)
End Sub

' Zone technique masquee : une cellule de reserve par liste (mod_AffectationFacture la redimensionne).
' Ces noms DOIVENT etre au niveau de la FEUILLE (PoserNom).
Private Sub AF_ConstruireZoneTechnique(ByVal ws As Worksheet)

    ws.Range(AF_COL_LISTE_TYPES & "1").Value = "ListeTypes"
    ws.Range(AF_COL_LISTE_FACTURES & "1").Value = "ListeFactures"
    mod_InstallCommun.PoserNom ws, "afgListeTypes", ws.Range(AF_COL_LISTE_TYPES & "2")
    mod_InstallCommun.PoserNom ws, "afgListeFactures", ws.Range(AF_COL_LISTE_FACTURES & "2")
    ws.Range(AF_COL_LISTE_TYPES & ":" & AF_COL_LISTE_FACTURES).Columns.Hidden = True

End Sub

Private Sub AF_ConstruireBoutons(ByVal ws As Worksheet)

    Dim gauche As Double

    gauche = ws.Cells(1, 2).Left
    mod_InstallCommun.AjouterBoutonEntete ws, AF_LIGNE_BOUTONS, gauche, mod_Display.FR("Nouvelle facture"), _
        "AffNouvelleFacture", "btnAffNouvelle", 100
    mod_InstallCommun.AjouterBoutonEntete ws, AF_LIGNE_BOUTONS, gauche, mod_InstallCommun.CapValider(), _
        "AffValider", "btnAffValider", mod_InstallCommun.FRM_BTN_L
    mod_InstallCommun.AjouterBoutonEntete ws, AF_LIGNE_BOUTONS, gauche, mod_InstallCommun.CapAnnuler(), _
        "AffAnnuler", "btnAffAnnuler", mod_InstallCommun.FRM_BTN_L
    mod_InstallCommun.AjouterBoutonAide ws, AF_LIGNE_BOUTONS, gauche

End Sub
