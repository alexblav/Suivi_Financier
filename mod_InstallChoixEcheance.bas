Attribute VB_Name = "mod_InstallChoixEcheance"
Option Explicit

' =====================================================================================
' MODULE : mod_InstallChoixEcheance
'
' RÔLE (à lire en premier, même si vous débutez) :
'   Construit la feuille-formulaire frm_ChoixEcheance, qui permet à l'opérateur de désigner
'   À LA MAIN l'échéance correspondant à une opération pendant le contrôle des catégories
'   d'un import (quand la détection automatique ne l'a pas proposée). La logique de l'écran
'   est dans mod_ChoixEcheance ; le rapprochement lui-même dans mod_Echeancier
'   (RapprocherManuellement).
'
' INSTALLATION (une seule fois) :
'   1. Alt+F11, Fichier > Importer : mod_VarGlobales.bas, mod_ChoixEcheance.bas, CE FICHIER,
'      mod_Echeancier.bas, mod_ControleCategories.bas, mod_InstallControleCategories.bas
'      et mod_InstallAide.bas (tous mis à jour).
'   2. Ctrl+G, taper : InstallerChoixEcheance, puis Entrée. La macro :
'        - crée la feuille frm_ChoixEcheance (et son aide) ;
'        - ajoute le bouton « Échéancier... » sur frm_ControleCategories SANS reconstruire
'          cette feuille, et met son aide à jour.
'   3. Coller dans le module de la feuille frm_ChoixEcheance le code de
'      CodeBehind_frm_ChoixEcheance.txt (le message de fin d'installation donne son CodeName).
'
' RÉINSTALLATION : possible (la feuille est reconstruite après confirmation).
'
' À PROPOS DES ACCENTS : textes affichés via mod_Display.FR() ; commentaires en UTF-8.
' =====================================================================================

Private Const LIGNE_BOUTONS As Long = 2
Private Const LIGNE_INSTRUCTIONS As Long = 4
Private Const LIGNE_CORPS As Long = 6


' =====================================================================================
' MACRO PRINCIPALE
' =====================================================================================
Public Sub InstallerChoixEcheance()

    Dim wsPrecedente As Worksheet
    Dim ws As Worksheet

    Set wsPrecedente = ActiveSheet

    If Not ConstruireFormulaire(wsPrecedente) Then Exit Sub
    mod_InstallControleCategories.AjouterBoutonEcheancierControle

    Set ws = mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_CHOIX_ECHEANCE)
    MsgBox mod_Display.FR("Le choix manuel d'une {e2}ch{e2}ance est install{e2}.") & vbCrLf & vbCrLf & _
           mod_Display.FR("PROCHAINE {E2}TAPE (indispensable) : coller le code de verrouillage dans la feuille du formulaire.") & vbCrLf & _
           "Nom interne (CodeName) de la feuille " & mod_VarGlobales.NOM_FEUILLE_CHOIX_ECHEANCE & " : " & ws.CodeName & vbCrLf & vbCrLf & _
           mod_Display.FR("Voir le fichier CodeBehind_frm_ChoixEcheance.txt."), _
           vbInformation, mod_Display.FR("Installation termin{e2}e")

End Sub


' =====================================================================================
' FORMULAIRE frm_ChoixEcheance
' =====================================================================================
Private Function ConstruireFormulaire(ByVal wsPrecedente As Worksheet) As Boolean

    Dim ws As Worksheet
    Dim gauche As Double
    Dim j As Long
    Dim colLettre As String, colLibelles As String

    Set ws = mod_InstallCommun.PreparerFeuille(mod_VarGlobales.NOM_FEUILLE_CHOIX_ECHEANCE, True)
    If ws Is Nothing Then Exit Function

    Application.ScreenUpdating = False

    ' Colonnes : marge / étiquettes / valeurs / suite des valeurs / marge
    mod_InstallCommun.MettreEnForme ws, Array(2, 26, 44, 22, 22, 2), 30

    If mod_InstallCommun.PremiereLigneCorps(1, False) <> LIGNE_CORPS Or _
       mod_InstallCommun.LigneInstructions(1, False) <> LIGNE_INSTRUCTIONS Then
        MsgBox "mod_InstallChoixEcheance : constantes de lignes incoherentes avec mod_InstallCommun.", vbCritical
    End If
    mod_InstallCommun.PoserHauteursEntete ws, 1, False
    mod_InstallCommun.EcrireInstructions ws, LIGNE_INSTRUCTIONS, 2, 2, 3, 5, TexteInstructions()

    ' --- Opération importée (lecture seule) ---
    mod_InstallCommun.PoserTitreBloc ws.Range("B6:E6"), mod_Display.FR("Op{e2}ration import{e2}e")
    mod_InstallCommun.PoserEtiquette ws.Range("B7"), "Date"
    mod_InstallCommun.PoserChampLecture ws.Range("C7:E7"), "@"
    mod_InstallCommun.PoserEtiquette ws.Range("B8"), "Tiers"
    mod_InstallCommun.PoserChampLecture ws.Range("C8:E8"), "@"
    mod_InstallCommun.PoserEtiquette ws.Range("B9"), "Montant"
    mod_InstallCommun.PoserChampLecture ws.Range("C9:E9"), "@"
    ws.Range(mod_ChoixEcheance.CE_ADR_MONTANT).Font.Bold = True
    mod_InstallCommun.PoserEtiquette ws.Range("B10"), mod_Display.FR("Cat{e2}gorie")
    mod_InstallCommun.PoserChampLecture ws.Range("C10:E10"), "@"
    mod_InstallCommun.PoserEtiquette ws.Range("B11"), mod_Display.FR("Ch{e1}que")
    mod_InstallCommun.PoserChampLecture ws.Range("C11:E11"), "@"
    ws.Rows(12).RowHeight = mod_InstallCommun.FRM_H_SEP

    ' --- Choix de l'échéance ---
    mod_InstallCommun.PoserTitreBloc ws.Range("B13:E13"), mod_Display.FR("{E2}ch{e2}ance {a2} rapprocher")
    mod_InstallCommun.PoserEtiquette ws.Range("B14"), mod_Display.FR("{E2}ch{e2}ance")
    mod_InstallCommun.PoserChampSaisie ws.Range("C14:E14"), "@"

    ' --- Détail de l'échéance choisie : FORMULES qui relisent la zone technique masquée ---
    ' (aucun événement de feuille nécessaire : le détail suit le choix tout seul)
    mod_InstallCommun.PoserEtiquette ws.Range("B15"), mod_Display.FR("Libell{e2}")
    mod_InstallCommun.PoserEtiquette ws.Range("B16"), "Tiers"
    mod_InstallCommun.PoserEtiquette ws.Range("B17"), mod_Display.FR("Cat{e2}gorie")
    mod_InstallCommun.PoserEtiquette ws.Range("B18"), mod_Display.FR("Montant pr{e2}vu")
    mod_InstallCommun.PoserEtiquette ws.Range("B19"), mod_Display.FR("Date pr{e2}vue")
    mod_InstallCommun.PoserEtiquette ws.Range("B20"), mod_Display.FR("P{e2}riodicit{e2}")
    mod_InstallCommun.PoserEtiquette ws.Range("B21"), mod_Display.FR("Pourquoi pas d{e2}tect{e2}e")
    colLibelles = Split(ws.Cells(1, mod_ChoixEcheance.CE_COL_LISTE).Address(True, False), "$")(0)
    For j = 1 To mod_ChoixEcheance.CE_NB_DETAILS
        mod_InstallCommun.PoserChampLecture ws.Range(ws.Cells(mod_ChoixEcheance.CE_LIGNE_DETAIL + j - 1, 3), _
                                                     ws.Cells(mod_ChoixEcheance.CE_LIGNE_DETAIL + j - 1, 5)), "", (j = mod_ChoixEcheance.CE_NB_DETAILS)
        colLettre = Split(ws.Cells(1, mod_ChoixEcheance.CE_COL_LISTE + j).Address(True, False), "$")(0)
        ws.Cells(mod_ChoixEcheance.CE_LIGNE_DETAIL + j - 1, 3).Formula = _
            "=IFERROR(INDEX(" & colLettre & "$2:" & colLettre & "$" & mod_ChoixEcheance.CE_LIGNE_MAX & "," & _
            "MATCH($C$14,$" & colLibelles & "$2:$" & colLibelles & "$" & mod_ChoixEcheance.CE_LIGNE_MAX & ",0)),"""")"
    Next j
    ws.Rows(mod_ChoixEcheance.CE_LIGNE_DETAIL + mod_ChoixEcheance.CE_NB_DETAILS - 1).RowHeight = 40
    ws.Rows(22).RowHeight = mod_InstallCommun.FRM_H_SEP

    mod_InstallCommun.PoserMessage ws.Range("B23:E23")

    ' --- Zone technique (libellés + détails) : colonnes masquées ---
    ws.Cells(1, mod_ChoixEcheance.CE_COL_LISTE).Value = "LibellesEcheances"
    For j = 0 To mod_ChoixEcheance.CE_NB_DETAILS
        ws.Columns(mod_ChoixEcheance.CE_COL_LISTE + j).Hidden = True
    Next j

    ' --- Boutons de la ligne d'en-tête ---
    gauche = ws.Cells(1, 2).Left
    mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, "Rapprocher", _
        "ChoixEcheanceRapprocher", "btnChoixRapprocher", mod_InstallCommun.FRM_BTN_L
    mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, mod_Display.FR("D{e2}faire"), _
        "ChoixEcheanceDefaire", "btnChoixDefaire", mod_InstallCommun.FRM_BTN_L
    mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, mod_InstallCommun.CapAnnuler(), _
        "ChoixEcheanceAnnuler", "btnChoixAnnuler", mod_InstallCommun.FRM_BTN_L
    mod_InstallCommun.AjouterBoutonAide ws, LIGNE_BOUTONS, gauche

    mod_InstallCommun.TerminerFeuille ws, wsPrecedente
    Application.ScreenUpdating = True
    ConstruireFormulaire = True

End Function

Private Function TexteInstructions() As String

    Dim t As String

    t = "Ce formulaire permet de d{e2}signer {a2} la main l'{e2}ch{e2}ance de l'{e2}ch{e2}ancier qui correspond {a2} l'op{e2}ration affich{e2}e dans [f:frm_ControleCategories], "
    t = t & "quand la d{e2}tection automatique ne l'a pas propos{e2}e (par exemple parce que l'{e2}cart de montant ou de date est trop grand)."
    t = t & Chr(10) & "Le haut rappelle l'op{e2}ration import{e2}e. Elle n'est jamais modifi{e2}e : elle vient de la banque."
    t = t & Chr(10) & "[c:{E2}ch{e2}ance] : liste d{e2}roulante des {e2}ch{e2}ances de l'{e2}ch{e2}ancier, de la plus proche {a2} la plus {e2}loign{e2}e de l'op{e2}ration. "
    t = t & "Chaque entr{e2}e indique son num{e2}ro, sa date, son tiers, son montant et les {e2}carts de date et de montant avec l'op{e2}ration. "
    t = t & "Les {e2}ch{e2}ances d{e2}j{a2} rapproch{e2}es avec une autre op{e2}ration de cet import ne sont pas propos{e2}es."
    t = t & Chr(10) & "Les lignes en dessous d{e2}taillent l'{e2}ch{e2}ance choisie et expliquent pourquoi la d{e2}tection automatique ne l'a pas retenue."
    t = t & Chr(10) & "Quand l'{e2}ch{e2}ance vient d'une r{e1}gle de r{e2}currence et que l'op{e2}ration en diff{e1}re (montant, tiers, cat{e2}gorie, sous-cat{e2}gorie), "
    t = t & "une question vous est pos{e2}e apr{e1}s [b:Rapprocher] : mettre {a2} jour la r{e1}gle et ses {e2}ch{e2}ances {a2} venir avec les valeurs de l'op{e2}ration, ou non."
    t = t & Chr(10) & "Comme pour un rapprochement automatique, l'{e2}ch{e2}ance n'est retir{e2}e de l'{e2}ch{e2}ancier (et la r{e1}gle mise {a2} jour) qu'{a2} l'enregistrement de l'import ; "
    t = t & "si l'import est abandonn{e2}, l'{e2}ch{e2}ancier reste intact."
    t = t & Chr(10) & "Boutons :"
    t = t & Chr(10) & "- [b:Rapprocher] : lie l'op{e2}ration {a2} l'{e2}ch{e2}ance choisie et revient au contr{o2}le des cat{e2}gories."
    t = t & Chr(10) & "- [b:D{e2}faire] : annule le rapprochement d{e2}j{a2} fait pour cette op{e2}ration."
    t = t & Chr(10) & "- [b:Annuler] : revient au contr{o2}le des cat{e2}gories sans rien changer."

    TexteInstructions = t

End Function


' =====================================================================================
' OUTILS DÉVELOPPEUR (Ctrl+G)
' =====================================================================================
Public Sub AfficherFormulaireChoixEcheancePourEdition()
    mod_InstallCommun.AfficherPourEdition mod_VarGlobales.NOM_FEUILLE_CHOIX_ECHEANCE, "InstallerChoixEcheance", "MasquerFormulaireChoixEcheanceApresEdition"
End Sub

Public Sub MasquerFormulaireChoixEcheanceApresEdition()
    mod_InstallCommun.MasquerApresEdition mod_VarGlobales.NOM_FEUILLE_CHOIX_ECHEANCE
End Sub
