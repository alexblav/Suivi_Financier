Option Explicit

' =====================================================================================
' MODULE : mod_InstallRemediationOFX
'
' REMÉDIATION DEPUIS UN OFX - construction du formulaire. La logique est dans
' mod_RemediationOFX.
'
' Une seule macro à lancer : InstallerRemediationOFX. Elle crée :
'   1. la feuille frm_Remediation_OFX (masquée) : boutons, ligne de résumé, tableau des écarts
'      (rempli à chaque analyse) ;
'   2. sa feuille d'aide frm_Remediation_OFX_aide (texte dans mod_InstallAide) ;
'   3. le bouton "Remédiation OFX" sur la feuille Synthese (sous les autres boutons,
'      déplaçable à la main).
' Aucun code n'est à coller dans la feuille (pas d'événements) : tout passe par des boutons.
'
' INSTALLATION :
'   1. Alt+F11, Fichier > Importer : mod_RemediationOFX.bas, mod_InstallAide.bas (mis à jour),
'      mod_ImportOFX.bas et mod_Categories.bas (mis à jour) puis CE fichier.
'   2. Ctrl+G, taper : InstallerRemediationOFX, puis Entrée.
'   3. Sur Synthese, cliquer sur le bouton "Remédiation OFX".
'
' RÉINSTALLATION possible : le tableau d'écarts affiché est alors effacé (il faut le refaire
' avec Analyser un OFX ; les corrections "Appliqué" ne seraient plus annulables).
' =====================================================================================

' En-tête commun : 1 ligne de boutons + ligne de résumé (voir mod_InstallCommun) :
'   1 marge / 2 boutons / 3 fine ligne / 4 résumé / 5 fine ligne / 6 Instructions (masquée) /
'   7 fine ligne / 8 en-têtes du tableau / 9 première ligne d'écart
Private Const REM_LIGNE_BOUTONS As Long = 2
Private Const REM_LIGNE_INSTRUCTIONS As Long = 6
Private Const REM_COL_DERNIERE_VISIBLE As Long = 11     ' K : Commentaire
Private Const REM_COL_TECH_DEB As Long = 12             ' L : code
Private Const REM_COL_TECH_FIN As Long = 14             ' N : proposé (brut)
Private Const REM_COL_CHEMIN As Long = 16               ' P : chemin de l'OFX analysé (ligne 1)
Private Const SYN_BOUTON_DEFAUT_LIGNE As Long = 31      ' si aucun bouton n'est trouvé sur Synthese


' =====================================================================================
' MACRO PRINCIPALE
' =====================================================================================
Public Sub InstallerRemediationOFX()

    Dim ws As Worksheet
    Dim wsPrecedente As Worksheet
    Dim wsExistante As Worksheet
    Dim derniere As Long, i As Long, nbAppliquees As Long
    Dim gauche As Double

    Set wsPrecedente = ActiveSheet

    ' Garde-fou : ne pas perdre sans le dire la possibilité d'annuler des corrections appliquées.
    Set wsExistante = mod_InstallCommun.TrouverFeuille(mod_RemediationOFX.NOM_FEUILLE_REMEDIATION)
    If Not wsExistante Is Nothing Then
        derniere = wsExistante.Cells(wsExistante.Rows.Count, 4).End(xlUp).Row
        For i = mod_RemediationOFX.REM_LIGNE_ENTETES + 1 To derniere
            If CStr(wsExistante.Cells(i, 3).Value) = "Appliqué" Then nbAppliquees = nbAppliquees + 1
        Next i
        If nbAppliquees > 0 Then
            If MsgBox(nbAppliquees & mod_Display.FR(" correction(s) appliqu{e2}e(s) ne pourront plus {ea}tre annul{e2}es si le formulaire est reconstruit.") & _
                      vbCrLf & mod_Display.FR("Continuer ?"), vbYesNo + vbExclamation, _
                      mod_Display.FR("Donn{e2}es existantes")) = vbNo Then Exit Sub
        End If
    End If

    Set ws = mod_InstallCommun.PreparerFeuille(mod_RemediationOFX.NOM_FEUILLE_REMEDIATION, True)
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    ' Colonnes : A marge / B Corriger / C Statut / D ID / E Date / F Tiers / G Champ / H Valeur actuelle /
    ' I Valeur proposée / J Nature / K Commentaire / L-N techniques (masquées) / O marge / P chemin OFX (masqué)
    mod_InstallCommun.MettreEnForme ws, Array(2, 9, 24, 20, 11, 26, 22, 34, 34, 24, 60, 12, 12, 12, 2, 12), 40

    ConstruireEntete ws
    ConstruireTableau ws
    ConstruireBoutons ws
    FigerVoletsSousEntetes ws

    mod_InstallCommun.TerminerFeuille ws, wsPrecedente
    ConstruireBoutonSynthese
    Application.ScreenUpdating = True

    MsgBox mod_Display.FR("Le formulaire '") & mod_RemediationOFX.NOM_FEUILLE_REMEDIATION & mod_Display.FR("' est install{e2} (masqu{e2}).") & vbCrLf & vbCrLf & _
           mod_Display.FR("Sur Synthese, cliquez sur le bouton 'Rem{e2}diation OFX', puis sur 'Analyser un OFX'."), _
           vbInformation, mod_Display.FR("Installation termin{e2}e")

End Sub


' =====================================================================================
' En-tête commun : hauteurs, ligne de résumé et aide
' =====================================================================================
Private Sub ConstruireEntete(ByVal ws As Worksheet)

    If mod_InstallCommun.PremiereLigneCorps(1, True) <> mod_RemediationOFX.REM_LIGNE_ENTETES Or _
       mod_InstallCommun.LigneInstructions(1, True) <> REM_LIGNE_INSTRUCTIONS Or _
       mod_InstallCommun.LigneCompteur(1, True) <> mod_RemediationOFX.REM_LIGNE_RESUME Then
        MsgBox "mod_InstallRemediationOFX : constantes de lignes incoherentes avec mod_InstallCommun.", vbCritical
    End If

    mod_InstallCommun.PoserHauteursEntete ws, 1, True

    ' La ligne du "compteur" porte ici le résumé de l'analyse (écrit par mod_RemediationOFX).
    mod_InstallCommun.PoserCompteur ws, mod_RemediationOFX.REM_LIGNE_RESUME, 2, REM_COL_DERNIERE_VISIBLE
    ws.Cells(mod_RemediationOFX.REM_LIGNE_RESUME, 2).Value = mod_Display.FR("Aucune analyse : cliquez sur 'Analyser un OFX'.")
    mod_InstallCommun.PoserNom ws, "ResumeRemediationOFX", ws.Cells(mod_RemediationOFX.REM_LIGNE_RESUME, 2)

    mod_InstallCommun.EcrireInstructions ws, REM_LIGNE_INSTRUCTIONS, 2, 3, 4, REM_COL_DERNIERE_VISIBLE, TexteInstructions()

End Sub

' Texte de repli de l'aide (le contenu réel est dans mod_InstallAide.Def_RemediationOFX).
Private Function TexteInstructions() As String
    TexteInstructions = "Compare un OFX r{e2}cent aux op{e2}rations d{e2}j{a2} import{e2}es et affiche les {e2}carts avant correction."
End Function


' =====================================================================================
' Tableau des écarts : en-têtes (le corps est écrit à chaque analyse)
' =====================================================================================
Private Sub ConstruireTableau(ByVal ws As Worksheet)

    Dim l As Long

    l = mod_RemediationOFX.REM_LIGNE_ENTETES

    ws.Range(ws.Cells(l, 2), ws.Cells(l, REM_COL_TECH_FIN)).Value = Array( _
        "Corriger", "Statut", "ID_Transaction", "Date", mod_Display.FR("Tiers (actuel)"), "Champ", "Valeur actuelle", _
        mod_Display.FR("Valeur OFX / propos{e2}e"), "Nature", "Commentaire", "code", "avant (brut)", mod_Display.FR("propos{e2} (brut)"))
    mod_InstallCommun.PoserEnteteTableau ws.Range(ws.Cells(l, 2), ws.Cells(l, REM_COL_DERNIERE_VISIBLE))
    ws.Rows(l).RowHeight = mod_InstallCommun.FRM_H_LIGNE

    ' Colonnes techniques (restauration des valeurs, chemin de l'OFX) : masquées.
    ws.Range(ws.Columns(REM_COL_TECH_DEB), ws.Columns(REM_COL_TECH_FIN)).Hidden = True
    ws.Columns(REM_COL_CHEMIN).Hidden = True

End Sub


' =====================================================================================
' Boutons
' =====================================================================================
Private Sub ConstruireBoutons(ByVal ws As Worksheet)

    Dim gauche As Double

    gauche = ws.Cells(1, 2).Left

    mod_InstallCommun.AjouterBoutonEntete ws, REM_LIGNE_BOUTONS, gauche, "Analyser un OFX", _
        "RemedierDepuisOFX", "btnAnalyserOFX", 100
    mod_InstallCommun.AjouterBoutonEntete ws, REM_LIGNE_BOUTONS, gauche, "Appliquer", _
        "AppliquerRemediationOFX", "btnAppliquerRemediation", mod_InstallCommun.FRM_BTN_L
    mod_InstallCommun.AjouterBoutonEntete ws, REM_LIGNE_BOUTONS, gauche, mod_Display.FR("Annuler les corrections"), _
        "AnnulerRemediationOFX", "btnAnnulerRemediation", 130
    mod_InstallCommun.AjouterBoutonEntete ws, REM_LIGNE_BOUTONS, gauche, mod_InstallCommun.CapSortir(), _
        "SortirRemediationOFX", "btnSortirRemediation", mod_InstallCommun.FRM_BTN_L
    mod_InstallCommun.AjouterBoutonAide ws, REM_LIGNE_BOUTONS, gauche

End Sub


' =====================================================================================
' Figer les volets sous la ligne d'en-têtes du tableau
' =====================================================================================
Private Sub FigerVoletsSousEntetes(ByVal ws As Worksheet)

    ws.Activate
    ws.Range("A" & (mod_RemediationOFX.REM_LIGNE_ENTETES + 1)).Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select

End Sub


' =====================================================================================
' Bouton d'accès sur la feuille Synthese (déplaçable à la main ensuite)
' =====================================================================================
' Placé sous le bouton le plus bas de la colonne des boutons existants, pour ne rien chevaucher.
Private Sub ConstruireBoutonSynthese()

    Dim ws As Worksheet
    Dim btn As Button
    Dim shp As Shape
    Dim gauche As Double, haut As Double, bas As Double, basMax As Double
    Dim trouve As Boolean

    Set ws = mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_SYNTHESE)
    If ws Is Nothing Then
        MsgBox mod_Display.FR("La feuille Synthese est introuvable : le bouton d'acc{e1}s n'a pas {e2}t{e2} cr{e2}{e2}."), vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    ws.Shapes("btnRemediationOFX").Delete
    On Error GoTo 0

    gauche = ws.Cells(1, 1).Left + 12
    For Each shp In ws.Shapes
        If shp.Type = msoFormControl Then
            If Abs(shp.Left - gauche) < 30 Then
                bas = shp.Top + shp.Height
                If bas > basMax Then basMax = bas: trouve = True
            End If
        End If
    Next shp

    If trouve Then
        haut = basMax + 6
    Else
        haut = ws.Cells(SYN_BOUTON_DEFAUT_LIGNE, 1).Top
    End If

    Set btn = ws.Buttons.Add(gauche, haut, 152, 17.2)
    btn.Caption = mod_Display.FR("Rem{e2}diation OFX")
    btn.OnAction = "AfficherRemediationOFX"
    btn.Name = "btnRemediationOFX"

End Sub


' =====================================================================================
' OUTILS DÉVELOPPEUR (Ctrl+G)
' =====================================================================================
Public Sub AfficherFormulaireRemediationPourEdition()
    mod_InstallCommun.AfficherPourEdition mod_RemediationOFX.NOM_FEUILLE_REMEDIATION, "InstallerRemediationOFX", "MasquerFormulaireRemediationApresEdition"
End Sub

Public Sub MasquerFormulaireRemediationApresEdition()
    mod_InstallCommun.MasquerApresEdition mod_RemediationOFX.NOM_FEUILLE_REMEDIATION
End Sub
