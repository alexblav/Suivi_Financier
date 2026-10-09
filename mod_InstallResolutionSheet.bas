Option Explicit

' =====================================================================================
' MODULE : mod_InstallResolutionSheet
'
' RÔLE (phase 1 du chantier "UserForm -> feuille dédiée") :
'   Ce module contient UNIQUEMENT la construction de la mise en page statique de
'   la feuille qui remplace le UserForm "frmResolutionCategories".
'   La logique de remplissage des cas ambigus et de gestion des clics se trouve dans
'   le module mod_ResolutionCategories (et le code-behind de la feuille).
'
'   Pourquoi séparer ainsi ? Pour construire les murs (ce module) avant de brancher
'   l'électricité (mod_ResolutionCategories).
'
' REFONTE DU 09/10/2026 (harmonisation des formulaires) :
'   Construite avec les outils COMMUNS de mod_InstallCommun. Le compteur, qui était à
'   droite de la ligne de boutons, est désormais sur sa propre ligne (format commun
'   "Opération: x/y", voir mod_ResolutionCategories.MettreAJourCompteur). Le bloc
'   Instructions (étiquette + texte sur une ligne) est placé sous la ligne de boutons.
'   En conséquence, l'en-tête du tableau passe de la ligne 11 à la ligne 8, et la
'   première ligne de données de la ligne 12 à la ligne 9 : voir les constantes
'   LIGNE_ENTETES_TABLEAU et LIGNE_PREMIERE_DONNEE de mod_VarGlobales.
'   Les couleurs de fond des lignes de données (gris = lecture seule, jaune = choix
'   de la catégorie) sont posées par le programme à chaque ouverture, uniquement sur
'   les lignes réellement utilisées.
'
' À FAIRE POUR INSTALLER CE MODULE :
'   1. Alt+F11 pour ouvrir l'éditeur VBA.
'   2. Fichier > Importer un fichier... : mod_InstallCommun.bas, mod_VarGlobales.bas
'      (mis à jour), mod_ResolutionCategories.bas (mis à jour) puis ce fichier.
'   3. Dans la fenêtre Exécution immédiate (Ctrl+G), taper :
'        CreerFeuilleResolutionCategories
'      puis appuyer sur Entrée. La feuille est créée, mise en forme, puis masquée.
' =====================================================================================

' -------------------------------------------------------------------------------------
' CONSTANTES DE MISE EN PAGE
' -------------------------------------------------------------------------------------
' Le nom de la feuille (NOM_FEUILLE_RESOLUTION), les colonnes du tableau (COL_STATUT...) et
' ses lignes (LIGNE_ENTETES_TABLEAU, LIGNE_PREMIERE_DONNEE, NB_LIGNES_PREPAREES) sont dans
' mod_VarGlobales, car mod_ResolutionCategories et le code-behind de la feuille les utilisent.
'
' En-tête commun (voir mod_InstallCommun) : 1 ligne de boutons + compteur.
'   ligne 1 marge / 2 boutons / 3 fine ligne / 4 compteur / 5 fine ligne /
'   6 Instructions / 7 fine ligne / 8 en-têtes du tableau / 9 premier cas
Private Const RES_LIGNE_BOUTONS As Long = 2
Private Const RES_LIGNE_COMPTEUR As Long = 4
Private Const RES_LIGNE_INSTRUCTIONS As Long = 6


' =====================================================================================
' MACRO PRINCIPALE D'INSTALLATION
' À exécuter UNE SEULE FOIS (ou de nouveau pour réinitialiser complètement la mise
' en forme de la feuille).
' =====================================================================================
Sub CreerFeuilleResolutionCategories()

    Dim ws As Worksheet
    Dim wsPrecedente As Worksheet

    Set wsPrecedente = ActiveSheet

    Set ws = mod_InstallCommun.PreparerFeuille(NOM_FEUILLE_RESOLUTION, True)
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    ' Colonnes : marge / Statut / Date / Montant / Tiers / Catégorie(s) / marge
    mod_InstallCommun.MettreEnForme ws, Array(2, 8, 12, 12, 28, 24, 2), LIGNE_PREMIERE_DONNEE + NB_LIGNES_PREPAREES

    ConstruireEntete ws
    ConstruireTableauCas ws
    ConstruireZoneBoutons ws    ' en dernier : les boutons se placent d'après les hauteurs de lignes
    FigerVoletsSousEntetes ws

    mod_InstallCommun.TerminerFeuille ws, wsPrecedente
    Application.ScreenUpdating = True

    MsgBox mod_Display.FR("La feuille '") & NOM_FEUILLE_RESOLUTION & mod_Display.FR("' a {e2}t{e2} cr{e2}{e2}e et mise en forme,") & vbCrLf & _
           mod_Display.FR("puis masqu{e2}e (xlSheetVeryHidden).") & vbCrLf & vbCrLf & _
           mod_Display.FR("Pour la revoir {a2} l'{e2}cran et v{e2}rifier le rendu, ex{e2}cutez la macro :") & vbCrLf & _
           "   AfficherFeuilleResolutionPourEdition", _
           vbInformation, mod_Display.FR("Installation termin{e2}e")

End Sub


' =====================================================================================
' En-tête commun : hauteurs, compteur et bloc Instructions
' =====================================================================================
Private Sub ConstruireEntete(ByVal ws As Worksheet)

    If mod_InstallCommun.PremiereLigneCorps(1, True) <> LIGNE_ENTETES_TABLEAU Or _
       mod_InstallCommun.LigneInstructions(1, True) <> RES_LIGNE_INSTRUCTIONS Or _
       mod_InstallCommun.LigneCompteur(1, True) <> RES_LIGNE_COMPTEUR Then
        MsgBox "mod_InstallResolutionSheet : constantes de lignes incoherentes avec mod_InstallCommun.", vbCritical
    End If

    mod_InstallCommun.PoserHauteursEntete ws, 1, True

    ' Compteur "Opération: x/y", rempli par mod_ResolutionCategories.MettreAJourCompteur.
    ' On lui attribue un NOM DÉFINI (plage nommée) plutôt qu'une référence codée en dur :
    ' Range("CompteurCasRestants") est plus explicite que Range("B4").
    mod_InstallCommun.PoserCompteur ws, RES_LIGNE_COMPTEUR, 2, 6
    mod_InstallCommun.PoserNom ws, "CompteurCasRestants", ws.Cells(RES_LIGNE_COMPTEUR, 2)

    ' Étiquette sur les colonnes B:C (la colonne B seule, "Statut", est trop étroite),
    ' texte sur D:F.
    mod_InstallCommun.EcrireInstructions ws, RES_LIGNE_INSTRUCTIONS, 2, 3, 4, 6, TexteInstructions()

End Sub

Private Function TexteInstructions() As String

    Dim t As String

    t = "Les op{e2}rations list{e2}es n'ont pas pu {ea}tre cat{e2}goris{e2}es automatiquement : ce formulaire permet de le faire manuellement."
    t = t & Chr(10) & "Pour chaque op{e2}ration, choisissez une cat{e2}gorie dans la liste d{e2}roulante de la colonne [c:Cat{e2}gorie(s)]. "
    t = t & "Le [c:Statut] de la ligne passe automatiquement {a2} une coche verte d{e1}s que la cat{e2}gorie est choisie. "
    t = t & "Le compteur indique le nombre d'op{e2}rations d{e2}j{a2} trait{e2}es sur le total."
    t = t & Chr(10) & "Boutons :"
    t = t & Chr(10) & "- [b:R{e2}initialiser cette ligne] : efface le choix de la ligne s{e2}lectionn{e2}e pour recommencer."
    t = t & Chr(10) & "- [b:Enregistrer] : applique les choix faits et poursuit l'import. Les op{e2}rations sans cat{e2}gorie "
    t = t & "resteront sans cat{e2}gorie (un message de confirmation est demand{e2})."

    TexteInstructions = t

End Function


' =====================================================================================
' Construction des en-têtes et mise en forme du tableau
' =====================================================================================
Private Sub ConstruireTableauCas(ByVal ws As Worksheet)

    Dim derniereLigne As Long

    ' --- Ligne d'en-têtes (LIGNE_ENTETES_TABLEAU) ---
    ws.Range(COL_STATUT & LIGNE_ENTETES_TABLEAU).Value = "Statut"
    ws.Range(COL_DATE & LIGNE_ENTETES_TABLEAU).Value = "Date"
    ws.Range(COL_MONTANT & LIGNE_ENTETES_TABLEAU).Value = "Montant"
    ws.Range(COL_TIERS & LIGNE_ENTETES_TABLEAU).Value = "Tiers"
    ws.Range(COL_CATEGORIE & LIGNE_ENTETES_TABLEAU).Value = mod_Display.FR("Cat{e2}gorie(s)")
    mod_InstallCommun.PoserEnteteTableau ws.Range(COL_STATUT & LIGNE_ENTETES_TABLEAU & ":" & COL_CATEGORIE & LIGNE_ENTETES_TABLEAU)
    ws.Rows(LIGNE_ENTETES_TABLEAU).RowHeight = mod_InstallCommun.FRM_H_LIGNE

    ' --- Mise en forme "à blanc" des lignes de données préparées à l'avance ---
    ' On ne remplit PAS de données réelles ici (c'est fait à chaque ouverture, selon le
    ' nombre réel de cas ambigus). On prépare uniquement l'apparence (liseré léger,
    ' alignements, formats). Les COULEURS DE FOND (gris = lecture seule, jaune = choix
    ' de la catégorie) sont posées par mod_ResolutionCategories.RemplirTableauCas sur
    ' les seules lignes utilisées.
    derniereLigne = LIGNE_PREMIERE_DONNEE + NB_LIGNES_PREPAREES - 1

    With ws.Range(COL_STATUT & LIGNE_PREMIERE_DONNEE & ":" & COL_CATEGORIE & derniereLigne)
        .Font.Size = 10
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = mod_InstallCommun.CoulBordureLecture()
        .Borders(xlInsideHorizontal).Color = mod_InstallCommun.CoulBordureLecture()
    End With

    ws.Range(COL_STATUT & LIGNE_PREMIERE_DONNEE & ":" & COL_STATUT & derniereLigne).HorizontalAlignment = xlCenter
    ws.Range(COL_MONTANT & LIGNE_PREMIERE_DONNEE & ":" & COL_MONTANT & derniereLigne).HorizontalAlignment = xlRight
    ws.Range(COL_MONTANT & LIGNE_PREMIERE_DONNEE & ":" & COL_MONTANT & derniereLigne).NumberFormat = "#,##0.00 " & ChrW(8364)
    ws.Range(COL_DATE & LIGNE_PREMIERE_DONNEE & ":" & COL_DATE & derniereLigne).NumberFormat = "dd/mm/yyyy"

    ' Les colonnes Statut à Tiers ne sont jamais modifiables; seule la Catégorie l'est.
    ws.Range(COL_STATUT & LIGNE_PREMIERE_DONNEE & ":" & COL_TIERS & derniereLigne).Locked = True
    ws.Range(COL_CATEGORIE & LIGNE_PREMIERE_DONNEE & ":" & COL_CATEGORIE & derniereLigne).Locked = False

End Sub


' =====================================================================================
' Boutons
' =====================================================================================
Private Sub ConstruireZoneBoutons(ByVal ws As Worksheet)

    Dim gauche As Double

    gauche = ws.Cells(1, 2).Left

    ' Ajout d'un bouton de type "Contrôle de formulaire" (PAS ActiveX), comme dans les
    ' autres formulaires : plus léger et plus fiable face au problème de DPI.
    mod_InstallCommun.AjouterBoutonEntete ws, RES_LIGNE_BOUTONS, gauche, mod_Display.FR("R{e2}initialiser cette ligne"), _
        "ReinitialiserLigneSelectionnee", "btnReinitialiserLigne", 120
    mod_InstallCommun.AjouterBoutonEntete ws, RES_LIGNE_BOUTONS, gauche, mod_InstallCommun.CapEnregistrer(), _
        "TerminerEtAppliquerChoix", "btnTerminerResolution", mod_InstallCommun.FRM_BTN_L

End Sub


' =====================================================================================
' Figer les volets sous la ligne d'en-têtes du tableau
' =====================================================================================
Private Sub FigerVoletsSousEntetes(ByVal ws As Worksheet)

    ws.Activate

    ' Excel fige tout ce qui se trouve AU-DESSUS et À GAUCHE de la cellule active au
    ' moment de l'appel à FreezePanes : on sélectionne la première ligne de données.
    ws.Range(COL_STATUT & LIGNE_PREMIERE_DONNEE).Select

    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True

    ws.Range("A1").Select

End Sub


' =====================================================================================
' OUTILS DÉVELOPPEUR (Ctrl+G)
' =====================================================================================
Sub AfficherFeuilleResolutionPourEdition()
    mod_InstallCommun.AfficherPourEdition NOM_FEUILLE_RESOLUTION, "CreerFeuilleResolutionCategories", "MasquerFeuilleResolutionApresEdition"
End Sub

Sub MasquerFeuilleResolutionApresEdition()
    mod_InstallCommun.MasquerApresEdition NOM_FEUILLE_RESOLUTION
End Sub
