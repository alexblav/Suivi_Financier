Option Explicit

' =====================================================================================
' MODULE : mod_InstallSuiviSanteSheet
'
' RÔLE (phase 2a du chantier "Suivi Santé") :
'   Ce module contient UNIQUEMENT la construction de la mise en page statique de la
'   feuille masquée qui servira de formulaire opérateur pour valider les dépenses
'   de santé ("Frais, remb santé") nécessitant une intervention manuelle.
'   Il ne contient ENCORE AUCUNE logique de remplissage des cas ni de gestion des
'   clics : cela viendra en phase 2b (module mod_SuiviSanteFormulaire), une fois la
'   mise en page ci-dessous validée visuellement.
'
'   Ce module reprend exactement les mêmes principes que mod_InstallResolutionSheet
'   (déjà utilisé pour la résolution des catégories ambiguës) : construire les murs
'   avant de brancher l'électricité.
'
' À FAIRE POUR INSTALLER CE MODULE :
'   1. Alt+F11 pour ouvrir l'éditeur VBA.
'   2. Fichier > Importer un fichier... > sélectionner ce fichier .bas.
'   3. Dans la fenêtre Exécution immédiate (Ctrl+G), taper :
'        CreerFeuilleSuiviSante
'      puis appuyer sur Entrée. La feuille est créée, mise en forme, puis masquée.
'   4. Pour la revoir à l'écran : taper AfficherFeuilleSuiviSantePourEdition.
'      Pour la masquer de nouveau : taper MasquerFeuilleSuiviSanteApresEdition.
'
' À PROPOS DES ACCENTS DANS CE FICHIER :
'   Les textes affichés (titres, libellés et instructions) sont construits par
'   la fonction FR() ci-dessous, qui remplace les marqueurs ASCII ({e2}, {e1}, ...)
'   par les caractères accentués correspondants via ChrW(). Les commentaires du
'   fichier sont encodés en UTF-8.
' =====================================================================================


' -------------------------------------------------------------------------------------
' FR : petit traducteur de marqueurs ASCII vers caractères accentués (voir la note ci-dessus).
' -------------------------------------------------------------------------------------
' Marqueurs disponibles : {e2}=é, {e1}=è, {ea}=ê, {a2}=à, {c2}=ç, {o2}=ô,
' {i2}=î et {E2}=É.
Public Const NOM_FEUILLE_SUIVI_SANTE As String = "frm_SuiviSante"

' --- Zone des boutons (ligne 2) ---
Public Const SS_LIGNE_BOUTONS As Long = 2

' --- Zone des instructions ---
Public Const SS_LIGNE_TITRE_INSTRUCTIONS As Long = 5
Public Const SS_LIGNE_DEBUT_INSTRUCTIONS As Long = 6
Public Const SS_LIGNE_FIN_INSTRUCTIONS As Long = 9

' --- Bloc "Dépense à traiter" (lecture seule) ---
Public Const SS_LIGNE_TITRE_DEPENSE As Long = 11
Public Const SS_LIGNE_DEPENSE As Long = 12

' --- Bloc "Remboursements liés" (lecture seule, jusqu'à deux lignes) ---
Public Const SS_LIGNE_TITRE_REMBOURSEMENTS As Long = 14
Public Const SS_LIGNE_REMBOURSEMENT_1 As Long = 15
Public Const SS_LIGNE_REMBOURSEMENT_2 As Long = 16

' --- Solde calculé (lecture seule, recalculé en direct par une formule) ---
Public Const SS_LIGNE_SOLDE As Long = 18

' --- Zone de saisie opérateur ---
Public Const SS_LIGNE_BENEFICIAIRE As Long = 20
Public Const SS_LIGNE_TIERS_CORRIGE As Long = 21
Public Const SS_LIGNE_FRANCHISE As Long = 22
Public Const SS_LIGNE_DEPASSEMENT As Long = 23
Public Const SS_LIGNE_COMMENTAIRE As Long = 24

' --- Bloc "Informations complémentaires" (lecture seule), ajouté ultérieurement ---
' Démarre à la ligne 26 (deux lignes après Commentaire) pour ne pas décaler
' les lignes existantes : tous les noms et boutons déjà créés restent valides;
' le nouveau bloc est simplement ajouté à la suite.
Public Const SS_LIGNE_TITRE_CONTEXTE As Long = 26
Public Const SS_LIGNE_NUM_CHEQUE As Long = 27
Public Const SS_LIGNE_SPE_CONSULT As Long = 28
Public Const SS_LIGNE_NOTES As Long = 29

' --- Colonnes (une lettre = une colonne Excel). Le formulaire est organisé en
'     trois paires "libellé / valeur" par ligne (B/C, D/E, F/G) pour les blocs
'     affichant plusieurs informations sur une même ligne (ex. : Date, Tiers,
'     montant de la dépense). --------------------------------------------------
Public Const SS_COL_LIBELLE_1 As String = "B"
Public Const SS_COL_VALEUR_1 As String = "C"
Public Const SS_COL_LIBELLE_2 As String = "D"
Public Const SS_COL_VALEUR_2 As String = "E"
Public Const SS_COL_LIBELLE_3 As String = "F"
Public Const SS_COL_VALEUR_3 As String = "G"

' Colonne technique (masquée) où sont conservées les informations nécessaires au
' code de la phase 2b (numéro de ligne de la dépense en cours de traitement dans
' TblOperations). Elle n'est jamais visible par l'opérateur.
Public Const SS_COL_TECHNIQUE As String = "J"


' =====================================================================================
' MACRO PRINCIPALE D'INSTALLATION
' À exécuter UNE SEULE FOIS (ou de nouveau pour réinitialiser complètement la mise
' en forme de la feuille).
' =====================================================================================

Sub CreerFeuilleSuiviSante()

    Dim ws As Worksheet
    Dim reponseUtilisateur As VbMsgBoxResult

    Set ws = ObtenirFeuilleSansErreurSS(NOM_FEUILLE_SUIVI_SANTE)

    If Not ws Is Nothing Then
        reponseUtilisateur = MsgBox( _
            "La feuille '" & NOM_FEUILLE_SUIVI_SANTE & "' existe deja." & vbCrLf & _
            "Veux-tu la reconstruire entierement (sa mise en forme actuelle sera perdue) ?", _
            vbYesNo + vbQuestion, "Confirmation de reconstruction")

        If reponseUtilisateur = vbNo Then
            MsgBox "Installation annulee, aucune modification effectuee.", vbInformation
            Exit Sub
        End If

        ws.Visible = xlSheetVisible
        ws.Cells.Clear
        Call SupprimerFormesExistantesSS(ws)
        Call SupprimerNomsExistantsSS(ws)
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = NOM_FEUILLE_SUIVI_SANTE
    End If

    Call SS_AppliquerMiseEnFormeGenerale(ws)
    Call SS_ConstruireZoneBoutons(ws)
    Call SS_ConstruireZoneInstructions(ws)
    Call SS_ConstruireBlocDepense(ws)
    Call SS_ConstruireBlocRemboursements(ws)
    Call SS_ConstruireBlocSolde(ws)
    Call SS_ConstruireBlocSaisie(ws)
    Call SS_ConstruireZoneTechnique(ws)

    ws.Visible = xlSheetVeryHidden

    MsgBox "La feuille '" & NOM_FEUILLE_SUIVI_SANTE & "' a ete creee et mise en forme," & vbCrLf & _
           "puis masquee (xlSheetVeryHidden)." & vbCrLf & vbCrLf & _
           "Pour la revoir a l'ecran et verifier le rendu, execute la macro :" & vbCrLf & _
           "   AfficherFeuilleSuiviSantePourEdition", _
           vbInformation, "Installation terminee"

End Sub


' =====================================================================================
' SOUS-PROCÉDURE - Mise en forme générale
' =====================================================================================
Private Sub SS_AppliquerMiseEnFormeGenerale(ws As Worksheet)

    ws.Activate
    ActiveWindow.DisplayGridlines = False

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 20
    ws.Columns("C").ColumnWidth = 16
    ws.Columns("D").ColumnWidth = 16
    ws.Columns("E").ColumnWidth = 16
    ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 16
    ws.Columns("H").ColumnWidth = 2

    ' Colonne technique : largeur minimale et masquée, jamais visible par l'opérateur.
    ws.Columns(SS_COL_TECHNIQUE).ColumnWidth = 10
    ws.Columns(SS_COL_TECHNIQUE).Hidden = True

    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10

    ws.Range("A1").Select
    ActiveWindow.DisplayHeadings = False

End Sub


' =====================================================================================
' SOUS-PROCÉDURE - Zone des boutons (ligne 2) et compteur
' =====================================================================================
Private Sub SS_ConstruireZoneBoutons(ws As Worksheet)

    Dim boutonSuivant As Button
    Dim boutonValider As Button
    Dim zoneBoutonSuivant As Range
    Dim zoneBoutonValider As Range

    Set zoneBoutonSuivant = ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_BOUTONS & ":" & SS_COL_VALEUR_1 & SS_LIGNE_BOUTONS)
    Set zoneBoutonValider = ws.Range(SS_COL_LIBELLE_2 & SS_LIGNE_BOUTONS & ":" & SS_COL_VALEUR_2 & SS_LIGNE_BOUTONS)

    zoneBoutonSuivant.RowHeight = 22
    zoneBoutonValider.RowHeight = 22

    ' Boutons de type "Contrôle de formulaire" (pas ActiveX), comme dans le reste
    ' du classeur. Les macros référencées dans .OnAction n'existent pas encore : elles
    ' seront ajoutées en phase 2b. Cela ne génère aucune erreur tant que personne
    ' ne clique sur ces boutons.
    Set boutonSuivant = ws.Buttons.Add( _
        zoneBoutonSuivant.Left, zoneBoutonSuivant.Top, zoneBoutonSuivant.Width, zoneBoutonSuivant.Height)
    With boutonSuivant
        .Caption = mod_Display.FR("Cas suivant")
        .OnAction = "CasSuivantSuiviSante"
        .Name = "btnCasSuivantSuiviSante"
    End With

    Set boutonValider = ws.Buttons.Add( _
        zoneBoutonValider.Left, zoneBoutonValider.Top, zoneBoutonValider.Width, zoneBoutonValider.Height)
    With boutonValider
        .Caption = mod_Display.FR("Valider ce cas")
        .OnAction = "ValiderCasSuiviSante"
        .Name = "btnValiderCasSuiviSante"
    End With

    ' Compteur "X restant(s) sur Y", même principe que CompteurCasRestants, déjà
    ' utilisé pour la résolution des catégories : un nom défini plutôt qu'une
    ' référence de cellule codée en dur.
    With ws.Range(SS_COL_LIBELLE_3 & SS_LIGNE_BOUTONS & ":" & SS_COL_VALEUR_3 & SS_LIGNE_BOUTONS)
        .Merge
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .Font.Size = 9
        .Font.Color = RGB(120, 120, 120)
        .value = ""
    End With
    Call CreerNomSiAbsentSS(ws, "CompteurCasSante", ws.Range(SS_COL_LIBELLE_3 & SS_LIGNE_BOUTONS))

End Sub


' =====================================================================================
' SOUS-PROCÉDURE - Zone des instructions
' =====================================================================================
Private Sub SS_ConstruireZoneInstructions(ws As Worksheet)

    With ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_TITRE_INSTRUCTIONS)
        .value = mod_Display.FR("Instructions")
        .Font.Bold = True
        .Font.Size = 11
    End With

    Dim zoneTexte As Range
    Set zoneTexte = ws.Range( _
        SS_COL_LIBELLE_1 & SS_LIGNE_DEBUT_INSTRUCTIONS & ":" & _
        SS_COL_VALEUR_3 & SS_LIGNE_FIN_INSTRUCTIONS)

    With zoneTexte
        .Merge
        .value = mod_Display.FR("Cette d{e2}pense de sant{e2} n'est pas encore soldee ({e2}quilibre non atteint).") & Chr(10) & _
                 mod_Display.FR("V{e2}rifie les informations ci-dessous puis renseigne les champs demand{e2}s :") & Chr(10) & _
                 mod_Display.FR("1. Choisis le B{e2}n{e2}ficiaire dans la liste") & Chr(10) & _
                 mod_Display.FR("2. Si le Tiers importe n'est pas correct, choisis la bonne valeur dans 'Tiers corrig{e2}'") & Chr(10) & _
                 mod_Display.FR("3. Indique le montant de Franchise retenu par l'assurance (0 si aucun)") & Chr(10) & _
                 mod_Display.FR("4. Si les 2 remboursements sont arriv{e2}s, indique s'il s'agit d'un d{e2}passement d'honoraires") & Chr(10) & _
                 mod_Display.FR("5. Ajoute un commentaire si besoin, puis clique sur 'Valider ce cas'")
        .WrapText = True
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
        .Font.Size = 9
        .Font.Color = RGB(80, 80, 80)
        .Interior.Color = RGB(245, 245, 242)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(210, 210, 205)
        .Locked = True
    End With

    ws.rows(SS_LIGNE_DEBUT_INSTRUCTIONS & ":" & SS_LIGNE_FIN_INSTRUCTIONS).RowHeight = 15

End Sub


' =====================================================================================
' SOUS-PROCÉDURE - Bloc "Dépense à traiter" (lecture seule)
' =====================================================================================
Private Sub SS_ConstruireBlocDepense(ws As Worksheet)

    With ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_TITRE_DEPENSE & ":" & SS_COL_VALEUR_3 & SS_LIGNE_TITRE_DEPENSE)
        .Merge
        .value = mod_Display.FR("D{e2}pense {a2} traiter")
        .Font.Bold = True
        .Font.Size = 10.5
        .Interior.Color = RGB(250, 250, 248)
    End With

    ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_DEPENSE).value = mod_Display.FR("Date :")
    ws.Range(SS_COL_LIBELLE_2 & SS_LIGNE_DEPENSE).value = mod_Display.FR("Tiers importe :")
    ws.Range(SS_COL_LIBELLE_3 & SS_LIGNE_DEPENSE).value = mod_Display.FR("Montant :")

    Call SS_MettreEnFormeLibelles(ws, SS_LIGNE_DEPENSE)

    ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_DEPENSE).NumberFormat = "dd/mm/yyyy"
    ws.Range(SS_COL_VALEUR_3 & SS_LIGNE_DEPENSE).NumberFormat = "#,##0.00 " & ChrW(8364)

    Call SS_MettreEnFormeValeursLectureSeule(ws, SS_LIGNE_DEPENSE)

    Call CreerNomSiAbsentSS(ws, "ssDate", ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_DEPENSE))
    Call CreerNomSiAbsentSS(ws, "ssTiersImporte", ws.Range(SS_COL_VALEUR_2 & SS_LIGNE_DEPENSE))
    Call CreerNomSiAbsentSS(ws, "ssMontant", ws.Range(SS_COL_VALEUR_3 & SS_LIGNE_DEPENSE))

End Sub


' =====================================================================================
' SOUS-PROCÉDURE - Bloc "Remboursements liés" (lecture seule, jusqu'à deux lignes)
' =====================================================================================
Private Sub SS_ConstruireBlocRemboursements(ws As Worksheet)

    With ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_TITRE_REMBOURSEMENTS & ":" & SS_COL_VALEUR_3 & SS_LIGNE_TITRE_REMBOURSEMENTS)
        .Merge
        .value = mod_Display.FR("Remboursements li{e2}s (jusqu'{a2} 2)")
        .Font.Bold = True
        .Font.Size = 10.5
        .Interior.Color = RGB(250, 250, 248)
    End With

    ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_REMBOURSEMENT_1).value = mod_Display.FR("Remb. 1 - Date :")
    ws.Range(SS_COL_LIBELLE_2 & SS_LIGNE_REMBOURSEMENT_1).value = mod_Display.FR("Montant :")

    ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_REMBOURSEMENT_2).value = mod_Display.FR("Remb. 2 - Date :")
    ws.Range(SS_COL_LIBELLE_2 & SS_LIGNE_REMBOURSEMENT_2).value = mod_Display.FR("Montant :")

    Call SS_MettreEnFormeLibelles(ws, SS_LIGNE_REMBOURSEMENT_1)
    Call SS_MettreEnFormeLibelles(ws, SS_LIGNE_REMBOURSEMENT_2)

    ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_REMBOURSEMENT_1).NumberFormat = "dd/mm/yyyy"
    ws.Range(SS_COL_VALEUR_2 & SS_LIGNE_REMBOURSEMENT_1).NumberFormat = "#,##0.00 " & ChrW(8364)
    ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_REMBOURSEMENT_2).NumberFormat = "dd/mm/yyyy"
    ws.Range(SS_COL_VALEUR_2 & SS_LIGNE_REMBOURSEMENT_2).NumberFormat = "#,##0.00 " & ChrW(8364)

    Call SS_MettreEnFormeValeursLectureSeule(ws, SS_LIGNE_REMBOURSEMENT_1)
    Call SS_MettreEnFormeValeursLectureSeule(ws, SS_LIGNE_REMBOURSEMENT_2)

    Call CreerNomSiAbsentSS(ws, "ssRemb1Date", ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_REMBOURSEMENT_1))
    Call CreerNomSiAbsentSS(ws, "ssRemb1Montant", ws.Range(SS_COL_VALEUR_2 & SS_LIGNE_REMBOURSEMENT_1))
    Call CreerNomSiAbsentSS(ws, "ssRemb2Date", ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_REMBOURSEMENT_2))
    Call CreerNomSiAbsentSS(ws, "ssRemb2Montant", ws.Range(SS_COL_VALEUR_2 & SS_LIGNE_REMBOURSEMENT_2))

End Sub


' =====================================================================================
' SOUS-PROCÉDURE - Bloc "Solde actuel" (lecture seule, formule dynamique)
' =====================================================================================
' Le solde est calculé PAR UNE FORMULE EXCEL CLASSIQUE (pas par du VBA) : il est
' recalculé automatiquement et instantanément dès que l'opérateur modifie la
' Franchise, sans code supplémentaire. La formule exacte sera écrite en phase 2b
' (elle dépend des cellules ssMontant/ssRemb1Montant/ssRemb2Montant/ssFranchise; cette
' dernière n'existe pas encore en phase 2a). Pour l'instant, on prépare seulement
' la mise en forme de la cellule qui recevra cette formule.
Private Sub SS_ConstruireBlocSolde(ws As Worksheet)

    ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_SOLDE).value = mod_Display.FR("Solde actuel :")
    Call SS_MettreEnFormeLibelles(ws, SS_LIGNE_SOLDE)

    With ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_SOLDE & ":" & SS_COL_VALEUR_2 & SS_LIGNE_SOLDE)
        .Merge
        .NumberFormat = "#,##0.00 " & ChrW(8364)
        .Font.Bold = True
    End With
    Call SS_MettreEnFormeValeursLectureSeule(ws, SS_LIGNE_SOLDE)

    Call CreerNomSiAbsentSS(ws, "ssSolde", ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_SOLDE))

End Sub


' =====================================================================================
' SOUS-PROCÉDURE - Bloc de saisie opérateur (Bénéficiaire, Tiers corrigé, Franchise,
' Dépassement, Commentaire)
' =====================================================================================
Private Sub SS_ConstruireBlocSaisie(ws As Worksheet)

    Dim rngBeneficiaire As Range
    Dim rngTiersCorrige As Range
    Dim rngFranchise As Range
    Dim rngDepassement As Range
    Dim rngCommentaire As Range

    ' --- Bénéficiaire (toujours demandé : la colonne est vide sur toutes les
    '     lignes existantes; elle n'est jamais renseignée par l'import OFX) ---
    ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_BENEFICIAIRE).value = mod_Display.FR("B{e2}n{e2}ficiaire :")
    Call SS_MettreEnFormeLibelles(ws, SS_LIGNE_BENEFICIAIRE)
    Set rngBeneficiaire = ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_BENEFICIAIRE & ":" & SS_COL_VALEUR_2 & SS_LIGNE_BENEFICIAIRE)
    rngBeneficiaire.Merge
    Call SS_MettreEnFormeZoneSaisie(rngBeneficiaire)
    Call SS_AppliquerListeDeroulante(rngBeneficiaire, "Beneficiaires")
    Call CreerNomSiAbsentSS(ws, "ssBeneficiaire", ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_BENEFICIAIRE))

    ' --- Tiers corrigé (demandé uniquement si le Tiers importé ne figure pas déjà
    '     dans la liste Praticiens; ce contrôle sera fait en phase 2b. Ici, on prépare
    '     uniquement la cellule et sa liste déroulante) ---
    ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_TIERS_CORRIGE).value = mod_Display.FR("Tiers corrig{e2} :")
    Call SS_MettreEnFormeLibelles(ws, SS_LIGNE_TIERS_CORRIGE)
    Set rngTiersCorrige = ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_TIERS_CORRIGE & ":" & SS_COL_VALEUR_2 & SS_LIGNE_TIERS_CORRIGE)
    rngTiersCorrige.Merge
    Call SS_MettreEnFormeZoneSaisie(rngTiersCorrige)
    Call SS_AppliquerListeDeroulante(rngTiersCorrige, "Praticiens")
    Call CreerNomSiAbsentSS(ws, "ssTiersCorrige", ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_TIERS_CORRIGE))

    ' --- Franchise (saisie numérique libre, valeur par défaut : 0) ---
    ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_FRANCHISE).value = mod_Display.FR("Franchise :")
    Call SS_MettreEnFormeLibelles(ws, SS_LIGNE_FRANCHISE)
    Set rngFranchise = ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_FRANCHISE)
    Call SS_MettreEnFormeZoneSaisie(rngFranchise)
    rngFranchise.NumberFormat = "#,##0.00 " & ChrW(8364)
    Call CreerNomSiAbsentSS(ws, "ssFranchise", rngFranchise)

    ' --- Dépassement d'honoraires (liste Oui/Non, pertinent seulement si les deux
    '     remboursements sont présents; la phase 2b grisera ou masquera ce champ
    '     le cas échéant) ---
    ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_DEPASSEMENT).value = mod_Display.FR("D{e2}passement d'honoraires ? :")
    Call SS_MettreEnFormeLibelles(ws, SS_LIGNE_DEPASSEMENT)
    Set rngDepassement = ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_DEPASSEMENT)
    Call SS_MettreEnFormeZoneSaisie(rngDepassement)
    Call SS_AppliquerListeDeroulanteTexte(rngDepassement, mod_Display.FR("Oui,Non"))
    Call CreerNomSiAbsentSS(ws, "ssDepassement", rngDepassement)

    ' --- Commentaire (texte libre) ---
    ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_COMMENTAIRE).value = mod_Display.FR("Commentaire :")
    Call SS_MettreEnFormeLibelles(ws, SS_LIGNE_COMMENTAIRE)
    Set rngCommentaire = ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_COMMENTAIRE & ":" & SS_COL_VALEUR_3 & SS_LIGNE_COMMENTAIRE)
    rngCommentaire.Merge
    rngCommentaire.RowHeight = 32
    With rngCommentaire
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    Call SS_MettreEnFormeZoneSaisie(rngCommentaire)
    Call CreerNomSiAbsentSS(ws, "ssCommentaire", ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_COMMENTAIRE))

End Sub


' =====================================================================================
' SOUS-PROCÉDURE - Zone technique masquée (mémorise la ligne TblOperations en cours)
' =====================================================================================
Private Sub SS_ConstruireZoneTechnique(ws As Worksheet)
    ws.Range(SS_COL_TECHNIQUE & SS_LIGNE_BOUTONS).value = 0
    Call CreerNomSiAbsentSS(ws, "ssLigneEnCours", ws.Range(SS_COL_TECHNIQUE & SS_LIGNE_BOUTONS))
End Sub


' =====================================================================================
' SOUS-PROCÉDURES UTILITAIRES DE MISE EN FORME (évitent de répéter le même code de
' style pour chaque champ)
' =====================================================================================
Private Sub SS_MettreEnFormeLibelles(ws As Worksheet, ByVal ligne As Long)
    ' On formate chaque cellule de libellé individuellement (B, D, F) plutôt que
    ' toute la plage B:F, car certaines lignes n'utilisent que la colonne B
    ' comme libellé (les colonnes D/F servent alors à une valeur fusionnée, par
    ' exemple pour le champ Bénéficiaire).
    ws.Range(SS_COL_LIBELLE_1 & ligne).Font.Bold = False
    ws.Range(SS_COL_LIBELLE_1 & ligne).Font.Color = RGB(90, 90, 90)
    ws.Range(SS_COL_LIBELLE_2 & ligne).Font.Color = RGB(90, 90, 90)
    ws.Range(SS_COL_LIBELLE_3 & ligne).Font.Color = RGB(90, 90, 90)
    ws.rows(ligne).RowHeight = 18
End Sub

Private Sub SS_MettreEnFormeValeursLectureSeule(ws As Worksheet, ByVal ligne As Long)
    With ws.Range(SS_COL_VALEUR_1 & ligne & ":" & SS_COL_VALEUR_3 & ligne)
        .Locked = True
        .Interior.Color = RGB(248, 248, 246)
    End With
End Sub

Private Sub SS_MettreEnFormeZoneSaisie(ByVal rng As Range)
    With rng
        .Locked = False
        .Interior.Color = RGB(255, 255, 235)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 190, 140)
    End With
End Sub

Private Sub SS_AppliquerListeDeroulante(ByVal rng As Range, ByVal nomPlage As String)
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Formula1:="=" & nomPlage
    On Error GoTo 0
End Sub

Private Sub SS_AppliquerListeDeroulanteTexte(ByVal rng As Range, ByVal listeVirgules As String)
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Formula1:=listeVirgules
    On Error GoTo 0
End Sub


' =====================================================================================
' FONCTIONS UTILITAIRES
' =====================================================================================
Private Function ObtenirFeuilleSansErreurSS(nomFeuille As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nomFeuille)
    On Error GoTo 0
    Set ObtenirFeuilleSansErreurSS = ws
End Function

Private Sub SupprimerFormesExistantesSS(ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
End Sub

' Supprime les noms définis qui pointent vers cette feuille avant sa reconstruction,
' afin d'éviter les erreurs "nom déjà utilisé" si l'installation est relancée.
Private Sub SupprimerNomsExistantsSS(ws As Worksheet)
    Dim n As Name
    Dim i As Long
    For i = ThisWorkbook.Names.count To 1 Step -1
        Set n = ThisWorkbook.Names(i)
        On Error Resume Next
        If InStr(1, n.RefersTo, "'" & NOM_FEUILLE_SUIVI_SANTE & "'", vbTextCompare) > 0 Then
            n.Delete
        End If
        On Error GoTo 0
    Next i
End Sub

' Crée un nom défini uniquement s'il n'existe pas déjà (évite une erreur si la
' macro d'installation est relancée plusieurs fois sans appeler
' SupprimerNomsExistantsSS, par exemple lors de tests manuels).
Private Sub CreerNomSiAbsentSS(ws As Worksheet, ByVal nomCellule As String, ByVal rng As Range)
    On Error Resume Next
    ThisWorkbook.Names(nomCellule).Delete
    On Error GoTo 0
    ws.Names.Add Name:=nomCellule, RefersTo:=rng
End Sub


' =====================================================================================
' OUTILS DÉVELOPPEUR (réservés à toi, jamais accessibles dans l'utilisation normale
' du classeur) : basculent la visibilité de la feuille pour permettre de la retoucher.
' =====================================================================================
' =====================================================================================
' AjouterBoutonsAjoutListe : ajoute (ou remplace s'ils sont déjà présents) UNIQUEMENT
' les deux boutons "+" à côté des champs Bénéficiaire et Tiers corrigé, sans reconstruire
' le reste de la feuille (contrairement à CreerFeuilleSuiviSante, qui repartirait de zéro
' et supprimerait tes éventuels réglages manuels).
' À exécuter UNE SEULE FOIS dans la fenêtre Exécution immédiate (Ctrl+G) :
'      AjouterBoutonsAjoutListe
' =====================================================================================
Sub AjouterBoutonsAjoutListe()

    Dim ws As Worksheet
    Dim etaitMasquee As Boolean
    Dim zoneBtn As Range
    Dim btn As Button

    Set ws = ObtenirFeuilleSansErreurSS(NOM_FEUILLE_SUIVI_SANTE)
    If ws Is Nothing Then
        MsgBox "La feuille '" & NOM_FEUILLE_SUIVI_SANTE & "' n'existe pas encore." & vbCrLf & _
               "Execute d'abord CreerFeuilleSuiviSante.", vbExclamation
        Exit Sub
    End If

    ' On rend temporairement la feuille visible pour placer les boutons
    ' (Shapes.Add échoue parfois sur une feuille très masquée), puis on la
    ' masque de nouveau si elle l'était au départ.
    etaitMasquee = (ws.Visible <> xlSheetVisible)
    ws.Visible = xlSheetVisible

    SupprimerBoutonSiExisteSS ws, "btnAjouterBeneficiaire"
    SupprimerBoutonSiExisteSS ws, "btnAjouterTiers"

    Set zoneBtn = ws.Range(SS_COL_LIBELLE_3 & SS_LIGNE_BENEFICIAIRE)
    Set btn = ws.Buttons.Add(zoneBtn.Left, zoneBtn.Top, 26, zoneBtn.Height)
    With btn
        .Caption = "+"
        .OnAction = "AjouterBeneficiaire"
        .Name = "btnAjouterBeneficiaire"
    End With

    Set zoneBtn = ws.Range(SS_COL_LIBELLE_3 & SS_LIGNE_TIERS_CORRIGE)
    Set btn = ws.Buttons.Add(zoneBtn.Left, zoneBtn.Top, 26, zoneBtn.Height)
    With btn
        .Caption = "+"
        .OnAction = "AjouterTiersPraticien"
        .Name = "btnAjouterTiers"
    End With

    If etaitMasquee Then ws.Visible = xlSheetVeryHidden

    MsgBox "Boutons ajoutes avec succes.", vbInformation

End Sub

Private Sub SupprimerBoutonSiExisteSS(ws As Worksheet, ByVal nomBouton As String)
    On Error Resume Next
    ws.Buttons(nomBouton).Delete
    On Error GoTo 0
End Sub


' =====================================================================================
' AjouterChampsContexteSuiviSante : ajoute le bloc "Informations complémentaires"
' (Num_Cheque, Date_consult, Spe_Consult, Notes brutes, toutes en lecture seule)
' SANS reconstruire le reste de la feuille. À exécuter UNE SEULE FOIS dans la
' fenêtre Exécution immédiate (Ctrl+G) :
'      AjouterChampsContexteSuiviSante
' =====================================================================================
Sub AjouterChampsContexteSuiviSante()

    Dim ws As Worksheet
    Dim etaitMasquee As Boolean

    Set ws = ObtenirFeuilleSansErreurSS(NOM_FEUILLE_SUIVI_SANTE)
    If ws Is Nothing Then
        MsgBox "La feuille '" & NOM_FEUILLE_SUIVI_SANTE & "' n'existe pas encore." & vbCrLf & _
               "Execute d'abord CreerFeuilleSuiviSante.", vbExclamation
        Exit Sub
    End If

    etaitMasquee = (ws.Visible <> xlSheetVisible)
    ws.Visible = xlSheetVisible

    ' --- Titre du bloc ---
    With ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_TITRE_CONTEXTE & ":" & SS_COL_VALEUR_3 & SS_LIGNE_TITRE_CONTEXTE)
        .Merge
        .value = mod_Display.FR("Informations compl{e2}mentaires")
        .Font.Bold = True
        .Font.Size = 10.5
        .Interior.Color = RGB(250, 250, 248)
    End With

    ' --- Ligne Num_Cheque + Date_consult (deux paires sur la même ligne) ---
    ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_NUM_CHEQUE).value = "Num_Cheque :"
    ws.Range(SS_COL_LIBELLE_2 & SS_LIGNE_NUM_CHEQUE).value = mod_Display.FR("Date_consult :")
    Call SS_MettreEnFormeLibelles(ws, SS_LIGNE_NUM_CHEQUE)
    ws.Range(SS_COL_VALEUR_2 & SS_LIGNE_NUM_CHEQUE).NumberFormat = "dd/mm/yyyy"
    Call SS_MettreEnFormeValeursLectureSeule(ws, SS_LIGNE_NUM_CHEQUE)
    Call CreerNomSiAbsentSS(ws, "ssNumCheque", ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_NUM_CHEQUE))
    Call CreerNomSiAbsentSS(ws, "ssDateConsult", ws.Range(SS_COL_VALEUR_2 & SS_LIGNE_NUM_CHEQUE))

    ' --- Ligne Spe_Consult ---
    ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_SPE_CONSULT).value = mod_Display.FR("Sp{e2}cialit{e2} consult{e2}e :")
    Call SS_MettreEnFormeLibelles(ws, SS_LIGNE_SPE_CONSULT)
    With ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_SPE_CONSULT & ":" & SS_COL_VALEUR_2 & SS_LIGNE_SPE_CONSULT)
        .Merge
    End With
    Call SS_MettreEnFormeValeursLectureSeule(ws, SS_LIGNE_SPE_CONSULT)
    Call CreerNomSiAbsentSS(ws, "ssSpeConsult", ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_SPE_CONSULT))

    ' --- Ligne Notes (texte brut, potentiellement long -> renvoi à la ligne) ---
    ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_NOTES).value = "Notes :"
    Call SS_MettreEnFormeLibelles(ws, SS_LIGNE_NOTES)
    With ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_NOTES & ":" & SS_COL_VALEUR_3 & SS_LIGNE_NOTES)
        .Merge
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    ws.rows(SS_LIGNE_NOTES).RowHeight = 28
    Call SS_MettreEnFormeValeursLectureSeule(ws, SS_LIGNE_NOTES)
    Call CreerNomSiAbsentSS(ws, "ssNotes", ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_NOTES))

    If etaitMasquee Then ws.Visible = xlSheetVeryHidden

    MsgBox "Bloc 'Informations complementaires' ajoute avec succes.", vbInformation

End Sub


Sub AfficherFeuilleSuiviSantePourEdition()
    Dim ws As Worksheet
    Set ws = ObtenirFeuilleSansErreurSS(NOM_FEUILLE_SUIVI_SANTE)

    If ws Is Nothing Then
        MsgBox "La feuille '" & NOM_FEUILLE_SUIVI_SANTE & "' n'existe pas encore." & vbCrLf & _
               "Execute d'abord la macro CreerFeuilleSuiviSante.", vbExclamation
        Exit Sub
    End If

    ws.Visible = xlSheetVisible
    ws.Activate
    MsgBox "La feuille est maintenant visible et modifiable." & vbCrLf & _
           "Pense a la remasquer avec MasquerFeuilleSuiviSanteApresEdition" & vbCrLf & _
           "une fois tes retouches terminees.", vbInformation
End Sub

Sub MasquerFeuilleSuiviSanteApresEdition()
    Dim ws As Worksheet
    Set ws = ObtenirFeuilleSansErreurSS(NOM_FEUILLE_SUIVI_SANTE)

    If ws Is Nothing Then
        MsgBox "La feuille '" & NOM_FEUILLE_SUIVI_SANTE & "' n'existe pas.", vbExclamation
        Exit Sub
    End If

    ws.Visible = xlSheetVeryHidden
    MsgBox "La feuille est de nouveau masquee (xlSheetVeryHidden).", vbInformation
End Sub
