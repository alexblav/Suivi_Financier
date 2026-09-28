Attribute VB_Name = "mod_InstallSuiviSanteSheet"
Option Explicit

' =====================================================================================
' MODULE : mod_InstallSuiviSanteSheet
'
' ROLE (PHASE 2a du chantier "Suivi Sante") :
'   Ce module ne contient QUE la construction de la mise en page statique de la
'   feuille masquee qui servira de formulaire operateur pour valider les depenses
'   de sante ("Frais, remb sante") qui necessitent une intervention manuelle.
'   Il ne contient ENCORE AUCUNE logique de remplissage des cas, ni de gestion des
'   clics : cela viendra en Phase 2b (module mod_SuiviSanteFormulaire), une fois que
'   la mise en page ci-dessous aura ete validee visuellement.
'
'   Ce module reprend exactement les memes principes que mod_InstallResolutionSheet
'   (deja utilise pour la resolution des categories ambigues) : construire les murs
'   avant de brancher l'electricite.
'
' A FAIRE POUR INSTALLER CE MODULE :
'   1. Alt+F11 pour ouvrir l'editeur VBA
'   2. Fichier > Importer un fichier... > selectionner ce fichier .bas
'   3. Dans la fenetre Execution immediate (Ctrl+G), taper :
'        CreerFeuilleSuiviSante
'      puis appuyer sur Entree. La feuille est creee, mise en forme, puis masquee.
'   4. Pour la revoir a l'ecran : taper AfficherFeuilleSuiviSantePourEdition
'      Pour la remasquer ensuite : taper MasquerFeuilleSuiviSanteApresEdition
'
' A PROPOS DES ACCENTS DANS CE FICHIER :
'   Comme pour les autres modules de ce classeur, ce fichier ne contient VOLONTAIREMENT
'   aucun caractere accentue (probleme d'encodage deja rencontre et corrige sur ce
'   projet). Tous les textes affiches a l'ecran (titres, libelles, instructions) sont
'   construits par la petite fonction FR() ci-dessous, qui remplace des marqueurs
'   ASCII simples ({e2}, {e1}, ...) par le bon caractere accentue via ChrW(). Cela
'   permet de garder un code source lisible tout en restant 100% ASCII a l'import.
' =====================================================================================


' -------------------------------------------------------------------------------------
' FR : petit traducteur de marqueurs ASCII vers caracteres accentues (voir note ci-dessus)
' -------------------------------------------------------------------------------------
' Marqueurs disponibles : {e2}=e accent aigu (e), {e1}=e accent grave (e), {ea}=e accent
' circonflexe (e), {a2}=a accent grave (a), {c2}=c cedille (c), {o2}=o accent circonflexe,
' {i2}=i accent circonflexe, {E2}=E accent aigu majuscule.
Public Const NOM_FEUILLE_SUIVI_SANTE As String = "frm_SuiviSante"

' --- Zone des boutons (ligne 2) ---
Public Const SS_LIGNE_BOUTONS As Long = 2

' --- Zone des instructions ---
Public Const SS_LIGNE_TITRE_INSTRUCTIONS As Long = 5
Public Const SS_LIGNE_DEBUT_INSTRUCTIONS As Long = 6
Public Const SS_LIGNE_FIN_INSTRUCTIONS As Long = 9

' --- Bloc "Depense a traiter" (lecture seule) ---
Public Const SS_LIGNE_TITRE_DEPENSE As Long = 11
Public Const SS_LIGNE_DEPENSE As Long = 12

' --- Bloc "Remboursements lies" (lecture seule, jusqu'a 2 lignes) ---
Public Const SS_LIGNE_TITRE_REMBOURSEMENTS As Long = 14
Public Const SS_LIGNE_REMBOURSEMENT_1 As Long = 15
Public Const SS_LIGNE_REMBOURSEMENT_2 As Long = 16

' --- Solde calcule (lecture seule, recalcule en direct via une formule) ---
Public Const SS_LIGNE_SOLDE As Long = 18

' --- Zone de saisie operateur ---
Public Const SS_LIGNE_BENEFICIAIRE As Long = 20
Public Const SS_LIGNE_TIERS_CORRIGE As Long = 21
Public Const SS_LIGNE_FRANCHISE As Long = 22
Public Const SS_LIGNE_DEPASSEMENT As Long = 23
Public Const SS_LIGNE_COMMENTAIRE As Long = 24

' --- Bloc "Informations complementaires" (lecture seule), ajoute apres coup ---
' Demarre a la ligne 26 (2 lignes apres Commentaire) pour ne jamais avoir a
' decaler les lignes existantes ci-dessus : tous les noms/boutons deja crees
' restent valides, on ne fait qu'ajouter a la suite.
Public Const SS_LIGNE_TITRE_CONTEXTE As Long = 26
Public Const SS_LIGNE_NUM_CHEQUE As Long = 27
Public Const SS_LIGNE_SPE_CONSULT As Long = 28
Public Const SS_LIGNE_NOTES As Long = 29

' --- Colonnes (une lettre = une colonne Excel). Le formulaire est organise en
'     3 paires "libelle / valeur" par ligne (B/C, D/E, F/G) pour les blocs qui
'     affichent plusieurs informations sur une meme ligne (ex : Date, Tiers,
'     Montant de la depense). ---
Public Const SS_COL_LIBELLE_1 As String = "B"
Public Const SS_COL_VALEUR_1 As String = "C"
Public Const SS_COL_LIBELLE_2 As String = "D"
Public Const SS_COL_VALEUR_2 As String = "E"
Public Const SS_COL_LIBELLE_3 As String = "F"
Public Const SS_COL_VALEUR_3 As String = "G"

' Colonne technique (masquee) ou l'on conserve les informations internes dont le
' code de la Phase 2b aura besoin (numero de ligne de la depense en cours de
' traitement dans TblOperations). Jamais visible pour l'operateur.
Public Const SS_COL_TECHNIQUE As String = "J"


' =====================================================================================
' MACRO PRINCIPALE D'INSTALLATION
' A executer UNE SEULE FOIS (ou a nouveau pour reinitialiser completement la mise
' en forme de la feuille).
' =====================================================================================


' -------------------------------------------------------------------------------------
' CONSTANTES DE MISE EN PAGE
' -------------------------------------------------------------------------------------
' Comme pour frm_ResolutionCategories, on regroupe ici toutes les positions de
' cellules : si tu veux deplacer une zone plus tard, on change une seule ligne ici.
' Ces constantes seront reutilisees telles quelles en Phase 2b.

Private Function FR(ByVal texte As String) As String
    Dim r As String
    r = texte
    r = Replace(r, "{e2}", ChrW(233))
    r = Replace(r, "{e1}", ChrW(232))
    r = Replace(r, "{ea}", ChrW(234))
    r = Replace(r, "{a2}", ChrW(224))
    r = Replace(r, "{c2}", ChrW(231))
    r = Replace(r, "{o2}", ChrW(244))
    r = Replace(r, "{i2}", ChrW(238))
    r = Replace(r, "{E2}", ChrW(201))
    FR = r
End Function
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
' SOUS-PROCEDURE - Mise en forme generale
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

    ' Colonne technique : reduite au minimum et masquee, jamais vue par l'operateur.
    ws.Columns(SS_COL_TECHNIQUE).ColumnWidth = 10
    ws.Columns(SS_COL_TECHNIQUE).Hidden = True

    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10

    ws.Range("A1").Select
    ActiveWindow.DisplayHeadings = False

End Sub


' =====================================================================================
' SOUS-PROCEDURE - Zone des boutons (ligne 2) + compteur
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

    ' Boutons de type "Controle de formulaire" (pas ActiveX), comme le reste du
    ' classeur. Les macros referencees dans .OnAction n'existent pas encore : ce
    ' sera le contenu de la Phase 2b. Cela ne genere aucune erreur tant que
    ' personne ne clique sur les boutons.
    Set boutonSuivant = ws.Buttons.Add( _
        zoneBoutonSuivant.Left, zoneBoutonSuivant.Top, zoneBoutonSuivant.Width, zoneBoutonSuivant.Height)
    With boutonSuivant
        .Caption = FR("Cas suivant")
        .OnAction = "CasSuivantSuiviSante"
        .Name = "btnCasSuivantSuiviSante"
    End With

    Set boutonValider = ws.Buttons.Add( _
        zoneBoutonValider.Left, zoneBoutonValider.Top, zoneBoutonValider.Width, zoneBoutonValider.Height)
    With boutonValider
        .Caption = FR("Valider ce cas")
        .OnAction = "ValiderCasSuiviSante"
        .Name = "btnValiderCasSuiviSante"
    End With

    ' Compteur "X restant(s) sur Y", meme principe que CompteurCasRestants deja
    ' utilise pour la resolution des categories : un nom defini plutot qu'une
    ' reference de cellule en dur.
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
' SOUS-PROCEDURE - Zone des instructions
' =====================================================================================
Private Sub SS_ConstruireZoneInstructions(ws As Worksheet)

    With ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_TITRE_INSTRUCTIONS)
        .value = FR("Instructions")
        .Font.Bold = True
        .Font.Size = 11
    End With

    Dim zoneTexte As Range
    Set zoneTexte = ws.Range( _
        SS_COL_LIBELLE_1 & SS_LIGNE_DEBUT_INSTRUCTIONS & ":" & _
        SS_COL_VALEUR_3 & SS_LIGNE_FIN_INSTRUCTIONS)

    With zoneTexte
        .Merge
        .value = FR("Cette d{e2}pense de sant{e2} n'est pas encore soldee ({e2}quilibre non atteint).") & Chr(10) & _
                 FR("V{e2}rifie les informations ci-dessous puis renseigne les champs demand{e2}s :") & Chr(10) & _
                 FR("1. Choisis le B{e2}n{e2}ficiaire dans la liste") & Chr(10) & _
                 FR("2. Si le Tiers importe n'est pas correct, choisis la bonne valeur dans 'Tiers corrig{e2}'") & Chr(10) & _
                 FR("3. Indique le montant de Franchise retenu par l'assurance (0 si aucun)") & Chr(10) & _
                 FR("4. Si les 2 remboursements sont arriv{e2}s, indique s'il s'agit d'un d{e2}passement d'honoraires") & Chr(10) & _
                 FR("5. Ajoute un commentaire si besoin, puis clique sur 'Valider ce cas'")
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
' SOUS-PROCEDURE - Bloc "Depense a traiter" (lecture seule)
' =====================================================================================
Private Sub SS_ConstruireBlocDepense(ws As Worksheet)

    With ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_TITRE_DEPENSE & ":" & SS_COL_VALEUR_3 & SS_LIGNE_TITRE_DEPENSE)
        .Merge
        .value = FR("D{e2}pense {a2} traiter")
        .Font.Bold = True
        .Font.Size = 10.5
        .Interior.Color = RGB(250, 250, 248)
    End With

    ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_DEPENSE).value = FR("Date :")
    ws.Range(SS_COL_LIBELLE_2 & SS_LIGNE_DEPENSE).value = FR("Tiers importe :")
    ws.Range(SS_COL_LIBELLE_3 & SS_LIGNE_DEPENSE).value = FR("Montant :")

    Call SS_MettreEnFormeLibelles(ws, SS_LIGNE_DEPENSE)

    ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_DEPENSE).NumberFormat = "dd/mm/yyyy"
    ws.Range(SS_COL_VALEUR_3 & SS_LIGNE_DEPENSE).NumberFormat = "#,##0.00 " & ChrW(8364)

    Call SS_MettreEnFormeValeursLectureSeule(ws, SS_LIGNE_DEPENSE)

    Call CreerNomSiAbsentSS(ws, "ssDate", ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_DEPENSE))
    Call CreerNomSiAbsentSS(ws, "ssTiersImporte", ws.Range(SS_COL_VALEUR_2 & SS_LIGNE_DEPENSE))
    Call CreerNomSiAbsentSS(ws, "ssMontant", ws.Range(SS_COL_VALEUR_3 & SS_LIGNE_DEPENSE))

End Sub


' =====================================================================================
' SOUS-PROCEDURE - Bloc "Remboursements lies" (lecture seule, jusqu'a 2 lignes)
' =====================================================================================
Private Sub SS_ConstruireBlocRemboursements(ws As Worksheet)

    With ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_TITRE_REMBOURSEMENTS & ":" & SS_COL_VALEUR_3 & SS_LIGNE_TITRE_REMBOURSEMENTS)
        .Merge
        .value = FR("Remboursements li{e2}s (jusqu'{a2} 2)")
        .Font.Bold = True
        .Font.Size = 10.5
        .Interior.Color = RGB(250, 250, 248)
    End With

    ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_REMBOURSEMENT_1).value = FR("Remb. 1 - Date :")
    ws.Range(SS_COL_LIBELLE_2 & SS_LIGNE_REMBOURSEMENT_1).value = FR("Montant :")

    ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_REMBOURSEMENT_2).value = FR("Remb. 2 - Date :")
    ws.Range(SS_COL_LIBELLE_2 & SS_LIGNE_REMBOURSEMENT_2).value = FR("Montant :")

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
' SOUS-PROCEDURE - Bloc "Solde actuel" (lecture seule, formule live)
' =====================================================================================
' Le solde est calcule PAR UNE FORMULE EXCEL CLASSIQUE (pas par du VBA) : il se
' recalcule donc automatiquement et instantanement des que l'operateur modifie la
' Franchise, sans qu'aucun code ne soit necessaire pour ca. La formule exacte sera
' ecrite en Phase 2b (elle depend des cellules ssMontant/ssRemb1Montant/
' ssRemb2Montant/ssFranchise, qui n'existent pas encore a ce stade de la Phase 2a
' pour ssFranchise). Pour l'instant on ne fait que preparer la mise en forme de la
' cellule qui recevra cette formule.
Private Sub SS_ConstruireBlocSolde(ws As Worksheet)

    ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_SOLDE).value = FR("Solde actuel :")
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
' SOUS-PROCEDURE - Bloc de saisie operateur (Beneficiaire, Tiers corrige, Franchise,
' Depassement, Commentaire)
' =====================================================================================
Private Sub SS_ConstruireBlocSaisie(ws As Worksheet)

    Dim rngBeneficiaire As Range
    Dim rngTiersCorrige As Range
    Dim rngFranchise As Range
    Dim rngDepassement As Range
    Dim rngCommentaire As Range

    ' --- Beneficiaire (toujours demande : la colonne demarre vide sur toutes les
    '     lignes existantes, elle n'est jamais renseignee par l'import OFX) ---
    ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_BENEFICIAIRE).value = FR("B{e2}n{e2}ficiaire :")
    Call SS_MettreEnFormeLibelles(ws, SS_LIGNE_BENEFICIAIRE)
    Set rngBeneficiaire = ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_BENEFICIAIRE & ":" & SS_COL_VALEUR_2 & SS_LIGNE_BENEFICIAIRE)
    rngBeneficiaire.Merge
    Call SS_MettreEnFormeZoneSaisie(rngBeneficiaire)
    Call SS_AppliquerListeDeroulante(rngBeneficiaire, "Beneficiaires")
    Call CreerNomSiAbsentSS(ws, "ssBeneficiaire", ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_BENEFICIAIRE))

    ' --- Tiers corrige (demande seulement si le Tiers importe n'est pas deja une
    '     valeur valide de la liste Praticiens - ce controle sera fait en Phase 2b,
    '     ici on prepare seulement la cellule et sa liste deroulante) ---
    ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_TIERS_CORRIGE).value = FR("Tiers corrig{e2} :")
    Call SS_MettreEnFormeLibelles(ws, SS_LIGNE_TIERS_CORRIGE)
    Set rngTiersCorrige = ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_TIERS_CORRIGE & ":" & SS_COL_VALEUR_2 & SS_LIGNE_TIERS_CORRIGE)
    rngTiersCorrige.Merge
    Call SS_MettreEnFormeZoneSaisie(rngTiersCorrige)
    Call SS_AppliquerListeDeroulante(rngTiersCorrige, "Praticiens")
    Call CreerNomSiAbsentSS(ws, "ssTiersCorrige", ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_TIERS_CORRIGE))

    ' --- Franchise (saisie numerique libre, defaut 0) ---
    ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_FRANCHISE).value = FR("Franchise :")
    Call SS_MettreEnFormeLibelles(ws, SS_LIGNE_FRANCHISE)
    Set rngFranchise = ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_FRANCHISE)
    Call SS_MettreEnFormeZoneSaisie(rngFranchise)
    rngFranchise.NumberFormat = "#,##0.00 " & ChrW(8364)
    Call CreerNomSiAbsentSS(ws, "ssFranchise", rngFranchise)

    ' --- Depassement d'honoraires (liste Oui/Non, pertinent seulement quand les 2
    '     remboursements sont presents - la Phase 2b grisera/masquera ce champ le
    '     cas echeant) ---
    ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_DEPASSEMENT).value = FR("D{e2}passement d'honoraires ? :")
    Call SS_MettreEnFormeLibelles(ws, SS_LIGNE_DEPASSEMENT)
    Set rngDepassement = ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_DEPASSEMENT)
    Call SS_MettreEnFormeZoneSaisie(rngDepassement)
    Call SS_AppliquerListeDeroulanteTexte(rngDepassement, FR("Oui,Non"))
    Call CreerNomSiAbsentSS(ws, "ssDepassement", rngDepassement)

    ' --- Commentaire (texte libre) ---
    ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_COMMENTAIRE).value = FR("Commentaire :")
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
' SOUS-PROCEDURE - Zone technique masquee (memorise la ligne TblOperations en cours)
' =====================================================================================
Private Sub SS_ConstruireZoneTechnique(ws As Worksheet)
    ws.Range(SS_COL_TECHNIQUE & SS_LIGNE_BOUTONS).value = 0
    Call CreerNomSiAbsentSS(ws, "ssLigneEnCours", ws.Range(SS_COL_TECHNIQUE & SS_LIGNE_BOUTONS))
End Sub


' =====================================================================================
' SOUS-PROCEDURES UTILITAIRES DE MISE EN FORME (evitent de repeter le meme code de
' style a chaque champ)
' =====================================================================================
Private Sub SS_MettreEnFormeLibelles(ws As Worksheet, ByVal ligne As Long)
    ' On formate chaque cellule de libelle individuellement (B, D, F) plutot que
    ' la plage entiere B:F, car certaines lignes n'utilisent que la colonne B
    ' comme libelle (les colonnes D/F servant alors a une valeur fusionnee, par
    ' exemple pour le champ Beneficiaire).
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

' Supprime les noms definis qui pointent vers cette feuille avant une reconstruction,
' pour eviter les erreurs "nom deja utilise" si on relance l'installation.
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

' Cree un nom defini seulement s'il n'existe pas deja (evite une erreur si la
' macro d'installation est relancee plusieurs fois sans etre passee par
' SupprimerNomsExistantsSS, par exemple lors de tests manuels).
Private Sub CreerNomSiAbsentSS(ws As Worksheet, ByVal nomCellule As String, ByVal rng As Range)
    On Error Resume Next
    ThisWorkbook.Names(nomCellule).Delete
    On Error GoTo 0
    ws.Names.Add Name:=nomCellule, RefersTo:=rng
End Sub


' =====================================================================================
' OUTILS DEVELOPPEUR (reserves a toi, jamais accessibles depuis l'usage normal du
' classeur) : basculent la visibilite de la feuille pour pouvoir la retoucher.
' =====================================================================================
' =====================================================================================
' AjouterBoutonsAjoutListe : ajoute (ou remplace si deja presents) UNIQUEMENT
' les 2 boutons "+" a cote de Beneficiaire et Tiers corrige, sans reconstruire
' le reste de la feuille (contrairement a CreerFeuilleSuiviSante qui, elle,
' repart de zero et perdrait tes eventuels reglages manuels).
' A executer UNE SEULE FOIS, dans la fenetre Execution immediate (Ctrl+G) :
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

    ' On doit rendre la feuille visible le temps de placer les boutons
    ' (Shapes.Add echoue parfois sur une feuille tres masquee), puis on la
    ' remasque si elle etait masquee au depart.
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
' AjouterChampsContexteSuiviSante : ajoute le bloc "Informations complementaires"
' (Num_Cheque, Date_consult, Spe_Consult, Notes bruts, tous en lecture seule)
' SANS reconstruire le reste de la feuille. A executer UNE SEULE FOIS, dans la
' fenetre Execution immediate (Ctrl+G) :
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
        .value = FR("Informations compl{e2}mentaires")
        .Font.Bold = True
        .Font.Size = 10.5
        .Interior.Color = RGB(250, 250, 248)
    End With

    ' --- Ligne Num_Cheque + Date_consult (2 paires sur la meme ligne) ---
    ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_NUM_CHEQUE).value = "Num_Cheque :"
    ws.Range(SS_COL_LIBELLE_2 & SS_LIGNE_NUM_CHEQUE).value = FR("Date_consult :")
    Call SS_MettreEnFormeLibelles(ws, SS_LIGNE_NUM_CHEQUE)
    ws.Range(SS_COL_VALEUR_2 & SS_LIGNE_NUM_CHEQUE).NumberFormat = "dd/mm/yyyy"
    Call SS_MettreEnFormeValeursLectureSeule(ws, SS_LIGNE_NUM_CHEQUE)
    Call CreerNomSiAbsentSS(ws, "ssNumCheque", ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_NUM_CHEQUE))
    Call CreerNomSiAbsentSS(ws, "ssDateConsult", ws.Range(SS_COL_VALEUR_2 & SS_LIGNE_NUM_CHEQUE))

    ' --- Ligne Spe_Consult ---
    ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_SPE_CONSULT).value = FR("Sp{e2}cialit{e2} consult{e2}e :")
    Call SS_MettreEnFormeLibelles(ws, SS_LIGNE_SPE_CONSULT)
    With ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_SPE_CONSULT & ":" & SS_COL_VALEUR_2 & SS_LIGNE_SPE_CONSULT)
        .Merge
    End With
    Call SS_MettreEnFormeValeursLectureSeule(ws, SS_LIGNE_SPE_CONSULT)
    Call CreerNomSiAbsentSS(ws, "ssSpeConsult", ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_SPE_CONSULT))

    ' --- Ligne Notes (texte brut, potentiellement long -> renvoi a la ligne) ---
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
