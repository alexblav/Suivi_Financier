Option Explicit

' =====================================================================================
' MODULE : mod_InstallSuiviSanteSheet
'
' RÔLE (phase 2a du chantier "Suivi Santé") :
'   Ce module contient UNIQUEMENT la construction de la mise en page statique de la
'   feuille masquée qui sert de formulaire opérateur pour valider les dépenses
'   de santé ("Frais, remb santé") nécessitant une intervention manuelle.
'   La logique de remplissage des cas et de gestion des clics se trouve dans
'   mod_SuiviSanteFormulaire (phase 2b).
'
'   Ce module reprend les mêmes principes que les autres mod_Install* : construire
'   les murs avant de brancher l'électricité.
'
' REFONTE DU 09/10/2026 (harmonisation des formulaires) :
'   La feuille est construite avec les outils COMMUNS de mod_InstallCommun et reproduit
'   la feuille RÉELLE retouchée à la main (export frm_SuiviSante_structure.txt) :
'   bloc "Informations complémentaires" (chèque, date et spécialité de consultation,
'   notes) intégré au bloc "Dépense à traiter", boutons "+" à côté de Bénéficiaire et
'   Tiers corrigé, champ "Dépassement validé ?". Les anciennes macros de rattrapage
'   AjouterBoutonsAjoutListe et AjouterChampsContexteSuiviSante sont supprimées : la
'   construction complète les inclut. Le compteur est sur sa propre ligne, au format
'   commun "Opération: x/y" (voir mod_SuiviSanteFormulaire.MettreAJourCompteur).
'   Tous les champs sont désignés par des NOMS DÉFINIS (ssDate, ssMontant, ...) : ce
'   sont eux, et non les adresses de cellules, qu'utilise mod_SuiviSanteFormulaire.
'
' À FAIRE POUR INSTALLER CE MODULE :
'   1. Alt+F11 pour ouvrir l'éditeur VBA.
'   2. Fichier > Importer un fichier... : mod_InstallCommun.bas puis ce fichier.
'   3. Dans la fenêtre Exécution immédiate (Ctrl+G), taper :
'        CreerFeuilleSuiviSante
'      puis appuyer sur Entrée. La feuille est créée, mise en forme, puis masquée.
'   4. Pour la revoir à l'écran : AfficherFeuilleSuiviSantePourEdition.
'      Pour la masquer de nouveau : MasquerFeuilleSuiviSanteApresEdition.
'
' À PROPOS DES ACCENTS : les textes affichés sont construits par mod_Display.FR().
' =====================================================================================

Public Const NOM_FEUILLE_SUIVI_SANTE As String = "frm_SuiviSante"

' --- En-tête commun (voir mod_InstallCommun) : 1 ligne de boutons + compteur ---
'   ligne 1 marge / 2 boutons / 3 fine ligne / 4 compteur / 5 fine ligne /
'   6 Instructions / 7 fine ligne / 8 et suivantes : corps du formulaire
Public Const SS_LIGNE_BOUTONS As Long = 2
Private Const SS_LIGNE_COMPTEUR As Long = 4
Private Const SS_LIGNE_INSTRUCTIONS As Long = 6
Private Const SS_LIGNE_CORPS As Long = 8

' --- Bloc "Dépense à traiter" (lecture seule) ---
Public Const SS_LIGNE_TITRE_DEPENSE As Long = 8
Public Const SS_LIGNE_DEPENSE As Long = 9            ' Date opération / Tiers importé / Montant
Public Const SS_LIGNE_CHEQUE As Long = 10            ' N° de chèque / Date de consultation
Public Const SS_LIGNE_SPE_CONSULT As Long = 11
Public Const SS_LIGNE_NOTES As Long = 12

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

' --- Colonnes (une lettre = une colonne Excel). Le formulaire est organisé en
'     trois paires "libellé / valeur" par ligne (B/C, D/E, F/G) pour les blocs
'     affichant plusieurs informations sur une même ligne. -------------------------
Public Const SS_COL_LIBELLE_1 As String = "B"
Public Const SS_COL_VALEUR_1 As String = "C"
Public Const SS_COL_LIBELLE_2 As String = "D"
Public Const SS_COL_VALEUR_2 As String = "E"
Public Const SS_COL_LIBELLE_3 As String = "F"
Public Const SS_COL_VALEUR_3 As String = "G"

' Colonne technique (masquée) où est conservé le numéro de ligne de la dépense en cours
' de traitement dans sa table source (nom défini ssLigneEnCours). Jamais visible.
Public Const SS_COL_TECHNIQUE As String = "J"


' =====================================================================================
' MACRO PRINCIPALE D'INSTALLATION
' À exécuter UNE SEULE FOIS (ou de nouveau pour réinitialiser complètement la mise
' en forme de la feuille).
' =====================================================================================
Sub CreerFeuilleSuiviSante()

    Dim ws As Worksheet
    Dim wsPrecedente As Worksheet

    Set wsPrecedente = ActiveSheet

    Set ws = mod_InstallCommun.PreparerFeuille(NOM_FEUILLE_SUIVI_SANTE, True)
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    ' Colonnes : marge / 3 paires libellé-valeur / marge / (I) / technique masquée (J)
    mod_InstallCommun.MettreEnForme ws, Array(2, 21.8, 16, 16, 16, 14, 16, 2, 10.13, 10)
    ws.Columns(SS_COL_TECHNIQUE).Hidden = True

    SS_ConstruireEntete ws
    SS_ConstruireBlocDepense ws
    SS_ConstruireBlocRemboursements ws
    SS_ConstruireBlocSolde ws
    SS_ConstruireBlocSaisie ws
    SS_ConstruireZoneTechnique ws
    SS_ConstruireBoutons ws     ' en dernier : les boutons se placent d'après les hauteurs de lignes

    mod_InstallCommun.TerminerFeuille ws, wsPrecedente
    Application.ScreenUpdating = True

    MsgBox mod_Display.FR("La feuille '") & NOM_FEUILLE_SUIVI_SANTE & mod_Display.FR("' a {e2}t{e2} cr{e2}{e2}e et mise en forme,") & vbCrLf & _
           mod_Display.FR("puis masqu{e2}e (xlSheetVeryHidden).") & vbCrLf & vbCrLf & _
           mod_Display.FR("Pour la revoir {a2} l'{e2}cran et v{e2}rifier le rendu, ex{e2}cutez la macro :") & vbCrLf & _
           "   AfficherFeuilleSuiviSantePourEdition", _
           vbInformation, mod_Display.FR("Installation termin{e2}e")

End Sub


' =====================================================================================
' En-tête commun : hauteurs, compteur et bloc Instructions
' =====================================================================================
Private Sub SS_ConstruireEntete(ByVal ws As Worksheet)

    If mod_InstallCommun.PremiereLigneCorps(1, True) <> SS_LIGNE_CORPS Or _
       mod_InstallCommun.LigneInstructions(1, True) <> SS_LIGNE_INSTRUCTIONS Or _
       mod_InstallCommun.LigneCompteur(1, True) <> SS_LIGNE_COMPTEUR Then
        MsgBox "mod_InstallSuiviSanteSheet : constantes de lignes incoherentes avec mod_InstallCommun.", vbCritical
    End If

    mod_InstallCommun.PoserHauteursEntete ws, 1, True

    ' Compteur "Opération: x/y", rempli par mod_SuiviSanteFormulaire.MettreAJourCompteur;
    ' même principe que CompteurCasRestants (résolution des catégories) : un nom défini.
    mod_InstallCommun.PoserCompteur ws, SS_LIGNE_COMPTEUR, 2, 7
    mod_InstallCommun.PoserNom ws, "CompteurCasSante", ws.Cells(SS_LIGNE_COMPTEUR, 2)

    mod_InstallCommun.EcrireInstructions ws, SS_LIGNE_INSTRUCTIONS, 2, 2, 3, 7, SS_TexteInstructions()

End Sub

Private Function SS_TexteInstructions() As String

    Dim t As String

    t = "Cette d{e2}pense de sant{e2} n'est pas encore sold{e2}e ({e2}quilibre non atteint) : ce formulaire permet de la compl{e2}ter."
    t = t & Chr(10) & "Les cas sont pr{e2}sent{e2}s un par un. V{e2}rifiez le bloc D{e2}pense {a2} traiter et les remboursements li{e2}s, "
    t = t & "puis renseignez les champs demand{e2}s :"
    t = t & Chr(10) & "1. Choisissez le [c:B{e2}n{e2}ficiaire] dans la liste."
    t = t & Chr(10) & "2. Si le [c:Tiers import{e2}] n'est pas correct, choisissez la bonne valeur dans [c:Tiers corrig{e2}]."
    t = t & Chr(10) & "3. Indiquez le montant de [c:Franchise] retenu par l'assurance (0 si aucune)."
    t = t & Chr(10) & "4. Le champ [c:D{e2}passement valid{e2} ?] est accessible d{e1}s que le [c:Solde actuel] n'est pas nul, que 1 ou 2 "
    t = t & "remboursements soient arriv{e2}s. Si vous consid{e2}rez ce solde comme d{e2}finitif et normal ({a2} expliquer dans [c:Commentaire]), "
    t = t & "r{e2}pondez Oui : la d{e2}pense ne sera plus jamais repropos{e2}e, m{ea}me si son statut reste affich{e2} KO. "
    t = t & "Si un remboursement suppl{e2}mentaire doit encore arriver, laissez ce champ vide (ou r{e2}pondez Non) : la d{e2}pense sera repropos{e2}e au prochain passage."
    t = t & Chr(10) & "5. Ajoutez un [c:Commentaire] si besoin, puis cliquez sur [b:Valider]."
    t = t & Chr(10) & "Boutons :"
    t = t & Chr(10) & "- [b:Valider] : contr{o2}le la saisie, l'enregistre et passe au cas suivant."
    t = t & Chr(10) & "- [b:Passer] : passe au cas suivant sans rien enregistrer ; la d{e2}pense sera repropos{e2}e au prochain traitement."
    t = t & Chr(10) & "- [b:Sortir] : ferme le formulaire et revient sur [f:Synthese]. Le traitement peut {ea}tre relanc{e2} avec le bouton [b:Retraiter suivi de sant{e2}]."
    t = t & Chr(10) & "- [b:+] (en face de [c:B{e2}n{e2}ficiaire] et [c:Tiers corrig{e2}]) : ajoute une nouvelle valeur {a2} la liste."

    SS_TexteInstructions = t

End Function


' =====================================================================================
' Bloc "Dépense à traiter" (lecture seule)
' =====================================================================================
Private Sub SS_ConstruireBlocDepense(ByVal ws As Worksheet)

    mod_InstallCommun.PoserTitreBloc ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_TITRE_DEPENSE & ":" & SS_COL_VALEUR_3 & SS_LIGNE_TITRE_DEPENSE), _
                                     mod_Display.FR("D{e2}pense {a2} traiter")

    ' Ligne 1 : Date opération / Tiers importé / Montant
    mod_InstallCommun.PoserEtiquette ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_DEPENSE), mod_Display.FR("Date op{e2}ration")
    mod_InstallCommun.PoserEtiquette ws.Range(SS_COL_LIBELLE_2 & SS_LIGNE_DEPENSE), mod_Display.FR("Tiers import{e2}")
    mod_InstallCommun.PoserEtiquette ws.Range(SS_COL_LIBELLE_3 & SS_LIGNE_DEPENSE), "Montant"
    mod_InstallCommun.PoserChampLecture ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_DEPENSE), "dd/mm/yyyy"
    mod_InstallCommun.PoserChampLecture ws.Range(SS_COL_VALEUR_2 & SS_LIGNE_DEPENSE), "@"
    mod_InstallCommun.PoserChampLecture ws.Range(SS_COL_VALEUR_3 & SS_LIGNE_DEPENSE), "#,##0.00 " & ChrW(8364)
    mod_InstallCommun.PoserNom ws, "ssDate", ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_DEPENSE)
    mod_InstallCommun.PoserNom ws, "ssTiersImporte", ws.Range(SS_COL_VALEUR_2 & SS_LIGNE_DEPENSE)
    mod_InstallCommun.PoserNom ws, "ssMontant", ws.Range(SS_COL_VALEUR_3 & SS_LIGNE_DEPENSE)

    ' Ligne 2 : N° de chèque / Date de consultation
    mod_InstallCommun.PoserEtiquette ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_CHEQUE), mod_Display.FR("Num_Cheque")
    mod_InstallCommun.PoserEtiquette ws.Range(SS_COL_LIBELLE_2 & SS_LIGNE_CHEQUE), "Date_consult"
    mod_InstallCommun.PoserChampLecture ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_CHEQUE), "@"
    mod_InstallCommun.PoserChampLecture ws.Range(SS_COL_VALEUR_2 & SS_LIGNE_CHEQUE), "dd/mm/yyyy"
    mod_InstallCommun.PoserNom ws, "ssNumCheque", ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_CHEQUE)
    mod_InstallCommun.PoserNom ws, "ssDateConsult", ws.Range(SS_COL_VALEUR_2 & SS_LIGNE_CHEQUE)

    ' Ligne 3 : Spécialité consultée
    mod_InstallCommun.PoserEtiquette ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_SPE_CONSULT), mod_Display.FR("Sp{e2}cialit{e2} consult{e2}e")
    mod_InstallCommun.PoserChampLecture ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_SPE_CONSULT & ":" & SS_COL_VALEUR_2 & SS_LIGNE_SPE_CONSULT), "@"
    mod_InstallCommun.PoserNom ws, "ssSpeConsult", ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_SPE_CONSULT)

    ' Ligne 4 : Notes (texte brut, potentiellement long -> retour à la ligne)
    mod_InstallCommun.PoserEtiquette ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_NOTES), "Notes"
    ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_NOTES).VerticalAlignment = xlTop
    mod_InstallCommun.PoserChampLecture ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_NOTES & ":" & SS_COL_VALEUR_3 & SS_LIGNE_NOTES), "@", True
    mod_InstallCommun.PoserNom ws, "ssNotes", ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_NOTES)

    ws.Rows(SS_LIGNE_NOTES + 1).RowHeight = mod_InstallCommun.FRM_H_SEP

End Sub


' =====================================================================================
' Bloc "Remboursements liés" (lecture seule, jusqu'à deux lignes)
' =====================================================================================
Private Sub SS_ConstruireBlocRemboursements(ByVal ws As Worksheet)

    mod_InstallCommun.PoserTitreBloc ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_TITRE_REMBOURSEMENTS & ":" & SS_COL_VALEUR_3 & SS_LIGNE_TITRE_REMBOURSEMENTS), _
                                     mod_Display.FR("Remboursements li{e2}s (jusqu'{a2} 2)")

    mod_InstallCommun.PoserEtiquette ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_REMBOURSEMENT_1), "Remb. 1 - Date"
    mod_InstallCommun.PoserEtiquette ws.Range(SS_COL_LIBELLE_2 & SS_LIGNE_REMBOURSEMENT_1), "Montant"
    mod_InstallCommun.PoserEtiquette ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_REMBOURSEMENT_2), "Remb. 2 - Date"
    mod_InstallCommun.PoserEtiquette ws.Range(SS_COL_LIBELLE_2 & SS_LIGNE_REMBOURSEMENT_2), "Montant"

    mod_InstallCommun.PoserChampLecture ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_REMBOURSEMENT_1), "dd/mm/yyyy"
    mod_InstallCommun.PoserChampLecture ws.Range(SS_COL_VALEUR_2 & SS_LIGNE_REMBOURSEMENT_1), "#,##0.00 " & ChrW(8364)
    mod_InstallCommun.PoserChampLecture ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_REMBOURSEMENT_2), "dd/mm/yyyy"
    mod_InstallCommun.PoserChampLecture ws.Range(SS_COL_VALEUR_2 & SS_LIGNE_REMBOURSEMENT_2), "#,##0.00 " & ChrW(8364)

    mod_InstallCommun.PoserNom ws, "ssRemb1Date", ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_REMBOURSEMENT_1)
    mod_InstallCommun.PoserNom ws, "ssRemb1Montant", ws.Range(SS_COL_VALEUR_2 & SS_LIGNE_REMBOURSEMENT_1)
    mod_InstallCommun.PoserNom ws, "ssRemb2Date", ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_REMBOURSEMENT_2)
    mod_InstallCommun.PoserNom ws, "ssRemb2Montant", ws.Range(SS_COL_VALEUR_2 & SS_LIGNE_REMBOURSEMENT_2)

    ws.Rows(SS_LIGNE_REMBOURSEMENT_2 + 1).RowHeight = mod_InstallCommun.FRM_H_SEP

End Sub


' =====================================================================================
' Solde actuel (lecture seule, formule dynamique)
' =====================================================================================
' Le solde est calculé PAR UNE FORMULE EXCEL CLASSIQUE (pas par du VBA) : il est recalculé
' automatiquement et instantanément dès que l'opérateur modifie la Franchise. La formule
' est posée en fin d'installation (SS_ConstruireZoneTechnique), quand tous les noms
' qu'elle utilise existent; mod_SuiviSanteFormulaire la repose à chaque ouverture.
Private Sub SS_ConstruireBlocSolde(ByVal ws As Worksheet)

    mod_InstallCommun.PoserEtiquette ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_SOLDE), "Solde actuel"
    mod_InstallCommun.PoserChampLecture ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_SOLDE & ":" & SS_COL_VALEUR_2 & SS_LIGNE_SOLDE), _
                                        "#,##0.00 " & ChrW(8364)
    ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_SOLDE).Font.Bold = True
    mod_InstallCommun.PoserNom ws, "ssSolde", ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_SOLDE)

    ws.Rows(SS_LIGNE_SOLDE + 1).RowHeight = mod_InstallCommun.FRM_H_SEP

End Sub


' =====================================================================================
' Bloc de saisie opérateur (Bénéficiaire, Tiers corrigé, Franchise, Dépassement, Commentaire)
' =====================================================================================
Private Sub SS_ConstruireBlocSaisie(ByVal ws As Worksheet)

    Dim rng As Range

    ' --- Bénéficiaire (toujours demandé : la colonne est vide sur toutes les lignes
    '     existantes; elle n'est jamais renseignée par l'import OFX) ---
    mod_InstallCommun.PoserEtiquette ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_BENEFICIAIRE), mod_Display.FR("B{e2}n{e2}ficiaire")
    Set rng = ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_BENEFICIAIRE & ":" & SS_COL_VALEUR_2 & SS_LIGNE_BENEFICIAIRE)
    mod_InstallCommun.PoserChampSaisie rng, "@"
    mod_InstallCommun.PoserListeDeroulante rng, "Beneficiaires"
    mod_InstallCommun.PoserNom ws, "ssBeneficiaire", ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_BENEFICIAIRE)

    ' --- Tiers corrigé (demandé uniquement si le Tiers importé ne figure pas déjà dans
    '     la liste Praticiens; ce contrôle est fait par mod_SuiviSanteFormulaire) ---
    mod_InstallCommun.PoserEtiquette ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_TIERS_CORRIGE), mod_Display.FR("Tiers corrig{e2}")
    Set rng = ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_TIERS_CORRIGE & ":" & SS_COL_VALEUR_2 & SS_LIGNE_TIERS_CORRIGE)
    mod_InstallCommun.PoserChampSaisie rng, "@"
    mod_InstallCommun.PoserListeDeroulante rng, "Praticiens"
    mod_InstallCommun.PoserNom ws, "ssTiersCorrige", ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_TIERS_CORRIGE)

    ' --- Franchise (saisie numérique libre, valeur par défaut : 0) ---
    mod_InstallCommun.PoserEtiquette ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_FRANCHISE), "Franchise"
    Set rng = ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_FRANCHISE)
    mod_InstallCommun.PoserChampSaisie rng, "#,##0.00 " & ChrW(8364)
    mod_InstallCommun.PoserNom ws, "ssFranchise", rng

    ' --- Dépassement validé ? (liste Oui/Non, pertinent seulement si le solde n'est pas nul) ---
    mod_InstallCommun.PoserEtiquette ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_DEPASSEMENT), mod_Display.FR("D{e2}passement valid{e2} ?")
    Set rng = ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_DEPASSEMENT)
    mod_InstallCommun.PoserChampSaisie rng
    SS_AppliquerListeDeroulanteTexte rng, "Oui,Non"
    mod_InstallCommun.PoserNom ws, "ssDepassement", rng

    ' --- Commentaire (texte libre, sur 2 lignes) ---
    mod_InstallCommun.PoserEtiquette ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_COMMENTAIRE), "Commentaire"
    ws.Range(SS_COL_LIBELLE_1 & SS_LIGNE_COMMENTAIRE).VerticalAlignment = xlTop
    Set rng = ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_COMMENTAIRE & ":" & SS_COL_VALEUR_3 & SS_LIGNE_COMMENTAIRE)
    mod_InstallCommun.PoserChampSaisie rng, "@", True
    mod_InstallCommun.PoserNom ws, "ssCommentaire", ws.Range(SS_COL_VALEUR_1 & SS_LIGNE_COMMENTAIRE)

End Sub

Private Sub SS_AppliquerListeDeroulanteTexte(ByVal rng As Range, ByVal listeVirgules As String)
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=listeVirgules
    On Error GoTo 0
End Sub


' =====================================================================================
' Zone technique masquée (mémorise la ligne en cours) et formule du solde
' =====================================================================================
Private Sub SS_ConstruireZoneTechnique(ByVal ws As Worksheet)
    ws.Range(SS_COL_TECHNIQUE & SS_LIGNE_BOUTONS).Value = 0
    mod_InstallCommun.PoserNom ws, "ssLigneEnCours", ws.Range(SS_COL_TECHNIQUE & SS_LIGNE_BOUTONS)

    ' Tous les noms existent maintenant : on peut poser la formule du solde.
    On Error Resume Next
    ws.Range("ssSolde").Formula = "=ssMontant-ssRemb1Montant-ssRemb2Montant-ssFranchise"
    On Error GoTo 0
End Sub


' =====================================================================================
' Boutons
' =====================================================================================
Private Sub SS_ConstruireBoutons(ByVal ws As Worksheet)

    Dim gauche As Double

    gauche = ws.Cells(1, 2).Left

    mod_InstallCommun.AjouterBoutonEntete ws, SS_LIGNE_BOUTONS, gauche, mod_InstallCommun.CapPasser(), _
        "CasSuivantSuiviSante", "btnCasSuivantSuiviSante", mod_InstallCommun.FRM_BTN_L
    mod_InstallCommun.AjouterBoutonEntete ws, SS_LIGNE_BOUTONS, gauche, mod_InstallCommun.CapValider(), _
        "ValiderCasSuiviSante", "btnValiderCasSuiviSante", mod_InstallCommun.FRM_BTN_L
    mod_InstallCommun.AjouterBoutonEntete ws, SS_LIGNE_BOUTONS, gauche, mod_InstallCommun.CapSortir(), _
        "SortirSuiviSante", "btnSortirSuiviSante", mod_InstallCommun.FRM_BTN_L

    ' --- Boutons "+" : ajoutent une valeur aux listes Bénéficiaires et Praticiens ---
    mod_InstallCommun.AjouterBoutonCellule ws, ws.Range(SS_COL_LIBELLE_3 & SS_LIGNE_BENEFICIAIRE), mod_InstallCommun.CapAjouter(), _
        "AjouterBeneficiaire", "btnAjouterBeneficiaire", mod_InstallCommun.FRM_BTN_PLUS
    mod_InstallCommun.AjouterBoutonCellule ws, ws.Range(SS_COL_LIBELLE_3 & SS_LIGNE_TIERS_CORRIGE), mod_InstallCommun.CapAjouter(), _
        "AjouterTiersPraticien", "btnAjouterTiers", mod_InstallCommun.FRM_BTN_PLUS

End Sub


' =====================================================================================
' OUTILS DÉVELOPPEUR (Ctrl+G)
' =====================================================================================
Sub AfficherFeuilleSuiviSantePourEdition()
    mod_InstallCommun.AfficherPourEdition NOM_FEUILLE_SUIVI_SANTE, "CreerFeuilleSuiviSante", "MasquerFeuilleSuiviSanteApresEdition"
End Sub

Sub MasquerFeuilleSuiviSanteApresEdition()
    mod_InstallCommun.MasquerApresEdition NOM_FEUILLE_SUIVI_SANTE
End Sub
