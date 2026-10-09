Option Explicit

' =====================================================================================
' MODULE : mod_InstallFormulairesNotes
'
' RÔLE (phase 4a du chantier "Suivi Santé") :
'   Construit la mise en page STATIQUE de deux feuilles masquées, utilisées lorsque le
'   champ "Notes" d'une opération de santé ne peut pas être découpé automatiquement
'   (voir mod_ImportOFX, fonction EstDateValide) :
'
'     1) frm_RapprochementNotes : propose une recherche par 4 filtres en
'        cascade (Date -> Spécialité -> Bénéficiaire -> Montant) parmi les
'        clés "Notes" déjà valides ailleurs dans TblOperations, afin de retrouver
'        la bonne clé sans avoir à la ressaisir.
'
'     2) frm_GenerationCle : si aucune correspondance n'existe, permet de
'        générer une nouvelle cle "AAAAMMJJ;Specialite;Beneficiaire;Montant"
'        à partir de quatre champs.
'
'   Ce module fait UNIQUEMENT placer les cellules, noms et boutons : le
'   fonctionnement est dans mod_FormulairesNotes (phase 4b), qui retrouve les champs par
'   leurs NOMS DÉFINIS (rnDate, rnNotes, gcCleGeneree...) et non par leurs adresses.
'
' REFONTE DU 09/10/2026 (harmonisation des formulaires) :
'   Les deux feuilles sont construites avec les outils COMMUNS de mod_InstallCommun et
'   reproduisent les feuilles RÉELLES retouchées à la main (exports
'   frm_RapprochementNotes_structure.txt et frm_GenerationCle_structure.txt) :
'     - frm_RapprochementNotes : deux lignes de boutons (Pas de correspondance, Valider /
'       Passer, Sortir), bouton "Changer la catégorie" à côté de la catégorie, champs
'       Catégorie et Sous-catégorie, compteur au format commun "Opération: x/y";
'     - frm_GenerationCle : les noms gcNotes et gcNumCheque (posés à la main sur la
'       feuille réelle, et utilisés par mod_FormulairesNotes) sont désormais créés ici.
'   Les anciennes macros de rattrapage (AjouterBoutonPasserEtInstructions,
'   AjouterBoutonChangerCategorie) sont supprimées : la construction complète les inclut.
'   Autres changements : "Chéque" -> "Chèque", "Sous-Caégorie" -> "Sous-catégorie",
'   "le champs Notes" -> "le champ Notes".
'
' À FAIRE POUR INSTALLER CE MODULE :
'   1. Alt+F11, Fichier > Importer un fichier..., puis mod_InstallCommun.bas et ce fichier.
'   2. Dans la fenêtre Exécution immédiate (Ctrl+G), lancer :
'        CreerFeuilleRapprochementNotes
'        CreerFeuilleGenerationCle
'   3. Pour afficher une feuille : AfficherFeuilleNotesPourEdition "frm_RapprochementNotes"
'      (ou "frm_GenerationCle"). Pour la masquer de nouveau : MasquerFeuilleNotesApresEdition "..."
' =====================================================================================


Public Const NOM_FEUILLE_RAPPROCHEMENT As String = "frm_RapprochementNotes"
Public Const NOM_FEUILLE_GENERATION As String = "frm_GenerationCle"

' Colonnes communes aux deux feuilles (trois paires libellé/valeur par ligne).
Private Const FN_COL_LIBELLE_1 As String = "B"
Private Const FN_COL_VALEUR_1 As String = "C"
Private Const FN_COL_LIBELLE_2 As String = "D"
Private Const FN_COL_VALEUR_2 As String = "E"
Private Const FN_COL_LIBELLE_3 As String = "F"
Private Const FN_COL_VALEUR_3 As String = "G"
Private Const FN_COL_TECHNIQUE As String = "J"

' --- Mise en page de frm_RapprochementNotes ---
' En-tête commun (voir mod_InstallCommun) : 2 lignes de boutons + compteur.
'   ligne 1 marge / 2 et 3 boutons / 4 fine ligne / 5 compteur / 6 fine ligne /
'   7 Instructions / 8 fine ligne / 9 et suivantes : corps du formulaire
Private Const RN_LIGNE_BOUTONS As Long = 2              ' 1re rangée de boutons
Private Const RN_LIGNE_BOUTONS_2 As Long = 3            ' 2e rangée de boutons (Passer, Sortir)
Private Const RN_LIGNE_COMPTEUR As Long = 5
Private Const RN_LIGNE_INSTRUCTIONS As Long = 7
Private Const RN_LIGNE_CORPS As Long = 9
Private Const RN_LIGNE_TITRE_OPERATION As Long = 9
Private Const RN_LIGNE_OPERATION As Long = 10           ' Date opération / Chèque / Montant
Private Const RN_LIGNE_CATEGORIE As Long = 11           ' Catégorie / Sous-catégorie (affichage)
Private Const RN_LIGNE_NOTES As Long = 12
Private Const RN_LIGNE_TIERS As Long = 13
Private Const RN_LIGNE_TITRE_RECHERCHE As Long = 15
Private Const RN_LIGNE_DATE As Long = 16
Private Const RN_LIGNE_SPECIALITE As Long = 17
Private Const RN_LIGNE_BENEFICIAIRE As Long = 18
Private Const RN_LIGNE_MONTANT As Long = 19
Private Const RN_LIGNE_CLE_TROUVEE As Long = 21

' Colonnes techniques masquées : listes de candidats pour les quatre filtres en
' cascade, recalculées par la phase 4b à chaque changement de filtre. Prévues ici
' (phase 4a) uniquement pour que les noms définis existent dès le départ.
Public Const RN_COL_LISTE_DATES As String = "L"
Public Const RN_COL_LISTE_SPECIALITES As String = "M"
Public Const RN_COL_LISTE_BENEFICIAIRES As String = "N"
Public Const RN_COL_LISTE_MONTANTS As String = "O"
Public Const RN_COL_LISTE_CLES As String = "P"   ' clé brute associée à chaque montant de la colonne O (même ligne)

' --- Mise en page de frm_GenerationCle ---
' En-tête commun : 1 ligne de boutons, pas de compteur.
'   ligne 1 marge / 2 boutons / 3 fine ligne / 4 Instructions / 5 fine ligne /
'   6 et suivantes : corps du formulaire
Private Const GC_LIGNE_BOUTONS As Long = 2
Private Const GC_LIGNE_INSTRUCTIONS As Long = 4
Private Const GC_LIGNE_CORPS As Long = 6
Private Const GC_LIGNE_TITRE_OPERATION As Long = 6
Private Const GC_LIGNE_OPERATION As Long = 7            ' Date / Tiers
Private Const GC_LIGNE_CHEQUE As Long = 8               ' Num. chèque / Notes
Private Const GC_LIGNE_TITRE_GENERATION As Long = 10
Private Const GC_LIGNE_DATE_CONSULT As Long = 11
Private Const GC_LIGNE_SPECIALITE As Long = 12
Private Const GC_LIGNE_BENEFICIAIRE As Long = 13
Private Const GC_LIGNE_MONTANT As Long = 14
Private Const GC_LIGNE_CLE_GENEREE As Long = 16


' =====================================================================================
' CreerFeuilleRapprochementNotes
' =====================================================================================
' Les NOMS (rnDate, rnNotes, ...) et leurs cellules sont ceux qu'attend
' mod_FormulairesNotes (17 noms). C'est la seule chose qui compte pour le
' fonctionnement : tout le reste est de la présentation.
Sub CreerFeuilleRapprochementNotes()

    Dim ws As Worksheet
    Dim wsPrecedente As Worksheet

    Set wsPrecedente = ActiveSheet

    Set ws = mod_InstallCommun.PreparerFeuille(NOM_FEUILLE_RAPPROCHEMENT, True)
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    ' Colonnes A à P : marge / B à G (formulaire) / marge / I / J technique / K / L à P listes techniques
    mod_InstallCommun.MettreEnForme ws, Array(2, 17.5, 22.33, 14, 25.87, 14, 16, 2, 10.13, 10, 9.67, 12, 12, 12, 12, 30)

    RN_ConstruireEntete ws
    RN_ConstruireBlocOperation ws
    RN_ConstruireBlocRecherche ws
    RN_ConstruireZoneTechnique ws
    RN_ConstruireBoutons ws     ' en dernier : les boutons se placent d'après les hauteurs de lignes

    mod_InstallCommun.TerminerFeuille ws, wsPrecedente
    Application.ScreenUpdating = True

    MsgBox "La feuille '" & NOM_FEUILLE_RAPPROCHEMENT & mod_Display.FR("' a {e2}t{e2} cr{e2}{e2}e et masqu{e2}e.") & vbCrLf & _
           "Pour la revoir : AfficherFeuilleNotesPourEdition " & Chr(34) & NOM_FEUILLE_RAPPROCHEMENT & Chr(34), _
           vbInformation, mod_Display.FR("Installation termin{e2}e")

End Sub

Private Sub RN_ConstruireEntete(ByVal ws As Worksheet)

    If mod_InstallCommun.PremiereLigneCorps(2, True) <> RN_LIGNE_CORPS Or _
       mod_InstallCommun.LigneInstructions(2, True) <> RN_LIGNE_INSTRUCTIONS Or _
       mod_InstallCommun.LigneCompteur(2, True) <> RN_LIGNE_COMPTEUR Then
        MsgBox "mod_InstallFormulairesNotes : constantes de lignes RN incoherentes avec mod_InstallCommun.", vbCritical
    End If

    mod_InstallCommun.PoserHauteursEntete ws, 2, True

    ' Compteur "Opération: x/y", rempli par mod_FormulairesNotes.
    mod_InstallCommun.PoserCompteur ws, RN_LIGNE_COMPTEUR, 2, 7
    mod_InstallCommun.PoserNom ws, "rnCompteurCas", ws.Cells(RN_LIGNE_COMPTEUR, 2)

    mod_InstallCommun.EcrireInstructions ws, RN_LIGNE_INSTRUCTIONS, 2, 2, 3, 7, RN_TexteInstructions()

End Sub

' Texte des instructions. Construit ligne par ligne dans une variable : une seule
' instruction VBA ne peut pas dépasser 24 suites de ligne (" _").
Private Function RN_TexteInstructions() As String

    Dim t As String

    t = "Cette op{e2}ration de sant{e2} n'a pas pu {ea}tre rapproch{e2}e automatiquement car le champ [c:Notes] n'a pas {e2}t{e2} correctement renseign{e2}. "
    t = t & "Ce formulaire permet de v{e2}rifier si une autre op{e2}ration est d{e2}j{a2} enregistr{e2}e pour cette consultation."
    t = t & Chr(10) & "S{e2}lectionnez la [c:Date consultation] dans la liste. Si elle est absente, cliquez sur [b:Pas de correspondance] pour cr{e2}er une nouvelle cl{e2}. "
    t = t & "Si elle est pr{e2}sente, continuez {a2} descendre : [c:Sp{e2}cialit{e2}], puis [c:B{e2}n{e2}ficiaire], puis [c:Montant], "
    t = t & "pour retrouver la cl{e2} parmi celles d{e2}j{a2} connues ; elle s'affiche dans [c:Cl{e2} trouv{e2}e]."
    t = t & Chr(10) & "Boutons :"
    t = t & Chr(10) & "- [b:Valider] : applique la cl{e2} trouv{e2}e {a2} cette op{e2}ration. ATTENTION : elle ne sera PLUS JAMAIS repropos{e2}e "
    t = t & "(rapprochement consid{e2}r{e2} r{e2}gl{e2} d{e2}finitivement)."
    t = t & Chr(10) & "- [b:Pas de correspondance] : ouvre [f:frm_GenerationCle] pour cr{e2}er une toute nouvelle cl{e2} (aucune cl{e2} existante ne convient)."
    t = t & Chr(10) & "- [b:Passer] : ne modifie RIEN sur cette op{e2}ration ; elle sera automatiquement repropos{e2}e la prochaine fois "
    t = t & "que la v{e2}rification sera relanc{e2}e."
    t = t & Chr(10) & "- [b:Sortir] : ferme le formulaire et revient sur [f:Synthese]. Le traitement peut {ea}tre relanc{e2} avec le bouton [b:Retraiter suivi de sant{e2}]."
    t = t & Chr(10) & "- [b:Changer la cat{e2}gorie] : si ce n'est pas une d{e2}pense de sant{e2} (erreur {a2} l'import), corrige sa cat{e2}gorie. "
    t = t & "Elle sort de l'analyse sant{e2} et ses colonnes sant{e2} sont vid{e2}es. Si elle reste en Frais, remb sant{e2} (ou en cas d'annulation), "
    t = t & "ce formulaire revient sur la m{ea}me op{e2}ration."
    t = t & Chr(10) & "ATTENTION, ligne de ventilation : c'est TOUTE la ventilation de l'op{e2}ration qui est rouverte puis r{e2}{e2}crite. "
    t = t & "Les rapprochements sant{e2} d{e2}j{a2} faits sur ses AUTRES lignes sont perdus et seront {a2} refaire."

    RN_TexteInstructions = t

End Function

' Bloc "Opération concernée" (lecture seule). Grille de 4 lignes, 3 paires libellé/valeur :
'   10 : Date opération | Chèque         | Montant
'   11 : Catégorie      | Sous-catégorie | (bouton "Changer la catégorie")
'   12 : Notes (valeur fusionnée de C à G)
'   13 : Tiers (valeur fusionnée de C à G)
Private Sub RN_ConstruireBlocOperation(ByVal ws As Worksheet)

    mod_InstallCommun.PoserTitreBloc ws.Range(FN_COL_LIBELLE_1 & RN_LIGNE_TITRE_OPERATION & ":" & FN_COL_VALEUR_3 & RN_LIGNE_TITRE_OPERATION), _
                                     mod_Display.FR("Op{e2}ration concern{e2}e")

    ' Ligne 10 : Date opération / Chèque / Montant
    mod_InstallCommun.PoserEtiquette ws.Range(FN_COL_LIBELLE_1 & RN_LIGNE_OPERATION), mod_Display.FR("Date op{e2}ration")
    mod_InstallCommun.PoserEtiquette ws.Range(FN_COL_LIBELLE_2 & RN_LIGNE_OPERATION), mod_Display.FR("Num. ch{e1}que")
    mod_InstallCommun.PoserEtiquette ws.Range(FN_COL_LIBELLE_3 & RN_LIGNE_OPERATION), "Montant"
    mod_InstallCommun.PoserChampLecture ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_OPERATION), "dd/mm/yyyy"
    mod_InstallCommun.PoserChampLecture ws.Range(FN_COL_VALEUR_2 & RN_LIGNE_OPERATION), "@"
    mod_InstallCommun.PoserChampLecture ws.Range(FN_COL_VALEUR_3 & RN_LIGNE_OPERATION), "#,##0.00 " & ChrW(8364)
    mod_InstallCommun.PoserNom ws, "rnDateOp", ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_OPERATION)
    mod_InstallCommun.PoserNom ws, "rnNumCheque", ws.Range(FN_COL_VALEUR_2 & RN_LIGNE_OPERATION)
    mod_InstallCommun.PoserNom ws, "rnMontantOp", ws.Range(FN_COL_VALEUR_3 & RN_LIGNE_OPERATION)

    ' Ligne 11 : Catégorie / Sous-catégorie (affichage), remplies par
    ' mod_FormulairesNotes.AfficherRapprochementPourLigne via les noms rnCategorie
    ' et rnSousCategorie.
    mod_InstallCommun.PoserEtiquette ws.Range(FN_COL_LIBELLE_1 & RN_LIGNE_CATEGORIE), mod_Display.FR("Cat{e2}gorie")
    mod_InstallCommun.PoserEtiquette ws.Range(FN_COL_LIBELLE_2 & RN_LIGNE_CATEGORIE), mod_Display.FR("Sous-cat{e2}gorie")
    mod_InstallCommun.PoserChampLecture ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_CATEGORIE), "@"
    mod_InstallCommun.PoserChampLecture ws.Range(FN_COL_VALEUR_2 & RN_LIGNE_CATEGORIE), "@"
    mod_InstallCommun.PoserNom ws, "rnCategorie", ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_CATEGORIE)
    mod_InstallCommun.PoserNom ws, "rnSousCategorie", ws.Range(FN_COL_VALEUR_2 & RN_LIGNE_CATEGORIE)

    ' Ligne 12 : Notes (valeur fusionnée sur C:G, retour à la ligne)
    mod_InstallCommun.PoserEtiquette ws.Range(FN_COL_LIBELLE_1 & RN_LIGNE_NOTES), "Notes"
    ws.Range(FN_COL_LIBELLE_1 & RN_LIGNE_NOTES).VerticalAlignment = xlTop
    mod_InstallCommun.PoserChampLecture ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_NOTES & ":" & FN_COL_VALEUR_3 & RN_LIGNE_NOTES), "@", True
    mod_InstallCommun.PoserNom ws, "rnNotes", ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_NOTES)

    ' Ligne 13 : Tiers (valeur fusionnée sur C:G : "[ligne de ventilation]" peut s'y ajouter)
    mod_InstallCommun.PoserEtiquette ws.Range(FN_COL_LIBELLE_1 & RN_LIGNE_TIERS), "Tiers"
    mod_InstallCommun.PoserChampLecture ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_TIERS & ":" & FN_COL_VALEUR_3 & RN_LIGNE_TIERS), "@"
    mod_InstallCommun.PoserNom ws, "rnTiersOp", ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_TIERS)

    ws.Rows(RN_LIGNE_TIERS + 1).RowHeight = mod_InstallCommun.FRM_H_SEP

End Sub

' Bloc "Recherche de la clé (filtres en cascade)" et "Clé trouvée".
Private Sub RN_ConstruireBlocRecherche(ByVal ws As Worksheet)

    mod_InstallCommun.PoserTitreBloc ws.Range(FN_COL_LIBELLE_1 & RN_LIGNE_TITRE_RECHERCHE & ":" & FN_COL_VALEUR_3 & RN_LIGNE_TITRE_RECHERCHE), _
                                     mod_Display.FR("Recherche de la cl{e2} (filtres en cascade)")

    RN_PoserFiltre ws, RN_LIGNE_DATE, "Date consultation", "rnDate"
    RN_PoserFiltre ws, RN_LIGNE_SPECIALITE, mod_Display.FR("Sp{e2}cialit{e2}"), "rnSpecialite"
    RN_PoserFiltre ws, RN_LIGNE_BENEFICIAIRE, mod_Display.FR("B{e2}n{e2}ficiaire"), "rnBeneficiaire"
    RN_PoserFiltre ws, RN_LIGNE_MONTANT, "Montant", "rnMontant"

    ws.Rows(RN_LIGNE_MONTANT + 1).RowHeight = mod_InstallCommun.FRM_H_SEP

    ' --- Clé trouvée (lecture seule) ---
    mod_InstallCommun.PoserEtiquette ws.Range(FN_COL_LIBELLE_1 & RN_LIGNE_CLE_TROUVEE), mod_Display.FR("Cl{e2} trouv{e2}e")
    mod_InstallCommun.PoserChampLecture ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_CLE_TROUVEE & ":" & FN_COL_VALEUR_3 & RN_LIGNE_CLE_TROUVEE), "@"
    ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_CLE_TROUVEE).Font.Bold = True
    mod_InstallCommun.PoserNom ws, "rnCleTrouvee", ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_CLE_TROUVEE)

End Sub

' Pose un filtre de la recherche en cascade : étiquette en B, zone de saisie fusionnée
' de C à F (modifiable), puis le nom sur la 1re cellule de la zone.
' Le format Texte ("@") évite qu'Excel ne transforme une date choisie dans la liste
' en nombre (même précaution que dans mod_FormulairesNotes).
Private Sub RN_PoserFiltre(ByVal ws As Worksheet, ByVal ligne As Long, ByVal libelle As String, ByVal nomCellule As String)
    mod_InstallCommun.PoserEtiquette ws.Range(FN_COL_LIBELLE_1 & ligne), libelle
    mod_InstallCommun.PoserChampSaisie ws.Range(FN_COL_VALEUR_1 & ligne & ":" & FN_COL_LIBELLE_3 & ligne), "@"
    ws.Range(FN_COL_VALEUR_1 & ligne).HorizontalAlignment = xlCenter
    mod_InstallCommun.PoserNom ws, nomCellule, ws.Range(FN_COL_VALEUR_1 & ligne)
End Sub

' Zone technique masquée : listes de candidats des filtres et ligne en cours.
Private Sub RN_ConstruireZoneTechnique(ByVal ws As Worksheet)

    ws.Range(RN_COL_LISTE_DATES & "1").Value = "ListeDates"
    ws.Range(RN_COL_LISTE_SPECIALITES & "1").Value = "ListeSpecialites"
    ws.Range(RN_COL_LISTE_BENEFICIAIRES & "1").Value = "ListeBeneficiaires"
    ws.Range(RN_COL_LISTE_MONTANTS & "1").Value = "ListeMontants"
    ws.Range(RN_COL_LISTE_CLES & "1").Value = "ListeCles"

    ' Chaque liste commence par UNE cellule de réserve (ligne 2), que
    ' mod_FormulairesNotes redimensionne dynamiquement (EcrireListeEtRedefinirNom).
    ' Ces 5 noms DOIVENT être au niveau de la FEUILLE : EcrireListeEtRedefinirNom les
    ' modifie via ws.Names(...). C'est ce que fait PoserNom.
    mod_InstallCommun.PoserNom ws, "rnListeDates", ws.Range(RN_COL_LISTE_DATES & "2")
    mod_InstallCommun.PoserNom ws, "rnListeSpecialites", ws.Range(RN_COL_LISTE_SPECIALITES & "2")
    mod_InstallCommun.PoserNom ws, "rnListeBeneficiaires", ws.Range(RN_COL_LISTE_BENEFICIAIRES & "2")
    mod_InstallCommun.PoserNom ws, "rnListeMontants", ws.Range(RN_COL_LISTE_MONTANTS & "2")
    mod_InstallCommun.PoserNom ws, "rnListeCles", ws.Range(RN_COL_LISTE_CLES & "2")

    ws.Range(RN_COL_LISTE_DATES & ":" & RN_COL_LISTE_CLES).Columns.Hidden = True

    ws.Range(FN_COL_TECHNIQUE & RN_LIGNE_BOUTONS).Value = 0
    mod_InstallCommun.PoserNom ws, "rnLigneEnCours", ws.Range(FN_COL_TECHNIQUE & RN_LIGNE_BOUTONS)
    ws.Columns(FN_COL_TECHNIQUE).Hidden = True

End Sub

' Boutons : 2 lignes en haut (largeurs relevées sur la feuille réelle, export du
' 08/10/2026) et "Changer la catégorie" à côté de la catégorie de l'opération.
Private Sub RN_ConstruireBoutons(ByVal ws As Worksheet)

    Dim gauche As Double

    gauche = ws.Cells(1, 2).Left
    mod_InstallCommun.AjouterBoutonEntete ws, RN_LIGNE_BOUTONS, gauche, mod_Display.FR("Pas de correspondance"), _
        "PasDeCorrespondanceNotes", "btnPasDeCorrespondance", 120
    mod_InstallCommun.AjouterBoutonEntete ws, RN_LIGNE_BOUTONS, gauche, mod_InstallCommun.CapValider(), _
        "ValiderRapprochementNotes", "btnValiderRapprochement", 112.5

    gauche = ws.Cells(1, 2).Left
    mod_InstallCommun.AjouterBoutonEntete ws, RN_LIGNE_BOUTONS_2, gauche, mod_InstallCommun.CapPasser(), _
        "PasserRapprochementNotes", "btnPasserRapprochement", 120
    mod_InstallCommun.AjouterBoutonEntete ws, RN_LIGNE_BOUTONS_2, gauche, mod_InstallCommun.CapSortir(), _
        "SortirRapprochementNotes", "btnSortirRapprochement", 112.5

    mod_InstallCommun.AjouterBoutonCellule ws, ws.Range(FN_COL_LIBELLE_3 & RN_LIGNE_CATEGORIE), mod_Display.FR("Changer la cat{e2}gorie"), _
        "ChangerCategorieRapprochementNotes", "btnChangerCategorie", 120.7

End Sub


' =====================================================================================
' MACRO D'INSTALLATION - frm_GenerationCle
' =====================================================================================
Sub CreerFeuilleGenerationCle()

    Dim ws As Worksheet
    Dim wsPrecedente As Worksheet

    Set wsPrecedente = ActiveSheet

    Set ws = mod_InstallCommun.PreparerFeuille(NOM_FEUILLE_GENERATION, True)
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    ' Colonnes : marge / B à G (formulaire) / marge / I / J technique masquée
    mod_InstallCommun.MettreEnForme ws, Array(2, 22, 16, 16, 16, 14, 16, 2, 10.13, 10)

    GC_ConstruireEntete ws
    GC_ConstruireBlocOperation ws
    GC_ConstruireBlocGeneration ws
    GC_ConstruireBoutons ws     ' en dernier : les boutons se placent d'après les hauteurs de lignes

    ' --- Zone technique masquée ---
    ws.Range(FN_COL_TECHNIQUE & GC_LIGNE_BOUTONS).Value = 0
    mod_InstallCommun.PoserNom ws, "gcLigneEnCours", ws.Range(FN_COL_TECHNIQUE & GC_LIGNE_BOUTONS)
    ws.Columns(FN_COL_TECHNIQUE).Hidden = True

    mod_InstallCommun.TerminerFeuille ws, wsPrecedente
    Application.ScreenUpdating = True

    MsgBox "La feuille '" & NOM_FEUILLE_GENERATION & mod_Display.FR("' a {e2}t{e2} cr{e2}{e2}e et masqu{e2}e.") & vbCrLf & _
           "Pour la revoir : AfficherFeuilleNotesPourEdition " & Chr(34) & NOM_FEUILLE_GENERATION & Chr(34), _
           vbInformation, mod_Display.FR("Installation termin{e2}e")

End Sub

Private Sub GC_ConstruireEntete(ByVal ws As Worksheet)

    If mod_InstallCommun.PremiereLigneCorps(1, False) <> GC_LIGNE_CORPS Or _
       mod_InstallCommun.LigneInstructions(1, False) <> GC_LIGNE_INSTRUCTIONS Then
        MsgBox "mod_InstallFormulairesNotes : constantes de lignes GC incoherentes avec mod_InstallCommun.", vbCritical
    End If

    mod_InstallCommun.PoserHauteursEntete ws, 1, False
    mod_InstallCommun.EcrireInstructions ws, GC_LIGNE_INSTRUCTIONS, 2, 2, 3, 7, GC_TexteInstructions()

End Sub

Private Function GC_TexteInstructions() As String

    Dim t As String

    t = "Aucune cl{e2} existante ne correspond {a2} cette op{e2}ration : ce formulaire permet de cr{e2}er une nouvelle cl{e2} de rapprochement "
    t = t & "([c:Cl{e2} g{e2}n{e2}r{e2}e]) de la forme AAAAMMJJ;Sp{e2}cialit{e2};B{e2}n{e2}ficiaire;Montant."
    t = t & Chr(10) & "Renseignez les 3 champs [c:Date consultation] (JJ/MM/AAAA), [c:Sp{e2}cialit{e2}] et [c:B{e2}n{e2}ficiaire] (le [c:Montant] est repris de l'op{e2}ration), "
    t = t & "cliquez sur [b:G{e2}n{e2}rer] pour construire la cl{e2} (elle est aussi copi{e2}e dans le presse-papier), v{e2}rifiez-la, puis cliquez sur [b:Valider]."
    t = t & Chr(10) & "Boutons :"
    t = t & Chr(10) & "- [b:G{e2}n{e2}rer] : construit la cl{e2} {a2} partir des champs saisis."
    t = t & Chr(10) & "- [b:Valider] : {e2}crit la cl{e2} g{e2}n{e2}r{e2}e dans l'op{e2}ration. ATTENTION : une fois [b:Valider] cliqu{e2}, la cl{e2} est {e2}crite "
    t = t & "D{E2}FINITIVEMENT, m{ea}me approximative ou incorrecte : l'op{e2}ration ne sera PLUS JAMAIS repropos{e2}e pour correction automatique "
    t = t & "(seule une modification manuelle dans [f:Import_Data] pourrait la faire r{e2}appara{i2}tre)."
    t = t & Chr(10) & "- [b:Passer] : laisse cette op{e2}ration de c{o2}t{e2} pour y revenir plus tard (utile si vous n'{ea}tes pas certain des informations)."
    t = t & Chr(10) & "- [b:Retour] : revient sur [f:frm_RapprochementNotes] pour la m{ea}me op{e2}ration, sans rien modifier (par exemple apr{e1}s une erreur de clic sur [b:Pas de correspondance])."
    t = t & Chr(10) & "- [b:+] (en face de [c:Sp{e2}cialit{e2}] et [c:B{e2}n{e2}ficiaire]) : ajoute une nouvelle valeur {a2} la liste."

    GC_TexteInstructions = t

End Function

' Bloc "Opération concernée" (lecture seule)
'   7 : Date | Tiers (valeur fusionnée de E à G)
'   8 : Num. chèque | Notes (valeur fusionnée de E à G, sur 2 lignes)
Private Sub GC_ConstruireBlocOperation(ByVal ws As Worksheet)

    mod_InstallCommun.PoserTitreBloc ws.Range(FN_COL_LIBELLE_1 & GC_LIGNE_TITRE_OPERATION & ":" & FN_COL_VALEUR_3 & GC_LIGNE_TITRE_OPERATION), _
                                     mod_Display.FR("Op{e2}ration concern{e2}e")

    mod_InstallCommun.PoserEtiquette ws.Range(FN_COL_LIBELLE_1 & GC_LIGNE_OPERATION), "Date"
    mod_InstallCommun.PoserEtiquette ws.Range(FN_COL_LIBELLE_2 & GC_LIGNE_OPERATION), "Tiers"
    mod_InstallCommun.PoserChampLecture ws.Range(FN_COL_VALEUR_1 & GC_LIGNE_OPERATION), "dd/mm/yyyy"
    mod_InstallCommun.PoserChampLecture ws.Range(FN_COL_VALEUR_2 & GC_LIGNE_OPERATION & ":" & FN_COL_VALEUR_3 & GC_LIGNE_OPERATION), "@"
    mod_InstallCommun.PoserNom ws, "gcDateOp", ws.Range(FN_COL_VALEUR_1 & GC_LIGNE_OPERATION)
    mod_InstallCommun.PoserNom ws, "gcTiersOp", ws.Range(FN_COL_VALEUR_2 & GC_LIGNE_OPERATION)

    mod_InstallCommun.PoserEtiquette ws.Range(FN_COL_LIBELLE_1 & GC_LIGNE_CHEQUE), mod_Display.FR("Num. ch{e1}que")
    mod_InstallCommun.PoserEtiquette ws.Range(FN_COL_LIBELLE_2 & GC_LIGNE_CHEQUE), "Notes"
    ws.Range(FN_COL_LIBELLE_2 & GC_LIGNE_CHEQUE).VerticalAlignment = xlTop
    mod_InstallCommun.PoserChampLecture ws.Range(FN_COL_VALEUR_1 & GC_LIGNE_CHEQUE), "@", True
    mod_InstallCommun.PoserChampLecture ws.Range(FN_COL_VALEUR_2 & GC_LIGNE_CHEQUE & ":" & FN_COL_VALEUR_3 & GC_LIGNE_CHEQUE), "@", True
    ' (noms créés à la main sur la feuille réelle, et utilisés par mod_FormulairesNotes)
    mod_InstallCommun.PoserNom ws, "gcNumCheque", ws.Range(FN_COL_VALEUR_1 & GC_LIGNE_CHEQUE)
    mod_InstallCommun.PoserNom ws, "gcNotes", ws.Range(FN_COL_VALEUR_2 & GC_LIGNE_CHEQUE)

    ws.Rows(GC_LIGNE_CHEQUE + 1).RowHeight = mod_InstallCommun.FRM_H_SEP

End Sub

' Bloc "Génération de la clé"
Private Sub GC_ConstruireBlocGeneration(ByVal ws As Worksheet)

    Dim rng As Range

    mod_InstallCommun.PoserTitreBloc ws.Range(FN_COL_LIBELLE_1 & GC_LIGNE_TITRE_GENERATION & ":" & FN_COL_VALEUR_3 & GC_LIGNE_TITRE_GENERATION), _
                                     mod_Display.FR("G{e2}n{e2}ration de la cl{e2}")

    ' --- Date de consultation ---
    mod_InstallCommun.PoserEtiquette ws.Range(FN_COL_LIBELLE_1 & GC_LIGNE_DATE_CONSULT), "Date consultation"
    Set rng = ws.Range(FN_COL_VALEUR_1 & GC_LIGNE_DATE_CONSULT)
    mod_InstallCommun.PoserChampSaisie rng, "dd/mm/yyyy"
    mod_InstallCommun.PoserNom ws, "gcDateConsult", rng

    ' --- Spécialité (liste Specialites; le bouton "+" est posé avec les autres boutons) ---
    mod_InstallCommun.PoserEtiquette ws.Range(FN_COL_LIBELLE_1 & GC_LIGNE_SPECIALITE), mod_Display.FR("Sp{e2}cialit{e2}")
    Set rng = ws.Range(FN_COL_VALEUR_1 & GC_LIGNE_SPECIALITE & ":" & FN_COL_VALEUR_2 & GC_LIGNE_SPECIALITE)
    mod_InstallCommun.PoserChampSaisie rng, "@"
    mod_InstallCommun.PoserListeDeroulante rng, "Specialites"
    mod_InstallCommun.PoserNom ws, "gcSpecialite", ws.Range(FN_COL_VALEUR_1 & GC_LIGNE_SPECIALITE)

    ' --- Bénéficiaire (liste Beneficiaires) ---
    mod_InstallCommun.PoserEtiquette ws.Range(FN_COL_LIBELLE_1 & GC_LIGNE_BENEFICIAIRE), mod_Display.FR("B{e2}n{e2}ficiaire")
    Set rng = ws.Range(FN_COL_VALEUR_1 & GC_LIGNE_BENEFICIAIRE & ":" & FN_COL_VALEUR_2 & GC_LIGNE_BENEFICIAIRE)
    mod_InstallCommun.PoserChampSaisie rng, "@"
    mod_InstallCommun.PoserListeDeroulante rng, "Beneficiaires"
    mod_InstallCommun.PoserNom ws, "gcBeneficiaire", ws.Range(FN_COL_VALEUR_1 & GC_LIGNE_BENEFICIAIRE)

    ' --- Montant (repris de l'opération, lecture seule) ---
    mod_InstallCommun.PoserEtiquette ws.Range(FN_COL_LIBELLE_1 & GC_LIGNE_MONTANT), "Montant"
    mod_InstallCommun.PoserChampLecture ws.Range(FN_COL_VALEUR_1 & GC_LIGNE_MONTANT), "#,##0.00 " & ChrW(8364)
    mod_InstallCommun.PoserNom ws, "gcMontant", ws.Range(FN_COL_VALEUR_1 & GC_LIGNE_MONTANT)

    ws.Rows(GC_LIGNE_MONTANT + 1).RowHeight = mod_InstallCommun.FRM_H_SEP

    ' --- Clé générée (lecture seule) ---
    mod_InstallCommun.PoserEtiquette ws.Range(FN_COL_LIBELLE_1 & GC_LIGNE_CLE_GENEREE), mod_Display.FR("Cl{e2} g{e2}n{e2}r{e2}e")
    mod_InstallCommun.PoserChampLecture ws.Range(FN_COL_VALEUR_1 & GC_LIGNE_CLE_GENEREE & ":" & FN_COL_VALEUR_3 & GC_LIGNE_CLE_GENEREE), "@"
    ws.Range(FN_COL_VALEUR_1 & GC_LIGNE_CLE_GENEREE).Font.Bold = True
    mod_InstallCommun.PoserNom ws, "gcCleGeneree", ws.Range(FN_COL_VALEUR_1 & GC_LIGNE_CLE_GENEREE)

End Sub

Private Sub GC_ConstruireBoutons(ByVal ws As Worksheet)

    Dim gauche As Double

    gauche = ws.Cells(1, 2).Left
    mod_InstallCommun.AjouterBoutonEntete ws, GC_LIGNE_BOUTONS, gauche, mod_Display.FR("G{e2}n{e2}rer"), _
        "GenererCleNotes", "btnGenererCle", mod_InstallCommun.FRM_BTN_L
    mod_InstallCommun.AjouterBoutonEntete ws, GC_LIGNE_BOUTONS, gauche, mod_InstallCommun.CapValider(), _
        "ValiderGenerationCle", "btnValiderGeneration", mod_InstallCommun.FRM_BTN_L
    mod_InstallCommun.AjouterBoutonEntete ws, GC_LIGNE_BOUTONS, gauche, mod_InstallCommun.CapPasser(), _
        "PasserGenerationCle", "btnPasserGeneration", mod_InstallCommun.FRM_BTN_L
    mod_InstallCommun.AjouterBoutonEntete ws, GC_LIGNE_BOUTONS, gauche, mod_InstallCommun.CapRetour(), _
        "RetourGenerationCle", "btnRetourGeneration", mod_InstallCommun.FRM_BTN_L

    mod_InstallCommun.AjouterBoutonCellule ws, ws.Range(FN_COL_LIBELLE_3 & GC_LIGNE_SPECIALITE), mod_InstallCommun.CapAjouter(), _
        "AjouterSpecialite", "btnAjouterSpecialite", mod_InstallCommun.FRM_BTN_PLUS
    mod_InstallCommun.AjouterBoutonCellule ws, ws.Range(FN_COL_LIBELLE_3 & GC_LIGNE_BENEFICIAIRE), mod_InstallCommun.CapAjouter(), _
        "AjouterBeneficiaireDepuisGeneration", "btnAjouterBeneficiaireGC", mod_InstallCommun.FRM_BTN_PLUS

End Sub
