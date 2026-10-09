Option Explicit

' =====================================================================================
' MODULE : mod_InstallFormulairesNotes
'
' RÔLE (phase 4a du chantier "Suivi Santé") :
'   Construit la mise en page STATIQUE (aucune logique de clic pour l'instant)
'   de deux feuilles masquées, utilisées lorsque le champ "Notes" d'une opération
'   de santé ne peut pas être découpé automatiquement (voir mod_ImportOFX,
'   fonction EstDateValide) :
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
'   Comme pour frm_SuiviSante (mod_InstallSuiviSanteSheet), ce module fait
'   UNIQUEMENT placer les cellules, noms et boutons : aucune action n'est encore
'   fonctionnelle. Ce sera l'objet de la phase 4b (mod_FormulairesNotes).
'
' CE MODULE EST VOLONTAIREMENT AUTONOME (il ne réutilise pas les fonctions
' internes de mod_InstallSuiviSanteSheet) : elles sont Private dans leur module,
' et dupliquer ces quelques dizaines de lignes de mise en forme est plus sûr
' que de multiplier les dépendances entre fichiers. Voir le mécanisme équivalent,
' déjà dupliqué volontairement, dans mod_SuiviSanteFormulaire.
'
' À PROPOS DES ACCENTS : même convention que dans le reste du chantier Suivi Santé :
' les textes affichés sont construits via la fonction FR(); les commentaires sont en UTF-8.
'
' A FAIRE POUR INSTALLER CE MODULE :
'   1. Alt+F11, Fichier > Importer un fichier..., puis choisir ce fichier .bas.
'   2. Dans la fenêtre Exécution immédiate (Ctrl+G), lancer :
'        CreerFeuilleRapprochementNotes
'        CreerFeuilleGenerationCle
'   3. Pour afficher une feuille : AfficherFeuilleNotesPourEdition "frm_RapprochementNotes"
'      (ou "frm_GenerationCle"). Pour la masquer de nouveau : MasquerFeuilleNotesApresEdition "..."
'
' MISE À JOUR DU 08/10/2026 (frm_RapprochementNotes) :
'   CreerFeuilleRapprochementNotes reproduit désormais la feuille telle que
'   l'opérateur l'a retouchée à la main (export frm_RapprochementNotes_structure.txt).
'   Sur une feuille DÉJÀ en place, il n'est PAS nécessaire de la reconstruire :
'     - AjouterBoutonChangerCategorie remet le bouton "Changer la catégorie" à sa
'       place (sans doublon) et réécrit les instructions ;
'     - une reconstruction complète (CreerFeuilleRapprochementNotes, réponse "Oui")
'       donne le même résultat, mais efface la mise en forme partielle (gras,
'       couleurs) que l'opérateur aurait posée à la main dans les instructions.
'   frm_GenerationCle n'a pas été revue (pas d'export de sa structure).
' =====================================================================================


Public Const NOM_FEUILLE_RAPPROCHEMENT As String = "frm_RapprochementNotes"
Public Const NOM_FEUILLE_GENERATION As String = "frm_GenerationCle"

' Colonnes communes aux deux feuilles (trois paires libellé/valeur par ligne, comme dans frm_SuiviSante).
Public Const FN_COL_LIBELLE_1 As String = "B"
Public Const FN_COL_VALEUR_1 As String = "C"
Public Const FN_COL_LIBELLE_2 As String = "D"
Public Const FN_COL_VALEUR_2 As String = "E"
Public Const FN_COL_LIBELLE_3 As String = "F"
Public Const FN_COL_VALEUR_3 As String = "G"
Public Const FN_COL_TECHNIQUE As String = "J"

' --- Mise en page de frm_RapprochementNotes ---
' MODIFIÉ le 08/10/2026 : valeurs alignées sur la feuille RÉELLE, retouchée à la main
' par l'opérateur (relevé : frm_RapprochementNotes_structure.txt, export du 08/10/2026).
' Ces constantes ne sont utilisées QUE dans ce module (vérifié) : changer leur valeur
' n'a aucun effet sur mod_FormulairesNotes, qui passe uniquement par les noms "rn...".
Public Const RN_LIGNE_BOUTONS As Long = 2              ' 1re rangée de boutons + compteur "X sur N"
Public Const RN_LIGNE_TITRE_INSTRUCTIONS As Long = 5
Public Const RN_LIGNE_DEBUT_INSTRUCTIONS As Long = 6
Public Const RN_LIGNE_FIN_INSTRUCTIONS As Long = 9
Public Const RN_LIGNE_TITRE_OPERATION As Long = 11
Public Const RN_LIGNE_OPERATION As Long = 12           ' Date opération / Chèque / Montant
Public Const RN_LIGNE_TITRE_RECHERCHE As Long = 17     ' (était 14)
Public Const RN_LIGNE_DATE As Long = 18                ' (était 15)
Public Const RN_LIGNE_SPECIALITE As Long = 19          ' (était 16)
Public Const RN_LIGNE_BENEFICIAIRE As Long = 20        ' (était 17)
Public Const RN_LIGNE_MONTANT As Long = 21             ' (était 18)
Public Const RN_LIGNE_CLE_TROUVEE As Long = 23         ' (était 20)

' Ajout du 08/10/2026 : lignes et réglages qui n'existaient pas dans la version
' d'origine. Déclarés Private (et non Public) : ce sont de simples réglages de mise en
' page, utiles à ce seul module ; ils n'ont donc pas leur place dans mod_VarGlobales.
Private Const RN_LIGNE_BOUTONS_2 As Long = 3           ' 2e rangée de boutons (Passer, Sortir)
Private Const RN_LIGNE_CATEGORIE As Long = 13          ' Catégorie / Sous-catégorie (affichage)
Private Const RN_LIGNE_NOTES As Long = 14
Private Const RN_LIGNE_TIERS As Long = 15
Private Const RN_LIGNE_ESPACE_CLE As Long = 22         ' fine ligne d'espacement avant "Clé trouvée"
' Hauteurs des 4 lignes d'instructions (6 à 8, puis 9) : relevées sur la feuille, la
' dernière agrandie (79,9 -> 120) pour loger le paragraphe "Changer la catégorie".
' À ajuster à l'œil si besoin.
Private Const RN_HAUTEUR_INSTRUCTIONS As Double = 34.9
Private Const RN_HAUTEUR_INSTRUCTIONS_DERNIERE As Double = 120

' Colonnes techniques masquées : listes de candidats pour les quatre filtres en
' cascade, recalculées par la phase 4b à chaque changement de filtre. Prévues ici
' (phase 4a) uniquement pour que les noms définis existent dès le départ.
Public Const RN_COL_LISTE_DATES As String = "L"
Public Const RN_COL_LISTE_SPECIALITES As String = "M"
Public Const RN_COL_LISTE_BENEFICIAIRES As String = "N"
Public Const RN_COL_LISTE_MONTANTS As String = "O"
Public Const RN_COL_LISTE_CLES As String = "P"   ' clé brute associée à chaque montant de la colonne O (même ligne)

' --- Mise en page de frm_GenerationCle ---
Public Const GC_LIGNE_BOUTONS As Long = 2
Public Const GC_LIGNE_TITRE_INSTRUCTIONS As Long = 5
Public Const GC_LIGNE_DEBUT_INSTRUCTIONS As Long = 6
Public Const GC_LIGNE_FIN_INSTRUCTIONS As Long = 9
Public Const GC_LIGNE_TITRE_OPERATION As Long = 11
Public Const GC_LIGNE_OPERATION As Long = 12
Public Const GC_LIGNE_TITRE_GENERATION As Long = 14
Public Const GC_LIGNE_DATE_CONSULT As Long = 15
Public Const GC_LIGNE_SPECIALITE As Long = 16
Public Const GC_LIGNE_BENEFICIAIRE As Long = 17
Public Const GC_LIGNE_MONTANT As Long = 18
Public Const GC_LIGNE_CLE_GENEREE As Long = 20


' =====================================================================================
' CreerFeuilleRapprochementNotes
' =====================================================================================
' RÉÉCRITE le 08/10/2026 : la feuille avait été retouchée à la main par l'opérateur
' (nouveaux champs, boutons "Sortir" et "Changer la catégorie", textes revus,
' largeurs et hauteurs). Cette macro reproduit désormais FIDÈLEMENT la feuille réelle,
' d'après l'export frm_RapprochementNotes_structure.txt du 08/10/2026. Une
' reconstruction ne fait donc plus perdre ces retouches.
'
' Seules exceptions volontaires, signalées à l'opérateur :
'   - 3 fautes de frappe corrigées : "Chéque" -> "Chèque", "Sous-Caégorie" ->
'     "Sous-catégorie" (avec un espace avant ":" comme les autres libellés), et
'     "le champs Notes" -> "le champ Notes" dans les instructions ;
'   - un paragraphe ajouté aux instructions pour le bouton "Changer la catégorie"
'     (avec le point de vigilance sur les ventilations) ;
'   - la mise en forme PARTIELLE des instructions (quelques mots en gras ou en
'     couleur) n'est pas reproduite : l'export ne dit pas quels mots étaient concernés ;
'   - la ligne 20 (Bénéficiaire) prend la même hauteur (18) que les 3 autres filtres ;
'   - les cellules d'affichage Catégorie / Sous-catégorie (C13, E13) reçoivent un
'     format Texte au lieu des formats date/nombre copiés de la ligne 12.
'
' Les NOMS (rnDate, rnNotes, ...) et leurs cellules sont exactement ceux qu'attend
' mod_FormulairesNotes (17 noms, tous vérifiés). C'est la seule chose qui compte pour
' le fonctionnement : tout le reste est de la présentation.
' =====================================================================================
Sub CreerFeuilleRapprochementNotes()

    Dim ws As Worksheet
    Dim reponse As VbMsgBoxResult

    Set ws = ObtenirFeuilleSansErreurFN(NOM_FEUILLE_RAPPROCHEMENT)

    If Not ws Is Nothing Then
        reponse = MsgBox("La feuille '" & NOM_FEUILLE_RAPPROCHEMENT & "' existe deja." & vbCrLf & _
                  "Veux-tu la reconstruire entierement (sa mise en forme actuelle sera perdue) ?", _
                          vbYesNo + vbQuestion, "Confirmation de reconstruction")
        If reponse = vbNo Then
            MsgBox "Installation annulee, aucune modification effectuee.", vbInformation
            Exit Sub
        End If
        ws.Visible = xlSheetVisible
        ws.Cells.Clear
        Call SupprimerFormesExistantesFN(ws)
        Call SupprimerNomsExistantsFN(ws, NOM_FEUILLE_RAPPROCHEMENT)
        ' AJOUT 08/10/2026 : Cells.Clear efface contenu et formats, mais PAS les
        ' hauteurs de lignes ni les largeurs de colonnes. On les remet donc à leur
        ' valeur standard avant de reposer celles de cette feuille ; sinon une
        ' ancienne hauteur (par exemple celle d'une ligne déplacée) resterait en place.
        ws.rows.UseStandardHeight = True
        ws.Columns.UseStandardWidth = True
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = NOM_FEUILLE_RAPPROCHEMENT
    End If

    ' Mise en forme commune aux 2 feuilles (police, quadrillage, colonnes A et H...).
    ' ATTENTION : cette procédure sert AUSSI à frm_GenerationCle ; on ne la modifie
    ' donc pas. Les largeurs propres à cette feuille sont reposées juste après.
    Call FN_AppliquerMiseEnFormeGenerale(ws)

    ' --- Largeurs de colonnes relevées sur la feuille (08/10/2026) ---
    ws.Columns("B").ColumnWidth = 14.4
    ws.Columns("C").ColumnWidth = 22.33
    ws.Columns("D").ColumnWidth = 11.27
    ws.Columns("E").ColumnWidth = 25.87
    ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 16
    ws.Columns("K").ColumnWidth = 9.67

    ' --- Hauteurs de lignes relevées sur la feuille (08/10/2026) ---
    ws.rows(RN_LIGNE_BOUTONS).RowHeight = 22.2
    ws.rows(RN_LIGNE_BOUTONS_2).RowHeight = 22.2
    ws.rows(RN_LIGNE_TITRE_INSTRUCTIONS).RowHeight = 14.3
    ws.rows(RN_LIGNE_TITRE_OPERATION).RowHeight = 13.9
    ws.rows(RN_LIGNE_OPERATION).RowHeight = 18
    ws.rows(RN_LIGNE_CATEGORIE).RowHeight = 18
    ws.rows(RN_LIGNE_TITRE_RECHERCHE - 1).RowHeight = 18      ' ligne vide avant le titre
    ws.rows(RN_LIGNE_TITRE_RECHERCHE).RowHeight = 18
    ws.rows(RN_LIGNE_DATE & ":" & RN_LIGNE_MONTANT).RowHeight = 18
    ws.rows(RN_LIGNE_ESPACE_CLE).RowHeight = 6.4
    ' (les lignes d'instructions 6 à 9 sont réglées par RN_EcrireInstructions)

    ' --- Boutons (5) et compteur "X sur N" ---
    Call RN_PoserTousLesBoutons(ws)

    With ws.Range(FN_COL_LIBELLE_3 & RN_LIGNE_BOUTONS & ":" & FN_COL_VALEUR_3 & RN_LIGNE_BOUTONS)
        .Merge
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .Font.Size = 9
        .Font.Color = RGB(120, 120, 120)
    End With
    Call CreerNomSiAbsentFN(ws, "rnCompteurCas", ws.Range(FN_COL_LIBELLE_3 & RN_LIGNE_BOUTONS))

    ' --- Instructions (titre + texte) ---
    With ws.Range(FN_COL_LIBELLE_1 & RN_LIGNE_TITRE_INSTRUCTIONS)
        .value = mod_Display.FR("Instructions")
        .Font.Bold = True
        .Font.Size = 11
    End With
    Call RN_EcrireInstructions(ws)

    ' --- Bloc "Opération concernée" (lecture seule) ---
    ' Grille de 4 lignes (12 à 15), 3 paires libellé/valeur par ligne :
    '   12 : Date opération | Chèque        | Montant
    '   13 : Catégorie      | Sous-catégorie | (vide)
    '   14 : Notes (valeur fusionnée de C à G)
    '   15 : Tiers          | (vide)         | (vide)
    With ws.Range(FN_COL_LIBELLE_1 & RN_LIGNE_TITRE_OPERATION & ":" & FN_COL_VALEUR_3 & RN_LIGNE_TITRE_OPERATION)
        .Merge
        .value = mod_Display.FR("Op{e2}ration concern{e2}e")
        .Font.Bold = True
        .Font.Size = 10.5
        .Interior.Color = RGB(250, 250, 248)
    End With

    ' Fond gris très clair des zones de valeur, comme sur la feuille : tout le bloc
    ' C12:G15, sauf les cellules des libellés de la colonne D aux lignes 12 et 13, plus
    ' les libellés "Notes" et "Tiers" de la colonne B.
    ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_OPERATION & ":" & FN_COL_VALEUR_3 & RN_LIGNE_TIERS).Interior.Color = RGB(248, 248, 246)
    ws.Range(FN_COL_LIBELLE_2 & RN_LIGNE_OPERATION & ":" & FN_COL_LIBELLE_2 & RN_LIGNE_CATEGORIE).Interior.ColorIndex = xlColorIndexNone
    ws.Range(FN_COL_LIBELLE_1 & RN_LIGNE_NOTES & ":" & FN_COL_LIBELLE_1 & RN_LIGNE_TIERS).Interior.Color = RGB(248, 248, 246)
    ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_OPERATION & ":" & FN_COL_VALEUR_3 & RN_LIGNE_TIERS).Locked = True
    ' Cellules vides du bloc, en gris moyen (comme sur la feuille).
    ws.Range(FN_COL_LIBELLE_3 & RN_LIGNE_CATEGORIE).Font.Color = RGB(90, 90, 90)
    ws.Range(FN_COL_LIBELLE_2 & RN_LIGNE_TIERS).Font.Color = RGB(90, 90, 90)
    ws.Range(FN_COL_LIBELLE_3 & RN_LIGNE_TIERS).Font.Color = RGB(90, 90, 90)

    ' Ligne 12 : Date opération / Chèque / Montant
    Call RN_StyleLibelle(ws.Range(FN_COL_LIBELLE_1 & RN_LIGNE_OPERATION), mod_Display.FR("Date op{e2}ration :"), False)
    Call RN_StyleLibelle(ws.Range(FN_COL_LIBELLE_2 & RN_LIGNE_OPERATION), mod_Display.FR("Ch{e1}que :"), True)
    Call RN_StyleLibelle(ws.Range(FN_COL_LIBELLE_3 & RN_LIGNE_OPERATION), "Montant :", False)
    ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_OPERATION).NumberFormat = "dd/mm/yyyy"
    ws.Range(FN_COL_VALEUR_2 & RN_LIGNE_OPERATION).NumberFormat = "#,#00"     ' format relevé sur la feuille
    ws.Range(FN_COL_VALEUR_3 & RN_LIGNE_OPERATION).NumberFormat = "#,##0.00 " & ChrW(8364)
    Call CreerNomSiAbsentFN(ws, "rnDateOp", ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_OPERATION))
    Call CreerNomSiAbsentFN(ws, "rnNumCheque", ws.Range(FN_COL_VALEUR_2 & RN_LIGNE_OPERATION))
    Call CreerNomSiAbsentFN(ws, "rnMontantOp", ws.Range(FN_COL_VALEUR_3 & RN_LIGNE_OPERATION))

    ' Ligne 13 : Catégorie / Sous-catégorie. AFFICHAGE SEULEMENT pour l'instant :
    ' aucun nom n'est posé sur C13/E13 et mod_FormulairesNotes ne les remplit pas
    ' (cellules ajoutées à la main par l'opérateur ; leur alimentation reste à décider).
    Call RN_StyleLibelle(ws.Range(FN_COL_LIBELLE_1 & RN_LIGNE_CATEGORIE), mod_Display.FR("Cat{e2}gorie :"), False)
    Call RN_StyleLibelle(ws.Range(FN_COL_LIBELLE_2 & RN_LIGNE_CATEGORIE), mod_Display.FR("Sous-cat{e2}gorie :"), True)
    ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_CATEGORIE).NumberFormat = "@"
    ws.Range(FN_COL_VALEUR_2 & RN_LIGNE_CATEGORIE).NumberFormat = "@"

    ' Ligne 14 : Notes (valeur fusionnée sur C:G, centrée, retour à la ligne)
    Call RN_StyleLibelle(ws.Range(FN_COL_LIBELLE_1 & RN_LIGNE_NOTES), "Notes :", False)
    With ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_NOTES & ":" & FN_COL_VALEUR_3 & RN_LIGNE_NOTES)
        .Merge
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
    Call CreerNomSiAbsentFN(ws, "rnNotes", ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_NOTES))

    ' Ligne 15 : Tiers
    Call RN_StyleLibelle(ws.Range(FN_COL_LIBELLE_1 & RN_LIGNE_TIERS), "Tiers :", False)
    Call CreerNomSiAbsentFN(ws, "rnTiersOp", ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_TIERS))

    ' --- Bloc "Recherche de la clé (filtres en cascade)" ---
    With ws.Range(FN_COL_LIBELLE_1 & RN_LIGNE_TITRE_RECHERCHE & ":" & FN_COL_VALEUR_3 & RN_LIGNE_TITRE_RECHERCHE)
        .Merge
        .value = mod_Display.FR("Recherche de la cl{e2} (filtres en cascade)")
        .Font.Bold = True
        .Font.Size = 10.5
        .Interior.Color = RGB(250, 250, 248)
    End With

    ' Les 4 filtres : libellé en B, zone de saisie fusionnée de C à F (au lieu de C:D
    ' dans la version d'origine). Le nom pointe sur la 1re cellule de la fusion.
    Call RN_PoserFiltre(ws, RN_LIGNE_DATE, mod_Display.FR("Date consultation :"), "rnDate")
    Call RN_PoserFiltre(ws, RN_LIGNE_SPECIALITE, mod_Display.FR("Sp{e2}cialit{e2} :"), "rnSpecialite")
    Call RN_PoserFiltre(ws, RN_LIGNE_BENEFICIAIRE, mod_Display.FR("B{e2}n{e2}ficiaire :"), "rnBeneficiaire")
    Call RN_PoserFiltre(ws, RN_LIGNE_MONTANT, "Montant :", "rnMontant")

    ' --- Clé trouvée (lecture seule) ---
    Call RN_StyleLibelle(ws.Range(FN_COL_LIBELLE_1 & RN_LIGNE_CLE_TROUVEE), mod_Display.FR("Cl{e2} trouv{e2}e :"), False)
    With ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_CLE_TROUVEE & ":" & FN_COL_VALEUR_3 & RN_LIGNE_CLE_TROUVEE)
        .Merge
        .Font.Bold = True
        .Interior.Color = RGB(248, 248, 246)
        .Locked = True
    End With
    Call CreerNomSiAbsentFN(ws, "rnCleTrouvee", ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_CLE_TROUVEE))

    ' --- Zone technique masquée (inchangée) ---
    ws.Columns(RN_COL_LISTE_DATES).ColumnWidth = 12
    ws.Columns(RN_COL_LISTE_SPECIALITES).ColumnWidth = 12
    ws.Columns(RN_COL_LISTE_BENEFICIAIRES).ColumnWidth = 12
    ws.Columns(RN_COL_LISTE_MONTANTS).ColumnWidth = 12
    ws.Columns(RN_COL_LISTE_CLES).ColumnWidth = 30
    ws.Range(RN_COL_LISTE_DATES & ":" & RN_COL_LISTE_CLES).Columns.Hidden = True

    ws.Range(RN_COL_LISTE_DATES & "1").value = "ListeDates"
    ws.Range(RN_COL_LISTE_SPECIALITES & "1").value = "ListeSpecialites"
    ws.Range(RN_COL_LISTE_BENEFICIAIRES & "1").value = "ListeBeneficiaires"
    ws.Range(RN_COL_LISTE_MONTANTS & "1").value = "ListeMontants"
    ws.Range(RN_COL_LISTE_CLES & "1").value = "ListeCles"

    ' Chaque liste commence par UNE cellule de réserve (ligne 2), que
    ' mod_FormulairesNotes redimensionne dynamiquement (EcrireListeEtRedefinirNom).
    ' Ces 5 noms DOIVENT être au niveau de la FEUILLE : EcrireListeEtRedefinirNom les
    ' modifie via ws.Names(...). C'est ce que fait CreerNomSiAbsentFN.
    Call CreerNomSiAbsentFN(ws, "rnListeDates", ws.Range(RN_COL_LISTE_DATES & "2"))
    Call CreerNomSiAbsentFN(ws, "rnListeSpecialites", ws.Range(RN_COL_LISTE_SPECIALITES & "2"))
    Call CreerNomSiAbsentFN(ws, "rnListeBeneficiaires", ws.Range(RN_COL_LISTE_BENEFICIAIRES & "2"))
    Call CreerNomSiAbsentFN(ws, "rnListeMontants", ws.Range(RN_COL_LISTE_MONTANTS & "2"))
    Call CreerNomSiAbsentFN(ws, "rnListeCles", ws.Range(RN_COL_LISTE_CLES & "2"))

    ws.Range(FN_COL_TECHNIQUE & RN_LIGNE_BOUTONS).value = 0
    Call CreerNomSiAbsentFN(ws, "rnLigneEnCours", ws.Range(FN_COL_TECHNIQUE & RN_LIGNE_BOUTONS))
    ws.Columns(FN_COL_TECHNIQUE).Hidden = True

    ws.Range("A1").Select
    ws.Visible = xlSheetVeryHidden

    MsgBox "La feuille '" & NOM_FEUILLE_RAPPROCHEMENT & "' a ete creee et masquee." & vbCrLf & _
           "Pour la revoir : AfficherFeuilleNotesPourEdition " & Chr(34) & NOM_FEUILLE_RAPPROCHEMENT & Chr(34), _
           vbInformation, "Installation terminee"

End Sub


' =====================================================================================
' OUTILS PROPRES À frm_RapprochementNotes (ajout du 08/10/2026)
' =====================================================================================

' Pose les 5 boutons de la feuille, aux positions relevées sur la feuille réelle
' (export du 08/10/2026 ; valeurs en points : gauche, haut, largeur, hauteur).
' Ces positions ne suivent pas une grille de cellules : l'opérateur a placé les
' boutons à la main. Les positions dépendent des largeurs de colonnes et des
' hauteurs de lignes posées par CreerFeuilleRapprochementNotes.
' Les boutons "Sortir" et "Changer la catégorie" s'appelaient "Button 12" et
' "Button 13" (noms automatiques d'Excel) : ils reçoivent ici un vrai nom.
Private Sub RN_PoserTousLesBoutons(ByVal ws As Worksheet)
    Call RN_PoserBouton(ws, "btnPasDeCorrespondance", mod_Display.FR("Pas de correspondance"), "PasDeCorrespondanceNotes", 7.5, 6.7, 120, 18.4)
    Call RN_PoserBouton(ws, "btnValiderRapprochement", "Valider", "ValiderRapprochementNotes", 129, 8.2, 112.5, 16.9)
    Call RN_PoserBouton(ws, "btnPasserRapprochement", "Passer", "PasserRapprochementNotes", 7.5, 36, 118.5, 22.1)
    Call RN_PoserBouton(ws, "btnSortirRapprochement", "Sortir", "SortirRapprochementNotes", 127.5, 36, 112.5, 22.1)
    Call RN_PoserBoutonChangerCategorie(ws)
End Sub

' Bouton "Changer la catégorie" seul (réutilisé par AjouterBoutonChangerCategorie).
Private Sub RN_PoserBoutonChangerCategorie(ByVal ws As Worksheet)
    Call RN_PoserBouton(ws, "btnChangerCategorie", mod_Display.FR("Changer la cat{e2}gorie"), "ChangerCategorieRapprochementNotes", 626.3, 319.5, 120.7, 21)
End Sub

' Pose UN bouton (contrôle de formulaire). Avant cela, on supprime tout bouton
' existant qui porte le MÊME NOM ou qui lance la MÊME MACRO : on évite ainsi les
' doublons, même pour un bouton créé à la main sous un nom automatique (par exemple
' "Button 13"), qu'une recherche par nom seule n'aurait pas trouvé.
Private Sub RN_PoserBouton(ByVal ws As Worksheet, ByVal nomBouton As String, ByVal libelle As String, _
                           ByVal macroCible As String, ByVal gauche As Double, ByVal haut As Double, _
                           ByVal largeur As Double, ByVal hauteur As Double)

    Dim i As Long
    Dim macroForme As String
    Dim aSupprimer As Boolean
    Dim btn As Button

    For i = ws.Shapes.count To 1 Step -1        ' à l'envers : on supprime en parcourant
        aSupprimer = (ws.Shapes(i).Name = nomBouton)
        If Not aSupprimer Then
            macroForme = ""
            On Error Resume Next                 ' certaines formes n'ont pas de macro
            macroForme = ws.Shapes(i).OnAction
            On Error GoTo 0
            ' La macro est mémorisée sous la forme "Classeur.xlsm!NomMacro" (ou parfois
            ' "NomMacro" seul) : on compare donc la fin du texte.
            If macroForme <> "" Then
                aSupprimer = (macroForme = macroCible) Or (Right(macroForme, Len(macroCible) + 1) = "!" & macroCible)
            End If
        End If
        If aSupprimer Then ws.Shapes(i).Delete
    Next i

    Set btn = ws.Buttons.Add(gauche, haut, largeur, hauteur)
    With btn
        .Caption = libelle
        .OnAction = macroCible
        .Name = nomBouton
    End With

End Sub

' Écrit un libellé dans le style de cette feuille : gras, taille 10, bleu foncé
' RGB(31, 73, 125). alignerADroite = True pour les libellés de la colonne D.
Private Sub RN_StyleLibelle(ByVal cellule As Range, ByVal texte As String, ByVal alignerADroite As Boolean)
    With cellule
        .value = texte
        .Font.Bold = True
        .Font.Size = 10
        .Font.Color = RGB(31, 73, 125)
        If alignerADroite Then .HorizontalAlignment = xlRight
    End With
End Sub

' Pose un filtre de la recherche en cascade : libellé en B, zone de saisie fusionnée
' de C à F (fond jaune pâle, bordure, déverrouillée, format Texte, centrée), puis le
' nom sur la 1re cellule de la zone.
' Le format Texte ("@") évite qu'Excel ne transforme une date choisie dans la liste
' en nombre (même précaution que dans mod_FormulairesNotes).
Private Sub RN_PoserFiltre(ByVal ws As Worksheet, ByVal ligne As Long, ByVal libelle As String, ByVal nomCellule As String)
    Dim zone As Range
    Call RN_StyleLibelle(ws.Range(FN_COL_LIBELLE_1 & ligne), libelle, False)
    Set zone = ws.Range(FN_COL_VALEUR_1 & ligne & ":" & FN_COL_LIBELLE_3 & ligne)
    zone.Merge
    Call FN_MettreEnFormeZoneSaisie(zone)
    zone.NumberFormat = "@"
    zone.HorizontalAlignment = xlCenter
    Call CreerNomSiAbsentFN(ws, nomCellule, ws.Range(FN_COL_VALEUR_1 & ligne))
End Sub

' Écrit le texte des instructions (zone fusionnée B6:G9) avec sa mise en forme et les
' hauteurs de ses 4 lignes. Utilisée À LA FOIS par CreerFeuilleRapprochementNotes et
' par MettreAJourInstructionsRapprochement : le texte n'existe donc qu'à UN seul
' endroit (RN_TexteInstructions), ce qui évite que les deux versions divergent, comme
' c'était arrivé avant le 08/10/2026.
Private Sub RN_EcrireInstructions(ByVal ws As Worksheet)
    With ws.Range(FN_COL_LIBELLE_1 & RN_LIGNE_DEBUT_INSTRUCTIONS & ":" & FN_COL_VALEUR_3 & RN_LIGNE_FIN_INSTRUCTIONS)
        .Merge
        .value = RN_TexteInstructions()
        .WrapText = True
        .VerticalAlignment = xlTop
        .Font.Size = 9
        .Font.Bold = False
        .Font.Color = RGB(80, 80, 80)
        .Interior.Color = RGB(245, 245, 242)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(210, 210, 205)
        .Locked = True
    End With
    ws.rows(RN_LIGNE_DEBUT_INSTRUCTIONS & ":" & (RN_LIGNE_FIN_INSTRUCTIONS - 1)).RowHeight = RN_HAUTEUR_INSTRUCTIONS
    ws.rows(RN_LIGNE_FIN_INSTRUCTIONS).RowHeight = RN_HAUTEUR_INSTRUCTIONS_DERNIERE
End Sub

' Texte des instructions : celui rédigé par l'opérateur sur la feuille (08/10/2026),
' complété par le paragraphe du bouton "Changer la catégorie". Construit ligne par
' ligne dans une variable : une seule instruction VBA ne peut pas dépasser 24 suites
' de ligne (" _"), et ce texte en demanderait davantage.
Private Function RN_TexteInstructions() As String
    Dim t As String
    t = mod_Display.FR("Cette op{e2}ration de sant{e2} n'a pas pu {ea}tre rapproch{e2}e automatiquement car le champ Notes n'a pas {e2}t{e2} correctement renseign{e2}")
    t = t & Chr(10) & mod_Display.FR("Ce formulaire permet de v{e2}rifier si une autre op{e2}ration est d{e2}j{a2} enregistr{e2}e pour cette consultation.")
    t = t & Chr(10) & mod_Display.FR("S{e2}lectionner la date de consultation dans la liste.")
    t = t & Chr(10) & mod_Display.FR("Si absente cliquer sur 'Pas de correspondance' pour cr{e2}er une nouvelle cl{e2}")
    t = t & Chr(10) & mod_Display.FR("Si pr{e2}sente continuer {a2} descendre: Sp{e2}cialit{e2}, puis B{e2}n{e2}ficiaire, puis Montant pour retrouver la cl{e2} parmi")
    t = t & Chr(10) & mod_Display.FR("celles d{e2}j{a2} connues.")
    t = t & Chr(10) & mod_Display.FR("Cons{e2}quence de chaque bouton :")
    t = t & Chr(10) & mod_Display.FR("- 'Valider' : applique la cl{e2} trouv{e2}e {a2} cette op{e2}ration. Elle ne sera PLUS JAMAIS repropos{e2}e (rapprochement consid{e2}r{e2} r{e2}gl{e2} d{e2}finitivement).")
    t = t & Chr(10) & mod_Display.FR("- 'Pas de correspondance' : ouvre un {e2}cran pour cr{e2}er une toute nouvelle cl{e2} (aucune cl{e2} existante ne convient).")
    t = t & Chr(10) & mod_Display.FR("- 'Passer' : ne modifie RIEN sur cette op{e2}ration. Elle sera automatiquement repropos{e2}e la prochaine fois que la v{e2}rification sera relanc{e2}e")
    t = t & Chr(10) & mod_Display.FR("- 'Sortir': Ferme la fen{ea}tre et revient sur la page Synthese. Il est possible de relancer le traitement avec le bouton: 'Retraiter suivi de sant{e2}'")
    ' --- Ajout du 08/10/2026 : bouton "Changer la catégorie" + point de vigilance ---
    t = t & Chr(10) & mod_Display.FR("- 'Changer la cat{e2}gorie' : si ce n'est pas une d{e2}pense de sant{e2} (erreur {a2} l'import), corrige sa cat{e2}gorie. Elle sort de l'analyse sant{e2} et ses colonnes sant{e2} sont vid{e2}es. Si elle reste en 'Frais, remb sant{e2}' (ou en cas d'annulation), ce formulaire revient sur la m{ea}me op{e2}ration.")
    t = t & Chr(10) & mod_Display.FR("  ATTENTION, ligne de ventilation : c'est TOUTE la ventilation de l'op{e2}ration qui est rouverte puis r{e2}{e2}crite. Les rapprochements sant{e2} d{e2}j{a2} faits sur ses AUTRES lignes sont perdus et seront {a2} refaire.")
    RN_TexteInstructions = t
End Function


' =====================================================================================
' MACRO D'INSTALLATION - frm_GenerationCle
' =====================================================================================
Sub CreerFeuilleGenerationCle()

    Dim ws As Worksheet
    Dim reponse As VbMsgBoxResult

    Set ws = ObtenirFeuilleSansErreurFN(NOM_FEUILLE_GENERATION)

    If Not ws Is Nothing Then
        reponse = MsgBox("La feuille '" & NOM_FEUILLE_GENERATION & "' existe deja." & vbCrLf & _
                          "Veux-tu la reconstruire entierement (sa mise en forme actuelle sera perdue) ?", _
                          vbYesNo + vbQuestion, "Confirmation de reconstruction")
        If reponse = vbNo Then
            MsgBox "Installation annulee, aucune modification effectuee.", vbInformation
            Exit Sub
        End If
        ws.Visible = xlSheetVisible
        ws.Cells.Clear
        Call SupprimerFormesExistantesFN(ws)
        Call SupprimerNomsExistantsFN(ws, NOM_FEUILLE_GENERATION)
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = NOM_FEUILLE_GENERATION
    End If

    Call FN_AppliquerMiseEnFormeGenerale(ws)

    ' --- Boutons ---
    Dim zoneBtn1 As Range, zoneBtn2 As Range
    Dim btn As Button

    Set zoneBtn1 = ws.Range(FN_COL_LIBELLE_1 & GC_LIGNE_BOUTONS & ":" & FN_COL_VALEUR_1 & GC_LIGNE_BOUTONS)
    zoneBtn1.RowHeight = 22
    Set btn = ws.Buttons.Add(zoneBtn1.Left, zoneBtn1.Top, zoneBtn1.Width, zoneBtn1.Height)
    With btn
        .Caption = mod_Display.FR("G{e2}n{e2}rer")
        .OnAction = "GenererCleNotes"
        .Name = "btnGenererCle"
    End With

    Set zoneBtn2 = ws.Range(FN_COL_LIBELLE_2 & GC_LIGNE_BOUTONS & ":" & FN_COL_VALEUR_2 & GC_LIGNE_BOUTONS)
    Set btn = ws.Buttons.Add(zoneBtn2.Left, zoneBtn2.Top, zoneBtn2.Width, zoneBtn2.Height)
    With btn
        .Caption = "Valider"
        .OnAction = "ValiderGenerationCle"
        .Name = "btnValiderGeneration"
    End With

    ' --- Instructions ---
    With ws.Range(FN_COL_LIBELLE_1 & GC_LIGNE_TITRE_INSTRUCTIONS)
        .value = mod_Display.FR("Instructions")
        .Font.Bold = True
        .Font.Size = 11
    End With
    With ws.Range(FN_COL_LIBELLE_1 & GC_LIGNE_DEBUT_INSTRUCTIONS & ":" & FN_COL_VALEUR_3 & GC_LIGNE_FIN_INSTRUCTIONS)
        .Merge
        .value = mod_Display.FR("Aucune cl{e2} existante ne correspond {a2} cette op{e2}ration.") & Chr(10) & _
                 mod_Display.FR("Renseigne les 4 champs ci-dessous, clique sur 'G{e2}n{e2}rer' pour construire la cl{e2}") & Chr(10) & _
                 mod_Display.FR("(elle sera aussi copi{e2}e dans le presse-papier), v{e2}rifie-la, puis clique sur 'Valider'.")
        .WrapText = True
        .VerticalAlignment = xlTop
        .Font.Size = 9
        .Font.Color = RGB(80, 80, 80)
        .Interior.Color = RGB(245, 245, 242)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(210, 210, 205)
        .Locked = True
    End With
    ws.rows(GC_LIGNE_DEBUT_INSTRUCTIONS & ":" & GC_LIGNE_FIN_INSTRUCTIONS).RowHeight = 15

    ' --- Bloc "Opération concernée" (lecture seule) ---
    With ws.Range(FN_COL_LIBELLE_1 & GC_LIGNE_TITRE_OPERATION & ":" & FN_COL_VALEUR_3 & GC_LIGNE_TITRE_OPERATION)
        .Merge
        .value = mod_Display.FR("Op{e2}ration concern{e2}e")
        .Font.Bold = True
        .Font.Size = 10.5
        .Interior.Color = RGB(250, 250, 248)
    End With
    ws.Range(FN_COL_LIBELLE_1 & GC_LIGNE_OPERATION).value = "Date :"
    ws.Range(FN_COL_LIBELLE_2 & GC_LIGNE_OPERATION).value = "Tiers :"
    Call FN_MettreEnFormeLibelles(ws, GC_LIGNE_OPERATION)
    ws.Range(FN_COL_VALEUR_1 & GC_LIGNE_OPERATION).NumberFormat = "dd/mm/yyyy"
    Call FN_MettreEnFormeValeursLectureSeule(ws, GC_LIGNE_OPERATION)
    Call CreerNomSiAbsentFN(ws, "gcDateOp", ws.Range(FN_COL_VALEUR_1 & GC_LIGNE_OPERATION))
    Call CreerNomSiAbsentFN(ws, "gcTiersOp", ws.Range(FN_COL_VALEUR_2 & GC_LIGNE_OPERATION))

    ' --- Bloc "Génération de la clé" ---
    With ws.Range(FN_COL_LIBELLE_1 & GC_LIGNE_TITRE_GENERATION & ":" & FN_COL_VALEUR_3 & GC_LIGNE_TITRE_GENERATION)
        .Merge
        .value = mod_Display.FR("G{e2}n{e2}ration de la cl{e2}")
        .Font.Bold = True
        .Font.Size = 10.5
        .Interior.Color = RGB(250, 250, 248)
    End With

    ws.Range(FN_COL_LIBELLE_1 & GC_LIGNE_DATE_CONSULT).value = mod_Display.FR("Date_consult (JJ/MM/AAAA) :")
    Call FN_MettreEnFormeLibelles(ws, GC_LIGNE_DATE_CONSULT)
    Dim rngDateConsult As Range
    Set rngDateConsult = ws.Range(FN_COL_VALEUR_1 & GC_LIGNE_DATE_CONSULT)
    Call FN_MettreEnFormeZoneSaisie(rngDateConsult)
    rngDateConsult.NumberFormat = "dd/mm/yyyy"
    Call CreerNomSiAbsentFN(ws, "gcDateConsult", rngDateConsult)

    ws.Range(FN_COL_LIBELLE_1 & GC_LIGNE_SPECIALITE).value = mod_Display.FR("Sp{e2}cialit{e2} :")
    Call FN_MettreEnFormeLibelles(ws, GC_LIGNE_SPECIALITE)
    Dim rngSpecialite As Range
    Set rngSpecialite = ws.Range(FN_COL_VALEUR_1 & GC_LIGNE_SPECIALITE & ":" & FN_COL_VALEUR_2 & GC_LIGNE_SPECIALITE)
    rngSpecialite.Merge
    Call FN_MettreEnFormeZoneSaisie(rngSpecialite)
    Call FN_AppliquerListeDeroulante(rngSpecialite, "Specialites")
    Call CreerNomSiAbsentFN(ws, "gcSpecialite", ws.Range(FN_COL_VALEUR_1 & GC_LIGNE_SPECIALITE))

    Dim btnPlusSpecialite As Button
    Dim zonePlusSpecialite As Range
    Set zonePlusSpecialite = ws.Range(FN_COL_LIBELLE_3 & GC_LIGNE_SPECIALITE)
    Set btnPlusSpecialite = ws.Buttons.Add(zonePlusSpecialite.Left, zonePlusSpecialite.Top, 26, zonePlusSpecialite.Height)
    With btnPlusSpecialite
        .Caption = "+"
        .OnAction = "AjouterSpecialite"
        .Name = "btnAjouterSpecialite"
    End With

    ws.Range(FN_COL_LIBELLE_1 & GC_LIGNE_BENEFICIAIRE).value = mod_Display.FR("B{e2}n{e2}ficiaire :")
    Call FN_MettreEnFormeLibelles(ws, GC_LIGNE_BENEFICIAIRE)
    Dim rngBeneficiaire As Range
    Set rngBeneficiaire = ws.Range(FN_COL_VALEUR_1 & GC_LIGNE_BENEFICIAIRE & ":" & FN_COL_VALEUR_2 & GC_LIGNE_BENEFICIAIRE)
    rngBeneficiaire.Merge
    Call FN_MettreEnFormeZoneSaisie(rngBeneficiaire)
    Call FN_AppliquerListeDeroulante(rngBeneficiaire, "Beneficiaires")
    Call CreerNomSiAbsentFN(ws, "gcBeneficiaire", ws.Range(FN_COL_VALEUR_1 & GC_LIGNE_BENEFICIAIRE))

    Dim btnPlusBeneficiaire As Button
    Dim zonePlusBeneficiaire As Range
    Set zonePlusBeneficiaire = ws.Range(FN_COL_LIBELLE_3 & GC_LIGNE_BENEFICIAIRE)
    Set btnPlusBeneficiaire = ws.Buttons.Add(zonePlusBeneficiaire.Left, zonePlusBeneficiaire.Top, 26, zonePlusBeneficiaire.Height)
    With btnPlusBeneficiaire
        .Caption = "+"
        .OnAction = "AjouterBeneficiaireDepuisGeneration"
        .Name = "btnAjouterBeneficiaireGC"
    End With

    ws.Range(FN_COL_LIBELLE_1 & GC_LIGNE_MONTANT).value = "Montant :"
    Call FN_MettreEnFormeLibelles(ws, GC_LIGNE_MONTANT)
    ws.Range(FN_COL_VALEUR_1 & GC_LIGNE_MONTANT).NumberFormat = "#,##0.00 " & ChrW(8364)
    Call FN_MettreEnFormeValeursLectureSeule(ws, GC_LIGNE_MONTANT)
    Call CreerNomSiAbsentFN(ws, "gcMontant", ws.Range(FN_COL_VALEUR_1 & GC_LIGNE_MONTANT))

    ' --- Clé générée (lecture seule) ---
    ws.Range(FN_COL_LIBELLE_1 & GC_LIGNE_CLE_GENEREE).value = mod_Display.FR("Cl{e2} g{e2}n{e2}r{e2}e :")
    Call FN_MettreEnFormeLibelles(ws, GC_LIGNE_CLE_GENEREE)
    With ws.Range(FN_COL_VALEUR_1 & GC_LIGNE_CLE_GENEREE & ":" & FN_COL_VALEUR_3 & GC_LIGNE_CLE_GENEREE)
        .Merge
        .Font.Bold = True
    End With
    Call FN_MettreEnFormeValeursLectureSeule(ws, GC_LIGNE_CLE_GENEREE)
    Call CreerNomSiAbsentFN(ws, "gcCleGeneree", ws.Range(FN_COL_VALEUR_1 & GC_LIGNE_CLE_GENEREE))

    ' --- Zone technique masquée ---
    ws.Range(FN_COL_TECHNIQUE & GC_LIGNE_BOUTONS).value = 0
    Call CreerNomSiAbsentFN(ws, "gcLigneEnCours", ws.Range(FN_COL_TECHNIQUE & GC_LIGNE_BOUTONS))
    ws.Columns(FN_COL_TECHNIQUE).Hidden = True

    ws.Visible = xlSheetVeryHidden

    MsgBox "La feuille '" & NOM_FEUILLE_GENERATION & "' a ete creee et masquee." & vbCrLf & _
           "Pour la revoir : AfficherFeuilleNotesPourEdition " & Chr(34) & NOM_FEUILLE_GENERATION & Chr(34), _
           vbInformation, "Installation terminee"

End Sub


' =====================================================================================
' MISE EN FORME GÉNÉRALE (commune aux deux feuilles)
' =====================================================================================
Private Sub FN_AppliquerMiseEnFormeGenerale(ws As Worksheet)

    ws.Activate
    ActiveWindow.DisplayGridlines = False

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 22
    ws.Columns("C").ColumnWidth = 16
    ws.Columns("D").ColumnWidth = 16
    ws.Columns("E").ColumnWidth = 16
    ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 16
    ws.Columns("H").ColumnWidth = 2

    ws.Columns(FN_COL_TECHNIQUE).ColumnWidth = 10

    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10

    ws.Range("A1").Select
    ActiveWindow.DisplayHeadings = False

End Sub


' =====================================================================================
' UTILITAIRES DE MISE EN FORME (communs aux deux feuilles, dans le même esprit
' que les fonctions SS_xxx de mod_InstallSuiviSanteSheet)
' =====================================================================================
Private Sub FN_MettreEnFormeLibelles(ws As Worksheet, ByVal ligne As Long)
    ws.Range(FN_COL_LIBELLE_1 & ligne).Font.Color = RGB(90, 90, 90)
    ws.Range(FN_COL_LIBELLE_2 & ligne).Font.Color = RGB(90, 90, 90)
    ws.Range(FN_COL_LIBELLE_3 & ligne).Font.Color = RGB(90, 90, 90)
    ws.rows(ligne).RowHeight = 18
End Sub

Private Sub FN_MettreEnFormeValeursLectureSeule(ws As Worksheet, ByVal ligne As Long)
    With ws.Range(FN_COL_VALEUR_1 & ligne & ":" & FN_COL_VALEUR_3 & ligne)
        .Locked = True
        .Interior.Color = RGB(248, 248, 246)
    End With
End Sub

Private Sub FN_MettreEnFormeZoneSaisie(ByVal rng As Range)
    With rng
        .Locked = False
        .Interior.Color = RGB(255, 255, 235)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 190, 140)
    End With
End Sub

Private Sub FN_AppliquerListeDeroulante(ByVal rng As Range, ByVal nomPlage As String)
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & nomPlage
    On Error GoTo 0
End Sub


' =====================================================================================
' FONCTIONS UTILITAIRES (création/suppression de feuilles, noms et formes)
' =====================================================================================
Private Function ObtenirFeuilleSansErreurFN(ByVal nomFeuille As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nomFeuille)
    On Error GoTo 0
    Set ObtenirFeuilleSansErreurFN = ws
End Function

Private Sub SupprimerFormesExistantesFN(ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
End Sub

Private Sub SupprimerNomsExistantsFN(ws As Worksheet, ByVal nomFeuille As String)
    Dim n As Name
    Dim i As Long
    For i = ThisWorkbook.Names.count To 1 Step -1
        Set n = ThisWorkbook.Names(i)
        On Error Resume Next
        If InStr(1, n.RefersTo, "'" & nomFeuille & "'", vbTextCompare) > 0 Then
            n.Delete
        End If
        On Error GoTo 0
    Next i
End Sub

Private Sub CreerNomSiAbsentFN(ws As Worksheet, ByVal nomCellule As String, ByVal rng As Range)
    On Error Resume Next
    ThisWorkbook.Names(nomCellule).Delete
    On Error GoTo 0
    ws.Names.Add Name:=nomCellule, RefersTo:=rng
End Sub


' =====================================================================================
' OUTILS DÉVELOPPEUR : basculent la visibilité d'une des deux feuilles pour
' permettre de la retoucher. Passer le nom exact de la feuille en paramètre.
' =====================================================================================
' =====================================================================================
' AjouterBoutonPasserEtInstructions : ajoute UNIQUEMENT le bouton "Passer" sur
' les deux feuilles (ligne 3, sous les boutons existants) et met à jour le texte
' des instructions pour préciser les conséquences de chaque action, SANS
' reconstruire le reste des feuilles. À exécuter UNE SEULE FOIS dans la fenêtre
' Exécution immédiate (Ctrl+G) :
'      AjouterBoutonPasserEtInstructions
' =====================================================================================
' MODIFIÉ le 08/10/2026 : pour frm_RapprochementNotes, le bouton "Passer" est
' désormais reposé à sa position RÉELLE (voir RN_PoserTousLesBoutons), et non plus
' en B3:C3 ; sinon cette macro le déplacerait et écraserait la mise en page de
' l'opérateur. La partie frm_GenerationCle est inchangée.
Sub AjouterBoutonPasserEtInstructions()

    Dim wsRn As Worksheet
    Dim etaitMasquee As Boolean

    Set wsRn = ObtenirFeuilleSansErreurFN(NOM_FEUILLE_RAPPROCHEMENT)
    If wsRn Is Nothing Then
        MsgBox "La feuille '" & NOM_FEUILLE_RAPPROCHEMENT & "' n'existe pas encore.", vbExclamation
    Else
        etaitMasquee = (wsRn.Visible <> xlSheetVisible)
        wsRn.Visible = xlSheetVisible
        Call RN_PoserBouton(wsRn, "btnPasserRapprochement", "Passer", "PasserRapprochementNotes", 7.5, 36, 118.5, 22.1)
        If etaitMasquee Then wsRn.Visible = xlSheetVeryHidden
    End If
    Call MettreAJourInstructionsRapprochement

    Call AjouterBoutonPasserSurFeuille(NOM_FEUILLE_GENERATION, GC_LIGNE_BOUTONS, "PasserGenerationCle", "btnPasserGeneration")
    Call MettreAJourInstructionsGeneration

    MsgBox "Bouton 'Passer' et instructions mis a jour sur les 2 feuilles.", vbInformation

End Sub

' =====================================================================================
' AjouterBoutonChangerCategorie (ajout du 07/10/2026)
' =====================================================================================
' Ajoute UNIQUEMENT le bouton "Changer la catégorie" sur frm_RapprochementNotes et
' met à jour le texte des instructions, SANS reconstruire le reste de la feuille.
' À exécuter dans la fenêtre Exécution immédiate (Ctrl+G) :
'      AjouterBoutonChangerCategorie
' Peut être relancée sans risque : tout bouton du même nom OU lançant la même macro
' (par exemple le "Button 13" posé à la main) est d'abord supprimé : pas de doublon.
' MODIFIÉ le 08/10/2026 : le bouton est posé à la position choisie par l'opérateur
' (à droite du bloc "Opération concernée", lignes 13-14), et non plus en D3:E3, où il
' aurait recouvert le bouton "Sortir".
' =====================================================================================
Sub AjouterBoutonChangerCategorie()

    Dim ws As Worksheet
    Dim etaitMasquee As Boolean

    Set ws = ObtenirFeuilleSansErreurFN(NOM_FEUILLE_RAPPROCHEMENT)
    If ws Is Nothing Then
        MsgBox "La feuille '" & NOM_FEUILLE_RAPPROCHEMENT & "' n'existe pas encore.", vbExclamation
        Exit Sub
    End If

    etaitMasquee = (ws.Visible <> xlSheetVisible)
    ws.Visible = xlSheetVisible
    Call RN_PoserBoutonChangerCategorie(ws)
    If etaitMasquee Then ws.Visible = xlSheetVeryHidden

    Call MettreAJourInstructionsRapprochement

    MsgBox mod_Display.FR("Bouton 'Changer la cat{e2}gorie' et instructions mis {a2} jour sur frm_RapprochementNotes."), vbInformation

End Sub

' Pose le bouton "Passer" sous la 1re rangée de boutons (colonnes B:C).
' NOTE du 08/10/2026 : n'est plus utilisée que pour frm_GenerationCle ;
' frm_RapprochementNotes a désormais ses propres outils (RN_PoserBouton), car ses
' boutons ont été placés à la main par l'opérateur. Les 3 paramètres facultatifs
' ajoutés le 07/10/2026 pour le bouton "Changer la catégorie" ont donc été retirés.
Private Sub AjouterBoutonPasserSurFeuille(ByVal nomFeuille As String, ByVal ligneBoutons As Long, ByVal macroCible As String, ByVal nomBouton As String)

    Dim ws As Worksheet
    Dim etaitMasquee As Boolean
    Dim zoneBtn As Range
    Dim btn As Button

    Set ws = ObtenirFeuilleSansErreurFN(nomFeuille)
    If ws Is Nothing Then
        MsgBox "La feuille '" & nomFeuille & "' n'existe pas encore.", vbExclamation
        Exit Sub
    End If

    etaitMasquee = (ws.Visible <> xlSheetVisible)
    ws.Visible = xlSheetVisible

    On Error Resume Next
    ws.Buttons(nomBouton).Delete
    On Error GoTo 0

    Set zoneBtn = ws.Range(FN_COL_LIBELLE_1 & (ligneBoutons + 1) & ":" & FN_COL_VALEUR_1 & (ligneBoutons + 1))
    zoneBtn.RowHeight = 22
    Set btn = ws.Buttons.Add(zoneBtn.Left, zoneBtn.Top, zoneBtn.Width, zoneBtn.Height)
    With btn
        .Caption = "Passer"
        .OnAction = macroCible
        .Name = nomBouton
    End With

    If etaitMasquee Then ws.Visible = xlSheetVeryHidden

End Sub

' Réécrit uniquement le texte des instructions de frm_RapprochementNotes,
' avec une explication complète de chaque bouton et de ses conséquences.
' MODIFIÉ le 08/10/2026 : le texte et sa mise en forme viennent désormais de
' RN_EcrireInstructions / RN_TexteInstructions, partagées avec
' CreerFeuilleRapprochementNotes. Avant, le texte était écrit en double (ici et dans
' la macro de construction), et les deux versions avaient fini par diverger.
Private Sub MettreAJourInstructionsRapprochement()

    Dim ws As Worksheet
    Dim etaitMasquee As Boolean

    Set ws = ObtenirFeuilleSansErreurFN(NOM_FEUILLE_RAPPROCHEMENT)
    If ws Is Nothing Then Exit Sub

    etaitMasquee = (ws.Visible <> xlSheetVisible)
    ws.Visible = xlSheetVisible

    Call RN_EcrireInstructions(ws)

    If etaitMasquee Then ws.Visible = xlSheetVeryHidden

End Sub

' Réécrit uniquement le texte des instructions de frm_GenerationCle, avec une
' mise en garde explicite sur la portée définitive d'une clé générée.
Private Sub MettreAJourInstructionsGeneration()

    Dim ws As Worksheet
    Dim etaitMasquee As Boolean

    Set ws = ObtenirFeuilleSansErreurFN(NOM_FEUILLE_GENERATION)
    If ws Is Nothing Then Exit Sub

    etaitMasquee = (ws.Visible <> xlSheetVisible)
    ws.Visible = xlSheetVisible

    With ws.Range(FN_COL_LIBELLE_1 & GC_LIGNE_DEBUT_INSTRUCTIONS & ":" & FN_COL_VALEUR_3 & GC_LIGNE_FIN_INSTRUCTIONS)
        .value = mod_Display.FR("Aucune cl{e2} existante ne correspond {a2} cette op{e2}ration.") & Chr(10) & _
                 mod_Display.FR("Renseigne les 4 champs, clique 'G{e2}n{e2}rer' pour construire la cl{e2} (elle est aussi") & Chr(10) & _
                 mod_Display.FR("copi{e2}e dans le presse-papier), v{e2}rifie-la, puis clique 'Valider'.") & Chr(10) & _
                 mod_Display.FR("ATTENTION : une fois 'Valider' cliqu{e2}, la cl{e2} est {e2}crite DEFINITIVEMENT, m{ea}me") & Chr(10) & _
                 mod_Display.FR("approximative ou incorrecte : cette op{e2}ration ne sera PLUS JAMAIS repropos{e2}e pour") & Chr(10) & _
                 mod_Display.FR("correction automatique (seule une modification manuelle dans Import_data pourrait") & Chr(10) & _
                 mod_Display.FR("la faire r{e2}apparaitre). Si tu n'es pas certain des informations, clique plut{o2}t sur") & Chr(10) & _
                 mod_Display.FR("'Passer' pour laisser cette op{e2}ration de c{o2}t{e2} et y revenir plus tard.")
    End With

    ws.rows(GC_LIGNE_DEBUT_INSTRUCTIONS & ":" & GC_LIGNE_FIN_INSTRUCTIONS).RowHeight = 30

    If etaitMasquee Then ws.Visible = xlSheetVeryHidden

End Sub


