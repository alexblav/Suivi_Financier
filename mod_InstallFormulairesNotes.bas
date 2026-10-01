Option Explicit

' =====================================================================================
' MODULE : mod_InstallFormulairesNotes
'
' ROLE (PHASE 4a du chantier "Suivi Sante") :
'   Construit la mise en page STATIQUE (aucune logique de clic pour l'instant)
'   de 2 feuilles masquees utilisees quand le champ "Notes" d'une operation de
'   sante ne peut pas etre decoupe automatiquement (voir mod_ImportOFX,
'   fonction EstDateValide) :
'
'     1) frm_RapprochementNotes : propose une recherche par 4 filtres en
'        cascade (Date -> Specialite -> Beneficiaire -> Montant) parmi les
'        cles "Notes" deja valides ailleurs dans TblOperations, pour
'        retrouver la bonne cle sans avoir a la retaper.
'
'     2) frm_GenerationCle : si aucune correspondance n'existe, permet de
'        generer une nouvelle cle "AAAAMMJJ;Specialite;Beneficiaire;Montant"
'        a partir de 4 champs.
'
'   Comme pour frm_SuiviSante (mod_InstallSuiviSanteSheet), ce module ne fait
'   QUE poser les cellules, noms et boutons : aucun clic n'est encore
'   fonctionnel. Ce sera l'objet de la Phase 4b (mod_FormulairesNotes).
'
' CE MODULE EST VOLONTAIREMENT AUTONOME (il ne reutilise pas les fonctions
' internes de mod_InstallSuiviSanteSheet) : elles sont Private a leur module,
' et dupliquer ces quelques dizaines de lignes de mise en forme est plus sur
' que de multiplier les dependances entre fichiers -- see le mecanisme
' equivalent, deja duplique volontairement, dans mod_SuiviSanteFormulaire.
'
' A PROPOS DES ACCENTS : meme convention que tout le chantier Suivi Sante :
' fichier 100% ASCII, textes accentues construits via la fonction FR().
'
' A FAIRE POUR INSTALLER CE MODULE :
'   1. Alt+F11, Fichier > Importer un fichier..., choisir ce fichier .bas
'   2. Dans la fenetre Execution immediate (Ctrl+G) :
'        CreerFeuilleRapprochementNotes
'        CreerFeuilleGenerationCle
'   3. Pour revoir une feuille a l'ecran : AfficherFeuilleNotesPourEdition "frm_RapprochementNotes"
'      (ou "frm_GenerationCle"). Pour la remasquer : MasquerFeuilleNotesApresEdition "..."
' =====================================================================================


Public Const NOM_FEUILLE_RAPPROCHEMENT As String = "frm_RapprochementNotes"
Public Const NOM_FEUILLE_GENERATION As String = "frm_GenerationCle"

' Colonnes communes aux 2 feuilles (3 paires libelle/valeur par ligne, comme frm_SuiviSante)
Public Const FN_COL_LIBELLE_1 As String = "B"
Public Const FN_COL_VALEUR_1 As String = "C"
Public Const FN_COL_LIBELLE_2 As String = "D"
Public Const FN_COL_VALEUR_2 As String = "E"
Public Const FN_COL_LIBELLE_3 As String = "F"
Public Const FN_COL_VALEUR_3 As String = "G"
Public Const FN_COL_TECHNIQUE As String = "J"

' --- Mise en page de frm_RapprochementNotes ---
Public Const RN_LIGNE_BOUTONS As Long = 2
Public Const RN_LIGNE_TITRE_INSTRUCTIONS As Long = 5
Public Const RN_LIGNE_DEBUT_INSTRUCTIONS As Long = 6
Public Const RN_LIGNE_FIN_INSTRUCTIONS As Long = 9
Public Const RN_LIGNE_TITRE_OPERATION As Long = 11
Public Const RN_LIGNE_OPERATION As Long = 12
Public Const RN_LIGNE_TITRE_RECHERCHE As Long = 14
Public Const RN_LIGNE_DATE As Long = 15
Public Const RN_LIGNE_SPECIALITE As Long = 16
Public Const RN_LIGNE_BENEFICIAIRE As Long = 17
Public Const RN_LIGNE_MONTANT As Long = 18
Public Const RN_LIGNE_CLE_TROUVEE As Long = 20

' Colonnes techniques masquees : listes de candidats pour les 4 filtres en
' cascade, recalculees par la Phase 4b a chaque changement de filtre. Prevues
' ici (Phase 4a) uniquement pour que les noms definis existent des le depart.
Public Const RN_COL_LISTE_DATES As String = "L"
Public Const RN_COL_LISTE_SPECIALITES As String = "M"
Public Const RN_COL_LISTE_BENEFICIAIRES As String = "N"
Public Const RN_COL_LISTE_MONTANTS As String = "O"
Public Const RN_COL_LISTE_CLES As String = "P"   ' cle brute correspondant a chaque montant de la colonne O (meme ligne)

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
' MACRO D'INSTALLATION - frm_RapprochementNotes
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
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = NOM_FEUILLE_RAPPROCHEMENT
    End If

    Call FN_AppliquerMiseEnFormeGenerale(ws)

    ' --- Boutons + compteur ---
    Dim zoneBtn1 As Range, zoneBtn2 As Range
    Dim btn As Button

    Set zoneBtn1 = ws.Range(FN_COL_LIBELLE_1 & RN_LIGNE_BOUTONS & ":" & FN_COL_VALEUR_1 & RN_LIGNE_BOUTONS)
    zoneBtn1.RowHeight = 22
    Set btn = ws.Buttons.Add(zoneBtn1.Left, zoneBtn1.Top, zoneBtn1.Width, zoneBtn1.Height)
    With btn
        .Caption = mod_Display.FR("Pas de correspondance")
        .OnAction = "PasDeCorrespondanceNotes"
        .Name = "btnPasDeCorrespondance"
    End With

    Set zoneBtn2 = ws.Range(FN_COL_LIBELLE_2 & RN_LIGNE_BOUTONS & ":" & FN_COL_VALEUR_2 & RN_LIGNE_BOUTONS)
    Set btn = ws.Buttons.Add(zoneBtn2.Left, zoneBtn2.Top, zoneBtn2.Width, zoneBtn2.Height)
    With btn
        .Caption = "Valider"
        .OnAction = "ValiderRapprochementNotes"
        .Name = "btnValiderRapprochement"
    End With

    With ws.Range(FN_COL_LIBELLE_3 & RN_LIGNE_BOUTONS & ":" & FN_COL_VALEUR_3 & RN_LIGNE_BOUTONS)
        .Merge
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .Font.Size = 9
        .Font.Color = RGB(120, 120, 120)
    End With
    Call CreerNomSiAbsentFN(ws, "rnCompteurCas", ws.Range(FN_COL_LIBELLE_3 & RN_LIGNE_BOUTONS))

    ' --- Instructions ---
    With ws.Range(FN_COL_LIBELLE_1 & RN_LIGNE_TITRE_INSTRUCTIONS)
        .value = mod_Display.FR("Instructions")
        .Font.Bold = True
        .Font.Size = 11
    End With
    With ws.Range(FN_COL_LIBELLE_1 & RN_LIGNE_DEBUT_INSTRUCTIONS & ":" & FN_COL_VALEUR_3 & RN_LIGNE_FIN_INSTRUCTIONS)
        .Merge
        .value = mod_Display.FR("Le champ Notes de cette op{e2}ration de sant{e2} n'a pas pu {ea}tre d{e2}cod{e2}.") & Chr(10) & _
                 mod_Display.FR("Essaie de retrouver la bonne cl{e2} en filtrant ci-dessous (Date, puis Sp{e2}cialit{e2},") & Chr(10) & _
                 mod_Display.FR("puis B{e2}n{e2}ficiaire, puis Montant si besoin), puis clique sur 'Valider'.") & Chr(10) & _
                 mod_Display.FR("Si aucune correspondance ne convient, clique sur 'Pas de correspondance' pour en") & Chr(10) & _
                 mod_Display.FR("cr{e2}er une nouvelle.")
        .WrapText = True
        .VerticalAlignment = xlTop
        .Font.Size = 9
        .Font.Color = RGB(80, 80, 80)
        .Interior.Color = RGB(245, 245, 242)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(210, 210, 205)
        .Locked = True
    End With
    ws.rows(RN_LIGNE_DEBUT_INSTRUCTIONS & ":" & RN_LIGNE_FIN_INSTRUCTIONS).RowHeight = 15

    ' --- Bloc "Operation concernee" (lecture seule) ---
    With ws.Range(FN_COL_LIBELLE_1 & RN_LIGNE_TITRE_OPERATION & ":" & FN_COL_VALEUR_3 & RN_LIGNE_TITRE_OPERATION)
        .Merge
        .value = mod_Display.FR("Op{e2}ration concern{e2}e")
        .Font.Bold = True
        .Font.Size = 10.5
        .Interior.Color = RGB(250, 250, 248)
    End With
    ws.Range(FN_COL_LIBELLE_1 & RN_LIGNE_OPERATION).value = "Date :"
    ws.Range(FN_COL_LIBELLE_2 & RN_LIGNE_OPERATION).value = "Tiers :"
    ws.Range(FN_COL_LIBELLE_3 & RN_LIGNE_OPERATION).value = "Montant :"
    Call FN_MettreEnFormeLibelles(ws, RN_LIGNE_OPERATION)
    ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_OPERATION).NumberFormat = "dd/mm/yyyy"
    ws.Range(FN_COL_VALEUR_3 & RN_LIGNE_OPERATION).NumberFormat = "#,##0.00 " & ChrW(8364)
    Call FN_MettreEnFormeValeursLectureSeule(ws, RN_LIGNE_OPERATION)
    Call CreerNomSiAbsentFN(ws, "rnDateOp", ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_OPERATION))
    Call CreerNomSiAbsentFN(ws, "rnTiersOp", ws.Range(FN_COL_VALEUR_2 & RN_LIGNE_OPERATION))
    Call CreerNomSiAbsentFN(ws, "rnMontantOp", ws.Range(FN_COL_VALEUR_3 & RN_LIGNE_OPERATION))

    ' --- Bloc "Recherche par filtres en cascade" ---
    With ws.Range(FN_COL_LIBELLE_1 & RN_LIGNE_TITRE_RECHERCHE & ":" & FN_COL_VALEUR_3 & RN_LIGNE_TITRE_RECHERCHE)
        .Merge
        .value = mod_Display.FR("Recherche de la cl{e2} (filtres en cascade)")
        .Font.Bold = True
        .Font.Size = 10.5
        .Interior.Color = RGB(250, 250, 248)
    End With

    ws.Range(FN_COL_LIBELLE_1 & RN_LIGNE_DATE).value = "Date :"
    Call FN_MettreEnFormeLibelles(ws, RN_LIGNE_DATE)
    Dim rngDate As Range
    Set rngDate = ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_DATE & ":" & FN_COL_VALEUR_2 & RN_LIGNE_DATE)
    rngDate.Merge
    Call FN_MettreEnFormeZoneSaisie(rngDate)
    Call CreerNomSiAbsentFN(ws, "rnDate", ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_DATE))

    ws.Range(FN_COL_LIBELLE_1 & RN_LIGNE_SPECIALITE).value = mod_Display.FR("Sp{e2}cialit{e2} :")
    Call FN_MettreEnFormeLibelles(ws, RN_LIGNE_SPECIALITE)
    Dim rngSpecialite As Range
    Set rngSpecialite = ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_SPECIALITE & ":" & FN_COL_VALEUR_2 & RN_LIGNE_SPECIALITE)
    rngSpecialite.Merge
    Call FN_MettreEnFormeZoneSaisie(rngSpecialite)
    Call CreerNomSiAbsentFN(ws, "rnSpecialite", ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_SPECIALITE))

    ws.Range(FN_COL_LIBELLE_1 & RN_LIGNE_BENEFICIAIRE).value = mod_Display.FR("B{e2}n{e2}ficiaire :")
    Call FN_MettreEnFormeLibelles(ws, RN_LIGNE_BENEFICIAIRE)
    Dim rngBeneficiaire As Range
    Set rngBeneficiaire = ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_BENEFICIAIRE & ":" & FN_COL_VALEUR_2 & RN_LIGNE_BENEFICIAIRE)
    rngBeneficiaire.Merge
    Call FN_MettreEnFormeZoneSaisie(rngBeneficiaire)
    Call CreerNomSiAbsentFN(ws, "rnBeneficiaire", ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_BENEFICIAIRE))

    ws.Range(FN_COL_LIBELLE_1 & RN_LIGNE_MONTANT).value = "Montant :"
    Call FN_MettreEnFormeLibelles(ws, RN_LIGNE_MONTANT)
    Dim rngMontant As Range
    Set rngMontant = ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_MONTANT & ":" & FN_COL_VALEUR_2 & RN_LIGNE_MONTANT)
    rngMontant.Merge
    Call FN_MettreEnFormeZoneSaisie(rngMontant)
    Call CreerNomSiAbsentFN(ws, "rnMontant", ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_MONTANT))

    ' --- Cle trouvee (lecture seule) ---
    ws.Range(FN_COL_LIBELLE_1 & RN_LIGNE_CLE_TROUVEE).value = mod_Display.FR("Cl{e2} trouv{e2}e :")
    Call FN_MettreEnFormeLibelles(ws, RN_LIGNE_CLE_TROUVEE)
    With ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_CLE_TROUVEE & ":" & FN_COL_VALEUR_3 & RN_LIGNE_CLE_TROUVEE)
        .Merge
        .Font.Bold = True
    End With
    Call FN_MettreEnFormeValeursLectureSeule(ws, RN_LIGNE_CLE_TROUVEE)
    Call CreerNomSiAbsentFN(ws, "rnCleTrouvee", ws.Range(FN_COL_VALEUR_1 & RN_LIGNE_CLE_TROUVEE))

    ' --- Zone technique masquee ---
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

    ' Chaque liste demarre avec UNE cellule de reserve (ligne 2), que la
    ' Phase 4b redimensionnera dynamiquement, comme deja fait pour les
    ' plages Beneficiaires/Praticiens (voir AjouterValeurDansListe).
    Call CreerNomSiAbsentFN(ws, "rnListeDates", ws.Range(RN_COL_LISTE_DATES & "2"))
    Call CreerNomSiAbsentFN(ws, "rnListeSpecialites", ws.Range(RN_COL_LISTE_SPECIALITES & "2"))
    Call CreerNomSiAbsentFN(ws, "rnListeBeneficiaires", ws.Range(RN_COL_LISTE_BENEFICIAIRES & "2"))
    Call CreerNomSiAbsentFN(ws, "rnListeMontants", ws.Range(RN_COL_LISTE_MONTANTS & "2"))
    Call CreerNomSiAbsentFN(ws, "rnListeCles", ws.Range(RN_COL_LISTE_CLES & "2"))

    ws.Range(FN_COL_TECHNIQUE & RN_LIGNE_BOUTONS).value = 0
    Call CreerNomSiAbsentFN(ws, "rnLigneEnCours", ws.Range(FN_COL_TECHNIQUE & RN_LIGNE_BOUTONS))
    ws.Columns(FN_COL_TECHNIQUE).Hidden = True

    ws.Visible = xlSheetVeryHidden

    MsgBox "La feuille '" & NOM_FEUILLE_RAPPROCHEMENT & "' a ete creee et masquee." & vbCrLf & _
           "Pour la revoir : AfficherFeuilleNotesPourEdition " & Chr(34) & NOM_FEUILLE_RAPPROCHEMENT & Chr(34), _
           vbInformation, "Installation terminee"

End Sub


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

    ' --- Bloc "Operation concernee" (lecture seule) ---
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

    ' --- Bloc "Generation de la cle" ---
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

    ' --- Cle generee (lecture seule) ---
    ws.Range(FN_COL_LIBELLE_1 & GC_LIGNE_CLE_GENEREE).value = mod_Display.FR("Cl{e2} g{e2}n{e2}r{e2}e :")
    Call FN_MettreEnFormeLibelles(ws, GC_LIGNE_CLE_GENEREE)
    With ws.Range(FN_COL_VALEUR_1 & GC_LIGNE_CLE_GENEREE & ":" & FN_COL_VALEUR_3 & GC_LIGNE_CLE_GENEREE)
        .Merge
        .Font.Bold = True
    End With
    Call FN_MettreEnFormeValeursLectureSeule(ws, GC_LIGNE_CLE_GENEREE)
    Call CreerNomSiAbsentFN(ws, "gcCleGeneree", ws.Range(FN_COL_VALEUR_1 & GC_LIGNE_CLE_GENEREE))

    ' --- Zone technique masquee ---
    ws.Range(FN_COL_TECHNIQUE & GC_LIGNE_BOUTONS).value = 0
    Call CreerNomSiAbsentFN(ws, "gcLigneEnCours", ws.Range(FN_COL_TECHNIQUE & GC_LIGNE_BOUTONS))
    ws.Columns(FN_COL_TECHNIQUE).Hidden = True

    ws.Visible = xlSheetVeryHidden

    MsgBox "La feuille '" & NOM_FEUILLE_GENERATION & "' a ete creee et masquee." & vbCrLf & _
           "Pour la revoir : AfficherFeuilleNotesPourEdition " & Chr(34) & NOM_FEUILLE_GENERATION & Chr(34), _
           vbInformation, "Installation terminee"

End Sub


' =====================================================================================
' MISE EN FORME GENERALE (commune aux 2 feuilles)
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
' UTILITAIRES DE MISE EN FORME (communs aux 2 feuilles, meme esprit que les
' fonctions SS_xxx de mod_InstallSuiviSanteSheet)
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
' FONCTIONS UTILITAIRES (creation/suppression feuille, noms, formes)
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
' OUTILS DEVELOPPEUR : basculent la visibilite d'une des 2 feuilles pour
' pouvoir la retoucher. Passe le nom exact de la feuille en parametre.
' =====================================================================================
' =====================================================================================
' AjouterBoutonPasserEtInstructions : ajoute UNIQUEMENT le bouton "Passer" sur
' les 2 feuilles (ligne 3, sous les boutons existants) et met a jour le texte
' des instructions (plus complet sur les consequences de chaque action), SANS
' reconstruire le reste des feuilles. A executer UNE SEULE FOIS, dans la
' fenetre Execution immediate (Ctrl+G) :
'      AjouterBoutonPasserEtInstructions
' =====================================================================================
Sub AjouterBoutonPasserEtInstructions()

    Call AjouterBoutonPasserSurFeuille(NOM_FEUILLE_RAPPROCHEMENT, RN_LIGNE_BOUTONS, "PasserRapprochementNotes", "btnPasserRapprochement")
    Call MettreAJourInstructionsRapprochement

    Call AjouterBoutonPasserSurFeuille(NOM_FEUILLE_GENERATION, GC_LIGNE_BOUTONS, "PasserGenerationCle", "btnPasserGeneration")
    Call MettreAJourInstructionsGeneration

    MsgBox "Bouton 'Passer' et instructions mis a jour sur les 2 feuilles.", vbInformation

End Sub

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

' Reecrit uniquement le texte des instructions de frm_RapprochementNotes,
' avec une explication complete de chaque bouton et de sa consequence.
Private Sub MettreAJourInstructionsRapprochement()

    Dim ws As Worksheet
    Dim etaitMasquee As Boolean

    Set ws = ObtenirFeuilleSansErreurFN(NOM_FEUILLE_RAPPROCHEMENT)
    If ws Is Nothing Then Exit Sub

    etaitMasquee = (ws.Visible <> xlSheetVisible)
    ws.Visible = xlSheetVisible

    With ws.Range(FN_COL_LIBELLE_1 & RN_LIGNE_DEBUT_INSTRUCTIONS & ":" & FN_COL_VALEUR_3 & RN_LIGNE_FIN_INSTRUCTIONS)
        .value = mod_Display.FR("Cette op{e2}ration de sant{e2} n'a pas pu {ea}tre rapproch{e2}e automatiquement.") & Chr(10) & _
                 mod_Display.FR("Filtre par Date (les plus r{e2}centes en premier), puis Sp{e2}cialit{e2}, puis B{e2}n{e2}ficiaire,") & Chr(10) & _
                 mod_Display.FR("puis Montant si plusieurs choix restent possibles, pour retrouver la cl{e2} parmi") & Chr(10) & _
                 mod_Display.FR("celles d{e2}j{a2} connues. Cons{e2}quence de chaque bouton :") & Chr(10) & _
                 mod_Display.FR("- 'Valider' : applique la cl{e2} trouv{e2}e {a2} cette op{e2}ration. Elle ne sera PLUS JAMAIS") & Chr(10) & _
                 mod_Display.FR("  reproposee (rapprochement considere regle definitivement).") & Chr(10) & _
                 mod_Display.FR("- 'Pas de correspondance' : ouvre un {e2}cran pour cr{e2}er une toute nouvelle cl{e2}") & Chr(10) & _
                 mod_Display.FR("  (aucune cl{e2} existante ne convient).") & Chr(10) & _
                 mod_Display.FR("- 'Passer' : ne modifie RIEN sur cette op{e2}ration. Elle sera automatiquement") & Chr(10) & _
                 mod_Display.FR("  repropos{e2}e la prochaine fois que tu relanceras la v{e2}rification")
    End With

    ws.rows(RN_LIGNE_DEBUT_INSTRUCTIONS & ":" & RN_LIGNE_FIN_INSTRUCTIONS).RowHeight = 30

    If etaitMasquee Then ws.Visible = xlSheetVeryHidden

End Sub

' Reecrit uniquement le texte des instructions de frm_GenerationCle, avec une
' mise en garde explicite sur la portee definitive d'une cle generee.
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


