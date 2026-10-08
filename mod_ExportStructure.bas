Option Explicit

' =====================================================================================
' MODULE : mod_ExportStructure
' Création : 08/10/2026 (demande opérateur)
'
' RÔLE
'   Écrire dans un fichier texte (UTF-8, fins de ligne LF) la STRUCTURE d'une feuille
'   Excel telle qu'elle existe réellement dans le classeur : largeurs de colonnes,
'   hauteurs de lignes, cellules fusionnées, textes et mises en forme de la zone de
'   présentation, tableaux (ListObjects), listes déroulantes, mises en forme
'   conditionnelles, boutons et formes, noms définis, réglages de la fenêtre
'   (volets figés, quadrillage, zoom...).
'
' POURQUOI
'   Les feuilles-formulaires (frm_RechercheOperations, frm_ControleCategories...)
'   sont construites par du code (modules mod_Install*), puis souvent retouchées à la
'   main par l'opérateur (largeurs, hauteurs, positions...). Ce module permet de
'   « photographier » la feuille réelle, de pousser le fichier obtenu sur GitHub, et
'   ainsi de remettre le code de construction en accord avec la réalité, sans rien
'   deviner.
'
' CE QUE CE MODULE NE FAIT JAMAIS
'   - Il ne MODIFIE rien dans le classeur : il ne fait que LIRE. Seule exception,
'     temporaire et restaurée : une feuille masquée est rendue visible un instant
'     pour lire les réglages de sa fenêtre (voir LireFenetre).
'   - Il n'exporte PAS le CONTENU des tableaux (lignes de données des ListObjects) :
'     seule leur structure (colonnes, formats, listes) est décrite. C'est voulu : le
'     dépôt GitHub est PUBLIC, et vos opérations bancaires ne doivent jamais s'y
'     retrouver. ATTENTION : une feuille de données qui n'est PAS un tableau (des
'     valeurs saisies directement dans des cellules) verrait, elle, ses valeurs
'     exportées dans la limite de EXPORT_NB_LIGNES_MAX lignes. Relisez le fichier
'     avant de le pousser si vous exportez une telle feuille.
'
' COMMENT L'UTILISER
'   - Alt+F8 > ExporterUneFeuille        : demande quelle feuille exporter.
'   - Alt+F8 > ExporterToutesLesFeuilles : un fichier par feuille du classeur.
'   - Ctrl+G (fenêtre Exécution) :
'         ExporterStructureFeuille "frm_RechercheOperations"
'   Les fichiers sont écrits dans un sous-dossier EXPORT_SOUS_DOSSIER, à côté du
'   classeur (voir DossierExport pour le cas OneDrive). Le chemin est rappelé en fin
'   de traitement.
'
' CONSTANTES UTILISÉES (déclarées dans mod_VarGlobales, règle du projet) :
'   EXPORT_SOUS_DOSSIER, EXPORT_NB_LIGNES_MAX, EXPORT_NB_CELLULES_MAX
'
' LECTURE DU FICHIER PRODUIT
'   - Une section par thème, titre précédé de "##".
'   - Les nombres sont écrits avec un POINT décimal (ex. 8.43), directement
'     réutilisables dans du code VBA, quel que soit le réglage régional du poste.
'   - Les couleurs sont écrites sous la forme RGB(r, g, b), réutilisable telle quelle.
'   - Un retour à la ligne à l'intérieur d'une cellule est écrit "\n".
' =====================================================================================


' =====================================================================================
' POINTS D'ENTRÉE (visibles dans Alt+F8 : ils n'ont pas de paramètre)
' =====================================================================================

' Demande à l'opérateur le numéro ou le nom de la feuille, puis l'exporte.
Public Sub ExporterUneFeuille()

    Dim ws As Worksheet
    Dim liste As String
    Dim choix As String
    Dim numero As Long

    ' Liste numérotée des feuilles, pour que l'opérateur n'ait pas à retaper un nom
    ' (les noms des feuilles-formulaires sont longs et faciles à mal orthographier).
    numero = 0
    For Each ws In ThisWorkbook.Worksheets
        numero = numero + 1
        liste = liste & numero & " - " & ws.Name & vbLf
    Next ws

    ' Une InputBox n'affiche qu'environ 1 000 caractères : au-delà, la liste complète
    ' est écrite dans la fenêtre Exécution (Ctrl+G) et on en montre le début.
    Debug.Print liste
    If Len(liste) > 800 Then
        liste = Left$(liste, 800) & "..." & vbLf & mod_Display.FR("(liste compl{e1}te : fen{ea}tre Ex{e2}cution, Ctrl+G)") & vbLf
    End If

    choix = InputBox(mod_Display.FR("Num{e2}ro ou nom de la feuille {a2} exporter :") & vbLf & vbLf & liste, _
                     mod_Display.FR("Export de la structure d'une feuille"), ActiveSheet.Name)
    choix = Trim$(choix)
    If choix = "" Then Exit Sub   ' bouton Annuler, ou rien saisi

    ' Un nombre est compris comme un NUMÉRO dans la liste, sinon comme un NOM.
    Set ws = Nothing
    On Error Resume Next
    If IsNumeric(choix) Then
        Set ws = ThisWorkbook.Worksheets(CLng(choix))
    Else
        Set ws = ThisWorkbook.Worksheets(choix)
    End If
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox mod_Display.FR("Feuille introuvable : ") & choix, vbExclamation
        Exit Sub
    End If

    ExporterStructureFeuille ws.Name

End Sub

' Exporte toutes les feuilles du classeur, un fichier par feuille.
Public Sub ExporterToutesLesFeuilles()

    Dim ws As Worksheet
    Dim dossier As String
    Dim nb As Long

    dossier = DossierExport()
    If dossier = "" Then Exit Sub

    nb = 0
    For Each ws In ThisWorkbook.Worksheets
        If EcrireStructureFeuille(ws, dossier, False) <> "" Then nb = nb + 1
    Next ws

    MsgBox nb & mod_Display.FR(" feuille(s) export{e2}e(s) dans :") & vbLf & dossier, vbInformation, _
           mod_Display.FR("Export de structure")

End Sub


' =====================================================================================
' POINT D'ENTRÉE AVEC PARAMÈTRE (utilisable depuis Ctrl+G ou depuis un autre module)
' =====================================================================================
Public Sub ExporterStructureFeuille(ByVal nomFeuille As String)

    Dim ws As Worksheet
    Dim dossier As String
    Dim chemin As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nomFeuille)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox mod_Display.FR("Feuille introuvable : ") & nomFeuille, vbExclamation
        Exit Sub
    End If

    dossier = DossierExport()
    If dossier = "" Then Exit Sub

    chemin = EcrireStructureFeuille(ws, dossier, True)
    If chemin <> "" Then
        MsgBox mod_Display.FR("Structure export{e2}e dans :") & vbLf & chemin, vbInformation, _
               mod_Display.FR("Export de structure")
    End If

End Sub


' =====================================================================================
' TRAITEMENT D'UNE FEUILLE : construit le texte puis l'écrit sur le disque.
' Renvoie le chemin du fichier écrit, ou "" en cas d'échec.
' afficherErreur = False en export groupé : une feuille en échec ne doit pas
' interrompre les suivantes par une boîte de dialogue (le motif est écrit dans la
' fenêtre Exécution, Ctrl+G).
' =====================================================================================
Private Function EcrireStructureFeuille(ByVal ws As Worksheet, ByVal dossier As String, _
                                        ByVal afficherErreur As Boolean) As String

    Dim contenu As String
    Dim chemin As String
    Dim ecranAvant As Boolean

    ecranAvant = Application.ScreenUpdating
    Application.ScreenUpdating = False

    On Error GoTo Erreur

    contenu = ConstruireStructure(ws)
    chemin = dossier & "\" & NomFichierSur(ws.Name) & "_structure.txt"
    EcrireFichierUTF8 chemin, contenu

    Application.ScreenUpdating = ecranAvant
    EcrireStructureFeuille = chemin
    Exit Function

Erreur:
    ' On remet TOUJOURS l'affichage dans son état d'origine, même en cas d'erreur
    ' (leçon de l'ancien Synthese_Care, qui laissait un réglage d'Excel désactivé).
    Application.ScreenUpdating = ecranAvant
    Debug.Print "Export '" & ws.Name & "' : erreur " & Err.Number & " - " & Err.Description
    If afficherErreur Then
        MsgBox mod_Display.FR("Export impossible pour la feuille '") & ws.Name & "' :" & vbLf & _
               Err.Description, vbExclamation
    End If
    EcrireStructureFeuille = ""

End Function


' =====================================================================================
' CONSTRUCTION DU TEXTE : une section par thème. Chaque sous-procédure ajoute ses
' lignes au texte "t", passé par référence (ByRef) : c'est le même texte qui
' s'allonge au fil des sections.
' =====================================================================================
Private Function ConstruireStructure(ByVal ws As Worksheet) As String

    Dim t As String
    Dim zone As Range            ' zone de présentation analysée en détail
    Dim plageTableaux As Range   ' réunion de tous les tableaux (exclus du détail des cellules)

    Set zone = ZoneAnalysee(ws)
    Set plageTableaux = ReunionTableaux(ws)

    SectionEntete ws, zone, t
    LireFenetre ws, t
    SectionColonnes ws, zone, t
    SectionLignes ws, zone, t
    SectionCellules ws, zone, plageTableaux, t
    SectionTableaux ws, t
    SectionValidations ws, plageTableaux, t
    SectionMisesEnFormeConditionnelles ws, t
    SectionFormes ws, t
    SectionNoms ws, t

    Ajouter t, ""
    Ajouter t, "# FIN DE L'EXPORT"

    ConstruireStructure = t

End Function


' -------------------------------------------------------------------------------------
' Zone analysée : de A1 jusqu'à la dernière colonne utilisée (+2 colonnes de marge,
' pour voir aussi les largeurs juste après le contenu) et jusqu'à la dernière ligne
' utilisée, plafonnée à EXPORT_NB_LIGNES_MAX (les feuilles-formulaires ont leur mise
' en page en haut ; au-delà, ce sont des données).
' -------------------------------------------------------------------------------------
Private Function ZoneAnalysee(ByVal ws As Worksheet) As Range

    Dim derniereLigne As Long
    Dim derniereColonne As Long

    With ws.UsedRange
        derniereLigne = .Row + .rows.count - 1
        derniereColonne = .Column + .Columns.count - 1 + 2
    End With

    If derniereLigne > EXPORT_NB_LIGNES_MAX Then derniereLigne = EXPORT_NB_LIGNES_MAX
    If derniereLigne < 1 Then derniereLigne = 1
    If derniereColonne > ws.Columns.count Then derniereColonne = ws.Columns.count

    Set ZoneAnalysee = ws.Range(ws.Cells(1, 1), ws.Cells(derniereLigne, derniereColonne))

End Function

' -------------------------------------------------------------------------------------
' Réunion (Union) des plages de tous les tableaux de la feuille, ou Nothing s'il n'y
' en a aucun. Sert à exclure leurs cellules du détail « cellule par cellule » :
' leur structure est décrite à part (SectionTableaux), et leur contenu jamais.
' -------------------------------------------------------------------------------------
Private Function ReunionTableaux(ByVal ws As Worksheet) As Range

    Dim lo As ListObject
    Dim resultat As Range

    For Each lo In ws.ListObjects
        If resultat Is Nothing Then
            Set resultat = lo.Range
        Else
            Set resultat = Union(resultat, lo.Range)
        End If
    Next lo

    Set ReunionTableaux = resultat

End Function


' =====================================================================================
' SECTION 1 : en-tête (identité de la feuille et réglages généraux)
' =====================================================================================
Private Sub SectionEntete(ByVal ws As Worksheet, ByVal zone As Range, ByRef t As String)

    Ajouter t, "# STRUCTURE DE LA FEUILLE : " & ws.Name
    Ajouter t, "Classeur           = " & ThisWorkbook.Name
    Ajouter t, "CodeName           = " & ws.CodeName
    Ajouter t, "Visibilite         = " & NomVisibilite(ws.Visible)
    Ajouter t, "Exporte le         = " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Ajouter t, "Version Excel      = " & Application.Version
    Ajouter t, "Plage utilisee     = " & ws.UsedRange.Address(False, False)
    Ajouter t, "Zone detaillee     = " & zone.Address(False, False) & _
               "  (lignes plafonnees a EXPORT_NB_LIGNES_MAX = " & EXPORT_NB_LIGNES_MAX & ")"
    Ajouter t, "Largeur standard   = " & NombreVBA(ws.StandardWidth)
    Ajouter t, "Hauteur standard   = " & NombreVBA(ws.StandardHeight)
    Ajouter t, "Protection contenu = " & OuiNon(ws.ProtectContents)
    Ajouter t, "Nb tableaux        = " & ws.ListObjects.count
    Ajouter t, "Nb formes          = " & ws.Shapes.count
    Ajouter t, "Contenu des tableaux NON exporte (confidentialite)."

End Sub


' =====================================================================================
' SECTION 2 : réglages de la FENÊTRE (volets figés, quadrillage, zoom...).
' Ces réglages appartiennent à la fenêtre et non à la feuille : on ne peut les lire
' que sur la feuille ACTIVE. On active donc la feuille un instant, puis on remet
' tout exactement dans l'état d'origine (feuille active, visibilité, événements).
' Application.EnableEvents est coupé pendant cette opération pour ne déclencher
' AUCUN code de feuille (Worksheet_Activate / Worksheet_Deactivate des
' feuilles-formulaires, qui pilotent les boucles modales du projet).
' =====================================================================================
Private Sub LireFenetre(ByVal ws As Worksheet, ByRef t As String)

    Dim feuilleAvant As Object
    Dim visibiliteAvant As XlSheetVisibility
    Dim evenementsAvant As Boolean
    Dim messageErreur As String

    Ajouter t, ""
    Ajouter t, "## FENETRE"

    Set feuilleAvant = ActiveSheet
    visibiliteAvant = ws.Visible
    evenementsAvant = Application.EnableEvents

    On Error GoTo Erreur
    Application.EnableEvents = False
    ThisWorkbook.Activate
    If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
    ws.Activate

    With ActiveWindow
        Ajouter t, "Zoom             = " & .Zoom
        Ajouter t, "FreezePanes      = " & OuiNon(.FreezePanes)
        Ajouter t, "SplitRow         = " & .SplitRow & "   (volets figes sous la ligne " & .SplitRow & ")"
        Ajouter t, "SplitColumn      = " & .SplitColumn
        Ajouter t, "ScrollRow        = " & .ScrollRow
        Ajouter t, "ScrollColumn     = " & .ScrollColumn
        Ajouter t, "DisplayGridlines = " & OuiNon(.DisplayGridlines)
        Ajouter t, "DisplayHeadings  = " & OuiNon(.DisplayHeadings)
        Ajouter t, "DisplayZeros     = " & OuiNon(.DisplayZeros)
    End With
    GoTo Restauration

Erreur:
    messageErreur = Err.Description
    ' "Resume" termine proprement la gestion d'erreur avant la restauration
    ' (sans lui, une 2e erreur dans la restauration ne serait plus interceptée).
    Resume Restauration

Restauration:
    On Error Resume Next
    If messageErreur <> "" Then Ajouter t, "(lecture de la fenetre impossible : " & messageErreur & ")"
    If Not feuilleAvant Is Nothing Then feuilleAvant.Activate
    ws.Visible = visibiliteAvant
    Application.EnableEvents = evenementsAvant
    On Error GoTo 0

End Sub


' =====================================================================================
' SECTION 3 : colonnes (largeur et masquage), de A jusqu'au bout de la zone
' =====================================================================================
Private Sub SectionColonnes(ByVal ws As Worksheet, ByVal zone As Range, ByRef t As String)

    Dim c As Long
    Dim ligne As String

    Ajouter t, ""
    Ajouter t, "## COLONNES (lettre | largeur | masquee)"

    For c = 1 To zone.Columns.count
        ligne = LettreColonne(ws, c) & " | " & NombreVBA(ws.Columns(c).ColumnWidth)
        If ws.Columns(c).Hidden Then ligne = ligne & " | MASQUEE"
        Ajouter t, ligne
    Next c

End Sub


' =====================================================================================
' SECTION 4 : lignes. Seules les lignes dont la hauteur DIFFÈRE de la hauteur
' standard, ou qui sont masquées, sont listées (sinon le fichier serait illisible).
' =====================================================================================
Private Sub SectionLignes(ByVal ws As Worksheet, ByVal zone As Range, ByRef t As String)

    Dim r As Long
    Dim ligne As String
    Dim nb As Long

    Ajouter t, ""
    Ajouter t, "## LIGNES NON STANDARD (numero | hauteur | masquee) -- hauteur standard = " & NombreVBA(ws.StandardHeight)

    nb = 0
    For r = 1 To zone.rows.count
        If ws.rows(r).Hidden Or Abs(ws.rows(r).RowHeight - ws.StandardHeight) > 0.01 Then
            ligne = r & " | " & NombreVBA(ws.rows(r).RowHeight)
            If ws.rows(r).Hidden Then ligne = ligne & " | MASQUEE"
            Ajouter t, ligne
            nb = nb + 1
        End If
    Next r
    If nb = 0 Then Ajouter t, "(aucune)"

End Sub


' =====================================================================================
' SECTION 5 : cellules de la zone de présentation (hors tableaux).
' Une cellule est listée si elle a un contenu, une formule, un fond coloré, du gras,
' une bordure, ou si elle est le coin haut-gauche d'une fusion. Pour une plage
' fusionnée, seule la 1re cellule est décrite (les autres sont vides par nature).
' =====================================================================================
Private Sub SectionCellules(ByVal ws As Worksheet, ByVal zone As Range, ByVal plageTableaux As Range, _
                            ByRef t As String)

    Dim c As Range
    Dim nb As Long
    Dim policeNormale As String, taillenormale As Double

    ' Police du style "Normal" du classeur : on ne signale une police que si elle
    ' en diffère (sinon chaque ligne répéterait "Calibri 11").
    policeNormale = "Calibri": taillenormale = 11
    On Error Resume Next
    policeNormale = ThisWorkbook.Styles("Normal").Font.Name
    taillenormale = ThisWorkbook.Styles("Normal").Font.Size
    On Error GoTo 0

    Ajouter t, ""
    Ajouter t, "## CELLULES DE LA ZONE (hors tableaux)"
    Ajouter t, "## adresse | valeur ou formule | fusion | police | fond | format | alignement | autres"

    nb = 0
    For Each c In zone.Cells
        If Not DansPlage(c, plageTableaux) Then
            If CelluleADecrire(c) Then
                Ajouter t, DecrireCellule(c, policeNormale, taillenormale)
                nb = nb + 1
                If nb >= EXPORT_NB_CELLULES_MAX Then
                    Ajouter t, "(arret : EXPORT_NB_CELLULES_MAX = " & EXPORT_NB_CELLULES_MAX & " cellules atteint)"
                    Exit For
                End If
            End If
        End If
    Next c
    If nb = 0 Then Ajouter t, "(aucune)"

End Sub

' Indique si une cellule mérite d'être décrite (voir l'en-tête de SectionCellules).
Private Function CelluleADecrire(ByVal c As Range) As Boolean

    ' Cellule fusionnée mais PAS la première de sa fusion : rien à dire.
    If c.MergeCells Then
        If c.Address <> c.MergeArea.Cells(1, 1).Address Then Exit Function
        CelluleADecrire = True
        Exit Function
    End If

    If c.HasFormula Then CelluleADecrire = True: Exit Function
    If Not IsEmpty(c.Value2) Then CelluleADecrire = True: Exit Function
    If c.Interior.ColorIndex <> xlColorIndexNone Then CelluleADecrire = True: Exit Function
    If ValeurVraie(c.Font.Bold) Then CelluleADecrire = True: Exit Function
    If ABordure(c) Then CelluleADecrire = True: Exit Function

End Function

' Décrit une cellule sur UNE ligne de texte.
Private Function DecrireCellule(ByVal c As Range, ByVal policeNormale As String, _
                                ByVal taillenormale As Double) As String

    Dim s As String
    Dim police As String
    Dim valeurTexte As String
    Dim alignH As String, alignV As String

    s = c.Address(False, False)

    ' --- Valeur ou formule (la formule est donnée en anglais, comme dans le code VBA)
    If c.HasFormula Then
        valeurTexte = c.Formula
    ElseIf IsError(c.Value2) Then
        valeurTexte = "#ERREUR"
    Else
        valeurTexte = CStr(c.Value2)
    End If
    s = s & " | """ & Echapper(valeurTexte) & """"

    ' --- Fusion
    If c.MergeCells Then s = s & " | fusion=" & c.MergeArea.Address(False, False)

    ' --- Police : uniquement ce qui diffère du style Normal. Une valeur Null signifie
    '     une mise en forme PARTIELLE du texte (technique Characters, cf. SurlignerMotsRO).
    police = ""
    If IsNull(c.Font.Name) Then
        police = police & " nom=mixte"
    ElseIf c.Font.Name <> policeNormale Then
        police = police & " nom=" & c.Font.Name
    End If
    If IsNull(c.Font.Size) Then
        police = police & " taille=mixte"
    ElseIf c.Font.Size <> taillenormale Then
        police = police & " taille=" & NombreVBA(c.Font.Size)
    End If
    If IsNull(c.Font.Bold) Then
        police = police & " gras=mixte"
    ElseIf c.Font.Bold Then
        police = police & " gras"
    End If
    If ValeurVraie(c.Font.Italic) Then police = police & " italique"
    If IsNull(c.Font.Color) Then
        police = police & " couleur=mixte (mise en forme partielle)"
    ElseIf c.Font.Color <> 0 Then
        police = police & " couleur=" & CouleurVBA(c.Font.Color)
    End If
    If police <> "" Then s = s & " | police:" & police

    ' --- Fond
    If c.Interior.ColorIndex <> xlColorIndexNone Then s = s & " | fond=" & CouleurVBA(c.Interior.Color)

    ' --- Format de nombre (seulement s'il n'est pas "Standard")
    If c.NumberFormat <> "General" Then s = s & " | format=""" & c.NumberFormat & """"

    ' --- Alignement (seulement s'il n'est pas celui par défaut)
    alignH = NomAlignementH(c.HorizontalAlignment)
    alignV = NomAlignementV(c.VerticalAlignment)
    If alignH <> "" Then s = s & " | alignH=" & alignH
    If alignV <> "" Then s = s & " | alignV=" & alignV

    ' --- Autres réglages utiles à la reconstruction
    If ValeurVraie(c.WrapText) Then s = s & " | retour_ligne"
    If ABordure(c) Then s = s & " | bordures"
    If Not ValeurVraie(c.Locked) Then s = s & " | non_verrouillee"

    DecrireCellule = s

End Function


' =====================================================================================
' SECTION 6 : tableaux (ListObjects) -- STRUCTURE uniquement, jamais les données.
' Pour chaque colonne : position, lettre, largeur, masquage, format et liste
' déroulante lus sur la 1re ligne de données (quand le tableau en a une).
' =====================================================================================
Private Sub SectionTableaux(ByVal ws As Worksheet, ByRef t As String)

    Dim lo As ListObject
    Dim lc As ListColumn
    Dim nbLignes As Long
    Dim styleTableau As String
    Dim ligne As String
    Dim colonneFeuille As Long

    Ajouter t, ""
    Ajouter t, "## TABLEAUX"
    If ws.ListObjects.count = 0 Then
        Ajouter t, "(aucun)"
        Exit Sub
    End If

    For Each lo In ws.ListObjects

        nbLignes = 0
        If Not lo.DataBodyRange Is Nothing Then nbLignes = lo.DataBodyRange.rows.count

        ' Le style peut être un objet TableStyle ou un simple texte selon les versions.
        styleTableau = ""
        On Error Resume Next
        styleTableau = lo.TableStyle.Name
        If styleTableau = "" Then styleTableau = CStr(lo.TableStyle)
        On Error GoTo 0

        Ajouter t, ""
        Ajouter t, "### Tableau " & lo.Name
        Ajouter t, "Plage            = " & lo.Range.Address(False, False)
        Ajouter t, "Ligne d'en-tete  = " & lo.HeaderRowRange.Row
        Ajouter t, "Nb lignes donnees= " & nbLignes & " (contenu non exporte)"
        Ajouter t, "Filtre auto      = " & OuiNon(lo.ShowAutoFilter)
        Ajouter t, "Ligne de total   = " & OuiNon(lo.ShowTotals)
        Ajouter t, "Style            = " & styleTableau
        Ajouter t, "Colonnes (index | nom | lettre | largeur | masquee | format | liste)"

        For Each lc In lo.ListColumns
            colonneFeuille = lc.Range.Column
            ligne = lc.index & " | " & lc.Name & " | " & LettreColonne(ws, colonneFeuille) & _
                    " | " & NombreVBA(ws.Columns(colonneFeuille).ColumnWidth)
            If ws.Columns(colonneFeuille).Hidden Then ligne = ligne & " | MASQUEE" Else ligne = ligne & " | -"
            If Not lc.DataBodyRange Is Nothing Then
                ligne = ligne & " | format=""" & lc.DataBodyRange.Cells(1, 1).NumberFormat & """"
                ligne = ligne & " | " & DecrireValidation(lc.DataBodyRange.Cells(1, 1))
            Else
                ligne = ligne & " | (tableau vide : format et liste non lisibles)"
            End If
            Ajouter t, ligne
        Next lc

    Next lo

End Sub


' =====================================================================================
' SECTION 7 : listes déroulantes et autres validations HORS tableaux
' =====================================================================================
Private Sub SectionValidations(ByVal ws As Worksheet, ByVal plageTableaux As Range, ByRef t As String)

    Dim plage As Range
    Dim bloc As Range
    Dim nb As Long

    Ajouter t, ""
    Ajouter t, "## VALIDATIONS HORS TABLEAUX (plage | description)"

    ' SpecialCells provoque une erreur s'il n'existe aucune validation : on
    ' l'intercepte et on considère simplement qu'il n'y en a pas.
    On Error Resume Next
    Set plage = ws.Cells.SpecialCells(xlCellTypeAllValidation)
    On Error GoTo 0

    nb = 0
    If Not plage Is Nothing Then
        For Each bloc In plage.Areas
            If Not DansPlage(bloc.Cells(1, 1), plageTableaux) Then
                Ajouter t, bloc.Address(False, False) & " | " & DecrireValidation(bloc.Cells(1, 1))
                nb = nb + 1
            End If
        Next bloc
    End If
    If nb = 0 Then Ajouter t, "(aucune)"

End Sub

' Décrit la validation d'une cellule, ou "sans liste" si elle n'en a pas.
Private Function DecrireValidation(ByVal c As Range) As String

    Dim typeV As Long
    Dim s As String

    ' Lire .Validation.Type provoque une erreur quand la cellule n'a AUCUNE
    ' validation : c'est la seule façon fiable de le savoir.
    On Error Resume Next
    typeV = c.Validation.Type
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        DecrireValidation = "sans liste"
        Exit Function
    End If

    s = "validation=" & NomTypeValidation(typeV)
    s = s & " Formula1=""" & Echapper(c.Validation.Formula1) & """"
    If typeV <> xlValidateList Then s = s & " Formula2=""" & Echapper(c.Validation.Formula2) & """"
    s = s & " alerte=" & NomAlerte(c.Validation.AlertStyle)
    If c.Validation.ErrorTitle <> "" Then s = s & " titreErreur=""" & Echapper(c.Validation.ErrorTitle) & """"
    If c.Validation.ErrorMessage <> "" Then s = s & " messageErreur=""" & Echapper(c.Validation.ErrorMessage) & """"
    If c.Validation.InputMessage <> "" Then s = s & " messageSaisie=""" & Echapper(c.Validation.InputMessage) & """"
    On Error GoTo 0

    DecrireValidation = s

End Function


' =====================================================================================
' SECTION 8 : mises en forme conditionnelles de toute la feuille
' =====================================================================================
Private Sub SectionMisesEnFormeConditionnelles(ByVal ws As Worksheet, ByRef t As String)

    Dim nb As Long
    Dim i As Long
    Dim fc As Object
    Dim adresse As String, typeFC As String, formule As String, couleur As String

    Ajouter t, ""
    Ajouter t, "## MISES EN FORME CONDITIONNELLES (plage | type | formule | effet)"

    On Error Resume Next
    nb = ws.Cells.FormatConditions.count
    On Error GoTo 0

    For i = 1 To nb
        adresse = "?": typeFC = "?": formule = "": couleur = ""
        Set fc = Nothing
        ' Chaque propriété est lue séparément sous "On Error Resume Next" : certaines
        ' n'existent pas pour tous les types de règle (barres de données, icônes...).
        On Error Resume Next
        Set fc = ws.Cells.FormatConditions(i)
        adresse = fc.AppliesTo.Address(False, False)
        typeFC = CStr(fc.Type)
        formule = fc.Formula1
        If Not IsNull(fc.Font.Color) Then couleur = " police=" & CouleurVBA(fc.Font.Color)
        If ValeurVraie(fc.Font.Bold) Then couleur = couleur & " gras"
        On Error GoTo 0
        Ajouter t, adresse & " | type=" & typeFC & " | """ & Echapper(formule) & """ |" & couleur
    Next i
    If nb = 0 Then Ajouter t, "(aucune)"

End Sub


' =====================================================================================
' SECTION 9 : formes et boutons (position, ancrage dans les cellules, texte, macro)
' Les positions sont en points ; les cellules d'ancrage (haut-gauche / bas-droite)
' sont plus parlantes pour reconstruire un bouton à partir d'une plage, comme le
' fait mod_Display.ConstruireBoutons.
' =====================================================================================
Private Sub SectionFormes(ByVal ws As Worksheet, ByRef t As String)

    Dim shp As Shape
    Dim texteForme As String, macro As String, ancrage As String, genre As String

    Ajouter t, ""
    Ajouter t, "## FORMES ET BOUTONS (nom | genre | ancrage | gauche,haut,largeur,hauteur | texte | macro | visible)"

    If ws.Shapes.count = 0 Then
        Ajouter t, "(aucune)"
        Exit Sub
    End If

    For Each shp In ws.Shapes
        texteForme = "": macro = "": ancrage = "": genre = ""

        On Error Resume Next
        ' Un bouton "Contrôle de formulaire" expose son texte via .Caption ; une
        ' forme ordinaire via son cadre de texte (TextFrame2).
        texteForme = shp.OLEFormat.Object.Caption
        If Err.Number <> 0 Or texteForme = "" Then
            Err.Clear
            texteForme = shp.TextFrame2.TextRange.Text
        End If
        Err.Clear
        macro = shp.OnAction
        ancrage = shp.TopLeftCell.Address(False, False) & ":" & shp.BottomRightCell.Address(False, False)
        genre = NomGenreForme(shp)
        On Error GoTo 0

        Ajouter t, shp.Name & " | " & genre & " | " & ancrage & " | " & _
                   NombreVBA(Round(shp.Left, 1)) & "," & NombreVBA(Round(shp.Top, 1)) & "," & _
                   NombreVBA(Round(shp.Width, 1)) & "," & NombreVBA(Round(shp.Height, 1)) & _
                   " | """ & Echapper(texteForme) & """ | " & macro & " | " & OuiNon(shp.Visible)
    Next shp

End Sub


' =====================================================================================
' SECTION 10 : noms définis liés à cette feuille (portée feuille, ou portée classeur
' pointant vers cette feuille). Rappel (piège déjà rencontré sur ce projet) : la
' PORTÉE d'un nom compte -- un nom créé par ws.Names.Add n'est pas un nom créé par
' ThisWorkbook.Names.Add. La colonne "portee" le précise.
' =====================================================================================
Private Sub SectionNoms(ByVal ws As Worksheet, ByRef t As String)

    Dim nm As Name
    Dim reference As String
    Dim portee As String
    Dim nb As Long

    Ajouter t, ""
    Ajouter t, "## NOMS DEFINIS (nom | portee | fait reference a | visible)"

    nb = 0
    For Each nm In ThisWorkbook.Names
        reference = ""
        portee = "classeur"
        On Error Resume Next
        reference = nm.RefersTo
        If TypeOf nm.Parent Is Worksheet Then portee = "feuille " & nm.Parent.Name
        On Error GoTo 0

        If portee = "feuille " & ws.Name _
           Or InStr(1, reference, ws.Name & "!", vbTextCompare) > 0 _
           Or InStr(1, reference, ws.Name & "'!", vbTextCompare) > 0 Then
            Ajouter t, nm.Name & " | " & portee & " | " & Echapper(reference) & " | " & OuiNon(nm.Visible)
            nb = nb + 1
        End If
    Next nm
    If nb = 0 Then Ajouter t, "(aucun)"

End Sub


' =====================================================================================
' ÉCRITURE DU FICHIER ET DOSSIER DE DESTINATION
' =====================================================================================

' Dossier d'export, créé au besoin. Cas particulier OneDrive / SharePoint :
' ThisWorkbook.Path renvoie alors une adresse web ("https://..."), inutilisable pour
' créer un dossier ; on se rabat dans ce cas sur le dossier Documents de l'utilisateur.
Private Function DossierExport() As String

    Dim baseDossier As String

    baseDossier = ThisWorkbook.Path
    If baseDossier = "" Or LCase$(Left$(baseDossier, 4)) = "http" Then
        baseDossier = Environ$("USERPROFILE") & "\Documents"
    End If
    baseDossier = baseDossier & "\" & EXPORT_SOUS_DOSSIER

    On Error Resume Next
    If Dir(baseDossier, vbDirectory) = "" Then MkDir baseDossier
    If Err.Number <> 0 Then
        MsgBox mod_Display.FR("Impossible de cr{e2}er le dossier d'export :") & vbLf & baseDossier & vbLf & _
               Err.Description, vbExclamation
        DossierExport = ""
        Exit Function
    End If
    On Error GoTo 0

    DossierExport = baseDossier

End Function

' Écrit le texte en UTF-8 SANS BOM (3 octets invisibles en tête de fichier que
' ADODB.Stream ajoute d'office, et que GitHub ou un comparateur de fichiers
' afficheraient comme une différence parasite). Les fins de ligne sont déjà des LF
' (voir Ajouter).
Private Sub EcrireFichierUTF8(ByVal chemin As String, ByVal contenu As String)

    Dim fluxTexte As Object
    Dim fluxBinaire As Object

    Set fluxTexte = CreateObject("ADODB.Stream")
    fluxTexte.Type = 2              ' 2 = mode texte
    fluxTexte.Charset = "utf-8"
    fluxTexte.Open
    fluxTexte.WriteText contenu

    ' Retrait du BOM : on repasse en mode binaire, on saute les 3 premiers octets et
    ' on recopie le reste dans un second flux, qui est celui enregistré sur le disque.
    fluxTexte.Position = 0
    fluxTexte.Type = 1              ' 1 = mode binaire
    fluxTexte.Position = 3

    Set fluxBinaire = CreateObject("ADODB.Stream")
    fluxBinaire.Type = 1
    fluxBinaire.Open
    fluxTexte.CopyTo fluxBinaire
    fluxBinaire.SaveToFile chemin, 2   ' 2 = écraser le fichier s'il existe déjà

    fluxBinaire.Close
    fluxTexte.Close

End Sub


' =====================================================================================
' PETITS UTILITAIRES DE MISE EN FORME DU TEXTE
' =====================================================================================

' Ajoute une ligne au texte, terminée par un LF seul (vbLf), comme demandé pour les
' fichiers du projet.
Private Sub Ajouter(ByRef t As String, ByVal ligne As String)
    t = t & ligne & vbLf
End Sub

' Rend une valeur lisible sur une seule ligne : retours à la ligne -> "\n",
' tabulations -> "\t".
Private Function Echapper(ByVal s As String) As String
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    Echapper = s
End Function

' Nombre écrit avec un POINT décimal, quel que soit le réglage régional (Str$ utilise
' toujours le point ; Trim$ retire l'espace qu'il place devant les nombres positifs).
Private Function NombreVBA(ByVal x As Double) As String
    NombreVBA = Trim$(Str$(x))
End Function

' Couleur Excel (nombre Long) -> texte "RGB(r, g, b)" réutilisable dans le code.
' Excel range la couleur sous la forme rouge + vert*256 + bleu*65536.
Private Function CouleurVBA(ByVal couleur As Long) As String
    CouleurVBA = "RGB(" & (couleur Mod 256) & ", " & ((couleur \ 256) Mod 256) & ", " & _
                 ((couleur \ 65536) Mod 256) & ")"
End Function

Private Function OuiNon(ByVal v As Variant) As String
    If ValeurVraie(v) Then OuiNon = "Oui" Else OuiNon = "Non"
End Function

' Vrai uniquement si v vaut réellement True (Null, vide ou erreur -> Faux), pour ne
' jamais planter sur une mise en forme "mixte".
Private Function ValeurVraie(ByVal v As Variant) As Boolean
    If IsNull(v) Or IsEmpty(v) Or IsError(v) Then Exit Function
    On Error Resume Next
    ValeurVraie = CBool(v)
    On Error GoTo 0
End Function

' La cellule c appartient-elle à la plage p ? (p peut valoir Nothing)
Private Function DansPlage(ByVal c As Range, ByVal p As Range) As Boolean
    If p Is Nothing Then Exit Function
    DansPlage = Not (Intersect(c, p) Is Nothing)
End Function

' La cellule a-t-elle au moins une bordure sur l'un de ses 4 côtés ?
Private Function ABordure(ByVal c As Range) As Boolean
    Dim cote As Variant
    For Each cote In Array(xlEdgeLeft, xlEdgeTop, xlEdgeRight, xlEdgeBottom)
        If Not IsNull(c.Borders(cote).LineStyle) Then
            If c.Borders(cote).LineStyle <> xlLineStyleNone Then ABordure = True: Exit Function
        End If
    Next cote
End Function

' Numéro de colonne -> lettre(s) ("A", "AB"...). Address(True, False) renvoie par
' exemple "AB$1" ; on garde ce qui précède le "$".
Private Function LettreColonne(ByVal ws As Worksheet, ByVal numeroColonne As Long) As String
    LettreColonne = Split(ws.Cells(1, numeroColonne).Address(True, False), "$")(0)
End Function

' Remplace les caractères interdits dans un nom de fichier Windows.
Private Function NomFichierSur(ByVal nom As String) As String
    Dim interdit As Variant
    For Each interdit In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        nom = Replace(nom, CStr(interdit), "_")
    Next interdit
    NomFichierSur = nom
End Function

Private Function NomVisibilite(ByVal v As XlSheetVisibility) As String
    Select Case v
        Case xlSheetVisible:    NomVisibilite = "visible"
        Case xlSheetHidden:     NomVisibilite = "masquee (xlSheetHidden)"
        Case xlSheetVeryHidden: NomVisibilite = "tres masquee (xlSheetVeryHidden)"
        Case Else:              NomVisibilite = CStr(v)
    End Select
End Function

' Alignement horizontal ; "" pour l'alignement par défaut (Standard).
Private Function NomAlignementH(ByVal v As Variant) As String
    If IsNull(v) Then NomAlignementH = "mixte": Exit Function
    Select Case v
        Case xlGeneral:                NomAlignementH = ""
        Case xlLeft:                   NomAlignementH = "xlLeft"
        Case xlCenter:                 NomAlignementH = "xlCenter"
        Case xlRight:                  NomAlignementH = "xlRight"
        Case xlJustify:                NomAlignementH = "xlJustify"
        Case xlCenterAcrossSelection:  NomAlignementH = "xlCenterAcrossSelection"
        Case xlFill:                   NomAlignementH = "xlFill"
        Case Else:                     NomAlignementH = CStr(v)
    End Select
End Function

' Alignement vertical ; "" pour l'alignement par défaut (en bas).
Private Function NomAlignementV(ByVal v As Variant) As String
    If IsNull(v) Then NomAlignementV = "mixte": Exit Function
    Select Case v
        Case xlBottom:   NomAlignementV = ""
        Case xlTop:      NomAlignementV = "xlTop"
        Case xlCenter:   NomAlignementV = "xlCenter"
        Case xlJustify:  NomAlignementV = "xlJustify"
        Case Else:       NomAlignementV = CStr(v)
    End Select
End Function

Private Function NomTypeValidation(ByVal v As Long) As String
    Select Case v
        Case xlValidateList:        NomTypeValidation = "xlValidateList"
        Case xlValidateWholeNumber: NomTypeValidation = "xlValidateWholeNumber"
        Case xlValidateDecimal:     NomTypeValidation = "xlValidateDecimal"
        Case xlValidateDate:        NomTypeValidation = "xlValidateDate"
        Case xlValidateTime:        NomTypeValidation = "xlValidateTime"
        Case xlValidateTextLength:  NomTypeValidation = "xlValidateTextLength"
        Case xlValidateCustom:      NomTypeValidation = "xlValidateCustom"
        Case xlValidateInputOnly:   NomTypeValidation = "xlValidateInputOnly"
        Case Else:                  NomTypeValidation = CStr(v)
    End Select
End Function

Private Function NomAlerte(ByVal v As Long) As String
    Select Case v
        Case xlValidAlertStop:        NomAlerte = "xlValidAlertStop"
        Case xlValidAlertWarning:     NomAlerte = "xlValidAlertWarning"
        Case xlValidAlertInformation: NomAlerte = "xlValidAlertInformation"
        Case Else:                    NomAlerte = CStr(v)
    End Select
End Function

' Genre d'une forme : bouton de formulaire, case à cocher, zone de texte...
Private Function NomGenreForme(ByVal shp As Shape) As String
    If shp.Type = msoFormControl Then
        Select Case shp.FormControlType
            Case xlButtonControl: NomGenreForme = "Bouton (controle de formulaire)"
            Case xlCheckBox:      NomGenreForme = "Case a cocher"
            Case xlDropDown:      NomGenreForme = "Liste deroulante"
            Case xlLabel:         NomGenreForme = "Etiquette"
            Case xlGroupBox:      NomGenreForme = "Zone de groupe"
            Case Else:            NomGenreForme = "Controle de formulaire " & shp.FormControlType
        End Select
    ElseIf shp.Type = msoTextBox Then
        NomGenreForme = "Zone de texte"
    ElseIf shp.Type = msoAutoShape Then
        NomGenreForme = "Forme automatique"
    ElseIf shp.Type = msoPicture Then
        NomGenreForme = "Image"
    Else
        NomGenreForme = "Type " & shp.Type
    End If
End Function
