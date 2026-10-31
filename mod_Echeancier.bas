Attribute VB_Name = "mod_Echeancier"
Option Explicit

Private mFiltresMemorises As Object   ' filtres levés avant suppression, par nom de tableau

' =====================================================================================
' MODULE : mod_Echeancier
'
' RÔLE (à lire en premier) :
'   Logique de l'ÉCHÉANCIER : opérations à venir (récurrentes ou ponctuelles) qui
'   alimentent le solde prévisionnel à 15 et 30 jours de la feuille Synthese.
'   La construction des feuilles est dans mod_InstallEcheancier (partie 1/2).
'
' PRINCIPES VALIDÉS PAR L'OPÉRATEUR (09/10/2026) :
'   - Deux tableaux : TblRecurrences (les RÈGLES : "loyer, tous les mois...") et
'     TblEcheances (les OCCURRENCES à venir, générées à partir des règles, ou saisies
'     directement quand l'opération est ponctuelle).
'   - AUCUN HISTORIQUE : quand une opération importée est rapprochée d'une échéance et
'     que l'opérateur valide, l'échéance est SUPPRIMÉE de TblEcheances. Seule
'     l'opération reste (dans TblOperations). Si l'échéance venait d'une règle, la
'     règle continue de produire les occurrences suivantes.
'   - Une échéance NON rapprochée reste dans la liste, et reste comptée dans le solde
'     prévisionnel, jusqu'à ce que l'opérateur la supprime lui-même.
'   - Le rapprochement n'utilise pas que le Tiers : la Catégorie / Sous-catégorie
'     peuvent suffire (achat dont on ne connaît pas le commerçant). La date peut aussi
'     être imprécise : précision "Mois" = l'opération peut tomber n'importe quel jour du
'     mois prévu.
'   - Le solde de départ est SAISI par l'opérateur (les exports OFX ne contiennent pas
'     de solde) : voir le bloc "Prévisionnel de trésorerie" de la feuille Synthese.
'
' POINTS D'ENTRÉE :
'   - TraiterApresImport ........ appelée par mod_ImportOFX à la fin de chaque import.
'   - SaisirEcheance ............ bouton de la feuille Synthese (et de la liste).
'   - DeclarerRecurrenteDepuisOperation ... bouton "Rendre récurrente" de l'écran de
'                                 recherche (donc aussi de l'écran "Dernier import").
'   - AfficherEcheancier / AfficherRecurrences / SortirEcheancier / Supprimer... /
'     ActualiserRecurrences ..... boutons des feuilles de liste.
'
' À PROPOS DES ACCENTS : textes affichés via mod_Display.FR(); commentaires en UTF-8.
' =====================================================================================

' --- Réglages (à modifier ici si besoin, un seul endroit) ---------------------------
' Les règles produisent des occurrences jusqu'à cet horizon (en jours à partir d'aujourd'hui).
' Chaque règle active a toujours AU MOINS une occurrence dans la liste, même plus lointaine.
Public Const ECH_HORIZON_JOURS As Long = 60
' Fenêtre de rapprochement, précision "Jour" : l'opération peut arriver de FEN_AVANT jours
' avant à FEN_APRES jours après la date prévue. Règle hebdomadaire : +/- FEN_HEBDO jours.
Public Const FEN_AVANT As Long = 5
Public Const FEN_APRES As Long = 10
Public Const FEN_HEBDO As Long = 3
' Tolérance sur le montant (en %) quand l'échéance n'en précise aucune.
Public Const TOLERANCE_DEFAUT As Double = 10
' Garde-fou doublon à la saisie : écart de montant (en %) en dessous duquel deux échéances sont jugées identiques.
Public Const DOUBLON_TOLERANCE_PCT As Double = 20

' --- Valeurs possibles des colonnes de type liste ------------------------------------
Public Const PERIODE_PONCTUELLE As String = "Ponctuelle"
Public Const PERIODE_HEBDO As String = "Hebdomadaire"
Public Const PERIODE_MENSUELLE As String = "Mensuelle"
Public Const PERIODE_TRIMESTRIELLE As String = "Trimestrielle"
Public Const PERIODE_ANNUELLE As String = "Annuelle"
Public Const PRECISION_JOUR As String = "Jour"
Public Const PRECISION_MOIS As String = "Mois"

' --- Adresses des champs du formulaire frm_Echeance (colonne B = étiquettes, C = saisie) ---
Public Const EF_ADR_LIBELLE As String = "C6"
Public Const EF_ADR_TIERS As String = "C7"
Public Const EF_ADR_CAT As String = "C8"
Public Const EF_ADR_SOUS As String = "C9"
Public Const EF_ADR_MONTANT As String = "C10"
Public Const EF_ADR_CHEQUE As String = "C11"
Public Const EF_ADR_PERIODE As String = "C12"
Public Const EF_ADR_DATE As String = "C13"
Public Const EF_ADR_DATEFIN As String = "C14"
Public Const EF_ADR_PRECISION As String = "C15"
Public Const EF_ADR_TOLERANCE As String = "C16"
Public Const EF_ADR_MESSAGE As String = "B18"      ' zone fusionnée B18:E18
' Zones techniques masquées du formulaire : sous-catégories (Z), périodicités (AA), précisions (AB).
Public Const EF_COL_AIDE_SOUS As Long = 26
Public Const EF_COL_AIDE_PERIODE As Long = 27
Public Const EF_COL_AIDE_PRECISION As Long = 28
Public Const EF_LIGNE_AIDE_MAX As Long = 300
' Liste des Tiers connus (colonne AC) : plus longue, car il y a bien plus de Tiers que de sous-catégories.
Public Const EF_COL_AIDE_TIERS As Long = 29
Public Const EF_LIGNE_AIDE_TIERS_MAX As Long = 5000

' Une échéance telle que saisie dans le formulaire.
Public Type TypeEcheance
    Libelle As String
    Tiers As String
    Categorie As String
    SousCategorie As String
    Montant As Double          ' signé : négatif = dépense, positif = entrée d'argent
    NumCheque As String        ' facultatif, échéance ponctuelle uniquement ("" = pas de chèque)
    Periodicite As String
    DatePrevue As Date
    DateFin As Date            ' 0 = pas de date de fin (règle sans limite); sans effet pour une ponctuelle
    PrecisionDate As String    ' "Jour" ou "Mois"
    TolerancePct As Double
End Type

' --- État du formulaire de saisie -----------------------------------------------------
Public g_EchEnCours As Boolean        ' Vrai tant que le formulaire est ouvert (boucle d'attente)
Public g_EchVerrouActif As Boolean    ' Vrai tant que Worksheet_Deactivate doit réactiver la feuille
Private g_EchValide As Boolean
Private g_EchRes As TypeEcheance
Private g_EchNomFeuillePrec As String
' État du rapprochement fait AVANT l'écriture de l'import (voir PreRapprocherAvantEcriture).
Private g_PreRapprochementFait As Boolean
Private g_EcheancesValidees As Object      ' ID_Echeance validés par l'opérateur -> True
Private g_OpSignature As Object            ' ID_Transaction -> empreinte (date, tiers, montant, catégorie...) déjà examinée
Private g_OpEcheance As Object             ' ID_Transaction -> ID_Echeance validée ("" si aucune)
Private g_OpManuel As Object               ' ID_Transaction -> True : rapprochement choisi à la main (prioritaire, voir RapprocherManuellement)
Private g_MajEcheances As Object           ' ID_Echeance rapprochée à la main -> valeurs de l'opération à reporter sur sa règle (appliquées après l'import)


' =====================================================================================
' OUTILS DE LECTURE / ÉCRITURE DES TABLEAUX
' =====================================================================================
Private Function ObtenirTable(ByVal nomFeuille As String, ByVal nomTable As String) As ListObject
    On Error Resume Next
    Set ObtenirTable = ThisWorkbook.Worksheets(nomFeuille).ListObjects(nomTable)
    On Error GoTo 0
End Function

Private Function TableEcheances() As ListObject
    Set TableEcheances = ObtenirTable(mod_VarGlobales.NOM_FEUILLE_ECHEANCIER, mod_VarGlobales.NOM_TABLE_ECHEANCES)
End Function

Private Function TableRecurrences() As ListObject
    Set TableRecurrences = ObtenirTable(mod_VarGlobales.NOM_FEUILLE_RECURRENCES, mod_VarGlobales.NOM_TABLE_RECURRENCES)
End Function

' Vrai si les deux tableaux existent (sinon l'échéancier n'est pas installé : les appels
' automatiques, comme celui de l'import, ne font alors rien).
Public Function EcheancierInstalle() As Boolean
    Dim a As ListObject, b As ListObject
    Set a = TableEcheances()
    Set b = TableRecurrences()
    EcheancierInstalle = Not (a Is Nothing) And Not (b Is Nothing)
End Function

' Numéro de colonne d'un tableau, retrouvé par son NOM (jamais par un numéro fixe).
Private Function ColIdx(ByVal lo As ListObject, ByVal nomCol As String) As Long
    On Error Resume Next
    ColIdx = lo.ListColumns(nomCol).Index
    On Error GoTo 0
End Function

' Vrai si le tableau n'a pas de colonne de ce nom (utilisé par mod_InstallEcheancier).
Public Function ColonneAbsente(ByVal lo As ListObject, ByVal nomCol As String) As Boolean
    ColonneAbsente = (ColIdx(lo, nomCol) = 0)
End Function

' Un tableau Excel a toujours au moins une ligne : "vide" = une seule ligne dont la
' première colonne (l'identifiant) est vide.
Private Function TableVide(ByVal lo As ListObject) As Boolean
    If lo.DataBodyRange Is Nothing Then
        TableVide = True
    ElseIf lo.ListRows.Count = 1 Then
        TableVide = (mod_DataStructure.CellText(lo.DataBodyRange.Cells(1, 1).Value2) = "")
    End If
End Function

' Renvoie la ligne à remplir : la ligne vide existante, sinon une nouvelle ligne.
Private Function NouvelleLigne(ByVal lo As ListObject) As ListRow
    If lo.DataBodyRange Is Nothing Then
        Set NouvelleLigne = lo.ListRows.Add
    ElseIf TableVide(lo) Then
        Set NouvelleLigne = lo.ListRows(1)
    Else
        Set NouvelleLigne = lo.ListRows.Add
    End If
End Function

' Lit une colonne entière dans un tableau 1 à N (gère le cas d'une seule ligne).
Private Function LireColonne(ByVal lo As ListObject, ByVal nomCol As String) As Variant()
    Dim v As Variant, r() As Variant
    Dim i As Long, n As Long
    n = lo.ListRows.Count
    ReDim r(1 To n)
    v = lo.ListColumns(nomCol).DataBodyRange.Value2
    If IsArray(v) Then
        For i = 1 To n
            r(i) = v(i, 1)
        Next i
    Else
        r(1) = v
    End If
    LireColonne = r
End Function

Private Function ProchainID(ByVal lo As ListObject, ByVal nomCol As String) As Long
    Dim v() As Variant
    Dim i As Long, mx As Long
    If lo.DataBodyRange Is Nothing Then
        ProchainID = 1
        Exit Function
    End If
    v = LireColonne(lo, nomCol)
    For i = 1 To UBound(v)
        If Not IsEmpty(v(i)) And Not IsError(v(i)) Then
            If IsNumeric(v(i)) Then
                If CLng(v(i)) > mx Then mx = CLng(v(i))
            End If
        End If
    Next i
    ProchainID = mx + 1
End Function

' Valeur de cellule -> Date (0 si ce n'est pas une date).
Private Function VersDate(ByVal v As Variant) As Date
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then
        VersDate = CDate(CDbl(v))
    ElseIf IsDate(v) Then
        VersDate = CDate(v)
    End If
End Function


' =====================================================================================
' CALCUL DES DATES
' =====================================================================================
' Ajoute n mois à une date en gardant le "jour d'ancrage" (le 31 donne le 28 en février,
' puis redevient le 31 en mars : le jour ne dérive pas).
Private Function AjouterMois(ByVal d As Date, ByVal n As Long, ByVal jourAncre As Long) As Date
    Dim premier As Date
    Dim dernierJour As Long
    premier = DateSerial(Year(d), Month(d) + n, 1)
    dernierJour = Day(DateSerial(Year(premier), Month(premier) + 1, 0))
    If jourAncre < 1 Then jourAncre = Day(d)
    If jourAncre > dernierJour Then jourAncre = dernierJour
    AjouterMois = DateSerial(Year(premier), Month(premier), jourAncre)
End Function

Private Function DateSuivante(ByVal d As Date, ByVal periode As String, ByVal jourAncre As Long) As Date
    Select Case periode
        Case PERIODE_HEBDO:         DateSuivante = d + 7
        Case PERIODE_MENSUELLE:     DateSuivante = AjouterMois(d, 1, jourAncre)
        Case PERIODE_TRIMESTRIELLE: DateSuivante = AjouterMois(d, 3, jourAncre)
        Case PERIODE_ANNUELLE:      DateSuivante = AjouterMois(d, 12, jourAncre)
        Case Else:                  DateSuivante = DateSerial(2999, 12, 31)   ' périodicité inconnue : on arrête la génération
    End Select
End Function


' =====================================================================================
' ENREGISTREMENT D'UNE ÉCHÉANCE (saisie manuelle, ou déclarée depuis une opération)
' =====================================================================================
' Ponctuelle -> une seule ligne dans TblEcheances.
' Récurrente -> une règle dans TblRecurrences, puis génération de ses occurrences.
Public Function EnregistrerEcheance(ByRef e As TypeEcheance) As Boolean

    Dim loE As ListObject, loR As ListObject
    Dim lr As ListRow
    Dim idRec As Long

    Set loE = TableEcheances()
    Set loR = TableRecurrences()
    If loE Is Nothing Or loR Is Nothing Then
        MsgBox mod_Display.FR("L'{e2}ch{e2}ancier n'est pas install{e2}. Ex{e2}cutez d'abord la macro InstallerEcheancier."), vbExclamation
        Exit Function
    End If

    If e.PrecisionDate = PRECISION_MOIS Then
        e.DatePrevue = DateSerial(Year(e.DatePrevue), Month(e.DatePrevue), 1)
    End If

    If e.Periodicite = PERIODE_PONCTUELLE Then
        AjouterOccurrence loE, e, 0
    Else
        idRec = ProchainID(loR, "ID_Recurrence")
        Set lr = NouvelleLigne(loR)
        With lr.Range
            .Cells(1, ColIdx(loR, "ID_Recurrence")).Value = idRec
            .Cells(1, ColIdx(loR, "Libelle")).Value = e.Libelle
            .Cells(1, ColIdx(loR, "Tiers")).Value = e.Tiers
            .Cells(1, ColIdx(loR, "Categorie")).Value = e.Categorie
            .Cells(1, ColIdx(loR, "SousCategorie")).Value = e.SousCategorie
            .Cells(1, ColIdx(loR, "Montant")).Value = e.Montant
            .Cells(1, ColIdx(loR, "Periodicite")).Value = e.Periodicite
            .Cells(1, ColIdx(loR, "DateProchaine")).Value = e.DatePrevue
            If e.DateFin <> 0 Then .Cells(1, ColIdx(loR, "DateFin")).Value = e.DateFin
            .Cells(1, ColIdx(loR, "PrecisionDate")).Value = e.PrecisionDate
            If e.PrecisionDate = PRECISION_MOIS Then
                .Cells(1, ColIdx(loR, "JourAncre")).Value = 1
            Else
                .Cells(1, ColIdx(loR, "JourAncre")).Value = Day(e.DatePrevue)
            End If
            .Cells(1, ColIdx(loR, "TolerancePct")).Value = e.TolerancePct
            .Cells(1, ColIdx(loR, "Actif")).Value = "Oui"
        End With
        GenererHorizon
    End If

    EnregistrerEcheance = True

End Function

' Ajoute une occurrence dans TblEcheances (idRec = 0 : échéance ponctuelle, sans règle).
Private Sub AjouterOccurrence(ByVal loE As ListObject, ByRef e As TypeEcheance, ByVal idRec As Long)

    Dim idEch As Long
    Dim lr As ListRow

    idEch = ProchainID(loE, "ID_Echeance")
    Set lr = NouvelleLigne(loE)
    With lr.Range
        .Cells(1, ColIdx(loE, "ID_Echeance")).Value = idEch
        If idRec > 0 Then .Cells(1, ColIdx(loE, "ID_Recurrence")).Value = idRec
        .Cells(1, ColIdx(loE, "Libelle")).Value = e.Libelle
        .Cells(1, ColIdx(loE, "Tiers")).Value = e.Tiers
        .Cells(1, ColIdx(loE, "Categorie")).Value = e.Categorie
        .Cells(1, ColIdx(loE, "SousCategorie")).Value = e.SousCategorie
        .Cells(1, ColIdx(loE, "Montant")).Value = e.Montant
        If ColIdx(loE, "NumCheque") > 0 Then .Cells(1, ColIdx(loE, "NumCheque")).Value = e.NumCheque
        .Cells(1, ColIdx(loE, "DatePrevue")).NumberFormat = "dd/mm/yyyy"
        .Cells(1, ColIdx(loE, "DatePrevue")).Value = e.DatePrevue
        .Cells(1, ColIdx(loE, "PrecisionDate")).Value = e.PrecisionDate
        .Cells(1, ColIdx(loE, "Periodicite")).Value = e.Periodicite
        .Cells(1, ColIdx(loE, "TolerancePct")).Value = e.TolerancePct
    End With

End Sub


' =====================================================================================
' GÉNÉRATION DES OCCURRENCES À PARTIR DES RÈGLES
' =====================================================================================
' Pour chaque règle ACTIVE : crée les occurrences jusqu'à l'horizon (ECH_HORIZON_JOURS),
' et au moins une si la règle n'en a plus aucune dans la liste (après un rapprochement,
' par exemple). La colonne DateProchaine de la règle mémorise la prochaine date à créer :
' une occurrence supprimée ou rapprochée n'est donc jamais recréée.
Public Sub GenererHorizon()

    Dim loE As ListObject, loR As ListObject
    Dim nbOcc As Object
    Dim v() As Variant
    Dim rg As Range
    Dim e As TypeEcheance
    Dim i As Long, garde As Long, idRec As Long, ancre As Long
    Dim d As Date, limite As Date, dFin As Date
    Dim cle As String, periode As String
    Dim cId As Long, cAct As Long, cDt As Long, cJour As Long
    Dim termine() As Boolean

    Set loE = TableEcheances()
    Set loR = TableRecurrences()
    If loE Is Nothing Or loR Is Nothing Then Exit Sub
    If TableVide(loR) Then
        TrierEcheances
        Exit Sub
    End If

    limite = Date + ECH_HORIZON_JOURS

    ' Nombre d'occurrences actuellement présentes pour chaque règle.
    Set nbOcc = CreateObject("Scripting.Dictionary")
    If Not TableVide(loE) Then
        v = LireColonne(loE, "ID_Recurrence")
        For i = 1 To UBound(v)
            If Not IsEmpty(v(i)) And Not IsError(v(i)) Then
                If IsNumeric(v(i)) Then
                    cle = CStr(CLng(v(i)))
                    If nbOcc.Exists(cle) Then
                        nbOcc(cle) = nbOcc(cle) + 1
                    Else
                        nbOcc.Add cle, 1
                    End If
                End If
            End If
        Next i
    End If

    cId = ColIdx(loR, "ID_Recurrence")
    cAct = ColIdx(loR, "Actif")
    cDt = ColIdx(loR, "DateProchaine")
    cJour = ColIdx(loR, "JourAncre")

    ReDim termine(1 To loR.ListRows.Count)
    For i = 1 To loR.ListRows.Count
        Set rg = loR.ListRows(i).Range
        If mod_DataStructure.CellText(rg.Cells(1, cId).Value2) <> "" And _
           StrComp(mod_DataStructure.CellText(rg.Cells(1, cAct).Value2), "Oui", vbTextCompare) = 0 Then

            idRec = mod_DataStructure.ToLong(rg.Cells(1, cId).Value2)
            d = VersDate(rg.Cells(1, cDt).Value2)
            dFin = VersDate(rg.Cells(1, ColIdx(loR, "DateFin")).Value2)   ' 0 = pas de date de fin
            If d <> 0 Then
                periode = mod_DataStructure.CellText(rg.Cells(1, ColIdx(loR, "Periodicite")).Value2)
                ancre = mod_DataStructure.ToLong(rg.Cells(1, cJour).Value2)

                e.NumCheque = ""
                e.Libelle = mod_DataStructure.CellText(rg.Cells(1, ColIdx(loR, "Libelle")).Value2)
                e.Tiers = mod_DataStructure.CellText(rg.Cells(1, ColIdx(loR, "Tiers")).Value2)
                e.Categorie = mod_DataStructure.CellText(rg.Cells(1, ColIdx(loR, "Categorie")).Value2)
                e.SousCategorie = mod_DataStructure.CellText(rg.Cells(1, ColIdx(loR, "SousCategorie")).Value2)
                e.Montant = mod_DataStructure.ToDouble(rg.Cells(1, ColIdx(loR, "Montant")).Value2)
                e.Periodicite = periode
                e.PrecisionDate = mod_DataStructure.CellText(rg.Cells(1, ColIdx(loR, "PrecisionDate")).Value2)
                e.TolerancePct = ToleranceDe(rg.Cells(1, ColIdx(loR, "TolerancePct")).Value2)

                cle = CStr(idRec)
                garde = 0
                Do While garde < 500
                    If dFin <> 0 Then
                        If d > dFin Then Exit Do      ' règle terminée : plus aucune occurrence
                    End If
                    If d > limite And nbOcc.Exists(cle) Then Exit Do
                    e.DatePrevue = d
                    AjouterOccurrence loE, e, idRec
                    nbOcc(cle) = 1
                    d = DateSuivante(d, periode, ancre)
                    garde = garde + 1
                Loop
                rg.Cells(1, cDt).Value = d
                ' Règle épuisée : date de fin atteinte et plus aucune échéance dans la liste.
                ' Elle n'a plus d'utilité (pas d'historique conservé) : elle est retirée plus bas.
                If dFin <> 0 Then
                    If d > dFin And Not nbOcc.Exists(cle) Then termine(i) = True
                End If
            End If
        End If
    Next i

    LeverFiltre loR
    For i = loR.ListRows.Count To 1 Step -1
        If termine(i) Then loR.ListRows(i).Delete
    Next i
    RestaurerFiltre loR

    TrierEcheances

End Sub

' Force le format JJ/MM/AAAA sur la colonne DatePrevue, puis trie la liste par date prévue
' croissante (la plus proche en haut ; les dates vides en bas).
Public Sub TrierEcheances()

    Dim lo As ListObject
    Dim c As Long

    On Error GoTo Fin
    Set lo = TableEcheances()
    If lo Is Nothing Then Exit Sub
    If TableVide(lo) Then Exit Sub
    c = ColIdx(lo, "DatePrevue")
    If c = 0 Then Exit Sub

    lo.ListColumns(c).DataBodyRange.NumberFormat = "dd/mm/yyyy"
    With lo.Sort
        .SortFields.Clear
        .SortFields.Add Key:=lo.ListColumns(c).DataBodyRange, SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
Fin:

End Sub

Private Function ToleranceDe(ByVal v As Variant) As Double
    If IsEmpty(v) Or IsError(v) Then
        ToleranceDe = TOLERANCE_DEFAUT
    ElseIf Not IsNumeric(v) Then
        ToleranceDe = TOLERANCE_DEFAUT
    ElseIf mod_DataStructure.CellText(v) = "" Then
        ToleranceDe = TOLERANCE_DEFAUT
    Else
        ToleranceDe = CDbl(v)
    End If
End Function


' =====================================================================================
' APRÈS CHAQUE IMPORT (appelée par mod_ImportOFX)
' =====================================================================================
' 1. complète les occurrences des règles ; 2. propose les rapprochements entre les
' opérations qui viennent d'être importées et les échéances ; 3. complète à nouveau
' (une règle dont l'occurrence vient d'être rapprochée en retrouve une).
' Le solde prévisionnel de Synthese est fait de formules : il se met à jour tout seul.
' Sans effet tant que l'échéancier n'est pas installé : l'import ne peut pas en souffrir.
Public Sub TraiterApresImport(ByRef listeID() As String, ByVal nb As Long)

    On Error GoTo Erreur

    If Not EcheancierInstalle() Then Exit Sub

    GenererHorizon
    If g_PreRapprochementFait Then
        ' Les rapprochements ont déjà été proposés AVANT l'écriture (PreRapprocherAvantEcriture) :
        ' on applique seulement ceux que l'opérateur a validés, sans reposer la question.
        AppliquerMajEcheancier              ' mises à jour acceptées lors d'un rapprochement manuel
        SupprimerEcheancesParID g_EcheancesValidees
        Set g_EcheancesValidees = Nothing
        Set g_MajEcheances = Nothing
        g_PreRapprochementFait = False
    ElseIf nb > 0 Then
        Rapprocher listeID, nb
    End If
    GenererHorizon
    Exit Sub

Erreur:
    MsgBox mod_Display.FR("L'{e2}ch{e2}ancier n'a pas pu {ea}tre mis {a2} jour apr{e1}s l'import :") & vbCrLf & _
           Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
           mod_Display.FR("L'import lui-m{ea}me est termin{e2} et correct."), vbExclamation, mod_Display.FR("{E2}ch{e2}ancier")

End Sub

' Variante sans nouvelle opération (import qui n'a rien ajouté).
Public Sub MettreAJourEcheancier()
    On Error Resume Next
    If EcheancierInstalle() Then GenererHorizon
    On Error GoTo 0
End Sub


' =====================================================================================
' RAPPROCHEMENT opérations importées <-> échéances
' =====================================================================================
' Une échéance est candidate pour une opération si :
'   - le montant a le même signe et l'écart reste dans la tolérance de l'échéance ;
'   - la date est dans la fenêtre (précision "Mois" : même mois que la date prévue) ;
'   - le Tiers de l'opération CONTIENT le Tiers de l'échéance (si renseigné) ;
'   - la Catégorie / Sous-catégorie sont égales (si renseignées).
' Chaque opération et chaque échéance ne sont utilisées qu'une fois : on retient d'abord
' les paires les plus proches (écart de date, puis de montant). L'opérateur valide.
Private Sub Rapprocher(ByRef listeID() As String, ByVal nb As Long)

    Dim loOps As ListObject
    Dim dictNouveaux As Object
    Dim cID() As Variant, cDt() As Variant, cTiers() As Variant, cMont() As Variant, cCat() As Variant, cSous() As Variant, cChq() As Variant
    Dim opIdx() As Long
    Dim nOp As Long, i As Long

    If TableVide(TableEcheances()) Then Exit Sub

    On Error Resume Next
    Set loOps = ThisWorkbook.Worksheets("Import_data").ListObjects("TblOperations")
    On Error GoTo 0
    If loOps Is Nothing Then Exit Sub
    If loOps.DataBodyRange Is Nothing Then Exit Sub

    ' --- Opérations de CET import ---
    Set dictNouveaux = CreateObject("Scripting.Dictionary")
    For i = 1 To nb
        dictNouveaux(CStr(listeID(i))) = True
    Next i

    cID = LireColonne(loOps, "ID_Transaction")
    cDt = LireColonne(loOps, "Date_Comptable")
    cTiers = LireColonne(loOps, "Tiers")
    cMont = LireColonne(loOps, "Montant")
    cCat = LireColonne(loOps, "Categorie")
    cSous = LireColonne(loOps, "SousCategorie")
    cChq = LireColonneSure(loOps, "Num_Cheque")

    ReDim opIdx(1 To nb)
    For i = 1 To UBound(cID)
        If dictNouveaux.Exists(CStr(cID(i))) Then
            If nOp < nb Then
                nOp = nOp + 1
                opIdx(nOp) = i
            End If
        End If
    Next i
    If nOp = 0 Then Exit Sub

    SupprimerEcheancesParID ChoisirRapprochements(cDt, cTiers, cMont, cCat, cSous, cChq, opIdx, nOp, _
                                                  mod_Display.FR("Rapprochement avec l'{e2}ch{e2}ancier"))

End Sub

' =====================================================================================
' RAPPROCHEMENT AVANT L'ÉCRITURE DE L'IMPORT (appelée par mod_ImportOFX)
' =====================================================================================
' Appelée une fois le contrôle des catégories terminé (catégories, sous-catégories et Tiers
' définitifs) et AVANT l'écriture dans TblOperations. Les opérations sont encore en mémoire :
' on les compare aux échéances, on affiche les cas trouvés et l'opérateur confirme. Rien n'est
' supprimé ici : les échéances validées sont mémorisées, puis retirées par TraiterApresImport
' une fois l'import réellement écrit. Si l'import est abandonné, les échéances restent intactes.
' Tous les tableaux sont indexés 1..nb, dans l'ordre des opérations importées.
Public Sub PreRapprocherAvantEcriture(ByRef cID() As Variant, ByRef cDt() As Variant, ByRef cTiers() As Variant, ByRef cMont() As Variant, _
                                      ByRef cCat() As Variant, ByRef cSous() As Variant, ByRef cChq() As Variant, _
                                      ByVal nb As Long)

    Dim i As Long

    g_PreRapprochementFait = False

    On Error GoTo Erreur
    If nb <= 0 Then Exit Sub
    If g_OpSignature Is Nothing Then InitialiserRapprochementImport

    ' Les opérations déjà examinées dans le formulaire de contrôle (bouton Suivant) et restées
    ' inchangées sont ignorées par RapprocherOperationImport : seules les autres sont proposées ici.
    For i = 1 To nb
        RapprocherOperationImport CStr(cID(i)), cDt(i), cTiers(i), cMont(i), cCat(i), cSous(i), cChq(i)
    Next i
    g_PreRapprochementFait = True
    Exit Sub

Erreur:
    ' Échec de la détection : l'import continue, et le rapprochement habituel aura lieu après l'écriture.
    MsgBox mod_Display.FR("La reconnaissance des {e2}ch{e2}ances avant l'{e2}criture a {e2}chou{e2} :") & vbCrLf & _
           Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
           mod_Display.FR("L'import continue ; le rapprochement sera propos{e2} apr{e1}s l'{e2}criture."), vbExclamation, mod_Display.FR("{E2}ch{e2}ancier")
    g_PreRapprochementFait = False
    Set g_EcheancesValidees = Nothing
    Set g_MajEcheances = Nothing

End Sub

' Remet à zéro les rapprochements mémorisés. Appelée par l'import avant le contrôle des catégories.
Public Sub InitialiserRapprochementImport()
    Set g_EcheancesValidees = CreateObject("Scripting.Dictionary")
    Set g_OpSignature = CreateObject("Scripting.Dictionary")
    Set g_OpEcheance = CreateObject("Scripting.Dictionary")
    Set g_OpManuel = CreateObject("Scripting.Dictionary")
    Set g_MajEcheances = CreateObject("Scripting.Dictionary")
    g_PreRapprochementFait = False
End Sub

' Examine UNE opération en cours d'import : si elle correspond à une ou plusieurs échéances, un
' message « Échéancier » la présente avec chaque candidate (OUI / NON : candidate suivante /
' ANNULER : pas de rapprochement pour cette opération). Appelée depuis le formulaire de contrôle
' des catégories (bouton Suivant) puis, pour les opérations restantes, avant l'écriture.
' Une opération déjà examinée n'est réexaminée que si ses données ont changé (catégorie, Tiers...).
' Rien n'est supprimé ici : voir TraiterApresImport.
Public Sub RapprocherOperationImport(ByVal idOp As String, ByVal d As Variant, ByVal tiers As Variant, _
                                     ByVal montant As Variant, ByVal cat As Variant, ByVal sous As Variant, _
                                     ByVal cheque As Variant)

    Dim cDt() As Variant, cTiers() As Variant, cMont() As Variant
    Dim cCat() As Variant, cSous() As Variant, cChq() As Variant
    Dim opIdx() As Long
    Dim res As Object
    Dim sig As String, ancien As String
    Dim cle As Variant

    On Error GoTo Erreur

    If g_OpSignature Is Nothing Then InitialiserRapprochementImport
    If Not EcheancierInstalle() Then Exit Sub
    If TableVide(TableEcheances()) Then Exit Sub

    ' Rapprochement choisi à la main (RapprocherManuellement) : prioritaire, jamais remis en cause ici.
    If g_OpManuel.Exists(idOp) Then Exit Sub

    sig = SignatureOperation(d, tiers, montant, cat, sous, cheque)

    If g_OpSignature.Exists(idOp) Then
        If g_OpSignature(idOp) = sig Then Exit Sub
        ancien = CStr(g_OpEcheance(idOp))
        If ancien <> "" Then
            If g_EcheancesValidees.Exists(ancien) Then g_EcheancesValidees.Remove ancien
        End If
    End If
    g_OpSignature(idOp) = sig
    g_OpEcheance(idOp) = ""

    ReDim cDt(1 To 1): ReDim cTiers(1 To 1): ReDim cMont(1 To 1)
    ReDim cCat(1 To 1): ReDim cSous(1 To 1): ReDim cChq(1 To 1): ReDim opIdx(1 To 1)
    cDt(1) = d: cTiers(1) = tiers: cMont(1) = montant: cCat(1) = cat: cSous(1) = sous: cChq(1) = cheque
    opIdx(1) = 1

    Set res = ChoisirRapprochements(cDt, cTiers, cMont, cCat, cSous, cChq, opIdx, 1, _
                                    mod_Display.FR("{E2}ch{e2}ancier"), g_EcheancesValidees)
    For Each cle In res.Keys
        g_OpEcheance(idOp) = CStr(cle)
        g_EcheancesValidees(CStr(cle)) = True
    Next cle
    Exit Sub

Erreur:
    MsgBox mod_Display.FR("La reconnaissance de l'{e2}ch{e2}ance a {e2}chou{e2} pour cette op{e2}ration :") & vbCrLf & _
           Err.Number & " - " & Err.Description, vbExclamation, mod_Display.FR("{E2}ch{e2}ancier")

End Sub

' Empreinte d'une opération (date, tiers, montant, catégorie, sous-catégorie, chèque) : sert à ne pas
' réexaminer une opération dont rien n'a changé.
Private Function SignatureOperation(ByVal d As Variant, ByVal tiers As Variant, ByVal montant As Variant, _
                                    ByVal cat As Variant, ByVal sous As Variant, ByVal cheque As Variant) As String
    SignatureOperation = Format(VersDate(d), "yyyymmdd") & "|" & mod_DataStructure.CellText(tiers) & "|" & _
          Format(mod_DataStructure.ToDouble(montant), "0.00") & "|" & mod_DataStructure.CellText(cat) & "|" & _
          mod_DataStructure.CellText(sous) & "|" & mod_DataStructure.CellText(cheque)
End Function


' =====================================================================================
' RAPPROCHEMENT MANUEL (bouton « Échéancier... » du contrôle des catégories)
' =====================================================================================
' Quand la détection automatique n'a pas reconnu une échéance (écart de montant ou de date trop
' grand, Tiers différent...), l'opérateur la désigne lui-même dans frm_ChoixEcheance.
'
' RÈGLES :
'   - L'OPÉRATION N'EST JAMAIS MODIFIÉE (elle vient de la banque) : seul l'échéancier peut l'être.
'   - Comme pour un rapprochement automatique, rien n'est supprimé tout de suite : l'échéance est
'     retirée à l'écriture de l'import (TraiterApresImport). Si l'import est abandonné, tout reste intact.
'   - Si l'échéance vient d'une règle de récurrence ET que l'opération diffère d'elle (montant,
'     Tiers, catégorie, sous-catégorie), on demande à l'opérateur s'il faut aligner la règle et ses
'     échéances à venir sur les valeurs de l'opération. La modification est elle aussi différée
'     jusqu'à l'écriture de l'import (AppliquerMajEcheancier). Une échéance ponctuelle disparaît au
'     rapprochement : il n'y a alors rien à mettre à jour.
'   - Un rapprochement manuel est prioritaire : la détection automatique ne le remet jamais en cause
'     (g_OpManuel), même si l'opérateur modifie ensuite la catégorie ou le Tiers de l'opération.
'
' Renvoie le texte à afficher dans le formulaire de contrôle ("" si rien n'a changé).
Public Function RapprocherManuellement(ByVal idOp As String, ByVal d As Variant, ByVal tiers As Variant, _
                                       ByVal montant As Variant, ByVal cat As Variant, ByVal sous As Variant, _
                                       ByVal cheque As Variant) As String

    Dim loE As ListObject
    Dim eV As Variant
    Dim iId As Long, iLib As Long, iTiers As Long, iCat As Long, iSous As Long, iMont As Long
    Dim iDate As Long, iPrec As Long, iPer As Long, iTol As Long, iChq As Long
    Dim nE As Long, nC As Long, nPrises As Long
    Dim i As Long, k As Long, j As Long, tmp As Long
    Dim dOp As Date, dE As Date
    Dim mOp As Double, mEch As Double, sc As Double
    Dim tiersOp As String, catOp As String, sousOp As String, chqOp As String
    Dim idE As String, idCourant As String, idChoisi As String
    Dim prec As String, periode As String, chqE As String, qui As String
    Dim texteJ As String, textePct As String, raisons As String, sig As String
    Dim idU() As String, libU() As String, detU() As String, scoreU() As Double, ligneU() As Long
    Dim ordre() As Long
    Dim libS() As String, detS() As String, idS() As String, ligneS() As Long
    Dim libCourant As String, libChoisi As String, info As String
    Dim lignesDiff As String
    Dim reponse As VbMsgBoxResult
    Dim resultat As Long
    Dim maj As Variant

    On Error GoTo Erreur

    If g_OpSignature Is Nothing Or g_EcheancesValidees Is Nothing Or g_MajEcheances Is Nothing Or g_OpManuel Is Nothing Then InitialiserRapprochementImport
    If Not EcheancierInstalle() Then
        MsgBox mod_Display.FR("L'{e2}ch{e2}ancier n'est pas install{e2}. Ex{e2}cutez d'abord la macro InstallerEcheancier."), vbExclamation
        Exit Function
    End If
    If TableVide(TableEcheances()) Then
        MsgBox mod_Display.FR("L'{e2}ch{e2}ancier ne contient aucune {e2}ch{e2}ance."), vbInformation, mod_Display.FR("{E2}ch{e2}ancier")
        Exit Function
    End If
    Set loE = TableEcheances()

    dOp = VersDate(d)
    mOp = mod_DataStructure.ToDouble(montant)
    tiersOp = mod_DataStructure.CellText(tiers)
    catOp = mod_DataStructure.CellText(cat)
    sousOp = mod_DataStructure.CellText(sous)
    chqOp = NormaliserCheque(cheque)
    sig = SignatureOperation(d, tiers, montant, cat, sous, cheque)

    If g_OpEcheance.Exists(idOp) Then idCourant = CStr(g_OpEcheance(idOp))

    ' --- Échéances ---
    eV = loE.DataBodyRange.Value2
    nE = UBound(eV, 1)
    iId = ColIdx(loE, "ID_Echeance"): iLib = ColIdx(loE, "Libelle"): iTiers = ColIdx(loE, "Tiers")
    iCat = ColIdx(loE, "Categorie"): iSous = ColIdx(loE, "SousCategorie"): iMont = ColIdx(loE, "Montant")
    iDate = ColIdx(loE, "DatePrevue"): iPrec = ColIdx(loE, "PrecisionDate"): iPer = ColIdx(loE, "Periodicite")
    iTol = ColIdx(loE, "TolerancePct")
    iChq = ColIdx(loE, "NumCheque")      ' 0 si la colonne n'existe pas encore

    ReDim idU(1 To nE): ReDim libU(1 To nE): ReDim detU(1 To nE, 1 To mod_ChoixEcheance.CE_NB_DETAILS)
    ReDim scoreU(1 To nE): ReDim ligneU(1 To nE)

    ' --- Échéances candidates : toutes, sauf celles déjà rapprochées avec UNE AUTRE opération ---
    For i = 1 To nE
        If Not IsError(eV(i, iId)) Then
            If mod_DataStructure.CellText(eV(i, iId)) <> "" Then
                idE = CStr(eV(i, iId))
                If g_EcheancesValidees.Exists(idE) And idE <> idCourant Then
                    nPrises = nPrises + 1
                Else
                    dE = VersDate(eV(i, iDate))
                    mEch = mod_DataStructure.ToDouble(eV(i, iMont))
                    prec = mod_DataStructure.CellText(eV(i, iPrec))
                    periode = mod_DataStructure.CellText(eV(i, iPer))
                    chqE = ""
                    If iChq > 0 Then chqE = NormaliserCheque(eV(i, iChq))

                    ' Écarts de date et de montant, pour aider l'opérateur à reconnaître l'échéance.
                    If dE = 0 Then
                        texteJ = mod_Display.FR("date inconnue")
                    ElseIf prec = PRECISION_MOIS Then
                        j = (Year(dOp) - Year(dE)) * 12 + Month(dOp) - Month(dE)
                        If j = 0 Then
                            texteJ = mod_Display.FR("m{ea}me mois")
                        Else
                            texteJ = Format(j, "+0;-0") & " mois"
                        End If
                    Else
                        texteJ = Format(CLng(dOp - dE), "+0;-0;0") & " j"
                    End If
                    If (mOp > 0) <> (mEch > 0) Then
                        textePct = mod_Display.FR("signe oppos{e2}")
                    ElseIf mEch = 0 Then
                        textePct = ""
                    Else
                        textePct = Format((Abs(mOp) - Abs(mEch)) / Abs(mEch) * 100#, "+0;-0;0") & " %"
                    End If

                    ' Pourquoi la détection automatique n'a pas retenu cette échéance.
                    raisons = ""
                    If chqE <> "" Then
                        If chqE <> chqOp Then raisons = mod_Display.FR("le num{e2}ro de ch{e1}que ne correspond pas")
                    Else
                        If Not FenetreOk(dOp, dE, prec, periode) Then raisons = mod_Display.FR("date hors fen{ea}tre")
                        If Not MontantOk(mOp, mEch, ToleranceDe(eV(i, iTol))) Then
                            If raisons <> "" Then raisons = raisons & " ; "
                            raisons = raisons & mod_Display.FR("montant hors tol{e2}rance (") & ToleranceDe(eV(i, iTol)) & " %)"
                        End If
                        If Not TextesOk(tiersOp, catOp, sousOp, mod_DataStructure.CellText(eV(i, iTiers)), _
                                        mod_DataStructure.CellText(eV(i, iCat)), mod_DataStructure.CellText(eV(i, iSous))) Then
                            If raisons <> "" Then raisons = raisons & " ; "
                            raisons = raisons & mod_Display.FR("Tiers ou cat{e2}gorie diff{e2}rents")
                        End If
                    End If
                    If raisons = "" Then raisons = mod_Display.FR("tous les crit{e1}res sont remplis")

                    ' Libellé de la liste : « #12  05/11/2026  EDF  -118,00 €  (+12 j, +37 %) »
                    qui = mod_DataStructure.CellText(eV(i, iTiers))
                    If qui = "" Then qui = mod_DataStructure.CellText(eV(i, iLib))
                    If qui = "" Then qui = mod_DataStructure.CellText(eV(i, iCat))
                    If Len(qui) > 30 Then qui = Left$(qui, 30)

                    nC = nC + 1
                    idU(nC) = idE
                    ligneU(nC) = i
                    libU(nC) = LibelleListe("#" & idE & "  " & Format(dE, "dd/mm/yyyy") & "  " & qui & "  " & _
                                            Format(mEch, "#,##0.00") & " " & ChrW(8364) & "  (" & texteJ & IIf(textePct <> "", ", " & textePct, "") & ")")
                    detU(nC, 1) = mod_DataStructure.CellText(eV(i, iLib))
                    detU(nC, 2) = mod_DataStructure.CellText(eV(i, iTiers))
                    detU(nC, 3) = mod_DataStructure.CellText(eV(i, iCat))
                    If mod_DataStructure.CellText(eV(i, iSous)) <> "" Then detU(nC, 3) = detU(nC, 3) & " / " & mod_DataStructure.CellText(eV(i, iSous))
                    detU(nC, 4) = Format(mEch, "#,##0.00") & " " & ChrW(8364)
                    detU(nC, 5) = Format(dE, "dd/mm/yyyy")
                    If prec = PRECISION_MOIS Then detU(nC, 5) = detU(nC, 5) & mod_Display.FR(" (mois entier)")
                    If periode = "" Then periode = PERIODE_PONCTUELLE
                    detU(nC, 6) = periode
                    detU(nC, 7) = mod_Display.FR("Ecart : ") & texteJ & IIf(textePct <> "", " ; " & textePct, "") & _
                                  mod_Display.FR(". Non d{e2}tect{e2}e car : ") & raisons

                    ' Classement : numéro de chèque identique, puis signe identique, puis écart de date, puis de montant.
                    sc = Abs(CDbl(dOp) - CDbl(dE)) * 1000# + Abs(Abs(mOp) - Abs(mEch))
                    If (mOp > 0) <> (mEch > 0) Then sc = sc + 1E+12
                    If chqE <> "" And chqE = chqOp Then sc = -1
                    scoreU(nC) = sc
                    If idE = idCourant Then libCourant = libU(nC)
                End If
            End If
        End If
    Next i

    If nC = 0 Then
        MsgBox mod_Display.FR("Aucune {e2}ch{e2}ance n'est disponible : toutes sont d{e2}j{a2} rapproch{e2}es avec d'autres op{e2}rations de cet import."), _
               vbInformation, mod_Display.FR("{E2}ch{e2}ancier")
        Exit Function
    End If

    ' --- Tri par score croissant (tri par insertion sur un tableau d'indices) ---
    ReDim ordre(1 To nC)
    For k = 1 To nC
        ordre(k) = k
    Next k
    For k = 2 To nC
        tmp = ordre(k)
        j = k - 1
        Do While j >= 1
            If scoreU(ordre(j)) <= scoreU(tmp) Then Exit Do
            ordre(j + 1) = ordre(j)
            j = j - 1
        Loop
        ordre(j + 1) = tmp
    Next k
    ReDim libS(1 To nC): ReDim idS(1 To nC): ReDim ligneS(1 To nC)
    ReDim detS(1 To nC, 1 To mod_ChoixEcheance.CE_NB_DETAILS)
    For k = 1 To nC
        libS(k) = libU(ordre(k))
        idS(k) = idU(ordre(k))
        ligneS(k) = ligneU(ordre(k))
        For j = 1 To mod_ChoixEcheance.CE_NB_DETAILS
            detS(k, j) = detU(ordre(k), j)
        Next j
    Next k

    info = mod_Display.FR("Les {e2}ch{e2}ances sont class{e2}es de la plus proche {a2} la plus {e2}loign{e2}e de l'op{e2}ration. ")
    If nPrises > 0 Then
        info = info & nPrises & mod_Display.FR(" {e2}ch{e2}ance(s) d{e2}j{a2} rapproch{e2}e(s) avec une autre op{e2}ration ne sont pas propos{e2}es.")
    End If

    resultat = mod_ChoixEcheance.OuvrirChoixEcheance(Format(dOp, "dd/mm/yyyy"), tiersOp, _
                   Format(mOp, "#,##0.00") & " " & ChrW(8364), LibelleCategorie(catOp, sousOp), _
                   mod_DataStructure.CellText(cheque), libS, detS, nC, libCourant, info, libChoisi)

    Select Case resultat

        Case mod_ChoixEcheance.CE_DEFAIRE
            LibererEcheance idCourant
            g_OpEcheance(idOp) = ""
            If g_OpManuel.Exists(idOp) Then g_OpManuel.Remove idOp
            g_OpSignature(idOp) = sig        ' la détection automatique ne repose pas la question tout de suite
            RapprocherManuellement = mod_Display.FR("Rapprochement avec l'{e2}ch{e2}ance #") & idCourant & mod_Display.FR(" d{e2}fait.")

        Case mod_ChoixEcheance.CE_RAPPROCHER
            For k = 1 To nC
                If libS(k) = libChoisi Then
                    idChoisi = idS(k)
                    Exit For
                End If
            Next k
            If idChoisi = "" Then Exit Function

            ' Faut-il aligner la règle de récurrence sur les valeurs de l'opération ? (jamais l'inverse)
            maj = ProposerMajEcheance(loE, eV, ligneS(k), tiersOp, mOp, catOp, sousOp, lignesDiff)
            If IsArray(maj) Then
                reponse = MsgBox(mod_Display.FR("L'{e2}ch{e2}ance choisie (#") & idChoisi & mod_Display.FR(") provient d'une r{e1}gle de r{e2}currence. ") & _
                                 mod_Display.FR("Les valeurs de l'op{e2}ration diff{e1}rent de celles de l'{e2}ch{e2}ancier :") & vbCrLf & vbCrLf & _
                                 lignesDiff & vbCrLf & _
                                 mod_Display.FR("Mettre {a2} jour la r{e1}gle et ses {e2}ch{e2}ances {a2} venir avec les valeurs de l'op{e2}ration ?") & vbCrLf & _
                                 mod_Display.FR("(l'op{e2}ration elle-m{ea}me n'est jamais modifi{e2}e)") & vbCrLf & vbCrLf & _
                                 mod_Display.FR("OUI : rapprocher et mettre {a2} jour l'{e2}ch{e2}ancier.") & vbCrLf & _
                                 mod_Display.FR("NON : rapprocher sans modifier l'{e2}ch{e2}ancier.") & vbCrLf & _
                                 mod_Display.FR("ANNULER : ne pas rapprocher."), _
                                 vbYesNoCancel + vbQuestion, mod_Display.FR("Mise {a2} jour de l'{e2}ch{e2}ancier"))
                If reponse = vbCancel Then Exit Function
            End If

            ' Une éventuelle autre échéance déjà liée à cette opération est libérée.
            If idCourant <> "" And idCourant <> idChoisi Then LibererEcheance idCourant
            LibererEcheance idChoisi
            g_EcheancesValidees(idChoisi) = True
            g_OpEcheance(idOp) = idChoisi
            g_OpManuel(idOp) = True
            g_OpSignature(idOp) = sig
            If IsArray(maj) Then
                If reponse = vbYes Then g_MajEcheances.Add idChoisi, maj
            End If

            RapprocherManuellement = mod_Display.FR("Op{e2}ration rapproch{e2}e {a2} la main avec l'{e2}ch{e2}ance #") & idChoisi & _
                mod_Display.FR(" : elle sera retir{e2}e de l'{e2}ch{e2}ancier {a2} l'enregistrement")
            If IsArray(maj) Then
                If reponse = vbYes Then
                    RapprocherManuellement = RapprocherManuellement & mod_Display.FR(" et la r{e1}gle sera mise {a2} jour.")
                Else
                    RapprocherManuellement = RapprocherManuellement & "."
                End If
            Else
                RapprocherManuellement = RapprocherManuellement & "."
            End If

    End Select
    Exit Function

Erreur:
    MsgBox mod_Display.FR("Le rapprochement manuel a {e2}chou{e2} :") & vbCrLf & Err.Number & " - " & Err.Description, _
           vbExclamation, mod_Display.FR("{E2}ch{e2}ancier")

End Function

' Retire une échéance des rapprochements validés (et de sa mise à jour éventuelle).
Private Sub LibererEcheance(ByVal idEcheance As String)
    If idEcheance = "" Then Exit Sub
    If g_EcheancesValidees.Exists(idEcheance) Then g_EcheancesValidees.Remove idEcheance
    If g_MajEcheances.Exists(idEcheance) Then g_MajEcheances.Remove idEcheance
End Sub

' « Catégorie / Sous-catégorie » (ou la catégorie seule) pour l'affichage.
Private Function LibelleCategorie(ByVal cat As String, ByVal sous As String) As String
    If sous <> "" Then
        LibelleCategorie = cat & " / " & sous
    Else
        LibelleCategorie = cat
    End If
End Function

' Libellé d'une entrée de la liste : sans * ? ~ (jokers de la fonction EQUIV qui relit le libellé) et
' limité à 200 caractères (EQUIV refuse les valeurs plus longues que 255).
Private Function LibelleListe(ByVal texte As String) As String
    texte = Replace(texte, "*", "-")
    texte = Replace(texte, "?", "-")
    texte = Replace(texte, "~", "-")
    If Len(texte) > 200 Then texte = Left$(texte, 200)
    LibelleListe = texte
End Function

' Compare l'opération à l'échéance choisie. Si cette échéance vient d'une règle de récurrence qui existe
' encore ET que l'opération diffère d'elle, renvoie le tableau des valeurs à reporter
' (Montant ?, valeur, Tiers ?, valeur, Catégorie ?, valeur, Sous-catégorie ?, valeur), voir EcrireMajLigne.
' Sinon renvoie Empty (échéance ponctuelle : elle disparaît au rapprochement, rien à mettre à jour).
' lignesDiff reçoit la liste des différences, une par ligne, pour la question posée à l'opérateur.
' Un champ n'est comparé que s'il est RENSEIGNÉ dans l'échéance (c'est alors un critère de
' reconnaissance) ; le Tiers de l'échéance est un mot-clé, il n'est changé que s'il n'est pas déjà
' contenu dans celui de l'opération ; le montant n'est repris que si le signe est le même.
Private Function ProposerMajEcheance(ByVal loE As ListObject, ByRef eV As Variant, ByVal ligne As Long, _
                                     ByVal tiersOp As String, ByVal mOp As Double, ByVal catOp As String, _
                                     ByVal sousOp As String, ByRef lignesDiff As String) As Variant

    Dim loR As ListObject
    Dim iRec As Long, idRec As Long
    Dim mEch As Double
    Dim tiersE As String, catE As String, sousE As String
    Dim majMontant As Boolean, majTiers As Boolean, majCat As Boolean, majSous As Boolean
    Dim fleche As String

    lignesDiff = ""
    iRec = ColIdx(loE, "ID_Recurrence")
    If iRec = 0 Then Exit Function
    If IsError(eV(ligne, iRec)) Then Exit Function
    idRec = mod_DataStructure.ToLong(eV(ligne, iRec))
    If idRec <= 0 Then Exit Function

    Set loR = TableRecurrences()
    If loR Is Nothing Then Exit Function
    If TableVide(loR) Then Exit Function
    If LigneRegle(loR, idRec) = 0 Then Exit Function

    fleche = " " & ChrW(8594) & " "

    mEch = mod_DataStructure.ToDouble(eV(ligne, ColIdx(loE, "Montant")))
    If (mOp > 0) = (mEch > 0) Then
        If Abs(Abs(mOp) - Abs(mEch)) > 0.005 Then
            majMontant = True
            lignesDiff = lignesDiff & "- Montant : " & Format(mEch, "#,##0.00") & fleche & Format(mOp, "#,##0.00") & vbCrLf
        End If
    End If

    tiersE = mod_DataStructure.CellText(eV(ligne, ColIdx(loE, "Tiers")))
    If tiersE <> "" And tiersOp <> "" Then
        If InStr(1, tiersOp, tiersE, vbTextCompare) = 0 Then
            majTiers = True
            lignesDiff = lignesDiff & "- Tiers : " & tiersE & fleche & tiersOp & vbCrLf
        End If
    End If

    catE = mod_DataStructure.CellText(eV(ligne, ColIdx(loE, "Categorie")))
    If catE <> "" And catOp <> "" Then
        If StrComp(catOp, catE, vbTextCompare) <> 0 Then
            majCat = True
            lignesDiff = lignesDiff & mod_Display.FR("- Cat{e2}gorie : ") & catE & fleche & catOp & vbCrLf
        End If
    End If

    sousE = mod_DataStructure.CellText(eV(ligne, ColIdx(loE, "SousCategorie")))
    If sousE <> "" And sousOp <> "" Then
        If StrComp(sousOp, sousE, vbTextCompare) <> 0 Then
            majSous = True
            lignesDiff = lignesDiff & mod_Display.FR("- Sous-cat{e2}gorie : ") & sousE & fleche & sousOp & vbCrLf
        End If
    End If

    If Not (majMontant Or majTiers Or majCat Or majSous) Then Exit Function
    ProposerMajEcheance = Array(majMontant, mOp, majTiers, tiersOp, majCat, catOp, majSous, sousOp)

End Function

' Numéro de la ligne (1 = première ligne de données) de la règle dans TblRecurrences ; 0 si introuvable.
Private Function LigneRegle(ByVal loR As ListObject, ByVal idRec As Long) As Long
    Dim v() As Variant
    Dim i As Long
    v = LireColonne(loR, "ID_Recurrence")
    For i = 1 To UBound(v)
        If Not IsError(v(i)) Then
            If Not IsEmpty(v(i)) Then
                If IsNumeric(v(i)) Then
                    If CLng(v(i)) = idRec Then
                        LigneRegle = i
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
End Function

' Numéro de la règle d'une échéance (lue par son ID_Echeance) ; 0 si l'échéance n'a pas de règle ou n'existe plus.
Private Function RegleDeEcheance(ByVal loE As ListObject, ByVal idEcheance As String) As Long
    Dim vId() As Variant, vRec() As Variant
    Dim i As Long
    If ColIdx(loE, "ID_Recurrence") = 0 Then Exit Function
    vId = LireColonne(loE, "ID_Echeance")
    vRec = LireColonne(loE, "ID_Recurrence")
    For i = 1 To UBound(vId)
        If Not IsError(vId(i)) Then
            If CStr(vId(i)) = idEcheance Then
                If Not IsError(vRec(i)) Then
                    If Not IsEmpty(vRec(i)) Then
                        If IsNumeric(vRec(i)) Then RegleDeEcheance = CLng(vRec(i))
                    End If
                End If
                Exit Function
            End If
        End If
    Next i
End Function

' Reporte sur UNE ligne d'un tableau (TblRecurrences ou TblEcheances : mêmes noms de colonnes) les
' valeurs retenues par l'opérateur. maj : voir ProposerMajEcheance.
Private Sub EcrireMajLigne(ByVal lo As ListObject, ByVal ligne As Long, ByRef maj As Variant)
    Dim col As Long
    If maj(0) Then
        col = ColIdx(lo, "Montant")
        If col > 0 Then lo.DataBodyRange.Cells(ligne, col).Value = maj(1)
    End If
    If maj(2) Then
        col = ColIdx(lo, "Tiers")
        If col > 0 Then lo.DataBodyRange.Cells(ligne, col).Value = maj(3)
    End If
    If maj(4) Then
        col = ColIdx(lo, "Categorie")
        If col > 0 Then lo.DataBodyRange.Cells(ligne, col).Value = maj(5)
    End If
    If maj(6) Then
        col = ColIdx(lo, "SousCategorie")
        If col > 0 Then lo.DataBodyRange.Cells(ligne, col).Value = maj(7)
    End If
End Sub

' Appliquée par TraiterApresImport, une fois l'import écrit : pour chaque échéance rapprochée à la main
' dont l'opérateur a accepté la mise à jour, la règle d'origine et TOUTES ses échéances encore dans la
' liste prennent les valeurs de l'opération. Les occurrences futures, créées ensuite par GenererHorizon,
' reprennent celles de la règle.
Private Sub AppliquerMajEcheancier()

    Dim loE As ListObject, loR As ListObject
    Dim cle As Variant
    Dim maj As Variant
    Dim vRec() As Variant
    Dim idRec As Long, ligneR As Long, i As Long

    If g_MajEcheances Is Nothing Then Exit Sub
    If g_MajEcheances.Count = 0 Then Exit Sub
    Set loE = TableEcheances()
    Set loR = TableRecurrences()
    If loE Is Nothing Or loR Is Nothing Then Exit Sub
    If TableVide(loE) Or TableVide(loR) Then Exit Sub

    For Each cle In g_MajEcheances.Keys
        maj = g_MajEcheances(cle)
        idRec = RegleDeEcheance(loE, CStr(cle))
        If idRec > 0 Then
            ligneR = LigneRegle(loR, idRec)
            If ligneR > 0 Then EcrireMajLigne loR, ligneR, maj
            vRec = LireColonne(loE, "ID_Recurrence")
            For i = 1 To UBound(vRec)
                If Not IsError(vRec(i)) Then
                    If Not IsEmpty(vRec(i)) Then
                        If IsNumeric(vRec(i)) Then
                            If CLng(vRec(i)) = idRec Then EcrireMajLigne loE, i, maj
                        End If
                    End If
                End If
            Next i
        End If
    Next cle

End Sub

' Noyau commun : cherche les paires (opération, échéance) compatibles, affiche les propositions et
' fait valider l'opérateur. Renvoie un dictionnaire des ID_Echeance validés (vide si aucun).
' Les opérations sont désignées par opIdx(1..nOp), positions dans les tableaux c*.
Private Function ChoisirRapprochements(ByRef cDt() As Variant, ByRef cTiers() As Variant, ByRef cMont() As Variant, _
                                       ByRef cCat() As Variant, ByRef cSous() As Variant, ByRef cChq() As Variant, _
                                       ByRef opIdx() As Long, ByVal nOp As Long, ByVal titre As String, _
                                       Optional ByVal exclus As Object = Nothing) As Object

    Dim loE As ListObject
    Dim valides As Object
    Dim eV As Variant
    Dim chqE As String, chqOp As String, iChq As Long
    Dim nE As Long
    Dim i As Long, o As Long, k As Long, p As Long
    Dim dOp As Date, dE As Date
    Dim mOp As Double, mEch As Double, meilleur As Double
    Dim tiersOp As String, catOp As String, sousOp As String
    Dim prec As String, periode As String
    Dim pO() As Long, pE() As Long, pS() As Double, nP As Long
    Dim usedE() As Boolean, refus() As Boolean, nAcc As Long
    Dim reponse As VbMsgBoxResult
    Dim iId As Long, iLib As Long, iTiers As Long, iCat As Long, iSous As Long, iMont As Long, iDate As Long, iPrec As Long, iPer As Long, iTol As Long

    Set valides = CreateObject("Scripting.Dictionary")
    Set ChoisirRapprochements = valides

    If TableVide(TableEcheances()) Then Exit Function
    Set loE = TableEcheances()

    ' --- Échéances ---
    eV = loE.DataBodyRange.Value2
    nE = UBound(eV, 1)
    iId = ColIdx(loE, "ID_Echeance"): iLib = ColIdx(loE, "Libelle"): iTiers = ColIdx(loE, "Tiers")
    iCat = ColIdx(loE, "Categorie"): iSous = ColIdx(loE, "SousCategorie"): iMont = ColIdx(loE, "Montant")
    iDate = ColIdx(loE, "DatePrevue"): iPrec = ColIdx(loE, "PrecisionDate"): iPer = ColIdx(loE, "Periodicite")
    iTol = ColIdx(loE, "TolerancePct")
    iChq = ColIdx(loE, "NumCheque")      ' 0 si la colonne n'existe pas encore

    ' --- Toutes les paires compatibles ---
    ReDim pO(1 To 1): ReDim pE(1 To 1): ReDim pS(1 To 1)
    For o = 1 To nOp
        dOp = VersDate(cDt(opIdx(o)))
        mOp = mod_DataStructure.ToDouble(cMont(opIdx(o)))
        tiersOp = mod_DataStructure.CellText(cTiers(opIdx(o)))
        catOp = mod_DataStructure.CellText(cCat(opIdx(o)))
        sousOp = mod_DataStructure.CellText(cSous(opIdx(o)))
        chqOp = NormaliserCheque(cChq(opIdx(o)))
        For i = 1 To nE
            If mod_DataStructure.CellText(eV(i, iId)) <> "" Then
                dE = VersDate(eV(i, iDate))
                mEch = mod_DataStructure.ToDouble(eV(i, iMont))
                prec = mod_DataStructure.CellText(eV(i, iPrec))
                periode = mod_DataStructure.CellText(eV(i, iPer))
                chqE = ""
                If iChq > 0 Then chqE = NormaliserCheque(eV(i, iChq))
                If dE <> 0 And mEch <> 0 Then
                    If chqE <> "" Then
                        ' Échéance avec numéro de chèque : le numéro identifie l'opération, à lui seul
                        ' (ni la date, ni la tolérance, ni le Tiers ne sont examinés). Seul le signe du
                        ' montant doit concorder. Ces paires passent avant toutes les autres (score -1).
                        If chqOp = chqE And (mOp > 0) = (mEch > 0) Then
                            nP = nP + 1
                            If nP > UBound(pO) Then
                                ReDim Preserve pO(1 To nP): ReDim Preserve pE(1 To nP): ReDim Preserve pS(1 To nP)
                            End If
                            pO(nP) = o
                            pE(nP) = i
                            pS(nP) = -1
                        End If
                    ElseIf FenetreOk(dOp, dE, prec, periode) And _
                       MontantOk(mOp, mEch, ToleranceDe(eV(i, iTol))) And _
                       TextesOk(tiersOp, catOp, sousOp, mod_DataStructure.CellText(eV(i, iTiers)), _
                                mod_DataStructure.CellText(eV(i, iCat)), mod_DataStructure.CellText(eV(i, iSous))) Then
                        nP = nP + 1
                        If nP > UBound(pO) Then
                            ReDim Preserve pO(1 To nP): ReDim Preserve pE(1 To nP): ReDim Preserve pS(1 To nP)
                        End If
                        pO(nP) = o
                        pE(nP) = i
                        pS(nP) = Abs(dOp - dE) * 1000# + Abs(Abs(mOp) - Abs(mEch))
                    End If
                End If
            End If
        Next i
    Next o
    If nP = 0 Then Exit Function
    ReDim refus(1 To nP)

    ' --- Validation opération par opération ---
    ' Pour chaque opération, les échéances candidates sont proposées une à une, de la plus
    ' proche à la plus éloignée. OUI : rapprochée (l'échéance n'est plus proposée aux autres
    ' opérations). NON : candidate suivante. ANNULER : on passe à l'opération suivante.
    ReDim usedE(1 To nE)
    nAcc = 0
    ' Échéances déjà rapprochées par une autre opération de cet import : plus proposées.
    If Not exclus Is Nothing Then
        For i = 1 To nE
            If exclus.Exists(CStr(eV(i, iId))) Then usedE(i) = True
        Next i
    End If
    For o = 1 To nOp
        Do
            p = 0
            meilleur = 1E+30
            For k = 1 To nP
                If pO(k) = o Then
                    If Not usedE(pE(k)) And Not refus(k) Then
                        If pS(k) < meilleur Then
                            meilleur = pS(k)
                            p = k
                        End If
                    End If
                End If
            Next k
            If p = 0 Then Exit Do

            reponse = MsgBox(mod_Display.FR("{E2}ch{e2}ancier :") & vbCrLf & vbCrLf & _
                             mod_Display.FR("L'op{e2}ration en cours semble correspondre {a2} cette entr{e2}e de l'{e2}ch{e2}ancier :") & vbCrLf & vbCrLf & _
                             mod_Display.FR("Op{e2}ration : ") & TexteOperation(cDt(opIdx(o)), cTiers(opIdx(o)), cMont(opIdx(o))) & vbCrLf & vbCrLf & _
                             TexteEcheance(loE, eV, pE(p)) & vbCrLf & _
                             mod_Display.FR("Confirmez-vous ce rapprochement ?") & vbCrLf & vbCrLf & _
                             mod_Display.FR("OUI : rapprocher.  NON : chercher un autre candidat.  ANNULER : ne pas rapprocher cette op{e2}ration."), _
                             vbYesNoCancel + vbQuestion, titre & IIf(nOp > 1, " (" & o & "/" & nOp & ")", ""))
            If reponse = vbYes Then
                valides(CStr(eV(pE(p), iId))) = True
                usedE(pE(p)) = True
                Exit Do
            ElseIf reponse = vbCancel Then
                Exit Do
            Else
                refus(p) = True
            End If
        Loop
    Next o

End Function

' Description d'une opération importée pour le message de validation.
Private Function TexteOperation(ByVal d As Variant, ByVal tiers As Variant, ByVal montant As Variant) As String
    TexteOperation = Format(VersDate(d), "dd/mm/yyyy") & "  " & mod_DataStructure.CellText(tiers) & "  " & _
                     Format(mod_DataStructure.ToDouble(montant), "#,##0.00")
End Function

' Toutes les valeurs renseignées d'une ligne de l'échéancier, "Colonne : valeur", une par ligne.
Private Function TexteEcheance(ByVal lo As ListObject, ByRef eV As Variant, ByVal ligne As Long) As String
    Dim c As Long, t As String, nom As String, fmt As String, v As Variant
    For c = 1 To lo.ListColumns.Count
        v = eV(ligne, c)
        If Not IsError(v) Then
            If mod_DataStructure.CellText(v) <> "" Then
                nom = lo.ListColumns(c).Name
                fmt = lo.ListColumns(c).DataBodyRange.Cells(1, 1).NumberFormat
                If IsNumeric(v) And (InStr(fmt, "yy") > 0 Or InStr(fmt, "dd") > 0) Then
                    t = t & nom & " : " & Format(VersDate(v), IIf(InStr(fmt, "dd") > 0, "dd/mm/yyyy", "mm/yyyy")) & vbCrLf
                Else
                    t = t & nom & " : " & mod_DataStructure.CellText(v) & vbCrLf
                End If
            End If
        End If
    Next c
    TexteEcheance = t
End Function

' Supprime de TblEcheances les échéances dont l'ID figure dans le dictionnaire (rapprochées et validées).
Private Sub SupprimerEcheancesParID(ByVal valides As Object)

    Dim loE As ListObject
    Dim iId As Long, i As Long
    Dim v As Variant

    If valides Is Nothing Then Exit Sub
    If valides.Count = 0 Then Exit Sub
    If TableVide(TableEcheances()) Then Exit Sub
    Set loE = TableEcheances()
    iId = ColIdx(loE, "ID_Echeance")
    If iId = 0 Then Exit Sub

    LeverFiltre loE
    v = loE.DataBodyRange.Value2
    ' De bas en haut : les numéros de lignes restantes ne bougent pas.
    For i = UBound(v, 1) To 1 Step -1
        If valides.Exists(CStr(v(i, iId))) Then loE.ListRows(i).Delete
    Next i
    RestaurerFiltre loE

End Sub

' Lit une colonne si elle existe ; sinon renvoie un tableau de cellules vides (même taille).
Private Function LireColonneSure(ByVal lo As ListObject, ByVal nomCol As String) As Variant()
    Dim r() As Variant
    If ColIdx(lo, nomCol) > 0 Then
        LireColonneSure = LireColonne(lo, nomCol)
    Else
        ReDim r(1 To lo.ListRows.Count)
        LireColonneSure = r
    End If
End Function

' Numéro de chèque "comparable" : chiffres seuls, sans zéros en tête ("" si vide ou invalide).
' Sert à comparer le numéro saisi avec celui de la banque, quelle que soit la façon de l'écrire.
Private Function NormaliserCheque(ByVal v As Variant) As String
    Dim t As String, r As String, c As String
    Dim i As Long
    t = mod_DataStructure.CellText(v)
    For i = 1 To Len(t)
        c = Mid$(t, i, 1)
        If c >= "0" And c <= "9" Then r = r & c
    Next i
    Do While Left$(r, 1) = "0"
        r = Mid$(r, 2)
    Loop
    NormaliserCheque = r
End Function

' Si ce numéro de chèque est déjà passé en banque (colonne Num_Cheque de TblOperations),
' renvoie la description de l'opération ; sinon "".
Private Function ChequeDejaPasse(ByVal numNormalise As String) As String

    Dim lo As ListObject
    Dim cChq() As Variant, cDt() As Variant, cTiers() As Variant, cMont() As Variant
    Dim i As Long

    On Error Resume Next
    Set lo = ThisWorkbook.Worksheets("Import_data").ListObjects("TblOperations")
    On Error GoTo 0
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If ColIdx(lo, "Num_Cheque") = 0 Then Exit Function

    cChq = LireColonne(lo, "Num_Cheque")
    For i = 1 To UBound(cChq)
        If NormaliserCheque(cChq(i)) = numNormalise Then
            cDt = LireColonne(lo, "Date_Comptable")
            cTiers = LireColonne(lo, "Tiers")
            cMont = LireColonne(lo, "Montant")
            ChequeDejaPasse = Format(VersDate(cDt(i)), "dd/mm/yyyy") & ", " & _
                              mod_DataStructure.CellText(cTiers(i)) & ", " & _
                              Format(mod_DataStructure.ToDouble(cMont(i)), "#,##0.00")
            Exit Function
        End If
    Next i

End Function

' Si ce numéro de chèque est déjà saisi dans l'échéancier, renvoie la description de
' l'échéance ; sinon "".
Private Function ChequeDansEcheancier(ByVal numNormalise As String) As String

    Dim lo As ListObject
    Dim cChq() As Variant, cLib() As Variant, cDt() As Variant
    Dim i As Long

    Set lo = TableEcheances()
    If lo Is Nothing Then Exit Function
    If TableVide(lo) Then Exit Function
    If ColIdx(lo, "NumCheque") = 0 Then Exit Function

    cChq = LireColonne(lo, "NumCheque")
    For i = 1 To UBound(cChq)
        If NormaliserCheque(cChq(i)) = numNormalise Then
            cLib = LireColonne(lo, "Libelle")
            cDt = LireColonne(lo, "DatePrevue")
            ChequeDansEcheancier = mod_DataStructure.CellText(cLib(i)) & ", " & Format(VersDate(cDt(i)), "dd/mm/yyyy")
            Exit Function
        End If
    Next i

End Function

' Montant existant et montant saisi de même signe, à +/- DOUBLON_TOLERANCE_PCT % près
' (écart mesuré par rapport au montant déjà présent dans l'échéancier).
Private Function MontantProche(ByVal mExistant As Double, ByVal mSaisi As Double) As Boolean
    If mExistant = 0 Or mSaisi = 0 Then Exit Function
    If (mExistant > 0) <> (mSaisi > 0) Then Exit Function
    MontantProche = (Abs(mSaisi - mExistant) <= Abs(mExistant) * DOUBLON_TOLERANCE_PCT / 100 + 0.005)
End Function

' Garde-fou doublon. Deux échéances sont jugées identiques si Tiers, Catégorie et
' Sous-catégorie sont les mêmes (casse ignorée), le Montant est proche (+/- 20 %) ET :
'   - ponctuelle : la date tombe le même jour (le même mois si l'une des deux est "Mois") ;
'   - récurrente : une règle ACTIVE de même périodicité existe déjà.
' Renvoie la description de l'échéance ou de la règle existante, "" si aucun doublon.
Private Function DoublonEcheancier(ByRef e As TypeEcheance) As String

    Dim lo As ListObject
    Dim cT() As Variant, cC() As Variant, cS() As Variant, cM() As Variant
    Dim cD() As Variant, cP() As Variant, cA() As Variant, cPer() As Variant
    Dim i As Long
    Dim dE As Date
    Dim memeMois As Boolean
    Dim txtMontant As String

    If e.Periodicite = PERIODE_PONCTUELLE Then
        Set lo = TableEcheances()
    Else
        Set lo = TableRecurrences()
    End If
    If lo Is Nothing Then Exit Function
    If TableVide(lo) Then Exit Function

    cT = LireColonne(lo, "Tiers")
    cC = LireColonne(lo, "Categorie")
    cS = LireColonne(lo, "SousCategorie")
    cM = LireColonne(lo, "Montant")
    cPer = LireColonne(lo, "Periodicite")
    If e.Periodicite = PERIODE_PONCTUELLE Then
        cD = LireColonne(lo, "DatePrevue")
        cP = LireColonne(lo, "PrecisionDate")
    Else
        cA = LireColonne(lo, "Actif")
    End If

    For i = 1 To UBound(cT)
        If StrComp(mod_DataStructure.CellText(cT(i)), e.Tiers, vbTextCompare) = 0 _
           And StrComp(mod_DataStructure.CellText(cC(i)), e.Categorie, vbTextCompare) = 0 _
           And StrComp(mod_DataStructure.CellText(cS(i)), e.SousCategorie, vbTextCompare) = 0 _
           And MontantProche(mod_DataStructure.ToDouble(cM(i)), e.Montant) Then
            txtMontant = Format(mod_DataStructure.ToDouble(cM(i)), "#,##0.00") & " " & ChrW(8364)

            If e.Periodicite = PERIODE_PONCTUELLE Then
                dE = VersDate(cD(i))
                memeMois = (Year(dE) = Year(e.DatePrevue) And Month(dE) = Month(e.DatePrevue))
                If memeMois Then
                    If mod_DataStructure.CellText(cP(i)) = PRECISION_MOIS Or e.PrecisionDate = PRECISION_MOIS Then
                        DoublonEcheancier = "{e2}ch{e2}ance du " & Format(dE, "mm/yyyy") & ", montant " & txtMontant
                        Exit Function
                    ElseIf Day(dE) = Day(e.DatePrevue) Then
                        DoublonEcheancier = "{e2}ch{e2}ance du " & Format(dE, "dd/mm/yyyy") & ", montant " & txtMontant
                        Exit Function
                    End If
                End If
            Else
                If StrComp(mod_DataStructure.CellText(cA(i)), "Oui", vbTextCompare) = 0 _
                   And StrComp(mod_DataStructure.CellText(cPer(i)), e.Periodicite, vbTextCompare) = 0 Then
                    DoublonEcheancier = "r{e2}gle " & LCase$(e.Periodicite) & " d{e2}j{a2} active, montant " & txtMontant
                    Exit Function
                End If
            End If
        End If
    Next i

End Function

Private Function FenetreOk(ByVal dOp As Date, ByVal dE As Date, ByVal prec As String, ByVal periode As String) As Boolean
    Dim ecart As Double
    If prec = PRECISION_MOIS Then
        FenetreOk = (Year(dOp) = Year(dE) And Month(dOp) = Month(dE))
    Else
        ecart = dOp - dE
        If periode = PERIODE_HEBDO Then
            FenetreOk = (ecart >= -FEN_HEBDO And ecart <= FEN_HEBDO)
        Else
            FenetreOk = (ecart >= -FEN_AVANT And ecart <= FEN_APRES)
        End If
    End If
End Function

Private Function MontantOk(ByVal mOp As Double, ByVal mEch As Double, ByVal tolPct As Double) As Boolean
    If (mOp > 0) <> (mEch > 0) Then Exit Function
    MontantOk = (Abs(Abs(mOp) - Abs(mEch)) <= Abs(mEch) * tolPct / 100# + 0.005)
End Function

' Au moins un critère de texte (Tiers ou Catégorie) doit être renseigné dans l'échéance.
Private Function TextesOk(ByVal tiersOp As String, ByVal catOp As String, ByVal sousOp As String, _
                          ByVal tiersE As String, ByVal catE As String, ByVal sousE As String) As Boolean
    If tiersE = "" And catE = "" Then Exit Function
    If tiersE <> "" Then
        If InStr(1, tiersOp, tiersE, vbTextCompare) = 0 Then Exit Function
    End If
    If catE <> "" Then
        If StrComp(catOp, catE, vbTextCompare) <> 0 Then Exit Function
    End If
    If sousE <> "" Then
        If StrComp(sousOp, sousE, vbTextCompare) <> 0 Then Exit Function
    End If
    TextesOk = True
End Function


' =====================================================================================
' SAISIE : boutons "Saisir une échéance" (Synthese et liste)
' =====================================================================================
Public Sub SaisirEcheance()

    Dim e As TypeEcheance

    If Not EcheancierInstalle() Then
        MsgBox mod_Display.FR("L'{e2}ch{e2}ancier n'est pas install{e2}. Ex{e2}cutez d'abord la macro InstallerEcheancier."), vbExclamation
        Exit Sub
    End If

    ' Périodicité laissée vide au départ : sans périodicité, l'opération est unique.
    e.Periodicite = ""
    e.PrecisionDate = PRECISION_JOUR
    e.TolerancePct = TOLERANCE_DEFAUT
    e.DatePrevue = Date

    If OuvrirFormEcheance(e) Then
        If EnregistrerEcheance(e) Then MsgBox ConfirmationTexte(e), vbInformation, mod_Display.FR("{E2}ch{e2}ance enregistr{e2}e")
    End If

End Sub

Private Function ConfirmationTexte(ByRef e As TypeEcheance) As String
    If e.Periodicite = PERIODE_PONCTUELLE Then
        ConfirmationTexte = mod_Display.FR("{E2}ch{e2}ance ponctuelle ajout{e2}e {a2} la liste.")
    Else
        ConfirmationTexte = mod_Display.FR("R{e2}gle de r{e2}currence enregistr{e2}e (") & LCase$(e.Periodicite) & _
                            mod_Display.FR("). Ses prochaines occurrences sont dans la liste des {e2}ch{e2}ances.")
    End If
End Function

' Bouton "Rendre récurrente" de l'écran de recherche : préremplit le formulaire avec
' les données de l'opération sélectionnée (Tiers, Catégorie, Montant...), la prochaine
' date étant calculée sur le rythme mensuel (modifiable dans le formulaire).
Public Sub DeclarerRecurrenteDepuisOperation(ByVal idTransaction As String)

    Dim loOps As ListObject
    Dim cID() As Variant, cDt() As Variant, cTiers() As Variant, cMont() As Variant, cCat() As Variant, cSous() As Variant
    Dim i As Long, trouve As Long
    Dim e As TypeEcheance
    Dim d As Date

    If Not EcheancierInstalle() Then
        MsgBox mod_Display.FR("L'{e2}ch{e2}ancier n'est pas install{e2}. Ex{e2}cutez d'abord la macro InstallerEcheancier."), vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set loOps = ThisWorkbook.Worksheets("Import_data").ListObjects("TblOperations")
    On Error GoTo 0
    If loOps Is Nothing Then Exit Sub
    If loOps.DataBodyRange Is Nothing Then Exit Sub

    cID = LireColonne(loOps, "ID_Transaction")
    For i = 1 To UBound(cID)
        If CStr(cID(i)) = idTransaction Then
            trouve = i
            Exit For
        End If
    Next i
    If trouve = 0 Then
        MsgBox mod_Display.FR("Op{e2}ration introuvable dans TblOperations."), vbExclamation
        Exit Sub
    End If

    cDt = LireColonne(loOps, "Date_Comptable")
    cTiers = LireColonne(loOps, "Tiers")
    cMont = LireColonne(loOps, "Montant")
    cCat = LireColonne(loOps, "Categorie")
    cSous = LireColonne(loOps, "SousCategorie")

    e.Tiers = mod_DataStructure.CellText(cTiers(trouve))
    e.Libelle = e.Tiers
    e.Categorie = mod_DataStructure.CellText(cCat(trouve))
    e.SousCategorie = mod_DataStructure.CellText(cSous(trouve))
    e.Montant = mod_DataStructure.ToDouble(cMont(trouve))
    e.PrecisionDate = PRECISION_JOUR
    e.TolerancePct = TOLERANCE_DEFAUT

    e.Periodicite = PERIODE_MENSUELLE

    ' Prochaine date : même jour du mois, le mois suivant l'opération, puis on avance
    ' jusqu'à une date à venir (les occurrences passées ont déjà été importées).
    d = VersDate(cDt(trouve))
    If d = 0 Then d = Date
    Dim ancre As Long
    ancre = Day(d)
    d = DateSuivante(d, PERIODE_MENSUELLE, ancre)
    Do While d <= Date
        d = DateSuivante(d, PERIODE_MENSUELLE, ancre)
    Loop
    e.DatePrevue = d

    If OuvrirFormEcheance(e) Then
        If EnregistrerEcheance(e) Then MsgBox ConfirmationTexte(e), vbInformation, mod_Display.FR("{E2}ch{e2}ance enregistr{e2}e")
    End If

End Sub


' =====================================================================================
' FORMULAIRE frm_Echeance (même principe que mod_NouvelleCategorie)
' =====================================================================================
' Ouvre le formulaire prérempli avec "e" et attend sa fermeture. Renvoie True si
' l'opérateur a validé ; "e" contient alors les valeurs saisies.
Public Function OuvrirFormEcheance(ByRef e As TypeEcheance) As Boolean

    Dim ws As Worksheet

    On Error GoTo Erreur

    Set ws = mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_ECHEANCE)
    If ws Is Nothing Then
        MsgBox mod_Display.FR("La feuille '") & mod_VarGlobales.NOM_FEUILLE_ECHEANCE & mod_Display.FR("' est introuvable.") & vbCrLf & _
               mod_Display.FR("Ex{e2}cutez d'abord la macro InstallerEcheancier."), vbExclamation
        Exit Function
    End If

    g_EchNomFeuillePrec = ActiveSheet.Name
    g_EchValide = False
    RemplirChampsEch ws, e

    ws.Visible = xlSheetVisible
    ws.Activate
    mod_InstallCommun.MasquerQuadrillage
    ws.Range(EF_ADR_LIBELLE).Select

    ' Les 5 listes déroulantes sont posées une seconde fois, feuille affichée et active :
    ' elles ne dépendent ainsi pas de l'état de la feuille pendant son remplissage.
    AppliquerListesEch ws

    Application.EnableEvents = True
    g_EchEnCours = True
    g_EchVerrouActif = True
    Do While g_EchEnCours
        DoEvents
    Loop

    On Error Resume Next
    mod_InstallCommun.ActiverFeuille g_EchNomFeuillePrec
    On Error GoTo 0

    If g_EchValide Then
        e = g_EchRes
        OuvrirFormEcheance = True
    End If
    Exit Function

Erreur:
    Application.EnableEvents = True
    g_EchEnCours = False
    g_EchVerrouActif = False
    MsgBox mod_Display.FR("Erreur inattendue dans la saisie de l'{e2}ch{e2}ance :") & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
    OuvrirFormEcheance = False

End Function

Private Sub RemplirChampsEch(ByVal ws As Worksheet, ByRef e As TypeEcheance)

    Dim evenementsAvant As Boolean
    Dim numErr As Long

    evenementsAvant = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Sortie

    ws.Range(EF_ADR_LIBELLE).Value = e.Libelle
    ws.Range(EF_ADR_TIERS).Value = e.Tiers
    ws.Range(EF_ADR_CAT).Value = e.Categorie
    ws.Range(EF_ADR_SOUS).Value = e.SousCategorie
    If e.Montant <> 0 Then ws.Range(EF_ADR_MONTANT).Value = e.Montant Else ws.Range(EF_ADR_MONTANT).ClearContents
    ws.Range(EF_ADR_CHEQUE).Value = e.NumCheque
    ws.Range(EF_ADR_PERIODE).Value = e.Periodicite
    If e.DatePrevue <> 0 Then ws.Range(EF_ADR_DATE).Value = e.DatePrevue Else ws.Range(EF_ADR_DATE).ClearContents
    If e.DateFin <> 0 Then ws.Range(EF_ADR_DATEFIN).Value = e.DateFin Else ws.Range(EF_ADR_DATEFIN).ClearContents
    ws.Range(EF_ADR_PRECISION).Value = e.PrecisionDate
    ws.Range(EF_ADR_TOLERANCE).Value = e.TolerancePct
    ws.Range(EF_ADR_MESSAGE).Value = ""

    ' Garde-fous de saisie (dates, montant, tolérance), reposés à chaque ouverture.
    mod_InstallCommun.PoserValidationMontant ws.Range(EF_ADR_MONTANT)
    mod_InstallCommun.PoserValidationDate ws.Range(EF_ADR_DATE)
    mod_InstallCommun.PoserValidationDate ws.Range(EF_ADR_DATEFIN)
    mod_InstallCommun.PoserValidationMontant ws.Range(EF_ADR_TOLERANCE), mod_InstallCommun.MONTANT_POSITIF_OU_NUL

    ' Tiers : liste des Tiers connus dans les opérations (saisie libre acceptée : un mot-clé suffit).
    RemplirListeTiersEch ws
    ' Catégorie : liste stricte, comme dans les autres formulaires; le bouton + en crée une nouvelle.
    PoserListeCategorie ws
    ' Périodicité et précision : listes strictes.
    PoserListe ws.Range(EF_ADR_PERIODE), "=ListePeriodicites", True
    PoserListe ws.Range(EF_ADR_PRECISION), "=ListePrecisionsDate", True

    RemplirListeSousCatEch ws, e.Categorie

Sortie:
    numErr = Err.Number
    Application.EnableEvents = evenementsAvant
    If numErr <> 0 Then Err.Raise numErr, "RemplirChampsEch", Err.Description

End Sub

Private Sub PoserListe(ByVal cible As Range, ByVal formule As String, ByVal strict As Boolean, _
                       Optional ByVal titreErreur As String = "", Optional ByVal messageErreur As String = "")
    With cible.Validation
        .Delete
        If strict Then
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=formule
            .ShowError = True
            If titreErreur <> "" Then .ErrorTitle = titreErreur
            If messageErreur <> "" Then .ErrorMessage = messageErreur
        Else
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Formula1:=formule
            .ShowError = False
        End If
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

' Sous-catégories connues de la catégorie choisie (colonne technique masquée + nom dynamique).
Private Sub RemplirListeSousCatEch(ByVal ws As Worksheet, ByVal categorie As String)

    Dim liste() As String
    Dim sortie() As Variant
    Dim nb As Long, r As Long
    Dim colLettre As String

    ws.Range(ws.Cells(2, EF_COL_AIDE_SOUS), ws.Cells(EF_LIGNE_AIDE_MAX, EF_COL_AIDE_SOUS)).ClearContents

    ' Catégorie choisie : ses sous-catégories. Sinon : toutes les sous-catégories connues.
    If categorie <> "" Then
        liste = mod_Categories.ObtenirSousCategories(categorie)
    Else
        liste = ToutesSousCategories()
    End If
    On Error Resume Next
    nb = UBound(liste) - LBound(liste) + 1
    If Err.Number <> 0 Then nb = 0
    Err.Clear
    On Error GoTo 0

    If nb > 0 Then
        ReDim sortie(1 To nb, 1 To 1)
        For r = 1 To nb
            sortie(r, 1) = liste(LBound(liste) + r - 1)
        Next r
        With ws.Cells(2, EF_COL_AIDE_SOUS).Resize(nb, 1)
            .NumberFormat = "@"
            .Value2 = sortie
            .Sort Key1:=.Cells(1, 1), Order1:=xlAscending, Header:=xlNo
        End With
    End If

    colLettre = Split(ws.Cells(1, EF_COL_AIDE_SOUS).Address(True, False), "$")(0)
    ThisWorkbook.Names.Add Name:="ListeSousCatEcheance", _
        RefersTo:="=OFFSET(" & mod_VarGlobales.NOM_FEUILLE_ECHEANCE & "!$" & colLettre & "$2,0,0," & _
                  "MAX(1,COUNTA(" & mod_VarGlobales.NOM_FEUILLE_ECHEANCE & "!$" & colLettre & "$2:$" & colLettre & "$" & _
                  EF_LIGNE_AIDE_MAX & ")),1)"

    PoserListe ws.Range(EF_ADR_SOUS), "=ListeSousCatEcheance", False

End Sub

Private Sub AppliquerListesEch(ByVal ws As Worksheet)
    Dim evenementsAvant As Boolean
    evenementsAvant = Application.EnableEvents
    Application.EnableEvents = False
    On Error Resume Next
    PoserListe ws.Range(EF_ADR_TIERS), "=ListeTiersEcheance", False
    PoserListeCategorie ws
    PoserListe ws.Range(EF_ADR_SOUS), "=ListeSousCatEcheance", False
    PoserListe ws.Range(EF_ADR_PERIODE), "=ListePeriodicites", True
    PoserListe ws.Range(EF_ADR_PRECISION), "=ListePrecisionsDate", True
    On Error GoTo 0
    Application.EnableEvents = evenementsAvant
End Sub

Private Sub PoserListeCategorie(ByVal ws As Worksheet)
    PoserListe ws.Range(EF_ADR_CAT), "=ListeCategories", True, _
        mod_Display.FR("Cat{e2}gorie inconnue"), _
        mod_Display.FR("Choisissez une cat{e2}gorie dans la liste, ou laissez le champ vide.") & vbCrLf & _
        mod_Display.FR("Pour cr{e2}er une nouvelle cat{e2}gorie, utilisez le bouton '+'.")
End Sub

' Tiers distincts connus dans TblOperations, triés, écrits dans une colonne technique masquée
' (comme pour les sous-catégories) et reliés au champ Tiers par un nom dynamique.
Private Sub RemplirListeTiersEch(ByVal ws As Worksheet)

    Dim lo As ListObject
    Dim v() As Variant
    Dim vus As Object
    Dim sortie() As Variant
    Dim i As Long, n As Long
    Dim t As String, colLettre As String
    Dim rg As Range

    ws.Range(ws.Cells(2, EF_COL_AIDE_TIERS), ws.Cells(EF_LIGNE_AIDE_TIERS_MAX, EF_COL_AIDE_TIERS)).ClearContents

    On Error Resume Next
    Set lo = ThisWorkbook.Worksheets("Import_data").ListObjects("TblOperations")
    On Error GoTo 0

    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            v = LireColonne(lo, "Tiers")
            Set vus = CreateObject("Scripting.Dictionary")
            vus.CompareMode = 1     ' sans tenir compte des majuscules
            ReDim sortie(1 To EF_LIGNE_AIDE_TIERS_MAX - 1, 1 To 1)
            For i = 1 To UBound(v)
                t = mod_DataStructure.CellText(v(i))
                If t <> "" Then
                    If Not vus.Exists(t) Then
                        vus.Add t, True
                        If n < EF_LIGNE_AIDE_TIERS_MAX - 1 Then
                            n = n + 1
                            sortie(n, 1) = t
                        End If
                    End If
                End If
            Next i
        End If
    End If

    If n > 0 Then
        Set rg = ws.Cells(2, EF_COL_AIDE_TIERS).Resize(n, 1)
        rg.NumberFormat = "@"
        rg.Value2 = ArrayRognee(sortie, n)
        rg.Sort Key1:=rg.Cells(1, 1), Order1:=xlAscending, Header:=xlNo
    End If

    colLettre = Split(ws.Cells(1, EF_COL_AIDE_TIERS).Address(True, False), "$")(0)
    ThisWorkbook.Names.Add Name:="ListeTiersEcheance", _
        RefersTo:="=OFFSET(" & mod_VarGlobales.NOM_FEUILLE_ECHEANCE & "!$" & colLettre & "$2,0,0," & _
                  "MAX(1,COUNTA(" & mod_VarGlobales.NOM_FEUILLE_ECHEANCE & "!$" & colLettre & "$2:$" & colLettre & "$" & _
                  EF_LIGNE_AIDE_TIERS_MAX & ")),1)"

    PoserListe ws.Range(EF_ADR_TIERS), "=ListeTiersEcheance", False

End Sub

' Recopie les n premières lignes d'un tableau (1 colonne) dans un tableau de la bonne taille.
Private Function ArrayRognee(ByRef source() As Variant, ByVal n As Long) As Variant
    Dim r() As Variant
    Dim i As Long
    ReDim r(1 To n, 1 To 1)
    For i = 1 To n
        r(i, 1) = source(i, 1)
    Next i
    ArrayRognee = r
End Function

' --- Boutons "+" (à côté de Catégorie et de Sous-catégorie) ---
' Ouvre le formulaire de création d'une catégorie / sous-catégorie (frm_NouvelleCategorie),
' prérempli avec les valeurs actuelles, comme dans les autres formulaires. Si l'opérateur
' valide, la paire créée est reprise dans le formulaire.
Public Sub EchNouvelleCategorie()

    Dim ws As Worksheet
    Dim cat As String, sous As String, catRes As String, sousRes As String
    Dim ok As Boolean

    If Not g_EchEnCours Then Exit Sub
    On Error GoTo Erreur
    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_ECHEANCE)

    cat = mod_DataStructure.CellText(ws.Range(EF_ADR_CAT).Value)
    sous = mod_DataStructure.CellText(ws.Range(EF_ADR_SOUS).Value)

    ' On suspend le verrou d'activation pour que l'ouverture de l'autre formulaire ne
    ' ramène pas de force l'opérateur ici (même principe que mod_ControleCategories).
    g_EchVerrouActif = False
    ok = mod_NouvelleCategorie.OuvrirNouvelleCategorie(cat, sous, catRes, sousRes)
    g_EchVerrouActif = True
    ws.Activate

    If ok Then
        mod_Categories.RafraichirListesCategories
        Application.EnableEvents = False
        PoserListeCategorie ws
        ws.Range(EF_ADR_CAT).Value = catRes
        ws.Range(EF_ADR_SOUS).Value = sousRes
        RemplirListeSousCatEch ws, catRes
        Application.EnableEvents = True
    End If
    Exit Sub

Erreur:
    g_EchVerrouActif = True
    Application.EnableEvents = True
    MsgBox mod_Display.FR("Erreur dans EchNouvelleCategorie : ") & Err.Number & " - " & Err.Description, vbCritical

End Sub


' Appelée par le code de la feuille (Worksheet_Change) : si la catégorie change, l'ancienne
' sous-catégorie n'a plus de sens.
Public Sub EchTraiterChangement(ByVal ws As Worksheet, ByVal Target As Range)

    Dim numErr As Long
    Dim catTrouvee As String

    If Not g_EchEnCours Then Exit Sub
    If Target.Cells.Count > 1 Then Exit Sub

    On Error GoTo Sortie

    If Target.Address(False, False) = EF_ADR_CAT Then
        Application.EnableEvents = False
        ws.Range(EF_ADR_SOUS).ClearContents
        RemplirListeSousCatEch ws, mod_DataStructure.CellText(Target.Value)

    ElseIf Target.Address(False, False) = EF_ADR_SOUS Then
        ' Sous-catégorie choisie alors qu'aucune catégorie n'est renseignée : si cette
        ' sous-catégorie n'appartient qu'à UNE catégorie, on renseigne celle-ci.
        If mod_DataStructure.CellText(ws.Range(EF_ADR_CAT).Value) = "" And mod_DataStructure.CellText(Target.Value) <> "" Then
            catTrouvee = CategorieUniquePour(mod_DataStructure.CellText(Target.Value))
            If catTrouvee <> "" Then
                Application.EnableEvents = False
                ws.Range(EF_ADR_CAT).Value = catTrouvee
                RemplirListeSousCatEch ws, catTrouvee
            End If
        End If
    End If

Sortie:
    numErr = Err.Number
    Application.EnableEvents = True
    If numErr <> 0 Then MsgBox mod_Display.FR("Erreur lors du changement de cat{e2}gorie : ") & Err.Description, vbExclamation

End Sub

' Table des catégories / sous-catégories (feuille Param).
Private Function TableCategories() As ListObject
    On Error Resume Next
    Set TableCategories = ThisWorkbook.Worksheets(mod_Categories.NOM_FEUILLE_PARAM).ListObjects("TblCategories")
    On Error GoTo 0
End Function

' Toutes les sous-catégories distinctes connues (tableau non alloué si aucune).
Private Function ToutesSousCategories() As String()

    Dim lo As ListObject
    Dim v() As Variant
    Dim vus As Object
    Dim liste() As String
    Dim i As Long, n As Long
    Dim t As String

    Set lo = TableCategories()
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    v = LireColonne(lo, "SousCategorie")
    Set vus = CreateObject("Scripting.Dictionary")
    vus.CompareMode = 1
    ReDim liste(1 To UBound(v))
    For i = 1 To UBound(v)
        t = mod_DataStructure.CellText(v(i))
        If t <> "" Then
            If Not vus.Exists(t) Then
                vus.Add t, True
                n = n + 1
                If n < EF_LIGNE_AIDE_MAX Then liste(n) = t Else n = n - 1
            End If
        End If
    Next i
    If n > 0 Then
        ReDim Preserve liste(1 To n)
        ToutesSousCategories = liste
    End If

End Function

' Catégorie à laquelle appartient une sous-catégorie, si elle n'en a qu'une ("" sinon).
Private Function CategorieUniquePour(ByVal sous As String) As String

    Dim lo As ListObject
    Dim vCat() As Variant, vSous() As Variant
    Dim i As Long
    Dim trouvee As String, c As String

    Set lo = TableCategories()
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    vCat = LireColonne(lo, "Categorie")
    vSous = LireColonne(lo, "SousCategorie")
    For i = 1 To UBound(vSous)
        If StrComp(mod_DataStructure.CellText(vSous(i)), sous, vbTextCompare) = 0 Then
            c = mod_DataStructure.CellText(vCat(i))
            If trouvee = "" Then
                trouvee = c
            ElseIf StrComp(trouvee, c, vbTextCompare) <> 0 Then
                Exit Function      ' plusieurs catégories possibles : on ne choisit pas à la place de l'opérateur
            End If
        End If
    Next i
    CategorieUniquePour = trouvee

End Function

' Verrou "modal" : tant que le formulaire est ouvert, impossible de le quitter en changeant d'onglet.
Public Sub EchVerrouiller(ByVal ws As Worksheet)
    If g_EchVerrouActif And Not mod_Aide.AideOuverte() Then ws.Activate
End Sub

' --- Bouton "Valider" ---
Public Sub EchValider()

    Dim ws As Worksheet
    Dim lib As String, tiers As String, cat As String, sous As String, periode As String, prec As String
    Dim vMont As Variant, vDate As Variant, vFin As Variant, vTol As Variant
    Dim nMont As Double, nTol As Double
    Dim dPrevue As Date, dFinLue As Date
    Dim chq As String, chqNorm As String, trouve As String, k As Long, chiffres As Boolean
    Dim listeCat() As String, listeSous() As String
    Dim e As TypeEcheance

    If Not g_EchEnCours Then Exit Sub
    On Error GoTo Erreur
    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_ECHEANCE)

    lib = mod_DataStructure.CellText(ws.Range(EF_ADR_LIBELLE).Value)
    tiers = mod_DataStructure.CellText(ws.Range(EF_ADR_TIERS).Value)
    cat = mod_DataStructure.CellText(ws.Range(EF_ADR_CAT).Value)
    sous = mod_DataStructure.CellText(ws.Range(EF_ADR_SOUS).Value)
    vMont = ws.Range(EF_ADR_MONTANT).Value
    chq = mod_DataStructure.CellText(ws.Range(EF_ADR_CHEQUE).Value)
    periode = mod_DataStructure.CellText(ws.Range(EF_ADR_PERIODE).Value)
    vDate = ws.Range(EF_ADR_DATE).Value
    vFin = ws.Range(EF_ADR_DATEFIN).Value
    prec = mod_DataStructure.CellText(ws.Range(EF_ADR_PRECISION).Value)
    vTol = ws.Range(EF_ADR_TOLERANCE).Value

    If tiers = "" And cat = "" And chq = "" Then
        AfficherMessageEch ws, "Renseignez au moins le Tiers, la Cat{e2}gorie ou le num{e2}ro de ch{e1}que : c'est ce qui permettra de reconna{i2}tre l'op{e2}ration quand elle sera import{e2}e."
        Exit Sub
    End If
    If cat = "" And sous <> "" Then
        AfficherMessageEch ws, "Une sous-cat{e2}gorie suppose une cat{e2}gorie."
        Exit Sub
    End If
    If cat <> "" Then
        listeCat = mod_Categories.ObtenirCategories()
        If Not ExisteDansListe(listeCat, cat) Then
            AfficherMessageEch ws, "Cat{e2}gorie inconnue : choisissez-la dans la liste (pour en cr{e2}er une, utilisez le bouton +)."
            Exit Sub
        End If
        If sous <> "" Then
            listeSous = mod_Categories.ObtenirSousCategories(cat)
            If Not ExisteDansListe(listeSous, sous) Then
                AfficherMessageEch ws, "Sous-cat{e2}gorie inconnue pour cette cat{e2}gorie : choisissez-la dans la liste (pour en cr{e2}er une, utilisez le bouton +)."
                Exit Sub
            End If
        End If
    End If
    If IsEmpty(vMont) Or IsError(vMont) Then
        AfficherMessageEch ws, "Le montant est obligatoire (n{e2}gatif pour une d{e2}pense, positif pour une entr{e2}e d'argent)."
        Exit Sub
    End If
    If Not mod_DataStructure.TryLireMontant(vMont, nMont) Then
        AfficherMessageEch ws, "Le montant n'est pas valide. " & mod_DataStructure.MessageMontantAttendu()
        Exit Sub
    End If
    If nMont = 0 Then
        AfficherMessageEch ws, "Le montant ne peut pas {ea}tre nul."
        Exit Sub
    End If
    ' Périodicité vide = opération unique (ponctuelle).
    If periode = "" Then periode = PERIODE_PONCTUELLE
    Select Case periode
        Case PERIODE_PONCTUELLE, PERIODE_HEBDO, PERIODE_MENSUELLE, PERIODE_TRIMESTRIELLE, PERIODE_ANNUELLE
        Case Else
            AfficherMessageEch ws, "Choisissez une p{e2}riodicit{e2} dans la liste."
            Exit Sub
    End Select
    If chq <> "" Then
        If periode <> PERIODE_PONCTUELLE Then
            AfficherMessageEch ws, "Un num{e2}ro de ch{e1}que ne concerne qu'une op{e2}ration unique : effacez-le, ou laissez la p{e2}riodicit{e2} vide."
            Exit Sub
        End If
        chiffres = True
        For k = 1 To Len(chq)
            If Mid$(chq, k, 1) < "0" Or Mid$(chq, k, 1) > "9" Then chiffres = False
        Next k
        If Not chiffres Or Len(chq) > 10 Then
            AfficherMessageEch ws, "Le num{e2}ro de ch{e1}que ne doit contenir que des chiffres (10 au maximum), sans espace."
            Exit Sub
        End If
        chqNorm = NormaliserCheque(chq)
        If chqNorm = "" Then
            AfficherMessageEch ws, "Ce num{e2}ro de ch{e1}que n'est pas valide (il ne peut pas {ea}tre z{e2}ro)."
            Exit Sub
        End If
        trouve = ChequeDejaPasse(chqNorm)
        If trouve <> "" Then
            AfficherMessageEch ws, "Ce ch{e1}que est d{e2}j{a2} pass{e2} en banque (op{e2}ration du " & trouve & ")."
            Exit Sub
        End If
        trouve = ChequeDansEcheancier(chqNorm)
        If trouve <> "" Then
            AfficherMessageEch ws, "Ce num{e2}ro de ch{e1}que est d{e2}j{a2} saisi dans l'{e2}ch{e2}ancier (" & trouve & ")."
            Exit Sub
        End If
    End If
    If prec <> PRECISION_JOUR And prec <> PRECISION_MOIS Then
        AfficherMessageEch ws, "Choisissez la pr{e2}cision de la date dans la liste (Jour ou Mois)."
        Exit Sub
    End If
    If IsEmpty(vDate) Or IsError(vDate) Then
        AfficherMessageEch ws, "La date est obligatoire."
        Exit Sub
    End If
    If Not mod_DataStructure.TryLireDate(vDate, dPrevue) Then
        AfficherMessageEch ws, "La date n'est pas valide. " & mod_DataStructure.MessageDateAttendue()
        Exit Sub
    End If
    If Not IsEmpty(vFin) Then
        If periode = PERIODE_PONCTUELLE Then
            AfficherMessageEch ws, "Une date de fin n'a de sens que pour une op{e2}ration r{e2}currente : effacez-la, ou choisissez une p{e2}riodicit{e2}."
            Exit Sub
        End If
        If Not mod_DataStructure.TryLireDate(vFin, dFinLue) Then
            AfficherMessageEch ws, "La date de fin n'est pas valide. " & mod_DataStructure.MessageDateAttendue()
            Exit Sub
        End If
        If dFinLue < dPrevue Then
            AfficherMessageEch ws, "La date de fin ne peut pas {ea}tre ant{e2}rieure {a2} la date pr{e2}vue."
            Exit Sub
        End If
    End If
    If Not IsEmpty(vTol) Then
        If Not mod_DataStructure.TryLireMontant(vTol, nTol) Then
            AfficherMessageEch ws, "La tol{e2}rance doit {ea}tre un nombre (pourcentage, exemple : 10)."
            Exit Sub
        End If
        If nTol < 0 Or nTol > 100 Then
            AfficherMessageEch ws, "La tol{e2}rance doit {ea}tre comprise entre 0 et 100."
            Exit Sub
        End If
    End If

    e.Tiers = tiers
    e.Categorie = cat
    e.SousCategorie = sous
    If lib <> "" Then
        e.Libelle = lib
    ElseIf tiers <> "" Then
        e.Libelle = tiers
    ElseIf sous <> "" Then
        e.Libelle = sous
    Else
        e.Libelle = cat
    End If
    e.Montant = nMont
    e.NumCheque = chq
    e.Periodicite = periode
    e.DatePrevue = dPrevue
    If IsEmpty(vFin) Then e.DateFin = 0 Else e.DateFin = dFinLue
    e.PrecisionDate = prec
    If IsEmpty(vTol) Then e.TolerancePct = TOLERANCE_DEFAUT Else e.TolerancePct = nTol

    ' Garde-fou doublon : l'opérateur peut passer outre (deux achats identiques le même jour existent).
    trouve = DoublonEcheancier(e)
    If trouve <> "" Then
        If MsgBox(mod_Display.FR("Une op{e2}ration identique existe d{e2}j{a2} dans l'{e2}ch{e2}ancier (" & trouve & _
                  ") : m{ea}me tiers, cat{e2}gorie, sous-cat{e2}gorie et montant {a2} " & DOUBLON_TOLERANCE_PCT & " % pr{e1}s (montant saisi : " & Format(e.Montant, "#,##0.00") & " " & ChrW(8364) & ")." & vbCrLf & vbCrLf & _
                  "Voulez-vous quand m{ea}me l'ajouter ?"), vbQuestion + vbYesNo + vbDefaultButton2, _
                  mod_Display.FR("Doublon possible")) <> vbYes Then
            AfficherMessageEch ws, "Doublon {e2}vit{e2} : modifiez la saisie ou annulez."
            Exit Sub
        End If
    End If

    g_EchRes = e
    g_EchValide = True
    FermerFeuilleEch ws
    Exit Sub

Erreur:
    Application.EnableEvents = True
    MsgBox mod_Display.FR("Erreur dans EchValider : ") & Err.Number & " - " & Err.Description, vbCritical

End Sub

' --- Bouton "Annuler" ---
Public Sub EchAnnuler()

    Dim ws As Worksheet

    If Not g_EchEnCours Then Exit Sub
    On Error GoTo Erreur

    If MsgBox(mod_Display.FR("Annuler la saisie ? Rien ne sera enregistr{e2}."), vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_ECHEANCE)
    g_EchValide = False
    FermerFeuilleEch ws
    Exit Sub

Erreur:
    Application.EnableEvents = True
    MsgBox mod_Display.FR("Erreur dans EchAnnuler : ") & Err.Number & " - " & Err.Description, vbCritical

End Sub

Private Sub AfficherMessageEch(ByVal ws As Worksheet, ByVal texte As String)
    ws.Range(EF_ADR_MESSAGE).Value = mod_Display.FR(texte)
End Sub

' g_EchEnCours doit passer à False AVANT de masquer la feuille, sinon le verrou la réactiverait.
Private Sub FermerFeuilleEch(ByVal ws As Worksheet)
    g_EchEnCours = False
    g_EchVerrouActif = False
    ws.Range(ws.Cells(2, EF_COL_AIDE_SOUS), ws.Cells(EF_LIGNE_AIDE_MAX, EF_COL_AIDE_SOUS)).ClearContents
    ws.Range(ws.Cells(2, EF_COL_AIDE_TIERS), ws.Cells(EF_LIGNE_AIDE_TIERS_MAX, EF_COL_AIDE_TIERS)).ClearContents
    mod_Aide.MasquerFormulaire ws
End Sub

Private Function ExisteDansListe(ByRef liste() As String, ByVal valeur As String) As Boolean
    Dim i As Long
    On Error GoTo Fin
    For i = LBound(liste) To UBound(liste)
        If StrComp(liste(i), valeur, vbTextCompare) = 0 Then
            ExisteDansListe = True
            Exit Function
        End If
    Next i
Fin:
End Function


' =====================================================================================
' FEUILLES DE LISTE (échéances à venir / règles de récurrence)
' =====================================================================================
Public Sub AfficherEcheancier()
    AfficherListe mod_VarGlobales.NOM_FEUILLE_ECHEANCIER
End Sub

Public Sub AfficherRecurrences()
    AfficherListe mod_VarGlobales.NOM_FEUILLE_RECURRENCES
End Sub

Private Sub AfficherListe(ByVal nomFeuille As String)

    Dim ws As Worksheet, autre As Worksheet
    Dim nomAutre As String

    Set ws = mod_InstallCommun.TrouverFeuille(nomFeuille)
    If ws Is Nothing Then
        MsgBox mod_Display.FR("L'{e2}ch{e2}ancier n'est pas install{e2}. Ex{e2}cutez d'abord la macro InstallerEcheancier."), vbExclamation
        Exit Sub
    End If

    ' Les occurrences sont complétées avant l'affichage.
    GenererHorizon

    If nomFeuille = mod_VarGlobales.NOM_FEUILLE_ECHEANCIER Then
        nomAutre = mod_VarGlobales.NOM_FEUILLE_RECURRENCES
    Else
        nomAutre = mod_VarGlobales.NOM_FEUILLE_ECHEANCIER
    End If

    ws.Visible = xlSheetVisible
    ws.Activate
    mod_InstallCommun.MasquerQuadrillage

    Set autre = mod_InstallCommun.TrouverFeuille(nomAutre)
    If Not autre Is Nothing Then mod_Aide.MasquerFormulaire autre

End Sub

Public Sub SortirEcheancier()

    Dim ws As Worksheet

    Set ws = ActiveSheet
    If ws.Name <> mod_VarGlobales.NOM_FEUILLE_ECHEANCIER And ws.Name <> mod_VarGlobales.NOM_FEUILLE_RECURRENCES Then Exit Sub

    On Error Resume Next
    mod_InstallCommun.AfficherSynthese
    On Error GoTo 0
    mod_Aide.MasquerFormulaire ws

End Sub

' Excel refuse de supprimer une ligne d'un tableau filtré (erreur 1004 : impossible de déplacer
' les cellules d'une plage filtrée). Les filtres de colonne sont donc levés avant toute
' suppression ; les numéros de ligne restent valables (ils sont calculés sur l'ordre du tableau).
Private Sub LeverFiltre(ByVal lo As ListObject)

    Dim memo As Variant

    If mFiltresMemorises Is Nothing Then Set mFiltresMemorises = CreateObject("Scripting.Dictionary")
    If mFiltresMemorises.Exists(lo.Name) Then mFiltresMemorises.Remove lo.Name
    If Not lo.ShowAutoFilter Then Exit Sub
    If Not lo.AutoFilter.FilterMode Then Exit Sub

    memo = MemoriserFiltres(lo)
    If Not IsEmpty(memo) Then mFiltresMemorises.Add lo.Name, memo
    lo.AutoFilter.ShowAllData

End Sub

' Mémorise les filtres de colonne actifs : une colonne par filtre, 4 lignes
' (n° de colonne, opérateur, critère 1, critère 2). Certains filtres (regroupement de dates
' par année/mois, par exemple) refusent la lecture de leurs critères (erreur 1004) : on
' garde alors les valeurs actuellement visibles de la colonne, réappliquées comme une liste
' de valeurs cochées (opérateur 0 dans la mémoire, voir RestaurerFiltre).
Private Function MemoriserFiltres(ByVal lo As ListObject) As Variant

    Dim res() As Variant
    Dim n As Long, f As Long
    Dim fl As Filter
    Dim op As Variant, c1 As Variant, c2 As Variant
    Dim ok As Boolean
    Dim vus As Object, c As Range
    Dim cles As Variant

    For f = 1 To lo.ListColumns.Count
        Set fl = Nothing
        On Error Resume Next
        Set fl = lo.AutoFilter.Filters(f)
        On Error GoTo 0
        If Not fl Is Nothing Then
            If fl.On Then
                n = n + 1
                ReDim Preserve res(1 To 4, 1 To n)
                res(1, n) = f
                op = Empty: c1 = Empty: c2 = Empty
                ok = True
                On Error Resume Next
                op = fl.Operator
                c1 = fl.Criteria1
                If Err.Number <> 0 Then ok = False
                Err.Clear
                c2 = fl.Criteria2
                Err.Clear
                On Error GoTo 0
                If ok Then
                    res(2, n) = op
                    res(3, n) = c1
                    res(4, n) = c2
                Else
                    Set vus = CreateObject("Scripting.Dictionary")
                    If Not lo.ListColumns(f).DataBodyRange Is Nothing Then
                        For Each c In lo.ListColumns(f).DataBodyRange.SpecialCells(xlCellTypeVisible).Cells
                            If Not vus.Exists(c.Text) Then vus.Add c.Text, True
                        Next c
                    End If
                    cles = vus.Keys
                    res(2, n) = -1
                    res(3, n) = cles
                    res(4, n) = Empty
                End If
            End If
        End If
    Next f

    If n > 0 Then MemoriserFiltres = res

End Function

' Remet, une fois les suppressions faites, les filtres de colonne levés par LeverFiltre
' (même principe que l'écran de recherche : l'opérateur retrouve son filtre).
Private Sub RestaurerFiltre(ByVal lo As ListObject)

    Dim memo As Variant
    Dim k As Long
    Dim c1 As Variant

    If mFiltresMemorises Is Nothing Then Exit Sub
    If Not mFiltresMemorises.Exists(lo.Name) Then Exit Sub
    memo = mFiltresMemorises(lo.Name)
    mFiltresMemorises.Remove lo.Name

    For k = 1 To UBound(memo, 2)
        c1 = memo(3, k)
        On Error Resume Next
        If memo(2, k) = -1 Then
            lo.Range.AutoFilter Field:=memo(1, k), Criteria1:=c1, Operator:=xlFilterValues
        ElseIf memo(2, k) = 0 Then
            lo.Range.AutoFilter Field:=memo(1, k), Criteria1:=c1
        ElseIf IsEmpty(memo(4, k)) Then
            If memo(2, k) = xlFilterValues And Not IsArray(c1) Then c1 = Array(c1)
            lo.Range.AutoFilter Field:=memo(1, k), Criteria1:=c1, Operator:=memo(2, k)
        Else
            lo.Range.AutoFilter Field:=memo(1, k), Criteria1:=c1, Operator:=memo(2, k), Criteria2:=memo(4, k)
        End If
        On Error GoTo 0
    Next k

End Sub

' Numéros (dans le tableau) des lignes touchées par la sélection courante.
Private Function LignesSelectionnees(ByVal lo As ListObject) As Object

    Dim d As Object
    Dim inter As Range, c As Range
    Dim n As Long

    Set d = CreateObject("Scripting.Dictionary")
    Set LignesSelectionnees = d
    If lo.DataBodyRange Is Nothing Then Exit Function

    On Error Resume Next
    Set inter = Intersect(Selection, lo.DataBodyRange)
    On Error GoTo 0
    If inter Is Nothing Then Exit Function

    ' Les lignes masquées par un filtre de colonne sont ignorées : une sélection qui les
    ' enjambe ne doit pas les supprimer.
    For Each c In inter.Cells
        If Not c.EntireRow.Hidden Then
            n = c.Row - lo.DataBodyRange.Row + 1
            If Not d.Exists(n) Then d.Add n, True
        End If
    Next c

End Function

' Bouton "Supprimer la sélection" de la liste des échéances.
' - Échéances ponctuelles : suppression simple.
' - Échéances issues d'une règle : l'opérateur choisit entre supprimer seulement ces
'   occurrences (la règle continue) ou ARRÊTER la récurrence à partir de la première
'   occurrence sélectionnée (la règle reçoit une date de fin, les occurrences suivantes
'   disparaissent) : cas d'un remboursement anticipé.
Public Sub SupprimerEcheanceSelection()

    Dim lo As ListObject, loR As ListObject
    Dim sel As Object
    Dim i As Long, nbRegle As Long, cIdRec As Long
    Dim v As Variant

    Set lo = TableEcheances()
    Set loR = TableRecurrences()
    If lo Is Nothing Or loR Is Nothing Then Exit Sub
    If ActiveSheet.Name <> mod_VarGlobales.NOM_FEUILLE_ECHEANCIER Then Exit Sub

    Set sel = LignesSelectionnees(lo)
    If sel.Count = 0 Then
        MsgBox mod_Display.FR("S{e2}lectionnez d'abord une ou plusieurs lignes du tableau."), vbInformation
        Exit Sub
    End If

    cIdRec = ColIdx(lo, "ID_Recurrence")
    For i = 1 To lo.ListRows.Count
        If sel.Exists(i) Then
            If mod_DataStructure.CellText(lo.ListRows(i).Range.Cells(1, cIdRec).Value2) <> "" Then nbRegle = nbRegle + 1
        End If
    Next i

    If nbRegle = 0 Then
        If MsgBox(sel.Count & mod_Display.FR(" {e2}ch{e2}ance(s) {a2} supprimer ?"), vbYesNo + vbQuestion, _
                  mod_Display.FR("Supprimer des {e2}ch{e2}ances")) = vbNo Then Exit Sub
        LeverFiltre lo
        For i = lo.ListRows.Count To 1 Step -1
            If sel.Exists(i) Then lo.ListRows(i).Delete
        Next i
        RestaurerFiltre lo
        Exit Sub
    End If

    Select Case MsgBox(nbRegle & mod_Display.FR(" des {e2}ch{e2}ances s{e2}lectionn{e2}es viennent d'une r{e1}gle de r{e2}currence.") & vbCrLf & vbCrLf & _
                       mod_Display.FR("OUI : supprimer seulement ces {e2}ch{e2}ances. La r{e1}gle continue (si elle n'a plus aucune {e2}ch{e2}ance dans la liste, la suivante est recr{e2}{e2}e aussit{o2}t).") & vbCrLf & vbCrLf & _
                       mod_Display.FR("NON : arr{ea}ter la r{e2}currence {a2} partir de la premi{e1}re {e2}ch{e2}ance s{e2}lectionn{e2}e : ces {e2}ch{e2}ances disparaissent et la r{e1}gle ne produira plus rien {a2} partir de cette date.") & vbCrLf & vbCrLf & _
                       mod_Display.FR("ANNULER : ne rien faire."), _
                       vbYesNoCancel + vbQuestion, mod_Display.FR("{E2}ch{e2}ances issues d'une r{e1}gle"))
        Case vbYes
            LeverFiltre lo
            For i = lo.ListRows.Count To 1 Step -1
                If sel.Exists(i) Then lo.ListRows(i).Delete
            Next i
            RestaurerFiltre lo
            GenererHorizon      ' une règle qui n'a plus d'occurrence en retrouve une
        Case vbNo
            ArreterRecurrences lo, loR, sel
    End Select

End Sub

' Arrête les récurrences des échéances sélectionnées : pour chaque règle concernée, date de
' fin = veille de la première occurrence sélectionnée, et suppression de toutes ses
' occurrences à partir de cette date. Les échéances ponctuelles sélectionnées sont supprimées aussi.
Private Sub ArreterRecurrences(ByVal lo As ListObject, ByVal loR As ListObject, ByVal sel As Object)

    Dim coupures As Object
    Dim aSupprimer() As Boolean
    Dim i As Long, j As Long, cIdRec As Long, cDatePrev As Long, cIdR As Long, cFin As Long, cLib As Long
    Dim v As Variant, cle As Variant
    Dim d As Date
    Dim msg As String, libelle As String

    cIdRec = ColIdx(lo, "ID_Recurrence")
    cDatePrev = ColIdx(lo, "DatePrevue")

    ' 1. date de coupure de chaque règle = plus petite date sélectionnée
    Set coupures = CreateObject("Scripting.Dictionary")
    For i = 1 To lo.ListRows.Count
        If sel.Exists(i) Then
            v = lo.ListRows(i).Range.Cells(1, cIdRec).Value2
            If mod_DataStructure.CellText(v) <> "" Then
                d = VersDate(lo.ListRows(i).Range.Cells(1, cDatePrev).Value2)
                cle = CStr(CLng(v))
                If Not coupures.Exists(cle) Then
                    coupures.Add cle, CDbl(d)
                ElseIf CDbl(d) < coupures(cle) Then
                    coupures(cle) = CDbl(d)
                End If
            End If
        End If
    Next i

    ' 2. confirmation
    cIdR = ColIdx(loR, "ID_Recurrence")
    cLib = ColIdx(loR, "Libelle")
    cFin = ColIdx(loR, "DateFin")
    msg = mod_Display.FR("La r{e2}currence va {ea}tre arr{ea}t{e2}e :") & vbCrLf & vbCrLf
    For Each cle In coupures.Keys
        libelle = "(r" & ChrW(232) & "gle " & cle & ")"
        For j = 1 To loR.ListRows.Count
            If mod_DataStructure.CellText(loR.ListRows(j).Range.Cells(1, cIdR).Value2) = CStr(cle) Then
                libelle = mod_DataStructure.CellText(loR.ListRows(j).Range.Cells(1, cLib).Value2)
                Exit For
            End If
        Next j
        msg = msg & "- " & libelle & mod_Display.FR(" : plus aucune {e2}ch{e2}ance {a2} partir du ") & Format(CDate(coupures(cle)), "dd/mm/yyyy") & vbCrLf
    Next cle
    msg = msg & vbCrLf & mod_Display.FR("Continuer ?")
    If MsgBox(msg, vbYesNo + vbQuestion, mod_Display.FR("Arr{e2}ter la r{e2}currence")) = vbNo Then Exit Sub

    ' 3. date de fin des règles
    For Each cle In coupures.Keys
        For j = 1 To loR.ListRows.Count
            If mod_DataStructure.CellText(loR.ListRows(j).Range.Cells(1, cIdR).Value2) = CStr(cle) Then
                d = CDate(coupures(cle)) - 1
                v = VersDate(loR.ListRows(j).Range.Cells(1, cFin).Value2)
                If v = 0 Or d < v Then loR.ListRows(j).Range.Cells(1, cFin).Value = d
                Exit For
            End If
        Next j
    Next cle

    ' 4. suppression : occurrences de ces règles à partir de la coupure + échéances ponctuelles sélectionnées
    ReDim aSupprimer(1 To lo.ListRows.Count)
    For i = 1 To lo.ListRows.Count
        v = lo.ListRows(i).Range.Cells(1, cIdRec).Value2
        If mod_DataStructure.CellText(v) <> "" Then
            cle = CStr(CLng(v))
            If coupures.Exists(cle) Then
                If CDbl(VersDate(lo.ListRows(i).Range.Cells(1, cDatePrev).Value2)) >= coupures(cle) Then aSupprimer(i) = True
            End If
        ElseIf sel.Exists(i) Then
            aSupprimer(i) = True
        End If
    Next i
    LeverFiltre lo
    For i = lo.ListRows.Count To 1 Step -1
        If aSupprimer(i) Then lo.ListRows(i).Delete
    Next i
    RestaurerFiltre lo

    GenererHorizon      ' sans effet pour ces règles : leur date de fin est posée

End Sub

' Bouton "Supprimer la règle" de la liste des récurrences (supprime aussi ses occurrences à venir).
Public Sub SupprimerRecurrenceSelection()

    Dim loR As ListObject, loE As ListObject
    Dim sel As Object, ids As Object
    Dim i As Long, cId As Long, cIdRec As Long
    Dim v As Variant

    Set loR = TableRecurrences()
    Set loE = TableEcheances()
    If loR Is Nothing Or loE Is Nothing Then Exit Sub
    If ActiveSheet.Name <> mod_VarGlobales.NOM_FEUILLE_RECURRENCES Then Exit Sub

    Set sel = LignesSelectionnees(loR)
    If sel.Count = 0 Then
        MsgBox mod_Display.FR("S{e2}lectionnez d'abord une ou plusieurs lignes du tableau."), vbInformation
        Exit Sub
    End If
    If MsgBox(sel.Count & mod_Display.FR(" r{e2}gle(s) {a2} supprimer, ainsi que leurs {e2}ch{e2}ances {a2} venir ?"), _
              vbYesNo + vbQuestion, mod_Display.FR("Supprimer des r{e2}gles")) = vbNo Then Exit Sub

    Set ids = CreateObject("Scripting.Dictionary")
    cId = ColIdx(loR, "ID_Recurrence")
    For i = 1 To loR.ListRows.Count
        If sel.Exists(i) Then
            v = loR.ListRows(i).Range.Cells(1, cId).Value2
            If mod_DataStructure.CellText(v) <> "" Then ids(CStr(CLng(v))) = True
        End If
    Next i

    ' Occurrences liées (de bas en haut)
    If Not TableVide(loE) Then
        cIdRec = ColIdx(loE, "ID_Recurrence")
        LeverFiltre loE
        For i = loE.ListRows.Count To 1 Step -1
            v = loE.ListRows(i).Range.Cells(1, cIdRec).Value2
            If mod_DataStructure.CellText(v) <> "" Then
                If IsNumeric(v) Then
                    If ids.Exists(CStr(CLng(v))) Then loE.ListRows(i).Delete
                End If
            End If
        Next i
        RestaurerFiltre loE
    End If

    LeverFiltre loR
    For i = loR.ListRows.Count To 1 Step -1
        If sel.Exists(i) Then loR.ListRows(i).Delete
    Next i
    RestaurerFiltre loR

End Sub

' Bouton "Actualiser" de la liste des règles : après modification d'une règle (montant,
' tiers, périodicité, actif...), ses occurrences déjà générées sont recréées avec les
' nouvelles valeurs. La date de départ est celle de la plus ancienne occurrence encore
' dans la liste (donc les occurrences non rapprochées restent à venir).
Public Sub ActualiserRecurrences()

    Dim loR As ListObject, loE As ListObject
    Dim minDate As Object
    Dim i As Long, cIdRec As Long, cDatePrev As Long, cId As Long, cDateProch As Long
    Dim v As Variant, d As Date
    Dim cle As String

    Set loR = TableRecurrences()
    Set loE = TableEcheances()
    If loR Is Nothing Or loE Is Nothing Then Exit Sub

    If Not TableVide(loE) Then
        Set minDate = CreateObject("Scripting.Dictionary")
        cIdRec = ColIdx(loE, "ID_Recurrence")
        cDatePrev = ColIdx(loE, "DatePrevue")

        ' 1. plus petite date par règle
        For i = 1 To loE.ListRows.Count
            v = loE.ListRows(i).Range.Cells(1, cIdRec).Value2
            If mod_DataStructure.CellText(v) <> "" Then
                If IsNumeric(v) Then
                    cle = CStr(CLng(v))
                    d = VersDate(loE.ListRows(i).Range.Cells(1, cDatePrev).Value2)
                    If d <> 0 Then
                        If Not minDate.Exists(cle) Then
                            minDate.Add cle, CDbl(d)
                        ElseIf CDbl(d) < minDate(cle) Then
                            minDate(cle) = CDbl(d)
                        End If
                    End If
                End If
            End If
        Next i

        ' 2. suppression de ces occurrences (de bas en haut)
        LeverFiltre loE
        For i = loE.ListRows.Count To 1 Step -1
            v = loE.ListRows(i).Range.Cells(1, cIdRec).Value2
            If mod_DataStructure.CellText(v) <> "" Then
                If IsNumeric(v) Then
                    If minDate.Exists(CStr(CLng(v))) Then loE.ListRows(i).Delete
                End If
            End If
        Next i
        RestaurerFiltre loE

        ' 3. la règle repart de sa plus ancienne occurrence
        cId = ColIdx(loR, "ID_Recurrence")
        cDateProch = ColIdx(loR, "DateProchaine")
        For i = 1 To loR.ListRows.Count
            v = loR.ListRows(i).Range.Cells(1, cId).Value2
            If mod_DataStructure.CellText(v) <> "" Then
                If IsNumeric(v) Then
                    cle = CStr(CLng(v))
                    If minDate.Exists(cle) Then loR.ListRows(i).Range.Cells(1, cDateProch).Value = CDate(minDate(cle))
                End If
            End If
        Next i
    End If

    GenererHorizon
    MsgBox mod_Display.FR("Les occurrences des r{e2}gles ont {e2}t{e2} recalcul{e2}es."), vbInformation, mod_Display.FR("R{e2}currences")

End Sub


' =====================================================================================
' SOLDE DE RÉFÉRENCE AVEC CONTRÔLE D'ÉCART (bouton de la feuille Synthese)
' =====================================================================================
' L'opérateur saisit le solde réel du compte (tel que la banque l'affiche) et sa date.
' Avant d'enregistrer, le programme calcule le solde que le classeur attendait à cette date
' (ancien solde de référence + opérations enregistrées depuis) et affiche l'ÉCART avec le
' solde saisi : un écart non nul révèle une opération manquante ou en double, une date de
' solde mal choisie, ou une opération modifiée depuis. L'opérateur décide ensuite
' d'enregistrer ou non le nouveau solde. Le dernier écart reste affiché sur Synthese.
Public Sub SaisirSoldeReference()

    Dim ws As Worksheet
    Dim loOps As ListObject
    Dim ancienM As Variant, ancienD As Variant
    Dim sMont As String, sDate As String, defMont As String
    Dim nouveauM As Double, calcule As Double, ecart As Double
    Dim nouvelleD As Date, ancD As Date
    Dim controle As Boolean
    Dim msg As String
    Dim titre As String

    titre = mod_Display.FR("Solde de r{e2}f{e2}rence")

    Set ws = mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_SYNTHESE)
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    Set loOps = ThisWorkbook.Worksheets("Import_data").ListObjects("TblOperations")
    ancienM = ws.Range("soldeRefMontant").Value
    ancienD = ws.Range("soldeRefDate").Value
    On Error GoTo 0
    If Not IsEmpty(ancienM) Then
        If IsNumeric(ancienM) Then defMont = Format(CDbl(ancienM), "0.00")
    End If

    sMont = InputBox(mod_Display.FR("Solde du compte tel que votre banque l'affiche (en euros, n{e2}gatif si le compte est {a2} d{e2}couvert) :"), titre, defMont)
    If Trim$(sMont) = "" Then Exit Sub
    If Not mod_DataStructure.TryLireMontant(sMont, nouveauM) Then
        MsgBox mod_Display.FR("Le montant saisi n'est pas valide.") & vbCrLf & vbCrLf & mod_DataStructure.MessageMontantAttendu(), vbExclamation, titre
        Exit Sub
    End If

    sDate = InputBox(mod_Display.FR("Date de ce solde (jj/mm/aaaa) : c'est la date {a2} laquelle la banque affiche ce solde, pas forc{e2}ment celle de l'import. Les op{e2}rations dat{e2}es de ce jour ou avant sont consid{e2}r{e2}es comme d{e2}j{a2} comprises dans le solde."), _
                     titre, Format(Date, "dd/mm/yyyy"))
    If Trim$(sDate) = "" Then Exit Sub
    If Not mod_DataStructure.TryLireDate(sDate, nouvelleD) Then
        MsgBox mod_Display.FR("La date saisie n'est pas valide.") & vbCrLf & vbCrLf & mod_DataStructure.MessageDateAttendue(), vbExclamation, titre
        Exit Sub
    End If
    nouvelleD = Int(nouvelleD)

    ' --- contrôle : possible seulement s'il existe déjà un solde de référence ---
    If Not IsEmpty(ancienM) And Not IsEmpty(ancienD) Then
        If IsNumeric(ancienM) And IsDate(ancienD) And Not loOps Is Nothing Then
            ancD = CDate(ancienD)
            If nouvelleD >= ancD Then
                calcule = CDbl(ancienM) + SommeOperations(loOps, ancD, nouvelleD)
            Else
                calcule = CDbl(ancienM) - SommeOperations(loOps, nouvelleD, ancD)
            End If
            ecart = Round(nouveauM - calcule, 2)
            controle = True
        End If
    End If

    If controle Then
        msg = mod_Display.FR("Solde calcul{e2} par le classeur au ") & Format(nouvelleD, "dd/mm/yyyy") & " : " & Format(calcule, "#,##0.00") & " EUR" & vbCrLf & _
              mod_Display.FR("Solde saisi : ") & Format(nouveauM, "#,##0.00") & " EUR" & vbCrLf & _
              mod_Display.FR("{E2}cart : ") & Format(ecart, "#,##0.00") & " EUR" & vbCrLf & vbCrLf
        If ecart = 0 Then
            msg = msg & mod_Display.FR("Aucun {e2}cart : les op{e2}rations enregistr{e2}es sont coh{e2}rentes avec votre banque.")
        Else
            msg = msg & mod_Display.FR("Un {e2}cart peut venir : d'une op{e2}ration non import{e2}e ou import{e2}e en double, d'une date de solde mal choisie, ") & _
                        mod_Display.FR("ou d'une op{e2}ration modifi{e2}e ou supprim{e2}e depuis le dernier solde.")
        End If
        msg = msg & vbCrLf & vbCrLf & mod_Display.FR("Enregistrer ce solde comme nouveau solde de r{e2}f{e2}rence ?")
        If MsgBox(msg, vbYesNo + IIf(ecart = 0, vbInformation, vbExclamation), titre) = vbNo Then Exit Sub
    End If

    ws.Range("soldeRefMontant").Value = nouveauM
    ws.Range("soldeRefDate").Value = nouvelleD
    On Error Resume Next
    If controle Then ws.Range("soldeEcart").Value = ecart Else ws.Range("soldeEcart").ClearContents
    On Error GoTo 0

    If Not controle Then
        MsgBox mod_Display.FR("Solde de r{e2}f{e2}rence enregistr{e2}. Aucun contr{o2}le n'a {e2}t{e2} possible (c'est le premier solde saisi)."), vbInformation, titre
    End If

End Sub

' Somme des montants des opérations dont la date comptable est > dDebut et <= dFin.
Private Function SommeOperations(ByVal lo As ListObject, ByVal dDebut As Date, ByVal dFin As Date) As Double

    Dim cMont() As Variant, cDt() As Variant
    Dim i As Long
    Dim d As Date
    Dim total As Double

    If lo.DataBodyRange Is Nothing Then Exit Function
    cMont = LireColonne(lo, "Montant")
    cDt = LireColonne(lo, "Date_Comptable")
    For i = 1 To UBound(cMont)
        d = VersDate(cDt(i))
        If d > dDebut And d <= dFin Then total = total + mod_DataStructure.ToDouble(cMont(i))
    Next i
    SommeOperations = total

End Function
