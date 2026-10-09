Attribute VB_Name = "mod_Echeancier"
Option Explicit

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
Public Const EF_ADR_PERIODE As String = "C11"
Public Const EF_ADR_DATE As String = "C12"
Public Const EF_ADR_DATEFIN As String = "C13"
Public Const EF_ADR_PRECISION As String = "C14"
Public Const EF_ADR_TOLERANCE As String = "C15"
Public Const EF_ADR_MESSAGE As String = "B17"      ' zone fusionnée B17:E17
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
    If TableVide(loR) Then Exit Sub

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

    For i = loR.ListRows.Count To 1 Step -1
        If termine(i) Then loR.ListRows(i).Delete
    Next i


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
    If nb > 0 Then Rapprocher listeID, nb
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

    Dim loOps As ListObject, loE As ListObject
    Dim dictNouveaux As Object
    Dim cID() As Variant, cDt() As Variant, cTiers() As Variant, cMont() As Variant, cCat() As Variant, cSous() As Variant
    Dim eV As Variant
    Dim opIdx() As Long
    Dim nOp As Long, nE As Long
    Dim i As Long, o As Long, k As Long, p As Long
    Dim dOp As Date, dE As Date
    Dim mOp As Double, mEch As Double, score As Double, meilleur As Double
    Dim tiersOp As String, catOp As String, sousOp As String
    Dim prec As String, periode As String
    Dim pO() As Long, pE() As Long, pS() As Double, nP As Long
    Dim usedO() As Boolean, usedE() As Boolean
    Dim accO() As Long, accE() As Long, nAcc As Long
    Dim ligne() As String
    Dim reponse As VbMsgBoxResult
    Dim msg As String
    Dim aSupprimer() As Boolean
    Dim iId As Long, iLib As Long, iTiers As Long, iCat As Long, iSous As Long, iMont As Long, iDate As Long, iPrec As Long, iPer As Long, iTol As Long

    If TableVide(TableEcheances()) Then Exit Sub
    Set loE = TableEcheances()

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

    ' --- Échéances ---
    eV = loE.DataBodyRange.Value2
    nE = UBound(eV, 1)
    iId = ColIdx(loE, "ID_Echeance"): iLib = ColIdx(loE, "Libelle"): iTiers = ColIdx(loE, "Tiers")
    iCat = ColIdx(loE, "Categorie"): iSous = ColIdx(loE, "SousCategorie"): iMont = ColIdx(loE, "Montant")
    iDate = ColIdx(loE, "DatePrevue"): iPrec = ColIdx(loE, "PrecisionDate"): iPer = ColIdx(loE, "Periodicite")
    iTol = ColIdx(loE, "TolerancePct")

    ' --- Toutes les paires compatibles ---
    ReDim pO(1 To 1): ReDim pE(1 To 1): ReDim pS(1 To 1)
    For o = 1 To nOp
        dOp = VersDate(cDt(opIdx(o)))
        mOp = mod_DataStructure.ToDouble(cMont(opIdx(o)))
        tiersOp = mod_DataStructure.CellText(cTiers(opIdx(o)))
        catOp = mod_DataStructure.CellText(cCat(opIdx(o)))
        sousOp = mod_DataStructure.CellText(cSous(opIdx(o)))
        For i = 1 To nE
            If mod_DataStructure.CellText(eV(i, iId)) <> "" Then
                dE = VersDate(eV(i, iDate))
                mEch = mod_DataStructure.ToDouble(eV(i, iMont))
                prec = mod_DataStructure.CellText(eV(i, iPrec))
                periode = mod_DataStructure.CellText(eV(i, iPer))
                If dE <> 0 And mEch <> 0 Then
                    If FenetreOk(dOp, dE, prec, periode) And _
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
    If nP = 0 Then Exit Sub

    ' --- Choix des paires : les plus proches d'abord, chaque élément une seule fois ---
    ReDim usedO(1 To nOp): ReDim usedE(1 To nE)
    ReDim accO(1 To nOp): ReDim accE(1 To nOp)
    Do
        p = 0
        meilleur = 1E+30
        For k = 1 To nP
            If Not usedO(pO(k)) And Not usedE(pE(k)) Then
                If pS(k) < meilleur Then
                    meilleur = pS(k)
                    p = k
                End If
            End If
        Next k
        If p = 0 Then Exit Do
        usedO(pO(p)) = True
        usedE(pE(p)) = True
        nAcc = nAcc + 1
        accO(nAcc) = pO(p)
        accE(nAcc) = pE(p)
    Loop
    If nAcc = 0 Then Exit Sub

    ' --- Texte de chaque proposition ---
    ReDim ligne(1 To nAcc)
    For k = 1 To nAcc
        ligne(k) = Format(VersDate(cDt(opIdx(accO(k)))), "dd/mm/yyyy") & "  " & _
                   Left$(mod_DataStructure.CellText(cTiers(opIdx(accO(k)))), 28) & "  " & _
                   Format(mod_DataStructure.ToDouble(cMont(opIdx(accO(k)))), "#,##0.00") & "  =  " & _
                   mod_DataStructure.CellText(eV(accE(k), iLib)) & " (" & mod_Display.FR("pr{e2}vue ") & _
                   IIf(mod_DataStructure.CellText(eV(accE(k), iPrec)) = PRECISION_MOIS, _
                       Format(VersDate(eV(accE(k), iDate)), "mm/yyyy"), _
                       Format(VersDate(eV(accE(k), iDate)), "dd/mm/yyyy")) & ")"
    Next k

    ' --- Validation par l'opérateur ---
    msg = nAcc & mod_Display.FR(" rapprochement(s) propos{e2}(s) entre les op{e2}rations import{e2}es et l'{e2}ch{e2}ancier :") & vbCrLf & vbCrLf
    For k = 1 To nAcc
        If k > 12 Then
            msg = msg & mod_Display.FR("... et ") & (nAcc - 12) & " autre(s)." & vbCrLf
            Exit For
        End If
        msg = msg & ligne(k) & vbCrLf
    Next k
    msg = msg & vbCrLf & mod_Display.FR("OUI : valider tout (les {e2}ch{e2}ances rapproch{e2}es disparaissent de la liste).") & vbCrLf & _
                mod_Display.FR("NON : examiner un par un.") & vbCrLf & _
                mod_Display.FR("ANNULER : ne rien rapprocher (les {e2}ch{e2}ances restent {a2} venir).")
    reponse = MsgBox(msg, vbYesNoCancel + vbQuestion, mod_Display.FR("Rapprochement avec l'{e2}ch{e2}ancier"))
    If reponse = vbCancel Then Exit Sub

    ReDim aSupprimer(1 To nE)
    For k = 1 To nAcc
        If reponse = vbYes Then
            aSupprimer(accE(k)) = True
        Else
            If MsgBox(ligne(k) & vbCrLf & vbCrLf & mod_Display.FR("Valider ce rapprochement ?"), vbYesNo + vbQuestion, _
                      mod_Display.FR("Rapprochement ") & k & "/" & nAcc) = vbYes Then
                aSupprimer(accE(k)) = True
            End If
        End If
    Next k

    ' --- Suppression des échéances rapprochées (de bas en haut : les numéros ne bougent pas) ---
    For i = nE To 1 Step -1
        If aSupprimer(i) Then loE.ListRows(i).Delete
    Next i

End Sub

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
    e.Periodicite = PERIODE_MENSUELLE
    e.PrecisionDate = PRECISION_JOUR
    e.TolerancePct = TOLERANCE_DEFAUT

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
    ThisWorkbook.Worksheets(g_EchNomFeuillePrec).Activate
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
    ws.Range(EF_ADR_PERIODE).Value = e.Periodicite
    If e.DatePrevue <> 0 Then ws.Range(EF_ADR_DATE).Value = e.DatePrevue Else ws.Range(EF_ADR_DATE).ClearContents
    If e.DateFin <> 0 Then ws.Range(EF_ADR_DATEFIN).Value = e.DateFin Else ws.Range(EF_ADR_DATEFIN).ClearContents
    ws.Range(EF_ADR_PRECISION).Value = e.PrecisionDate
    ws.Range(EF_ADR_TOLERANCE).Value = e.TolerancePct
    ws.Range(EF_ADR_MESSAGE).Value = ""

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
    If g_EchVerrouActif Then ws.Activate
End Sub

' --- Bouton "Valider" ---
Public Sub EchValider()

    Dim ws As Worksheet
    Dim lib As String, tiers As String, cat As String, sous As String, periode As String, prec As String
    Dim vMont As Variant, vDate As Variant, vFin As Variant, vTol As Variant
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
    periode = mod_DataStructure.CellText(ws.Range(EF_ADR_PERIODE).Value)
    vDate = ws.Range(EF_ADR_DATE).Value
    vFin = ws.Range(EF_ADR_DATEFIN).Value
    prec = mod_DataStructure.CellText(ws.Range(EF_ADR_PRECISION).Value)
    vTol = ws.Range(EF_ADR_TOLERANCE).Value

    If tiers = "" And cat = "" Then
        AfficherMessageEch ws, "Renseignez au moins le Tiers ou la Cat{e2}gorie : c'est ce qui permettra de reconna{i2}tre l'op{e2}ration quand elle sera import{e2}e."
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
    If Not IsNumeric(vMont) Then
        AfficherMessageEch ws, "Le montant doit {ea}tre un nombre."
        Exit Sub
    End If
    If CDbl(vMont) = 0 Then
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
    If prec <> PRECISION_JOUR And prec <> PRECISION_MOIS Then
        AfficherMessageEch ws, "Choisissez la pr{e2}cision de la date dans la liste (Jour ou Mois)."
        Exit Sub
    End If
    If IsEmpty(vDate) Or IsError(vDate) Then
        AfficherMessageEch ws, "La date est obligatoire."
        Exit Sub
    End If
    If Not IsDate(vDate) Then
        AfficherMessageEch ws, "La date n'est pas valide (exemple : 15/03/2027)."
        Exit Sub
    End If
    If Not IsEmpty(vFin) Then
        If periode = PERIODE_PONCTUELLE Then
            AfficherMessageEch ws, "Une date de fin n'a de sens que pour une op{e2}ration r{e2}currente : effacez-la, ou choisissez une p{e2}riodicit{e2}."
            Exit Sub
        End If
        If Not IsDate(vFin) Then
            AfficherMessageEch ws, "La date de fin n'est pas valide (exemple : 15/03/2028)."
            Exit Sub
        End If
        If CDate(vFin) < CDate(vDate) Then
            AfficherMessageEch ws, "La date de fin ne peut pas {ea}tre ant{e2}rieure {a2} la date pr{e2}vue."
            Exit Sub
        End If
    End If
    If Not IsEmpty(vTol) Then
        If Not IsNumeric(vTol) Then
            AfficherMessageEch ws, "La tol{e2}rance doit {ea}tre un nombre (pourcentage)."
            Exit Sub
        End If
        If CDbl(vTol) < 0 Or CDbl(vTol) > 100 Then
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
    e.Montant = CDbl(vMont)
    e.Periodicite = periode
    e.DatePrevue = CDate(vDate)
    If IsEmpty(vFin) Then e.DateFin = 0 Else e.DateFin = CDate(vFin)
    e.PrecisionDate = prec
    If IsEmpty(vTol) Then e.TolerancePct = TOLERANCE_DEFAUT Else e.TolerancePct = CDbl(vTol)

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
    ws.Visible = xlSheetVeryHidden
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
    If Not autre Is Nothing Then autre.Visible = xlSheetVeryHidden

End Sub

Public Sub SortirEcheancier()

    Dim ws As Worksheet

    Set ws = ActiveSheet
    If ws.Name <> mod_VarGlobales.NOM_FEUILLE_ECHEANCIER And ws.Name <> mod_VarGlobales.NOM_FEUILLE_RECURRENCES Then Exit Sub

    On Error Resume Next
    ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_SYNTHESE).Activate
    On Error GoTo 0
    ws.Visible = xlSheetVeryHidden

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

    For Each c In inter.Cells
        n = c.Row - lo.DataBodyRange.Row + 1
        If Not d.Exists(n) Then d.Add n, True
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
        For i = lo.ListRows.Count To 1 Step -1
            If sel.Exists(i) Then lo.ListRows(i).Delete
        Next i
        Exit Sub
    End If

    Select Case MsgBox(nbRegle & mod_Display.FR(" des {e2}ch{e2}ances s{e2}lectionn{e2}es viennent d'une r{e1}gle de r{e2}currence.") & vbCrLf & vbCrLf & _
                       mod_Display.FR("OUI : supprimer seulement ces {e2}ch{e2}ances. La r{e1}gle continue (si elle n'a plus aucune {e2}ch{e2}ance dans la liste, la suivante est recr{e2}{e2}e aussit{o2}t).") & vbCrLf & vbCrLf & _
                       mod_Display.FR("NON : arr{ea}ter la r{e2}currence {a2} partir de la premi{e1}re {e2}ch{e2}ance s{e2}lectionn{e2}e : ces {e2}ch{e2}ances disparaissent et la r{e1}gle ne produira plus rien {a2} partir de cette date.") & vbCrLf & vbCrLf & _
                       mod_Display.FR("ANNULER : ne rien faire."), _
                       vbYesNoCancel + vbQuestion, mod_Display.FR("{E2}ch{e2}ances issues d'une r{e1}gle"))
        Case vbYes
            For i = lo.ListRows.Count To 1 Step -1
                If sel.Exists(i) Then lo.ListRows(i).Delete
            Next i
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
    For i = lo.ListRows.Count To 1 Step -1
        If aSupprimer(i) Then lo.ListRows(i).Delete
    Next i

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
        For i = loE.ListRows.Count To 1 Step -1
            v = loE.ListRows(i).Range.Cells(1, cIdRec).Value2
            If mod_DataStructure.CellText(v) <> "" Then
                If IsNumeric(v) Then
                    If ids.Exists(CStr(CLng(v))) Then loE.ListRows(i).Delete
                End If
            End If
        Next i
    End If

    For i = loR.ListRows.Count To 1 Step -1
        If sel.Exists(i) Then loR.ListRows(i).Delete
    Next i

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
        For i = loE.ListRows.Count To 1 Step -1
            v = loE.ListRows(i).Range.Cells(1, cIdRec).Value2
            If mod_DataStructure.CellText(v) <> "" Then
                If IsNumeric(v) Then
                    If minDate.Exists(CStr(CLng(v))) Then loE.ListRows(i).Delete
                End If
            End If
        Next i

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
