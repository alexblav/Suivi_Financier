Attribute VB_Name = "mod_InstallEcheancier"
Option Explicit

' =====================================================================================
' MODULE : mod_InstallEcheancier
'
' ÉCHÉANCIER / PRÉVISIONNEL DE TRÉSORERIE - PARTIE 1/2 (construction)
'
' RÔLE (à lire en premier, même si vous débutez) :
'   Ce module CONSTRUIT tout ce que l'échéancier ajoute au classeur. La logique
'   (rapprochement, génération des occurrences, formulaire) est dans mod_Echeancier.
'   Une seule macro à lancer : InstallerEcheancier. Elle crée, dans l'ordre :
'     1. frm_Echeancier  : la liste des échéances à venir (tableau TblEcheances).
'     2. frm_Recurrences : la liste des règles de récurrence (tableau TblRecurrences).
'     3. frm_Echeance    : le formulaire de saisie d'une échéance.
'     4. le bloc "Prévisionnel de trésorerie" et ses 2 boutons, sur la feuille Synthese.
'     5. le bouton "Rendre récurrente" de l'écran de recherche (frm_RechercheOperations).
'   Les 3 feuilles frm_* sont masquées : on les ouvre avec les boutons.
'
' INSTALLATION (une seule fois) :
'   1. Alt+F11, Fichier > Importer : mod_VarGlobales.bas (mis à jour), mod_Echeancier.bas,
'      CE fichier, mod_ImportOFX.bas et mod_RechercheOperations.bas (mis à jour),
'      mod_InstallRechercheOperations.bas (mis à jour).
'   2. Ctrl+G, taper : InstallerEcheancier, puis Entrée.
'   3. Coller dans le module de la feuille frm_Echeance le code de
'      CodeBehind_frm_Echeance.txt (le message de fin d'installation donne son CodeName).
'
' RÉINSTALLATION : possible. Si les listes contiennent des données, une confirmation
' est demandée avant de les reconstruire (les données seraient PERDUES).
'
' SOLDE PRÉVISIONNEL : le bloc de Synthese est fait de FORMULES Excel :
'   Solde actuel = Solde de référence saisi + somme des opérations postérieures à la date du solde
'   Solde J+15   = Solde actuel + somme des échéances dont la date prévue est <= aujourd'hui + 15
'   Solde J+30   = idem à 30 jours
' Elles se recalculent donc toutes seules après chaque import ou saisie d'échéance.
' =====================================================================================

' --- Formulaire frm_Echeance : l'adresse des champs est dans mod_Echeancier (EF_ADR_...) ---
Private Const LIGNE_BOUTONS As Long = 2
Private Const LIGNE_INSTRUCTIONS As Long = 4
Private Const LIGNE_CORPS As Long = 6
' Les boutons + sont décalés vers la droite (en points) : la flèche de la liste déroulante
' du champ voisin s'affiche sur le bord droit de la cellule et ne doit pas être masquée.
Private Const DECALAGE_BOUTON_PLUS As Double = 24

' --- Listes : en-tête sur la ligne 6, données à partir de la ligne 7 ---
Private Const LIGNE_ENTETE_LISTE As Long = 6

' --- Bloc de la feuille Synthese ---
Private Const SYN_LIGNE_TITRE As Long = 14


' =====================================================================================
' MACRO PRINCIPALE
' =====================================================================================
Public Sub InstallerEcheancier()

    Dim wsPrecedente As Worksheet

    Set wsPrecedente = ActiveSheet

    If Not ConstruireListeEcheances(wsPrecedente) Then Exit Sub
    If Not ConstruireListeRecurrences(wsPrecedente) Then Exit Sub
    If Not ConstruireFormulaire(wsPrecedente) Then Exit Sub
    ConstruireBlocSynthese
    AjouterBoutonRecurrenteRO

    MsgBox mod_Display.FR("L'{e2}ch{e2}ancier est install{e2}.") & vbCrLf & vbCrLf & _
           mod_Display.FR("PROCHAINE {E2}TAPE (indispensable) : coller le code des {e2}v{e2}nements dans la feuille du formulaire.") & vbCrLf & _
           "Nom interne (CodeName) de la feuille " & mod_VarGlobales.NOM_FEUILLE_ECHEANCE & " : " & _
           mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_ECHEANCE).CodeName & vbCrLf & vbCrLf & _
           mod_Display.FR("Voir le fichier CodeBehind_frm_Echeance.txt."), _
           vbInformation, mod_Display.FR("Installation termin{e2}e")

End Sub


' =====================================================================================
' LISTES (échéances à venir, règles de récurrence)
' =====================================================================================
Private Function ConstruireListeEcheances(ByVal wsPrecedente As Worksheet) As Boolean

    Dim entetes As Variant, largeurs As Variant, formats As Variant, boutons As Variant
    Dim t As String

    entetes = Array("ID_Echeance", "ID_Recurrence", "Libelle", "Tiers", "Categorie", "SousCategorie", "Montant", _
                    "NumCheque", "DatePrevue", "PrecisionDate", "Periodicite", "TolerancePct")
    largeurs = Array(12, 13, 28, 24, 20, 24, 11, 13, 12, 13, 14, 12)
    formats = Array("0", "0", "@", "@", "@", "@", "#,##0.00", "@", "dd/mm/yyyy", "@", "@", "0")

    boutons = Array( _
        mod_Display.FR("Ajouter une {e2}ch{e2}ance"), "SaisirEcheance", "btnAjouterEcheance", 118, _
        mod_Display.FR("Supprimer la s{e2}lection"), "SupprimerEcheanceSelection", "btnSupprimerEcheance", 118, _
        mod_Display.FR("Voir les r{e2}gles"), "AfficherRecurrences", "btnVoirRecurrences", 100, _
        mod_InstallCommun.CapSortir(), "SortirEcheancier", "btnSortirEcheancier", mod_InstallCommun.FRM_BTN_L)

    t = "Cette liste contient les {e2}ch{e2}ances {a2} venir : op{e2}rations pr{e2}vues qui ne sont pas encore pass{e2}es sur le compte. "
    t = t & "Elles servent au calcul du solde pr{e2}visionnel de [f:Synthese]. Les {e2}ch{e2}ances d'une r{e1}gle y sont ajout{e2}es "
    t = t & "automatiquement (voir [f:frm_Recurrences])."
    t = t & Chr(10) & "Apr{e1}s chaque import, les op{e2}rations import{e2}es sont compar{e2}es aux {e2}ch{e2}ances et l'import propose les rapprochements. "
    t = t & "Une fois un rapprochement valid{e2}, l'{e2}ch{e2}ance dispara{i2}t de cette liste : seule l'op{e2}ration est conserv{e2}e. "
    t = t & "Une {e2}ch{e2}ance non rapproch{e2}e reste ici, reste compt{e2}e dans le solde pr{e2}visionnel et, si sa date est pass{e2}e, "
    t = t & "est compt{e2}e dans '{E2}ch. en retard' de [f:Synthese] (pour une pr{e2}cision Mois, une fois le mois {e2}coul{e2}) : c'est {a2} vous de la supprimer."
    t = t & Chr(10) & "Vous pouvez corriger directement les cellules (montant, date, tiers...). [c:PrecisionDate] vaut Jour ou Mois (Mois : l'op{e2}ration "
    t = t & "peut arriver n'importe quel jour du mois de [c:DatePrevue], mais l'{e2}ch{e2}ance est compt{e2}e dans le solde pr{e2}visionnel d{e1}s le 1er du mois). [c:TolerancePct] : {e2}cart de montant accept{e2} pour le rapprochement. "
    t = t & "[c:NumCheque] : num{e2}ro du ch{e1}que attendu (vide si aucun) ; son contr{o2}le est automatique et silencieux (aucun message) : la cellule devient rouge si le num{e2}ro est invalide ou d{e2}j{a2} pass{e2} en banque, orange s'il figure deux fois dans cette liste. "
    t = t & "[c:ID_Recurrence] indique la r{e1}gle d'origine (vide pour une op{e2}ration ponctuelle) ; modifier une {e2}ch{e2}ance ne modifie pas sa r{e1}gle."
    t = t & Chr(10) & "Boutons :"
    t = t & Chr(10) & "- [b:Ajouter une {e2}ch{e2}ance] : ouvre le formulaire de saisie."
    t = t & Chr(10) & "- [b:Supprimer la s{e2}lection] : supprime les lignes s{e2}lectionn{e2}es. Pour une {e2}ch{e2}ance qui vient d'une r{e1}gle, "
    t = t & "deux choix vous sont propos{e2}s : Oui = supprimer seulement ces {e2}ch{e2}ances (la r{e1}gle continue ; si elle n'a plus aucune {e2}ch{e2}ance dans la liste, "
    t = t & "la suivante est recr{e2}{e2}e aussit{o2}t) ; Non = arr{ea}ter la r{e2}currence {a2} partir de la premi{e1}re {e2}ch{e2}ance s{e2}lectionn{e2}e "
    t = t & "(ces {e2}ch{e2}ances disparaissent et la r{e1}gle ne produira plus rien {a2} partir de cette date, puis dispara{i2}t de [f:frm_Recurrences] quand il ne lui reste plus d'{e2}ch{e2}ance : cas d'un remboursement anticip{e2}, par exemple)."
    t = t & Chr(10) & "- [b:Voir les r{e1}gles] : affiche la liste des r{e1}gles de r{e2}currence."
    t = t & Chr(10) & "- [b:Sortir] : revient sur [f:Synthese]."

    ConstruireListeEcheances = ConstruireListe(mod_VarGlobales.NOM_FEUILLE_ECHEANCIER, mod_VarGlobales.NOM_TABLE_ECHEANCES, _
                                               entetes, largeurs, formats, t, boutons, wsPrecedente)

End Function

Private Function ConstruireListeRecurrences(ByVal wsPrecedente As Worksheet) As Boolean

    Dim entetes As Variant, largeurs As Variant, formats As Variant, boutons As Variant
    Dim t As String

    entetes = Array("ID_Recurrence", "Libelle", "Tiers", "Categorie", "SousCategorie", "Montant", "Periodicite", _
                    "DateProchaine", "DateFin", "PrecisionDate", "JourAncre", "TolerancePct", "Actif")
    largeurs = Array(13, 28, 24, 20, 24, 11, 14, 14, 12, 13, 10, 12, 8)
    formats = Array("0", "@", "@", "@", "@", "#,##0.00", "@", "dd/mm/yyyy", "dd/mm/yyyy", "@", "0", "0", "@")

    boutons = Array( _
        mod_Display.FR("Ajouter une {e2}ch{e2}ance"), "SaisirEcheance", "btnAjouterEcheanceR", 118, _
        mod_Display.FR("Actualiser les occurrences"), "ActualiserRecurrences", "btnActualiserRecurrences", 135, _
        mod_Display.FR("Supprimer la r{e2}gle"), "SupprimerRecurrenceSelection", "btnSupprimerRecurrence", 118, _
        mod_Display.FR("Voir les {e2}ch{e2}ances"), "AfficherEcheancier", "btnVoirEcheancier", 108, _
        mod_InstallCommun.CapSortir(), "SortirEcheancier", "btnSortirRecurrences", mod_InstallCommun.FRM_BTN_L)

    t = "Cette liste contient les r{e1}gles de r{e2}currence. Une r{e1}gle produit seule ses {e2}ch{e2}ances dans [f:frm_Echeancier] : "
    t = t & "celles des " & mod_Echeancier.ECH_HORIZON_JOURS & " prochains jours, et toujours au moins une. La liste est compl{e2}t{e2}e "
    t = t & "{a2} l'ouverture du classeur, {a2} chaque import et {a2} l'ouverture de ces listes. Une r{e1}gle sans [c:DateFin] ne s'arr{ea}te jamais : "
    t = t & "supprimez-la ou mettez-la en pause pour l'arr{ea}ter."
    t = t & Chr(10) & "[c:DateProchaine] : prochaine date {a2} cr{e2}er (n'y touchez que pour d{e2}caler la r{e1}gle). [c:DateFin] : facultative ; "
    t = t & "apr{e1}s cette date la r{e1}gle ne cr{e2}e plus rien (elle est retir{e2}e de cette liste d{e1}s qu'il ne lui reste plus aucune {e2}ch{e2}ance : derni{e1}re {e2}ch{e2}ance rapproch{e2}e ou supprim{e2}e). "
    t = t & "[c:Actif] : Oui ou Non (Non = r{e1}gle en pause, aucune {e2}ch{e2}ance cr{e2}{e2}e). [c:JourAncre] : jour du mois conserv{e2} "
    t = t & "d'une occurrence {a2} l'autre."
    t = t & Chr(10) & "Vous pouvez modifier une r{e1}gle directement dans le tableau (montant, tiers, p{e2}riodicit{e2}, date de fin...). "
    t = t & "La modification ne s'applique qu'aux {e2}ch{e2}ances cr{e2}{e2}es ensuite ; pour l'appliquer aussi {a2} celles d{e2}j{a2} dans la liste, "
    t = t & "cliquez sur [b:Actualiser les occurrences]."
    t = t & Chr(10) & "Boutons :"
    t = t & Chr(10) & "- [b:Ajouter une {e2}ch{e2}ance] : ouvre le formulaire de saisie (une p{e2}riodicit{e2} cr{e2}e une r{e1}gle ; sans p{e2}riodicit{e2}, l'op{e2}ration est unique)."
    t = t & Chr(10) & "- [b:Actualiser les occurrences] : pour chaque r{e1}gle, supprime ses {e2}ch{e2}ances actuellement dans la liste (y compris "
    t = t & "celles non rapproch{e2}es) puis les recr{e2}e avec les valeurs de la r{e1}gle, en repartant de la plus ancienne. Une r{e1}gle en pause "
    t = t & "perd ainsi ses {e2}ch{e2}ances."
    t = t & Chr(10) & "- [b:Supprimer la r{e1}gle] : supprime les r{e1}gles s{e2}lectionn{e2}es ET leurs {e2}ch{e2}ances {a2} venir (pour arr{ea}ter une r{e2}currence "
    t = t & "pour de bon). Pour l'arr{ea}ter {a2} une date pr{e2}cise, renseignez plut{o2}t [c:DateFin] puis cliquez sur [b:Actualiser les occurrences]."
    t = t & Chr(10) & "- [b:Voir les {e2}ch{e2}ances] : affiche la liste des {e2}ch{e2}ances."
    t = t & Chr(10) & "- [b:Sortir] : revient sur [f:Synthese]."

    ConstruireListeRecurrences = ConstruireListe(mod_VarGlobales.NOM_FEUILLE_RECURRENCES, mod_VarGlobales.NOM_TABLE_RECURRENCES, _
                                                 entetes, largeurs, formats, t, boutons, wsPrecedente)

End Function

' Construction commune aux deux listes.
'   boutons : suite de quadruplets (légende, macro, nom du bouton, largeur).
Private Function ConstruireListe(ByVal nomFeuille As String, ByVal nomTable As String, ByVal entetes As Variant, _
                                 ByVal largeurs As Variant, ByVal formats As Variant, ByVal texteInstr As String, _
                                 ByVal boutons As Variant, ByVal wsPrecedente As Worksheet) As Boolean

    Dim ws As Worksheet, wsExistante As Worksheet
    Dim lo As ListObject
    Dim w() As Variant
    Dim nb As Long, i As Long
    Dim gauche As Double
    Dim donneesPresentes As Boolean
    Dim sauve As Variant, enteteSauve As Variant
    Dim nbSauve As Long, k As Long, r As Long, jEnt As Long
    Dim colonne() As Variant

    ' Garde-fou : les données déjà saisies sont CONSERVÉES. Elles sont mises de côté, la feuille est
    ' reconstruite, puis chaque colonne est remise par son NOM (une colonne nouvelle reste vide).
    Set wsExistante = mod_InstallCommun.TrouverFeuille(nomFeuille)
    If Not wsExistante Is Nothing Then
        On Error Resume Next
        Set lo = wsExistante.ListObjects(nomTable)
        On Error GoTo 0
        If Not lo Is Nothing Then
            If Not lo.DataBodyRange Is Nothing Then
                If Application.WorksheetFunction.CountA(lo.ListColumns(1).DataBodyRange) > 0 Then donneesPresentes = True
            End If
        End If
        If donneesPresentes Then
            sauve = lo.DataBodyRange.Value2
            enteteSauve = lo.HeaderRowRange.Value2
            nbSauve = UBound(sauve, 1)
        End If
    End If

    Set ws = mod_InstallCommun.PreparerFeuille(nomFeuille, , True)
    If ws Is Nothing Then Exit Function

    Application.ScreenUpdating = False

    nb = UBound(entetes) - LBound(entetes) + 1
    ReDim w(0 To nb)
    w(0) = 2
    For i = 0 To nb - 1
        w(i + 1) = largeurs(LBound(largeurs) + i)
    Next i
    mod_InstallCommun.MettreEnForme ws, w

    ' En-tête commun : 1 ligne de boutons, pas de compteur.
    mod_InstallCommun.PoserHauteursEntete ws, 1, False
    mod_InstallCommun.EcrireInstructions ws, LIGNE_INSTRUCTIONS, 2, 2, 3, nb + 1, texteInstr

    ' Tableau (une ligne vide, indispensable pour créer un ListObject)
    For i = 0 To nb - 1
        ws.Cells(LIGNE_ENTETE_LISTE, i + 2).Value = entetes(LBound(entetes) + i)
    Next i
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(LIGNE_ENTETE_LISTE, 2), ws.Cells(LIGNE_ENTETE_LISTE + 1, nb + 1)), , xlYes)
    lo.Name = nomTable
    lo.TableStyle = ""
    lo.ShowAutoFilter = False
    If nbSauve > 1 Then lo.Resize ws.Range(ws.Cells(LIGNE_ENTETE_LISTE, 2), ws.Cells(LIGNE_ENTETE_LISTE + nbSauve, nb + 1))
    mod_InstallCommun.PoserEnteteTableau lo.HeaderRowRange
    For i = 1 To nb
        lo.ListColumns(i).DataBodyRange.NumberFormat = formats(LBound(formats) + i - 1)
    Next i
    With lo.DataBodyRange
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = mod_InstallCommun.CoulBordureSaisie()
        .Interior.Color = mod_InstallCommun.CoulFondSaisie()
    End With

    ' Remise en place des données conservées, colonne par colonne (par le nom de l'en-tête).
    If nbSauve > 0 Then
        For i = 1 To nb
            jEnt = 0
            For k = 1 To UBound(enteteSauve, 2)
                If CStr(enteteSauve(1, k)) = CStr(entetes(LBound(entetes) + i - 1)) Then
                    jEnt = k
                    Exit For
                End If
            Next k
            If jEnt > 0 Then
                ReDim colonne(1 To nbSauve, 1 To 1)
                For r = 1 To nbSauve
                    colonne(r, 1) = sauve(r, jEnt)
                Next r
                lo.ListColumns(i).DataBodyRange.Value2 = colonne
            End If
        Next i
    End If

    ' Contrôle silencieux des numéros de chèque (colonne NumCheque de la liste des échéances).
    If nomTable = mod_VarGlobales.NOM_TABLE_ECHEANCES Then PoserControleCheques ws, lo

    ' Boutons (en dernier : ils se placent d'après les hauteurs de lignes)
    gauche = ws.Cells(1, 2).Left
    For i = LBound(boutons) To UBound(boutons) Step 4
        mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, CStr(boutons(i)), CStr(boutons(i + 1)), _
                                              CStr(boutons(i + 2)), CDbl(boutons(i + 3))
    Next i

    mod_InstallCommun.TerminerFeuille ws, wsPrecedente
    Application.ScreenUpdating = True
    ConstruireListe = True

End Function


' =====================================================================================
' FORMULAIRE frm_Echeance
' =====================================================================================
Private Function ConstruireFormulaire(ByVal wsPrecedente As Worksheet) As Boolean

    Dim ws As Worksheet
    Dim gauche As Double
    Dim i As Long

    Set ws = mod_InstallCommun.PreparerFeuille(mod_VarGlobales.NOM_FEUILLE_ECHEANCE)
    If ws Is Nothing Then Exit Function

    Application.ScreenUpdating = False

    mod_InstallCommun.MettreEnForme ws, Array(2, 30, 44, 22, 22, 2)

    If mod_InstallCommun.PremiereLigneCorps(1, False) <> LIGNE_CORPS Or _
       mod_InstallCommun.LigneInstructions(1, False) <> LIGNE_INSTRUCTIONS Then
        MsgBox "mod_InstallEcheancier : constantes de lignes incoherentes avec mod_InstallCommun.", vbCritical
    End If
    mod_InstallCommun.PoserHauteursEntete ws, 1, False
    mod_InstallCommun.EcrireInstructions ws, LIGNE_INSTRUCTIONS, 2, 2, 3, 5, TexteInstructionsFormulaire()

    ' --- Champs ---
    mod_InstallCommun.PoserEtiquette ws.Range("B6"), mod_Display.FR("Libell{e2} (facultatif)")
    mod_InstallCommun.PoserChampSaisie ws.Range(mod_Echeancier.EF_ADR_LIBELLE), "@"

    mod_InstallCommun.PoserEtiquette ws.Range("B7"), mod_Display.FR("Tiers (contient)")
    mod_InstallCommun.PoserChampSaisie ws.Range(mod_Echeancier.EF_ADR_TIERS), "@"

    mod_InstallCommun.PoserEtiquette ws.Range("B8"), mod_Display.FR("Cat{e2}gorie")
    mod_InstallCommun.PoserChampSaisie ws.Range(mod_Echeancier.EF_ADR_CAT), "@"

    mod_InstallCommun.PoserEtiquette ws.Range("B9"), mod_Display.FR("Sous-cat{e2}gorie")
    mod_InstallCommun.PoserChampSaisie ws.Range(mod_Echeancier.EF_ADR_SOUS), "@"

    mod_InstallCommun.PoserEtiquette ws.Range("B10"), mod_Display.FR("Montant (- = d{e2}pense)")
    mod_InstallCommun.PoserChampSaisie ws.Range(mod_Echeancier.EF_ADR_MONTANT), "#,##0.00"

    mod_InstallCommun.PoserEtiquette ws.Range("B11"), mod_Display.FR("Num{e2}ro de ch{e1}que (facultatif)")
    mod_InstallCommun.PoserChampSaisie ws.Range(mod_Echeancier.EF_ADR_CHEQUE), "@"

    mod_InstallCommun.PoserEtiquette ws.Range("B12"), mod_Display.FR("P{e2}riodicit{e2}")
    mod_InstallCommun.PoserChampSaisie ws.Range(mod_Echeancier.EF_ADR_PERIODE), "@"

    mod_InstallCommun.PoserEtiquette ws.Range("B13"), mod_Display.FR("Date pr{e2}vue (ou 1{e1}re date)")
    mod_InstallCommun.PoserChampSaisie ws.Range(mod_Echeancier.EF_ADR_DATE), "dd/mm/yyyy"

    mod_InstallCommun.PoserEtiquette ws.Range("B14"), mod_Display.FR("Date de fin (facultative)")
    mod_InstallCommun.PoserChampSaisie ws.Range(mod_Echeancier.EF_ADR_DATEFIN), "dd/mm/yyyy"

    mod_InstallCommun.PoserEtiquette ws.Range("B15"), mod_Display.FR("Pr{e2}cision de la date")
    mod_InstallCommun.PoserChampSaisie ws.Range(mod_Echeancier.EF_ADR_PRECISION), "@"

    mod_InstallCommun.PoserEtiquette ws.Range("B16"), mod_Display.FR("Tol{e2}rance sur le montant (%)")
    mod_InstallCommun.PoserChampSaisie ws.Range(mod_Echeancier.EF_ADR_TOLERANCE), "0"

    ws.Rows(17).RowHeight = mod_InstallCommun.FRM_H_SEP

    mod_InstallCommun.PoserMessage ws.Range("B18:E18")
    ws.Range(mod_Echeancier.EF_ADR_MESSAGE).Font.Color = RGB(192, 80, 0)

    ' --- Listes techniques (périodicités, précisions) : colonnes masquées ---
    ws.Cells(1, mod_Echeancier.EF_COL_AIDE_PERIODE).Value = "Periodicites"
    ws.Cells(2, mod_Echeancier.EF_COL_AIDE_PERIODE).Value = mod_Echeancier.PERIODE_PONCTUELLE
    ws.Cells(3, mod_Echeancier.EF_COL_AIDE_PERIODE).Value = mod_Echeancier.PERIODE_HEBDO
    ws.Cells(4, mod_Echeancier.EF_COL_AIDE_PERIODE).Value = mod_Echeancier.PERIODE_MENSUELLE
    ws.Cells(5, mod_Echeancier.EF_COL_AIDE_PERIODE).Value = mod_Echeancier.PERIODE_TRIMESTRIELLE
    ws.Cells(6, mod_Echeancier.EF_COL_AIDE_PERIODE).Value = mod_Echeancier.PERIODE_ANNUELLE
    ws.Cells(1, mod_Echeancier.EF_COL_AIDE_PRECISION).Value = "Precisions"
    ws.Cells(2, mod_Echeancier.EF_COL_AIDE_PRECISION).Value = mod_Echeancier.PRECISION_JOUR
    ws.Cells(3, mod_Echeancier.EF_COL_AIDE_PRECISION).Value = mod_Echeancier.PRECISION_MOIS
    PoserNomListe "ListePeriodicites", ws, mod_Echeancier.EF_COL_AIDE_PERIODE, 2, 6
    PoserNomListe "ListePrecisionsDate", ws, mod_Echeancier.EF_COL_AIDE_PRECISION, 2, 3
    For i = mod_Echeancier.EF_COL_AIDE_SOUS To mod_Echeancier.EF_COL_AIDE_TIERS
        ws.Columns(i).Hidden = True
    Next i

    ' --- Boutons "+" à côté de Catégorie et de Sous-catégorie (ouvrent frm_NouvelleCategorie) ---
    mod_InstallCommun.AjouterBoutonCellule ws, ws.Range("D8"), mod_InstallCommun.CapAjouter(), _
        "EchNouvelleCategorie", "btnEchNouvelleCategorie", mod_InstallCommun.FRM_BTN_PLUS, DECALAGE_BOUTON_PLUS
    mod_InstallCommun.AjouterBoutonCellule ws, ws.Range("D9"), mod_InstallCommun.CapAjouter(), _
        "EchNouvelleCategorie", "btnEchNouvelleSousCategorie", mod_InstallCommun.FRM_BTN_PLUS, DECALAGE_BOUTON_PLUS

    ' --- Boutons de la ligne d'en-tête ---
    gauche = ws.Cells(1, 2).Left
    mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, mod_InstallCommun.CapValider(), _
        "EchValider", "btnEchValider", mod_InstallCommun.FRM_BTN_L
    mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, mod_InstallCommun.CapAnnuler(), _
        "EchAnnuler", "btnEchAnnuler", mod_InstallCommun.FRM_BTN_L

    mod_InstallCommun.TerminerFeuille ws, wsPrecedente
    Application.ScreenUpdating = True
    ConstruireFormulaire = True

End Function

Private Sub PoserNomListe(ByVal nom As String, ByVal ws As Worksheet, ByVal col As Long, ByVal ligneDeb As Long, ByVal ligneFin As Long)
    ThisWorkbook.Names.Add Name:=nom, RefersTo:="=" & ws.Name & "!" & _
        ws.Range(ws.Cells(ligneDeb, col), ws.Cells(ligneFin, col)).Address(True, True)
End Sub

Private Function TexteInstructionsFormulaire() As String

    Dim t As String

    t = "Ce formulaire d{e2}clare une op{e2}ration {a2} venir. Quand une op{e2}ration import{e2}e lui correspond, l'import propose de les rapprocher."
    t = t & Chr(10) & "Reconnaissance : renseignez au moins le [c:Tiers], la [c:Cat{e2}gorie] ou le [c:Num{e2}ro de ch{e1}que] (plusieurs si vous voulez). Le [c:Tiers] se choisit dans la liste "
    t = t & "des Tiers d{e2}j{a2} connus dans les op{e2}rations, ou se tape librement : il est cherch{e2} dans le Tiers de l'op{e2}ration import{e2}e, "
    t = t & "sans tenir compte des majuscules, donc un mot-cl{e2} suffit (ex. EDF). La [c:Cat{e2}gorie] et la [c:Sous-cat{e2}gorie] se choisissent "
    t = t & "dans les listes (la liste des sous-cat{e2}gories d{e2}pend de la cat{e2}gorie choisie ; sans cat{e2}gorie, elles sont toutes propos{e2}es, "
    t = t & "et choisir une sous-cat{e2}gorie qui n'appartient qu'{a2} une cat{e2}gorie renseigne celle-ci). "
    t = t & "Si vous renseignez plusieurs crit{e1}res (Tiers, Cat{e2}gorie, Sous-cat{e2}gorie), l'op{e2}ration doit tous les respecter."
    t = t & Chr(10) & "[c:Montant] : n{e2}gatif pour une d{e2}pense, positif pour une entr{e2}e d'argent ; une op{e2}ration de signe contraire n'est jamais rapproch{e2}e."
    t = t & Chr(10) & "[c:Num{e2}ro de ch{e1}que] : facultatif, pour une op{e2}ration unique seulement. Chiffres uniquement (10 au maximum). "
    t = t & "Le formulaire refuse un num{e2}ro d{e2}j{a2} pass{e2} en banque ou d{e2}j{a2} saisi dans l'{e2}ch{e2}ancier. "
    t = t & "Avec un num{e2}ro, l'op{e2}ration est reconnue {a2} l'import par ce seul num{e2}ro (le signe du montant doit concorder) : "
    t = t & "le Tiers, la Cat{e2}gorie, la date et la tol{e2}rance ne sont plus examin{e2}s."
    t = t & Chr(10) & "[c:P{e2}riodicit{e2}] : laiss{e2}e vide (ou Ponctuelle), l'op{e2}ration est unique. Hebdomadaire, Mensuelle, Trimestrielle ou Annuelle "
    t = t & "cr{e2}e une r{e1}gle qui produit seule les {e2}ch{e2}ances suivantes, sans limite, jusqu'{a2} ce que vous la supprimiez, "
    t = t & "la mettiez en pause ou qu'elle atteigne sa date de fin."
    t = t & Chr(10) & "[c:Date pr{e2}vue] : date de l'op{e2}ration, ou de la premi{e1}re pour une r{e1}gle (le jour est conserv{e2} chaque mois ; "
    t = t & "un 31 devient le dernier jour des mois plus courts). [c:Date de fin] : facultative, r{e2}serv{e2}e aux op{e2}rations r{e2}currentes ; "
    t = t & "la derni{e1}re occurrence cr{e2}{e2}e est la derni{e1}re dont la date ne d{e2}passe pas la date de fin. Laiss{e2}e vide, la r{e1}gle ne s'arr{ea}te jamais."
    t = t & Chr(10) & "[c:Pr{e2}cision de la date] : Jour = l'op{e2}ration est reconnue de " & mod_Echeancier.FEN_AVANT & " jours avant "
    t = t & "{a2} " & mod_Echeancier.FEN_APRES & " jours apr{e1}s la date pr{e2}vue (" & mod_Echeancier.FEN_HEBDO & " jours de chaque c{o2}t{e2} pour une r{e1}gle hebdomadaire). "
    t = t & "Mois = elle est reconnue n'importe quel jour du mois de la date pr{e2}vue : si vous ne connaissez pas le jour, mettez une date quelconque de ce mois."
    t = t & Chr(10) & "[c:Tol{e2}rance] : {e2}cart de montant accept{e2}, en pourcentage du montant pr{e2}vu (" & mod_Echeancier.TOLERANCE_DEFAUT & " si vide ; 0 = montant exact)."
    t = t & Chr(10) & "Si un champ est incorrect, un message appara{i2}t sous le formulaire et rien n'est enregistr{e2}."
    t = t & Chr(10) & "Boutons :"
    t = t & Chr(10) & "- [b:+] (en face de Cat{e2}gorie et de Sous-cat{e2}gorie) : ouvre [f:frm_NouvelleCategorie] pour cr{e2}er une cat{e2}gorie ou une "
    t = t & "sous-cat{e2}gorie absente des listes ; une fois valid{e2}e, elle est reprise dans ce formulaire."
    t = t & Chr(10) & "- [b:Valider] : enregistre l'{e2}ch{e2}ance (ou la r{e1}gle) puis referme le formulaire."
    t = t & Chr(10) & "- [b:Annuler] : referme le formulaire sans rien enregistrer."

    TexteInstructionsFormulaire = t

End Function


' =====================================================================================
' BLOC "Prévisionnel de trésorerie" sur la feuille Synthese
' =====================================================================================
' Placé sous le bouton Importer (lignes 14 à 20, colonnes A et B) ; 2 boutons à droite.
' Les 2 cellules de saisie portent un nom (soldeRefMontant, soldeRefDate) utilisé par les formules.
Private Sub ConstruireBlocSynthese()

    Dim ws As Worksheet
    Dim l As Long, k As Long
    Dim libelles(0 To 7) As String
    Dim neuf(0 To 7) As Boolean
    Dim coulBloc As Long
    Dim f As String, c3 As String

    Set ws = mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_SYNTHESE)
    If ws Is Nothing Then
        MsgBox mod_Display.FR("La feuille Synthese est introuvable : le bloc pr{e2}visionnel n'a pas {e2}t{e2} cr{e2}{e2}."), vbExclamation
        Exit Sub
    End If

    ' NON DESTRUCTIF : le bloc a pu être déplacé ou retouché à la main. Sa position est retrouvée
    ' par le nom soldeRefMontant ; ce qui existe déjà (libellés, valeurs saisies, formules, formats,
    ' noms, boutons) n'est JAMAIS réécrit, seul ce qui manque est ajouté.
    l = LigneTitreBloc(ws)
    coulBloc = RGB(141, 180, 226)   ' même bleu que les paramètres de la feuille

    libelles(0) = mod_Display.FR("Pr{e2}visionnel de tr{e2}sorerie")
    libelles(1) = mod_Display.FR("Solde de r{e2}f.")
    libelles(2) = "Date du solde"
    libelles(3) = "Solde actuel"
    libelles(4) = mod_Display.FR("Solde {a2} J+15")
    libelles(5) = mod_Display.FR("Solde {a2} J+30")
    libelles(6) = mod_Display.FR("{E2}ch. en retard")
    libelles(7) = mod_Display.FR("{E2}cart du solde")

    For k = 0 To 7
        If IsEmpty(ws.Cells(l + k, 1).Value) Then
            neuf(k) = True
            ws.Range(ws.Cells(l + k, 1), ws.Cells(l + k, 2)).Interior.Color = coulBloc
            ws.Cells(l + k, 1).Value = libelles(k)
        End If
    Next k
    If neuf(0) Then ws.Cells(l, 1).Font.Bold = True

    ' Cellules de saisie (mise en forme seulement si la ligne vient d'être créée)
    If neuf(1) Then
        With ws.Cells(l + 1, 2)
            .NumberFormat = "#,##0.00"
            .Interior.Color = mod_InstallCommun.CoulFondSaisie()
            .Locked = False
        End With
    End If
    If neuf(2) Then
        With ws.Cells(l + 2, 2)
            .NumberFormat = "dd/mm/yyyy"
            .Interior.Color = mod_InstallCommun.CoulFondSaisie()
            .Locked = False
        End With
    End If

    AjouterNomSiAbsent "soldeRefMontant", ws.Cells(l + 1, 2)
    AjouterNomSiAbsent "soldeRefDate", ws.Cells(l + 2, 2)
    ' Écart constaté lors de la dernière saisie du solde (voir mod_Echeancier.SaisirSoldeReference)
    AjouterNomSiAbsent "soldeEcart", ws.Cells(l + 7, 2)
    If neuf(7) Then
        With ws.Cells(l + 7, 2)
            .NumberFormat = "#,##0.00;[Red]-#,##0.00"
            .Interior.Color = mod_InstallCommun.CoulFondLecture()
            .HorizontalAlignment = xlRight
        End With
    End If

    ' Formules (syntaxe anglaise pour .Formula, quelle que soit la langue d'Excel)
    c3 = ws.Cells(l + 3, 2).Address(False, False)
    If FormuleAReparer(ws.Cells(l + 3, 2)) Then
        f = "=IF(OR(soldeRefMontant="""",soldeRefDate=""""),""" & mod_Display.FR("{a2} renseigner") & """," & _
            "soldeRefMontant+SUMIFS(TblOperations[Montant],TblOperations[Date_Comptable],"">""&soldeRefDate))"
        ws.Cells(l + 3, 2).Formula = f
    End If
    If FormuleAReparer(ws.Cells(l + 4, 2)) Then
        ws.Cells(l + 4, 2).Formula = "=IF(ISNUMBER(" & c3 & ")," & c3 & "+SUMIFS(" & mod_VarGlobales.NOM_TABLE_ECHEANCES & _
            "[Montant]," & mod_VarGlobales.NOM_TABLE_ECHEANCES & "[DatePrevue],""<=""&(TODAY()+15)),"""")"
    End If
    If FormuleAReparer(ws.Cells(l + 5, 2)) Then
        ws.Cells(l + 5, 2).Formula = "=IF(ISNUMBER(" & c3 & ")," & c3 & "+SUMIFS(" & mod_VarGlobales.NOM_TABLE_ECHEANCES & _
            "[Montant]," & mod_VarGlobales.NOM_TABLE_ECHEANCES & "[DatePrevue],""<=""&(TODAY()+30)),"""")"
    End If
    ' En retard = date prévue passée ; pour une précision "Mois", seulement une fois le mois écoulé.
    If FormuleAReparer(ws.Cells(l + 6, 2)) Then
        ws.Cells(l + 6, 2).Formula = "=SUMPRODUCT((" & mod_VarGlobales.NOM_TABLE_ECHEANCES & "[DatePrevue]<>"""")*((" & _
            mod_VarGlobales.NOM_TABLE_ECHEANCES & "[DatePrevue]+(" & mod_VarGlobales.NOM_TABLE_ECHEANCES & "[PrecisionDate]=""Mois"")*(DATE(YEAR(" & _
            mod_VarGlobales.NOM_TABLE_ECHEANCES & "[DatePrevue]),MONTH(" & mod_VarGlobales.NOM_TABLE_ECHEANCES & "[DatePrevue])+1,0)-" & _
            mod_VarGlobales.NOM_TABLE_ECHEANCES & "[DatePrevue]))<TODAY()))"
    End If
    For k = 3 To 5
        If neuf(k) Then
            With ws.Cells(l + k, 2)
                .NumberFormat = "#,##0.00;[Red]-#,##0.00"
                .Interior.Color = mod_InstallCommun.CoulFondLecture()
                .HorizontalAlignment = xlRight
            End With
        End If
    Next k
    If neuf(6) Then
        With ws.Cells(l + 6, 2)
            .NumberFormat = "0"
            .Interior.Color = mod_InstallCommun.CoulFondLecture()
            .HorizontalAlignment = xlRight
        End With
    End If

    PoserBoutonsSynthese ws, l

End Sub

' Une cellule est à (re)écrire si elle n'a pas de formule, ou si sa formule est cassée (#REF!).
' Cas typique : reconstruire une liste de l'échéancier supprime puis recrée le tableau Excel
' (TblEcheances) ; les formules de Synthese qui le citaient deviennent alors #REF! et sont
' réécrites ici, sans toucher au format ni aux autres cellules.
Private Function FormuleAReparer(ByVal cellule As Range) As Boolean
    If Not cellule.HasFormula Then
        FormuleAReparer = True
    ElseIf InStr(1, cellule.Formula, "#REF!", vbTextCompare) > 0 Then
        FormuleAReparer = True
    End If
End Function

' Ligne du titre du bloc : celle au-dessus de la cellule nommée soldeRefMontant si elle existe
' (le bloc a pu être déplacé), sinon la position par défaut.
Private Function LigneTitreBloc(ByVal ws As Worksheet) As Long
    Dim n As Name
    On Error Resume Next
    Set n = ThisWorkbook.Names("soldeRefMontant")
    If Not n Is Nothing Then LigneTitreBloc = n.RefersToRange.Row - 1
    On Error GoTo 0
    If LigneTitreBloc < 1 Then LigneTitreBloc = SYN_LIGNE_TITRE
End Function

Private Sub AjouterNomSiAbsent(ByVal nom As String, ByVal cellule As Range)
    Dim n As Name
    On Error Resume Next
    Set n = ThisWorkbook.Names(nom)
    On Error GoTo 0
    If n Is Nothing Then
        ThisWorkbook.Names.Add Name:=nom, RefersTo:="=" & cellule.Worksheet.Name & "!" & cellule.Address(True, True)
    End If
End Sub

Private Function FormeSiExiste(ByVal ws As Worksheet, ByVal nom As String) As Shape
    On Error Resume Next
    Set FormeSiExiste = ws.Shapes(nom)
    On Error GoTo 0
End Function

Private Function AjouterBoutonSynthese(ByVal ws As Worksheet, ByVal nom As String, ByVal legende As String, _
                                       ByVal macro As String, ByVal gauche As Double, ByVal haut As Double, _
                                       ByVal largeur As Double, ByVal hauteur As Double) As Shape
    Dim btn As Button
    Set btn = ws.Buttons.Add(gauche, haut, largeur, hauteur)
    btn.Caption = legende
    btn.OnAction = macro
    btn.Name = nom
    Set AjouterBoutonSynthese = ws.Shapes(nom)
End Function

' Boutons du bloc. Les boutons déjà présents ne sont jamais déplacés ni recréés. Le bouton
' "Saisir / contrôler le solde" se place sous "Voir l'échéancier", à la même largeur et au même
' pas vertical que les deux premiers ; le cadre qui entoure ces boutons (Gestion prévisionnelle)
' est agrandi vers le bas si nécessaire pour l'englober.
Private Sub PoserBoutonsSynthese(ByVal ws As Worksheet, ByVal l As Long)

    Dim bS As Shape, bV As Shape, bSolde As Shape
    Dim pas As Double, haut As Double

    Set bS = FormeSiExiste(ws, "btnSaisirEcheance")
    If bS Is Nothing Then
        Set bS = AjouterBoutonSynthese(ws, "btnSaisirEcheance", mod_Display.FR("Saisir une {e2}ch{e2}ance"), "SaisirEcheance", _
                                       ws.Cells(l + 1, 3).Left + 6, ws.Cells(l + 1, 3).Top, 170, 20)
    End If
    Set bV = FormeSiExiste(ws, "btnVoirEcheancier")
    If bV Is Nothing Then
        Set bV = AjouterBoutonSynthese(ws, "btnVoirEcheancier", mod_Display.FR("Voir l'{e2}ch{e2}ancier"), "AfficherEcheancier", _
                                       ws.Cells(l + 3, 3).Left + 6, ws.Cells(l + 3, 3).Top, 170, 20)
    End If

    If FormeSiExiste(ws, "btnSaisirSolde") Is Nothing Then
        pas = bV.Top - bS.Top
        If pas < bV.Height + 4 Then pas = bV.Height + 7
        haut = bV.Top + pas
        Set bSolde = AjouterBoutonSynthese(ws, "btnSaisirSolde", mod_Display.FR("Saisir / contr{o2}ler le solde"), "SaisirSoldeReference", _
                                           bV.Left, haut, bV.Width, bV.Height)
        ElargirCadre ws, bV, haut + bV.Height
    End If

End Sub

' Agrandit vers le bas le plus petit cadre (forme automatique) qui contient la forme de référence,
' pour qu'il englobe aussi un bouton ajouté dessous.
Private Sub ElargirCadre(ByVal ws As Worksheet, ByVal ref As Shape, ByVal basVoulu As Double)

    Dim sh As Shape, cadre As Shape
    Dim aire As Double, aireMin As Double

    For Each sh In ws.Shapes
        If sh.Type = 1 Then        ' 1 = msoAutoShape (cadre arrondi, pas un bouton ni une zone de texte)
            If sh.Left <= ref.Left And sh.Top <= ref.Top And _
               sh.Left + sh.Width >= ref.Left + ref.Width And sh.Top + sh.Height >= ref.Top + ref.Height Then
                aire = sh.Width * sh.Height
                If cadre Is Nothing Or aire < aireMin Then
                    Set cadre = sh
                    aireMin = aire
                End If
            End If
        End If
    Next sh

    If cadre Is Nothing Then Exit Sub
    If cadre.Top + cadre.Height < basVoulu + 8 Then cadre.Height = basVoulu + 8 - cadre.Top

End Sub

' Reconstruit seulement le bloc de la feuille Synthese (formules, écart, boutons), sans toucher
' aux listes ni aux valeurs déjà saisies (solde de référence, date). Sans danger : à lancer après
' une mise à jour du code pour obtenir les nouveaux boutons ou formules.
Public Sub MettreAJourBlocSynthese()
    ConstruireBlocSynthese
    MsgBox mod_Display.FR("Le bloc Pr{e2}visionnel de la feuille Synthese a {e2}t{e2} mis {a2} jour."), vbInformation
End Sub

' =====================================================================================
' CONTRÔLE SILENCIEUX DES NUMÉROS DE CHÈQUE (colonne NumCheque de TblEcheances)
' =====================================================================================
' Aucune macro, aucun message : une mise en forme conditionnelle colore la cellule, en direct,
' à chaque modification du tableau.
'   ROUGE  : numéro invalide (pas uniquement des chiffres) ou DÉJÀ PASSÉ EN BANQUE
'            (présent dans la colonne Num_Cheque de TblOperations).
'   ORANGE : numéro présent plusieurs fois dans l'échéancier.
' Les numéros sont comparés comme des nombres : 0001234 et 1234 sont le même chèque.
' Deux noms définis servent aux formules (une mise en forme conditionnelle n'accepte pas
' les références de tableau directes) : ChequesEcheances et ChequesBanque.
Private Sub PoserControleCheques(ByVal ws As Worksheet, ByVal lo As ListObject)

    Dim rng As Range
    Dim ref As String, fRouge As String, fOrange As String
    Dim fc As Object

    On Error GoTo Fin      ' ne jamais bloquer l'installation pour une mise en forme

    ThisWorkbook.Names.Add Name:="ChequesEcheances", RefersTo:="=" & lo.Name & "[NumCheque]"
    On Error Resume Next
    ThisWorkbook.Names.Add Name:="ChequesBanque", RefersTo:="=TblOperations[Num_Cheque]"
    If Err.Number <> 0 Then
        Err.Clear
        ThisWorkbook.Names.Add Name:="ChequesBanque", RefersTo:="=""""", Visible:=True   ' pas de tableau d'opérations : liste vide
    End If
    On Error GoTo Fin

    Set rng = lo.ListColumns("NumCheque").DataBodyRange

    ' Les références relatives d'une mise en forme conditionnelle posée par programme sont
    ' interprétées par rapport à la cellule active : on se place donc sur la première cellule.
    ws.Activate
    rng.Cells(1, 1).Select
    ref = rng.Cells(1, 1).Address(False, False)

    rng.FormatConditions.Delete

    fRouge = "=AND(" & ref & "<>"""",IF(ISERROR(--" & ref & "),TRUE," & _
             "SUMPRODUCT(--(IFERROR(--ChequesBanque,-1)=--" & ref & "))>0))"
    fOrange = "=AND(" & ref & "<>"""",NOT(ISERROR(--" & ref & ")),SUMPRODUCT(--(IFERROR(--ChequesEcheances,-1)=--" & ref & "))>1)"

    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:=fRouge)
    fc.Interior.Color = RGB(255, 199, 206)
    fc.Font.Color = RGB(156, 0, 6)
    fc.StopIfTrue = True

    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:=fOrange)
    fc.Interior.Color = RGB(255, 235, 156)
    fc.Font.Color = RGB(156, 87, 0)

Fin:
End Sub

' Pose (ou repose) seulement le contrôle des chèques sur une liste déjà installée, sans la
' reconstruire : à lancer après une mise à jour du code. Sans danger pour les données.
Public Sub MettreAJourControleCheques()

    Dim ws As Worksheet, wsPrec As Worksheet
    Dim lo As ListObject
    Dim etat As XlSheetVisibility

    Set ws = mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_ECHEANCIER)
    If ws Is Nothing Then
        MsgBox mod_Display.FR("L'{e2}ch{e2}ancier n'est pas install{e2}. Ex{e2}cutez d'abord la macro InstallerEcheancier."), vbExclamation
        Exit Sub
    End If
    On Error Resume Next
    Set lo = ws.ListObjects(mod_VarGlobales.NOM_TABLE_ECHEANCES)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    If mod_Echeancier.ColonneAbsente(lo, "NumCheque") Then
        MsgBox mod_Display.FR("La colonne NumCheque n'existe pas encore : lancez InstallerEcheancier."), vbExclamation
        Exit Sub
    End If

    Set wsPrec = ActiveSheet
    etat = ws.Visible
    ws.Visible = xlSheetVisible
    PoserControleCheques ws, lo
    ws.Visible = etat
    On Error Resume Next
    wsPrec.Activate
    On Error GoTo 0

    MsgBox mod_Display.FR("Le contr{o2}le des num{e2}ros de ch{e1}que est en place."), vbInformation

End Sub

Private Sub SupprimerBoutonSiExiste(ByVal ws As Worksheet, ByVal nom As String)
    On Error Resume Next
    ws.Shapes(nom).Delete
    On Error GoTo 0
End Sub


' =====================================================================================
' Bouton "Rendre récurrente" de l'écran de recherche
' =====================================================================================
' Ajouté à la suite des boutons existants de frm_RechercheOperations. Aussi appelée par
' InstallerEcheancier : la feuille de recherche n'est donc pas reconstruite. Sans danger
' si le bouton existe déjà (il est remplacé).
Public Sub AjouterBoutonRecurrenteRO()

    Dim ws As Worksheet
    Dim gauche As Double
    Dim ref As Shape

    Set ws = mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_RECHERCHE)
    If ws Is Nothing Then Exit Sub

    SupprimerBoutonSiExiste ws, "btnRendreRecurrenteRO"

    On Error Resume Next
    Set ref = ws.Shapes("btnDecalerBudgetOperationRO")
    On Error GoTo 0
    If ref Is Nothing Then
        gauche = ws.Cells(1, 1).Left + 575
    Else
        gauche = ref.Left + ref.Width + mod_InstallCommun.FRM_BTN_ECART
    End If

    mod_InstallCommun.AjouterBoutonEntete ws, mod_InstallRechercheOperations.RO_LIGNE_BOUTONS, gauche, _
        mod_Display.FR("Rendre r{e2}currente"), "RendreRecurrenteRO", "btnRendreRecurrenteRO", 105

End Sub


' =====================================================================================
' OUTILS DÉVELOPPEUR (Ctrl+G)
' =====================================================================================
Public Sub AfficherFormulaireEcheancePourEdition()
    mod_InstallCommun.AfficherPourEdition mod_VarGlobales.NOM_FEUILLE_ECHEANCE, "InstallerEcheancier", "MasquerFormulaireEcheanceApresEdition"
End Sub

Public Sub MasquerFormulaireEcheanceApresEdition()
    mod_InstallCommun.MasquerApresEdition mod_VarGlobales.NOM_FEUILLE_ECHEANCE
End Sub
