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
                    "DatePrevue", "PrecisionDate", "Periodicite", "TolerancePct")
    largeurs = Array(12, 13, 28, 24, 20, 24, 11, 12, 13, 14, 12)
    formats = Array("0", "0", "@", "@", "@", "@", "#,##0.00", "dd/mm/yyyy", "@", "@", "0")

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

    ' Garde-fou : ne jamais écraser des données sans le dire.
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
            If MsgBox(mod_Display.FR("La feuille '") & nomFeuille & mod_Display.FR("' contient d{e2}j{a2} des donn{e2}es.") & vbCrLf & _
                      mod_Display.FR("La reconstruire les EFFACERAIT. Continuer ?"), _
                      vbYesNo + vbExclamation, mod_Display.FR("Donn{e2}es existantes")) = vbNo Then Exit Function
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

    mod_InstallCommun.PoserEtiquette ws.Range("B11"), mod_Display.FR("P{e2}riodicit{e2}")
    mod_InstallCommun.PoserChampSaisie ws.Range(mod_Echeancier.EF_ADR_PERIODE), "@"

    mod_InstallCommun.PoserEtiquette ws.Range("B12"), mod_Display.FR("Date pr{e2}vue (ou 1{e1}re date)")
    mod_InstallCommun.PoserChampSaisie ws.Range(mod_Echeancier.EF_ADR_DATE), "dd/mm/yyyy"

    mod_InstallCommun.PoserEtiquette ws.Range("B13"), mod_Display.FR("Date de fin (facultative)")
    mod_InstallCommun.PoserChampSaisie ws.Range(mod_Echeancier.EF_ADR_DATEFIN), "dd/mm/yyyy"

    mod_InstallCommun.PoserEtiquette ws.Range("B14"), mod_Display.FR("Pr{e2}cision de la date")
    mod_InstallCommun.PoserChampSaisie ws.Range(mod_Echeancier.EF_ADR_PRECISION), "@"

    mod_InstallCommun.PoserEtiquette ws.Range("B15"), mod_Display.FR("Tol{e2}rance sur le montant (%)")
    mod_InstallCommun.PoserChampSaisie ws.Range(mod_Echeancier.EF_ADR_TOLERANCE), "0"

    ws.Rows(16).RowHeight = mod_InstallCommun.FRM_H_SEP

    mod_InstallCommun.PoserMessage ws.Range("B17:E17")
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
    t = t & Chr(10) & "Reconnaissance : renseignez au moins le [c:Tiers] ou la [c:Cat{e2}gorie] (les deux si vous voulez). Le [c:Tiers] se choisit dans la liste "
    t = t & "des Tiers d{e2}j{a2} connus dans les op{e2}rations, ou se tape librement : il est cherch{e2} dans le Tiers de l'op{e2}ration import{e2}e, "
    t = t & "sans tenir compte des majuscules, donc un mot-cl{e2} suffit (ex. EDF). La [c:Cat{e2}gorie] et la [c:Sous-cat{e2}gorie] se choisissent "
    t = t & "dans les listes (la liste des sous-cat{e2}gories d{e2}pend de la cat{e2}gorie choisie ; sans cat{e2}gorie, elles sont toutes propos{e2}es, "
    t = t & "et choisir une sous-cat{e2}gorie qui n'appartient qu'{a2} une cat{e2}gorie renseigne celle-ci). "
    t = t & "Si vous renseignez plusieurs crit{e1}res (Tiers, Cat{e2}gorie, Sous-cat{e2}gorie), l'op{e2}ration doit tous les respecter."
    t = t & Chr(10) & "[c:Montant] : n{e2}gatif pour une d{e2}pense, positif pour une entr{e2}e d'argent ; une op{e2}ration de signe contraire n'est jamais rapproch{e2}e."
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
    Dim l As Long
    Dim btn As Button
    Dim coulBloc As Long
    Dim f As String

    Set ws = mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_SYNTHESE)
    If ws Is Nothing Then
        MsgBox mod_Display.FR("La feuille Synthese est introuvable : le bloc pr{e2}visionnel n'a pas {e2}t{e2} cr{e2}{e2}."), vbExclamation
        Exit Sub
    End If

    l = SYN_LIGNE_TITRE
    coulBloc = RGB(141, 180, 226)   ' même bleu que les paramètres de la feuille

    ws.Range(ws.Cells(l, 1), ws.Cells(l + 6, 2)).Interior.Color = coulBloc
    ws.Cells(l, 1).Value = mod_Display.FR("Pr{e2}visionnel de tr{e2}sorerie")
    ws.Cells(l, 1).Font.Bold = True
    ws.Cells(l + 1, 1).Value = mod_Display.FR("Solde de r{e2}f.")
    ws.Cells(l + 2, 1).Value = "Date du solde"
    ws.Cells(l + 3, 1).Value = "Solde actuel"
    ws.Cells(l + 4, 1).Value = mod_Display.FR("Solde {a2} J+15")
    ws.Cells(l + 5, 1).Value = mod_Display.FR("Solde {a2} J+30")
    ws.Cells(l + 6, 1).Value = mod_Display.FR("{E2}ch. en retard")

    ' Cellules de saisie
    With ws.Cells(l + 1, 2)
        .NumberFormat = "#,##0.00"
        .Interior.Color = mod_InstallCommun.CoulFondSaisie()
        .Locked = False
    End With
    With ws.Cells(l + 2, 2)
        .NumberFormat = "dd/mm/yyyy"
        .Interior.Color = mod_InstallCommun.CoulFondSaisie()
        .Locked = False
    End With
    ThisWorkbook.Names.Add Name:="soldeRefMontant", RefersTo:="=" & ws.Name & "!" & ws.Cells(l + 1, 2).Address(True, True)
    ThisWorkbook.Names.Add Name:="soldeRefDate", RefersTo:="=" & ws.Name & "!" & ws.Cells(l + 2, 2).Address(True, True)

    ' Formules (syntaxe anglaise pour .Formula, quelle que soit la langue d'Excel)
    f = "=IF(OR(soldeRefMontant="""",soldeRefDate=""""),""" & mod_Display.FR("{a2} renseigner") & """," & _
        "soldeRefMontant+SUMIFS(" & "TblOperations" & "[Montant]," & _
        "TblOperations" & "[Date_Comptable],"">""&soldeRefDate))"
    ws.Cells(l + 3, 2).Formula = f

    ws.Cells(l + 4, 2).Formula = "=IF(ISNUMBER(" & ws.Cells(l + 3, 2).Address(False, False) & ")," & _
        ws.Cells(l + 3, 2).Address(False, False) & "+SUMIFS(" & mod_VarGlobales.NOM_TABLE_ECHEANCES & "[Montant]," & _
        mod_VarGlobales.NOM_TABLE_ECHEANCES & "[DatePrevue],""<=""&(TODAY()+15)),"""")"
    ws.Cells(l + 5, 2).Formula = "=IF(ISNUMBER(" & ws.Cells(l + 3, 2).Address(False, False) & ")," & _
        ws.Cells(l + 3, 2).Address(False, False) & "+SUMIFS(" & mod_VarGlobales.NOM_TABLE_ECHEANCES & "[Montant]," & _
        mod_VarGlobales.NOM_TABLE_ECHEANCES & "[DatePrevue],""<=""&(TODAY()+30)),"""")"
    ' En retard = date prévue passée ; pour une précision "Mois", seulement une fois le mois écoulé.
    ws.Cells(l + 6, 2).Formula = "=SUMPRODUCT((" & mod_VarGlobales.NOM_TABLE_ECHEANCES & "[DatePrevue]<>"""")*((" & _
        mod_VarGlobales.NOM_TABLE_ECHEANCES & "[DatePrevue]+(" & mod_VarGlobales.NOM_TABLE_ECHEANCES & "[PrecisionDate]=""Mois"")*(DATE(YEAR(" & _
        mod_VarGlobales.NOM_TABLE_ECHEANCES & "[DatePrevue]),MONTH(" & mod_VarGlobales.NOM_TABLE_ECHEANCES & "[DatePrevue])+1,0)-" & _
        mod_VarGlobales.NOM_TABLE_ECHEANCES & "[DatePrevue]))<TODAY()))"

    With ws.Range(ws.Cells(l + 3, 2), ws.Cells(l + 5, 2))
        .NumberFormat = "#,##0.00;[Red]-#,##0.00"
        .Interior.Color = mod_InstallCommun.CoulFondLecture()
        .HorizontalAlignment = xlRight
    End With
    With ws.Cells(l + 6, 2)
        .NumberFormat = "0"
        .Interior.Color = mod_InstallCommun.CoulFondLecture()
        .HorizontalAlignment = xlRight
    End With

    ' Boutons
    SupprimerBoutonSiExiste ws, "btnSaisirEcheance"
    SupprimerBoutonSiExiste ws, "btnVoirEcheancier"
    Set btn = ws.Buttons.Add(ws.Cells(l + 1, 3).Left + 6, ws.Cells(l + 1, 3).Top, 170, 20)
    btn.Caption = mod_Display.FR("Saisir une {e2}ch{e2}ance")
    btn.OnAction = "SaisirEcheance"
    btn.Name = "btnSaisirEcheance"
    Set btn = ws.Buttons.Add(ws.Cells(l + 3, 3).Left + 6, ws.Cells(l + 3, 3).Top, 170, 20)
    btn.Caption = mod_Display.FR("Voir l'{e2}ch{e2}ancier")
    btn.OnAction = "AfficherEcheancier"
    btn.Name = "btnVoirEcheancier"

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
