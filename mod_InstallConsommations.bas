Attribute VB_Name = "mod_InstallConsommations"
Option Explicit

' =====================================================================================
' MODULE : mod_InstallConsommations
'
' SUIVI DES CONSOMMATIONS (eau, électricité, gaz) - construction des feuilles.
' La logique est dans mod_Consommations.
'
' Une seule macro à lancer : InstallerConsommations. Elle crée, dans l'ordre :
'     1. frm_Releve      : le formulaire de saisie d'un relevé.
'     2. frm_Contrat     : le formulaire de saisie d'un contrat (opérateur).
'     3. frm_Contrats    : la liste des contrats (tableau TblContrats).
'     4. frm_Releves     : la liste des relevés (tableau TblReleves).
'     5. frm_Consommations : l'écran principal (paramètres, tableau mensuel, graphique).
'     6. le bouton "Consommations" sur la feuille Synthese.
'   Les 5 feuilles frm_* sont masquées : on les ouvre avec les boutons.
'
' INSTALLATION (une seule fois) :
'   1. Alt+F11, Fichier > Importer : mod_VarGlobales.bas (mis à jour), mod_Consommations.bas
'      et CE fichier.
'   2. Ctrl+G, taper : InstallerConsommations, puis Entrée.
'   3. Coller dans le module de la feuille frm_Releve le code de CodeBehind_frm_Releve.txt,
'      et dans celui de la feuille frm_Contrat le code de CodeBehind_frm_Contrat.txt (le
'      message de fin d'installation donne les CodeName).
'   4. Sur Synthese, cliquer sur le bouton "Consommations", puis sur Contrats pour
'      déclarer les contrats en cours, puis sur Saisir un relevé.
'
' RÉINSTALLATION : possible. Si les listes contiennent des données, une confirmation
' est demandée avant de les reconstruire (les données seraient PERDUES). L'écran
' principal, lui, revient à ses paramètres par défaut.
' =====================================================================================

Private Const LIGNE_BOUTONS As Long = 2
Private Const LIGNE_INSTRUCTIONS As Long = 4
Private Const LIGNE_CORPS As Long = 6
Private Const LIGNE_ENTETE_LISTE As Long = 6
Private Const SYN_LIGNE_BOUTON As Long = 25


' =====================================================================================
' MACRO PRINCIPALE
' =====================================================================================
Public Sub InstallerConsommations()

    Dim wsPrecedente As Worksheet

    Set wsPrecedente = ActiveSheet

    If Not ConstruireFormulaireReleve(wsPrecedente) Then Exit Sub
    If Not ConstruireFormulaireContrat(wsPrecedente) Then Exit Sub
    If Not ConstruireListeContrats(wsPrecedente) Then Exit Sub
    If Not ConstruireListeReleves(wsPrecedente) Then Exit Sub
    If Not ConstruireEcranConso(wsPrecedente) Then Exit Sub
    ConstruireBoutonSynthese

    MsgBox mod_Display.FR("Le suivi des consommations est install{e2}.") & vbCrLf & vbCrLf & _
           mod_Display.FR("PROCHAINE {E2}TAPE (indispensable) : coller le code des {e2}v{e2}nements dans les deux feuilles de formulaire.") & vbCrLf & _
           "Nom interne (CodeName) de la feuille " & mod_VarGlobales.NOM_FEUILLE_RELEVE & " : " & _
           mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_RELEVE).CodeName & vbCrLf & _
           "Nom interne (CodeName) de la feuille " & mod_VarGlobales.NOM_FEUILLE_CONTRAT & " : " & _
           mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_CONTRAT).CodeName & vbCrLf & vbCrLf & _
           mod_Display.FR("Voir les fichiers CodeBehind_frm_Releve.txt et CodeBehind_frm_Contrat.txt."), _
           vbInformation, mod_Display.FR("Installation termin{e2}e")

End Sub


' =====================================================================================
' LISTES (relevés, contrats)
' =====================================================================================
Private Function ConstruireListeContrats(ByVal wsPrecedente As Worksheet) As Boolean

    Dim entetes As Variant, largeurs As Variant, formats As Variant, boutons As Variant
    Dim t As String

    entetes = Array("ID_Contrat", "TypeConso", "Operateur", "Unite", "DateDebut", "DateFin", "AbonnementMensuel", "Reference")
    largeurs = Array(11, 18, 24, 9, 12, 12, 18, 24)
    formats = Array("0", "@", "@", "@", "dd/mm/yyyy", "dd/mm/yyyy", "#,##0.00", "@")

    boutons = Array( _
        mod_Display.FR("Ajouter un contrat"), "AjouterContrat", "btnAjouterContrat", 118, _
        mod_Display.FR("Supprimer la s{e2}lection"), "SupprimerContratSelection", "btnSupprimerContrat", 118, _
        mod_Display.FR("Voir les relev{e2}s"), "AfficherReleves", "btnVoirRelevesC", 108, _
        "Retour", "RetourConsommations", "btnRetourContrats", mod_InstallCommun.FRM_BTN_L)

    t = "Cette liste contient les contrats de fourniture : un type de consommation chez un op{e2}rateur, entre deux dates. "
    t = t & "Le contrat donne {a2} chaque relev{e2} son op{e2}rateur, son unit{e2} (m3, kWh...) et l'abonnement mensuel propos{e2} par d{e2}faut."
    t = t & Chr(10) & "Pour changer d'op{e2}rateur, ajoutez simplement le nouveau contrat avec [b:Ajouter un contrat] : s'il y a un contrat en cours "
    t = t & "pour ce type, le programme propose de le cl{o2}turer la veille. L'historique reste complet. "
    t = t & "Vous pouvez corriger directement les cellules ([c:DateFin] vide = contrat en cours) ; deux contrats d'un m{e1}me type ne doivent pas se chevaucher."
    t = t & Chr(10) & "[c:AbonnementMensuel] n'est qu'une valeur de d{e2}part : l'abonnement r{e2}el est port{e2} par chaque relev{e2} et peut {e2}voluer en cours de contrat."
    t = t & Chr(10) & "Boutons :"
    t = t & Chr(10) & "- [b:Ajouter un contrat] : ouvre le formulaire de saisie."
    t = t & Chr(10) & "- [b:Supprimer la s{e2}lection] : supprime les lignes s{e2}lectionn{e2}es (les relev{e2}s rattach{e2}s sont conserv{e2}s)."
    t = t & Chr(10) & "- [b:Voir les relev{e2}s] : affiche la liste des relev{e2}s."
    t = t & Chr(10) & "- [b:Retour] : revient sur [f:frm_Consommations]."

    ConstruireListeContrats = ConstruireListe(mod_VarGlobales.NOM_FEUILLE_CONTRATS, mod_VarGlobales.NOM_TABLE_CONTRATS, _
                                              entetes, largeurs, formats, t, boutons, wsPrecedente)

End Function

Private Function ConstruireListeReleves(ByVal wsPrecedente As Worksheet) As Boolean

    Dim entetes As Variant, largeurs As Variant, formats As Variant, boutons As Variant
    Dim t As String

    entetes = Array("ID_Releve", "TypeConso", "Operateur", "DateDebut", "DateFin", "Quantite", "PrixUnitaire", _
                    "AbonnementMensuel", "NbMois", "Abonnement", "Montant", "Facture", "Commentaire")
    largeurs = Array(10, 18, 22, 12, 12, 11, 12, 18, 9, 12, 12, 16, 28)
    formats = Array("0", "@", "@", "dd/mm/yyyy", "dd/mm/yyyy", "#,##0.00", "#,##0.0000", "#,##0.00", "0.00", "#,##0.00", "#,##0.00", "@", "@")

    boutons = Array( _
        mod_Display.FR("Ajouter un relev{e2}"), "SaisirReleve", "btnAjouterReleve", 118, _
        mod_Display.FR("Supprimer la s{e2}lection"), "SupprimerReleveSelection", "btnSupprimerReleve", 118, _
        "Contrats", "AfficherContrats", "btnVoirContratsR", 80, _
        "Retour", "RetourConsommations", "btnRetourReleves", mod_InstallCommun.FRM_BTN_L)

    t = "Cette liste contient tous les relev{e2}s de consommation. Chaque ligne est une p{e2}riode (de [c:DateDebut] {a2} [c:DateFin] incluse) avec sa quantit{e2}, "
    t = t & "SON prix d'unit{e2} et SON abonnement mensuel : une facture dont le prix change plusieurs fois se saisit en plusieurs lignes cons{e2}cutives "
    t = t & "(un prix par ligne), avec le m{e1}me [c:Facture]."
    t = t & Chr(10) & "Vous pouvez corriger directement les cellules jaunes. Les colonnes grises sont calcul{e2}es : [c:Operateur] (contrat en vigueur {a2} la date de d{e2}but), "
    t = t & "[c:NbMois] (un mois complet vaut 1 ; une p{e2}riode partielle compte ses jours au prorata du mois), [c:Abonnement] (abonnement mensuel x NbMois) "
    t = t & "et [c:Montant] (quantit{e2} x prix + abonnement)."
    t = t & Chr(10) & "Le formulaire v{e2}rifie qu'une p{e2}riode ne chevauche pas un autre relev{e2} du m{e1}me type et ne d{e2}borde pas de son contrat ; "
    t = t & "ces contr{o2}les ne s'appliquent pas {a2} une correction faite directement dans la liste."
    t = t & Chr(10) & "Boutons :"
    t = t & Chr(10) & "- [b:Ajouter un relev{e2}] : ouvre le formulaire de saisie."
    t = t & Chr(10) & "- [b:Supprimer la s{e2}lection] : supprime les lignes s{e2}lectionn{e2}es."
    t = t & Chr(10) & "- [b:Contrats] : affiche la liste des contrats."
    t = t & Chr(10) & "- [b:Retour] : revient sur [f:frm_Consommations]."

    ConstruireListeReleves = ConstruireListe(mod_VarGlobales.NOM_FEUILLE_RELEVES, mod_VarGlobales.NOM_TABLE_RELEVES, _
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
    Dim col As Variant

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

    ' Particularités de chaque tableau.
    If nomTable = mod_VarGlobales.NOM_TABLE_RELEVES Then
        mod_Consommations.PoserFormulesReleves lo
        ' Colonnes calculées : fond gris (non modifiables).
        For Each col In Array("Operateur", "NbMois", "Abonnement", "Montant")
            lo.ListColumns(CStr(col)).DataBodyRange.Interior.Color = mod_InstallCommun.CoulFondLecture()
        Next col
    End If
    PoserListeTypeSurColonne lo

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

' Liste déroulante des types sur toute la colonne TypeConso d'un tableau.
Private Sub PoserListeTypeSurColonne(ByVal lo As ListObject)
    Dim rng As Range
    Set rng = lo.ListColumns("TypeConso").DataBodyRange
    On Error Resume Next
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=ListeTypesConso"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    On Error GoTo 0
End Sub


' =====================================================================================
' FORMULAIRE frm_Releve
' =====================================================================================
Private Function ConstruireFormulaireReleve(ByVal wsPrecedente As Worksheet) As Boolean

    Dim ws As Worksheet
    Dim gauche As Double
    Dim i As Long, tps As Variant

    Set ws = mod_InstallCommun.PreparerFeuille(mod_VarGlobales.NOM_FEUILLE_RELEVE)
    If ws Is Nothing Then Exit Function

    Application.ScreenUpdating = False

    mod_InstallCommun.MettreEnForme ws, Array(2, 32, 30, 20, 20, 2)

    If mod_InstallCommun.PremiereLigneCorps(1, False) <> LIGNE_CORPS Or _
       mod_InstallCommun.LigneInstructions(1, False) <> LIGNE_INSTRUCTIONS Then
        MsgBox "mod_InstallConsommations : constantes de lignes incoherentes avec mod_InstallCommun.", vbCritical
    End If
    mod_InstallCommun.PoserHauteursEntete ws, 1, False
    mod_InstallCommun.EcrireInstructions ws, LIGNE_INSTRUCTIONS, 2, 2, 3, 5, TexteInstructionsReleve()

    ' --- Champs ---
    mod_InstallCommun.PoserEtiquette ws.Range("B6"), "Type de consommation"
    mod_InstallCommun.PoserChampSaisie ws.Range(mod_Consommations.RF_ADR_TYPE), "@"

    mod_InstallCommun.PoserEtiquette ws.Range("B7"), mod_Display.FR("Op{e2}rateur (contrat)")
    mod_InstallCommun.PoserChampLecture ws.Range("C7:E7"), "@"

    mod_InstallCommun.PoserEtiquette ws.Range("B8"), mod_Display.FR("D{e2}but de la p{e2}riode")
    mod_InstallCommun.PoserChampDate ws.Range(mod_Consommations.RF_ADR_DEBUT)

    mod_InstallCommun.PoserEtiquette ws.Range("B9"), mod_Display.FR("Fin de la p{e2}riode (incluse)")
    mod_InstallCommun.PoserChampDate ws.Range(mod_Consommations.RF_ADR_FIN)

    mod_InstallCommun.PoserEtiquette ws.Range("B10"), mod_Display.FR("Quantit{e2} consomm{e2}e")
    mod_InstallCommun.PoserChampMontant ws.Range(mod_Consommations.RF_ADR_QTE), "#,##0.00", mod_InstallCommun.MONTANT_POSITIF_OU_NUL
    mod_InstallCommun.PoserChampLecture ws.Range(mod_Consommations.RF_ADR_UNITE), "@"

    mod_InstallCommun.PoserEtiquette ws.Range("B11"), mod_Display.FR("Prix de l'unit{e2}")
    mod_InstallCommun.PoserChampMontant ws.Range(mod_Consommations.RF_ADR_PRIX), "#,##0.0000", mod_InstallCommun.MONTANT_POSITIF_OU_NUL
    mod_InstallCommun.PoserChampLecture ws.Range(mod_Consommations.RF_ADR_UNITE_PRIX), "@"

    mod_InstallCommun.PoserEtiquette ws.Range("B12"), "Abonnement mensuel"
    mod_InstallCommun.PoserChampMontant ws.Range(mod_Consommations.RF_ADR_ABO), "#,##0.00", mod_InstallCommun.MONTANT_POSITIF_OU_NUL
    mod_InstallCommun.PoserChampLecture ws.Range(mod_Consommations.RF_ADR_UNITE_ABO), "@"

    mod_InstallCommun.PoserEtiquette ws.Range("B13"), mod_Display.FR("N{o2} de facture (facultatif)")
    mod_InstallCommun.PoserChampSaisie ws.Range("C13:E13"), "@"

    mod_InstallCommun.PoserEtiquette ws.Range("B14"), "Commentaire (facultatif)"
    mod_InstallCommun.PoserChampSaisie ws.Range("C14:E14"), "@"

    ws.Rows(15).RowHeight = mod_InstallCommun.FRM_H_SEP

    ' --- Montants calculés (formules, mises à jour en direct) ---
    mod_InstallCommun.PoserEtiquette ws.Range("B16"), "Consommation (" & ChrW(8364) & ")"
    mod_InstallCommun.PoserChampLecture ws.Range(mod_Consommations.RF_ADR_CALC_CONSO), "#,##0.00"
    ws.Range(mod_Consommations.RF_ADR_CALC_CONSO).Formula = _
        "=IF(AND(ISNUMBER(" & mod_Consommations.RF_ADR_QTE & "),ISNUMBER(" & mod_Consommations.RF_ADR_PRIX & ")),ROUND(" & _
        mod_Consommations.RF_ADR_QTE & "*" & mod_Consommations.RF_ADR_PRIX & ",2),"""")"

    mod_InstallCommun.PoserEtiquette ws.Range("B17"), mod_Display.FR("Abonnement de la p{e2}riode (") & ChrW(8364) & ")"
    mod_InstallCommun.PoserChampLecture ws.Range(mod_Consommations.RF_ADR_CALC_ABO), "#,##0.00"
    ws.Range(mod_Consommations.RF_ADR_CALC_ABO).Formula = _
        "=IF(AND(ISNUMBER(" & mod_Consommations.RF_ADR_DEBUT & "),ISNUMBER(" & mod_Consommations.RF_ADR_FIN & ")),IF(" & _
        mod_Consommations.RF_ADR_FIN & "<" & mod_Consommations.RF_ADR_DEBUT & ","""",ROUND(N(" & mod_Consommations.RF_ADR_ABO & ")*" & _
        mod_Consommations.FormuleNbMois(mod_Consommations.RF_ADR_DEBUT, mod_Consommations.RF_ADR_FIN) & ",2)),"""")"

    mod_InstallCommun.PoserEtiquette ws.Range("B18"), mod_Display.FR("Total de la p{e2}riode (") & ChrW(8364) & ")"
    mod_InstallCommun.PoserChampLecture ws.Range(mod_Consommations.RF_ADR_CALC_TOTAL), "#,##0.00"
    ws.Range(mod_Consommations.RF_ADR_CALC_TOTAL).Formula = _
        "=IF(AND(ISNUMBER(" & mod_Consommations.RF_ADR_CALC_CONSO & "),ISNUMBER(" & mod_Consommations.RF_ADR_CALC_ABO & ")),(" & _
        mod_Consommations.RF_ADR_CALC_CONSO & ")+(" & mod_Consommations.RF_ADR_CALC_ABO & "),"""")"
    ws.Range(mod_Consommations.RF_ADR_CALC_TOTAL).Font.Bold = True

    ws.Rows(19).RowHeight = mod_InstallCommun.FRM_H_SEP

    mod_InstallCommun.PoserMessage ws.Range("B20:E20")
    ws.Range(mod_Consommations.RF_ADR_MESSAGE).Font.Color = RGB(192, 80, 0)

    ' --- Liste technique des types (colonne masquée) ---
    ws.Cells(1, mod_Consommations.RF_COL_AIDE_TYPE).Value = "Types"
    tps = mod_Consommations.TypesConso()
    For i = 0 To UBound(tps)
        ws.Cells(2 + i, mod_Consommations.RF_COL_AIDE_TYPE).Value = tps(i)
    Next i
    PoserNomListe "ListeTypesConso", ws, mod_Consommations.RF_COL_AIDE_TYPE, 2, 1 + UBound(tps) + 1
    ws.Columns(mod_Consommations.RF_COL_AIDE_TYPE).Hidden = True
    mod_InstallCommun.PoserListeDeroulante ws.Range(mod_Consommations.RF_ADR_TYPE), "ListeTypesConso"

    ' --- Boutons de la ligne d'en-tête ---
    gauche = ws.Cells(1, 2).Left
    mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, mod_InstallCommun.CapValider(), _
        "RelValider", "btnRelValider", mod_InstallCommun.FRM_BTN_L
    mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, mod_Display.FR("Valider + suivant"), _
        "RelValiderSuivant", "btnRelValiderSuivant", 115
    mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, mod_InstallCommun.CapAnnuler(), _
        "RelAnnuler", "btnRelAnnuler", mod_InstallCommun.FRM_BTN_L

    mod_InstallCommun.TerminerFeuille ws, wsPrecedente
    Application.ScreenUpdating = True
    ConstruireFormulaireReleve = True

End Function

Private Function TexteInstructionsReleve() As String

    Dim t As String

    t = "Ce formulaire enregistre une p{e2}riode de consommation avec sa quantit{e2} et son prix. Choisissez d'abord le [c:Type de consommation] : "
    t = t & "l'op{e2}rateur et l'unit{e2} s'affichent d'apr{e1}s le contrat en vigueur {a2} la date de d{e2}but, et le formulaire propose la date de d{e2}but "
    t = t & "(lendemain du dernier relev{e2}), le dernier prix et le dernier abonnement connus pour ce contrat. Vous pouvez tout modifier."
    t = t & Chr(10) & "[c:D{e2}but de la p{e2}riode] et [c:Fin de la p{e2}riode] : dates libres, la fin est incluse. Une p{e2}riode ne peut pas chevaucher un autre relev{e2} du m{e1}me type, "
    t = t & "ni d{e2}border de son contrat. Pour une facture dont le prix de l'unit{e2} change plusieurs fois, saisissez UN relev{e2} par prix "
    t = t & "(chacun avec ses dates, sa quantit{e2} et son prix) : le bouton [b:Valider + suivant] enregistre le relev{e2} puis rouvre le formulaire "
    t = t & "sur la p{e2}riode suivante, avec les m{e2}mes prix et le m{e2}me num{e2}ro de facture."
    t = t & Chr(10) & "[c:Quantit{e2} consomm{e2}e] : en unit{e2}s du contrat (m3, kWh...). [c:Prix de l'unit{e2}] : prix hors abonnement, en euros par unit{e2} ; 0 est accept{e2} (eau recycl{e2}e, par exemple). "
    t = t & "[c:Abonnement mensuel] : montant par mois ; il peut changer d'un relev{e2} {a2} l'autre, et vous pouvez l'ajuster ici. "
    t = t & "L'abonnement de la p{e2}riode se calcule au prorata : un mois complet vaut 1 mois, une p{e2}riode partielle compte ses jours rapport{e2}s au mois, "
    t = t & "donc des relev{e2}s qui se partagent un m{e2}me mois totalisent exactement un abonnement."
    t = t & Chr(10) & "Les trois montants du bas se calculent en direct. Si un champ est incorrect, un message appara{i2}t sous le formulaire et rien n'est enregistr{e2}."
    t = t & Chr(10) & "Boutons :"
    t = t & Chr(10) & "- [b:Valider] : enregistre le relev{e2} puis referme le formulaire."
    t = t & Chr(10) & "- [b:Valider + suivant] : enregistre puis rouvre le formulaire sur la p{e2}riode suivante."
    t = t & Chr(10) & "- [b:Annuler] : referme le formulaire sans rien enregistrer."

    TexteInstructionsReleve = t

End Function


' =====================================================================================
' FORMULAIRE frm_Contrat
' =====================================================================================
Private Function ConstruireFormulaireContrat(ByVal wsPrecedente As Worksheet) As Boolean

    Dim ws As Worksheet
    Dim gauche As Double

    Set ws = mod_InstallCommun.PreparerFeuille(mod_VarGlobales.NOM_FEUILLE_CONTRAT)
    If ws Is Nothing Then Exit Function

    Application.ScreenUpdating = False

    mod_InstallCommun.MettreEnForme ws, Array(2, 32, 30, 20, 20, 2)

    mod_InstallCommun.PoserHauteursEntete ws, 1, False
    mod_InstallCommun.EcrireInstructions ws, LIGNE_INSTRUCTIONS, 2, 2, 3, 5, TexteInstructionsContrat()

    mod_InstallCommun.PoserEtiquette ws.Range("B6"), "Type de consommation"
    mod_InstallCommun.PoserChampSaisie ws.Range(mod_Consommations.CF_ADR_TYPE), "@"

    mod_InstallCommun.PoserEtiquette ws.Range("B7"), mod_Display.FR("Op{e2}rateur (fournisseur)")
    mod_InstallCommun.PoserChampSaisie ws.Range("C7:E7"), "@"

    mod_InstallCommun.PoserEtiquette ws.Range("B8"), mod_Display.FR("Unit{e2} de mesure")
    mod_InstallCommun.PoserChampSaisie ws.Range(mod_Consommations.CF_ADR_UNITE), "@"

    mod_InstallCommun.PoserEtiquette ws.Range("B9"), mod_Display.FR("D{e2}but du contrat")
    mod_InstallCommun.PoserChampDate ws.Range(mod_Consommations.CF_ADR_DEBUT)

    mod_InstallCommun.PoserEtiquette ws.Range("B10"), "Fin du contrat (facultative)"
    mod_InstallCommun.PoserChampDate ws.Range(mod_Consommations.CF_ADR_FIN)

    mod_InstallCommun.PoserEtiquette ws.Range("B11"), "Abonnement mensuel (" & ChrW(8364) & ")"
    mod_InstallCommun.PoserChampMontant ws.Range(mod_Consommations.CF_ADR_ABO), "#,##0.00", mod_InstallCommun.MONTANT_POSITIF_OU_NUL

    mod_InstallCommun.PoserEtiquette ws.Range("B12"), mod_Display.FR("R{e2}f{e2}rence / N{o2} de compteur")
    mod_InstallCommun.PoserChampSaisie ws.Range("C12:E12"), "@"

    ws.Rows(13).RowHeight = mod_InstallCommun.FRM_H_SEP

    mod_InstallCommun.PoserMessage ws.Range("B14:E14")
    ws.Range(mod_Consommations.CF_ADR_MESSAGE).Font.Color = RGB(192, 80, 0)

    mod_InstallCommun.PoserListeDeroulante ws.Range(mod_Consommations.CF_ADR_TYPE), "ListeTypesConso"

    gauche = ws.Cells(1, 2).Left
    mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, mod_InstallCommun.CapValider(), _
        "CtrValider", "btnCtrValider", mod_InstallCommun.FRM_BTN_L
    mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, mod_InstallCommun.CapAnnuler(), _
        "CtrAnnuler", "btnCtrAnnuler", mod_InstallCommun.FRM_BTN_L

    mod_InstallCommun.TerminerFeuille ws, wsPrecedente
    Application.ScreenUpdating = True
    ConstruireFormulaireContrat = True

End Function

Private Function TexteInstructionsContrat() As String

    Dim t As String

    t = "Ce formulaire d{e2}clare un contrat : un type de consommation chez un op{e2}rateur, {a2} partir d'une date. "
    t = t & "Les relev{e2}s retrouvent leur op{e2}rateur et leur unit{e2} gr{a2}ce {a2} leurs dates."
    t = t & Chr(10) & "CHANGEMENT D'OP{E2}RATEUR :cr{e2}ez simplement le nouveau contrat avec sa [c:D{e2}but du contrat]. S'il y a un contrat en cours "
    t = t & "pour ce type, le programme propose de le cl{o2}turer la veille du nouveau. Les relev{e2}s d{e2}j{a2} saisis gardent leur ancien op{e2}rateur."
    t = t & Chr(10) & "[c:Unit{e2} de mesure] : propos{e2}e d'apr{e1}s le type (m3 pour l'eau, kWh pour l'{e2}lectricit{e2} et le gaz), modifiable. "
    t = t & "[c:Fin du contrat (facultative)] : laissez vide pour un contrat en cours. "
    t = t & "[c:Abonnement mensuel] : valeur de d{e2}part propos{e2}e dans les relev{e2}s ; elle peut {e2}voluer ensuite, relev{e2} par relev{e2}."
    t = t & Chr(10) & "Si un champ est incorrect, un message appara{i2}t sous le formulaire et rien n'est enregistr{e2}."
    t = t & Chr(10) & "Boutons :"
    t = t & Chr(10) & "- [b:Valider] : enregistre le contrat puis referme le formulaire."
    t = t & Chr(10) & "- [b:Annuler] : referme le formulaire sans rien enregistrer."

    TexteInstructionsContrat = t

End Function


' =====================================================================================
' ÉCRAN PRINCIPAL frm_Consommations
' =====================================================================================
Private Function ConstruireEcranConso(ByVal wsPrecedente As Worksheet) As Boolean

    Dim ws As Worksheet
    Dim co As ChartObject
    Dim r As Range
    Dim l As Long, ent As Long

    Set ws = mod_InstallCommun.PreparerFeuille(mod_VarGlobales.NOM_FEUILLE_CONSO, True, True)
    If ws Is Nothing Then Exit Function

    Application.ScreenUpdating = False

    ' A marge ; B:G tableau mensuel ; H marge ; I:Q graphique.
    mod_InstallCommun.MettreEnForme ws, Array(2, 24, 12, 12, 12, 12, 26, 2, 10, 10, 10, 10, 10, 10, 10, 10, 10), 100

    mod_InstallCommun.PoserHauteursEntete ws, 1, False
    mod_InstallCommun.EcrireInstructions ws, LIGNE_INSTRUCTIONS, 2, 2, 3, 17, TexteInstructionsConso()

    ConstruireListesTechniques ws

    ' --- Paramètres ---
    mod_InstallCommun.PoserTitreBloc ws.Range("B6:G6"), mod_Display.FR("Param{e1}tres")
    mod_InstallCommun.PoserEtiquette ws.Range("B7"), "Type"
    mod_InstallCommun.PoserEtiquette ws.Range("B8"), "Indicateur"
    mod_InstallCommun.PoserEtiquette ws.Range("B9"), mod_Display.FR("P{e2}riode")
    mod_InstallCommun.PoserEtiquette ws.Range("B10"), mod_Display.FR("Ann{e2}e")
    mod_InstallCommun.PoserEtiquette ws.Range("B11"), mod_Display.FR("Mois (fin de p{e2}riode)")
    mod_InstallCommun.PoserEtiquette ws.Range("B12"), "Forme du graphique"
    mod_InstallCommun.PoserEtiquette ws.Range("B13"), mod_Display.FR("R{e2}partition des relev{e2}s")
    For l = 7 To 13
        mod_InstallCommun.PoserChampSaisie ws.Range(ws.Cells(l, 3), ws.Cells(l, 5))
    Next l

    ws.Range(mod_Consommations.CO_ADR_TYPE).Value = mod_Consommations.LibTous()
    ws.Range(mod_Consommations.CO_ADR_INDIC).Value = mod_Consommations.LibIndMontant()
    ws.Range(mod_Consommations.CO_ADR_PERIODE).Value = mod_Consommations.LibPeriode12()
    ws.Range(mod_Consommations.CO_ADR_ANNEE).Value = Year(Date)
    ws.Range(mod_Consommations.CO_ADR_MOIS).Value = Month(Date)
    ws.Range(mod_Consommations.CO_ADR_FORME).Value = mod_Consommations.LibFormeBarres()
    ws.Range(mod_Consommations.CO_ADR_REPART).Value = mod_Consommations.LibRepartProrata()

    mod_InstallCommun.PoserListeDeroulante ws.Range(mod_Consommations.CO_ADR_TYPE), "ListeTypeConsoGraph"
    mod_InstallCommun.PoserListeDeroulante ws.Range(mod_Consommations.CO_ADR_INDIC), "ListeIndicConso"
    mod_InstallCommun.PoserListeDeroulante ws.Range(mod_Consommations.CO_ADR_PERIODE), "ListePeriodeConso"
    mod_InstallCommun.PoserListeDeroulante ws.Range(mod_Consommations.CO_ADR_ANNEE), "ListeAnneesConso"
    mod_InstallCommun.PoserListeDeroulante ws.Range(mod_Consommations.CO_ADR_MOIS), "ListeMoisConso"
    mod_InstallCommun.PoserListeDeroulante ws.Range(mod_Consommations.CO_ADR_FORME), "ListeFormeConso"
    mod_InstallCommun.PoserListeDeroulante ws.Range(mod_Consommations.CO_ADR_REPART), "ListeRepartConso"

    ' Ligne de résumé (écrite par le programme)
    mod_InstallCommun.PoserChampLecture ws.Range("B15:Q15")
    ws.Range("B15").Font.Bold = True
    ws.Range("B15").Font.Color = mod_InstallCommun.CoulEtiquette()

    ' --- Tableau mensuel ---
    ent = mod_Consommations.CO_LIGNE_ENTETE
    ws.Cells(ent, 2).Value = "Mois"
    ws.Cells(ent, 3).Value = mod_Display.FR("Quantit{e2}")
    ws.Cells(ent, 4).Value = "Prix unitaire"
    ws.Cells(ent, 5).Value = "Abonnement (" & ChrW(8364) & ")"
    ws.Cells(ent, 6).Value = "Montant (" & ChrW(8364) & ")"
    ws.Cells(ent, 7).Value = mod_Display.FR("Op{e2}rateur(s)")
    mod_InstallCommun.PoserEnteteTableau ws.Range(ws.Cells(ent, 2), ws.Cells(ent, 7))
    ws.Range(ws.Cells(ent, 3), ws.Cells(ent, 6)).HorizontalAlignment = xlRight

    mod_InstallCommun.PoserCorpsTableauLecture ws.Range(ws.Cells(ent + 1, 2), ws.Cells(ent + mod_Consommations.CO_NB_LIGNES_MAX + 1, 7))
    ws.Range(ws.Cells(ent + 1, 2), ws.Cells(ent + mod_Consommations.CO_NB_LIGNES_MAX + 1, 2)).NumberFormat = "mmm yyyy"
    ws.Range(ws.Cells(ent + 1, 2), ws.Cells(ent + mod_Consommations.CO_NB_LIGNES_MAX + 1, 2)).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(ent + 1, 3), ws.Cells(ent + mod_Consommations.CO_NB_LIGNES_MAX + 1, 3)).NumberFormat = "#,##0.00"
    ws.Range(ws.Cells(ent + 1, 4), ws.Cells(ent + mod_Consommations.CO_NB_LIGNES_MAX + 1, 4)).NumberFormat = "#,##0.0000"
    ws.Range(ws.Cells(ent + 1, 5), ws.Cells(ent + mod_Consommations.CO_NB_LIGNES_MAX + 1, 6)).NumberFormat = "#,##0.00"

    ' --- Graphique (ses séries sont posées par mod_Consommations à chaque actualisation) ---
    Set r = ws.Range(ws.Cells(ent, 9), ws.Cells(ent + 28, 17))
    Set co = ws.ChartObjects.Add(r.Left, r.Top, r.Width, r.Height)
    co.Name = mod_Consommations.CO_GRAPH
    co.Chart.ChartType = xlColumnStacked
    co.Chart.PlotVisibleOnly = False      ' les données du graphique sont dans des colonnes masquées
    co.Chart.HasTitle = True
    co.Chart.ChartTitle.Text = "Consommations"

    ' Colonnes techniques (données du graphique, listes déroulantes) : masquées.
    ws.Range(ws.Columns(mod_Consommations.CO_COL_H_LIB), ws.Columns(mod_Consommations.CO_COL_LISTE_ANNEE)).EntireColumn.Hidden = True

    ' --- Boutons (en dernier : ils se placent d'après les hauteurs de lignes) ---
    ConstruireBoutonsConso ws

    mod_InstallCommun.TerminerFeuille ws, wsPrecedente
    Application.ScreenUpdating = True
    ConstruireEcranConso = True

End Function

Private Sub ConstruireBoutonsConso(ByVal ws As Worksheet)

    Dim gauche As Double

    gauche = ws.Cells(1, 2).Left

    mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, "Actualiser", _
        "ActualiserConsommations", "btnActualiserConso", 80
    mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, mod_Display.FR("Saisir un relev{e2}"), _
        "SaisirReleve", "btnSaisirReleveConso", 108
    mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, mod_Display.FR("Voir les relev{e2}s"), _
        "AfficherReleves", "btnVoirRelevesConso", 108
    mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, "Contrats", _
        "AfficherContrats", "btnVoirContratsConso", 80
    mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, mod_InstallCommun.CapSortir(), _
        "SortirConsommations", "btnSortirConso", mod_InstallCommun.FRM_BTN_L

End Sub

Private Function TexteInstructionsConso() As String

    Dim t As String

    t = "Cet {e2}cran suit l'{e2}volution des consommations d'eau, d'{e2}lectricit{e2} et de gaz, mois apr{e1}s mois. "
    t = t & "[b:Actualiser] recalcule le tableau et le graphique apr{e1}s un changement de param{e1}tre."
    t = t & Chr(10) & "[c:Type] : un type de consommation, ou (tous) pour comparer les montants. [c:Indicateur] : Quantit{e2}, Prix unitaire "
    t = t & "(moyenne pond{e2}r{e2}e par les quantit{e2}s, abonnement exclu), Abonnement ou Montant total (consommation + abonnement). "
    t = t & "Pour (tous), seuls l'abonnement et le montant sont comparables."
    t = t & Chr(10) & "[c:P{e2}riode] : les 12, 36, 60 ou 120 derniers mois (jusqu'{a2} 10 ans) se terminant au [c:Mois (fin de p{e2}riode)] de l'[c:Ann{e2}e] choisie, ou l'ann{e2}e civile enti{e1}re."
    t = t & Chr(10) & "[c:R{e2}partition des relev{e2}s] : 'Au prorata des jours' r{e2}partit chaque relev{e2} sur les mois qu'il couvre (quantit{e2} et "
    t = t & "consommation au prorata des jours, abonnement au prorata de chaque mois) ; 'Mois de fin du relev{e2}' l'impute enti{e1}rement au mois de sa date de fin."
    t = t & Chr(10) & "[c:Forme du graphique] : Barres (empil{e2}es par op{e2}rateur, ou par type pour (tous)), Courbes, ou Barres + courbe du prix "
    t = t & "(pour un type choisi : les barres montrent l'indicateur et le prix unitaire se lit sur l'axe de droite)."
    t = t & Chr(10) & "Boutons :"
    t = t & Chr(10) & "- [b:Saisir un relev{e2}] : ouvre le formulaire de saisie d'une p{e2}riode de consommation."
    t = t & Chr(10) & "- [b:Voir les relev{e2}s] : affiche la liste des relev{e2}s (correction directe possible)."
    t = t & Chr(10) & "- [b:Contrats] : affiche la liste des contrats (op{e2}rateurs) ; c'est l{a2} qu'on d{e2}clare un changement d'op{e2}rateur."
    t = t & Chr(10) & "- [b:Sortir] : revient sur [f:Synthese]."

    TexteInstructionsConso = t

End Function


' =====================================================================================
' Listes techniques (colonnes masquées) et noms
' =====================================================================================
Private Sub ConstruireListesTechniques(ByVal ws As Worksheet)

    Dim i As Long
    Dim tps As Variant

    ws.Cells(2, mod_Consommations.CO_COL_LISTE_TYPE).Value = mod_Consommations.LibTous()
    tps = mod_Consommations.TypesConso()
    For i = 0 To UBound(tps)
        ws.Cells(3 + i, mod_Consommations.CO_COL_LISTE_TYPE).Value = tps(i)
    Next i

    ws.Cells(2, mod_Consommations.CO_COL_LISTE_INDIC).Value = mod_Consommations.LibIndQuantite()
    ws.Cells(3, mod_Consommations.CO_COL_LISTE_INDIC).Value = mod_Consommations.LibIndPrix()
    ws.Cells(4, mod_Consommations.CO_COL_LISTE_INDIC).Value = mod_Consommations.LibIndAbo()
    ws.Cells(5, mod_Consommations.CO_COL_LISTE_INDIC).Value = mod_Consommations.LibIndMontant()

    ws.Cells(2, mod_Consommations.CO_COL_LISTE_PERIODE).Value = mod_Consommations.LibPeriode12()
    ws.Cells(3, mod_Consommations.CO_COL_LISTE_PERIODE).Value = mod_Consommations.LibPeriode36()
    ws.Cells(4, mod_Consommations.CO_COL_LISTE_PERIODE).Value = mod_Consommations.LibPeriode60()
    ws.Cells(5, mod_Consommations.CO_COL_LISTE_PERIODE).Value = mod_Consommations.LibPeriode120()
    ws.Cells(6, mod_Consommations.CO_COL_LISTE_PERIODE).Value = mod_Consommations.LibPeriodeAnnee()

    ws.Cells(2, mod_Consommations.CO_COL_LISTE_FORME).Value = mod_Consommations.LibFormeBarres()
    ws.Cells(3, mod_Consommations.CO_COL_LISTE_FORME).Value = mod_Consommations.LibFormeCourbes()
    ws.Cells(4, mod_Consommations.CO_COL_LISTE_FORME).Value = mod_Consommations.LibFormeMixte()

    ws.Cells(2, mod_Consommations.CO_COL_LISTE_REPART).Value = mod_Consommations.LibRepartProrata()
    ws.Cells(3, mod_Consommations.CO_COL_LISTE_REPART).Value = mod_Consommations.LibRepartFin()

    For i = 1 To 12
        ws.Cells(1 + i, mod_Consommations.CO_COL_LISTE_MOIS).Value = i
    Next i

    ws.Cells(2, mod_Consommations.CO_COL_LISTE_ANNEE).Value = Year(Date)

    PoserNomListe "ListeTypeConsoGraph", ws, mod_Consommations.CO_COL_LISTE_TYPE, 2, 2 + UBound(tps) + 1
    PoserNomListe "ListeIndicConso", ws, mod_Consommations.CO_COL_LISTE_INDIC, 2, 5
    PoserNomListe "ListePeriodeConso", ws, mod_Consommations.CO_COL_LISTE_PERIODE, 2, 6
    PoserNomListe "ListeFormeConso", ws, mod_Consommations.CO_COL_LISTE_FORME, 2, 4
    PoserNomListe "ListeRepartConso", ws, mod_Consommations.CO_COL_LISTE_REPART, 2, 3
    PoserNomListe "ListeMoisConso", ws, mod_Consommations.CO_COL_LISTE_MOIS, 2, 13

    ' Liste dont la longueur varie (remplie par mod_Consommations à chaque actualisation).
    PoserNomDynamique "ListeAnneesConso", ws, mod_Consommations.CO_COL_LISTE_ANNEE, 60

End Sub

Private Sub PoserNomListe(ByVal nom As String, ByVal ws As Worksheet, ByVal col As Long, ByVal ligneDeb As Long, ByVal ligneFin As Long)
    On Error Resume Next
    ThisWorkbook.Names(nom).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nom, RefersTo:="='" & ws.Name & "'!" & _
        ws.Range(ws.Cells(ligneDeb, col), ws.Cells(ligneFin, col)).Address(True, True)
End Sub

' Plage qui s'adapte au nombre de valeurs écrites à partir de la ligne 2.
Private Sub PoserNomDynamique(ByVal nom As String, ByVal ws As Worksheet, ByVal col As Long, ByVal ligneMax As Long)
    Dim debut As String, plage As String
    debut = ws.Cells(2, col).Address(True, True)
    plage = ws.Range(ws.Cells(2, col), ws.Cells(ligneMax, col)).Address(True, True)
    On Error Resume Next
    ThisWorkbook.Names(nom).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nom, RefersTo:="=OFFSET('" & ws.Name & "'!" & debut & ",0,0,MAX(1,COUNTA('" & ws.Name & "'!" & plage & ")),1)"
End Sub


' =====================================================================================
' Bouton d'accès sur la feuille Synthese (déplaçable à la main ensuite)
' =====================================================================================
Private Sub ConstruireBoutonSynthese()

    Dim ws As Worksheet
    Dim btn As Button

    Set ws = mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_SYNTHESE)
    If ws Is Nothing Then
        MsgBox mod_Display.FR("La feuille Synthese est introuvable : le bouton d'acc{e1}s n'a pas {e2}t{e2} cr{e2}{e2}."), vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    ws.Shapes("btnConsommations").Delete
    On Error GoTo 0

    Set btn = ws.Buttons.Add(ws.Cells(SYN_LIGNE_BOUTON, 1).Left + 12, ws.Cells(SYN_LIGNE_BOUTON, 1).Top, 152, 28)
    btn.Caption = "Consommations"
    btn.OnAction = "AfficherConsommations"
    btn.Name = "btnConsommations"

End Sub


' =====================================================================================
' OUTILS DÉVELOPPEUR (Ctrl+G)
' =====================================================================================
Public Sub AfficherFeuilleConsommationsPourEdition()
    mod_InstallCommun.AfficherPourEdition mod_VarGlobales.NOM_FEUILLE_CONSO, "InstallerConsommations", "MasquerFeuilleConsommationsApresEdition"
End Sub

Public Sub MasquerFeuilleConsommationsApresEdition()
    mod_InstallCommun.MasquerApresEdition mod_VarGlobales.NOM_FEUILLE_CONSO
End Sub

Public Sub AfficherFormulaireRelevePourEdition()
    mod_InstallCommun.AfficherPourEdition mod_VarGlobales.NOM_FEUILLE_RELEVE, "InstallerConsommations", "MasquerFormulaireReleveApresEdition"
End Sub

Public Sub MasquerFormulaireReleveApresEdition()
    mod_InstallCommun.MasquerApresEdition mod_VarGlobales.NOM_FEUILLE_RELEVE
End Sub

Public Sub AfficherFormulaireContratPourEdition()
    mod_InstallCommun.AfficherPourEdition mod_VarGlobales.NOM_FEUILLE_CONTRAT, "InstallerConsommations", "MasquerFormulaireContratApresEdition"
End Sub

Public Sub MasquerFormulaireContratApresEdition()
    mod_InstallCommun.MasquerApresEdition mod_VarGlobales.NOM_FEUILLE_CONTRAT
End Sub
