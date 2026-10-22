Attribute VB_Name = "mod_InstallConsommations"
Option Explicit

' =====================================================================================
' MODULE : mod_InstallConsommations
'
' SUIVI DES CONSOMMATIONS (eau, électricité, gaz) - construction des feuilles.
' La logique est dans mod_Consommations (contrats, graphique) et mod_Factures (factures).
'
' Une seule macro à lancer : InstallerConsommations. Elle crée, dans l'ordre :
'     0. les feuilles de DONNÉES masquées (voir plus bas), sans jamais les effacer ;
'     1. frm_Facture     : le formulaire de saisie d'une facture (en-tête + lignes d'abonnement,
'                          de consommation et de taxes).
'     2. frm_Contrat     : le formulaire de saisie d'un contrat (opérateur).
'     3. frm_Contrats    : la liste des contrats.
'     4. frm_Factures    : la liste des factures (lecture seule, avec totaux et écart).
'     5. frm_Consommations : l'écran principal (paramètres, tableau mensuel, graphique).
'     6. le bouton "Consommations" sur la feuille Synthese.
'   Les 5 feuilles frm_* sont masquées : on les ouvre avec les boutons.
'
' INSTALLATION :
'   1. Alt+F11, Fichier > Importer : mod_VarGlobales.bas, mod_Consommations.bas,
'      mod_Factures.bas, mod_InstallConsommations.bas, mod_InstallAide.bas et mod_Synchro.bas.
'   2. Ctrl+G, taper : InstallerConsommations, puis Entrée.
'   3. Coller dans le module de la feuille frm_Facture le code de CodeBehind_frm_Facture.txt,
'      et dans celui de la feuille frm_Contrat le code de CodeBehind_frm_Contrat.txt (le
'      message de fin d'installation donne les CodeName).
'   4. Sur Synthese, cliquer sur le bouton "Consommations", puis sur Contrats pour
'      déclarer les contrats en cours, puis sur Saisir une facture.
'
' DONNÉES : TblFactures, TblAbonnements, TblConsos, TblTaxes et TblContrats sont sur des
' feuilles masquées (Conso_Factures, Conso_Abonnements, Conso_Consommations, Conso_Taxes,
' Conso_Contrats), comme TblVentilations sur Ventilations. Elles ne sont JAMAIS reconstruites :
' réinstaller ne touche pas aux données. Les écrans frm_Contrats et frm_Factures n'en montrent
' qu'une copie. Anciennes versions : les contrats saisis sur frm_Contrats sont déplacés
' automatiquement ; l'ancien modèle "relevé" (une période unique par facture) est abandonné :
' ses feuilles (frm_Releve, frm_Releves, Conso_Releves) sont supprimées après confirmation.
'
' RÉINSTALLATION : possible sans perte. L'écran principal revient à ses paramètres par défaut.
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

    If Not PreparerToutesLesDonnees() Then Exit Sub
    If Not SupprimerAnciensReleves() Then Exit Sub
    If Not ConstruireFormulaireFacture(wsPrecedente) Then Exit Sub
    If Not ConstruireFormulaireContrat(wsPrecedente) Then Exit Sub
    If Not ConstruireListeContrats(wsPrecedente) Then Exit Sub
    If Not ConstruireListeFactures(wsPrecedente) Then Exit Sub
    If Not ConstruireEcranConso(wsPrecedente) Then Exit Sub
    ConstruireBoutonSynthese

    MsgBox mod_Display.FR("Le suivi des consommations est install{e2}.") & vbCrLf & vbCrLf & _
           mod_Display.FR("PROCHAINE {E2}TAPE (indispensable) : coller le code des {e2}v{e2}nements dans les deux feuilles de formulaire.") & vbCrLf & _
           "Nom interne (CodeName) de la feuille " & mod_VarGlobales.NOM_FEUILLE_FACTURE & " : " & _
           mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_FACTURE).CodeName & vbCrLf & _
           "Nom interne (CodeName) de la feuille " & mod_VarGlobales.NOM_FEUILLE_CONTRAT & " : " & _
           mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_CONTRAT).CodeName & vbCrLf & vbCrLf & _
           mod_Display.FR("Voir les fichiers CodeBehind_frm_Facture.txt et CodeBehind_frm_Contrat.txt."), _
           vbInformation, mod_Display.FR("Installation termin{e2}e")

End Sub


' =====================================================================================
' FEUILLES DE DONNÉES (Conso_*) : créées une seule fois, jamais effacées
' =====================================================================================
Private Function PreparerToutesLesDonnees() As Boolean

    Dim lignes As Variant
    Dim e As String

    e = "ID_Facture|Libelle|DateDebut|DateFin|Quantite|Unite|PrixUnitaireHT|TauxTVA|MontantHT|MontantTTC"
    lignes = Split(e, "|")

    If Not PreparerDonnees(mod_VarGlobales.NOM_FEUILLE_DATA_CONTRATS, mod_VarGlobales.NOM_TABLE_CONTRATS, _
        Split("ID_Contrat|TypeConso|Operateur|Unite|DateDebut|DateFin|AbonnementMensuel|Reference", "|"), _
        mod_VarGlobales.NOM_FEUILLE_CONTRATS) Then Exit Function
    If Not PreparerDonnees(mod_VarGlobales.NOM_FEUILLE_DATA_FACTURES, mod_VarGlobales.NOM_TABLE_FACTURES, _
        Split("ID_Facture|TypeConso|Operateur|NumeroFacture|DateFacture|DebutPeriode|FinPeriode|TotalFactureTTC|Commentaire", "|"), "") Then Exit Function
    If Not PreparerDonnees(mod_VarGlobales.NOM_FEUILLE_DATA_ABOS, mod_VarGlobales.NOM_TABLE_ABOS, lignes, "") Then Exit Function
    If Not PreparerDonnees(mod_VarGlobales.NOM_FEUILLE_DATA_CONSOS, mod_VarGlobales.NOM_TABLE_CONSOS, lignes, "") Then Exit Function
    If Not PreparerDonnees(mod_VarGlobales.NOM_FEUILLE_DATA_TAXES, mod_VarGlobales.NOM_TABLE_TAXES, lignes, "") Then Exit Function
    PreparerToutesLesDonnees = True

End Function

' Crée la feuille de données et son tableau s'ils n'existent pas. S'il existe déjà, on n'y touche pas.
' "nomAncienne" : ancienne feuille d'écran qui portait ce tableau (ses données sont alors déplacées).
Private Function PreparerDonnees(ByVal nomData As String, ByVal nomTable As String, ByVal entetes As Variant, _
                                 ByVal nomAncienne As String) As Boolean

    Dim ws As Worksheet, wsOld As Worksheet
    Dim lo As ListObject, loOld As ListObject
    Dim nb As Long, i As Long, n As Long
    Dim noms() As String, vals() As Variant

    ' Déjà là : on n'y touche pas.
    Set ws = mod_InstallCommun.TrouverFeuille(nomData)
    If Not ws Is Nothing Then
        On Error Resume Next
        Set lo = ws.ListObjects(nomTable)
        On Error GoTo 0
        If Not lo Is Nothing Then
            PreparerDonnees = True
            Exit Function
        End If
    End If

    Application.ScreenUpdating = False
    nb = UBound(entetes) - LBound(entetes) + 1

    ' Ancienne version : le tableau (et ses données) était sur l'écran frm_*. On le recopie
    ' puis on le supprime pour libérer son nom.
    If nomAncienne <> "" Then
        Set wsOld = mod_InstallCommun.TrouverFeuille(nomAncienne)
        If Not wsOld Is Nothing Then
            On Error Resume Next
            Set loOld = wsOld.ListObjects(nomTable)
            On Error GoTo 0
        End If
    End If
    If Not loOld Is Nothing Then
        If Not loOld.DataBodyRange Is Nothing Then
            n = loOld.ListRows.Count
            ReDim noms(1 To loOld.ListColumns.Count)
            ReDim vals(1 To loOld.ListColumns.Count)
            For i = 1 To loOld.ListColumns.Count
                noms(i) = loOld.ListColumns(i).Name
                If Not loOld.ListColumns(i).DataBodyRange.Cells(1, 1).HasFormula Then
                    vals(i) = loOld.ListColumns(i).DataBodyRange.Value2
                End If
            Next i
        End If
        loOld.Delete
    End If

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nomData
    End If
    ws.Cells.Clear
    For i = 0 To nb - 1
        ws.Cells(1, i + 1).Value = entetes(LBound(entetes) + i)
    Next i
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(2, nb)), , xlYes)
    lo.Name = nomTable
    If n > 1 Then lo.Resize ws.Range(ws.Cells(1, 1), ws.Cells(n + 1, nb))

    If n > 0 Then
        For i = 1 To UBound(noms)
            If Not IsEmpty(vals(i)) Then
                On Error Resume Next
                lo.ListColumns(noms(i)).DataBodyRange.Value2 = vals(i)
                On Error GoTo 0
            End If
        Next i
    End If
    For i = 1 To nb
        lo.ListColumns(i).DataBodyRange.NumberFormat = mod_Consommations.FormatColonneConso(CStr(entetes(LBound(entetes) + i - 1)))
    Next i
    ws.Columns.AutoFit

    ws.Visible = xlSheetHidden
    Application.ScreenUpdating = True
    PreparerDonnees = True

End Function

' Ancien modèle "relevé" : ses feuilles sont supprimées. Si Conso_Releves contient des données,
' une confirmation est demandée (elles seraient perdues).
Private Function SupprimerAnciensReleves() As Boolean

    Dim noms As Variant
    Dim ws As Worksheet, wsRel As Worksheet
    Dim lo As ListObject
    Dim i As Long
    Dim aDesDonnees As Boolean

    Set wsRel = mod_InstallCommun.TrouverFeuille("Conso_Releves")
    If Not wsRel Is Nothing Then
        On Error Resume Next
        Set lo = wsRel.ListObjects("TblReleves")
        On Error GoTo 0
        If Not lo Is Nothing Then
            If Not lo.DataBodyRange Is Nothing Then
                aDesDonnees = (Application.WorksheetFunction.CountA(lo.DataBodyRange) > 0)
            End If
        End If
    End If
    If aDesDonnees Then
        If MsgBox(mod_Display.FR("L'ancienne feuille 'Conso_Releves' contient des relev{e2}s. Le nouveau mod{e1}le (factures avec lignes d'abonnement, de consommation et de taxes) ne les reprend pas.") & vbCrLf & _
                  mod_Display.FR("Les supprimer maintenant ?"), vbYesNo + vbExclamation, mod_Display.FR("Anciens relev{e2}s")) = vbNo Then
            MsgBox mod_Display.FR("Installation annul{e2}e : rien n'a {e2}t{e2} modifi{e2}."), vbInformation
            Exit Function
        End If
    End If

    noms = Array("frm_Releve", "frm_Releve_aide", "frm_Releves", "frm_Releves_aide", "Conso_Releves")
    ' Événements coupés : l'ancienne feuille frm_Releve contient du code (Worksheet_Deactivate) qui
    ' appelait des procédures désormais supprimées ; il ne doit pas se déclencher pendant la suppression.
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    For i = LBound(noms) To UBound(noms)
        Set ws = mod_InstallCommun.TrouverFeuille(CStr(noms(i)))
        If Not ws Is Nothing Then
            On Error Resume Next
            ws.Visible = xlSheetVisible
            ws.Delete
            On Error GoTo 0
        End If
    Next i
    Application.DisplayAlerts = True
    Application.EnableEvents = True

    ' Contrôle : une feuille qui résiste à la suppression est signalée (plus d'échec silencieux).
    Dim restantes As String
    For i = LBound(noms) To UBound(noms)
        If Not mod_InstallCommun.TrouverFeuille(CStr(noms(i))) Is Nothing Then restantes = restantes & vbCrLf & " - " & noms(i)
    Next i
    If restantes <> "" Then
        MsgBox mod_Display.FR("Ces anciennes feuilles n'ont pas pu {ea}tre supprim{e2}es automatiquement (supprimez-les {a2} la main, clic droit sur l'onglet) :") & restantes, _
               vbExclamation, mod_Display.FR("Anciens relev{e2}s")
    End If
    SupprimerAnciensReleves = True

End Function


' Outil développeur (Ctrl+G) : supprime seulement les anciennes feuilles "relevé", sans reconstruire le reste.
Public Sub NettoyerAnciensReleves()
    If SupprimerAnciensReleves() Then MsgBox mod_Display.FR("Nettoyage termin{e2}."), vbInformation
End Sub


' =====================================================================================
' LISTES (contrats, factures)
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
        mod_Display.FR("Voir les factures"), "AfficherFactures", "btnVoirFacturesC", 108, _
        "Retour", "RetourConsommations", "btnRetourContrats", mod_InstallCommun.FRM_BTN_L)

    t = "Cette liste contient les contrats de fourniture : un type de consommation chez un op{e2}rateur, entre deux dates. "
    t = t & "Le contrat sert {a2} proposer l'op{e2}rateur et l'unit{e2} (m3, kWh...) quand on saisit une facture."
    t = t & Chr(10) & "Pour changer d'op{e2}rateur, ajoutez simplement le nouveau contrat avec [b:Ajouter un contrat] : s'il y a un contrat en cours "
    t = t & "pour ce type, le programme propose de le cl{o2}turer la veille. L'historique reste complet. "
    t = t & "Vous pouvez corriger directement les cellules ([c:DateFin] vide = contrat en cours) ; deux contrats d'un m{e1}me type ne doivent pas se chevaucher."
    t = t & Chr(10) & "[c:AbonnementMensuel] est donn{e2} {a2} titre indicatif : l'abonnement r{e2}el est port{e2} par les lignes d'abonnement de chaque facture."
    t = t & Chr(10) & "Boutons :"
    t = t & Chr(10) & "- [b:Ajouter un contrat] : ouvre le formulaire de saisie."
    t = t & Chr(10) & "- [b:Supprimer la s{e2}lection] : supprime les lignes s{e2}lectionn{e2}es (les factures d{e2}j{a2} saisies sont conserv{e2}es)."
    t = t & Chr(10) & "- [b:Voir les factures] : affiche la liste des factures."
    t = t & Chr(10) & "- [b:Retour] : revient sur [f:frm_Consommations]."

    ConstruireListeContrats = ConstruireListe(mod_VarGlobales.NOM_FEUILLE_CONTRATS, mod_VarGlobales.NOM_TABLE_CONTRATS_AFF, _
                                              entetes, largeurs, formats, t, boutons, wsPrecedente)

End Function

Private Function ConstruireListeFactures(ByVal wsPrecedente As Worksheet) As Boolean

    Dim entetes As Variant, largeurs As Variant, formats As Variant, boutons As Variant
    Dim t As String

    entetes = Array("ID_Facture", "TypeConso", "Operateur", "NumeroFacture", "DateFacture", "DebutPeriode", "FinPeriode", _
                    "TotalHT", "TotalTVA", "TotalTTCCalcule", "TotalFactureTTC", "Ecart", "Commentaire")
    largeurs = Array(11, 18, 22, 16, 12, 12, 12, 12, 12, 16, 16, 10, 28)
    formats = Array("0", "@", "@", "@", "dd/mm/yyyy", "dd/mm/yyyy", "dd/mm/yyyy", "#,##0.00", "#,##0.00", "#,##0.00", "#,##0.00", "#,##0.00", "@")

    boutons = Array( _
        mod_Display.FR("Ajouter une facture"), "SaisirFacture", "btnAjouterFacture", 118, _
        mod_Display.FR("Modifier la s{e2}lection"), "ModifierFactureSelection", "btnModifierFacture", 118, _
        mod_Display.FR("Copier la s{e2}lection"), "CopierFactureSelection", "btnCopierFacture", 112, _
        mod_Display.FR("Supprimer la s{e2}lection"), "SupprimerFactureSelection", "btnSupprimerFacture", 118, _
        "Contrats", "AfficherContrats", "btnVoirContratsF", 80, _
        "Retour", "RetourConsommations", "btnRetourFactures", mod_InstallCommun.FRM_BTN_L)

    t = "Cette liste r{e2}capitule les factures de consommation. Une facture se compose d'un en-t{ea}te et de lignes d'abonnement, de consommation et de taxes, "
    t = t & "chacune avec ses propres dates, son prix unitaire HT, son taux de TVA et son montant TTC."
    t = t & Chr(10) & "La liste est en lecture seule : on ajoute ou on corrige une facture avec le formulaire. [c:TotalTTCCalcule] est la somme des lignes saisies "
    t = t & "(indicatif) ; [c:TotalFactureTTC] est le montant r{e2}ellement factur{e2} ; [c:Ecart] est leur diff{e2}rence."
    t = t & Chr(10) & "Boutons :"
    t = t & Chr(10) & "- [b:Ajouter une facture] : ouvre le formulaire de saisie."
    t = t & Chr(10) & "- [b:Modifier la s{e2}lection] : rouvre le formulaire sur la facture s{e2}lectionn{e2}e."
    t = t & Chr(10) & "- [b:Copier la s{e2}lection] : ouvre le formulaire avec le type, l'op{e2}rateur, les libell{e2}s, les unit{e2}s et les taux de TVA de la facture s{e2}lectionn{e2}e ; num{e2}ro, dates, quantit{e2}s et prix sont laiss{e2}s vides. Utile pour saisir la facture suivante."
    t = t & Chr(10) & "- [b:Supprimer la s{e2}lection] : supprime la facture et toutes ses lignes."
    t = t & Chr(10) & "- [b:Contrats] : affiche la liste des contrats."
    t = t & Chr(10) & "- [b:Retour] : revient sur [f:frm_Consommations]."

    ConstruireListeFactures = ConstruireListe(mod_VarGlobales.NOM_FEUILLE_FACTURES, mod_VarGlobales.NOM_TABLE_FACTURES_AFF, _
                                              entetes, largeurs, formats, t, boutons, wsPrecedente)

End Function

' Construction commune aux deux listes (l'écran affiche une copie des données, voir mod_Consommations).
'   boutons : suite de quadruplets (légende, macro, nom du bouton, largeur).
Private Function ConstruireListe(ByVal nomFeuille As String, ByVal nomTable As String, ByVal entetes As Variant, _
                                 ByVal largeurs As Variant, ByVal formats As Variant, ByVal texteInstr As String, _
                                 ByVal boutons As Variant, ByVal wsPrecedente As Worksheet) As Boolean

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim w() As Variant
    Dim nb As Long, i As Long
    Dim gauche As Double
    
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

    PoserListeTypeSurColonne lo
    If nomTable = mod_VarGlobales.NOM_TABLE_FACTURES_AFF Then
        mod_Factures.RafraichirListeFactures
    Else
        mod_Consommations.RafraichirListeContrats
    End If

    ' Boutons (en dernier : ils se placent d'après les hauteurs de lignes)
    gauche = ws.Cells(1, 2).Left
    For i = LBound(boutons) To UBound(boutons) Step 4
        mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, CStr(boutons(i)), CStr(boutons(i + 1)), _
                                              CStr(boutons(i + 2)), CDbl(boutons(i + 3))
    Next i
    mod_InstallCommun.AjouterBoutonAide ws, LIGNE_BOUTONS, gauche

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
' FORMULAIRE frm_Facture
' =====================================================================================
Private Function ConstruireFormulaireFacture(ByVal wsPrecedente As Worksheet) As Boolean

    Dim ws As Worksheet
    Dim gauche As Double
    Dim i As Long, tps As Variant

    Set ws = mod_InstallCommun.PreparerFeuille(mod_VarGlobales.NOM_FEUILLE_FACTURE)
    If ws Is Nothing Then Exit Function

    Application.ScreenUpdating = False

    ' A marge ; B libellé ; C début ; D fin ; E quantité ; F unité ; G prix HT ; H TVA ; I HT ; J TTC ; K marge
    mod_InstallCommun.MettreEnForme ws, Array(2, 44, 12, 12, 11, 8, 13, 8, 12, 12, 2), 30

    If mod_InstallCommun.PremiereLigneCorps(1, False) <> LIGNE_CORPS Or _
       mod_InstallCommun.LigneInstructions(1, False) <> LIGNE_INSTRUCTIONS Then
        MsgBox "mod_InstallConsommations : constantes de lignes incoherentes avec mod_InstallCommun.", vbCritical
    End If
    mod_InstallCommun.PoserHauteursEntete ws, 1, False
    mod_InstallCommun.EcrireInstructions ws, LIGNE_INSTRUCTIONS, 2, 2, 3, 10, TexteInstructionsFacture()

    ' --- En-tête de la facture ---
    mod_InstallCommun.PoserEtiquette ws.Range("B6"), "Type de consommation"
    mod_InstallCommun.PoserChampSaisie ws.Range("C6:D6"), "@"

    mod_InstallCommun.PoserEtiquette ws.Range("B7"), mod_Display.FR("Op{e2}rateur (fournisseur)")
    mod_InstallCommun.PoserChampSaisie ws.Range("C7:E7"), "@"

    mod_InstallCommun.PoserEtiquette ws.Range("B8"), mod_Display.FR("Num{e2}ro de facture (facultatif)")
    mod_InstallCommun.PoserChampSaisie ws.Range("C8:D8"), "@"

    mod_InstallCommun.PoserEtiquette ws.Range("B9"), "Date de la facture (facultative)"
    mod_InstallCommun.PoserChampDate ws.Range(mod_Factures.FF_ADR_DATE)

    mod_InstallCommun.PoserEtiquette ws.Range("B10"), mod_Display.FR("D{e2}but de la p{e2}riode (facultatif)")
    mod_InstallCommun.PoserChampDate ws.Range(mod_Factures.FF_ADR_DEBUT)

    mod_InstallCommun.PoserEtiquette ws.Range("B11"), "Fin de la p" & ChrW(233) & "riode (facultative)"
    mod_InstallCommun.PoserChampDate ws.Range(mod_Factures.FF_ADR_FIN)

    mod_InstallCommun.PoserEtiquette ws.Range("B12"), mod_Display.FR("Total TTC de la facture (facultatif)")
    mod_InstallCommun.PoserChampMontant ws.Range(mod_Factures.FF_ADR_TOTAL), "#,##0.00", mod_InstallCommun.MONTANT_POSITIF_OU_NUL

    mod_InstallCommun.PoserEtiquette ws.Range("B13"), "Commentaire (facultatif)"
    mod_InstallCommun.PoserChampSaisie ws.Range("C13:J13"), "@"

    ws.Rows(14).RowHeight = mod_InstallCommun.FRM_H_SEP

    ' --- Récapitulatif (indicatif), EN HAUT : les blocs de lignes grandissent vers le bas ---
    mod_InstallCommun.PoserTitreBloc ws.Range(ws.Cells(mod_Factures.FF_RECAP_TITRE, 2), ws.Cells(mod_Factures.FF_RECAP_TITRE, 10)), _
        mod_Display.FR("R{e2}capitulatif (indicatif, d'apr{e1}s les lignes saisies)")
    mod_InstallCommun.PoserEtiquette ws.Range("B16"), "Total HT"
    mod_InstallCommun.PoserChampLecture ws.Range(mod_Factures.FF_ADR_RECAP_HT), "#,##0.00"
    mod_InstallCommun.PoserEtiquette ws.Range("B17"), "TVA"
    mod_InstallCommun.PoserChampLecture ws.Range(mod_Factures.FF_ADR_RECAP_TVA), "#,##0.00"
    mod_InstallCommun.PoserEtiquette ws.Range("B18"), mod_Display.FR("Total TTC calcul{e2}")
    mod_InstallCommun.PoserChampLecture ws.Range(mod_Factures.FF_ADR_RECAP_TTC), "#,##0.00"
    ws.Range(mod_Factures.FF_ADR_RECAP_TTC).Font.Bold = True
    mod_InstallCommun.PoserEtiquette ws.Range("B19"), mod_Display.FR("Total TTC de la facture (saisi)")
    mod_InstallCommun.PoserChampLecture ws.Range(mod_Factures.FF_ADR_RECAP_SAISI), "#,##0.00"
    mod_InstallCommun.PoserEtiquette ws.Range("B20"), mod_Display.FR("{E2}cart (facture - calcul)")
    mod_InstallCommun.PoserChampLecture ws.Range(mod_Factures.FF_ADR_RECAP_ECART), "#,##0.00"
    ws.Rows(21).RowHeight = mod_InstallCommun.FRM_H_SEP

    mod_InstallCommun.PoserMessage ws.Range("B22:J22")
    ws.Range(mod_Factures.FF_ADR_MESSAGE).Font.Color = RGB(192, 80, 0)
    ws.Rows(23).RowHeight = mod_InstallCommun.FRM_H_SEP

    ' --- Liste technique des types (colonne masquée) ---
    ws.Cells(1, mod_Factures.FF_COL_AIDE_TYPE).Value = "Types"
    tps = mod_Consommations.TypesConso()
    For i = 0 To UBound(tps)
        ws.Cells(2 + i, mod_Factures.FF_COL_AIDE_TYPE).Value = tps(i)
    Next i
    PoserNomListe "ListeTypesConso", ws, mod_Factures.FF_COL_AIDE_TYPE, 2, 1 + UBound(tps) + 1
    ws.Columns(mod_Factures.FF_COL_AIDE_TYPE).Hidden = True
    ws.Columns(mod_Factures.FF_COL_AIDE_UNITE).Hidden = True
    ws.Columns(mod_Factures.FF_COL_AIDE_OPER).Hidden = True     ' liste des opérateurs en cours, remplie à l'ouverture
    mod_InstallCommun.PoserListeDeroulante ws.Range(mod_Factures.FF_ADR_TYPE), "ListeTypesConso"

    ' Les trois blocs de lignes (abonnement, consommation, taxes), FF_NB_DEFAUT lignes chacun :
    ' ils sont redessinés à chaque ouverture et grandissent avec leur bouton "+".
    mod_Factures.DessinerBlocs ws, mod_Factures.FF_NB_DEFAUT, mod_Factures.FF_NB_DEFAUT, mod_Factures.FF_NB_DEFAUT

    ' --- Boutons de la ligne d'en-tête ---
    gauche = ws.Cells(1, 2).Left
    mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, mod_InstallCommun.CapValider(), _
        "FacValider", "btnFacValider", mod_InstallCommun.FRM_BTN_L
    mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, mod_InstallCommun.CapAnnuler(), _
        "FacAnnuler", "btnFacAnnuler", mod_InstallCommun.FRM_BTN_L
    mod_InstallCommun.AjouterBoutonAide ws, LIGNE_BOUTONS, gauche

    mod_InstallCommun.TerminerFeuille ws, wsPrecedente
    Application.ScreenUpdating = True
    ConstruireFormulaireFacture = True

End Function

Private Function TexteInstructionsFacture() As String

    Dim t As String

    t = "Ce formulaire enregistre une facture telle qu'elle est imprim{e2}e par le fournisseur : un en-t{ea}te, puis des lignes d'ABONNEMENT, de CONSOMMATION et de TAXES. "
    t = t & "Chaque ligne a ses propres dates, qui ne co{i2}ncident pas forc{e2}ment avec celles des autres lignes (le prix change en cours de facture)."
    t = t & Chr(10) & "[c:Type de consommation] et [c:Op{e2}rateur] : l'op{e2}rateur du contrat en vigueur est propos{e2} quand vous choisissez le type. "
    t = t & "Les dates et le total de la facture sont facultatifs mais utiles : le total saisi est compar{e2} au total calcul{e2} d'apr{e1}s les lignes."
    t = t & Chr(10) & "[c:Boutons + et -] (en haut {a2} droite de chaque bloc) : + ajoute une ligne au bloc et d{e2}cale les blocs suivants vers le bas ; - retire la derni{e1}re ligne si elle est vide. Chaque bloc commence avec 3 lignes. Le taux de [c:TVA] est pr{e2}rempli (celui de la derni{e1}re facture du m{ea}me type, sinon un taux usuel) et reste modifiable ; [c:Unit{e2}] se choisit dans une liste, et une unit{e2} nouvelle peut {ea}tre saisie apr{e1}s confirmation."
    t = t & Chr(10) & "[c:Lignes] : une ligne par p{e2}riode de prix. [c:D{e2}but] et [c:Fin] : dates libres, fin incluse. Abonnement : les jours se calculent tout seuls. "
    t = t & "Consommation et taxes : la quantit{e2} est en kWh, m3 ou en euros (base d'une taxe), et l'unit{e2} se saisit librement. "
    t = t & "[c:Prix HT] : prix unitaire hors taxes (jusqu'{a2} 6 d{e2}cimales). [c:TVA] : taux de la ligne (5,5 % ou 20 %, par exemple)."
    t = t & Chr(10) & "[c:Montant HT] = quantit{e2} x prix, arrondi au centime ; [c:Montant TTC] = HT x (1 + TVA), arrondi au centime. Ils se calculent en direct. "
    t = t & "Les lignes vides sont ignor{e2}es. Le r{e2}capitulatif du bas est indicatif : un {e2}cart avec la facture n'emp{ea}che jamais l'enregistrement."
    t = t & Chr(10) & "Si un champ est incorrect, un message appara{i2}t sous le formulaire et rien n'est enregistr{e2}. "
    t = t & "Si une p{e2}riode de consommation ou d'abonnement en chevauche une autre du m{ea}me type, une confirmation est demand{e2}e."
    t = t & Chr(10) & "Boutons :"
    t = t & Chr(10) & "- [b:Valider] : enregistre la facture puis referme le formulaire."
    t = t & Chr(10) & "- [b:Annuler] : referme le formulaire sans rien enregistrer."

    TexteInstructionsFacture = t

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

    mod_InstallCommun.PoserEtiquette ws.Range("B12"), mod_Display.FR("R{e2}f{e2}rence / num{e2}ro de compteur")
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
    mod_InstallCommun.AjouterBoutonAide ws, LIGNE_BOUTONS, gauche

    mod_InstallCommun.TerminerFeuille ws, wsPrecedente
    Application.ScreenUpdating = True
    ConstruireFormulaireContrat = True

End Function

Private Function TexteInstructionsContrat() As String

    Dim t As String

    t = "Ce formulaire d{e2}clare un contrat : un type de consommation chez un op{e2}rateur, {a2} partir d'une date. "
    t = t & "L'op{e2}rateur et l'unit{e2} sont propos{e2}s {a2} la saisie d'une facture d'apr{e1}s ses dates."
    t = t & Chr(10) & "CHANGEMENT D'OP{E2}RATEUR :cr{e2}ez simplement le nouveau contrat avec sa [c:D{e2}but du contrat]. S'il y a un contrat en cours "
    t = t & "pour ce type, le programme propose de le cl{o2}turer la veille du nouveau. Les factures d{e2}j{a2} saisies gardent leur ancien op{e2}rateur."
    t = t & Chr(10) & "[c:Unit{e2} de mesure] : propos{e2}e d'apr{e1}s le type (m3 pour l'eau, kWh pour l'{e2}lectricit{e2} et le gaz), modifiable. "
    t = t & "[c:Fin du contrat (facultative)] : laissez vide pour un contrat en cours. "
    t = t & "[c:Abonnement mensuel] : donn{e2} {a2} titre indicatif ; l'abonnement r{e2}el est saisi sur chaque facture."
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
    mod_InstallCommun.MettreEnForme ws, Array(2, 24, 12, 12, 12, 12, 12, 12, 26, 2, 10, 10, 10, 10, 10, 10, 10, 10, 10), 100

    mod_InstallCommun.PoserHauteursEntete ws, 1, False
    mod_InstallCommun.EcrireInstructions ws, LIGNE_INSTRUCTIONS, 2, 2, 3, 19, TexteInstructionsConso()

    ConstruireListesTechniques ws

    ' --- Paramètres ---
    mod_InstallCommun.PoserTitreBloc ws.Range("B6:G6"), mod_Display.FR("Param{e1}tres")
    mod_InstallCommun.PoserEtiquette ws.Range("B7"), "Type"
    mod_InstallCommun.PoserEtiquette ws.Range("B8"), "Indicateur"
    mod_InstallCommun.PoserEtiquette ws.Range("B9"), mod_Display.FR("P{e2}riode")
    mod_InstallCommun.PoserEtiquette ws.Range("B10"), mod_Display.FR("Ann{e2}e")
    mod_InstallCommun.PoserEtiquette ws.Range("B11"), mod_Display.FR("Mois (fin de p{e2}riode)")
    mod_InstallCommun.PoserEtiquette ws.Range("B12"), "Forme du graphique"
    mod_InstallCommun.PoserEtiquette ws.Range("B13"), mod_Display.FR("R{e2}partition des lignes")
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
    mod_InstallCommun.PoserChampLecture ws.Range("B15:S15")
    ws.Range("B15").Font.Bold = True
    ws.Range("B15").Font.Color = mod_InstallCommun.CoulEtiquette()

    ' --- Tableau mensuel ---
    ent = mod_Consommations.CO_LIGNE_ENTETE
    ws.Cells(ent, 2).Value = "Mois"
    ws.Cells(ent, 3).Value = mod_Display.FR("Quantit{e2}")
    ws.Cells(ent, 4).Value = "Prix unitaire HT"
    ws.Cells(ent, 5).Value = "Consommation (" & ChrW(8364) & ")"
    ws.Cells(ent, 6).Value = "Abonnement (" & ChrW(8364) & ")"
    ws.Cells(ent, 7).Value = "Taxes (" & ChrW(8364) & ")"
    ws.Cells(ent, 8).Value = "Montant total (" & ChrW(8364) & ")"
    ws.Cells(ent, 9).Value = mod_Display.FR("Op{e2}rateur(s)")
    mod_InstallCommun.PoserEnteteTableau ws.Range(ws.Cells(ent, 2), ws.Cells(ent, 1 + mod_Consommations.CO_NB_COL_TABLE))
    ws.Range(ws.Cells(ent, 3), ws.Cells(ent, 8)).HorizontalAlignment = xlRight

    mod_InstallCommun.PoserCorpsTableauLecture ws.Range(ws.Cells(ent + 1, 2), ws.Cells(ent + mod_Consommations.CO_NB_LIGNES_MAX + 1, 1 + mod_Consommations.CO_NB_COL_TABLE))
    ws.Range(ws.Cells(ent + 1, 2), ws.Cells(ent + mod_Consommations.CO_NB_LIGNES_MAX + 1, 2)).NumberFormat = "mmm yyyy"
    ws.Range(ws.Cells(ent + 1, 2), ws.Cells(ent + mod_Consommations.CO_NB_LIGNES_MAX + 1, 2)).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(ent + 1, 3), ws.Cells(ent + mod_Consommations.CO_NB_LIGNES_MAX + 1, 3)).NumberFormat = "#,##0.00"
    ws.Range(ws.Cells(ent + 1, 4), ws.Cells(ent + mod_Consommations.CO_NB_LIGNES_MAX + 1, 4)).NumberFormat = "#,##0.00000"
    ws.Range(ws.Cells(ent + 1, 5), ws.Cells(ent + mod_Consommations.CO_NB_LIGNES_MAX + 1, 8)).NumberFormat = "#,##0.00"

    ' --- Graphique (ses séries sont posées par mod_Consommations à chaque actualisation) ---
    Set r = ws.Range(ws.Cells(ent, mod_Consommations.CO_COL_GRAPH_DEB), ws.Cells(ent + 28, mod_Consommations.CO_COL_GRAPH_FIN))
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
    mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, mod_Display.FR("Saisir une facture"), _
        "SaisirFacture", "btnSaisirFactureConso", 112
    mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, mod_Display.FR("Voir les factures"), _
        "AfficherFactures", "btnVoirFacturesConso", 108
    mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, "Contrats", _
        "AfficherContrats", "btnVoirContratsConso", 80
    mod_InstallCommun.AjouterBoutonEntete ws, LIGNE_BOUTONS, gauche, mod_InstallCommun.CapSortir(), _
        "SortirConsommations", "btnSortirConso", mod_InstallCommun.FRM_BTN_L
    mod_InstallCommun.AjouterBoutonAide ws, LIGNE_BOUTONS, gauche

End Sub

Private Function TexteInstructionsConso() As String

    Dim t As String

    t = "Cet {e2}cran suit l'{e2}volution des consommations d'eau, d'{e2}lectricit{e2} et de gaz, mois apr{e1}s mois, d'apr{e1}s les lignes des factures. "
    t = t & "[b:Actualiser] recalcule le tableau et le graphique apr{e1}s un changement de param{e1}tre."
    t = t & Chr(10) & "[c:Type] : un type de consommation, ou (tous) pour comparer les montants. [c:Indicateur] : Quantit{e2}, Prix unitaire HT "
    t = t & "(moyenne pond{e2}r{e2}e par les quantit{e2}s, abonnement et taxes exclus), Consommation, Abonnement, Taxes (montants TTC) ou Montant total (leur somme). "
    t = t & "Pour (tous), seuls les montants sont comparables."
    t = t & Chr(10) & "[c:P{e2}riode] : les 12, 36, 60 ou 120 derniers mois (jusqu'{a2} 10 ans) se terminant au [c:Mois (fin de p{e2}riode)] de l'[c:Ann{e2}e] choisie, ou l'ann{e2}e civile enti{e1}re."
    t = t & Chr(10) & "[c:R{e2}partition des lignes] : 'Au prorata des jours' r{e2}partit chaque ligne de facture sur les mois qu'elle couvre (chaque ligne a ses propres dates) ; "
    t = t & "'Mois de fin de la ligne' l'impute enti{e1}rement au mois de sa date de fin."
    t = t & Chr(10) & "[c:Forme du graphique] : Barres (empil{e2}es par op{e2}rateur, ou par type pour (tous)), Courbes, ou Barres + courbe du prix "
    t = t & "(pour un type choisi : les barres montrent l'indicateur et le prix unitaire se lit sur l'axe de droite)."
    t = t & Chr(10) & "Boutons :"
    t = t & Chr(10) & "- [b:Saisir une facture] : ouvre le formulaire de saisie d'une facture (abonnement, consommation, taxes)."
    t = t & Chr(10) & "- [b:Voir les factures] : affiche la liste des factures (modification et suppression)."
    t = t & Chr(10) & "- [b:Contrats] : affiche la liste des contrats (op{e2}rateurs)."
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
    ws.Cells(4, mod_Consommations.CO_COL_LISTE_INDIC).Value = mod_Consommations.LibIndConso()
    ws.Cells(5, mod_Consommations.CO_COL_LISTE_INDIC).Value = mod_Consommations.LibIndAbo()
    ws.Cells(6, mod_Consommations.CO_COL_LISTE_INDIC).Value = mod_Consommations.LibIndTaxes()
    ws.Cells(7, mod_Consommations.CO_COL_LISTE_INDIC).Value = mod_Consommations.LibIndMontant()

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
    PoserNomListe "ListeIndicConso", ws, mod_Consommations.CO_COL_LISTE_INDIC, 2, 7
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

Public Sub AfficherFormulaireFacturePourEdition()
    mod_InstallCommun.AfficherPourEdition mod_VarGlobales.NOM_FEUILLE_FACTURE, "InstallerConsommations", "MasquerFormulaireFactureApresEdition"
End Sub

Public Sub MasquerFormulaireFactureApresEdition()
    mod_InstallCommun.MasquerApresEdition mod_VarGlobales.NOM_FEUILLE_FACTURE
End Sub

Public Sub AfficherFormulaireContratPourEdition()
    mod_InstallCommun.AfficherPourEdition mod_VarGlobales.NOM_FEUILLE_CONTRAT, "InstallerConsommations", "MasquerFormulaireContratApresEdition"
End Sub

Public Sub MasquerFormulaireContratApresEdition()
    mod_InstallCommun.MasquerApresEdition mod_VarGlobales.NOM_FEUILLE_CONTRAT
End Sub
