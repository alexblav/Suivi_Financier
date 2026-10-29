Attribute VB_Name = "mod_Synthese"
Option Explicit

' =====================================================================================
' MODULE : mod_Synthese
'
' FEUILLE "SYNTHESE" : etat de la base en un coup d'oeil (ajout 28/10/2026, demande operateur).
'
' REORGANISATION DES FEUILLES
'   - L'ancienne feuille "Synthese" (boutons + parametres de recherche) s'appelle desormais "Menu"
'     (constante mod_VarGlobales.NOM_FEUILLE_MENU). C'est l'ecran central : tous les modules qui
'     l'affichaient ou la masquaient pendant les formulaires utilisent cette constante.
'   - La NOUVELLE feuille "Synthese" (mod_VarGlobales.NOM_FEUILLE_SYNTHESE) contient :
'       * a GAUCHE (colonnes A:B)  : le bloc "Previsionnel de tresorerie" (deplace depuis Menu avec ses
'         cellules nommees soldeRefMontant, soldeRefDate, soldeEcart : les formules suivent) ;
'       * a DROITE (colonnes D:H)  : les rappels, un bloc par sujet, chacun avec son titre :
'           1. Echeances en retard                 (TblEcheances)
'           2. Factures non liees : Eau et Electricite, gaz, chauffage (TblOperations)
'           3. Operations de sante a traiter ou en attente (TblOperations)
'
' MACROS
'   - CreerFeuilleSynthese : a lancer UNE FOIS (Ctrl+G). Renomme l'ancienne Synthese en Menu, cree
'     la nouvelle Synthese, y deplace le bloc de tresorerie, pose la mise en page et le bouton
'     "Actualiser", puis remplit les rappels. Relancable sans danger (rien n'est refait deux fois).
'   - ActualiserSynthese   : recalcule les rappels. Appelee par le bouton "Actualiser" et, a chaque
'     affichage de la feuille, par ThisWorkbook.Workbook_SheetActivate.
'
' REGLES DES RAPPELS (a valider par l'operateur)
'   - Echeance en retard : meme regle que "Ech. en retard" du bloc de tresorerie (date prevue
'     depassee ; pour une precision "Mois", une fois le mois ecoule).
'   - Facture non liee : operation de DEPENSE (montant negatif) de categorie ou sous-categorie
'     "Eau" ou "Electricite, gaz, chauffage" (accents et casse ignores) dont l'ID_Transaction
'     n'est cite dans la colonne ID_Transactions de TblFactures (mod_Factures.FAC_COL_LIEN ; une
'     facture peut etre liee a plusieurs operations, separees par ";").
'     Le bouton "Affecter une facture" de l'entete du bloc ouvre le formulaire frm_AffectationFacture
'     (macro AffecterFactureOperation, voir mod_AffectationFacture) pour relier l'operation
'     selectionnee a une facture existante ou a une facture neuve.
'   - Sante : operations de sante (categorie "Sante, prevoyance" ou sous-categorie "Frais, remb sante")
'     dont StatutSante = "KO" et dont le depassement d'honoraires n'est pas valide (ce sont les cas que
'     propose l'ecran de suivi sante). "En attente" = il reste un solde positif (depense pas encore
'     entierement remboursee) ; "A traiter" = les autres cas (solde nul ou negatif, remboursement sans
'     depense...). Les parts de TblVentilations ne sont pas reprises.
'
' FICHIER 100 % ASCII : les textes accentues passent par mod_Display.FR().
' =====================================================================================

Public Const SYN_LIGNE_DEBUT As Long = 3          ' premiere ligne des blocs (titre du bloc de tresorerie et du 1er rappel)
Private Const SYN_COL_RAPPELS As Long = 4         ' colonne D : debut des rappels
Private Const SYN_NB_COL As Long = 5              ' D:H
Private Const SYN_NB_MAX As Long = 15             ' lignes affichees au maximum par rappel
Private Const SYN_COL_ID_OPERATION As Long = 9    ' colonne I (masquee) : ID_Transaction de chaque operation du bloc "Factures non liees"

Private Const NOM_BOUTON_ACTUALISER As String = "btnActualiserSynthese"
Private Const NOM_BOUTON_AFFECTER As String = "btnAffecterFacture"
Private Const NOM_BOUTON_AFFICHER_TOUT As String = "btnAfficherToutFactures"


' =====================================================================================
' MACRO D'INSTALLATION
' =====================================================================================
Public Sub CreerFeuilleSynthese()

    Dim wsMenu As Worksheet, wsSyn As Worksheet
    Dim nomBloc As Name
    Dim celBloc As Range
    Dim ligneTitre As Long
    Dim reponse As VbMsgBoxResult

    Set wsMenu = mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_MENU)
    Set wsSyn = mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_SYNTHESE)

    ' --- 1. L'ancienne Synthese devient Menu -------------------------------------------
    If wsMenu Is Nothing Then
        If wsSyn Is Nothing Then
            MsgBox mod_Display.FR("Aucune feuille 'Synthese' ni 'Menu' n'a {e2}t{e2} trouv{e2}e : rien {a2} renommer."), vbExclamation
            Exit Sub
        End If
        reponse = MsgBox(mod_Display.FR("La feuille 'Synthese' actuelle va {ea}tre renomm{e2}e 'Menu'.") & vbCrLf & _
                         mod_Display.FR("Une nouvelle feuille 'Synthese' sera cr{e2}{e2}e, avec le bloc Pr{e2}visionnel de tr{e2}sorerie et les rappels.") & _
                         vbCrLf & vbCrLf & "Continuer ?", vbYesNo + vbQuestion, mod_Display.FR("Cr{e2}ation de la feuille Synthese"))
        If reponse <> vbYes Then Exit Sub
        wsSyn.Name = mod_VarGlobales.NOM_FEUILLE_MENU      ' les formules et les noms qui la citent suivent tout seuls
        Set wsMenu = wsSyn
        Set wsSyn = Nothing
    End If

    ' --- 2. Nouvelle feuille Synthese ---------------------------------------------------
    If wsSyn Is Nothing Then
        Set wsSyn = ThisWorkbook.Worksheets.Add(After:=wsMenu)
        wsSyn.Name = mod_VarGlobales.NOM_FEUILLE_SYNTHESE
    End If

    ' --- 3. Deplacement du bloc de tresorerie (seulement s'il est encore sur Menu) -------
    Set nomBloc = NomSiExiste("soldeRefMontant")
    If Not nomBloc Is Nothing Then
        On Error Resume Next
        Set celBloc = nomBloc.RefersToRange
        On Error GoTo 0
        If Not celBloc Is Nothing Then
            If StrComp(celBloc.Worksheet.Name, wsMenu.Name, vbTextCompare) = 0 Then
                ligneTitre = celBloc.Row - 1
                wsMenu.Range(wsMenu.Cells(ligneTitre, 1), wsMenu.Cells(ligneTitre + 7, 2)).Cut _
                    Destination:=wsSyn.Cells(SYN_LIGNE_DEBUT, 1)
                Application.CutCopyMode = False
            End If
        End If
    End If

    ' Complete ou repare le bloc (cree s'il n'existait pas, reecrit les formules cassees)
    mod_InstallEcheancier.ConstruireBlocSynthese

    ' --- 4. Mise en page, bouton, contenu -----------------------------------------------
    MettreEnPage wsSyn
    PoserBoutonActualiser wsSyn
    ActualiserSynthese

    wsMenu.Activate
    MsgBox mod_Display.FR("La feuille Synthese est en place. L'ancienne feuille s'appelle d{e2}sormais 'Menu'.") & vbCrLf & vbCrLf & _
           mod_Display.FR("Pour que les rappels se mettent {a2} jour {a2} chaque affichage, coller l'{e2}v{e2}nement Workbook_SheetActivate de ThisWorkbook.bas dans le module ThisWorkbook."), _
           vbInformation, mod_Display.FR("Installation termin{e2}e")

End Sub


Private Sub MettreEnPage(ByVal ws As Worksheet)

    ws.Columns(1).ColumnWidth = 26
    ws.Columns(2).ColumnWidth = 16
    ws.Columns(3).ColumnWidth = 3
    ws.Columns(4).ColumnWidth = 12
    ws.Columns(5).ColumnWidth = 46
    ws.Columns(6).ColumnWidth = 30
    ws.Columns(7).ColumnWidth = 13
    ws.Columns(8).ColumnWidth = 36

    With ws.Cells(1, 1)
        .Value = mod_Display.FR("Synth{e1}se de l'{e2}tat de la base")
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = mod_InstallCommun.CoulNomFeuille()
    End With
    ws.Cells(2, 1).Font.Italic = True
    ws.Cells(2, 1).Font.Color = RGB(110, 110, 110)

    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Range("A1").Select
    On Error GoTo 0

End Sub

Private Sub PoserBoutonActualiser(ByVal ws As Worksheet)

    Dim btn As Button
    Dim cel As Range

    On Error Resume Next
    ws.Shapes(NOM_BOUTON_ACTUALISER).Delete
    On Error GoTo 0

    Set cel = ws.Cells(1, SYN_COL_RAPPELS + SYN_NB_COL - 1)
    Set btn = ws.Buttons.Add(cel.Left + cel.Width - 110, cel.Top + 2, 108, 20)
    btn.Caption = "Actualiser"
    btn.OnAction = "ActualiserSynthese"
    btn.Name = NOM_BOUTON_ACTUALISER

End Sub


' =====================================================================================
' MACRO D'ACTUALISATION
' =====================================================================================
Public Sub ActualiserSynthese()

    Dim ws As Worksheet
    Dim ligne As Long
    Dim evtAvant As Boolean, ecranAvant As Boolean

    Set ws = mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_SYNTHESE)
    If ws Is Nothing Then Exit Sub

    evtAvant = Application.EnableEvents
    ecranAvant = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo Fin

    ws.Range(ws.Cells(SYN_LIGNE_DEBUT, SYN_COL_RAPPELS), ws.Cells(ws.Rows.Count, SYN_COL_ID_OPERATION)).Clear
    ws.Columns(SYN_COL_ID_OPERATION).Hidden = True

    ligne = SYN_LIGNE_DEBUT
    ligne = BlocEcheancesEnRetard(ws, ligne)
    ligne = BlocFacturesNonLiees(ws, ligne)
    ligne = BlocSante(ws, ligne)

    ws.Cells(2, 1).Value = mod_Display.FR("Actualis{e2} le ") & Format(Now, "dd/mm/yyyy hh:nn")

Fin:
    Application.EnableEvents = evtAvant
    Application.ScreenUpdating = ecranAvant

End Sub


' =====================================================================================
' BLOC 1 : ECHEANCES EN RETARD
' =====================================================================================
Private Function BlocEcheancesEnRetard(ByVal ws As Worksheet, ByVal ligneDebut As Long) As Long

    Dim lo As ListObject
    Dim nbL As Long, i As Long, nb As Long
    Dim vDate As Variant, vPrec As Variant, vLib As Variant, vTiers As Variant, vCat As Variant, vMont As Variant
    Dim cles() As Double, vals() As Variant
    Dim res As Variant
    Dim jourPrevu As Double, limite As Double
    Dim texte As String
    Dim libTitre As String, msgVide As String

    libTitre = mod_Display.FR("{E2}ch{e2}ances en retard")
    msgVide = mod_Display.FR("Aucune {e2}ch{e2}ance en retard.")

    Set lo = TableParNom(mod_VarGlobales.NOM_FEUILLE_ECHEANCIER, mod_VarGlobales.NOM_TABLE_ECHEANCES)
    If lo Is Nothing Then
        msgVide = mod_Display.FR("L'{e2}ch{e2}ancier n'est pas install{e2}.")
    ElseIf Not lo.DataBodyRange Is Nothing Then
        nbL = lo.DataBodyRange.Rows.Count
        vDate = ValeursColonne(lo, "DatePrevue", nbL)
        vPrec = ValeursColonne(lo, "PrecisionDate", nbL)
        vLib = ValeursColonne(lo, "Libelle", nbL)
        vTiers = ValeursColonne(lo, "Tiers", nbL)
        vCat = ValeursColonne(lo, "Categorie", nbL)
        vMont = ValeursColonne(lo, "Montant", nbL)
        ReDim cles(1 To nbL)
        ReDim vals(1 To nbL, 1 To SYN_NB_COL)

        For i = 1 To nbL
            If IsNumeric(Cellule(vDate, i)) And Not IsEmpty(Cellule(vDate, i)) Then
                jourPrevu = CDbl(Cellule(vDate, i))
                limite = jourPrevu
                ' Precision "Mois" : en retard seulement une fois le mois ecoule (dernier jour du mois)
                If mod_DataStructure.CellText(Cellule(vPrec, i)) = "Mois" Then
                    limite = CDbl(DateSerial(Year(CDate(jourPrevu)), Month(CDate(jourPrevu)) + 1, 0))
                End If
                If limite < CDbl(Date) Then
                    nb = nb + 1
                    cles(nb) = jourPrevu
                    texte = mod_DataStructure.CellText(Cellule(vLib, i))
                    If mod_DataStructure.CellText(Cellule(vTiers, i)) <> "" Then
                        If texte <> "" Then texte = texte & " - "
                        texte = texte & mod_DataStructure.CellText(Cellule(vTiers, i))
                    End If
                    vals(nb, 1) = jourPrevu
                    vals(nb, 2) = texte
                    vals(nb, 3) = mod_DataStructure.CellText(Cellule(vCat, i))
                    vals(nb, 4) = mod_DataStructure.ToDouble(Cellule(vMont, i))
                    vals(nb, 5) = CLng(CDbl(Date) - limite) & " j"
                End If
            End If
        Next i
    End If

    If nb > 0 Then Extraire cles, vals, nb, False, res

    BlocEcheancesEnRetard = EcrireBloc(ws, ligneDebut, libTitre, _
        Array(mod_Display.FR("Date pr{e2}vue"), mod_Display.FR("{E2}ch{e2}ance"), mod_Display.FR("Cat{e2}gorie"), "Montant", "Retard"), _
        res, nb, msgVide)

End Function


' =====================================================================================
' BLOC 2 : OPERATIONS EAU / ELECTRICITE SANS FACTURE LIEE
' =====================================================================================
Private Function BlocFacturesNonLiees(ByVal ws As Worksheet, ByVal ligneDebut As Long) As Long

    Dim lo As ListObject
    Dim nbL As Long, i As Long, nb As Long
    Dim vId As Variant, vDate As Variant, vTiers As Variant, vCat As Variant, vSous As Variant, vMont As Variant
    Dim cles() As Double, vals() As Variant
    Dim res As Variant
    Dim liens As Object
    Dim catTexte As String, sousTexte As String
    Dim libTitre As String, msgVide As String
    Dim ids() As String
    Dim nbAff As Long

    libTitre = mod_Display.FR("Factures non li{e2}es : Eau, {E2}lectricit{e2} - gaz - chauffage")
    msgVide = mod_Display.FR("Toutes les op{e2}rations Eau / {E2}lectricit{e2} ont une facture li{e2}e.")

    Set lo = mod_DonneesTable.GetOperationsValue(mod_VarGlobales.NOM_FEUILLE_DONNEES, "TblOperations")
    If lo Is Nothing Then
        msgVide = mod_Display.FR("Tableau des op{e2}rations introuvable.")
    ElseIf Not lo.DataBodyRange Is Nothing Then
        nbL = lo.DataBodyRange.Rows.Count
        Set liens = ChargerLiensFactures()
        vId = ValeursColonne(lo, "ID_Transaction", nbL)
        vDate = ValeursColonne(lo, "Date_Comptable", nbL)
        vTiers = ValeursColonne(lo, "Tiers", nbL)
        vCat = ValeursColonne(lo, "Categorie", nbL)
        vSous = ValeursColonne(lo, "SousCategorie", nbL)
        vMont = ValeursColonne(lo, "Montant", nbL)
        ReDim cles(1 To nbL)
        ReDim vals(1 To nbL, 1 To SYN_NB_COL + 1)      ' derniere colonne : ID_Transaction (non affiche dans D:H)

        For i = 1 To nbL
            If mod_DataStructure.ToDouble(Cellule(vMont, i)) < 0 Then
                catTexte = mod_DataStructure.CellText(Cellule(vCat, i))
                sousTexte = mod_DataStructure.CellText(Cellule(vSous, i))
                If EstCategorieFacturee(catTexte) Or EstCategorieFacturee(sousTexte) Then
                    If Not liens.Exists(mod_DataStructure.CellText(Cellule(vId, i))) Then
                        nb = nb + 1
                        If IsNumeric(Cellule(vDate, i)) And Not IsEmpty(Cellule(vDate, i)) Then cles(nb) = CDbl(Cellule(vDate, i))
                        vals(nb, 1) = cles(nb)
                        vals(nb, 2) = mod_DataStructure.CellText(Cellule(vTiers, i))
                        If sousTexte <> "" And sousTexte <> catTexte Then
                            vals(nb, 3) = catTexte & " / " & sousTexte
                        Else
                            vals(nb, 3) = catTexte
                        End If
                        vals(nb, 4) = mod_DataStructure.ToDouble(Cellule(vMont, i))
                        vals(nb, 5) = mod_Display.FR("Facture {a2} relier")
                        vals(nb, SYN_NB_COL + 1) = mod_DataStructure.CellText(Cellule(vId, i))
                    End If
                End If
            End If
        Next i
    End If

    If nb > 0 Then
        Extraire cles, vals, nb, True, res
        nbAff = UBound(res, 1)
        ReDim ids(1 To nbAff)
        For i = 1 To nbAff
            ids(i) = res(i, SYN_NB_COL + 1)
        Next i
        ReDim Preserve res(1 To nbAff, 1 To SYN_NB_COL)
    End If

    BlocFacturesNonLiees = EcrireBloc(ws, ligneDebut, libTitre, _
        Array("Date", "Tiers", mod_Display.FR("Cat{e2}gorie"), "Montant", "Facture"), _
        res, nb, msgVide)

    ' ID_Transaction de chaque ligne affichee (colonne masquee) : c'est elle qui designe l'operation a relier.
    If nb > 0 Then
        With ws.Range(ws.Cells(ligneDebut + 2, SYN_COL_ID_OPERATION), ws.Cells(ligneDebut + 1 + nbAff, SYN_COL_ID_OPERATION))
            .NumberFormat = "@"
            For i = 1 To nbAff
                .Cells(i, 1).Value = ids(i)
            Next i
        End With
    End If

    PoserBoutonAffecter ws, ligneDebut

End Function

Private Sub PoserBoutonAffecter(ByVal ws As Worksheet, ByVal ligneTitre As Long)

    Dim btn As Button
    Dim cel As Range

    On Error Resume Next
    ws.Shapes(NOM_BOUTON_AFFECTER).Delete
    ws.Shapes(NOM_BOUTON_AFFICHER_TOUT).Delete
    On Error GoTo 0

    Set cel = ws.Cells(ligneTitre, SYN_COL_RAPPELS + SYN_NB_COL - 1)
    Set btn = ws.Buttons.Add(cel.Left + cel.Width - 150, cel.Top + 1, 148, cel.Height - 2)
    btn.Caption = mod_Display.FR("Affecter une facture")
    btn.OnAction = "AffecterFactureOperation"
    btn.Name = NOM_BOUTON_AFFECTER

    Set btn = ws.Buttons.Add(cel.Left + cel.Width - 150 - 92, cel.Top + 1, 90, cel.Height - 2)
    btn.Caption = "Afficher tout"
    btn.OnAction = "AfficherToutesFacturesNonLiees"
    btn.Name = NOM_BOUTON_AFFICHER_TOUT

End Sub

' Bouton "Afficher tout" : toutes les operations du bloc (sans la limite de SYN_NB_MAX lignes) dans
' frm_RechercheOperations, ou elles peuvent aussi etre reliees a une facture (bouton "Affecter une facture").
Public Sub AfficherToutesFacturesNonLiees()
    mod_RechercheOperations.RechercherOperations "FacturesNonLiees"
End Sub

' =====================================================================================
' AFFECTATION D'UNE FACTURE A L'OPERATION SELECTIONNEE (bouton de l'entete du bloc 2)
' =====================================================================================
' 1. L'operateur clique une ligne du bloc "Factures non liees" puis sur le bouton.
' 2. Le formulaire frm_AffectationFacture (mod_AffectationFacture) reprend l'operation en entete et
'    propose, par filtres progressifs (type de consommation, puis date de la facture), la facture
'    qui l'a reglee. Valider ecrit le lien, Annuler n'ecrit rien, Nouvelle facture ouvre la saisie
'    d'une facture deja liee a l'operation.
' Plusieurs operations peuvent etre liees a une meme facture (echeances).
Public Sub AffecterFactureOperation()

    Dim ws As Worksheet
    Dim ligne As Long
    Dim idOperation As String
    Dim montantOp As Double, jourOp As Double

    Set ws = mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_SYNTHESE)
    If ws Is Nothing Then Exit Sub
    If Not ActiveSheet Is ws Then ws.Activate

    ' --- Operation selectionnee ---
    If TypeName(Selection) <> "Range" Then
        MsgBox mod_Display.FR("S{e2}lectionnez d'abord une op{e2}ration du tableau Factures non li{e2}es."), vbInformation
        Exit Sub
    End If
    ligne = Selection.Row
    If Selection.Rows.Count > 1 Then ligne = 0
    If ligne > 0 Then idOperation = mod_DataStructure.CellText(ws.Cells(ligne, SYN_COL_ID_OPERATION).Value2)
    If idOperation = "" Then
        MsgBox mod_Display.FR("S{e2}lectionnez d'abord UNE ligne d'op{e2}ration du tableau Factures non li{e2}es, puis cliquez sur le bouton."), vbInformation
        Exit Sub
    End If
    If IsNumeric(ws.Cells(ligne, SYN_COL_RAPPELS).Value2) Then jourOp = CDbl(ws.Cells(ligne, SYN_COL_RAPPELS).Value2)
    montantOp = mod_DataStructure.ToDouble(ws.Cells(ligne, SYN_COL_RAPPELS + 3).Value2)

    If mod_AffectationFacture.AffecterFacture(idOperation, jourOp, _
            mod_DataStructure.CellText(ws.Cells(ligne, SYN_COL_RAPPELS + 1).Value2), _
            mod_DataStructure.CellText(ws.Cells(ligne, SYN_COL_RAPPELS + 2).Value2), montantOp) Then
        ActualiserSynthese
    End If

End Sub

' Vrai si le texte est la categorie Eau ou Electricite, gaz, chauffage (accents et casse ignores).
' Public : la meme regle sert au prefiltre "FacturesNonLiees" de frm_RechercheOperations.
Public Function EstCategorieFacturee(ByVal texte As String) As Boolean
    Dim t As String
    t = NormaliserTexte(texte)
    If t = "" Then Exit Function
    EstCategorieFacturee = (t = "eau") Or (t = "electricite, gaz, chauffage")
End Function

Private Function NormaliserTexte(ByVal texte As String) As String
    Dim t As String
    t = LCase$(Trim$(texte))
    t = Replace(t, ChrW(233), "e")
    t = Replace(t, ChrW(232), "e")
    t = Replace(t, ChrW(234), "e")
    NormaliserTexte = t
End Function

' ID_Transaction des operations deja reliees a une facture. Vide tant que TblFactures n'a pas la
' colonne de liaison (voir mod_Factures.FAC_COL_LIEN) : aucune operation n'est alors consideree comme liee.
Public Function ChargerLiensFactures() As Object

    Dim dic As Object
    Dim lo As ListObject
    Dim nbL As Long, i As Long
    Dim vLien As Variant
    Dim cle As String
    Dim element As Variant

    Set dic = CreateObject("Scripting.Dictionary")
    Set lo = mod_Factures.TableFactures()
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            nbL = lo.DataBodyRange.Rows.Count
            vLien = ValeursColonne(lo, mod_Factures.FAC_COL_LIEN, nbL)
            For i = 1 To nbL
                ' Une facture peut etre reglee par plusieurs operations : la cellule en contient la liste.
                For Each element In Split(mod_DataStructure.CellText(Cellule(vLien, i)), mod_Factures.FAC_SEP_LIENS)
                    cle = Trim$(CStr(element))
                    If cle <> "" Then
                        If Not dic.Exists(cle) Then dic.Add cle, True
                    End If
                Next element
            Next i
        End If
    End If
    Set ChargerLiensFactures = dic

End Function


' =====================================================================================
' BLOC 3 : OPERATIONS DE SANTE A TRAITER OU EN ATTENTE
' =====================================================================================
Private Function BlocSante(ByVal ws As Worksheet, ByVal ligneDebut As Long) As Long

    Dim lo As ListObject
    Dim nbL As Long, i As Long, nb As Long
    Dim vDate As Variant, vTiers As Variant, vCat As Variant, vSous As Variant, vMont As Variant
    Dim vStatut As Variant, vSolde As Variant, vDepass As Variant
    Dim cles() As Double, vals() As Variant
    Dim res As Variant
    Dim catTexte As String, sousTexte As String
    Dim soldeSante As Double
    Dim libTitre As String, msgVide As String

    libTitre = mod_Display.FR("Sant{e2} : op{e2}rations {a2} traiter ou en attente")
    msgVide = mod_Display.FR("Aucune op{e2}ration de sant{e2} {a2} traiter.")

    Set lo = mod_DonneesTable.GetOperationsValue(mod_VarGlobales.NOM_FEUILLE_DONNEES, "TblOperations")
    If lo Is Nothing Then
        msgVide = mod_Display.FR("Tableau des op{e2}rations introuvable.")
    ElseIf Not lo.DataBodyRange Is Nothing Then
        nbL = lo.DataBodyRange.Rows.Count
        vDate = ValeursColonne(lo, "Date_Comptable", nbL)
        vTiers = ValeursColonne(lo, "Tiers", nbL)
        vCat = ValeursColonne(lo, "Categorie", nbL)
        vSous = ValeursColonne(lo, "SousCategorie", nbL)
        vMont = ValeursColonne(lo, "Montant", nbL)
        vStatut = ValeursColonne(lo, "StatutSante", nbL)
        vSolde = ValeursColonne(lo, "SoldeSante", nbL)
        vDepass = ValeursColonne(lo, "DepassementHoraires", nbL)
        ReDim cles(1 To nbL)
        ReDim vals(1 To nbL, 1 To SYN_NB_COL)

        For i = 1 To nbL
            If UCase$(mod_DataStructure.CellText(Cellule(vStatut, i))) = "KO" And Not EstVrai(Cellule(vDepass, i)) Then
                catTexte = mod_DataStructure.CellText(Cellule(vCat, i))
                sousTexte = mod_DataStructure.CellText(Cellule(vSous, i))
                If catTexte = mod_VarGlobales.CATEGORIE_SANTE Or sousTexte = mod_VarGlobales.SOUS_CATEGORIE_SANTE Then
                    nb = nb + 1
                    If IsNumeric(Cellule(vDate, i)) And Not IsEmpty(Cellule(vDate, i)) Then cles(nb) = CDbl(Cellule(vDate, i))
                    vals(nb, 1) = cles(nb)
                    vals(nb, 2) = mod_DataStructure.CellText(Cellule(vTiers, i))
                    If sousTexte <> "" And sousTexte <> catTexte Then
                        vals(nb, 3) = catTexte & " / " & sousTexte
                    Else
                        vals(nb, 3) = catTexte
                    End If
                    vals(nb, 4) = mod_DataStructure.ToDouble(Cellule(vMont, i))
                    soldeSante = mod_DataStructure.ToDouble(Cellule(vSolde, i))
                    If soldeSante > 0.005 Then
                        vals(nb, 5) = mod_Display.FR("En attente de remboursement (reste ") & Format(soldeSante, "#,##0.00") & ")"
                    Else
                        vals(nb, 5) = mod_Display.FR("{A2} traiter")
                    End If
                End If
            End If
        Next i
    End If

    If nb > 0 Then Extraire cles, vals, nb, False, res

    BlocSante = EcrireBloc(ws, ligneDebut, libTitre, _
        Array("Date", "Tiers", mod_Display.FR("Cat{e2}gorie"), "Montant", "Statut"), _
        res, nb, msgVide)

End Function

Private Function EstVrai(ByVal v As Variant) As Boolean
    If VarType(v) = vbBoolean Then
        EstVrai = v
    Else
        Select Case UCase$(mod_DataStructure.CellText(v))
            Case "VRAI", "TRUE", "OUI"
                EstVrai = True
        End Select
    End If
End Function


' =====================================================================================
' ECRITURE D'UN BLOC : titre (avec le nombre d'elements), en-tetes, lignes, message si vide
' =====================================================================================
' res : tableau (1..k, 1..SYN_NB_COL) deja trie, k = Min(nbTotal, SYN_NB_MAX) ; nbTotal : nombre reel d'elements.
' Retourne la premiere ligne libre apres le bloc (une ligne vide est laissee entre deux blocs).
Private Function EcrireBloc(ByVal ws As Worksheet, ByVal ligneDebut As Long, ByVal libTitre As String, _
                            ByVal entetes As Variant, ByRef res As Variant, ByVal nbTotal As Long, _
                            ByVal msgVide As String) As Long

    Dim cDeb As Long, cFin As Long
    Dim l As Long, k As Long, nbAff As Long
    Dim zone As Range

    cDeb = SYN_COL_RAPPELS
    cFin = SYN_COL_RAPPELS + SYN_NB_COL - 1
    l = ligneDebut

    ' Titre
    With ws.Range(ws.Cells(l, cDeb), ws.Cells(l, cFin))
        .Interior.Color = RGB(141, 180, 226)
        .Font.Bold = True
    End With
    ws.Cells(l, cDeb).Value = libTitre & "  (" & nbTotal & ")"
    l = l + 1

    ' En-tetes
    With ws.Range(ws.Cells(l, cDeb), ws.Cells(l, cFin))
        .Interior.Color = RGB(221, 235, 247)
        .Font.Bold = True
    End With
    For k = 0 To SYN_NB_COL - 1
        ws.Cells(l, cDeb + k).Value = entetes(k)
    Next k
    ws.Cells(l, cDeb + 3).HorizontalAlignment = xlRight
    l = l + 1

    If nbTotal = 0 Then
        With ws.Cells(l, cDeb)
            .Value = msgVide
            .Font.Italic = True
            .Font.Color = RGB(110, 110, 110)
        End With
        EcrireBloc = l + 2
        Exit Function
    End If

    nbAff = nbTotal
    If nbAff > SYN_NB_MAX Then nbAff = SYN_NB_MAX

    Set zone = ws.Range(ws.Cells(l, cDeb), ws.Cells(l + nbAff - 1, cFin))
    ws.Range(ws.Cells(l, cDeb), ws.Cells(l + nbAff - 1, cDeb)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(l, cDeb), ws.Cells(l + nbAff - 1, cDeb)).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(l, cDeb + 3), ws.Cells(l + nbAff - 1, cDeb + 3)).NumberFormat = "#,##0.00"
    zone.Value = res
    With zone.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Color = RGB(200, 200, 200)
    End With
    With zone.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(200, 200, 200)
    End With
    l = l + nbAff

    If nbTotal > nbAff Then
        With ws.Cells(l, cDeb)
            .Value = "... " & mod_Display.FR("et ") & (nbTotal - nbAff) & mod_Display.FR(" autre(s) non affich{e2}(es).")
            .Font.Italic = True
            .Font.Color = RGB(110, 110, 110)
        End With
        l = l + 1
    End If

    EcrireBloc = l + 1

End Function


' =====================================================================================
' OUTILS
' =====================================================================================
' Trie les nb premieres lignes de vals selon cles (insertion, stable) et copie les SYN_NB_MAX premieres dans res.
Private Sub Extraire(ByRef cles() As Double, ByRef vals() As Variant, ByVal nb As Long, _
                     ByVal decroissant As Boolean, ByRef res As Variant)

    Dim ordre() As Long
    Dim i As Long, j As Long, k As Long, courant As Long
    Dim nbAff As Long
    Dim avant As Boolean

    ReDim ordre(1 To nb)
    For i = 1 To nb
        courant = i
        j = i - 1
        Do While j >= 1
            If decroissant Then
                avant = cles(courant) > cles(ordre(j))
            Else
                avant = cles(courant) < cles(ordre(j))
            End If
            If Not avant Then Exit Do
            ordre(j + 1) = ordre(j)
            j = j - 1
        Loop
        ordre(j + 1) = courant
    Next i

    nbAff = nb
    If nbAff > SYN_NB_MAX Then nbAff = SYN_NB_MAX
    ReDim res(1 To nbAff, 1 To UBound(vals, 2))
    For i = 1 To nbAff
        For k = 1 To UBound(vals, 2)
            res(i, k) = vals(ordre(i), k)
        Next k
    Next i

End Sub

' Valeurs d'une colonne de tableau sous forme de tableau (1..n, 1..1) ; Empty si la colonne n'existe pas.
Private Function ValeursColonne(ByVal lo As ListObject, ByVal nomCol As String, ByVal nbLignes As Long) As Variant

    Dim col As ListColumn
    Dim r() As Variant

    On Error Resume Next
    Set col = lo.ListColumns(nomCol)
    On Error GoTo 0
    If col Is Nothing Then Exit Function

    If nbLignes = 1 Then
        ReDim r(1 To 1, 1 To 1)
        r(1, 1) = col.DataBodyRange.Cells(1, 1).Value2
        ValeursColonne = r
    Else
        ValeursColonne = col.DataBodyRange.Value2
    End If

End Function

Private Function Cellule(ByRef valeurs As Variant, ByVal i As Long) As Variant
    If IsArray(valeurs) Then Cellule = valeurs(i, 1)
End Function

Private Function TableParNom(ByVal nomFeuille As String, ByVal nomTable As String) As ListObject
    Dim ws As Worksheet
    Set ws = mod_InstallCommun.TrouverFeuille(nomFeuille)
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    Set TableParNom = ws.ListObjects(nomTable)
    On Error GoTo 0
End Function

Private Function NomSiExiste(ByVal nomDefini As String) As Name
    On Error Resume Next
    Set NomSiExiste = ThisWorkbook.Names(nomDefini)
    On Error GoTo 0
End Function
