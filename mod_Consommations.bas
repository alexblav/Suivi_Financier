Attribute VB_Name = "mod_Consommations"
Option Explicit

' =====================================================================================
' MODULE : mod_Consommations
'
' SUIVI DES CONSOMMATIONS (eau, électricité, gaz) - logique des contrats et du graphique
' mensuel. La saisie des FACTURES est dans mod_Factures ; la construction des feuilles est
' dans mod_InstallConsommations (macro InstallerConsommations).
'
' CE QUE FAIT LE MODULE :
'   - TblContrats : un contrat = un type de consommation chez un opérateur, entre deux dates
'     (feuille de données masquée Conso_Contrats, écran frm_Contrats). Le contrat sert à
'     proposer l'opérateur et l'unité à la saisie d'une facture.
'   - L'écran frm_Consommations montre l'évolution mois par mois (barres, courbes, ou
'     barres + courbe du prix) pour un type, ou les montants de tous les types, d'après les
'     LIGNES des factures (abonnement, consommation, taxes), chacune avec ses propres dates.
'
' RÈGLES DE CALCUL DU GRAPHIQUE :
'   - Chaque ligne de facture est RÉPARTIE sur les mois qu'elle couvre au prorata des jours
'     (quantité, montant HT et montant TTC) ; l'autre choix met toute la ligne dans le mois
'     de sa date de fin.
'   - Consommation = lignes de consommation (TTC) ; Abonnement = lignes d'abonnement (TTC) ;
'     Taxes = lignes de taxes (TTC) ; Montant total = leur somme.
'   - Prix unitaire d'un mois = montant HT de consommation / quantité (moyenne pondérée),
'     abonnement et taxes exclus.
'
' POINTS D'ENTRÉE (boutons) : AfficherConsommations (feuille Synthese), ActualiserConsommations,
'   AfficherFactures, AfficherContrats, RetourConsommations, SortirConsommations,
'   AjouterContrat, SupprimerContratSelection.
' Appelées par le code des feuilles-formulaires : CtrTraiterChangement, CtrVerrouiller
'   (voir CodeBehind_frm_Contrat.txt) ; FacTraiterChangement, FacVerrouiller de mod_Factures
'   (voir CodeBehind_frm_Facture.txt).
'
' NOTE : les petits outils de lecture de tableau (ObtenirTable, ColIdx, TableVide...) sont
' recopiés de mod_Echeancier, où ils sont privés. Une factorisation est possible plus tard.
'
' À PROPOS DES ACCENTS : textes affichés via mod_Display.FR(); commentaires en UTF-8.
' =====================================================================================

' --- Disposition de l'écran frm_Consommations (partagée avec mod_InstallConsommations) ---
Public Const CO_ADR_TYPE As String = "C7"
Public Const CO_ADR_INDIC As String = "C8"
Public Const CO_ADR_PERIODE As String = "C9"
Public Const CO_ADR_ANNEE As String = "C10"
Public Const CO_ADR_MOIS As String = "C11"
Public Const CO_ADR_FORME As String = "C12"
Public Const CO_ADR_REPART As String = "C13"
Public Const CO_ADR_MESSAGE As String = "B15"
Public Const CO_LIGNE_ENTETE As Long = 17       ' en-tête du tableau mensuel
Public Const CO_NB_LIGNES_MAX As Long = 120     ' lignes de données (120 mois = 10 ans au plus)
Public Const CO_NB_COL_TABLE As Long = 8        ' colonnes B à I du tableau mensuel
Public Const CO_COL_GRAPH_DEB As Long = 11      ' le graphique occupe les colonnes K à S
Public Const CO_COL_GRAPH_FIN As Long = 19
Public Const CO_GRAPH As String = "chConsommation"

' --- Colonnes techniques masquées de frm_Consommations ---
' Données du graphique : mois, prix unitaire moyen, puis une colonne par série (opérateur ou type).
Public Const CO_NB_SERIES As Long = 6
Public Const CO_COL_H_LIB As Long = 30          ' AD
Public Const CO_COL_H_PRIX As Long = 31         ' AE
Public Const CO_COL_H_SERIE As Long = 32        ' AF à AK
Public Const CO_COL_LISTE_TYPE As Long = 40     ' AN
Public Const CO_COL_LISTE_INDIC As Long = 41
Public Const CO_COL_LISTE_PERIODE As Long = 42
Public Const CO_COL_LISTE_FORME As Long = 43
Public Const CO_COL_LISTE_REPART As Long = 44
Public Const CO_COL_LISTE_MOIS As Long = 45
Public Const CO_COL_LISTE_ANNEE As Long = 46    ' AT

' --- Formulaire frm_Contrat ---
Public Const CF_ADR_TYPE As String = "C6"
Public Const CF_ADR_OPERATEUR As String = "C7"
Public Const CF_ADR_UNITE As String = "C8"
Public Const CF_ADR_DEBUT As String = "C9"
Public Const CF_ADR_FIN As String = "C10"
Public Const CF_ADR_ABO As String = "C11"
Public Const CF_ADR_REFERENCE As String = "C12"
Public Const CF_ADR_MESSAGE As String = "B14"           ' zone fusionnée B14:E14

' Un contrat tel que saisi dans le formulaire.
Public Type TypeContrat
    Typ As String
    Operateur As String
    Unite As String
    Debut As Date
    Fin As Date                 ' 0 = contrat en cours (pas de date de fin)
    AboMensuel As Double
    Reference As String
End Type

' --- État des formulaires de saisie des contrats ---
Private g_CtrEnCours As Boolean
Private g_CtrVerrouActif As Boolean
Private g_CtrValide As Boolean
Private g_CtrRes As TypeContrat
Private g_CtrLigneAClore As Long
Private g_CtrNomFeuillePrec As String

' --- Contrats chargés en mémoire (voir ChargerContrats) ---
' Les dates sont des nombres entiers (numéro de série Excel) ; 0 = absente.
Private c_Nb As Long
Private c_Typ() As String
Private c_Oper() As String
Private c_Unite() As String
Private c_Deb() As Double
Private c_Fin() As Double
Private c_Abo() As Double
Private c_Ligne() As Long        ' numéro de la ligne dans TblContrats


' =====================================================================================
' LIBELLÉS (fonctions : FR() ne peut pas servir dans une constante)
' =====================================================================================
Public Function TypeAdduction() As String
    TypeAdduction = "Adduction eau"
End Function

Public Function TypeRecyclage() As String
    TypeRecyclage = "Recyclage eau"
End Function

Public Function TypeElec() As String
    TypeElec = mod_Display.FR("Fourniture {e2}lec")
End Function

Public Function TypeGaz() As String
    TypeGaz = "Fourniture gaz"
End Function

Public Function TypesConso() As Variant
    TypesConso = Array(TypeAdduction(), TypeRecyclage(), TypeElec(), TypeGaz())
End Function

Public Function LibTous() As String
    LibTous = "(tous)"
End Function

Public Function LibIndQuantite() As String
    LibIndQuantite = mod_Display.FR("Quantit{e2}")
End Function

Public Function LibIndPrix() As String
    LibIndPrix = "Prix unitaire HT"
End Function

Public Function LibIndConso() As String
    LibIndConso = "Consommation"
End Function

Public Function LibIndTaxes() As String
    LibIndTaxes = "Taxes"
End Function

Public Function LibIndAbo() As String
    LibIndAbo = "Abonnement"
End Function

Public Function LibIndMontant() As String
    LibIndMontant = "Montant total"
End Function

Public Function LibPeriode12() As String
    LibPeriode12 = "12 derniers mois"
End Function

Public Function LibPeriode36() As String
    LibPeriode36 = "36 derniers mois"
End Function

Public Function LibPeriode60() As String
    LibPeriode60 = "60 derniers mois"
End Function

Public Function LibPeriode120() As String
    LibPeriode120 = "120 derniers mois"
End Function

Public Function LibPeriodeAnnee() As String
    LibPeriodeAnnee = mod_Display.FR("Ann{e2}e")
End Function

Public Function LibFormeBarres() As String
    LibFormeBarres = "Barres"
End Function

Public Function LibFormeCourbes() As String
    LibFormeCourbes = "Courbes"
End Function

Public Function LibFormeMixte() As String
    LibFormeMixte = "Barres + courbe du prix"
End Function

Public Function LibRepartProrata() As String
    LibRepartProrata = "Au prorata des jours"
End Function

Public Function LibRepartFin() As String
    LibRepartFin = mod_Display.FR("Mois de fin de la ligne")
End Function

Public Function LibSansContrat() As String
    LibSansContrat = "(sans contrat)"
End Function

Public Function LibAutres() As String
    LibAutres = "Autres"
End Function

' Unité proposée par défaut selon le type (modifiable dans le contrat).
Public Function UniteDefaut(ByVal typ As String) As String
    If StrComp(typ, TypeAdduction(), vbTextCompare) = 0 Or StrComp(typ, TypeRecyclage(), vbTextCompare) = 0 Then
        UniteDefaut = "m" & ChrW(179)
    Else
        UniteDefaut = "kWh"
    End If
End Function

Private Function EstUniteDefaut(ByVal u As String) As Boolean
    EstUniteDefaut = (StrComp(u, "m" & ChrW(179), vbTextCompare) = 0 Or StrComp(u, "kWh", vbTextCompare) = 0)
End Function

Public Function TypeConnu(ByVal typ As String) As Boolean
    Dim t As Variant
    For Each t In TypesConso()
        If StrComp(CStr(t), typ, vbTextCompare) = 0 Then
            TypeConnu = True
            Exit Function
        End If
    Next t
End Function

Public Function SymboleEuro() As String
    SymboleEuro = ChrW(8364)
End Function


' =====================================================================================
' OUTILS DE LECTURE / ÉCRITURE DES TABLEAUX
' =====================================================================================
Public Function ObtenirTable(ByVal nomFeuille As String, ByVal nomTable As String) As ListObject
    On Error Resume Next
    Set ObtenirTable = ThisWorkbook.Worksheets(nomFeuille).ListObjects(nomTable)
    On Error GoTo 0
End Function

' Les contrats sont sur une feuille de données masquée ; l'écran frm_Contrats n'en montre
' qu'une copie modifiable (TblContratsAff), réécrite dans les données quand on quitte l'écran.
Private Function TableContrats() As ListObject
    Set TableContrats = ObtenirTable(mod_VarGlobales.NOM_FEUILLE_DATA_CONTRATS, mod_VarGlobales.NOM_TABLE_CONTRATS)
End Function

Private Function TableAffContrats() As ListObject
    Set TableAffContrats = ObtenirTable(mod_VarGlobales.NOM_FEUILLE_CONTRATS, mod_VarGlobales.NOM_TABLE_CONTRATS_AFF)
End Function

' Lettre de la colonne de numéro donné (10 -> "J").
Public Function AdresseCol(ByVal col As Long) As String
    AdresseCol = Split(ThisWorkbook.Worksheets(1).Columns(col).Address(False, False), ":")(0)
End Function

' Format d'affichage d'une colonne des tableaux de consommations (par son NOM).
Public Function FormatColonneConso(ByVal nom As String) As String
    Select Case nom
        Case "ID_Contrat", "ID_Facture": FormatColonneConso = "0"
        Case "DateDebut", "DateFin", "DateFacture", "DebutPeriode", "FinPeriode": FormatColonneConso = "dd/mm/yyyy"
        Case "Quantite": FormatColonneConso = "#,##0.00"
        Case "PrixUnitaireHT": FormatColonneConso = "#,##0.00000"
        Case "TauxTVA": FormatColonneConso = "0.0%"
        Case "AbonnementMensuel", "MontantHT", "MontantTTC", "TotalFactureTTC", "TotalHT", "TotalTVA", "TotalTTCCalcule", "Ecart"
            FormatColonneConso = "#,##0.00"
        Case Else: FormatColonneConso = "@"
    End Select
End Function

' Copie les colonnes d'un tableau vers l'autre (par NOM de colonne).
Private Sub CopierListe(ByVal src As ListObject, ByVal dst As ListObject)

    Dim n As Long, idx As Long
    Dim cd As ListColumn
    Dim vide As Boolean

    vide = (src.DataBodyRange Is Nothing)
    n = IIf(vide, 1, src.ListRows.Count)

    Do While dst.ListRows.Count > 1
        dst.ListRows(dst.ListRows.Count).Delete
    Loop
    If n > 1 Then dst.Resize dst.Range.Resize(n + 1)

    For Each cd In dst.ListColumns
        idx = ColIdx(src, cd.Name)
        If idx = 0 Or vide Then
            cd.DataBodyRange.ClearContents
        Else
            cd.DataBodyRange.Value2 = src.ListColumns(idx).DataBodyRange.Value2
        End If
    Next cd

End Sub

Private Sub MettreEnFormeCorps(ByVal lo As ListObject, ByVal affichage As Boolean)

    Dim cd As ListColumn

    For Each cd In lo.ListColumns
        cd.DataBodyRange.NumberFormat = FormatColonneConso(cd.Name)
    Next cd
    If Not affichage Then Exit Sub

    With lo.DataBodyRange
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = mod_InstallCommun.CoulBordureSaisie()
        .Interior.Color = mod_InstallCommun.CoulFondSaisie()
    End With
    On Error Resume Next
    With lo.ListColumns("TypeConso").DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=ListeTypesConso"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    On Error GoTo 0

End Sub

' Liste des factures (lecture seule) : formats, fond gris, bordures.
Public Sub MettreEnFormeListeFactures(ByVal lo As ListObject)

    Dim cd As ListColumn

    For Each cd In lo.ListColumns
        cd.DataBodyRange.NumberFormat = FormatColonneConso(cd.Name)
    Next cd
    mod_InstallCommun.PoserCorpsTableauLecture lo.DataBodyRange

End Sub

' Contrats : données -> écran (à l'ouverture de la liste, et après une saisie par formulaire).
Public Sub RafraichirListeContrats()
    Dim src As ListObject, dst As ListObject
    Set src = TableContrats()
    Set dst = TableAffContrats()
    If src Is Nothing Or dst Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    CopierListe src, dst
    MettreEnFormeCorps dst, True
    Application.ScreenUpdating = True
End Sub

' Contrats : écran -> données (en quittant la liste). Les lignes gardent leur numéro d'ordre.
Public Sub EnregistrerListeContrats()
    Dim src As ListObject, dst As ListObject
    Set src = TableAffContrats()
    Set dst = TableContrats()
    If src Is Nothing Or dst Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    CopierListe src, dst
    MettreEnFormeCorps dst, False
    Application.ScreenUpdating = True
End Sub

' Si la liste des contrats est l'écran actif, elle est écrite dans les données.
Public Sub EnregistrerListesSiOuvertes()
    If ActiveSheet Is Nothing Then Exit Sub
    If ActiveSheet.Name = mod_VarGlobales.NOM_FEUILLE_CONTRATS Then EnregistrerListeContrats
End Sub

' Vrai si les tableaux du suivi existent (sinon le suivi n'est pas installé).
Public Function ConsoInstallee() As Boolean
    ConsoInstallee = mod_Factures.FacturesInstallees() And Not (TableContrats() Is Nothing)
End Function

Public Function ColIdx(ByVal lo As ListObject, ByVal nomCol As String) As Long
    On Error Resume Next
    ColIdx = lo.ListColumns(nomCol).Index
    On Error GoTo 0
End Function

' Un tableau Excel a toujours au moins une ligne : "vide" = une seule ligne dont la
' première colonne (l'identifiant) est vide.
Public Function TableVide(ByVal lo As ListObject) As Boolean
    If lo.DataBodyRange Is Nothing Then
        TableVide = True
    ElseIf lo.ListRows.Count = 1 Then
        TableVide = (mod_DataStructure.CellText(lo.DataBodyRange.Cells(1, 1).Value2) = "")
    End If
End Function

Public Function NouvelleLigne(ByVal lo As ListObject) As ListRow
    If lo.DataBodyRange Is Nothing Then
        Set NouvelleLigne = lo.ListRows.Add
    ElseIf TableVide(lo) Then
        Set NouvelleLigne = lo.ListRows(1)
    Else
        Set NouvelleLigne = lo.ListRows.Add
    End If
End Function

Public Function LireColonne(ByVal lo As ListObject, ByVal nomCol As String) As Variant()
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

Public Function ProchainID(ByVal lo As ListObject, ByVal nomCol As String) As Long
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

' Valeur de cellule -> numéro de jour entier (0 si ce n'est pas une date).
Public Function VersJour(ByVal v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then
        VersJour = Int(CDbl(v))
    ElseIf IsDate(v) Then
        VersJour = Int(CDbl(CDate(v)))
    End If
End Function

' Vrai si la cellule contient quelque chose (autre que des espaces) : permet de distinguer
' un champ laisse vide d'un champ rempli avec une valeur invalide.
Public Function SaisieNonVide(ByVal v As Variant) As Boolean
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If IsError(v) Then
        SaisieNonVide = True
    Else
        SaisieNonVide = (Trim$(CStr(v)) <> "")
    End If
End Function

Public Function TexteDate(ByVal jour As Double) As String
    TexteDate = Format$(CDate(jour), "dd/mm/yyyy")
End Function

' Chargement des contrats en mémoire (c_Ligne = position dans le tableau).
Private Function ChargerContrats() As Boolean

    Dim lo As ListObject
    Dim vT() As Variant, vO() As Variant, vU() As Variant, vD() As Variant, vF() As Variant, vA() As Variant
    Dim i As Long, n As Long
    Dim t As String

    c_Nb = 0
    Set lo = TableContrats()
    If lo Is Nothing Then Exit Function
    ChargerContrats = True
    If TableVide(lo) Then Exit Function

    n = lo.ListRows.Count
    vT = LireColonne(lo, "TypeConso")
    vO = LireColonne(lo, "Operateur")
    vU = LireColonne(lo, "Unite")
    vD = LireColonne(lo, "DateDebut")
    vF = LireColonne(lo, "DateFin")
    vA = LireColonne(lo, "AbonnementMensuel")

    ReDim c_Typ(1 To n)
    ReDim c_Oper(1 To n)
    ReDim c_Unite(1 To n)
    ReDim c_Deb(1 To n)
    ReDim c_Fin(1 To n)
    ReDim c_Abo(1 To n)
    ReDim c_Ligne(1 To n)

    For i = 1 To n
        t = mod_DataStructure.CellText(vT(i))
        If t <> "" Then
            c_Nb = c_Nb + 1
            c_Typ(c_Nb) = t
            c_Oper(c_Nb) = mod_DataStructure.CellText(vO(i))
            c_Unite(c_Nb) = mod_DataStructure.CellText(vU(i))
            c_Deb(c_Nb) = VersJour(vD(i))
            c_Fin(c_Nb) = VersJour(vF(i))
            c_Abo(c_Nb) = mod_DataStructure.ToDouble(vA(i))
            c_Ligne(c_Nb) = i
        End If
    Next i

End Function



' =====================================================================================
' CONTRATS : RECHERCHE PAR TYPE ET PAR DATE
' =====================================================================================
' Indice du contrat du type donné en vigueur le jour donné (0 = aucun).
Private Function ContratPour(ByVal typ As String, ByVal jour As Double) As Long
    Dim i As Long
    For i = 1 To c_Nb
        If StrComp(c_Typ(i), typ, vbTextCompare) = 0 Then
            If jour >= c_Deb(i) Then
                If c_Fin(i) = 0 Or jour <= c_Fin(i) Then
                    ContratPour = i
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Private Function NomOperateur(ByVal typ As String, ByVal jour As Double) As String
    Dim k As Long
    k = ContratPour(typ, jour)
    If k > 0 Then
        NomOperateur = c_Oper(k)
    End If
    If NomOperateur = "" Then NomOperateur = LibSansContrat()
End Function

' Unité à afficher pour un type : celle de la dernière ligne de consommation saisie, sinon celle
' du contrat le plus récent, sinon l'unité par défaut.
Private Function UniteDuType(ByVal typ As String, ByRef t() As TypeLigne, ByVal nb As Long) As String
    Dim i As Long, k As Long
    Dim meilleur As Double
    For i = 1 To nb
        If t(i).Nat = mod_Factures.NAT_CONSO And StrComp(t(i).Typ, typ, vbTextCompare) = 0 And t(i).Unite <> "" Then
            If t(i).Fin >= meilleur Then
                meilleur = t(i).Fin
                UniteDuType = t(i).Unite
            End If
        End If
    Next i
    If UniteDuType <> "" Then Exit Function
    meilleur = 0
    For i = 1 To c_Nb
        If StrComp(c_Typ(i), typ, vbTextCompare) = 0 And c_Unite(i) <> "" Then
            If k = 0 Or c_Deb(i) >= meilleur Then
                k = i
                meilleur = c_Deb(i)
            End If
        End If
    Next i
    If k > 0 Then UniteDuType = c_Unite(k) Else UniteDuType = UniteDefaut(typ)
End Function

' Opérateurs (sans doublon) des contrats du type donné (tous les types si typ est vide) qui ne sont pas
' terminés : sans date de fin, ou dont la fin n'est pas passée. Renvoie leur nombre ; noms() (base 1).
Public Function OperateursEnCours(ByVal typ As String, ByRef noms() As String) As Long
    Dim i As Long, j As Long, n As Long
    Dim deja As Boolean
    ReDim noms(1 To 1)
    ChargerContrats
    For i = 1 To c_Nb
        If (typ = "" Or StrComp(c_Typ(i), typ, vbTextCompare) = 0) And c_Oper(i) <> "" Then
            If c_Fin(i) = 0 Or c_Fin(i) >= CDbl(Date) Then
                deja = False
                For j = 1 To n
                    If StrComp(noms(j), c_Oper(i), vbTextCompare) = 0 Then deja = True
                Next j
                If Not deja Then
                    n = n + 1
                    If n > UBound(noms) Then ReDim Preserve noms(1 To n + 7)
                    noms(n) = c_Oper(i)
                End If
            End If
        End If
    Next i
    OperateursEnCours = n
End Function

' Vrai si un contrat (en cours ou terminé) de ce type existe pour cet opérateur.
Public Function OperateurConnu(ByVal typ As String, ByVal oper As String) As Boolean
    Dim i As Long
    ChargerContrats
    For i = 1 To c_Nb
        If StrComp(c_Typ(i), typ, vbTextCompare) = 0 And StrComp(c_Oper(i), oper, vbTextCompare) = 0 Then
            OperateurConnu = True
            Exit Function
        End If
    Next i
End Function

' Opérateur et unité du contrat du type donné en vigueur le jour donné. Faux si aucun contrat.
Public Function OperateurContrat(ByVal typ As String, ByVal jour As Double, ByRef oper As String, ByRef unite As String) As Boolean
    Dim k As Long
    ChargerContrats
    k = ContratPour(typ, jour)
    If k > 0 Then
        oper = c_Oper(k)
        unite = c_Unite(k)
        OperateurContrat = True
    End If
End Function


' =====================================================================================
' ÉCRAN frm_Consommations : CALCUL, TABLEAU, GRAPHIQUE
' =====================================================================================
Private Function FeuilleConso() As Worksheet
    Set FeuilleConso = mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_CONSO)
End Function

Public Sub ActualiserConsommations()

    On Error GoTo Erreur
    Application.ScreenUpdating = False
    RafraichirConso
    Application.ScreenUpdating = True
    Exit Sub

Erreur:
    Application.ScreenUpdating = True
    MsgBox mod_Display.FR("Erreur pendant le calcul des consommations.") & vbLf & _
           "Erreur " & Err.Number & " : " & Err.Description, vbExclamation, mod_Display.FR("Suivi des consommations")

End Sub

Private Sub RafraichirConso()

    Dim ws As Worksheet
    Dim typ As String, indic As String, indicSerie As String, periode As String, forme As String, repart As String
    Dim annee As Long, mois As Long, idxDeb As Long, nbM As Long
    Dim tous As Boolean, prorata As Boolean
    Dim note As String
    Dim serNoms() As String, nS As Long
    Dim aQ() As Double, aHT() As Double, aCo() As Double, aAb() As Double, aTx() As Double, aHas() As Boolean
    Dim tabT() As Variant, tabH() As Variant
    Dim p As Long, s As Long, ent As Long, idx As Long
    Dim q As Double, ht As Double, co As Double, ab As Double, tx As Double, has As Boolean
    Dim totQ As Double, totHT As Double, totCo As Double, totAb As Double, totTx As Double
    Dim oper As String
    Dim nbLignesPrises As Long
    Dim unite As String, texte As String
    Dim lignes() As TypeLigne, nbL As Long

    Set ws = FeuilleConso()
    If ws Is Nothing Or Not ConsoInstallee() Then
        MsgBox mod_Display.FR("Le suivi des consommations n'est pas install{e2} : lancez d'abord InstallerConsommations."), vbExclamation
        Exit Sub
    End If

    If Not LireParametresConso(ws, typ, indic, periode, annee, mois, forme, repart, idxDeb, nbM) Then Exit Sub
    ChargerContrats
    nbL = mod_Factures.ChargerLignes(lignes)
    MettreAJourListesConso ws, lignes, nbL

    tous = (StrComp(typ, LibTous(), vbTextCompare) = 0)
    prorata = (StrComp(repart, LibRepartProrata(), vbTextCompare) = 0)

    ' Combinaisons sans sens : pour "tous les types", pas de quantité ni de prix (unités
    ' différentes) et pas de courbe du prix.
    indicSerie = indic
    If tous Then
        If StrComp(indic, LibIndQuantite(), vbTextCompare) = 0 Or StrComp(indic, LibIndPrix(), vbTextCompare) = 0 Then
            indic = LibIndMontant()
            indicSerie = indic
            note = " Pour tous les types, seuls les montants sont comparables : le montant total est affich{e2}."
        End If
        If StrComp(forme, LibFormeMixte(), vbTextCompare) = 0 Then
            forme = LibFormeBarres()
            note = note & " La courbe du prix n'existe que pour un type choisi."
        End If
    ElseIf StrComp(forme, LibFormeMixte(), vbTextCompare) = 0 And StrComp(indic, LibIndPrix(), vbTextCompare) = 0 Then
        indicSerie = LibIndQuantite()      ' le prix est déjà la courbe : les barres montrent la quantité
    End If

    CalculerMois lignes, nbL, typ, tous, prorata, idxDeb, nbM, serNoms, nS, aQ, aHT, aCo, aAb, aTx, aHas, nbLignesPrises

    ' ---------------- Tableau mensuel (visible) et données du graphique (masquées) ----------------
    ent = CO_LIGNE_ENTETE
    ws.Range(ws.Cells(ent + 1, 2), ws.Cells(ent + CO_NB_LIGNES_MAX + 1, 1 + CO_NB_COL_TABLE)).ClearContents
    ws.Range(ws.Cells(ent + 1, 2), ws.Cells(ent + CO_NB_LIGNES_MAX + 1, 1 + CO_NB_COL_TABLE)).Font.Bold = False
    ws.Range(ws.Cells(ent, CO_COL_H_LIB), ws.Cells(ent + CO_NB_LIGNES_MAX, CO_COL_H_SERIE + CO_NB_SERIES - 1)).ClearContents

    ReDim tabT(1 To nbM + 1, 1 To CO_NB_COL_TABLE)
    ReDim tabH(1 To nbM, 1 To 2 + CO_NB_SERIES)

    For p = 1 To nbM
        idx = idxDeb + p - 1
        tabT(p, 1) = DateSerial(idx \ 12, (idx Mod 12) + 1, 1)
        tabH(p, 1) = tabT(p, 1)
        q = 0: ht = 0: co = 0: ab = 0: tx = 0: oper = ""
        For s = 1 To nS
            q = q + aQ(p, s)
            ht = ht + aHT(p, s)
            co = co + aCo(p, s)
            ab = ab + aAb(p, s)
            tx = tx + aTx(p, s)
            If aHas(p, s) And Not tous Then
                If oper <> "" Then oper = oper & " / "
                oper = oper & serNoms(s)
            End If
            tabH(p, 2 + s) = ValeurIndic(indicSerie, aQ(p, s), aHT(p, s), aCo(p, s), aAb(p, s), aTx(p, s), aHas(p, s))
        Next s
        has = (q <> 0 Or ht <> 0 Or co <> 0 Or ab <> 0 Or tx <> 0)
        If has Then
            If Not tous Then
                tabT(p, 2) = q
                If q > 0 Then
                    tabT(p, 3) = ht / q
                    tabH(p, 2) = ht / q
                End If
            End If
            tabT(p, 4) = co
            tabT(p, 5) = ab
            tabT(p, 6) = tx
            tabT(p, 7) = co + ab + tx
            tabT(p, 8) = oper
        End If
        totQ = totQ + q
        totHT = totHT + ht
        totCo = totCo + co
        totAb = totAb + ab
        totTx = totTx + tx
    Next p

    ' Ligne de total, juste sous le dernier mois.
    tabT(nbM + 1, 1) = "Total"
    If Not tous Then
        tabT(nbM + 1, 2) = totQ
        If totQ > 0 Then tabT(nbM + 1, 3) = totHT / totQ
    End If
    tabT(nbM + 1, 4) = totCo
    tabT(nbM + 1, 5) = totAb
    tabT(nbM + 1, 6) = totTx
    tabT(nbM + 1, 7) = totCo + totAb + totTx

    ' Colonne des mois au format "mmm yyyy" partout (le mot "Total" s'y affiche tel quel).
    ws.Range(ws.Cells(ent + 1, 2), ws.Cells(ent + CO_NB_LIGNES_MAX + 1, 2)).NumberFormat = "mmm yyyy"
    ws.Range(ws.Cells(ent + 1, 2), ws.Cells(ent + nbM + 1, 1 + CO_NB_COL_TABLE)).Value = tabT
    ws.Range(ws.Cells(ent + nbM + 1, 2), ws.Cells(ent + nbM + 1, 1 + CO_NB_COL_TABLE)).Font.Bold = True

    ' Données du graphique : en-têtes, puis une ligne par mois.
    ws.Cells(ent, CO_COL_H_LIB).Value = "Mois"
    ws.Cells(ent, CO_COL_H_PRIX).Value = "Prix unitaire HT"
    For s = 1 To nS
        ws.Cells(ent, CO_COL_H_SERIE + s - 1).Value = serNoms(s)
    Next s
    ws.Range(ws.Cells(ent + 1, CO_COL_H_LIB), ws.Cells(ent + nbM, CO_COL_H_SERIE + CO_NB_SERIES - 1)).Value = tabH
    ws.Range(ws.Cells(ent + 1, CO_COL_H_LIB), ws.Cells(ent + nbM, CO_COL_H_LIB)).NumberFormat = "mmm yyyy"

    ' Libellé de la colonne Quantité / Prix avec l'unité (ou rien pour "tous").
    If tous Then
        ws.Cells(ent, 3).Value = mod_Display.FR("Quantit{e2}")
        ws.Cells(ent, 4).Value = "Prix unitaire HT"
        unite = ""
    Else
        unite = UniteDuType(typ, lignes, nbL)
        ws.Cells(ent, 3).Value = mod_Display.FR("Quantit{e2} (") & unite & ")"
        ws.Cells(ent, 4).Value = "Prix HT (" & SymboleEuro() & "/" & unite & ")"
    End If

    ' Résumé.
    texte = IIf(tous, mod_Display.FR("Tous les types"), typ) & " -- " & _
            Format$(DateSerial(idxDeb \ 12, (idxDeb Mod 12) + 1, 1), "mmm yyyy") & mod_Display.FR(" {a2} ") & _
            Format$(DateSerial((idxDeb + nbM - 1) \ 12, ((idxDeb + nbM - 1) Mod 12) + 1, 1), "mmm yyyy") & _
            " -- " & nbLignesPrises & mod_Display.FR(" ligne(s) de facture prise(s) en compte")
    If nbLignesPrises = 0 Then texte = texte & mod_Display.FR(". Aucune facture sur cette p{e2}riode.")
    ws.Range(CO_ADR_MESSAGE).Value = texte & mod_Display.FR(note)

    ConstruireGraphiqueConso ws, nbM, nS, serNoms, indic, indicSerie, forme, typ, tous, unite

End Sub

' Lit et contrôle les paramètres de l'écran. Renvoie Faux (avec un message) s'ils sont incorrects.
Private Function LireParametresConso(ByVal ws As Worksheet, ByRef typ As String, ByRef indic As String, _
                                     ByRef periode As String, ByRef annee As Long, ByRef mois As Long, _
                                     ByRef forme As String, ByRef repart As String, _
                                     ByRef idxDeb As Long, ByRef nbM As Long) As Boolean

    Dim idxFin As Long

    typ = mod_DataStructure.CellText(ws.Range(CO_ADR_TYPE).Value2)
    indic = mod_DataStructure.CellText(ws.Range(CO_ADR_INDIC).Value2)
    periode = mod_DataStructure.CellText(ws.Range(CO_ADR_PERIODE).Value2)
    annee = mod_DataStructure.ToLong(ws.Range(CO_ADR_ANNEE).Value2)
    mois = mod_DataStructure.ToLong(ws.Range(CO_ADR_MOIS).Value2)
    forme = mod_DataStructure.CellText(ws.Range(CO_ADR_FORME).Value2)
    repart = mod_DataStructure.CellText(ws.Range(CO_ADR_REPART).Value2)

    If typ = "" Then typ = LibTous()
    If indic = "" Then indic = LibIndMontant()
    If forme = "" Then forme = LibFormeBarres()
    If repart = "" Then repart = LibRepartProrata()

    If annee < 1900 Or annee > 2200 Then
        MsgBox mod_Display.FR("L'ann{e2}e de r{e2}f{e2}rence est incorrecte."), vbExclamation
        Exit Function
    End If
    If mois < 1 Or mois > 12 Then
        MsgBox mod_Display.FR("Le mois de r{e2}f{e2}rence est incorrect (1 {a2} 12)."), vbExclamation
        Exit Function
    End If

    idxFin = annee * 12 + mois - 1
    If StrComp(periode, LibPeriodeAnnee(), vbTextCompare) = 0 Then
        idxDeb = annee * 12
        nbM = 12
    ElseIf StrComp(periode, LibPeriode36(), vbTextCompare) = 0 Then
        idxDeb = idxFin - 35
        nbM = 36
    ElseIf StrComp(periode, LibPeriode60(), vbTextCompare) = 0 Then
        idxDeb = idxFin - 59
        nbM = 60
    ElseIf StrComp(periode, LibPeriode120(), vbTextCompare) = 0 Then
        idxDeb = idxFin - 119
        nbM = 120
    Else
        idxDeb = idxFin - 11
        nbM = 12
    End If

    LireParametresConso = True

End Function

' Répartit les lignes de factures sur les mois de la période. Pour chaque (mois, série) :
'   aQ = quantité consommée, aHT = montant HT de consommation, aCo = consommation TTC,
'   aAb = abonnement TTC, aTx = taxes TTC.
' Série = opérateur (un seul type choisi) ou type (tous les types).
Private Sub CalculerMois(ByRef t() As TypeLigne, ByVal nb As Long, ByVal typ As String, ByVal tous As Boolean, _
                         ByVal prorata As Boolean, ByVal idxDeb As Long, ByVal nbM As Long, _
                         ByRef serNoms() As String, ByRef nS As Long, _
                         ByRef aQ() As Double, ByRef aHT() As Double, ByRef aCo() As Double, _
                         ByRef aAb() As Double, ByRef aTx() As Double, _
                         ByRef aHas() As Boolean, ByRef nbPris As Long)

    Dim dic As Object
    Dim j As Long, i As Long, s As Long, p As Long
    Dim cle As String
    Dim iDeb As Long, iFin As Long, y As Long, m As Long
    Dim dMois As Double, fMois As Double, a As Double, b As Double
    Dim frac As Double, totalJours As Double

    ReDim serNoms(1 To CO_NB_SERIES)
    ReDim aQ(1 To nbM, 1 To CO_NB_SERIES)
    ReDim aHT(1 To nbM, 1 To CO_NB_SERIES)
    ReDim aCo(1 To nbM, 1 To CO_NB_SERIES)
    ReDim aAb(1 To nbM, 1 To CO_NB_SERIES)
    ReDim aTx(1 To nbM, 1 To CO_NB_SERIES)
    ReDim aHas(1 To nbM, 1 To CO_NB_SERIES)
    nS = 0
    nbPris = 0

    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = 1     ' sans tenir compte des majuscules

    For j = 1 To nb
        If tous Or StrComp(t(j).Typ, typ, vbTextCompare) = 0 Then

            If tous Then
                cle = t(j).Typ
            Else
                cle = t(j).Operateur
                If cle = "" Then cle = LibSansContrat()
            End If

            ' Série de cette ligne (au-delà de CO_NB_SERIES, tout le reste est regroupé dans "Autres").
            If dic.Exists(cle) Then
                s = dic(cle)
            ElseIf nS < CO_NB_SERIES Then
                nS = nS + 1
                serNoms(nS) = cle
                dic.Add cle, nS
                s = nS
            Else
                s = CO_NB_SERIES
                serNoms(s) = LibAutres()
                dic.Add cle, s
            End If

            iDeb = Year(CDate(t(j).Debut)) * 12 + Month(CDate(t(j).Debut)) - 1
            iFin = Year(CDate(t(j).Fin)) * 12 + Month(CDate(t(j).Fin)) - 1

            If prorata Then
                totalJours = t(j).Fin - t(j).Debut + 1
                For i = iDeb To iFin
                    p = i - idxDeb + 1
                    If p >= 1 And p <= nbM Then
                        y = i \ 12
                        m = (i Mod 12) + 1
                        dMois = CDbl(DateSerial(y, m, 1))
                        fMois = CDbl(DateSerial(y, m + 1, 0))
                        a = t(j).Debut
                        If dMois > a Then a = dMois
                        b = t(j).Fin
                        If fMois < b Then b = fMois
                        frac = (b - a + 1) / totalJours
                        Ajouter t(j), frac, p, s, aQ, aHT, aCo, aAb, aTx
                        aHas(p, s) = True
                    End If
                Next i
                If iFin >= idxDeb And iDeb <= idxDeb + nbM - 1 Then nbPris = nbPris + 1
            Else
                p = iFin - idxDeb + 1
                If p >= 1 And p <= nbM Then
                    Ajouter t(j), 1, p, s, aQ, aHT, aCo, aAb, aTx
                    aHas(p, s) = True
                    nbPris = nbPris + 1
                End If
            End If
        End If
    Next j

End Sub

' Ajoute la fraction "frac" d'une ligne de facture au mois p, série s, dans la bonne catégorie.
Private Sub Ajouter(ByRef l As TypeLigne, ByVal frac As Double, ByVal p As Long, ByVal s As Long, _
                    ByRef aQ() As Double, ByRef aHT() As Double, ByRef aCo() As Double, _
                    ByRef aAb() As Double, ByRef aTx() As Double)
    Select Case l.Nat
        Case mod_Factures.NAT_CONSO
            aQ(p, s) = aQ(p, s) + l.Qte * frac
            aHT(p, s) = aHT(p, s) + l.HT * frac
            aCo(p, s) = aCo(p, s) + l.TTC * frac
        Case mod_Factures.NAT_ABO
            aAb(p, s) = aAb(p, s) + l.TTC * frac
        Case Else
            aTx(p, s) = aTx(p, s) + l.TTC * frac
    End Select
End Sub

Private Function ValeurIndic(ByVal indic As String, ByVal q As Double, ByVal ht As Double, ByVal co As Double, _
                             ByVal ab As Double, ByVal tx As Double, ByVal has As Boolean) As Variant

    If Not has Then Exit Function       ' Empty : le graphique laisse un trou

    If StrComp(indic, LibIndQuantite(), vbTextCompare) = 0 Then
        ValeurIndic = q
    ElseIf StrComp(indic, LibIndPrix(), vbTextCompare) = 0 Then
        If q > 0 Then ValeurIndic = ht / q
    ElseIf StrComp(indic, LibIndConso(), vbTextCompare) = 0 Then
        ValeurIndic = co
    ElseIf StrComp(indic, LibIndAbo(), vbTextCompare) = 0 Then
        ValeurIndic = ab
    ElseIf StrComp(indic, LibIndTaxes(), vbTextCompare) = 0 Then
        ValeurIndic = tx
    Else
        ValeurIndic = co + ab + tx
    End If

End Function

' Années proposées : de la plus ancienne année de facture (ou de l'année en cours) à l'année suivante.
Private Sub MettreAJourListesConso(ByVal ws As Worksheet, ByRef t() As TypeLigne, ByVal nb As Long)

    Dim i As Long, anMin As Long, anMax As Long, an As Long

    anMin = Year(Date)
    anMax = Year(Date) + 1
    For i = 1 To nb
        an = Year(CDate(t(i).Debut))
        If an < anMin Then anMin = an
        an = Year(CDate(t(i).Fin))
        If an + 1 > anMax Then anMax = an + 1
    Next i

    ' Périodes proposées (liste reposée à chaque actualisation : un classeur déjà installé
    ' reçoit ainsi les nouvelles périodes sans réinstallation).
    ws.Range(ws.Cells(2, CO_COL_LISTE_PERIODE), ws.Cells(10, CO_COL_LISTE_PERIODE)).ClearContents
    ws.Cells(2, CO_COL_LISTE_PERIODE).Value = LibPeriode12()
    ws.Cells(3, CO_COL_LISTE_PERIODE).Value = LibPeriode36()
    ws.Cells(4, CO_COL_LISTE_PERIODE).Value = LibPeriode60()
    ws.Cells(5, CO_COL_LISTE_PERIODE).Value = LibPeriode120()
    ws.Cells(6, CO_COL_LISTE_PERIODE).Value = LibPeriodeAnnee()
    ThisWorkbook.Names.Add Name:="ListePeriodeConso", RefersTo:="='" & ws.Name & "'!" & _
        ws.Range(ws.Cells(2, CO_COL_LISTE_PERIODE), ws.Cells(6, CO_COL_LISTE_PERIODE)).Address(True, True)

    ' Mise en forme de la zone du tableau jusqu'à CO_NB_LIGNES_MAX lignes.
    mod_InstallCommun.PoserCorpsTableauLecture ws.Range(ws.Cells(CO_LIGNE_ENTETE + 1, 2), ws.Cells(CO_LIGNE_ENTETE + CO_NB_LIGNES_MAX + 1, 1 + CO_NB_COL_TABLE))
    ws.Range(ws.Cells(CO_LIGNE_ENTETE + 1, 2), ws.Cells(CO_LIGNE_ENTETE + CO_NB_LIGNES_MAX + 1, 2)).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(CO_LIGNE_ENTETE + 1, 3), ws.Cells(CO_LIGNE_ENTETE + CO_NB_LIGNES_MAX + 1, 3)).NumberFormat = "#,##0.00"
    ws.Range(ws.Cells(CO_LIGNE_ENTETE + 1, 4), ws.Cells(CO_LIGNE_ENTETE + CO_NB_LIGNES_MAX + 1, 4)).NumberFormat = "#,##0.00000"
    ws.Range(ws.Cells(CO_LIGNE_ENTETE + 1, 5), ws.Cells(CO_LIGNE_ENTETE + CO_NB_LIGNES_MAX + 1, 8)).NumberFormat = "#,##0.00"

    ws.Range(ws.Cells(2, CO_COL_LISTE_ANNEE), ws.Cells(60, CO_COL_LISTE_ANNEE)).ClearContents
    For an = anMin To anMax
        If an - anMin + 2 <= 60 Then ws.Cells(an - anMin + 2, CO_COL_LISTE_ANNEE).Value = an
    Next an

End Sub

Private Sub ConstruireGraphiqueConso(ByVal ws As Worksheet, ByVal nbM As Long, ByVal nS As Long, _
                                     ByRef serNoms() As String, ByVal indic As String, ByVal indicSerie As String, _
                                     ByVal forme As String, ByVal typ As String, ByVal tous As Boolean, _
                                     ByVal unite As String)

    Dim co As ChartObject
    Dim se As Series
    Dim s As Long, ent As Long
    Dim courbes As Boolean, mixte As Boolean
    Dim xr As Range

    On Error Resume Next
    Set co = ws.ChartObjects(CO_GRAPH)
    On Error GoTo 0
    If co Is Nothing Then Exit Sub

    Do While co.Chart.SeriesCollection.Count > 0
        co.Chart.SeriesCollection(1).Delete
    Loop

    co.Chart.HasTitle = True
    If nS = 0 Then
        co.Chart.ChartType = xlColumnStacked
        co.Chart.ChartTitle.Text = mod_Display.FR("Aucune facture sur la p{e2}riode")
        Exit Sub
    End If

    ent = CO_LIGNE_ENTETE
    Set xr = ws.Range(ws.Cells(ent + 1, CO_COL_H_LIB), ws.Cells(ent + nbM, CO_COL_H_LIB))
    courbes = (StrComp(forme, LibFormeCourbes(), vbTextCompare) = 0)
    mixte = (StrComp(forme, LibFormeMixte(), vbTextCompare) = 0)

    If courbes Then co.Chart.ChartType = xlLineMarkers Else co.Chart.ChartType = xlColumnStacked

    For s = 1 To nS
        Set se = co.Chart.SeriesCollection.NewSeries
        se.Name = serNoms(s)
        se.Values = ws.Range(ws.Cells(ent + 1, CO_COL_H_SERIE + s - 1), ws.Cells(ent + nbM, CO_COL_H_SERIE + s - 1))
        se.XValues = xr
        If courbes Then se.ChartType = xlLineMarkers Else se.ChartType = xlColumnStacked
    Next s

    If mixte Then
        Set se = co.Chart.SeriesCollection.NewSeries
        se.Name = "Prix unitaire"
        se.Values = ws.Range(ws.Cells(ent + 1, CO_COL_H_PRIX), ws.Cells(ent + nbM, CO_COL_H_PRIX))
        se.XValues = xr
        se.ChartType = xlLineMarkers
        se.AxisGroup = xlSecondary
    End If

    co.Chart.HasLegend = True
    co.Chart.Legend.Position = xlLegendPositionBottom
    co.Chart.DisplayBlanksAs = xlNotPlotted

    ' Axe des mois : étiquettes textuelles (un graduation par mois), pas d'axe temporel.
    On Error Resume Next
    co.Chart.Axes(xlCategory).CategoryType = xlCategoryScale
    co.Chart.Axes(xlCategory).TickLabels.NumberFormat = "mmm yy"
    If nbM <= 12 Then
        co.Chart.Axes(xlCategory).TickLabelSpacing = 1
    ElseIf nbM <= 36 Then
        co.Chart.Axes(xlCategory).TickLabelSpacing = 3
    ElseIf nbM <= 60 Then
        co.Chart.Axes(xlCategory).TickLabelSpacing = 6
    Else
        co.Chart.Axes(xlCategory).TickLabelSpacing = 12
    End If

    ' Titres des axes.
    co.Chart.Axes(xlValue, xlPrimary).HasTitle = True
    co.Chart.Axes(xlValue, xlPrimary).AxisTitle.Text = TitreAxe(indicSerie, unite)
    If mixte Then
        co.Chart.Axes(xlValue, xlSecondary).HasTitle = True
        co.Chart.Axes(xlValue, xlSecondary).AxisTitle.Text = TitreAxe(LibIndPrix(), unite)
    End If
    On Error GoTo 0

    If tous Then
        co.Chart.ChartTitle.Text = mod_Display.FR("Tous les types -- ") & indic
    Else
        co.Chart.ChartTitle.Text = typ & " -- " & indic
        If mixte Then co.Chart.ChartTitle.Text = typ & mod_Display.FR(" -- ") & indicSerie & " et prix unitaire"
    End If

End Sub

Private Function TitreAxe(ByVal indic As String, ByVal unite As String) As String
    If StrComp(indic, LibIndQuantite(), vbTextCompare) = 0 Then
        TitreAxe = indic & IIf(unite <> "", " (" & unite & ")", "")
    ElseIf StrComp(indic, LibIndPrix(), vbTextCompare) = 0 Then
        TitreAxe = indic & " (" & SymboleEuro() & IIf(unite <> "", "/" & unite, "") & ")"
    Else
        TitreAxe = indic & " (" & SymboleEuro() & ")"
    End If
End Function


' =====================================================================================
' NAVIGATION ENTRE LES ÉCRANS
' =====================================================================================
' Affiche un écran du suivi (graphique, factures ou contrats) et masque les deux autres
' ainsi que la feuille Synthese.
Private Sub MontrerEcran(ByVal nomFeuille As String)

    Dim ws As Worksheet, autre As Worksheet
    Dim noms As Variant
    Dim i As Long

    Set ws = mod_InstallCommun.TrouverFeuille(nomFeuille)
    If ws Is Nothing Then
        MsgBox mod_Display.FR("Le suivi des consommations n'est pas install{e2} : lancez d'abord InstallerConsommations (Ctrl+G)."), vbExclamation
        Exit Sub
    End If

    EnregistrerListesSiOuvertes
    If nomFeuille = mod_VarGlobales.NOM_FEUILLE_FACTURES Then mod_Factures.RafraichirListeFactures
    If nomFeuille = mod_VarGlobales.NOM_FEUILLE_CONTRATS Then RafraichirListeContrats

    ws.Visible = xlSheetVisible
    ws.Activate
    mod_InstallCommun.MasquerQuadrillage

    noms = Array(mod_VarGlobales.NOM_FEUILLE_CONSO, mod_VarGlobales.NOM_FEUILLE_FACTURES, mod_VarGlobales.NOM_FEUILLE_CONTRATS)
    For i = LBound(noms) To UBound(noms)
        If noms(i) <> nomFeuille Then
            Set autre = mod_InstallCommun.TrouverFeuille(CStr(noms(i)))
            If Not autre Is Nothing Then mod_Aide.MasquerFormulaire autre
        End If
    Next i

    ' Synthese est masquee par MasquerQuadrillage (regle commune des formulaires).

End Sub

' Ouverture depuis la feuille Synthese.
Public Sub AfficherConsommations()
    MontrerEcran mod_VarGlobales.NOM_FEUILLE_CONSO
    If Not FeuilleConso() Is Nothing Then ActualiserConsommations
End Sub

' Retour sur l'écran graphique depuis une liste : il est recalculé.
Public Sub RetourConsommations()
    MontrerEcran mod_VarGlobales.NOM_FEUILLE_CONSO
    If Not FeuilleConso() Is Nothing Then ActualiserConsommations
End Sub

Public Sub AfficherFactures()
    MontrerEcran mod_VarGlobales.NOM_FEUILLE_FACTURES
End Sub

Public Sub AfficherContrats()
    MontrerEcran mod_VarGlobales.NOM_FEUILLE_CONTRATS
End Sub

' Sortie vers Synthese (bouton Sortir de l'écran graphique).
Public Sub SortirConsommations()

    Dim ws As Worksheet
    Dim noms As Variant
    Dim i As Long

    EnregistrerListesSiOuvertes
    mod_InstallCommun.AfficherSynthese

    noms = Array(mod_VarGlobales.NOM_FEUILLE_CONSO, mod_VarGlobales.NOM_FEUILLE_FACTURES, mod_VarGlobales.NOM_FEUILLE_CONTRATS)
    For i = LBound(noms) To UBound(noms)
        Set ws = mod_InstallCommun.TrouverFeuille(CStr(noms(i)))
        If Not ws Is Nothing Then mod_Aide.MasquerFormulaire ws
    Next i

End Sub


' =====================================================================================
' SAISIE D'UN CONTRAT (formulaire frm_Contrat)
' =====================================================================================
' Ajoute un contrat. Si un contrat du même type est en cours et commence avant le
' nouveau, il est proposé de le clôturer la veille (changement d'opérateur).
Public Sub AjouterContrat()
    EnregistrerListesSiOuvertes
    AjouterContratCore
    RafraichirListeContrats
End Sub

Private Sub AjouterContratCore()

    Dim c As TypeContrat
    Dim ligneAClore As Long

    If Not ConsoInstallee() Then
        MsgBox mod_Display.FR("Le suivi des consommations n'est pas install{e2}. Ex{e2}cutez d'abord la macro InstallerConsommations."), vbExclamation
        Exit Sub
    End If

    If OuvrirFormContrat(c, ligneAClore) Then
        If EnregistrerContrat(c, ligneAClore) Then
            If ActiveSheet.Name = mod_VarGlobales.NOM_FEUILLE_CONSO Then ActualiserConsommations
            MsgBox mod_Display.FR("Contrat enregistr{e2}."), vbInformation, mod_Display.FR("Suivi des consommations")
        End If
    End If

End Sub

Private Function EnregistrerContrat(ByRef c As TypeContrat, ByVal ligneAClore As Long) As Boolean

    Dim lo As ListObject
    Dim lr As ListRow

    Set lo = TableContrats()
    If lo Is Nothing Then Exit Function

    ' Clôture du contrat en cours (la veille du nouveau).
    If ligneAClore > 0 Then
        lo.ListRows(ligneAClore).Range.Cells(1, ColIdx(lo, "DateFin")).Value = c.Debut - 1
    End If

    Set lr = NouvelleLigne(lo)
    With lr.Range
        .Cells(1, ColIdx(lo, "ID_Contrat")).Value = ProchainID(lo, "ID_Contrat")
        .Cells(1, ColIdx(lo, "TypeConso")).Value = c.Typ
        .Cells(1, ColIdx(lo, "Operateur")).Value = c.Operateur
        .Cells(1, ColIdx(lo, "Unite")).Value = c.Unite
        .Cells(1, ColIdx(lo, "DateDebut")).Value = c.Debut
        If c.Fin <> 0 Then .Cells(1, ColIdx(lo, "DateFin")).Value = c.Fin
        .Cells(1, ColIdx(lo, "AbonnementMensuel")).Value = c.AboMensuel
        .Cells(1, ColIdx(lo, "Reference")).Value = c.Reference
    End With

    EnregistrerContrat = True

End Function

Private Function OuvrirFormContrat(ByRef c As TypeContrat, ByRef ligneAClore As Long) As Boolean

    Dim ws As Worksheet
    Dim evenementsAvant As Boolean

    On Error GoTo Erreur

    Set ws = mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_CONTRAT)
    If ws Is Nothing Then
        MsgBox mod_Display.FR("La feuille '") & mod_VarGlobales.NOM_FEUILLE_CONTRAT & mod_Display.FR("' est introuvable.") & vbCrLf & _
               mod_Display.FR("Ex{e2}cutez d'abord la macro InstallerConsommations."), vbExclamation
        Exit Function
    End If

    g_CtrNomFeuillePrec = ActiveSheet.Name
    g_CtrValide = False
    g_CtrLigneAClore = 0

    evenementsAvant = Application.EnableEvents
    Application.EnableEvents = False
    ws.Range(CF_ADR_TYPE).Value = c.Typ
    ws.Range(CF_ADR_OPERATEUR).Value = c.Operateur
    ws.Range(CF_ADR_UNITE).Value = c.Unite
    If c.Debut <> 0 Then ws.Range(CF_ADR_DEBUT).Value = c.Debut Else ws.Range(CF_ADR_DEBUT).ClearContents
    If c.Fin <> 0 Then ws.Range(CF_ADR_FIN).Value = c.Fin Else ws.Range(CF_ADR_FIN).ClearContents
    If c.AboMensuel <> 0 Then ws.Range(CF_ADR_ABO).Value = c.AboMensuel Else ws.Range(CF_ADR_ABO).ClearContents
    ws.Range(CF_ADR_REFERENCE).Value = c.Reference
    ws.Range(CF_ADR_MESSAGE).Value = ""
    mod_InstallCommun.PoserValidationDate ws.Range(CF_ADR_DEBUT)
    mod_InstallCommun.PoserValidationDate ws.Range(CF_ADR_FIN)
    mod_InstallCommun.PoserValidationMontant ws.Range(CF_ADR_ABO), mod_InstallCommun.MONTANT_POSITIF_OU_NUL
    Application.EnableEvents = evenementsAvant

    ws.Visible = xlSheetVisible
    ws.Activate
    mod_InstallCommun.MasquerQuadrillage
    ws.Range(CF_ADR_TYPE).Select

    On Error Resume Next
    mod_InstallCommun.PoserListeDeroulante ws.Range(CF_ADR_TYPE), "ListeTypesConso", True
    On Error GoTo Erreur

    Application.EnableEvents = True
    g_CtrEnCours = True
    g_CtrVerrouActif = True
    Do While g_CtrEnCours
        DoEvents
    Loop

    On Error Resume Next
    mod_InstallCommun.ActiverFeuille g_CtrNomFeuillePrec
    On Error GoTo 0

    If g_CtrValide Then
        c = g_CtrRes
        ligneAClore = g_CtrLigneAClore
        OuvrirFormContrat = True
    End If
    Exit Function

Erreur:
    Application.EnableEvents = True
    g_CtrEnCours = False
    g_CtrVerrouActif = False
    MsgBox mod_Display.FR("Erreur inattendue dans la saisie du contrat :") & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
    OuvrirFormContrat = False

End Function

' Appelée par le code de la feuille (Worksheet_Change) : le choix d'un type propose son unité.
Public Sub CtrTraiterChangement(ByVal ws As Worksheet, ByVal Target As Range)

    Dim numErr As Long
    Dim u As String

    If Not g_CtrEnCours Then Exit Sub
    If Target.Cells.Count > 1 Then Exit Sub

    On Error GoTo Sortie

    If Target.Address(False, False) = CF_ADR_TYPE Then
        u = mod_DataStructure.CellText(ws.Range(CF_ADR_UNITE).Value)
        If u = "" Or EstUniteDefaut(u) Then
            Application.EnableEvents = False
            ws.Range(CF_ADR_UNITE).Value = UniteDefaut(mod_DataStructure.CellText(Target.Value))
        End If
    End If

Sortie:
    numErr = Err.Number
    Application.EnableEvents = True
    If numErr <> 0 Then MsgBox mod_Display.FR("Erreur lors du changement de type : ") & Err.Description, vbExclamation

End Sub

Public Sub CtrVerrouiller(ByVal ws As Worksheet)
    If g_CtrVerrouActif And Not mod_Aide.AideOuverte() Then ws.Activate
End Sub

Public Sub CtrValider()

    Dim ws As Worksheet
    Dim typ As String, oper As String
    Dim dDeb As Double, dFin As Double, jFin As Double
    Dim vAbo As Variant
    Dim nAbo As Double
    Dim dTmp As Date
    Dim j As Long, aClore As Long
    Dim e As TypeContrat

    If Not g_CtrEnCours Then Exit Sub
    On Error GoTo Erreur
    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_CONTRAT)

    typ = mod_DataStructure.CellText(ws.Range(CF_ADR_TYPE).Value)
    oper = mod_DataStructure.CellText(ws.Range(CF_ADR_OPERATEUR).Value)
    dDeb = VersJour(ws.Range(CF_ADR_DEBUT).Value)
    dFin = VersJour(ws.Range(CF_ADR_FIN).Value)
    vAbo = ws.Range(CF_ADR_ABO).Value

    If typ = "" Or Not TypeConnu(typ) Then
        AfficherMessageCtr ws, "Choisissez le type de consommation dans la liste."
        Exit Sub
    End If
    If oper = "" Then
        AfficherMessageCtr ws, "Renseignez l'op{e2}rateur (fournisseur)."
        Exit Sub
    End If
    If SaisieNonVide(ws.Range(CF_ADR_DEBUT).Value) And Not mod_DataStructure.TryLireDate(ws.Range(CF_ADR_DEBUT).Value, dTmp) Then
        AfficherMessageCtr ws, "La date de d{e2}but n'est pas valide. " & mod_DataStructure.MessageDateAttendue()
        Exit Sub
    End If
    If SaisieNonVide(ws.Range(CF_ADR_FIN).Value) And Not mod_DataStructure.TryLireDate(ws.Range(CF_ADR_FIN).Value, dTmp) Then
        AfficherMessageCtr ws, "La date de fin n'est pas valide. " & mod_DataStructure.MessageDateAttendue()
        Exit Sub
    End If
    If dDeb = 0 Then
        AfficherMessageCtr ws, "Renseignez la date de d{e2}but du contrat."
        Exit Sub
    End If
    If dFin <> 0 And dFin < dDeb Then
        AfficherMessageCtr ws, "La date de fin est ant{e2}rieure {a2} la date de d{e2}but."
        Exit Sub
    End If
    If Not SaisieNonVide(vAbo) Then
        nAbo = 0
    ElseIf Not mod_DataStructure.TryLireMontant(vAbo, nAbo) Then
        AfficherMessageCtr ws, "L'abonnement mensuel n'est pas valide. " & mod_DataStructure.MessageMontantAttendu()
        Exit Sub
    ElseIf nAbo < 0 Then
        AfficherMessageCtr ws, "L'abonnement mensuel ne peut pas {ea}tre n{e2}gatif."
        Exit Sub
    End If

    ChargerContrats

    ' Un contrat du même type en cours (sans date de fin) qui commence avant : on propose de le clôturer.
    For j = 1 To c_Nb
        If StrComp(c_Typ(j), typ, vbTextCompare) = 0 And c_Fin(j) = 0 And c_Deb(j) < dDeb Then
            If MsgBox(mod_Display.FR("Le contrat de ") & c_Oper(j) & mod_Display.FR(" (depuis le ") & TexteDate(c_Deb(j)) & _
                      mod_Display.FR(") est en cours.") & vbCrLf & vbCrLf & _
                      mod_Display.FR("Le cl{o2}turer le ") & TexteDate(dDeb - 1) & mod_Display.FR(" (veille du nouveau contrat) ?"), _
                      vbYesNo + vbQuestion, mod_Display.FR("Changement d'op{e2}rateur")) = vbNo Then
                AfficherMessageCtr ws, "Un seul contrat peut {ea}tre en cours pour un type : cl{o2}turez l'ancien ou donnez une date de fin."
                Exit Sub
            End If
            aClore = c_Ligne(j)
            c_Fin(j) = dDeb - 1           ' pour le contrôle de chevauchement ci-dessous
            Exit For
        End If
    Next j

    ' Chevauchement avec les autres contrats du même type.
    For j = 1 To c_Nb
        If StrComp(c_Typ(j), typ, vbTextCompare) = 0 Then
            jFin = c_Fin(j)
            If jFin = 0 Then jFin = 2958465          ' 31/12/9999
            If dFin = 0 Then
                If jFin >= dDeb Then
                    AfficherMessageCtr ws, "Ce contrat chevauche celui de " & c_Oper(j) & " (du " & TexteDate(c_Deb(j)) & IIf(c_Fin(j) = 0, " en cours", " au " & TexteDate(c_Fin(j))) & ")."
                    Exit Sub
                End If
            ElseIf dDeb <= jFin And dFin >= c_Deb(j) Then
                AfficherMessageCtr ws, "Ce contrat chevauche celui de " & c_Oper(j) & " (du " & TexteDate(c_Deb(j)) & IIf(c_Fin(j) = 0, " en cours", " au " & TexteDate(c_Fin(j))) & ")."
                Exit Sub
            End If
        End If
    Next j

    e.Typ = typ
    e.Operateur = oper
    e.Unite = mod_DataStructure.CellText(ws.Range(CF_ADR_UNITE).Value)
    If e.Unite = "" Then e.Unite = UniteDefaut(typ)
    e.Debut = CDate(dDeb)
    If dFin <> 0 Then e.Fin = CDate(dFin)
    e.AboMensuel = nAbo
    e.Reference = mod_DataStructure.CellText(ws.Range(CF_ADR_REFERENCE).Value)

    g_CtrRes = e
    g_CtrLigneAClore = aClore
    g_CtrValide = True
    FermerFeuilleCtr ws
    Exit Sub

Erreur:
    Application.EnableEvents = True
    MsgBox mod_Display.FR("Erreur dans CtrValider : ") & Err.Number & " - " & Err.Description, vbCritical

End Sub

Public Sub CtrAnnuler()

    Dim ws As Worksheet

    If Not g_CtrEnCours Then Exit Sub
    On Error GoTo Erreur

    If MsgBox(mod_Display.FR("Annuler la saisie ? Rien ne sera enregistr{e2}."), vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_CONTRAT)
    g_CtrValide = False
    FermerFeuilleCtr ws
    Exit Sub

Erreur:
    Application.EnableEvents = True
    MsgBox mod_Display.FR("Erreur dans CtrAnnuler : ") & Err.Number & " - " & Err.Description, vbCritical

End Sub

Private Sub AfficherMessageCtr(ByVal ws As Worksheet, ByVal texte As String)
    ws.Range(CF_ADR_MESSAGE).Value = mod_Display.FR(texte)
End Sub

Private Sub FermerFeuilleCtr(ByVal ws As Worksheet)
    g_CtrEnCours = False
    g_CtrVerrouActif = False
    mod_Aide.MasquerFormulaire ws
End Sub


' =====================================================================================
' SUPPRESSION (liste frm_Contrats ; celle des factures est dans mod_Factures)
' =====================================================================================
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

Public Sub SupprimerContratSelection()

    Dim lo As ListObject
    Dim sel As Object
    Dim i As Long

    Set lo = TableAffContrats()
    If lo Is Nothing Then Exit Sub
    If ActiveSheet.Name <> mod_VarGlobales.NOM_FEUILLE_CONTRATS Then Exit Sub

    Set sel = LignesSelectionnees(lo)
    If sel.Count = 0 Then
        MsgBox mod_Display.FR("S{e2}lectionnez d'abord une ou plusieurs lignes du tableau."), vbInformation
        Exit Sub
    End If

    If MsgBox(sel.Count & mod_Display.FR(" contrat(s) {a2} supprimer ? Les factures d{e2}j{a2} saisies sont conserv{e2}es (leur op{e2}rateur est m{e2}moris{e2} sur la facture)."), _
              vbYesNo + vbQuestion, mod_Display.FR("Supprimer des contrats")) = vbNo Then Exit Sub

    For i = lo.ListRows.Count To 1 Step -1
        If sel.Exists(i) Then lo.ListRows(i).Delete
    Next i
    EnregistrerListeContrats

End Sub
