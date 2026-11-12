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
'   - Un indicateur créé par l'opérateur = lignes TTC de cet indicateur ; il compte dans le Montant total.
'   - La liste Indicateur ne propose que les indicateurs ayant des lignes pour le type choisi
'     (blocs de colonnes CO_COL_LISTE_INDIC_TYPE, nom ListeIndicConso reposé à chaque actualisation).
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
Public Const CO_ADR_INDIC As String = "C8"      ' les champs C8 à C12 n'apparaissent qu'au fil de la saisie (voir AppliquerVisibilite)
Public Const CO_ADR_DEBUT As String = "C9"      ' premier mois de la période (liste des mois en base)
Public Const CO_ADR_FIN As String = "C10"       ' dernier mois (au plus 11 mois après le début)
Public Const CO_ADR_FORME As String = "C11"
Public Const CO_ADR_REPART As String = "C12"
Public Const CO_ADR_MESSAGE As String = "B15"
Public Const CO_LIGNE_ENTETE As Long = 17       ' en-tête du tableau mensuel
Public Const CO_NB_LIGNES_MAX As Long = 120     ' lignes de données (120 mois = 10 ans au plus)
Public Const CO_NB_MOIS_MAX As Long = 120       ' mois au plus entre Début et Fin (12 = un an ; ne pas dépasser CO_NB_LIGNES_MAX)
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
Public Const CO_COL_LISTE_FORME As Long = 43
Public Const CO_COL_LISTE_REPART As Long = 44
Public Const CO_COL_LISTE_DATES As Long = 46    ' AT à BD : mois disponibles en base, une colonne par choix de type (listes Début et Fin)
Public Const CO_NB_COL_DATES As Long = 11       ' (tous) + au plus 10 types de consommation
Public Const CO_COL_LISTE_DATES_DESC As Long = 57   ' une colonne par choix de type : les mêmes mois, du plus récent au plus ancien (liste Fin)
Public Const CO_NB_LISTE_DATES As Long = 200    ' mois au plus dans chaque colonne
Public Const CO_COL_LISTE_INDIC_TYPE As Long = 68   ' une colonne par choix de type : indicateurs réellement alimentés, une colonne par choix de type (liste Indicateur)
Public Const CO_NB_LISTE_INDIC As Long = 50    ' indicateurs au plus dans chaque colonne
Public Const CO_NB_EXTRA As Long = 12           ' indicateurs créés par l'opérateur suivis au plus dans le calcul (les suivants sont regroupés dans le dernier)

' --- Formulaire frm_Contrat ---
Public Const CF_ADR_TYPE As String = "C6"
Public Const CF_ADR_OPERATEUR As String = "C7"
Public Const CF_ADR_UNITE As String = "C8"
Public Const CF_ADR_DEBUT As String = "C9"
Public Const CF_ADR_FIN As String = "C10"
Public Const CF_ADR_ABO As String = "C11"
Public Const CF_ADR_REFERENCE As String = "C12"
Public Const CF_ADR_MESSAGE As String = "B14"           ' zone fusionnée B14:E14

' --- Sens d'un type de consommation (ajout 11/11/2026) ---
' Un type est un ACHAT (valeur par défaut : on paie le fournisseur, l'opération bancaire est négative) ou une VENTE (revente
' d'électricité : on encaisse, l'opération est positive). Le sens du type est proposé à la saisie d'une facture, qui le
' mémorise dans sa propre colonne Sens (voir mod_Factures.FAC_COL_SENS). Colonne de TblTypesConso créée au premier besoin.
Public Const SENS_ACHAT As String = "Achat"
Public Const SENS_VENTE As String = "Vente"
Public Const TYPE_COL_SENS As String = "Sens"

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

' --- Premier et dernier mois (année * 12 + mois - 1) couverts par les factures, par choix de type ((tous), puis chaque type) ; -1 = aucun ---
Private g_MinType() As Long
Private g_MaxType() As Long

' --- Indicateurs créés par l'opérateur (hors Abonnement, Consommation, Taxes) présents dans le calcul en cours ---
Private g_ExNb As Long
Private g_ExNoms(1 To CO_NB_EXTRA) As String
Private g_ExDic As Object
Private g_ExQty(1 To CO_NB_EXTRA) As Boolean    ' l'indicateur créé a une vraie quantité (m3, kWh)
Private g_ExDer As Object       ' libellés "Prix HT X" et "Montant X" -> "rang|genre"

' --- Montants répartis : (mois, série, groupe, champ). Groupe 1 Consommation, 2 Abonnement, 3 Taxes, 3 + rang = indicateur créé ---
Private g_A() As Double
Private g_Has() As Boolean      ' (mois, série, groupe) : au moins une ligne du groupe tombe dans ce mois
Private Const GRP_CONSO As Long = 1
Private Const GRP_ABO As Long = 2
Private Const GRP_TAXE As Long = 3
Private Const CO_NB_GRP As Long = 3 + CO_NB_EXTRA
Private Const F_QTE As Long = 1
Private Const F_HT As Long = 2
Private Const F_TTC As Long = 3

' --- Genre de valeur d'un indicateur ---
Private Const K_QTE As Long = 1         ' la quantité (m3, kWh)
Private Const K_PRIX As Long = 2        ' le prix unitaire HT
Private Const K_MONTANT As Long = 3     ' le montant TTC du groupe
Private Const K_TOTAL As Long = 4       ' le montant TTC de toutes les lignes

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

' Types de base : proposés tant que le tableau TblTypesConso n'existe pas (avant la réinstallation) et semés dedans.
Public Function TypesDeBase() As Variant
    TypesDeBase = Array(TypeAdduction(), TypeRecyclage(), TypeElec(), TypeGaz())
End Function

' Les types de consommation : ceux de TblTypesConso (feuille masquée Conso_Types), dans l'ordre du tableau, au plus
' CO_NB_COL_DATES - 1 (un bloc de colonnes techniques par type dans frm_Consommations). Tableau de base 0.
Public Function TypesConso() As Variant

    Dim lo As ListObject
    Dim vTyp() As Variant
    Dim res() As Variant
    Dim i As Long, nbTyp As Long
    Dim s As String

    Set lo = TableTypes()
    If Not lo Is Nothing Then
        If Not TableVide(lo) Then
            vTyp = LireColonne(lo, "TypeConso")
            ReDim res(0 To lo.ListRows.Count - 1)
            For i = 1 To lo.ListRows.Count
                s = mod_DataStructure.CellText(vTyp(i))
                If s <> "" And nbTyp < CO_NB_COL_DATES - 1 Then
                    res(nbTyp) = s
                    nbTyp = nbTyp + 1
                End If
            Next i
        End If
    End If

    If nbTyp = 0 Then
        TypesConso = TypesDeBase()
    Else
        ReDim Preserve res(0 To nbTyp - 1)
        TypesConso = res
    End If

End Function

Public Function TableTypes() As ListObject
    Set TableTypes = ObtenirTable(mod_VarGlobales.NOM_FEUILLE_DATA_TYPES, mod_VarGlobales.NOM_TABLE_TYPES)
End Function

' --- Sens d'un type de consommation (ajout 11/11/2026) ---
' Un type est un ACHAT ou une VENTE : voir les constantes SENS_ACHAT et SENS_VENTE en tête de module.

' Colonne « Sens » de TblTypesConso, créée au premier besoin. À la création, les types dont le nom contient « vente »
' (Revente Electricité...) reçoivent le sens Vente, les autres le sens Achat.
Public Function ColonneSensType(ByVal lo As ListObject) As Long

    Dim col As ListColumn
    Dim i As Long, idxTyp As Long

    ColonneSensType = ColIdx(lo, TYPE_COL_SENS)
    If ColonneSensType > 0 Then Exit Function
    Set col = lo.ListColumns.Add
    col.Name = TYPE_COL_SENS
    ColonneSensType = col.Index
    If lo.DataBodyRange Is Nothing Then Exit Function
    col.DataBodyRange.NumberFormat = "@"
    idxTyp = ColIdx(lo, "TypeConso")
    For i = 1 To lo.ListRows.Count
        If InStr(1, mod_DataStructure.CellText(lo.DataBodyRange.Cells(i, idxTyp).Value2), "vente", vbTextCompare) > 0 Then
            lo.DataBodyRange.Cells(i, ColonneSensType).Value = SENS_VENTE
        ElseIf mod_DataStructure.CellText(lo.DataBodyRange.Cells(i, idxTyp).Value2) <> "" Then
            lo.DataBodyRange.Cells(i, ColonneSensType).Value = SENS_ACHAT
        End If
    Next i

End Function

' Sens (SENS_ACHAT ou SENS_VENTE) mémorisé avec un type ; Achat si le type est inconnu ou sans sens.
Public Function SensDuType(ByVal typ As String) As String

    Dim lo As ListObject
    Dim i As Long, idxTyp As Long, idxSens As Long

    SensDuType = SENS_ACHAT
    Set lo = TableTypes()
    If lo Is Nothing Then Exit Function
    idxSens = ColonneSensType(lo)
    If TableVide(lo) Then Exit Function
    idxTyp = ColIdx(lo, "TypeConso")
    For i = 1 To lo.ListRows.Count
        If StrComp(mod_DataStructure.CellText(lo.DataBodyRange.Cells(i, idxTyp).Value2), typ, vbTextCompare) = 0 Then
            If StrComp(mod_DataStructure.CellText(lo.DataBodyRange.Cells(i, idxSens).Value2), SENS_VENTE, vbTextCompare) = 0 Then SensDuType = SENS_VENTE
            Exit Function
        End If
    Next i

End Function

' Macro (Alt+F8) : change le sens (achat / vente) d'un type de consommation existant. N'agit que sur les factures
' saisies ensuite : chaque facture garde le sens qu'elle a enregistré.
Public Sub ModifierSensTypeConso()

    Dim lo As ListObject
    Dim nomType As String
    Dim i As Long, idxTyp As Long, idxSens As Long
    Dim rep As VbMsgBoxResult

    Set lo = TableTypes()
    If lo Is Nothing Then
        MsgBox mod_Display.FR("La liste des types de consommation n'existe pas encore. Ex{e2}cutez d'abord la macro InstallerConsommations."), vbExclamation
        Exit Sub
    End If
    idxSens = ColonneSensType(lo)
    idxTyp = ColIdx(lo, "TypeConso")
    nomType = Trim$(InputBox(mod_Display.FR("Nom du type de consommation {a2} modifier (par exemple Revente Electricit{e2}) :"), _
                             mod_Display.FR("Sens d'un type de consommation")))
    If nomType = "" Then Exit Sub
    If TableVide(lo) Then Exit Sub
    For i = 1 To lo.ListRows.Count
        If StrComp(mod_DataStructure.CellText(lo.DataBodyRange.Cells(i, idxTyp).Value2), nomType, vbTextCompare) = 0 Then
            rep = MsgBox(mod_Display.FR("Type : ") & mod_DataStructure.CellText(lo.DataBodyRange.Cells(i, idxTyp).Value2) & vbCrLf & _
                         mod_Display.FR("Sens actuel : ") & SensDuType(nomType) & vbCrLf & vbCrLf & _
                         mod_Display.FR("Oui : VENTE (revente, les op{e2}rations li{e2}es sont des encaissements).") & vbCrLf & _
                         mod_Display.FR("Non : ACHAT (consommation, les op{e2}rations li{e2}es sont des paiements).") & vbCrLf & _
                         mod_Display.FR("Annuler : ne rien changer."), vbYesNoCancel + vbQuestion, mod_Display.FR("Sens d'un type de consommation"))
            If rep = vbCancel Then Exit Sub
            lo.DataBodyRange.Cells(i, idxSens).NumberFormat = "@"
            If rep = vbYes Then lo.DataBodyRange.Cells(i, idxSens).Value = SENS_VENTE Else lo.DataBodyRange.Cells(i, idxSens).Value = SENS_ACHAT
            Exit Sub
        End If
    Next i
    MsgBox mod_Display.FR("Type introuvable : ") & nomType, vbExclamation, mod_Display.FR("Sens d'un type de consommation")

End Sub

Public Function LibTous() As String
    LibTous = "(tous)"
End Function

Public Function LibIndQuantite() As String
    LibIndQuantite = "Consommation"          ' la quantité consommée, dans l'unité du type (m³, kWh)
End Function

Public Function LibIndPrix() As String
    LibIndPrix = "Prix unitaire HT"
End Function

Public Function LibIndConso() As String
    LibIndConso = "Montant consommation"     ' le montant TTC des lignes de consommation (€)
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

' Unité mémorisée avec un type dans TblTypesConso ("" si le type n'y est pas ou n'a pas d'unité).
Private Function UniteDuTypeMemorisee(ByVal typ As String) As String

    Dim lo As ListObject
    Dim vTyp() As Variant, vUni() As Variant
    Dim i As Long

    Set lo = TableTypes()
    If lo Is Nothing Then Exit Function
    If TableVide(lo) Then Exit Function
    If ColIdx(lo, "Unite") = 0 Then Exit Function
    vTyp = LireColonne(lo, "TypeConso")
    vUni = LireColonne(lo, "Unite")
    For i = 1 To lo.ListRows.Count
        If StrComp(mod_DataStructure.CellText(vTyp(i)), typ, vbTextCompare) = 0 Then
            UniteDuTypeMemorisee = mod_DataStructure.CellText(vUni(i))
            Exit Function
        End If
    Next i

End Function

' Unité proposée par défaut selon le type (modifiable dans le contrat) : celle du tableau des types, sinon m3 pour l'eau et kWh.
Public Function UniteDefaut(ByVal typ As String) As String
    UniteDefaut = UniteDuTypeMemorisee(typ)
    If UniteDefaut <> "" Then Exit Function
    If StrComp(typ, TypeAdduction(), vbTextCompare) = 0 Or StrComp(typ, TypeRecyclage(), vbTextCompare) = 0 Then
        UniteDefaut = "m" & ChrW(179)
    Else
        UniteDefaut = "kWh"
    End If
End Function

' Vrai si l'unité est une de celles proposées automatiquement (m3, kWh ou l'unité d'un type de la liste) :
' elle peut alors être remplacée quand on change de type.
Private Function EstUniteDefaut(ByVal u As String) As Boolean

    Dim tps As Variant
    Dim i As Long

    EstUniteDefaut = (StrComp(u, "m" & ChrW(179), vbTextCompare) = 0 Or StrComp(u, "kWh", vbTextCompare) = 0)
    If EstUniteDefaut Then Exit Function
    tps = TypesConso()
    For i = 0 To UBound(tps)
        If StrComp(UniteDuTypeMemorisee(CStr(tps(i))), u, vbTextCompare) = 0 Then
            EstUniteDefaut = True
            Exit Function
        End If
    Next i

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
        Case "ID_Contrat", "ID_Facture", "OrdreBloc", "NbOperations": FormatColonneConso = "0"
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
        ' ListRows.Add échoue (erreur 50290) dans certains contextes, par exemple appelé depuis un bouton pendant
        ' qu'un formulaire attend (boucle DoEvents) : on agrandit alors le tableau d'une ligne par Resize.
        On Error Resume Next
        Set NouvelleLigne = lo.ListRows.Add
        If Err.Number <> 0 Then
            Err.Clear
            Set NouvelleLigne = Nothing
            lo.Resize lo.Range.Resize(lo.Range.Rows.Count + 1)
            If Err.Number = 0 Then Set NouvelleLigne = lo.ListRows(lo.ListRows.Count)
        End If
        On Error GoTo 0
        If NouvelleLigne Is Nothing Then Set NouvelleLigne = lo.ListRows.Add     ' l'erreur d'origine remonte
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
    Application.EnableEvents = False        ' les écritures de Début/Fin ne doivent pas relancer ConsoTraiterChangement
    RafraichirConso
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

Erreur:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox mod_Display.FR("Erreur pendant le calcul des consommations.") & vbLf & _
           "Erreur " & Err.Number & " : " & Err.Description, vbExclamation, mod_Display.FR("Suivi des consommations")

End Sub

Private Sub RafraichirConso()

    Dim ws As Worksheet
    Dim typ As String, indic As String, indicSerie As String, forme As String, repart As String
    Dim idxDeb As Long, nbM As Long
    Dim nbRows As Long, r As Long, g As Long, nbGraph As Long
    Dim hasM() As Boolean
    Dim tous As Boolean, prorata As Boolean, mixte As Boolean
    Dim note As String
    Dim serNoms() As String, nS As Long
    Dim aHas() As Boolean
    Dim tabT() As Variant, tabH() As Variant
    Dim p As Long, s As Long, ent As Long, idx As Long, k As Long
    Dim q As Double, ht As Double, ttc As Double, tot As Double, totS As Double
    Dim totQ As Double, totHT As Double, totTTC As Double, totAll As Double
    Dim oper As String
    Dim nbLignesPrises As Long
    Dim unite As String, texte As String, nomRef As String
    Dim lignes() As TypeLigne, nbL As Long
    Dim etape As Long
    Dim gref As Long, kind As Long, kindBarres As Long
    Dim has As Boolean

    Set ws = FeuilleConso()
    If ws Is Nothing Or Not ConsoInstallee() Then
        MsgBox mod_Display.FR("Le suivi des consommations n'est pas install{e2} : lancez d'abord InstallerConsommations."), vbExclamation
        Exit Sub
    End If

    ChargerContrats
    nbL = mod_Factures.ChargerLignes(lignes, 0, True)      ' avec les indicateurs créés par l'opérateur
    MettreAJourListesConso ws, lignes, nbL
    AppliquerVisibilite ws
    etape = EtapeSaisie(ws)
    If etape < 3 Then
        EffacerResultats ws, etape
        Exit Sub
    End If
    If Not LireParametresConso(ws, typ, indic, forme, repart, idxDeb, nbM) Then Exit Sub

    tous = (StrComp(typ, LibTous(), vbTextCompare) = 0)
    prorata = (StrComp(repart, LibRepartProrata(), vbTextCompare) = 0)

    ChargerExtras lignes, nbL, typ, tous

    ' Indicateur choisi -> groupe de lignes de référence (colonnes quantité, prix et montants du tableau) et genre de valeur.
    If Not ResoudreIndic(indic, gref, kind) Then
        indic = LibIndMontant()         ' indicateur sans ligne pour ce type : le montant total est affiché
        ResoudreIndic indic, gref, kind
        note = " L'indicateur choisi n'a aucune ligne pour ce type : le montant total est affich{e2}."
    End If

    ' Combinaisons sans sens : pour "tous les types", pas de quantité ni de prix (unités
    ' différentes) et pas de courbe du prix.
    mixte = (StrComp(forme, LibFormeMixte(), vbTextCompare) = 0)
    If tous Then
        If kind = K_QTE Or kind = K_PRIX Then
            indic = LibIndMontant()
            ResoudreIndic indic, gref, kind
            note = note & " Pour tous les types, seuls les montants sont comparables : le montant total est affich{e2}."
        End If
        If mixte Then
            mixte = False
            forme = LibFormeBarres()
            note = note & " La courbe du prix n'existe que pour un type choisi."
        End If
    End If

    ' Avec la courbe du prix, un prix choisi comme indicateur est déjà la courbe : les barres montrent la quantité.
    indicSerie = indic
    kindBarres = kind
    If mixte And kind = K_PRIX Then
        kindBarres = K_QTE
        indicSerie = NomGroupe(gref)
    End If

    CalculerMois lignes, nbL, typ, tous, prorata, idxDeb, nbM, serNoms, nS, aHas, nbLignesPrises
    nomRef = NomGroupe(gref)
    If Len(nomRef) > 28 Then nomRef = Left$(nomRef, 27) & "."

    ' ---------------- Tableau mensuel (visible) et données du graphique (masquées) ----------------
    ent = CO_LIGNE_ENTETE
    ws.Range(ws.Cells(ent + 1, 2), ws.Cells(ent + CO_NB_LIGNES_MAX + 1, 1 + CO_NB_COL_TABLE)).ClearContents
    ws.Range(ws.Cells(ent + 1, 2), ws.Cells(ent + CO_NB_LIGNES_MAX + 1, 1 + CO_NB_COL_TABLE)).Font.Bold = False
    ws.Range(ws.Cells(ent, CO_COL_H_LIB), ws.Cells(ent + CO_NB_LIGNES_MAX, CO_COL_H_SERIE + CO_NB_SERIES - 1)).ClearContents

    ' Seuls les mois qui ont des données sont affichés (tableau et graphique).
    ReDim hasM(1 To nbM)
    nbRows = 0
    For p = 1 To nbM
        For s = 1 To nS
            If SerieAvecDonnees(p, s, kind, gref, aHas) Then hasM(p) = True
        Next s
        If hasM(p) Then nbRows = nbRows + 1
    Next p
    nbGraph = nbRows

    ReDim tabT(1 To nbRows + 1, 1 To CO_NB_COL_TABLE)
    ReDim tabH(1 To IIf(nbGraph > 0, nbGraph, 1), 1 To 2 + CO_NB_SERIES)

    ' Les mois sont parcourus du plus ancien au plus récent : le graphique est dans cet ordre
    ' (ligne g) ; le tableau est dans l'ordre inverse, la facture la plus récente en premier.
    ' Colonnes du tableau : 2 quantité et 3 prix HT du groupe de référence, 4 HT et 5 TTC du groupe,
    ' 6 autres indicateurs (total moins le groupe), 7 montant total, 8 opérateurs.
    g = 0
    For p = 1 To nbM
        If Not hasM(p) Then GoTo MoisSuivant
        g = g + 1
        r = nbRows - g + 1
        idx = idxDeb + p - 1
        tabT(r, 1) = DateSerial(idx \ 12, (idx Mod 12) + 1, 1)
        tabH(g, 1) = tabT(r, 1)
        q = 0: ht = 0: ttc = 0: tot = 0: oper = ""
        For s = 1 To nS
            totS = 0
            For k = 1 To CO_NB_GRP
                totS = totS + g_A(p, s, k, F_TTC)
            Next k
            q = q + g_A(p, s, gref, F_QTE)
            ht = ht + g_A(p, s, gref, F_HT)
            ttc = ttc + g_A(p, s, gref, F_TTC)
            tot = tot + totS
            has = SerieAvecDonnees(p, s, kind, gref, aHas)
            If has And Not tous Then
                If oper <> "" Then oper = oper & " / "
                oper = oper & serNoms(s)
            End If
            tabH(g, 2 + s) = ValeurIndic(kindBarres, g_A(p, s, gref, F_QTE), g_A(p, s, gref, F_HT), _
                                         g_A(p, s, gref, F_TTC), totS, has)
        Next s
        If Not tous Then
            tabT(r, 2) = q
            If q > 0 Then
                tabT(r, 3) = ht / q
                tabH(g, 2) = ht / q
            End If
        End If
        tabT(r, 4) = ht
        tabT(r, 5) = ttc
        tabT(r, 6) = tot - ttc
        tabT(r, 7) = tot
        tabT(r, 8) = oper
        totQ = totQ + q
        totHT = totHT + ht
        totTTC = totTTC + ttc
        totAll = totAll + tot
MoisSuivant:
    Next p

    ' Ligne de total, juste sous le dernier mois affiché.
    tabT(nbRows + 1, 1) = "Total"
    If Not tous Then
        tabT(nbRows + 1, 2) = totQ
        If totQ > 0 Then tabT(nbRows + 1, 3) = totHT / totQ
    End If
    tabT(nbRows + 1, 4) = totHT
    tabT(nbRows + 1, 5) = totTTC
    tabT(nbRows + 1, 6) = totAll - totTTC
    tabT(nbRows + 1, 7) = totAll

    ' Colonne des mois au format "mmm yyyy" partout (le mot "Total" s'y affiche tel quel).
    ws.Range(ws.Cells(ent + 1, 2), ws.Cells(ent + CO_NB_LIGNES_MAX + 1, 2)).NumberFormat = "mmm yyyy"
    ws.Range(ws.Cells(ent + 1, 2), ws.Cells(ent + nbRows + 1, 1 + CO_NB_COL_TABLE)).Value = tabT
    ws.Range(ws.Cells(ent + nbRows + 1, 2), ws.Cells(ent + nbRows + 1, 1 + CO_NB_COL_TABLE)).Font.Bold = True

    ' Données du graphique : en-têtes, puis une ligne par mois non vide (ordre chronologique).
    ws.Cells(ent, CO_COL_H_LIB).Value = "Mois"
    ws.Cells(ent, CO_COL_H_PRIX).Value = "Prix unitaire HT"
    For s = 1 To nS
        ws.Cells(ent, CO_COL_H_SERIE + s - 1).Value = serNoms(s)
    Next s
    If nbGraph > 0 Then
        ws.Range(ws.Cells(ent + 1, CO_COL_H_LIB), ws.Cells(ent + nbGraph, CO_COL_H_SERIE + CO_NB_SERIES - 1)).Value = tabH
        ws.Range(ws.Cells(ent + 1, CO_COL_H_LIB), ws.Cells(ent + nbGraph, CO_COL_H_LIB)).NumberFormat = "mmm yyyy"
    End If

    ' En-têtes : ils nomment le groupe de référence (l'indicateur choisi) et son unité (rien pour "tous").
    If tous Then
        unite = ""
        ws.Cells(ent, 3).Value = "Quantit" & ChrW(233)
        ws.Cells(ent, 4).Value = "Prix unitaire HT"
    Else
        unite = UniteGroupe(typ, lignes, nbL, gref)
        ws.Cells(ent, 3).Value = nomRef & IIf(unite <> "", " (" & unite & ")", "")
        ws.Cells(ent, 4).Value = "Prix HT (" & SymboleEuro() & IIf(unite <> "", "/" & unite, "") & ")"
    End If
    ws.Cells(ent, 5).Value = "HT " & nomRef & " (" & SymboleEuro() & ")"
    ws.Cells(ent, 6).Value = "TTC " & nomRef & " (" & SymboleEuro() & ")"
    ws.Cells(ent, 7).Value = "Autres indicateurs (" & SymboleEuro() & ")"
    ws.Cells(ent, 8).Value = "Montant total (" & SymboleEuro() & ")"

    ' Résumé.
    texte = IIf(tous, mod_Display.FR("Tous les types"), typ) & " -- " & _
            Format$(DateSerial(idxDeb \ 12, (idxDeb Mod 12) + 1, 1), "mmm yyyy") & mod_Display.FR(" {a2} ") & _
            Format$(DateSerial((idxDeb + nbM - 1) \ 12, ((idxDeb + nbM - 1) Mod 12) + 1, 1), "mmm yyyy") & _
            " -- " & nbLignesPrises & mod_Display.FR(" ligne(s) de facture prise(s) en compte")
    If nbLignesPrises = 0 Then texte = texte & mod_Display.FR(". Aucune facture sur cette p{e2}riode.")
    ws.Range(CO_ADR_MESSAGE).Value = texte & mod_Display.FR(note)

    AjusterLargeursTableau ws, nbRows

    If nbGraph = 0 Then nS = 0      ' aucun mois à tracer : le graphique indique qu'il n'y a pas de facture
    ConstruireGraphiqueConso ws, nbGraph, nS, serNoms, indic, indicSerie, kindBarres, forme, typ, tous, unite

End Sub

' Lit et contrôle les paramètres de l'écran. Renvoie Faux (avec un message) s'ils sont incorrects.
Private Function LireParametresConso(ByVal ws As Worksheet, ByRef typ As String, ByRef indic As String, _
                                     ByRef forme As String, ByRef repart As String, _
                                     ByRef idxDeb As Long, ByRef nbM As Long) As Boolean

    Dim idxFin As Long, c As Long

    typ = mod_DataStructure.CellText(ws.Range(CO_ADR_TYPE).Value2)
    indic = mod_DataStructure.CellText(ws.Range(CO_ADR_INDIC).Value2)
    forme = mod_DataStructure.CellText(ws.Range(CO_ADR_FORME).Value2)
    repart = mod_DataStructure.CellText(ws.Range(CO_ADR_REPART).Value2)

    If forme = "" Then forme = LibFormeBarres()
    If repart = "" Then repart = LibRepartProrata()

    ' Début : un mois couvert par les factures du type choisi (sinon la saisie reprend à Début, par
    ' exemple après un changement de type fait sans le code des événements de la feuille).
    c = IndexTypeListe(typ)
    idxDeb = IndexMois(ws.Range(CO_ADR_DEBUT).Value2)
    If c < 0 Then
        EffacerResultats ws, 0
        Exit Function
    End If
    If g_MinType(c) < 0 Or idxDeb < g_MinType(c) Or idxDeb > g_MaxType(c) Then
        ws.Range(CO_ADR_DEBUT).MergeArea.ClearContents
        ws.Range(CO_ADR_FIN).MergeArea.ClearContents
        AppliquerVisibilite ws
        EffacerResultats ws, 2
        Exit Function
    End If

    ' Fin : vide, elle va jusqu'au dernier mois du type (dans la limite de CO_NB_MOIS_MAX mois).
    idxFin = IndexMois(ws.Range(CO_ADR_FIN).Value2)
    If idxFin < 0 Then
        idxFin = g_MaxType(c)
        If idxFin - idxDeb >= CO_NB_MOIS_MAX Then idxFin = idxDeb + CO_NB_MOIS_MAX - 1
    End If

    If idxFin < idxDeb Then
        MsgBox mod_Display.FR("La date de fin ne peut pas {e1}tre ant{e2}rieure {a1} la date de d{e2}but."), vbExclamation
        Exit Function
    End If
    If idxFin - idxDeb >= CO_NB_MOIS_MAX Then
        MsgBox mod_Display.FR("La p{e2}riode ne peut pas d{e2}passer ") & CO_NB_MOIS_MAX & mod_Display.FR(" mois : choisissez une date de fin plus proche du d{e2}but."), vbExclamation
        Exit Function
    End If

    nbM = idxFin - idxDeb + 1
    EcrireMois ws.Range(CO_ADR_DEBUT), idxDeb
    EcrireMois ws.Range(CO_ADR_FIN), idxFin

    LireParametresConso = True

End Function

' Rang du type dans le bloc des mois (0 = (tous), 1 à 4 = types) ; -1 si le texte n'est pas un type.
Private Function IndexTypeListe(ByVal typ As String) As Long
    Dim tps As Variant, i As Long
    IndexTypeListe = -1
    If StrComp(typ, LibTous(), vbTextCompare) = 0 Then
        IndexTypeListe = 0
        Exit Function
    End If
    tps = TypesConso()
    For i = 0 To UBound(tps)
        If StrComp(typ, tps(i), vbTextCompare) = 0 Then IndexTypeListe = i + 1
    Next i
End Function

' Ajuste la largeur des colonnes du tableau mensuel à leur contenu (en-têtes compris), sans descendre
' sous la largeur d'origine de l'écran. Seules les cellules du tableau comptent : les paramètres au-dessus
' et le résumé ne sont pas pris en compte.
Private Sub AjusterLargeursTableau(ByVal ws As Worksheet, ByVal nbRows As Long)

    Dim largeurs As Variant
    Dim j As Long
    Dim w As Double

    largeurs = Array(24, 12, 12, 12, 12, 12, 12, 26)       ' colonnes B à I, comme à la construction de l'écran
    ws.Range(ws.Cells(CO_LIGNE_ENTETE, 2), ws.Cells(CO_LIGNE_ENTETE + nbRows + 1, 1 + CO_NB_COL_TABLE)).Columns.AutoFit
    For j = 0 To CO_NB_COL_TABLE - 1
        w = ws.Columns(2 + j).ColumnWidth + 2
        If w < largeurs(j) Then w = largeurs(j)
        ws.Columns(2 + j).ColumnWidth = w
    Next j

End Sub

' Étape de la saisie d'après les cellules : 0 = type à choisir, 1 = indicateur, 2 = début, 3 = complète.
' Fin est facultative (vide : jusqu'au dernier mois du type).
Private Function EtapeCellules(ByVal ws As Worksheet) As Long
    If mod_DataStructure.CellText(ws.Range(CO_ADR_TYPE).Value2) = "" Then Exit Function
    EtapeCellules = 1
    If mod_DataStructure.CellText(ws.Range(CO_ADR_INDIC).Value2) = "" Then Exit Function
    EtapeCellules = 2
    If IndexMois(ws.Range(CO_ADR_DEBUT).Value2) < 0 Then Exit Function
    EtapeCellules = 3
End Function

' Même étape, avec le contrôle du type et du début : à appeler après MettreAJourListesConso.
Private Function EtapeSaisie(ByVal ws As Worksheet) As Long
    Dim c As Long, idx As Long
    EtapeSaisie = EtapeCellules(ws)
    c = IndexTypeListe(mod_DataStructure.CellText(ws.Range(CO_ADR_TYPE).Value2))
    If c < 0 Then
        EtapeSaisie = 0
    ElseIf EtapeSaisie = 3 Then
        idx = IndexMois(ws.Range(CO_ADR_DEBUT).Value2)
        If g_MinType(c) < 0 Or idx < g_MinType(c) Or idx > g_MaxType(c) Then EtapeSaisie = 2
    End If
End Function

' Les champs apparaissent au fil de la saisie : Type d'abord ; puis Indicateur ; puis Début ;
' puis Fin, Forme du graphique et Répartition des lignes.
Public Sub AppliquerVisibilite(ByVal ws As Worksheet)
    Dim etape As Long
    etape = EtapeCellules(ws)
    ws.Rows(ws.Range(CO_ADR_INDIC).Row).Hidden = (etape < 1)
    ws.Rows(ws.Range(CO_ADR_DEBUT).Row).Hidden = (etape < 2)
    ws.Rows(ws.Range(CO_ADR_FIN).Row).Hidden = (etape < 3)
    ws.Rows(ws.Range(CO_ADR_FORME).Row).Hidden = (etape < 3)
    ws.Rows(ws.Range(CO_ADR_REPART).Row).Hidden = (etape < 3)
End Sub

' Appelée par le code de la feuille (Worksheet_Change, voir CodeBehind_frm_Consommations.txt) :
' un nouveau type efface Début et Fin (leurs listes changent) ; un nouveau début efface Fin.
Public Sub ConsoTraiterChangement(ByVal ws As Worksheet, ByVal Target As Range)

    If Intersect(Target, ws.Range("B7:G13")) Is Nothing Then Exit Sub

    On Error GoTo Sortie
    Application.EnableEvents = False
    If Not Intersect(Target, ws.Range(CO_ADR_TYPE).MergeArea) Is Nothing Then
        ws.Range(CO_ADR_DEBUT).MergeArea.ClearContents
        ws.Range(CO_ADR_FIN).MergeArea.ClearContents
        ' L'indicateur déjà choisi est conservé s'il existe aussi pour le nouveau type, sinon il est vidé.
        If Not IndicProposePourType(ws) Then ws.Range(CO_ADR_INDIC).MergeArea.ClearContents
    ElseIf Not Intersect(Target, ws.Range(CO_ADR_DEBUT).MergeArea) Is Nothing Then
        ws.Range(CO_ADR_FIN).MergeArea.ClearContents
    End If
    AppliquerVisibilite ws

    ' Le focus passe sur le champ qui vient d'apparaître : sinon la flèche de liste du champ
    ' précédent reste affichée et propose encore l'ancienne liste.
    On Error Resume Next
    If Not Intersect(Target, ws.Range(CO_ADR_TYPE).MergeArea) Is Nothing Then
        If mod_DataStructure.CellText(ws.Range(CO_ADR_TYPE).Value2) <> "" Then ws.Range(CO_ADR_INDIC).Select
    ElseIf Not Intersect(Target, ws.Range(CO_ADR_INDIC).MergeArea) Is Nothing Then
        If mod_DataStructure.CellText(ws.Range(CO_ADR_INDIC).Value2) <> "" Then ws.Range(CO_ADR_DEBUT).Select
    ElseIf Not Intersect(Target, ws.Range(CO_ADR_DEBUT).MergeArea) Is Nothing Then
        If IndexMois(ws.Range(CO_ADR_DEBUT).Value2) >= 0 Then ws.Range(CO_ADR_FIN).Select
    End If
    On Error GoTo 0

Sortie:
    Application.EnableEvents = True

End Sub

' Vrai si l'indicateur de la cellule figure dans la liste des indicateurs du type choisi (dernière actualisation).
Private Function IndicProposePourType(ByVal ws As Worksheet) As Boolean
    Dim ind As String, typ As String
    Dim c As Long
    ind = mod_DataStructure.CellText(ws.Range(CO_ADR_INDIC).Value2)
    If ind = "" Then
        IndicProposePourType = True
        Exit Function
    End If
    typ = mod_DataStructure.CellText(ws.Range(CO_ADR_TYPE).Value2)
    c = IndexTypeListe(typ)
    If c < 0 Then Exit Function
    IndicProposePourType = (Application.WorksheetFunction.CountIf(ws.Range(ws.Cells(2, CO_COL_LISTE_INDIC_TYPE + c), _
        ws.Cells(CO_NB_LISTE_INDIC + 1, CO_COL_LISTE_INDIC_TYPE + c)), ind) > 0)
End Function

' Saisie incomplète : le tableau, le graphique et le résumé sont vidés et un message dit quoi choisir.
Private Sub EffacerResultats(ByVal ws As Worksheet, ByVal etape As Long)
    Dim co As ChartObject
    Dim msg As String
    ws.Range(ws.Cells(CO_LIGNE_ENTETE + 1, 2), ws.Cells(CO_LIGNE_ENTETE + CO_NB_LIGNES_MAX + 1, 1 + CO_NB_COL_TABLE)).ClearContents
    ws.Range(ws.Cells(CO_LIGNE_ENTETE, CO_COL_H_LIB), ws.Cells(CO_LIGNE_ENTETE + CO_NB_LIGNES_MAX, CO_COL_H_SERIE + CO_NB_SERIES - 1)).ClearContents
    Select Case etape
        Case 0: msg = "Choisissez d'abord un type de consommation."
        Case 1: msg = "Choisissez un indicateur."
        Case Else: msg = "Choisissez le mois de d{e2}but (la fin est facultative)."
    End Select
    On Error Resume Next
    Set co = ws.ChartObjects(CO_GRAPH)
    On Error GoTo 0
    If Not co Is Nothing Then
        Do While co.Chart.SeriesCollection.Count > 0
            co.Chart.SeriesCollection(1).Delete
        Loop
        co.Chart.HasTitle = True
        co.Chart.ChartTitle.Text = mod_Display.FR(msg)
    End If
    ws.Range(CO_ADR_MESSAGE).Value = mod_Display.FR(msg)
End Sub

' Numéro de mois (année * 12 + mois - 1) d'une date de cellule ; -1 si la cellule ne contient pas de date.
Private Function IndexMois(ByVal v As Variant) As Long
    IndexMois = -1
    If IsEmpty(v) Then Exit Function
    If VarType(v) <> vbDouble Then Exit Function
    If v < 1 Then Exit Function
    IndexMois = Year(CDate(v)) * 12 + Month(CDate(v)) - 1
End Function

Private Sub EcrireMois(ByVal c As Range, ByVal idx As Long)
    c.NumberFormat = "mmm yyyy"
    c.Value = DateSerial(idx \ 12, (idx Mod 12) + 1, 1)
End Sub

' Répartit les lignes de factures sur les mois de la période. Pour chaque (mois, série, groupe) le tableau g_A
' garde la quantité, le montant HT et le montant TTC. Un groupe = une famille de base (1 Consommation, 2 Abonnement,
' 3 Taxes) ou un indicateur créé par l'opérateur (groupe 3 + son rang, voir ChargerExtras).
' Série = opérateur (un seul type choisi) ou type (tous les types).
Private Sub CalculerMois(ByRef t() As TypeLigne, ByVal nb As Long, ByVal typ As String, ByVal tous As Boolean, _
                         ByVal prorata As Boolean, ByVal idxDeb As Long, ByVal nbM As Long, _
                         ByRef serNoms() As String, ByRef nS As Long, _
                         ByRef aHas() As Boolean, ByRef nbPris As Long)

    Dim dic As Object
    Dim j As Long, i As Long, s As Long, p As Long, grp As Long
    Dim cle As String
    Dim iDeb As Long, iFin As Long, y As Long, m As Long
    Dim dMois As Double, fMois As Double, a As Double, b As Double
    Dim frac As Double, totalJours As Double

    ReDim serNoms(1 To CO_NB_SERIES)
    ReDim g_A(1 To nbM, 1 To CO_NB_SERIES, 1 To CO_NB_GRP, 1 To 3)
    ReDim aHas(1 To nbM, 1 To CO_NB_SERIES)
    ReDim g_Has(1 To nbM, 1 To CO_NB_SERIES, 1 To CO_NB_GRP)
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

            grp = GroupeDeLigne(t(j))

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
                        Ajouter t(j), frac, p, s, grp
                        aHas(p, s) = True
                    End If
                Next i
                If iFin >= idxDeb And iDeb <= idxDeb + nbM - 1 Then nbPris = nbPris + 1
            Else
                p = iFin - idxDeb + 1
                If p >= 1 And p <= nbM Then
                    Ajouter t(j), 1, p, s, grp
                    aHas(p, s) = True
                    nbPris = nbPris + 1
                End If
            End If
        End If
    Next j

End Sub

' Vrai si la série a des données à montrer ce mois-là : au moins une ligne du groupe de référence (indicateur choisi), ou
' une ligne quelconque pour le montant total. Un mois qui n'a que d'autres indicateurs n'est pas affiché (sinon des 0).
Private Function SerieAvecDonnees(ByVal p As Long, ByVal s As Long, ByVal kind As Long, ByVal gref As Long, _
                                  ByRef aHas() As Boolean) As Boolean
    If kind = K_TOTAL Then
        SerieAvecDonnees = aHas(p, s)
    Else
        SerieAvecDonnees = g_Has(p, s, gref)
    End If
End Function

' Groupe d'une ligne : 1 Consommation, 2 Abonnement, 3 Taxes, 3 + rang pour un indicateur créé.
Private Function GroupeDeLigne(ByRef l As TypeLigne) As Long
    Dim k As Long
    Select Case l.Nat
        Case mod_Factures.NAT_CONSO: GroupeDeLigne = GRP_CONSO
        Case mod_Factures.NAT_ABO: GroupeDeLigne = GRP_ABO
        Case mod_Factures.NAT_TAXE: GroupeDeLigne = GRP_TAXE
        Case Else
            k = IndexExtra(l.Indic)
            If k = 0 Then k = g_ExNb        ' au-delà de CO_NB_EXTRA indicateurs : regroupés dans le dernier
            If k = 0 Then k = 1             ' ne devrait pas arriver (ChargerExtras a vu toutes les lignes)
            GroupeDeLigne = GRP_TAXE + k
    End Select
End Function

' Ajoute la fraction "frac" d'une ligne de facture au mois p, série s, groupe grp (quantité, HT, TTC).
Private Sub Ajouter(ByRef l As TypeLigne, ByVal frac As Double, ByVal p As Long, ByVal s As Long, ByVal grp As Long)
    g_A(p, s, grp, F_QTE) = g_A(p, s, grp, F_QTE) + l.Qte * frac
    g_A(p, s, grp, F_HT) = g_A(p, s, grp, F_HT) + l.HT * frac
    g_A(p, s, grp, F_TTC) = g_A(p, s, grp, F_TTC) + l.TTC * frac
    g_Has(p, s, grp) = True
End Sub

' Vrai si l'unité est une vraie quantité consommée (m3, kWh...) : l'indicateur a alors une quantité et un prix unitaire.
Private Function UniteQuantifiable(ByVal u As String) As Boolean
    Dim norm As String
    u = LCase$(Trim$(u))
    Select Case u
        Case "m3", "m" & ChrW(179), "kwh", "mwh", "l", "litre", "litres"
            UniteQuantifiable = True
        Case Else
            ' Unité de consommation ajoutée par l'opérateur (par exemple stère).
            If u <> "" Then UniteQuantifiable = mod_Unites.EstUniteConso(u, norm)
    End Select
End Function

' Indicateurs créés par l'opérateur (ligne d'indicateur autre qu'Abonnement, Consommation, Taxes) qui ont des lignes
' pour le type choisi (tous les types si tous) : ils reçoivent un rang 1..CO_NB_EXTRA ; au-delà, ils sont regroupés dans
' le dernier. g_ExQty dit si l'indicateur a une vraie quantité (m3, kWh) : il propose alors trois choix dans la liste
' (X = la quantité, "Prix HT X", "Montant X") ; g_ExDer relie ces deux derniers libellés à (rang, genre).
Private Sub ChargerExtras(ByRef t() As TypeLigne, ByVal nb As Long, ByVal typ As String, ByVal tous As Boolean)
    Dim j As Long, k As Long
    g_ExNb = 0
    Set g_ExDic = CreateObject("Scripting.Dictionary")
    g_ExDic.CompareMode = 1
    Set g_ExDer = CreateObject("Scripting.Dictionary")
    g_ExDer.CompareMode = 1
    For k = 1 To CO_NB_EXTRA
        g_ExQty(k) = False
    Next k
    For j = 1 To nb
        If t(j).Nat = mod_Factures.NAT_AUTRE And (tous Or StrComp(t(j).Typ, typ, vbTextCompare) = 0) Then
            If Not g_ExDic.Exists(t(j).Indic) Then
                If g_ExNb < CO_NB_EXTRA Then
                    g_ExNb = g_ExNb + 1
                    g_ExNoms(g_ExNb) = t(j).Indic
                    g_ExDic.Add t(j).Indic, g_ExNb
                End If
            End If
            k = IndexExtra(t(j).Indic)
            If k = 0 Then k = g_ExNb
            If k > 0 Then
                If UniteQuantifiable(t(j).Unite) Then g_ExQty(k) = True
            End If
        End If
    Next j
    For k = 1 To g_ExNb
        If g_ExQty(k) Then
            g_ExDer(LibPrixDe(g_ExNoms(k))) = k & "|" & K_PRIX
            g_ExDer(LibMontantDe(g_ExNoms(k))) = k & "|" & K_MONTANT
        End If
    Next k
End Sub

Public Function LibPrixDe(ByVal indic As String) As String
    LibPrixDe = "Prix HT " & indic
End Function

Public Function LibMontantDe(ByVal indic As String) As String
    LibMontantDe = "Montant " & indic
End Function

' Rang d'un indicateur créé dans le calcul en cours ; 0 s'il n'en fait pas partie.
Private Function IndexExtra(ByVal indic As String) As Long
    If g_ExDic Is Nothing Then Exit Function
    If g_ExDic.Exists(indic) Then IndexExtra = g_ExDic(indic)
End Function

' Traduit le libellé de la liste Indicateur : groupe de lignes de référence (grp) et genre de valeur (kind).
' Renvoie Faux si le libellé ne correspond à rien pour le type choisi.
Private Function ResoudreIndic(ByVal indic As String, ByRef grp As Long, ByRef kind As Long) As Boolean

    Dim k As Long
    Dim v As Variant

    ResoudreIndic = True
    If StrComp(indic, LibIndQuantite(), vbTextCompare) = 0 Then
        grp = GRP_CONSO: kind = K_QTE
    ElseIf StrComp(indic, LibIndPrix(), vbTextCompare) = 0 Then
        grp = GRP_CONSO: kind = K_PRIX
    ElseIf StrComp(indic, LibIndConso(), vbTextCompare) = 0 Then
        grp = GRP_CONSO: kind = K_MONTANT
    ElseIf StrComp(indic, LibIndAbo(), vbTextCompare) = 0 Then
        grp = GRP_ABO: kind = K_MONTANT
    ElseIf StrComp(indic, LibIndTaxes(), vbTextCompare) = 0 Then
        grp = GRP_TAXE: kind = K_MONTANT
    ElseIf StrComp(indic, LibIndMontant(), vbTextCompare) = 0 Then
        grp = GRP_CONSO: kind = K_TOTAL          ' la consommation reste la référence des colonnes quantité et prix
    Else
        k = IndexExtra(indic)
        If k > 0 Then
            grp = GRP_TAXE + k
            kind = IIf(g_ExQty(k), K_QTE, K_MONTANT)
        ElseIf g_ExDer.Exists(indic) Then
            v = Split(CStr(g_ExDer(indic)), "|")
            grp = GRP_TAXE + CLng(v(0))
            kind = CLng(v(1))
        Else
            ResoudreIndic = False
        End If
    End If

End Function

' Nom du groupe de référence, pour les en-têtes du tableau.
Private Function NomGroupe(ByVal grp As Long) As String
    Select Case grp
        Case GRP_CONSO: NomGroupe = LibIndQuantite()
        Case GRP_ABO: NomGroupe = LibIndAbo()
        Case GRP_TAXE: NomGroupe = LibIndTaxes()
        Case Else: NomGroupe = g_ExNoms(grp - GRP_TAXE)
    End Select
End Function

' Unité de la dernière ligne (par date de fin) du groupe pour le type donné ; pour la consommation, à défaut,
' celle du contrat ou du type.
Private Function UniteGroupe(ByVal typ As String, ByRef t() As TypeLigne, ByVal nb As Long, ByVal grp As Long) As String
    Dim i As Long
    Dim meilleur As Double
    For i = 1 To nb
        If StrComp(t(i).Typ, typ, vbTextCompare) = 0 And t(i).Unite <> "" Then
            If GroupeDeLigne(t(i)) = grp And t(i).Fin >= meilleur Then
                meilleur = t(i).Fin
                UniteGroupe = t(i).Unite
            End If
        End If
    Next i
    If UniteGroupe = "" And grp = GRP_CONSO Then UniteGroupe = UniteDuType(typ, t, nb)
End Function

' Valeur d'un point du graphique pour une série.
Private Function ValeurIndic(ByVal kind As Long, ByVal q As Double, ByVal ht As Double, ByVal ttc As Double, _
                             ByVal total As Double, ByVal has As Boolean) As Variant

    If Not has Then Exit Function       ' Empty : le graphique laisse un trou

    Select Case kind
        Case K_QTE: ValeurIndic = q
        Case K_PRIX
            If q > 0 Then ValeurIndic = ht / q
        Case K_MONTANT: ValeurIndic = ttc
        Case Else: ValeurIndic = total
    End Select

End Function

' Liste du paramètre Type de l'écran : (tous) puis les types de consommation, reposée à chaque actualisation
' (un type ajouté depuis frm_Contrat y apparaît sans réinstallation).
Private Sub PoserListeTypeGraph(ByVal ws As Worksheet, ByVal tps As Variant)

    Dim i As Long
    Dim adr As String

    ws.Range(ws.Cells(2, CO_COL_LISTE_TYPE), ws.Cells(CO_NB_COL_DATES + 1, CO_COL_LISTE_TYPE)).ClearContents
    ws.Cells(2, CO_COL_LISTE_TYPE).Value = LibTous()
    For i = 0 To UBound(tps)
        ws.Cells(3 + i, CO_COL_LISTE_TYPE).Value = tps(i)
    Next i
    adr = "='" & ws.Name & "'!" & ws.Range(ws.Cells(2, CO_COL_LISTE_TYPE), ws.Cells(2 + UBound(tps) + 1, CO_COL_LISTE_TYPE)).Address(True, True)
    On Error Resume Next
    ThisWorkbook.Names("ListeTypeConsoGraph").RefersTo = adr
    If Err.Number <> 0 Then
        Err.Clear
        ThisWorkbook.Names.Add Name:="ListeTypeConsoGraph", RefersTo:=adr
    End If
    On Error GoTo 0

End Sub

' Liste des types de la plage nommée ListeTypesConso (listes déroulantes de frm_Contrat, frm_Contrats et frm_Facture) :
' réécrite depuis TblTypesConso, dans la colonne technique masquée de frm_Facture.
Public Sub RafraichirListeTypes()

    Dim ws As Worksheet
    Dim tps As Variant
    Dim i As Long
    Dim adr As String

    Set ws = mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_FACTURE)
    If ws Is Nothing Then Exit Sub
    tps = TypesConso()
    ws.Range(ws.Cells(1, mod_Factures.FF_COL_AIDE_TYPE), ws.Cells(20, mod_Factures.FF_COL_AIDE_TYPE)).ClearContents
    ws.Cells(1, mod_Factures.FF_COL_AIDE_TYPE).Value = "Types"
    For i = 0 To UBound(tps)
        ws.Cells(2 + i, mod_Factures.FF_COL_AIDE_TYPE).Value = tps(i)
    Next i
    adr = "='" & ws.Name & "'!" & ws.Range(ws.Cells(2, mod_Factures.FF_COL_AIDE_TYPE), ws.Cells(2 + UBound(tps), mod_Factures.FF_COL_AIDE_TYPE)).Address(True, True)
    On Error Resume Next
    ThisWorkbook.Names("ListeTypesConso").RefersTo = adr
    If Err.Number <> 0 Then
        Err.Clear
        ThisWorkbook.Names.Add Name:="ListeTypesConso", RefersTo:=adr
    End If
    On Error GoTo 0

End Sub

' Mois disponibles en base (listes Début et Fin) et mise en forme du tableau.
Private Sub MettreAJourListesConso(ByVal ws As Worksheet, ByRef t() As TypeLigne, ByVal nb As Long)

    Dim i As Long, c As Long, iDeb As Long, iFin As Long, m As Long, n As Long
    Dim dic As Object
    Dim tps As Variant
    Dim tabMois() As Variant, tabMoisDesc() As Variant
    Dim nbMoisType() As Long
    Dim lib As String, bloc As String, entetes As String, prem As String, cType As String, cDeb As String
    Dim blocD As String, premD As String, colD As String, finMax As String
    Dim k As String, colK As String
    Dim retenue As Boolean
    Dim tabInd() As Variant
    Dim aConso As Boolean, aAbo As Boolean, aTaxe As Boolean
    Dim nbInd As Long
    Dim dicInd As Object
    Dim cleInd As Variant
    Dim premI As String, blocI As String

    ' Un bloc de colonnes par choix de type ((tous) puis chaque type) : les mois couverts par les
    ' lignes de factures de ce type, c'est-à-dire la période effectivement saisie (hors trous entre
    ' deux factures). Ils alimentent les listes Début et Fin, qui suivent donc le type choisi.
    tps = TypesConso()
    PoserListeTypeGraph ws, tps
    RafraichirListeTypes
    ReDim g_MinType(0 To CO_NB_COL_DATES - 1)
    ReDim g_MaxType(0 To CO_NB_COL_DATES - 1)
    ReDim tabMois(1 To CO_NB_LISTE_DATES, 1 To CO_NB_COL_DATES)
    ReDim tabMoisDesc(1 To CO_NB_LISTE_DATES, 1 To CO_NB_COL_DATES)
    ReDim nbMoisType(0 To CO_NB_COL_DATES - 1)
    ReDim tabInd(1 To CO_NB_LISTE_INDIC, 1 To CO_NB_COL_DATES)

    For c = 0 To CO_NB_COL_DATES - 1
        Set dic = CreateObject("Scripting.Dictionary")
        g_MinType(c) = -1
        g_MaxType(c) = -1
        aConso = False: aAbo = False: aTaxe = False
        nbInd = 0
        Set dicInd = CreateObject("Scripting.Dictionary")
        dicInd.CompareMode = 1
        For i = 1 To nb
            retenue = (c = 0)
            If Not retenue Then
                If c - 1 <= UBound(tps) Then retenue = (StrComp(t(i).Typ, tps(c - 1), vbTextCompare) = 0)
            End If
            If retenue Then
                Select Case t(i).Nat
                    Case mod_Factures.NAT_CONSO: aConso = True
                    Case mod_Factures.NAT_ABO: aAbo = True
                    Case mod_Factures.NAT_TAXE: aTaxe = True
                    Case Else
                        If Not dicInd.Exists(t(i).Indic) And dicInd.Count < CO_NB_EXTRA Then dicInd.Add t(i).Indic, False
                        If dicInd.Exists(t(i).Indic) Then
                            If UniteQuantifiable(t(i).Unite) Then dicInd(t(i).Indic) = True
                        End If
                End Select
                iDeb = Year(CDate(t(i).Debut)) * 12 + Month(CDate(t(i).Debut)) - 1
                iFin = Year(CDate(t(i).Fin)) * 12 + Month(CDate(t(i).Fin)) - 1
                If iFin - iDeb > 240 Then iFin = iDeb + 240
                For m = iDeb To iFin
                    dic(m) = True
                    If g_MinType(c) < 0 Or m < g_MinType(c) Then g_MinType(c) = m
                    If m > g_MaxType(c) Then g_MaxType(c) = m
                Next m
            End If
        Next i
        If c = 0 Then
            ws.Cells(1, CO_COL_LISTE_DATES).Value = LibTous()
        ElseIf c - 1 <= UBound(tps) Then
            ws.Cells(1, CO_COL_LISTE_DATES + c).Value = tps(c - 1)
        Else
            ws.Cells(1, CO_COL_LISTE_DATES + c).ClearContents      ' colonne de type non utilisée
        End If
        n = 0
        If g_MinType(c) >= 0 Then
            For m = g_MinType(c) To g_MaxType(c)
                If dic.Exists(m) And n < CO_NB_LISTE_DATES Then
                    n = n + 1
                    tabMois(n, c + 1) = DateSerial(m \ 12, (m Mod 12) + 1, 1)
                End If
            Next m
        End If
        nbMoisType(c) = n
        For i = 1 To n
            tabMoisDesc(i, c + 1) = tabMois(n - i + 1, c + 1)
        Next i

        ' Indicateurs proposés pour ce choix de type : seulement ceux qui ont des lignes de facture.
        ' Quantité et prix n'existent que pour un type choisi (les unités diffèrent d'un type à l'autre).
        If c = 0 Then
            ws.Cells(1, CO_COL_LISTE_INDIC_TYPE).Value = LibTous()
        ElseIf c - 1 <= UBound(tps) Then
            ws.Cells(1, CO_COL_LISTE_INDIC_TYPE + c).Value = tps(c - 1)
        Else
            ws.Cells(1, CO_COL_LISTE_INDIC_TYPE + c).ClearContents
        End If
        If aConso Then
            If c > 0 Then
                nbInd = nbInd + 1: tabInd(nbInd, c + 1) = LibIndQuantite()
                nbInd = nbInd + 1: tabInd(nbInd, c + 1) = LibIndPrix()
            End If
            nbInd = nbInd + 1: tabInd(nbInd, c + 1) = LibIndConso()
        End If
        If aAbo Then nbInd = nbInd + 1: tabInd(nbInd, c + 1) = LibIndAbo()
        If aTaxe Then nbInd = nbInd + 1: tabInd(nbInd, c + 1) = LibIndTaxes()
        ' Indicateur créé : sans quantité (m3, kWh), un seul choix (son montant) ; avec quantité, la quantité (pour un type
        ' choisi), son prix HT (idem) et son montant.
        For Each cleInd In dicInd.Keys
            If dicInd(cleInd) Then
                If c > 0 Then
                    nbInd = nbInd + 1: tabInd(nbInd, c + 1) = CStr(cleInd)
                    nbInd = nbInd + 1: tabInd(nbInd, c + 1) = LibPrixDe(CStr(cleInd))
                End If
                nbInd = nbInd + 1: tabInd(nbInd, c + 1) = LibMontantDe(CStr(cleInd))
            Else
                nbInd = nbInd + 1: tabInd(nbInd, c + 1) = CStr(cleInd)
            End If
        Next cleInd
        If aConso Or aAbo Or aTaxe Or dicInd.Count > 0 Then nbInd = nbInd + 1: tabInd(nbInd, c + 1) = LibIndMontant()
    Next c

    With ws.Range(ws.Cells(2, CO_COL_LISTE_INDIC_TYPE), ws.Cells(CO_NB_LISTE_INDIC + 1, CO_COL_LISTE_INDIC_TYPE + CO_NB_COL_DATES - 1))
        .ClearContents
        .Value = tabInd
    End With
    ws.Range(ws.Columns(CO_COL_H_LIB), ws.Columns(CO_COL_LISTE_INDIC_TYPE + CO_NB_COL_DATES - 1)).EntireColumn.Hidden = True

    With ws.Range(ws.Cells(2, CO_COL_LISTE_DATES), ws.Cells(CO_NB_LISTE_DATES + 1, CO_COL_LISTE_DATES + CO_NB_COL_DATES - 1))
        .ClearContents
        .NumberFormat = "mmm yyyy"
        .Value = tabMois
    End With
    With ws.Range(ws.Cells(2, CO_COL_LISTE_DATES_DESC), ws.Cells(CO_NB_LISTE_DATES + 1, CO_COL_LISTE_DATES_DESC + CO_NB_COL_DATES - 1))
        .ClearContents
        .NumberFormat = "mmm yyyy"
        .Value = tabMoisDesc
    End With

    ' Listes reposées à chaque actualisation (un classeur déjà installé les reçoit sans réinstallation).
    ' Elles se recalculent seules quand on change de type ou de début. Sans type choisi, elles sont vides.
    ' Début : tous les mois du type. Fin : les mois du type à partir de Début (au plus CO_NB_MOIS_MAX mois),
    ' ce qui interdit une fin antérieure au début.
    lib = "'" & ws.Name & "'!"
    prem = lib & ws.Cells(2, CO_COL_LISTE_DATES).Address(True, True)
    bloc = lib & ws.Range(ws.Cells(2, CO_COL_LISTE_DATES), ws.Cells(CO_NB_LISTE_DATES + 1, CO_COL_LISTE_DATES + CO_NB_COL_DATES - 1)).Address(True, True)
    entetes = lib & ws.Range(ws.Cells(1, CO_COL_LISTE_DATES), ws.Cells(1, CO_COL_LISTE_DATES + CO_NB_COL_DATES - 1)).Address(True, True)
    cType = lib & ws.Range(CO_ADR_TYPE).Address(True, True)
    cDeb = lib & ws.Range(CO_ADR_DEBUT).Address(True, True)
    k = "IFERROR(MATCH(" & cType & "," & entetes & ",0),0)"            ' rang du type dans le bloc, 0 = aucun
    colK = "INDEX(" & bloc & ",0," & k & ")"                          ' colonne des mois du type
    PoserNomConso "ListeDebutConso", "=IF(" & k & "=0,OFFSET(" & prem & "," & CO_NB_LISTE_DATES & ",0,1,1)," & _
        "OFFSET(" & prem & ",0," & k & "-1,MAX(1,COUNT(" & colK & ")),1))"
    ' Indicateur : uniquement ceux qui ont des lignes pour le type choisi (colonne du type dans le bloc des indicateurs).
    premI = lib & ws.Cells(2, CO_COL_LISTE_INDIC_TYPE).Address(True, True)
    blocI = lib & ws.Range(ws.Cells(2, CO_COL_LISTE_INDIC_TYPE), ws.Cells(CO_NB_LISTE_INDIC + 1, CO_COL_LISTE_INDIC_TYPE + CO_NB_COL_DATES - 1)).Address(True, True)
    PoserNomConso "ListeIndicConso", "=IF(" & k & "=0,OFFSET(" & premI & "," & CO_NB_LISTE_INDIC & ",0,1,1)," & _
        "OFFSET(" & premI & ",0," & k & "-1,MAX(1,COUNTA(INDEX(" & blocI & ",0," & k & "))),1))"
    ' Fin : liste du plus récent au plus ancien (bloc descendant). Les mois au-delà de Début + CO_NB_MOIS_MAX - 1
    ' sont en tête de colonne : on les saute, puis on prend les mois jusqu'à Début compris.
    premD = lib & ws.Cells(2, CO_COL_LISTE_DATES_DESC).Address(True, True)
    blocD = lib & ws.Range(ws.Cells(2, CO_COL_LISTE_DATES_DESC), ws.Cells(CO_NB_LISTE_DATES + 1, CO_COL_LISTE_DATES_DESC + CO_NB_COL_DATES - 1)).Address(True, True)
    colD = "INDEX(" & blocD & ",0," & k & ")"
    finMax = "EDATE(" & cDeb & "," & (CO_NB_MOIS_MAX - 1) & ")"
    PoserNomConso "ListeFinConso", "=IF(" & k & "=0,OFFSET(" & prem & "," & CO_NB_LISTE_DATES & ",0,1,1)," & _
        "OFFSET(" & premD & ",COUNTIF(" & colD & ","">""&" & finMax & ")," & k & "-1," & _
        "MAX(1,COUNTIFS(" & colD & ","">=""&" & cDeb & "," & colD & ",""<=""&" & finMax & ")),1))"

    ' Mise en forme de la zone du tableau jusqu'à CO_NB_LIGNES_MAX lignes.
    mod_InstallCommun.PoserCorpsTableauLecture ws.Range(ws.Cells(CO_LIGNE_ENTETE + 1, 2), ws.Cells(CO_LIGNE_ENTETE + CO_NB_LIGNES_MAX + 1, 1 + CO_NB_COL_TABLE))
    ws.Range(ws.Cells(CO_LIGNE_ENTETE + 1, 2), ws.Cells(CO_LIGNE_ENTETE + CO_NB_LIGNES_MAX + 1, 2)).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(CO_LIGNE_ENTETE + 1, 3), ws.Cells(CO_LIGNE_ENTETE + CO_NB_LIGNES_MAX + 1, 3)).NumberFormat = "#,##0.00"
    ws.Range(ws.Cells(CO_LIGNE_ENTETE + 1, 4), ws.Cells(CO_LIGNE_ENTETE + CO_NB_LIGNES_MAX + 1, 4)).NumberFormat = "#,##0.00000"
    ws.Range(ws.Cells(CO_LIGNE_ENTETE + 1, 5), ws.Cells(CO_LIGNE_ENTETE + CO_NB_LIGNES_MAX + 1, 8)).NumberFormat = "#,##0.00"

End Sub

Private Sub PoserNomConso(ByVal nom As String, ByVal formule As String)
    On Error Resume Next
    ThisWorkbook.Names(nom).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nom, RefersTo:=formule
End Sub

Private Sub ConstruireGraphiqueConso(ByVal ws As Worksheet, ByVal nbM As Long, ByVal nS As Long, _
                                     ByRef serNoms() As String, ByVal indic As String, ByVal indicSerie As String, _
                                     ByVal kindBarres As Long, ByVal forme As String, ByVal typ As String, ByVal tous As Boolean, _
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

    If courbes Then co.Chart.ChartType = xlLine Else co.Chart.ChartType = xlColumnStacked

    For s = 1 To nS
        Set se = co.Chart.SeriesCollection.NewSeries
        se.Name = serNoms(s)
        se.Values = ws.Range(ws.Cells(ent + 1, CO_COL_H_SERIE + s - 1), ws.Cells(ent + nbM, CO_COL_H_SERIE + s - 1))
        se.XValues = xr
        If courbes Then se.ChartType = xlLine Else se.ChartType = xlColumnStacked
    Next s

    If mixte Then
        Set se = co.Chart.SeriesCollection.NewSeries
        se.Name = "Prix unitaire"
        se.Values = ws.Range(ws.Cells(ent + 1, CO_COL_H_PRIX), ws.Cells(ent + nbM, CO_COL_H_PRIX))
        se.XValues = xr
        se.ChartType = xlLine
        se.AxisGroup = xlSecondary
    End If

    co.Chart.HasLegend = True
    co.Chart.Legend.Position = xlLegendPositionBottom
    ' Courbes sans points : un mois sans valeur est enjambé, sinon un opérateur présent un seul mois ne serait pas visible.
    If courbes Or mixte Then co.Chart.DisplayBlanksAs = xlInterpolated Else co.Chart.DisplayBlanksAs = xlNotPlotted

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
    co.Chart.Axes(xlValue, xlPrimary).AxisTitle.Text = TitreAxe(kindBarres, indicSerie, unite)
    If mixte Then
        co.Chart.Axes(xlValue, xlSecondary).HasTitle = True
        co.Chart.Axes(xlValue, xlSecondary).AxisTitle.Text = TitreAxe(K_PRIX, LibIndPrix(), unite)
    End If
    On Error GoTo 0

    If tous Then
        co.Chart.ChartTitle.Text = mod_Display.FR("Tous les types -- ") & indic
    Else
        co.Chart.ChartTitle.Text = typ & " -- " & indic
        If mixte Then co.Chart.ChartTitle.Text = typ & mod_Display.FR(" -- ") & indicSerie & " et prix unitaire"
    End If

End Sub

Private Function TitreAxe(ByVal kind As Long, ByVal indic As String, ByVal unite As String) As String
    If kind = K_QTE Then
        TitreAxe = indic & IIf(unite <> "", " (" & unite & ")", "")
    ElseIf kind = K_PRIX Then
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
    If nomFeuille = mod_VarGlobales.NOM_FEUILLE_FACTURES Then mod_Factures.FigerEnteteListeActive

    noms = Array(mod_VarGlobales.NOM_FEUILLE_CONSO, mod_VarGlobales.NOM_FEUILLE_FACTURES, mod_VarGlobales.NOM_FEUILLE_CONTRATS)
    For i = LBound(noms) To UBound(noms)
        If noms(i) <> nomFeuille Then
            Set autre = mod_InstallCommun.TrouverFeuille(CStr(noms(i)))
            If Not autre Is Nothing Then mod_Aide.MasquerFormulaire autre
        End If
    Next i

    ' Synthese est masquee par MasquerQuadrillage (regle commune des formulaires).

End Sub

' Ouverture depuis la feuille Synthese : l'écran s'ouvre vide (la recherche précédente est oubliée).
Public Sub AfficherConsommations()
    ReinitialiserConsommations
    MontrerEcran mod_VarGlobales.NOM_FEUILLE_CONSO
    If Not FeuilleConso() Is Nothing Then ActualiserConsommations
End Sub

' Remet l'écran des consommations à zéro : plus de type, d'indicateur, de début ni de fin ; forme et répartition par
' défaut ; tableau, graphique et résumé vidés. Appelée à la sortie et à l'ouverture depuis Synthese.
Public Sub ReinitialiserConsommations()

    Dim ws As Worksheet
    Dim evtAvant As Boolean

    Set ws = FeuilleConso()
    If ws Is Nothing Then Exit Sub

    evtAvant = Application.EnableEvents
    On Error GoTo Sortie
    Application.EnableEvents = False
    ws.Range(CO_ADR_TYPE).MergeArea.ClearContents
    ws.Range(CO_ADR_INDIC).MergeArea.ClearContents
    ws.Range(CO_ADR_DEBUT).MergeArea.ClearContents
    ws.Range(CO_ADR_FIN).MergeArea.ClearContents
    ws.Range(CO_ADR_FORME).MergeArea.Value = LibFormeBarres()
    ws.Range(CO_ADR_REPART).MergeArea.Value = LibRepartProrata()
    AppliquerVisibilite ws
    EffacerResultats ws, 0
    ws.Range(CO_ADR_TYPE).Select

Sortie:
    Application.EnableEvents = evtAvant

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

' Diagnostic (Ctrl+G) : état de visibilité de toutes les feuilles et nombre de feuilles sélectionnées.
' Sert à comprendre un refus de masquer une feuille (Excel refuse de masquer la dernière feuille
' visible, ou une feuille quand plusieurs feuilles sont sélectionnées en groupe).
Public Sub DiagnostiquerFeuilles()

    Dim sh As Object
    Dim t As String, etat As String

    For Each sh In ThisWorkbook.Sheets
        Select Case sh.Visible
            Case xlSheetVisible: etat = "VISIBLE"
            Case xlSheetHidden: etat = "masquee"
            Case Else: etat = "tres masquee"
        End Select
        t = t & sh.Name & " : " & etat & vbLf
    Next sh
    t = t & vbLf & "Feuille active : " & ActiveSheet.Name & vbLf & "Feuilles selectionnees : " & ActiveWindow.SelectedSheets.Count
    MsgBox t, vbInformation, "Diagnostic des feuilles"

End Sub

' Sortie vers Synthese (bouton Sortir de l'écran graphique).
Public Sub SortirConsommations()

    Dim ws As Worksheet
    Dim noms As Variant
    Dim i As Long

    EnregistrerListesSiOuvertes
    ReinitialiserConsommations
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
    RafraichirListeTypes
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

' Bouton « + » du champ Type de frm_Contrat : ajoute un type de consommation à la liste (par exemple « Production
' électricité »), avec son unité de mesure par défaut, et le sélectionne dans le formulaire.
Public Sub AjouterTypeConso()

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lr As ListRow
    Dim tps As Variant
    Dim nomType As String, uniteType As String
    Dim i As Long, idxSens As Long
    Dim sensRep As VbMsgBoxResult

    If Not g_CtrEnCours Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_CONTRAT)

    Set lo = TableTypes()
    If lo Is Nothing Then
        MsgBox mod_Display.FR("La liste des types de consommation n'existe pas encore. Ex{e2}cutez d'abord la macro InstallerConsommations."), vbExclamation
        Exit Sub
    End If

    tps = TypesConso()
    If UBound(tps) + 1 >= CO_NB_COL_DATES - 1 Then
        MsgBox mod_Display.FR("La liste est pleine : ") & (CO_NB_COL_DATES - 1) & mod_Display.FR(" types au plus."), vbExclamation, mod_Display.FR("Nouveau type de consommation")
        Exit Sub
    End If

    nomType = Trim$(InputBox(mod_Display.FR("Nom du nouveau type de consommation (par exemple Production {e2}lectricit{e2}). ") & _
                             mod_Display.FR("Il sera propos{e2} dans les listes des contrats et des factures :"), _
                             mod_Display.FR("Nouveau type de consommation")))
    If nomType = "" Then Exit Sub
    nomType = Replace(Replace(nomType, vbCr, " "), vbLf, " ")
    If Len(nomType) > 40 Then
        MsgBox mod_Display.FR("Le nom du type est limit{e2} {a2} 40 caract{e1}res."), vbExclamation, mod_Display.FR("Nouveau type de consommation")
        Exit Sub
    End If
    If StrComp(nomType, LibTous(), vbTextCompare) = 0 Then
        MsgBox mod_Display.FR("Ce nom est r{e2}serv{e2} : choisissez-en un autre."), vbExclamation, mod_Display.FR("Nouveau type de consommation")
        Exit Sub
    End If

    ' Un type déjà connu n'est pas recréé : on le sélectionne simplement (même orthographe que dans la liste).
    For i = 0 To UBound(tps)
        If StrComp(CStr(tps(i)), nomType, vbTextCompare) = 0 Then
            MsgBox mod_Display.FR("Ce type existe d{e2}j{a2} : il est s{e2}lectionn{e2}."), vbInformation, mod_Display.FR("Nouveau type de consommation")
            ws.Range(CF_ADR_TYPE).Value = CStr(tps(i))
            Exit Sub
        End If
    Next i

    uniteType = Trim$(InputBox(mod_Display.FR("Unit{e2} de mesure de ce type (kWh, m") & ChrW(179) & mod_Display.FR(", litres...). Elle sera propos{e2}e par d{e2}faut dans les contrats et les factures :"), _
                               mod_Display.FR("Nouveau type de consommation"), "kWh"))
    If uniteType = "" Then Exit Sub
    uniteType = mod_Unites.DeclarerUnite(uniteType, True)     ' unité nouvelle : ajoutée à la liste (consommation ou non ?)

    ' Sens du type : achat (on paie un fournisseur) ou vente (on encaisse, par exemple la revente d'électricité).
    sensRep = MsgBox(mod_Display.FR("Ce type est-il une REVENTE (vous {e2}mettez les factures et encaissez les paiements) ?") & vbCrLf & vbCrLf & _
                     mod_Display.FR("Oui : vente (les op{e2}rations li{e2}es aux factures sont des encaissements).") & vbCrLf & _
                     mod_Display.FR("Non : achat / consommation (les op{e2}rations li{e2}es sont des paiements).") & vbCrLf & _
                     mod_Display.FR("Annuler : ne pas cr{e2}er le type."), _
                     vbYesNoCancel + vbQuestion + vbDefaultButton2, mod_Display.FR("Nouveau type de consommation"))
    If sensRep = vbCancel Then Exit Sub

    idxSens = ColonneSensType(lo)
    Set lr = NouvelleLigne(lo)
    lr.Range.Cells(1, ColIdx(lo, "TypeConso")).Value = nomType
    lr.Range.Cells(1, ColIdx(lo, "Unite")).Value = uniteType
    lr.Range.Cells(1, idxSens).NumberFormat = "@"
    If sensRep = vbYes Then lr.Range.Cells(1, idxSens).Value = SENS_VENTE Else lr.Range.Cells(1, idxSens).Value = SENS_ACHAT

    RafraichirListeTypes
    ws.Range(CF_ADR_TYPE).Value = nomType       ' déclenche la proposition de l'unité (CtrTraiterChangement)

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
    ' Une unité nouvelle (par exemple stère) est ajoutée à la liste des unités, après avoir demandé si c'est une unité de consommation.
    e.Unite = mod_Unites.DeclarerUnite(e.Unite, True)
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
