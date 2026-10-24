Attribute VB_Name = "mod_Factures"
Option Explicit

' =====================================================================================
' MODULE : mod_Factures
'
' SUIVI DES CONSOMMATIONS - SAISIE DES FACTURES (ajout 19/10/2026). Remplace les anciens
' "relevés" (une seule période commune à toute la facture). La construction des feuilles est
' dans mod_InstallConsommations ; le graphique mensuel est dans mod_Consommations.
'
' MODÈLE (calqué sur la facture du fournisseur) :
'   - TblFactures      : une ligne par facture (type, opérateur, N°, date, période facturée,
'                        total TTC réellement facturé).
'   - TblAbonnements   : les lignes d'ABONNEMENT d'une facture. Chacune a SA période, son
'   - TblConsos        : nombre de jours (abonnement) ou sa quantité, SON prix unitaire HT,
'   - TblTaxes         : SON taux de TVA, et ses montants HT et TTC. Consommation et taxes
'                        ont leurs propres périodes, qui ne coïncident pas avec celles de
'                        l'abonnement ni entre elles.
'   Les quatre tableaux sont sur des feuilles de données masquées (Conso_Factures,
'   Conso_Abonnements, Conso_Consommations, Conso_Taxes), jamais reconstruites. Les lignes
'   sont reliées à leur facture par ID_Facture.
'
' RÈGLES DE CALCUL (identiques à celles des factures de l'opérateur) :
'   - Abonnement : quantité et unité se saisissent comme pour les autres blocs (par exemple 31 jours).
'   - Montant HT d'une ligne = ROUND(quantité x prix unitaire HT ; 2).
'   - Montant TTC d'une ligne = ROUND(HT x (1 + TVA) ; 2).
'   - Le total de la facture est CALCULÉ À TITRE INDICATIF (somme des lignes) et comparé au
'     total facturé saisi : l'écart est affiché mais n'empêche jamais l'enregistrement.
'
' BLOCS LIBRES : en plus des trois blocs fixes (modèle de toute nouvelle facture), l'opérateur peut ajouter des blocs
'   nommés (bouton "Ajouter un bloc" du formulaire), de même forme. Leurs lignes sont enregistrées dans le tableau
'   TblBlocsLibres (feuille Conso_Blocs), avec le nom et le rang du bloc ; elles sont comptées dans les totaux de la
'   facture (récapitulatif, liste) mais pas encore dans le graphique.
'
' SAISIE : formulaire frm_Facture (en-tête + trois blocs de lignes, extensibles),
'   ouvert par SaisirFacture (nouvelle facture) ou ModifierFactureSelection (liste
'   frm_Factures). Les montants du formulaire se calculent en direct (formules).
'
' À PROPOS DES ACCENTS : textes affichés via mod_Display.FR(); commentaires en UTF-8.
' =====================================================================================

' --- Natures de lignes ---
Public Const NAT_ABO As Long = 1
Public Const NAT_CONSO As Long = 2
Public Const NAT_TAXE As Long = 3
Public Const NAT_LIBRE As Long = 4          ' ligne d'un bloc ajouté par l'opérateur (TblBlocsLibres)

' --- Formulaire frm_Facture : en-tête (B = étiquettes, C = saisie) ---
Public Const FF_ADR_TYPE As String = "C6"             ' fusion C6:D6
Public Const FF_ADR_OPERATEUR As String = "C7"        ' fusion C7:E7
Public Const FF_ADR_NUMERO As String = "C8"           ' fusion C8:D8
Public Const FF_ADR_DATE As String = "C9"
Public Const FF_ADR_DEBUT As String = "C10"
Public Const FF_ADR_FIN As String = "C11"
Public Const FF_ADR_TOTAL As String = "C12"
Public Const FF_ADR_COMMENT As String = "C13"         ' fusion C13:J13

' --- Blocs de lignes : colonnes ---
Public Const FF_COL_LIB As Long = 2
Public Const FF_COL_DEB As Long = 3
Public Const FF_COL_FIN As Long = 4
Public Const FF_COL_QTE As Long = 5
Public Const FF_COL_UNITE As Long = 6
Public Const FF_COL_PRIX As Long = 7
Public Const FF_COL_TVA As Long = 8
Public Const FF_COL_HT As Long = 9
Public Const FF_COL_TTC As Long = 10

' --- Récapitulatif (indicatif), placé EN HAUT pour que les blocs puissent grandir vers le bas ---
Public Const FF_RECAP_TITRE As Long = 15
Public Const FF_ADR_RECAP_HT As String = "C16"
Public Const FF_ADR_RECAP_TVA As String = "C17"
Public Const FF_ADR_RECAP_TTC As String = "C18"
Public Const FF_ADR_RECAP_SAISI As String = "C19"
Public Const FF_ADR_RECAP_ECART As String = "C20"
Public Const FF_ADR_MESSAGE As String = "B22"         ' fusion B22:J22
Public Const FF_COL_AIDE_TYPE As Long = 26            ' colonne technique masquée (liste des types)
Public Const FF_COL_AIDE_OPER As Long = 27            ' colonne technique masquée (opérateurs en cours du type choisi)
Public Const FF_COL_AIDE_UNITE As Long = 28           ' colonne technique masquée (unités proposées)
Public Const FF_LARG_LIBELLE As Double = 44           ' largeur de la colonne B (libellés des lignes)

' --- Blocs de lignes : ils commencent ligne 24 et GRANDISSENT (bouton "+" du titre de chaque bloc) ---
' Chaque bloc occupe : titre, en-tête, n lignes, sous-total, séparateur (n + 4 lignes).
Public Const FF_BLOCS_DEBUT As Long = 24
Public Const FF_NB_DEFAUT As Long = 3           ' lignes affichées au départ dans chaque bloc
Public Const FF_MAX_LIGNES As Long = 40         ' limite d'un bloc
Public Const FF_MAX_BLOCS_LIBRES As Long = 6    ' blocs que l'opérateur peut ajouter aux trois blocs fixes
Public Const FF_MAX_BLOCS As Long = 9           ' 3 blocs fixes + FF_MAX_BLOCS_LIBRES
Public Const FF_LONG_NOM_BLOC As Long = 60      ' longueur maximale du nom d'un bloc ajouté

' Modes de ModifierStructure (déclarations de module : elles doivent précéder toute procédure).
Private Const MODE_PLUS As Long = 1
Private Const MODE_MOINS As Long = 2
Private Const MODE_AJOUT_BLOC As Long = 3
Private Const MODE_SUPPR_BLOC As Long = 4

' Une facture telle que saisie dans le formulaire.
Public Type TypeFacture
    ID As Long                  ' 0 = nouvelle facture
    Typ As String
    Operateur As String
    Numero As String
    DateFac As Double           ' 0 = non renseignée
    Debut As Double
    Fin As Double
    TotalSaisi As Double
    TotalRenseigne As Boolean
    Commentaire As String
End Type

' Une ligne de facture (abonnement, consommation ou taxe).
Public Type TypeLigne
    Nat As Long                 ' 1 abonnement, 2 consommation, 3 taxes, 4 bloc libre
    Bloc As String              ' nom du bloc (blocs libres seulement)
    RangBloc As Long            ' rang du bloc libre dans la facture (1, 2...)
    IDFacture As Long
    Typ As String               ' type de la facture (renseigné au chargement)
    Operateur As String         ' opérateur de la facture (renseigné au chargement)
    Libelle As String
    Debut As Double
    Fin As Double
    Qte As Double
    Unite As String
    Prix As Double              ' prix unitaire HT
    Tva As Double               ' 0,055 pour 5,5 %
    HT As Double
    TTC As Double
    Copie As Boolean            ' ligne issue d'une copie : dates, quantité et prix laissés vides
End Type

' --- État du formulaire de saisie ---
Private g_FacEnCours As Boolean
Private g_FacVerrouActif As Boolean
Private g_FacValide As Boolean
Private g_FacRes As TypeFacture
Private g_FacLignes() As TypeLigne
Private g_FacNb As Long
Private g_FacNomFeuillePrec As String
Private g_IdEnEdition As Long
Private g_NbUnites As Long
Private g_AutoCompleteAvant As Boolean
Private g_NbBlocs As Long                     ' blocs affichés : 3 fixes + les blocs libres
Private g_N(1 To FF_MAX_BLOCS) As Long        ' nombre de lignes affichées dans chaque bloc
Private g_NomBloc(1 To FF_MAX_BLOCS) As String ' nom des blocs libres (indices 4 et suivants)


' =====================================================================================
' ACCÈS AUX TABLEAUX
' =====================================================================================
Public Function TableFactures() As ListObject
    Set TableFactures = mod_Consommations.ObtenirTable(mod_VarGlobales.NOM_FEUILLE_DATA_FACTURES, mod_VarGlobales.NOM_TABLE_FACTURES)
End Function

Public Function TableLignes(ByVal nat As Long) As ListObject
    Select Case nat
        Case NAT_ABO
            Set TableLignes = mod_Consommations.ObtenirTable(mod_VarGlobales.NOM_FEUILLE_DATA_ABOS, mod_VarGlobales.NOM_TABLE_ABOS)
        Case NAT_CONSO
            Set TableLignes = mod_Consommations.ObtenirTable(mod_VarGlobales.NOM_FEUILLE_DATA_CONSOS, mod_VarGlobales.NOM_TABLE_CONSOS)
        Case NAT_LIBRE
            Set TableLignes = mod_Consommations.ObtenirTable(mod_VarGlobales.NOM_FEUILLE_DATA_BLOCS, mod_VarGlobales.NOM_TABLE_BLOCS)
        Case Else
            Set TableLignes = mod_Consommations.ObtenirTable(mod_VarGlobales.NOM_FEUILLE_DATA_TAXES, mod_VarGlobales.NOM_TABLE_TAXES)
    End Select
End Function

Private Function TableAffFactures() As ListObject
    Set TableAffFactures = mod_Consommations.ObtenirTable(mod_VarGlobales.NOM_FEUILLE_FACTURES, mod_VarGlobales.NOM_TABLE_FACTURES_AFF)
End Function

Public Function FacturesInstallees() As Boolean
    FacturesInstallees = Not (TableFactures() Is Nothing) And Not (TableLignes(NAT_ABO) Is Nothing) And _
                         Not (TableLignes(NAT_CONSO) Is Nothing) And Not (TableLignes(NAT_TAXE) Is Nothing) And _
                         Not (TableLignes(NAT_LIBRE) Is Nothing)
End Function

' Nombre de lignes affichées dans un bloc, et positions (le formulaire est redessiné quand ce nombre change).
Private Function BlocN(ByVal b As Long) As Long
    BlocN = g_N(b)
End Function

Private Function BlocTitre(ByVal b As Long) As Long
    Dim t As Long, k As Long
    t = FF_BLOCS_DEBUT
    For k = 1 To b - 1
        t = t + g_N(k) + 4
    Next k
    BlocTitre = t
End Function

Private Function BlocL1(ByVal b As Long) As Long
    BlocL1 = BlocTitre(b) + 2
End Function

Private Function NomNature(ByVal nat As Long) As String
    Select Case nat
        Case NAT_ABO: NomNature = "Abonnement"
        Case NAT_CONSO: NomNature = "Consommation"
        Case NAT_LIBRE: NomNature = "Bloc ajout" & ChrW(233)
        Case Else: NomNature = "Taxes et contributions"
    End Select
End Function

' Formules des montants d'une ligne du formulaire (syntaxe anglaise).
Public Function FormuleHT(ByVal ligne As Long) As String
    Dim e As String, g As String
    e = mod_Consommations.AdresseCol(FF_COL_QTE) & ligne
    g = mod_Consommations.AdresseCol(FF_COL_PRIX) & ligne
    FormuleHT = "=IF(AND(ISNUMBER(" & e & "),ISNUMBER(" & g & ")),ROUND(" & e & "*" & g & ",2),"""")"
End Function

Public Function FormuleTTC(ByVal ligne As Long) As String
    Dim i As String, h As String
    i = mod_Consommations.AdresseCol(FF_COL_HT) & ligne
    h = mod_Consommations.AdresseCol(FF_COL_TVA) & ligne
    FormuleTTC = "=IF(ISNUMBER(" & i & "),ROUND(" & i & "*(1+N(" & h & ")),2),"""")"
End Function

' =====================================================================================
' CHARGEMENT DES LIGNES (utilisé par le graphique et par les contrôles)
' =====================================================================================
' Remplit "t" (base 1) avec toutes les lignes valides des tableaux, avec le type et l'opérateur de leur
' facture. Renvoie leur nombre. Si "sauf" > 0, les lignes de cette facture sont ignorées (contrôle de
' chevauchement lors d'une modification). Les lignes des BLOCS LIBRES (TblBlocsLibres) ne sont renvoyées que
' si "inclureLibres" : le graphique ne les utilise pas encore (leur rattachement aux indicateurs reste à décider).
Public Function ChargerLignes(ByRef t() As TypeLigne, Optional ByVal sauf As Long = 0, Optional ByVal inclureLibres As Boolean = False) As Long

    Dim loF As ListObject, lo As ListObject
    Dim dTyp As Object, dOp As Object
    Dim vID() As Variant, vT() As Variant, vO() As Variant
    Dim vI() As Variant, vL() As Variant, vD() As Variant, vF() As Variant, vQ() As Variant
    Dim vU() As Variant, vP() As Variant, vV() As Variant, vH() As Variant, vC() As Variant
    Dim vB() As Variant, vR() As Variant
    Dim i As Long, n As Long, nat As Long, nb As Long, id As Long, natMax As Long
    Dim d As Double, f As Double
    Dim cle As String

    ReDim t(1 To 1)
    Set loF = TableFactures()
    If loF Is Nothing Then Exit Function

    Set dTyp = CreateObject("Scripting.Dictionary")
    Set dOp = CreateObject("Scripting.Dictionary")
    If Not mod_Consommations.TableVide(loF) Then
        vID = mod_Consommations.LireColonne(loF, "ID_Facture")
        vT = mod_Consommations.LireColonne(loF, "TypeConso")
        vO = mod_Consommations.LireColonne(loF, "Operateur")
        For i = 1 To loF.ListRows.Count
            If IsNumeric(vID(i)) And Not IsEmpty(vID(i)) Then
                cle = CStr(CLng(vID(i)))
                dTyp(cle) = mod_DataStructure.CellText(vT(i))
                dOp(cle) = mod_DataStructure.CellText(vO(i))
            End If
        Next i
    End If

    If inclureLibres Then natMax = NAT_LIBRE Else natMax = NAT_TAXE
    For nat = NAT_ABO To natMax
        Set lo = TableLignes(nat)
        If Not lo Is Nothing Then
            If Not mod_Consommations.TableVide(lo) Then
                n = lo.ListRows.Count
                vI = mod_Consommations.LireColonne(lo, "ID_Facture")
                vL = mod_Consommations.LireColonne(lo, "Libelle")
                vD = mod_Consommations.LireColonne(lo, "DateDebut")
                vF = mod_Consommations.LireColonne(lo, "DateFin")
                vQ = mod_Consommations.LireColonne(lo, "Quantite")
                vU = mod_Consommations.LireColonne(lo, "Unite")
                vP = mod_Consommations.LireColonne(lo, "PrixUnitaireHT")
                vV = mod_Consommations.LireColonne(lo, "TauxTVA")
                vH = mod_Consommations.LireColonne(lo, "MontantHT")
                vC = mod_Consommations.LireColonne(lo, "MontantTTC")
                If nat = NAT_LIBRE Then
                    vB = mod_Consommations.LireColonne(lo, "Bloc")
                    vR = mod_Consommations.LireColonne(lo, "RangBloc")
                End If
                For i = 1 To n
                    If IsNumeric(vI(i)) And Not IsEmpty(vI(i)) Then
                        id = CLng(vI(i))
                        cle = CStr(id)
                        d = mod_Consommations.VersJour(vD(i))
                        f = mod_Consommations.VersJour(vF(i))
                        If dTyp.Exists(cle) And d > 0 And f >= d And id <> sauf Then
                            nb = nb + 1
                            If nb > UBound(t) Then ReDim Preserve t(1 To nb + 63)
                            With t(nb)
                                .Nat = nat
                                If nat = NAT_LIBRE Then
                                    .Bloc = mod_DataStructure.CellText(vB(i))
                                    .RangBloc = mod_DataStructure.ToLong(vR(i))
                                End If
                                .IDFacture = id
                                .Typ = dTyp(cle)
                                .Operateur = dOp(cle)
                                .Libelle = mod_DataStructure.CellText(vL(i))
                                .Debut = d
                                .Fin = f
                                .Qte = mod_DataStructure.ToDouble(vQ(i))
                                .Unite = mod_DataStructure.CellText(vU(i))
                                .Prix = mod_DataStructure.ToDouble(vP(i))
                                .Tva = mod_DataStructure.ToDouble(vV(i))
                                .HT = mod_DataStructure.ToDouble(vH(i))
                                .TTC = mod_DataStructure.ToDouble(vC(i))
                            End With
                        End If
                    End If
                Next i
            End If
        End If
    Next nat

    ChargerLignes = nb

End Function


' Arrondi à 2 décimales comme la fonction ARRONDI d'Excel (le Round de VBA arrondit au pair).
Public Function Arrondi(ByVal x As Double) As Double
    Arrondi = Application.WorksheetFunction.Round(x, 2)
End Function


' =====================================================================================
' POINTS D'ENTRÉE (boutons)
' =====================================================================================
Public Sub SaisirFacture()

    Dim f As TypeFacture

    mod_Consommations.EnregistrerListesSiOuvertes
    If Not FacturesInstallees() Then
        MsgBox mod_Display.FR("Le suivi des consommations n'est pas install{e2}. Ex{e2}cutez d'abord la macro InstallerConsommations."), vbExclamation
        Exit Sub
    End If

    g_FacNb = 0
    ReDim g_FacLignes(1 To 1)
    f.ID = 0
    TerminerSaisie f

End Sub

' Modifie la facture sélectionnée dans la liste frm_Factures.
Public Sub ModifierFactureSelection()

    Dim lo As ListObject
    Dim id As Long
    Dim f As TypeFacture

    If Not FacturesInstallees() Then Exit Sub
    id = IdFactureSelectionnee(lo)
    If id = 0 Then Exit Sub
    If Not ChargerFacture(id, f) Then
        MsgBox mod_Display.FR("Cette facture est introuvable."), vbExclamation
        Exit Sub
    End If
    TerminerSaisie f

End Sub

' Copie la facture sélectionnée pour en saisir une nouvelle : mêmes type, opérateur, libellés, unités et taux
' de TVA ; numéro, dates, total, commentaire, quantités et prix sont laissés vides (ils changent à chaque facture).
Public Sub CopierFactureSelection()

    Dim lo As ListObject
    Dim id As Long, k As Long
    Dim f As TypeFacture

    If Not FacturesInstallees() Then Exit Sub
    id = IdFactureSelectionnee(lo)
    If id = 0 Then Exit Sub
    If Not ChargerFacture(id, f) Then
        MsgBox mod_Display.FR("Cette facture est introuvable."), vbExclamation
        Exit Sub
    End If

    f.ID = 0
    f.Numero = ""
    f.DateFac = 0
    f.Debut = 0
    f.Fin = 0
    f.TotalSaisi = 0
    f.TotalRenseigne = False
    f.Commentaire = ""
    For k = 1 To g_FacNb
        g_FacLignes(k).IDFacture = 0
        g_FacLignes(k).Debut = 0
        g_FacLignes(k).Fin = 0
        g_FacLignes(k).Qte = 0
        g_FacLignes(k).Prix = 0
        g_FacLignes(k).HT = 0
        g_FacLignes(k).TTC = 0
        g_FacLignes(k).Copie = True
    Next k
    TerminerSaisie f

End Sub

Private Sub TerminerSaisie(ByRef f As TypeFacture)

    If OuvrirFormFacture(f) Then
        If EnregistrerFacture(g_FacRes) Then
            RafraichirListeFactures
            If ActiveSheet.Name = mod_VarGlobales.NOM_FEUILLE_CONSO Then mod_Consommations.ActualiserConsommations
            MsgBox mod_Display.FR("Facture enregistr{e2}e."), vbInformation, mod_Display.FR("Suivi des consommations")
        End If
    End If

End Sub

' Numéro de la facture sélectionnée dans la liste (0 + message si aucune).
Private Function IdFactureSelectionnee(ByRef lo As ListObject) As Long

    Dim inter As Range, c As Range
    Dim lignes As Object
    Dim n As Long
    Dim v As Variant

    Set lo = TableAffFactures()
    If lo Is Nothing Then Exit Function
    If ActiveSheet.Name <> mod_VarGlobales.NOM_FEUILLE_FACTURES Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    Set inter = SelectionVisible(lo)
    If inter Is Nothing Then
        MsgBox mod_Display.FR("S{e2}lectionnez d'abord une ligne du tableau."), vbInformation
        Exit Function
    End If

    ' Une seule ligne (avec un filtre actif, la sélection peut compter plusieurs zones).
    Set lignes = CreateObject("Scripting.Dictionary")
    For Each c In inter.Cells
        If Not lignes.Exists(c.Row) Then lignes.Add c.Row, True
    Next c
    If lignes.Count > 1 Then
        MsgBox mod_Display.FR("S{e2}lectionnez une seule facture."), vbInformation
        Exit Function
    End If

    n = inter.Cells(1, 1).Row - lo.DataBodyRange.Row + 1
    v = lo.ListColumns("ID_Facture").DataBodyRange.Cells(n, 1).Value2
    If IsNumeric(v) And Not IsEmpty(v) Then IdFactureSelectionnee = CLng(v)

End Function

' Cellules VISIBLES de la sélection dans le tableau de la liste. Avec un filtre actif, une sélection qui
' traverse des lignes masquées les contient aussi : sans cette précaution, une suppression toucherait
' des factures que l'opérateur ne voit pas.
Private Function SelectionVisible(ByVal lo As ListObject) As Range

    Dim inter As Range

    On Error Resume Next
    Set inter = Intersect(Selection, lo.DataBodyRange)
    On Error GoTo 0
    If inter Is Nothing Then Exit Function

    If inter.Cells.Count = 1 Then
        Set SelectionVisible = inter        ' SpecialCells sur une seule cellule porterait sur toute la feuille
    Else
        On Error Resume Next
        Set SelectionVisible = inter.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
    End If

End Function

' Charge une facture et ses lignes dans g_FacLignes / g_FacNb.
Private Function ChargerFacture(ByVal id As Long, ByRef f As TypeFacture) As Boolean

    Dim lo As ListObject
    Dim i As Long, n As Long, nat As Long
    Dim v As Variant

    Set lo = TableFactures()
    If lo Is Nothing Then Exit Function
    If mod_Consommations.TableVide(lo) Then Exit Function

    n = lo.ListRows.Count
    For i = 1 To n
        v = lo.ListColumns("ID_Facture").DataBodyRange.Cells(i, 1).Value2
        If IsNumeric(v) And Not IsEmpty(v) Then
            If CLng(v) = id Then
                With lo.ListRows(i).Range
                    f.ID = id
                    f.Typ = mod_DataStructure.CellText(.Cells(1, mod_Consommations.ColIdx(lo, "TypeConso")).Value2)
                    f.Operateur = mod_DataStructure.CellText(.Cells(1, mod_Consommations.ColIdx(lo, "Operateur")).Value2)
                    f.Numero = mod_DataStructure.CellText(.Cells(1, mod_Consommations.ColIdx(lo, "NumeroFacture")).Value2)
                    f.DateFac = mod_Consommations.VersJour(.Cells(1, mod_Consommations.ColIdx(lo, "DateFacture")).Value2)
                    f.Debut = mod_Consommations.VersJour(.Cells(1, mod_Consommations.ColIdx(lo, "DebutPeriode")).Value2)
                    f.Fin = mod_Consommations.VersJour(.Cells(1, mod_Consommations.ColIdx(lo, "FinPeriode")).Value2)
                    v = .Cells(1, mod_Consommations.ColIdx(lo, "TotalFactureTTC")).Value2
                    If mod_Consommations.SaisieNonVide(v) Then
                        f.TotalSaisi = mod_DataStructure.ToDouble(v)
                        f.TotalRenseigne = True
                    End If
                    f.Commentaire = mod_DataStructure.CellText(.Cells(1, mod_Consommations.ColIdx(lo, "Commentaire")).Value2)
                End With
                ChargerFacture = True
                Exit For
            End If
        End If
    Next i
    If Not ChargerFacture Then Exit Function

    ' Lignes de la facture.
    g_FacNb = 0
    ReDim g_FacLignes(1 To FF_MAX_BLOCS * FF_MAX_LIGNES)
    For nat = NAT_ABO To NAT_LIBRE
        Set lo = TableLignes(nat)
        If Not lo Is Nothing Then
            If Not mod_Consommations.TableVide(lo) Then
                For i = 1 To lo.ListRows.Count
                    v = lo.ListColumns("ID_Facture").DataBodyRange.Cells(i, 1).Value2
                    If IsNumeric(v) And Not IsEmpty(v) Then
                        If CLng(v) = id Then
                            g_FacNb = g_FacNb + 1
                            If g_FacNb > UBound(g_FacLignes) Then ReDim Preserve g_FacLignes(1 To g_FacNb + 20)
                            With lo.ListRows(i).Range
                                g_FacLignes(g_FacNb).Nat = nat
                                g_FacLignes(g_FacNb).IDFacture = id
                                g_FacLignes(g_FacNb).Libelle = mod_DataStructure.CellText(.Cells(1, mod_Consommations.ColIdx(lo, "Libelle")).Value2)
                                g_FacLignes(g_FacNb).Debut = mod_Consommations.VersJour(.Cells(1, mod_Consommations.ColIdx(lo, "DateDebut")).Value2)
                                g_FacLignes(g_FacNb).Fin = mod_Consommations.VersJour(.Cells(1, mod_Consommations.ColIdx(lo, "DateFin")).Value2)
                                g_FacLignes(g_FacNb).Qte = mod_DataStructure.ToDouble(.Cells(1, mod_Consommations.ColIdx(lo, "Quantite")).Value2)
                                g_FacLignes(g_FacNb).Unite = mod_DataStructure.CellText(.Cells(1, mod_Consommations.ColIdx(lo, "Unite")).Value2)
                                g_FacLignes(g_FacNb).Prix = mod_DataStructure.ToDouble(.Cells(1, mod_Consommations.ColIdx(lo, "PrixUnitaireHT")).Value2)
                                g_FacLignes(g_FacNb).Tva = mod_DataStructure.ToDouble(.Cells(1, mod_Consommations.ColIdx(lo, "TauxTVA")).Value2)
                                If nat = NAT_LIBRE Then
                                    g_FacLignes(g_FacNb).Bloc = mod_DataStructure.CellText(.Cells(1, mod_Consommations.ColIdx(lo, "Bloc")).Value2)
                                    g_FacLignes(g_FacNb).RangBloc = mod_DataStructure.ToLong(.Cells(1, mod_Consommations.ColIdx(lo, "RangBloc")).Value2)
                                End If
                            End With
                        End If
                    End If
                Next i
            End If
        End If
    Next nat

End Function


' =====================================================================================
' ENREGISTREMENT
' =====================================================================================
Private Function EnregistrerFacture(ByRef f As TypeFacture) As Boolean

    Dim lo As ListObject
    Dim lr As ListRow
    Dim i As Long, nat As Long

    Set lo = TableFactures()
    If lo Is Nothing Then Exit Function

    ' Tous les tableaux nécessaires doivent exister AVANT de supprimer quoi que ce soit : l'enregistrement d'une
    ' facture modifiée remplace ses anciennes lignes, un échec en cours de route les ferait perdre.
    For i = 1 To g_FacNb
        If TableLignes(g_FacLignes(i).Nat) Is Nothing Then
            MsgBox mod_Display.FR("Le tableau des lignes de type '") & NomNature(g_FacLignes(i).Nat) & mod_Display.FR("' est introuvable : relancez InstallerConsommations. Rien n'a {e2}t{e2} enregistr{e2}."), vbExclamation
            Exit Function
        End If
    Next i

    If f.ID = 0 Then
        f.ID = mod_Consommations.ProchainID(lo, "ID_Facture")
    Else
        SupprimerDonneesFacture f.ID
    End If

    Set lr = mod_Consommations.NouvelleLigne(lo)
    With lr.Range
        .Cells(1, mod_Consommations.ColIdx(lo, "ID_Facture")).Value = f.ID
        .Cells(1, mod_Consommations.ColIdx(lo, "TypeConso")).Value = f.Typ
        .Cells(1, mod_Consommations.ColIdx(lo, "Operateur")).Value = f.Operateur
        .Cells(1, mod_Consommations.ColIdx(lo, "NumeroFacture")).Value = f.Numero
        If f.DateFac > 0 Then .Cells(1, mod_Consommations.ColIdx(lo, "DateFacture")).Value = CDate(f.DateFac)
        If f.Debut > 0 Then .Cells(1, mod_Consommations.ColIdx(lo, "DebutPeriode")).Value = CDate(f.Debut)
        If f.Fin > 0 Then .Cells(1, mod_Consommations.ColIdx(lo, "FinPeriode")).Value = CDate(f.Fin)
        If f.TotalRenseigne Then .Cells(1, mod_Consommations.ColIdx(lo, "TotalFactureTTC")).Value = f.TotalSaisi
        .Cells(1, mod_Consommations.ColIdx(lo, "Commentaire")).Value = f.Commentaire
    End With

    For i = 1 To g_FacNb
        Set lo = TableLignes(g_FacLignes(i).Nat)
        If lo Is Nothing Then Exit Function
        Set lr = mod_Consommations.NouvelleLigne(lo)
        With lr.Range
            .Cells(1, mod_Consommations.ColIdx(lo, "ID_Facture")).Value = f.ID
            .Cells(1, mod_Consommations.ColIdx(lo, "Libelle")).Value = g_FacLignes(i).Libelle
            .Cells(1, mod_Consommations.ColIdx(lo, "DateDebut")).Value = CDate(g_FacLignes(i).Debut)
            .Cells(1, mod_Consommations.ColIdx(lo, "DateFin")).Value = CDate(g_FacLignes(i).Fin)
            .Cells(1, mod_Consommations.ColIdx(lo, "Quantite")).Value = g_FacLignes(i).Qte
            .Cells(1, mod_Consommations.ColIdx(lo, "Unite")).Value = g_FacLignes(i).Unite
            .Cells(1, mod_Consommations.ColIdx(lo, "PrixUnitaireHT")).Value = g_FacLignes(i).Prix
            .Cells(1, mod_Consommations.ColIdx(lo, "TauxTVA")).Value = g_FacLignes(i).Tva
            .Cells(1, mod_Consommations.ColIdx(lo, "MontantHT")).Value = g_FacLignes(i).HT
            .Cells(1, mod_Consommations.ColIdx(lo, "MontantTTC")).Value = g_FacLignes(i).TTC
            If g_FacLignes(i).Nat = NAT_LIBRE Then
                .Cells(1, mod_Consommations.ColIdx(lo, "Bloc")).Value = g_FacLignes(i).Bloc
                .Cells(1, mod_Consommations.ColIdx(lo, "RangBloc")).Value = g_FacLignes(i).RangBloc
            End If
        End With
    Next i

    EnregistrerFacture = True

End Function

' Supprime, dans tous les tableaux (en-tête, trois natures de lignes, blocs libres), la facture et ses lignes.
Private Sub SupprimerDonneesFacture(ByVal id As Long)

    Dim lo As ListObject
    Dim i As Long, nat As Long
    Dim v As Variant

    For nat = 0 To NAT_LIBRE
        If nat = 0 Then Set lo = TableFactures() Else Set lo = TableLignes(nat)
        If Not lo Is Nothing Then
            If Not mod_Consommations.TableVide(lo) Then
                For i = lo.ListRows.Count To 1 Step -1
                    v = lo.ListColumns("ID_Facture").DataBodyRange.Cells(i, 1).Value2
                    If IsNumeric(v) And Not IsEmpty(v) Then
                        If CLng(v) = id Then lo.ListRows(i).Delete
                    End If
                Next i
            End If
        End If
    Next nat

End Sub

Public Sub SupprimerFactureSelection()

    Dim lo As ListObject
    Dim inter As Range, c As Range
    Dim ids As Object
    Dim n As Long, i As Long, id As Long
    Dim v As Variant

    Set lo = TableAffFactures()
    If lo Is Nothing Then Exit Sub
    If ActiveSheet.Name <> mod_VarGlobales.NOM_FEUILLE_FACTURES Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Set inter = SelectionVisible(lo)
    If inter Is Nothing Then
        MsgBox mod_Display.FR("S{e2}lectionnez d'abord une ou plusieurs lignes du tableau."), vbInformation
        Exit Sub
    End If

    Set ids = CreateObject("Scripting.Dictionary")
    For Each c In inter.Cells
        n = c.Row - lo.DataBodyRange.Row + 1
        v = lo.ListColumns("ID_Facture").DataBodyRange.Cells(n, 1).Value2
        If IsNumeric(v) And Not IsEmpty(v) Then
            If Not ids.Exists(CLng(v)) Then ids.Add CLng(v), True
        End If
    Next c
    If ids.Count = 0 Then Exit Sub

    If MsgBox(ids.Count & mod_Display.FR(" facture(s) et toutes leurs lignes (abonnement, consommation, taxes) {a2} supprimer ?"), _
              vbYesNo + vbQuestion, mod_Display.FR("Supprimer des factures")) = vbNo Then Exit Sub

    For Each v In ids.Keys
        SupprimerDonneesFacture CLng(v)
    Next v
    RafraichirListeFactures

End Sub


' =====================================================================================
' LISTE DES FACTURES (frm_Factures) : copie de lecture des données, recalculée à l'ouverture
' =====================================================================================
' Calcule le total HT et TTC d'une facture à partir de ses lignes.
Private Sub TotauxFacture(ByRef t() As TypeLigne, ByVal nb As Long, ByVal id As Long, ByRef ht As Double, ByRef ttc As Double)
    Dim i As Long
    ht = 0: ttc = 0
    For i = 1 To nb
        If t(i).IDFacture = id Then
            ht = ht + t(i).HT
            ttc = ttc + t(i).TTC
        End If
    Next i
End Sub

Public Sub RafraichirListeFactures()

    Dim src As ListObject, dst As ListObject
    Dim t() As TypeLigne
    Dim nbL As Long, n As Long, i As Long, p As Long, s As Long
    Dim donnees() As Variant
    Dim ht As Double, ttc As Double
    Dim id As Long
    Dim vTot As Variant
    Dim nbCol As Long
    Dim ordre() As Long
    Dim filtres As Variant

    Set src = TableFactures()
    Set dst = TableAffFactures()
    If src Is Nothing Or dst Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    ' Les filtres de colonne (flèches d'Excel) sont mémorisés puis levés pendant la reconstruction du tableau,
    ' et réappliqués à la fin : l'opérateur retrouve sa vue filtrée après chaque rafraîchissement.
    If Not dst.ShowAutoFilter Then dst.ShowAutoFilter = True
    filtres = MemoriserFiltres(dst)
    On Error Resume Next
    If dst.AutoFilter.FilterMode Then dst.AutoFilter.ShowAllData
    On Error GoTo 0

    Do While dst.ListRows.Count > 1
        dst.ListRows(dst.ListRows.Count).Delete
    Loop
    dst.DataBodyRange.ClearContents

    If Not mod_Consommations.TableVide(src) Then
        nbL = ChargerLignes(t, 0, True)
        n = src.ListRows.Count
        OrdreParDateDecroissante src, ordre
        If n > 1 Then dst.Resize dst.Range.Resize(n + 1)
        nbCol = dst.ListColumns.Count
        ReDim donnees(1 To n, 1 To nbCol)
        For i = 1 To n
            s = ordre(i)
            id = 0
            If IsNumeric(src.ListColumns("ID_Facture").DataBodyRange.Cells(s, 1).Value2) Then
                id = mod_DataStructure.ToLong(src.ListColumns("ID_Facture").DataBodyRange.Cells(s, 1).Value2)
            End If
            TotauxFacture t, nbL, id, ht, ttc
            For p = 1 To nbCol
                Select Case dst.ListColumns(p).Name
                    Case "TotalHT": donnees(i, p) = ht
                    Case "TotalTVA": donnees(i, p) = ttc - ht
                    Case "TotalTTCCalcule": donnees(i, p) = ttc
                    Case "Ecart"
                        vTot = src.ListColumns("TotalFactureTTC").DataBodyRange.Cells(s, 1).Value2
                        If mod_Consommations.SaisieNonVide(vTot) Then donnees(i, p) = Arrondi(mod_DataStructure.ToDouble(vTot) - ttc)
                    Case Else
                        If mod_Consommations.ColIdx(src, dst.ListColumns(p).Name) > 0 Then
                            donnees(i, p) = src.ListColumns(dst.ListColumns(p).Name).DataBodyRange.Cells(s, 1).Value2
                        End If
                End Select
            Next p
        Next i
        dst.DataBodyRange.Value2 = donnees
    End If

    mod_Consommations.MettreEnFormeListeFactures dst
    RestaurerFiltres dst, filtres
    Application.ScreenUpdating = True

End Sub

' Numéros de ligne du tableau source classés par DateFacture décroissante (les plus récentes en haut) ;
' les factures sans date en bas, à égalité le numéro de facture le plus grand d'abord.
Private Sub OrdreParDateDecroissante(ByVal src As ListObject, ByRef ordre() As Long)

    Dim n As Long, i As Long, j As Long, k As Long
    Dim dates() As Double, ids() As Double
    Dim vD As Variant, vI As Variant

    n = src.ListRows.Count
    ReDim ordre(1 To n)
    ReDim dates(1 To n)
    ReDim ids(1 To n)
    vD = mod_Consommations.LireColonne(src, "DateFacture")
    vI = mod_Consommations.LireColonne(src, "ID_Facture")
    For i = 1 To n
        dates(i) = mod_Consommations.VersJour(vD(i))
        If IsNumeric(vI(i)) And Not IsEmpty(vI(i)) Then ids(i) = CDbl(vI(i))
        ordre(i) = i
    Next i

    ' Tri par insertion (quelques centaines de factures au plus).
    For i = 2 To n
        k = ordre(i)
        j = i - 1
        Do While j >= 1
            If dates(ordre(j)) > dates(k) Or (dates(ordre(j)) = dates(k) And ids(ordre(j)) > ids(k)) Then Exit Do
            ordre(j + 1) = ordre(j)
            j = j - 1
        Loop
        ordre(j + 1) = k
    Next i

End Sub

' Mémorise les filtres de colonne actifs d'un tableau : une colonne du résultat par filtre
' (champ, opérateur, critère 1, critère 2) ; Empty si aucun. Même principe que dans mod_RechercheOperations.
Private Function MemoriserFiltres(ByVal tbl As ListObject) As Variant

    Dim res() As Variant
    Dim nb As Long, f As Long
    Dim fc As Filter

    If Not tbl.ShowAutoFilter Then Exit Function
    For f = 1 To tbl.ListColumns.Count
        Set fc = Nothing
        On Error Resume Next
        Set fc = tbl.AutoFilter.Filters(f)
        On Error GoTo 0
        If Not fc Is Nothing Then
            If fc.On Then
                nb = nb + 1
                ReDim Preserve res(1 To 4, 1 To nb)
                res(1, nb) = f
                res(2, nb) = fc.Operator
                res(3, nb) = fc.Criteria1
                res(4, nb) = Empty
                On Error Resume Next
                res(4, nb) = fc.Criteria2
                On Error GoTo 0
            End If
        End If
    Next f
    If nb > 0 Then MemoriserFiltres = res

End Function

Private Sub RestaurerFiltres(ByVal tbl As ListObject, ByVal filtres As Variant)

    Dim f As Long
    Dim crit As Variant
    Dim op As Long

    If IsEmpty(filtres) Then Exit Sub
    On Error Resume Next
    For f = 1 To UBound(filtres, 2)
        crit = filtres(3, f)
        op = filtres(2, f)
        If op = xlFilterValues Then
            If Not IsArray(crit) Then crit = Array(crit)
        End If
        If op = 0 Then
            tbl.Range.AutoFilter Field:=filtres(1, f), Criteria1:=crit
        ElseIf IsEmpty(filtres(4, f)) Then
            tbl.Range.AutoFilter Field:=filtres(1, f), Criteria1:=crit, Operator:=op
        Else
            tbl.Range.AutoFilter Field:=filtres(1, f), Criteria1:=crit, Operator:=op, Criteria2:=filtres(4, f)
        End If
    Next f
    On Error GoTo 0

End Sub


' =====================================================================================
' FORMULAIRE frm_Facture : OUVERTURE, VALIDATION, FERMETURE
' =====================================================================================
' Ouvre le formulaire prérempli avec "f" et g_FacLignes, et attend sa fermeture.
' Renvoie Vrai si l'opérateur a validé (le résultat est dans g_FacRes / g_FacLignes).
Private Function OuvrirFormFacture(ByRef f As TypeFacture) As Boolean

    Dim ws As Worksheet

    On Error GoTo Erreur

    Set ws = mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_FACTURE)
    If ws Is Nothing Then
        MsgBox mod_Display.FR("La feuille '") & mod_VarGlobales.NOM_FEUILLE_FACTURE & mod_Display.FR("' est introuvable.") & vbCrLf & _
               mod_Display.FR("Ex{e2}cutez d'abord la macro InstallerConsommations."), vbExclamation
        Exit Function
    End If

    g_FacNomFeuillePrec = ActiveSheet.Name
    g_FacValide = False
    g_IdEnEdition = f.ID
    ws.Visible = xlSheetVisible
    ws.Activate
    mod_InstallCommun.MasquerQuadrillage
    Application.ScreenUpdating = False
    RemplirFormulaire ws, f
    Application.ScreenUpdating = True
    ws.Range(FF_ADR_TYPE).Select


    Application.EnableEvents = True
    ' La saisie semi-automatique d'Excel complèterait "M" en "m³" (unité déjà présente dans la colonne) :
    ' elle est coupée pendant la saisie de la facture puis rétablie telle qu'elle était.
    g_AutoCompleteAvant = Application.EnableAutoComplete
    Application.EnableAutoComplete = False
    g_FacEnCours = True
    g_FacVerrouActif = True
    Do While g_FacEnCours
        DoEvents
    Loop
    Application.EnableAutoComplete = g_AutoCompleteAvant

    On Error Resume Next
    mod_InstallCommun.ActiverFeuille g_FacNomFeuillePrec
    On Error GoTo 0

    OuvrirFormFacture = g_FacValide
    Exit Function

Erreur:
    Application.EnableEvents = True
    Application.EnableAutoComplete = g_AutoCompleteAvant
    g_FacEnCours = False
    g_FacVerrouActif = False
    MsgBox mod_Display.FR("Erreur inattendue dans la saisie de la facture :") & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
    OuvrirFormFacture = False

End Function

Private Sub RemplirFormulaire(ByVal ws As Worksheet, ByRef f As TypeFacture)

    Dim evenementsAvant As Boolean
    Dim numErr As Long
    Dim k As Long, b As Long, r As Long
    Dim compte(1 To FF_MAX_BLOCS) As Long

    evenementsAvant = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Sortie

    ws.Range(FF_ADR_TYPE).Value = f.Typ
    ws.Range(FF_ADR_OPERATEUR).Value = f.Operateur
    ws.Range(FF_ADR_NUMERO).Value = f.Numero
    EcrireDate ws.Range(FF_ADR_DATE), f.DateFac
    EcrireDate ws.Range(FF_ADR_DEBUT), f.Debut
    EcrireDate ws.Range(FF_ADR_FIN), f.Fin
    If f.TotalRenseigne Then ws.Range(FF_ADR_TOTAL).Value = f.TotalSaisi Else ws.Range(FF_ADR_TOTAL).ClearContents
    ws.Range(FF_ADR_COMMENT).Value = f.Commentaire
    ws.Range(FF_ADR_MESSAGE).Value = ""
    PoserListeType ws
    MettreAJourListeOperateurs ws, f.Typ

    ' Blocs : les trois blocs fixes + les blocs libres retrouvés dans les lignes de la facture (dans l'ordre
    ' où ils ont été enregistrés). Chaque bloc affiche au moins FF_NB_DEFAUT lignes.
    g_NbBlocs = 3
    For b = 1 To FF_MAX_BLOCS
        g_NomBloc(b) = ""
        compte(b) = 0
    Next b
    For k = 1 To g_FacNb
        b = BlocDeLigne(g_FacLignes(k), True)
        If b > 0 Then compte(b) = compte(b) + 1
    Next k
    For b = 1 To g_NbBlocs
        g_N(b) = PlusGrand(compte(b), FF_NB_DEFAUT)
        compte(b) = 0
    Next b
    DessinerBlocs ws

    ' Les lignes de la facture sont réparties dans les blocs.
    For k = 1 To g_FacNb
        b = BlocDeLigne(g_FacLignes(k), False)
        If b > 0 Then
            compte(b) = compte(b) + 1
            r = BlocL1(b) + compte(b) - 1
            ws.Cells(r, FF_COL_LIB).Value = g_FacLignes(k).Libelle
            ws.Cells(r, FF_COL_UNITE).Value = g_FacLignes(k).Unite
            ws.Cells(r, FF_COL_TVA).Value = g_FacLignes(k).Tva
            If Not g_FacLignes(k).Copie Then
                ws.Cells(r, FF_COL_DEB).Value = CDate(g_FacLignes(k).Debut)
                ws.Cells(r, FF_COL_FIN).Value = CDate(g_FacLignes(k).Fin)
                ws.Cells(r, FF_COL_QTE).Value = g_FacLignes(k).Qte
                ws.Cells(r, FF_COL_PRIX).Value = g_FacLignes(k).Prix
            End If
        End If
    Next k

    RemplirTvaDefaut ws, f.Typ

    mod_InstallCommun.PoserValidationDate ws.Range(FF_ADR_DATE)
    mod_InstallCommun.PoserValidationDate ws.Range(FF_ADR_DEBUT)
    mod_InstallCommun.PoserValidationDate ws.Range(FF_ADR_FIN)
    mod_InstallCommun.PoserValidationMontant ws.Range(FF_ADR_TOTAL), mod_InstallCommun.MONTANT_POSITIF_OU_NUL

Sortie:
    numErr = Err.Number
    Application.EnableEvents = evenementsAvant
    If numErr <> 0 Then Err.Raise numErr, "RemplirFormulaire", Err.Description

End Sub

' Liste déroulante du TYPE : posée directement sur la plage technique de la feuille (colonne masquée),
' sans passer par un nom défini du classeur (un nom cassé ferait disparaître la liste).
Private Sub PoserListeType(ByVal ws As Worksheet)

    Dim tps As Variant
    Dim i As Long

    tps = mod_Consommations.TypesConso()
    ws.Columns(FF_COL_AIDE_TYPE).Hidden = True
    ws.Range(ws.Cells(1, FF_COL_AIDE_TYPE), ws.Cells(20, FF_COL_AIDE_TYPE)).ClearContents
    ws.Cells(1, FF_COL_AIDE_TYPE).Value = "Types"
    For i = 0 To UBound(tps)
        ws.Cells(2 + i, FF_COL_AIDE_TYPE).Value = tps(i)
    Next i

    On Error Resume Next
    With ws.Range(FF_ADR_TYPE).MergeArea.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="=" & ws.Range(ws.Cells(2, FF_COL_AIDE_TYPE), ws.Cells(2 + UBound(tps), FF_COL_AIDE_TYPE)).Address(True, True)
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
        .ErrorTitle = "Type"
        .ErrorMessage = mod_Display.FR("Choisissez le type de consommation dans la liste.")
    End With
    On Error GoTo 0

End Sub

' Liste déroulante de l'opérateur : les fournisseurs dont le contrat n'est pas terminé pour le type choisi
' (tous les types tant qu'aucun type n'est choisi). Aucune saisie hors liste : un nouveau fournisseur doit
' d'abord faire l'objet d'un contrat (bouton Contrats de l'écran Consommations). Même principe que le type :
' plage de la feuille, pas de nom défini.
Private Sub MettreAJourListeOperateurs(ByVal ws As Worksheet, ByVal typ As String)

    Dim noms() As String
    Dim n As Long, i As Long
    Dim cel As Range
    Dim msg As String

    Set cel = ws.Range(FF_ADR_OPERATEUR)
    ws.Columns(FF_COL_AIDE_OPER).Hidden = True
    ws.Range(ws.Cells(1, FF_COL_AIDE_OPER), ws.Cells(60, FF_COL_AIDE_OPER)).ClearContents

    msg = mod_Display.FR("Cet op{e2}rateur n'a pas de contrat en cours pour ce type de consommation. ") & _
          mod_Display.FR("Pour utiliser un nouveau fournisseur, commencez par saisir un nouveau contrat ") & _
          mod_Display.FR("(bouton Contrats de l'{e2}cran Consommations), puis revenez saisir la facture.")

    n = mod_Consommations.OperateursEnCours(typ, noms)
    If n > 0 Then
        ws.Cells(1, FF_COL_AIDE_OPER).Value = "Operateurs"
        For i = 1 To n
            ws.Cells(1 + i, FF_COL_AIDE_OPER).Value = noms(i)
        Next i
    End If

    On Error Resume Next
    With cel.MergeArea.Validation
        .Delete
        If n > 0 Then
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Formula1:="=" & ws.Range(ws.Cells(2, FF_COL_AIDE_OPER), ws.Cells(1 + n, FF_COL_AIDE_OPER)).Address(True, True)
            .InCellDropdown = True
        Else
            ' Aucun contrat en cours : toute saisie est refusée avec le même message.
            .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Formula1:="=FALSE"
        End If
        .IgnoreBlank = True
        .ShowError = True
        .ErrorTitle = mod_Display.FR("Op{e2}rateur inconnu")
        .ErrorMessage = msg
    End With
    On Error GoTo 0

End Sub

' Unités proposées dans la liste déroulante : unités usuelles + toutes celles déjà utilisées dans les lignes
' enregistrées. Une unité nouvelle peut être saisie (après confirmation) : elle sera proposée ensuite.
Private Sub MettreAJourListeUnites(ByVal ws As Worksheet)

    Dim base As Variant
    Dim t() As TypeLigne
    Dim nb As Long, i As Long, j As Long, n As Long
    Dim liste() As String
    Dim deja As Boolean
    Dim u As String

    base = Array("kWh", "m" & ChrW(179), ChrW(8364), "MWh", "L", "jours", "mois", "M", "MM")
    ReDim liste(1 To UBound(base) + 1 + 60)
    For i = 0 To UBound(base)
        n = n + 1
        liste(n) = CStr(base(i))
    Next i

    nb = ChargerLignes(t, 0, True)
    For i = 1 To nb
        u = t(i).Unite
        If u <> "" And n < UBound(liste) Then
            deja = False
            For j = 1 To n
                If StrComp(liste(j), u, vbTextCompare) = 0 Then deja = True
            Next j
            If Not deja Then
                n = n + 1
                liste(n) = u
            End If
        End If
    Next i

    ws.Columns(FF_COL_AIDE_UNITE).Hidden = True
    ws.Range(ws.Cells(1, FF_COL_AIDE_UNITE), ws.Cells(UBound(liste) + 1, FF_COL_AIDE_UNITE)).ClearContents
    ws.Cells(1, FF_COL_AIDE_UNITE).Value = "Unites"
    For i = 1 To n
        ws.Cells(1 + i, FF_COL_AIDE_UNITE).Value = liste(i)
    Next i
    g_NbUnites = n

End Sub

' Taux de TVA proposés pour un type, bloc par bloc : celui de la dernière ligne enregistrée du même type dans le
' même bloc, sinon un taux usuel (eau : 5,5 % ; électricité et gaz : abonnement 5,5 %, consommation et taxes
' 20 % ; bloc libre : 20 %). tv(b) vaut -1 tant qu'aucun type n'est choisi.
Private Sub TvaDefauts(ByVal typ As String, ByRef tv() As Double)

    Dim t() As TypeLigne
    Dim nb As Long, i As Long, b As Long
    Dim fin(1 To FF_MAX_BLOCS) As Double
    Dim eau As Boolean

    For b = 1 To FF_MAX_BLOCS
        tv(b) = -1
    Next b
    If typ = "" Then Exit Sub

    eau = (StrComp(typ, mod_Consommations.TypeAdduction(), vbTextCompare) = 0 Or _
           StrComp(typ, mod_Consommations.TypeRecyclage(), vbTextCompare) = 0)
    For b = 1 To g_NbBlocs
        Select Case b
            Case NAT_ABO
                tv(b) = 0.055
            Case NAT_CONSO, NAT_TAXE
                If eau Then tv(b) = 0.055 Else tv(b) = 0.2
            Case Else
                tv(b) = 0.2
        End Select
    Next b

    nb = ChargerLignes(t, 0, True)
    For i = 1 To nb
        If StrComp(t(i).Typ, typ, vbTextCompare) = 0 Then
            b = BlocDeLigne(t(i), False)
            If b > 0 Then
                If t(i).Fin >= fin(b) Then
                    fin(b) = t(i).Fin
                    tv(b) = t(i).Tva
                End If
            End If
        End If
    Next i

End Sub

' Remplit le taux de TVA des lignes qui n'ont encore AUCUNE autre saisie (une ligne déjà commencée garde son taux).
Private Sub RemplirTvaDefaut(ByVal ws As Worksheet, ByVal typ As String)

    Dim tv(1 To FF_MAX_BLOCS) As Double
    Dim b As Long, i As Long, r As Long

    TvaDefauts typ, tv
    For b = 1 To g_NbBlocs
        If tv(b) >= 0 Then
            For i = 0 To BlocN(b) - 1
                r = BlocL1(b) + i
                If LigneVide(ws, r) Then
                    ws.Cells(r, FF_COL_TVA).Value = tv(b)
                    If b = NAT_ABO And Not mod_Consommations.SaisieNonVide(ws.Cells(r, FF_COL_UNITE).Value) Then ws.Cells(r, FF_COL_UNITE).Value = "jours"
                End If
            Next i
        End If
    Next b

End Sub

Private Function PlusGrand(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then PlusGrand = a Else PlusGrand = b
End Function

' =====================================================================================
' BLOCS DE LIGNES : GÉOMÉTRIE ET CORRESPONDANCE AVEC LES LIGNES DE FACTURE
' =====================================================================================
' Un "bloc" du formulaire porte un numéro b : 1 = abonnement, 2 = consommation, 3 = taxes (blocs fixes, modèle
' de toute nouvelle facture), puis 4 et suivants = blocs LIBRES nommés par l'opérateur (bouton "Ajouter un bloc").
Private Function NatDuBloc(ByVal b As Long) As Long
    If b <= 3 Then NatDuBloc = b Else NatDuBloc = NAT_LIBRE
End Function

Private Function NomDuBloc(ByVal b As Long) As String
    If b <= 3 Then NomDuBloc = NomNature(b) Else NomDuBloc = g_NomBloc(b)
End Function

' Numéro de bloc d'une ligne de facture (0 si introuvable). Un bloc libre inconnu est créé si "creer".
Private Function BlocDeLigne(ByRef l As TypeLigne, ByVal creer As Boolean) As Long
    Dim b As Long
    If l.Nat <= 3 Then
        BlocDeLigne = l.Nat
        Exit Function
    End If
    For b = 4 To g_NbBlocs
        If StrComp(g_NomBloc(b), l.Bloc, vbTextCompare) = 0 Then
            BlocDeLigne = b
            Exit Function
        End If
    Next b
    If creer And g_NbBlocs < FF_MAX_BLOCS Then
        g_NbBlocs = g_NbBlocs + 1
        g_NomBloc(g_NbBlocs) = l.Bloc
        BlocDeLigne = g_NbBlocs
    End If
End Function

Private Function LibelleBloc(ByVal b As Long) As String
    Select Case b
        Case NAT_ABO: LibelleBloc = "Abonnement (une ligne par p" & ChrW(233) & "riode de prix)"
        Case NAT_CONSO: LibelleBloc = "Consommation (une ligne par p" & ChrW(233) & "riode de prix)"
        Case NAT_TAXE: LibelleBloc = "Taxes et contributions"
        Case Else: LibelleBloc = g_NomBloc(b)
    End Select
End Function

' Blocs fixes par défaut (nouvelle facture) : trois blocs de FF_NB_DEFAUT lignes. Appelé par l'installateur.
Public Sub DessinerBlocsParDefaut(ByVal ws As Worksheet)
    Dim b As Long
    g_NbBlocs = 3
    For b = 1 To FF_MAX_BLOCS
        g_NomBloc(b) = ""
        g_N(b) = IIf(b <= 3, FF_NB_DEFAUT, 0)
    Next b
    DessinerBlocs ws
End Sub

' (Re)dessine tous les blocs d'après g_NbBlocs, g_N et g_NomBloc : zone effacée puis reconstruite.
' Les valeurs saisies sont perdues : l'appelant les relit avant et les réécrit après (voir ModifierStructure).
Private Sub DessinerBlocs(ByVal ws As Worksheet)

    Dim b As Long, i As Long, r As Long, titre As Long, l1 As Long, n As Long
    Dim entetes As Variant
    Dim zone As Range
    Dim ht As String, ttc As String
    Dim derniere As Long

    For b = 1 To g_NbBlocs
        g_N(b) = Borne(g_N(b))
    Next b

    ws.Columns(FF_COL_LIB).ColumnWidth = FF_LARG_LIBELLE
    MettreAJourListeUnites ws

    ' Boutons de la version précédente du dessin.
    On Error Resume Next
    For b = 1 To FF_MAX_BLOCS
        ws.Shapes("btnPlus" & b).Delete
        ws.Shapes("btnMoins" & b).Delete
        ws.Shapes("btnSuppr" & b).Delete
    Next b
    On Error GoTo 0

    ' Zone des blocs : colonnes A à K seulement (les listes techniques masquées sont plus à droite).
    derniere = FF_BLOCS_DEBUT + FF_MAX_BLOCS * (FF_MAX_LIGNES + 4)
    Set zone = ws.Range(ws.Cells(FF_BLOCS_DEBUT, 1), ws.Cells(derniere, FF_COL_TTC + 1))
    zone.UnMerge
    zone.Clear
    On Error Resume Next
    zone.Validation.Delete
    On Error GoTo 0
    zone.Font.Name = "Calibri"
    zone.Font.Size = 10
    zone.VerticalAlignment = xlCenter
    ws.Range(ws.Rows(FF_BLOCS_DEBUT), ws.Rows(derniere)).RowHeight = mod_InstallCommun.FRM_H_LIGNE

    entetes = Array("Libell" & ChrW(233), "D" & ChrW(233) & "but", "Fin", "Quantit" & ChrW(233), "Unit" & ChrW(233), "Prix HT", "TVA", "Montant HT", "Montant TTC")
    ht = mod_Consommations.AdresseCol(FF_COL_HT)
    ttc = mod_Consommations.AdresseCol(FF_COL_TTC)

    For b = 1 To g_NbBlocs
        titre = BlocTitre(b)
        l1 = BlocL1(b)
        n = g_N(b)

        mod_InstallCommun.PoserTitreBloc ws.Range(ws.Cells(titre, FF_COL_LIB), ws.Cells(titre, FF_COL_TTC)), LibelleBloc(b)

        For i = 0 To 8
            ws.Cells(titre + 1, FF_COL_LIB + i).Value = entetes(i)
        Next i
        mod_InstallCommun.PoserEnteteTableau ws.Range(ws.Cells(titre + 1, FF_COL_LIB), ws.Cells(titre + 1, FF_COL_TTC))
        ws.Range(ws.Cells(titre + 1, FF_COL_QTE), ws.Cells(titre + 1, FF_COL_TTC)).HorizontalAlignment = xlRight
        ws.Cells(titre + 1, FF_COL_UNITE).HorizontalAlignment = xlLeft

        For i = 0 To n - 1
            r = l1 + i
            mod_InstallCommun.PoserChampSaisie ws.Cells(r, FF_COL_LIB), "@"
            mod_InstallCommun.PoserChampDate ws.Cells(r, FF_COL_DEB)
            mod_InstallCommun.PoserChampDate ws.Cells(r, FF_COL_FIN)
            ' Tous les blocs ont les mêmes colonnes : quantité et unité se saisissent partout.
            mod_InstallCommun.PoserChampMontant ws.Cells(r, FF_COL_QTE), "#,##0.00", mod_InstallCommun.MONTANT_POSITIF_OU_NUL
            mod_InstallCommun.PoserChampSaisie ws.Cells(r, FF_COL_UNITE), "@"
            mod_InstallCommun.PoserChampSaisie ws.Cells(r, FF_COL_PRIX), "#,##0.00000"      ' contrôle : PoserGardeFousBloc (prix négatif = remboursement)
            mod_InstallCommun.PoserChampSaisie ws.Cells(r, FF_COL_TVA), "0.0%"
            mod_InstallCommun.PoserChampLecture ws.Cells(r, FF_COL_HT), "#,##0.00"
            ws.Cells(r, FF_COL_HT).Formula = FormuleHT(r)
            mod_InstallCommun.PoserChampLecture ws.Cells(r, FF_COL_TTC), "#,##0.00"
            ws.Cells(r, FF_COL_TTC).Formula = FormuleTTC(r)
        Next i

        ' Sous-total du bloc, puis une fine ligne de séparation.
        r = l1 + n
        mod_InstallCommun.PoserEtiquette ws.Cells(r, FF_COL_LIB), "Sous-total"
        mod_InstallCommun.PoserChampLecture ws.Cells(r, FF_COL_HT), "#,##0.00"
        ws.Cells(r, FF_COL_HT).Formula = "=SUM(" & ht & l1 & ":" & ht & (l1 + n - 1) & ")"
        mod_InstallCommun.PoserChampLecture ws.Cells(r, FF_COL_TTC), "#,##0.00"
        ws.Cells(r, FF_COL_TTC).Formula = "=SUM(" & ttc & l1 & ":" & ttc & (l1 + n - 1) & ")"
        ws.Range(ws.Cells(r, FF_COL_HT), ws.Cells(r, FF_COL_TTC)).Font.Bold = True
        ws.Rows(r + 1).RowHeight = mod_InstallCommun.FRM_H_SEP

        ' Boutons du titre : "+" (ajoute une ligne), "-" (retire la dernière ligne si elle est vide) et, pour un
        ' bloc libre, "x" (supprime le bloc). Une même macro sert à tous les blocs : le numéro est lu dans le nom du bouton.
        mod_InstallCommun.AjouterBoutonCellule ws, ws.Cells(titre, FF_COL_TTC), "+", "FacAjouterLigne", "btnPlus" & b, 24, 34
        mod_InstallCommun.AjouterBoutonCellule ws, ws.Cells(titre, FF_COL_TTC), ChrW(8722), "FacRetirerLigne", "btnMoins" & b, 24, 6
        If b > 3 Then
            mod_InstallCommun.AjouterBoutonCellule ws, ws.Cells(titre, FF_COL_HT), ChrW(215), "FacSupprimerBloc", "btnSuppr" & b, 24, 34
        End If
    Next b

    PoserFormulesRecap ws
    For b = 1 To g_NbBlocs
        PoserGardeFousBloc ws, b
    Next b

End Sub

Private Function Borne(ByVal n As Long) As Long
    If n < 1 Then
        Borne = 1
    ElseIf n > FF_MAX_LIGNES Then
        Borne = FF_MAX_LIGNES
    Else
        Borne = n
    End If
End Function

' Formules du récapitulatif (en haut) : elles pointent vers les sous-totaux de TOUS les blocs, dont la position varie.
Private Sub PoserFormulesRecap(ByVal ws As Worksheet)

    Dim b As Long
    Dim fHT As String, fTTC As String
    Dim ht As String, ttc As String

    ht = mod_Consommations.AdresseCol(FF_COL_HT)
    ttc = mod_Consommations.AdresseCol(FF_COL_TTC)
    fHT = "="
    fTTC = "="
    For b = 1 To g_NbBlocs
        If b > 1 Then
            fHT = fHT & "+"
            fTTC = fTTC & "+"
        End If
        fHT = fHT & "N(" & ht & (BlocL1(b) + g_N(b)) & ")"
        fTTC = fTTC & "N(" & ttc & (BlocL1(b) + g_N(b)) & ")"
    Next b

    ws.Range(FF_ADR_RECAP_HT).Formula = fHT
    ws.Range(FF_ADR_RECAP_TTC).Formula = fTTC
    ws.Range(FF_ADR_RECAP_TVA).Formula = "=" & FF_ADR_RECAP_TTC & "-" & FF_ADR_RECAP_HT
    ws.Range(FF_ADR_RECAP_SAISI).Formula = "=IF(ISNUMBER(" & FF_ADR_TOTAL & ")," & FF_ADR_TOTAL & ","""")"
    ws.Range(FF_ADR_RECAP_ECART).Formula = "=IF(ISNUMBER(" & FF_ADR_RECAP_SAISI & "),ROUND(" & _
        FF_ADR_RECAP_SAISI & "-" & FF_ADR_RECAP_TTC & ",2),"""")"

End Sub

' =====================================================================================
' BOUTONS : + / - (lignes), AJOUTER UN BLOC, SUPPRIMER UN BLOC
' =====================================================================================
' Numéro de bloc lu dans le nom du bouton qui a déclenché la macro ("btnPlus5" -> 5).
Private Function BlocDuBouton(ByVal prefixe As String) As Long
    Dim nom As String
    On Error Resume Next
    nom = CStr(Application.Caller)
    On Error GoTo 0
    If Left$(nom, Len(prefixe)) = prefixe Then BlocDuBouton = CLng(Val(Mid$(nom, Len(prefixe) + 1)))
End Function

Public Sub FacAjouterLigne()
    ModifierStructure MODE_PLUS, BlocDuBouton("btnPlus"), ""
End Sub

Public Sub FacRetirerLigne()
    ModifierStructure MODE_MOINS, BlocDuBouton("btnMoins"), ""
End Sub

Public Sub FacSupprimerBloc()
    ModifierStructure MODE_SUPPR_BLOC, BlocDuBouton("btnSuppr"), ""
End Sub

' Bouton "Ajouter un bloc" : crée, sous les blocs existants, un bloc de même forme (mêmes colonnes) dont le nom est
' donné par l'opérateur. Ses lignes sont enregistrées dans TblBlocsLibres avec ce nom.
Public Sub FacAjouterBloc()

    Dim ws As Worksheet
    Dim nom As String
    Dim b As Long

    If Not g_FacEnCours Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_FACTURE)

    If g_NbBlocs - 3 >= FF_MAX_BLOCS_LIBRES Then
        AfficherMessage ws, "Nombre maximal de blocs ajout{e2}s atteint (" & FF_MAX_BLOCS_LIBRES & ")."
        Exit Sub
    End If

    nom = Trim$(InputBox(mod_Display.FR("Nom du nouveau bloc (exemples : Redevance, Avoir, Location du compteur) :"), mod_Display.FR("Ajouter un bloc")))
    If nom = "" Then Exit Sub
    nom = Replace(Replace(nom, vbCr, " "), vbLf, " ")
    If Len(nom) > FF_LONG_NOM_BLOC Then
        AfficherMessage ws, "Le nom du bloc est limit{e2} {a2} " & FF_LONG_NOM_BLOC & " caract{e1}res."
        Exit Sub
    End If
    For b = 1 To g_NbBlocs
        If StrComp(NomDuBloc(b), nom, vbTextCompare) = 0 Then
            AfficherMessage ws, "Un bloc porte d{e2}j{a2} ce nom."
            Exit Sub
        End If
    Next b
    If StrComp(nom, "Taxes", vbTextCompare) = 0 Then
        AfficherMessage ws, "Ce nom est d{e2}j{a2} pris par le bloc Taxes et contributions."
        Exit Sub
    End If

    ModifierStructure MODE_AJOUT_BLOC, 0, nom

End Sub

' Modifie la structure du formulaire (ajout / retrait d'une ligne, ajout / suppression d'un bloc) : les blocs
' suivants sont décalés, le formulaire est redessiné, et les valeurs déjà saisies sont relues avant et réécrites après.
Private Sub ModifierStructure(ByVal mode As Long, ByVal b As Long, ByVal nom As String)

    Dim ws As Worksheet
    Dim brut(1 To FF_MAX_BLOCS) As Variant
    Dim ancN(1 To FF_MAX_BLOCS) As Long
    Dim ancNb As Long
    Dim ob As Long, nb As Long, i As Long, c As Long, m As Long, k As Long, r As Long
    Dim evenementsAvant As Boolean
    Dim numErr As Long
    Dim vide As Boolean

    If Not g_FacEnCours Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_FACTURE)

    Select Case mode
        Case MODE_PLUS, MODE_MOINS, MODE_SUPPR_BLOC
            If b < 1 Or b > g_NbBlocs Then Exit Sub
    End Select

    Select Case mode
        Case MODE_PLUS
            If g_N(b) >= FF_MAX_LIGNES Then
                AfficherMessage ws, "Ce bloc est limit{e2} {a2} " & FF_MAX_LIGNES & " lignes."
                Exit Sub
            End If
        Case MODE_MOINS
            If g_N(b) <= 1 Then
                AfficherMessage ws, "Un bloc garde au moins une ligne."
                Exit Sub
            End If
            If Not LigneVide(ws, BlocL1(b) + g_N(b) - 1) Then
                AfficherMessage ws, "La derni{e1}re ligne de ce bloc n'est pas vide : effacez-la d'abord."
                Exit Sub
            End If
        Case MODE_AJOUT_BLOC
            If g_NbBlocs >= FF_MAX_BLOCS Then Exit Sub
        Case MODE_SUPPR_BLOC
            If b <= 3 Then Exit Sub
            vide = True
            For i = 0 To g_N(b) - 1
                If Not LigneVide(ws, BlocL1(b) + i) Then vide = False
            Next i
            If Not vide Then
                If MsgBox(mod_Display.FR("Supprimer le bloc '") & g_NomBloc(b) & mod_Display.FR("' et les lignes qu'il contient ?"), _
                          vbYesNo + vbQuestion + vbDefaultButton2, mod_Display.FR("Supprimer un bloc")) = vbNo Then Exit Sub
            End If
    End Select

    evenementsAvant = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo Sortie

    ' Relecture des valeurs saisies (colonnes Libellé à TVA) de tous les blocs.
    ancNb = g_NbBlocs
    For ob = 1 To ancNb
        ancN(ob) = g_N(ob)
        brut(ob) = ws.Range(ws.Cells(BlocL1(ob), FF_COL_LIB), ws.Cells(BlocL1(ob) + ancN(ob) - 1, FF_COL_TVA)).Value
    Next ob

    ' Nouvelle structure.
    Select Case mode
        Case MODE_PLUS
            g_N(b) = g_N(b) + 1
        Case MODE_MOINS
            g_N(b) = g_N(b) - 1
        Case MODE_AJOUT_BLOC
            g_NbBlocs = g_NbBlocs + 1
            g_N(g_NbBlocs) = FF_NB_DEFAUT
            g_NomBloc(g_NbBlocs) = nom
        Case MODE_SUPPR_BLOC
            For k = b To g_NbBlocs - 1
                g_N(k) = g_N(k + 1)
                g_NomBloc(k) = g_NomBloc(k + 1)
            Next k
            g_N(g_NbBlocs) = 0
            g_NomBloc(g_NbBlocs) = ""
            g_NbBlocs = g_NbBlocs - 1
    End Select

    DessinerBlocs ws

    ' Réécriture : chaque bloc conservé reprend ses valeurs (dans un bloc raccourci, seules les lignes restantes).
    For ob = 1 To ancNb
        If Not (mode = MODE_SUPPR_BLOC And ob = b) Then
            nb = ob
            If mode = MODE_SUPPR_BLOC And ob > b Then nb = ob - 1
            m = ancN(ob)
            If g_N(nb) < m Then m = g_N(nb)
            For i = 1 To m
                For c = 1 To (FF_COL_TVA - FF_COL_LIB + 1)
                    If Not IsEmpty(brut(ob)(i, c)) Then
                        ws.Cells(BlocL1(nb) + i - 1, FF_COL_LIB + c - 1).Value = brut(ob)(i, c)
                    End If
                Next c
            Next i
        End If
    Next ob

    RemplirTvaDefaut ws, mod_DataStructure.CellText(ws.Range(FF_ADR_TYPE).Value)
    ws.Range(FF_ADR_MESSAGE).Value = ""
    Select Case mode
        Case MODE_PLUS
            ws.Cells(BlocL1(b) + g_N(b) - 1, FF_COL_LIB).Select
        Case MODE_AJOUT_BLOC
            ws.Cells(BlocL1(g_NbBlocs), FF_COL_LIB).Select
    End Select

Sortie:
    numErr = Err.Number
    Application.EnableEvents = evenementsAvant
    Application.ScreenUpdating = True
    If numErr <> 0 Then MsgBox mod_Display.FR("Erreur lors de la modification du formulaire : ") & Err.Description, vbExclamation

End Sub

' Vrai si la ligne r d'un bloc ne contient aucune saisie.
' La TVA (préremplie) et l'unité ne comptent pas : une ligne sans autre saisie reste vide.
Private Function LigneVide(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    Dim c As Long
    For c = FF_COL_LIB To FF_COL_PRIX
        If c <> FF_COL_UNITE Then
            If mod_Consommations.SaisieNonVide(ws.Cells(r, c).Value) Then Exit Function
        End If
    Next c
    LigneVide = True
End Function

Private Sub EcrireDate(ByVal c As Range, ByVal jour As Double)
    If jour > 0 Then c.Value = CDate(jour) Else c.ClearContents
End Sub

' Prix unitaire HT : un prix NÉGATIF est admis (remboursement, avoir). Pas de contrôle bloquant sur la cellule
' (une validation « décimal » refusait des saisies correctes) : seule une bulle d'aide s'affiche, et la validation
' du formulaire relit la valeur (virgule ou point décimal, signe moins) avant tout enregistrement.
Private Sub PoserAidePrix(ByVal c As Range)
    On Error Resume Next
    With c.Validation
        .Delete
        .Add Type:=xlValidateInputOnly
        .ShowInput = True
        .InputTitle = "Prix HT"
        .InputMessage = mod_Display.FR("Prix unitaire hors taxes, avec une virgule (exemple : 0,09010). Un remboursement se saisit en n{e2}gatif (exemple : -0,33810).")
    End With
    On Error GoTo 0
End Sub

' Validation des cellules de dates, quantités, prix et TVA d'un bloc.
Public Sub PoserGardeFousBloc(ByVal ws As Worksheet, ByVal b As Long)

    Dim l1 As Long, l2 As Long, r As Long

    l1 = BlocL1(b)
    l2 = l1 + BlocN(b) - 1
    ' Les validations de mod_InstallCommun ne portent que sur la première cellule de la plage : une cellule à la fois.
    For r = l1 To l2
        mod_InstallCommun.PoserValidationDate ws.Cells(r, FF_COL_DEB)
        mod_InstallCommun.PoserValidationDate ws.Cells(r, FF_COL_FIN)
        mod_InstallCommun.PoserValidationMontant ws.Cells(r, FF_COL_QTE), mod_InstallCommun.MONTANT_POSITIF_OU_NUL
        PoserAidePrix ws.Cells(r, FF_COL_PRIX)
    Next r
    If g_NbUnites > 0 Then
        On Error Resume Next
        With ws.Range(ws.Cells(l1, FF_COL_UNITE), ws.Cells(l2, FF_COL_UNITE)).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, _
                 Formula1:="=" & ws.Range(ws.Cells(2, FF_COL_AIDE_UNITE), ws.Cells(1 + g_NbUnites, FF_COL_AIDE_UNITE)).Address(True, True)
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowError = True
            .ErrorTitle = mod_Display.FR("Unit{e2} nouvelle")
            .ErrorMessage = mod_Display.FR("Cette unit{e2} n'est pas dans la liste. L'utiliser quand m{ea}me ? Elle sera ensuite propos{e2}e pour les prochaines lignes.")
        End With
        On Error GoTo 0
    End If
    On Error Resume Next
    With ws.Range(ws.Cells(l1, FF_COL_TVA), ws.Cells(l2, FF_COL_TVA)).Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="0", Formula2:="1"
        .IgnoreBlank = True
        .ShowInput = True
        .InputTitle = "TVA"
        .InputMessage = mod_Display.FR("Taux de TVA en pourcentage (exemple : 5,5 % ou 20 %).")
        .ShowError = True
        .ErrorTitle = mod_Display.FR("Taux invalide")
        .ErrorMessage = mod_Display.FR("Saisissez un taux entre 0 % et 100 % (exemple : 20 %).")
    End With
    On Error GoTo 0

End Sub

' Appelée par le code de la feuille (Worksheet_Change) : à la saisie du type, l'opérateur
' du contrat en vigueur est proposé s'il est encore vide.
Public Sub FacTraiterChangement(ByVal ws As Worksheet, ByVal Target As Range)

    Dim typ As String, oper As String, unite As String
    Dim d As Double
    Dim numErr As Long

    If Not g_FacEnCours Then Exit Sub
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Address(False, False) <> FF_ADR_TYPE Then Exit Sub

    On Error GoTo Sortie
    Application.EnableEvents = False
    typ = mod_DataStructure.CellText(ws.Range(FF_ADR_TYPE).Value)
    If typ = "" Then GoTo Sortie
    MettreAJourListeOperateurs ws, typ
    RemplirTvaDefaut ws, typ
    If mod_DataStructure.CellText(ws.Range(FF_ADR_OPERATEUR).Value) = "" Then
        d = mod_Consommations.VersJour(ws.Range(FF_ADR_DEBUT).Value)
        If d = 0 Then d = CDbl(Date)
        If mod_Consommations.OperateurContrat(typ, d, oper, unite) Then ws.Range(FF_ADR_OPERATEUR).Value = oper
    End If

Sortie:
    numErr = Err.Number
    Application.EnableEvents = True

End Sub

' Verrou "modal" : tant que le formulaire est ouvert, impossible de le quitter en changeant d'onglet.
Public Sub FacVerrouiller(ByVal ws As Worksheet)
    If g_FacVerrouActif And Not mod_Aide.AideOuverte() Then ws.Activate
End Sub

Public Sub FacValider()

    Dim ws As Worksheet
    Dim f As TypeFacture
    Dim tmp As TypeFacture
    Dim vTot As Variant
    Dim dTmp As Date
    Dim dDeb As Double, dFin As Double
    Dim msg As String

    If Not g_FacEnCours Then Exit Sub
    On Error GoTo Erreur
    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_FACTURE)

    ' L'identifiant de la facture en cours de modification est conservé entre l'ouverture et la validation.
    f.ID = g_IdEnEdition

    f.Typ = mod_DataStructure.CellText(ws.Range(FF_ADR_TYPE).Value)
    If f.Typ = "" Or Not mod_Consommations.TypeConnu(f.Typ) Then
        AfficherMessage ws, "Choisissez le type de consommation dans la liste."
        Exit Sub
    End If
    f.Operateur = mod_DataStructure.CellText(ws.Range(FF_ADR_OPERATEUR).Value)
    If f.Operateur = "" Then
        AfficherMessage ws, "Renseignez l'op{e2}rateur (fournisseur)."
        Exit Sub
    End If
    If Not mod_Consommations.OperateurConnu(f.Typ, f.Operateur) Then
        AfficherMessage ws, "Cet op{e2}rateur n'a pas de contrat pour ce type de consommation. Saisissez d'abord un nouveau contrat (bouton Contrats de l'{e2}cran Consommations), puis revenez saisir la facture."
        Exit Sub
    End If
    f.Numero = mod_DataStructure.CellText(ws.Range(FF_ADR_NUMERO).Value)

    If Not LireDateFacultative(ws, FF_ADR_DATE, "date de la facture", f.DateFac) Then Exit Sub
    If Not LireDateFacultative(ws, FF_ADR_DEBUT, "d{e2}but de la p{e2}riode", f.Debut) Then Exit Sub
    If Not LireDateFacultative(ws, FF_ADR_FIN, "fin de la p{e2}riode", f.Fin) Then Exit Sub
    If f.Debut > 0 And f.Fin > 0 And f.Fin < f.Debut Then
        AfficherMessage ws, "La fin de la p{e2}riode est ant{e2}rieure {a2} son d{e2}but."
        Exit Sub
    End If

    vTot = ws.Range(FF_ADR_TOTAL).Value
    If mod_Consommations.SaisieNonVide(vTot) Then
        If Not mod_DataStructure.TryLireMontant(vTot, f.TotalSaisi) Then
            AfficherMessage ws, "Le total de la facture n'est pas valide. " & mod_DataStructure.MessageMontantAttendu()
            Exit Sub
        End If
        f.TotalRenseigne = True
    End If
    f.Commentaire = mod_DataStructure.CellText(ws.Range(FF_ADR_COMMENT).Value)

    If Not LireLignes(ws, f) Then Exit Sub

    If g_FacNb = 0 Then
        AfficherMessage ws, "Saisissez au moins une ligne (abonnement, consommation ou taxe)."
        Exit Sub
    End If

    msg = ControlerChevauchements(f)
    If msg <> "" Then
        If MsgBox(mod_Display.FR(msg) & vbCrLf & vbCrLf & mod_Display.FR("Enregistrer quand m{ea}me ?"), _
                  vbYesNo + vbExclamation + vbDefaultButton2, mod_Display.FR("Chevauchement de p{e2}riodes")) = vbNo Then Exit Sub
    End If

    g_FacRes = f
    g_FacValide = True
    FermerFeuille ws
    Exit Sub

Erreur:
    Application.EnableEvents = True
    MsgBox mod_Display.FR("Erreur dans FacValider : ") & Err.Number & " - " & Err.Description, vbCritical

End Sub

Private Function LireDateFacultative(ByVal ws As Worksheet, ByVal adr As String, ByVal nom As String, ByRef jour As Double) As Boolean

    Dim v As Variant
    Dim d As Date

    v = ws.Range(adr).Value
    jour = 0
    If Not mod_Consommations.SaisieNonVide(v) Then
        LireDateFacultative = True
        Exit Function
    End If
    If Not mod_DataStructure.TryLireDate(v, d) Then
        AfficherMessage ws, "La " & nom & " n'est pas valide. " & mod_DataStructure.MessageDateAttendue()
        Exit Function
    End If
    jour = CDbl(d)
    LireDateFacultative = True

End Function

' Lit les trois blocs et remplit g_FacLignes / g_FacNb (HT et TTC calculés comme sur la facture).
' Renvoie Faux (avec un message) si une ligne est incorrecte.
Private Function LireLignes(ByVal ws As Worksheet, ByRef f As TypeFacture) As Boolean

    Dim nat As Long, i As Long, r As Long
    Dim vLib As Variant, vDeb As Variant, vFin As Variant, vQte As Variant, vUni As Variant, vPrix As Variant, vTva As Variant
    Dim dDeb As Date, dFin As Date
    Dim q As Double, p As Double, tv As Double
    Dim texteLigne As String

    g_FacNb = 0
    ReDim g_FacLignes(1 To FF_MAX_BLOCS * FF_MAX_LIGNES)

    ' "nat" désigne ici le NUMÉRO DU BLOC (1 à 3 : blocs fixes ; 4 et plus : blocs libres).
    For nat = 1 To g_NbBlocs
        For i = 0 To BlocN(nat) - 1
            r = BlocL1(nat) + i
            vLib = ws.Cells(r, FF_COL_LIB).Value
            vDeb = ws.Cells(r, FF_COL_DEB).Value
            vFin = ws.Cells(r, FF_COL_FIN).Value
            vQte = ws.Cells(r, FF_COL_QTE).Value
            vUni = ws.Cells(r, FF_COL_UNITE).Value
            vPrix = ws.Cells(r, FF_COL_PRIX).Value
            vTva = ws.Cells(r, FF_COL_TVA).Value

            ' Ligne vide : ignorée (la TVA et l'unité, préremplies, ne comptent pas).
            If Not (mod_Consommations.SaisieNonVide(vLib) Or mod_Consommations.SaisieNonVide(vDeb) Or mod_Consommations.SaisieNonVide(vFin) Or _
                    mod_Consommations.SaisieNonVide(vPrix) Or mod_Consommations.SaisieNonVide(vQte)) Then GoTo Suivante

            texteLigne = NomDuBloc(nat) & ", ligne " & (i + 1) & " : "

            If Not mod_Consommations.SaisieNonVide(vDeb) Or Not mod_DataStructure.TryLireDate(vDeb, dDeb) Then
                AfficherMessage ws, texteLigne & "date de d{e2}but absente ou invalide."
                Exit Function
            End If
            If Not mod_Consommations.SaisieNonVide(vFin) Or Not mod_DataStructure.TryLireDate(vFin, dFin) Then
                AfficherMessage ws, texteLigne & "date de fin absente ou invalide."
                Exit Function
            End If
            If dFin < dDeb Then
                AfficherMessage ws, texteLigne & "la date de fin est ant{e2}rieure {a2} la date de d{e2}but."
                Exit Function
            End If

            If Not mod_Consommations.SaisieNonVide(vQte) Or Not mod_DataStructure.TryLireMontant(vQte, q, 4) Then
                AfficherMessage ws, texteLigne & "quantit{e2} absente ou invalide."
                Exit Function
            End If
            If q < 0 Then
                AfficherMessage ws, texteLigne & "la quantit{e2} ne peut pas {ea}tre n{e2}gative."
                Exit Function
            End If
            If Not mod_Consommations.SaisieNonVide(vPrix) Or Not mod_DataStructure.TryLireMontant(vPrix, p, 6) Then
                AfficherMessage ws, texteLigne & "prix unitaire HT absent ou invalide (6 d{e2}cimales au maximum)."
                Exit Function
            End If
            If Not mod_Consommations.SaisieNonVide(vTva) Then
                AfficherMessage ws, texteLigne & "renseignez le taux de TVA (0 % si non applicable)."
                Exit Function
            End If
            If Not mod_DataStructure.TryLireMontant(vTva, tv, 6) Then
                AfficherMessage ws, texteLigne & "taux de TVA invalide."
                Exit Function
            End If
            ' Tolérance : 20 saisi dans une cellule non pourcentage = 20 %.
            If tv > 1 Then tv = tv / 100
            If tv < 0 Or tv > 1 Then
                AfficherMessage ws, texteLigne & "le taux de TVA doit {ea}tre compris entre 0 % et 100 %."
                Exit Function
            End If

            g_FacNb = g_FacNb + 1
            With g_FacLignes(g_FacNb)
                .Nat = NatDuBloc(nat)
                If nat > 3 Then
                    .Bloc = g_NomBloc(nat)
                    .RangBloc = nat - 3
                End If
                .Libelle = mod_DataStructure.CellText(vLib)
                .Debut = CDbl(dDeb)
                .Fin = CDbl(dFin)
                .Qte = q
                .Unite = mod_DataStructure.CellText(vUni)
                .Prix = p
                .Tva = tv
                .HT = Arrondi(q * p)
                .TTC = Arrondi(.HT * (1 + tv))
            End With
Suivante:
        Next i
    Next nat

    LireLignes = True

End Function

' Détecte les chevauchements qui trahissent une double saisie : deux lignes de consommation
' (ou d'abonnement) de la même facture qui se recouvrent, ou une période déjà couverte par
' une autre facture du même type. Renvoie le texte d'avertissement, ou "".
Private Function ControlerChevauchements(ByRef f As TypeFacture) As String

    Dim i As Long, j As Long, n As Long
    Dim t() As TypeLigne

    ' Lignes de la même facture entre elles (consommation et abonnement).
    For i = 1 To g_FacNb
        If (g_FacLignes(i).Nat = NAT_ABO Or g_FacLignes(i).Nat = NAT_CONSO) Then
            For j = i + 1 To g_FacNb
                If g_FacLignes(j).Nat = g_FacLignes(i).Nat Then
                    If g_FacLignes(i).Debut <= g_FacLignes(j).Fin And g_FacLignes(i).Fin >= g_FacLignes(j).Debut Then
                        ControlerChevauchements = "Deux lignes de " & LCase$(NomNature(g_FacLignes(i).Nat)) & " de cette facture se chevauchent (du " & _
                            mod_Consommations.TexteDate(g_FacLignes(j).Debut) & " au " & mod_Consommations.TexteDate(g_FacLignes(j).Fin) & ")."
                        Exit Function
                    End If
                End If
            Next j
        End If
    Next i

    ' Autres factures du même type.
    n = ChargerLignes(t, f.ID)
    For i = 1 To g_FacNb
        If (g_FacLignes(i).Nat = NAT_ABO Or g_FacLignes(i).Nat = NAT_CONSO) Then
            For j = 1 To n
                If t(j).Nat = g_FacLignes(i).Nat And StrComp(t(j).Typ, f.Typ, vbTextCompare) = 0 Then
                    If g_FacLignes(i).Debut <= t(j).Fin And g_FacLignes(i).Fin >= t(j).Debut Then
                        ControlerChevauchements = "Une ligne de " & LCase$(NomNature(g_FacLignes(i).Nat)) & " chevauche la facture n" & ChrW(176) & _
                            t(j).IDFacture & " du m" & ChrW(234) & "me type (p" & ChrW(233) & "riode du " & mod_Consommations.TexteDate(t(j).Debut) & _
                            " au " & mod_Consommations.TexteDate(t(j).Fin) & ")."
                        Exit Function
                    End If
                End If
            Next j
        End If
    Next i

End Function

Public Sub FacAnnuler()

    Dim ws As Worksheet

    If Not g_FacEnCours Then Exit Sub
    On Error GoTo Erreur

    If MsgBox(mod_Display.FR("Annuler la saisie ? Rien ne sera enregistr{e2}."), vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_FACTURE)
    g_FacValide = False
    FermerFeuille ws
    Exit Sub

Erreur:
    Application.EnableEvents = True
    MsgBox mod_Display.FR("Erreur dans FacAnnuler : ") & Err.Number & " - " & Err.Description, vbCritical

End Sub

Private Sub AfficherMessage(ByVal ws As Worksheet, ByVal texte As String)
    ws.Range(FF_ADR_MESSAGE).Value = mod_Display.FR(texte)
End Sub

' g_FacEnCours doit passer à False AVANT de masquer la feuille, sinon le verrou la réactiverait.
Private Sub FermerFeuille(ByVal ws As Worksheet)
    g_FacEnCours = False
    g_FacVerrouActif = False
    mod_Aide.MasquerFormulaire ws
End Sub
