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
'   - Abonnement : quantité = nombre de jours de la période (début et fin inclus).
'   - Montant HT d'une ligne = ROUND(quantité x prix unitaire HT ; 2).
'   - Montant TTC d'une ligne = ROUND(HT x (1 + TVA) ; 2).
'   - Le total de la facture est CALCULÉ À TITRE INDICATIF (somme des lignes) et comparé au
'     total facturé saisi : l'écart est affiché mais n'empêche jamais l'enregistrement.
'
' SAISIE : formulaire frm_Facture (en-tête + trois blocs de lignes de capacité fixe),
'   ouvert par SaisirFacture (nouvelle facture) ou ModifierFactureSelection (liste
'   frm_Factures). Les montants du formulaire se calculent en direct (formules).
'
' À PROPOS DES ACCENTS : textes affichés via mod_Display.FR(); commentaires en UTF-8.
' =====================================================================================

' --- Natures de lignes ---
Public Const NAT_ABO As Long = 1
Public Const NAT_CONSO As Long = 2
Public Const NAT_TAXE As Long = 3

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
    Nat As Long
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
Private g_N(NAT_ABO To NAT_TAXE) As Long      ' nombre de lignes affichées dans chaque bloc


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
        Case Else
            Set TableLignes = mod_Consommations.ObtenirTable(mod_VarGlobales.NOM_FEUILLE_DATA_TAXES, mod_VarGlobales.NOM_TABLE_TAXES)
    End Select
End Function

Private Function TableAffFactures() As ListObject
    Set TableAffFactures = mod_Consommations.ObtenirTable(mod_VarGlobales.NOM_FEUILLE_FACTURES, mod_VarGlobales.NOM_TABLE_FACTURES_AFF)
End Function

Public Function FacturesInstallees() As Boolean
    FacturesInstallees = Not (TableFactures() Is Nothing) And Not (TableLignes(NAT_ABO) Is Nothing) And _
                         Not (TableLignes(NAT_CONSO) Is Nothing) And Not (TableLignes(NAT_TAXE) Is Nothing)
End Function

' Nombre de lignes affichées dans un bloc, et positions (le formulaire est redessiné quand ce nombre change).
Private Function BlocN(ByVal nat As Long) As Long
    BlocN = g_N(nat)
End Function

Private Function BlocTitre(ByVal nat As Long) As Long
    Dim t As Long, k As Long
    t = FF_BLOCS_DEBUT
    For k = NAT_ABO To nat - 1
        t = t + g_N(k) + 4
    Next k
    BlocTitre = t
End Function

Private Function BlocL1(ByVal nat As Long) As Long
    BlocL1 = BlocTitre(nat) + 2
End Function

Private Function NomNature(ByVal nat As Long) As String
    Select Case nat
        Case NAT_ABO: NomNature = "Abonnement"
        Case NAT_CONSO: NomNature = "Consommation"
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

' Nombre de jours d'une ligne d'abonnement (formule de la cellule Quantité).
Public Function FormuleJours(ByVal ligne As Long) As String
    Dim c As String, d As String
    c = mod_Consommations.AdresseCol(FF_COL_DEB) & ligne
    d = mod_Consommations.AdresseCol(FF_COL_FIN) & ligne
    FormuleJours = "=IF(AND(ISNUMBER(" & c & "),ISNUMBER(" & d & ")),IF(" & d & ">=" & c & "," & d & "-" & c & "+1,""""),"""")"
End Function


' =====================================================================================
' CHARGEMENT DES LIGNES (utilisé par le graphique et par les contrôles)
' =====================================================================================
' Remplit "t" (base 1) avec toutes les lignes valides des trois tableaux, avec le type et
' l'opérateur de leur facture. Renvoie leur nombre. Si "sauf" > 0, les lignes de cette
' facture sont ignorées (contrôle de chevauchement lors d'une modification).
Public Function ChargerLignes(ByRef t() As TypeLigne, Optional ByVal sauf As Long = 0) As Long

    Dim loF As ListObject, lo As ListObject
    Dim dTyp As Object, dOp As Object
    Dim vID() As Variant, vT() As Variant, vO() As Variant
    Dim vI() As Variant, vL() As Variant, vD() As Variant, vF() As Variant, vQ() As Variant
    Dim vU() As Variant, vP() As Variant, vV() As Variant, vH() As Variant, vC() As Variant
    Dim i As Long, n As Long, nat As Long, nb As Long, id As Long
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

    For nat = NAT_ABO To NAT_TAXE
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

    Dim inter As Range

    Set lo = TableAffFactures()
    If lo Is Nothing Then Exit Function
    If ActiveSheet.Name <> mod_VarGlobales.NOM_FEUILLE_FACTURES Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    On Error Resume Next
    Set inter = Intersect(Selection, lo.DataBodyRange)
    On Error GoTo 0
    If inter Is Nothing Then
        MsgBox mod_Display.FR("S{e2}lectionnez d'abord une ligne du tableau."), vbInformation
        Exit Function
    End If
    If inter.Rows.Count > 1 Then
        MsgBox mod_Display.FR("S{e2}lectionnez une seule facture."), vbInformation
        Exit Function
    End If

    If IsNumeric(lo.ListColumns("ID_Facture").DataBodyRange.Cells(inter.Row - lo.DataBodyRange.Row + 1, 1).Value2) Then
        IdFactureSelectionnee = CLng(lo.ListColumns("ID_Facture").DataBodyRange.Cells(inter.Row - lo.DataBodyRange.Row + 1, 1).Value2)
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
    ReDim g_FacLignes(1 To 3 * FF_MAX_LIGNES)
    For nat = NAT_ABO To NAT_TAXE
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
        End With
    Next i

    EnregistrerFacture = True

End Function

' Supprime, dans les quatre tableaux, la facture et toutes ses lignes.
Private Sub SupprimerDonneesFacture(ByVal id As Long)

    Dim lo As ListObject
    Dim i As Long, nat As Long
    Dim v As Variant

    For nat = 0 To NAT_TAXE
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

    On Error Resume Next
    Set inter = Intersect(Selection, lo.DataBodyRange)
    On Error GoTo 0
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
    Dim nbL As Long, n As Long, i As Long, p As Long
    Dim donnees() As Variant
    Dim ht As Double, ttc As Double
    Dim id As Long
    Dim vTot As Variant
    Dim nbCol As Long

    Set src = TableFactures()
    Set dst = TableAffFactures()
    If src Is Nothing Or dst Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    Do While dst.ListRows.Count > 1
        dst.ListRows(dst.ListRows.Count).Delete
    Loop
    dst.DataBodyRange.ClearContents

    If Not mod_Consommations.TableVide(src) Then
        nbL = ChargerLignes(t)
        n = src.ListRows.Count
        If n > 1 Then dst.Resize dst.Range.Resize(n + 1)
        nbCol = dst.ListColumns.Count
        ReDim donnees(1 To n, 1 To nbCol)
        For i = 1 To n
            id = 0
            If IsNumeric(src.ListColumns("ID_Facture").DataBodyRange.Cells(i, 1).Value2) Then
                id = mod_DataStructure.ToLong(src.ListColumns("ID_Facture").DataBodyRange.Cells(i, 1).Value2)
            End If
            TotauxFacture t, nbL, id, ht, ttc
            For p = 1 To nbCol
                Select Case dst.ListColumns(p).Name
                    Case "TotalHT": donnees(i, p) = ht
                    Case "TotalTVA": donnees(i, p) = ttc - ht
                    Case "TotalTTCCalcule": donnees(i, p) = ttc
                    Case "Ecart"
                        vTot = src.ListColumns("TotalFactureTTC").DataBodyRange.Cells(i, 1).Value2
                        If mod_Consommations.SaisieNonVide(vTot) Then donnees(i, p) = Arrondi(mod_DataStructure.ToDouble(vTot) - ttc)
                    Case Else
                        If mod_Consommations.ColIdx(src, dst.ListColumns(p).Name) > 0 Then
                            donnees(i, p) = src.ListColumns(dst.ListColumns(p).Name).DataBodyRange.Cells(i, 1).Value2
                        End If
                End Select
            Next p
        Next i
        dst.DataBodyRange.Value2 = donnees
    End If

    mod_Consommations.MettreEnFormeListeFactures dst
    Application.ScreenUpdating = True

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
    g_FacEnCours = True
    g_FacVerrouActif = True
    Do While g_FacEnCours
        DoEvents
    Loop

    On Error Resume Next
    mod_InstallCommun.ActiverFeuille g_FacNomFeuillePrec
    On Error GoTo 0

    OuvrirFormFacture = g_FacValide
    Exit Function

Erreur:
    Application.EnableEvents = True
    g_FacEnCours = False
    g_FacVerrouActif = False
    MsgBox mod_Display.FR("Erreur inattendue dans la saisie de la facture :") & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
    OuvrirFormFacture = False

End Function

Private Sub RemplirFormulaire(ByVal ws As Worksheet, ByRef f As TypeFacture)

    Dim evenementsAvant As Boolean
    Dim numErr As Long
    Dim nat As Long, k As Long, r As Long
    Dim compte(NAT_ABO To NAT_TAXE) As Long

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

    ' Nombre de lignes affichées par bloc : FF_NB_DEFAUT au moins, davantage si la facture en a plus.
    For k = 1 To g_FacNb
        compte(g_FacLignes(k).Nat) = compte(g_FacLignes(k).Nat) + 1
    Next k
    DessinerBlocs ws, PlusGrand(compte(NAT_ABO), FF_NB_DEFAUT), PlusGrand(compte(NAT_CONSO), FF_NB_DEFAUT), PlusGrand(compte(NAT_TAXE), FF_NB_DEFAUT)

    ' Les lignes de la facture sont réparties dans les blocs.
    compte(NAT_ABO) = 0: compte(NAT_CONSO) = 0: compte(NAT_TAXE) = 0
    For k = 1 To g_FacNb
        nat = g_FacLignes(k).Nat
        compte(nat) = compte(nat) + 1
        r = BlocL1(nat) + compte(nat) - 1
        ws.Cells(r, FF_COL_LIB).Value = g_FacLignes(k).Libelle
        If nat <> NAT_ABO Then ws.Cells(r, FF_COL_UNITE).Value = g_FacLignes(k).Unite
        ws.Cells(r, FF_COL_TVA).Value = g_FacLignes(k).Tva
        If Not g_FacLignes(k).Copie Then
            ws.Cells(r, FF_COL_DEB).Value = CDate(g_FacLignes(k).Debut)
            ws.Cells(r, FF_COL_FIN).Value = CDate(g_FacLignes(k).Fin)
            If nat <> NAT_ABO Then ws.Cells(r, FF_COL_QTE).Value = g_FacLignes(k).Qte
            ws.Cells(r, FF_COL_PRIX).Value = g_FacLignes(k).Prix
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

    base = Array("kWh", "m" & ChrW(179), ChrW(8364), "MWh", "L", "jours", "mois")
    ReDim liste(1 To UBound(base) + 1 + 60)
    For i = 0 To UBound(base)
        n = n + 1
        liste(n) = CStr(base(i))
    Next i

    nb = ChargerLignes(t)
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

' Taux de TVA proposés pour un type : celui de la dernière ligne enregistrée du même type et de la même
' nature, sinon un taux usuel (eau : 5,5 % ; électricité et gaz : abonnement 5,5 %, consommation et taxes 20 %).
' tv(nat) vaut -1 tant qu'aucun type n'est choisi.
Private Sub TvaDefauts(ByVal typ As String, ByRef tv() As Double)

    Dim t() As TypeLigne
    Dim nb As Long, i As Long, nat As Long
    Dim fin(NAT_ABO To NAT_TAXE) As Double
    Dim eau As Boolean

    For nat = NAT_ABO To NAT_TAXE
        tv(nat) = -1
    Next nat
    If typ = "" Then Exit Sub

    eau = (StrComp(typ, mod_Consommations.TypeAdduction(), vbTextCompare) = 0 Or _
           StrComp(typ, mod_Consommations.TypeRecyclage(), vbTextCompare) = 0)
    tv(NAT_ABO) = 0.055
    If eau Then
        tv(NAT_CONSO) = 0.055
        tv(NAT_TAXE) = 0.055
    Else
        tv(NAT_CONSO) = 0.2
        tv(NAT_TAXE) = 0.2
    End If

    nb = ChargerLignes(t)
    For i = 1 To nb
        If StrComp(t(i).Typ, typ, vbTextCompare) = 0 Then
            If t(i).Fin >= fin(t(i).Nat) Then
                fin(t(i).Nat) = t(i).Fin
                tv(t(i).Nat) = t(i).Tva
            End If
        End If
    Next i

End Sub

' Remplit le taux de TVA des lignes qui n'ont encore AUCUNE autre saisie (une ligne déjà commencée garde son taux).
Private Sub RemplirTvaDefaut(ByVal ws As Worksheet, ByVal typ As String)

    Dim tv(NAT_ABO To NAT_TAXE) As Double
    Dim nat As Long, i As Long, r As Long

    TvaDefauts typ, tv
    For nat = NAT_ABO To NAT_TAXE
        If tv(nat) >= 0 Then
            For i = 0 To BlocN(nat) - 1
                r = BlocL1(nat) + i
                If LigneVide(ws, nat, r) Then ws.Cells(r, FF_COL_TVA).Value = tv(nat)
            Next i
        End If
    Next nat

End Sub

Private Function PlusGrand(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then PlusGrand = a Else PlusGrand = b
End Function

' =====================================================================================
' DESSIN DES BLOCS DE LIGNES (appelé à l'installation, à chaque ouverture et par les boutons + / -)
' =====================================================================================
Private Function LibelleBloc(ByVal nat As Long) As String
    Select Case nat
        Case NAT_ABO: LibelleBloc = "Abonnement (une ligne par p" & ChrW(233) & "riode de prix)"
        Case NAT_CONSO: LibelleBloc = "Consommation (une ligne par p" & ChrW(233) & "riode de prix)"
        Case Else: LibelleBloc = "Taxes et contributions"
    End Select
End Function

Private Function SuffixeMacro(ByVal nat As Long) As String
    Select Case nat
        Case NAT_ABO: SuffixeMacro = "Abo"
        Case NAT_CONSO: SuffixeMacro = "Conso"
        Case Else: SuffixeMacro = "Taxe"
    End Select
End Function

' (Re)dessine les trois blocs : zone effacée puis reconstruite avec n1, n2, n3 lignes.
' Les valeurs saisies sont perdues : l'appelant les relit avant et les réécrit après (voir AjusterBloc).
Public Sub DessinerBlocs(ByVal ws As Worksheet, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long)

    Dim nat As Long, i As Long, r As Long, titre As Long, l1 As Long, n As Long
    Dim entetes As Variant
    Dim zone As Range
    Dim ht As String, ttc As String

    g_N(NAT_ABO) = Borne(n1)
    g_N(NAT_CONSO) = Borne(n2)
    g_N(NAT_TAXE) = Borne(n3)

    ws.Columns(FF_COL_LIB).ColumnWidth = FF_LARG_LIBELLE
    MettreAJourListeUnites ws

    ' Boutons + et - de la version précédente du dessin.
    On Error Resume Next
    For nat = NAT_ABO To NAT_TAXE
        ws.Shapes("btnPlus" & nat).Delete
        ws.Shapes("btnMoins" & nat).Delete
    Next nat
    On Error GoTo 0

    Set zone = ws.Range(ws.Rows(FF_BLOCS_DEBUT), ws.Rows(FF_BLOCS_DEBUT + 3 * (FF_MAX_LIGNES + 4)))
    zone.UnMerge
    zone.Clear
    On Error Resume Next
    zone.Validation.Delete
    On Error GoTo 0
    zone.Font.Name = "Calibri"
    zone.Font.Size = 10
    zone.VerticalAlignment = xlCenter
    zone.RowHeight = mod_InstallCommun.FRM_H_LIGNE

    entetes = Array("Libell" & ChrW(233), "D" & ChrW(233) & "but", "Fin", "Quantit" & ChrW(233), "Unit" & ChrW(233), "Prix HT", "TVA", "Montant HT", "Montant TTC")
    ht = mod_Consommations.AdresseCol(FF_COL_HT)
    ttc = mod_Consommations.AdresseCol(FF_COL_TTC)

    For nat = NAT_ABO To NAT_TAXE
        titre = BlocTitre(nat)
        l1 = BlocL1(nat)
        n = g_N(nat)

        mod_InstallCommun.PoserTitreBloc ws.Range(ws.Cells(titre, FF_COL_LIB), ws.Cells(titre, FF_COL_TTC)), LibelleBloc(nat)

        For i = 0 To 8
            ws.Cells(titre + 1, FF_COL_LIB + i).Value = entetes(i)
        Next i
        If nat = NAT_ABO Then ws.Cells(titre + 1, FF_COL_QTE).Value = "Jours"
        mod_InstallCommun.PoserEnteteTableau ws.Range(ws.Cells(titre + 1, FF_COL_LIB), ws.Cells(titre + 1, FF_COL_TTC))
        ws.Range(ws.Cells(titre + 1, FF_COL_QTE), ws.Cells(titre + 1, FF_COL_TTC)).HorizontalAlignment = xlRight
        ws.Cells(titre + 1, FF_COL_UNITE).HorizontalAlignment = xlLeft

        For i = 0 To n - 1
            r = l1 + i
            mod_InstallCommun.PoserChampSaisie ws.Cells(r, FF_COL_LIB), "@"
            mod_InstallCommun.PoserChampDate ws.Cells(r, FF_COL_DEB)
            mod_InstallCommun.PoserChampDate ws.Cells(r, FF_COL_FIN)
            If nat = NAT_ABO Then
                mod_InstallCommun.PoserChampLecture ws.Cells(r, FF_COL_QTE), "#,##0"
                ws.Cells(r, FF_COL_QTE).Formula = FormuleJours(r)
                mod_InstallCommun.PoserChampLecture ws.Cells(r, FF_COL_UNITE), "@"
                ws.Cells(r, FF_COL_UNITE).Value = "J"
            Else
                mod_InstallCommun.PoserChampMontant ws.Cells(r, FF_COL_QTE), "#,##0.00", mod_InstallCommun.MONTANT_POSITIF_OU_NUL
                mod_InstallCommun.PoserChampSaisie ws.Cells(r, FF_COL_UNITE), "@"
            End If
            mod_InstallCommun.PoserChampMontant ws.Cells(r, FF_COL_PRIX), "#,##0.00000", mod_InstallCommun.MONTANT_POSITIF_OU_NUL
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

        ' Boutons "+" (ajoute une ligne au bloc) et "-" (retire la dernière ligne, si elle est vide).
        mod_InstallCommun.AjouterBoutonCellule ws, ws.Cells(titre, FF_COL_TTC), "+", "FacAjouterLigne" & SuffixeMacro(nat), "btnPlus" & nat, 24, 34
        mod_InstallCommun.AjouterBoutonCellule ws, ws.Cells(titre, FF_COL_TTC), ChrW(8722), "FacRetirerLigne" & SuffixeMacro(nat), "btnMoins" & nat, 24, 6
    Next nat

    PoserFormulesRecap ws
    For nat = NAT_ABO To NAT_TAXE
        PoserGardeFousBloc ws, nat
    Next nat

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

' Formules du récapitulatif (en haut) : elles pointent vers les sous-totaux des blocs, dont la position varie.
Private Sub PoserFormulesRecap(ByVal ws As Worksheet)

    Dim nat As Long
    Dim fHT As String, fTTC As String
    Dim ht As String, ttc As String

    ht = mod_Consommations.AdresseCol(FF_COL_HT)
    ttc = mod_Consommations.AdresseCol(FF_COL_TTC)
    fHT = "="
    fTTC = "="
    For nat = NAT_ABO To NAT_TAXE
        If nat > NAT_ABO Then
            fHT = fHT & "+"
            fTTC = fTTC & "+"
        End If
        fHT = fHT & "N(" & ht & (BlocL1(nat) + g_N(nat)) & ")"
        fTTC = fTTC & "N(" & ttc & (BlocL1(nat) + g_N(nat)) & ")"
    Next nat

    ws.Range(FF_ADR_RECAP_HT).Formula = fHT
    ws.Range(FF_ADR_RECAP_TTC).Formula = fTTC
    ws.Range(FF_ADR_RECAP_TVA).Formula = "=" & FF_ADR_RECAP_TTC & "-" & FF_ADR_RECAP_HT
    ws.Range(FF_ADR_RECAP_SAISI).Formula = "=IF(ISNUMBER(" & FF_ADR_TOTAL & ")," & FF_ADR_TOTAL & ","""")"
    ws.Range(FF_ADR_RECAP_ECART).Formula = "=IF(ISNUMBER(" & FF_ADR_RECAP_SAISI & "),ROUND(" & _
        FF_ADR_RECAP_SAISI & "-" & FF_ADR_RECAP_TTC & ",2),"""")"

End Sub

' =====================================================================================
' BOUTONS "+" ET "-" DES BLOCS
' =====================================================================================
Public Sub FacAjouterLigneAbo()
    AjusterBloc NAT_ABO, 1
End Sub

Public Sub FacAjouterLigneConso()
    AjusterBloc NAT_CONSO, 1
End Sub

Public Sub FacAjouterLigneTaxe()
    AjusterBloc NAT_TAXE, 1
End Sub

Public Sub FacRetirerLigneAbo()
    AjusterBloc NAT_ABO, -1
End Sub

Public Sub FacRetirerLigneConso()
    AjusterBloc NAT_CONSO, -1
End Sub

Public Sub FacRetirerLigneTaxe()
    AjusterBloc NAT_TAXE, -1
End Sub

' Ajoute (delta = 1) ou retire (delta = -1) une ligne à un bloc : les blocs suivants sont décalés.
' Le formulaire est redessiné ; les valeurs déjà saisies sont relues avant et réécrites après.
Private Sub AjusterBloc(ByVal nat As Long, ByVal delta As Long)

    Dim ws As Worksheet
    Dim brut(NAT_ABO To NAT_TAXE) As Variant
    Dim n(NAT_ABO To NAT_TAXE) As Long
    Dim k As Long, i As Long, c As Long, m As Long
    Dim evenementsAvant As Boolean
    Dim numErr As Long
    Dim rDerniere As Long

    If Not g_FacEnCours Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_FACTURE)

    If delta > 0 And g_N(nat) >= FF_MAX_LIGNES Then
        AfficherMessage ws, "Ce bloc est limit{e2} {a2} " & FF_MAX_LIGNES & " lignes."
        Exit Sub
    End If
    If delta < 0 Then
        If g_N(nat) <= 1 Then
            AfficherMessage ws, "Un bloc garde au moins une ligne."
            Exit Sub
        End If
        rDerniere = BlocL1(nat) + g_N(nat) - 1
        If Not LigneVide(ws, nat, rDerniere) Then
            AfficherMessage ws, "La derni{e1}re ligne de ce bloc n'est pas vide : effacez-la d'abord."
            Exit Sub
        End If
    End If

    evenementsAvant = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo Sortie

    ' Relecture des valeurs saisies (colonnes Libellé à TVA).
    For k = NAT_ABO To NAT_TAXE
        n(k) = g_N(k)
        brut(k) = ws.Range(ws.Cells(BlocL1(k), FF_COL_LIB), ws.Cells(BlocL1(k) + n(k) - 1, FF_COL_TVA)).Value
    Next k

    DessinerBlocs ws, n(NAT_ABO) + IIf(nat = NAT_ABO, delta, 0), _
                      n(NAT_CONSO) + IIf(nat = NAT_CONSO, delta, 0), n(NAT_TAXE) + IIf(nat = NAT_TAXE, delta, 0)

    ' Réécriture : les lignes conservées reprennent leurs valeurs (les cellules calculées sont épargnées).
    For k = NAT_ABO To NAT_TAXE
        m = n(k)
        If g_N(k) < m Then m = g_N(k)
        For i = 1 To m
            For c = 1 To (FF_COL_TVA - FF_COL_LIB + 1)
                If k = NAT_ABO And (FF_COL_LIB + c - 1 = FF_COL_QTE Or FF_COL_LIB + c - 1 = FF_COL_UNITE) Then
                    ' formule des jours / unité fixe : on ne touche pas
                ElseIf Not IsEmpty(brut(k)(i, c)) Then
                    ws.Cells(BlocL1(k) + i - 1, FF_COL_LIB + c - 1).Value = brut(k)(i, c)
                End If
            Next c
        Next i
    Next k

    RemplirTvaDefaut ws, mod_DataStructure.CellText(ws.Range(FF_ADR_TYPE).Value)
    ws.Range(FF_ADR_MESSAGE).Value = ""
    If delta > 0 Then ws.Cells(BlocL1(nat) + g_N(nat) - 1, FF_COL_LIB).Select

Sortie:
    numErr = Err.Number
    Application.EnableEvents = evenementsAvant
    Application.ScreenUpdating = True
    If numErr <> 0 Then MsgBox mod_Display.FR("Erreur lors de l'ajustement du bloc : ") & Err.Description, vbExclamation

End Sub

' Vrai si la ligne r du bloc ne contient aucune saisie.
Private Function LigneVide(ByVal ws As Worksheet, ByVal nat As Long, ByVal r As Long) As Boolean
    Dim c As Long
    ' La TVA (préremplie) et l'unité ne comptent pas : une ligne sans autre saisie reste vide.
    For c = FF_COL_LIB To FF_COL_PRIX
        If Not (c = FF_COL_UNITE Or (nat = NAT_ABO And c = FF_COL_QTE)) Then
            If mod_Consommations.SaisieNonVide(ws.Cells(r, c).Value) Then Exit Function
        End If
    Next c
    LigneVide = True
End Function

Private Sub EcrireDate(ByVal c As Range, ByVal jour As Double)
    If jour > 0 Then c.Value = CDate(jour) Else c.ClearContents
End Sub

' Validation des cellules de dates, quantités, prix et TVA d'un bloc.
Public Sub PoserGardeFousBloc(ByVal ws As Worksheet, ByVal nat As Long)

    Dim l1 As Long, l2 As Long, r As Long

    l1 = BlocL1(nat)
    l2 = l1 + BlocN(nat) - 1
    ' Les validations de mod_InstallCommun ne portent que sur la première cellule de la plage : une cellule à la fois.
    For r = l1 To l2
        mod_InstallCommun.PoserValidationDate ws.Cells(r, FF_COL_DEB)
        mod_InstallCommun.PoserValidationDate ws.Cells(r, FF_COL_FIN)
        If nat <> NAT_ABO Then
            mod_InstallCommun.PoserValidationMontant ws.Cells(r, FF_COL_QTE), mod_InstallCommun.MONTANT_POSITIF_OU_NUL
        End If
        mod_InstallCommun.PoserValidationMontant ws.Cells(r, FF_COL_PRIX), mod_InstallCommun.MONTANT_POSITIF_OU_NUL
    Next r
    If nat <> NAT_ABO And g_NbUnites > 0 Then
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
    ReDim g_FacLignes(1 To 3 * FF_MAX_LIGNES)

    For nat = NAT_ABO To NAT_TAXE
        For i = 0 To BlocN(nat) - 1
            r = BlocL1(nat) + i
            vLib = ws.Cells(r, FF_COL_LIB).Value
            vDeb = ws.Cells(r, FF_COL_DEB).Value
            vFin = ws.Cells(r, FF_COL_FIN).Value
            vQte = ws.Cells(r, FF_COL_QTE).Value
            vUni = ws.Cells(r, FF_COL_UNITE).Value
            vPrix = ws.Cells(r, FF_COL_PRIX).Value
            vTva = ws.Cells(r, FF_COL_TVA).Value

            ' Ligne vide : ignorée (la quantité d'un abonnement est calculée, elle ne compte pas).
            If Not (mod_Consommations.SaisieNonVide(vLib) Or mod_Consommations.SaisieNonVide(vDeb) Or mod_Consommations.SaisieNonVide(vFin) Or _
                    mod_Consommations.SaisieNonVide(vPrix) Or _
                    (nat <> NAT_ABO And mod_Consommations.SaisieNonVide(vQte))) Then GoTo Suivante

            texteLigne = NomNature(nat) & ", ligne " & (i + 1) & " : "

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

            If nat = NAT_ABO Then
                q = CDbl(dFin) - CDbl(dDeb) + 1
            Else
                If Not mod_Consommations.SaisieNonVide(vQte) Or Not mod_DataStructure.TryLireMontant(vQte, q, 4) Then
                    AfficherMessage ws, texteLigne & "quantit{e2} absente ou invalide."
                    Exit Function
                End If
                If q < 0 Then
                    AfficherMessage ws, texteLigne & "la quantit{e2} ne peut pas {ea}tre n{e2}gative."
                    Exit Function
                End If
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
                .Nat = nat
                .Libelle = mod_DataStructure.CellText(vLib)
                .Debut = CDbl(dDeb)
                .Fin = CDbl(dFin)
                .Qte = q
                If nat = NAT_ABO Then .Unite = "J" Else .Unite = mod_DataStructure.CellText(vUni)
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
        If g_FacLignes(i).Nat <> NAT_TAXE Then
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
        If g_FacLignes(i).Nat <> NAT_TAXE Then
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
