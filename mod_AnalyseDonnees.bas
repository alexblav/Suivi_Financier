Attribute VB_Name = "mod_AnalyseDonnees"
Option Explicit

' =====================================================================================
' MODULE : mod_AnalyseDonnees
'
' SOCLE COMMUN des écrans d'analyse (frm_Analyse, frm_Evolution). Aucun écran n'en
' dépend pour son affichage : chacun appelle ChargerLignes puis lit les tableaux L...
'
' RÈGLES DE CHARGEMENT :
'   - Source : TblOperations + TblVentilations. Une opération "Ventilé" est REMPLACÉE
'     par ses parts (sinon double comptage). Une part prend le mois de budget de
'     l'opération d'origine, et le signe de l'opération d'origine.
'   - Mois de budget (donc les décalages de budget sont respectés) ; à défaut, mois de la
'     date comptable.
'   - Une ligne = une opération ou une part. Le mois est un INDICE : annee * 12 + (mois - 1).
' =====================================================================================

Private Const VEN_FEUILLE As String = "Ventilations"
Private Const VEN_TABLE As String = "TblVentilations"

' --- Données chargées en mémoire (remplies par ChargerLignes) -----------------------
Public LCat() As String
Public LSous() As String
Public LMont() As Double
Public LIdx() As Long
Public LNb As Long
Public LAnneeMin As Long
Public LAnneeMax As Long


' =====================================================================================
' LIBELLÉS
' =====================================================================================
Public Function LibAutres() As String
    LibAutres = "Autres"
End Function

Public Function LibSansCategorie() As String
    LibSansCategorie = mod_Display.FR("(sans cat{e2}gorie)")
End Function

Public Function LibSansSousCategorie() As String
    LibSansSousCategorie = mod_Display.FR("(sans sous-cat{e2}gorie)")
End Function

Public Function CategorieVentile() As String
    CategorieVentile = "Ventil" & ChrW(233)
End Function


' =====================================================================================
' INDICES DE MOIS
' =====================================================================================
Public Function IndiceMois(ByVal annee As Long, ByVal mois As Long) As Long
    IndiceMois = annee * 12 + (mois - 1)
End Function

Public Function AnneeDeIndice(ByVal idx As Long) As Long
    AnneeDeIndice = idx \ 12
End Function

Public Function MoisDeIndice(ByVal idx As Long) As Long
    MoisDeIndice = (idx Mod 12) + 1
End Function

Public Function TexteMois(ByVal idx As Long) As String
    TexteMois = Format(MoisDeIndice(idx), "00") & "/" & AnneeDeIndice(idx)
End Function

' Mois de budget d'une ligne ; à défaut (colonnes vides) on prend le mois de la date.
Public Function IndiceDeLigne(ByVal annee As Long, ByVal mois As Long, ByVal dateOp As Variant) As Long
    If annee < 1900 Or mois < 1 Or mois > 12 Then
        If IsDate(dateOp) Then
            annee = Year(CDate(dateOp))
            mois = Month(CDate(dateOp))
        Else
            Exit Function      ' 0 = ligne inutilisable
        End If
    End If
    IndiceDeLigne = IndiceMois(annee, mois)
End Function


' =====================================================================================
' CHARGEMENT
' =====================================================================================
' Charge TblOperations et TblVentilations en mémoire (voir règles en tête de module).
Public Function ChargerLignes() As Boolean

    Dim tblOp As ListObject, tblVen As ListObject
    Dim dOp As Variant, dVen As Variant
    Dim ixID As Long, ixCat As Long, ixSous As Long, ixMont As Long, ixMois As Long, ixAnnee As Long, ixDate As Long
    Dim vxID As Long, vxCat As Long, vxSous As Long, vxMont As Long
    Dim nOp As Long, nVen As Long, i As Long, p As Long
    Dim cat As String
    Dim opIdx() As Long, opMont() As Double
    Dim parents As Object
    Dim idx As Long
    Dim mt As Double

    LNb = 0
    LAnneeMin = 0
    LAnneeMax = 0

    Set tblOp = mod_DonneesTable.GetOperationsTable()
    If tblOp Is Nothing Then Exit Function
    If tblOp.DataBodyRange Is Nothing Then
        MsgBox mod_Display.FR("TblOperations ne contient aucune ligne."), vbInformation
        Exit Function
    End If

    ixID = mod_Display.GetColumnIndex(tblOp, "ID_Transaction")
    ixCat = mod_Display.GetColumnIndex(tblOp, "Categorie")
    ixSous = mod_Display.GetColumnIndex(tblOp, "SousCategorie")
    ixMont = mod_Display.GetColumnIndex(tblOp, "Montant")
    ixMois = mod_Display.GetColumnIndex(tblOp, "MoisBudget")
    ixAnnee = mod_Display.GetColumnIndex(tblOp, "AnneeBudget")
    ixDate = mod_Display.GetColumnIndex(tblOp, "Date_Comptable")
    If ixCat = 0 Or ixMont = 0 Then
        MsgBox mod_Display.FR("Colonnes Categorie ou Montant introuvables dans TblOperations."), vbExclamation
        Exit Function
    End If

    dOp = tblOp.DataBodyRange.Value
    nOp = UBound(dOp, 1)

    ' Parts de ventilation (facultatives)
    Set tblVen = Nothing
    On Error Resume Next
    Set tblVen = ThisWorkbook.Worksheets(VEN_FEUILLE).ListObjects(VEN_TABLE)
    On Error GoTo 0
    If Not tblVen Is Nothing Then
        If Not tblVen.DataBodyRange Is Nothing Then
            vxID = mod_Display.GetColumnIndex(tblVen, "ID_Transaction")
            vxCat = mod_Display.GetColumnIndex(tblVen, "Categorie")
            vxSous = mod_Display.GetColumnIndex(tblVen, "SousCategorie")
            vxMont = mod_Display.GetColumnIndex(tblVen, "Montant")
            If vxID <> 0 And vxCat <> 0 And vxMont <> 0 Then
                dVen = tblVen.DataBodyRange.Value
                nVen = UBound(dVen, 1)
            End If
        End If
    End If

    ReDim LCat(1 To nOp + nVen)
    ReDim LSous(1 To nOp + nVen)
    ReDim LMont(1 To nOp + nVen)
    ReDim LIdx(1 To nOp + nVen)
    ReDim opIdx(1 To nOp)
    ReDim opMont(1 To nOp)

    Set parents = CreateObject("Scripting.Dictionary")
    parents.CompareMode = 1

    For i = 1 To nOp
        mt = mod_DataStructure.ToDouble(dOp(i, ixMont))
        opMont(i) = mt
        idx = 0
        If ixMois <> 0 And ixAnnee <> 0 And ixDate <> 0 Then
            idx = IndiceDeLigne(mod_DataStructure.ToLong(dOp(i, ixAnnee)), mod_DataStructure.ToLong(dOp(i, ixMois)), dOp(i, ixDate))
        ElseIf ixDate <> 0 Then
            idx = IndiceDeLigne(0, 0, dOp(i, ixDate))
        End If
        opIdx(i) = idx
        If ixID <> 0 Then parents(mod_DataStructure.CellText(dOp(i, ixID))) = i

        cat = mod_DataStructure.CellText(dOp(i, ixCat))
        ' Une opération ventilée est remplacée par ses parts (ajoutées plus bas).
        If idx > 0 And StrComp(cat, CategorieVentile(), vbTextCompare) <> 0 Then
            LNb = LNb + 1
            LCat(LNb) = cat
            If ixSous <> 0 Then LSous(LNb) = mod_DataStructure.CellText(dOp(i, ixSous))
            LMont(LNb) = mt
            LIdx(LNb) = idx
        End If
    Next i

    If nVen > 0 And ixID <> 0 Then
        For i = 1 To nVen
            If parents.Exists(mod_DataStructure.CellText(dVen(i, vxID))) Then
                p = parents(mod_DataStructure.CellText(dVen(i, vxID)))
                If opIdx(p) > 0 Then
                    LNb = LNb + 1
                    LCat(LNb) = mod_DataStructure.CellText(dVen(i, vxCat))
                    If vxSous <> 0 Then LSous(LNb) = mod_DataStructure.CellText(dVen(i, vxSous))
                    ' Le signe est celui de l'opération d'origine, quel que soit le signe saisi sur la part.
                    mt = Abs(mod_DataStructure.ToDouble(dVen(i, vxMont)))
                    If opMont(p) < 0 Then mt = -mt
                    LMont(LNb) = mt
                    LIdx(LNb) = opIdx(p)
                End If
            End If
        Next i
    End If

    For i = 1 To LNb
        If LAnneeMin = 0 Or AnneeDeIndice(LIdx(i)) < LAnneeMin Then LAnneeMin = AnneeDeIndice(LIdx(i))
        If AnneeDeIndice(LIdx(i)) > LAnneeMax Then LAnneeMax = AnneeDeIndice(LIdx(i))
    Next i

    If LNb = 0 Then
        MsgBox mod_Display.FR("Aucune op{e2}ration exploitable (mois de budget manquant ?)."), vbInformation
        Exit Function
    End If

    ChargerLignes = True

End Function
