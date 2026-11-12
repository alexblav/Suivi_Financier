Attribute VB_Name = "mod_Unites"
Option Explicit

' =====================================================================================
' MODULE : mod_Unites
'
' Gestion des UNITES DE MESURE (kWh, m3, stere, litres, jours...).
'
' Les unites sont memorisees dans le tableau TblUnites (feuille masquee Conso_Unites), colonnes :
'   Unite : le texte de l'unite ;
'   Conso : Oui si c'est une unite de CONSOMMATION (celle qu'une ligne d'un bloc d'indicateur Consommation peut porter),
'           Non sinon (euros, jours, mois...).
'
' - La liste de la cellule Unite d'un bloc Consommation (frm_Facture) est limitee aux unites Conso = Oui ;
'   celle des autres blocs propose toutes les unites.
' - Une unite nouvelle saisie dans un contrat ou dans un nouveau type est enregistree ici : on demande alors
'   si c'est une unite de consommation.
' - La macro GererUnites (bouton "Gerer" de frm_Contrat) permet d'ajouter une unite, de la supprimer (si
'   rien ne l'utilise) et de dire si c'est une unite de consommation.
' - Le tableau est cree tout seul au premier besoin ; il est alimente avec les unites usuelles et celles des
'   types et des contrats existants (qui sont des unites de consommation).
' =====================================================================================

Private Const FEUILLE_UNITES As String = "Conso_Unites"
Private Const TABLE_UNITES As String = "TblUnites"
Private Const COL_UNITE As String = "Unite"
Private Const COL_CONSO As String = "Conso"
Private Const VAL_OUI As String = "Oui"
Private Const VAL_NON As String = "Non"
Private Const TITRE_MSG As String = "Unit{e2}s de mesure"

Private Function Cube() As String
    Cube = "m" & ChrW(179)
End Function

' Ecriture normalisee d'une unite saisie : espaces et retours supprimes, "m3" devient "m" + exposant 3.
Public Function NormaliserUnite(ByVal u As String) As String
    Dim s As String
    s = Trim$(Replace(Replace(u, vbCr, " "), vbLf, " "))
    If StrComp(s, "m3", vbTextCompare) = 0 Then s = Cube()
    NormaliserUnite = s
End Function

' =====================================================================================
' TABLEAU DES UNITES
' =====================================================================================

' Le tableau TblUnites ; il est cree (et alimente) s'il n'existe pas encore. Nothing en cas d'echec.
Public Function TableUnites() As ListObject

    Dim lo As ListObject

    Set lo = mod_Consommations.ObtenirTable(FEUILLE_UNITES, TABLE_UNITES)
    If lo Is Nothing Then Set lo = CreerTableUnites()
    Set TableUnites = lo

End Function

Private Function CreerTableUnites() As ListObject

    Dim ws As Worksheet, wsPrec As Worksheet
    Dim lo As ListObject
    Dim base As Variant
    Dim vals() As Variant
    Dim i As Long, n As Long
    Dim evt As Boolean, ecran As Boolean

    evt = Application.EnableEvents
    ecran = Application.ScreenUpdating
    On Error GoTo Erreur

    Set wsPrec = ActiveSheet
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Set ws = mod_InstallCommun.TrouverFeuille(FEUILLE_UNITES)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = FEUILLE_UNITES
    End If
    ws.Cells.Clear
    ws.Range("A:B").NumberFormat = "@"
    ws.Cells(1, 1).Value = COL_UNITE
    ws.Cells(1, 2).Value = COL_CONSO
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B2"), , xlYes)
    lo.Name = TABLE_UNITES

    ' Unites usuelles : (unite, est une unite de consommation).
    base = Array("kWh", VAL_OUI, Cube(), VAL_OUI, "MWh", VAL_NON, "L", VAL_NON, ChrW(8364), VAL_NON, _
                 "jours", VAL_NON, "mois", VAL_NON, "M", VAL_NON, "MM", VAL_NON)
    n = (UBound(base) + 1) \ 2
    ReDim vals(1 To n, 1 To 2)
    For i = 1 To n
        vals(i, 1) = base(2 * i - 2)
        vals(i, 2) = base(2 * i - 1)
    Next i
    lo.Resize ws.Range(ws.Cells(1, 1), ws.Cells(n + 1, 2))
    lo.DataBodyRange.Value2 = vals
    ws.Columns("A:B").AutoFit
    ws.Visible = xlSheetHidden

    If Not wsPrec Is Nothing Then wsPrec.Activate
    Application.EnableEvents = evt
    Application.ScreenUpdating = ecran
    Set CreerTableUnites = lo

    ' Unites des types et des contrats existants (ex. "stere" saisi avant la creation de ce tableau).
    SynchroniserUnites
    Exit Function

Erreur:
    Application.EnableEvents = evt
    Application.ScreenUpdating = ecran
    Set CreerTableUnites = Nothing

End Function

' Rang (1 a n) de l'unite dans le tableau, 0 si elle n'y est pas.
Private Function IndexUnite(ByVal lo As ListObject, ByVal u As String) As Long

    Dim v() As Variant
    Dim i As Long
    Dim cible As String

    cible = NormaliserUnite(u)
    If cible = "" Then Exit Function
    If mod_Consommations.TableVide(lo) Then Exit Function
    v = mod_Consommations.LireColonne(lo, COL_UNITE)
    For i = 1 To lo.ListRows.Count
        If StrComp(NormaliserUnite(mod_DataStructure.CellText(v(i))), cible, vbTextCompare) = 0 Then
            IndexUnite = i
            Exit Function
        End If
    Next i

End Function

Private Sub AjouterUnite(ByVal lo As ListObject, ByVal u As String, ByVal conso As Boolean)

    Dim lr As ListRow

    Set lr = mod_Consommations.NouvelleLigne(lo)
    lr.Range.Cells(1, 1).NumberFormat = "@"
    lr.Range.Cells(1, 2).NumberFormat = "@"
    lr.Range.Cells(1, 1).Value = u
    If conso Then lr.Range.Cells(1, 2).Value = VAL_OUI Else lr.Range.Cells(1, 2).Value = VAL_NON

End Sub

' Ajoute au tableau les unites des types et des contrats qui n'y figurent pas (comme unites de consommation).
Public Sub SynchroniserUnites()

    Dim lo As ListObject

    Set lo = TableUnites()
    If lo Is Nothing Then Exit Sub
    AjouterUnitesDe lo, mod_Consommations.TableTypes()
    AjouterUnitesDe lo, mod_Consommations.ObtenirTable(mod_VarGlobales.NOM_FEUILLE_DATA_CONTRATS, mod_VarGlobales.NOM_TABLE_CONTRATS)

End Sub

Private Sub AjouterUnitesDe(ByVal lo As ListObject, ByVal src As ListObject)

    Dim v() As Variant
    Dim i As Long
    Dim s As String

    If src Is Nothing Then Exit Sub
    If mod_Consommations.TableVide(src) Then Exit Sub
    If mod_Consommations.ColIdx(src, COL_UNITE) = 0 Then Exit Sub
    v = mod_Consommations.LireColonne(src, COL_UNITE)
    For i = 1 To src.ListRows.Count
        s = NormaliserUnite(mod_DataStructure.CellText(v(i)))
        If s <> "" Then
            If IndexUnite(lo, s) = 0 Then AjouterUnite lo, s, True
        End If
    Next i

End Sub

' =====================================================================================
' INTERROGATION
' =====================================================================================

' Vrai si l'unite est une unite de consommation. "norm" reçoit son ecriture normalisee (celle du tableau).
Public Function EstUniteConso(ByVal u As String, ByRef norm As String) As Boolean

    Dim lo As ListObject
    Dim v() As Variant, w() As Variant
    Dim i As Long
    Dim s As String

    s = NormaliserUnite(u)
    norm = s
    If s = "" Then Exit Function

    Set lo = TableUnites()
    If lo Is Nothing Then
        ' Repli si le tableau est inaccessible : kWh et m3.
        If StrComp(s, "kWh", vbTextCompare) = 0 Then
            norm = "kWh"
            EstUniteConso = True
        ElseIf StrComp(s, Cube(), vbTextCompare) = 0 Then
            EstUniteConso = True
        End If
        Exit Function
    End If

    i = IndexUnite(lo, s)
    If i = 0 Then Exit Function
    v = mod_Consommations.LireColonne(lo, COL_UNITE)
    w = mod_Consommations.LireColonne(lo, COL_CONSO)
    norm = NormaliserUnite(mod_DataStructure.CellText(v(i)))
    EstUniteConso = (StrComp(mod_DataStructure.CellText(w(i)), VAL_OUI, vbTextCompare) = 0)

End Function

' Liste (tableau a partir de 0) des unites : uniquement celles de consommation si consoSeule, sinon toutes.
Public Function ListeUnites(ByVal consoSeule As Boolean) As Variant

    Dim lo As ListObject
    Dim v() As Variant, w() As Variant
    Dim res() As String
    Dim i As Long, n As Long
    Dim s As String

    ReDim res(0 To 0)
    Set lo = TableUnites()
    If Not lo Is Nothing Then
        If Not mod_Consommations.TableVide(lo) Then
            v = mod_Consommations.LireColonne(lo, COL_UNITE)
            w = mod_Consommations.LireColonne(lo, COL_CONSO)
            ReDim res(0 To lo.ListRows.Count - 1)
            For i = 1 To lo.ListRows.Count
                s = NormaliserUnite(mod_DataStructure.CellText(v(i)))
                If s <> "" Then
                    If Not consoSeule Or StrComp(mod_DataStructure.CellText(w(i)), VAL_OUI, vbTextCompare) = 0 Then
                        res(n) = s
                        n = n + 1
                    End If
                End If
            Next i
        End If
    End If

    If n = 0 Then
        ReDim res(0 To 1)
        res(0) = "kWh"
        res(1) = Cube()
        ListeUnites = res
        Exit Function
    End If
    ReDim Preserve res(0 To n - 1)
    ListeUnites = res

End Function

' Utilisations d'une unite : contrats, types et lignes de facture (dont celles d'un bloc Consommation).
Private Sub Usages(ByVal u As String, ByRef nbContrats As Long, ByRef nbTypes As Long, _
                   ByRef nbLignes As Long, ByRef nbLignesConso As Long)

    Dim src As ListObject
    Dim v() As Variant, w() As Variant
    Dim i As Long

    nbContrats = 0: nbTypes = 0: nbLignes = 0: nbLignesConso = 0

    Set src = mod_Consommations.ObtenirTable(mod_VarGlobales.NOM_FEUILLE_DATA_CONTRATS, mod_VarGlobales.NOM_TABLE_CONTRATS)
    If Not src Is Nothing Then
        If Not mod_Consommations.TableVide(src) And mod_Consommations.ColIdx(src, COL_UNITE) > 0 Then
            v = mod_Consommations.LireColonne(src, COL_UNITE)
            For i = 1 To src.ListRows.Count
                If StrComp(NormaliserUnite(mod_DataStructure.CellText(v(i))), u, vbTextCompare) = 0 Then nbContrats = nbContrats + 1
            Next i
        End If
    End If

    Set src = mod_Consommations.TableTypes()
    If Not src Is Nothing Then
        If Not mod_Consommations.TableVide(src) And mod_Consommations.ColIdx(src, COL_UNITE) > 0 Then
            v = mod_Consommations.LireColonne(src, COL_UNITE)
            For i = 1 To src.ListRows.Count
                If StrComp(NormaliserUnite(mod_DataStructure.CellText(v(i))), u, vbTextCompare) = 0 Then nbTypes = nbTypes + 1
            Next i
        End If
    End If

    Set src = mod_Factures.TableLignes()
    If Not src Is Nothing Then
        If Not mod_Consommations.TableVide(src) And mod_Consommations.ColIdx(src, COL_UNITE) > 0 Then
            v = mod_Consommations.LireColonne(src, COL_UNITE)
            w = mod_Consommations.LireColonne(src, "Indicateur")
            For i = 1 To src.ListRows.Count
                If StrComp(NormaliserUnite(mod_DataStructure.CellText(v(i))), u, vbTextCompare) = 0 Then
                    nbLignes = nbLignes + 1
                    If LCase$(Trim$(mod_DataStructure.CellText(w(i)))) = "consommation" Then nbLignesConso = nbLignesConso + 1
                End If
            Next i
        End If
    End If

End Sub

' =====================================================================================
' ENREGISTREMENT D'UNE UNITE SAISIE (contrat, nouveau type)
' =====================================================================================

' Enregistre l'unite si elle est nouvelle (apres avoir demande s'il s'agit d'une unite de consommation quand "demander"
' est vrai ; sinon elle est creee comme unite de consommation). Renvoie son ecriture normalisee.
Public Function DeclarerUnite(ByVal u As String, Optional ByVal demander As Boolean = True) As String

    Dim lo As ListObject
    Dim v() As Variant
    Dim i As Long
    Dim s As String
    Dim conso As Boolean

    s = NormaliserUnite(u)
    DeclarerUnite = s
    If s = "" Then Exit Function

    Set lo = TableUnites()
    If lo Is Nothing Then Exit Function
    i = IndexUnite(lo, s)
    If i > 0 Then
        v = mod_Consommations.LireColonne(lo, COL_UNITE)
        DeclarerUnite = NormaliserUnite(mod_DataStructure.CellText(v(i)))
        Exit Function
    End If

    conso = True
    If demander Then
        conso = (MsgBox(mod_Display.FR("L'unit{e2} '") & s & mod_Display.FR("' est nouvelle : elle va {ea}tre ajout{e2}e {a2} la liste des unit{e2}s.") & vbCrLf & vbCrLf & _
                        mod_Display.FR("Est-ce une unit{e2} de CONSOMMATION (une quantit{e2} consomm{e2}e, comme kWh ou m") & ChrW(179) & mod_Display.FR(") ?") & vbCrLf & _
                        mod_Display.FR("Oui : elle pourra servir dans un bloc d'indicateur Consommation.") & vbCrLf & _
                        mod_Display.FR("Non : unit{e2} d'un autre usage (jours, mois, euros...)."), _
                        vbYesNo + vbQuestion, mod_Display.FR(TITRE_MSG)) = vbYes)
    End If
    AjouterUnite lo, s, conso

End Function

' =====================================================================================
' MACRO DE GESTION (bouton "Gerer" de frm_Contrat, ou Alt+F8)
' =====================================================================================
Public Sub GererUnites()

    Dim lo As ListObject
    Dim v() As Variant, w() As Variant
    Dim msg As String, rep As String, cmd As String, arg As String
    Dim i As Long, n As Long

    On Error GoTo Erreur
    Set lo = TableUnites()
    If lo Is Nothing Then
        MsgBox mod_Display.FR("Le tableau des unit{e2}s est inaccessible."), vbExclamation, mod_Display.FR(TITRE_MSG)
        Exit Sub
    End If
    SynchroniserUnites

    Do
        msg = mod_Display.FR("Unit{e2}s de mesure (C = unit{e2} de consommation) :") & vbCrLf
        n = 0
        If Not mod_Consommations.TableVide(lo) Then
            n = lo.ListRows.Count
            v = mod_Consommations.LireColonne(lo, COL_UNITE)
            w = mod_Consommations.LireColonne(lo, COL_CONSO)
            For i = 1 To n
                msg = msg & i & ".  " & mod_DataStructure.CellText(v(i))
                If StrComp(mod_DataStructure.CellText(w(i)), VAL_OUI, vbTextCompare) = 0 Then msg = msg & "   [C]"
                msg = msg & vbCrLf
            Next i
        End If
        msg = msg & vbCrLf & mod_Display.FR("A nom : ajouter       S n : supprimer       C n : basculer ""consommation""") & vbCrLf & _
              mod_Display.FR("(n = num{e2}ro de la liste). Laissez vide pour terminer.")

        rep = Trim$(InputBox(msg, mod_Display.FR(TITRE_MSG), ""))
        If rep = "" Then Exit Do

        cmd = UCase$(Left$(rep, 1))
        arg = Trim$(Mid$(rep, 2))
        Select Case cmd
            Case "A"
                ActionAjouter lo, arg
            Case "S"
                ActionSupprimer lo, arg
            Case "C"
                ActionBasculer lo, arg
            Case Else
                MsgBox mod_Display.FR("Commande non reconnue : tapez A, S ou C."), vbInformation, mod_Display.FR(TITRE_MSG)
        End Select
    Loop
    Exit Sub

Erreur:
    MsgBox mod_Display.FR("Erreur dans la gestion des unit{e2}s : ") & Err.Number & " - " & Err.Description, vbCritical

End Sub

Private Sub ActionAjouter(ByVal lo As ListObject, ByVal arg As String)

    Dim s As String

    s = arg
    If s = "" Then s = InputBox(mod_Display.FR("Nom de la nouvelle unit{e2} (par exemple st{e1}re) :"), mod_Display.FR(TITRE_MSG))
    s = NormaliserUnite(s)
    If s = "" Then Exit Sub
    If Len(s) > 20 Then
        MsgBox mod_Display.FR("Le nom d'une unit{e2} est limit{e2} {a2} 20 caract{e1}res."), vbExclamation, mod_Display.FR(TITRE_MSG)
        Exit Sub
    End If
    If IndexUnite(lo, s) > 0 Then
        MsgBox mod_Display.FR("Cette unit{e2} existe d{e2}j{a2}."), vbInformation, mod_Display.FR(TITRE_MSG)
        Exit Sub
    End If
    DeclarerUnite s, True

End Sub

' Numero de ligne saisi ("S 3" -> 3), 0 si invalide (un message est affiche).
Private Function NumeroSaisi(ByVal lo As ListObject, ByVal arg As String) As Long

    Dim k As Long

    If Not IsNumeric(arg) Or Trim$(arg) = "" Then
        MsgBox mod_Display.FR("Indiquez le num{e2}ro de l'unit{e2} dans la liste, par exemple S 3."), vbInformation, mod_Display.FR(TITRE_MSG)
        Exit Function
    End If
    k = CLng(arg)
    If k < 1 Or k > lo.ListRows.Count Or mod_Consommations.TableVide(lo) Then
        MsgBox mod_Display.FR("Ce num{e2}ro n'est pas dans la liste."), vbInformation, mod_Display.FR(TITRE_MSG)
        Exit Function
    End If
    NumeroSaisi = k

End Function

Private Sub ActionSupprimer(ByVal lo As ListObject, ByVal arg As String)

    Dim k As Long
    Dim u As String, txt As String
    Dim nbC As Long, nbT As Long, nbL As Long, nbLC As Long
    Dim v() As Variant

    k = NumeroSaisi(lo, arg)
    If k = 0 Then Exit Sub
    v = mod_Consommations.LireColonne(lo, COL_UNITE)
    u = NormaliserUnite(mod_DataStructure.CellText(v(k)))

    Usages u, nbC, nbT, nbL, nbLC
    If nbC + nbT + nbL > 0 Then
        txt = mod_Display.FR("L'unit{e2} '") & u & mod_Display.FR("' est utilis{e2}e et ne peut pas {ea}tre supprim{e2}e :") & vbCrLf
        If nbC > 0 Then txt = txt & "- " & nbC & " contrat(s)" & vbCrLf
        If nbT > 0 Then txt = txt & "- " & nbT & mod_Display.FR(" type(s) de consommation") & vbCrLf
        If nbL > 0 Then txt = txt & "- " & nbL & " ligne(s) de facture" & vbCrLf
        MsgBox txt, vbExclamation, mod_Display.FR(TITRE_MSG)
        Exit Sub
    End If

    If MsgBox(mod_Display.FR("Supprimer l'unit{e2} '") & u & "' ?", vbYesNo + vbQuestion + vbDefaultButton2, mod_Display.FR(TITRE_MSG)) = vbNo Then Exit Sub

    If lo.ListRows.Count = 1 Then
        lo.ListRows(1).Range.ClearContents
    Else
        lo.ListRows(k).Delete
    End If

End Sub

Private Sub ActionBasculer(ByVal lo As ListObject, ByVal arg As String)

    Dim k As Long
    Dim u As String
    Dim nbC As Long, nbT As Long, nbL As Long, nbLC As Long
    Dim v() As Variant, w() As Variant
    Dim estConso As Boolean

    k = NumeroSaisi(lo, arg)
    If k = 0 Then Exit Sub
    v = mod_Consommations.LireColonne(lo, COL_UNITE)
    w = mod_Consommations.LireColonne(lo, COL_CONSO)
    u = NormaliserUnite(mod_DataStructure.CellText(v(k)))
    estConso = (StrComp(mod_DataStructure.CellText(w(k)), VAL_OUI, vbTextCompare) = 0)

    If estConso Then
        ' Retirer "consommation" : impossible tant qu'un contrat, un type ou une ligne de consommation l'utilise.
        Usages u, nbC, nbT, nbL, nbLC
        If nbC + nbT + nbLC > 0 Then
            MsgBox mod_Display.FR("L'unit{e2} '") & u & mod_Display.FR("' ne peut plus {ea}tre une unit{e2} de consommation : elle est utilis{e2}e par ") & _
                   nbC & " contrat(s), " & nbT & " type(s) et " & nbLC & mod_Display.FR(" ligne(s) de bloc Consommation. Changez d'abord ces unit{e2}s."), _
                   vbExclamation, mod_Display.FR(TITRE_MSG)
            Exit Sub
        End If
        lo.ListRows(k).Range.Cells(1, 2).Value = VAL_NON
    Else
        lo.ListRows(k).Range.Cells(1, 2).Value = VAL_OUI
    End If

End Sub
