Attribute VB_Name = "mod_AffectationFacture"
Option Explicit

' =====================================================================================
' MODULE : mod_AffectationFacture
'
' ROLE (ajout 29/10/2026) :
'   Logique du formulaire frm_AffectationFacture (construit par mod_InstallAffectationFacture) :
'   relier une operation bancaire a la facture qui l'a reglee. Point d'entree : AffecterFacture,
'   appele par mod_Synthese.AffecterFactureOperation (bouton du tableau "Factures non liees").
'
'   DEROULEMENT
'     1. Le formulaire reprend l'operation en entete (date, tiers, categorie, montant).
'     2. Filtre 1 : Type de consommation (types presents dans TblFactures).
'     3. Filtre 2 : Date de la facture - les factures du type choisi, la plus recente en premier,
'        chaque ligne indiquant date, operateur, numero et total TTC.
'     4. Valider ecrit le lien (mod_Factures.LierFacture : l'ID_Transaction de l'operation est AJOUTE a
'        ceux deja lies a la facture, car une facture peut etre reglee en plusieurs echeances).
'        Annuler ferme sans rien ecrire. Nouvelle facture ferme le formulaire et ouvre la saisie
'        d'une facture deja liee a l'operation (mod_Factures.SaisirFacturePourOperation) ; si l'operateur
'        l'abandonne, le formulaire reapparait.
'
'   Meme mecanique que mod_FormulairesNotes : pas de code-behind, la cascade des filtres est geree par
'   SONDAGE dans la boucle d'attente (Do While g_SaisieEnCours : DoEvents : Loop). Le verrou
'   g_SaisieEnCours est celui de mod_ResolutionCategories.
'
' FICHIER 100 % ASCII : les textes accentues passent par mod_Display.FR().
' =====================================================================================

Private Const COL_TYPE_VIDE As String = "(sans type)"

Private m_Action As String            ' "Valider", "Nouvelle" ou "Annuler" : bouton clique
Private m_Attente As Boolean          ' Vrai tant que la boucle d'attente tourne reellement
Private m_IdChoisi As Long            ' facture retenue au clic sur Valider
Private m_FeuilleRetour As String     ' feuille sur laquelle on revient (Synthese, ou frm_RechercheOperations)

' Factures de TblFactures, chargees a chaque ouverture (index 1..m_NbFac, ordre du tableau).
Private m_NbFac As Long
Private m_IdFac() As Long
Private m_TypeFac() As String
Private m_LibFac() As String          ' libelle propose dans la liste (unique)
Private m_DetailFac() As String       ' texte de "Facture retenue"
Private m_Ordre() As Long             ' index des factures, date decroissante

' Liste affichee dans le filtre "Date de la facture" : index (dans m_*Fac) de chaque ligne proposee.
Private m_NbListe As Long
Private m_IndexListe() As Long


' =====================================================================================
' POINT D'ENTREE
' =====================================================================================
' Vrai si une facture a ete liee a l'operation (ou creee puis liee) : l'appelant doit alors actualiser.
' feuilleRetour : feuille reactivee a la fermeture du formulaire (Synthese si omise).
Public Function AffecterFacture(ByVal idOperation As String, ByVal jourOp As Double, ByVal tiers As String, _
                                ByVal categorie As String, ByVal montantOp As Double, _
                                Optional ByVal feuilleRetour As String = "") As Boolean

    Dim suite As String
    Dim idFacture As Long

    m_FeuilleRetour = feuilleRetour
    If m_FeuilleRetour = "" Then m_FeuilleRetour = mod_VarGlobales.NOM_FEUILLE_SYNTHESE

    If Not mod_Factures.FacturesInstallees() Then
        MsgBox mod_Display.FR("Le suivi des consommations n'est pas install{e2}. Ex{e2}cutez d'abord la macro InstallerConsommations."), vbExclamation
        Exit Function
    End If
    If mod_InstallCommun.TrouverFeuille(mod_InstallAffectationFacture.NOM_FEUILLE_AFFECTATION) Is Nothing Then
        MsgBox mod_Display.FR("Le formulaire d'affectation n'est pas install{e2} : lancez la macro CreerFeuilleAffectationFacture (Ctrl+G)."), vbExclamation
        Exit Function
    End If

    Do
        suite = AfficherFormulaire(jourOp, tiers, categorie, montantOp, idFacture)
        Select Case suite
            Case "Valider"
                If mod_Factures.LierFacture(idFacture, idOperation) Then
                    MsgBox mod_Display.FR("Facture affect{e2}e. L'op{e2}ration quitte la liste des factures non li{e2}es."), vbInformation, mod_Display.FR("Affecter une facture")
                    AffecterFacture = True
                Else
                    MsgBox mod_Display.FR("Facture introuvable : rien n'a {e2}t{e2} modifi{e2}."), vbExclamation
                End If
                Exit Do
            Case "Nouvelle"
                ' Le formulaire est deja ferme. Facture abandonnee : on revient ici, sur la meme operation.
                If mod_Factures.SaisirFacturePourOperation(idOperation) <> 0 Then
                    AffecterFacture = True
                    Exit Do
                End If
            Case Else
                Exit Do
        End Select
    Loop

End Function


' =====================================================================================
' AFFICHAGE ET BOUCLE D'ATTENTE
' =====================================================================================
' Affiche le formulaire, attend un bouton, le referme et revient sur la feuille de retour (Synthese par defaut). Renvoie l'action choisie
' ("" en cas d'erreur) ; idFacture : la facture retenue pour "Valider".
Private Function AfficherFormulaire(ByVal jourOp As Double, ByVal tiers As String, ByVal categorie As String, _
                                    ByVal montantOp As Double, ByRef idFacture As Long) As String

    Dim ws As Worksheet
    Dim dernierType As String, dernierLib As String
    Dim pos As Long
    Dim msgErr As String

    Set ws = mod_InstallCommun.TrouverFeuille(mod_InstallAffectationFacture.NOM_FEUILLE_AFFECTATION)
    If ws Is Nothing Then Exit Function

    ChargerFactures

    ' Feuille affichee EN PREMIER : redefinir un nom sur une feuille tres masquee provoque l'erreur 1004.
    ws.Visible = xlSheetVisible
    ws.Activate
    mod_InstallCommun.MasquerQuadrillage

    If jourOp > 0 Then ws.Range("afgDateOp").Value = CDate(jourOp) Else ws.Range("afgDateOp").Value = ""
    ws.Range("afgTiersOp").Value = tiers
    ws.Range("afgCategorie").Value = categorie
    ws.Range("afgMontantOp").Value = montantOp

    ws.Range("afgType").NumberFormat = "@"
    ws.Range("afgFacture").NumberFormat = "@"
    ws.Range("afgType").Value = ""
    ws.Range("afgFacture").Value = ""
    ws.Range("afgFactureRetenue").Value = ""

    EcrireListeTypes ws
    m_NbListe = 0
    EcrireListe ws, "afgListeFactures", mod_InstallAffectationFacture.AF_COL_LISTE_FACTURES, Nothing, 0
    PoserListe ws.Range("afgType"), "afgListeTypes"
    PoserListe ws.Range("afgFacture"), "afgListeFactures"
    ws.Range("afgType").Select

    m_Action = ""
    m_IdChoisi = 0
    g_SaisieEnCours = True
    m_Attente = True
    On Error GoTo ErreurBoucle
    Do While g_SaisieEnCours
        DoEvents

        If CStr(ws.Range("afgType").Value) <> dernierType Then
            dernierType = CStr(ws.Range("afgType").Value)
            dernierLib = ""
            ws.Range("afgFacture").Value = ""
            ws.Range("afgFactureRetenue").Value = ""
            EcrireListeFactures ws, dernierType
        End If

        If CStr(ws.Range("afgFacture").Value) <> dernierLib Then
            dernierLib = CStr(ws.Range("afgFacture").Value)
            pos = PositionDansListe(dernierLib)
            If pos > 0 Then
                ws.Range("afgFactureRetenue").Value = m_DetailFac(m_IndexListe(pos))
            Else
                ws.Range("afgFactureRetenue").Value = ""
            End If
        End If
    Loop
    On Error GoTo 0
    m_Attente = False

    mod_Aide.MasquerFormulaire ws
    mod_InstallCommun.ActiverFeuille m_FeuilleRetour

    idFacture = m_IdChoisi
    AfficherFormulaire = m_Action
    Exit Function

ErreurBoucle:
    ' Erreur PENDANT l'attente : sans ce traitement l'ecran resterait affiche avec des boutons sans effet.
    msgErr = "Erreur " & Err.Number & " : " & Err.Description
    m_Attente = False
    g_SaisieEnCours = False
    On Error Resume Next
    mod_Aide.MasquerFormulaire ws
    mod_InstallCommun.ActiverFeuille m_FeuilleRetour
    On Error GoTo 0
    MsgBox mod_Display.FR("L'affectation de la facture s'est interrompue sur une erreur :") & vbCrLf & msgErr & vbCrLf & vbCrLf & _
           mod_Display.FR("Rien n'a {e2}t{e2} modifi{e2}."), vbCritical

End Function


' =====================================================================================
' BOUTONS
' =====================================================================================
' Vrai si une boucle d'attente ecoute reellement les boutons. Sinon (macro interrompue, projet VBA
' reinitialise), l'ecran est orphelin : on le referme et on previent l'operateur.
Private Function BoucleActive() As Boolean
    If m_Attente Then
        BoucleActive = True
        Exit Function
    End If
    g_SaisieEnCours = False
    On Error Resume Next
    mod_Aide.MasquerFormulaire ActiveSheet
    mod_InstallCommun.AfficherSynthese
    On Error GoTo 0
    MsgBox mod_Display.FR("Cet {e2}cran n'est plus rattach{e2} {a2} un traitement en cours. Rien n'a {e2}t{e2} modifi{e2}. Relancez l'affectation depuis la feuille Synthese."), vbExclamation
End Function

Public Sub AffValider()

    Dim ws As Worksheet
    Dim pos As Long

    If Not BoucleActive() Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(mod_InstallAffectationFacture.NOM_FEUILLE_AFFECTATION)

    pos = PositionDansListe(CStr(ws.Range("afgFacture").Value))
    If pos = 0 Then
        MsgBox mod_Display.FR("Choisissez d'abord un type de consommation, puis une facture dans la liste ") & _
               mod_Display.FR("(ou cliquez sur Nouvelle facture, ou sur Annuler)."), vbInformation
        Exit Sub
    End If

    m_IdChoisi = m_IdFac(m_IndexListe(pos))
    m_Action = "Valider"
    g_SaisieEnCours = False

End Sub

Public Sub AffAnnuler()
    If Not BoucleActive() Then Exit Sub
    m_Action = "Annuler"
    g_SaisieEnCours = False
End Sub

Public Sub AffNouvelleFacture()
    If Not BoucleActive() Then Exit Sub
    m_Action = "Nouvelle"
    g_SaisieEnCours = False
End Sub


' =====================================================================================
' DONNEES ET LISTES
' =====================================================================================
' Charge TblFactures en memoire : un libelle de liste et un texte de detail par facture, et l'ordre
' date decroissante (les factures sans date en dernier, a egalite l'ordre du tableau).
Private Sub ChargerFactures()

    Dim lo As ListObject
    Dim v As Variant
    Dim i As Long, j As Long, n As Long, nbL As Long
    Dim cId As Long, cTyp As Long, cOper As Long, cNum As Long, idxDateFac As Long, cTot As Long, cLien As Long
    Dim jour() As Double
    Dim texte As String, oper As String, num As String, liens As String, typ As String, lib As String, suffixe As String
    Dim nbLiens As Long
    Dim vus As Object
    Dim element As Variant
    Dim courant As Long

    m_NbFac = 0
    Set lo = mod_Factures.TableFactures()
    If lo Is Nothing Then Exit Sub
    If mod_Consommations.TableVide(lo) Then Exit Sub

    v = lo.DataBodyRange.Value2
    nbL = lo.DataBodyRange.Rows.Count
    cId = mod_Consommations.ColIdx(lo, "ID_Facture")
    cTyp = mod_Consommations.ColIdx(lo, "TypeConso")
    cOper = mod_Consommations.ColIdx(lo, "Operateur")
    cNum = mod_Consommations.ColIdx(lo, "NumeroFacture")
    idxDateFac = mod_Consommations.ColIdx(lo, "DateFacture")
    cTot = mod_Consommations.ColIdx(lo, "TotalFactureTTC")
    cLien = mod_Consommations.ColIdx(lo, mod_Factures.FAC_COL_LIEN)

    ReDim m_IdFac(1 To nbL)
    ReDim m_TypeFac(1 To nbL)
    ReDim m_LibFac(1 To nbL)
    ReDim m_DetailFac(1 To nbL)
    ReDim jour(1 To nbL)
    Set vus = CreateObject("Scripting.Dictionary")

    For i = 1 To nbL
        If IsNumeric(v(i, cId)) And Not IsEmpty(v(i, cId)) Then
            n = n + 1
            m_IdFac(n) = CLng(v(i, cId))
            typ = mod_DataStructure.CellText(v(i, cTyp))
            If typ = "" Then typ = COL_TYPE_VIDE
            m_TypeFac(n) = typ
            jour(n) = mod_Consommations.VersJour(v(i, idxDateFac))
            oper = mod_DataStructure.CellText(v(i, cOper))
            num = mod_DataStructure.CellText(v(i, cNum))

            ' Libelle de la liste : date | operateur | numero | total TTC
            If jour(n) > 0 Then lib = Format(CDate(jour(n)), "dd/mm/yyyy") Else lib = mod_Display.FR("(sans date)")
            texte = oper
            If num <> "" Then
                If texte <> "" Then texte = texte & " | "
                texte = texte & "n" & ChrW(176) & " " & num
            End If
            If texte <> "" Then lib = lib & " | " & texte
            If mod_Consommations.SaisieNonVide(v(i, cTot)) Then
                lib = lib & " | " & Format(mod_DataStructure.ToDouble(v(i, cTot)), "#,##0.00") & " " & ChrW(8364)
            End If
            If vus.Exists(lib) Then lib = lib & " #" & m_IdFac(n)       ' deux factures identiques : on les distingue
            vus.Add lib, True
            m_LibFac(n) = lib

            ' Detail de "Facture retenue" : rappel complet et operations deja liees
            liens = ""
            If cLien > 0 Then liens = mod_DataStructure.CellText(v(i, cLien))
            nbLiens = 0
            For Each element In Split(liens, mod_Factures.FAC_SEP_LIENS)
                If Trim$(CStr(element)) <> "" Then nbLiens = nbLiens + 1
            Next element
            suffixe = mod_Display.FR("aucune op{e2}ration li{e2}e pour l'instant")
            If nbLiens = 1 Then suffixe = mod_Display.FR("d{e2}j{a2} li{e2}e {a2} 1 op{e2}ration")
            If nbLiens > 1 Then suffixe = mod_Display.FR("d{e2}j{a2} li{e2}e {a2} ") & nbLiens & mod_Display.FR(" op{e2}rations")
            m_DetailFac(n) = typ & " - " & lib & " - " & suffixe
        End If
    Next i
    m_NbFac = n
    If n = 0 Then Exit Sub

    ' Ordre : date decroissante (tri par insertion, stable).
    ReDim m_Ordre(1 To n)
    For i = 1 To n
        courant = i
        j = i - 1
        Do While j >= 1
            If Not (jour(courant) > jour(m_Ordre(j))) Then Exit Do
            m_Ordre(j + 1) = m_Ordre(j)
            j = j - 1
        Loop
        m_Ordre(j + 1) = courant
    Next i

End Sub

' Filtre 1 : types distincts (sans tenir compte de la casse), par ordre alphabetique.
Private Sub EcrireListeTypes(ByVal ws As Worksheet)

    Dim types() As String
    Dim nb As Long, i As Long, j As Long
    Dim existe As Boolean

    If m_NbFac > 0 Then
        ReDim types(1 To m_NbFac)
        For i = 1 To m_NbFac
            existe = False
            For j = 1 To nb
                If StrComp(types(j), m_TypeFac(i), vbTextCompare) = 0 Then existe = True: Exit For
            Next j
            If Not existe Then
                nb = nb + 1
                j = nb - 1
                Do While j >= 1
                    If StrComp(types(j), m_TypeFac(i), vbTextCompare) <= 0 Then Exit Do
                    types(j + 1) = types(j)
                    j = j - 1
                Loop
                types(j + 1) = m_TypeFac(i)
            End If
        Next i
    End If
    EcrireListe ws, "afgListeTypes", mod_InstallAffectationFacture.AF_COL_LISTE_TYPES, types, nb

End Sub

' Filtre 2 : factures du type choisi, la plus recente en premier (vide tant qu'aucun type n'est choisi).
Private Sub EcrireListeFactures(ByVal ws As Worksheet, ByVal typeFiltre As String)

    Dim libs() As String
    Dim k As Long, nb As Long

    m_NbListe = 0
    If typeFiltre <> "" And m_NbFac > 0 Then
        ReDim libs(1 To m_NbFac)
        ReDim m_IndexListe(1 To m_NbFac)
        For k = 1 To m_NbFac
            If StrComp(m_TypeFac(m_Ordre(k)), typeFiltre, vbTextCompare) = 0 Then
                nb = nb + 1
                m_IndexListe(nb) = m_Ordre(k)
                libs(nb) = m_LibFac(m_Ordre(k))
            End If
        Next k
        m_NbListe = nb
    End If
    EcrireListe ws, "afgListeFactures", mod_InstallAffectationFacture.AF_COL_LISTE_FACTURES, libs, m_NbListe

End Sub

' Rang (dans la liste affichee) du libelle choisi, 0 s'il n'y est pas.
Private Function PositionDansListe(ByVal libelle As String) As Long

    Dim k As Long

    If libelle = "" Then Exit Function
    For k = 1 To m_NbListe
        If m_LibFac(m_IndexListe(k)) = libelle Then
            PositionDansListe = k
            Exit Function
        End If
    Next k

End Function

' Ecrit nb valeurs a partir de la ligne 2 d'une colonne technique puis redefinit le nom pour qu'il couvre
' exactement ces cellules. Format Texte force : une valeur ressemblant a une date ne doit pas etre convertie.
Private Sub EcrireListe(ByVal ws As Worksheet, ByVal nomListe As String, ByVal colonne As String, _
                        ByRef valeurs As Variant, ByVal nb As Long)

    Dim i As Long

    ws.Range(colonne & "2:" & colonne & "1000").NumberFormat = "@"
    ws.Range(colonne & "2:" & colonne & "1000").ClearContents

    If nb < 1 Then
        ws.Names(nomListe).RefersTo = "=" & ws.Range(colonne & "2").Address(External:=True)
        Exit Sub
    End If

    For i = 1 To nb
        ws.Range(colonne & (1 + i)).Value = valeurs(i)
    Next i
    ws.Names(nomListe).RefersTo = "=" & ws.Range(colonne & "2:" & colonne & (1 + nb)).Address(External:=True)

End Sub

' Liste deroulante de validation alimentee par un nom defini de la feuille.
Private Sub PoserListe(ByVal rng As Range, ByVal nomPlage As String)
    Dim cible As Range
    Set cible = rng.Cells(1, 1).MergeArea
    On Error Resume Next
    cible.Validation.Delete
    cible.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & nomPlage
    cible.Validation.IgnoreBlank = True
    cible.Validation.InCellDropdown = True
    On Error GoTo 0
End Sub
