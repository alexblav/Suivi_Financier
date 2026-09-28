Attribute VB_Name = "mod_DernierImport"
Option Explicit

' =====================================================================================
' MODULE : mod_DernierImport
'
' ROLE (PHASE 3, partie ergonomie) :
'   Memorise la liste des operations ajoutees par le DERNIER import OFX (leurs
'   ID_Transaction et la date/heure de l'import), et sait la reafficher sur la
'   feuille Synthese, a la maniere des autres requetes deja en place
'   (Synthese_Care utilise le meme mecanisme PrepareOutputArea / MonArray /
'   RecupPosArray / RecupColSortieIndex / MiseEnPage).
'
'   Deux points d'entree :
'     - MemoriserDernierImport   : appelee automatiquement par ImporterOperationsOFX
'                                   a la fin de chaque import (rien a faire cote operateur)
'     - AfficherDernierImportSurSynthese : appelee automatiquement apres un import,
'                                   MAIS AUSSI utilisable a la demande, plus tard,
'                                   depuis un bouton sur Synthese (a affecter
'                                   manuellement a cette macro : clic droit sur le
'                                   bouton > Affecter une macro > AfficherDernierImportSurSynthese)
'
' STOCKAGE : une feuille technique tres masquee ("TechDernierImport") contient :
'     A1 = "DateImport"      B1 = date/heure du dernier import
'     A2 = "NbOperations"    B2 = nombre d'ID_Transaction memorises
'     A4, A5, A6...          liste des ID_Transaction (un par ligne, a partir de A4)
'   Cette structure simple est deliberement extensible : si un jour tu veux
'   conserver plusieurs imports (pas seulement le dernier), il suffira d'ajouter
'   une colonne "NumeroImport" sans tout redessiner.
'
' A PROPOS DES ACCENTS : meme convention que les autres modules du chantier
' Suivi Sante : fichier 100% ASCII, textes accentues affiches a l'operateur
' construits via la fonction FR().
' =====================================================================================

Private Const NOM_FEUILLE_TECH As String = "TechDernierImport"
Private Const LIGNE_DEBUT_LISTE_ID As Long = 4


Private Function FR(ByVal texte As String) As String
    Dim r As String
    r = texte
    r = Replace(r, "{e2}", ChrW(233))
    r = Replace(r, "{e1}", ChrW(232))
    r = Replace(r, "{ea}", ChrW(234))
    r = Replace(r, "{a2}", ChrW(224))
    r = Replace(r, "{c2}", ChrW(231))
    r = Replace(r, "{o2}", ChrW(244))
    r = Replace(r, "{i2}", ChrW(238))
    r = Replace(r, "{E2}", ChrW(201))
    FR = r
End Function


' =====================================================================================
' MemoriserDernierImport : enregistre la liste des ID_Transaction du dernier
' import, avec la date/heure, dans la feuille technique (ecrase le contenu
' precedent : on ne garde que le tout dernier import pour l'instant).
' =====================================================================================
Public Sub MemoriserDernierImport(ByRef listeID() As String, ByVal nb As Long)

    Dim wsTech As Worksheet
    Dim i As Long

    Set wsTech = ObtenirOuCreerFeuilleTech()

    wsTech.Cells.Clear

    wsTech.Range("A1").Value = "DateImport"
    wsTech.Range("B1").Value = Now
    wsTech.Range("B1").NumberFormat = "dd/mm/yyyy hh:mm:ss"

    wsTech.Range("A2").Value = "NbOperations"
    wsTech.Range("B2").Value = nb

    wsTech.Range("A" & (LIGNE_DEBUT_LISTE_ID - 1)).Value = "ID_Transaction"

    For i = 1 To nb
        wsTech.Cells(LIGNE_DEBUT_LISTE_ID - 1 + i, 1).Value = listeID(i)
    Next i

End Sub


' =====================================================================================
' AfficherDernierImportSurSynthese : reaffiche, sur Synthese, la liste des
' operations du dernier import memorise. Peut etre appelee automatiquement en
' fin d'import, ou a la demande (bouton).
' =====================================================================================
Public Sub AfficherDernierImportSurSynthese()

    Dim wsTech As Worksheet
    Dim nb As Long
    Dim i As Long
    Dim listeID As Object   ' Scripting.Dictionary : recherche rapide par ID_Transaction

    Dim tabResultat() As Variant
    Dim idxRes As Long
    Dim nbLigne As Long

    Set wsTech = ObtenirFeuilleTechSiExiste()
    If wsTech Is Nothing Then
        MsgBox FR("Aucun import n'a encore {e2}t{e2} enregistr{e2}."), vbInformation, FR("Dernier import")
        Exit Sub
    End If

    nb = 0
    On Error Resume Next
    nb = CLng(wsTech.Range("B2").Value)
    On Error GoTo 0

    If nb <= 0 Then
        MsgBox FR("Aucun import n'a encore {e2}t{e2} enregistr{e2}."), vbInformation, FR("Dernier import")
        Exit Sub
    End If

    Set listeID = CreateObject("Scripting.Dictionary")
    For i = 1 To nb
        listeID(CStr(wsTech.Cells(LIGNE_DEBUT_LISTE_ID - 1 + i, 1).Value)) = True
    Next i

    ' --- Recuperation de TblOperations (variables PUBLIQUES de mod_Synthese,
    '     surtout ne pas les redeclarer ici avec Dim : voir la mesaventure
    '     rencontree sur mod_SuiviSanteFormulaire pour la meme raison) ---
    Set tbl = mod_DonneesTable.GetOperationsTable()
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    mod_Display.RecupIndexCol

    tblData = tbl.Range.Value   ' avec la ligne d'entetes, comme dans Synthese_Care
    tblDataLineTotal = UBound(tblData, 1)

    Set wsSynthese = mod_Criteres.GetFeuilleSynthese()

    MonArray = Array("Date", "Tiers", "Montant", "Cat" & ChrW(233) & "gorie", "Notes", "StatutSante", "SoldeSante")
    mod_Display.PrepareOutputArea wsSynthese, MonArray
    nbColonne = UBound(MonArray) - LBound(MonArray) + 1

    mod_Display.RecupPosArray
    mod_Display.RecupColSortieIndex

    ReDim tabResultat(1 To tblDataLineTotal, 1 To nbColonne)

    idxRes = 0
    For nbLigne = 2 To tblDataLineTotal
        If listeID.Exists(mod_DataStructure.CellText(tblData(nbLigne, colID))) Then
            idxRes = idxRes + 1
            If posDate <> 0 Then tabResultat(idxRes, posDate) = tblData(nbLigne, colDate)
            If posTiers <> 0 Then tabResultat(idxRes, posTiers) = tblData(nbLigne, colTiers)
            If posMontant <> 0 Then tabResultat(idxRes, posMontant) = tblData(nbLigne, colMontant)
            If posCategorie <> 0 Then tabResultat(idxRes, posCategorie) = tblData(nbLigne, colCategorie)
            If posNotes <> 0 Then tabResultat(idxRes, posNotes) = tblData(nbLigne, colNotes)
            If posStatutSante <> 0 Then tabResultat(idxRes, posStatutSante) = tblData(nbLigne, colStatutSante)
            If posSoldeSante <> 0 Then tabResultat(idxRes, posSoldeSante) = tblData(nbLigne, colSoldeSante)
        End If
    Next nbLigne

    If idxRes = 0 Then
        MsgBox FR("Aucune des op{e2}rations du dernier import n'a {e2}t{e2} retrouv{e2}e dans TblOperations."), _
               vbInformation, FR("Dernier import")
        Exit Sub
    End If

    plageSortieEcriture.Resize(idxRes, nbColonne).Value = tabResultat
    mod_Display.MiseEnPage idxRes

    wsSynthese.Activate

End Sub


' =====================================================================================
' FONCTIONS UTILITAIRES (creation / recuperation de la feuille technique)
' =====================================================================================
Private Function ObtenirFeuilleTechSiExiste() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(NOM_FEUILLE_TECH)
    On Error GoTo 0
    Set ObtenirFeuilleTechSiExiste = ws
End Function

Private Function ObtenirOuCreerFeuilleTech() As Worksheet
    Dim ws As Worksheet
    Set ws = ObtenirFeuilleTechSiExiste()
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = NOM_FEUILLE_TECH
        ws.Visible = xlSheetVeryHidden
    End If
    Set ObtenirOuCreerFeuilleTech = ws
End Function
