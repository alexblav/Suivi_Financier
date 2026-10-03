Option Explicit

' =====================================================================================
' MODULE : mod_DernierImport
'
' ROLE (PHASE 3, partie ergonomie -- ALLEGE EN PHASE 6) :
'   Memorise la liste des operations ajoutees par le DERNIER import OFX (leurs
'   ID_Transaction et la date/heure de l'import) dans une feuille technique.
'
'   Point d'entree unique de ce module desormais :
'     - MemoriserDernierImport : appelee automatiquement par ImporterOperationsOFX
'                                 a la fin de chaque import (rien a faire cote operateur)
'
'   PHASE 6 : la fonction d'AFFICHAGE (AfficherDernierImportSurSynthese) a ete
'   retiree de ce module. L'affichage se fait desormais via l'ecran central
'   mod_RechercheOperations.RechercherOperations("DernierImport"), qui relit
'   cette meme feuille technique. Ce module ne s'occupe donc plus que de
'   MEMORISER, pas d'afficher.
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
' construits via la fonction mod_Display.FR().
' =====================================================================================

'Private Const NOM_FEUILLE_TECH As String = "TechDernierImport"
Private Const LIGNE_DEBUT_LISTE_ID As Long = 4

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

    wsTech.Range("A1").value = "DateImport"
    wsTech.Range("B1").value = Now
    wsTech.Range("B1").NumberFormat = "dd/mm/yyyy hh:mm:ss"

    wsTech.Range("A2").value = "NbOperations"
    wsTech.Range("B2").value = nb

    wsTech.Range("A" & (LIGNE_DEBUT_LISTE_ID - 1)).value = "ID_Transaction"

    ' Forcer la colonne en TEXTE *avant* d'ecrire les ID_Transaction, sinon Excel
    ' convertit silencieusement les longs identifiants numeriques en nombre (meme
    ' piege que celui deja traite dans mod_ImportOFX.bas pour TblOperations).
    wsTech.Range(wsTech.Cells(LIGNE_DEBUT_LISTE_ID, 1), wsTech.Cells(LIGNE_DEBUT_LISTE_ID - 1 + nb, 1)).NumberFormat = "@"

    For i = 1 To nb
        wsTech.Cells(LIGNE_DEBUT_LISTE_ID - 1 + i, 1).value = listeID(i)
    Next i

End Sub


' PHASE 6 : AfficherDernierImportSurSynthese est supprimee. L'affichage du
' dernier import se fait desormais dans l'ecran central
' frm_RechercheOperations (mod_RechercheOperations.RechercherOperations
' "DernierImport", appelee par mod_ImportOFX a la fin de chaque import, ou
' a la demande depuis n'importe quel bouton). Elle relit la meme feuille
' technique "TechDernierImport" que celle-ci alimente toujours
' (MemoriserDernierImport, juste au-dessus, est INCHANGEE).
'' =====================================================================================
'' FONCTIONS UTILITAIRES (creation / recuperation de la feuille technique)
'' =====================================================================================
Private Function ObtenirFeuilleTechSiExiste() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_TECH)
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
