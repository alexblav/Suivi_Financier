Option Explicit

' =====================================================================================
' MODULE : mod_DernierImport
'
' RÔLE (PHASE 3, volet ergonomie, allégé en phase 6) :
'   Mémorise dans une feuille technique la liste des opérations ajoutées par le
'   DERNIER import OFX (leurs ID_Transaction et la date/heure de l'import).
'
'   Point d'entrée unique de ce module :
'     - MemoriserDernierImport : appelée automatiquement par ImporterOperationsOFX
'                                à la fin de chaque import (aucune action opérateur requise).
'
'   PHASE 6 : la fonction d'affichage (AfficherDernierImportSurSynthese) a été
'   retirée de ce module. L'affichage se fait désormais dans l'écran central
'   mod_RechercheOperations.RechercherOperations("DernierImport"), qui relit
'   cette même feuille technique. Ce module ne s'occupe donc plus que de
'   MÉMORISER, et non d'afficher.
'
' STOCKAGE : une feuille technique très masquée ("TechDernierImport") contient :
'     A1 = "DateImport"      B1 = date/heure du dernier import
'     A2 = "NbOperations"    B2 = nombre d'ID_Transaction mémorisés
'     A4, A5, A6...          liste des ID_Transaction (un par ligne, à partir de A4)
'   Cette structure simple est délibérément extensible : pour conserver plusieurs
'   imports (pas seulement le dernier), il suffira d'ajouter une colonne
'   "NumeroImport", sans tout redessiner.
'
' À PROPOS DES ACCENTS : même convention que dans les autres modules du chantier
' Suivi Santé : les textes affichés à l'opérateur sont construits via
' mod_Display.FR(); les commentaires du fichier sont encodés en UTF-8.
' =====================================================================================

'Private Const NOM_FEUILLE_TECH As String = "TechDernierImport"
Private Const LIGNE_DEBUT_LISTE_ID As Long = 4

' =====================================================================================
' MemoriserDernierImport : enregistre dans la feuille technique la liste des
' ID_Transaction du dernier import et sa date/heure (écrase le contenu précédent :
' pour l'instant, seul le dernier import est conservé).
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

    ' Forcer la colonne au format TEXTE avant d'écrire les ID_Transaction, sinon Excel
    ' convertit silencieusement les longs identifiants numériques en nombres (même
    ' piège que celui déjà traité dans mod_ImportOFX.bas pour TblOperations).
    wsTech.Range(wsTech.Cells(LIGNE_DEBUT_LISTE_ID, 1), wsTech.Cells(LIGNE_DEBUT_LISTE_ID - 1 + nb, 1)).NumberFormat = "@"

    For i = 1 To nb
        wsTech.Cells(LIGNE_DEBUT_LISTE_ID - 1 + i, 1).value = listeID(i)
    Next i

End Sub


' PHASE 6 : AfficherDernierImportSurSynthese a été supprimée. L'affichage du
' dernier import se fait désormais dans l'écran central frm_RechercheOperations
' (mod_RechercheOperations.RechercherOperations("DernierImport")), appelée par
' mod_ImportOFX à la fin de chaque import ou à la demande depuis n'importe quel bouton.
' Cette procédure relit la même feuille technique "TechDernierImport", toujours
' alimentée par MemoriserDernierImport (juste au-dessus, inchangée).
' =====================================================================================
' FONCTIONS UTILITAIRES (création et récupération de la feuille technique)
' =====================================================================================
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
