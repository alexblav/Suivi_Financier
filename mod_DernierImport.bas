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

'Private Const NOM_FEUILLE_TECH As String = "TechDernierImport"
Private Const LIGNE_DEBUT_LISTE_ID As Long = 4


'Private Function FR(ByVal texte As String) As String
'    Dim r As String
'    r = texte
'    r = Replace(r, "{e2}", ChrW(233))
'    r = Replace(r, "{e1}", ChrW(232))
'    r = Replace(r, "{ea}", ChrW(234))
'    r = Replace(r, "{a2}", ChrW(224))
'    r = Replace(r, "{c2}", ChrW(231))
'    r = Replace(r, "{o2}", ChrW(244))
'    r = Replace(r, "{i2}", ChrW(238))
'    r = Replace(r, "{E2}", ChrW(201))
'    FR = r
'End Function


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

    For i = 1 To nb
        wsTech.Cells(LIGNE_DEBUT_LISTE_ID - 1 + i, 1).value = listeID(i)
    Next i

End Sub


' =====================================================================================
' AfficherDernierImportSurSynthese : reaffiche, sur Synthese, la liste des
' operations du dernier import memorise. Peut etre appelee automatiquement en
' fin d'import, ou a la demande (bouton).
' =====================================================================================
Public Sub AfficherDernierImportSurSynthese()

    Dim wsTech As Worksheet
    Const LIGNE_DEBUT_LISTE_ID As Long = 4
    Dim reponse As VbMsgBoxResult
    Dim nb As Long
    Dim debPlageTravail As Range
    Dim ligneAffichage As Long
    Dim lignesDepenses() As Long
    Dim valeursDepenses() As Double
    Dim nombreDepenses As Long
    Dim nombreSurligne As Long

    'Dim i As Long
    Dim listeID As Object   ' Scripting.Dictionary : recherche rapide par ID_Transaction

    'Dim tabResultat() As Variant
    Dim idxRes As Long
    Dim nbLigne As Long

    ' 1. Récupération des pointeurs vers la feuille et le tableau
    Set wsSynthese = mod_Criteres.GetFeuille(NOM_FEUILLE_SYNTHESE)
    Set wsResultat = mod_Criteres.GetFeuille(NOM_FEUILLE_RESULTAT)
    Set tbl = mod_DonneesTable.GetOperationsValue(NOM_FEUILLE_DONNEES, "TblOperations")
    If tbl Is Nothing Then
        MsgBox "Le tableau ne contient aucune ligne de données.", vbExclamation
        Exit Sub
    End If
    If tbl.DataBodyRange Is Nothing Then
        MsgBox "Le tableau est vide.", vbExclamation
        Exit Sub
    End If

    ' 1. Verifier si la feuille de travail existe deja, pour eviter d'ecraser du travail sans prevenir.
    If Not wsResultat Is Nothing Then
        ' La feuille existe deja : on demande confirmation avant de tout reconstruire,
        ' car cela va effacer sa mise en forme actuelle.
        reponse = MsgBox(FR("La feuille '" & NOM_FEUILLE_RESULTAT & "' existe deja." & vbCrLf & _
                            "Voulez-vous la reconstruire enti{e1}rement (sa mise en forme actuelle sera perdue) ?"), _
                            vbYesNo + vbQuestion, "Confirmation de reconstruction")
        If reponse = vbNo Then
            MsgBox FR("Installation annul{e2}e, aucune modification effectu{e2}e."), vbInformation
            Exit Sub
        End If
        
        ' On la rend visible temporairement : impossible de la modifier/supprimer
        ' proprement tant qu'elle est en xlSheetVeryHidden.
        wsResultat.Visible = xlSheetVisible
        wsResultat.Cells.Clear
        Call mod_Display.SupprimerFormesExistantesFN(wsResultat)
        Call mod_Display.SupprimerNomsExistantsFN(wsResultat, NOM_FEUILLE_RESULTAT)
    Else
        ' La feuille n'existe pas encore : on la cree, positionnee en derniere position
        ' pour ne pas perturber l'ordre des onglets existants (Accueil, Synthese...).
        Set wsResultat = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        wsResultat.Name = NOM_FEUILLE_RESULTAT
    End If
    
    ' On charge la feuille TechDernierImport
    Set wsTech = mod_Criteres.GetFeuille(mod_VarGlobales.NOM_FEUILLE_TECH)
    If wsTech Is Nothing Then
        MsgBox FR("Aucun import n'a encore {e2}t{e2} enregistr{e2}."), vbInformation, FR("Dernier import")
        Exit Sub
    End If

    nb = 0
    On Error Resume Next
    nb = CLng(wsTech.Range("B2").value)
    On Error GoTo 0

    If nb <= 0 Then
        MsgBox FR("Aucun import n'a encore {e2}t{e2} enregistr{e2}."), vbInformation, FR("Dernier import")
        Exit Sub
    End If
    
    ' Les 3 premières lignes de wsTech contiennent:
    'DateImport
    'NbOperations
    'ID_Transaction
    ' Elle ne doivent pas être prise en compte
    Set listeID = CreateObject("Scripting.Dictionary")
    For nbLigne = 1 To nb
        listeID(CStr(wsTech.Cells(LIGNE_DEBUT_LISTE_ID - 1 + nbLigne, 1).value)) = True
    Next nbLigne
    
    ' On masque la feuille Synthese. Cette option est prise pour éviter à l'opérateur de se déplacer en dehors de la feuille crée
    wsSynthese.Visible = xlSheetVeryHidden
    
    ' 2. Charge TOUT le tableau de données (en-têtes incluses en ligne 1)
    tblData = tbl.Range.value
    tblDataLineTotal = UBound(tblData, 1)
    
    ' 3. Récupération des index de colonne dans la base de données
    mod_Display.RecupIndexCol
    
    ' 4. Autorise ou non le double click
    AllowDetailDoubleClick = False
    ' Paramètre de navigation
    RecherOperations = False
    ' Désactive les événements
    Application.EnableEvents = False
    
    ' 5. Construction de la zone des boutons
    ' On détermine la position du bouton
    Set zoneBouton = wsResultat.Range(cellSortieDep).Offset(0, 0)
    
    'On définit son titre
    texteBouton = "Sortir"
    nomMacroBouton = "Sortir"
    Call mod_Display.ConstruireBoutons(wsResultat, zoneBouton, texteBouton, nomMacroBouton)

    ' 6. Construction de la zone des instructions
    ' On fournit le message à afficher en remplaçant les carractères accentué par les balise de la fonction FR
    Message = "Derni{e1}re données importées" & Chr(10) & _
                "Liste les op{e2}rations ajout{e2}es au dernier import" & Chr(10) & _
                 "{A2} la fin sortez avec le bouton ""Sortir"""
                 
    ' On fournit le titre, la position de la cellule dans laquelle on veut écrire de titre
    Titre = "Instructions"
    Set PositionTitre = wsResultat.Range(cellSortieDep).Offset(2, 0)
    
    ' On détermine la plage de début et de fin de la zone
    Set debutZone = wsResultat.Range(cellSortieDep).Offset(3, 0)
    Set finZone = wsResultat.Range(cellSortieDep).Offset(5, 5)
    
    Call mod_Display.ConstruireZoneTexte(wsResultat, Titre, Message, PositionTitre, debutZone, finZone)

    ' 7. Affiche les entêtes du tableau de sortie
    ' Définit le début de la page de travail
    Set debPlageTravail = wsResultat.Range(cellSortieDep).Offset(7, 0)
    
    MonArray = Array("Date", "Tiers", "Montant", "Cat" & ChrW(233) & "gorie", "Notes", "StatutSante", "SoldeSante")
    Call mod_Display.PrepareOutputArea(wsResultat, MonArray, debPlageTravail)
    
    nbColonne = UBound(MonArray) - LBound(MonArray) + 1
    
    ' Récupére la position du champs dans l'ARRAY (Application.Match est naturellement insensible à la casse)
    mod_Display.RecupPosArray
    
    ' Récupération des index de colonne dans la feuille de sortie
    Call mod_Display.RecupPosSortieIndex(wsResultat, debPlageTravail)
    
    ' Taille maximale du tableau de résultat = nombre total de lignes source
    ReDim tabResultat(1 To tblDataLineTotal, 1 To nbColonne)

    idxRes = 0
    For nbLigne = 2 To tblDataLineTotal
        If listeID.Exists(mod_DataStructure.CellText(tblData(nbLigne, colID))) Then
            idxRes = idxRes + 1
            tabResultat(idxRes, posDate) = tblData(nbLigne, colDate)
            tabResultat(idxRes, posTiers) = tblData(nbLigne, colTiers)
            tabResultat(idxRes, posMontant) = tblData(nbLigne, colMontant)
            tabResultat(idxRes, posCategorie) = tblData(nbLigne, colCategorie)
            tabResultat(idxRes, posNotes) = tblData(nbLigne, colNotes)
            tabResultat(idxRes, posStatutSante) = tblData(nbLigne, colStatutSante)
            tabResultat(idxRes, posSoldeSante) = tblData(nbLigne, colSoldeSante)
        End If
    Next nbLigne

    If idxRes = 0 Then
        MsgBox FR("Aucune des op{e2}rations du dernier import n'a {e2}t{e2} retrouv{e2}e dans TblOperations."), _
               vbInformation, FR("Dernier import")
        Exit Sub
    ElseIf idxRes > 0 Then
        Application.ScreenUpdating = False
        
        ' Injection directe du tableau mémoire dans la plage d'affichage
        ' On redimentionne la taille de la plage pour pas voir s'afficher des erreur type #N/A dans les cellules en trop
        plageSortieEcriture.Resize(idxRes, nbColonne).value = tabResultat
        
        ' Mise en couleur conditionnel
        ligneAffichage = 2
        For nbLigne = 1 To idxRes
            ' Une entrée (montant positif) est affichée en vert.
            If tabResultat(nbLigne, posMontant) > 0 Then
                plageSortieEcriture.Cells(nbLigne, posMontant).Font.Color = RGB(0, 128, 0)
            End If
            ' Une dépense (montant négatif) est mémorisée comme candidate au surlignage.
'            If tabResultat(nbLigne, posMontant) < 0 Then
'                mod_Rapports.AddExpenseForHighlight lignesDepenses, valeursDepenses, nombreDepenses, ligneAffichage, Abs(tabResultat(nbLigne, posMontant))
'            End If
            ligneAffichage = ligneAffichage + 1
        Next nbLigne
    
'        ' Déterminer combien de dépenses surligner, puis les surligner.
'        nombreSurligne = mod_Rapports.GetTopCount(critNbOperations, nombreDepenses)
'        mod_Rapports.HighlightTopRows lignesDepenses, valeursDepenses, nombreDepenses, nombreSurligne
        
        ' Mise en forme rapide des colonnes
        mod_Display.MiseEnPage wsResultat, idxRes
        
         ' Réactive les événements
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If

End Sub


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
