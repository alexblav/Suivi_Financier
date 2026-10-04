Option Explicit

' =====================================================================================
' MODULE : mod_DecalagesBudget
'
' ROLE (ajout 03/10/2026, demande operateur) :
'   Avant ce module, UNE SEULE regle etait codee en dur dans
'   mod_ImportOFX.CalculerBudget : les salaires verses par "DRFIP OCCITANIE ET HTE"
'   etaient affectes au mois SUIVANT celui de l'operation (le virement de fin de
'   mois correspond au salaire du mois qui commence). Cette regle a casse
'   silencieusement lors du passage a la hierarchie Categorie/Sous-categorie
'   (Phase 5) : le test comparait la colonne Categorie a "Salaire/Revenus
'   d'activite", qui est devenue une SOUS-categorie de "Revenus" - plus aucune
'   operation ne pouvait donc jamais correspondre.
'
'   Ce module remplace cette regle unique par un TABLEAU DE CORRESPONDANCE,
'   modifiable par l'operateur sans toucher au code : la feuille Param, tableau
'   "TblDecalagesBudget" (colonnes S a V, voir mod_VarGlobales). Chaque ligne du
'   tableau precise un critere (Tiers / Categorie / Sous-categorie - chaque
'   champ laisse VIDE agit comme un joker, "n'importe quelle valeur") et le
'   decalage a appliquer, en nombre de mois (+1, -1, 0, ou toute autre valeur
'   entiere si besoin un jour).
'
'   Une operation qui ne correspond a AUCUNE ligne du tableau garde le
'   comportement actuel (aucun decalage, budget = mois de l'operation).
'
' COLONNES DU TABLEAU "TblDecalagesBudget" (feuille Param, S:V) :
'   S = Tiers          : texte exact du Tiers, ou VIDE = n'importe quel Tiers.
'   T = Categorie      : categorie PARENTE (ex: "Revenus"), ou VIDE = joker.
'   U = SousCategorie  : sous-categorie (ex: "Salaire/Revenus d'activite"), ou
'                        VIDE = joker.
'   V = Decalage       : nombre de mois a ajouter a Date_Comptable pour obtenir
'                        le mois de budget (entier, positif ou negatif).
'
' COMMENT INSTALLER / RE-INSTALLER LE TABLEAU :
'   Ctrl+G (fenetre Execution), taper :  InitialiserTableDecalagesBudget
'   Sans danger si le tableau existe deja : la procedure ne fait rien dans ce cas
'   (meme principe que mod_Categories.AjouterColonneSousCategorie).
'
' COMMENT AJOUTER UNE NOUVELLE REGLE DE DECALAGE :
'   Directement dans le tableau TblDecalagesBudget (feuille Param) : ajouter une
'   ligne, remplir les criteres voulus (laisser vide ce qui ne doit pas filtrer)
'   et le decalage. Aucune modification de code necessaire.
'
' AJOUT 03/10/2026 (demande operateur) : DECALAGE MANUEL D'UNE OPERATION PRECISE.
'   En plus des regles generales ci-dessus (qui peuvent s'appliquer a PLUSIEURS
'   operations partageant le meme Tiers/Categorie/SousCategorie), l'operateur peut
'   desormais forcer un decalage pour UNE SEULE operation, depuis l'ecran de
'   recherche (bouton "Decaler le budget de cette operation", voir
'   mod_RechercheOperations.DecalerBudgetOperationRO). Ce choix est ecrit dans une
'   nouvelle colonne technique de TblOperations, "DecalageManuel" (voir
'   mod_VarGlobales.NOM_COL_DECALAGE_MANUEL / colDecalageManuel), qui MARQUE la
'   ligne : si un futur recalcul automatique global devait un jour exister, il
'   devra imperativement commencer par verifier cette colonne et ne JAMAIS modifier
'   une ligne ou elle est renseignee (non vide) - le choix de l'operateur doit
'   toujours primer sur les regles generales pour cette operation precise.
' =====================================================================================


' =====================================================================================
' INSTALLATION : cree le tableau TblDecalagesBudget s'il n'existe pas encore, avec en
' premiere ligne la regle DRFIP (migree depuis l'ancien codage en dur de
' mod_ImportOFX.CalculerBudget). Ne fait rien si le tableau existe deja : ne JAMAIS
' ecraser des regles que l'operateur aurait deja ajoutees ou modifiees.
' =====================================================================================
Public Sub InitialiserTableDecalagesBudget()

    Dim wsParam As Worksheet
    Dim tbl As ListObject
    Dim colDebut As Long, nbCol As Long

    colDebut = mod_VarGlobales.DECALAGES_BUDGET_COL_DEBUT
    nbCol = mod_VarGlobales.DECALAGES_BUDGET_NB_COL

    On Error Resume Next
    Set wsParam = ThisWorkbook.Worksheets(mod_Categories.NOM_FEUILLE_PARAM)
    On Error GoTo 0
    If wsParam Is Nothing Then
        MsgBox "La feuille " & mod_Categories.NOM_FEUILLE_PARAM & " est introuvable.", vbExclamation, "Decalages de budget"
        Exit Sub
    End If

    On Error Resume Next
    Set tbl = wsParam.ListObjects(mod_VarGlobales.NOM_TABLE_DECALAGES_BUDGET)
    On Error GoTo 0
    If Not tbl Is Nothing Then
        MsgBox "Le tableau " & mod_VarGlobales.NOM_TABLE_DECALAGES_BUDGET & " existe deja. Rien n'a ete modifie.", _
               vbInformation, "Decalages de budget"
        Exit Sub
    End If

    ' --- En-tetes ---
    wsParam.Range(wsParam.Cells(1, colDebut), wsParam.Cells(1, colDebut + nbCol - 1)).value = _
        Array("Tiers", "Categorie", "SousCategorie", "Decalage")

    ' --- Format TEXTE sur les 3 premieres colonnes avant d'ecrire (meme precaution que
    ' partout ailleurs dans ce projet : un Tiers ou une Categorie qui ressemble a un
    ' nombre ne doit jamais etre convertie silencieusement). La colonne Decalage, elle,
    ' reste un nombre normal (Standard). ---
    wsParam.Range(wsParam.Cells(2, colDebut), wsParam.Cells(2, colDebut + nbCol - 2)).NumberFormat = "@"

    ' --- Migration de l'ancienne regle codee en dur (DRFIP) ---
    wsParam.Cells(2, colDebut).value = "DRFIP OCCITANIE ET HTE"          ' Tiers
    wsParam.Cells(2, colDebut + 1).value = "Revenus"                     ' Categorie (parente)
    wsParam.Cells(2, colDebut + 2).value = "Salaire/Revenus d'activit" & Chr(233)  ' SousCategorie
    wsParam.Cells(2, colDebut + 3).value = 1                             ' Decalage (+1 mois)

    Set tbl = wsParam.ListObjects.Add(xlSrcRange, _
        wsParam.Range(wsParam.Cells(1, colDebut), wsParam.Cells(2, colDebut + nbCol - 1)), , xlYes)
    tbl.Name = mod_VarGlobales.NOM_TABLE_DECALAGES_BUDGET

    MsgBox "Tableau " & mod_VarGlobales.NOM_TABLE_DECALAGES_BUDGET & " cree en " & mod_Categories.NOM_FEUILLE_PARAM & "!S1, avec la regle DRFIP en 1ere ligne.", _
           vbInformation, "Decalages de budget"

End Sub


' =====================================================================================
' LECTURE : renvoie le decalage (en mois) a appliquer pour une operation donnee, ou 0
' si aucune ligne du tableau ne correspond (comportement identique a aujourd'hui :
' budget = mois de l'operation). Chaque critere laisse VIDE dans le tableau agit comme
' un joker ("n'importe quelle valeur") ; la comparaison ignore la casse et les espaces
' en debut/fin (memes conventions que le reste du projet).
'
' En cas de PLUSIEURS lignes correspondantes, c'est la PREMIERE rencontree (de haut en
' bas dans le tableau) qui s'applique : a l'operateur de placer ses regles les plus
' specifiques en premier si un jour plusieurs regles pouvaient se chevaucher.
' =====================================================================================
Public Function ObtenirDecalageBudget(ByVal tiers As String, ByVal categorie As String, ByVal sousCategorie As String) As Long

    Dim wsParam As Worksheet
    Dim tbl As ListObject
    Dim donnees As Variant
    Dim i As Long
    Dim critTiers As String, critCategorie As String, critSousCategorie As String

    ObtenirDecalageBudget = 0   ' valeur par defaut : aucun decalage

    On Error Resume Next
    Set wsParam = ThisWorkbook.Worksheets(mod_Categories.NOM_FEUILLE_PARAM)
    On Error GoTo 0
    If wsParam Is Nothing Then Exit Function

    On Error Resume Next
    Set tbl = wsParam.ListObjects(mod_VarGlobales.NOM_TABLE_DECALAGES_BUDGET)
    On Error GoTo 0
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    donnees = tbl.DataBodyRange.value   ' colonnes 1 a 4 = Tiers/Categorie/SousCategorie/Decalage

    For i = 1 To UBound(donnees, 1)
        critTiers = Trim(mod_DataStructure.CellText(donnees(i, 1)))
        critCategorie = Trim(mod_DataStructure.CellText(donnees(i, 2)))
        critSousCategorie = Trim(mod_DataStructure.CellText(donnees(i, 3)))

        If (critTiers = "" Or StrComp(critTiers, Trim(tiers), vbTextCompare) = 0) And _
           (critCategorie = "" Or StrComp(critCategorie, Trim(categorie), vbTextCompare) = 0) And _
           (critSousCategorie = "" Or StrComp(critSousCategorie, Trim(sousCategorie), vbTextCompare) = 0) Then
            ObtenirDecalageBudget = CLng(donnees(i, 4))
            Exit Function
        End If
    Next i

End Function


' =====================================================================================
' CALCUL DE DATE (ajout 03/10/2026, demande operateur) : SEULE fonction de tout le
' projet qui transforme une date comptable + un decalage (en mois) en date de budget
' (toujours ramenee au 1er du mois). Reutilisee a la fois par mod_ImportOFX.CalculerBudget
' (decalage lu dans TblDecalagesBudget) et par AppliquerDecalageManuel ci-dessous
' (decalage impose par l'operateur pour une seule operation) : ainsi, quel que soit le
' chemin de code empreinte, le budget d'une operation est TOUJOURS calcule de la meme
' maniere - aucune formule dupliquee ailleurs dans le projet.
' DateSerial() gere seul le changement d'annee (mois 13 -> janvier annee+1, et de la
' meme facon mois 0 -> decembre annee-1 pour un decalage negatif).
' =====================================================================================
Public Function CalculerDateBudget(ByVal dateComptable As Date, ByVal decalage As Long) As Date
    CalculerDateBudget = DateSerial(Year(dateComptable), Month(dateComptable) + decalage, 1)
End Function


' =====================================================================================
' INSTALLATION : ajoute la colonne technique "DecalageManuel" a la fin de TblOperations,
' si elle n'existe pas deja (meme principe, et meme prudence, que
' mod_Categories.AjouterColonneSousCategorie : ajout TOUJOURS a la fin du tableau, pour
' ne deplacer aucune colonne existante ; ne fait rien si la colonne existe deja).
' A executer UNE FOIS, par Ctrl+G : AjouterColonneDecalageManuel
' =====================================================================================
Public Sub AjouterColonneDecalageManuel()

    Dim tblOps As ListObject
    Dim colonne As ListColumn

    Set tblOps = mod_DonneesTable.GetOperationsTable()
    If tblOps Is Nothing Then Exit Sub

    On Error Resume Next
    Set colonne = tblOps.ListColumns(mod_VarGlobales.NOM_COL_DECALAGE_MANUEL)
    On Error GoTo 0

    If Not colonne Is Nothing Then
        MsgBox "La colonne " & mod_VarGlobales.NOM_COL_DECALAGE_MANUEL & " existe deja dans " & _
               tblOps.Name & ". Rien n'a ete modifie.", vbInformation, "Decalage manuel"
        Exit Sub
    End If

    ' ListColumns.Add sans argument = ajout apres la derniere colonne du tableau.
    Set colonne = tblOps.ListColumns.Add
    colonne.Name = mod_VarGlobales.NOM_COL_DECALAGE_MANUEL

    MsgBox "Colonne " & mod_VarGlobales.NOM_COL_DECALAGE_MANUEL & " ajoutee a la fin de " & tblOps.Name & "." & vbCrLf & _
           "Elle reste vide pour toutes les operations tant que l'operateur ne force pas de decalage manuel.", _
           vbInformation, "Decalage manuel"

End Sub


' =====================================================================================
' ECRITURE : force le decalage de budget d'UNE SEULE operation, identifiee par son
' ID_Transaction. Retrouve la ligne dans TblOperations (meme mecanisme de recherche
' que mod_RechercheOperations.RevoirVentilationRO : lecture de colID colonne par
' colonne, pas de formule), ecrit le decalage choisi dans la colonne technique
' "DecalageManuel" (ce qui MARQUE la ligne, voir l'en-tete de ce module), puis
' recalcule Budget/MoisBudget/AnneeBudget via CalculerDateBudget ci-dessus - jamais
' via ObtenirDecalageBudget, puisque l'operateur a choisi ce decalage explicitement et
' qu'il ne doit pas etre recalcule a partir des regles generales.
'
' Renvoie True si la ligne a bien ete retrouvee et mise a jour, False sinon (et affiche
' lui-meme un message d'erreur dans ce cas - l'appelant n'a rien d'autre a faire).
' =====================================================================================
Public Function AppliquerDecalageManuel(ByVal idTransaction As String, ByVal nouveauDecalage As Long) As Boolean

    Dim tblOp As ListObject
    Dim donneesOp As Variant
    Dim i As Long, ligneParent As Long
    Dim dateOp As Date
    Dim nouveauBudget As Date

    AppliquerDecalageManuel = False

    Set tblOp = mod_DonneesTable.GetOperationsTable()
    If tblOp Is Nothing Or tblOp.DataBodyRange Is Nothing Then
        MsgBox "Le tableau TblOperations est introuvable.", vbExclamation, "Decalage manuel"
        Exit Function
    End If

    mod_Display.RecupIndexCol
    If colDecalageManuel = 0 Then
        MsgBox "La colonne " & mod_VarGlobales.NOM_COL_DECALAGE_MANUEL & " n'existe pas encore." & vbCrLf & _
               "Execute d'abord (Ctrl+G) : AjouterColonneDecalageManuel", vbExclamation, "Decalage manuel"
        Exit Function
    End If

    donneesOp = tblOp.DataBodyRange.value

    ligneParent = 0
    For i = 1 To UBound(donneesOp, 1)
        If mod_DataStructure.CellText(donneesOp(i, colID)) = idTransaction Then
            ligneParent = i
            Exit For
        End If
    Next i

    If ligneParent = 0 Then
        MsgBox "Operation introuvable dans TblOperations (ID '" & idTransaction & "').", vbExclamation, "Decalage manuel"
        Exit Function
    End If

    dateOp = CDate(donneesOp(ligneParent, colDate))
    nouveauBudget = CalculerDateBudget(dateOp, nouveauDecalage)

    tblOp.DataBodyRange.Cells(ligneParent, colDecalageManuel).value = nouveauDecalage
    tblOp.DataBodyRange.Cells(ligneParent, colBudget).value = nouveauBudget
    tblOp.DataBodyRange.Cells(ligneParent, colBudget).NumberFormat = "mm/yyyy"
    tblOp.DataBodyRange.Cells(ligneParent, colMoisBud).value = Month(nouveauBudget)
    tblOp.DataBodyRange.Cells(ligneParent, colAnneeBud).value = Year(nouveauBudget)

    AppliquerDecalageManuel = True

End Function


' =====================================================================================
' ECRITURE : ajoute une nouvelle REGLE GENERALE dans TblDecalagesBudget (feuille
' Param), a partir de criteres deja connus (typiquement relus sur une operation
' selectionnee dans l'ecran de recherche - voir
' mod_RechercheOperations.AjouterDecalageDepuisRO). Contrairement a
' AppliquerDecalageManuel ci-dessus (qui ne modifie qu'UNE operation), une regle
' ajoutee ici s'appliquera a TOUTES les operations qui correspondent aux criteres,
' y compris celles qui seront importees plus tard.
'
' Meme precaution que partout ailleurs dans ce module : format TEXTE sur les 3
' premieres colonnes avant d'ecrire (pattern repris de
' mod_Categories.AjouterCategoriePersonnalisee), pour qu'un Tiers qui ressemble a un
' nombre ne soit jamais converti silencieusement.
'
' Renvoie True si la ligne a bien ete ajoutee, False si le tableau est introuvable
' (et affiche lui-meme un message d'erreur dans ce cas).
' =====================================================================================
Public Function AjouterRegleDecalage(ByVal tiers As String, ByVal categorie As String, _
                                      ByVal sousCategorie As String, ByVal decalage As Long) As Boolean

    Dim wsParam As Worksheet
    Dim tbl As ListObject
    Dim nouvelleLigne As ListRow
    Dim colDebut As Long, nbCol As Long

    AjouterRegleDecalage = False

    colDebut = mod_VarGlobales.DECALAGES_BUDGET_COL_DEBUT
    nbCol = mod_VarGlobales.DECALAGES_BUDGET_NB_COL

    On Error Resume Next
    Set wsParam = ThisWorkbook.Worksheets(mod_Categories.NOM_FEUILLE_PARAM)
    On Error GoTo 0
    If wsParam Is Nothing Then
        MsgBox "La feuille " & mod_Categories.NOM_FEUILLE_PARAM & " est introuvable.", vbExclamation, "Decalages de budget"
        Exit Function
    End If

    On Error Resume Next
    Set tbl = wsParam.ListObjects(mod_VarGlobales.NOM_TABLE_DECALAGES_BUDGET)
    On Error GoTo 0
    If tbl Is Nothing Then
        MsgBox "Le tableau " & mod_VarGlobales.NOM_TABLE_DECALAGES_BUDGET & " n'existe pas encore." & vbCrLf & _
               "Execute d'abord (Ctrl+G) : InitialiserTableDecalagesBudget", vbExclamation, "Decalages de budget"
        Exit Function
    End If

    Set nouvelleLigne = tbl.ListRows.Add
    ' Format TEXTE sur les 3 premieres colonnes (Tiers/Categorie/SousCategorie) AVANT
    ' d'y ecrire quoi que ce soit - meme precaution qu'a l'installation du tableau.
    wsParam.Range(wsParam.Cells(nouvelleLigne.Range.Row, colDebut), _
                  wsParam.Cells(nouvelleLigne.Range.Row, colDebut + nbCol - 2)).NumberFormat = "@"

    wsParam.Cells(nouvelleLigne.Range.Row, colDebut).value = tiers               ' Tiers
    wsParam.Cells(nouvelleLigne.Range.Row, colDebut + 1).value = categorie       ' Categorie
    wsParam.Cells(nouvelleLigne.Range.Row, colDebut + 2).value = sousCategorie   ' SousCategorie
    wsParam.Cells(nouvelleLigne.Range.Row, colDebut + 3).value = decalage        ' Decalage

    AjouterRegleDecalage = True

End Function
