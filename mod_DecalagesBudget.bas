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
