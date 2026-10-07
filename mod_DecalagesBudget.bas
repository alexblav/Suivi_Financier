Option Explicit

' =====================================================================================
' MODULE : mod_DecalagesBudget
'
' RÔLE (ajout du 03/10/2026, demande opérateur) :
'   Avant ce module, une seule règle était codée en dur dans
'   mod_ImportOFX.CalculerBudget : les salaires versés par "DRFIP OCCITANIE ET HTE"
'   étaient affectés au mois SUIVANT celui de l'opération (le virement de fin de
'   mois correspond au salaire du mois qui commence). Cette règle a cessé de
'   fonctionner silencieusement lors du passage à la hiérarchie Catégorie/Sous-catégorie
'   (phase 5) : le test comparait la colonne Categorie à "Salaire/Revenus
'   d'activite", devenue une SOUS-CATÉGORIE de "Revenus". Aucune opération ne pouvait
'   donc correspondre.
'
'   Ce module remplace cette règle par un TABLEAU DE CORRESPONDANCE modifiable
'   par l'opérateur sans toucher au code : le tableau "TblDecalagesBudget" de la
'   feuille Param (colonnes S à V, voir mod_VarGlobales). Chaque ligne précise un
'   critère (Tiers / Categorie / Sous-categorie; un champ laissé VIDE agit comme
'   un joker, c'est-à-dire "n'importe quelle valeur") et le décalage à appliquer,
'   exprimé en nombre de mois (+1, -1, 0 ou toute autre valeur entière si nécessaire).
'
'   Une opération qui ne correspond à AUCUNE ligne du tableau conserve le
'   comportement actuel (aucun décalage, budget = mois de l'opération).
'
' COLONNES DU TABLEAU "TblDecalagesBudget" (feuille Param, S:V) :
'   S = Tiers          : texte exact du Tiers, ou VIDE = n'importe quel Tiers.
'   T = Categorie      : catégorie PARENTE (ex. : "Revenus"), ou VIDE = joker.
'   U = SousCategorie  : sous-catégorie (ex. : "Salaire/Revenus d'activite"), ou
'                        VIDE = joker.
'   V = Decalage       : nombre de mois à ajouter à Date_Comptable pour obtenir
'                        le mois de budget (entier positif ou négatif).
'
' COMMENT INSTALLER / RÉINSTALLER LE TABLEAU :
'   Ctrl+G (fenêtre Exécution), taper : InitialiserTableDecalagesBudget
'   Sans danger si le tableau existe déjà : la procédure ne fait rien dans ce cas
'   (même principe que mod_Categories.AjouterColonneSousCategorie).
'
' COMMENT AJOUTER UNE NOUVELLE RÈGLE DE DÉCALAGE :
'   Directement dans le tableau TblDecalagesBudget (feuille Param) : ajouter une
'   ligne, remplir les critères voulus (laisser vide ce qui ne doit pas filtrer)
'   et le décalage. Aucune modification du code n'est nécessaire.
'
' AJOUT du 03/10/2026 (demande opérateur) : DÉCALAGE MANUEL D'UNE OPÉRATION PRÉCISE.
'   En plus des règles générales ci-dessus (qui peuvent s'appliquer à PLUSIEURS
'   opérations partageant le même Tiers/Categorie/SousCategorie), l'opérateur peut
'   désormais forcer un décalage pour UNE SEULE opération depuis l'écran de recherche
'   (bouton "Decaler le budget de cette operation", voir
'   mod_RechercheOperations.DecalerBudgetOperationRO). Ce choix est écrit dans une
'   nouvelle colonne technique de TblOperations, "DecalageManuel" (voir
'   mod_VarGlobales.NOM_COL_DECALAGE_MANUEL / colDecalageManuel), qui MARQUE la ligne.
'   Si un recalcul automatique global devait un jour exister, il devra impérativement
'   vérifier cette colonne en premier et ne JAMAIS modifier une ligne où elle est
'   renseignée (non vide). Le choix de l'opérateur doit toujours primer sur les règles
'   générales pour cette opération précise.
' =====================================================================================


' =====================================================================================
' INSTALLATION : crée le tableau TblDecalagesBudget s'il n'existe pas encore, avec en
' première ligne la règle DRFIP (migrée depuis l'ancien codage en dur de
' mod_ImportOFX.CalculerBudget). Ne fait rien si le tableau existe déjà : ne JAMAIS
' écraser les règles que l'opérateur aurait déjà ajoutées ou modifiées.
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

    ' --- En-têtes ---
    wsParam.Range(wsParam.Cells(1, colDebut), wsParam.Cells(1, colDebut + nbCol - 1)).value = _
        Array("Tiers", "Categorie", "SousCategorie", "Decalage")

    ' --- Format TEXTE sur les trois premières colonnes avant d'écrire (même précaution
    ' que partout ailleurs dans ce projet : un Tiers ou une Categorie ressemblant à un
    ' nombre ne doit jamais être converti silencieusement). La colonne Decalage, elle,
    ' reste un nombre normal (Standard). ---
    wsParam.Range(wsParam.Cells(2, colDebut), wsParam.Cells(2, colDebut + nbCol - 2)).NumberFormat = "@"

    ' --- Migration de l'ancienne règle codée en dur (DRFIP) ---
    wsParam.Cells(2, colDebut).value = "DRFIP OCCITANIE ET HTE"          ' Tiers
    wsParam.Cells(2, colDebut + 1).value = "Revenus"                     ' Categorie (parente)
    wsParam.Cells(2, colDebut + 2).value = "Salaire/Revenus d'activit" & Chr(233)  ' SousCategorie
    wsParam.Cells(2, colDebut + 3).value = 1                             ' Décalage (+1 mois)

    Set tbl = wsParam.ListObjects.Add(xlSrcRange, _
        wsParam.Range(wsParam.Cells(1, colDebut), wsParam.Cells(2, colDebut + nbCol - 1)), , xlYes)
    tbl.Name = mod_VarGlobales.NOM_TABLE_DECALAGES_BUDGET

    MsgBox "Tableau " & mod_VarGlobales.NOM_TABLE_DECALAGES_BUDGET & " cree en " & mod_Categories.NOM_FEUILLE_PARAM & "!S1, avec la regle DRFIP en 1ere ligne.", _
           vbInformation, "Decalages de budget"

End Sub


' =====================================================================================
' LECTURE : renvoie le décalage (en mois) à appliquer pour une opération donnée, ou 0
' si aucune ligne du tableau ne correspond (comportement identique à aujourd'hui :
' budget = mois de l'opération). Chaque critère laissé VIDE dans le tableau agit comme
' un joker ("n'importe quelle valeur"); la comparaison ignore la casse et les espaces
' en début/fin (mêmes conventions que le reste du projet).
'
' En cas de PLUSIEURS lignes correspondantes, c'est la PREMIÈRE rencontrée (de haut en
' bas dans le tableau) qui s'applique : à l'opérateur de placer ses règles les plus
' spécifiques en premier si plusieurs règles venaient à se chevaucher.
' =====================================================================================
Public Function ObtenirDecalageBudget(ByVal tiers As String, ByVal categorie As String, ByVal sousCategorie As String) As Long

    Dim wsParam As Worksheet
    Dim tbl As ListObject
    Dim donnees As Variant
    Dim i As Long
    Dim critTiers As String, critCategorie As String, critSousCategorie As String

    ObtenirDecalageBudget = 0   ' valeur par défaut : aucun décalage

    On Error Resume Next
    Set wsParam = ThisWorkbook.Worksheets(mod_Categories.NOM_FEUILLE_PARAM)
    On Error GoTo 0
    If wsParam Is Nothing Then Exit Function

    On Error Resume Next
    Set tbl = wsParam.ListObjects(mod_VarGlobales.NOM_TABLE_DECALAGES_BUDGET)
    On Error GoTo 0
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    donnees = tbl.DataBodyRange.value   ' colonnes 1 à 4 = Tiers/Categorie/SousCategorie/Decalage

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
' CALCUL DE DATE (ajout du 03/10/2026, demande opérateur) : SEULE fonction du projet
' qui transforme une date comptable et un décalage (en mois) en date de budget
' (toujours ramenée au premier du mois). Elle est réutilisée par mod_ImportOFX.CalculerBudget
' (décalage lu dans TblDecalagesBudget) et par AppliquerDecalageManuel ci-dessous
' (décalage imposé par l'opérateur pour une seule opération). Ainsi, quel que soit le
' chemin de code emprunté, le budget d'une opération est TOUJOURS calculé de la même
' manière, sans formule dupliquée ailleurs dans le projet.
' DateSerial() gère seul le changement d'année (mois 13 -> janvier de l'année suivante;
' de même, mois 0 -> décembre de l'année précédente pour un décalage négatif).
' =====================================================================================
Public Function CalculerDateBudget(ByVal dateComptable As Date, ByVal decalage As Long) As Date
    CalculerDateBudget = DateSerial(Year(dateComptable), Month(dateComptable) + decalage, 1)
End Function


' =====================================================================================
' INSTALLATION : ajoute la colonne technique "DecalageManuel" à la fin de TblOperations,
' si elle n'existe pas déjà (même principe et même prudence que
' mod_Categories.AjouterColonneSousCategorie : ajout TOUJOURS à la fin du tableau afin
' de ne déplacer aucune colonne existante; ne fait rien si la colonne existe déjà).
' À exécuter UNE FOIS, par Ctrl+G : AjouterColonneDecalageManuel
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

    ' ListColumns.Add sans argument ajoute la colonne après la dernière du tableau.
    Set colonne = tblOps.ListColumns.Add
    colonne.Name = mod_VarGlobales.NOM_COL_DECALAGE_MANUEL

    MsgBox "Colonne " & mod_VarGlobales.NOM_COL_DECALAGE_MANUEL & " ajoutee a la fin de " & tblOps.Name & "." & vbCrLf & _
           "Elle reste vide pour toutes les operations tant que l'operateur ne force pas de decalage manuel.", _
           vbInformation, "Decalage manuel"

End Sub


' =====================================================================================
' ÉCRITURE : force le décalage de budget d'UNE SEULE opération, identifiée par son
' ID_Transaction. Retrouve la ligne dans TblOperations (même mécanisme de recherche
' que mod_RechercheOperations.RevoirVentilationRO : lecture de colID colonne par colonne,
' sans formule), écrit le décalage choisi dans la colonne technique "DecalageManuel"
' (ce qui MARQUE la ligne; voir l'en-tête du module), puis recalcule
' Budget/MoisBudget/AnneeBudget via CalculerDateBudget ci-dessus, jamais via
' ObtenirDecalageBudget : l'opérateur a choisi ce décalage explicitement, il ne doit
' donc pas être recalculé à partir des règles générales.
'
' Renvoie True si la ligne a été trouvée et mise à jour, False sinon (et affiche alors
' un message d'erreur; l'appelant n'a rien d'autre à faire).
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
' ÉCRITURE : ajoute une nouvelle RÈGLE GÉNÉRALE dans TblDecalagesBudget (feuille Param),
' à partir de critères déjà connus (généralement lus sur une opération sélectionnée
' dans l'écran de recherche; voir mod_RechercheOperations.AjouterDecalageDepuisRO).
' Contrairement à AppliquerDecalageManuel (qui ne modifie qu'UNE opération), une règle
' ajoutée ici s'appliquera à TOUTES les opérations correspondant aux critères, y compris
' celles qui seront importées plus tard.
'
' Même précaution que partout ailleurs dans ce module : format TEXTE sur les trois
' premières colonnes avant d'écrire (méthode reprise de
' mod_Categories.AjouterCategoriePersonnalisee), afin qu'un Tiers ressemblant à un
' nombre ne soit jamais converti silencieusement.
'
' Renvoie True si la ligne a été ajoutée, False si le tableau est introuvable
' (et affiche alors un message d'erreur).
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
    ' Format TEXTE sur les trois premières colonnes (Tiers/Categorie/SousCategorie) AVANT
    ' d'y écrire quoi que ce soit, par précaution comme lors de l'installation du tableau.
    wsParam.Range(wsParam.Cells(nouvelleLigne.Range.Row, colDebut), _
                  wsParam.Cells(nouvelleLigne.Range.Row, colDebut + nbCol - 2)).NumberFormat = "@"

    wsParam.Cells(nouvelleLigne.Range.Row, colDebut).value = tiers               ' Tiers
    wsParam.Cells(nouvelleLigne.Range.Row, colDebut + 1).value = categorie       ' Categorie
    wsParam.Cells(nouvelleLigne.Range.Row, colDebut + 2).value = sousCategorie   ' SousCategorie
    wsParam.Cells(nouvelleLigne.Range.Row, colDebut + 3).value = decalage        ' Decalage

    AjouterRegleDecalage = True

End Function
