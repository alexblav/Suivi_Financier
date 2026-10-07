Option Explicit

' =====================================================================================
' MODULE : mod_Categories
'
' PHASE 1 du chantier « Catégorie / Sous-catégorie / Ventilation ».
'
' RÔLE DE CE MODULE (à lire en premier, même si vous débutez) :
'
'   Aujourd'hui, une opération n'a qu'UNE catégorie, sous forme de texte simple
'   (ex. : "Frais, remb santé"). Le nouveau fonctionnement demande DEUX niveaux :
'        Catégorie      (ex. : "Santé, prévoyance")
'        Sous-catégorie (ex. : "Frais, remb santé")
'
'   Les fichiers sources de la banque (OFX + CSV) ne fournissent toujours qu'UNE
'   catégorie : c'est la "CATÉGORIE SOURCE". Il faut donc un TABLEAU DE
'   CORRESPONDANCE qui indique, pour chaque catégorie source :
'        "quand la banque m'envoie X, je la range dans Catégorie Y / Sous-catégorie Z".
'
'   Ce module met en place les deux briques de base, RIEN d'autre :
'     1. La colonne "SousCategorie" ajoutée à la fin de TblOperations.
'     2. Le tableau "TblCategories" dans la feuille Param (colonnes K à N), prérempli
'        avec toutes les catégories déjà présentes dans vos opérations.
'
'   CE MODULE NE MODIFIE AUCUNE DONNÉE EXISTANTE de TblOperations : les valeurs de
'   la colonne Categorie ne sont pas touchées. La conversion des anciennes opérations
'   viendra dans une phase ultérieure, après votre validation du tableau.
'
'   Il ne modifie PAS non plus les autres modules (mod_ImportOFX, mod_Display, ...).
'
' COMMENT UTILISER (dans l'ordre) :
'   1. Enregistrer une COPIE du classeur (précaution : une colonne va être ajoutée).
'   2. Ctrl+G (fenêtre Exécution), taper : PreparerPhase1Categories, puis Entrée.
'   3. Sur la feuille Param qui s'affiche, vérifier ou corriger le tableau K:N,
'      puis mettre "Oui" dans la colonne "Verifie" pour chaque ligne contrôlée.
'   4. Quand vous avez fini : MasquerParam (Ctrl+G) pour cacher la feuille.
'
' À PROPOS DES ACCENTS : les textes destinés à Excel sont construits à l'exécution
' avec ChrW() ou la fonction AccentsFR (voir plus bas). Les commentaires sont en UTF-8.
' =====================================================================================

' --- Nom de la feuille et du tableau de correspondance ---------------------------------
Public Const NOM_FEUILLE_PARAM As String = "Param"
Public Const NOM_TABLE_CATEGORIES As String = "TblCategories"

' --- Nom de la colonne ajoutée dans TblOperations -----------------------------------
Public Const NOM_COL_SOUS_CATEGORIE As String = "SousCategorie"

' --- Position du tableau dans la feuille Param ----------------------------------------
' Les colonnes A à I de Param sont déjà utilisées (listes Mois, Années, Médecins...).
' On place donc le nouveau tableau plus à droite, en laissant la colonne J vide.
Private Const COL_DEBUT As Long = 11        ' colonne K : première colonne du tableau
Private Const NB_COL_TABLE As Long = 4      ' K, L, M, N = 4 colonnes
Private Const COL_LISTE As Long = 16        ' colonne P : liste technique des catégories

' Nom (au niveau du classeur) de la liste servant aux menus déroulants
Private Const NOM_LISTE_CATEGORIES As String = "ListeCategories"

' Les 4 colonnes du tableau, dans l'ordre :
'   K = CategorieSource : le texte EXACT envoyé par la banque (ne pas modifier)
'   L = Categorie       : la catégorie de rangement dans le nouvel outil
'   M = SousCategorie   : la sous-catégorie (peut rester vide)
'   N = Verifie         : "Oui" quand l'opérateur a contrôlé la ligne


' =====================================================================================
' PROCÉDURE D'ENSEMBLE : enchaîne les trois étapes de la phase 1.
' (Chaque étape reste utilisable seule; voir plus bas.)
' =====================================================================================
Public Sub PreparerPhase1Categories()

    ' Étape 1 : colonne SousCategorie dans TblOperations
    AjouterColonneSousCategorie

    ' Étape 2 : création ou mise à jour du tableau de correspondance
    InitialiserTableCategories

    ' Étape 3 : on affiche la feuille Param pour que l'opérateur puisse vérifier.
    AfficherParamCategories

End Sub


' =====================================================================================
' ÉTAPE 1 : ajoute la colonne "SousCategorie" à la fin de TblOperations
' =====================================================================================
' Pourquoi "à la fin" ? Parce que tout votre code lit les colonnes par nom
' (voir RecupIndexCol) ou par position pour les colonnes 1 à 20. En ajoutant la
' nouvelle colonne après la dernière (colonne U), aucune position existante ne bouge.
' Cette procédure peut être relancée sans danger : si la colonne existe déjà,
' elle ne fait rien.
Public Sub AjouterColonneSousCategorie()

    Dim tblOps As ListObject
    Dim colonne As ListColumn

    ' On réutilise la fonction existante qui retrouve TblOperations
    ' (elle affiche elle-même un message si le tableau est introuvable).
    Set tblOps = mod_DonneesTable.GetOperationsTable()
    If tblOps Is Nothing Then Exit Sub

    ' On vérifie si la colonne existe déjà. "On Error Resume Next" évite un plantage :
    ' demander une colonne inexistante provoque une erreur que l'on peut interpréter
    ' (colonne restée à Nothing = elle n'existe pas).
    On Error Resume Next
    Set colonne = tblOps.ListColumns(NOM_COL_SOUS_CATEGORIE)
    On Error GoTo 0

    If Not colonne Is Nothing Then
         MsgBox "La colonne " & NOM_COL_SOUS_CATEGORIE & " existe deja dans " & _
             tblOps.Name & ". Rien n'a ete modifie.", vbInformation, "Phase 1"
        Exit Sub
    End If

    ' ListColumns.Add sans argument ajoute une colonne après la dernière du tableau.
    Set colonne = tblOps.ListColumns.Add
    colonne.Name = NOM_COL_SOUS_CATEGORIE

    MsgBox "Colonne " & NOM_COL_SOUS_CATEGORIE & " ajoutee a la fin de " & tblOps.Name & ".", _
           vbInformation, "Phase 1"

End Sub


' =====================================================================================
' ÉTAPE 2 : crée ou complète le tableau de correspondance TblCategories (feuille Param)
' =====================================================================================
' - Première exécution : le tableau est créé avec UNE ligne par catégorie source
'   trouvée dans TblOperations (toutes les valeurs distinctes, triées de A à Z).
' - Aux exécutions suivantes, seules les NOUVELLES catégories source sont ajoutées en
'   bas. Les saisies de l'opérateur ne sont JAMAIS écrasées.
'
' Valeurs proposées au départ (l'opérateur peut tout modifier) :
'   - Categorie = le texte source recopié tel quel; SousCategorie et Verifie sont vides.
'   - Cas particulier décrit : "Frais, remb santé" est une sous-catégorie
'     de "Santé, prévoyance"; ces valeurs sont préremplies et Verifie = "Oui".
Public Sub InitialiserTableCategories()

    Dim tblOps As ListObject, tblCat As ListObject
    Dim wsParam As Worksheet
    Dim colonneCat As ListColumn
    Dim donnees As Variant, existantes As Variant
    Dim dejaVu As Object
    Dim nouvelles() As String
    Dim nbNouvelles As Long
    Dim i As Long
    Dim texte As String
    Dim Sortie() As Variant
    Dim premiereLigne As Long, derniereLigne As Long
    Dim nbAVerifier As Long, verifies As Variant

    ' --- Récupérer TblOperations et la feuille Param -------------------------------
    Set tblOps = mod_DonneesTable.GetOperationsTable()
    If tblOps Is Nothing Then Exit Sub
    If tblOps.DataBodyRange Is Nothing Then
        MsgBox "TblOperations ne contient aucune ligne.", vbInformation, "Phase 1"
        Exit Sub
    End If

    On Error Resume Next
    Set colonneCat = tblOps.ListColumns("Categorie")
    Set wsParam = ThisWorkbook.Worksheets(NOM_FEUILLE_PARAM)
    On Error GoTo 0

    If colonneCat Is Nothing Then
        MsgBox "La colonne Categorie est introuvable dans TblOperations.", vbExclamation, "Phase 1"
        Exit Sub
    End If
    If wsParam Is Nothing Then
        MsgBox "La feuille " & NOM_FEUILLE_PARAM & " est introuvable.", vbExclamation, "Phase 1"
        Exit Sub
    End If

    ' --- Lire toutes les catégories source en mémoire (un seul aller-retour = rapide) ---
    donnees = LireColonne(colonneCat.DataBodyRange)

    ' --- "dejaVu" : dictionnaire qui mémorise les catégories déjà traitées ------------
    ' CompareMode = 1 (vbTextCompare) : on ignore majuscules/minuscules.
    ' Il doit être réglé AVANT d'ajouter le moindre élément.
    Set dejaVu = CreateObject("Scripting.Dictionary")
    dejaVu.CompareMode = 1

    ' Si le tableau existe déjà, on y lit les sources connues pour éviter les doublons.
    Set tblCat = TrouverTable(wsParam, NOM_TABLE_CATEGORIES)
    If Not tblCat Is Nothing Then
        If Not tblCat.DataBodyRange Is Nothing Then
            existantes = LireColonne(tblCat.ListColumns("CategorieSource").DataBodyRange)
            For i = 1 To UBound(existantes, 1)
                texte = mod_DataStructure.CellText(existantes(i, 1))
                If texte <> "" Then
                    If Not dejaVu.Exists(texte) Then dejaVu.Add texte, True
                End If
            Next i
        End If
    End If

    ' --- Repérer les catégories source qui ne sont pas encore dans le tableau -----------
    ReDim nouvelles(1 To UBound(donnees, 1))
    nbNouvelles = 0
    For i = 1 To UBound(donnees, 1)
        ' CellText : lecture "sécurisée" (jamais d'erreur, espaces inutiles retirés)
        texte = mod_DataStructure.CellText(donnees(i, 1))
        If texte <> "" Then
            If Not dejaVu.Exists(texte) Then
                dejaVu.Add texte, True
                nbNouvelles = nbNouvelles + 1
                nouvelles(nbNouvelles) = texte
            End If
        End If
    Next i

    ' --- Rien de nouveau : on s'assure seulement que les listes sont à jour -------------
    If nbNouvelles = 0 Then
        If Not tblCat Is Nothing Then
            RafraichirListesCategories
            PoserListesDeroulantes tblCat
        End If
        MsgBox "Aucune nouvelle categorie source a ajouter.", vbInformation, "Phase 1"
        Exit Sub
    End If

    ' --- Trier les nouvelles catégories de A à Z ------------------------------------------
    TrierTextes nouvelles, nbNouvelles

    ' --- Préparer le bloc à écrire (tableau en mémoire : 4 colonnes) -----------------------
    ReDim Sortie(1 To nbNouvelles, 1 To NB_COL_TABLE)
    For i = 1 To nbNouvelles
        Sortie(i, 1) = nouvelles(i)      ' CategorieSource
        Sortie(i, 2) = nouvelles(i)      ' Categorie : proposition = même texte
        Sortie(i, 3) = ""                ' SousCategorie : vide
        Sortie(i, 4) = ""                ' Verifie : à faire par l'opérateur

        ' Cas particulier décrit par l'opérateur : "Frais, remb santé" est une
        ' sous-catégorie de "Santé, prévoyance".
        If StrComp(nouvelles(i), CategorieSourceSante(), vbTextCompare) = 0 Then
            Sortie(i, 2) = CategorieSanteNouvelle()
            Sortie(i, 3) = CategorieSourceSante()
            Sortie(i, 4) = "Oui"
        End If
    Next i

    ' --- Écriture dans la feuille Param -----------------------------------------------------
    If tblCat Is Nothing Then
        ' PREMIÈRE FOIS : en-têtes, données et création du tableau Excel
        wsParam.Range(wsParam.Cells(1, COL_DEBUT), wsParam.Cells(1, COL_DEBUT + NB_COL_TABLE - 1)).value = _
            Array("CategorieSource", "Categorie", "SousCategorie", "Verifie")
        premiereLigne = 2
    Else
        ' Tableau déjà présent : on écrit sous la dernière ligne remplie. On se base sur
        ' la colonne L (Categorie), toujours remplie, contrairement à la colonne K
        ' (une catégorie créée à la main n'a pas de catégorie source).
        derniereLigne = wsParam.Cells(wsParam.rows.count, COL_DEBUT + 1).End(xlUp).Row
        premiereLigne = derniereLigne + 1
    End If

    ' Format TEXTE avant d'écrire : évite qu'Excel transforme une valeur en date ou en nombre
    ' (piège déjà rencontré dans ce projet).
    With wsParam.Range(wsParam.Cells(premiereLigne, COL_DEBUT), _
                       wsParam.Cells(premiereLigne + nbNouvelles - 1, COL_DEBUT + NB_COL_TABLE - 1))
        .NumberFormat = "@"
        .Value2 = Sortie
    End With

    If tblCat Is Nothing Then
        Set tblCat = wsParam.ListObjects.Add(xlSrcRange, _
            wsParam.Range(wsParam.Cells(1, COL_DEBUT), _
                          wsParam.Cells(premiereLigne + nbNouvelles - 1, COL_DEBUT + NB_COL_TABLE - 1)), , xlYes)
        tblCat.Name = NOM_TABLE_CATEGORIES
    Else
        ' On agrandit le tableau Excel pour qu'il englobe les lignes ajoutées.
        tblCat.Resize wsParam.Range(wsParam.Cells(1, COL_DEBUT), _
                          wsParam.Cells(premiereLigne + nbNouvelles - 1, COL_DEBUT + NB_COL_TABLE - 1))
    End If

    ' --- Listes déroulantes ---------------------------------------------------------------------
    RafraichirListesCategories
    PoserListesDeroulantes tblCat

    ' --- Petit bilan : combien de lignes restent à contrôler ? -----------------------------------
    verifies = LireColonne(tblCat.ListColumns("Verifie").DataBodyRange)
    For i = 1 To UBound(verifies, 1)
        If StrComp(mod_DataStructure.CellText(verifies(i, 1)), "Oui", vbTextCompare) <> 0 Then
            nbAVerifier = nbAVerifier + 1
        End If
    Next i

    MsgBox nbNouvelles & " categorie(s) source ajoutee(s) au tableau " & NOM_TABLE_CATEGORIES & "." & _
           vbCrLf & vbCrLf & _
           "Lignes restant a verifier (colonne Verifie <> Oui) : " & nbAVerifier & vbCrLf & vbCrLf & _
           mod_Display.FR("Pour chaque ligne : contr{o2}lez Cat{e2}gorie et Sous-cat{e2}gorie, puis tapez Oui dans Verifie."), _
           vbInformation, "Phase 1"

End Sub


' =====================================================================================
' Met à jour la liste TECHNIQUE des catégories distinctes (colonne P de Param).
' =====================================================================================
' Cette liste alimente le menu déroulant de la colonne "Categorie". Elle est
' recalculée à partir du tableau : toute catégorie ajoutée au tableau apparaît
' dans le menu après l'appel de cette procédure. Elle sera aussi réutilisée par
' les futurs formulaires (phases 2 et 3).
'
' Pourquoi ne pas mettre les catégories dans le menu sous forme de texte séparé par
' des virgules ? Parce que vos catégories contiennent elles-mêmes des virgules
' ("Alimentation, supermarché") : Excel les couperait en morceaux. On passe donc
' par des cellules, et le menu pointe vers cette plage.
Public Sub RafraichirListesCategories()

    Dim wsParam As Worksheet
    Dim tblCat As ListObject
    Dim donnees As Variant
    Dim dejaVu As Object
    Dim liste() As String
    Dim Sortie() As Variant
    Dim nb As Long, i As Long
    Dim texte As String

    On Error Resume Next
    Set wsParam = ThisWorkbook.Worksheets(NOM_FEUILLE_PARAM)
    On Error GoTo 0
    If wsParam Is Nothing Then Exit Sub

    Set tblCat = TrouverTable(wsParam, NOM_TABLE_CATEGORIES)
    If tblCat Is Nothing Then Exit Sub
    If tblCat.DataBodyRange Is Nothing Then Exit Sub

    donnees = LireColonne(tblCat.ListColumns("Categorie").DataBodyRange)

    Set dejaVu = CreateObject("Scripting.Dictionary")
    dejaVu.CompareMode = 1
    ReDim liste(1 To UBound(donnees, 1))

    For i = 1 To UBound(donnees, 1)
        texte = mod_DataStructure.CellText(donnees(i, 1))
        If texte <> "" Then
            If Not dejaVu.Exists(texte) Then
                dejaVu.Add texte, True
                nb = nb + 1
                liste(nb) = texte
            End If
        End If
    Next i
    If nb = 0 Then Exit Sub

    TrierTextes liste, nb

    ' On efface l'ancienne liste (sous l'en-tête), puis on écrit la nouvelle.
    wsParam.Range(wsParam.Cells(2, COL_LISTE), wsParam.Cells(wsParam.rows.count, COL_LISTE)).ClearContents
    wsParam.Cells(1, COL_LISTE).value = "Liste categories (zone technique - ne pas modifier)"

    ReDim Sortie(1 To nb, 1 To 1)
    For i = 1 To nb
        Sortie(i, 1) = liste(i)
    Next i
    wsParam.Cells(2, COL_LISTE).Resize(nb, 1).NumberFormat = "@"
    wsParam.Cells(2, COL_LISTE).Resize(nb, 1).Value2 = Sortie

    ' Nom "ListeCategories" : plage DYNAMIQUE (OFFSET) qui s'adapte au nombre de
    ' catégories écrites. Créé au niveau du CLASSEUR (ThisWorkbook.Names), comme les
    ' autres noms du projet, pour éviter la confusion de portée déjà rencontrée.
    ThisWorkbook.Names.Add Name:=NOM_LISTE_CATEGORIES, _
        RefersTo:="=OFFSET(" & NOM_FEUILLE_PARAM & "!$" & ColonneEnLettre(COL_LISTE) & "$2,0,0," & _
                  "MAX(1,COUNTA(" & NOM_FEUILLE_PARAM & "!$" & ColonneEnLettre(COL_LISTE) & "$2:$" & _
                  ColonneEnLettre(COL_LISTE) & "$500)),1)"

End Sub


' =====================================================================================
' Pose les menus déroulants du tableau TblCategories
' =====================================================================================
'   - colonne Categorie : liste des catégories existantes, MAIS l'opérateur peut
'     saisir une nouvelle catégorie (alerte "avertissement", sans blocage).
'   - colonne Verifie   : uniquement "Oui" ou vide.
' La colonne SousCategorie reste en saisie libre pour l'instant : sa liste
' dépendante de la catégorie choisie sera gérée dans les formulaires (phase 2).
Private Sub PoserListesDeroulantes(ByVal tblCat As ListObject)

    If tblCat.DataBodyRange Is Nothing Then Exit Sub

    With tblCat.ListColumns("Categorie").DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, _
             Formula1:="=" & NOM_LISTE_CATEGORIES
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    With tblCat.ListColumns("Verifie").DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Oui"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

End Sub


' =====================================================================================
' Affiche la feuille Param (normalement masquée) et se place sur le tableau.
' =====================================================================================
Public Sub AfficherParamCategories()

    Dim wsParam As Worksheet

    On Error Resume Next
    Set wsParam = ThisWorkbook.Worksheets(NOM_FEUILLE_PARAM)
    On Error GoTo 0
    If wsParam Is Nothing Then Exit Sub

    wsParam.Visible = xlSheetVisible
    wsParam.Activate
    wsParam.Cells(1, COL_DEBUT).Select

End Sub


' =====================================================================================
' Remasque la feuille Param (à lancer quand le contrôle du tableau est terminé).
' =====================================================================================
Public Sub MasquerParam()

    Dim wsParam As Worksheet

    On Error Resume Next
    Set wsParam = ThisWorkbook.Worksheets(NOM_FEUILLE_PARAM)
    On Error GoTo 0
    If wsParam Is Nothing Then Exit Sub

    ' Excel refuse de masquer la seule feuille visible : on repasse par Accueil si besoin.
    On Error Resume Next
    ThisWorkbook.Worksheets("Accueil").Activate
    On Error GoTo 0

    wsParam.Visible = xlSheetHidden

End Sub


' =====================================================================================
' PHASE 3 : création d'une catégorie ou sous-catégorie par l'opérateur
' =====================================================================================
' Ces deux fonctions sont utilisées par le formulaire de création (mod_NouvelleCategorie).
' Elles se trouvent ici, et non dans ce nouveau module, car TblCategories appartient
' à mod_Categories : c'est le seul endroit du projet qui sait comment ce tableau est
' construit (nom de la feuille, position des colonnes...).

' ---------------------------------------------------------------------------------------
' AjouterCategoriePersonnalisee : enregistre dans TblCategories une paire (Categorie,
' SousCategorie) choisie ou créée par l'opérateur, si elle n'y figure pas déjà.
' ---------------------------------------------------------------------------------------
' ByRef categorie / sousCategorie : en entrée, les valeurs saisies ou choisies par
' l'opérateur. En sortie (si la fonction renvoie True), ces variables prennent leur
' forme "canonique" : si "sante" existe déjà sous la forme "Sante" dans le tableau,
' categorie devient "Sante", afin d'éviter que deux graphies de la même catégorie
' cohabitent (ex. : "Sante" et "sante" traitées comme deux catégories différentes
' ailleurs dans le classeur).
'
' Renvoie False uniquement si la catégorie est vide ou si TblCategories est introuvable
' (le formulaire appelant doit alors afficher un message et s'arrêter).
Public Function AjouterCategoriePersonnalisee(ByRef categorie As String, ByRef sousCategorie As String) As Boolean

    Dim wsParam As Worksheet, tblCat As ListObject
    Dim cat As Variant, sous As Variant
    Dim r As Long, ligneCible As Long
    Dim texteCat As String, texteSous As String
    Dim casCatTrouvee As Boolean, casCatCanonique As String

    categorie = Trim(categorie)
    sousCategorie = Trim(sousCategorie)
    If categorie = "" Then Exit Function

    On Error Resume Next
    Set wsParam = ThisWorkbook.Worksheets(NOM_FEUILLE_PARAM)
    On Error GoTo 0
    If wsParam Is Nothing Then Exit Function

    Set tblCat = TrouverTable(wsParam, NOM_TABLE_CATEGORIES)
    If tblCat Is Nothing Then Exit Function
    If tblCat.DataBodyRange Is Nothing Then Exit Function

    cat = LireColonne(tblCat.ListColumns("Categorie").DataBodyRange)
    sous = LireColonne(tblCat.ListColumns("SousCategorie").DataBodyRange)

    ' --- Recherche d'une correspondance existante ------------------------------------------
    ' NIVEAU 1 du garde-fou orthographique (normalisation silencieuse, sans question) :
    ' on ignore les majuscules/minuscules, les accents et les espaces superflus.
    ' "Sante" et "Sant{e2}" sont ainsi reconnues comme une seule catégorie.
    For r = 1 To UBound(cat, 1)
        texteCat = mod_DataStructure.CellText(cat(r, 1))
        texteSous = mod_DataStructure.CellText(sous(r, 1))

        If NormaliserTexte(texteCat) = NormaliserTexte(categorie) Then
            casCatTrouvee = True
            casCatCanonique = texteCat            ' on garde la première graphie rencontrée

            If NormaliserTexte(texteSous) = NormaliserTexte(sousCategorie) Then
                ' La paire existe déjà (mêmes lettres, sans tenir compte des accents,
                ' espaces ou majuscules) : rien à ajouter. On renvoie les libellés EXACTS
                ' déjà présents dans le tableau.
                categorie = texteCat
                sousCategorie = texteSous
                AjouterCategoriePersonnalisee = True
                Exit Function
            End If
        End If
    Next r

    ' La catégorie existe déjà (dans une autre paire) : on reprend sa graphie exacte
    ' pour éviter que "Sante" et "sante" deviennent deux catégories différentes.
    If casCatTrouvee Then categorie = casCatCanonique

    ' --- Aucune paire correspondante : ajout d'une nouvelle ligne au tableau -------------
    ' CategorieSource reste VIDE : cette ligne ne provient pas d'un import bancaire,
    ' mais d'une création manuelle par l'opérateur.
    ligneCible = tblCat.ListRows.Add.Range.Row

    With wsParam.Range(wsParam.Cells(ligneCible, COL_DEBUT), wsParam.Cells(ligneCible, COL_DEBUT + NB_COL_TABLE - 1))
        .NumberFormat = "@"        ' format TEXTE avant écriture (voir le piège déjà rencontré)
    End With
    wsParam.Cells(ligneCible, COL_DEBUT).value = ""             ' CategorieSource
    wsParam.Cells(ligneCible, COL_DEBUT + 1).value = categorie  ' Categorie
    wsParam.Cells(ligneCible, COL_DEBUT + 2).value = sousCategorie   ' SousCategorie
    wsParam.Cells(ligneCible, COL_DEBUT + 3).value = "Oui"      ' Verifie : creee volontairement

    ' La liste déroulante des catégories (phase 1) doit refléter cet ajout immédiatement.
    RafraichirListesCategories

    AjouterCategoriePersonnalisee = True

End Function

' ---------------------------------------------------------------------------------------
' ObtenirSousCategories : renvoie, triées par ordre alphabétique, les sous-catégories
' (non vides) déjà associées à une catégorie donnée dans TblCategories.
' ---------------------------------------------------------------------------------------
' Si la catégorie n'a aucune sous-catégorie connue (ou n'existe pas encore), la fonction
' renvoie un tableau NON ALLOUÉ (comportement normal d'une fonction de type tableau
' qui ne lui affecte aucune valeur). L'appelant doit le vérifier avant d'utiliser le
' résultat, par exemple avec un test "On Error Resume Next : x = UBound(...)".
Public Function ObtenirSousCategories(ByVal categorie As String) As String()

    Dim wsParam As Worksheet, tblCat As ListObject
    Dim cat As Variant, sous As Variant
    Dim dejaVu As Object
    Dim liste() As String
    Dim nb As Long, r As Long
    Dim texteCat As String, texteSous As String

    categorie = Trim(categorie)
    If categorie = "" Then Exit Function

    On Error Resume Next
    Set wsParam = ThisWorkbook.Worksheets(NOM_FEUILLE_PARAM)
    On Error GoTo 0
    If wsParam Is Nothing Then Exit Function

    Set tblCat = TrouverTable(wsParam, NOM_TABLE_CATEGORIES)
    If tblCat Is Nothing Then Exit Function
    If tblCat.DataBodyRange Is Nothing Then Exit Function

    cat = LireColonne(tblCat.ListColumns("Categorie").DataBodyRange)
    sous = LireColonne(tblCat.ListColumns("SousCategorie").DataBodyRange)

    Set dejaVu = CreateObject("Scripting.Dictionary")
    dejaVu.CompareMode = 1
    ReDim liste(1 To UBound(cat, 1))

    For r = 1 To UBound(cat, 1)
        texteCat = mod_DataStructure.CellText(cat(r, 1))
        If StrComp(texteCat, categorie, vbTextCompare) = 0 Then
            texteSous = mod_DataStructure.CellText(sous(r, 1))
            If texteSous <> "" Then
                If Not dejaVu.Exists(texteSous) Then
                    dejaVu.Add texteSous, True
                    nb = nb + 1
                    liste(nb) = texteSous
                End If
            End If
        End If
    Next r

    If nb > 0 Then
        ReDim Preserve liste(1 To nb)
        TrierTextes liste, nb
        ObtenirSousCategories = liste
    End If
    ' Si nb = 0, la fonction renvoie son tableau par défaut (non alloué) : c'est voulu.

End Function


' ---------------------------------------------------------------------------------------
' ObtenirCategories : renvoie, triées par ordre alphabétique, toutes les catégories
' distinctes déjà présentes dans TblCategories.
' ---------------------------------------------------------------------------------------
' Même principe que pour ObtenirSousCategories : renvoie un tableau NON ALLOUÉ si le
' tableau est introuvable ou vide.
Public Function ObtenirCategories() As String()

    Dim wsParam As Worksheet, tblCat As ListObject
    Dim cat As Variant
    Dim dejaVu As Object
    Dim liste() As String
    Dim nb As Long, r As Long
    Dim texteCat As String

    On Error Resume Next
    Set wsParam = ThisWorkbook.Worksheets(NOM_FEUILLE_PARAM)
    On Error GoTo 0
    If wsParam Is Nothing Then Exit Function

    Set tblCat = TrouverTable(wsParam, NOM_TABLE_CATEGORIES)
    If tblCat Is Nothing Then Exit Function
    If tblCat.DataBodyRange Is Nothing Then Exit Function

    cat = LireColonne(tblCat.ListColumns("Categorie").DataBodyRange)

    Set dejaVu = CreateObject("Scripting.Dictionary")
    dejaVu.CompareMode = 1
    ReDim liste(1 To UBound(cat, 1))

    For r = 1 To UBound(cat, 1)
        texteCat = mod_DataStructure.CellText(cat(r, 1))
        If texteCat <> "" Then
            If Not dejaVu.Exists(texteCat) Then
                dejaVu.Add texteCat, True
                nb = nb + 1
                liste(nb) = texteCat
            End If
        End If
    Next r

    If nb > 0 Then
        ReDim Preserve liste(1 To nb)
        TrierTextes liste, nb
        ObtenirCategories = liste
    End If

End Function


' =====================================================================================
' GARDE-FOU ORTHOGRAPHIQUE (niveau 2 : détection de ressemblances)
' =====================================================================================
' Le niveau 1 (accents, espaces et majuscules ignorés) est déjà intégré dans
' AjouterCategoriePersonnalisee ci-dessus : il ne pose aucune question et reconnaît
' silencieusement deux graphies d'une même valeur.
'
' Le niveau 2, lui, POSE UNE QUESTION à l'opérateur lorsque le texte saisi ressemble
' BEAUCOUP à une valeur existante SANS lui être identique (typiquement, à cause d'une
' faute de frappe). Il est utilisé par mod_NouvelleCategorie (phase 3), juste avant
' l'enregistrement.

' ---------------------------------------------------------------------------------------
' NormaliserTexte : réduit un texte à une forme minimale pour le COMPARER (jamais pour
' l'afficher ou l'enregistrer; c'est uniquement un outil de comparaison interne).
' ---------------------------------------------------------------------------------------
' - minuscules
' - accents retirés (e2 -> e, e1 -> e, etc.)
' - espaces superflus réduits à un seul espace, début et fin supprimés
Public Function NormaliserTexte(ByVal texte As String) As String

    Dim r As String

    r = Trim(texte)

    ' Lettres accentuées courantes (minuscules et majuscules) remplacées par une lettre simple.
    r = Replace(r, ChrW(233), "e"): r = Replace(r, ChrW(201), "e")    ' é aigu
    r = Replace(r, ChrW(232), "e"): r = Replace(r, ChrW(200), "e")    ' è grave
    r = Replace(r, ChrW(234), "e"): r = Replace(r, ChrW(202), "e")    ' ê circonflexe
    r = Replace(r, ChrW(235), "e"): r = Replace(r, ChrW(203), "e")    ' ë tréma
    r = Replace(r, ChrW(224), "a"): r = Replace(r, ChrW(192), "a")    ' à grave
    r = Replace(r, ChrW(226), "a"): r = Replace(r, ChrW(194), "a")    ' â circonflexe
    r = Replace(r, ChrW(238), "i"): r = Replace(r, ChrW(206), "i")    ' î circonflexe
    r = Replace(r, ChrW(239), "i"): r = Replace(r, ChrW(207), "i")    ' ï tréma
    r = Replace(r, ChrW(244), "o"): r = Replace(r, ChrW(212), "o")    ' ô circonflexe
    r = Replace(r, ChrW(249), "u"): r = Replace(r, ChrW(217), "u")    ' ù grave
    r = Replace(r, ChrW(251), "u"): r = Replace(r, ChrW(219), "u")    ' û circonflexe
    r = Replace(r, ChrW(252), "u"): r = Replace(r, ChrW(220), "u")    ' ü tréma
    r = Replace(r, ChrW(231), "c"): r = Replace(r, ChrW(199), "c")    ' ç cédille

    r = LCase(r)

    ' Espaces multiples -> un seul espace (boucle simple : le nombre d'espaces
    ' consécutifs dans une catégorie est toujours très faible; la performance
    ' n'est pas un enjeu ici).
    Do While InStr(r, "  ") > 0
        r = Replace(r, "  ", " ")
    Loop

    NormaliserTexte = Trim(r)

End Function

' ---------------------------------------------------------------------------------------
' DistanceLevenshtein : nombre minimal de lettres à ajouter, retirer ou remplacer pour
' transformer "a" en "b". Méthode standard de détection d'une faute de frappe : deux mots
' identiques ont une distance de 0; une seule lettre différente donne une distance de 1.
' ---------------------------------------------------------------------------------------
Public Function DistanceLevenshtein(ByVal a As String, ByVal b As String) As Long

    Dim la As Long, lb As Long
    Dim i As Long, j As Long
    Dim cout As Long
    Dim d() As Long

    la = Len(a)
    lb = Len(b)
    ReDim d(0 To la, 0 To lb)

    For i = 0 To la
        d(i, 0) = i        ' transformer "a" en "" coûte Len(a) suppressions
    Next i
    For j = 0 To lb
        d(0, j) = j        ' transformer "" en "b" coûte Len(b) ajouts
    Next j

    For i = 1 To la
        For j = 1 To lb
            If Mid$(a, i, 1) = Mid$(b, j, 1) Then
                cout = 0
            Else
                cout = 1
            End If
            ' Le meilleur des 3 chemins : suppression, ajout, ou remplacement.
            d(i, j) = PlusPetit(d(i - 1, j) + 1, d(i, j - 1) + 1, d(i - 1, j - 1) + cout)
        Next j
    Next i

    DistanceLevenshtein = d(la, lb)

End Function

' ---------------------------------------------------------------------------------------
' TrouverCorrespondanceProche : cherche, dans une liste de textes existants, celui qui
' ressemble le plus à "texte" (distance <= seuil), en ignorant les correspondances DÉJÀ
' IDENTIQUES après normalisation (le niveau 1 s'en charge ailleurs; ce n'est pas son rôle).
' ---------------------------------------------------------------------------------------
' texte      : valeur que l'opérateur vient de saisir.
' liste      : valeurs existantes à comparer (peut être un tableau NON ALLOUÉ; dans
'              ce cas, la fonction renvoie simplement False, sans erreur).
' seuil      : distance maximale considérée comme une simple faute de frappe (2 conseillé).
' longueurMin: en dessous de cette longueur (texte normalisé), aucune correction n'est
'              proposée; sur un texte très court, une faible distance n'est pas significative.
' trouve     : en sortie, texte existant trouvé, EXACTEMENT comme écrit dans le tableau
'              (jamais normalisé), prêt à être affiché ou réutilisé tel quel.
' Renvoie True si une correspondance proche a été trouvée.
Public Function TrouverCorrespondanceProche(ByVal texte As String, ByRef liste() As String, _
                                            ByVal seuil As Long, ByVal longueurMin As Long, _
                                            ByRef trouve As String) As Boolean

    Dim texteNorm As String
    Dim candidatNorm As String
    Dim borneInf As Long, borneSup As Long
    Dim i As Long
    Dim d As Long
    Dim meilleureDistance As Long

    trouve = ""

    texteNorm = NormaliserTexte(texte)
    If Len(texteNorm) < longueurMin Then Exit Function

    ' La liste peut être un tableau non alloué (aucune valeur existante) : on le détecte
    ' sans erreur, avec le même idiome que EstTableauAlloue ailleurs dans le projet.
    On Error Resume Next
    borneInf = LBound(liste)
    borneSup = UBound(liste)
    If Err.Number <> 0 Then Exit Function
    On Error GoTo 0

    meilleureDistance = seuil + 1      ' rien trouvé pour l'instant (valeur hors seuil)

    For i = borneInf To borneSup
        candidatNorm = NormaliserTexte(liste(i))
        If candidatNorm <> texteNorm Then      ' identique après normalisation = niveau 1, pas ici
            d = DistanceLevenshtein(texteNorm, candidatNorm)
            If d <= seuil And d < meilleureDistance Then
                meilleureDistance = d
                trouve = liste(i)               ' graphie EXACTE du candidat, pas la forme normalisée
            End If
        End If
    Next i

    TrouverCorrespondanceProche = (trouve <> "")

End Function

' Renvoie le plus petit de trois nombres (utilisé par DistanceLevenshtein).
Private Function PlusPetit(ByVal a As Long, ByVal b As Long, ByVal c As Long) As Long
    Dim m As Long
    m = a
    If b < m Then m = b
    If c < m Then m = c
    PlusPetit = m
End Function


' ---------------------------------------------------------------------------------------
' ObtenirToutesSousCategories : renvoie, triées, TOUTES les sous-catégories distinctes
' du tableau, toutes catégories confondues. Utilisée par la grille de ventilation
' (phase 4) comme aide à la saisie (liste générale, non filtrée ligne par ligne) :
' contrairement à ObtenirSousCategories, elle ne dépend pas d'une catégorie précise.
' ---------------------------------------------------------------------------------------
Public Function ObtenirToutesSousCategories() As String()

    Dim wsParam As Worksheet, tblCat As ListObject
    Dim sous As Variant
    Dim dejaVu As Object
    Dim liste() As String
    Dim nb As Long, r As Long
    Dim texteSous As String

    On Error Resume Next
    Set wsParam = ThisWorkbook.Worksheets(NOM_FEUILLE_PARAM)
    On Error GoTo 0
    If wsParam Is Nothing Then Exit Function

    Set tblCat = TrouverTable(wsParam, NOM_TABLE_CATEGORIES)
    If tblCat Is Nothing Then Exit Function
    If tblCat.DataBodyRange Is Nothing Then Exit Function

    sous = LireColonne(tblCat.ListColumns("SousCategorie").DataBodyRange)

    Set dejaVu = CreateObject("Scripting.Dictionary")
    dejaVu.CompareMode = 1
    ReDim liste(1 To UBound(sous, 1))

    For r = 1 To UBound(sous, 1)
        texteSous = mod_DataStructure.CellText(sous(r, 1))
        If texteSous <> "" Then
            If Not dejaVu.Exists(texteSous) Then
                dejaVu.Add texteSous, True
                nb = nb + 1
                liste(nb) = texteSous
            End If
        End If
    Next r

    If nb > 0 Then
        ReDim Preserve liste(1 To nb)
        TrierTextes liste, nb
        ObtenirToutesSousCategories = liste
    End If

End Function

' ---------------------------------------------------------------------------------------
' FormeCanonique : recherche "texte" dans "liste" sans tenir compte des accents, espaces
' ou majuscules, et renvoie la graphie exacte trouvée dans la liste (ex. : l'opérateur
' saisit "sante" et la liste contient "Sante"; la fonction renvoie "Sante"). Si aucune
' valeur ne correspond, elle renvoie "texte" tel quel.
' ---------------------------------------------------------------------------------------
Public Function FormeCanonique(ByVal texte As String, ByRef liste() As String) As String

    Dim borneInf As Long, borneSup As Long
    Dim i As Long
    Dim texteNorm As String

    FormeCanonique = texte
    If Trim(texte) = "" Then Exit Function

    On Error Resume Next
    borneInf = LBound(liste)
    borneSup = UBound(liste)
    If Err.Number <> 0 Then Exit Function      ' liste non allouée : rien à chercher
    On Error GoTo 0

    texteNorm = NormaliserTexte(texte)
    For i = borneInf To borneSup
        If NormaliserTexte(liste(i)) = texteNorm Then
            FormeCanonique = liste(i)
            Exit Function
        End If
    Next i

End Function


' =====================================================================================
' OUTILS INTERNES (privés : invisibles dans la liste des macros)
' =====================================================================================

' Retrouve un tableau Excel par son nom dans une feuille. Renvoie Nothing s'il n'existe pas.
Private Function TrouverTable(ByVal ws As Worksheet, ByVal nomTable As String) As ListObject
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If lo.Name = nomTable Then
            Set TrouverTable = lo
            Exit Function
        End If
    Next lo
End Function

' Lit une plage en mémoire et renvoie TOUJOURS un tableau à deux dimensions.
' (Piège VBA : quand la plage ne contient qu'UNE cellule, .Value2 renvoie une simple
' valeur et non un tableau; cette fonction uniformise le résultat.)
Private Function LireColonne(ByVal plage As Range) As Variant
    Dim t() As Variant
    If plage.Cells.count = 1 Then
        ReDim t(1 To 1, 1 To 1)
        t(1, 1) = plage.Value2
        LireColonne = t
    Else
        LireColonne = plage.Value2
    End If
End Function

' Tri alphabétique (insensible à la casse) des n premiers éléments d'un tableau de textes.
' Le tri par insertion est simple à lire et suffisamment rapide pour quelques centaines
' d'éléments.
Private Sub TrierTextes(ByRef t() As String, ByVal n As Long)
    Dim i As Long, j As Long
    Dim cle As String
    For i = 2 To n
        cle = t(i)
        j = i - 1
        Do While j >= 1
            If StrComp(t(j), cle, vbTextCompare) <= 0 Then Exit Do
            t(j + 1) = t(j)
            j = j - 1
        Loop
        t(j + 1) = cle
    Next i
End Sub

' Convertit un numéro de colonne en lettre (1 -> A, 16 -> P). Valable jusqu'à la colonne Z.
Private Function ColonneEnLettre(ByVal numero As Long) As String
    ColonneEnLettre = Chr$(64 + numero)
End Function

' Catégorie SOURCE "Frais, remb santé" (accent construit avec ChrW pour rester compatible à l'import).
Private Function CategorieSourceSante() As String
    CategorieSourceSante = "Frais, remb sant" & ChrW(233)
End Function

' Nouvelle catégorie parente : "Santé, prévoyance" (telle qu'écrite dans vos données).
Private Function CategorieSanteNouvelle() As String
    CategorieSanteNouvelle = "Sant" & ChrW(233) & ", pr" & ChrW(233) & "voyance"
End Function

