Attribute VB_Name = "mod_Categories"
Option Explicit

' =====================================================================================
' MODULE : mod_Categories
'
' PHASE 1 du chantier "Categorie / Sous-categorie / Ventilation".
'
' ROLE DE CE MODULE (a lire en premier, meme si vous debutez) :
'
'   Aujourd'hui, une operation n'a qu'UNE categorie, sous forme de texte simple
'   (ex : "Frais, remb sante"). Le nouveau fonctionnement demande DEUX niveaux :
'        Categorie      (ex : "Sante, prevoyance")
'        Sous-categorie (ex : "Frais, remb sante")
'
'   Les fichiers sources de la banque (OFX + CSV) ne fournissent toujours qu'UNE
'   categorie : c'est la "CATEGORIE SOURCE". Il faut donc un TABLEAU DE
'   CORRESPONDANCE qui dit, pour chaque categorie source :
'        "quand la banque m'envoie X, je la range dans Categorie Y / Sous-categorie Z".
'
'   Ce module met en place les deux briques de base, RIEN d'autre :
'     1. La colonne "SousCategorie" ajoutee a la fin de TblOperations.
'     2. Le tableau "TblCategories" dans la feuille Param (colonnes K a N), pre-rempli
'        avec toutes les categories deja presentes dans vos operations.
'
'   CE MODULE NE MODIFIE AUCUNE DONNEE EXISTANTE de TblOperations : les valeurs de
'   la colonne Categorie ne sont pas touchees. La conversion des anciennes operations
'   viendra dans une phase ulterieure, apres votre validation du tableau.
'
'   Il ne modifie PAS non plus les autres modules (mod_ImportOFX, mod_Display, ...).
'
' COMMENT UTILISER (dans l'ordre) :
'   1. Enregistrer une COPIE du classeur (precaution, une colonne va etre ajoutee).
'   2. Ctrl+G (fenetre Execution), taper :  PreparerPhase1Categories   puis Entree.
'   3. Sur la feuille Param qui s'affiche, verifier / corriger le tableau K:N,
'      puis mettre "Oui" dans la colonne "Verifie" pour chaque ligne controlee.
'   4. Quand vous avez fini : MasquerParam (Ctrl+G) pour cacher la feuille.
'
' A PROPOS DES ACCENTS : ce fichier est 100% ASCII. Les lettres accentuees sont
' fabriquees a l'execution avec ChrW() ou avec la fonction AccentsFR (voir plus bas).
' =====================================================================================

' --- Nom de la feuille et du tableau de correspondance ---------------------------------
Public Const NOM_FEUILLE_PARAM As String = "Param"
Public Const NOM_TABLE_CATEGORIES As String = "TblCategories"

' --- Nom de la colonne ajoutee dans TblOperations -----------------------------------
Public Const NOM_COL_SOUS_CATEGORIE As String = "SousCategorie"

' --- Position du tableau dans la feuille Param ----------------------------------------
' Les colonnes A a I de Param sont deja utilisees (listes Mois, Annees, Medecins...).
' On place donc le nouveau tableau plus a droite, en laissant la colonne J vide.
Private Const COL_DEBUT As Long = 11        ' colonne K : premiere colonne du tableau
Private Const NB_COL_TABLE As Long = 4      ' K, L, M, N = 4 colonnes
Private Const COL_LISTE As Long = 16        ' colonne P : liste technique des categories

' Nom (au niveau du classeur) de la liste servant aux menus deroulants
Private Const NOM_LISTE_CATEGORIES As String = "ListeCategories"

' Les 4 colonnes du tableau, dans l'ordre :
'   K = CategorieSource : le texte EXACT envoye par la banque (ne pas modifier)
'   L = Categorie       : la categorie de rangement dans le nouvel outil
'   M = SousCategorie   : la sous-categorie (peut rester vide)
'   N = Verifie         : "Oui" quand l'operateur a controle la ligne


' =====================================================================================
' PROCEDURE D'ENSEMBLE : enchaine les 3 etapes de la Phase 1.
' (Chaque etape reste utilisable seule, voir plus bas.)
' =====================================================================================
Public Sub PreparerPhase1Categories()

    ' Etape 1 : colonne SousCategorie dans TblOperations
    AjouterColonneSousCategorie

    ' Etape 2 : creation / mise a jour du tableau de correspondance
    InitialiserTableCategories

    ' Etape 3 : on montre la feuille Param pour que l'operateur puisse verifier
    AfficherParamCategories

End Sub


' =====================================================================================
' ETAPE 1 : ajoute la colonne "SousCategorie" a la fin de TblOperations
' =====================================================================================
' Pourquoi "a la fin" ? Parce que tout votre code lit les colonnes par NOM
' (voir RecupIndexCol) ou par position pour les colonnes 1 a 20. En ajoutant la
' nouvelle colonne apres la derniere (colonne U), aucune position existante ne bouge.
' Cette procedure peut etre relancee sans danger : si la colonne existe deja,
' elle ne fait rien.
Public Sub AjouterColonneSousCategorie()

    Dim tblOps As ListObject
    Dim colonne As ListColumn

    ' On reutilise la fonction existante qui retrouve TblOperations
    ' (elle affiche elle-meme un message si le tableau est introuvable).
    Set tblOps = mod_DonneesTable.GetOperationsTable()
    If tblOps Is Nothing Then Exit Sub

    ' On cherche si la colonne existe deja. "On Error Resume Next" evite un plantage :
    ' demander une colonne inexistante provoque une erreur, que l'on sait interpreter
    ' (colonne restee a Nothing = elle n'existe pas).
    On Error Resume Next
    Set colonne = tblOps.ListColumns(NOM_COL_SOUS_CATEGORIE)
    On Error GoTo 0

    If Not colonne Is Nothing Then
        MsgBox "La colonne " & NOM_COL_SOUS_CATEGORIE & " existe deja dans " & _
               tblOps.Name & ". Rien n'a ete modifie.", vbInformation, "Phase 1"
        Exit Sub
    End If

    ' ListColumns.Add sans argument = ajout apres la derniere colonne du tableau.
    Set colonne = tblOps.ListColumns.Add
    colonne.Name = NOM_COL_SOUS_CATEGORIE

    MsgBox "Colonne " & NOM_COL_SOUS_CATEGORIE & " ajoutee a la fin de " & tblOps.Name & ".", _
           vbInformation, "Phase 1"

End Sub


' =====================================================================================
' ETAPE 2 : cree ou complete le tableau de correspondance TblCategories (feuille Param)
' =====================================================================================
' - 1ere execution : le tableau est cree, avec UNE ligne par categorie source
'   trouvee dans TblOperations (toutes les valeurs distinctes, triees A a Z).
' - Executions suivantes : seules les categories source NOUVELLES sont ajoutees en
'   bas. Ce que l'operateur a deja saisi n'est JAMAIS ecrase.
'
' Valeurs proposees au depart (l'operateur peut tout modifier) :
'   - Categorie = le texte source recopie tel quel, SousCategorie vide, Verifie vide.
'   - Cas particulier que vous avez decrit : "Frais, remb sante" est une sous-categorie
'     de "Sante, prevoyance" -> pre-rempli ainsi, et Verifie = "Oui".
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
    Dim sortie() As Variant
    Dim premiereLigne As Long, derniereLigne As Long
    Dim nbAVerifier As Long, verifies As Variant

    ' --- Recuperer TblOperations et la feuille Param -------------------------------
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

    ' --- Lire toutes les categories source en memoire (un seul aller-retour = rapide) ---
    donnees = LireColonne(colonneCat.DataBodyRange)

    ' --- "dejaVu" : dictionnaire qui memorise les categories deja traitees ------------
    ' CompareMode = 1 (vbTextCompare) : on ignore majuscules/minuscules.
    ' Il doit etre regle AVANT d'ajouter le moindre element.
    Set dejaVu = CreateObject("Scripting.Dictionary")
    dejaVu.CompareMode = 1

    ' Si le tableau existe deja, on y lit les sources connues pour ne pas les redoubler.
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

    ' --- Reperer les categories source qui ne sont pas encore dans le tableau -----------
    ReDim nouvelles(1 To UBound(donnees, 1))
    nbNouvelles = 0
    For i = 1 To UBound(donnees, 1)
        ' CellText : lecture "securisee" (jamais d'erreur, espaces inutiles retires)
        texte = mod_DataStructure.CellText(donnees(i, 1))
        If texte <> "" Then
            If Not dejaVu.Exists(texte) Then
                dejaVu.Add texte, True
                nbNouvelles = nbNouvelles + 1
                nouvelles(nbNouvelles) = texte
            End If
        End If
    Next i

    ' --- Rien de nouveau : on s'assure seulement que les listes sont a jour -------------
    If nbNouvelles = 0 Then
        If Not tblCat Is Nothing Then
            RafraichirListesCategories
            PoserListesDeroulantes tblCat
        End If
        MsgBox "Aucune nouvelle categorie source a ajouter.", vbInformation, "Phase 1"
        Exit Sub
    End If

    ' --- Trier les nouvelles categories de A a Z ------------------------------------------
    TrierTextes nouvelles, nbNouvelles

    ' --- Preparer le bloc a ecrire (tableau en memoire : 4 colonnes) ----------------------
    ReDim sortie(1 To nbNouvelles, 1 To NB_COL_TABLE)
    For i = 1 To nbNouvelles
        sortie(i, 1) = nouvelles(i)      ' CategorieSource
        sortie(i, 2) = nouvelles(i)      ' Categorie : proposition = meme texte
        sortie(i, 3) = ""                ' SousCategorie : vide
        sortie(i, 4) = ""                ' Verifie : a faire par l'operateur

        ' Cas particulier decrit par l'operateur : "Frais, remb sante" est une
        ' sous-categorie de "Sante, prevoyance".
        If StrComp(nouvelles(i), CategorieSourceSante(), vbTextCompare) = 0 Then
            sortie(i, 2) = CategorieSanteNouvelle()
            sortie(i, 3) = CategorieSourceSante()
            sortie(i, 4) = "Oui"
        End If
    Next i

    ' --- Ecriture dans la feuille Param -----------------------------------------------------
    If tblCat Is Nothing Then
        ' PREMIERE FOIS : en-tetes + donnees + creation du tableau Excel
        wsParam.Range(wsParam.Cells(1, COL_DEBUT), wsParam.Cells(1, COL_DEBUT + NB_COL_TABLE - 1)).Value = _
            Array("CategorieSource", "Categorie", "SousCategorie", "Verifie")
        premiereLigne = 2
    Else
        ' Tableau deja present : on ecrit sous la derniere ligne remplie. On se base sur la
        ' colonne L (Categorie) car elle est toujours remplie, contrairement a la colonne K
        ' (une categorie creee a la main n'a pas de categorie source).
        derniereLigne = wsParam.Cells(wsParam.Rows.Count, COL_DEBUT + 1).End(xlUp).Row
        premiereLigne = derniereLigne + 1
    End If

    ' Format TEXTE avant d'ecrire : evite qu'Excel transforme une valeur en date ou en nombre
    ' (piege deja rencontre dans ce projet).
    With wsParam.Range(wsParam.Cells(premiereLigne, COL_DEBUT), _
                       wsParam.Cells(premiereLigne + nbNouvelles - 1, COL_DEBUT + NB_COL_TABLE - 1))
        .NumberFormat = "@"
        .Value2 = sortie
    End With

    If tblCat Is Nothing Then
        Set tblCat = wsParam.ListObjects.Add(xlSrcRange, _
            wsParam.Range(wsParam.Cells(1, COL_DEBUT), _
                          wsParam.Cells(premiereLigne + nbNouvelles - 1, COL_DEBUT + NB_COL_TABLE - 1)), , xlYes)
        tblCat.Name = NOM_TABLE_CATEGORIES
    Else
        ' On agrandit le tableau Excel pour qu'il englobe les lignes ajoutees.
        tblCat.Resize wsParam.Range(wsParam.Cells(1, COL_DEBUT), _
                          wsParam.Cells(premiereLigne + nbNouvelles - 1, COL_DEBUT + NB_COL_TABLE - 1))
    End If

    ' --- Listes deroulantes ---------------------------------------------------------------------
    RafraichirListesCategories
    PoserListesDeroulantes tblCat

    ' --- Petit bilan : combien de lignes restent a controler ? ------------------------------------
    verifies = LireColonne(tblCat.ListColumns("Verifie").DataBodyRange)
    For i = 1 To UBound(verifies, 1)
        If StrComp(mod_DataStructure.CellText(verifies(i, 1)), "Oui", vbTextCompare) <> 0 Then
            nbAVerifier = nbAVerifier + 1
        End If
    Next i

    MsgBox nbNouvelles & " categorie(s) source ajoutee(s) au tableau " & NOM_TABLE_CATEGORIES & "." & _
           vbCrLf & vbCrLf & _
           "Lignes restant a verifier (colonne Verifie <> Oui) : " & nbAVerifier & vbCrLf & vbCrLf & _
           AccentsFR("Pour chaque ligne : contr{o1}lez Cat{e2}gorie et Sous-cat{e2}gorie, puis tapez Oui dans Verifie."), _
           vbInformation, "Phase 1"

End Sub


' =====================================================================================
' Met a jour la liste TECHNIQUE des categories distinctes (colonne P de Param).
' =====================================================================================
' Cette liste alimente le menu deroulant de la colonne "Categorie". Elle est
' RECALCULEE a partir du tableau, donc toute categorie ajoutee au tableau apparait
' dans le menu apres un appel a cette procedure. Elle sera aussi reutilisee par les
' futurs formulaires (Phases 2 et 3).
'
' Pourquoi ne pas mettre les categories dans le menu sous forme de texte separe par
' des virgules ? Parce que vos categories contiennent elles-memes des virgules
' ("Alimentation, supermarche") : Excel les couperait en morceaux. On passe donc
' par des CELLULES, et le menu pointe vers cette plage.
Public Sub RafraichirListesCategories()

    Dim wsParam As Worksheet
    Dim tblCat As ListObject
    Dim donnees As Variant
    Dim dejaVu As Object
    Dim liste() As String
    Dim sortie() As Variant
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

    ' On efface l'ancienne liste (sous l'en-tete), puis on ecrit la nouvelle.
    wsParam.Range(wsParam.Cells(2, COL_LISTE), wsParam.Cells(wsParam.Rows.Count, COL_LISTE)).ClearContents
    wsParam.Cells(1, COL_LISTE).Value = "Liste categories (zone technique - ne pas modifier)"

    ReDim sortie(1 To nb, 1 To 1)
    For i = 1 To nb
        sortie(i, 1) = liste(i)
    Next i
    wsParam.Cells(2, COL_LISTE).Resize(nb, 1).NumberFormat = "@"
    wsParam.Cells(2, COL_LISTE).Resize(nb, 1).Value2 = sortie

    ' Nom "ListeCategories" : plage DYNAMIQUE (OFFSET) qui s'adapte au nombre de
    ' categories ecrites. Cree au niveau du CLASSEUR (ThisWorkbook.Names), comme les
    ' autres noms du projet, pour eviter la confusion de portee deja rencontree.
    ThisWorkbook.Names.Add Name:=NOM_LISTE_CATEGORIES, _
        RefersTo:="=OFFSET(" & NOM_FEUILLE_PARAM & "!$" & ColonneEnLettre(COL_LISTE) & "$2,0,0," & _
                  "MAX(1,COUNTA(" & NOM_FEUILLE_PARAM & "!$" & ColonneEnLettre(COL_LISTE) & "$2:$" & _
                  ColonneEnLettre(COL_LISTE) & "$500)),1)"

End Sub


' =====================================================================================
' Pose les menus deroulants du tableau TblCategories
' =====================================================================================
'   - colonne Categorie : liste des categories existantes, MAIS l'operateur peut
'     taper une categorie nouvelle (alerte "avertissement", pas de blocage).
'   - colonne Verifie   : uniquement "Oui" ou vide.
' La colonne SousCategorie reste en saisie libre pour l'instant : sa liste
' dependante de la categorie choisie sera geree dans les formulaires (Phase 2).
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
' Affiche la feuille Param (normalement masquee) et se place sur le tableau
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
' Remasque la feuille Param (a lancer quand le controle du tableau est termine)
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
' PHASE 3 : creation d'une categorie / sous-categorie par l'operateur
' =====================================================================================
' Ces deux fonctions sont utilisees par le formulaire de creation (mod_NouvelleCategorie).
' Elles vivent ici, et non dans ce nouveau module, car TblCategories est "la propriete"
' de mod_Categories : c'est le seul endroit du projet qui sait comment ce tableau est
' construit (nom de la feuille, position des colonnes...).

' ---------------------------------------------------------------------------------------
' AjouterCategoriePersonnalisee : enregistre une paire (Categorie, SousCategorie) choisie
' ou creee par l'operateur dans le tableau TblCategories, si elle n'y est pas deja.
' ---------------------------------------------------------------------------------------
' ByRef categorie / sousCategorie : en ENTREE, ce que l'operateur a tape ou choisi.
' En SORTIE (si la fonction renvoie True), ces deux variables sont remplacees par leur
' forme "canonique" : si "sante" existait deja sous la forme "Sante" dans le tableau,
' categorie devient "Sante", pour eviter que deux orthographes de la meme categorie
' cohabitent (ex: "Sante" et "sante" traites comme deux categories differentes ailleurs
' dans le classeur).
'
' Renvoie False uniquement si la categorie est vide, ou si TblCategories est introuvable
' (le formulaire appelant doit alors afficher un message et ne pas continuer).
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
    ' NIVEAU 1 du garde-fou orthographique (normalisation silencieuse, aucune question
    ' posee) : on ignore les majuscules/minuscules, les accents et les espaces en trop.
    ' "Sante" et "Sant{e2}" sont ainsi reconnus comme LA MEME categorie, sans interruption.
    For r = 1 To UBound(cat, 1)
        texteCat = mod_DataStructure.CellText(cat(r, 1))
        texteSous = mod_DataStructure.CellText(sous(r, 1))

        If NormaliserTexte(texteCat) = NormaliserTexte(categorie) Then
            casCatTrouvee = True
            casCatCanonique = texteCat            ' on garde la 1ere ecriture rencontree

            If NormaliserTexte(texteSous) = NormaliserTexte(sousCategorie) Then
                ' La paire existe deja (memes lettres, accents/espaces/casse mis a part) :
                ' rien a ajouter. On renvoie les libelles EXACTS deja presents dans le tableau.
                categorie = texteCat
                sousCategorie = texteSous
                AjouterCategoriePersonnalisee = True
                Exit Function
            End If
        End If
    Next r

    ' La categorie existe deja (sous une autre paire) : on reprend son ecriture exacte,
    ' pour eviter que "Sante" et "sante" deviennent deux categories differentes.
    If casCatTrouvee Then categorie = casCatCanonique

    ' --- Aucune paire correspondante : ajout d'une nouvelle ligne au tableau -------------
    ' CategorieSource reste VIDE : cette ligne ne provient pas d'un import bancaire,
    ' mais d'une creation manuelle de l'operateur.
    ligneCible = tblCat.ListRows.Add.Range.Row

    With wsParam.Range(wsParam.Cells(ligneCible, COL_DEBUT), wsParam.Cells(ligneCible, COL_DEBUT + NB_COL_TABLE - 1))
        .NumberFormat = "@"        ' format TEXTE avant ecriture (voir piege deja rencontre)
    End With
    wsParam.Cells(ligneCible, COL_DEBUT).Value = ""             ' CategorieSource
    wsParam.Cells(ligneCible, COL_DEBUT + 1).Value = categorie  ' Categorie
    wsParam.Cells(ligneCible, COL_DEBUT + 2).Value = sousCategorie   ' SousCategorie
    wsParam.Cells(ligneCible, COL_DEBUT + 3).Value = "Oui"      ' Verifie : creee volontairement

    ' La liste deroulante des categories (Phase 1) doit refleter cet ajout immediatement.
    RafraichirListesCategories

    AjouterCategoriePersonnalisee = True

End Function

' ---------------------------------------------------------------------------------------
' ObtenirSousCategories : renvoie, triees par ordre alphabetique, les sous-categories
' (non vides) deja associees a une categorie donnee dans TblCategories.
' ---------------------------------------------------------------------------------------
' Si la categorie n'a aucune sous-categorie connue (ou n'existe pas encore), la fonction
' renvoie un tableau NON ALLOUE (comportement normal d'une fonction de type tableau qui
' ne lui affecte jamais de valeur). L'appelant doit verifier cela avant d'utiliser le
' resultat, par exemple avec un test "On Error Resume Next : x = UBound(...)".
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
    ' Si nb = 0, la fonction renvoie son tableau par defaut (non alloue) : c'est voulu.

End Function


' ---------------------------------------------------------------------------------------
' ObtenirCategories : renvoie, triees par ordre alphabetique, toutes les categories
' distinctes deja presentes dans TblCategories.
' ---------------------------------------------------------------------------------------
' Meme principe que ObtenirSousCategories : renvoie un tableau NON ALLOUE si le tableau
' est introuvable ou vide.
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
' GARDE-FOU ORTHOGRAPHIQUE (Niveau 2 : detection de ressemblance)
' =====================================================================================
' Le Niveau 1 (accents, espaces, majuscules ignores) est deja integre dans
' AjouterCategoriePersonnalisee ci-dessus : il ne pose jamais de question, il "reconnait"
' silencieusement deux ecritures de la meme chose.
'
' Le Niveau 2, lui, POSE UNE QUESTION a l'operateur quand le texte tape ressemble
' BEAUCOUP a une valeur existante SANS lui etre identique (typiquement : une faute de
' frappe). Il est utilise par mod_NouvelleCategorie (Phase 3), juste avant l'enregistrement.

' ---------------------------------------------------------------------------------------
' NormaliserTexte : reduit un texte a une forme minimale pour le COMPARER (jamais pour
' l'afficher ou l'enregistrer : c'est uniquement un outil de comparaison interne).
' ---------------------------------------------------------------------------------------
' - minuscules
' - accents retires (e2 -> e, e1 -> e, etc.)
' - espaces en trop reduits a un seul espace, debut/fin coupes
Public Function NormaliserTexte(ByVal texte As String) As String

    Dim r As String

    r = Trim(texte)

    ' Lettres accentuees courantes (minuscules et majuscules) -> lettre simple.
    r = Replace(r, ChrW(233), "e") : r = Replace(r, ChrW(201), "e")   ' e aigu
    r = Replace(r, ChrW(232), "e") : r = Replace(r, ChrW(200), "e")   ' e grave
    r = Replace(r, ChrW(234), "e") : r = Replace(r, ChrW(202), "e")   ' e circonflexe
    r = Replace(r, ChrW(235), "e") : r = Replace(r, ChrW(203), "e")   ' e trema
    r = Replace(r, ChrW(224), "a") : r = Replace(r, ChrW(192), "a")   ' a grave
    r = Replace(r, ChrW(226), "a") : r = Replace(r, ChrW(194), "a")   ' a circonflexe
    r = Replace(r, ChrW(238), "i") : r = Replace(r, ChrW(206), "i")   ' i circonflexe
    r = Replace(r, ChrW(239), "i") : r = Replace(r, ChrW(207), "i")   ' i trema
    r = Replace(r, ChrW(244), "o") : r = Replace(r, ChrW(212), "o")   ' o circonflexe
    r = Replace(r, ChrW(249), "u") : r = Replace(r, ChrW(217), "u")   ' u grave
    r = Replace(r, ChrW(251), "u") : r = Replace(r, ChrW(219), "u")   ' u circonflexe
    r = Replace(r, ChrW(252), "u") : r = Replace(r, ChrW(220), "u")   ' u trema
    r = Replace(r, ChrW(231), "c") : r = Replace(r, ChrW(199), "c")   ' c cedille

    r = LCase(r)

    ' Espaces multiples -> un seul espace (boucle simple : le nombre d'espaces
    ' consecutifs dans une categorie est toujours tres petit, la performance n'est pas
    ' un enjeu ici).
    Do While InStr(r, "  ") > 0
        r = Replace(r, "  ", " ")
    Loop

    NormaliserTexte = Trim(r)

End Function

' ---------------------------------------------------------------------------------------
' DistanceLevenshtein : nombre minimal de lettres a ajouter, retirer ou remplacer pour
' transformer "a" en "b". Methode standard pour detecter une faute de frappe : deux mots
' identiques ont une distance de 0, une seule lettre differente donne 1, etc.
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
        d(i, 0) = i        ' transformer "a" en "" coute la Len(a) suppressions
    Next i
    For j = 0 To lb
        d(0, j) = j        ' transformer "" en "b" coute la Len(b) ajouts
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
' ressemble le plus a "texte" (distance <= seuil), en ignorant les correspondances DEJA
' IDENTIQUES apres normalisation (Niveau 1 s'en charge ailleurs, ce n'est pas son role).
' ---------------------------------------------------------------------------------------
' texte      : ce que l'operateur vient de taper.
' liste      : les valeurs existantes a comparer (peut etre un tableau NON ALLOUE : dans
'              ce cas, la fonction renvoie simplement False, sans erreur).
' seuil      : distance maximale consideree comme "une simple faute de frappe" (2 conseille).
' longueurMin: en dessous de cette longueur (texte normalise), on ne propose jamais de
'              correction : sur un texte tres court, une petite distance ne veut rien dire.
' trouve     : (en sortie) le texte existant trouve, EXACTEMENT comme ecrit dans le
'              tableau (jamais normalise), pret a etre affiche ou reutilise tel quel.
' Renvoie True si une correspondance proche a ete trouvee.
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

    ' La liste peut etre un tableau jamais alloue (aucune valeur existante) : on le
    ' detecte sans planter, via le meme idiome que EstTableauAlloue ailleurs dans le projet.
    On Error Resume Next
    borneInf = LBound(liste)
    borneSup = UBound(liste)
    If Err.Number <> 0 Then Exit Function
    On Error GoTo 0

    meilleureDistance = seuil + 1      ' rien trouve pour l'instant (valeur "hors seuil")

    For i = borneInf To borneSup
        candidatNorm = NormaliserTexte(liste(i))
        If candidatNorm <> texteNorm Then      ' identique apres normalisation = Niveau 1, pas ici
            d = DistanceLevenshtein(texteNorm, candidatNorm)
            If d <= seuil And d < meilleureDistance Then
                meilleureDistance = d
                trouve = liste(i)               ' forme EXACTE du candidat, pas la normalisee
            End If
        End If
    Next i

    TrouverCorrespondanceProche = (trouve <> "")

End Function

' Renvoie le plus petit de 3 nombres (utilise par DistanceLevenshtein).
Private Function PlusPetit(ByVal a As Long, ByVal b As Long, ByVal c As Long) As Long
    Dim m As Long
    m = a
    If b < m Then m = b
    If c < m Then m = c
    PlusPetit = m
End Function


' ---------------------------------------------------------------------------------------
' ObtenirToutesSousCategories : renvoie, triees, TOUTES les sous-categories distinctes
' du tableau, toutes categories confondues. Utilisee par la grille de ventilation
' (Phase 4) comme simple aide a la saisie (liste "large", pas filtree ligne par ligne) :
' contrairement a ObtenirSousCategories, elle ne depend pas d'une categorie precise.
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
' FormeCanonique : recherche "texte" dans "liste" en ignorant accents/espaces/majuscules,
' et renvoie l'ECRITURE EXACTE trouvee dans la liste (ex : l'operateur tape "sante", la
' liste contient "Sante" -> renvoie "Sante"). Si rien n'est trouve, renvoie "texte" tel quel.
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
    If Err.Number <> 0 Then Exit Function      ' liste non allouee : rien a chercher
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
' OUTILS INTERNES (prives : invisibles dans la liste des macros)
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

' Lit une plage en memoire et renvoie TOUJOURS un tableau a 2 dimensions.
' (Piege VBA : quand la plage ne contient qu'UNE cellule, .Value2 renvoie une simple
'  valeur et non un tableau ; cette fonction uniformise le resultat.)
Private Function LireColonne(ByVal plage As Range) As Variant
    Dim t() As Variant
    If plage.Cells.Count = 1 Then
        ReDim t(1 To 1, 1 To 1)
        t(1, 1) = plage.Value2
        LireColonne = t
    Else
        LireColonne = plage.Value2
    End If
End Function

' Tri alphabetique (insensible a la casse) des n premiers elements d'un tableau de textes.
' "Tri par insertion" : simple a lire, largement assez rapide pour quelques centaines
' d'elements.
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

' Convertit un numero de colonne en lettre (1 -> A, 16 -> P). Valable jusqu'a la colonne Z.
Private Function ColonneEnLettre(ByVal numero As Long) As String
    ColonneEnLettre = Chr$(64 + numero)
End Function

' Categorie SOURCE "Frais, remb sante" (avec accent, fabrique par ChrW pour rester en ASCII)
Private Function CategorieSourceSante() As String
    CategorieSourceSante = "Frais, remb sant" & ChrW(233)
End Function

' Nouvelle categorie parente : "Sante, prevoyance" (telle qu'ecrite dans vos donnees)
Private Function CategorieSanteNouvelle() As String
    CategorieSanteNouvelle = "Sant" & ChrW(233) & ", pr" & ChrW(233) & "voyance"
End Function

' Traduit des codes en lettres accentuees (meme principe que la fonction FR du projet).
'   {e2} = e accent aigu   {e1} = e accent grave   {ea} = e accent circonflexe
'   {a2} = a accent grave  {o1} = o accent circonflexe
Private Function AccentsFR(ByVal texte As String) As String
    Dim r As String
    r = texte
    r = Replace(r, "{e2}", ChrW(233))
    r = Replace(r, "{e1}", ChrW(232))
    r = Replace(r, "{ea}", ChrW(234))
    r = Replace(r, "{a2}", ChrW(224))
    r = Replace(r, "{o1}", ChrW(244))
    AccentsFR = r
End Function
