' =====================================================================================
' PROPOSITION (NON APPLIQUEE) - A VALIDER PAR L'OPERATEUR AVANT TOUTE INTEGRATION
'
' Bloc de remplacement de la Sub AppliquerLignesMarquees du module
' mod_SyntheseRechercheOperations (de "Public Sub AppliquerLignesMarquees()" jusqu'a
' son "End Sub"). Aucune autre procedure n'est touchee.
' Fichier 100% ASCII : les accents des messages passent par FR() comme convenu.
' =====================================================================================

' =====================================================================================
' AppliquerLignesMarquees : applique Categorie/Notes des lignes marquees "OUI"
'
' PRINCIPE (meme structure que RechercherOperations et les autres macros) :
'   1. On REUTILISE les variables globales deja renseignees par RechercherOperations :
'        - tbl                 : le tableau TblOperations
'        - plageSortieEcriture : la zone de donnees affichee a l'ecran
'        - posValider, posCategorie, posNotes, posID : position de chaque champ dans
'                                 la zone ecran (1 = 1ere colonne de la zone)
'        - colID, colCategorie, colNotes : position de chaque champ dans TblOperations
'   2. On lit l'ecran EN BLOC dans un tableau memoire (une seule lecture de la feuille)
'   3. On parcourt ce tableau en memoire (rapide) pour reperer les lignes marquees
'   4. On n'ecrit dans TblOperations QUE les cellules concernees (jamais une colonne entiere)
'   5. On efface la marque, puis on recalcule le suivi sante
' =====================================================================================
Public Sub AppliquerLignesMarquees()

    ' --- Variables LOCALES uniquement : aucune ne porte le nom d'une variable globale
    '     de mod_VarGlobales (sinon elle la "masquerait" silencieusement) ---
    Dim plageTable As Range           ' zone des donnees de TblOperations (sans l'entete)
    Dim tabEcran As Variant           ' copie memoire de la zone affichee a l'ecran
    Dim nbLignesEcran As Long         ' nombre de lignes de cette copie memoire
    Dim indexParID As Object          ' dictionnaire : ID_Transaction -> ligne dans tblData
    Dim idTexte As String             ' un ID lu dans tblData (construction du dictionnaire)
    Dim idCible As String             ' l'ID de la ligne marquee, lu a l'ecran
    Dim i As Long                     ' ligne courante DANS la zone ecran (1 = 1ere operation)
    Dim ligneTbl As Long              ' ligne courante DANS tblData (1 = entete, 2 = 1ere operation)
    Dim ligneCible As Long            ' ligne de la marquee dans tblData (meme repere que ligneTbl)
    Dim categorieValeur As String
    Dim notesOrigine As String, seg0 As String
    Dim verrouillee As Boolean
    Dim nbAppliquees As Long

    ' 1. Garde-fous : les variables globales existent-elles encore ?
    '    (elles sont vides si aucune recherche n'a ete lancee, ou apres un "Reset" de VBA)
    If tbl Is Nothing Or plageSortieEcriture Is Nothing Then
        MsgBox FR("Aucune recherche n'est affich{e2}e. Cliquez d'abord sur 'Rechercher'."), vbExclamation
        Exit Sub
    End If
    If tbl.DataBodyRange Is Nothing Then
        MsgBox "Le tableau est vide.", vbExclamation
        Exit Sub
    End If

    ' Si une erreur survient, on passe par GestionErreur pour REACTIVER les evenements
    ' (sans cela Excel resterait "sourd" aux modifications de cellules)
    On Error GoTo GestionErreur

    Set plageTable = tbl.DataBodyRange

    ' 2. Relecture FRAICHE de la table (meme lecture que RechercherOperations : entete en ligne 1)
    '    On NE se fie PAS au contenu actuel de la globale tblData : d'autres macros
    '    (VerifierNotesSante...) la rechargent SANS la ligne d'entete, ce qui decalerait
    '    toutes les lignes d'un cran. La relire ici la rend fiable et a jour.
    tblData = tbl.Range.value
    tblDataLineTotal = UBound(tblData, 1)

    ' 3. Recuperation des index de colonne dans la base de donnees
    mod_Display.RecupIndexCol

    ' 4. Lecture de la zone ecran EN BLOC : 1 seul echange avec la feuille, ensuite tout
    '    se passe en memoire. Les lignes masquees par un filtre sont bien incluses.
    tabEcran = plageSortieEcriture.value
    nbLignesEcran = UBound(tabEcran, 1)

    ' 5. Index ID_Transaction -> ligne de tblData, construit UNE SEULE FOIS
    '    (plutot qu'un balayage complet de la table pour chaque ligne marquee).
    '    Les lignes SANS ID sont exclues : sinon toutes les lignes vides partageraient la
    '    meme cle "" et une ligne marquee sans ID serait appliquee sur la DERNIERE ligne vide.
    Set indexParID = CreateObject("Scripting.Dictionary")
    For ligneTbl = 2 To tblDataLineTotal
        idTexte = mod_DataStructure.CellText(tblData(ligneTbl, colID))
        If idTexte <> "" Then indexParID(idTexte) = ligneTbl
    Next ligneTbl

    ' 6. Application des lignes marquees
    '    Evenements coupes : nos ecritures ne doivent pas declencher Workbook_SheetChange
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    nbAppliquees = 0

    For i = 1 To nbLignesEcran
        ' Une ligne est marquee si sa colonne "Valider" contient OUI (majuscules ou non)
        If LCase$(mod_DataStructure.CellText(tabEcran(i, posValider))) = "oui" Then

            ' On retrouve la ligne d'origine grace a son ID_Transaction
            idCible = mod_DataStructure.CellText(tabEcran(i, posID))

            If indexParID.Exists(idCible) Then
                ligneCible = indexParID(idCible)

                ' Categorie : toujours appliquee
                ' (ligneCible - 1 car plageTable n'a pas de ligne d'entete, contrairement a tblData)
                categorieValeur = mod_DataStructure.CellText(tabEcran(i, posCategorie))
                plageTable.Cells(ligneCible - 1, colCategorie).value = categorieValeur

                ' Notes : verrouillage sur la valeur ORIGINALE (celle de TblOperations avant
                ' cette application), pas sur ce qui est affiche a l'ecran. C'est le meme test
                ' que celui de RechercherOperations qui grise la cellule.
                notesOrigine = mod_DataStructure.CellText(tblData(ligneCible, colNotes))
                seg0 = mod_ImportOFX.SegmentTexte(notesOrigine, ";", 0)
                verrouillee = mod_ImportOFX.EstDateValide(seg0)

                If Not verrouillee Then
                    plageTable.Cells(ligneCible - 1, colNotes).value = CStr(tabEcran(i, posNotes))
                End If

                ' On efface la marque : IMPORTANT de le faire AVANT les recalculs plus bas,
                ' sinon Workbook_SheetDeactivate reproposerait d'appliquer ces memes lignes
                plageSortieEcriture.Cells(i, posValider).value = ""

                nbAppliquees = nbAppliquees + 1
            End If
        End If
    Next i

    ' On remet les evenements AVANT les recalculs : les feuilles de rapprochement ouvertes par
    ' VerifierNotesSante ont besoin des evenements pour fonctionner
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ' 7. Recalcul du suivi sante (uniquement si quelque chose a ete applique)
    If nbAppliquees > 0 Then
        mod_FormulairesNotes.VerifierNotesSante
        mod_SuiviSante.CalculerSuiviSante AfficherResume:=False
    End If

    MsgBox nbAppliquees & FR(" ligne(s) appliqu{e2}e(s)."), vbInformation
    Exit Sub

GestionErreur:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox FR("Erreur pendant l'application des lignes marqu{e2}es :") & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical

End Sub
