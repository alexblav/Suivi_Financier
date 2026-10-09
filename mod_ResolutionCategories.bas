Option Explicit

' =====================================================================================
' MODULE : mod_ResolutionCategories
'
' RÔLE (phase 2 du chantier "UserForm -> feuille dédiée") :
'   Ce module contient toute la logique de la feuille frm_ResolutionCategories :
'     - remplissage du tableau à partir des variables g_CasTexte/g_CasCandidats;
'     - affichage bloquant remplaçant l'ancien frmResolutionCategories.Show vbModal;
'     - actions des deux boutons (Réinitialiser cette ligne / Terminer et appliquer).
'
'   Il s'appuie sur la mise en page déjà construite par mod_InstallResolutionSheet
'   (mêmes constantes de position des cellules, réutilisées telles quelles).
'
'   ATTENTION : la gestion des événements Worksheet_Change et Worksheet_Deactivate
'   ne peut PAS être placée dans ce module. En VBA, les événements d'une feuille
'   doivent obligatoirement se trouver dans le module de code de LA FEUILLE ELLE-MÊME
'   (le "code-behind"). Ce code est fourni à part, avec des instructions précises
'   pour le coller au bon endroit (voir la section correspondante plus bas).
' =====================================================================================

' --- Variable de synchronisation pour le comportement "modal" -----------------------
' Tant qu'elle vaut True, la boucle d'attente (voir AfficherFeuilleResolutionEtAttendre)
' ne rend pas la main a mod_ImportOFX, et le code-behind de la feuille empeche
' l'utilisateur de changer d'onglet (voir Worksheet_Deactivate, fourni a part).
Public g_SaisieEnCours As Boolean

' --- Garde-fou anti-recursion pendant le remplissage automatique de la feuille ------
' Quand RemplirTableauCas() ecrit des valeurs dans les cellules, cela declenche
' normalement l'evenement Worksheet_Change (comme si l'utilisateur avait tape).
' On ne veut PAS que la logique de "l'utilisateur vient de choisir une categorie"
' se declenche pendant qu'on remplit le tableau nous-memes au demarrage : ce
' drapeau permet au code-behind de la feuille de savoir qu'il doit ignorer les
' changements en cours pendant cette phase de remplissage automatique.
Public g_ChargementEnCours As Boolean

' --- Memorise quelle feuille etait active avant l'ouverture, pour y revenir apres ---
Private g_NomFeuillePrecedente As String

' --- Zone technique cachee : une ligne par cas ambigu, une colonne par categorie
' candidate. Sert de SOURCE aux listes deroulantes (voir explication detaillee
' dans RemplirTableauCas). Placee loin a droite du tableau visible pour ne
' jamais interferer avec la mise en page, et masquee (colonnes cachees). ---
Public Const COL_AIDE_CANDIDATS_DEBUT As String = "Z"
Public Const NB_COL_AIDE_MAX As Long = 15   ' marge large : jusqu'a 15 categories candidates par cas

' =====================================================================================
' PROCEDURE PRINCIPALE - remplace l'ancien "frmResolutionCategories.Show vbModal"
' A appeler depuis mod_ImportOFX exactement a l'endroit ou se trouvait l'ancien appel.
' Cette procedure ne rend la main a l'appelant qu'une fois que l'utilisateur a
' clique sur "Terminer et appliquer" -- comme le faisait le UserForm modal.
' =====================================================================================
Public Sub AfficherFeuilleResolutionEtAttendre()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_RESOLUTION)

    ' On retient la feuille actuellement affichee, pour y revenir automatiquement
    ' une fois la resolution terminee (l'utilisateur ne doit pas "atterrir" sur
    ' une feuille inattendue apres avoir clique sur Terminer).
    g_NomFeuillePrecedente = ActiveSheet.Name

    ' On remplit le tableau AVANT de rendre la feuille visible, pour que
    ' l'utilisateur ne voie jamais un tableau vide se remplir sous ses yeux.
    Call RemplirTableauCas(ws)

    ' On revele la feuille et on la place au premier plan.
    ws.Visible = xlSheetVisible
    ws.Activate
    mod_InstallCommun.MasquerQuadrillage   ' quadrillage et en-tetes toujours masques (09/10/2026)
    ws.Range(COL_CATEGORIE & LIGNE_PREMIERE_DONNEE).Select

    ' --- Verrouillage "modal" ---
    g_SaisieEnCours = True

    ' Boucle d'attente active : le code reste bloque ICI, sans planter Excel
    ' (DoEvents laisse Excel traiter les clics de l'utilisateur), jusqu'a ce que
    ' g_SaisieEnCours repasse a False -- ce que fera la macro TerminerEtAppliquerChoix
    ' quand l'utilisateur cliquera sur le bouton "Terminer et appliquer".
    Do While g_SaisieEnCours
        DoEvents
    Loop

    ' A ce stade, l'utilisateur a clique sur "Terminer et appliquer".
    ' g_CasChoix() a deja ete rempli au fil de l'eau par le code-behind de la
    ' feuille (Worksheet_Change) -- mod_ImportOFX peut le relire normalement,
    ' exactement comme avant avec le UserForm.

End Sub


' =====================================================================================
' Remplit le tableau de la feuille a partir de g_NbCasAmbigus / g_CasTexte / g_CasCandidats
' =====================================================================================
Private Sub RemplirTableauCas(ws As Worksheet)

    Dim i As Long, ligneCible As Long
    Dim partiesTexte() As String
    Dim DateTexte As String, montantTexte As String, tiersTexte As String
    Dim montantNombre As Double

    ' On leve le garde-fou : le code-behind de la feuille (Worksheet_Change)
    ' saura qu'il doit ignorer les evenements declenches par CE remplissage,
    ' et non par une vraie action de l'utilisateur.
    g_ChargementEnCours = True

    ' --- Nettoyage prealable de la zone de donnees (au cas ou une resolution
    ' precedente aurait laisse des lignes remplies) ---
    Dim derniereLigneEfface As Long
    derniereLigneEfface = LIGNE_PREMIERE_DONNEE + NB_LIGNES_PREPAREES - 1
    With ws.Range(COL_STATUT & LIGNE_PREMIERE_DONNEE & ":" & COL_CATEGORIE & derniereLigneEfface)
        .ClearContents
        ' On supprime aussi d'anciennes listes deroulantes (validation de donnees)
        ' pour repartir sur une base propre a chaque nouvel import.
        .Validation.Delete
        ' Ajout 09/10/2026 : les fonds (gris / jaune) sont reposes ci-dessous, sur les
        ' seules lignes utilisees (voir mod_InstallCommun pour la charte de couleurs).
        .Interior.ColorIndex = xlColorIndexNone
    End With

    ' On nettoie egalement la zone technique cachee (les anciennes listes de
    ' categories candidates ecrites lors d'un import precedent).
    Dim colAideIndex As Long
    colAideIndex = ws.Range(COL_AIDE_CANDIDATS_DEBUT & "1").Column
    ws.Range(ws.Cells(LIGNE_PREMIERE_DONNEE, colAideIndex), _
             ws.Cells(derniereLigneEfface, colAideIndex + NB_COL_AIDE_MAX - 1)).ClearContents

    ' On s'assure que cette zone technique reste invisible pour l'utilisateur
    ' (operation sans effet si les colonnes sont deja masquees -- sans danger
    ' de la refaire a chaque remplissage).
    ws.Range(ws.Cells(1, colAideIndex), ws.Cells(1, colAideIndex + NB_COL_AIDE_MAX - 1)) _
        .EntireColumn.Hidden = True

    ' --- Remplissage ligne par ligne ---
    For i = 1 To g_NbCasAmbigus
        ligneCible = LIGNE_PREMIERE_DONNEE + i - 1

        ' g_CasTexte(i) est de la forme "dd/mm/yyyy  |  0,00 EUR  |  Nom du tiers"
        ' (le separateur decimal est la virgule car Format() suit les parametres
        ' regionaux francais du poste -- coherent avec le reste du classeur).
        partiesTexte = Split(g_CasTexte(i), "|")
        DateTexte = Trim(partiesTexte(0))
        montantTexte = Trim(partiesTexte(1))
        tiersTexte = Trim(partiesTexte(2))

        ' On retire le suffixe " EUR" et on reconvertit en nombre reel, pour que
        ' la colonne Montant reste un vrai nombre (alignement a droite, tri
        ' eventuel, format monetaire deja prepare en Phase 1) plutot qu'un texte.
        montantTexte = Replace(montantTexte, "EUR", "")
        montantTexte = Trim(montantTexte)
        montantNombre = CDbl(montantTexte)   ' CDbl respecte la virgule francaise

        ' --- Ecriture des cellules de la ligne ---
        ' Ajout 09/10/2026 : gris = cellules non modifiables (Statut a Tiers),
        ' jaune pale = cellule modifiable (choix de la categorie).
        ws.Range(COL_STATUT & ligneCible & ":" & COL_TIERS & ligneCible).Interior.Color = mod_InstallCommun.CoulFondLecture()
        ws.Range(COL_CATEGORIE & ligneCible).Interior.Color = mod_InstallCommun.CoulFondSaisie()

        ws.Range(COL_STATUT & ligneCible).value = "?"
        ws.Range(COL_STATUT & ligneCible).Font.Color = RGB(150, 150, 150)
        ws.Range(COL_STATUT & ligneCible).Font.Bold = False

        ws.Range(COL_DATE & ligneCible).value = DateTexte
        ws.Range(COL_MONTANT & ligneCible).value = montantNombre
        ws.Range(COL_TIERS & ligneCible).value = tiersTexte

        ' --- Liste deroulante de categorie, PROPRE A CETTE LIGNE ---
        ' g_CasCandidats(i) est au format "Categorie1;Categorie2;...". On aurait pu
        ' donner directement cette chaine a Formula1, MAIS Excel n'interprete pas
        ' toujours ";" comme un separateur d'elements dans ce contexte precis (c'est
        ' ce qui a cause le bug observe : toute la chaine consideree comme UNE seule
        ' valeur). Solution fiable : on decoupe nous-memes la chaine avec Split()
        ' (fonction VBA, independante des reglages d'Excel), on ecrit chaque
        ' categorie candidate dans une cellule separee de la zone technique cachee,
        ' puis on donne a la liste deroulante une REFERENCE DE PLAGE (ex. $Z$12:$AA$12)
        ' plutot qu'une chaine de texte. Une plage de cellules ne peut pas etre
        ' mal interpretee : chaque cellule est forcement un element distinct.
        Dim candidatsTableau() As String
        Dim nbCandidatsCase As Long
        Dim k As Long

        candidatsTableau = Split(g_CasCandidats(i), ";")
        nbCandidatsCase = UBound(candidatsTableau) - LBound(candidatsTableau) + 1

        For k = 0 To nbCandidatsCase - 1
            ws.Cells(ligneCible, colAideIndex + k).value = Trim(candidatsTableau(k))
        Next k

        Dim adresseDebut As String, adresseFin As String
        adresseDebut = ws.Cells(ligneCible, colAideIndex).Address(True, True)
        adresseFin = ws.Cells(ligneCible, colAideIndex + nbCandidatsCase - 1).Address(True, True)

        With ws.Range(COL_CATEGORIE & ligneCible).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:="=" & adresseDebut & ":" & adresseFin
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowError = True
            .ErrorTitle = "Categorie invalide"
            .ErrorMessage = "Merci de choisir une categorie dans la liste proposee."
        End With

        ' Si un choix avait deja ete fait auparavant (cas rare : reprise apres
        ' interruption), on le reaffiche et on marque la ligne comme traitee.
        If g_CasChoix(i) <> "" Then
            ws.Range(COL_CATEGORIE & ligneCible).value = g_CasChoix(i)
            ws.Range(COL_STATUT & ligneCible).value = ChrW(&H2713)   ' caractere "check" (V)
            ws.Range(COL_STATUT & ligneCible).Font.Color = RGB(30, 130, 76)
            ws.Range(COL_STATUT & ligneCible).Font.Bold = True
        End If
    Next i

    Call MettreAJourCompteur(ws)

    g_ChargementEnCours = False

End Sub


' =====================================================================================
' Met a jour le texte "X restant(s) sur Y" dans la cellule nommee CompteurCasRestants.
' Appelee au demarrage ET a chaque fois qu'une categorie est choisie/effacee.
' PUBLIC car elle sera aussi appelee depuis le code-behind de la feuille (Worksheet_Change).
' =====================================================================================
Public Sub MettreAJourCompteur(ws As Worksheet)

    Dim i As Long, nbRestants As Long
    For i = 1 To g_NbCasAmbigus
        If g_CasChoix(i) = "" Then nbRestants = nbRestants + 1
    Next i

    ' Format commun a tous les formulaires : "Operation: x/y" (09/10/2026). Ici x = nombre
    ' d'operations deja traitees, y = nombre total d'operations a traiter.
    ws.Range("CompteurCasRestants").value = mod_InstallCommun.TexteCompteur(g_NbCasAmbigus - nbRestants, g_NbCasAmbigus)

End Sub


' =====================================================================================
' ACTION DU BOUTON "Réinitialiser cette ligne"
' Efface le choix de categorie de la ligne actuellement selectionnee. L'effacement
' de la cellule declenche automatiquement Worksheet_Change (code-behind), qui se
' charge lui-meme de remettre le statut a "?" et de mettre a jour g_CasChoix et
' le compteur -- on ne duplique donc pas cette logique ici.
' =====================================================================================
Public Sub ReinitialiserLigneSelectionnee()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_RESOLUTION)

    Dim ligneSelection As Long
    ligneSelection = ActiveCell.Row

    ' Garde-fou : on verifie que la selection est bien a l'interieur de la zone
    ' de donnees du tableau, pour ne pas effacer autre chose par erreur si
    ' l'utilisateur avait clique ailleurs (ex. dans le bloc d'instructions).
    If ligneSelection < LIGNE_PREMIERE_DONNEE Or ligneSelection > LIGNE_PREMIERE_DONNEE + g_NbCasAmbigus - 1 Then
        MsgBox "Selectionne d'abord une ligne d'operation dans le tableau, " & _
               "puis clique sur ce bouton pour reinitialiser son choix.", vbExclamation
        Exit Sub
    End If

    ws.Range(COL_CATEGORIE & ligneSelection).ClearContents

End Sub


' =====================================================================================
' ACTION DU BOUTON "Terminer et appliquer"
' Ferme la feuille et rend la main a mod_ImportOFX (deverrouille la boucle d'attente).
' =====================================================================================
Public Sub TerminerEtAppliquerChoix()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_RESOLUTION)

    ' --- Avertissement si des cas restent non traites, comme le faisait
    ' implicitement l'ancien formulaire (les cas non traites gardent
    ' g_CasChoix(i) = "" et ne seront simplement pas modifies dans le tableau
    ' final -- ce comportement est conserve a l'identique). ---
    Dim i As Long, nbRestants As Long
    For i = 1 To g_NbCasAmbigus
        If g_CasChoix(i) = "" Then nbRestants = nbRestants + 1
    Next i

    If nbRestants > 0 Then
        Dim reponse As VbMsgBoxResult
        reponse = MsgBox( _
            nbRestants & " operation(s) n'ont pas encore de categorie choisie." & vbCrLf & _
            "Elles resteront sans categorie dans le tableau final." & vbCrLf & vbCrLf & _
            "Veux-tu vraiment terminer maintenant ?", _
            vbYesNo + vbQuestion, "Cas non resolus")
        If reponse = vbNo Then Exit Sub   ' l'utilisateur reste sur la feuille
    End If

    ' --- Deverrouillage : la boucle d'attente dans AfficherFeuilleResolutionEtAttendre
    ' va se terminer au prochain passage (elle teste g_SaisieEnCours en boucle). ---
    g_SaisieEnCours = False

    ' --- Nettoyage de la feuille pour la prochaine utilisation ---
    Dim derniereLigneEfface As Long
    derniereLigneEfface = LIGNE_PREMIERE_DONNEE + NB_LIGNES_PREPAREES - 1
    With ws.Range(COL_STATUT & LIGNE_PREMIERE_DONNEE & ":" & COL_CATEGORIE & derniereLigneEfface)
        .ClearContents
        .Validation.Delete
        .Interior.ColorIndex = xlColorIndexNone
    End With

    ' Nettoyage de la zone technique cachee (listes de categories candidates)
    Dim colAideIndex As Long
    colAideIndex = ws.Range(COL_AIDE_CANDIDATS_DEBUT & "1").Column
    ws.Range(ws.Cells(LIGNE_PREMIERE_DONNEE, colAideIndex), _
             ws.Cells(derniereLigneEfface, colAideIndex + NB_COL_AIDE_MAX - 1)).ClearContents

    ws.Range("CompteurCasRestants").value = ""

    ' --- Masquage et retour a la feuille sur laquelle etait l'utilisateur avant ---
    ws.Visible = xlSheetVeryHidden

    On Error Resume Next   ' au cas ou la feuille precedente aurait ete renommee/supprimee entre-temps
    ThisWorkbook.Worksheets(g_NomFeuillePrecedente).Activate
    On Error GoTo 0

End Sub
