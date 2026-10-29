Option Explicit
' Ce module regroupe les macros qui préparent une ZONE D'AFFICHAGE sur une
' feuille (nettoyage, mise en place des en-tétes) AVANT que d'autres modules
' n'y écrivent des données. Il ne contient volontairement AUCUNE logique de
' calcul ou de filtrage : uniquement de la mise en forme / nettoyage visuel.

' PHASE 6 : l'ancien Workbook_SheetBeforeDoubleClick (qui appelait
' mod_Actions.DoubleClick) et le bloc de Workbook_SheetChange lie a
' NOM_FEUILLE_RESULTAT/RecherOperations ont ete supprimes a l'epoque : les 2
' etaient devenus du code mort avec la suppression de la famille d'ecrans
' "Synthese_*" et de l'ancien systeme de double-clic sur les totaux (remplace
' par des boutons explicites).
'
' AJOUT 01/10/2026 (ergonomie, retour operateur) : un NOUVEAU
' Workbook_SheetBeforeDoubleClick est recree ci-dessous, mais volontairement
' tres cible -- contrairement a l'ancien (generique, declenche sur toutes les
' feuilles), celui-ci ne fait quelque chose QUE sur la feuille
' frm_RechercheOperations, et UNIQUEMENT dans les colonnes "Ventile" et
' "Categorie". Sur toute autre feuille, ou toute autre colonne de cette
' feuille, Excel garde son comportement normal (double-clic = entrer en mode
' edition de la cellule).
'
' MISE A JOUR 02/10/2026 (apres discussion avec l'operateur) :
'   - Colonne "Ventile" : un double-clic rouvre desormais TOUJOURS
'     mod_RechercheOperations.RevoirVentilationRO, que la ligne soit deja
'     ventilee (Oui -- revoir/supprimer le detail) ou non (cellule vide --
'     demarrer une nouvelle ventilation) : c'est RevoirVentilationRO qui fait
'     la distinction et tout le travail, ce handler se contente de
'     selectionner la bonne ligne puis de l'appeler.
'   - Colonne "Categorie" (nouveau) : un double-clic ouvre
'     mod_RechercheOperations.EditerCategorieRO, qui reutilise le formulaire
'     habituel de controle des categories (mod_ControleCategories) pour
'     corriger la Categorie/SousCategorie de cette seule operation. Si la
'     ligne est deja ventilee, EditerCategorieRO affiche elle-meme un message
'     d'information explicatif plutot que de ne rien faire silencieusement.
Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)

    ' On ne s'interesse qu'a l'ecran central de recherche.
    If Sh.Name <> mod_VarGlobales.NOM_FEUILLE_RECHERCHE Then Exit Sub

    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = Sh.ListObjects(mod_VarGlobales.NOM_TABLE_RECHERCHE)
    On Error GoTo 0
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    ' On ne s'interesse qu'a un double-clic dans une ligne de donnees du
    ' tableau (jamais dans l'en-tete, ni en dehors du tableau).
    If Intersect(Target, tbl.DataBodyRange) Is Nothing Then Exit Sub

    ' Colonnes "Ventile" et "Categorie" : on retrouve leur position par leur
    ' NOM (jamais par un numero fige), comme partout ailleurs sur ce chantier.
    Dim colVentile As Long, colCategorie As Long
    On Error Resume Next
    colVentile = tbl.ListColumns("Ventile").index
    colCategorie = tbl.ListColumns("Categorie").index
    On Error GoTo 0

    If colVentile <> 0 And Target.Column = tbl.Range.Columns(colVentile).Column Then
        ' Cancel = True empeche Excel de passer la cellule en mode edition
        ' (comportement par defaut d'un double-clic) : on prend la main a la place.
        Cancel = True
        Target.Select
        mod_RechercheOperations.RevoirVentilationRO
        Exit Sub
    End If

    If colCategorie <> 0 And Target.Column = tbl.Range.Columns(colCategorie).Column Then
        Cancel = True
        Target.Select
        mod_RechercheOperations.EditerCategorieRO
        Exit Sub
    End If
 
    ' AJOUT 07/10/2026, REVU 08/10/2026 (décision opérateur) : un double-clic sur une
    ' cellule SousCategorie se comporte EXACTEMENT comme sur Categorie, dans tous les
    ' préfiltres : il ouvre le formulaire habituel de contrôle des catégories
    ' (mod_RechercheOperations.EditerCategorieRO), qui modifie la catégorie ET la
    ' sous-catégorie ensemble. Après validation, EditerCategorieRO recalcule déjà les
    ' statuts santé EN SILENCE (CalculerSuiviSante AfficherResume:=False) puis recharge
    ' l'écran dans le même préfiltre : l'ancien double-clic de "recalcul" devient
    ' donc inutile et a été supprimé (RecalculerSuiviSanteRO, PrefiltreActifRO).
    ' La variable locale s'appelle colSousCat (et non colSousCategorie) pour ne pas
    ' masquer la variable publique du même nom déclarée dans mod_VarGlobales.
    Dim colSousCat As Long
    On Error Resume Next
    colSousCat = tbl.ListColumns("SousCategorie").index
    On Error GoTo 0
 
    ' Deux "If" imbriqués, et non un seul "If ... And ...". En VBA, un "And" évalue
    ' TOUJOURS ses deux côtés : si colSousCat valait 0, tbl.Range.Columns(0)
    ' provoquerait une erreur, même avec colSousCat <> 0 déjà faux.
    If colSousCat <> 0 Then
        If Target.Column = tbl.Range.Columns(colSousCat).Column Then
            Cancel = True
            Target.Select
            mod_RechercheOperations.EditerCategorieRO
            Exit Sub
        End If
    End If
 
End Sub

' AJOUT 11/10/2026 (pointage des opérations, voir mod_Pointage) : toute modification d'une
' cellule de pointage est interceptée ICI, pour tous les écrans à la fois (colonne Pointe de
' frm_RechercheOperations et champ "Pointé" des formulaires). Ce point d'entrée unique évite de
' coller du code dans chaque feuille. mod_Pointage sort aussitôt si la cellule modifiée n'est
' pas une cellule de pointage : le coût est négligeable pour toutes les autres feuilles.
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    mod_Pointage.TraiterChangementFeuille Sh, Target
End Sub

' AJOUT 09/10/2026 (échéancier) : à l'ouverture du classeur, complète les occurrences des règles
' de récurrence jusqu'à l'horizon, pour que le solde prévisionnel ne manque jamais d'échéances
' même si aucun import n'a été fait depuis longtemps. Sans effet si l'échéancier n'est pas installé.
'
' AJOUT 11/10/2026 (déploiement et sauvegardes, voir mod_Synchro) : la synchronisation avec le
' fichier central du réseau passe EN PREMIER. Si l'échéancier était complété avant, ce
' complément compterait comme une modification locale et créerait un conflit inutile.
Private Sub Workbook_Open()
    mod_Synchro.SynchroniserALOuverture
    mod_Echeancier.MettreAJourEcheancier
    PositionnerSurSynthese
End Sub

' AJOUT 29/10/2026 : le classeur s'ouvre toujours sur la feuille Synthese (si elle existe et est
' visible). L'activation declenche Workbook_SheetActivate, donc les rappels sont recalcules.
Private Sub PositionnerSurSynthese()
    Dim feuilleSynthese As Worksheet
    On Error Resume Next
    Set feuilleSynthese = Me.Worksheets(mod_VarGlobales.NOM_FEUILLE_SYNTHESE)
    If Not feuilleSynthese Is Nothing Then
        If feuilleSynthese.Visible = xlSheetVisible Then feuilleSynthese.Activate
    End If
    On Error GoTo 0
End Sub

' AJOUT 11/10/2026 : à la fermeture, envoi des données vers le fichier central. Si la
' synchronisation échoue et que l'opérateur refuse de fermer sans elle, la fermeture est annulée.
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If Not mod_Synchro.SynchroniserALaFermeture() Then Cancel = True
End Sub

' AJOUT 19/10/2026 : les listes de relevés et de contrats (consommations) sont une copie
' modifiable des feuilles de données Conso_Releves / Conso_Contrats ; une liste encore
' ouverte est écrite dans les données avant l'enregistrement.
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    mod_Consommations.EnregistrerListesSiOuvertes
End Sub

' AJOUT 28/10/2026 (feuille Synthese, voir mod_Synthese) : les rappels (echeances en retard,
' factures non liees, sante a traiter) sont recalcules a chaque affichage de la feuille Synthese.
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    If StrComp(Sh.Name, mod_VarGlobales.NOM_FEUILLE_SYNTHESE, vbTextCompare) = 0 Then mod_Synthese.ActualiserSynthese
End Sub
