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
    If Sh.Name <> mod_InstallRechercheOperations.NOM_FEUILLE_RECHERCHE Then Exit Sub

    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = Sh.ListObjects(mod_InstallRechercheOperations.NOM_TABLE_RECHERCHE)
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

End Sub