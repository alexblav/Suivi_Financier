Option Explicit
' Ce module regroupe les macros qui préparent une ZONE D'AFFICHAGE sur une
' feuille (nettoyage, mise en place des en-têtes) AVANT que d'autres modules
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
' frm_RechercheOperations, et UNIQUEMENT dans la colonne "Ventile" : un
' double-clic y rouvre directement le detail de la ventilation de la ligne
' concernee (meme resultat que de selectionner la ligne puis cliquer sur le
' bouton "Revoir la ventilation" -- mod_RechercheOperations.RevoirVentilationRO
' fait d'ailleurs tout le travail de validation, ce handler ne fait que
' selectionner la bonne ligne puis l'appeler). Sur toute autre feuille, ou
' toute autre colonne de cette feuille, Excel garde son comportement normal
' (double-clic = entrer en mode edition de la cellule).
Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)

    ' On ne s'interesse qu'a l'ecran central de recherche.
    If Sh.Name <> mod_InstallRechercheOperations.NOM_FEUILLE_RECHERCHE Then Exit Sub

    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = Sh.ListObjects(mod_InstallRechercheOperations.NOM_TABLE_RECHERCHE)
    On Error GoTo 0
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    ' Colonne "Ventile" : on retrouve sa position par son NOM (jamais par un
    ' numero fige), comme partout ailleurs sur ce chantier.
    Dim colVentile As Long
    On Error Resume Next
    colVentile = tbl.ListColumns("Ventile").index
    On Error GoTo 0
    If colVentile = 0 Then Exit Sub

    ' Le double-clic doit tomber dans la colonne Ventile ET dans une ligne de
    ' donnees du tableau (jamais dans l'en-tete, ni une autre colonne).
    If Target.Column <> tbl.Range.Columns(colVentile).Column Then Exit Sub
    If Intersect(Target, tbl.DataBodyRange) Is Nothing Then Exit Sub

    ' Cancel = True empeche Excel de passer la cellule en mode edition
    ' (comportement par defaut d'un double-clic) : on prend la main a la place.
    Cancel = True
    Target.Select
    mod_RechercheOperations.RevoirVentilationRO

End Sub