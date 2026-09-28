Option Explicit
' Ce module regroupe les macros qui préparent une ZONE D'AFFICHAGE sur une
' feuille (nettoyage, mise en place des en-têtes) AVANT que d'autres modules
' n'y écrivent des données. Il ne contient volontairement AUCUNE logique de
' calcul ou de filtrage : uniquement de la mise en forme / nettoyage visuel.

' PHASE 6 : Workbook_SheetBeforeDoubleClick (qui appelait mod_Actions.DoubleClick)
' et le bloc de Workbook_SheetChange lié à NOM_FEUILLE_RESULTAT/RecherOperations
' sont supprimés ici. Les 2 étaient devenus du code mort avec la suppression de
' la famille d'écrans "Synthese_*" et de l'ancien système de double-clic sur les
' totaux (remplacé par des boutons explicites) : plus rien, nulle part dans le
' classeur, ne met plus jamais RecherOperations à True ni AllowDetailDoubleClick
' à True. Si un jour un comportement similaire est de nouveau nécessaire sur
' l'écran central frm_RechercheOperations, il faudra le recréer spécifiquement
' pour cette feuille (mod_InstallRechercheOperations.NOM_FEUILLE_RECHERCHE),
' pas le récupérer tel quel ici.