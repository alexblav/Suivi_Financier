Option Explicit
 
' =====================================================================================
' MODULE : Mod_SyntheseCare
'
' RÔLE : point d'entrée du suivi santé. La macro garde son nom historique
' « Synthese_Care » pour que le bouton de la feuille Synthese qui l'appelle
' continue de fonctionner sans aucune retouche.
'
' REFONTE du 07/10/2026 (validée par l'opérateur, « piste 2 ») : l'ancien écran
' dédié (feuille frm_Resultat construite ici, avec son propre rapprochement OK/KO)
' est abandonné. Il présentait 4 défauts :
'   1. il testait la colonne Categorie, qui contient la catégorie PARENTE depuis
'      la Phase 1 : il ne trouvait donc plus aucune ligne ;
'   2. il ignorait les parts de ventilation (TblVentilations) ;
'   3. il recalculait son propre statut OK/KO, différent de celui calculé par
'      mod_SuiviSante.CalculerSuiviSante (franchise, écarts de montant et
'      dépassement d'honoraires ignorés) ;
'   4. il laissait Application.EnableEvents à False quand aucune ligne n'était
'      trouvée, ce qui désactivait tous les événements du classeur.
' Désormais, l'affichage passe par l'écran central frm_RechercheOperations, avec le
' préfiltre "SuiviSante" (voir mod_RechercheOperations.LigneOpRetenuePourPrefiltre).
' Le statut affiché est celui de la colonne StatutSante, seule source de vérité.
'
' REMARQUE : le type mod_Rapports.TOperation n'est plus utilisé par ce module. Il
' est volontairement laissé en place jusqu'au passage de nettoyage prévu après la
' mise en production.
' =====================================================================================
 
Public Sub Synthese_Care()
    mod_RechercheOperations.RechercherOperations "SuiviSante"
End Sub