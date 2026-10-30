Attribute VB_Name = "mod_InstallAide"
Option Explicit

' =====================================================================================
' MODULE : mod_InstallAide
'
' RÔLE : contenu des feuilles d'aide frm_Nomform_aide (voir mod_InstallCommun.EcrireInstructions
' et mod_Aide). Pour chaque formulaire :
'   - un TITRE ;
'   - une INTRODUCTION (le principe du formulaire) ;
'   - la liste des CHAMPS : ce qui est attendu dans chacun ;
'   - la liste des BOUTONS : ce que fait chacun.
' Chaque élément des deux listes est écrit dans sa propre cellule de la feuille d'aide.
'
' FORMAT : une liste est un texte dont chaque ligne (séparée par Chr(10)) est "Nom|Description".
' La description accepte les balises du bloc Instructions : [b:Bouton], [c:Champ],
' [f:Feuille], ATTENTION ; les accents s'écrivent avec les marqueurs de mod_Display.FR
' ({e2} é, {e1} è, {ea} ê, {a2} à, {c2} ç, {o2} ô, {i2} î, {E2} É). Pas de "ù" ni de "°" dans ces
' marqueurs : écrire "ou" et "num{e2}ro". Le marqueur {crayon} est remplacé par le crayon
' des boutons Éditer.
' Pour modifier une aide : changer le texte ici, puis relancer la macro de construction
' du formulaire concerné (liste dans mod_InstallCommun).
' =====================================================================================

' Renvoie Vrai si le formulaire "nomForm" a un contenu d'aide défini ici.
Public Function ChargerAide(ByVal nomForm As String, ByRef titre As String, ByRef intro As String, _
                            ByRef champs As String, ByRef boutons As String) As Boolean

    ChargerAide = True
    Select Case nomForm
        Case "frm_Analyse":              Def_Analyse titre, intro, champs, boutons
        Case "frm_Evolution":            Def_Evolution titre, intro, champs, boutons
        Case "frm_Consommations":        Def_Consommations titre, intro, champs, boutons
        Case "frm_Factures":             Def_Factures titre, intro, champs, boutons
        Case "frm_Contrats":             Def_Contrats titre, intro, champs, boutons
        Case "frm_Facture":              Def_Facture titre, intro, champs, boutons
        Case "frm_Contrat":              Def_Contrat titre, intro, champs, boutons
        Case "frm_Echeancier":           Def_Echeancier titre, intro, champs, boutons
        Case "frm_Recurrences":          Def_Recurrences titre, intro, champs, boutons
        Case "frm_Echeance":             Def_Echeance titre, intro, champs, boutons
        Case "frm_ControleCategories":   Def_ControleCategories titre, intro, champs, boutons
        Case "frm_RapprochementNotes":   Def_RapprochementNotes titre, intro, champs, boutons
        Case "frm_GenerationCle":        Def_GenerationCle titre, intro, champs, boutons
        Case "frm_NouvelleCategorie":    Def_NouvelleCategorie titre, intro, champs, boutons
        Case "frm_RechercheOperations":  Def_RechercheOperations titre, intro, champs, boutons
        Case "frm_ResolutionCategories": Def_ResolutionCategories titre, intro, champs, boutons
        Case "frm_SuiviSante":           Def_SuiviSante titre, intro, champs, boutons
        Case "frm_Ventilation":          Def_Ventilation titre, intro, champs, boutons
        Case "frm_Remediation_OFX":      Def_RemediationOFX titre, intro, champs, boutons
        Case "frm_AffectationFacture":   Def_AffectationFacture titre, intro, champs, boutons
        Case Else:                       ChargerAide = False
    End Select

    If ChargerAide Then
        intro = Replace(intro, "{crayon}", mod_InstallCommun.CapEditer())
        champs = Replace(champs, "{crayon}", mod_InstallCommun.CapEditer())
        boutons = Replace(boutons, "{crayon}", mod_InstallCommun.CapEditer())
    End If

End Function

' Ajoute un élément "Nom|Description" à une liste.
Private Sub L(ByRef liste As String, ByVal nom As String, ByVal description As String)
    If Len(liste) > 0 Then liste = liste & Chr(10)
    liste = liste & nom & "|" & description
End Sub

' Pointé : même fonctionnement sur tous les formulaires qui l'affichent.
Private Sub LPointe(ByRef liste As String)
    L liste, "Point{e2}", "Oui si l'op{e2}ration est d{e2}j{a2} analys{e2}e, vide sinon. Le changement est enregistr{e2} tout de suite ; retirer un pointage demande une confirmation."
End Sub


' =====================================================================================
' ANALYSE ET EVOLUTION
' =====================================================================================
Private Sub Def_Analyse(ByRef titre As String, ByRef intro As String, ByRef champs As String, ByRef boutons As String)

    titre = "Analyse budg{e2}taire : r{e2}partition des d{e2}penses"
    intro = "Cet {e2}cran r{e2}partit les d{e2}penses par cat{e2}gorie sur une p{e2}riode, en tableau et en graphique. " & _
            "Les montants sont des d{e2}penses nettes des remboursements, par mois de budget. " & _
            "Apr{e1}s tout changement de param{e1}tre, cliquez sur [b:Actualiser]. Cliquez une ligne du tableau pour la s{e2}lectionner."

    L champs, "P{e2}riode", "Liste : Mois (le mois choisi), Ann{e2}e (l'ann{e2}e de budget enti{e1}re) ou 12 derniers mois."
    L champs, "Ann{e2}e de budget", "Liste : l'ann{e2}e analys{e2}e. Reprise de la feuille [f:Synthese] {a2} l'ouverture."
    L champs, "Mois de budget", "Liste : 1 {a2} 12. Sert pour la p{e2}riode Mois et pour la fin des 12 derniers mois."
    L champs, "Cat{e2}gorie {a2} d{e2}tailler", "Liste : (toutes) pour la vue g{e2}n{e2}rale, ou une cat{e2}gorie pour voir ses sous-cat{e2}gories."
    L champs, "Graphique de r{e2}partition", "Liste : Barres ou Anneau."
    L champs, "Cat{e2}gories exclues", "Jusqu'{a2} 6 cat{e2}gories (liste, cliquez une case) jamais compt{e2}es dans l'analyse. Par d{e2}faut : Virements internes et Revenus."
    L champs, "Tableau de r{e2}partition", "Lecture seule : cat{e2}gorie ou sous-cat{e2}gorie, d{e2}penses en euros, part en % et nombre d'op{e2}rations."
    L champs, "Ligne de r{e2}sum{e2}", "Lecture seule : rappelle la p{e2}riode et les totaux de l'affichage."

    L boutons, "Actualiser", "Recalcule le tableau et le graphique avec les param{e1}tres affich{e2}s."
    L boutons, "D{e2}tailler la s{e2}lection", "Affiche les sous-cat{e2}gories de la cat{e2}gorie de la ligne s{e2}lectionn{e2}e dans le tableau."
    L boutons, "Toutes les cat{e2}gories", "Revient {a2} la vue g{e2}n{e2}rale, toutes cat{e2}gories."
    L boutons, "Voir les op{e2}rations", "Ouvre [f:frm_RechercheOperations] sur les op{e2}rations de la ligne s{e2}lectionn{e2}e (ou de tout l'affichage si aucune ligne n'est s{e2}lectionn{e2}e). Au retour, l'analyse est recalcul{e2}e."
    L boutons, "Aide", "Ouvre cette aide, sans fermer l'analyse."
    L boutons, "Sortir", "Ferme l'analyse (et cette aide si elle est ouverte) et revient sur [f:Synthese]."

End Sub

Private Sub Def_Evolution(ByRef titre As String, ByRef intro As String, ByRef champs As String, ByRef boutons As String)

    titre = "{E2}volution des d{e2}penses dans le temps"
    intro = "Cet {e2}cran montre l'{e2}volution des d{e2}penses mois apr{e1}s mois (ou trimestre, ou ann{e2}e), en tableau et en graphique. " & _
            "Les montants sont des d{e2}penses nettes des remboursements, par mois de budget. " & _
            "Sans choix de cat{e2}gorie, toutes les cat{e2}gories sont prises, sauf celles de Cat{e2}gories exclues."

    L champs, "P{e2}riode", "Liste : Mois en cours, Mois dernier, 3, 6, 12, 24 ou 36 derniers mois, ou Tous (tout l'historique)."
    L champs, "Pas de temps", "Liste : Automatique (choisi selon la dur{e2}e), Mois, Trimestre ou Ann{e2}e. Regroupe les colonnes du tableau."
    L champs, "Graphique", "Liste : Barres empil{e2}es, Courbes ou Barres group{e2}es."
    L champs, "D{e2}tail par", "Liste : Automatique, Cat{e2}gorie ou Sous-cat{e2}gorie. Avec une seule cat{e2}gorie choisie, Automatique passe aux sous-cat{e2}gories."
    L champs, "Cat{e2}gories (toutes si vide)", "Liste de cat{e2}gories {a2} analyser (cliquez une case). Vide = toutes."
    L champs, "Sous-cat{e2}gories (toutes si vide)", "Liste de sous-cat{e2}gories {a2} analyser. Vide = toutes."
    L champs, "Cat{e2}gories exclues", "Cat{e2}gories jamais compt{e2}es (par d{e2}faut Virements internes et Revenus). ATTENTION : elles ne s'appliquent plus d{e1}s qu'un choix est fait dans la liste Cat{e2}gories."
    L champs, "Ligne de message", "Lecture seule : p{e2}riode, pas de temps et totaux du r{e2}sultat affich{e2}."

    L boutons, "Actualiser", "Recalcule le tableau et le graphique apr{e1}s un changement de param{e1}tre."
    L boutons, "R{e2}initialiser les s{e2}lections", "Vide les listes de cat{e2}gories et de sous-cat{e2}gories (retour {a2} toutes)."
    L boutons, "Aide", "Ouvre cette aide, sans fermer l'{e2}cran."
    L boutons, "Sortir", "Ferme l'{e2}cran (et cette aide si elle est ouverte) et revient sur [f:Synthese]."

End Sub


' =====================================================================================
' CONSOMMATIONS (eau, électricité, gaz)
' =====================================================================================
Private Sub Def_Consommations(ByRef titre As String, ByRef intro As String, ByRef champs As String, ByRef boutons As String)

    titre = "Suivi des consommations"
    intro = "Cet {e2}cran suit l'{e2}volution des consommations d'eau, d'{e2}lectricit{e2} et de gaz, mois apr{e1}s mois, en tableau et en graphique. " & _
            "Apr{e1}s tout changement de param{e1}tre, cliquez sur [b:Actualiser]."

    L champs, "Type", "Choisir en premier (les autres champs sont masqu{e2}s tant qu'il est vide) : un type de consommation, ou (tous) pour comparer les montants des diff{e2}rents types. Il d{e2}termine les mois propos{e2}s dans D{e2}but et Fin."
    L champs, "Indicateur", "Consommation (la quantit{e2} consomm{e2}e, en m3 pour l'eau, en kWh pour l'{e2}nergie), Prix unitaire HT (moyenne pond{e2}r{e2}e par les quantit{e2}s, abonnement et taxes exclus), Montant consommation, Abonnement, Taxes (montants TTC en euros) ou Montant total (leur somme). Pour (tous), seuls les montants sont comparables."
    L champs, "D{e2}but", "Appara{i2}t apr{e1}s l'indicateur. Premier mois de la p{e2}riode, {a2} choisir parmi les mois couverts par les factures du type choisi."
    L champs, "Fin", "Appara{i2}t apr{e1}s le d{e2}but. Dernier mois de la p{e2}riode, parmi les mois du type {a2} partir du d{e2}but. Laiss{e2} vide, la p{e2}riode va jusqu'au dernier mois du type."
    L champs, "Forme du graphique", "Barres (empil{e2}es par op{e2}rateur, ou par type pour (tous)), Courbes, ou Barres + courbe du prix (pour un type choisi : les barres montrent l'indicateur, le prix unitaire se lit sur l'axe de droite)."
    L champs, "R{e2}partition des lignes", "Au prorata des jours : chaque ligne de facture est r{e2}partie sur les mois qu'elle couvre. Mois de fin de la ligne : elle est imput{e2}e enti{e1}rement au mois de sa date de fin."

    L boutons, "Actualiser", "Recalcule le tableau et le graphique."
    L boutons, "Saisir une facture", "Ouvre [f:frm_Facture] pour saisir une facture (abonnement, consommation, taxes)."
    L boutons, "Voir les factures", "Affiche [f:frm_Factures], la liste des factures (modification et suppression)."
    L boutons, "Contrats", "Affiche [f:frm_Contrats], la liste des contrats ; c'est l{a2} qu'on d{e2}clare un changement d'op{e2}rateur."
    L boutons, "Aide", "Ouvre cette aide, sans fermer l'{e2}cran."
    L boutons, "Sortir", "Ferme l'{e2}cran (et cette aide si elle est ouverte) et revient sur [f:Synthese]."

End Sub

Private Sub Def_Factures(ByRef titre As String, ByRef intro As String, ByRef champs As String, ByRef boutons As String)

    titre = "Liste des factures de consommation"
    intro = "Cette liste r{e2}capitule les factures. Une facture se compose d'un en-t{ea}te et de lignes d'abonnement, de consommation et de taxes, " & _
            "chacune avec ses propres dates, son prix unitaire HT, son taux de TVA et son montant TTC. La liste est en lecture seule : " & _
            "on ajoute ou on corrige une facture avec le formulaire ([b:Ajouter une facture], [b:Modifier la s{e2}lection]). Les factures sont class{e2}es par date de facture, les plus r{e2}centes en haut ; les fl{e1}ches de l'en-t{ea}te filtrent chaque colonne et le filtre est conserv{e2} quand la liste est actualis{e2}e."

    L champs, "ID_Facture", "Num{e2}ro de la facture, attribu{e2} automatiquement."
    L champs, "TypeConso", "Type de consommation (eau, eau recycl{e2}e, {e2}lectricit{e2}, gaz)."
    L champs, "Operateur", "Fournisseur de la facture."
    L champs, "NumeroFacture", "Num{e2}ro imprim{e2} sur la facture (facultatif)."
    L champs, "DateFacture", "Date de la facture (facultative)."
    L champs, "DebutPeriode", "D{e2}but de la p{e2}riode factur{e2}e (facultatif)."
    L champs, "FinPeriode", "Fin de la p{e2}riode factur{e2}e (facultatif)."
    L champs, "TotalHT", "Calcul{e2} : somme des montants HT des lignes."
    L champs, "TotalTVA", "Calcul{e2} : TVA des lignes."
    L champs, "TotalTTCCalcule", "Calcul{e2} : somme des montants TTC des lignes (indicatif)."
    L champs, "TotalFactureTTC", "Montant TTC r{e2}ellement indiqu{e2} sur la facture."
    L champs, "Ecart", "TotalFactureTTC - TotalTTCCalcule ; vide si le total de la facture n'a pas {e2}t{e2} saisi."
    L champs, "Commentaire", "Texte libre (facultatif)."

    L boutons, "Ajouter une facture", "Ouvre [f:frm_Facture] pour saisir une facture."
    L boutons, "Consulter la s{e2}lection", "Ouvre [f:frm_Facture] sur la facture s{e2}lectionn{e2}e en consultation : un seul bouton Sortir, qui abandonne en silence toute modification (rien ne peut {ea}tre enregistr{e2})."
    L boutons, "Modifier la s{e2}lection", "Rouvre [f:frm_Facture] sur la facture s{e2}lectionn{e2}e."
    L boutons, "Copier la s{e2}lection", "Ouvre [f:frm_Facture] pr{e2}rempli avec le type, l'op{e2}rateur, les libell{e2}s, les unit{e2}s et les taux de TVA de la facture s{e2}lectionn{e2}e ; num{e2}ro, dates, quantit{e2}s et prix sont laiss{e2}s vides. Utile pour saisir la facture suivante."
    L boutons, "Supprimer la s{e2}lection", "Supprime la facture s{e2}lectionn{e2}e et toutes ses lignes."
    L boutons, "Contrats", "Affiche [f:frm_Contrats], la liste des contrats."
    L boutons, "Retour", "Revient sur [f:frm_Consommations]."
    L boutons, "Aide", "Ouvre cette aide, sans fermer la liste."

End Sub

Private Sub Def_Contrats(ByRef titre As String, ByRef intro As String, ByRef champs As String, ByRef boutons As String)

    titre = "Liste des contrats de fourniture"
    intro = "Un contrat est un type de consommation chez un op{e2}rateur, entre deux dates. Il sert {a2} proposer l'op{e2}rateur et l'unit{e2} quand on saisit une facture. " & _
            "Pour changer d'op{e2}rateur, ajoutez simplement le nouveau contrat : s'il y a un contrat en cours pour ce type, le programme propose de le cl{o2}turer la veille. L'historique reste complet. " & _
            "Les cellules se corrigent directement ; deux contrats d'un m{ea}me type ne doivent pas se chevaucher."

    L champs, "ID_Contrat", "Num{e2}ro du contrat, attribu{e2} automatiquement."
    L champs, "TypeConso", "Type de consommation (eau, eau recycl{e2}e, {e2}lectricit{e2}, gaz)."
    L champs, "Operateur", "Nom du fournisseur."
    L champs, "Unite", "Unit{e2} de mesure (m3, kWh...)."
    L champs, "DateDebut", "Premier jour du contrat (JJ/MM/AAAA)."
    L champs, "DateFin", "Dernier jour du contrat ; vide = contrat en cours."
    L champs, "AbonnementMensuel", "Donn{e2} {a2} titre indicatif. L'abonnement r{e2}el est saisi sur les lignes d'abonnement de chaque facture."
    L champs, "Reference", "R{e2}f{e2}rence du contrat ou num{e2}ro de compteur (facultatif)."

    L boutons, "Ajouter un contrat", "Ouvre [f:frm_Contrat] pour saisir un contrat."
    L boutons, "Supprimer la s{e2}lection", "Supprime les lignes s{e2}lectionn{e2}es (les factures d{e2}j{a2} saisies sont conserv{e2}es)."
    L boutons, "Voir les factures", "Affiche [f:frm_Factures], la liste des factures."
    L boutons, "Retour", "Revient sur [f:frm_Consommations]."
    L boutons, "Aide", "Ouvre cette aide, sans fermer la liste."

End Sub

Private Sub Def_Facture(ByRef titre As String, ByRef intro As String, ByRef champs As String, ByRef boutons As String)

    titre = "Saisie d'une facture de consommation"
    intro = "Ce formulaire enregistre une facture telle qu'elle est imprim{e2}e par le fournisseur : un en-t{ea}te, puis des lignes d'ABONNEMENT, de CONSOMMATION et de TAXES. " & _
            "Chaque ligne a ses propres dates, qui ne co{i2}ncident pas forc{e2}ment avec celles des autres lignes. " & _
            "Les montants se calculent en direct ; le r{e2}capitulatif du bas est indicatif et n'emp{ea}che jamais l'enregistrement. " & _
            "Si un champ est incorrect, un message appara{i2}t sous le formulaire et rien n'est enregistr{e2}."

    L champs, "Type de consommation", "Liste : eau, eau recycl{e2}e, {e2}lectricit{e2}, gaz. Propose l'op{e2}rateur du contrat en vigueur."
    L champs, "Op{e2}rateur (fournisseur)", "Nom du fournisseur de la facture (obligatoire). Seuls les fournisseurs dont le contrat n'est pas termin{e2} pour le type choisi sont accept{e2}s : pour un nouveau fournisseur, saisissez d'abord un nouveau contrat (bouton Contrats de l'{e2}cran Consommations)."
    L champs, "Num{e2}ro de facture (facultatif)", "Num{e2}ro imprim{e2} sur la facture."
    L champs, "Date de la facture (facultative)", "Date JJ/MM/AAAA."
    L champs, "D{e2}but / fin de la p{e2}riode (facultatifs)", "P{e2}riode globale couverte par la facture."
    L champs, "Total TTC de la facture (facultatif)", "Montant TTC imprim{e2} ; compar{e2} au total calcul{e2} d'apr{e1}s les lignes."
    L champs, "Commentaire (facultatif)", "Texte libre."
    L champs, "Boutons + et -", "En haut {a2} droite de chaque bloc : + ajoute une ligne au bloc (les blocs suivants sont d{e2}cal{e2}s vers le bas), - retire la derni{e1}re ligne si elle est vide. Chaque bloc commence avec 3 lignes."
    L champs, "Abonnement : ligne", "Libell{e2}, d{e2}but, fin (incluse), quantit{e2} (nombre de jours ou de mois), unit{e2} (propos{e2}e : jours), prix HT par unit{e2} (n{e2}gatif pour un remboursement) et TVA. Les montants se calculent."
    L champs, "Unit{e2} (ligne de consommation)", "Obligatoirement m3 ou kWh pour un bloc d'indicateur Consommation (la liste de la cellule se limite {a2} ces deux valeurs ; m3 est converti automatiquement en m cube)."
    L champs, "Consommation : ligne", "Libell{e2}, d{e2}but, fin, quantit{e2} (kWh, m3...), unit{e2}, prix unitaire HT et TVA."
    L champs, "Taxes : ligne", "Libell{e2}, d{e2}but, fin, quantit{e2} (kWh ou euros de base), unit{e2}, prix HT et TVA."
    L champs, "Montant HT / Montant TTC", "Calcul{e2}s en direct : HT = quantit{e2} x prix, TTC = HT x (1 + TVA), arrondis au centime."
    L champs, "R{e2}capitulatif", "Totaux HT, TVA, TTC calcul{e2}s, total saisi et {e2}cart."

    L boutons, "Valider", "Enregistre la facture puis referme le formulaire."
    L boutons, "Ajouter un bloc", "Demande un nom, puis l'indicateur que le bloc alimente (liste num{e2}rot{e2}e ; 0 pour cr{e2}er un nouvel indicateur, qui sera propos{e2} dans les listes suivantes), et cr{e2}e sous les blocs existants un bloc de m{ea}me forme (m{ea}mes colonnes). Son titre rappelle l'indicateur entre parenth{e1}ses. Ses lignes sont enregistr{e2}es avec ce nom et cet indicateur et comptent dans les totaux de la facture. Neuf blocs au plus. Pour remettre un bloc du mod{e1}le supprim{e2}, tapez son nom : Abonnement, Consommation ou Taxes et contributions (l'indicateur n'est alors pas redemand{e2})."
    L boutons, "Indicateur (titre d'un bloc)", "Corrige l'indicateur d'un bloc affect{e2} par erreur (liste num{e2}rot{e2}e, 0 pour en cr{e2}er un). Si ce bloc figure dans d'autres factures, elles sont list{e2}es et une confirmation est demand{e2}e : le changement leur est appliqu{e2} quand vous validez la facture en cours (rien n'est modifi{e2} ailleurs si vous abandonnez). Un bloc ne peut pas passer {a2} Consommation si une facture l'utilise avec une unit{e2} autre que m3 ou kWh."
    L boutons, "+ / - / x (titre d'un bloc)", "+ ajoute une ligne au bloc (les blocs suivants descendent) ; - retire la derni{e1}re ligne si elle est vide ; x supprime le bloc, m{ea}me un bloc du mod{e1}le (confirmation s'il contient des lignes) : seule cette facture est concern{e2}e, le mod{e1}le des nouvelles factures ne change pas."
    L boutons, "Annuler", "Referme le formulaire sans rien enregistrer."
    L boutons, "Aide", "Ouvre cette aide, sans fermer le formulaire."

End Sub

Private Sub Def_Contrat(ByRef titre As String, ByRef intro As String, ByRef champs As String, ByRef boutons As String)

    titre = "Saisie d'un contrat de fourniture"
    intro = "Ce formulaire d{e2}clare un contrat : un type de consommation chez un op{e2}rateur, {a2} partir d'une date. " & _
            "CHANGEMENT D'OP{E2}RATEUR : cr{e2}ez simplement le nouveau contrat avec sa date de d{e2}but ; s'il y a un contrat en cours pour ce type, " & _
            "le programme propose de le cl{o2}turer la veille. Les factures d{e2}j{a2} saisies gardent leur ancien op{e2}rateur. " & _
            "Si un champ est incorrect, un message appara{i2}t sous le formulaire et rien n'est enregistr{e2}."

    L champs, "Type de consommation", "Liste : eau, eau recycl{e2}e, {e2}lectricit{e2}, gaz."
    L champs, "Op{e2}rateur (fournisseur)", "Nom du fournisseur."
    L champs, "Unit{e2} de mesure", "Propos{e2}e d'apr{e1}s le type (m3 pour l'eau, kWh pour l'{e2}lectricit{e2} et le gaz), modifiable."
    L champs, "D{e2}but du contrat", "Date JJ/MM/AAAA."
    L champs, "Fin du contrat (facultative)", "Date JJ/MM/AAAA ; laissez vide pour un contrat en cours."
    L champs, "Abonnement mensuel (euros)", "Donn{e2} {a2} titre indicatif ; l'abonnement r{e2}el est saisi sur chaque facture."
    L champs, "R{e2}f{e2}rence / num{e2}ro de compteur", "Texte libre (facultatif)."

    L boutons, "Valider", "Enregistre le contrat puis referme le formulaire."
    L boutons, "Annuler", "Referme le formulaire sans rien enregistrer."
    L boutons, "Aide", "Ouvre cette aide, sans fermer le formulaire."

End Sub


' =====================================================================================
' ECHEANCIER
' =====================================================================================
Private Sub Def_Echeancier(ByRef titre As String, ByRef intro As String, ByRef champs As String, ByRef boutons As String)

    titre = "{E2}ch{e2}ancier : op{e2}rations {a2} venir"
    intro = "Cette liste contient les op{e2}rations pr{e2}vues qui ne sont pas encore pass{e2}es sur le compte. Elles servent au calcul du solde pr{e2}visionnel de [f:Synthese]. " & _
            "Pendant l'import, apr{e1}s le contr{o2}le des cat{e2}gories et avant l'{e2}criture, chaque op{e2}ration qui correspond {a2} une {e2}ch{e2}ance vous est pr{e2}sent{e2}e avec cette {e2}ch{e2}ance (OUI : rapprocher ; NON : candidate suivante ; ANNULER : pas de rapprochement pour cette op{e2}ration) ; une fois l'import {e2}crit, l'{e2}ch{e2}ance rapproch{e2}e dispara{i2}t de la liste. " & _
            "Une {e2}ch{e2}ance non rapproch{e2}e reste ici, reste compt{e2}e dans le solde pr{e2}visionnel et, si sa date est pass{e2}e, dans '{E2}ch. en retard' de [f:Synthese] : c'est {a2} vous de la supprimer. " & _
            "Les cellules se corrigent directement. Les fl{e1}ches de filtre de l'en-t{ea}te de chaque colonne permettent de restreindre la liste affich{e2}e ; les lignes masqu{e2}es ne sont jamais supprim{e2}es par le bouton Supprimer la s{e2}lection."

    L champs, "ID_Echeance", "Num{e2}ro de l'{e2}ch{e2}ance, attribu{e2} automatiquement."
    L champs, "ID_Recurrence", "R{e1}gle d'origine (vide pour une op{e2}ration ponctuelle). Modifier une {e2}ch{e2}ance ne modifie pas sa r{e1}gle."
    L champs, "Libelle", "Texte libre (facultatif)."
    L champs, "Tiers", "Mot-cl{e2} cherch{e2} dans le Tiers de l'op{e2}ration import{e2}e, sans tenir compte des majuscules."
    L champs, "Categorie, SousCategorie", "Crit{e1}res de reconnaissance {a2} l'import (tous ceux renseign{e2}s doivent {ea}tre respect{e2}s)."
    L champs, "Montant", "N{e2}gatif pour une d{e2}pense, positif pour une entr{e2}e d'argent."
    L champs, "NumCheque", "Num{e2}ro du ch{e1}que attendu (vide si aucun). Contr{o2}le silencieux : cellule rouge si le num{e2}ro est invalide ou d{e2}j{a2} pass{e2} en banque, orange s'il figure deux fois dans cette liste."
    L champs, "DatePrevue", "Date pr{e2}vue de l'op{e2}ration (JJ/MM/AAAA)."
    L champs, "PrecisionDate", "Jour ou Mois. Mois : l'op{e2}ration peut arriver n'importe quel jour du mois, mais est compt{e2}e dans le solde pr{e2}visionnel d{e1}s le 1er du mois."
    L champs, "Periodicite", "Ponctuelle, Hebdomadaire, Mensuelle, Trimestrielle ou Annuelle."
    L champs, "TolerancePct", "{E2}cart de montant accept{e2} pour le rapprochement, en % du montant pr{e2}vu."

    L boutons, "Ajouter une {e2}ch{e2}ance", "Ouvre [f:frm_Echeance] pour saisir une op{e2}ration {a2} venir."
    L boutons, "Supprimer la s{e2}lection", "Supprime les lignes s{e2}lectionn{e2}es. Pour une {e2}ch{e2}ance issue d'une r{e1}gle, deux choix : Oui = supprimer seulement ces {e2}ch{e2}ances (la r{e1}gle continue) ; Non = arr{ea}ter la r{e2}currence {a2} partir de la premi{e1}re {e2}ch{e2}ance s{e2}lectionn{e2}e."
    L boutons, "Voir les r{e1}gles", "Affiche [f:frm_Recurrences], la liste des r{e1}gles de r{e2}currence."
    L boutons, "Aide", "Ouvre cette aide, sans fermer la liste."
    L boutons, "Sortir", "Ferme la liste (et cette aide si elle est ouverte) et revient sur [f:Synthese]."

End Sub

Private Sub Def_Recurrences(ByRef titre As String, ByRef intro As String, ByRef champs As String, ByRef boutons As String)

    titre = "R{e1}gles de r{e2}currence"
    intro = "Une r{e1}gle produit seule ses {e2}ch{e2}ances dans [f:frm_Echeancier] : celles des prochains jours, et toujours au moins une. " & _
            "La liste est compl{e2}t{e2}e {a2} l'ouverture du classeur, {a2} chaque import et {a2} l'ouverture de ces listes. " & _
            "Une r{e1}gle sans DateFin ne s'arr{ea}te jamais : supprimez-la ou mettez-la en pause. " & _
            "Une modification ne s'applique qu'aux {e2}ch{e2}ances cr{e2}{e2}es ensuite ; pour l'appliquer aussi {a2} celles d{e2}j{a2} dans la liste, cliquez sur [b:Actualiser les occurrences]."

    L champs, "ID_Recurrence", "Num{e2}ro de la r{e1}gle, attribu{e2} automatiquement."
    L champs, "Libelle, Tiers, Categorie, SousCategorie, Montant", "M{ea}mes significations que dans [f:frm_Echeancier] : texte libre, mot-cl{e2} du tiers, crit{e1}res de cat{e2}gorie, montant (n{e2}gatif = d{e2}pense)."
    L champs, "Periodicite", "Hebdomadaire, Mensuelle, Trimestrielle ou Annuelle."
    L champs, "DateProchaine", "Prochaine date {a2} cr{e2}er. N'y touchez que pour d{e2}caler la r{e1}gle."
    L champs, "DateFin", "Facultative. Apr{e1}s cette date la r{e1}gle ne cr{e2}e plus rien ; elle est retir{e2}e de la liste d{e1}s qu'il ne lui reste plus aucune {e2}ch{e2}ance."
    L champs, "PrecisionDate", "Jour ou Mois (voir [f:frm_Echeancier])."
    L champs, "JourAncre", "Jour du mois conserv{e2} d'une occurrence {a2} l'autre (un 31 devient le dernier jour des mois plus courts)."
    L champs, "TolerancePct", "{E2}cart de montant accept{e2} pour le rapprochement, en %."
    L champs, "Actif", "Oui ou Non. Non = r{e1}gle en pause : aucune {e2}ch{e2}ance cr{e2}{e2}e."

    L boutons, "Ajouter une {e2}ch{e2}ance", "Ouvre [f:frm_Echeance] (une p{e2}riodicit{e2} cr{e2}e une r{e1}gle ; sans p{e2}riodicit{e2}, l'op{e2}ration est unique)."
    L boutons, "Actualiser les occurrences", "Pour chaque r{e1}gle, supprime ses {e2}ch{e2}ances actuellement dans la liste (y compris celles non rapproch{e2}es) puis les recr{e2}e avec les valeurs de la r{e1}gle. ATTENTION : une r{e1}gle en pause perd ainsi ses {e2}ch{e2}ances."
    L boutons, "Supprimer la r{e1}gle", "Supprime les r{e1}gles s{e2}lectionn{e2}es ET leurs {e2}ch{e2}ances {a2} venir. Pour arr{ea}ter {a2} une date pr{e2}cise, renseignez plut{o2}t DateFin puis cliquez sur [b:Actualiser les occurrences]."
    L boutons, "Voir les {e2}ch{e2}ances", "Affiche [f:frm_Echeancier]."
    L boutons, "Aide", "Ouvre cette aide, sans fermer la liste."
    L boutons, "Sortir", "Ferme la liste (et cette aide si elle est ouverte) et revient sur [f:Synthese]."

End Sub

Private Sub Def_Echeance(ByRef titre As String, ByRef intro As String, ByRef champs As String, ByRef boutons As String)

    titre = "Saisie d'une {e2}ch{e2}ance"
    intro = "Ce formulaire d{e2}clare une op{e2}ration {a2} venir. Quand une op{e2}ration import{e2}e lui correspond, l'import propose de les rapprocher. " & _
            "RECONNAISSANCE : renseignez au moins le Tiers, la Cat{e2}gorie ou le Num{e2}ro de ch{e1}que. Si plusieurs crit{e1}res sont renseign{e2}s, l'op{e2}ration doit tous les respecter. " & _
            "Si un champ est incorrect, un message appara{i2}t sous le formulaire et rien n'est enregistr{e2}."

    L champs, "Libell{e2} (facultatif)", "Texte libre."
    L champs, "Tiers (contient)", "Choisi dans la liste des Tiers d{e2}j{a2} connus ou tap{e2} librement. Cherch{e2} dans le Tiers de l'op{e2}ration import{e2}e, sans tenir compte des majuscules : un mot-cl{e2} suffit (ex. EDF)."
    L champs, "Cat{e2}gorie", "Liste. Utilisez [b:+] si la cat{e2}gorie manque."
    L champs, "Sous-cat{e2}gorie", "Liste d{e2}pendant de la cat{e2}gorie (sans cat{e2}gorie, toutes sont propos{e2}es ; choisir une sous-cat{e2}gorie propre {a2} une cat{e2}gorie renseigne celle-ci)."
    L champs, "Montant (- = d{e2}pense)", "N{e2}gatif pour une d{e2}pense, positif pour une entr{e2}e d'argent ; une op{e2}ration de signe contraire n'est jamais rapproch{e2}e."
    L champs, "Num{e2}ro de ch{e1}que (facultatif)", "Pour une op{e2}ration unique seulement. Chiffres uniquement (10 au maximum), refus{e2} s'il est d{e2}j{a2} pass{e2} en banque ou d{e2}j{a2} saisi. Avec un num{e2}ro, l'op{e2}ration est reconnue par ce seul num{e2}ro (le signe du montant doit concorder)."
    L champs, "P{e2}riodicit{e2}", "Vide ou Ponctuelle : op{e2}ration unique. Hebdomadaire, Mensuelle, Trimestrielle ou Annuelle : cr{e2}e une r{e1}gle qui produit seule les {e2}ch{e2}ances suivantes."
    L champs, "Date pr{e2}vue (ou 1{e1}re date)", "Date JJ/MM/AAAA de l'op{e2}ration, ou de la premi{e1}re pour une r{e1}gle."
    L champs, "Date de fin (facultative)", "R{e2}serv{e2}e aux op{e2}rations r{e2}currentes. Vide : la r{e1}gle ne s'arr{ea}te jamais."
    L champs, "Pr{e2}cision de la date", "Jour : l'op{e2}ration est reconnue autour de la date pr{e2}vue. Mois : n'importe quel jour du mois de la date pr{e2}vue."
    L champs, "Tol{e2}rance sur le montant (%)", "{E2}cart de montant accept{e2}, en % du montant pr{e2}vu ; 0 = montant exact ; valeur par d{e2}faut si vide."

    L boutons, "+ (Cat{e2}gorie, Sous-cat{e2}gorie)", "Ouvre [f:frm_NouvelleCategorie] pour cr{e2}er une cat{e2}gorie ou une sous-cat{e2}gorie absente des listes ; une fois valid{e2}e, elle est reprise ici."
    L boutons, "Valider", "Enregistre l'{e2}ch{e2}ance (ou la r{e1}gle) puis referme le formulaire."
    L boutons, "Annuler", "Referme le formulaire sans rien enregistrer."
    L boutons, "Aide", "Ouvre cette aide, sans fermer le formulaire."

End Sub


' =====================================================================================
' IMPORT : contrôle, ventilation, résolution, notes, santé
' =====================================================================================
Private Sub Def_ControleCategories(ByRef titre As String, ByRef intro As String, ByRef champs As String, ByRef boutons As String)

    titre = "Contr{o2}le des cat{e2}gories"
    intro = "Ce formulaire contr{o2}le la cat{e2}gorie et la sous-cat{e2}gorie des op{e2}rations qui viennent d'{ea}tre import{e2}es. " & _
            "Les op{e2}rations sont pr{e2}sent{e2}es une par une : v{e2}rifiez les informations, corrigez si besoin la cat{e2}gorie puis la sous-cat{e2}gorie, puis passez {a2} la suivante. " & _
            "Si des op{e2}rations n'ont pas pu {ea}tre reli{e2}es {a2} une sous-cat{e2}gorie interne, vous {ea}tes repositionn{e2} sur la premi{e1}re d'entre elles. " & _
            "ATTENTION : vos choix ne sont pris en compte qu'apr{e1}s un clic sur [b:Enregistrer]."

    L champs, "Compteur", "Num{e2}ro de l'op{e2}ration affich{e2}e sur le total {a2} contr{o2}ler."
    L champs, "Date", "Lecture seule : date de l'op{e2}ration."
    L champs, "Tiers", "Modifiable, sauf pour une op{e2}ration de sant{e2}."
    L champs, "Libell{e2} / Notes", "Modifiable, sauf pour une op{e2}ration de sant{e2}."
    L champs, "Montant", "Lecture seule : montant de l'op{e2}ration."
    L champs, "Ch{e1}que", "Lecture seule : num{e2}ro de ch{e1}que, s'il y en a un."
    L champs, "Cat{e2}gorie source (banque)", "Lecture seule : cat{e2}gorie fournie par la banque."
    L champs, "Cat{e2}gorie", "Liste d{e2}roulante : la cat{e2}gorie retenue. Utilisez [b:+] si elle manque."
    L champs, "Sous-cat{e2}gorie", "Liste d{e2}roulante d{e2}pendant de la cat{e2}gorie."
    LPointe champs

    L boutons, "< Pr{e2}c{e2}dent  /  Suivant >", "Naviguent entre les op{e2}rations ; les choix d{e2}j{a2} faits sont conserv{e2}s."
    L boutons, "Ventiler", "R{e2}partit le montant de l'op{e2}ration entre plusieurs sous-cat{e2}gories (ouvre [f:frm_Ventilation])."
    L boutons, "+", "Cr{e2}e une nouvelle cat{e2}gorie ou sous-cat{e2}gorie (ouvre [f:frm_NouvelleCategorie])."
    L boutons, "Enregistrer", "Applique tous les choix faits et poursuit l'import."
    L boutons, "Annuler", "Abandonne l'import en cours."
    L boutons, "Aide", "Ouvre cette aide, sans fermer le formulaire."

End Sub

Private Sub Def_Ventilation(ByRef titre As String, ByRef intro As String, ByRef champs As String, ByRef boutons As String)

    titre = "Ventilation d'une op{e2}ration"
    intro = "Ce formulaire ventile une op{e2}ration bancaire entre plusieurs cat{e2}gories et sous-cat{e2}gories. " & _
            "Le haut rappelle l'op{e2}ration et le calcul ; le tableau liste les lignes d{e2}j{a2} ajout{e2}es ; la zone du bas sert {a2} saisir une nouvelle ligne. " & _
            "La ventilation ne peut {ea}tre enregistr{e2}e que lorsque le reste {a2} ventiler est exactement nul. " & _
            "ATTENTION : une ligne retir{e2}e du tableau avec [b:{crayon}] est perdue si vous ne cliquez pas ensuite sur [b:Ajouter la ligne]."

    L champs, "Date, Tiers, Libell{e2} / Notes, Montant, Ch{e1}que", "Lecture seule : rappel de l'op{e2}ration {a2} ventiler."
    L champs, "Cat{e2}gorie actuelle", "Lecture seule : cat{e2}gorie de l'op{e2}ration avant ventilation."
    LPointe champs
    L champs, "Montant {a2} ventiler", "Lecture seule : montant de l'op{e2}ration."
    L champs, "Total saisi", "Lecture seule : somme des lignes du tableau."
    L champs, "Reste {a2} ventiler", "Lecture seule : montant {a2} ventiler moins total saisi. Doit valoir 0 pour enregistrer."
    L champs, "Cat{e2}gorie (zone de saisie)", "Liste. Utilisez [b:+] si elle manque."
    L champs, "Sous-cat{e2}gorie (zone de saisie)", "Liste d{e2}pendant de la cat{e2}gorie."
    L champs, "Montant (zone de saisie)", "Montant de la ligne, de m{ea}me signe que l'op{e2}ration."
    L champs, "Notes (zone de saisie)", "Texte libre (facultatif) propre {a2} la ligne."

    L boutons, "Enregistrer", "Enregistre toutes les lignes et referme le formulaire."
    L boutons, "Annuler", "Abandonne la saisie en cours sans rien enregistrer."
    L boutons, "Supprimer la ventilation", "Visible seulement si l'op{e2}ration est d{e2}j{a2} ventil{e2}e : efface toutes ses lignes et rend {a2} l'op{e2}ration sa cat{e2}gorie d'origine."
    L boutons, "{crayon}  (en face de chaque ligne)", "Retire la ligne du tableau et la recopie dans la zone de saisie pour la corriger."
    L boutons, "+", "Cr{e2}e une nouvelle cat{e2}gorie ou sous-cat{e2}gorie (ouvre [f:frm_NouvelleCategorie])."
    L boutons, "Ajouter la ligne", "Ajoute la ligne saisie au tableau (ou remplace la ligne en cours de correction)."
    L boutons, "Effacer la saisie", "Vide la zone de saisie."
    L boutons, "Aide", "Ouvre cette aide, sans fermer le formulaire."

End Sub

Private Sub Def_NouvelleCategorie(ByRef titre As String, ByRef intro As String, ByRef champs As String, ByRef boutons As String)

    titre = "Nouvelle cat{e2}gorie ou sous-cat{e2}gorie"
    intro = "Ce formulaire cr{e2}e une nouvelle cat{e2}gorie, ou une nouvelle sous-cat{e2}gorie d'une cat{e2}gorie existante. " & _
            "Il s'ouvre avec le bouton [b:+] de [f:frm_ControleCategories], de [f:frm_Ventilation] ou de [f:frm_Echeance]."

    L champs, "Cat{e2}gorie", "Choisissez une cat{e2}gorie existante dans la liste, ou tapez-en une nouvelle."
    L champs, "Sous-cat{e2}gorie", "Choisissez une sous-cat{e2}gorie d{e2}j{a2} connue de cette cat{e2}gorie, ou tapez-en une nouvelle. Ce champ peut rester vide."

    L boutons, "Valider", "Enregistre la cat{e2}gorie et la sous-cat{e2}gorie saisies, puis referme le formulaire."
    L boutons, "Annuler", "Referme le formulaire sans rien cr{e2}er."
    L boutons, "Aide", "Ouvre cette aide, sans fermer le formulaire."

End Sub

Private Sub Def_ResolutionCategories(ByRef titre As String, ByRef intro As String, ByRef champs As String, ByRef boutons As String)

    titre = "R{e2}solution des cat{e2}gories"
    intro = "Les op{e2}rations list{e2}es n'ont pas pu {ea}tre cat{e2}goris{e2}es automatiquement : ce formulaire permet de le faire manuellement. " & _
            "Pour chacune, choisissez une cat{e2}gorie dans la liste d{e2}roulante. Le compteur indique le nombre d'op{e2}rations d{e2}j{a2} trait{e2}es sur le total."

    L champs, "Statut", "Lecture seule : passe automatiquement {a2} une coche verte d{e1}s que la cat{e2}gorie est choisie."
    L champs, "Date", "Lecture seule : date de l'op{e2}ration."
    L champs, "Montant", "Lecture seule : montant de l'op{e2}ration."
    L champs, "Tiers", "Lecture seule : tiers de l'op{e2}ration."
    L champs, "Cat{e2}gorie(s)", "Liste d{e2}roulante (cellule jaune) : la cat{e2}gorie {a2} retenir pour cette op{e2}ration."

    L boutons, "R{e2}initialiser cette ligne", "Efface le choix de la ligne s{e2}lectionn{e2}e pour recommencer."
    L boutons, "Enregistrer", "Applique les choix faits et poursuit l'import. Les op{e2}rations sans cat{e2}gorie resteront sans cat{e2}gorie (une confirmation est demand{e2}e)."
    L boutons, "Aide", "Ouvre cette aide, sans fermer le formulaire."

End Sub

Private Sub Def_RapprochementNotes(ByRef titre As String, ByRef intro As String, ByRef champs As String, ByRef boutons As String)

    titre = "Rapprochement des notes de sant{e2}"
    intro = "Cette op{e2}ration de sant{e2} n'a pas pu {ea}tre rapproch{e2}e automatiquement car son champ Notes n'a pas {e2}t{e2} correctement renseign{e2}. " & _
            "Ce formulaire v{e2}rifie si une autre op{e2}ration est d{e2}j{a2} enregistr{e2}e pour cette consultation. " & _
            "S{e2}lectionnez la date de consultation dans la liste ; si elle est absente, cliquez sur [b:Pas de correspondance]. " & _
            "Sinon continuez {a2} descendre (sp{e2}cialit{e2}, b{e2}n{e2}ficiaire, montant) pour retrouver la cl{e2}, qui s'affiche dans Cl{e2} trouv{e2}e."

    L champs, "Date op{e2}ration, Num. ch{e1}que, Montant", "Lecture seule : rappel de l'op{e2}ration bancaire."
    L champs, "Cat{e2}gorie, Sous-cat{e2}gorie", "Lecture seule : cat{e2}gorie actuelle de l'op{e2}ration (modifiable avec [b:Changer la cat{e2}gorie])."
    L champs, "Notes, Tiers", "Lecture seule : notes et tiers de l'op{e2}ration."
    LPointe champs
    L champs, "Date consultation", "Liste des dates de consultation connues : premier crit{e1}re de recherche de la cl{e2}."
    L champs, "Sp{e2}cialit{e2}", "Liste, filtr{e2}e d'apr{e1}s la date choisie."
    L champs, "B{e2}n{e2}ficiaire", "Liste, filtr{e2}e d'apr{e1}s la date et la sp{e2}cialit{e2}."
    L champs, "Montant (recherche)", "Liste, filtr{e2}e d'apr{e1}s les choix pr{e2}c{e2}dents."
    L champs, "Cl{e2} trouv{e2}e", "Lecture seule : cl{e2} de rapprochement retrouv{e2}e parmi celles d{e2}j{a2} connues."

    L boutons, "Valider", "Applique la cl{e2} trouv{e2}e {a2} cette op{e2}ration. ATTENTION : elle ne sera PLUS JAMAIS repropos{e2}e (rapprochement consid{e2}r{e2} r{e2}gl{e2} d{e2}finitivement)."
    L boutons, "Pas de correspondance", "Ouvre [f:frm_GenerationCle] pour cr{e2}er une toute nouvelle cl{e2} (aucune cl{e2} existante ne convient)."
    L boutons, "Passer", "Ne modifie RIEN sur cette op{e2}ration ; elle sera repropos{e2}e la prochaine fois que la v{e2}rification sera relanc{e2}e."
    L boutons, "Sortir", "Ferme le formulaire (et cette aide si elle est ouverte) et revient sur [f:Synthese]. Le traitement peut {ea}tre relanc{e2} avec le bouton Retraiter suivi de sant{e2}."
    L boutons, "Changer la cat{e2}gorie", "Pour une op{e2}ration qui n'est pas une d{e2}pense de sant{e2} (erreur {a2} l'import) : corrige sa cat{e2}gorie ; elle sort de l'analyse sant{e2} et ses colonnes sant{e2} sont vid{e2}es. ATTENTION, pour une ligne de ventilation, TOUTE la ventilation de l'op{e2}ration est rouverte puis r{e2}{e2}crite : les rapprochements sant{e2} d{e2}j{a2} faits sur ses autres lignes sont perdus."
    L boutons, "Aide", "Ouvre cette aide, sans fermer le formulaire."

End Sub

Private Sub Def_GenerationCle(ByRef titre As String, ByRef intro As String, ByRef champs As String, ByRef boutons As String)

    titre = "G{e2}n{e2}ration d'une cl{e2} de rapprochement"
    intro = "Aucune cl{e2} existante ne correspond {a2} cette op{e2}ration : ce formulaire cr{e2}e une nouvelle cl{e2} de la forme AAAAMMJJ;Sp{e2}cialit{e2};B{e2}n{e2}ficiaire;Montant. " & _
            "Renseignez la date de consultation, la sp{e2}cialit{e2} et le b{e2}n{e2}ficiaire (le montant est repris de l'op{e2}ration), cliquez sur [b:G{e2}n{e2}rer], v{e2}rifiez la cl{e2}, puis cliquez sur [b:Valider]. " & _
            "Pour une d{e2}pense ventil{e2}e, deux blocs sont affich{e2}s : l'op{e2}ration parente (op{e2}ration bancaire d'origine) et l'op{e2}ration ventil{e2}e (la part concern{e2}e)."

    L champs, "Date, Tiers, Num. ch{e1}que, Notes", "Lecture seule : rappel de l'op{e2}ration (et, pour une ventilation, de l'op{e2}ration parente)."
    LPointe champs
    L champs, "Date consultation", "Date JJ/MM/AAAA de la consultation."
    L champs, "Sp{e2}cialit{e2}", "Liste. Utilisez [b:+] si elle manque."
    L champs, "B{e2}n{e2}ficiaire", "Liste. Utilisez [b:+] si elle manque."
    L champs, "Montant", "Lecture seule : repris de l'op{e2}ration."
    L champs, "Cl{e2} g{e2}n{e2}r{e2}e", "Lecture seule : cl{e2} construite par [b:G{e2}n{e2}rer], aussi copi{e2}e dans le presse-papier."

    L boutons, "G{e2}n{e2}rer", "Construit la cl{e2} {a2} partir des champs saisis."
    L boutons, "Valider", "{E2}crit la cl{e2} g{e2}n{e2}r{e2}e dans l'op{e2}ration. ATTENTION : l'{e2}criture est D{E2}FINITIVE, m{ea}me si la cl{e2} est approximative ou incorrecte : l'op{e2}ration ne sera plus jamais repropos{e2}e (seule une modification manuelle dans [f:Import_Data] la ferait r{e2}appara{i2}tre)."
    L boutons, "Passer", "Laisse cette op{e2}ration de c{o2}t{e2} pour y revenir plus tard."
    L boutons, "Retour", "Revient sur [f:frm_RapprochementNotes] pour la m{ea}me op{e2}ration, sans rien modifier."
    L boutons, "+ (Sp{e2}cialit{e2}, B{e2}n{e2}ficiaire)", "Ajoute une nouvelle valeur {a2} la liste."
    L boutons, "Aide", "Ouvre cette aide, sans fermer le formulaire."

End Sub

Private Sub Def_AffectationFacture(ByRef titre As String, ByRef intro As String, ByRef champs As String, ByRef boutons As String)

    titre = "Affectation d'une facture {a2} une op{e2}ration"
    intro = "Cet {e2}cran relie l'op{e2}ration bancaire affich{e2}e en haut {a2} la facture qui l'a r{e2}gl{e2}e (Eau, {E2}lectricit{e2} - gaz - chauffage). " & _
            "Choisissez le type de consommation, puis la facture dans la liste, la plus r{e2}cente en premier, et cliquez sur [b:Valider]. " & _
            "Une m{ea}me facture peut {ea}tre li{e2}e {a2} plusieurs op{e2}rations (r{e2}glement en plusieurs {e2}ch{e2}ances) : les factures d{e2}j{a2} li{e2}es restent donc propos{e2}es. " & _
            "Une fois li{e2}e, l'op{e2}ration quitte le tableau Factures non li{e2}es de la feuille [f:Synthese]."

    L champs, "Date op{e2}ration, Tiers, Cat{e2}gorie, Montant", "Lecture seule : rappel de l'op{e2}ration {a2} affecter."
    L champs, "Type de consommation", "Liste : premier crit{e1}re, les types des factures enregistr{e2}es."
    L champs, "Date de la facture", "Liste des factures du type choisi, la plus r{e2}cente en premier : date, op{e2}rateur, num{e2}ro et total TTC."
    L champs, "Facture retenue", "Lecture seule : rappel de la facture choisie et du nombre d'op{e2}rations qui lui sont d{e2}j{a2} li{e2}es."

    L boutons, "Valider", "{E2}crit le lien entre l'op{e2}ration et la facture retenue, puis revient sur [f:Synthese]."
    L boutons, "Nouvelle facture", "Aucune facture de la liste ne convient : ouvre la saisie d'une nouvelle facture, d{e2}j{a2} li{e2}e {a2} l'op{e2}ration. Si vous l'abandonnez, ce formulaire r{e2}appara{i2}t."
    L boutons, "Annuler", "Ferme le formulaire sans rien modifier et revient sur [f:Synthese]."
    L boutons, "Aide", "Ouvre cette aide, sans fermer le formulaire."

End Sub

Private Sub Def_SuiviSante(ByRef titre As String, ByRef intro As String, ByRef champs As String, ByRef boutons As String)

    titre = "Suivi sant{e2} : compl{e2}ter une d{e2}pense"
    intro = "Cette d{e2}pense de sant{e2} n'est pas encore sold{e2}e ({e2}quilibre non atteint) : ce formulaire permet de la compl{e2}ter. " & _
            "Les cas sont pr{e2}sent{e2}s un par un. V{e2}rifiez le bloc D{e2}pense {a2} traiter et les remboursements li{e2}s, puis renseignez les champs demand{e2}s, de haut en bas."

    L champs, "Date op{e2}ration, Tiers import{e2}, Montant", "Lecture seule : la d{e2}pense {a2} traiter."
    L champs, "Num_Cheque, Date_consult, Sp{e2}cialit{e2} consult{e2}e, Notes", "Lecture seule : informations de la d{e2}pense."
    L champs, "Remb. 1 - Date / Montant, Remb. 2 - Date / Montant", "Lecture seule : les remboursements d{e2}j{a2} li{e2}s {a2} la d{e2}pense."
    L champs, "Solde actuel", "Lecture seule : d{e2}pense moins remboursements. Nul = sold{e2}e."
    L champs, "B{e2}n{e2}ficiaire", "Liste : la personne concern{e2}e. Utilisez [b:+] si elle manque."
    L champs, "Tiers corrig{e2}", "Liste : {a2} renseigner si le tiers import{e2} n'est pas correct. Utilisez [b:+] si le praticien manque."
    L champs, "Franchise", "Montant retenu par l'assurance (0 si aucune)."
    L champs, "D{e2}passement valid{e2} ?", "Accessible d{e1}s que le solde n'est pas nul. Oui : vous consid{e2}rez ce solde comme d{e2}finitif et normal ({a2} expliquer dans le commentaire) ; la d{e2}pense ne sera plus jamais repropos{e2}e. Vide ou Non : la d{e2}pense sera repropos{e2}e au prochain passage (un remboursement peut encore arriver)."
    L champs, "Commentaire", "Texte libre (facultatif)."
    LPointe champs

    L boutons, "Valider", "Contr{o2}le la saisie, l'enregistre et passe au cas suivant."
    L boutons, "Passer", "Passe au cas suivant sans rien enregistrer ; la d{e2}pense sera repropos{e2}e au prochain traitement."
    L boutons, "Sortir", "Ferme le formulaire (et cette aide si elle est ouverte) et revient sur [f:Synthese]. Le traitement peut {ea}tre relanc{e2} avec le bouton Retraiter suivi de sant{e2}."
    L boutons, "+ (B{e2}n{e2}ficiaire, Tiers corrig{e2})", "Ajoute une nouvelle valeur {a2} la liste."
    L boutons, "Aide", "Ouvre cette aide, sans fermer le formulaire."

End Sub


' =====================================================================================
' RECHERCHE D'OPERATIONS
' =====================================================================================
Private Sub Def_RechercheOperations(ByRef titre As String, ByRef intro As String, ByRef champs As String, ByRef boutons As String)

    titre = "Recherche et correction des op{e2}rations"
    intro = "Cet {e2}cran recherche et corrige les op{e2}rations enregistr{e2}es, ainsi que leurs lignes de ventilation. Le tableau est charg{e2} {a2} l'ouverture " & _
            "(recherche globale, dernier import, mois budg{e2}taire, suivi sant{e2}...) ; la ligne du dessus rappelle le filtre actif et le nombre de lignes affich{e2}es. " & _
            "Utilisez les fl{e1}ches de filtre de l'en-t{ea}te de chaque colonne pour restreindre la liste. " & _
            "Double-clic sur une cellule Categorie ou SousCategorie : [f:frm_ControleCategories] s'ouvre pour l'{e2}diter. Double-clic sur la colonne Ventile : [f:frm_Ventilation] s'ouvre (nouvelle ventilation si vide, revue ou suppression si Oui)."

    L champs, "Date, Tiers, Montant", "Lecture seule : informations de l'op{e2}ration."
    L champs, "Categorie, SousCategorie", "Modifiables par double-clic (formulaire de contr{o2}le des cat{e2}gories) ; enregistrement imm{e2}diat."
    L champs, "Notes", "Modifiable directement dans la cellule ; enregistr{e2} d{e1}s la sortie de la cellule, avec contr{o2}le de la saisie (une note dont la cl{e2} sant{e2} est d{e2}j{a2} valide est verrouill{e2}e)."
    L champs, "Ventile", "Oui si l'op{e2}ration est ventil{e2}e. Double-clic pour d{e2}marrer ou revoir une ventilation."
    L champs, "Pointe", "Oui si l'op{e2}ration est d{e2}j{a2} analys{e2}e (liste). Enregistr{e2} tout de suite ; retirer un pointage demande confirmation. Une part de ventilation suit le pointage de son op{e2}ration. Filtrez la colonne pour ne garder que les op{e2}rations {a2} analyser."
    L champs, "Num_Cheque, Budget, Date_consult, Spe_consult, StatutSante, SoldeSante", "Colonnes d'information, affich{e2}es selon le filtre demand{e2} (budget, suivi sant{e2})."

    L boutons, "Recherche globale", "Recharge toutes les op{e2}rations, sans filtre."
    L boutons, "Sortir", "Ferme l'{e2}cran (et cette aide si elle est ouverte) et revient sur l'{e2}cran d'origine ([f:Synthese], bilan mensuel ou analyse)."
    L boutons, "Ajouter un d{e2}calage", "Ajoute un d{e2}calage de budget r{e2}current."
    L boutons, "D{e2}caler cette op{e2}ration", "D{e2}calage ponctuel d'une op{e2}ration ; le calcul repart toujours de la date de l'op{e2}ration (pour revenir au mois d'origine, indiquez 0)."
    L boutons, "Aide", "Ouvre cette aide, sans fermer l'{e2}cran."
    L boutons, "Rendre r{e2}currente", "Ouvre [f:frm_Echeance] pr{e2}rempli avec l'op{e2}ration s{e2}lectionn{e2}e (tiers, cat{e2}gorie, montant, p{e2}riodicit{e2} mensuelle, prochaine date). Vous pouvez tout modifier avant de valider ; l'op{e2}ration elle-m{ea}me n'est pas modifi{e2}e."
    L boutons, "Affecter une facture", "Relie l'op{e2}ration s{e2}lectionn{e2}e {a2} la facture qui l'a r{e2}gl{e2}e : ouvre [f:frm_AffectationFacture] (type de consommation, puis facture). Une part de ventilation ou une op{e2}ration d{e2}j{a2} li{e2}e est refus{e2}e. Ouvert par [b:Afficher tout] de la feuille [f:Synthese], l'{e2}cran ne montre que les factures non li{e2}es, sans limite de lignes."

End Sub


' =====================================================================================
' REMEDIATION DEPUIS UN OFX
' =====================================================================================
Private Sub Def_RemediationOFX(ByRef titre As String, ByRef intro As String, ByRef champs As String, ByRef boutons As String)

    titre = "Rem{e2}diation depuis un OFX"
    intro = "Ce formulaire compare un fichier OFX r{e2}cent aux op{e2}rations d{e2}j{a2} enregistr{e2}es (par identifiant bancaire) et affiche les {e2}carts AVANT toute correction. " & _
            "Il sert aussi aux contr{o2}les p{e2}riodiques : exportez un OFX frais, cliquez sur [b:Analyser un OFX], relisez le tableau, ajustez la colonne [c:Corriger], puis cliquez sur [b:Appliquer]. " & _
            "Rien n'est modifi{e2} pendant l'analyse. " & _
            "Les champs d'identit{e2} (dates, type, montant, num{e2}ro de ch{e1}que) viennent de la banque : l'OFX fait foi, ils sont pr{e2}-coch{e2}s Oui. " & _
            "Tiers, Notes et Cat{e2}gorie ne sont pr{e2}-coch{e2}s que pour une ligne manifestement m{e2}lang{e2}e (Tiers diff{e2}rent de l'OFX ET Notes reprenant le texte bancaire d'une autre op{e2}ration) ; " & _
            "sinon c'est Non, car l'op{e2}rateur a pu les corriger volontairement. " & _
            "Les op{e2}rations de l'OFX absentes du classeur sont seulement signal{e2}es : pour les ajouter, relancez l'import avec ce m{ea}me OFX. " & _
            "ATTENTION : le tableau affiche des donn{e2}es bancaires et reste dans le classeur (utile pour annuler) ; l'OFX peut {ea}tre effac{e2} apr{e1}s l'application."

    L champs, "Corriger", "Cellule jaune : Oui = la correction sera {e2}crite par Appliquer ; Non = ignor{e2}e ; - = information seulement."
    L champs, "Statut", "Lecture seule : vide avant application, puis Appliqu{e2}, Ignor{e2} (la valeur a chang{e2} depuis l'analyse, ou d{e2}calage manuel pour le budget) ou Annul{e2}."
    L champs, "ID_Transaction", "Lecture seule : identifiant bancaire de l'op{e2}ration. Les corrections sont retrouv{e2}es par cet identifiant."
    L champs, "Date", "Lecture seule : date comptable de l'op{e2}ration dans le classeur."
    L champs, "Tiers (actuel)", "Lecture seule : Tiers enregistr{e2} dans le classeur."
    L champs, "Champ", "Lecture seule : la donn{e2}e concern{e2}e par l'{e2}cart."
    L champs, "Valeur actuelle", "Lecture seule : valeur enregistr{e2}e dans le classeur."
    L champs, "Valeur OFX / propos{e2}e", "Lecture seule : valeur de l'OFX, ou valeur recalcul{e2}e (budget, cat{e2}gorie)."
    L champs, "Nature", "Lecture seule : Identit{e2} (l'OFX fait foi), Tiers, Notes d'une autre op{e2}ration, Cat{e2}gorie, Budget, Pointage ou Information."
    L champs, "Commentaire", "Lecture seule : explication ; pour une ligne m{e2}lang{e2}e, l'op{e2}ration dont semblent provenir ses valeurs."

    L boutons, "Analyser un OFX", "Choisir l'OFX {a2} comparer puis remplir le tableau des {e2}carts. Aucune donn{e2}e n'est modifi{e2}e."
    L boutons, "Appliquer", "{E2}crit les corrections dont [c:Corriger] vaut Oui et dont le [c:Statut] est vide. Une copie pr{e2}alable du classeur est recommand{e2}e ; une confirmation est demand{e2}e."
    L boutons, "Annuler les corrections", "Restaure les valeurs d'avant pour les lignes {a2} [c:Statut] Appliqu{e2}."
    L boutons, "Sortir", "Revient sur [f:Synthese]."
    L boutons, "Aide", "Ouvre cette aide, sans fermer le formulaire."

End Sub
