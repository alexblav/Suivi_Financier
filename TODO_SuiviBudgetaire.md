# Module de suivi budgétaire – proposition et TODO

Créé le 10/10/2026 · branche `Refonte_recherche`

## 1. Ce que font les autres (benchmark)

| Outil | Analyses marquantes | À garder ? |
| --- | --- | --- |
| Monarch / Mint | Répartition des dépenses par catégorie, « bilan du mois » automatique (top catégories), évolution mois par mois, budget « flex » (fixe / variable / non mensuel) | **Oui** : répartition, évolution, top catégories. Fixe/variable : plus tard |
| YNAB | Budget par catégorie, réel vs prévu | **Oui, en phase 2** (budget par catégorie) |
| Bankin' / Finary | Camembert par catégorie, comparaison avec le mois précédent, dépenses récurrentes détectées | Oui : comparaison N-1 / mois précédent. Récurrentes : déjà couvert par l'échéancier |
| Tableaux de bord « budget vs réel » | Écart = réel − budget (+ en %), moyenne glissante 3 mois, alertes de seuil, drill-down résumé → détail | **Oui** : écart, moyenne 3 mois, drill-down vers les opérations |
| Sunburst / treemap | Hiérarchie catégorie → sous-catégorie | Sunburst Excel natif possible, mais peu lisible et non cliquable en VBA : **barres + anneau** à la place (voir décisions) |

Sources : [Monarch (Experian)](https://experian.com/blogs/ask-experian/monarch-money-review/), [meilleures applis budget (Finary)](https://finary.com/fr/blog/finance-perso/budget/meilleures-applications-budget), [Budget vs Actual dashboard](https://www.clearpointstrategy.com/dashboards/budget-vs-actual-dashboard), [bonnes pratiques de dashboard](https://adminlte.io/blog/25-best-dashboard-design-examples-for-2026-real-products-principles/).

## 2. Les trois demandes + extras retenus

1. **Répartition des dépenses** par catégorie, puis sous-catégorie (clic/sélection d'une catégorie = zoom sur ses sous-catégories).
2. **Lien vers `frm_RechercheOperations`** : un nouveau préfiltre `Analyse` affiche les opérations derrière une barre/ligne du graphique, avec toutes les fonctions de l'écran (recatégoriser, ventiler, décaler…).
3. **Évolution dans le temps** par catégorie (courbe ou barres empilées par mois de budget).
4. Extras proposés (à valider) : comparaison avec la période précédente, top 5 catégories en hausse/baisse, moyenne glissante 3 mois, puis budget cible par catégorie avec écart.

## 3. Décisions de conception (recommandées)

- **Nouvelle feuille-écran `frm_Analyse`**, même principe que les autres formulaires-feuilles (cachée, bouton Sortir, Synthese masquée pendant l'affichage). Un module `mod_Analyse` (calculs, sans dessin) + `mod_InstallAnalyse` (construction de la feuille), comme pour l'échéancier.
- **Source des données** : TblOperations + TblVentilations. Une opération « Ventilé » est **remplacée par ses parts** (sinon double comptage et catégorie « Ventilé » dans le graphique). Le bilan mensuel actuel, lui, ne compte que TblOperations : à garder en tête.
- **Période** : on réutilise `critAnnee` / `critMois` de Synthese (mois de budget, donc les décalages sont respectés) + option « année entière » et « 12 derniers mois ».
- **Exclusions** : catégories « Virements internes » et revenus hors des dépenses (liste d'exclusion modifiable dans Param, pas codée en dur).
- **Signe** : dépenses = montants négatifs, affichées en positif dans les graphiques.
- **Graphiques** : graphiques Excel natifs créés par VBA (anneau + barres horizontales triées, courbes/barres empilées pour le temps), alimentés par une zone de calcul de la feuille. Le sunburst natif est écarté (pas d'événement de clic exploitable en VBA, rendu limité).
- **Drill-down** : un tableau à côté du graphique (catégorie, montant, %, nb d'opérations). Double-clic sur une ligne = zoom sur les sous-catégories ; bouton « Voir les opérations » = ouverture de la recherche. Plus fiable que le clic sur le graphique.
- **Préfiltre `Analyse`** dans `RechercherOperations` : paramètre = catégorie, sous-catégorie, période (le mécanisme `prefiltre` + `param` existe déjà). Il faudra aussi inclure les parts ventilées (`inclureVentilations`).

## 4. TODO (un module à la fois)

> **État 10/10/2026 : V1 écrite, non testée dans Excel.** Phases 1, 2 et 3 codées ensemble dans `mod_Analyse` / `mod_InstallAnalyse` (cases cochées = codé, à valider par test). Installation : voir l'en-tête de `mod_InstallAnalyse`.

### Phase 0 – Cadrage
- [ ] Valider la liste des décisions de la section 3 (surtout : exclusions, période par défaut, pas de sunburst)
- [ ] Lister les colonnes de TblOperations / TblVentilations utiles (Categorie, SousCategorie, Montant, MoisBudget, AnneeBudget) et vérifier leurs noms exacts

### Phase 1 – Répartition par catégorie / sous-catégorie
- [x] `mod_Analyse` : lecture des 2 tableaux en mémoire, remplacement des opérations « Ventilé » par leurs parts, filtre période, exclusions
- [x] `mod_Analyse` : agrégation catégorie → sous-catégorie (montant, % du total, nombre d'opérations) dans un dictionnaire
- [x] Liste d'exclusion de catégories : saisie directement sur l'écran (bloc « Catégories exclues »), pas dans Param
- [x] `mod_InstallAnalyse` : feuille `frm_Analyse` (en-tête, sélecteur de période, tableau de répartition, zone graphique, boutons)
- [x] Graphique en barres triées (catégories) + anneau ; zoom sur une catégorie = sous-catégories
- [x] Bouton d'accès sur Synthese + Sortir
- [ ] Test : total des dépenses de l'analyse = total du bilan mensuel (à ventilations près)

### Phase 2 – Lien avec l'écran de recherche
- [x] Nouveau préfiltre `Analyse` dans `mod_RechercheOperations` (catégorie, sous-catégorie, période en `param`)
- [x] Texte de filtre actif (`DecrireFiltreActifRO`) et colonnes visibles (`DefinirColonnesVisibles`)
- [x] Bouton « Voir les opérations » + retour sur `frm_Analyse` à la sortie de la recherche (comme `DetailTotal` avec le bilan)
- [ ] Test : nombre de lignes de la recherche = nombre d'opérations du tableau d'analyse
- [ ] Après recatégorisation depuis la recherche, l'analyse se recalcule au retour

### Phase 3 – Évolution dans le temps
- [x] `mod_Analyse` : matrice catégorie × mois de budget (12 derniers mois ou année)
- [x] Graphique courbes (catégories choisies) et barres empilées (toutes catégories)
- [x] Sélecteur : catégorie ou sous-catégorie, nombre de mois
- [ ] Moyenne glissante 3 mois
- [ ] Clic sur un mois / une catégorie du tableau = recherche filtrée (réutilise le préfiltre de la phase 2)

### Phase 4 – Comparaison et alertes (optionnel)
- [ ] Comparaison avec le mois précédent et le même mois N-1 (écart en € et en %)
- [ ] Top 5 hausses / baisses de catégories
- [ ] Mise en forme : rouge/vert avec libellé (accessibilité)

### Phase 5 – Budget cible (optionnel)
- [ ] Tableau `TblBudgets` dans Param : catégorie, montant mensuel cible
- [ ] Réel vs budget, écart = réel − budget et %, barre de progression
- [ ] Seuil d'alerte (ex. 80 % / 100 %) et lien avec le prévisionnel de l'échéancier

### Transverse
- [ ] Documenter dans README (section « Analyse budgétaire ») et regénérer `frm_Analyse_structure.txt` via `mod_ExportStructure`
- [ ] Utiliser les constantes existantes (`mod_VarGlobales`) plutôt que de recopier les noms de feuilles/tables
