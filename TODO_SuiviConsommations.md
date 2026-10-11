# Suivi des consommations (eau, électricité, gaz) – décisions et TODO

Créé le 11/10/2026 · branche `Refonte_recherche`

## 1. La demande

Suivre la consommation d'électricité, de gaz et d'eau (adduction et recyclage) : historiser, **pour chaque période**, le nombre d'unités consommées et le prix de l'unité, avec une vue graphique (barres ou courbes) de leur évolution. L'opérateur doit pouvoir changer en cours de route (nouveau contrat).

Deux précisions de l'opérateur, validées le 11/10/2026 :

- **L'abonnement peut évoluer en cours de contrat**, comme le prix. Il est donc porté par **chaque relevé** ; le contrat n'en donne que la valeur de départ, et elle reste ajustable relevé par relevé.
- **Les périodes sont personnalisées** (début et fin libres) : certaines factures voient le prix de l'unité changer plusieurs fois dans le mois. Une facture à plusieurs prix se saisit en plusieurs relevés consécutifs (même N° de facture).

## 2. Décisions de conception

- **Deux tableaux** : `TblContrats` (type, opérateur, unité, dates, abonnement de départ, référence) et `TblReleves` (type, début, fin, quantité, prix de l'unité, abonnement mensuel, N° de facture, commentaire, et 4 colonnes calculées : `Operateur`, `NbMois`, `Abonnement`, `Montant`).
- **Cinq feuilles-écrans masquées**, même principe que l'échéancier : `frm_Consommations` (paramètres, tableau mensuel, graphique), `frm_Releves` et `frm_Contrats` (listes), `frm_Releve` et `frm_Contrat` (formulaires modaux). Un module `mod_Consommations` (logique) + `mod_InstallConsommations` (construction), constantes de noms dans `mod_VarGlobales`.
- **Changement d'opérateur** = nouveau contrat ; le contrat en cours du même type est clôturé la veille (avec confirmation). Chaque relevé retrouve son opérateur par ses dates ; une période ne peut pas déborder de son contrat (le formulaire coupe le relevé en deux).
- **NbMois d'une période** = somme, pour chaque mois touché, de (jours couverts / jours du mois). Un mois complet vaut 1, et des relevés qui se partagent un mois totalisent exactement 1 abonnement. Même règle en formule Excel (colonne `NbMois`, formulaire) et en VBA (`NbMoisPeriode`) : **à garder identiques**.
- **Graphique** : répartition au prorata des jours par défaut (sinon : tout dans le mois de fin du relevé) ; barres empilées par opérateur (un type) ou par type (tous), courbes, ou barres + courbe du prix sur axe secondaire. Prix mensuel = moyenne pondérée par les quantités, abonnement exclu.
- **Pas de lien avec les opérations bancaires** en V1 : relevés saisis d'après les factures.
- **Libellés** : « Montant » et non « Coût » (la balise `{u3}` n'existe pas dans `FR()`).

## 3. Installation

1. Importer `mod_VarGlobales.bas` (mis à jour), `mod_Consommations.bas`, `mod_InstallConsommations.bas`.
2. Ctrl+G : `InstallerConsommations`.
3. Coller `CodeBehind_frm_Releve.txt` dans la feuille `frm_Releve` et `CodeBehind_frm_Contrat.txt` dans la feuille `frm_Contrat` (les CodeName sont donnés par le message de fin d'installation).

## 4. TODO

> **État 11/10/2026 : V1 écrite, non testée dans Excel.** Les cases cochées = codé, à valider par test.

### Phase 1 – Données et saisie
- [x] Constantes de noms dans `mod_VarGlobales`
- [x] `mod_InstallConsommations` : formulaires `frm_Releve` / `frm_Contrat`, listes `frm_Releves` / `frm_Contrats` (tableaux, colonnes calculées, listes déroulantes), bouton sur Synthese
- [x] `mod_Consommations` : saisie d'un relevé (pré-remplissage du dernier prix et abonnement, « Valider + suivant »), saisie d'un contrat (clôture du contrat en cours), contrôles de chevauchement
- [x] Suppression de la sélection dans les deux listes
- [ ] Test : installer, créer un contrat par type, saisir un relevé par type ; les colonnes grises de `TblReleves` se calculent
- [ ] Test : changement d'opérateur (clôture la veille, opérateur correct sur les anciens et nouveaux relevés)
- [ ] Test : facture à deux prix (« Valider + suivant ») ; abonnement total du mois = 1 abonnement
- [ ] Test : relevé qui chevauche un autre / déborde d'un contrat / sans contrat → message, rien d'enregistré
- [ ] Vérifier que la formule `Operateur` (`LOOKUP`) s'affiche (sinon la colonne reste vide : sans conséquence sur le graphique)

### Phase 2 – Écran et graphique
- [x] Tableau mensuel + total, données masquées du graphique
- [x] Barres empilées / courbes / barres + courbe du prix (axe secondaire), axes titrés
- [x] Répartition au prorata des jours ou mois de fin
- [x] Périodes : 12 mois, 36 mois, année civile
- [ ] Test : total de la colonne Montant = somme des `Montant` de `TblReleves` (période complète, répartition « Mois de fin »)
- [ ] Test : un relevé à cheval sur deux mois se répartit au prorata des jours ; la somme des deux mois = le relevé
- [ ] Test : « (tous) » avec Quantité ou Prix → bascule sur Montant avec le message ; forme mixte avec « (tous) » → barres
- [ ] Test : plus de 6 opérateurs pour un type → série « Autres »
- [ ] Test : axe des mois (étiquettes mmm yy, une graduation sur trois à 36 mois)

### Phase 3 – À faire ensuite (optionnel)
- [ ] Comparaison avec l'année précédente (N-1) : courbe superposée
- [ ] Économie du recyclage de l'eau = volume recyclé × prix de l'adduction du même mois
- [ ] Total annuel et variation en % par rapport à N-1
- [ ] Rapprochement avec les opérations bancaires (Tiers + montant), comme l'échéancier
- [ ] Prix en tranches (eau) ou heures pleines / creuses (électricité) : aujourd'hui, un prix moyen par relevé
- [ ] Contrôles à la saisie directe dans les listes (chevauchement, contrat)

### Transverse
- [ ] Générer les `…_structure.txt` des 5 feuilles avec `mod_ExportStructure`
- [x] Documenter dans le README (section 10) et régénérer `README.html` (`python build_docs.py`)
- [ ] Factoriser les petits outils de lecture de tableau recopiés de `mod_Echeancier` (`ObtenirTable`, `TableVide`, `NouvelleLigne`, `LireColonne`, `ProchainID`)
