Option Explicit

' =====================================================================================
' MODULE : mod_VarGlobales

' Cellule de démarrage de la plage d'écriture feuille Resultat
Public Const cellSortieDep As String = "A1"

' 1. Définition des couleurs (exemple avec un thème bleu)
Public Const coulEntete As Long = &H794E1F     ' Bleu principal (en-tête)
Public Const coulClair1 As Long = &HF8F2EE  ' Nuance 1 (très claire)
Public Const coulClair2 As Long = &HF8E3DA   ' Nuance 2 (légèrement plus soutenue)
Public Const coulBloc As Long = &H868878 ' Couleur des cellules de blocage
 
' --- Couleurs communes des MONTANTS (ajout 08/10/2026, décision opérateur) ----------
' Une seule définition pour tout le classeur, utilisée par la fonction commune
' mod_Display.AppliquerCouleurMontant. Changer une couleur ici la change partout.
' RAPPEL : une couleur VBA s'écrit &HBBVVRR& (Bleu, Vert, Rouge, dans cet ordre),
' c'est-à-dire l'inverse de l'écriture RGB habituelle ; le "&" final force le type
' Long (sans lui, &H8000 serait lu comme un nombre NÉGATIF de type Integer).
Public Const coulMontantPositif As Long = &H8000&   ' vert foncé = RGB(0, 128, 0)   : entrée d'argent
Public Const coulMontantNegatif As Long = &H0&      ' noir       = RGB(0, 0, 0)     : dépense
Public Const coulMontantAlerte As Long = &HFF&      ' rouge      = RGB(255, 0, 0)   : dépense à surveiller
 
' Défini si l'obtion double click doit être activé ou non sur un feuille
Public AllowDetailDoubleClick As Boolean
Public RecherOperations As Boolean

' Stocke le numéro d'index d'une colonne dans le tableau TblOperations
Public colCategorie As Long, colMontant As Long, colNotes As Long, colDate As Long, colLibelle As Long, colDateConsult As Long, colSpeConsult As Long
Public colType As Long, colBudget As Long, colTiers As Long, colMoisBud As Long, colAnneeBud As Long, colCheque As Long
Public colID As Long, colStatutSante As Long, colSoldeSante As Long, colDepassementHoraires As Long, colCommentaireSante As Long, colFranchise As Long
Public colBeneficiaire As Long

' Décalage manuel d'une operation précise (ajout 03/10/2026, demande opérateur) : vaut
' 0 (ou vide) pour la grande majorité des operations, qui suivent la règle générale
' lue dans TblDecalagesBudget. Quand l'opérateur force un décalage pour UNE opération
' en particulier (écran de recherche, bouton "Décaler le budget de cette operation"),
' cette colonne garde la valeur choisie : elle MARQUE la ligne, pour qu'un futur
' recalcul automatique ne vienne jamais l'écraser silencieusement (voir
' mod_DecalagesBudget.AppliquerDecalageManuel).
Public colDecalageManuel As Long

' --- PHASE 1/6 (gestion des catégories a 2 niveaux) --------------------------------
' Index de la colonne SousCategorie dans TblOperations. Vaut 0 tant que la colonne
' n'existe pas encore (feuille Param / Phase 1 pas encore installee) : tout code qui
' utilisé cette variable DOIT vérifier qu'elle est différent de 0 avant de s'en servir,
' exactement comme pour colCategorie, colMontant, etc.
Public colSousCategorie As Long

' --- Sous-categorie technique "Frais, remb sante" (ajout 03/10/2026) --------------
' Regroupee ici, au lieu d'etre recopiee en dur dans chaque module (demande operateur :
' regrouper les constantes globales), car PLUSIEURS modules doivent reconnaitre cette
' sous-categorie precise :
'   - mod_ImportOFX : y lit, dans le champ Notes, une date de consultation placee
'     AVANT le premier ";" (voir CalculerBudget/segment0 a l'import).
'   - mod_ControleCategories : interdit desormais la modification libre du Tiers et
'     de la Notes sur l'ecran de controle des categories pour cette sous-categorie,
'     justement pour ne pas abimer cette date avant meme qu'elle soit lue.
' Les 2 modules DOIVENT utiliser cette meme constante (jamais leur propre texte en
' dur) pour rester surs de comparer exactement la meme chaine.
Public Const SOUS_CATEGORIE_SANTE As String = "Frais, remb santé"
' --- Catégorie PARENTE de la sous-catégorie santé (ajout 08/10/2026) -------------
' Utilisée par le préfiltre "SuiviSante" de frm_RechercheOperations (écran ouvert par
' Synthese_Care) pour remonter AUSSI les opérations rangées dans cette catégorie
' mais sous une autre sous-catégorie : une dépense de santé mal sous-catégorisée se
' retrouve ainsi facilement. Même règle d'écriture que SOUS_CATEGORIE_SANTE : un seul
' endroit pour ce texte, jamais recopié en dur ailleurs.
' NOTE : mod_Categories contient encore une fonction privée CategorieSanteNouvelle()
' qui renvoie le même texte ; elle pourra être raccordée à cette constante lors du
' prochain passage de nettoyage (non fait ici, faute de validation).
Public Const CATEGORIE_SANTE As String = "Santé, prévoyance"

' --- Colonnes de suivi santé à vider quand une opération SORT de l'analyse santé ---
' (ajout du 07/10/2026, bouton "Changer la catégorie" de frm_RapprochementNotes)
' Liste des colonnes de TblOperations remplies UNIQUEMENT pour une ligne
' "Frais, remb santé" (initialisées à l'import par mod_ImportOFX, puis par le
' rapprochement et le suivi santé). Quand l'opérateur corrige la catégorie d'une
' opération qui n'était pas une dépense de santé, ces colonnes n'ont plus de sens :
' mod_ControleCategories.NettoyerColonnesSante les vide une par une (appelée par la
' fonction partagée ModifierCategorieOperation, donc depuis l'écran de recherche ET
' depuis l'écran de rapprochement).
'   - Les noms sont séparés par ";" (découpés avec Split au moment de l'emploi).
'   - La colonne "Notes" n'y figure VOLONTAIREMENT PAS : dans TblOperations, elle
'     contient le libellé bancaire d'origine, qui reste une donnée utile.
'   - Même liste (hors Notes) que celle utilisée par mod_InstallVentilation pour
'     créer les colonnes santé de TblVentilations.
Public Const COLONNES_SUIVI_SANTE As String = "Date_consult;Spe_Consult;StatutSante;SoldeSante;DepassementHoraires;CommentaireSante;Franchise;Beneficiaire"

' Stocke l'index d'une colonne, calculé à partir d'un des champs de l'ARRAY, pour toute la feuille de sortie.
' EXEMPLE Si on veut poser les entêtes à partir de E1 dans synthese,
' Array("Date", "Tiers", "Montant"): posDate=5,posTiers=6,posMontant=7
Public posSortieCategorie As Long, posSortieMontant As Long, posSortieNotes As Long, posSortieDate As Long, posSortieLibelle As Long, posSortieDateConsult As Long, posSortieSpeConsult As Long
Public posSortieType As Long, posSortieBudget As Long, posSortieTiers As Long, posSortieMoisBudget As Long, posSortieAnneeBudget As Long, posSortieCheque As Long
Public posSortieStatutSante As Long, posSortieSoldeSante As Long, posSortieID As Long

' Stocke la position d'une valeur dans un ARRAY
' EXEMPLE Array("Date", "Tiers", "Montant"): posDate=1,posTiers=2,posMontant=3
' A utiliser avec plageSortieEcriture
Public posCategorie As Long, posMontant As Long, posNotes As Long, posDate As Long, posDateConsult As Long, posSpeConsult As Long
Public posCheque As Long, posMoisBudget As Long, posAnneeBudget As Long, posType As Long, posBudget As Long, posTiers As Long, posMoisBud As Long, posAnneeBud As Long
Public posStatutSante As Long, posSoldeSante As Long, posID As Long, posValider As Long

' Stocke les valeurs des critères de recherche dans la feuille Synthese.
Public critAnnee As String, critMois As String, critNbOperations As Long, critMontantMin As Double, critTriChamps As String, critTriOrdre As String

' En vu de faciliter la relecture de nombreuses variables sont déclarées au niveau global
' pour ne pas surcharger les macros
Public wsSynthese As Worksheet
Public wsResultat As Worksheet
Public tbl As ListObject
Public tblData As Variant
Public tblDataLineTotal As Long
Public MonArray As Variant, nbColonne As Long, tblSortie As Variant
Public plageSortieEnTetes As Range, plageSortieEcriture As Range

' Les noms de feuille
Public Const NOM_FEUILLE_DETAIL_POSITIF As String = "frm_Detail_Positif"
Public Const NOM_FEUILLE_DETAIL_NEGATIF As String = "frm_Detail_Negatif"
Public Const NOM_FEUILLE_DONNEES As String = "Import_Data"
Public Const NOM_FEUILLE_RAPPROCHEMENT As String = "frm_RapprochementNotes"
Public Const NOM_FEUILLE_GENERATION As String = "frm_GenerationCle"
Public Const NOM_FEUILLE_RESULTAT As String = "frm_Resultat"
Public Const NOM_FEUILLE_SYNTHESE As String = "Synthese"
Public Const NOM_TABLE_RECHERCHE As String = "TblRechercheOperations"
Public Const NOM_FEUILLE_RECHERCHE As String = "frm_RechercheOperations"
Public Const NOM_FEUILLE_TECH As String = "TechDernierImport"
Public Const NOM_FEUILLE_SUIVI_SANTE As String = "frm_SuiviSante"
Public Const NOM_FEUILLE_RESOLUTION As String = "frm_ResolutionCategories"
 
' Variables de construction d'une zone de texte
Public PositionTitre As Range, debutZone As Range, finZone As Range
Public Titre As String, Message As String

' Variables de construction d'un bouton
Public zoneBouton As Range, texteBouton As String, nomMacroBouton As String

'' -------------------------------------------------------------------------------------
'' CONSTANTES DE MISE EN PAGE
'' -------------------------------------------------------------------------------------
'' Regrouper ici toutes les positions de cellules evite d'avoir des "nombres magiques"
'' disperses dans le code. Si un jour tu veux deplacer une zone, tu changes UNE seule
'' ligne ici plutot que de chercher partout dans le code.
'' Ces constantes seront reutilisees telles quelles dans le module de la Phase 2.
' --- Colonnes du tableau (une lettre = une colonne Excel) ---
Public Const COL_STATUT As String = "B"
Public Const COL_DATE As String = "C"
Public Const COL_MONTANT As String = "D"
Public Const COL_TIERS As String = "E"
Public Const COL_CATEGORIE As String = "F"
' --- Zone du tableau des cas ambigus ---
Public Const LIGNE_ENTETES_TABLEAU As Long = 8    ' ligne des libelles de colonnes (09/10/2026 : etait 11, voir mod_InstallResolutionSheet)
Public Const LIGNE_PREMIERE_DONNEE As Long = 9    ' premiere ligne ou s'affichera un cas (09/10/2026 : etait 12)



' Nombre maximal de cas que la mise en forme du tableau preparera a l'avance.
' Ce n'est PAS une limite dure : en Phase 2, si jamais il y avait plus de cas que
' cela, le code etendra la mise en forme automatiquement. Ce nombre sert juste a
' preparer une zone confortable des l'installation, pour un rendu propre immediat.
Public Const NB_LIGNES_PREPAREES As Long = 200

'' -------------------------------------------------------------------------------------
'' TABLEAU DES DECALAGES DE BUDGET (ajout 03/10/2026, demande operateur)
'' -------------------------------------------------------------------------------------
'' Remplace la regle codee en dur qui existait dans mod_ImportOFX.CalculerBudget
'' (seul le cas "DRFIP OCCITANIE ET HTE" + "Salaire/Revenus d'activite" beneficiait
'' d'un decalage d'un mois). Desormais, N'IMPORTE QUELLE combinaison Tiers / Categorie /
'' Sous-categorie peut avoir son propre decalage (+1, -1, ou toute autre valeur entiere),
'' en ajoutant simplement une ligne dans ce tableau - voir mod_DecalagesBudget.bas pour
'' la logique de lecture, et mod_Categories.NOM_FEUILLE_PARAM pour la feuille (reutilisee
'' telle quelle, pas de redeclaration de "Param" ici - un seul nom, un seul endroit).
Public Const NOM_TABLE_DECALAGES_BUDGET As String = "TblDecalagesBudget"
' Les colonnes A a R de Param sont deja utilisees (TblCategories en K:N notamment) :
' on place ce nouveau tableau plus a droite, a partir de la colonne S.
Public Const DECALAGES_BUDGET_COL_DEBUT As Long = 19   ' colonne S
Public Const DECALAGES_BUDGET_NB_COL As Long = 4        ' S,T,U,V = Tiers/Categorie/SousCategorie/Decalage

' Nom de la colonne technique ajoutee a TblOperations (ajout 03/10/2026) pour
' marquer/memoriser le decalage manuel d'une operation precise - voir colDecalageManuel
' ci-dessus et mod_DecalagesBudget.AjouterColonneDecalageManuel/AppliquerDecalageManuel.
Public Const NOM_COL_DECALAGE_MANUEL As String = "DecalageManuel"

'' -------------------------------------------------------------------------------------
'' EXPORT DE LA STRUCTURE DES FEUILLES (ajout 08/10/2026, voir mod_ExportStructure)
'' -------------------------------------------------------------------------------------
' Sous-dossier créé à côté du classeur pour recevoir les fichiers d'export (un
' fichier "<NomFeuille>_structure.txt" par feuille exportée).
Public Const EXPORT_SOUS_DOSSIER As String = "Export_Structure"
' Nombre maximal de lignes analysées en détail (cellules, hauteurs de lignes). Les
' feuilles-formulaires ont leur mise en page en haut de feuille ; au-delà, ce sont
' des données, qu'on ne veut pas exporter. À augmenter si une mise en page descend
' plus bas (ex. frm_ResolutionCategories prépare NB_LIGNES_PREPAREES = 200 lignes).
Public Const EXPORT_NB_LIGNES_MAX As Long = 100
' Garde-fou : nombre maximal de cellules décrites par feuille.
Public Const EXPORT_NB_CELLULES_MAX As Long = 5000

