Attribute VB_Name = "mod_VarGlobales"
Option Explicit

' Cellule de démarrage de la plage d'écriture feuille RÃ©sultat
Public Const cellSortieDep As String = "A1"

' 1. Définition des couleurs (exemple avec un thème bleu)
Public Const coulEntete As Long = &H794E1F     ' Bleu principal (en-tête)
Public Const coulClair1 As Long = &HF8F2EE  ' Nuance 1 (très claire)
Public Const coulClair2 As Long = &HF8E3DA   ' Nuance 2 (légèrement plus soutenue)

' Défini si l'obtion double click doit être activé ou non sur un feuille
Public AllowDetailDoubleClick As Boolean

' Stocke le numéro d'index d'une colonne dans le tableau TblOperations
Public colCategorie As Long, colMontant As Long, colNotes As Long, colDate As Long, colLibelle As Long, colDateConsult As Long, colSpeConsult As Long
Public colType As Long, colBudget As Long, colTiers As Long, colMoisBud As Long, colAnneeBud As Long, colCheque As Long
Public colID As Long, colStatutSante As Long, colSoldeSante As Long, colDepassementHoraires As Long, colCommentaireSante As Long, colFranchise As Long
Public colBeneficiaire As Long

' --- PHASE 1/6 (gestion des catÃ©gories a 2 niveaux) --------------------------------
' Index de la colonne SousCategorie dans TblOperations. Vaut 0 tant que la colonne
' n'existe pas encore (feuille Param / Phase 1 pas encore installee) : tout code qui
' utilisÃ© cette variable DOIT vÃ©rifier qu'elle est diffÃ©rent de 0 avant de s'en servir,
' exactement comme pour colCategorie, colMontant, etc.
Public colSousCategorie As Long

' Stoque l'index d'une colonne, calculé à partir d'un des champ de l'ARRAY, sur toute la feuille de sortie
' EXEMPLE Si on veut poser les entêtes à partir de E1 dans synthese,
' Array("Date", "Tiers", "Montant"): posDate=5,posTiers=6,posMontant=7
Public colSortieCategorie As Long, colSortieMontant As Long, colSortieNotes As Long, colSortieDate As Long, colSortieLibelle As Long, colSortieDateConsult As Long, colSortieSpeConsult As Long
Public colSortieType As Long, colSortieBudget As Long, colSortieTiers As Long, colSortieMoisBudget As Long, colSortieAnneeBudget As Long, colSortieCheque As Long
Public colSortieStatutSante As Long, colSortieSoldeSante As Long

' Stocke la position d'une valeur dans un ARRAY
' EXEMPLE Array("Date", "Tiers", "Montant"): posDate=1,posTiers=2,posMontant=3
' A utiliser avec plageSortieEcriture
Public posCategorie As Long, posMontant As Long, posNotes As Long, posDate As Long, posDateConsult As Long, posSpeConsult As Long
Public posCheque As Long, posMoisBudget As Long, posAnneeBudget As Long, posType As Long, posBudget As Long, posTiers As Long, posMoisBud As Long, posAnneeBud As Long
Public posStatutSante As Long, posSoldeSante

' Stoque les valeurs des critères de recherche dans la feuille Synthese
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
Public Const NOM_FEUILLE_DONNEES As String = "Import_Data"
Public Const NOM_FEUILLE_RAPPROCHEMENT As String = "frm_RapprochementNotes"
Public Const NOM_FEUILLE_GENERATION As String = "frm_GenerationCle"
Public Const NOM_FEUILLE_RESULTAT As String = "frm_Resultat"
Public Const NOM_FEUILLE_SYNTHESE As String = "Synthese"

' Variables de construction d'une zone de texte
Public PositionTitre As Range, debutZone As Range, finZone As Range
Public Titre As String, Message As String

' Variables de construction d'un bouton
Public zoneBouton As Range, texteBouton As String, nomMacroBouton As String

