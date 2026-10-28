Attribute VB_Name = "mod_Factures"
Option Explicit

' =====================================================================================
' MODULE : mod_Factures
'
' SUIVI DES CONSOMMATIONS - SAISIE DES FACTURES (ajout 19/10/2026). Remplace les anciens
' "relevés" (une seule période commune à toute la facture). La construction des feuilles est
' dans mod_InstallConsommations ; le graphique mensuel est dans mod_Consommations.
'
' MODÈLE (calqué sur la facture du fournisseur) :
'   - TblFactures : une ligne par facture (type, opérateur, N°, date, période facturée, total TTC facturé).
'   - TblLignes   : TOUTES les lignes de TOUTES les factures (feuille masquée Conso_Lignes). Une ligne appartient à un
'                   BLOC (colonne Bloc : son nom, unique dans la facture) placé à une POSITION (OrdreBloc), et à une
'                   INDICATEUR (colonne Indicateur : Abonnement, Consommation, Taxes, Autre ou un indicateur créé) qui sert aux totaux et au
'                   graphique. Chaque ligne a SA période, sa quantité, SON prix unitaire HT, SON taux de TVA et ses
'                   montants HT et TTC. Les lignes sont reliées à leur facture par ID_Facture.
'   Un formulaire de facture neuve propose le MODÈLE : trois blocs (Abonnement, Consommation, Taxes et contributions).
'   Chaque bloc peut être supprimé et d'autres peuvent être ajoutés (bouton Ajouter un bloc) ; à la réouverture, la
'   facture retrouve exactement ses blocs, dans le même ordre.
'
' RÈGLES DE CALCUL (identiques à celles des factures de l'opérateur) :
'   - Abonnement : quantité et unité se saisissent comme pour les autres blocs (par exemple 31 jours).
'   - Montant HT d'une ligne = ROUND(quantité x prix unitaire HT ; 2).
'   - Montant TTC d'une ligne = ROUND(HT x (1 + TVA) ; 2).
'   - Le total de la facture est CALCULÉ À TITRE INDICATIF (somme des lignes) et comparé au
'     total facturé saisi : l'écart est affiché mais n'empêche jamais l'enregistrement.
'
' SAISIE : formulaire frm_Facture (en-tête + trois blocs de lignes, extensibles),
'   ouvert par SaisirFacture (nouvelle facture) ou ModifierFactureSelection (liste
'   frm_Factures). Les montants du formulaire se calculent en direct (formules).
'
' À PROPOS DES ACCENTS : textes affichés via mod_Display.FR(); commentaires en UTF-8.
' =====================================================================================

' --- Natures de lignes ---
Public Const NAT_ABO As Long = 1
Public Const NAT_CONSO As Long = 2
Public Const NAT_TAXE As Long = 3
Public Const NAT_AUTRE As Long = 4          ' famille "Autre" : blocs ajoutés qui ne sont ni abonnement, ni consommation, ni taxes

' --- Formulaire frm_Facture : en-tête (B = étiquettes, C = saisie) ---
Public Const FF_ADR_TYPE As String = "C6"             ' fusion C6:D6
Public Const FF_ADR_OPERATEUR As String = "C7"        ' fusion C7:E7
Public Const FF_ADR_NUMERO As String = "C8"           ' fusion C8:D8
Public Const FF_ADR_DATE As String = "C9"
Public Const FF_ADR_DEBUT As String = "C10"
Public Const FF_ADR_FIN As String = "C11"
Public Const FF_ADR_TOTAL As String = "C12"
Public Const FF_ADR_COMMENT As String = "C13"         ' fusion C13:J13

' --- Blocs de lignes : colonnes ---
Public Const FF_COL_LIB As Long = 2
Public Const FF_COL_DEB As Long = 3
Public Const FF_COL_FIN As Long = 4
Public Const FF_COL_QTE As Long = 5
Public Const FF_COL_UNITE As Long = 6
Public Const FF_COL_PRIX As Long = 7
Public Const FF_COL_TVA As Long = 8
Public Const FF_COL_HT As Long = 9
Public Const FF_COL_TTC As Long = 10

' --- Récapitulatif (indicatif), placé EN HAUT pour que les blocs puissent grandir vers le bas ---
Public Const FF_RECAP_TITRE As Long = 15
Public Const FF_ADR_RECAP_HT As String = "C16"
Public Const FF_ADR_RECAP_TVA As String = "C17"
Public Const FF_ADR_RECAP_TTC As String = "C18"
Public Const FF_ADR_RECAP_SAISI As String = "C19"
Public Const FF_ADR_RECAP_ECART As String = "C20"
Public Const FF_ADR_MESSAGE As String = "B22"         ' fusion B22:J22
Public Const FF_COL_AIDE_TYPE As Long = 26            ' colonne technique masquée (liste des types)
Public Const FF_COL_AIDE_OPER As Long = 27            ' colonne technique masquée (opérateurs en cours du type choisi)
Public Const FF_COL_AIDE_UNITE As Long = 28           ' colonne technique masquée (unités proposées)
Public Const FF_COL_AIDE_UNITE_CONSO As Long = 29     ' colonne technique masquée (unités admises pour la consommation : kWh, m³)
Public Const FF_LARG_LIBELLE As Double = 44           ' largeur de la colonne B (libellés des lignes)

' --- Blocs de lignes : ils commencent ligne 24 et GRANDISSENT (bouton "+" du titre de chaque bloc) ---
' Chaque bloc occupe : titre, en-tête, n lignes, sous-total, séparateur (n + 4 lignes).
Public Const FF_BLOCS_DEBUT As Long = 24
Public Const FF_NB_DEFAUT As Long = 3           ' lignes affichées au départ dans chaque bloc
Public Const FF_MAX_LIGNES As Long = 40         ' limite d'un bloc
Public Const FF_MAX_BLOCS As Long = 9           ' blocs au plus dans un formulaire (le modèle en compte 3)
Public Const FF_LONG_NOM_BLOC As Long = 60      ' longueur maximale du nom d'un bloc ajouté
Public Const FF_LONG_INDIC As Long = 40         ' longueur maximale du nom d'un indicateur
Public Const FF_LIGNE_BOUTONS As Long = 2       ' ligne des boutons du formulaire

' Modes de ModifierStructure (déclarations de module : elles doivent précéder toute procédure).
Private Const MODE_PLUS As Long = 1
Private Const MODE_MOINS As Long = 2
Private Const MODE_AJOUT_BLOC As Long = 3
Private Const MODE_SUPPR_BLOC As Long = 4
Private Const MODE_INDIC As Long = 5

' Une facture telle que saisie dans le formulaire.
Public Type TypeFacture
    ID As Long                  ' 0 = nouvelle facture
    Typ As String
    Operateur As String
    Numero As String
    DateFac As Double           ' 0 = non renseignée
    Debut As Double
    Fin As Double
    TotalSaisi As Double
    TotalRenseigne As Boolean
    Commentaire As String
End Type

' Une ligne de facture (abonnement, consommation ou taxe).
Public Type TypeLigne
    Nat As Long                 ' code de l'indicateur : 1 abonnement, 2 consommation, 3 taxes, 4 tout autre indicateur
    Indic As String             ' NOM de l'indicateur alimenté par la ligne (colonne Indicateur)
    Bloc As String              ' nom du bloc qui porte la ligne (colonne Bloc)
    Ordre As Long               ' POSITION du bloc dans la facture (1 = premier bloc affiché), colonne OrdreBloc
    IDFacture As Long
    Typ As String               ' type de la facture (renseigné au chargement)
    Operateur As String         ' opérateur de la facture (renseigné au chargement)
    Libelle As String
    Debut As Double
    Fin As Double
    Qte As Double
    Unite As String
    Prix As Double              ' prix unitaire HT
    Tva As Double               ' 0,055 pour 5,5 %
    HT As Double
    TTC As Double
    Copie As Boolean            ' ligne issue d'une copie : dates, quantité et prix laissés vides
End Type

' --- État du formulaire de saisie ---
Private g_FacEnCours As Boolean
Private g_FacVerrouActif As Boolean
Private g_FacValide As Boolean
Private g_FacRes As TypeFacture
Private g_FacLignes() As TypeLigne
Private g_FacNb As Long
Private g_FacNomFeuillePrec As String
Private g_IdEnEdition As Long
Private g_NbUnites As Long
Private g_AutoCompleteAvant As Boolean
Private g_NbBlocs As Long                     ' blocs affichés : 3 fixes + les blocs libres
Private g_N(1 To FF_MAX_BLOCS) As Long        ' nombre de lignes affichées dans chaque bloc
Private g_Nat(1 To FF_MAX_BLOCS) As Long      ' code de l'indicateur de chaque bloc : 1 abonnement, 2 consommation, 3 taxes, 4 autre
Private g_Indic(1 To FF_MAX_BLOCS) As String  ' nom de l'indicateur de chaque bloc (affiché entre parenthèses dans son titre)
Private g_FacConsultation As Boolean          ' formulaire ouvert en consultation : rien n'est enregistré
Private g_Pend As Object                      ' changements d'indicateur d'un bloc à répercuter sur d'autres factures à l'enregistrement
Private g_NomBloc(1 To FF_MAX_BLOCS) As String ' nom des blocs libres


' =====================================================================================
' ACCÈS AUX TABLEAUX
' =====================================================================================
Public Function TableFactures() As ListObject
    Set TableFactures = mod_Consommations.ObtenirTable(mod_VarGlobales.NOM_FEUILLE_DATA_FACTURES, mod_VarGlobales.NOM_TABLE_FACTURES)
End Function

' Toutes les lignes de toutes les factures (tous blocs, toutes familles) sont dans UN SEUL tableau.
Public Function TableLignes() As ListObject
    Set TableLignes = mod_Consommations.ObtenirTable(mod_VarGlobales.NOM_FEUILLE_DATA_LIGNES, mod_VarGlobales.NOM_TABLE_LIGNES)
End Function

Private Function TableAffFactures() As ListObject
    Set TableAffFactures = mod_Consommations.ObtenirTable(mod_VarGlobales.NOM_FEUILLE_FACTURES, mod_VarGlobales.NOM_TABLE_FACTURES_AFF)
End Function

Public Function FacturesInstallees() As Boolean
    FacturesInstallees = Not (TableFactures() Is Nothing) And Not (TableLignes() Is Nothing)
End Function

' Nombre de lignes affichées dans un bloc, et positions (le formulaire est redessiné quand ce nombre change).
Private Function BlocN(ByVal b As Long) As Long
    BlocN = g_N(b)
End Function

Private Function BlocTitre(ByVal b As Long) As Long
    Dim t As Long, k As Long
    t = FF_BLOCS_DEBUT
    For k = 1 To b - 1
        t = t + g_N(k) + 4
    Next k
    BlocTitre = t
End Function

Private Function BlocL1(ByVal b As Long) As Long
    BlocL1 = BlocTitre(b) + 2
End Function

' Nom d'un des quatre indicateurs de base (valeur de la colonne Indicateur).
Private Function NomNature(ByVal nat As Long) As String
    Select Case nat
        Case NAT_ABO: NomNature = "Abonnement"
        Case NAT_CONSO: NomNature = "Consommation"
        Case NAT_TAXE: NomNature = "Taxes"
        Case Else: NomNature = "Autre"
    End Select
End Function

' Code de l'indicateur d'après son nom : 1 à 3 pour Abonnement, Consommation et Taxes, 4 (NAT_AUTRE) pour tout autre
' indicateur, qu'il soit « Autre » ou créé par l'opérateur. Le graphique ne connaît pour l'instant que les trois premiers.
Private Function FamilleDepuisTexte(ByVal txt As String) As Long
    Select Case LCase$(Trim$(txt))
        Case "abonnement": FamilleDepuisTexte = NAT_ABO
        Case "consommation": FamilleDepuisTexte = NAT_CONSO
        Case "taxes", "taxe", "taxes et contributions": FamilleDepuisTexte = NAT_TAXE
        Case Else: FamilleDepuisTexte = NAT_AUTRE
    End Select
End Function

' =====================================================================================
' INDICATEURS : ce que peut alimenter un bloc (colonne Indicateur de TblLignes)
' =====================================================================================
Public Function TableIndicateurs() As ListObject
    Set TableIndicateurs = mod_Consommations.ObtenirTable(mod_VarGlobales.NOM_FEUILLE_DATA_INDIC, mod_VarGlobales.NOM_TABLE_INDIC)
End Function

Private Sub AjouterNomListe(ByRef noms() As String, ByRef nbNoms As Long, ByVal txt As String)
    Dim i As Long
    txt = Trim$(txt)
    If txt = "" Then Exit Sub
    For i = 1 To nbNoms
        If StrComp(noms(i), txt, vbTextCompare) = 0 Then Exit Sub
    Next i
    nbNoms = nbNoms + 1
    If nbNoms > UBound(noms) Then ReDim Preserve noms(1 To nbNoms + 15)
    noms(nbNoms) = txt
End Sub

' Indicateurs proposés, sans doublon : les quatre de base (Abonnement, Consommation, Taxes, Autre), puis ceux créés
' par l'opérateur (TblIndicateurs), puis ceux déjà utilisés par des lignes enregistrées. Renvoie leur nombre ;
' les noms sont dans noms() (base 1).
Public Function ListeIndicateurs(ByRef noms() As String) As Long

    Dim nbNoms As Long, i As Long
    Dim lo As ListObject
    Dim vals() As Variant

    ReDim noms(1 To 20)
    For i = NAT_ABO To NAT_AUTRE
        AjouterNomListe noms, nbNoms, NomNature(i)
    Next i

    Set lo = TableIndicateurs()
    If Not lo Is Nothing Then
        If Not mod_Consommations.TableVide(lo) Then
            vals = mod_Consommations.LireColonne(lo, "Indicateur")
            For i = 1 To lo.ListRows.Count
                AjouterNomListe noms, nbNoms, mod_DataStructure.CellText(vals(i))
            Next i
        End If
    End If

    Set lo = TableLignes()
    If Not lo Is Nothing Then
        If Not mod_Consommations.TableVide(lo) And mod_Consommations.ColIdx(lo, "Indicateur") > 0 Then
            vals = mod_Consommations.LireColonne(lo, "Indicateur")
            For i = 1 To lo.ListRows.Count
                AjouterNomListe noms, nbNoms, mod_DataStructure.CellText(vals(i))
            Next i
        End If
    End If

    ListeIndicateurs = nbNoms

End Function

' Mémorise un indicateur créé par l'opérateur : il sera proposé dans toutes les listes futures.
Private Sub MemoriserIndicateur(ByVal nom As String)

    Dim lo As ListObject
    Dim lr As ListRow
    Dim vals() As Variant
    Dim i As Long

    Set lo = TableIndicateurs()
    If lo Is Nothing Then Exit Sub
    If Not mod_Consommations.TableVide(lo) Then
        vals = mod_Consommations.LireColonne(lo, "Indicateur")
        For i = 1 To lo.ListRows.Count
            If StrComp(mod_DataStructure.CellText(vals(i)), nom, vbTextCompare) = 0 Then Exit Sub
        Next i
    End If
    Set lr = mod_Consommations.NouvelleLigne(lo)
    lr.Range.Cells(1, mod_Consommations.ColIdx(lo, "Indicateur")).Value = nom

End Sub

' Nom des trois blocs du MODÈLE d'une nouvelle facture (un par famille : abonnement, consommation, taxes).
Public Function NomBlocModele(ByVal nat As Long) As String
    Select Case nat
        Case NAT_ABO: NomBlocModele = "Abonnement"
        Case NAT_CONSO: NomBlocModele = "Consommation"
        Case NAT_TAXE: NomBlocModele = "Taxes et contributions"
    End Select
End Function

' Formules des montants d'une ligne du formulaire (syntaxe anglaise).
Public Function FormuleHT(ByVal ligne As Long) As String
    Dim e As String, g As String
    e = mod_Consommations.AdresseCol(FF_COL_QTE) & ligne
    g = mod_Consommations.AdresseCol(FF_COL_PRIX) & ligne
    FormuleHT = "=IF(AND(ISNUMBER(" & e & "),ISNUMBER(" & g & ")),ROUND(" & e & "*" & g & ",2),"""")"
End Function

Public Function FormuleTTC(ByVal ligne As Long) As String
    Dim i As String, h As String
    i = mod_Consommations.AdresseCol(FF_COL_HT) & ligne
    h = mod_Consommations.AdresseCol(FF_COL_TVA) & ligne
    FormuleTTC = "=IF(ISNUMBER(" & i & "),ROUND(" & i & "*(1+N(" & h & ")),2),"""")"
End Function

' =====================================================================================
' CHARGEMENT DES LIGNES (utilisé par le graphique et par les contrôles)
' =====================================================================================
' Remplit "t" (base 1) avec toutes les lignes valides de TblLignes, avec le type et l'opérateur de leur facture.
' Renvoie leur nombre. Si "sauf" > 0, les lignes de cette facture sont ignorées (contrôle de chevauchement lors
' d'une modification). Les lignes de famille "Autre" ne sont renvoyées que si "inclureAutres" : le graphique ne les
' utilise pas encore (leur rattachement aux indicateurs reste à décider).
Public Function ChargerLignes(ByRef t() As TypeLigne, Optional ByVal sauf As Long = 0, Optional ByVal inclureAutres As Boolean = False) As Long

    Dim loF As ListObject, lo As ListObject
    Dim dTyp As Object, dOp As Object
    Dim vID() As Variant, vT() As Variant, vO() As Variant
    Dim vI() As Variant, vL() As Variant, vD() As Variant, vF() As Variant, vQ() As Variant
    Dim vU() As Variant, vP() As Variant, vV() As Variant, vH() As Variant, vC() As Variant
    Dim vB() As Variant, vFam() As Variant, vOr() As Variant
    Dim i As Long, n As Long, nat As Long, nb As Long, id As Long
    Dim d As Double, f As Double
    Dim cle As String

    ReDim t(1 To 1)
    Set loF = TableFactures()
    If loF Is Nothing Then Exit Function

    Set dTyp = CreateObject("Scripting.Dictionary")
    Set dOp = CreateObject("Scripting.Dictionary")
    If Not mod_Consommations.TableVide(loF) Then
        vID = mod_Consommations.LireColonne(loF, "ID_Facture")
        vT = mod_Consommations.LireColonne(loF, "TypeConso")
        vO = mod_Consommations.LireColonne(loF, "Operateur")
        For i = 1 To loF.ListRows.Count
            If IsNumeric(vID(i)) And Not IsEmpty(vID(i)) Then
                cle = CStr(CLng(vID(i)))
                dTyp(cle) = mod_DataStructure.CellText(vT(i))
                dOp(cle) = mod_DataStructure.CellText(vO(i))
            End If
        Next i
    End If

    Set lo = TableLignes()
    If lo Is Nothing Then Exit Function
    If mod_Consommations.TableVide(lo) Then Exit Function

    n = lo.ListRows.Count
    vI = mod_Consommations.LireColonne(lo, "ID_Facture")
    vL = mod_Consommations.LireColonne(lo, "Libelle")
    vD = mod_Consommations.LireColonne(lo, "DateDebut")
    vF = mod_Consommations.LireColonne(lo, "DateFin")
    vQ = mod_Consommations.LireColonne(lo, "Quantite")
    vU = mod_Consommations.LireColonne(lo, "Unite")
    vP = mod_Consommations.LireColonne(lo, "PrixUnitaireHT")
    vV = mod_Consommations.LireColonne(lo, "TauxTVA")
    vH = mod_Consommations.LireColonne(lo, "MontantHT")
    vC = mod_Consommations.LireColonne(lo, "MontantTTC")
    vB = mod_Consommations.LireColonne(lo, "Bloc")
    vFam = mod_Consommations.LireColonne(lo, "Indicateur")
    vOr = mod_Consommations.LireColonne(lo, "OrdreBloc")

    For i = 1 To n
        If IsNumeric(vI(i)) And Not IsEmpty(vI(i)) Then
            id = CLng(vI(i))
            cle = CStr(id)
            d = mod_Consommations.VersJour(vD(i))
            f = mod_Consommations.VersJour(vF(i))
            nat = FamilleDepuisTexte(mod_DataStructure.CellText(vFam(i)))
            If dTyp.Exists(cle) And d > 0 And f >= d And id <> sauf And (inclureAutres Or nat <> NAT_AUTRE) Then
                nb = nb + 1
                If nb > UBound(t) Then ReDim Preserve t(1 To nb + 63)
                With t(nb)
                    .Nat = nat
                    .Indic = mod_DataStructure.CellText(vFam(i))
                    If .Indic = "" Then .Indic = NomNature(nat)
                    .Bloc = mod_DataStructure.CellText(vB(i))
                    .Ordre = mod_DataStructure.ToLong(vOr(i))
                    .IDFacture = id
                    .Typ = dTyp(cle)
                    .Operateur = dOp(cle)
                    .Libelle = mod_DataStructure.CellText(vL(i))
                    .Debut = d
                    .Fin = f
                    .Qte = mod_DataStructure.ToDouble(vQ(i))
                    .Unite = mod_DataStructure.CellText(vU(i))
                    .Prix = mod_DataStructure.ToDouble(vP(i))
                    .Tva = mod_DataStructure.ToDouble(vV(i))
                    .HT = mod_DataStructure.ToDouble(vH(i))
                    .TTC = mod_DataStructure.ToDouble(vC(i))
                End With
            End If
        End If
    Next i

    ChargerLignes = nb

End Function


' Vrai si l'unité convient à l'indicateur Consommation : kWh ou m³ (sans tenir compte des majuscules ; « m3 » est
' accepté). "norm" reçoit l'écriture normalisée (kWh, m³).
Private Function UniteConsommationValide(ByVal u As String, ByRef norm As String) As Boolean
    Select Case LCase$(Trim$(u))
        Case "kwh"
            norm = "kWh"
            UniteConsommationValide = True
        Case "m3", "m" & ChrW(179)
            norm = "m" & ChrW(179)
            UniteConsommationValide = True
    End Select
End Function

' Arrondi à 2 décimales comme la fonction ARRONDI d'Excel (le Round de VBA arrondit au pair).
Public Function Arrondi(ByVal x As Double) As Double
    Arrondi = Application.WorksheetFunction.Round(x, 2)
End Function


' =====================================================================================
' POINTS D'ENTRÉE (boutons)
' =====================================================================================
Public Sub SaisirFacture()

    Dim f As TypeFacture

    mod_Consommations.EnregistrerListesSiOuvertes
    If Not FacturesInstallees() Then
        MsgBox mod_Display.FR("Le suivi des consommations n'est pas install{e2}. Ex{e2}cutez d'abord la macro InstallerConsommations."), vbExclamation
        Exit Sub
    End If

    g_FacNb = 0
    ReDim g_FacLignes(1 To 1)
    f.ID = 0
    TerminerSaisie f

End Sub

' Modifie la facture sélectionnée dans la liste frm_Factures.
Public Sub ModifierFactureSelection()

    Dim lo As ListObject
    Dim id As Long
    Dim f As TypeFacture

    If Not FacturesInstallees() Then Exit Sub
    id = IdFactureSelectionnee(lo)
    If id = 0 Then Exit Sub
    If Not ChargerFacture(id, f) Then
        MsgBox mod_Display.FR("Cette facture est introuvable."), vbExclamation
        Exit Sub
    End If
    TerminerSaisie f

End Sub

' Copie la facture sélectionnée pour en saisir une nouvelle : mêmes type, opérateur, libellés, unités et taux
' de TVA ; numéro, dates, total, commentaire, quantités et prix sont laissés vides (ils changent à chaque facture).
' CONSULTATION : ouvre la facture sélectionnée comme « Modifier la sélection », mais le formulaire n'a qu'un bouton
' Sortir : tout ce qui aurait pu être modifié est abandonné en silence, rien n'est enregistré.
Public Sub ConsulterFactureSelection()

    Dim lo As ListObject
    Dim id As Long
    Dim f As TypeFacture

    If Not FacturesInstallees() Then Exit Sub
    id = IdFactureSelectionnee(lo)
    If id = 0 Then Exit Sub
    If Not ChargerFacture(id, f) Then
        MsgBox mod_Display.FR("Cette facture est introuvable."), vbExclamation
        Exit Sub
    End If

    g_FacConsultation = True
    OuvrirFormFacture f
    g_FacConsultation = False

End Sub

' Bouton Sortir du formulaire en consultation : referme sans rien enregistrer et sans demander de confirmation.
Public Sub FacSortirConsultation()

    Dim ws As Worksheet

    If Not g_FacEnCours Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_FACTURE)
    g_FacValide = False
    FermerFeuille ws

End Sub

' Boutons du formulaire selon le mode : en consultation, seul « Sortir » (et l'aide) reste disponible.
Private Sub PreparerBoutonsMode(ByVal ws As Worksheet)

    Dim nomBtn As Variant
    Dim gauche As Double
    Dim b As Long

    On Error Resume Next
    ws.Shapes("btnFacSortir").Delete
    On Error GoTo 0

    For Each nomBtn In Array("btnFacValider", "btnFacAnnuler", "btnFacAjouterBloc")
        On Error Resume Next
        If g_FacConsultation Then ws.Shapes(CStr(nomBtn)).Visible = 0 Else ws.Shapes(CStr(nomBtn)).Visible = -1
        On Error GoTo 0
    Next nomBtn

    If g_FacConsultation Then
        gauche = ws.Cells(1, 2).Left
        On Error Resume Next
        gauche = ws.Shapes("btnFacValider").Left
        On Error GoTo 0
        mod_InstallCommun.AjouterBoutonEntete ws, FF_LIGNE_BOUTONS, gauche, mod_InstallCommun.CapSortir(), _
            "FacSortirConsultation", "btnFacSortir", mod_InstallCommun.FRM_BTN_L
        ' Plus de boutons de structure (+, -, x) : la facture ne se modifie pas.
        On Error Resume Next
        For b = 1 To FF_MAX_BLOCS
            ws.Shapes("btnPlus" & b).Delete
            ws.Shapes("btnMoins" & b).Delete
            ws.Shapes("btnSuppr" & b).Delete
            ws.Shapes("btnIndic" & b).Delete
        Next b
        On Error GoTo 0
        AfficherMessage ws, "Consultation : rien de ce que vous modifiez ici n'est enregistr{e2}. Utilisez Sortir pour fermer."
    End If

End Sub

Public Sub CopierFactureSelection()

    Dim lo As ListObject
    Dim id As Long, k As Long
    Dim f As TypeFacture

    If Not FacturesInstallees() Then Exit Sub
    id = IdFactureSelectionnee(lo)
    If id = 0 Then Exit Sub
    If Not ChargerFacture(id, f) Then
        MsgBox mod_Display.FR("Cette facture est introuvable."), vbExclamation
        Exit Sub
    End If

    f.ID = 0
    f.Numero = ""
    f.DateFac = 0
    f.Debut = 0
    f.Fin = 0
    f.TotalSaisi = 0
    f.TotalRenseigne = False
    f.Commentaire = ""
    For k = 1 To g_FacNb
        g_FacLignes(k).IDFacture = 0
        g_FacLignes(k).Debut = 0
        g_FacLignes(k).Fin = 0
        g_FacLignes(k).Qte = 0
        g_FacLignes(k).Prix = 0
        g_FacLignes(k).HT = 0
        g_FacLignes(k).TTC = 0
        g_FacLignes(k).Copie = True
    Next k
    TerminerSaisie f

End Sub

Private Sub TerminerSaisie(ByRef f As TypeFacture)

    If OuvrirFormFacture(f) Then
        If EnregistrerFacture(g_FacRes) Then
            AppliquerChangementsIndicateurs g_FacRes.ID
            RafraichirListeFactures
            If ActiveSheet.Name = mod_VarGlobales.NOM_FEUILLE_CONSO Then mod_Consommations.ActualiserConsommations
            MsgBox mod_Display.FR("Facture enregistr{e2}e."), vbInformation, mod_Display.FR("Suivi des consommations")
        End If
    End If

End Sub

' Numéro de la facture sélectionnée dans la liste (0 + message si aucune).
Private Function IdFactureSelectionnee(ByRef lo As ListObject) As Long

    Dim inter As Range, c As Range
    Dim lignes As Object
    Dim n As Long
    Dim v As Variant

    Set lo = TableAffFactures()
    If lo Is Nothing Then Exit Function
    If ActiveSheet.Name <> mod_VarGlobales.NOM_FEUILLE_FACTURES Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    Set inter = SelectionVisible(lo)
    If inter Is Nothing Then
        MsgBox mod_Display.FR("S{e2}lectionnez d'abord une ligne du tableau."), vbInformation
        Exit Function
    End If

    ' Une seule ligne (avec un filtre actif, la sélection peut compter plusieurs zones).
    Set lignes = CreateObject("Scripting.Dictionary")
    For Each c In inter.Cells
        If Not lignes.Exists(c.Row) Then lignes.Add c.Row, True
    Next c
    If lignes.Count > 1 Then
        MsgBox mod_Display.FR("S{e2}lectionnez une seule facture."), vbInformation
        Exit Function
    End If

    n = inter.Cells(1, 1).Row - lo.DataBodyRange.Row + 1
    v = lo.ListColumns("ID_Facture").DataBodyRange.Cells(n, 1).Value2
    If IsNumeric(v) And Not IsEmpty(v) Then IdFactureSelectionnee = CLng(v)

End Function

' Cellules VISIBLES de la sélection dans le tableau de la liste. Avec un filtre actif, une sélection qui
' traverse des lignes masquées les contient aussi : sans cette précaution, une suppression toucherait
' des factures que l'opérateur ne voit pas.
Private Function SelectionVisible(ByVal lo As ListObject) As Range

    Dim inter As Range

    On Error Resume Next
    Set inter = Intersect(Selection, lo.DataBodyRange)
    On Error GoTo 0
    If inter Is Nothing Then Exit Function

    If inter.Cells.Count = 1 Then
        Set SelectionVisible = inter        ' SpecialCells sur une seule cellule porterait sur toute la feuille
    Else
        On Error Resume Next
        Set SelectionVisible = inter.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
    End If

End Function

' Charge une facture et ses lignes dans g_FacLignes / g_FacNb.
Private Function ChargerFacture(ByVal id As Long, ByRef f As TypeFacture) As Boolean

    Dim lo As ListObject
    Dim i As Long, n As Long, nat As Long
    Dim v As Variant
    Dim ord As Long, idxOrd As Long

    Set lo = TableFactures()
    If lo Is Nothing Then Exit Function
    If mod_Consommations.TableVide(lo) Then Exit Function

    n = lo.ListRows.Count
    For i = 1 To n
        v = lo.ListColumns("ID_Facture").DataBodyRange.Cells(i, 1).Value2
        If IsNumeric(v) And Not IsEmpty(v) Then
            If CLng(v) = id Then
                With lo.ListRows(i).Range
                    f.ID = id
                    f.Typ = mod_DataStructure.CellText(.Cells(1, mod_Consommations.ColIdx(lo, "TypeConso")).Value2)
                    f.Operateur = mod_DataStructure.CellText(.Cells(1, mod_Consommations.ColIdx(lo, "Operateur")).Value2)
                    f.Numero = mod_DataStructure.CellText(.Cells(1, mod_Consommations.ColIdx(lo, "NumeroFacture")).Value2)
                    f.DateFac = mod_Consommations.VersJour(.Cells(1, mod_Consommations.ColIdx(lo, "DateFacture")).Value2)
                    f.Debut = mod_Consommations.VersJour(.Cells(1, mod_Consommations.ColIdx(lo, "DebutPeriode")).Value2)
                    f.Fin = mod_Consommations.VersJour(.Cells(1, mod_Consommations.ColIdx(lo, "FinPeriode")).Value2)
                    v = .Cells(1, mod_Consommations.ColIdx(lo, "TotalFactureTTC")).Value2
                    If mod_Consommations.SaisieNonVide(v) Then
                        f.TotalSaisi = mod_DataStructure.ToDouble(v)
                        f.TotalRenseigne = True
                    End If
                    f.Commentaire = mod_DataStructure.CellText(.Cells(1, mod_Consommations.ColIdx(lo, "Commentaire")).Value2)
                End With
                ChargerFacture = True
                Exit For
            End If
        End If
    Next i
    If Not ChargerFacture Then Exit Function

    ' Lignes de la facture (tous blocs).
    g_FacNb = 0
    ReDim g_FacLignes(1 To FF_MAX_BLOCS * FF_MAX_LIGNES)
    Set lo = TableLignes()
    If lo Is Nothing Then Exit Function
    If mod_Consommations.TableVide(lo) Then Exit Function

    idxOrd = mod_Consommations.ColIdx(lo, "OrdreBloc")
    For i = 1 To lo.ListRows.Count
        v = lo.ListColumns("ID_Facture").DataBodyRange.Cells(i, 1).Value2
        If IsNumeric(v) And Not IsEmpty(v) Then
            If CLng(v) = id Then
                g_FacNb = g_FacNb + 1
                If g_FacNb > UBound(g_FacLignes) Then ReDim Preserve g_FacLignes(1 To g_FacNb + 20)
                With lo.ListRows(i).Range
                    g_FacLignes(g_FacNb).Indic = mod_DataStructure.CellText(.Cells(1, mod_Consommations.ColIdx(lo, "Indicateur")).Value2)
                    nat = FamilleDepuisTexte(g_FacLignes(g_FacNb).Indic)
                    If g_FacLignes(g_FacNb).Indic = "" Then g_FacLignes(g_FacNb).Indic = NomNature(nat)
                    g_FacLignes(g_FacNb).Nat = nat
                    g_FacLignes(g_FacNb).IDFacture = id
                    g_FacLignes(g_FacNb).Bloc = mod_DataStructure.CellText(.Cells(1, mod_Consommations.ColIdx(lo, "Bloc")).Value2)
                    g_FacLignes(g_FacNb).Libelle = mod_DataStructure.CellText(.Cells(1, mod_Consommations.ColIdx(lo, "Libelle")).Value2)
                    g_FacLignes(g_FacNb).Debut = mod_Consommations.VersJour(.Cells(1, mod_Consommations.ColIdx(lo, "DateDebut")).Value2)
                    g_FacLignes(g_FacNb).Fin = mod_Consommations.VersJour(.Cells(1, mod_Consommations.ColIdx(lo, "DateFin")).Value2)
                    g_FacLignes(g_FacNb).Qte = mod_DataStructure.ToDouble(.Cells(1, mod_Consommations.ColIdx(lo, "Quantite")).Value2)
                    g_FacLignes(g_FacNb).Unite = mod_DataStructure.CellText(.Cells(1, mod_Consommations.ColIdx(lo, "Unite")).Value2)
                    g_FacLignes(g_FacNb).Prix = mod_DataStructure.ToDouble(.Cells(1, mod_Consommations.ColIdx(lo, "PrixUnitaireHT")).Value2)
                    g_FacLignes(g_FacNb).Tva = mod_DataStructure.ToDouble(.Cells(1, mod_Consommations.ColIdx(lo, "TauxTVA")).Value2)
                    ord = 0
                    If idxOrd > 0 Then ord = mod_DataStructure.ToLong(.Cells(1, idxOrd).Value2)
                    If ord <= 0 Then ord = nat
                    g_FacLignes(g_FacNb).Ordre = ord
                    ' Un bloc sans nom (donnée ancienne) prend le nom du bloc du modèle de sa famille.
                    If g_FacLignes(g_FacNb).Bloc = "" Then
                        If nat <= NAT_TAXE Then g_FacLignes(g_FacNb).Bloc = NomBlocModele(nat) Else g_FacLignes(g_FacNb).Bloc = g_FacLignes(g_FacNb).Indic
                    End If
                End With
            End If
        End If
    Next i

End Function


' =====================================================================================
' ENREGISTREMENT
' =====================================================================================
Private Function EnregistrerFacture(ByRef f As TypeFacture) As Boolean

    Dim lo As ListObject
    Dim lr As ListRow
    Dim i As Long, nat As Long

    Set lo = TableFactures()
    If lo Is Nothing Then Exit Function

    ' Le tableau des lignes doit exister AVANT de supprimer quoi que ce soit : l'enregistrement d'une facture
    ' modifiée remplace ses anciennes lignes, un échec en cours de route les ferait perdre.
    If TableLignes() Is Nothing Then
        MsgBox mod_Display.FR("Le tableau des lignes de facture est introuvable : relancez InstallerConsommations. Rien n'a {e2}t{e2} enregistr{e2}."), vbExclamation
        Exit Function
    End If

    If f.ID = 0 Then
        f.ID = mod_Consommations.ProchainID(lo, "ID_Facture")
    Else
        SupprimerDonneesFacture f.ID
    End If

    Set lr = mod_Consommations.NouvelleLigne(lo)
    With lr.Range
        .Cells(1, mod_Consommations.ColIdx(lo, "ID_Facture")).Value = f.ID
        .Cells(1, mod_Consommations.ColIdx(lo, "TypeConso")).Value = f.Typ
        .Cells(1, mod_Consommations.ColIdx(lo, "Operateur")).Value = f.Operateur
        .Cells(1, mod_Consommations.ColIdx(lo, "NumeroFacture")).Value = f.Numero
        If f.DateFac > 0 Then .Cells(1, mod_Consommations.ColIdx(lo, "DateFacture")).Value = CDate(f.DateFac)
        If f.Debut > 0 Then .Cells(1, mod_Consommations.ColIdx(lo, "DebutPeriode")).Value = CDate(f.Debut)
        If f.Fin > 0 Then .Cells(1, mod_Consommations.ColIdx(lo, "FinPeriode")).Value = CDate(f.Fin)
        If f.TotalRenseigne Then .Cells(1, mod_Consommations.ColIdx(lo, "TotalFactureTTC")).Value = f.TotalSaisi
        .Cells(1, mod_Consommations.ColIdx(lo, "Commentaire")).Value = f.Commentaire
    End With

    Set lo = TableLignes()
    For i = 1 To g_FacNb
        Set lr = mod_Consommations.NouvelleLigne(lo)
        With lr.Range
            .Cells(1, mod_Consommations.ColIdx(lo, "ID_Facture")).Value = f.ID
            .Cells(1, mod_Consommations.ColIdx(lo, "OrdreBloc")).Value = g_FacLignes(i).Ordre
            .Cells(1, mod_Consommations.ColIdx(lo, "Bloc")).Value = g_FacLignes(i).Bloc
            If g_FacLignes(i).Indic = "" Then g_FacLignes(i).Indic = NomNature(g_FacLignes(i).Nat)
            .Cells(1, mod_Consommations.ColIdx(lo, "Indicateur")).Value = g_FacLignes(i).Indic
            .Cells(1, mod_Consommations.ColIdx(lo, "Libelle")).Value = g_FacLignes(i).Libelle
            .Cells(1, mod_Consommations.ColIdx(lo, "DateDebut")).Value = CDate(g_FacLignes(i).Debut)
            .Cells(1, mod_Consommations.ColIdx(lo, "DateFin")).Value = CDate(g_FacLignes(i).Fin)
            .Cells(1, mod_Consommations.ColIdx(lo, "Quantite")).Value = g_FacLignes(i).Qte
            .Cells(1, mod_Consommations.ColIdx(lo, "Unite")).Value = g_FacLignes(i).Unite
            .Cells(1, mod_Consommations.ColIdx(lo, "PrixUnitaireHT")).Value = g_FacLignes(i).Prix
            .Cells(1, mod_Consommations.ColIdx(lo, "TauxTVA")).Value = g_FacLignes(i).Tva
            .Cells(1, mod_Consommations.ColIdx(lo, "MontantHT")).Value = g_FacLignes(i).HT
            .Cells(1, mod_Consommations.ColIdx(lo, "MontantTTC")).Value = g_FacLignes(i).TTC
        End With
    Next i

    EnregistrerFacture = True

End Function

' Supprime la facture (TblFactures) et toutes ses lignes (TblLignes).
Private Sub SupprimerDonneesFacture(ByVal id As Long)

    Dim lo As ListObject
    Dim i As Long, nat As Long
    Dim v As Variant

    For nat = 0 To 1
        If nat = 0 Then Set lo = TableFactures() Else Set lo = TableLignes()
        If Not lo Is Nothing Then
            If Not mod_Consommations.TableVide(lo) Then
                For i = lo.ListRows.Count To 1 Step -1
                    v = lo.ListColumns("ID_Facture").DataBodyRange.Cells(i, 1).Value2
                    If IsNumeric(v) And Not IsEmpty(v) Then
                        If CLng(v) = id Then lo.ListRows(i).Delete
                    End If
                Next i
            End If
        End If
    Next nat

End Sub

Public Sub SupprimerFactureSelection()

    Dim lo As ListObject
    Dim inter As Range, c As Range
    Dim ids As Object
    Dim n As Long, i As Long, id As Long
    Dim v As Variant

    Set lo = TableAffFactures()
    If lo Is Nothing Then Exit Sub
    If ActiveSheet.Name <> mod_VarGlobales.NOM_FEUILLE_FACTURES Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Set inter = SelectionVisible(lo)
    If inter Is Nothing Then
        MsgBox mod_Display.FR("S{e2}lectionnez d'abord une ou plusieurs lignes du tableau."), vbInformation
        Exit Sub
    End If

    Set ids = CreateObject("Scripting.Dictionary")
    For Each c In inter.Cells
        n = c.Row - lo.DataBodyRange.Row + 1
        v = lo.ListColumns("ID_Facture").DataBodyRange.Cells(n, 1).Value2
        If IsNumeric(v) And Not IsEmpty(v) Then
            If Not ids.Exists(CLng(v)) Then ids.Add CLng(v), True
        End If
    Next c
    If ids.Count = 0 Then Exit Sub

    If MsgBox(ids.Count & mod_Display.FR(" facture(s) et toutes leurs lignes (abonnement, consommation, taxes) {a2} supprimer ?"), _
              vbYesNo + vbQuestion, mod_Display.FR("Supprimer des factures")) = vbNo Then Exit Sub

    For Each v In ids.Keys
        SupprimerDonneesFacture CLng(v)
    Next v
    RafraichirListeFactures

End Sub


' =====================================================================================
' LISTE DES FACTURES (frm_Factures) : copie de lecture des données, recalculée à l'ouverture
' =====================================================================================
' Calcule le total HT et TTC d'une facture à partir de ses lignes.
Private Sub TotauxFacture(ByRef t() As TypeLigne, ByVal nb As Long, ByVal id As Long, ByRef ht As Double, ByRef ttc As Double)
    Dim i As Long
    ht = 0: ttc = 0
    For i = 1 To nb
        If t(i).IDFacture = id Then
            ht = ht + t(i).HT
            ttc = ttc + t(i).TTC
        End If
    Next i
End Sub

Public Sub RafraichirListeFactures()

    Dim src As ListObject, dst As ListObject
    Dim t() As TypeLigne
    Dim nbL As Long, n As Long, i As Long, p As Long, s As Long
    Dim donnees() As Variant
    Dim ht As Double, ttc As Double
    Dim id As Long
    Dim vTot As Variant
    Dim nbCol As Long
    Dim ordre() As Long
    Dim filtres As Variant

    Set src = TableFactures()
    Set dst = TableAffFactures()
    If src Is Nothing Or dst Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    ' Les filtres de colonne (flèches d'Excel) sont mémorisés puis levés pendant la reconstruction du tableau,
    ' et réappliqués à la fin : l'opérateur retrouve sa vue filtrée après chaque rafraîchissement.
    If Not dst.ShowAutoFilter Then dst.ShowAutoFilter = True
    filtres = MemoriserFiltres(dst)
    On Error Resume Next
    If dst.AutoFilter.FilterMode Then dst.AutoFilter.ShowAllData
    On Error GoTo 0

    Do While dst.ListRows.Count > 1
        dst.ListRows(dst.ListRows.Count).Delete
    Loop
    dst.DataBodyRange.ClearContents

    If Not mod_Consommations.TableVide(src) Then
        nbL = ChargerLignes(t, 0, True)
        n = src.ListRows.Count
        OrdreParDateDecroissante src, ordre
        If n > 1 Then dst.Resize dst.Range.Resize(n + 1)
        nbCol = dst.ListColumns.Count
        ReDim donnees(1 To n, 1 To nbCol)
        For i = 1 To n
            s = ordre(i)
            id = 0
            If IsNumeric(src.ListColumns("ID_Facture").DataBodyRange.Cells(s, 1).Value2) Then
                id = mod_DataStructure.ToLong(src.ListColumns("ID_Facture").DataBodyRange.Cells(s, 1).Value2)
            End If
            TotauxFacture t, nbL, id, ht, ttc
            For p = 1 To nbCol
                Select Case dst.ListColumns(p).Name
                    Case "TotalHT": donnees(i, p) = ht
                    Case "TotalTVA": donnees(i, p) = ttc - ht
                    Case "TotalTTCCalcule": donnees(i, p) = ttc
                    Case "Ecart"
                        vTot = src.ListColumns("TotalFactureTTC").DataBodyRange.Cells(s, 1).Value2
                        If mod_Consommations.SaisieNonVide(vTot) Then donnees(i, p) = Arrondi(mod_DataStructure.ToDouble(vTot) - ttc)
                    Case Else
                        If mod_Consommations.ColIdx(src, dst.ListColumns(p).Name) > 0 Then
                            donnees(i, p) = src.ListColumns(dst.ListColumns(p).Name).DataBodyRange.Cells(s, 1).Value2
                        End If
                End Select
            Next p
        Next i
        dst.DataBodyRange.Value2 = donnees
    End If

    mod_Consommations.MettreEnFormeListeFactures dst
    RestaurerFiltres dst, filtres
    Application.ScreenUpdating = True

End Sub

' Numéros de ligne du tableau source classés par DateFacture décroissante (les plus récentes en haut) ;
' les factures sans date en bas, à égalité le numéro de facture le plus grand d'abord.
Private Sub OrdreParDateDecroissante(ByVal src As ListObject, ByRef ordre() As Long)

    Dim n As Long, i As Long, j As Long, k As Long
    Dim dates() As Double, ids() As Double
    Dim vD As Variant, vI As Variant

    n = src.ListRows.Count
    ReDim ordre(1 To n)
    ReDim dates(1 To n)
    ReDim ids(1 To n)
    vD = mod_Consommations.LireColonne(src, "DateFacture")
    vI = mod_Consommations.LireColonne(src, "ID_Facture")
    For i = 1 To n
        dates(i) = mod_Consommations.VersJour(vD(i))
        If IsNumeric(vI(i)) And Not IsEmpty(vI(i)) Then ids(i) = CDbl(vI(i))
        ordre(i) = i
    Next i

    ' Tri par insertion (quelques centaines de factures au plus).
    For i = 2 To n
        k = ordre(i)
        j = i - 1
        Do While j >= 1
            If dates(ordre(j)) > dates(k) Or (dates(ordre(j)) = dates(k) And ids(ordre(j)) > ids(k)) Then Exit Do
            ordre(j + 1) = ordre(j)
            j = j - 1
        Loop
        ordre(j + 1) = k
    Next i

End Sub

' Mémorise les filtres de colonne actifs d'un tableau : une colonne du résultat par filtre
' (champ, opérateur, critère 1, critère 2) ; Empty si aucun. Même principe que dans mod_RechercheOperations.
Private Function MemoriserFiltres(ByVal tbl As ListObject) As Variant

    Dim res() As Variant
    Dim nb As Long, f As Long
    Dim fc As Filter

    If Not tbl.ShowAutoFilter Then Exit Function
    For f = 1 To tbl.ListColumns.Count
        Set fc = Nothing
        On Error Resume Next
        Set fc = tbl.AutoFilter.Filters(f)
        On Error GoTo 0
        If Not fc Is Nothing Then
            If fc.On Then
                nb = nb + 1
                ReDim Preserve res(1 To 4, 1 To nb)
                res(1, nb) = f
                res(2, nb) = fc.Operator
                res(3, nb) = fc.Criteria1
                res(4, nb) = Empty
                On Error Resume Next
                res(4, nb) = fc.Criteria2
                On Error GoTo 0
            End If
        End If
    Next f
    If nb > 0 Then MemoriserFiltres = res

End Function

Private Sub RestaurerFiltres(ByVal tbl As ListObject, ByVal filtres As Variant)

    Dim f As Long
    Dim crit As Variant
    Dim op As Long

    If IsEmpty(filtres) Then Exit Sub
    On Error Resume Next
    For f = 1 To UBound(filtres, 2)
        crit = filtres(3, f)
        op = filtres(2, f)
        If op = xlFilterValues Then
            If Not IsArray(crit) Then crit = Array(crit)
        End If
        If op = 0 Then
            tbl.Range.AutoFilter Field:=filtres(1, f), Criteria1:=crit
        ElseIf IsEmpty(filtres(4, f)) Then
            tbl.Range.AutoFilter Field:=filtres(1, f), Criteria1:=crit, Operator:=op
        Else
            tbl.Range.AutoFilter Field:=filtres(1, f), Criteria1:=crit, Operator:=op, Criteria2:=filtres(4, f)
        End If
    Next f
    On Error GoTo 0

End Sub


' =====================================================================================
' FORMULAIRE frm_Facture : OUVERTURE, VALIDATION, FERMETURE
' =====================================================================================
' Ouvre le formulaire prérempli avec "f" et g_FacLignes, et attend sa fermeture.
' Renvoie Vrai si l'opérateur a validé (le résultat est dans g_FacRes / g_FacLignes).
Private Function OuvrirFormFacture(ByRef f As TypeFacture) As Boolean

    Dim ws As Worksheet

    On Error GoTo Erreur

    Set ws = mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_FACTURE)
    If ws Is Nothing Then
        MsgBox mod_Display.FR("La feuille '") & mod_VarGlobales.NOM_FEUILLE_FACTURE & mod_Display.FR("' est introuvable.") & vbCrLf & _
               mod_Display.FR("Ex{e2}cutez d'abord la macro InstallerConsommations."), vbExclamation
        Exit Function
    End If

    g_FacNomFeuillePrec = ActiveSheet.Name
    g_FacValide = False
    g_IdEnEdition = f.ID
    Set g_Pend = CreateObject("Scripting.Dictionary")
    ws.Visible = xlSheetVisible
    ws.Activate
    mod_InstallCommun.MasquerQuadrillage
    Application.ScreenUpdating = False
    RemplirFormulaire ws, f
    PreparerBoutonsMode ws
    Application.ScreenUpdating = True
    ws.Range(FF_ADR_TYPE).Select


    Application.EnableEvents = True
    ' La saisie semi-automatique d'Excel complèterait "M" en "m³" (unité déjà présente dans la colonne) :
    ' elle est coupée pendant la saisie de la facture puis rétablie telle qu'elle était.
    g_AutoCompleteAvant = Application.EnableAutoComplete
    Application.EnableAutoComplete = False
    g_FacEnCours = True
    g_FacVerrouActif = True
    Do While g_FacEnCours
        DoEvents
    Loop
    Application.EnableAutoComplete = g_AutoCompleteAvant

    On Error Resume Next
    mod_InstallCommun.ActiverFeuille g_FacNomFeuillePrec
    On Error GoTo 0

    OuvrirFormFacture = g_FacValide
    Exit Function

Erreur:
    Application.EnableEvents = True
    Application.EnableAutoComplete = g_AutoCompleteAvant
    g_FacEnCours = False
    g_FacVerrouActif = False
    MsgBox mod_Display.FR("Erreur inattendue dans la saisie de la facture :") & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
    OuvrirFormFacture = False

End Function

Private Sub RemplirFormulaire(ByVal ws As Worksheet, ByRef f As TypeFacture)

    Dim evenementsAvant As Boolean
    Dim numErr As Long
    Dim k As Long, b As Long, r As Long
    Dim compte(1 To FF_MAX_BLOCS) As Long
    Dim ordMin(1 To FF_MAX_BLOCS) As Long
    Dim nat As Long

    evenementsAvant = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Sortie

    ws.Range(FF_ADR_TYPE).Value = f.Typ
    ws.Range(FF_ADR_OPERATEUR).Value = f.Operateur
    ws.Range(FF_ADR_NUMERO).Value = f.Numero
    EcrireDate ws.Range(FF_ADR_DATE), f.DateFac
    EcrireDate ws.Range(FF_ADR_DEBUT), f.Debut
    EcrireDate ws.Range(FF_ADR_FIN), f.Fin
    If f.TotalRenseigne Then ws.Range(FF_ADR_TOTAL).Value = f.TotalSaisi Else ws.Range(FF_ADR_TOTAL).ClearContents
    ws.Range(FF_ADR_COMMENT).Value = f.Commentaire
    ws.Range(FF_ADR_MESSAGE).Value = ""
    PoserListeType ws
    MettreAJourListeOperateurs ws, f.Typ

    ' Structure des blocs :
    '  - nouvelle facture (aucune ligne) : le MODÈLE, c'est-à-dire les trois blocs de départ de FF_NB_DEFAUT lignes ;
    '  - facture existante (modification ou copie) : EXACTEMENT la structure de la source, c'est-à-dire seulement
    '    les blocs qui contiennent des lignes, dans l'ordre où ils étaient affichés (colonne OrdreBloc), avec
    '    autant de lignes que la source.
    g_NbBlocs = 0
    For b = 1 To FF_MAX_BLOCS
        g_NomBloc(b) = ""
        g_Nat(b) = 0
        g_Indic(b) = ""
        g_N(b) = 0
        compte(b) = 0
        ordMin(b) = 0
    Next b
    If g_FacNb = 0 Then
        g_NbBlocs = 3
        For b = 1 To 3
            g_Nat(b) = b
            g_Indic(b) = NomNature(b)
            g_NomBloc(b) = NomBlocModele(b)
            g_N(b) = FF_NB_DEFAUT
        Next b
    Else
        ' Un bloc par (nature, nom) rencontré dans les lignes ; sa position est la plus petite de ses lignes.
        For k = 1 To g_FacNb
            b = BlocDeLigne(g_FacLignes(k), True)
            If b > 0 Then
                compte(b) = compte(b) + 1
                If ordMin(b) = 0 Or g_FacLignes(k).Ordre < ordMin(b) Then ordMin(b) = g_FacLignes(k).Ordre
            End If
        Next k
        TrierBlocs ordMin, compte
        For b = 1 To g_NbBlocs
            g_N(b) = PlusGrand(compte(b), 1)
            compte(b) = 0
        Next b
    End If
    DessinerBlocs ws

    ' Les lignes de la facture sont réparties dans les blocs.
    For k = 1 To g_FacNb
        b = BlocDeLigne(g_FacLignes(k), False)
        If b > 0 Then
            compte(b) = compte(b) + 1
            r = BlocL1(b) + compte(b) - 1
            ws.Cells(r, FF_COL_LIB).Value = g_FacLignes(k).Libelle
            ws.Cells(r, FF_COL_UNITE).Value = g_FacLignes(k).Unite
            ws.Cells(r, FF_COL_TVA).Value = g_FacLignes(k).Tva
            If Not g_FacLignes(k).Copie Then
                ws.Cells(r, FF_COL_DEB).Value = CDate(g_FacLignes(k).Debut)
                ws.Cells(r, FF_COL_FIN).Value = CDate(g_FacLignes(k).Fin)
                ws.Cells(r, FF_COL_QTE).Value = g_FacLignes(k).Qte
                ws.Cells(r, FF_COL_PRIX).Value = g_FacLignes(k).Prix
            End If
        End If
    Next k

    RemplirTvaDefaut ws, f.Typ

    mod_InstallCommun.PoserValidationDate ws.Range(FF_ADR_DATE)
    mod_InstallCommun.PoserValidationDate ws.Range(FF_ADR_DEBUT)
    mod_InstallCommun.PoserValidationDate ws.Range(FF_ADR_FIN)
    mod_InstallCommun.PoserValidationMontant ws.Range(FF_ADR_TOTAL), mod_InstallCommun.MONTANT_POSITIF_OU_NUL

Sortie:
    numErr = Err.Number
    Application.EnableEvents = evenementsAvant
    If numErr <> 0 Then Err.Raise numErr, "RemplirFormulaire", Err.Description

End Sub

' Liste déroulante du TYPE : posée directement sur la plage technique de la feuille (colonne masquée),
' sans passer par un nom défini du classeur (un nom cassé ferait disparaître la liste).
Private Sub PoserListeType(ByVal ws As Worksheet)

    Dim tps As Variant
    Dim i As Long

    tps = mod_Consommations.TypesConso()
    ws.Columns(FF_COL_AIDE_TYPE).Hidden = True
    ws.Range(ws.Cells(1, FF_COL_AIDE_TYPE), ws.Cells(20, FF_COL_AIDE_TYPE)).ClearContents
    ws.Cells(1, FF_COL_AIDE_TYPE).Value = "Types"
    For i = 0 To UBound(tps)
        ws.Cells(2 + i, FF_COL_AIDE_TYPE).Value = tps(i)
    Next i

    On Error Resume Next
    With ws.Range(FF_ADR_TYPE).MergeArea.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="=" & ws.Range(ws.Cells(2, FF_COL_AIDE_TYPE), ws.Cells(2 + UBound(tps), FF_COL_AIDE_TYPE)).Address(True, True)
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
        .ErrorTitle = "Type"
        .ErrorMessage = mod_Display.FR("Choisissez le type de consommation dans la liste.")
    End With
    On Error GoTo 0

End Sub

' Liste déroulante de l'opérateur : les fournisseurs dont le contrat n'est pas terminé pour le type choisi
' (tous les types tant qu'aucun type n'est choisi). Aucune saisie hors liste : un nouveau fournisseur doit
' d'abord faire l'objet d'un contrat (bouton Contrats de l'écran Consommations). Même principe que le type :
' plage de la feuille, pas de nom défini.
Private Sub MettreAJourListeOperateurs(ByVal ws As Worksheet, ByVal typ As String)

    Dim noms() As String
    Dim n As Long, i As Long
    Dim cel As Range
    Dim msg As String

    Set cel = ws.Range(FF_ADR_OPERATEUR)
    ws.Columns(FF_COL_AIDE_OPER).Hidden = True
    ws.Range(ws.Cells(1, FF_COL_AIDE_OPER), ws.Cells(60, FF_COL_AIDE_OPER)).ClearContents

    msg = mod_Display.FR("Cet op{e2}rateur n'a pas de contrat en cours pour ce type de consommation. ") & _
          mod_Display.FR("Pour utiliser un nouveau fournisseur, commencez par saisir un nouveau contrat ") & _
          mod_Display.FR("(bouton Contrats de l'{e2}cran Consommations), puis revenez saisir la facture.")

    n = mod_Consommations.OperateursEnCours(typ, noms)
    If n > 0 Then
        ws.Cells(1, FF_COL_AIDE_OPER).Value = "Operateurs"
        For i = 1 To n
            ws.Cells(1 + i, FF_COL_AIDE_OPER).Value = noms(i)
        Next i
    End If

    On Error Resume Next
    With cel.MergeArea.Validation
        .Delete
        If n > 0 Then
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Formula1:="=" & ws.Range(ws.Cells(2, FF_COL_AIDE_OPER), ws.Cells(1 + n, FF_COL_AIDE_OPER)).Address(True, True)
            .InCellDropdown = True
        Else
            ' Aucun contrat en cours : toute saisie est refusée avec le même message.
            .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Formula1:="=FALSE"
        End If
        .IgnoreBlank = True
        .ShowError = True
        .ErrorTitle = mod_Display.FR("Op{e2}rateur inconnu")
        .ErrorMessage = msg
    End With
    On Error GoTo 0

End Sub

' Unités proposées dans la liste déroulante : unités usuelles + toutes celles déjà utilisées dans les lignes
' enregistrées. Une unité nouvelle peut être saisie (après confirmation) : elle sera proposée ensuite.
Private Sub MettreAJourListeUnites(ByVal ws As Worksheet)

    Dim base As Variant
    Dim t() As TypeLigne
    Dim nb As Long, i As Long, j As Long, n As Long
    Dim liste() As String
    Dim deja As Boolean
    Dim u As String

    base = Array("kWh", "m" & ChrW(179), ChrW(8364), "MWh", "L", "jours", "mois", "M", "MM")
    ReDim liste(1 To UBound(base) + 1 + 60)
    For i = 0 To UBound(base)
        n = n + 1
        liste(n) = CStr(base(i))
    Next i

    nb = ChargerLignes(t, 0, True)
    For i = 1 To nb
        u = t(i).Unite
        If u <> "" And n < UBound(liste) Then
            deja = False
            For j = 1 To n
                If StrComp(liste(j), u, vbTextCompare) = 0 Then deja = True
            Next j
            If Not deja Then
                n = n + 1
                liste(n) = u
            End If
        End If
    Next i

    ws.Columns(FF_COL_AIDE_UNITE).Hidden = True
    ws.Range(ws.Cells(1, FF_COL_AIDE_UNITE), ws.Cells(UBound(liste) + 1, FF_COL_AIDE_UNITE)).ClearContents
    ws.Cells(1, FF_COL_AIDE_UNITE).Value = "Unites"
    For i = 1 To n
        ws.Cells(1 + i, FF_COL_AIDE_UNITE).Value = liste(i)
    Next i
    g_NbUnites = n

End Sub

' Taux de TVA proposés pour un type, bloc par bloc : celui de la dernière ligne enregistrée du même type dans le
' même bloc, sinon un taux usuel (eau : 5,5 % ; électricité et gaz : abonnement 5,5 %, consommation et taxes
' 20 % ; bloc libre : 20 %). tv(b) vaut -1 tant qu'aucun type n'est choisi.
Private Sub TvaDefauts(ByVal typ As String, ByRef tv() As Double)

    Dim t() As TypeLigne
    Dim nb As Long, i As Long, b As Long
    Dim fin(1 To FF_MAX_BLOCS) As Double
    Dim eau As Boolean

    For b = 1 To FF_MAX_BLOCS
        tv(b) = -1
    Next b
    If typ = "" Then Exit Sub

    eau = (StrComp(typ, mod_Consommations.TypeAdduction(), vbTextCompare) = 0 Or _
           StrComp(typ, mod_Consommations.TypeRecyclage(), vbTextCompare) = 0)
    For b = 1 To g_NbBlocs
        Select Case g_Nat(b)
            Case NAT_ABO
                tv(b) = 0.055
            Case NAT_CONSO, NAT_TAXE
                If eau Then tv(b) = 0.055 Else tv(b) = 0.2
            Case Else
                tv(b) = 0.2
        End Select
    Next b

    nb = ChargerLignes(t, 0, True)
    For i = 1 To nb
        If StrComp(t(i).Typ, typ, vbTextCompare) = 0 Then
            b = BlocDeLigne(t(i), False)
            If b > 0 Then
                If t(i).Fin >= fin(b) Then
                    fin(b) = t(i).Fin
                    tv(b) = t(i).Tva
                End If
            End If
        End If
    Next i

End Sub

' Remplit le taux de TVA des lignes qui n'ont encore AUCUNE autre saisie (une ligne déjà commencée garde son taux).
Private Sub RemplirTvaDefaut(ByVal ws As Worksheet, ByVal typ As String)

    Dim tv(1 To FF_MAX_BLOCS) As Double
    Dim b As Long, i As Long, r As Long
    Dim uniteCourante As String

    TvaDefauts typ, tv
    For b = 1 To g_NbBlocs
        If tv(b) >= 0 Then
            For i = 0 To BlocN(b) - 1
                r = BlocL1(b) + i
                If LigneVide(ws, r) Then
                    ws.Cells(r, FF_COL_TVA).Value = tv(b)
                    uniteCourante = mod_DataStructure.CellText(ws.Cells(r, FF_COL_UNITE).Value)
                    If g_Nat(b) = NAT_ABO And uniteCourante = "" Then ws.Cells(r, FF_COL_UNITE).Value = "jours"
                    ' Consommation : m³ pour l'eau, kWh pour l'électricité et le gaz ; une unité par défaut précédente est remplacée
                    ' quand on change de type.
                    If g_Nat(b) = NAT_CONSO And (uniteCourante = "" Or StrComp(uniteCourante, "kWh", vbTextCompare) = 0 Or uniteCourante = "m" & ChrW(179)) Then
                        ws.Cells(r, FF_COL_UNITE).Value = mod_Consommations.UniteDefaut(typ)
                    End If
                End If
            Next i
        End If
    Next b

End Sub

' Trie les blocs affichés par position croissante (ordre d'enregistrement de la facture). À position égale
' (anciennes données), la nature fait foi : abonnement, consommation, taxes, blocs libres.
Private Sub TrierBlocs(ByRef ordre() As Long, ByRef compte() As Long)
    Dim i As Long, j As Long
    Dim o As Long, c As Long, n As Long, nom As String, ind As String
    For i = 2 To g_NbBlocs
        o = ordre(i): c = compte(i): n = g_Nat(i): nom = g_NomBloc(i): ind = g_Indic(i)
        j = i - 1
        Do While j >= 1
            If ordre(j) < o Or (ordre(j) = o And g_Nat(j) <= n) Then Exit Do
            ordre(j + 1) = ordre(j)
            compte(j + 1) = compte(j)
            g_Nat(j + 1) = g_Nat(j)
            g_Indic(j + 1) = g_Indic(j)
            g_NomBloc(j + 1) = g_NomBloc(j)
            j = j - 1
        Loop
        ordre(j + 1) = o
        compte(j + 1) = c
        g_Nat(j + 1) = n
        g_Indic(j + 1) = ind
        g_NomBloc(j + 1) = nom
    Next i
End Sub

Private Function PlusGrand(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then PlusGrand = a Else PlusGrand = b
End Function

' =====================================================================================
' BLOCS DE LIGNES : GÉOMÉTRIE ET CORRESPONDANCE AVEC LES LIGNES DE FACTURE
' =====================================================================================
' Un "bloc" du formulaire porte un numéro b : 1 = abonnement, 2 = consommation, 3 = taxes (blocs fixes, modèle
' de toute nouvelle facture), puis 4 et suivants = blocs LIBRES nommés par l'opérateur (bouton "Ajouter un bloc").
Private Function NatDuBloc(ByVal b As Long) As Long
    NatDuBloc = g_Nat(b)
End Function

Private Function NomDuBloc(ByVal b As Long) As String
    NomDuBloc = g_NomBloc(b)
End Function

' Numéro du bloc qui porte une ligne de facture (0 si introuvable) : un bloc se reconnaît à son NOM (unique dans
' une facture, sans tenir compte des majuscules). Si "creer", le bloc manquant est ajouté à la fin.
Private Function BlocDeLigne(ByRef l As TypeLigne, ByVal creer As Boolean) As Long
    Dim b As Long
    For b = 1 To g_NbBlocs
        If StrComp(g_NomBloc(b), l.Bloc, vbTextCompare) = 0 Then
            BlocDeLigne = b
            Exit Function
        End If
    Next b
    If creer And g_NbBlocs < FF_MAX_BLOCS Then
        g_NbBlocs = g_NbBlocs + 1
        g_Nat(g_NbBlocs) = l.Nat
        g_Indic(g_NbBlocs) = l.Indic
        g_NomBloc(g_NbBlocs) = l.Bloc
        BlocDeLigne = g_NbBlocs
    End If
End Function

' Titre affiché d'un bloc : son nom, SUIVI SYSTÉMATIQUEMENT entre parenthèses de l'indicateur qu'il alimente, y compris
' pour les blocs du modèle (par exemple « Abonnement (Abonnement) ») : l'opérateur voit toujours ce qui est impacté.
Private Function LibelleBloc(ByVal b As Long) As String
    Dim indic As String
    indic = g_Indic(b)
    If indic = "" Then indic = NomNature(g_Nat(b))
    LibelleBloc = g_NomBloc(b) & " (" & indic & ")"
End Function

' Blocs du modèle (nouvelle facture) : trois blocs de FF_NB_DEFAUT lignes. Appelé par l'installateur.
Public Sub DessinerBlocsParDefaut(ByVal ws As Worksheet)
    Dim b As Long
    g_NbBlocs = 3
    For b = 1 To FF_MAX_BLOCS
        If b <= 3 Then
            g_Nat(b) = b
            g_Indic(b) = NomNature(b)
            g_NomBloc(b) = NomBlocModele(b)
            g_N(b) = FF_NB_DEFAUT
        Else
            g_Nat(b) = 0
            g_Indic(b) = ""
            g_NomBloc(b) = ""
            g_N(b) = 0
        End If
    Next b
    DessinerBlocs ws
End Sub

' (Re)dessine tous les blocs d'après g_NbBlocs, g_N et g_NomBloc : zone effacée puis reconstruite.
' Les valeurs saisies sont perdues : l'appelant les relit avant et les réécrit après (voir ModifierStructure).
Private Sub DessinerBlocs(ByVal ws As Worksheet)

    Dim b As Long, i As Long, r As Long, titre As Long, l1 As Long, n As Long
    Dim entetes As Variant
    Dim zone As Range
    Dim ht As String, ttc As String
    Dim derniere As Long

    For b = 1 To g_NbBlocs
        g_N(b) = Borne(g_N(b))
    Next b

    ws.Columns(FF_COL_LIB).ColumnWidth = FF_LARG_LIBELLE
    MettreAJourListeUnites ws

    ' Boutons de la version précédente du dessin.
    On Error Resume Next
    For b = 1 To FF_MAX_BLOCS
        ws.Shapes("btnPlus" & b).Delete
        ws.Shapes("btnMoins" & b).Delete
        ws.Shapes("btnSuppr" & b).Delete
        ws.Shapes("btnIndic" & b).Delete
    Next b
    On Error GoTo 0

    ' Zone des blocs : colonnes A à K seulement (les listes techniques masquées sont plus à droite).
    derniere = FF_BLOCS_DEBUT + FF_MAX_BLOCS * (FF_MAX_LIGNES + 4)
    Set zone = ws.Range(ws.Cells(FF_BLOCS_DEBUT, 1), ws.Cells(derniere, FF_COL_TTC + 1))
    zone.UnMerge
    zone.Clear
    On Error Resume Next
    zone.Validation.Delete
    On Error GoTo 0
    zone.Font.Name = "Calibri"
    zone.Font.Size = 10
    zone.VerticalAlignment = xlCenter
    ws.Range(ws.Rows(FF_BLOCS_DEBUT), ws.Rows(derniere)).RowHeight = mod_InstallCommun.FRM_H_LIGNE

    entetes = Array("Libell" & ChrW(233), "D" & ChrW(233) & "but", "Fin", "Quantit" & ChrW(233), "Unit" & ChrW(233), "Prix HT", "TVA", "Montant HT", "Montant TTC")
    ht = mod_Consommations.AdresseCol(FF_COL_HT)
    ttc = mod_Consommations.AdresseCol(FF_COL_TTC)

    For b = 1 To g_NbBlocs
        titre = BlocTitre(b)
        l1 = BlocL1(b)
        n = g_N(b)

        mod_InstallCommun.PoserTitreBloc ws.Range(ws.Cells(titre, FF_COL_LIB), ws.Cells(titre, FF_COL_TTC)), LibelleBloc(b)

        For i = 0 To 8
            ws.Cells(titre + 1, FF_COL_LIB + i).Value = entetes(i)
        Next i
        mod_InstallCommun.PoserEnteteTableau ws.Range(ws.Cells(titre + 1, FF_COL_LIB), ws.Cells(titre + 1, FF_COL_TTC))
        ws.Range(ws.Cells(titre + 1, FF_COL_QTE), ws.Cells(titre + 1, FF_COL_TTC)).HorizontalAlignment = xlRight
        ws.Cells(titre + 1, FF_COL_UNITE).HorizontalAlignment = xlLeft

        For i = 0 To n - 1
            r = l1 + i
            mod_InstallCommun.PoserChampSaisie ws.Cells(r, FF_COL_LIB), "@"
            mod_InstallCommun.PoserChampDate ws.Cells(r, FF_COL_DEB)
            mod_InstallCommun.PoserChampDate ws.Cells(r, FF_COL_FIN)
            ' Tous les blocs ont les mêmes colonnes : quantité et unité se saisissent partout.
            mod_InstallCommun.PoserChampMontant ws.Cells(r, FF_COL_QTE), "#,##0.00", mod_InstallCommun.MONTANT_POSITIF_OU_NUL
            mod_InstallCommun.PoserChampSaisie ws.Cells(r, FF_COL_UNITE), "@"
            mod_InstallCommun.PoserChampSaisie ws.Cells(r, FF_COL_PRIX), "#,##0.00000"      ' contrôle : PoserGardeFousBloc (prix négatif = remboursement)
            mod_InstallCommun.PoserChampSaisie ws.Cells(r, FF_COL_TVA), "0.0%"
            mod_InstallCommun.PoserChampLecture ws.Cells(r, FF_COL_HT), "#,##0.00"
            ws.Cells(r, FF_COL_HT).Formula = FormuleHT(r)
            mod_InstallCommun.PoserChampLecture ws.Cells(r, FF_COL_TTC), "#,##0.00"
            ws.Cells(r, FF_COL_TTC).Formula = FormuleTTC(r)
        Next i

        ' Sous-total du bloc, puis une fine ligne de séparation.
        r = l1 + n
        mod_InstallCommun.PoserEtiquette ws.Cells(r, FF_COL_LIB), "Sous-total"
        mod_InstallCommun.PoserChampLecture ws.Cells(r, FF_COL_HT), "#,##0.00"
        ws.Cells(r, FF_COL_HT).Formula = "=SUM(" & ht & l1 & ":" & ht & (l1 + n - 1) & ")"
        mod_InstallCommun.PoserChampLecture ws.Cells(r, FF_COL_TTC), "#,##0.00"
        ws.Cells(r, FF_COL_TTC).Formula = "=SUM(" & ttc & l1 & ":" & ttc & (l1 + n - 1) & ")"
        ws.Range(ws.Cells(r, FF_COL_HT), ws.Cells(r, FF_COL_TTC)).Font.Bold = True
        ws.Rows(r + 1).RowHeight = mod_InstallCommun.FRM_H_SEP

        ' Boutons du titre : "+" (ajoute une ligne), "-" (retire la dernière ligne si elle est vide) et "x"
        ' (supprime le bloc, y compris un bloc du modèle : seule la facture en cours est concernée). Une même macro sert à tous les blocs : le numéro est lu dans le nom du bouton.
        mod_InstallCommun.AjouterBoutonCellule ws, ws.Cells(titre, FF_COL_TTC), "+", "FacAjouterLigne", "btnPlus" & b, 24, 34
        mod_InstallCommun.AjouterBoutonCellule ws, ws.Cells(titre, FF_COL_TTC), ChrW(8722), "FacRetirerLigne", "btnMoins" & b, 24, 6
        mod_InstallCommun.AjouterBoutonCellule ws, ws.Cells(titre, FF_COL_HT), ChrW(215), "FacSupprimerBloc", "btnSuppr" & b, 24, 34
        mod_InstallCommun.AjouterBoutonCellule ws, ws.Cells(titre, FF_COL_PRIX), "Indicateur", "FacModifierIndicateur", "btnIndic" & b, 66, 2
    Next b

    PoserFormulesRecap ws
    For b = 1 To g_NbBlocs
        PoserGardeFousBloc ws, b
    Next b

End Sub

Private Function Borne(ByVal n As Long) As Long
    If n < 1 Then
        Borne = 1
    ElseIf n > FF_MAX_LIGNES Then
        Borne = FF_MAX_LIGNES
    Else
        Borne = n
    End If
End Function

' Formules du récapitulatif (en haut) : elles pointent vers les sous-totaux de TOUS les blocs, dont la position varie.
Private Sub PoserFormulesRecap(ByVal ws As Worksheet)

    Dim b As Long
    Dim fHT As String, fTTC As String
    Dim ht As String, ttc As String

    ht = mod_Consommations.AdresseCol(FF_COL_HT)
    ttc = mod_Consommations.AdresseCol(FF_COL_TTC)
    fHT = "="
    fTTC = "="
    If g_NbBlocs = 0 Then
        fHT = "=0"
        fTTC = "=0"
    End If
    For b = 1 To g_NbBlocs
        If b > 1 Then
            fHT = fHT & "+"
            fTTC = fTTC & "+"
        End If
        fHT = fHT & "N(" & ht & (BlocL1(b) + g_N(b)) & ")"
        fTTC = fTTC & "N(" & ttc & (BlocL1(b) + g_N(b)) & ")"
    Next b

    ws.Range(FF_ADR_RECAP_HT).Formula = fHT
    ws.Range(FF_ADR_RECAP_TTC).Formula = fTTC
    ws.Range(FF_ADR_RECAP_TVA).Formula = "=" & FF_ADR_RECAP_TTC & "-" & FF_ADR_RECAP_HT
    ws.Range(FF_ADR_RECAP_SAISI).Formula = "=IF(ISNUMBER(" & FF_ADR_TOTAL & ")," & FF_ADR_TOTAL & ","""")"
    ws.Range(FF_ADR_RECAP_ECART).Formula = "=IF(ISNUMBER(" & FF_ADR_RECAP_SAISI & "),ROUND(" & _
        FF_ADR_RECAP_SAISI & "-" & FF_ADR_RECAP_TTC & ",2),"""")"

End Sub

' =====================================================================================
' BOUTONS : + / - (lignes), AJOUTER UN BLOC, SUPPRIMER UN BLOC
' =====================================================================================
' Numéro de bloc lu dans le nom du bouton qui a déclenché la macro ("btnPlus5" -> 5).
Private Function BlocDuBouton(ByVal prefixe As String) As Long
    Dim nom As String
    On Error Resume Next
    nom = CStr(Application.Caller)
    On Error GoTo 0
    If Left$(nom, Len(prefixe)) = prefixe Then BlocDuBouton = CLng(Val(Mid$(nom, Len(prefixe) + 1)))
End Function

Public Sub FacAjouterLigne()
    ModifierStructure MODE_PLUS, BlocDuBouton("btnPlus"), ""
End Sub

Public Sub FacRetirerLigne()
    ModifierStructure MODE_MOINS, BlocDuBouton("btnMoins"), ""
End Sub

Public Sub FacSupprimerBloc()
    ModifierStructure MODE_SUPPR_BLOC, BlocDuBouton("btnSuppr"), ""
End Sub

' Bouton "Ajouter un bloc" : crée, sous les blocs existants, un bloc de même forme (mêmes colonnes) dont le nom est
' donné par l'opérateur. Le nom d'un bloc du modèle (Abonnement, Consommation, Taxes et contributions) le rétablit
' avec son indicateur d'origine ; pour tout autre nom, l'INDICATEUR alimenté par le bloc est choisi dans la liste
' (indicateurs de base + ceux déjà créés) ou créé à la volée : il est alors mémorisé (TblIndicateurs) et proposé dans
' toutes les listes futures. Le titre du bloc affiche cet indicateur entre parenthèses.
Public Sub FacAjouterBloc()

    Dim ws As Worksheet
    Dim nom As String, reponse As String, indic As String, txt As String
    Dim noms() As String
    Dim b As Long, nat As Long, nbInd As Long, choix As Long, defaut As Long

    If Not g_FacEnCours Or g_FacConsultation Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_FACTURE)

    If g_NbBlocs >= FF_MAX_BLOCS Then
        AfficherMessage ws, "Nombre maximal de blocs atteint (" & FF_MAX_BLOCS & ")."
        Exit Sub
    End If

    nom = Trim$(InputBox(mod_Display.FR("Nom du nouveau bloc (exemples : Redevance, Avoir, Location du compteur)." & vbCrLf & _
                         "Pour remettre un bloc du mod{e1}le que vous avez supprim{e2}, tapez son nom : Abonnement, Consommation ou Taxes et contributions."), _
                         mod_Display.FR("Ajouter un bloc")))
    If nom = "" Then Exit Sub
    nom = Replace(Replace(nom, vbCr, " "), vbLf, " ")
    If Len(nom) > FF_LONG_NOM_BLOC Then
        AfficherMessage ws, "Le nom du bloc est limit{e2} {a2} " & FF_LONG_NOM_BLOC & " caract{e1}res."
        Exit Sub
    End If

    ' Nom d'un bloc du modèle : indicateur d'origine, nom normalisé.
    nat = 0
    For b = NAT_ABO To NAT_TAXE
        If StrComp(nom, NomBlocModele(b), vbTextCompare) = 0 Then nat = b
    Next b
    If StrComp(nom, "Taxes", vbTextCompare) = 0 Then nat = NAT_TAXE
    If nat > 0 Then nom = NomBlocModele(nat)

    For b = 1 To g_NbBlocs
        If StrComp(g_NomBloc(b), nom, vbTextCompare) = 0 Then
            AfficherMessage ws, "Un bloc porte d{e2}j{a2} ce nom."
            Exit Sub
        End If
    Next b

    If nat > 0 Then
        indic = NomNature(nat)
    Else
        indic = ChoisirIndicateur(ws, NomNature(NAT_AUTRE))
        If indic = "" Then Exit Sub
    End If

    ModifierStructure MODE_AJOUT_BLOC, 0, nom, indic

End Sub

' Demande l'indicateur d'un bloc : liste numérotée (indicateurs de base + ceux déjà créés) avec 0 pour en créer un nouveau,
' mémorisé pour les listes futures. "defautNom" est proposé par défaut. Renvoie "" si l'opérateur annule ou si la
' saisie est invalide (un message est alors affiché).
Private Function ChoisirIndicateur(ByVal ws As Worksheet, ByVal defautNom As String) As String

    Dim noms() As String
    Dim txt As String, reponse As String, indic As String
    Dim b As Long, nbInd As Long, choix As Long, defaut As Long

    nbInd = ListeIndicateurs(noms)
    defaut = nbInd
    txt = mod_Display.FR("Indicateur aliment{e2} par ce bloc (totaux et graphique) :") & vbCrLf
    For b = 1 To nbInd
        txt = txt & b & " = " & noms(b) & vbCrLf
        If StrComp(noms(b), defautNom, vbTextCompare) = 0 Then defaut = b
    Next b
    txt = txt & mod_Display.FR("0 = Cr{e2}er un nouvel indicateur")
    reponse = Trim$(InputBox(txt, mod_Display.FR("Indicateur du bloc"), CStr(defaut)))
    If reponse = "" Then Exit Function

    choix = CLng(Val(reponse))
    If choix = 0 Then
        indic = Trim$(InputBox(mod_Display.FR("Nom du nouvel indicateur (") & FF_LONG_INDIC & mod_Display.FR(" caract{e1}res au plus). Il sera propos{e2} dans les listes suivantes :"), _
                               mod_Display.FR("Nouvel indicateur")))
        If indic = "" Then Exit Function
        indic = Replace(Replace(indic, vbCr, " "), vbLf, " ")
        If Len(indic) > FF_LONG_INDIC Then
            AfficherMessage ws, "Le nom de l'indicateur est limit{e2} {a2} " & FF_LONG_INDIC & " caract{e1}res."
            Exit Function
        End If
        ' Un indicateur déjà connu est réutilisé tel quel (même orthographe).
        For b = 1 To nbInd
            If StrComp(noms(b), indic, vbTextCompare) = 0 Then indic = noms(b)
        Next b
        MemoriserIndicateur indic
        ChoisirIndicateur = indic
    ElseIf choix >= 1 And choix <= nbInd Then
        ChoisirIndicateur = noms(choix)
    Else
        AfficherMessage ws, "Choix invalide : tapez un num{e2}ro de la liste, ou 0 pour cr{e2}er un indicateur."
    End If

End Function

' Bouton « Indicateur » du titre d'un bloc : corrige l'indicateur d'un bloc affecté par erreur. Si ce même bloc (même nom)
' figure dans d'autres factures, ces factures sont listées et une confirmation est demandée : le changement leur sera
' appliqué quand la facture en cours sera enregistrée (rien n'est modifié ailleurs si on abandonne le formulaire).
Public Sub FacModifierIndicateur()

    Dim ws As Worksheet
    Dim b As Long, nbImp As Long
    Dim nouvel As String, liste As String, conflits As String, txt As String
    Dim controle As Boolean

    If Not g_FacEnCours Or g_FacConsultation Then Exit Sub
    b = BlocDuBouton("btnIndic")
    If b < 1 Or b > g_NbBlocs Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_FACTURE)

    nouvel = ChoisirIndicateur(ws, g_Indic(b))
    If nouvel = "" Then Exit Sub
    If StrComp(nouvel, g_Indic(b), vbTextCompare) = 0 Then Exit Sub

    controle = (FamilleDepuisTexte(nouvel) = NAT_CONSO)
    nbImp = FacturesImpactees(g_NomBloc(b), g_IdEnEdition, controle, liste, conflits)

    If conflits <> "" Then
        MsgBox mod_Display.FR("Le bloc '") & g_NomBloc(b) & mod_Display.FR("' ne peut pas passer {a2} l'indicateur Consommation : des factures l'utilisent avec une unit{e2} autre que m") & ChrW(179) & _
               mod_Display.FR(" ou kWh (factures n") & ChrW(176) & " " & conflits & mod_Display.FR("). Corrigez d'abord ces factures."), vbExclamation, mod_Display.FR("Indicateur du bloc")
        Exit Sub
    End If

    If nbImp > 0 Then
        txt = mod_Display.FR("Le bloc '") & g_NomBloc(b) & mod_Display.FR("' est aussi utilis{e2} par ") & nbImp & mod_Display.FR(" autre(s) facture(s) :") & vbCrLf & vbCrLf & liste & vbCrLf & vbCrLf & _
              mod_Display.FR("Passer son indicateur de '") & g_Indic(b) & mod_Display.FR("' {a2} '") & nouvel & mod_Display.FR("' modifiera aussi ces factures. ") & _
              mod_Display.FR("Le changement leur sera appliqu{e2} quand vous validerez la facture en cours.") & vbCrLf & vbCrLf & mod_Display.FR("Confirmer ?")
        If MsgBox(txt, vbYesNo + vbQuestion + vbDefaultButton2, mod_Display.FR("Modifier l'indicateur d'un bloc")) = vbNo Then Exit Sub
        If g_Pend Is Nothing Then Set g_Pend = CreateObject("Scripting.Dictionary")
        g_Pend(LCase$(g_NomBloc(b))) = nouvel
    End If

    ModifierStructure MODE_INDIC, b, "", nouvel

End Sub

' Factures, autres que idCourant, dont au moins une ligne appartient à un bloc de ce nom. Renvoie leur nombre ;
' "liste" reçoit leur description (une par ligne, douze au plus) ; si controlerUnites, "conflits" reçoit les numéros de
' celles qui ont une ligne dont l'unité n'est ni kWh ni m³ (incompatible avec l'indicateur Consommation).
Private Function FacturesImpactees(ByVal nomBloc As String, ByVal idCourant As Long, ByVal controlerUnites As Boolean, _
                                   ByRef liste As String, ByRef conflits As String) As Long

    Dim lo As ListObject, loF As ListObject
    Dim vI() As Variant, vB() As Variant, vU() As Variant
    Dim vF() As Variant, vNum() As Variant, vDat() As Variant, vTyp() As Variant, vOp() As Variant
    Dim dic As Object, det As Object
    Dim i As Long, id As Long, nb As Long
    Dim cle As String, norm As String, d As String
    Dim k As Variant

    liste = ""
    conflits = ""
    Set lo = TableLignes()
    If lo Is Nothing Then Exit Function
    If mod_Consommations.TableVide(lo) Then Exit Function

    Set det = CreateObject("Scripting.Dictionary")
    Set loF = TableFactures()
    If Not loF Is Nothing Then
        If Not mod_Consommations.TableVide(loF) Then
            vF = mod_Consommations.LireColonne(loF, "ID_Facture")
            vNum = mod_Consommations.LireColonne(loF, "NumeroFacture")
            vDat = mod_Consommations.LireColonne(loF, "DateFacture")
            vTyp = mod_Consommations.LireColonne(loF, "TypeConso")
            vOp = mod_Consommations.LireColonne(loF, "Operateur")
            For i = 1 To loF.ListRows.Count
                If IsNumeric(vF(i)) And Not IsEmpty(vF(i)) Then
                    d = "n" & ChrW(176) & CLng(vF(i)) & " - " & mod_DataStructure.CellText(vTyp(i)) & " - " & mod_DataStructure.CellText(vOp(i))
                    If mod_DataStructure.CellText(vNum(i)) <> "" Then d = d & " - facture " & mod_DataStructure.CellText(vNum(i))
                    If mod_Consommations.VersJour(vDat(i)) > 0 Then d = d & " - " & mod_Consommations.TexteDate(mod_Consommations.VersJour(vDat(i)))
                    det(CStr(CLng(vF(i)))) = d
                End If
            Next i
        End If
    End If

    Set dic = CreateObject("Scripting.Dictionary")
    vI = mod_Consommations.LireColonne(lo, "ID_Facture")
    vB = mod_Consommations.LireColonne(lo, "Bloc")
    vU = mod_Consommations.LireColonne(lo, "Unite")
    For i = 1 To lo.ListRows.Count
        If IsNumeric(vI(i)) And Not IsEmpty(vI(i)) Then
            id = CLng(vI(i))
            If id <> idCourant And StrComp(mod_DataStructure.CellText(vB(i)), nomBloc, vbTextCompare) = 0 Then
                cle = CStr(id)
                If Not dic.Exists(cle) Then dic.Add cle, True
                If controlerUnites Then
                    If Not UniteConsommationValide(mod_DataStructure.CellText(vU(i)), norm) Then
                        If InStr(", " & conflits & ",", ", " & cle & ",") = 0 Then
                            If conflits = "" Then conflits = cle Else conflits = conflits & ", " & cle
                        End If
                    End If
                End If
            End If
        End If
    Next i

    nb = dic.Count
    i = 0
    For Each k In dic.Keys
        i = i + 1
        If i <= 12 Then
            If det.Exists(CStr(k)) Then liste = liste & det(CStr(k)) & vbCrLf Else liste = liste & "n" & ChrW(176) & CStr(k) & vbCrLf
        End If
    Next k
    If nb > 12 Then liste = liste & "... et " & (nb - 12) & " autre(s)"
    liste = RTrim$(liste)
    FacturesImpactees = nb

End Function

' Répercute, après l'enregistrement de la facture, les changements d'indicateur confirmés sur les autres factures qui
' utilisent un bloc de même nom. Un changement n'est appliqué que si le bloc existe toujours dans le formulaire avec
' cet indicateur (un bloc supprimé ou rechangé entre-temps n'entraîne rien).
Private Sub AppliquerChangementsIndicateurs(ByVal idFacture As Long)

    Dim lo As ListObject
    Dim vI() As Variant, vB() As Variant, vInd() As Variant
    Dim res() As Variant
    Dim cles As Collection
    Dim ids As Object
    Dim k As Variant
    Dim i As Long, b As Long, n As Long
    Dim actif As Boolean
    Dim cle As String

    If g_Pend Is Nothing Then Exit Sub
    If g_Pend.Count = 0 Then Exit Sub

    ' On ne garde que les changements encore valables.
    Set cles = New Collection
    For Each k In g_Pend.Keys
        actif = False
        For b = 1 To g_NbBlocs
            If LCase$(g_NomBloc(b)) = CStr(k) And StrComp(g_Indic(b), CStr(g_Pend(CStr(k))), vbTextCompare) = 0 Then actif = True
        Next b
        If actif Then cles.Add CStr(k)
    Next k
    If cles.Count = 0 Then Exit Sub

    Set lo = TableLignes()
    If lo Is Nothing Then Exit Sub
    If mod_Consommations.TableVide(lo) Then Exit Sub

    n = lo.ListRows.Count
    vI = mod_Consommations.LireColonne(lo, "ID_Facture")
    vB = mod_Consommations.LireColonne(lo, "Bloc")
    vInd = mod_Consommations.LireColonne(lo, "Indicateur")
    ReDim res(1 To n, 1 To 1)
    Set ids = CreateObject("Scripting.Dictionary")

    For i = 1 To n
        res(i, 1) = vInd(i)
        If IsNumeric(vI(i)) And Not IsEmpty(vI(i)) Then
            If CLng(vI(i)) <> idFacture Then
                cle = LCase$(mod_DataStructure.CellText(vB(i)))
                For Each k In cles
                    If CStr(k) = cle Then
                        If StrComp(mod_DataStructure.CellText(vInd(i)), CStr(g_Pend(CStr(k))), vbTextCompare) <> 0 Then
                            res(i, 1) = g_Pend(CStr(k))
                            ids(CStr(CLng(vI(i)))) = True
                        End If
                    End If
                Next k
            End If
        End If
    Next i

    If ids.Count > 0 Then
        lo.ListColumns("Indicateur").DataBodyRange.Value2 = res
        MsgBox ids.Count & mod_Display.FR(" autre(s) facture(s) mise(s) {a2} jour avec le nouvel indicateur du bloc."), vbInformation, mod_Display.FR("Indicateur du bloc")
    End If

End Sub

' Modifie la structure du formulaire (ajout / retrait d'une ligne, ajout / suppression d'un bloc) : les blocs
' suivants sont décalés, le formulaire est redessiné, et les valeurs déjà saisies sont relues avant et réécrites après.
Private Sub ModifierStructure(ByVal quoi As Long, ByVal b As Long, ByVal nom As String, Optional ByVal indic As String = "")

    Dim ws As Worksheet
    Dim brut(1 To FF_MAX_BLOCS) As Variant
    Dim ancN(1 To FF_MAX_BLOCS) As Long
    Dim ancNb As Long
    Dim ob As Long, nb As Long, i As Long, c As Long, m As Long, k As Long, r As Long
    Dim evenementsAvant As Boolean
    Dim numErr As Long
    Dim vide As Boolean

    If Not g_FacEnCours Or g_FacConsultation Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_FACTURE)

    Select Case quoi
        Case MODE_PLUS, MODE_MOINS, MODE_SUPPR_BLOC, MODE_INDIC
            If b < 1 Or b > g_NbBlocs Then Exit Sub
    End Select

    Select Case quoi
        Case MODE_PLUS
            If g_N(b) >= FF_MAX_LIGNES Then
                AfficherMessage ws, "Ce bloc est limit{e2} {a2} " & FF_MAX_LIGNES & " lignes."
                Exit Sub
            End If
        Case MODE_MOINS
            If g_N(b) <= 1 Then
                AfficherMessage ws, "Un bloc garde au moins une ligne."
                Exit Sub
            End If
            If Not LigneVide(ws, BlocL1(b) + g_N(b) - 1) Then
                AfficherMessage ws, "La derni{e1}re ligne de ce bloc n'est pas vide : effacez-la d'abord."
                Exit Sub
            End If
        Case MODE_AJOUT_BLOC
            If g_NbBlocs >= FF_MAX_BLOCS Then Exit Sub
        Case MODE_SUPPR_BLOC
            vide = True
            For i = 0 To g_N(b) - 1
                If Not LigneVide(ws, BlocL1(b) + i) Then vide = False
            Next i
            If Not vide Then
                If MsgBox(mod_Display.FR("Supprimer le bloc '") & NomDuBloc(b) & mod_Display.FR("' et les lignes qu'il contient ? (Seule cette facture est concern{e2}e : le mod{e1}le des nouvelles factures ne change pas.)"), _
                          vbYesNo + vbQuestion + vbDefaultButton2, mod_Display.FR("Supprimer un bloc")) = vbNo Then Exit Sub
            End If
    End Select

    evenementsAvant = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo Sortie

    ' Relecture des valeurs saisies (colonnes Libellé à TVA) de tous les blocs.
    ancNb = g_NbBlocs
    For ob = 1 To ancNb
        ancN(ob) = g_N(ob)
        brut(ob) = ws.Range(ws.Cells(BlocL1(ob), FF_COL_LIB), ws.Cells(BlocL1(ob) + ancN(ob) - 1, FF_COL_TVA)).Value
    Next ob

    ' Nouvelle structure.
    Select Case quoi
        Case MODE_PLUS
            g_N(b) = g_N(b) + 1
        Case MODE_MOINS
            g_N(b) = g_N(b) - 1
        Case MODE_INDIC
            g_Indic(b) = indic
            g_Nat(b) = FamilleDepuisTexte(indic)
        Case MODE_AJOUT_BLOC
            g_NbBlocs = g_NbBlocs + 1
            g_N(g_NbBlocs) = FF_NB_DEFAUT
            g_Indic(g_NbBlocs) = indic
            g_Nat(g_NbBlocs) = FamilleDepuisTexte(indic)
            g_NomBloc(g_NbBlocs) = nom
        Case MODE_SUPPR_BLOC
            For k = b To g_NbBlocs - 1
                g_N(k) = g_N(k + 1)
                g_Nat(k) = g_Nat(k + 1)
                g_Indic(k) = g_Indic(k + 1)
                g_NomBloc(k) = g_NomBloc(k + 1)
            Next k
            g_N(g_NbBlocs) = 0
            g_Nat(g_NbBlocs) = 0
            g_Indic(g_NbBlocs) = ""
            g_NomBloc(g_NbBlocs) = ""
            g_NbBlocs = g_NbBlocs - 1
    End Select

    DessinerBlocs ws

    ' Réécriture : chaque bloc conservé reprend ses valeurs (dans un bloc raccourci, seules les lignes restantes).
    For ob = 1 To ancNb
        If Not (quoi = MODE_SUPPR_BLOC And ob = b) Then
            nb = ob
            If quoi = MODE_SUPPR_BLOC And ob > b Then nb = ob - 1
            m = ancN(ob)
            If g_N(nb) < m Then m = g_N(nb)
            For i = 1 To m
                For c = 1 To (FF_COL_TVA - FF_COL_LIB + 1)
                    If Not IsEmpty(brut(ob)(i, c)) Then
                        ws.Cells(BlocL1(nb) + i - 1, FF_COL_LIB + c - 1).Value = brut(ob)(i, c)
                    End If
                Next c
            Next i
        End If
    Next ob

    RemplirTvaDefaut ws, mod_DataStructure.CellText(ws.Range(FF_ADR_TYPE).Value)
    ws.Range(FF_ADR_MESSAGE).Value = ""
    Select Case quoi
        Case MODE_PLUS
            ws.Cells(BlocL1(b) + g_N(b) - 1, FF_COL_LIB).Select
        Case MODE_AJOUT_BLOC
            ws.Cells(BlocL1(g_NbBlocs), FF_COL_LIB).Select
    End Select

Sortie:
    numErr = Err.Number
    Application.EnableEvents = evenementsAvant
    Application.ScreenUpdating = True
    If numErr <> 0 Then MsgBox mod_Display.FR("Erreur lors de la modification du formulaire : ") & Err.Description, vbExclamation

End Sub

' Vrai si la ligne r d'un bloc ne contient aucune saisie.
' La TVA (préremplie) et l'unité ne comptent pas : une ligne sans autre saisie reste vide.
Private Function LigneVide(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    Dim c As Long
    For c = FF_COL_LIB To FF_COL_PRIX
        If c <> FF_COL_UNITE Then
            If mod_Consommations.SaisieNonVide(ws.Cells(r, c).Value) Then Exit Function
        End If
    Next c
    LigneVide = True
End Function

Private Sub EcrireDate(ByVal c As Range, ByVal jour As Double)
    If jour > 0 Then c.Value = CDate(jour) Else c.ClearContents
End Sub

' Prix unitaire HT : un prix NÉGATIF est admis (remboursement, avoir). Pas de contrôle bloquant sur la cellule
' (une validation « décimal » refusait des saisies correctes) : seule une bulle d'aide s'affiche, et la validation
' du formulaire relit la valeur (virgule ou point décimal, signe moins) avant tout enregistrement.
Private Sub PoserAidePrix(ByVal c As Range)
    On Error Resume Next
    With c.Validation
        .Delete
        .Add Type:=xlValidateInputOnly
        .ShowInput = True
        .InputTitle = "Prix HT"
        .InputMessage = mod_Display.FR("Prix unitaire hors taxes, avec une virgule (exemple : 0,09010). Un remboursement se saisit en n{e2}gatif (exemple : -0,33810).")
    End With
    On Error GoTo 0
End Sub

' Validation des cellules de dates, quantités, prix et TVA d'un bloc.
Public Sub PoserGardeFousBloc(ByVal ws As Worksheet, ByVal b As Long)

    Dim l1 As Long, l2 As Long, r As Long

    l1 = BlocL1(b)
    l2 = l1 + BlocN(b) - 1
    ' Les validations de mod_InstallCommun ne portent que sur la première cellule de la plage : une cellule à la fois.
    For r = l1 To l2
        mod_InstallCommun.PoserValidationDate ws.Cells(r, FF_COL_DEB)
        mod_InstallCommun.PoserValidationDate ws.Cells(r, FF_COL_FIN)
        mod_InstallCommun.PoserValidationMontant ws.Cells(r, FF_COL_QTE), mod_InstallCommun.MONTANT_POSITIF_OU_NUL
        PoserAidePrix ws.Cells(r, FF_COL_PRIX)
    Next r
    If g_Nat(b) = NAT_CONSO Then
        ' Indicateur Consommation : l'unité est obligatoirement kWh ou m³ (liste stricte).
        ws.Columns(FF_COL_AIDE_UNITE_CONSO).Hidden = True
        ws.Cells(1, FF_COL_AIDE_UNITE_CONSO).Value = "UnitesConso"
        ws.Cells(2, FF_COL_AIDE_UNITE_CONSO).Value = "kWh"
        ws.Cells(3, FF_COL_AIDE_UNITE_CONSO).Value = "m" & ChrW(179)
        On Error Resume Next
        With ws.Range(ws.Cells(l1, FF_COL_UNITE), ws.Cells(l2, FF_COL_UNITE)).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Formula1:="=" & ws.Range(ws.Cells(2, FF_COL_AIDE_UNITE_CONSO), ws.Cells(3, FF_COL_AIDE_UNITE_CONSO)).Address(True, True)
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowError = True
            .ErrorTitle = mod_Display.FR("Unit{e2} de consommation")
            .ErrorMessage = mod_Display.FR("Une ligne de consommation a pour unit{e2} m") & ChrW(179) & mod_Display.FR(" ou kWh.")
        End With
        On Error GoTo 0
    ElseIf g_NbUnites > 0 Then
        On Error Resume Next
        With ws.Range(ws.Cells(l1, FF_COL_UNITE), ws.Cells(l2, FF_COL_UNITE)).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, _
                 Formula1:="=" & ws.Range(ws.Cells(2, FF_COL_AIDE_UNITE), ws.Cells(1 + g_NbUnites, FF_COL_AIDE_UNITE)).Address(True, True)
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowError = True
            .ErrorTitle = mod_Display.FR("Unit{e2} nouvelle")
            .ErrorMessage = mod_Display.FR("Cette unit{e2} n'est pas dans la liste. L'utiliser quand m{ea}me ? Elle sera ensuite propos{e2}e pour les prochaines lignes.")
        End With
        On Error GoTo 0
    End If
    On Error Resume Next
    With ws.Range(ws.Cells(l1, FF_COL_TVA), ws.Cells(l2, FF_COL_TVA)).Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="0", Formula2:="1"
        .IgnoreBlank = True
        .ShowInput = True
        .InputTitle = "TVA"
        .InputMessage = mod_Display.FR("Taux de TVA en pourcentage (exemple : 5,5 % ou 20 %).")
        .ShowError = True
        .ErrorTitle = mod_Display.FR("Taux invalide")
        .ErrorMessage = mod_Display.FR("Saisissez un taux entre 0 % et 100 % (exemple : 20 %).")
    End With
    On Error GoTo 0

End Sub

' Appelée par le code de la feuille (Worksheet_Change) : à la saisie du type, l'opérateur
' du contrat en vigueur est proposé s'il est encore vide.
Public Sub FacTraiterChangement(ByVal ws As Worksheet, ByVal Target As Range)

    Dim typ As String, oper As String, unite As String
    Dim d As Double
    Dim numErr As Long

    If Not g_FacEnCours Then Exit Sub
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Address(False, False) <> FF_ADR_TYPE Then Exit Sub

    On Error GoTo Sortie
    Application.EnableEvents = False
    typ = mod_DataStructure.CellText(ws.Range(FF_ADR_TYPE).Value)
    If typ = "" Then GoTo Sortie
    MettreAJourListeOperateurs ws, typ
    RemplirTvaDefaut ws, typ
    If mod_DataStructure.CellText(ws.Range(FF_ADR_OPERATEUR).Value) = "" Then
        d = mod_Consommations.VersJour(ws.Range(FF_ADR_DEBUT).Value)
        If d = 0 Then d = CDbl(Date)
        If mod_Consommations.OperateurContrat(typ, d, oper, unite) Then ws.Range(FF_ADR_OPERATEUR).Value = oper
    End If

Sortie:
    numErr = Err.Number
    Application.EnableEvents = True

End Sub

' Verrou "modal" : tant que le formulaire est ouvert, impossible de le quitter en changeant d'onglet.
Public Sub FacVerrouiller(ByVal ws As Worksheet)
    If g_FacVerrouActif And Not mod_Aide.AideOuverte() Then ws.Activate
End Sub

Public Sub FacValider()

    Dim ws As Worksheet
    Dim f As TypeFacture
    Dim tmp As TypeFacture
    Dim vTot As Variant
    Dim dTmp As Date
    Dim dDeb As Double, dFin As Double
    Dim msg As String

    If Not g_FacEnCours Or g_FacConsultation Then Exit Sub
    On Error GoTo Erreur
    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_FACTURE)

    ' L'identifiant de la facture en cours de modification est conservé entre l'ouverture et la validation.
    f.ID = g_IdEnEdition

    f.Typ = mod_DataStructure.CellText(ws.Range(FF_ADR_TYPE).Value)
    If f.Typ = "" Or Not mod_Consommations.TypeConnu(f.Typ) Then
        AfficherMessage ws, "Choisissez le type de consommation dans la liste."
        Exit Sub
    End If
    f.Operateur = mod_DataStructure.CellText(ws.Range(FF_ADR_OPERATEUR).Value)
    If f.Operateur = "" Then
        AfficherMessage ws, "Renseignez l'op{e2}rateur (fournisseur)."
        Exit Sub
    End If
    If Not mod_Consommations.OperateurConnu(f.Typ, f.Operateur) Then
        AfficherMessage ws, "Cet op{e2}rateur n'a pas de contrat pour ce type de consommation. Saisissez d'abord un nouveau contrat (bouton Contrats de l'{e2}cran Consommations), puis revenez saisir la facture."
        Exit Sub
    End If
    f.Numero = mod_DataStructure.CellText(ws.Range(FF_ADR_NUMERO).Value)

    If Not LireDateFacultative(ws, FF_ADR_DATE, "date de la facture", f.DateFac) Then Exit Sub
    If Not LireDateFacultative(ws, FF_ADR_DEBUT, "d{e2}but de la p{e2}riode", f.Debut) Then Exit Sub
    If Not LireDateFacultative(ws, FF_ADR_FIN, "fin de la p{e2}riode", f.Fin) Then Exit Sub
    If f.Debut > 0 And f.Fin > 0 And f.Fin < f.Debut Then
        AfficherMessage ws, "La fin de la p{e2}riode est ant{e2}rieure {a2} son d{e2}but."
        Exit Sub
    End If

    vTot = ws.Range(FF_ADR_TOTAL).Value
    If mod_Consommations.SaisieNonVide(vTot) Then
        If Not mod_DataStructure.TryLireMontant(vTot, f.TotalSaisi) Then
            AfficherMessage ws, "Le total de la facture n'est pas valide. " & mod_DataStructure.MessageMontantAttendu()
            Exit Sub
        End If
        f.TotalRenseigne = True
    End If
    f.Commentaire = mod_DataStructure.CellText(ws.Range(FF_ADR_COMMENT).Value)

    If Not LireLignes(ws, f) Then Exit Sub

    If g_FacNb = 0 Then
        AfficherMessage ws, "Saisissez au moins une ligne (abonnement, consommation ou taxe)."
        Exit Sub
    End If

    msg = ControlerChevauchements(f)
    If msg <> "" Then
        If MsgBox(mod_Display.FR(msg) & vbCrLf & vbCrLf & mod_Display.FR("Enregistrer quand m{ea}me ?"), _
                  vbYesNo + vbExclamation + vbDefaultButton2, mod_Display.FR("Chevauchement de p{e2}riodes")) = vbNo Then Exit Sub
    End If

    g_FacRes = f
    g_FacValide = True
    FermerFeuille ws
    Exit Sub

Erreur:
    Application.EnableEvents = True
    MsgBox mod_Display.FR("Erreur dans FacValider : ") & Err.Number & " - " & Err.Description, vbCritical

End Sub

Private Function LireDateFacultative(ByVal ws As Worksheet, ByVal adr As String, ByVal nom As String, ByRef jour As Double) As Boolean

    Dim v As Variant
    Dim d As Date

    v = ws.Range(adr).Value
    jour = 0
    If Not mod_Consommations.SaisieNonVide(v) Then
        LireDateFacultative = True
        Exit Function
    End If
    If Not mod_DataStructure.TryLireDate(v, d) Then
        AfficherMessage ws, "La " & nom & " n'est pas valide. " & mod_DataStructure.MessageDateAttendue()
        Exit Function
    End If
    jour = CDbl(d)
    LireDateFacultative = True

End Function

' Lit les trois blocs et remplit g_FacLignes / g_FacNb (HT et TTC calculés comme sur la facture).
' Renvoie Faux (avec un message) si une ligne est incorrecte.
Private Function LireLignes(ByVal ws As Worksheet, ByRef f As TypeFacture) As Boolean

    Dim nat As Long, i As Long, r As Long
    Dim vLib As Variant, vDeb As Variant, vFin As Variant, vQte As Variant, vUni As Variant, vPrix As Variant, vTva As Variant
    Dim dDeb As Date, dFin As Date
    Dim q As Double, p As Double, tv As Double
    Dim texteLigne As String
    Dim uniteNorm As String

    g_FacNb = 0
    ReDim g_FacLignes(1 To FF_MAX_BLOCS * FF_MAX_LIGNES)

    ' "nat" désigne ici le NUMÉRO DU BLOC (1 à 3 : blocs fixes ; 4 et plus : blocs libres).
    For nat = 1 To g_NbBlocs
        For i = 0 To BlocN(nat) - 1
            r = BlocL1(nat) + i
            vLib = ws.Cells(r, FF_COL_LIB).Value
            vDeb = ws.Cells(r, FF_COL_DEB).Value
            vFin = ws.Cells(r, FF_COL_FIN).Value
            vQte = ws.Cells(r, FF_COL_QTE).Value
            vUni = ws.Cells(r, FF_COL_UNITE).Value
            vPrix = ws.Cells(r, FF_COL_PRIX).Value
            vTva = ws.Cells(r, FF_COL_TVA).Value

            ' Ligne vide : ignorée (la TVA et l'unité, préremplies, ne comptent pas).
            If Not (mod_Consommations.SaisieNonVide(vLib) Or mod_Consommations.SaisieNonVide(vDeb) Or mod_Consommations.SaisieNonVide(vFin) Or _
                    mod_Consommations.SaisieNonVide(vPrix) Or mod_Consommations.SaisieNonVide(vQte)) Then GoTo Suivante

            texteLigne = NomDuBloc(nat) & ", ligne " & (i + 1) & " : "

            If Not mod_Consommations.SaisieNonVide(vDeb) Or Not mod_DataStructure.TryLireDate(vDeb, dDeb) Then
                AfficherMessage ws, texteLigne & "date de d{e2}but absente ou invalide."
                Exit Function
            End If
            If Not mod_Consommations.SaisieNonVide(vFin) Or Not mod_DataStructure.TryLireDate(vFin, dFin) Then
                AfficherMessage ws, texteLigne & "date de fin absente ou invalide."
                Exit Function
            End If
            If dFin < dDeb Then
                AfficherMessage ws, texteLigne & "la date de fin est ant{e2}rieure {a2} la date de d{e2}but."
                Exit Function
            End If

            If Not mod_Consommations.SaisieNonVide(vQte) Or Not mod_DataStructure.TryLireMontant(vQte, q, 4) Then
                AfficherMessage ws, texteLigne & "quantit{e2} absente ou invalide."
                Exit Function
            End If
            If q < 0 Then
                AfficherMessage ws, texteLigne & "la quantit{e2} ne peut pas {ea}tre n{e2}gative."
                Exit Function
            End If
            If Not mod_Consommations.SaisieNonVide(vPrix) Or Not mod_DataStructure.TryLireMontant(vPrix, p, 6) Then
                AfficherMessage ws, texteLigne & "prix unitaire HT absent ou invalide (6 d{e2}cimales au maximum)."
                Exit Function
            End If
            If Not mod_Consommations.SaisieNonVide(vTva) Then
                AfficherMessage ws, texteLigne & "renseignez le taux de TVA (0 % si non applicable)."
                Exit Function
            End If
            If Not mod_DataStructure.TryLireMontant(vTva, tv, 6) Then
                AfficherMessage ws, texteLigne & "taux de TVA invalide."
                Exit Function
            End If
            ' Tolérance : 20 saisi dans une cellule non pourcentage = 20 %.
            If tv > 1 Then tv = tv / 100
            If tv < 0 Or tv > 1 Then
                AfficherMessage ws, texteLigne & "le taux de TVA doit {ea}tre compris entre 0 % et 100 %."
                Exit Function
            End If

            ' Garde-fou : une ligne de l'indicateur Consommation a pour unité kWh ou m³ (m3 est accepté et converti en m³).
            uniteNorm = mod_DataStructure.CellText(vUni)
            If g_Nat(nat) = NAT_CONSO Then
                If Not UniteConsommationValide(uniteNorm, uniteNorm) Then
                    AfficherMessage ws, texteLigne & "l'unit{e2} d'une ligne de consommation doit {ea}tre m" & ChrW(179) & " ou kWh."
                    Exit Function
                End If
            End If

            g_FacNb = g_FacNb + 1
            With g_FacLignes(g_FacNb)
                .Nat = NatDuBloc(nat)             ' code de l'indicateur du bloc
                .Indic = g_Indic(nat)
                .Ordre = nat                      ' position du bloc dans le formulaire
                .Bloc = g_NomBloc(nat)
                .Libelle = mod_DataStructure.CellText(vLib)
                .Debut = CDbl(dDeb)
                .Fin = CDbl(dFin)
                .Qte = q
                .Unite = uniteNorm
                .Prix = p
                .Tva = tv
                .HT = Arrondi(q * p)
                .TTC = Arrondi(.HT * (1 + tv))
            End With
Suivante:
        Next i
    Next nat

    LireLignes = True

End Function

' Détecte les chevauchements qui trahissent une double saisie : deux lignes de consommation
' (ou d'abonnement) de la même facture qui se recouvrent, ou une période déjà couverte par
' une autre facture du même type. Renvoie le texte d'avertissement, ou "".
Private Function ControlerChevauchements(ByRef f As TypeFacture) As String

    Dim i As Long, j As Long, n As Long
    Dim t() As TypeLigne

    ' Lignes de la même facture entre elles (consommation et abonnement).
    For i = 1 To g_FacNb
        If (g_FacLignes(i).Nat = NAT_ABO Or g_FacLignes(i).Nat = NAT_CONSO) Then
            For j = i + 1 To g_FacNb
                If g_FacLignes(j).Nat = g_FacLignes(i).Nat Then
                    If g_FacLignes(i).Debut <= g_FacLignes(j).Fin And g_FacLignes(i).Fin >= g_FacLignes(j).Debut Then
                        ControlerChevauchements = "Deux lignes de " & LCase$(NomNature(g_FacLignes(i).Nat)) & " de cette facture se chevauchent (du " & _
                            mod_Consommations.TexteDate(g_FacLignes(j).Debut) & " au " & mod_Consommations.TexteDate(g_FacLignes(j).Fin) & ")."
                        Exit Function
                    End If
                End If
            Next j
        End If
    Next i

    ' Autres factures du même type.
    n = ChargerLignes(t, f.ID)
    For i = 1 To g_FacNb
        If (g_FacLignes(i).Nat = NAT_ABO Or g_FacLignes(i).Nat = NAT_CONSO) Then
            For j = 1 To n
                If t(j).Nat = g_FacLignes(i).Nat And StrComp(t(j).Typ, f.Typ, vbTextCompare) = 0 Then
                    If g_FacLignes(i).Debut <= t(j).Fin And g_FacLignes(i).Fin >= t(j).Debut Then
                        ControlerChevauchements = "Une ligne de " & LCase$(NomNature(g_FacLignes(i).Nat)) & " chevauche la facture n" & ChrW(176) & _
                            t(j).IDFacture & " du m" & ChrW(234) & "me type (p" & ChrW(233) & "riode du " & mod_Consommations.TexteDate(t(j).Debut) & _
                            " au " & mod_Consommations.TexteDate(t(j).Fin) & ")."
                        Exit Function
                    End If
                End If
            Next j
        End If
    Next i

End Function

Public Sub FacAnnuler()

    Dim ws As Worksheet

    If Not g_FacEnCours Then Exit Sub
    On Error GoTo Erreur

    If MsgBox(mod_Display.FR("Annuler la saisie ? Rien ne sera enregistr{e2}."), vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_FACTURE)
    g_FacValide = False
    FermerFeuille ws
    Exit Sub

Erreur:
    Application.EnableEvents = True
    MsgBox mod_Display.FR("Erreur dans FacAnnuler : ") & Err.Number & " - " & Err.Description, vbCritical

End Sub

Private Sub AfficherMessage(ByVal ws As Worksheet, ByVal texte As String)
    ws.Range(FF_ADR_MESSAGE).Value = mod_Display.FR(texte)
End Sub

' g_FacEnCours doit passer à False AVANT de masquer la feuille, sinon le verrou la réactiverait.
Private Sub FermerFeuille(ByVal ws As Worksheet)
    g_FacEnCours = False
    g_FacVerrouActif = False
    mod_Aide.MasquerFormulaire ws
End Sub
