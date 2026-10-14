Attribute VB_Name = "mod_Consommations"
Option Explicit

' =====================================================================================
' MODULE : mod_Consommations
'
' SUIVI DES CONSOMMATIONS (eau, électricité, gaz) - V1 (logique). La construction des
' feuilles est dans mod_InstallConsommations (macro InstallerConsommations, à lancer
' une seule fois).
'
' CE QUE FAIT LE MODULE :
'   - Deux tableaux : TblContrats (un contrat = un type de consommation chez un
'     opérateur, entre deux dates) et TblReleves (une ligne = une période de
'     consommation avec SA quantité et SON prix d'unité).
'   - Le changement d'opérateur se fait en créant un nouveau contrat : le contrat en
'     cours est clôturé la veille. L'historique reste complet.
'   - Le prix de l'unité ET l'abonnement sont portés par chaque relevé (pas par le
'     contrat) : un tarif qui change en cours de contrat est donc historisé. Le contrat
'     ne sert que de valeur de départ pour l'abonnement.
'   - Une période est LIBRE (date de début / date de fin) : une facture dont le prix
'     change plusieurs fois dans le mois se saisit en plusieurs relevés consécutifs
'     (un par prix), avec le même N° de facture.
'   - L'écran frm_Consommations montre l'évolution mois par mois (barres, courbes, ou
'     barres + courbe du prix) pour un type, ou le montant de tous les types.
'
' RÈGLES DE CALCUL :
'   - NbMois d'une période = somme, pour chaque mois touché, de (jours couverts / jours
'     du mois). Un mois complet vaut donc exactement 1, quelle que soit sa longueur, et
'     plusieurs relevés découpant un même mois totalisent exactement 1.
'   - Abonnement d'un relevé = AbonnementMensuel x NbMois.
'   - Montant d'un relevé = Quantité x PrixUnitaire + Abonnement.
'   - Graphique : par défaut, chaque relevé est RÉPARTI sur les mois qu'il couvre au
'     prorata des jours (quantité et montant de consommation) et au prorata de la part de
'     mois (abonnement). L'autre choix met tout le relevé dans le mois de sa date de fin.
'   - Prix unitaire d'un mois = moyenne pondérée par les quantités, abonnement exclu.
'
' POINTS D'ENTRÉE (boutons) : AfficherConsommations (feuille Synthese), ActualiserConsommations,
'   SaisirReleve, AjouterContrat, AfficherReleves, AfficherContrats, RetourConsommations,
'   SortirConsommations, SupprimerReleveSelection, SupprimerContratSelection.
' Appelées par le code des feuilles-formulaires : RelTraiterChangement, RelVerrouiller,
'   CtrTraiterChangement, CtrVerrouiller (voir CodeBehind_frm_Releve.txt et
'   CodeBehind_frm_Contrat.txt).
'
' NOTE : les petits outils de lecture de tableau (ObtenirTable, ColIdx, TableVide...) sont
' recopiés de mod_Echeancier, où ils sont privés. Une factorisation est possible plus tard.
'
' À PROPOS DES ACCENTS : textes affichés via mod_Display.FR(); commentaires en UTF-8.
' Le mot "coût" n'a pas de balise FR : le libellé employé est "Montant".
' =====================================================================================

' --- Disposition de l'écran frm_Consommations (partagée avec mod_InstallConsommations) ---
Public Const CO_ADR_TYPE As String = "C7"
Public Const CO_ADR_INDIC As String = "C8"
Public Const CO_ADR_PERIODE As String = "C9"
Public Const CO_ADR_ANNEE As String = "C10"
Public Const CO_ADR_MOIS As String = "C11"
Public Const CO_ADR_FORME As String = "C12"
Public Const CO_ADR_REPART As String = "C13"
Public Const CO_ADR_MESSAGE As String = "B15"
Public Const CO_LIGNE_ENTETE As Long = 17       ' en-tête du tableau mensuel
Public Const CO_NB_LIGNES_MAX As Long = 120     ' lignes de données (120 mois = 10 ans au plus)
Public Const CO_GRAPH As String = "chConsommation"

' --- Colonnes techniques masquées de frm_Consommations ---
' Données du graphique : mois, prix unitaire moyen, puis une colonne par série (opérateur ou type).
Public Const CO_NB_SERIES As Long = 6
Public Const CO_COL_H_LIB As Long = 30          ' AD
Public Const CO_COL_H_PRIX As Long = 31         ' AE
Public Const CO_COL_H_SERIE As Long = 32        ' AF à AK
Public Const CO_COL_LISTE_TYPE As Long = 40     ' AN
Public Const CO_COL_LISTE_INDIC As Long = 41
Public Const CO_COL_LISTE_PERIODE As Long = 42
Public Const CO_COL_LISTE_FORME As Long = 43
Public Const CO_COL_LISTE_REPART As Long = 44
Public Const CO_COL_LISTE_MOIS As Long = 45
Public Const CO_COL_LISTE_ANNEE As Long = 46    ' AT

' --- Formulaire frm_Releve (colonne B = étiquettes, C = saisie, D = unités) ---
Public Const RF_ADR_TYPE As String = "C6"
Public Const RF_ADR_OPERATEUR As String = "C7"          ' lecture, zone fusionnée C7:E7
Public Const RF_ADR_DEBUT As String = "C8"
Public Const RF_ADR_FIN As String = "C9"
Public Const RF_ADR_QTE As String = "C10"
Public Const RF_ADR_UNITE As String = "D10"             ' lecture
Public Const RF_ADR_PRIX As String = "C11"
Public Const RF_ADR_UNITE_PRIX As String = "D11"        ' lecture
Public Const RF_ADR_ABO As String = "C12"
Public Const RF_ADR_UNITE_ABO As String = "D12"         ' lecture
Public Const RF_ADR_FACTURE As String = "C13"
Public Const RF_ADR_COMMENT As String = "C14"
Public Const RF_ADR_CALC_CONSO As String = "C16"        ' formules de lecture
Public Const RF_ADR_CALC_ABO As String = "C17"
Public Const RF_ADR_CALC_TOTAL As String = "C18"
Public Const RF_ADR_MESSAGE As String = "B20"           ' zone fusionnée B20:E20
Public Const RF_COL_AIDE_TYPE As Long = 26              ' colonne technique masquée (liste des types)

' --- Formulaire frm_Contrat ---
Public Const CF_ADR_TYPE As String = "C6"
Public Const CF_ADR_OPERATEUR As String = "C7"
Public Const CF_ADR_UNITE As String = "C8"
Public Const CF_ADR_DEBUT As String = "C9"
Public Const CF_ADR_FIN As String = "C10"
Public Const CF_ADR_ABO As String = "C11"
Public Const CF_ADR_REFERENCE As String = "C12"
Public Const CF_ADR_MESSAGE As String = "B14"           ' zone fusionnée B14:E14

' Un relevé tel que saisi dans le formulaire.
Public Type TypeReleve
    Typ As String
    Debut As Date
    Fin As Date
    Qte As Double
    Prix As Double
    AboMensuel As Double
    Facture As String
    Commentaire As String
End Type

' Un contrat tel que saisi dans le formulaire.
Public Type TypeContrat
    Typ As String
    Operateur As String
    Unite As String
    Debut As Date
    Fin As Date                 ' 0 = contrat en cours (pas de date de fin)
    AboMensuel As Double
    Reference As String
End Type

' --- État des formulaires de saisie ---
Private g_RelEnCours As Boolean
Private g_RelVerrouActif As Boolean
Private g_RelValide As Boolean
Private g_RelSuivant As Boolean
Private g_RelRes As TypeReleve
Private g_RelNomFeuillePrec As String

Private g_CtrEnCours As Boolean
Private g_CtrVerrouActif As Boolean
Private g_CtrValide As Boolean
Private g_CtrRes As TypeContrat
Private g_CtrLigneAClore As Long
Private g_CtrNomFeuillePrec As String

' --- Données chargées en mémoire (voir ChargerReleves / ChargerContrats) ---
' Les dates sont des nombres entiers (numéro de série Excel) ; 0 = absente.
Private r_Nb As Long
Private r_Typ() As String
Private r_Deb() As Double
Private r_Fin() As Double
Private r_Qte() As Double
Private r_Prix() As Double
Private r_Abo() As Double

Private c_Nb As Long
Private c_Typ() As String
Private c_Oper() As String
Private c_Unite() As String
Private c_Deb() As Double
Private c_Fin() As Double
Private c_Abo() As Double
Private c_Ligne() As Long        ' numéro de la ligne dans TblContrats


' =====================================================================================
' LIBELLÉS (fonctions : FR() ne peut pas servir dans une constante)
' =====================================================================================
Public Function TypeAdduction() As String
    TypeAdduction = "Adduction eau"
End Function

Public Function TypeRecyclage() As String
    TypeRecyclage = "Recyclage eau"
End Function

Public Function TypeElec() As String
    TypeElec = mod_Display.FR("Fourniture {e2}lec")
End Function

Public Function TypeGaz() As String
    TypeGaz = "Fourniture gaz"
End Function

Public Function TypesConso() As Variant
    TypesConso = Array(TypeAdduction(), TypeRecyclage(), TypeElec(), TypeGaz())
End Function

Public Function LibTous() As String
    LibTous = "(tous)"
End Function

Public Function LibIndQuantite() As String
    LibIndQuantite = mod_Display.FR("Quantit{e2}")
End Function

Public Function LibIndPrix() As String
    LibIndPrix = "Prix unitaire"
End Function

Public Function LibIndAbo() As String
    LibIndAbo = "Abonnement"
End Function

Public Function LibIndMontant() As String
    LibIndMontant = "Montant total"
End Function

Public Function LibPeriode12() As String
    LibPeriode12 = "12 derniers mois"
End Function

Public Function LibPeriode36() As String
    LibPeriode36 = "36 derniers mois"
End Function

Public Function LibPeriode60() As String
    LibPeriode60 = "60 derniers mois"
End Function

Public Function LibPeriode120() As String
    LibPeriode120 = "120 derniers mois"
End Function

Public Function LibPeriodeAnnee() As String
    LibPeriodeAnnee = mod_Display.FR("Ann{e2}e")
End Function

Public Function LibFormeBarres() As String
    LibFormeBarres = "Barres"
End Function

Public Function LibFormeCourbes() As String
    LibFormeCourbes = "Courbes"
End Function

Public Function LibFormeMixte() As String
    LibFormeMixte = "Barres + courbe du prix"
End Function

Public Function LibRepartProrata() As String
    LibRepartProrata = "Au prorata des jours"
End Function

Public Function LibRepartFin() As String
    LibRepartFin = mod_Display.FR("Mois de fin du relev{e2}")
End Function

Public Function LibSansContrat() As String
    LibSansContrat = "(sans contrat)"
End Function

Public Function LibAutres() As String
    LibAutres = "Autres"
End Function

' Unité proposée par défaut selon le type (modifiable dans le contrat).
Public Function UniteDefaut(ByVal typ As String) As String
    If StrComp(typ, TypeAdduction(), vbTextCompare) = 0 Or StrComp(typ, TypeRecyclage(), vbTextCompare) = 0 Then
        UniteDefaut = "m" & ChrW(179)
    Else
        UniteDefaut = "kWh"
    End If
End Function

Private Function EstUniteDefaut(ByVal u As String) As Boolean
    EstUniteDefaut = (StrComp(u, "m" & ChrW(179), vbTextCompare) = 0 Or StrComp(u, "kWh", vbTextCompare) = 0)
End Function

Private Function TypeConnu(ByVal typ As String) As Boolean
    Dim t As Variant
    For Each t In TypesConso()
        If StrComp(CStr(t), typ, vbTextCompare) = 0 Then
            TypeConnu = True
            Exit Function
        End If
    Next t
End Function

Private Function SymboleEuro() As String
    SymboleEuro = ChrW(8364)
End Function


' =====================================================================================
' OUTILS DE LECTURE / ÉCRITURE DES TABLEAUX
' =====================================================================================
Private Function ObtenirTable(ByVal nomFeuille As String, ByVal nomTable As String) As ListObject
    On Error Resume Next
    Set ObtenirTable = ThisWorkbook.Worksheets(nomFeuille).ListObjects(nomTable)
    On Error GoTo 0
End Function

Private Function TableReleves() As ListObject
    Set TableReleves = ObtenirTable(mod_VarGlobales.NOM_FEUILLE_RELEVES, mod_VarGlobales.NOM_TABLE_RELEVES)
End Function

Private Function TableContrats() As ListObject
    Set TableContrats = ObtenirTable(mod_VarGlobales.NOM_FEUILLE_CONTRATS, mod_VarGlobales.NOM_TABLE_CONTRATS)
End Function

' Vrai si les deux tableaux existent (sinon le suivi n'est pas installé).
Public Function ConsoInstallee() As Boolean
    ConsoInstallee = Not (TableReleves() Is Nothing) And Not (TableContrats() Is Nothing)
End Function

Private Function ColIdx(ByVal lo As ListObject, ByVal nomCol As String) As Long
    On Error Resume Next
    ColIdx = lo.ListColumns(nomCol).Index
    On Error GoTo 0
End Function

' Un tableau Excel a toujours au moins une ligne : "vide" = une seule ligne dont la
' première colonne (l'identifiant) est vide.
Private Function TableVide(ByVal lo As ListObject) As Boolean
    If lo.DataBodyRange Is Nothing Then
        TableVide = True
    ElseIf lo.ListRows.Count = 1 Then
        TableVide = (mod_DataStructure.CellText(lo.DataBodyRange.Cells(1, 1).Value2) = "")
    End If
End Function

Private Function NouvelleLigne(ByVal lo As ListObject) As ListRow
    If lo.DataBodyRange Is Nothing Then
        Set NouvelleLigne = lo.ListRows.Add
    ElseIf TableVide(lo) Then
        Set NouvelleLigne = lo.ListRows(1)
    Else
        Set NouvelleLigne = lo.ListRows.Add
    End If
End Function

Private Function LireColonne(ByVal lo As ListObject, ByVal nomCol As String) As Variant()
    Dim v As Variant, r() As Variant
    Dim i As Long, n As Long
    n = lo.ListRows.Count
    ReDim r(1 To n)
    v = lo.ListColumns(nomCol).DataBodyRange.Value2
    If IsArray(v) Then
        For i = 1 To n
            r(i) = v(i, 1)
        Next i
    Else
        r(1) = v
    End If
    LireColonne = r
End Function

Private Function ProchainID(ByVal lo As ListObject, ByVal nomCol As String) As Long
    Dim v() As Variant
    Dim i As Long, mx As Long
    If lo.DataBodyRange Is Nothing Then
        ProchainID = 1
        Exit Function
    End If
    v = LireColonne(lo, nomCol)
    For i = 1 To UBound(v)
        If Not IsEmpty(v(i)) And Not IsError(v(i)) Then
            If IsNumeric(v(i)) Then
                If CLng(v(i)) > mx Then mx = CLng(v(i))
            End If
        End If
    Next i
    ProchainID = mx + 1
End Function

' Valeur de cellule -> numéro de jour entier (0 si ce n'est pas une date).
Private Function VersJour(ByVal v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then
        VersJour = Int(CDbl(v))
    ElseIf IsDate(v) Then
        VersJour = Int(CDbl(CDate(v)))
    End If
End Function

' Vrai si la cellule contient quelque chose (autre que des espaces) : permet de distinguer
' un champ laisse vide d'un champ rempli avec une valeur invalide.
Private Function SaisieNonVide(ByVal v As Variant) As Boolean
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If IsError(v) Then
        SaisieNonVide = True
    Else
        SaisieNonVide = (Trim$(CStr(v)) <> "")
    End If
End Function

Private Function TexteDate(ByVal jour As Double) As String
    TexteDate = Format$(CDate(jour), "dd/mm/yyyy")
End Function

' Chargement des relevés en mémoire. Les lignes incomplètes (type, début ou fin absent,
' fin avant début) sont ignorées.
Private Function ChargerReleves() As Boolean

    Dim lo As ListObject
    Dim vT() As Variant, vD() As Variant, vF() As Variant, vQ() As Variant, vP() As Variant, vA() As Variant
    Dim i As Long, n As Long
    Dim t As String, d As Double, f As Double

    r_Nb = 0
    Set lo = TableReleves()
    If lo Is Nothing Then Exit Function
    ChargerReleves = True
    If TableVide(lo) Then Exit Function

    n = lo.ListRows.Count
    vT = LireColonne(lo, "TypeConso")
    vD = LireColonne(lo, "DateDebut")
    vF = LireColonne(lo, "DateFin")
    vQ = LireColonne(lo, "Quantite")
    vP = LireColonne(lo, "PrixUnitaire")
    vA = LireColonne(lo, "AbonnementMensuel")

    ReDim r_Typ(1 To n)
    ReDim r_Deb(1 To n)
    ReDim r_Fin(1 To n)
    ReDim r_Qte(1 To n)
    ReDim r_Prix(1 To n)
    ReDim r_Abo(1 To n)

    For i = 1 To n
        t = mod_DataStructure.CellText(vT(i))
        d = VersJour(vD(i))
        f = VersJour(vF(i))
        If t <> "" And d > 0 And f >= d Then
            r_Nb = r_Nb + 1
            r_Typ(r_Nb) = t
            r_Deb(r_Nb) = d
            r_Fin(r_Nb) = f
            r_Qte(r_Nb) = mod_DataStructure.ToDouble(vQ(i))
            r_Prix(r_Nb) = mod_DataStructure.ToDouble(vP(i))
            r_Abo(r_Nb) = mod_DataStructure.ToDouble(vA(i))
        End If
    Next i

End Function

' Chargement des contrats en mémoire (c_Ligne = position dans le tableau).
Private Function ChargerContrats() As Boolean

    Dim lo As ListObject
    Dim vT() As Variant, vO() As Variant, vU() As Variant, vD() As Variant, vF() As Variant, vA() As Variant
    Dim i As Long, n As Long
    Dim t As String

    c_Nb = 0
    Set lo = TableContrats()
    If lo Is Nothing Then Exit Function
    ChargerContrats = True
    If TableVide(lo) Then Exit Function

    n = lo.ListRows.Count
    vT = LireColonne(lo, "TypeConso")
    vO = LireColonne(lo, "Operateur")
    vU = LireColonne(lo, "Unite")
    vD = LireColonne(lo, "DateDebut")
    vF = LireColonne(lo, "DateFin")
    vA = LireColonne(lo, "AbonnementMensuel")

    ReDim c_Typ(1 To n)
    ReDim c_Oper(1 To n)
    ReDim c_Unite(1 To n)
    ReDim c_Deb(1 To n)
    ReDim c_Fin(1 To n)
    ReDim c_Abo(1 To n)
    ReDim c_Ligne(1 To n)

    For i = 1 To n
        t = mod_DataStructure.CellText(vT(i))
        If t <> "" Then
            c_Nb = c_Nb + 1
            c_Typ(c_Nb) = t
            c_Oper(c_Nb) = mod_DataStructure.CellText(vO(i))
            c_Unite(c_Nb) = mod_DataStructure.CellText(vU(i))
            c_Deb(c_Nb) = VersJour(vD(i))
            c_Fin(c_Nb) = VersJour(vF(i))
            c_Abo(c_Nb) = mod_DataStructure.ToDouble(vA(i))
            c_Ligne(c_Nb) = i
        End If
    Next i

End Function


' =====================================================================================
' RÈGLES DE CALCUL
' =====================================================================================

' Nombre de mois d'une période (voir l'en-tête du module). Même résultat que la
' colonne NbMois de TblReleves (formule Excel, voir FormuleNbMois).
Private Function NbMoisPeriode(ByVal deb As Double, ByVal fin As Double) As Double

    Dim i As Long, iDeb As Long, iFin As Long
    Dim y As Long, m As Long
    Dim dMois As Double, fMois As Double, a As Double, b As Double
    Dim total As Double

    If fin < deb Then Exit Function
    iDeb = Year(CDate(deb)) * 12 + Month(CDate(deb)) - 1
    iFin = Year(CDate(fin)) * 12 + Month(CDate(fin)) - 1

    For i = iDeb To iFin
        y = i \ 12
        m = (i Mod 12) + 1
        dMois = CDbl(DateSerial(y, m, 1))
        fMois = CDbl(DateSerial(y, m + 1, 0))
        a = deb
        If dMois > a Then a = dMois
        b = fin
        If fMois < b Then b = fMois
        total = total + (b - a + 1) / (fMois - dMois + 1)
    Next i

    NbMoisPeriode = total

End Function

' Texte de la formule Excel du nombre de mois, pour deux cellules données (par exemple
' "[@DateDebut]" et "[@DateFin]" dans le tableau, ou "C8" et "C9" dans le formulaire).
' Syntaxe anglaise (propriété .Formula), quelle que soit la langue d'Excel.
Public Function FormuleNbMois(ByVal refDebut As String, ByVal refFin As String) As String

    Dim f As String

    f = "IF(AND(YEAR({d})=YEAR({f}),MONTH({d})=MONTH({f}))," & _
        "({f}-{d}+1)/DAY(EOMONTH({d},0))," & _
        "(EOMONTH({d},0)-{d}+1)/DAY(EOMONTH({d},0))" & _
        "+((YEAR({f})-YEAR({d}))*12+MONTH({f})-MONTH({d})-1)" & _
        "+DAY({f})/DAY(EOMONTH({f},0)))"
    f = Replace(f, "{d}", refDebut)
    f = Replace(f, "{f}", refFin)
    FormuleNbMois = f

End Function

' Pose les colonnes calculées de TblReleves (Operateur, NbMois, Abonnement, Montant).
' Appelée à l'installation et après chaque ajout de ligne : les formules sont ainsi
' toujours présentes, même si Excel ne les a pas recopiées tout seul.
Public Sub PoserFormulesReleves(ByVal lo As ListObject)

    Dim f As String

    If lo.DataBodyRange Is Nothing Then Exit Sub

    ' Opérateur : le contrat du même type en vigueur à la date de début.
    ' (Sans conséquence si TblContrats n'existe pas encore : la colonne reste vide.)
    f = "=IFERROR(LOOKUP(2,1/((" & mod_VarGlobales.NOM_TABLE_CONTRATS & "[TypeConso]=[@TypeConso])*(" & _
        mod_VarGlobales.NOM_TABLE_CONTRATS & "[DateDebut]<=[@DateDebut])*((" & _
        mod_VarGlobales.NOM_TABLE_CONTRATS & "[DateFin]="""")+(" & _
        mod_VarGlobales.NOM_TABLE_CONTRATS & "[DateFin]>=[@DateDebut])>0))," & _
        mod_VarGlobales.NOM_TABLE_CONTRATS & "[Operateur]),"""")"
    On Error Resume Next
    lo.ListColumns("Operateur").DataBodyRange.Formula = f
    On Error GoTo 0

    f = "=IF(OR([@DateDebut]="""",[@DateFin]="""",[@DateFin]<[@DateDebut]),""""," & _
        FormuleNbMois("[@DateDebut]", "[@DateFin]") & ")"
    lo.ListColumns("NbMois").DataBodyRange.Formula = f

    lo.ListColumns("Abonnement").DataBodyRange.Formula = _
        "=IF([@NbMois]="""","""",ROUND(N([@AbonnementMensuel])*[@NbMois],2))"

    lo.ListColumns("Montant").DataBodyRange.Formula = _
        "=IF([@NbMois]="""","""",ROUND(N([@Quantite])*N([@PrixUnitaire])+[@Abonnement],2))"

End Sub

' Indice du contrat du type donné en vigueur le jour donné (0 = aucun).
Private Function ContratPour(ByVal typ As String, ByVal jour As Double) As Long
    Dim i As Long
    For i = 1 To c_Nb
        If StrComp(c_Typ(i), typ, vbTextCompare) = 0 Then
            If jour >= c_Deb(i) Then
                If c_Fin(i) = 0 Or jour <= c_Fin(i) Then
                    ContratPour = i
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Private Function NomOperateur(ByVal typ As String, ByVal jour As Double) As String
    Dim k As Long
    k = ContratPour(typ, jour)
    If k > 0 Then
        NomOperateur = c_Oper(k)
    End If
    If NomOperateur = "" Then NomOperateur = LibSansContrat()
End Function

' Indice du relevé du type donné qui se termine le plus tard (0 = aucun). Si
' contrat > 0, seuls les relevés relevant de ce contrat sont considérés.
Private Function DernierReleve(ByVal typ As String, ByVal contrat As Long) As Long
    Dim j As Long
    Dim meilleur As Double
    For j = 1 To r_Nb
        If StrComp(r_Typ(j), typ, vbTextCompare) = 0 Then
            If contrat = 0 Or ContratPour(typ, r_Deb(j)) = contrat Then
                If r_Fin(j) > meilleur Then
                    meilleur = r_Fin(j)
                    DernierReleve = j
                End If
            End If
        End If
    Next j
End Function

' Unité à afficher pour un type : celle du contrat le plus récent, sinon l'unité par défaut.
Private Function UniteDuType(ByVal typ As String) As String
    Dim i As Long, k As Long
    Dim meilleur As Double
    For i = 1 To c_Nb
        If StrComp(c_Typ(i), typ, vbTextCompare) = 0 And c_Unite(i) <> "" Then
            If k = 0 Or c_Deb(i) >= meilleur Then
                k = i
                meilleur = c_Deb(i)
            End If
        End If
    Next i
    If k > 0 Then UniteDuType = c_Unite(k) Else UniteDuType = UniteDefaut(typ)
End Function


' =====================================================================================
' ÉCRAN frm_Consommations : CALCUL, TABLEAU, GRAPHIQUE
' =====================================================================================
Private Function FeuilleConso() As Worksheet
    Set FeuilleConso = mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_CONSO)
End Function

Public Sub ActualiserConsommations()

    On Error GoTo Erreur
    Application.ScreenUpdating = False
    RafraichirConso
    Application.ScreenUpdating = True
    Exit Sub

Erreur:
    Application.ScreenUpdating = True
    MsgBox mod_Display.FR("Erreur pendant le calcul des consommations.") & vbLf & _
           "Erreur " & Err.Number & " : " & Err.Description, vbExclamation, mod_Display.FR("Suivi des consommations")

End Sub

Private Sub RafraichirConso()

    Dim ws As Worksheet
    Dim typ As String, indic As String, indicSerie As String, periode As String, forme As String, repart As String
    Dim annee As Long, mois As Long, idxDeb As Long, nbM As Long
    Dim tous As Boolean, prorata As Boolean
    Dim note As String
    Dim serNoms() As String, nS As Long
    Dim aQ() As Double, aQP() As Double, aAbo() As Double, aHas() As Boolean
    Dim tabT() As Variant, tabH() As Variant
    Dim p As Long, s As Long, ent As Long, idx As Long
    Dim q As Double, qp As Double, ab As Double, has As Boolean
    Dim totQ As Double, totQP As Double, totAbo As Double
    Dim oper As String
    Dim nbLignesPrises As Long
    Dim unite As String, texte As String

    Set ws = FeuilleConso()
    If ws Is Nothing Or Not ConsoInstallee() Then
        MsgBox mod_Display.FR("Le suivi des consommations n'est pas install{e2} : lancez d'abord InstallerConsommations."), vbExclamation
        Exit Sub
    End If

    If Not LireParametresConso(ws, typ, indic, periode, annee, mois, forme, repart, idxDeb, nbM) Then Exit Sub
    ChargerContrats
    ChargerReleves
    MettreAJourListesConso ws

    tous = (StrComp(typ, LibTous(), vbTextCompare) = 0)
    prorata = (StrComp(repart, LibRepartProrata(), vbTextCompare) = 0)

    ' Combinaisons sans sens : pour "tous les types", pas de quantité ni de prix (unités
    ' différentes) et pas de courbe du prix.
    indicSerie = indic
    If tous Then
        If StrComp(indic, LibIndQuantite(), vbTextCompare) = 0 Or StrComp(indic, LibIndPrix(), vbTextCompare) = 0 Then
            indic = LibIndMontant()
            indicSerie = indic
            note = " Pour tous les types, seuls l'abonnement et le montant sont comparables : le montant est affich{e2}."
        End If
        If StrComp(forme, LibFormeMixte(), vbTextCompare) = 0 Then
            forme = LibFormeBarres()
            note = note & " La courbe du prix n'existe que pour un type choisi."
        End If
    ElseIf StrComp(forme, LibFormeMixte(), vbTextCompare) = 0 And StrComp(indic, LibIndPrix(), vbTextCompare) = 0 Then
        indicSerie = LibIndQuantite()      ' le prix est déjà la courbe : les barres montrent la quantité
    End If

    CalculerMois typ, tous, prorata, idxDeb, nbM, serNoms, nS, aQ, aQP, aAbo, aHas, nbLignesPrises

    ' ---------------- Tableau mensuel (visible) et données du graphique (masquées) ----------------
    ent = CO_LIGNE_ENTETE
    ws.Range(ws.Cells(ent + 1, 2), ws.Cells(ent + CO_NB_LIGNES_MAX + 1, 7)).ClearContents
    ws.Range(ws.Cells(ent + 1, 2), ws.Cells(ent + CO_NB_LIGNES_MAX + 1, 7)).Font.Bold = False
    ws.Range(ws.Cells(ent, CO_COL_H_LIB), ws.Cells(ent + CO_NB_LIGNES_MAX, CO_COL_H_SERIE + CO_NB_SERIES - 1)).ClearContents

    ReDim tabT(1 To nbM + 1, 1 To 6)
    ReDim tabH(1 To nbM, 1 To 2 + CO_NB_SERIES)

    For p = 1 To nbM
        idx = idxDeb + p - 1
        tabT(p, 1) = DateSerial(idx \ 12, (idx Mod 12) + 1, 1)
        tabH(p, 1) = tabT(p, 1)
        q = 0: qp = 0: ab = 0: oper = ""
        For s = 1 To nS
            q = q + aQ(p, s)
            qp = qp + aQP(p, s)
            ab = ab + aAbo(p, s)
            If aHas(p, s) And Not tous Then
                If oper <> "" Then oper = oper & " / "
                oper = oper & serNoms(s)
            End If
            tabH(p, 2 + s) = ValeurIndic(indicSerie, aQ(p, s), aQP(p, s), aAbo(p, s), aHas(p, s))
        Next s
        has = (q <> 0 Or qp <> 0 Or ab <> 0)
        If has Then
            If Not tous Then
                tabT(p, 2) = q
                If q > 0 Then
                    tabT(p, 3) = qp / q
                    tabH(p, 2) = qp / q
                End If
            End If
            tabT(p, 4) = ab
            tabT(p, 5) = qp + ab
            tabT(p, 6) = oper
        End If
        totQ = totQ + q
        totQP = totQP + qp
        totAbo = totAbo + ab
    Next p

    ' Ligne de total, juste sous le dernier mois.
    tabT(nbM + 1, 1) = "Total"
    If Not tous Then
        tabT(nbM + 1, 2) = totQ
        If totQ > 0 Then tabT(nbM + 1, 3) = totQP / totQ
    End If
    tabT(nbM + 1, 4) = totAbo
    tabT(nbM + 1, 5) = totQP + totAbo

    ' Colonne des mois au format "mmm yyyy" partout (le mot "Total" s'y affiche tel quel).
    ws.Range(ws.Cells(ent + 1, 2), ws.Cells(ent + CO_NB_LIGNES_MAX + 1, 2)).NumberFormat = "mmm yyyy"
    ws.Range(ws.Cells(ent + 1, 2), ws.Cells(ent + nbM + 1, 7)).Value = tabT
    ws.Range(ws.Cells(ent + nbM + 1, 2), ws.Cells(ent + nbM + 1, 7)).Font.Bold = True

    ' Données du graphique : en-têtes, puis une ligne par mois.
    ws.Cells(ent, CO_COL_H_LIB).Value = "Mois"
    ws.Cells(ent, CO_COL_H_PRIX).Value = "Prix unitaire"
    For s = 1 To nS
        ws.Cells(ent, CO_COL_H_SERIE + s - 1).Value = serNoms(s)
    Next s
    ws.Range(ws.Cells(ent + 1, CO_COL_H_LIB), ws.Cells(ent + nbM, CO_COL_H_SERIE + CO_NB_SERIES - 1)).Value = tabH
    ws.Range(ws.Cells(ent + 1, CO_COL_H_LIB), ws.Cells(ent + nbM, CO_COL_H_LIB)).NumberFormat = "mmm yyyy"

    ' Libellé de la colonne Quantité / Prix avec l'unité (ou rien pour "tous").
    If tous Then
        ws.Cells(ent, 3).Value = mod_Display.FR("Quantit{e2}")
        ws.Cells(ent, 4).Value = "Prix unitaire"
        unite = ""
    Else
        unite = UniteDuType(typ)
        ws.Cells(ent, 3).Value = mod_Display.FR("Quantit{e2} (") & unite & ")"
        ws.Cells(ent, 4).Value = "Prix (" & SymboleEuro() & "/" & unite & ")"
    End If

    ' Résumé.
    texte = IIf(tous, mod_Display.FR("Tous les types"), typ) & " -- " & _
            Format$(DateSerial(idxDeb \ 12, (idxDeb Mod 12) + 1, 1), "mmm yyyy") & mod_Display.FR(" {a2} ") & _
            Format$(DateSerial((idxDeb + nbM - 1) \ 12, ((idxDeb + nbM - 1) Mod 12) + 1, 1), "mmm yyyy") & _
            " -- " & nbLignesPrises & mod_Display.FR(" relev{e2}(s) pris en compte")
    If nbLignesPrises = 0 Then texte = texte & mod_Display.FR(". Aucun relev{e2} sur cette p{e2}riode.")
    ws.Range(CO_ADR_MESSAGE).Value = texte & mod_Display.FR(note)

    ConstruireGraphiqueConso ws, nbM, nS, serNoms, indic, indicSerie, forme, typ, tous, unite

End Sub

' Lit et contrôle les paramètres de l'écran. Renvoie Faux (avec un message) s'ils sont incorrects.
Private Function LireParametresConso(ByVal ws As Worksheet, ByRef typ As String, ByRef indic As String, _
                                     ByRef periode As String, ByRef annee As Long, ByRef mois As Long, _
                                     ByRef forme As String, ByRef repart As String, _
                                     ByRef idxDeb As Long, ByRef nbM As Long) As Boolean

    Dim idxFin As Long

    typ = mod_DataStructure.CellText(ws.Range(CO_ADR_TYPE).Value2)
    indic = mod_DataStructure.CellText(ws.Range(CO_ADR_INDIC).Value2)
    periode = mod_DataStructure.CellText(ws.Range(CO_ADR_PERIODE).Value2)
    annee = mod_DataStructure.ToLong(ws.Range(CO_ADR_ANNEE).Value2)
    mois = mod_DataStructure.ToLong(ws.Range(CO_ADR_MOIS).Value2)
    forme = mod_DataStructure.CellText(ws.Range(CO_ADR_FORME).Value2)
    repart = mod_DataStructure.CellText(ws.Range(CO_ADR_REPART).Value2)

    If typ = "" Then typ = LibTous()
    If indic = "" Then indic = LibIndMontant()
    If forme = "" Then forme = LibFormeBarres()
    If repart = "" Then repart = LibRepartProrata()

    If annee < 1900 Or annee > 2200 Then
        MsgBox mod_Display.FR("L'ann{e2}e de r{e2}f{e2}rence est incorrecte."), vbExclamation
        Exit Function
    End If
    If mois < 1 Or mois > 12 Then
        MsgBox mod_Display.FR("Le mois de r{e2}f{e2}rence est incorrect (1 {a2} 12)."), vbExclamation
        Exit Function
    End If

    idxFin = annee * 12 + mois - 1
    If StrComp(periode, LibPeriodeAnnee(), vbTextCompare) = 0 Then
        idxDeb = annee * 12
        nbM = 12
    ElseIf StrComp(periode, LibPeriode36(), vbTextCompare) = 0 Then
        idxDeb = idxFin - 35
        nbM = 36
    ElseIf StrComp(periode, LibPeriode60(), vbTextCompare) = 0 Then
        idxDeb = idxFin - 59
        nbM = 60
    ElseIf StrComp(periode, LibPeriode120(), vbTextCompare) = 0 Then
        idxDeb = idxFin - 119
        nbM = 120
    Else
        idxDeb = idxFin - 11
        nbM = 12
    End If

    LireParametresConso = True

End Function

' Répartit les relevés sur les mois de la période. Pour chaque (mois, série) :
'   aQ = quantité, aQP = quantité x prix (consommation facturée), aAbo = abonnement.
' Série = opérateur (un seul type choisi) ou type (tous les types).
Private Sub CalculerMois(ByVal typ As String, ByVal tous As Boolean, ByVal prorata As Boolean, _
                         ByVal idxDeb As Long, ByVal nbM As Long, _
                         ByRef serNoms() As String, ByRef nS As Long, _
                         ByRef aQ() As Double, ByRef aQP() As Double, ByRef aAbo() As Double, _
                         ByRef aHas() As Boolean, ByRef nbPris As Long)

    Dim dic As Object
    Dim j As Long, i As Long, s As Long, p As Long
    Dim cle As String
    Dim iDeb As Long, iFin As Long, y As Long, m As Long
    Dim dMois As Double, fMois As Double, a As Double, b As Double
    Dim jours As Double, totalJours As Double
    Dim q As Double, qp As Double, ab As Double

    ReDim serNoms(1 To CO_NB_SERIES)
    ReDim aQ(1 To nbM, 1 To CO_NB_SERIES)
    ReDim aQP(1 To nbM, 1 To CO_NB_SERIES)
    ReDim aAbo(1 To nbM, 1 To CO_NB_SERIES)
    ReDim aHas(1 To nbM, 1 To CO_NB_SERIES)
    nS = 0
    nbPris = 0

    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = 1     ' sans tenir compte des majuscules

    For j = 1 To r_Nb
        If tous Or StrComp(r_Typ(j), typ, vbTextCompare) = 0 Then

            If tous Then cle = r_Typ(j) Else cle = NomOperateur(r_Typ(j), r_Deb(j))

            ' Série de ce relevé (au-delà de CO_NB_SERIES, tout le reste est regroupé dans "Autres").
            If dic.Exists(cle) Then
                s = dic(cle)
            ElseIf nS < CO_NB_SERIES Then
                nS = nS + 1
                serNoms(nS) = cle
                dic.Add cle, nS
                s = nS
            Else
                s = CO_NB_SERIES
                serNoms(s) = LibAutres()
                dic.Add cle, s
            End If

            iDeb = Year(CDate(r_Deb(j))) * 12 + Month(CDate(r_Deb(j))) - 1
            iFin = Year(CDate(r_Fin(j))) * 12 + Month(CDate(r_Fin(j))) - 1

            If prorata Then
                totalJours = r_Fin(j) - r_Deb(j) + 1
                For i = iDeb To iFin
                    p = i - idxDeb + 1
                    If p >= 1 And p <= nbM Then
                        y = i \ 12
                        m = (i Mod 12) + 1
                        dMois = CDbl(DateSerial(y, m, 1))
                        fMois = CDbl(DateSerial(y, m + 1, 0))
                        a = r_Deb(j)
                        If dMois > a Then a = dMois
                        b = r_Fin(j)
                        If fMois < b Then b = fMois
                        jours = b - a + 1
                        aQ(p, s) = aQ(p, s) + r_Qte(j) * jours / totalJours
                        aQP(p, s) = aQP(p, s) + r_Qte(j) * r_Prix(j) * jours / totalJours
                        aAbo(p, s) = aAbo(p, s) + r_Abo(j) * jours / (fMois - dMois + 1)
                        aHas(p, s) = True
                    End If
                Next i
                If iFin >= idxDeb And iDeb <= idxDeb + nbM - 1 Then nbPris = nbPris + 1
            Else
                p = iFin - idxDeb + 1
                If p >= 1 And p <= nbM Then
                    aQ(p, s) = aQ(p, s) + r_Qte(j)
                    aQP(p, s) = aQP(p, s) + r_Qte(j) * r_Prix(j)
                    aAbo(p, s) = aAbo(p, s) + r_Abo(j) * NbMoisPeriode(r_Deb(j), r_Fin(j))
                    aHas(p, s) = True
                    nbPris = nbPris + 1
                End If
            End If
        End If
    Next j

End Sub

Private Function ValeurIndic(ByVal indic As String, ByVal q As Double, ByVal qp As Double, _
                             ByVal ab As Double, ByVal has As Boolean) As Variant

    If Not has Then Exit Function       ' Empty : le graphique laisse un trou

    If StrComp(indic, LibIndQuantite(), vbTextCompare) = 0 Then
        ValeurIndic = q
    ElseIf StrComp(indic, LibIndPrix(), vbTextCompare) = 0 Then
        If q > 0 Then ValeurIndic = qp / q
    ElseIf StrComp(indic, LibIndAbo(), vbTextCompare) = 0 Then
        ValeurIndic = ab
    Else
        ValeurIndic = qp + ab
    End If

End Function

' Années proposées : de la plus ancienne année de relevé (ou de l'année en cours) à l'année suivante.
Private Sub MettreAJourListesConso(ByVal ws As Worksheet)

    Dim i As Long, anMin As Long, anMax As Long, an As Long

    anMin = Year(Date)
    anMax = Year(Date) + 1
    For i = 1 To r_Nb
        an = Year(CDate(r_Deb(i)))
        If an < anMin Then anMin = an
        an = Year(CDate(r_Fin(i)))
        If an + 1 > anMax Then anMax = an + 1
    Next i

    ' Périodes proposées (liste reposée à chaque actualisation : un classeur déjà installé
    ' reçoit ainsi les nouvelles périodes sans réinstallation).
    ws.Range(ws.Cells(2, CO_COL_LISTE_PERIODE), ws.Cells(10, CO_COL_LISTE_PERIODE)).ClearContents
    ws.Cells(2, CO_COL_LISTE_PERIODE).Value = LibPeriode12()
    ws.Cells(3, CO_COL_LISTE_PERIODE).Value = LibPeriode36()
    ws.Cells(4, CO_COL_LISTE_PERIODE).Value = LibPeriode60()
    ws.Cells(5, CO_COL_LISTE_PERIODE).Value = LibPeriode120()
    ws.Cells(6, CO_COL_LISTE_PERIODE).Value = LibPeriodeAnnee()
    ThisWorkbook.Names.Add Name:="ListePeriodeConso", RefersTo:="='" & ws.Name & "'!" & _
        ws.Range(ws.Cells(2, CO_COL_LISTE_PERIODE), ws.Cells(6, CO_COL_LISTE_PERIODE)).Address(True, True)

    ' Mise en forme de la zone du tableau jusqu'à CO_NB_LIGNES_MAX lignes.
    mod_InstallCommun.PoserCorpsTableauLecture ws.Range(ws.Cells(CO_LIGNE_ENTETE + 1, 2), ws.Cells(CO_LIGNE_ENTETE + CO_NB_LIGNES_MAX + 1, 7))
    ws.Range(ws.Cells(CO_LIGNE_ENTETE + 1, 2), ws.Cells(CO_LIGNE_ENTETE + CO_NB_LIGNES_MAX + 1, 2)).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(CO_LIGNE_ENTETE + 1, 3), ws.Cells(CO_LIGNE_ENTETE + CO_NB_LIGNES_MAX + 1, 3)).NumberFormat = "#,##0.00"
    ws.Range(ws.Cells(CO_LIGNE_ENTETE + 1, 4), ws.Cells(CO_LIGNE_ENTETE + CO_NB_LIGNES_MAX + 1, 4)).NumberFormat = "#,##0.0000"
    ws.Range(ws.Cells(CO_LIGNE_ENTETE + 1, 5), ws.Cells(CO_LIGNE_ENTETE + CO_NB_LIGNES_MAX + 1, 6)).NumberFormat = "#,##0.00"

    ws.Range(ws.Cells(2, CO_COL_LISTE_ANNEE), ws.Cells(60, CO_COL_LISTE_ANNEE)).ClearContents
    For an = anMin To anMax
        If an - anMin + 2 <= 60 Then ws.Cells(an - anMin + 2, CO_COL_LISTE_ANNEE).Value = an
    Next an

End Sub

Private Sub ConstruireGraphiqueConso(ByVal ws As Worksheet, ByVal nbM As Long, ByVal nS As Long, _
                                     ByRef serNoms() As String, ByVal indic As String, ByVal indicSerie As String, _
                                     ByVal forme As String, ByVal typ As String, ByVal tous As Boolean, _
                                     ByVal unite As String)

    Dim co As ChartObject
    Dim se As Series
    Dim s As Long, ent As Long
    Dim courbes As Boolean, mixte As Boolean
    Dim xr As Range

    On Error Resume Next
    Set co = ws.ChartObjects(CO_GRAPH)
    On Error GoTo 0
    If co Is Nothing Then Exit Sub

    Do While co.Chart.SeriesCollection.Count > 0
        co.Chart.SeriesCollection(1).Delete
    Loop

    co.Chart.HasTitle = True
    If nS = 0 Then
        co.Chart.ChartType = xlColumnStacked
        co.Chart.ChartTitle.Text = mod_Display.FR("Aucun relev{e2} sur la p{e2}riode")
        Exit Sub
    End If

    ent = CO_LIGNE_ENTETE
    Set xr = ws.Range(ws.Cells(ent + 1, CO_COL_H_LIB), ws.Cells(ent + nbM, CO_COL_H_LIB))
    courbes = (StrComp(forme, LibFormeCourbes(), vbTextCompare) = 0)
    mixte = (StrComp(forme, LibFormeMixte(), vbTextCompare) = 0)

    If courbes Then co.Chart.ChartType = xlLineMarkers Else co.Chart.ChartType = xlColumnStacked

    For s = 1 To nS
        Set se = co.Chart.SeriesCollection.NewSeries
        se.Name = serNoms(s)
        se.Values = ws.Range(ws.Cells(ent + 1, CO_COL_H_SERIE + s - 1), ws.Cells(ent + nbM, CO_COL_H_SERIE + s - 1))
        se.XValues = xr
        If courbes Then se.ChartType = xlLineMarkers Else se.ChartType = xlColumnStacked
    Next s

    If mixte Then
        Set se = co.Chart.SeriesCollection.NewSeries
        se.Name = "Prix unitaire"
        se.Values = ws.Range(ws.Cells(ent + 1, CO_COL_H_PRIX), ws.Cells(ent + nbM, CO_COL_H_PRIX))
        se.XValues = xr
        se.ChartType = xlLineMarkers
        se.AxisGroup = xlSecondary
    End If

    co.Chart.HasLegend = True
    co.Chart.Legend.Position = xlLegendPositionBottom
    co.Chart.DisplayBlanksAs = xlNotPlotted

    ' Axe des mois : étiquettes textuelles (un graduation par mois), pas d'axe temporel.
    On Error Resume Next
    co.Chart.Axes(xlCategory).CategoryType = xlCategoryScale
    co.Chart.Axes(xlCategory).TickLabels.NumberFormat = "mmm yy"
    If nbM <= 12 Then
        co.Chart.Axes(xlCategory).TickLabelSpacing = 1
    ElseIf nbM <= 36 Then
        co.Chart.Axes(xlCategory).TickLabelSpacing = 3
    ElseIf nbM <= 60 Then
        co.Chart.Axes(xlCategory).TickLabelSpacing = 6
    Else
        co.Chart.Axes(xlCategory).TickLabelSpacing = 12
    End If

    ' Titres des axes.
    co.Chart.Axes(xlValue, xlPrimary).HasTitle = True
    co.Chart.Axes(xlValue, xlPrimary).AxisTitle.Text = TitreAxe(indicSerie, unite)
    If mixte Then
        co.Chart.Axes(xlValue, xlSecondary).HasTitle = True
        co.Chart.Axes(xlValue, xlSecondary).AxisTitle.Text = TitreAxe(LibIndPrix(), unite)
    End If
    On Error GoTo 0

    If tous Then
        co.Chart.ChartTitle.Text = mod_Display.FR("Tous les types -- ") & indic
    Else
        co.Chart.ChartTitle.Text = typ & " -- " & indic
        If mixte Then co.Chart.ChartTitle.Text = typ & mod_Display.FR(" -- ") & indicSerie & " et prix unitaire"
    End If

End Sub

Private Function TitreAxe(ByVal indic As String, ByVal unite As String) As String
    If StrComp(indic, LibIndQuantite(), vbTextCompare) = 0 Then
        TitreAxe = indic & IIf(unite <> "", " (" & unite & ")", "")
    ElseIf StrComp(indic, LibIndPrix(), vbTextCompare) = 0 Then
        TitreAxe = indic & " (" & SymboleEuro() & IIf(unite <> "", "/" & unite, "") & ")"
    Else
        TitreAxe = indic & " (" & SymboleEuro() & ")"
    End If
End Function


' =====================================================================================
' NAVIGATION ENTRE LES ÉCRANS
' =====================================================================================
' Affiche un écran du suivi (graphique, relevés ou contrats) et masque les deux autres
' ainsi que la feuille Synthese.
Private Sub MontrerEcran(ByVal nomFeuille As String)

    Dim ws As Worksheet, autre As Worksheet, wsSyn As Worksheet
    Dim noms As Variant
    Dim i As Long

    Set ws = mod_InstallCommun.TrouverFeuille(nomFeuille)
    If ws Is Nothing Then
        MsgBox mod_Display.FR("Le suivi des consommations n'est pas install{e2} : lancez d'abord InstallerConsommations (Ctrl+G)."), vbExclamation
        Exit Sub
    End If

    ws.Visible = xlSheetVisible
    ws.Activate
    mod_InstallCommun.MasquerQuadrillage

    noms = Array(mod_VarGlobales.NOM_FEUILLE_CONSO, mod_VarGlobales.NOM_FEUILLE_RELEVES, mod_VarGlobales.NOM_FEUILLE_CONTRATS)
    For i = LBound(noms) To UBound(noms)
        If noms(i) <> nomFeuille Then
            Set autre = mod_InstallCommun.TrouverFeuille(CStr(noms(i)))
            If Not autre Is Nothing Then mod_Aide.MasquerFormulaire autre
        End If
    Next i

    Set wsSyn = mod_Criteres.GetFeuille(mod_VarGlobales.NOM_FEUILLE_SYNTHESE)
    If Not wsSyn Is Nothing Then wsSyn.Visible = xlSheetVeryHidden

End Sub

' Ouverture depuis la feuille Synthese.
Public Sub AfficherConsommations()
    MontrerEcran mod_VarGlobales.NOM_FEUILLE_CONSO
    If Not FeuilleConso() Is Nothing Then ActualiserConsommations
End Sub

' Retour sur l'écran graphique depuis une liste : il est recalculé.
Public Sub RetourConsommations()
    MontrerEcran mod_VarGlobales.NOM_FEUILLE_CONSO
    If Not FeuilleConso() Is Nothing Then ActualiserConsommations
End Sub

Public Sub AfficherReleves()
    MontrerEcran mod_VarGlobales.NOM_FEUILLE_RELEVES
End Sub

Public Sub AfficherContrats()
    MontrerEcran mod_VarGlobales.NOM_FEUILLE_CONTRATS
End Sub

' Sortie vers Synthese (bouton Sortir de l'écran graphique).
Public Sub SortirConsommations()

    Dim wsSyn As Worksheet, ws As Worksheet
    Dim noms As Variant
    Dim i As Long

    Set wsSyn = mod_Criteres.GetFeuille(mod_VarGlobales.NOM_FEUILLE_SYNTHESE)
    If Not wsSyn Is Nothing Then
        wsSyn.Visible = xlSheetVisible
        wsSyn.Activate
    End If

    noms = Array(mod_VarGlobales.NOM_FEUILLE_CONSO, mod_VarGlobales.NOM_FEUILLE_RELEVES, mod_VarGlobales.NOM_FEUILLE_CONTRATS)
    For i = LBound(noms) To UBound(noms)
        Set ws = mod_InstallCommun.TrouverFeuille(CStr(noms(i)))
        If Not ws Is Nothing Then mod_Aide.MasquerFormulaire ws
    Next i

End Sub


' =====================================================================================
' SAISIE D'UN RELEVÉ (formulaire frm_Releve)
' =====================================================================================
Public Sub SaisirReleve()

    Dim r As TypeReleve
    Dim suivant As Boolean
    Dim nb As Long

    If Not ConsoInstallee() Then
        MsgBox mod_Display.FR("Le suivi des consommations n'est pas install{e2}. Ex{e2}cutez d'abord la macro InstallerConsommations."), vbExclamation
        Exit Sub
    End If

    ChargerContrats
    If c_Nb = 0 Then
        MsgBox mod_Display.FR("Aucun contrat n'est d{e2}clar{e2}. Cr{e2}ez d'abord un contrat (bouton Contrats) : il donne l'op{e2}rateur et l'unit{e2} des relev{e2}s."), _
               vbInformation, mod_Display.FR("Suivi des consommations")
        Exit Sub
    End If

    Do
        suivant = False
        If Not OuvrirFormReleve(r, suivant) Then Exit Do
        If Not EnregistrerReleve(r) Then Exit Do
        nb = nb + 1
        If Not suivant Then Exit Do
        ' Segment suivant de la même facture : il commence le lendemain, avec les mêmes prix.
        r.Debut = r.Fin + 1
        r.Fin = 0
        r.Qte = 0
    Loop

    If nb > 0 Then
        If ActiveSheet.Name = mod_VarGlobales.NOM_FEUILLE_CONSO Then ActualiserConsommations
        MsgBox nb & mod_Display.FR(" relev{e2}(s) enregistr{e2}(s)."), vbInformation, mod_Display.FR("Relev{e2} enregistr{e2}")
    End If

End Sub

Private Function EnregistrerReleve(ByRef r As TypeReleve) As Boolean

    Dim lo As ListObject
    Dim lr As ListRow

    Set lo = TableReleves()
    If lo Is Nothing Then Exit Function

    Set lr = NouvelleLigne(lo)
    With lr.Range
        .Cells(1, ColIdx(lo, "ID_Releve")).Value = ProchainID(lo, "ID_Releve")
        .Cells(1, ColIdx(lo, "TypeConso")).Value = r.Typ
        .Cells(1, ColIdx(lo, "DateDebut")).Value = r.Debut
        .Cells(1, ColIdx(lo, "DateFin")).Value = r.Fin
        .Cells(1, ColIdx(lo, "Quantite")).Value = r.Qte
        .Cells(1, ColIdx(lo, "PrixUnitaire")).Value = r.Prix
        .Cells(1, ColIdx(lo, "AbonnementMensuel")).Value = r.AboMensuel
        .Cells(1, ColIdx(lo, "Facture")).Value = r.Facture
        .Cells(1, ColIdx(lo, "Commentaire")).Value = r.Commentaire
    End With
    PoserFormulesReleves lo

    EnregistrerReleve = True

End Function

' Ouvre le formulaire prérempli avec "r" et attend sa fermeture. Renvoie Vrai si
' l'opérateur a validé ; "suivant" vaut Vrai s'il a demandé la saisie du segment suivant.
Private Function OuvrirFormReleve(ByRef r As TypeReleve, ByRef suivant As Boolean) As Boolean

    Dim ws As Worksheet

    On Error GoTo Erreur

    Set ws = mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_RELEVE)
    If ws Is Nothing Then
        MsgBox mod_Display.FR("La feuille '") & mod_VarGlobales.NOM_FEUILLE_RELEVE & mod_Display.FR("' est introuvable.") & vbCrLf & _
               mod_Display.FR("Ex{e2}cutez d'abord la macro InstallerConsommations."), vbExclamation
        Exit Function
    End If

    g_RelNomFeuillePrec = ActiveSheet.Name
    g_RelValide = False
    g_RelSuivant = False
    RemplirChampsRel ws, r

    ws.Visible = xlSheetVisible
    ws.Activate
    mod_InstallCommun.MasquerQuadrillage
    ws.Range(RF_ADR_TYPE).Select

    ' Liste des types posée une seconde fois, feuille affichée et active.
    On Error Resume Next
    mod_InstallCommun.PoserListeDeroulante ws.Range(RF_ADR_TYPE), "ListeTypesConso", True
    On Error GoTo Erreur

    Application.EnableEvents = True
    g_RelEnCours = True
    g_RelVerrouActif = True
    Do While g_RelEnCours
        DoEvents
    Loop

    On Error Resume Next
    ThisWorkbook.Worksheets(g_RelNomFeuillePrec).Activate
    On Error GoTo 0

    If g_RelValide Then
        r = g_RelRes
        suivant = g_RelSuivant
        OuvrirFormReleve = True
    End If
    Exit Function

Erreur:
    Application.EnableEvents = True
    g_RelEnCours = False
    g_RelVerrouActif = False
    MsgBox mod_Display.FR("Erreur inattendue dans la saisie du relev{e2} :") & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
    OuvrirFormReleve = False

End Function

Private Sub RemplirChampsRel(ByVal ws As Worksheet, ByRef r As TypeReleve)

    Dim evenementsAvant As Boolean
    Dim numErr As Long

    evenementsAvant = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Sortie

    ws.Range(RF_ADR_TYPE).Value = r.Typ
    If r.Debut <> 0 Then ws.Range(RF_ADR_DEBUT).Value = r.Debut Else ws.Range(RF_ADR_DEBUT).ClearContents
    If r.Fin <> 0 Then ws.Range(RF_ADR_FIN).Value = r.Fin Else ws.Range(RF_ADR_FIN).ClearContents
    If r.Qte <> 0 Then ws.Range(RF_ADR_QTE).Value = r.Qte Else ws.Range(RF_ADR_QTE).ClearContents
    If r.Typ <> "" Then
        ws.Range(RF_ADR_PRIX).Value = r.Prix
        ws.Range(RF_ADR_ABO).Value = r.AboMensuel
    Else
        ws.Range(RF_ADR_PRIX).ClearContents
        ws.Range(RF_ADR_ABO).ClearContents
    End If
    ws.Range(RF_ADR_FACTURE).Value = r.Facture
    ws.Range(RF_ADR_COMMENT).Value = r.Commentaire
    ws.Range(RF_ADR_MESSAGE).Value = ""

    ' Garde-fous de saisie reposés à chaque ouverture (dates, quantité, prix, abonnement).
    mod_InstallCommun.PoserValidationDate ws.Range(RF_ADR_DEBUT)
    mod_InstallCommun.PoserValidationDate ws.Range(RF_ADR_FIN)
    mod_InstallCommun.PoserValidationMontant ws.Range(RF_ADR_QTE), mod_InstallCommun.MONTANT_POSITIF_OU_NUL
    mod_InstallCommun.PoserValidationMontant ws.Range(RF_ADR_PRIX), mod_InstallCommun.MONTANT_POSITIF_OU_NUL
    mod_InstallCommun.PoserValidationMontant ws.Range(RF_ADR_ABO), mod_InstallCommun.MONTANT_POSITIF_OU_NUL

    ' Opérateur et unités (sans toucher aux prix saisis).
    RelRemplirContexte ws, False

Sortie:
    numErr = Err.Number
    Application.EnableEvents = evenementsAvant
    If numErr <> 0 Then Err.Raise numErr, "RemplirChampsRel", Err.Description

End Sub

' Met à jour l'opérateur et les unités affichés selon le type et la date de début.
' Si preremplir : propose aussi la date de début (lendemain du dernier relevé du type),
' le dernier prix et le dernier abonnement connus pour le contrat en vigueur.
Private Sub RelRemplirContexte(ByVal ws As Worksheet, ByVal preremplir As Boolean)

    Dim typ As String, unite As String
    Dim d As Double, f As Double
    Dim k As Long, j As Long
    Dim evenementsAvant As Boolean

    evenementsAvant = Application.EnableEvents
    Application.EnableEvents = False

    typ = mod_DataStructure.CellText(ws.Range(RF_ADR_TYPE).Value)
    If typ = "" Then
        ws.Range(RF_ADR_OPERATEUR).Value = ""
        ws.Range(RF_ADR_UNITE).Value = ""
        ws.Range(RF_ADR_UNITE_PRIX).Value = ""
        ws.Range(RF_ADR_UNITE_ABO).Value = ""
        Application.EnableEvents = evenementsAvant
        Exit Sub
    End If

    ChargerContrats
    ChargerReleves

    d = VersJour(ws.Range(RF_ADR_DEBUT).Value)
    If preremplir And d = 0 Then
        j = DernierReleve(typ, 0)
        If j > 0 Then
            f = r_Fin(j)
            d = f + 1
            ws.Range(RF_ADR_DEBUT).Value = CDate(d)
        End If
    End If
    If d = 0 Then d = CDbl(Date)

    k = ContratPour(typ, d)
    If k > 0 Then
        unite = c_Unite(k)
        ws.Range(RF_ADR_OPERATEUR).Value = c_Oper(k) & " (" & mod_Display.FR("contrat depuis le ") & TexteDate(c_Deb(k)) & ")"
    Else
        ws.Range(RF_ADR_OPERATEUR).Value = mod_Display.FR("Aucun contrat {a2} cette date : cr{e2}ez-le d'abord (bouton Contrats)")
    End If
    If unite = "" Then unite = UniteDefaut(typ)
    ws.Range(RF_ADR_UNITE).Value = unite
    ws.Range(RF_ADR_UNITE_PRIX).Value = SymboleEuro() & " / " & unite
    ws.Range(RF_ADR_UNITE_ABO).Value = SymboleEuro() & " / " & mod_Display.FR("mois")

    If preremplir And k > 0 Then
        j = DernierReleve(typ, k)
        If j > 0 Then
            ws.Range(RF_ADR_PRIX).Value = r_Prix(j)
            ws.Range(RF_ADR_ABO).Value = r_Abo(j)
        Else
            ws.Range(RF_ADR_PRIX).ClearContents
            ws.Range(RF_ADR_ABO).Value = c_Abo(k)
        End If
    End If

    Application.EnableEvents = evenementsAvant

End Sub

' Appelée par le code de la feuille (Worksheet_Change).
Public Sub RelTraiterChangement(ByVal ws As Worksheet, ByVal Target As Range)

    Dim numErr As Long

    If Not g_RelEnCours Then Exit Sub
    If Target.Cells.Count > 1 Then Exit Sub

    On Error GoTo Sortie

    Select Case Target.Address(False, False)
        Case RF_ADR_TYPE
            RelRemplirContexte ws, True
        Case RF_ADR_DEBUT
            RelRemplirContexte ws, False
    End Select

Sortie:
    numErr = Err.Number
    Application.EnableEvents = True
    If numErr <> 0 Then MsgBox mod_Display.FR("Erreur lors du changement de type : ") & Err.Description, vbExclamation

End Sub

' Verrou "modal" : tant que le formulaire est ouvert, impossible de le quitter en changeant d'onglet.
Public Sub RelVerrouiller(ByVal ws As Worksheet)
    If g_RelVerrouActif And Not mod_Aide.AideOuverte() Then ws.Activate
End Sub

Public Sub RelValider()
    RelValiderCommun False
End Sub

Public Sub RelValiderSuivant()
    RelValiderCommun True
End Sub

Private Sub RelValiderCommun(ByVal suivant As Boolean)

    Dim ws As Worksheet
    Dim typ As String
    Dim dDeb As Double, dFin As Double
    Dim vQte As Variant, vPrix As Variant, vAbo As Variant
    Dim nQte As Double, nPrix As Double, nAbo As Double
    Dim dTmp As Date
    Dim k As Long, j As Long
    Dim e As TypeReleve

    If Not g_RelEnCours Then Exit Sub
    On Error GoTo Erreur
    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_RELEVE)

    typ = mod_DataStructure.CellText(ws.Range(RF_ADR_TYPE).Value)
    dDeb = VersJour(ws.Range(RF_ADR_DEBUT).Value)
    dFin = VersJour(ws.Range(RF_ADR_FIN).Value)
    vQte = ws.Range(RF_ADR_QTE).Value
    vPrix = ws.Range(RF_ADR_PRIX).Value
    vAbo = ws.Range(RF_ADR_ABO).Value

    If typ = "" Or Not TypeConnu(typ) Then
        AfficherMessageRel ws, "Choisissez le type de consommation dans la liste."
        Exit Sub
    End If
    If SaisieNonVide(ws.Range(RF_ADR_DEBUT).Value) And Not mod_DataStructure.TryLireDate(ws.Range(RF_ADR_DEBUT).Value, dTmp) Then
        AfficherMessageRel ws, "La date de d{e2}but n'est pas valide. " & mod_DataStructure.MessageDateAttendue()
        Exit Sub
    End If
    If SaisieNonVide(ws.Range(RF_ADR_FIN).Value) And Not mod_DataStructure.TryLireDate(ws.Range(RF_ADR_FIN).Value, dTmp) Then
        AfficherMessageRel ws, "La date de fin n'est pas valide. " & mod_DataStructure.MessageDateAttendue()
        Exit Sub
    End If
    If dDeb = 0 Then
        AfficherMessageRel ws, "Renseignez la date de d{e2}but de la p{e2}riode."
        Exit Sub
    End If
    If dFin = 0 Then
        AfficherMessageRel ws, "Renseignez la date de fin de la p{e2}riode."
        Exit Sub
    End If
    If dFin < dDeb Then
        AfficherMessageRel ws, "La date de fin est ant{e2}rieure {a2} la date de d{e2}but."
        Exit Sub
    End If
    If Not SaisieNonVide(vQte) Then
        AfficherMessageRel ws, "Renseignez la quantit{e2} consomm{e2}e (0 si aucune)."
        Exit Sub
    End If
    If Not mod_DataStructure.TryLireMontant(vQte, nQte) Then
        AfficherMessageRel ws, "La quantit{e2} n'est pas valide. " & mod_DataStructure.MessageMontantAttendu()
        Exit Sub
    End If
    If nQte < 0 Then
        AfficherMessageRel ws, "La quantit{e2} ne peut pas {ea}tre n{e2}gative."
        Exit Sub
    End If
    If Not SaisieNonVide(vPrix) Then
        AfficherMessageRel ws, "Renseignez le prix de l'unit{e2} (0 si l'unit{e2} est gratuite)."
        Exit Sub
    End If
    If Not mod_DataStructure.TryLireMontant(vPrix, nPrix, 4) Then
        AfficherMessageRel ws, "Le prix de l'unit{e2} n'est pas valide (4 d{e2}cimales au maximum). " & mod_DataStructure.MessageMontantAttendu()
        Exit Sub
    End If
    If nPrix < 0 Then
        AfficherMessageRel ws, "Le prix de l'unit{e2} ne peut pas {ea}tre n{e2}gatif."
        Exit Sub
    End If
    If Not SaisieNonVide(vAbo) Then
        nAbo = 0
    ElseIf Not mod_DataStructure.TryLireMontant(vAbo, nAbo) Then
        AfficherMessageRel ws, "L'abonnement mensuel n'est pas valide. " & mod_DataStructure.MessageMontantAttendu()
        Exit Sub
    ElseIf nAbo < 0 Then
        AfficherMessageRel ws, "L'abonnement mensuel ne peut pas {ea}tre n{e2}gatif."
        Exit Sub
    End If

    ChargerContrats
    ChargerReleves

    k = ContratPour(typ, dDeb)
    If k = 0 Then
        AfficherMessageRel ws, "Aucun contrat de ce type ne couvre la date de d{e2}but. Cr{e2}ez-le d'abord (bouton Contrats)."
        Exit Sub
    End If
    If c_Fin(k) <> 0 And dFin > c_Fin(k) Then
        AfficherMessageRel ws, "La p{e2}riode d{e2}passe la fin du contrat (" & TexteDate(c_Fin(k)) & "). Coupez-la en deux relev{e2}s."
        Exit Sub
    End If

    For j = 1 To r_Nb
        If StrComp(r_Typ(j), typ, vbTextCompare) = 0 Then
            If dDeb <= r_Fin(j) And dFin >= r_Deb(j) Then
                AfficherMessageRel ws, "Cette p{e2}riode chevauche un relev{e2} d{e2}j{a2} saisi (du " & TexteDate(r_Deb(j)) & " au " & TexteDate(r_Fin(j)) & ")."
                Exit Sub
            End If
        End If
    Next j

    e.Typ = typ
    e.Debut = CDate(dDeb)
    e.Fin = CDate(dFin)
    e.Qte = nQte
    e.Prix = nPrix
    e.AboMensuel = nAbo
    e.Facture = mod_DataStructure.CellText(ws.Range(RF_ADR_FACTURE).Value)
    e.Commentaire = mod_DataStructure.CellText(ws.Range(RF_ADR_COMMENT).Value)

    g_RelRes = e
    g_RelSuivant = suivant
    g_RelValide = True
    FermerFeuilleRel ws
    Exit Sub

Erreur:
    Application.EnableEvents = True
    MsgBox mod_Display.FR("Erreur dans RelValider : ") & Err.Number & " - " & Err.Description, vbCritical

End Sub

Public Sub RelAnnuler()

    Dim ws As Worksheet

    If Not g_RelEnCours Then Exit Sub
    On Error GoTo Erreur

    If MsgBox(mod_Display.FR("Annuler la saisie ? Rien ne sera enregistr{e2}."), vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_RELEVE)
    g_RelValide = False
    FermerFeuilleRel ws
    Exit Sub

Erreur:
    Application.EnableEvents = True
    MsgBox mod_Display.FR("Erreur dans RelAnnuler : ") & Err.Number & " - " & Err.Description, vbCritical

End Sub

Private Sub AfficherMessageRel(ByVal ws As Worksheet, ByVal texte As String)
    ws.Range(RF_ADR_MESSAGE).Value = mod_Display.FR(texte)
End Sub

' g_RelEnCours doit passer à False AVANT de masquer la feuille, sinon le verrou la réactiverait.
Private Sub FermerFeuilleRel(ByVal ws As Worksheet)
    g_RelEnCours = False
    g_RelVerrouActif = False
    mod_Aide.MasquerFormulaire ws
End Sub


' =====================================================================================
' SAISIE D'UN CONTRAT (formulaire frm_Contrat)
' =====================================================================================
' Ajoute un contrat. Si un contrat du même type est en cours et commence avant le
' nouveau, il est proposé de le clôturer la veille (changement d'opérateur).
Public Sub AjouterContrat()

    Dim c As TypeContrat
    Dim ligneAClore As Long

    If Not ConsoInstallee() Then
        MsgBox mod_Display.FR("Le suivi des consommations n'est pas install{e2}. Ex{e2}cutez d'abord la macro InstallerConsommations."), vbExclamation
        Exit Sub
    End If

    If OuvrirFormContrat(c, ligneAClore) Then
        If EnregistrerContrat(c, ligneAClore) Then
            If ActiveSheet.Name = mod_VarGlobales.NOM_FEUILLE_CONSO Then ActualiserConsommations
            MsgBox mod_Display.FR("Contrat enregistr{e2}."), vbInformation, mod_Display.FR("Suivi des consommations")
        End If
    End If

End Sub

Private Function EnregistrerContrat(ByRef c As TypeContrat, ByVal ligneAClore As Long) As Boolean

    Dim lo As ListObject
    Dim lr As ListRow

    Set lo = TableContrats()
    If lo Is Nothing Then Exit Function

    ' Clôture du contrat en cours (la veille du nouveau).
    If ligneAClore > 0 Then
        lo.ListRows(ligneAClore).Range.Cells(1, ColIdx(lo, "DateFin")).Value = c.Debut - 1
    End If

    Set lr = NouvelleLigne(lo)
    With lr.Range
        .Cells(1, ColIdx(lo, "ID_Contrat")).Value = ProchainID(lo, "ID_Contrat")
        .Cells(1, ColIdx(lo, "TypeConso")).Value = c.Typ
        .Cells(1, ColIdx(lo, "Operateur")).Value = c.Operateur
        .Cells(1, ColIdx(lo, "Unite")).Value = c.Unite
        .Cells(1, ColIdx(lo, "DateDebut")).Value = c.Debut
        If c.Fin <> 0 Then .Cells(1, ColIdx(lo, "DateFin")).Value = c.Fin
        .Cells(1, ColIdx(lo, "AbonnementMensuel")).Value = c.AboMensuel
        .Cells(1, ColIdx(lo, "Reference")).Value = c.Reference
    End With

    EnregistrerContrat = True

End Function

Private Function OuvrirFormContrat(ByRef c As TypeContrat, ByRef ligneAClore As Long) As Boolean

    Dim ws As Worksheet
    Dim evenementsAvant As Boolean

    On Error GoTo Erreur

    Set ws = mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_CONTRAT)
    If ws Is Nothing Then
        MsgBox mod_Display.FR("La feuille '") & mod_VarGlobales.NOM_FEUILLE_CONTRAT & mod_Display.FR("' est introuvable.") & vbCrLf & _
               mod_Display.FR("Ex{e2}cutez d'abord la macro InstallerConsommations."), vbExclamation
        Exit Function
    End If

    g_CtrNomFeuillePrec = ActiveSheet.Name
    g_CtrValide = False
    g_CtrLigneAClore = 0

    evenementsAvant = Application.EnableEvents
    Application.EnableEvents = False
    ws.Range(CF_ADR_TYPE).Value = c.Typ
    ws.Range(CF_ADR_OPERATEUR).Value = c.Operateur
    ws.Range(CF_ADR_UNITE).Value = c.Unite
    If c.Debut <> 0 Then ws.Range(CF_ADR_DEBUT).Value = c.Debut Else ws.Range(CF_ADR_DEBUT).ClearContents
    If c.Fin <> 0 Then ws.Range(CF_ADR_FIN).Value = c.Fin Else ws.Range(CF_ADR_FIN).ClearContents
    If c.AboMensuel <> 0 Then ws.Range(CF_ADR_ABO).Value = c.AboMensuel Else ws.Range(CF_ADR_ABO).ClearContents
    ws.Range(CF_ADR_REFERENCE).Value = c.Reference
    ws.Range(CF_ADR_MESSAGE).Value = ""
    mod_InstallCommun.PoserValidationDate ws.Range(CF_ADR_DEBUT)
    mod_InstallCommun.PoserValidationDate ws.Range(CF_ADR_FIN)
    mod_InstallCommun.PoserValidationMontant ws.Range(CF_ADR_ABO), mod_InstallCommun.MONTANT_POSITIF_OU_NUL
    Application.EnableEvents = evenementsAvant

    ws.Visible = xlSheetVisible
    ws.Activate
    mod_InstallCommun.MasquerQuadrillage
    ws.Range(CF_ADR_TYPE).Select

    On Error Resume Next
    mod_InstallCommun.PoserListeDeroulante ws.Range(CF_ADR_TYPE), "ListeTypesConso", True
    On Error GoTo Erreur

    Application.EnableEvents = True
    g_CtrEnCours = True
    g_CtrVerrouActif = True
    Do While g_CtrEnCours
        DoEvents
    Loop

    On Error Resume Next
    ThisWorkbook.Worksheets(g_CtrNomFeuillePrec).Activate
    On Error GoTo 0

    If g_CtrValide Then
        c = g_CtrRes
        ligneAClore = g_CtrLigneAClore
        OuvrirFormContrat = True
    End If
    Exit Function

Erreur:
    Application.EnableEvents = True
    g_CtrEnCours = False
    g_CtrVerrouActif = False
    MsgBox mod_Display.FR("Erreur inattendue dans la saisie du contrat :") & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
    OuvrirFormContrat = False

End Function

' Appelée par le code de la feuille (Worksheet_Change) : le choix d'un type propose son unité.
Public Sub CtrTraiterChangement(ByVal ws As Worksheet, ByVal Target As Range)

    Dim numErr As Long
    Dim u As String

    If Not g_CtrEnCours Then Exit Sub
    If Target.Cells.Count > 1 Then Exit Sub

    On Error GoTo Sortie

    If Target.Address(False, False) = CF_ADR_TYPE Then
        u = mod_DataStructure.CellText(ws.Range(CF_ADR_UNITE).Value)
        If u = "" Or EstUniteDefaut(u) Then
            Application.EnableEvents = False
            ws.Range(CF_ADR_UNITE).Value = UniteDefaut(mod_DataStructure.CellText(Target.Value))
        End If
    End If

Sortie:
    numErr = Err.Number
    Application.EnableEvents = True
    If numErr <> 0 Then MsgBox mod_Display.FR("Erreur lors du changement de type : ") & Err.Description, vbExclamation

End Sub

Public Sub CtrVerrouiller(ByVal ws As Worksheet)
    If g_CtrVerrouActif And Not mod_Aide.AideOuverte() Then ws.Activate
End Sub

Public Sub CtrValider()

    Dim ws As Worksheet
    Dim typ As String, oper As String
    Dim dDeb As Double, dFin As Double, jFin As Double
    Dim vAbo As Variant
    Dim nAbo As Double
    Dim dTmp As Date
    Dim j As Long, aClore As Long
    Dim e As TypeContrat

    If Not g_CtrEnCours Then Exit Sub
    On Error GoTo Erreur
    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_CONTRAT)

    typ = mod_DataStructure.CellText(ws.Range(CF_ADR_TYPE).Value)
    oper = mod_DataStructure.CellText(ws.Range(CF_ADR_OPERATEUR).Value)
    dDeb = VersJour(ws.Range(CF_ADR_DEBUT).Value)
    dFin = VersJour(ws.Range(CF_ADR_FIN).Value)
    vAbo = ws.Range(CF_ADR_ABO).Value

    If typ = "" Or Not TypeConnu(typ) Then
        AfficherMessageCtr ws, "Choisissez le type de consommation dans la liste."
        Exit Sub
    End If
    If oper = "" Then
        AfficherMessageCtr ws, "Renseignez l'op{e2}rateur (fournisseur)."
        Exit Sub
    End If
    If SaisieNonVide(ws.Range(CF_ADR_DEBUT).Value) And Not mod_DataStructure.TryLireDate(ws.Range(CF_ADR_DEBUT).Value, dTmp) Then
        AfficherMessageCtr ws, "La date de d{e2}but n'est pas valide. " & mod_DataStructure.MessageDateAttendue()
        Exit Sub
    End If
    If SaisieNonVide(ws.Range(CF_ADR_FIN).Value) And Not mod_DataStructure.TryLireDate(ws.Range(CF_ADR_FIN).Value, dTmp) Then
        AfficherMessageCtr ws, "La date de fin n'est pas valide. " & mod_DataStructure.MessageDateAttendue()
        Exit Sub
    End If
    If dDeb = 0 Then
        AfficherMessageCtr ws, "Renseignez la date de d{e2}but du contrat."
        Exit Sub
    End If
    If dFin <> 0 And dFin < dDeb Then
        AfficherMessageCtr ws, "La date de fin est ant{e2}rieure {a2} la date de d{e2}but."
        Exit Sub
    End If
    If Not SaisieNonVide(vAbo) Then
        nAbo = 0
    ElseIf Not mod_DataStructure.TryLireMontant(vAbo, nAbo) Then
        AfficherMessageCtr ws, "L'abonnement mensuel n'est pas valide. " & mod_DataStructure.MessageMontantAttendu()
        Exit Sub
    ElseIf nAbo < 0 Then
        AfficherMessageCtr ws, "L'abonnement mensuel ne peut pas {ea}tre n{e2}gatif."
        Exit Sub
    End If

    ChargerContrats

    ' Un contrat du même type en cours (sans date de fin) qui commence avant : on propose de le clôturer.
    For j = 1 To c_Nb
        If StrComp(c_Typ(j), typ, vbTextCompare) = 0 And c_Fin(j) = 0 And c_Deb(j) < dDeb Then
            If MsgBox(mod_Display.FR("Le contrat de ") & c_Oper(j) & mod_Display.FR(" (depuis le ") & TexteDate(c_Deb(j)) & _
                      mod_Display.FR(") est en cours.") & vbCrLf & vbCrLf & _
                      mod_Display.FR("Le cl{o2}turer le ") & TexteDate(dDeb - 1) & mod_Display.FR(" (veille du nouveau contrat) ?"), _
                      vbYesNo + vbQuestion, mod_Display.FR("Changement d'op{e2}rateur")) = vbNo Then
                AfficherMessageCtr ws, "Un seul contrat peut {ea}tre en cours pour un type : cl{o2}turez l'ancien ou donnez une date de fin."
                Exit Sub
            End If
            aClore = c_Ligne(j)
            c_Fin(j) = dDeb - 1           ' pour le contrôle de chevauchement ci-dessous
            Exit For
        End If
    Next j

    ' Chevauchement avec les autres contrats du même type.
    For j = 1 To c_Nb
        If StrComp(c_Typ(j), typ, vbTextCompare) = 0 Then
            jFin = c_Fin(j)
            If jFin = 0 Then jFin = 2958465          ' 31/12/9999
            If dFin = 0 Then
                If jFin >= dDeb Then
                    AfficherMessageCtr ws, "Ce contrat chevauche celui de " & c_Oper(j) & " (du " & TexteDate(c_Deb(j)) & IIf(c_Fin(j) = 0, " en cours", " au " & TexteDate(c_Fin(j))) & ")."
                    Exit Sub
                End If
            ElseIf dDeb <= jFin And dFin >= c_Deb(j) Then
                AfficherMessageCtr ws, "Ce contrat chevauche celui de " & c_Oper(j) & " (du " & TexteDate(c_Deb(j)) & IIf(c_Fin(j) = 0, " en cours", " au " & TexteDate(c_Fin(j))) & ")."
                Exit Sub
            End If
        End If
    Next j

    e.Typ = typ
    e.Operateur = oper
    e.Unite = mod_DataStructure.CellText(ws.Range(CF_ADR_UNITE).Value)
    If e.Unite = "" Then e.Unite = UniteDefaut(typ)
    e.Debut = CDate(dDeb)
    If dFin <> 0 Then e.Fin = CDate(dFin)
    e.AboMensuel = nAbo
    e.Reference = mod_DataStructure.CellText(ws.Range(CF_ADR_REFERENCE).Value)

    g_CtrRes = e
    g_CtrLigneAClore = aClore
    g_CtrValide = True
    FermerFeuilleCtr ws
    Exit Sub

Erreur:
    Application.EnableEvents = True
    MsgBox mod_Display.FR("Erreur dans CtrValider : ") & Err.Number & " - " & Err.Description, vbCritical

End Sub

Public Sub CtrAnnuler()

    Dim ws As Worksheet

    If Not g_CtrEnCours Then Exit Sub
    On Error GoTo Erreur

    If MsgBox(mod_Display.FR("Annuler la saisie ? Rien ne sera enregistr{e2}."), vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_CONTRAT)
    g_CtrValide = False
    FermerFeuilleCtr ws
    Exit Sub

Erreur:
    Application.EnableEvents = True
    MsgBox mod_Display.FR("Erreur dans CtrAnnuler : ") & Err.Number & " - " & Err.Description, vbCritical

End Sub

Private Sub AfficherMessageCtr(ByVal ws As Worksheet, ByVal texte As String)
    ws.Range(CF_ADR_MESSAGE).Value = mod_Display.FR(texte)
End Sub

Private Sub FermerFeuilleCtr(ByVal ws As Worksheet)
    g_CtrEnCours = False
    g_CtrVerrouActif = False
    mod_Aide.MasquerFormulaire ws
End Sub


' =====================================================================================
' SUPPRESSION (listes frm_Releves et frm_Contrats)
' =====================================================================================
' Numéros (dans le tableau) des lignes touchées par la sélection courante.
Private Function LignesSelectionnees(ByVal lo As ListObject) As Object

    Dim d As Object
    Dim inter As Range, c As Range
    Dim n As Long

    Set d = CreateObject("Scripting.Dictionary")
    Set LignesSelectionnees = d
    If lo.DataBodyRange Is Nothing Then Exit Function

    On Error Resume Next
    Set inter = Intersect(Selection, lo.DataBodyRange)
    On Error GoTo 0
    If inter Is Nothing Then Exit Function

    For Each c In inter.Cells
        n = c.Row - lo.DataBodyRange.Row + 1
        If Not d.Exists(n) Then d.Add n, True
    Next c

End Function

Public Sub SupprimerReleveSelection()

    Dim lo As ListObject
    Dim sel As Object
    Dim i As Long

    Set lo = TableReleves()
    If lo Is Nothing Then Exit Sub
    If ActiveSheet.Name <> mod_VarGlobales.NOM_FEUILLE_RELEVES Then Exit Sub

    Set sel = LignesSelectionnees(lo)
    If sel.Count = 0 Then
        MsgBox mod_Display.FR("S{e2}lectionnez d'abord une ou plusieurs lignes du tableau."), vbInformation
        Exit Sub
    End If

    If MsgBox(sel.Count & mod_Display.FR(" relev{e2}(s) {a2} supprimer ?"), vbYesNo + vbQuestion, _
              mod_Display.FR("Supprimer des relev{e2}s")) = vbNo Then Exit Sub

    For i = lo.ListRows.Count To 1 Step -1
        If sel.Exists(i) Then lo.ListRows(i).Delete
    Next i

End Sub

Public Sub SupprimerContratSelection()

    Dim lo As ListObject
    Dim sel As Object
    Dim i As Long, k As Long, j As Long, nbReleves As Long
    Dim fin As Double
    Dim msg As String

    Set lo = TableContrats()
    If lo Is Nothing Then Exit Sub
    If ActiveSheet.Name <> mod_VarGlobales.NOM_FEUILLE_CONTRATS Then Exit Sub

    Set sel = LignesSelectionnees(lo)
    If sel.Count = 0 Then
        MsgBox mod_Display.FR("S{e2}lectionnez d'abord une ou plusieurs lignes du tableau."), vbInformation
        Exit Sub
    End If

    ' Relevés rattachés aux contrats sélectionnés (même type, début dans la durée du contrat).
    ChargerContrats
    ChargerReleves
    For k = 1 To c_Nb
        If sel.Exists(c_Ligne(k)) Then
            fin = c_Fin(k)
            If fin = 0 Then fin = 2958465
            For j = 1 To r_Nb
                If StrComp(r_Typ(j), c_Typ(k), vbTextCompare) = 0 And r_Deb(j) >= c_Deb(k) And r_Deb(j) <= fin Then nbReleves = nbReleves + 1
            Next j
        End If
    Next k

    msg = sel.Count & mod_Display.FR(" contrat(s) {a2} supprimer ?")
    If nbReleves > 0 Then
        msg = msg & vbCrLf & vbCrLf & mod_Display.FR("ATTENTION : ") & nbReleves & _
              mod_Display.FR(" relev{e2}(s) y sont rattach{e2}s. Ils ne seront pas supprim{e2}s mais appara{i2}tront sans op{e2}rateur ('") & _
              LibSansContrat() & "')."
    End If

    If MsgBox(msg, vbYesNo + vbQuestion, mod_Display.FR("Supprimer des contrats")) = vbNo Then Exit Sub

    For i = lo.ListRows.Count To 1 Step -1
        If sel.Exists(i) Then lo.ListRows(i).Delete
    Next i

End Sub
