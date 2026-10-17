Option Explicit

' ============================================================================
'  MODULE : mod_RemediationOFX
'  RÔLE   : Remédiation des données déjà importées. On fournit un OFX FRAÎCHEMENT
'           exporté; la macro le compare à TblOperations (par ID_Transaction = FITID),
'           AFFICHE les écarts dans le formulaire frm_Remediation_OFX, puis (sur demande)
'           corrige ceux que l'opérateur a laissés sur "Oui".
'
'  FORMULAIRE : feuille frm_Remediation_OFX (construite par mod_InstallRemediationOFX,
'  macro InstallerRemediationOFX), ouverte par le bouton de Synthese. Aide : frm_Remediation_OFX_aide.
'
'  MACROS (boutons du formulaire, ou Ctrl+G) :
'     AfficherRemediationOFX    : ouvre le formulaire depuis Synthese.
'     RemedierDepuisOFX         : choisit l'OFX, analyse, remplit le tableau du formulaire.
'     AppliquerRemediationOFX   : applique les lignes "Corriger = Oui" non encore traitées.
'     AnnulerRemediationOFX     : remet les valeurs d'avant pour les lignes "Appliqué".
'     SortirRemediationOFX      : revient sur Synthese.
'
'  Principe de sécurité :
'   - Les champs d'IDENTITÉ de l'opération (Date_Comptable, Date_Operation, Type, Montant,
'     Num_Cheque) viennent de la banque et ne sont jamais modifiés par l'opérateur : l'OFX fait
'     foi, ces écarts sont pré-cochés "Oui".
'   - Tiers / Notes / Catégorie peuvent avoir été corrigés légitimement par l'opérateur : ils ne
'     sont pré-cochés que pour une ligne manifestement MÉLANGÉE (Tiers différent de l'OFX ET
'     Notes qui reprennent le texte bancaire d'une AUTRE opération). Tout le reste est "Non".
'   - Rien n'est écrit pendant l'analyse. Chaque correction est retrouvée par ID_Transaction (jamais
'     par numéro de ligne), vérifiée contre la valeur vue à l'analyse, et mémorisée pour pouvoir
'     être annulée.
'   - Les opérations de l'OFX absentes du classeur sont seulement SIGNALÉES (les ajouter = relancer
'     l'import avec ce même OFX, qui ignore les FITID déjà présents).
'   - Aucune donnée ne quitte le classeur : le rapport est une feuille du classeur.
' ============================================================================

' --- Formulaire frm_Remediation_OFX (mise en page : mod_InstallRemediationOFX) ---
' En-tête commun (mod_InstallCommun) : 1 ligne de boutons + ligne de résumé (le "compteur") :
'   ligne 1 marge / 2 boutons / 3 fine ligne / 4 résumé / 5 fine ligne / 6 Instructions (masquée) /
'   7 fine ligne / 8 en-têtes du tableau / 9 première ligne d'écart
Public Const NOM_FEUILLE_REMEDIATION As String = "frm_Remediation_OFX"
Public Const REM_LIGNE_RESUME As Long = 4
Public Const REM_LIGNE_ENTETES As Long = 8
Private Const LIGNE_ENTETES As Long = REM_LIGNE_ENTETES
Private Const C_CORRIGER As Long = 2   ' la colonne A est une marge
Private Const C_STATUT As Long = 3
Private Const C_ID As Long = 4
Private Const C_DATE As Long = 5
Private Const C_TIERS As Long = 6
Private Const C_CHAMP As Long = 7
Private Const C_ACTUEL As Long = 8
Private Const C_PROPOSE As Long = 9
Private Const C_NATURE As Long = 10
Private Const C_COMMENT As Long = 11
Private Const C_CODE As Long = 12      ' technique (masquée) : DATE, DATEOP, TYPE, MONTANT, CHEQUE, TIERS, NOTES, CATSOUS, BUDGET, POINTE, MANQUANTE
Private Const C_AVANT As Long = 13     ' technique (masquée) : valeur d'avant, forme brute (restauration)
Private Const C_PROPRAW As Long = 14   ' technique (masquée) : valeur proposée, forme brute
Private Const C_CHEMIN As Long = 16    ' technique (masquée) : chemin de l'OFX analysé (ligne 1)
Private Const NB_COL_RAPPORT As Long = 13   ' nombre de colonnes de C_CORRIGER à C_PROPRAW

Private Const NAT_IDENTITE As String = "Identité (OFX fait foi)"
Private Const NAT_TIERS As String = "Tiers"
Private Const NAT_NOTES As String = "Notes d'une autre opération"
Private Const NAT_CATEGORIE As String = "Catégorie"
Private Const NAT_BUDGET As String = "Budget (recalcul)"
Private Const NAT_POINTE As String = "Pointage"
Private Const NAT_MANQUANTE As String = "Information"

Private Type OpOFX
    Fitid As String
    DatePost As Date
    DateUser As Variant
    TypeOp As String
    Montant As Double
    Tiers As String
    Memo As String
    Cheque As String
    Reconnu As Boolean
    Categorie As String
    Notes As String
End Type

' --- État partagé pendant une analyse / une application ---
Private m_ops() As OpOFX
Private m_nbOps As Long
Private m_dictOFX As Object          ' FITID -> indice dans m_ops
Private m_dictRow As Object          ' ID_Transaction -> numéro de ligne dans TblOperations
Private m_sources() As String
Private m_nbSources As Long
Private m_map As Object              ' CategorieSource -> Categorie & vbTab & SousCategorie
Private m_data As Variant            ' TblOperations.DataBodyRange.Value (analyse uniquement)
Private m_lo As ListObject
Private m_lignes As Collection

Private m_cID As Long, m_cDate As Long, m_cDateOp As Long, m_cType As Long, m_cTiers As Long
Private m_cNotes As Long, m_cMontant As Long, m_cCheque As Long, m_cCat As Long, m_cSous As Long
Private m_cBudget As Long, m_cMois As Long, m_cAnnee As Long, m_cDecal As Long, m_cPointe As Long


' ============================================================================
' MACRO 1 : analyse + rapport (+ application facultative)
' ============================================================================
Public Sub RemedierDepuisOFX()

    Dim cheminOFX As String
    Dim r As Long, nbLignes As Long
    Dim id As String
    Dim nbComparees As Long, nbHorsOFX As Long, nbManquantes As Long
    Dim nbPreOui As Long, nbLignesRapport As Long
    Dim ws As Worksheet
    Dim vu As Object
    Dim k As Long

    Set ws = TrouverFeuilleRapport()
    If ws Is Nothing Then
        MsgBox "Le formulaire " & NOM_FEUILLE_REMEDIATION & " n'est pas installe : lancez InstallerRemediationOFX (Ctrl+G).", _
               vbExclamation, "Remediation OFX"
        Exit Sub
    End If

    cheminOFX = mod_ImportOFX.ChoisirFichier("Fichiers OFX (*.ofx), *.ofx", _
        "Selectionner un OFX recemment exporte (comparaison avec le classeur)")
    If cheminOFX = "" Then Exit Sub

    Set m_lo = mod_DonneesTable.GetOperationsTable()
    If m_lo Is Nothing Then Exit Sub
    If m_lo.DataBodyRange Is Nothing Then
        MsgBox "TblOperations est vide : rien a comparer.", vbInformation, "Remediation OFX"
        Exit Sub
    End If
    If Not LireIndexColonnes() Then Exit Sub

    LireSourcesEtCorrespondance
    If Not ChargerOFX(cheminOFX) Then Exit Sub

    m_data = m_lo.DataBodyRange.Value
    nbLignes = UBound(m_data, 1)
    Set m_lignes = New Collection
    Set vu = CreateObject("Scripting.Dictionary")

    ' Index ID -> ligne (pour retrouver la ligne "donneuse" d'un mélange)
    Set m_dictRow = CreateObject("Scripting.Dictionary")
    For r = 1 To nbLignes
        id = mod_DataStructure.CellText(m_data(r, m_cID))
        If id <> "" Then
            If Not m_dictRow.Exists(id) Then m_dictRow.Add id, r
        End If
    Next r

    For r = 1 To nbLignes
        id = mod_DataStructure.CellText(m_data(r, m_cID))
        If id <> "" Then
            If m_dictOFX.Exists(id) Then
                k = m_dictOFX(id)
                If Not vu.Exists(id) Then vu.Add id, True
                nbComparees = nbComparees + 1
                AnalyserLigne r, k
            Else
                nbHorsOFX = nbHorsOFX + 1
            End If
        End If
    Next r

    ' Opérations de l'OFX absentes du classeur : simple information.
    For k = 1 To m_nbOps
        If Not m_dictRow.Exists(m_ops(k).Fitid) Then
            nbManquantes = nbManquantes + 1
            AjouterEcart "-", m_ops(k).Fitid, Format$(m_ops(k).DatePost, "dd/mm/yyyy"), m_ops(k).Tiers, _
                "Operation", "(absente du classeur)", "A importer", NAT_MANQUANTE, _
                "Montant " & Format$(m_ops(k).Montant, "0.00") & " EUR. Relancer l'import avec ce meme OFX : " & _
                "les FITID deja presents sont ignores, seules les manquantes sont ajoutees.", "MANQUANTE", "", ""
        End If
    Next k

    nbLignesRapport = m_lignes.Count
    EcrireRapport ws, cheminOFX, nbComparees, nbHorsOFX, nbManquantes
    nbPreOui = CompterAAppliquer(ws)

    ws.Activate
    mod_InstallCommun.MasquerQuadrillage

    If nbLignesRapport = 0 Then
        MsgBox nbComparees & " operation(s) comparee(s) avec l'OFX : aucun ecart." & vbCrLf & _
               nbHorsOFX & " operation(s) du classeur hors periode de l'OFX (non comparees).", _
               vbInformation, "Remediation OFX"
        Exit Sub
    End If

    MsgBox "Analyse terminee (aucune donnee modifiee)." & vbCrLf & vbCrLf & _
           "Operations comparees : " & nbComparees & vbCrLf & _
           "Lignes d'ecart : " & nbLignesRapport & vbCrLf & _
           "  dont pre-cochees ""Oui"" : " & nbPreOui & vbCrLf & _
           "Operations OFX absentes du classeur : " & nbManquantes & vbCrLf & _
           "Operations du classeur hors periode de l'OFX : " & nbHorsOFX & vbCrLf & vbCrLf & _
           "Relisez le tableau, ajustez la colonne Corriger (Oui / Non), puis cliquez sur le bouton Appliquer.", _
           vbInformation, "Remediation OFX"

End Sub


' ============================================================================
' MACRO 2 : application des corrections cochées "Oui"
' ============================================================================
Public Sub AppliquerRemediationOFX()

    Dim ws As Worksheet
    Dim derniere As Long, i As Long, r As Long
    Dim id As String, code As String, avant As String, propose As String, courant As String
    Dim nb As Long, nbOk As Long, nbIgnore As Long
    Dim evtAvant As Boolean, majAvant As Boolean
    Dim touche As Boolean, sante As Boolean
    Dim chemin As String

    Set ws = TrouverFeuilleRapport()
    If ws Is Nothing Then
        MsgBox "Aucune analyse : cliquez d'abord sur Analyser un OFX.", vbExclamation, "Remediation OFX"
        Exit Sub
    End If
    Set m_lo = mod_DonneesTable.GetOperationsTable()
    If m_lo Is Nothing Then Exit Sub
    If m_lo.DataBodyRange Is Nothing Then Exit Sub
    If Not LireIndexColonnes() Then Exit Sub

    derniere = ws.Cells(ws.Rows.Count, C_ID).End(xlUp).Row
    nb = CompterAAppliquer(ws)
    If nb = 0 Then
        MsgBox "Aucune ligne a appliquer (colonne Corriger = Oui et Statut vide).", vbInformation, "Remediation OFX"
        Exit Sub
    End If
    If MsgBox(nb & " correction(s) vont etre ecrites dans TblOperations." & vbCrLf & _
              "Les valeurs d'avant restent dans le tableau (le bouton Annuler les corrections les restaure)." & vbCrLf & vbCrLf & _
              "Une copie du classeur avant ce traitement est recommandee. Continuer ?", _
              vbYesNo + vbDefaultButton2 + vbExclamation, "Remediation OFX") <> vbYes Then Exit Sub

    ConstruireIndexLignes

    evtAvant = Application.EnableEvents
    majAvant = Application.ScreenUpdating
    On Error GoTo Erreur
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    For i = LIGNE_ENTETES + 1 To derniere
        If CStr(ws.Cells(i, C_CORRIGER).Value) = "Oui" And CStr(ws.Cells(i, C_STATUT).Value) = "" Then
            code = CStr(ws.Cells(i, C_CODE).Value)
            If code <> "" And code <> "MANQUANTE" Then
                id = CStr(ws.Cells(i, C_ID).Value)
                avant = CStr(ws.Cells(i, C_AVANT).Value)
                propose = CStr(ws.Cells(i, C_PROPRAW).Value)
                If Not m_dictRow.Exists(id) Then
                    ws.Cells(i, C_STATUT).Value = "Ignoré : opération introuvable"
                    nbIgnore = nbIgnore + 1
                Else
                    r = m_dictRow(id)
                    courant = LireBrut(code, r)
                    If courant <> avant Then
                        ws.Cells(i, C_STATUT).Value = "Ignoré : valeur modifiée depuis l'analyse"
                        nbIgnore = nbIgnore + 1
                    ElseIf code = "BUDGET" Then
                        If RecalculerBudget(r) Then
                            ws.Cells(i, C_STATUT).Value = "Appliqué"
                            nbOk = nbOk + 1
                        Else
                            ws.Cells(i, C_STATUT).Value = "Ignoré : décalage manuel ou données incomplètes"
                            nbIgnore = nbIgnore + 1
                        End If
                    Else
                        EcrireBrut code, r, propose
                        ws.Cells(i, C_STATUT).Value = "Appliqué"
                        nbOk = nbOk + 1
                        If code = "NOTES" Or code = "TIERS" Or code = "CATSOUS" Then touche = True
                    End If
                End If
            End If
        End If
    Next i

    Application.EnableEvents = evtAvant
    Application.ScreenUpdating = majAvant
    On Error GoTo 0

    If touche Then
        On Error Resume Next
        mod_SuiviSante.CalculerSuiviSante AfficherResume:=False
        sante = (Err.Number = 0)
        On Error GoTo 0
    End If

    MsgBox nbOk & " correction(s) appliquee(s)." & vbCrLf & nbIgnore & " ignoree(s) (voir colonne Statut)." & vbCrLf & vbCrLf & _
           "Pour revenir en arriere : bouton Annuler les corrections.", vbInformation, "Remediation OFX"

    ' Le fichier OFX contient des donnees bancaires en clair : on propose de l'effacer.
    MajResume ws
    chemin = CStr(ws.Cells(1, C_CHEMIN).Value)
    If chemin <> "" Then
        If Len(Dir$(chemin)) > 0 Then
            If MsgBox("Effacer le fichier OFX utilise (ecrase puis supprime) ?" & vbCrLf & chemin, _
                      vbYesNo + vbQuestion, "Remediation OFX") = vbYes Then
                Dim msg As String
                msg = mod_ImportOFX.SupprimerFichierSensible(chemin)
                If msg <> "" Then MsgBox Trim$(msg), vbExclamation, "Remediation OFX"
            End If
        End If
    End If
    Exit Sub

Erreur:
    Application.EnableEvents = evtAvant
    Application.ScreenUpdating = majAvant
    MsgBox "Erreur pendant l'application : " & Err.Number & " - " & Err.Description & vbCrLf & _
           "Les lignes deja marquees ""Applique"" restent annulables (bouton Annuler les corrections).", vbCritical, "Remediation OFX"
End Sub


' ============================================================================
' MACRO 3 : annulation des corrections appliquées
' ============================================================================
Public Sub AnnulerRemediationOFX()

    Dim ws As Worksheet
    Dim derniere As Long, i As Long, r As Long
    Dim id As String, code As String, avant As String, propose As String, courant As String
    Dim nb As Long, nbOk As Long, nbIgnore As Long
    Dim evtAvant As Boolean, majAvant As Boolean

    Set ws = TrouverFeuilleRapport()
    If ws Is Nothing Then
        MsgBox "Le formulaire de remediation n'est pas installe (InstallerRemediationOFX).", vbExclamation, "Remediation OFX"
        Exit Sub
    End If
    Set m_lo = mod_DonneesTable.GetOperationsTable()
    If m_lo Is Nothing Then Exit Sub
    If m_lo.DataBodyRange Is Nothing Then Exit Sub
    If Not LireIndexColonnes() Then Exit Sub

    derniere = ws.Cells(ws.Rows.Count, C_ID).End(xlUp).Row
    For i = LIGNE_ENTETES + 1 To derniere
        If CStr(ws.Cells(i, C_STATUT).Value) = "Appliqué" Then nb = nb + 1
    Next i
    If nb = 0 Then
        MsgBox "Aucune correction appliquee a annuler.", vbInformation, "Remediation OFX"
        Exit Sub
    End If
    If MsgBox(nb & " correction(s) vont etre annulees (valeurs d'avant restaurees). Continuer ?", _
              vbYesNo + vbDefaultButton2 + vbQuestion, "Remediation OFX") <> vbYes Then Exit Sub

    ConstruireIndexLignes

    evtAvant = Application.EnableEvents
    majAvant = Application.ScreenUpdating
    On Error GoTo Erreur
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' Ordre inverse : le Budget (recalculé en dernier à l'application) est restauré en premier.
    For i = derniere To LIGNE_ENTETES + 1 Step -1
        If CStr(ws.Cells(i, C_STATUT).Value) = "Appliqué" Then
            code = CStr(ws.Cells(i, C_CODE).Value)
            id = CStr(ws.Cells(i, C_ID).Value)
            avant = CStr(ws.Cells(i, C_AVANT).Value)
            propose = CStr(ws.Cells(i, C_PROPRAW).Value)
            If Not m_dictRow.Exists(id) Then
                ws.Cells(i, C_STATUT).Value = "Annulation impossible : opération introuvable"
                nbIgnore = nbIgnore + 1
            Else
                r = m_dictRow(id)
                If code = "BUDGET" Then
                    RestaurerBudget r, avant
                    ws.Cells(i, C_STATUT).Value = "Annulé"
                    nbOk = nbOk + 1
                Else
                    courant = LireBrut(code, r)
                    If courant <> propose Then
                        ws.Cells(i, C_STATUT).Value = "Annulation impossible : valeur modifiée depuis"
                        nbIgnore = nbIgnore + 1
                    Else
                        EcrireBrut code, r, avant
                        ws.Cells(i, C_STATUT).Value = "Annulé"
                        nbOk = nbOk + 1
                    End If
                End If
            End If
        End If
    Next i

    Application.EnableEvents = evtAvant
    Application.ScreenUpdating = majAvant
    On Error GoTo 0

    MajResume ws
    MsgBox nbOk & " correction(s) annulee(s), " & nbIgnore & " non annulee(s) (voir colonne Statut).", _
           vbInformation, "Remediation OFX"
    Exit Sub

Erreur:
    Application.EnableEvents = evtAvant
    Application.ScreenUpdating = majAvant
    MsgBox "Erreur pendant l'annulation : " & Err.Number & " - " & Err.Description, vbCritical, "Remediation OFX"
End Sub


' ============================================================================
' ANALYSE D'UNE LIGNE (aucune écriture)
' ============================================================================
Private Sub AnalyserLigne(ByVal r As Long, ByVal k As Long)

    Dim id As String, dateTxt As String, tiersAct As String, notesAct As String
    Dim catAct As String, sousAct As String
    Dim vD As Date, ok As Boolean
    Dim ecartDate As Boolean
    Dim tiersDiff As Boolean, notesIncoherent As Boolean, catDiff As Boolean
    Dim suspect As Boolean, isSante As Boolean
    Dim expNotes As String, expCat As String, expSous As String, morceaux() As String
    Dim montantAct As Double
    Dim stCheque As String, stType As String
    Dim donneur As String
    Dim propBudget As Date, budgetAct As Date, okB As Boolean
    Dim tiersFinal As String, catFinal As String, sousFinal As String

    id = m_ops(k).Fitid
    tiersAct = mod_DataStructure.CellText(m_data(r, m_cTiers))
    notesAct = mod_DataStructure.CellText(m_data(r, m_cNotes))
    catAct = mod_DataStructure.CellText(m_data(r, m_cCat))
    If m_cSous > 0 Then sousAct = mod_DataStructure.CellText(m_data(r, m_cSous))
    isSante = (sousAct = mod_VarGlobales.SOUS_CATEGORIE_SANTE)

    vD = VersDate(m_data(r, m_cDate), ok)
    If ok Then dateTxt = Format$(vD, "dd/mm/yyyy") Else dateTxt = mod_DataStructure.CellText(m_data(r, m_cDate))

    ' ---------- Champs d'identité : l'OFX fait foi ----------
    If (Not ok) Then
        AjouterEcart "Oui", id, dateTxt, tiersAct, "Date_Comptable", dateTxt, Format$(m_ops(k).DatePost, "dd/mm/yyyy"), _
            NAT_IDENTITE, "", "DATE", "", BrutDate(m_ops(k).DatePost)
        ecartDate = True
    ElseIf Int(CDbl(vD)) <> Int(CDbl(m_ops(k).DatePost)) Then
        AjouterEcart "Oui", id, dateTxt, tiersAct, "Date_Comptable", dateTxt, Format$(m_ops(k).DatePost, "dd/mm/yyyy"), _
            NAT_IDENTITE, "", "DATE", BrutDate(vD), BrutDate(m_ops(k).DatePost)
        ecartDate = True
    End If

    If m_cDateOp > 0 And Not IsEmpty(m_ops(k).DateUser) Then
        Dim vD2 As Date, ok2 As Boolean
        vD2 = VersDate(m_data(r, m_cDateOp), ok2)
        If (Not ok2) Then
            AjouterEcart "Oui", id, dateTxt, tiersAct, "Date_Operation", mod_DataStructure.CellText(m_data(r, m_cDateOp)), _
                Format$(m_ops(k).DateUser, "dd/mm/yyyy"), NAT_IDENTITE, "", "DATEOP", "", BrutDate(m_ops(k).DateUser)
        ElseIf Int(CDbl(vD2)) <> Int(CDbl(m_ops(k).DateUser)) Then
            AjouterEcart "Oui", id, dateTxt, tiersAct, "Date_Operation", Format$(vD2, "dd/mm/yyyy"), _
                Format$(m_ops(k).DateUser, "dd/mm/yyyy"), NAT_IDENTITE, "", "DATEOP", BrutDate(vD2), BrutDate(m_ops(k).DateUser)
        End If
    End If

    If m_cType > 0 And m_ops(k).TypeOp <> "" Then
        stType = mod_DataStructure.CellText(m_data(r, m_cType))
        If StrComp(stType, m_ops(k).TypeOp, vbTextCompare) <> 0 Then
            AjouterEcart "Oui", id, dateTxt, tiersAct, "Type_Operation", stType, m_ops(k).TypeOp, _
                NAT_IDENTITE, "", "TYPE", stType, m_ops(k).TypeOp
        End If
    End If

    montantAct = mod_DataStructure.ToDouble(m_data(r, m_cMontant))
    If Abs(montantAct - m_ops(k).Montant) > 0.005 Then
        AjouterEcart "Oui", id, dateTxt, tiersAct, "Montant", Format$(montantAct, "0.00"), Format$(m_ops(k).Montant, "0.00"), _
            NAT_IDENTITE, "Une ligne ventilee garde le montant de l'operation parente : verifier avant d'appliquer.", _
            "MONTANT", BrutNombre(montantAct), BrutNombre(m_ops(k).Montant)
    End If

    If m_cCheque > 0 Then
        stCheque = mod_DataStructure.CellText(m_data(r, m_cCheque))
        If Not ChequeEgal(stCheque, m_ops(k).Cheque) Then
            AjouterEcart "Oui", id, dateTxt, tiersAct, "Num_Cheque", stCheque, m_ops(k).Cheque, _
                NAT_IDENTITE, "", "CHEQUE", stCheque, m_ops(k).Cheque
        End If
    End If

    ' ---------- Tiers / Notes / Catégorie : indices de mélange ----------
    tiersDiff = (StrComp(tiersAct, m_ops(k).Tiers, vbTextCompare) <> 0)

    If Not isSante Then
        If Not NotesCompatibles(notesAct, m_ops(k).Memo, m_ops(k).Notes) Then
            notesIncoherent = LooksBancaire(notesAct)
        End If
    End If

    If m_ops(k).Reconnu And Not m_map Is Nothing Then
        If m_map.Exists(m_ops(k).Categorie) Then
            morceaux = Split(CStr(m_map(m_ops(k).Categorie)), vbTab)
            expCat = morceaux(0)
            If UBound(morceaux) >= 1 Then expSous = morceaux(1)
            catDiff = (StrComp(catAct, expCat, vbTextCompare) <> 0) Or (StrComp(sousAct, expSous, vbTextCompare) <> 0)
        End If
    End If

    suspect = (tiersDiff And notesIncoherent) Or (tiersDiff And isSante And catDiff)

    If tiersDiff Or notesIncoherent Then donneur = TrouverDonneur(id, tiersAct, notesAct)

    If tiersDiff Then
        AjouterEcart IIf(suspect, "Oui", "Non"), id, dateTxt, tiersAct, "Tiers", tiersAct, m_ops(k).Tiers, NAT_TIERS, _
            IIf(suspect, "Ligne melangee. ", "Peut-etre une correction voulue de l'operateur. ") & donneur, _
            "TIERS", tiersAct, m_ops(k).Tiers
    End If

    If notesIncoherent Then
        If m_ops(k).Reconnu Then expNotes = m_ops(k).Notes Else expNotes = m_ops(k).Memo
        AjouterEcart "Oui", id, dateTxt, tiersAct, "Notes", notesAct, expNotes, NAT_NOTES, _
            "Le texte bancaire des Notes ne correspond pas a cette operation. Une note saisie par l'operateur sur l'AUTRE " & _
            "operation est perdue ici : a ressaisir sur celle-ci. " & donneur, "NOTES", notesAct, expNotes
    End If

    If catDiff And (tiersDiff Or notesIncoherent) Then
        AjouterEcart IIf(suspect, "Oui", "Non"), id, dateTxt, tiersAct, "Categorie / Sous-categorie", _
            catAct & " / " & sousAct, expCat & " / " & expSous, NAT_CATEGORIE, _
            "Correspondance de la categorie source '" & m_ops(k).Categorie & "' (TblCategories).", _
            "CATSOUS", catAct & vbTab & sousAct, expCat & vbTab & expSous
    End If

    ' ---------- Budget : recalcul proposé si la date, le Tiers ou la catégorie sont corrigés ----------
    If ecartDate Or suspect Then
        tiersFinal = IIf(suspect, m_ops(k).Tiers, tiersAct)
        If suspect And catDiff Then
            catFinal = expCat: sousFinal = expSous
        Else
            catFinal = catAct: sousFinal = sousAct
        End If
        If m_cBudget > 0 And m_cMois > 0 And m_cAnnee > 0 Then
            If m_cDecal = 0 Or mod_DataStructure.CellText(m_data(r, IIf(m_cDecal > 0, m_cDecal, 1))) = "" Then
                Dim dBase As Date
                dBase = m_ops(k).DatePost
                propBudget = mod_ImportOFX.CalculerBudget(dBase, tiersFinal, catFinal, sousFinal)
                budgetAct = VersDate(m_data(r, m_cBudget), okB)
                If (Not okB) Then
                    AjouterEcart IIf(suspect Or ecartDate, "Oui", "Non"), id, dateTxt, tiersAct, "Budget", "(vide)", Format$(propBudget, "mm/yyyy"), _
                        NAT_BUDGET, "Recalcule a partir de la date OFX et des valeurs corrigees.", "BUDGET", "", BrutDate(propBudget)
                ElseIf Month(budgetAct) <> Month(propBudget) Or Year(budgetAct) <> Year(propBudget) Then
                    AjouterEcart IIf(suspect Or ecartDate, "Oui", "Non"), id, dateTxt, tiersAct, "Budget", Format$(budgetAct, "mm/yyyy"), _
                        Format$(propBudget, "mm/yyyy"), NAT_BUDGET, _
                        "Recalcule a l'application a partir des valeurs finales de la ligne (ignore si decalage manuel).", _
                        "BUDGET", BrutDate(budgetAct), BrutDate(propBudget)
                End If
            End If
        End If
    End If

    If suspect And m_cPointe > 0 Then
        If StrComp(mod_DataStructure.CellText(m_data(r, m_cPointe)), mod_VarGlobales.POINTE_OUI, vbTextCompare) = 0 Then
            AjouterEcart "Non", id, dateTxt, tiersAct, "Pointe", mod_VarGlobales.POINTE_OUI, "(vide)", NAT_POINTE, _
                "Ligne pointee alors que ses donnees etaient melangees : a remettre a analyser si besoin.", _
                "POINTE", mod_VarGlobales.POINTE_OUI, ""
        End If
    End If

End Sub

' Cherche dans l'OFX l'opération dont cette ligne a repris le Tiers ET les Notes.
Private Function TrouverDonneur(ByVal id As String, ByVal tiersAct As String, ByVal notesAct As String) As String
    Dim j As Long, rr As Long
    Dim texte As String
    For j = 1 To m_nbOps
        If m_ops(j).Fitid <> id Then
            If StrComp(m_ops(j).Tiers, tiersAct, vbTextCompare) = 0 Then
                If m_ops(j).Notes <> "" Or m_ops(j).Memo <> "" Then
                    If NotesCompatibles(notesAct, m_ops(j).Memo, m_ops(j).Notes) And notesAct <> "" Then
                        texte = "Tiers/Notes identiques a l'operation FITID " & m_ops(j).Fitid & " (" & _
                                Format$(m_ops(j).DatePost, "dd/mm/yyyy") & ", " & Format$(m_ops(j).Montant, "0.00") & " EUR)."
                        If m_dictRow.Exists(m_ops(j).Fitid) Then
                            rr = m_dictRow(m_ops(j).Fitid)
                            texte = texte & " Sa ligne dans le classeur porte actuellement le Tiers '" & _
                                    mod_DataStructure.CellText(m_data(rr, m_cTiers)) & "'."
                        Else
                            texte = texte & " Cette operation est absente du classeur."
                        End If
                        TrouverDonneur = texte
                        Exit Function
                    End If
                End If
            End If
        End If
    Next j
End Function

' Les Notes enregistrées sont-elles COMPATIBLES avec l'opération ? Vrai si vides (suppression
' volontaire), égales aux notes attendues ou au MEMO brut, ou si elles les contiennent (note de
' l'opérateur ajoutée devant ou derrière).
Private Function NotesCompatibles(ByVal notes As String, ByVal memo As String, ByVal notesAttendues As String) As Boolean
    notes = Trim$(notes)
    If notes = "" Then NotesCompatibles = True: Exit Function
    If notesAttendues <> "" Then
        If InStr(1, notes, notesAttendues, vbTextCompare) > 0 Then NotesCompatibles = True: Exit Function
    End If
    If memo <> "" Then
        If InStr(1, notes, memo, vbTextCompare) > 0 Then NotesCompatibles = True: Exit Function
        If InStr(1, memo, notes, vbTextCompare) > 0 Then NotesCompatibles = True: Exit Function
    End If
    If notesAttendues = "" And memo <> "" Then
        ' MEMO = catégorie seule : toute note saisie à la main reste compatible.
        If ExtraireSeulCategorie(memo) Then NotesCompatibles = True
    End If
End Function

Private Function ExtraireSeulCategorie(ByVal memo As String) As Boolean
    Dim c As String, n As String
    ExtraireSeulCategorie = mod_ImportOFX.ExtraireCategorieMemo(memo, m_sources, m_nbSources, c, n) And (n = "")
End Function

' Le texte ressemble-t-il à un libellé bancaire généré (et non à une note saisie à la main) ?
Private Function LooksBancaire(ByVal texte As String) As Boolean
    LooksBancaire = (InStr(1, texte, "CB ", vbBinaryCompare) > 0) Or (InStr(1, texte, " FACT ", vbBinaryCompare) > 0) _
                    Or (InStr(1, texte, "PRLV ", vbBinaryCompare) > 0) Or (InStr(1, texte, "VIR SEPA", vbBinaryCompare) > 0) _
                    Or (InStr(1, texte, "RETRAIT DAB", vbBinaryCompare) > 0)
End Function

Private Function ChequeEgal(ByVal a As String, ByVal b As String) As Boolean
    a = Trim$(a): b = Trim$(b)
    If a = "" And b = "" Then ChequeEgal = True: Exit Function
    If IsNumeric(a) And IsNumeric(b) Then
        ChequeEgal = (Val(a) = Val(b))
    Else
        ChequeEgal = (StrComp(a, b, vbTextCompare) = 0)
    End If
End Function


' ============================================================================
' CHARGEMENT DE L'OFX
' ============================================================================
Private Function ChargerOFX(ByVal chemin As String) As Boolean

    Dim xmlDoc As Object, liste As Object, noeud As Object
    Dim i As Long, n As Long
    Dim fitid As String, memo As String
    Dim cat As String, notes As String

    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.validateOnParse = False
    If Not xmlDoc.Load(chemin) Then
        MsgBox "Le fichier OFX n'a pas pu etre lu." & vbCrLf & xmlDoc.parseError.reason, vbCritical, "Remediation OFX"
        Exit Function
    End If
    Set liste = xmlDoc.getElementsByTagName("STMTTRN")
    If liste.Length = 0 Then
        MsgBox "Aucune transaction trouvee dans le fichier OFX.", vbExclamation, "Remediation OFX"
        Exit Function
    End If

    ReDim m_ops(1 To liste.Length)
    Set m_dictOFX = CreateObject("Scripting.Dictionary")
    m_nbOps = 0

    For i = 0 To liste.Length - 1
        Set noeud = liste.Item(i)
        fitid = mod_ImportOFX.LireValeurNoeud(noeud, "FITID")
        If fitid <> "" Then
            If Not m_dictOFX.Exists(fitid) Then
                m_nbOps = m_nbOps + 1
                n = m_nbOps
                m_ops(n).Fitid = fitid
                m_ops(n).DatePost = mod_ImportOFX.ConvertirDateOFX(mod_ImportOFX.LireValeurNoeud(noeud, "DTPOSTED"))
                m_ops(n).DateUser = mod_ImportOFX.ConvertirDateOFXVariant(mod_ImportOFX.LireValeurNoeud(noeud, "DTUSER"))
                m_ops(n).TypeOp = mod_ImportOFX.LireValeurNoeud(noeud, "TRNTYPE")
                m_ops(n).Montant = Val(mod_ImportOFX.LireValeurNoeud(noeud, "TRNAMT"))
                m_ops(n).Tiers = mod_ImportOFX.LireValeurNoeud(noeud, "NAME")
                memo = mod_ImportOFX.LireValeurNoeud(noeud, "MEMO")
                m_ops(n).Memo = memo
                m_ops(n).Cheque = CStr(mod_ImportOFX.NettoyerNumCheque(mod_ImportOFX.LireValeurNoeud(noeud, "CHECKNUM")))
                If mod_ImportOFX.ExtraireCategorieMemo(memo, m_sources, m_nbSources, cat, notes) Then
                    m_ops(n).Reconnu = True
                    m_ops(n).Categorie = cat
                    m_ops(n).Notes = notes
                Else
                    m_ops(n).Reconnu = False
                    m_ops(n).Categorie = ""
                    m_ops(n).Notes = memo       ' MEMO non reconnu : repris tel quel, comme à l'import
                End If
                m_dictOFX.Add fitid, n
            End If
        End If
    Next i

    ChargerOFX = (m_nbOps > 0)
End Function

' Catégories source connues + correspondance CategorieSource -> Catégorie/Sous-catégorie.
Private Sub LireSourcesEtCorrespondance()
    Dim ws As Worksheet, lo As ListObject
    Dim src As Variant, cat As Variant, sous As Variant
    Dim r As Long, s As String, c As String, u As String

    m_sources = mod_Categories.ObtenirCategoriesSource()
    On Error Resume Next
    m_nbSources = UBound(m_sources) - LBound(m_sources) + 1
    If Err.Number <> 0 Then m_nbSources = 0
    Err.Clear
    On Error GoTo 0

    Set m_map = CreateObject("Scripting.Dictionary")
    m_map.CompareMode = 1

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(mod_Categories.NOM_FEUILLE_PARAM)
    Set lo = ws.ListObjects(mod_Categories.NOM_TABLE_CATEGORIES)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    src = ColonneEn2D(lo.ListColumns("CategorieSource").DataBodyRange)
    cat = ColonneEn2D(lo.ListColumns("Categorie").DataBodyRange)
    sous = ColonneEn2D(lo.ListColumns("SousCategorie").DataBodyRange)

    For r = 1 To UBound(cat, 1)
        s = Trim$(mod_DataStructure.CellText(src(r, 1)))
        c = mod_DataStructure.CellText(cat(r, 1))
        u = mod_DataStructure.CellText(sous(r, 1))
        If s <> "" And c <> "" Then
            If Not m_map.Exists(s) Then m_map.Add s, c & vbTab & u    ' la 1re ligne l'emporte (comme ControleCategories)
        End If
    Next r
End Sub

Private Function ColonneEn2D(ByVal plage As Range) As Variant
    Dim t() As Variant
    If plage.Cells.Count = 1 Then
        ReDim t(1 To 1, 1 To 1)
        t(1, 1) = plage.Value2
        ColonneEn2D = t
    Else
        ColonneEn2D = plage.Value2
    End If
End Function


' ============================================================================
' LECTURE / ÉCRITURE DES CELLULES (toujours via l'ID, jamais via un numéro de ligne mémorisé)
' ============================================================================
Private Function LireIndexColonnes() As Boolean
    m_cID = IdxCol("ID_Transaction")
    m_cDate = IdxCol("Date_Comptable")
    m_cDateOp = IdxCol("Date_Operation")
    m_cType = IdxCol("Type_Operation")
    m_cTiers = IdxCol("Tiers")
    m_cNotes = IdxCol("Notes")
    m_cMontant = IdxCol("Montant")
    m_cCheque = IdxCol("Num_Cheque")
    m_cCat = IdxCol("Categorie")
    m_cSous = IdxCol("SousCategorie")
    m_cBudget = IdxCol("Budget")
    m_cMois = IdxCol("MoisBudget")
    m_cAnnee = IdxCol("AnneeBudget")
    m_cDecal = IdxCol(mod_VarGlobales.NOM_COL_DECALAGE_MANUEL)
    m_cPointe = IdxCol(mod_VarGlobales.NOM_COL_POINTE)

    If m_cID = 0 Or m_cDate = 0 Or m_cTiers = 0 Or m_cNotes = 0 Or m_cMontant = 0 Or m_cCat = 0 Then
        MsgBox "Colonnes indispensables introuvables dans TblOperations (ID_Transaction, Date_Comptable, Tiers, Notes, Montant, Categorie).", _
               vbCritical, "Remediation OFX"
        Exit Function
    End If
    LireIndexColonnes = True
End Function

Private Function IdxCol(ByVal nom As String) As Long
    On Error Resume Next
    IdxCol = m_lo.ListColumns(nom).Index
    On Error GoTo 0
End Function

Private Sub ConstruireIndexLignes()
    Dim vId As Variant, r As Long, id As String
    Set m_dictRow = CreateObject("Scripting.Dictionary")
    vId = ColonneEn2D(m_lo.ListColumns(m_cID).DataBodyRange)
    For r = 1 To UBound(vId, 1)
        id = mod_DataStructure.CellText(vId(r, 1))
        If id <> "" Then
            If Not m_dictRow.Exists(id) Then m_dictRow.Add id, r
        End If
    Next r
End Sub

Private Function Cellule(ByVal r As Long, ByVal c As Long) As Range
    Set Cellule = m_lo.DataBodyRange.Cells(r, c)
End Function

' Valeur actuelle d'une cellule sous la forme "brute" mémorisée dans le rapport.
Private Function LireBrut(ByVal code As String, ByVal r As Long) As String
    Dim ok As Boolean, d As Date
    Select Case code
        Case "DATE"
            d = VersDate(Cellule(r, m_cDate).Value, ok)
            If ok Then LireBrut = BrutDate(d)
        Case "DATEOP"
            d = VersDate(Cellule(r, m_cDateOp).Value, ok)
            If ok Then LireBrut = BrutDate(d)
        Case "TYPE": LireBrut = mod_DataStructure.CellText(Cellule(r, m_cType).Value)
        Case "MONTANT": LireBrut = BrutNombre(mod_DataStructure.ToDouble(Cellule(r, m_cMontant).Value))
        Case "CHEQUE": LireBrut = mod_DataStructure.CellText(Cellule(r, m_cCheque).Value)
        Case "TIERS": LireBrut = mod_DataStructure.CellText(Cellule(r, m_cTiers).Value)
        Case "NOTES": LireBrut = mod_DataStructure.CellText(Cellule(r, m_cNotes).Value)
        Case "CATSOUS"
            LireBrut = mod_DataStructure.CellText(Cellule(r, m_cCat).Value) & vbTab
            If m_cSous > 0 Then LireBrut = LireBrut & mod_DataStructure.CellText(Cellule(r, m_cSous).Value)
        Case "BUDGET"
            d = VersDate(Cellule(r, m_cBudget).Value, ok)
            If ok Then LireBrut = BrutDate(d)
        Case "POINTE": LireBrut = mod_DataStructure.CellText(Cellule(r, m_cPointe).Value)
    End Select
End Function

Private Sub EcrireBrut(ByVal code As String, ByVal r As Long, ByVal brut As String)
    Dim morceaux() As String
    Select Case code
        Case "DATE"
            Cellule(r, m_cDate).Value = DepuisBrutDate(brut)
            Cellule(r, m_cDate).NumberFormat = "dd/mm/yyyy"
        Case "DATEOP"
            If brut = "" Then
                Cellule(r, m_cDateOp).ClearContents
            Else
                Cellule(r, m_cDateOp).Value = DepuisBrutDate(brut)
                Cellule(r, m_cDateOp).NumberFormat = "dd/mm/yyyy"
            End If
        Case "TYPE": Cellule(r, m_cType).Value = brut
        Case "MONTANT": Cellule(r, m_cMontant).Value = Val(brut)
        Case "CHEQUE"
            If brut = "" Then Cellule(r, m_cCheque).ClearContents Else Cellule(r, m_cCheque).Value = brut
        Case "TIERS": EcrireTexte Cellule(r, m_cTiers), brut
        Case "NOTES": EcrireTexte Cellule(r, m_cNotes), brut
        Case "CATSOUS"
            morceaux = Split(brut, vbTab)
            EcrireTexte Cellule(r, m_cCat), morceaux(0)
            If m_cSous > 0 And UBound(morceaux) >= 1 Then EcrireTexte Cellule(r, m_cSous), morceaux(1)
        Case "POINTE"
            If brut = "" Then Cellule(r, m_cPointe).ClearContents Else Cellule(r, m_cPointe).Value = brut
    End Select
End Sub

' Écrit un texte sans qu'Excel l'interprète comme une formule ("=...", "+...", "-...").
Private Sub EcrireTexte(ByVal cel As Range, ByVal texte As String)
    If texte = "" Then
        cel.ClearContents
    ElseIf Left$(texte, 1) = "=" Or Left$(texte, 1) = "+" Or Left$(texte, 1) = "-" Or Left$(texte, 1) = "@" Then
        cel.Value = "'" & texte
    Else
        cel.Value = texte
    End If
End Sub

' Recalcule Budget / MoisBudget / AnneeBudget à partir des valeurs FINALES de la ligne.
' Renvoie False (sans rien écrire) s'il y a un décalage manuel ou des données incomplètes.
Private Function RecalculerBudget(ByVal r As Long) As Boolean
    Dim ok As Boolean, d As Date, b As Date
    If m_cBudget = 0 Or m_cMois = 0 Or m_cAnnee = 0 Then Exit Function
    If m_cDecal > 0 Then
        If mod_DataStructure.CellText(Cellule(r, m_cDecal).Value) <> "" Then Exit Function
    End If
    d = VersDate(Cellule(r, m_cDate).Value, ok)
    If Not ok Then Exit Function
    If m_cSous > 0 Then
        b = mod_ImportOFX.CalculerBudget(d, mod_DataStructure.CellText(Cellule(r, m_cTiers).Value), _
                mod_DataStructure.CellText(Cellule(r, m_cCat).Value), mod_DataStructure.CellText(Cellule(r, m_cSous).Value))
    Else
        b = mod_ImportOFX.CalculerBudget(d, mod_DataStructure.CellText(Cellule(r, m_cTiers).Value), _
                mod_DataStructure.CellText(Cellule(r, m_cCat).Value), "")
    End If
    EcrireBudget r, b
    RecalculerBudget = True
End Function

Private Sub EcrireBudget(ByVal r As Long, ByVal b As Date)
    Cellule(r, m_cBudget).Value = b
    Cellule(r, m_cBudget).NumberFormat = "mm/yyyy"
    Cellule(r, m_cMois).Value = Month(b)
    Cellule(r, m_cAnnee).Value = Year(b)
End Sub

Private Sub RestaurerBudget(ByVal r As Long, ByVal brut As String)
    If brut = "" Then
        Cellule(r, m_cBudget).ClearContents
        Cellule(r, m_cMois).ClearContents
        Cellule(r, m_cAnnee).ClearContents
    Else
        EcrireBudget r, DepuisBrutDate(brut)
    End If
End Sub


' ============================================================================
' CONVERSIONS NEUTRES (indépendantes des paramètres régionaux)
' ============================================================================
Private Function VersDate(ByVal v As Variant, ByRef ok As Boolean) As Date
    ok = False
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then
        VersDate = v: ok = True
    ElseIf IsNumeric(v) Then
        If CDbl(v) > 0 Then VersDate = CDate(CDbl(v)): ok = True
    ElseIf IsDate(v) Then
        VersDate = CDate(v): ok = True
    End If
End Function

Private Function BrutDate(ByVal d As Date) As String
    BrutDate = Format$(d, "yyyy-mm-dd")
End Function

Private Function DepuisBrutDate(ByVal brut As String) As Date
    DepuisBrutDate = DateSerial(CInt(Left$(brut, 4)), CInt(Mid$(brut, 6, 2)), CInt(Right$(brut, 2)))
End Function

Private Function BrutNombre(ByVal x As Double) As String
    BrutNombre = Replace(CStr(Round(x, 2)), ",", ".")
End Function


' ============================================================================
' FEUILLE DE RAPPORT
' ============================================================================
Private Sub AjouterEcart(ByVal corriger As String, ByVal id As String, ByVal dateTxt As String, ByVal tiers As String, _
                         ByVal champ As String, ByVal actuel As String, ByVal propose As String, ByVal nature As String, _
                         ByVal commentaire As String, ByVal code As String, ByVal avantBrut As String, ByVal propBrut As String)
    m_lignes.Add Array(corriger, "", id, dateTxt, tiers, champ, actuel, propose, nature, commentaire, code, avantBrut, propBrut)
End Sub

Private Function TrouverFeuilleRapport() As Worksheet
    On Error Resume Next
    Set TrouverFeuilleRapport = ThisWorkbook.Worksheets(NOM_FEUILLE_REMEDIATION)
    On Error GoTo 0
End Function

' Remplit le tableau du formulaire frm_Remediation_OFX avec les écarts trouvés (m_lignes).
' Le formulaire (feuille très masquée) est construit par mod_InstallRemediationOFX : ici on
' ne touche qu'à la zone sous la ligne d'en-têtes, qui est effacée puis réécrite.
Private Sub EcrireRapport(ByVal ws As Worksheet, ByVal cheminOFX As String, ByVal nbComparees As Long, _
                          ByVal nbHorsOFX As Long, ByVal nbManquantes As Long)

    Dim n As Long, i As Long, j As Long
    Dim tableau() As Variant
    Dim ligne As Variant
    Dim derniereAvant As Long, premiere As Long, derniere As Long
    Dim plage As Range
    Dim evtAvant As Boolean, majAvant As Boolean

    evtAvant = Application.EnableEvents
    majAvant = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ws.Visible = xlSheetVisible
    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    ' --- Effacement de l'analyse précédente ---
    premiere = LIGNE_ENTETES + 1
    derniereAvant = ws.Cells(ws.Rows.Count, C_ID).End(xlUp).Row
    If derniereAvant < premiere Then derniereAvant = premiere
    With ws.Range(ws.Cells(premiere, 1), ws.Cells(derniereAvant + 1, C_CHEMIN))
        .Clear
        .Validation.Delete
    End With
    ws.Range(ws.Rows(premiere), ws.Rows(derniereAvant + 1)).RowHeight = mod_InstallCommun.FRM_H_LIGNE

    ws.Cells(1, C_CHEMIN).Value = cheminOFX

    n = m_lignes.Count
    If n > 0 Then
        ReDim tableau(1 To n, 1 To NB_COL_RAPPORT)
        For i = 1 To n
            ligne = m_lignes(i)
            For j = 1 To NB_COL_RAPPORT
                tableau(i, j) = ligne(j - 1)
            Next j
        Next i
        derniere = premiere + n - 1

        ' Texte partout (identifiants, dates, valeurs brutes) : Excel ne doit rien convertir.
        Set plage = ws.Range(ws.Cells(premiere, C_CORRIGER), ws.Cells(derniere, C_PROPRAW))
        plage.NumberFormat = "@"
        plage.Value2 = tableau

        ' Mise en forme commune : lecture seule grise, colonne Corriger modifiable (jaune).
        Set plage = ws.Range(ws.Cells(premiere, C_STATUT), ws.Cells(derniere, C_COMMENT))
        mod_InstallCommun.PoserCorpsTableauLecture plage
        With ws.Range(ws.Cells(premiere, C_CORRIGER), ws.Cells(derniere, C_CORRIGER))
            .Interior.Color = mod_InstallCommun.CoulFondSaisie()
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.Color = mod_InstallCommun.CoulBordureSaisie()
            .HorizontalAlignment = xlCenter
            .Locked = False
            With .Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Oui,Non,-"
                .IgnoreBlank = True
                .InCellDropdown = True
            End With
        End With
        With ws.Range(ws.Cells(premiere, C_CORRIGER), ws.Cells(derniere, C_COMMENT))
            .Font.Name = "Calibri"
            .Font.Size = 10
            .VerticalAlignment = xlTop
        End With
        ws.Range(ws.Cells(premiere, C_COMMENT), ws.Cells(derniere, C_COMMENT)).WrapText = True
        ws.Range(ws.Cells(premiere, C_ACTUEL), ws.Cells(derniere, C_PROPOSE)).WrapText = True
        ws.Range(ws.Rows(premiere), ws.Rows(derniere)).AutoFit

        ' Filtre sur les en-têtes (Corriger, Nature... pour trier la revue).
        ws.Range(ws.Cells(LIGNE_ENTETES, C_CORRIGER), ws.Cells(derniere, C_COMMENT)).AutoFilter
    End If

    ws.Range(ws.Cells(REM_LIGNE_RESUME, C_CORRIGER), ws.Cells(REM_LIGNE_RESUME, C_CORRIGER)).Value = _
        "Analyse du " & Format$(Now, "dd/mm/yyyy hh:nn") & " : " & nbComparees & " operations comparees, " & n & _
        " ligne(s) d'ecart, " & nbManquantes & " absente(s) du classeur, " & nbHorsOFX & " du classeur hors periode OFX."

    Application.EnableEvents = evtAvant
    Application.ScreenUpdating = majAvant
End Sub

' Résumé de la ligne 4 après une application / annulation : lignes "Oui" en attente, appliquées.
Private Sub MajResume(ByVal ws As Worksheet)
    Dim derniere As Long, i As Long, nbAttente As Long, nbAppl As Long, base As String
    derniere = ws.Cells(ws.Rows.Count, C_ID).End(xlUp).Row
    For i = LIGNE_ENTETES + 1 To derniere
        If CStr(ws.Cells(i, C_STATUT).Value) = "Appliqué" Then nbAppl = nbAppl + 1
    Next i
    nbAttente = CompterAAppliquer(ws)
    base = CStr(ws.Cells(REM_LIGNE_RESUME, C_CORRIGER).Value)
    If InStr(base, " | ") > 0 Then base = Left$(base, InStr(base, " | ") - 1)
    ws.Cells(REM_LIGNE_RESUME, C_CORRIGER).Value = base & " | A appliquer : " & nbAttente & " | Appliquees : " & nbAppl
End Sub


' ============================================================================
' OUVERTURE / SORTIE DU FORMULAIRE (bouton de Synthese, bouton Sortir)
' ============================================================================
Public Sub AfficherRemediationOFX()

    Dim ws As Worksheet

    Set ws = TrouverFeuilleRapport()
    If ws Is Nothing Then
        MsgBox "Le formulaire " & NOM_FEUILLE_REMEDIATION & " n'est pas installe : lancez InstallerRemediationOFX (Ctrl+G).", _
               vbExclamation, "Remediation OFX"
        Exit Sub
    End If

    ' Synthese est masquee par MasquerQuadrillage (regle commune des formulaires) ;
    ' SortirRemediationOFX la reaffiche.
    ws.Visible = xlSheetVisible
    ws.Activate
    mod_InstallCommun.MasquerQuadrillage

End Sub

Public Sub SortirRemediationOFX()

    Dim ws As Worksheet

    mod_InstallCommun.AfficherSynthese

    Set ws = TrouverFeuilleRapport()
    If Not ws Is Nothing Then mod_Aide.MasquerFormulaire ws

End Sub

' Nombre de lignes du rapport à appliquer : Corriger = "Oui", Statut vide, ligne technique valide.
Private Function CompterAAppliquer(ByVal ws As Worksheet) As Long
    Dim derniere As Long, i As Long, nb As Long, code As String
    derniere = ws.Cells(ws.Rows.Count, C_ID).End(xlUp).Row
    For i = LIGNE_ENTETES + 1 To derniere
        If CStr(ws.Cells(i, C_CORRIGER).Value) = "Oui" And CStr(ws.Cells(i, C_STATUT).Value) = "" Then
            code = CStr(ws.Cells(i, C_CODE).Value)
            If code <> "" And code <> "MANQUANTE" Then nb = nb + 1
        End If
    Next i
    CompterAAppliquer = nb
End Function
