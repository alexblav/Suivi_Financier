Option Explicit

' =====================================================================================
' MODULE : mod_Synchro  (ajout 11/10/2026, déploiement et sauvegardes)
'
' CONTEXTE
'   Pour ne pas être bloqué par les règles de sécurité de Windows (macros d'un fichier
'   placé sur un partage réseau), chaque opérateur travaille sur une COPIE LOCALE du
'   classeur (sur son bureau). Une copie "centrale" est conservée sur le réseau ; ce
'   module garde les deux copies strictement identiques du point de vue des DONNÉES.
'
' QUAND ?
'   - À L'OUVERTURE  (ThisWorkbook.Workbook_Open)        : SynchroniserALOuverture
'   - À LA FERMETURE (ThisWorkbook.Workbook_BeforeClose) : SynchroniserALaFermeture
'   - À la demande   : SynchroniserMaintenant, DefinirFichierCentral, AfficherEtatSynchro
'
' PRINCIPE (comparaison à 3 voies, par "empreinte" des données)
'   Une empreinte (hash) est calculée sur toutes les données synchronisées, pour le
'   fichier local (L) et pour le fichier central (N). Le classeur local mémorise aussi
'   l'empreinte de la DERNIÈRE synchronisation réussie (B, propriété personnalisée du
'   classeur "SyncBaseHash"). On en déduit :
'       L = N                 -> déjà identiques, rien à faire
'       L = B et N <> B       -> seul le réseau a changé : on RÉCUPÈRE (on copie les données
'                                du central dans le classeur local, avec un résumé)
'       N = B et L <> B       -> seul le local a changé : on ENVOIE (le central est remplacé
'                                par une copie complète du classeur local)
'       L <> B et N <> B      -> CONFLIT (les deux ont changé) : l'opérateur choisit la
'                                version à conserver ; l'autre est sauvegardée avant
'                                d'être écrasée.
'   Après chaque récupération ou envoi, les empreintes sont RECALCULÉES pour vérifier que
'   les deux fichiers sont bien identiques ; sinon un avertissement liste les écarts.
'
' CE QUI EST SYNCHRONISÉ (voir ListeTables / ListeFeuilles / ListeNoms)
'   - les tableaux de données (colonnes retrouvées par leur NOM, jamais par leur position ;
'     les colonnes calculées par formule sont ignorées : elles se recalculent seules) ;
'   - les cellules des feuilles de paramètres (Param hors tableaux, TechDernierImport) ;
'   - les cellules nommées de réglage (soldeRefMontant, soldeRefDate).
'   NE SONT PAS synchronisés : les écrans de résultat (TblRechercheOperations, analyse,
'   consommations) qui se recalculent, et les choix d'affichage propres à chaque
'   opérateur (critères de Synthese, réglages des écrans). Pour ajouter un tableau, une
'   feuille ou une cellule nommée au périmètre, il suffit de l'ajouter dans les 3 fonctions
'   Liste... ci-dessous.
'
' OÙ EST MÉMORISÉ LE FICHIER CENTRAL ?
'   Dans le registre Windows de l'opérateur (SaveSetting/GetSetting), PAS dans le classeur :
'   le chemin est propre à chaque poste et ne fait donc jamais partie des données
'   synchronisées. Conseil : saisir un chemin UNC (\\serveur\partage\...) plutôt qu'une
'   lettre de lecteur mappée, qui peut changer d'un poste à l'autre.
'
' SÉCURITÉS
'   - un fichier verrou (<central>.verrou) évite deux synchronisations simultanées ;
'   - avant d'écraser le central, une sauvegarde est faite dans "Sauvegardes_Synchro"
'     (à côté du fichier central ; les 10 dernières sont gardées) ; en cas de conflit résolu
'     en faveur du réseau, le classeur local est sauvegardé de la même façon ;
'   - le central est lu à partir d'une copie temporaire en lecture seule (macros désactivées).
'   - classeur chiffré (Fichier > Informations > Protéger le classeur > Chiffrer avec mot de passe,
'     MÊME passphrase pour tous et pour le central) : le central est détecté comme chiffré (signature
'     OLE), la passphrase est demandée UNE fois par session Excel (au premier accès au central, donc
'     à l'ouverture) puis réutilisée à la fermeture. Elle reste en mémoire (mPassphrase + variable d'environnement du processus Excel), jamais
'     écrite sur disque ni dans le classeur. SaveCopyAs conserve le chiffrement : central et
'     sauvegardes restent chiffrés. Passphrase refusée : 3 essais, puis la synchro échoue.
' =====================================================================================

Private Const TITRE As String = "Synchronisation"
Private Const APP_REG As String = "SuiviFinancier"
Private Const SECTION_REG As String = "Synchro"
Private Const CLE_CHEMIN As String = "CheminCentral"
Private Const PROP_BASE As String = "SyncBaseHash"
Private Const PROP_DATE As String = "SyncDerniereDate"
Private Const DOSSIER_SAUVEGARDES As String = "Sauvegardes_Synchro"
Private Const NB_SAUVEGARDES As Long = 10
Private Const VERROU_MINUTES As Long = 15
Private Const VAR_SESSION As String = "SUIVIFIN_PASSPHRASE"

' Enregistrer automatiquement le classeur local à la fermeture, une fois la synchro faite
' (sinon l'empreinte "dernière synchronisation" ne serait pas conservée dans le fichier).
Private Const ENREGISTRER_A_LA_FERMETURE As Boolean = True

Private mEnCours As Boolean
Private mPassphrase As String       ' passphrase du fichier central, en mémoire seulement (jamais sur disque)
Private mWbN As Workbook            ' copie temporaire du central, ouverte en lecture seule
Private mCopieTemp As String        ' chemin de cette copie temporaire
Private mVerrou As String           ' chemin du fichier verrou posé par cette synchronisation
Private mResume As String           ' lignes du résumé affiché à l'opérateur
Private mNbChg As Long              ' nombre total de lignes/cellules modifiées par la récupération
Private mH1 As Double, mH2 As Double

' Sauvegarde de l'état d'Excel pendant la synchronisation
Private mEcran As Boolean, mEvt As Boolean, mCalc As Long

Private Function T(ByVal s As String) As String
    T = mod_Display.FR(s)
End Function

' =====================================================================================
' PÉRIMÈTRE DE LA SYNCHRONISATION
' =====================================================================================
Private Function ListeTables() As Variant
    ListeTables = Array("TblOperations", "TblVentilations", _
                        mod_VarGlobales.NOM_TABLE_ECHEANCES, mod_VarGlobales.NOM_TABLE_RECURRENCES, _
                        mod_VarGlobales.NOM_TABLE_FACTURES, mod_VarGlobales.NOM_TABLE_ABOS, mod_VarGlobales.NOM_TABLE_CONSOS, _
                        mod_VarGlobales.NOM_TABLE_TAXES, mod_VarGlobales.NOM_TABLE_BLOCS, mod_VarGlobales.NOM_TABLE_CONTRATS, _
                        mod_Categories.NOM_TABLE_CATEGORIES, mod_VarGlobales.NOM_TABLE_DECALAGES_BUDGET)
End Function

Private Function ListeFeuilles() As Variant
    ListeFeuilles = Array(mod_Categories.NOM_FEUILLE_PARAM, mod_VarGlobales.NOM_FEUILLE_TECH)
End Function

Private Function ListeNoms() As Variant
    ListeNoms = Array("soldeRefMontant", "soldeRefDate")
End Function

' =====================================================================================
' POINTS D'ENTRÉE
' =====================================================================================
Public Sub SynchroniserALOuverture()
    ' À l'ouverture, un échec n'empêche jamais de travailler : on continue simplement.
    Synchroniser "ouverture"
End Sub

' Renvoie Faux si la fermeture doit être annulée (l'opérateur a refusé de fermer sans synchro).
Public Function SynchroniserALaFermeture() As Boolean
    Dim ok As Boolean
    ' Une liste de consommations encore ouverte est d'abord écrite dans les données.
    mod_Consommations.EnregistrerListesSiOuvertes
    ok = Synchroniser("fermeture")
    If Not ok Then
        ok = (MsgBox(T("La synchronisation avec le fichier central n'a pas abouti.") & vbCrLf & _
                     T("Fermer quand m{e1}me le classeur ?") & vbCrLf & vbCrLf & _
                     T("(Les donn{e2}es seront resynchronis{e2}es {a2} la prochaine ouverture.)"), _
                     vbYesNo + vbExclamation + vbDefaultButton2, TITRE) = vbYes)
    ElseIf ENREGISTRER_A_LA_FERMETURE Then
        If Not ThisWorkbook.ReadOnly Then ThisWorkbook.Save
    End If
    SynchroniserALaFermeture = ok
End Function

Public Sub SynchroniserMaintenant()
    Synchroniser "manuel"
End Sub

' Permet de désigner un autre fichier central (ou d'en créer un nouveau).
Public Sub DefinirFichierCentral()
    Dim p As String
    p = ChoisirCentral("manuel", GetSetting(APP_REG, SECTION_REG, CLE_CHEMIN, ""))
    If Len(p) > 0 Then
        MsgBox T("Fichier central m{e2}moris{e2} :") & vbCrLf & p & vbCrLf & vbCrLf & _
               T("Lancez SynchroniserMaintenant (ou fermez le classeur) pour synchroniser."), vbInformation, TITRE
    End If
End Sub

Public Sub AfficherEtatSynchro()
    Dim p As String, s As String
    p = GetSetting(APP_REG, SECTION_REG, CLE_CHEMIN, "")
    s = T("Fichier central : ") & IIf(Len(p) = 0, T("(non d{e2}fini)"), p) & vbCrLf
    If Len(p) > 0 Then s = s & IIf(FichierExiste(p), T("Fichier accessible."), T("Fichier INTROUVABLE.")) & vbCrLf
    ' Chiffrement : le central doit l'être comme le classeur local (voir "Protection par passphrase").
    s = s & T("Classeur local chiffr{e2} : ") & IIf(ThisWorkbook.HasPassword, T("oui"), T("NON")) & vbCrLf
    If Len(p) > 0 Then
        If FichierExiste(p) Then s = s & T("Fichier central chiffr{e2} : ") & IIf(EstChiffre(p), T("oui"), T("NON")) & vbCrLf
    End If
    s = s & T("Passphrase m{e2}moris{e2}e pour cette session : ") & IIf(Len(mPassphrase) > 0 Or Len(LirePassphraseSession()) > 0, T("oui"), T("non")) & vbCrLf
    s = s & T("Derni{e1}re synchronisation r{e2}ussie : ") & LirePropriete(PROP_DATE)
    MsgBox s, vbInformation, TITRE
End Sub

Public Sub EffacerBarreEtat()
    Application.StatusBar = False
End Sub

' =====================================================================================
' SYNCHRONISATION (routine principale)
' Renvoie Vrai si les deux fichiers sont identiques à la fin (ou n'avaient rien à échanger).
' =====================================================================================
Private Function Synchroniser(ByVal moment As String) As Boolean

    Dim chemin As String, action As String, incompat As String
    Dim dL As Object, dN As Object, dApres As Object
    Dim hL As String, hN As String, base As String, hFinal As String
    Dim ecarts As String, txt As String, ok As Boolean, conflit As Boolean, msg As String

    If mEnCours Then Synchroniser = True: Exit Function
    mEnCours = True
    mResume = "": mNbChg = 0
    SauverEtat
    On Error GoTo Echec

    chemin = GetSetting(APP_REG, SECTION_REG, CLE_CHEMIN, "")

    ' Rien n'est configuré : on ne dérange pas l'opérateur à l'ouverture, on lui demandera à la fermeture.
    If Len(chemin) = 0 And moment = "ouverture" Then ok = True: GoTo Sortie
    ' Ce classeur EST le fichier central (ouvert directement) : rien à synchroniser.
    If StrComp(chemin, ThisWorkbook.FullName, vbTextCompare) = 0 Then ok = True: GoTo Sortie

    If Not FichierExiste(chemin) Then
        chemin = ChoisirCentral(moment, chemin)
        If Len(chemin) = 0 Then GoTo Sortie           ' l'opérateur renonce (ok reste Faux)
    End If

    Application.StatusBar = T("Synchronisation avec le fichier central...")
    If Not PoserVerrou(chemin) Then GoTo Sortie

    Set dL = CalculerEmpreintes(ThisWorkbook)
    hL = dL("TOTAL")

    If Not FichierExiste(chemin) Then
        action = "envoyer"                            ' premier envoi : création du fichier central
    Else
        OuvrirCentral chemin
        incompat = StructureIncompatible(mWbN)
        If Len(incompat) > 0 Then
            If MsgBox(T("Les deux fichiers n'ont pas la m{e2}me structure :") & vbCrLf & incompat & vbCrLf & vbCrLf & _
                      T("Cela arrive apr{e1}s une mise {a2} jour du classeur (colonne ajout{e2}e...).") & vbCrLf & _
                      T("Remplacer le fichier central par une copie compl{e1}te de VOTRE classeur ?") & vbCrLf & _
                      T("(une sauvegarde du fichier central est faite avant)"), _
                      vbYesNo + vbExclamation + vbDefaultButton2, TITRE) <> vbYes Then GoTo Sortie
            action = "envoyer"
        Else
            Set dN = CalculerEmpreintes(mWbN)
            hN = dN("TOTAL")
            base = LirePropriete(PROP_BASE)
            If hL = hN Then
                action = "rien"
            ElseIf Len(base) > 0 And hL = base Then
                action = "recuperer"
            ElseIf Len(base) > 0 And hN = base Then
                action = "envoyer"
            Else
                conflit = True
                action = ResoudreConflit(dL, dN, (Len(base) = 0))
                If Len(action) = 0 Then GoTo Sortie   ' annulé
            End If
        End If
    End If

    Select Case action
        Case "rien"
            ok = True
            hFinal = hL

        Case "recuperer"
            If conflit Then SauvegarderLocal
            RecupererDepuisCentral dL, dN
            Set dApres = CalculerEmpreintes(ThisWorkbook)
            ecarts = UnitesDifferentes(dApres, dN)
            If Len(ecarts) > 0 Then
                MsgBox T("Attention : apr{e1}s la r{e2}cup{e2}ration, ces {e2}l{e2}ments diff{e1}rent encore du fichier central :") & _
                       vbCrLf & ecarts, vbExclamation, TITRE
            Else
                ok = True
                hFinal = dApres("TOTAL")
            End If

        Case "envoyer"
            If Not dN Is Nothing Then txt = UnitesDifferentes(dL, dN)
            FermerCentral
            EnvoyerVersCentral chemin
            ' Contrôle : le central relu doit avoir exactement la même empreinte que le local.
            OuvrirCentral chemin
            Set dApres = CalculerEmpreintes(mWbN)
            ecarts = UnitesDifferentes(dL, dApres)
            If Len(ecarts) > 0 Then
                MsgBox T("Attention : apr{e1}s l'envoi, ces {e2}l{e2}ments diff{e1}rent encore dans le fichier central :") & _
                       vbCrLf & ecarts, vbExclamation, TITRE
            Else
                ok = True
                hFinal = hL
            End If
            If Len(txt) > 0 Then
                mResume = T("- Envoy{e2}s vers le fichier central : ") & txt & vbCrLf
            Else
                mResume = T("- Fichier central cr{e2}{e2} ou mis {a2} jour avec votre classeur.") & vbCrLf
            End If
    End Select

    If ok Then
        EcrirePropriete PROP_BASE, hFinal
        EcrirePropriete PROP_DATE, Format$(Now, "dd/mm/yyyy hh:nn")
    End If

    ' Résumé pour l'opérateur (à l'ouverture et à la demande ; jamais à la fermeture).
    If moment <> "fermeture" And ok Then
        If (action = "recuperer" Or action = "envoyer") Then
            txt = T("Fichier central : ") & chemin & vbCrLf & vbCrLf
            If action = "recuperer" Then
                txt = txt & T("Donn{e2}es r{e2}cup{e2}r{e2}es du fichier central :") & vbCrLf & mResume & vbCrLf & _
                      T("Les deux fichiers sont maintenant identiques.")
            Else
                txt = txt & T("Vos donn{e2}es ont {e2}t{e2} envoy{e2}es :") & vbCrLf & mResume & vbCrLf & _
                      T("Les deux fichiers sont maintenant identiques.")
            End If
            Application.StatusBar = False
            MsgBox txt, vbInformation, TITRE
        ElseIf moment = "manuel" Then
            MsgBox T("Les deux fichiers sont d{e2}j{a2} identiques. Rien {a2} synchroniser."), vbInformation, TITRE
        Else
            Application.StatusBar = T("Donn{e2}es {a2} jour avec le fichier central.")
            If moment = "ouverture" Then Application.OnTime Now + TimeSerial(0, 0, 6), "mod_Synchro.EffacerBarreEtat"
        End If
    End If

Sortie:
    On Error Resume Next
    FermerCentral
    LibererVerrou
    RestaurerEtat
    If moment <> "ouverture" Or Not ok Then Application.StatusBar = False
    mEnCours = False
    Synchroniser = ok
    Exit Function

Echec:
    msg = Err.Description
    If Err.Number = 70 Or Err.Number = 75 Or Err.Number = 76 Or Err.Number = 52 Then
        msg = msg & vbCrLf & T("(le fichier central est peut-{e1}tre ouvert par quelqu'un d'autre, ou le r{e2}seau est inaccessible)")
    End If
    On Error Resume Next
    MsgBox T("La synchronisation a {e2}chou{e2} :") & vbCrLf & msg, vbExclamation, TITRE
    ok = False
    Resume Sortie

End Function

' =====================================================================================
' CHOIX / MÉMORISATION DU FICHIER CENTRAL
' Renvoie le chemin retenu (et le mémorise), ou "" si l'opérateur renonce.
' =====================================================================================
Private Function ChoisirCentral(ByVal moment As String, ByVal ancien As String) As String

    Dim msg As String, rep As VbMsgBoxResult, v As Variant, s As String, ext As String

    ext = ExtensionDe(ThisWorkbook.Name)
    If Len(ancien) = 0 Then
        msg = T("Aucun fichier central n'est encore d{e2}fini.")
    ElseIf moment = "manuel" Then
        msg = T("Fichier central actuel :") & vbCrLf & ancien
    Else
        msg = T("Le fichier central est introuvable :") & vbCrLf & ancien
    End If
    msg = msg & vbCrLf & vbCrLf & _
          T("Oui = d{e2}signer un fichier central qui existe d{e2}j{a2}") & vbCrLf & _
          T("Non = cr{e2}er un nouveau fichier central (copie de ce classeur)") & vbCrLf & _
          T("Annuler = ne pas synchroniser cette fois")

    rep = MsgBox(msg, vbYesNoCancel + vbQuestion, TITRE)
    Select Case rep
        Case vbYes
            v = Application.GetOpenFilename(FileFilter:=T("Classeur Excel (*") & ext & "),*" & ext, _
                                            Title:=T("Emplacement du fichier central"))
            If VarType(v) = vbString Then ChoisirCentral = CStr(v)
        Case vbNo
            v = Application.GetSaveAsFilename(InitialFileName:=ThisWorkbook.Name, _
                                              FileFilter:=T("Classeur Excel (*") & ext & "),*" & ext, _
                                              Title:=T("Emplacement du nouveau fichier central"))
            If VarType(v) = vbString Then
                s = CStr(v)
                If LCase$(Right$(s, Len(ext))) <> LCase$(ext) Then s = s & ext
                ChoisirCentral = s
            End If
    End Select

    If Len(ChoisirCentral) > 0 Then
        If StrComp(ChoisirCentral, ThisWorkbook.FullName, vbTextCompare) = 0 Then
            MsgBox T("Le fichier central doit {e2}tre diff{e2}rent du classeur local."), vbExclamation, TITRE
            ChoisirCentral = ""
        Else
            SaveSetting APP_REG, SECTION_REG, CLE_CHEMIN, ChoisirCentral
        End If
    End If

End Function

' =====================================================================================
' CONFLIT : les deux fichiers ont changé depuis la dernière synchronisation.
' Renvoie "recuperer" (garder le réseau), "envoyer" (garder le local) ou "" (annulé).
' =====================================================================================
Private Function ResoudreConflit(ByVal dL As Object, ByVal dN As Object, ByVal premiere As Boolean) As String

    Dim msg As String, rep As VbMsgBoxResult

    If premiere Then
        msg = T("Premi{e1}re synchronisation : votre classeur et le fichier central contiennent des donn{e2}es diff{e1}rentes.")
    Else
        msg = T("Les donn{e2}es ont {e2}t{e2} modifi{e2}es des DEUX c{o2}t{e2}s depuis la derni{e1}re synchronisation.")
    End If
    msg = msg & vbCrLf & T("Diff{e2}rences : ") & UnitesDifferentes(dL, dN) & vbCrLf & vbCrLf & _
          T("Oui = garder la version du R{E2}SEAU (vos modifications locales sont remplac{e2}es)") & vbCrLf & _
          T("Non = garder MA version (le fichier central est remplac{e2})") & vbCrLf & _
          T("Annuler = ne rien faire") & vbCrLf & vbCrLf & _
          T("Dans les deux cas, la version {e2}cras{e2}e est sauvegard{e2}e dans le dossier ") & DOSSIER_SAUVEGARDES & "."

    rep = MsgBox(msg, vbYesNoCancel + vbExclamation + vbDefaultButton3, TITRE)
    Select Case rep
        Case vbYes: ResoudreConflit = "recuperer"
        Case vbNo:  ResoudreConflit = "envoyer"
    End Select

End Function

' =====================================================================================
' RÉCUPÉRATION : copie les données du central dans le classeur local
' (seulement les éléments dont l'empreinte diffère)
' =====================================================================================
Private Sub RecupererDepuisCentral(ByVal dL As Object, ByVal dN As Object)

    Dim u As Variant, k As String, nbIdentiques As Long
    Dim rs As Range, rd As Range

    ' 1. Les tableaux (d'abord : leur taille conditionne la lecture des feuilles)
    For Each u In ListeTables()
        k = "T:" & u
        If dN.Exists(k) Then
            If dL(k) <> dN(k) Then
                CopierTable TrouverTable(mWbN, CStr(u)), TrouverTable(ThisWorkbook, CStr(u)), CStr(u)
            Else
                nbIdentiques = nbIdentiques + 1
            End If
        End If
    Next u

    ' 2. Les feuilles de paramètres (cellules hors tableaux)
    For Each u In ListeFeuilles()
        k = "F:" & u
        If dN.Exists(k) Then
            If dL(k) <> dN(k) Then
                CopierFeuille TrouverFeuille(mWbN, CStr(u)), TrouverFeuille(ThisWorkbook, CStr(u)), CStr(u)
            Else
                nbIdentiques = nbIdentiques + 1
            End If
        End If
    Next u

    ' 3. Les cellules nommées
    For Each u In ListeNoms()
        k = "N:" & u
        If dN.Exists(k) Then
            If dL(k) <> dN(k) Then
                Set rs = PlageNom(mWbN, CStr(u))
                Set rd = PlageNom(ThisWorkbook, CStr(u))
                If Not rs Is Nothing And Not rd Is Nothing Then
                    rd.NumberFormat = rs.NumberFormat
                    rd.Value2 = rs.Value2
                    mResume = mResume & "- " & u & T(" : valeur mise {a2} jour") & vbCrLf
                    mNbChg = mNbChg + 1
                End If
            Else
                nbIdentiques = nbIdentiques + 1
            End If
        End If
    Next u

    If nbIdentiques > 0 Then
        mResume = mResume & T("- ") & nbIdentiques & T(" autre(s) {e2}l{e2}ment(s) d{e2}j{a2} identique(s).") & vbCrLf
    End If

End Sub

Private Sub CopierTable(ByVal src As ListObject, ByVal dst As ListObject, ByVal nom As String)

    Dim noms() As String, n As Long, k As Long, i As Long, idx As Long
    Dim sigS() As String, sigD() As String, nS As Long, nD As Long
    Dim corps As Variant, nCible As Long, a() As Variant
    Dim ajout As Long, retire As Long, fmt As String, ligne As String

    If src Is Nothing Or dst Is Nothing Then Exit Sub

    noms = ColonnesCopiables(src, n)
    sigD = SignaturesLignes(dst, noms, n, nD)
    sigS = SignaturesLignes(src, noms, n, nS)
    CompterDifferences sigD, nD, sigS, nS, ajout, retire

    corps = CorpsTable(src)
    nCible = nS
    If nCible < 1 Then nCible = 1
    AjusterLignes dst, nCible

    For k = 0 To n - 1
        idx = IndexColonne(src, noms(k))
        ReDim a(1 To nCible, 1 To 1)
        For i = 1 To nS
            a(i, 1) = corps(i, idx)
        Next i
        With dst.ListColumns(noms(k)).DataBodyRange
            If nS > 0 Then
                fmt = src.ListColumns(noms(k)).DataBodyRange.Cells(1, 1).NumberFormat
                .NumberFormat = fmt
            End If
            .Value2 = a
        End With
    Next k

    RecopierFormules dst

    ligne = "- " & nom & " : " & Format$(nS, "#,##0") & T(" ligne(s)")
    If ajout + retire > 0 Then
        ligne = ligne & " (" & ajout & T(" nouvelle(s) ou modifi{e2}e(s), ") & retire & T(" retir{e2}e(s) ou modifi{e2}e(s))")
    Else
        ligne = ligne & T(" (contenu identique)")
    End If
    mResume = mResume & ligne & vbCrLf
    mNbChg = mNbChg + ajout + retire

End Sub

' Amène le tableau à n lignes de données (n >= 1). Les cellules libérées sont vidées.
Private Sub AjusterLignes(ByVal lo As ListObject, ByVal n As Long)

    Dim nOld As Long, libere As Range

    If Not lo.DataBodyRange Is Nothing Then nOld = lo.DataBodyRange.Rows.Count
    If nOld = n Then Exit Sub

    If nOld = 0 Then
        lo.Resize lo.HeaderRowRange.Resize(n + 1)
    ElseIf n > nOld Then
        lo.Resize lo.Range.Resize(n + 1)
    Else
        Set libere = lo.DataBodyRange.Offset(n).Resize(nOld - n)
        lo.Resize lo.Range.Resize(n + 1)
        libere.Clear
    End If

End Sub

' Après un agrandissement, s'assure que les colonnes calculées sont remplies sur toutes les lignes.
Private Sub RecopierFormules(ByVal lo As ListObject)
    Dim lc As ListColumn
    On Error Resume Next
    For Each lc In lo.ListColumns
        If EstColonneFormule(lc) Then
            lc.DataBodyRange.FormulaR1C1 = lc.DataBodyRange.Cells(1, 1).FormulaR1C1
        End If
    Next lc
End Sub

Private Sub CopierFeuille(ByVal wsS As Worksheet, ByVal wsD As Worksheet, ByVal nom As String)

    Dim nl1 As Long, nc1 As Long, nl2 As Long, nc2 As Long, nl As Long, nc As Long
    Dim vs As Variant, vd As Variant, i As Long, j As Long, nb As Long

    If wsS Is Nothing Or wsD Is Nothing Then Exit Sub
    DimFeuille wsS, nl1, nc1
    DimFeuille wsD, nl2, nc2
    nl = IIf(nl1 > nl2, nl1, nl2)
    nc = IIf(nc1 > nc2, nc1, nc2)
    If nl = 0 Or nc = 0 Then Exit Sub

    vs = LireFeuille(wsS, nl, nc)
    vd = LireFeuille(wsD, nl, nc)
    For i = 1 To nl
        For j = 1 To nc
            If Etiquette(vs(i, j)) <> Etiquette(vd(i, j)) Then
                If Not wsD.Cells(i, j).HasFormula Then
                    If IsEmpty(vs(i, j)) Then
                        wsD.Cells(i, j).ClearContents
                    Else
                        wsD.Cells(i, j).NumberFormat = wsS.Cells(i, j).NumberFormat
                        wsD.Cells(i, j).Value2 = vs(i, j)
                    End If
                    nb = nb + 1
                End If
            End If
        Next j
    Next i

    If nb > 0 Then
        mResume = mResume & "- " & nom & " : " & nb & T(" cellule(s) mise(s) {a2} jour") & vbCrLf
        mNbChg = mNbChg + nb
    End If

End Sub

' =====================================================================================
' ENVOI : le central devient une copie complète du classeur local
' =====================================================================================
Private Sub EnvoyerVersCentral(ByVal chemin As String)

    Dim tmp As String, ext As String, dossier As String, motif As String, sauv As String

    ext = ExtensionDe(chemin)
    tmp = Left$(chemin, Len(chemin) - Len(ext)) & ".sync_tmp" & ext

    ' Sauvegarde de l'ancien central (s'il existe)
    If FichierExiste(chemin) Then
        sauv = CheminSauvegarde(chemin, "central", dossier, motif)
        FileCopy chemin, sauv
        PurgerSauvegardes dossier, motif
    End If

    If FichierExiste(tmp) Then Kill tmp
    ThisWorkbook.SaveCopyAs tmp
    If FichierExiste(chemin) Then Kill chemin
    Name tmp As chemin

End Sub

' Sauvegarde du classeur local avant qu'il soit rempli avec les données du réseau (cas d'un conflit).
Private Sub SauvegarderLocal()
    Dim dossier As String, motif As String, sauv As String
    sauv = CheminSauvegarde(ThisWorkbook.FullName, "local", dossier, motif)
    ThisWorkbook.SaveCopyAs sauv
    PurgerSauvegardes dossier, motif
End Sub

Private Function CheminSauvegarde(ByVal fichier As String, ByVal etiquette As String, _
                                  ByRef dossier As String, ByRef motif As String) As String
    Dim nom As String, ext As String
    dossier = DossierDe(fichier) & "\" & DOSSIER_SAUVEGARDES
    On Error Resume Next
    MkDir dossier
    On Error GoTo 0
    ext = ExtensionDe(fichier)
    nom = Mid$(fichier, InStrRev(fichier, "\") + 1)
    nom = Left$(nom, Len(nom) - Len(ext))
    motif = nom & "_*_" & etiquette & ext
    CheminSauvegarde = dossier & "\" & nom & "_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & etiquette & ext
End Function

' Ne garde que les NB_SAUVEGARDES plus récentes sauvegardes correspondant au motif.
Private Sub PurgerSauvegardes(ByVal dossier As String, ByVal motif As String)

    Dim noms() As String, n As Long, f As String, i As Long, j As Long, tmp As String

    f = Dir$(dossier & "\" & motif)
    Do While Len(f) > 0
        ReDim Preserve noms(0 To n)
        noms(n) = f
        n = n + 1
        f = Dir$
    Loop
    If n <= NB_SAUVEGARDES Then Exit Sub

    ' Tri alphabétique : l'horodatage dans le nom place les plus anciennes en premier.
    For i = 1 To n - 1
        tmp = noms(i)
        j = i - 1
        Do While j >= 0
            If StrComp(noms(j), tmp, vbTextCompare) <= 0 Then Exit Do
            noms(j + 1) = noms(j)
            j = j - 1
        Loop
        noms(j + 1) = tmp
    Next i
    On Error Resume Next
    For i = 0 To n - NB_SAUVEGARDES - 1
        Kill dossier & "\" & noms(i)
    Next i

End Sub

' =====================================================================================
' OUVERTURE DU CENTRAL (copie temporaire, lecture seule, macros désactivées)
' Une copie est nécessaire car Excel refuse d'ouvrir deux classeurs de même nom.
' =====================================================================================
Private Sub OuvrirCentral(ByVal chemin As String)

    Dim secu As Long, ext As String
    Dim chiffre As Boolean, essais As Long

    FermerCentral
    ext = ExtensionDe(chemin)
    mCopieTemp = Environ$("TEMP") & "\SuiviFinancier_sync_" & Format$(Now, "hhnnss") & ext
    FileCopy chemin, mCopieTemp
    chiffre = EstChiffre(mCopieTemp)

    secu = Application.AutomationSecurity
    Application.AutomationSecurity = msoAutomationSecurityForceDisable

    ' Central chiffré : on utilise la passphrase de la session, demandée une seule fois (3 essais).
    Do
        If chiffre And Len(mPassphrase) = 0 Then mPassphrase = LirePassphraseSession()
        If chiffre And Len(mPassphrase) = 0 Then
            If Not DemanderPassphrase(essais > 0) Then
                Application.AutomationSecurity = secu
                Err.Raise vbObjectError + 513, , T("Passphrase du fichier central non fournie.")
            End If
        End If
        essais = essais + 1
        On Error Resume Next
        If chiffre Then
            Set mWbN = Application.Workbooks.Open(Filename:=mCopieTemp, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, Password:=mPassphrase)
        Else
            Set mWbN = Application.Workbooks.Open(Filename:=mCopieTemp, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
        End If
        If Err.Number <> 0 Then
            Dim num As Long, desc As String
            num = Err.Number: desc = Err.Description
            Err.Clear
            On Error GoTo 0
            If chiffre And essais < 3 Then
                OublierPassphrase                     ' passphrase refusée : on la redemande
            Else
                Application.AutomationSecurity = secu
                Err.Raise num, , desc
            End If
        Else
            On Error GoTo 0
        End If
    Loop Until Not mWbN Is Nothing

    Application.AutomationSecurity = secu
    ThisWorkbook.Activate

End Sub

' Un classeur chiffré par mot de passe d'ouverture n'est plus une archive ZIP ("PK") mais un
' conteneur OLE (signature D0 CF 11 E0) : on le détecte sans avoir à l'ouvrir.
Private Function EstChiffre(ByVal chemin As String) As Boolean
    Dim f As Integer, b(0 To 3) As Byte
    On Error GoTo Fin
    f = FreeFile
    Open chemin For Binary Access Read As #f
    Get #f, , b
    Close #f
    EstChiffre = (b(0) = &HD0 And b(1) = &HCF And b(2) = &H11 And b(3) = &HE0)
    Exit Function
Fin:
    On Error Resume Next
    Close #f
End Function

' Demande la passphrase une seule fois par session Excel ; elle n'est JAMAIS écrite sur disque.
' Saisie MASQUÉE via le formulaire frm_Passphrase (créé par mod_InstallPassphrase.InstallerFormPassphrase) ;
' s'il n'est pas installé, repli sur une InputBox (saisie visible). Excel ne communique pas le mot de
' passe saisi à l'ouverture du classeur local : une seconde saisie est donc inévitable.
Private Function DemanderPassphrase(ByVal erreur As Boolean) As Boolean
    Dim v As String, msg As String, f As Object
    msg = T("Le fichier central est prot{e2}g{e2} par une passphrase.") & vbCrLf & _
          T("Saisissez-la (elle ne sera demand{e2}e qu'une fois pour cette session) :")
    If erreur Then msg = T("Passphrase incorrecte.") & vbCrLf & vbCrLf & msg

    ' Pendant la synchronisation l'écran est gelé (ScreenUpdating = False) : on le réactive le temps
    ' de la saisie, sinon le formulaire modal peut rester invisible et le classeur semble bloqué.
    Application.ScreenUpdating = True
    Application.StatusBar = False
    DoEvents

    On Error Resume Next
    Set f = VBA.UserForms.Add("frm_Passphrase")      ' liaison tardive : compile même si le formulaire est absent
    If Not f Is Nothing Then
        f.StartUpPosition = 2                         ' centré sur l'écran (pas sur la fenêtre Excel, peut-être cachée)
        Err.Clear
        f.Controls("lblMessage").Caption = msg        ' texte d'explication, posé directement sur l'étiquette
        f.Preparer msg                                ' remise à zéro de la saisie
        If Err.Number = 0 Then f.Show vbModal
        If Err.Number = 0 Then
            If f.Valide Then v = f.Saisie
            Unload f
        Else
            Err.Clear                                 ' formulaire défaillant : repli sur l'InputBox
            Set f = Nothing
        End If
    End If
    On Error GoTo 0
    If f Is Nothing Then v = InputBox(msg, TITRE)

    Application.ScreenUpdating = False

    If Len(v) > 0 Then
        mPassphrase = v
        EcrirePassphraseSession v
        DemanderPassphrase = True
    End If
End Function

' La variable mPassphrase est remise à zéro chaque fois que le projet VBA est réinitialisé (erreur non
' gérée, modification du code, fermeture puis réouverture du classeur dans la même fenêtre Excel...),
' ce qui redemandait la passphrase à la fermeture. On la double donc d'une variable d'environnement du
' PROCESSUS Excel : elle survit à ces réinitialisations, disparaît à la fermeture d'Excel et n'est
' JAMAIS écrite sur disque (ni dans le registre, ni dans le classeur).

Private Function LirePassphraseSession() As String
    On Error Resume Next
    LirePassphraseSession = CreateObject("WScript.Shell").Environment("Process")(VAR_SESSION)
End Function

Private Sub EcrirePassphraseSession(ByVal v As String)
    On Error Resume Next
    CreateObject("WScript.Shell").Environment("Process")(VAR_SESSION) = v
End Sub

Private Sub OublierPassphrase()
    mPassphrase = ""
    EcrirePassphraseSession ""
End Sub

Private Sub FermerCentral()
    On Error Resume Next
    If Not mWbN Is Nothing Then mWbN.Close SaveChanges:=False
    Set mWbN = Nothing
    If Len(mCopieTemp) > 0 Then Kill mCopieTemp
    mCopieTemp = ""
End Sub

' =====================================================================================
' VERROU (évite deux synchronisations simultanées sur le même fichier central)
' =====================================================================================
Private Function PoserVerrou(ByVal chemin As String) As Boolean

    Dim v As String, f As Integer, qui As String, age As Long

    v = chemin & ".verrou"
    Do While FichierExiste(v)
        age = DateDiff("n", FileDateTime(v), Now)
        If age >= VERROU_MINUTES Then Exit Do      ' verrou périmé (plantage d'un autre poste) : on l'écrase
        qui = ""
        On Error Resume Next
        f = FreeFile
        Open v For Input As #f
        Line Input #f, qui
        Close #f
        On Error GoTo 0
        If MsgBox(T("Le fichier central est en cours de synchronisation par :") & vbCrLf & qui & vbCrLf & vbCrLf & _
                  T("R{e2}essayer ?"), vbRetryCancel + vbExclamation, TITRE) <> vbRetry Then Exit Function
    Loop

    f = FreeFile
    Open v For Output As #f
    Print #f, Environ$("USERNAME") & " / " & Environ$("COMPUTERNAME") & " / " & Format$(Now, "dd/mm/yyyy hh:nn")
    Close #f
    mVerrou = v
    PoserVerrou = True

End Function

Private Sub LibererVerrou()
    On Error Resume Next
    If Len(mVerrou) > 0 Then Kill mVerrou
    mVerrou = ""
End Sub

' =====================================================================================
' COMPARAISON DES STRUCTURES
' Renvoie "" si les deux fichiers ont les mêmes tableaux / feuilles / noms / colonnes.
' =====================================================================================
Private Function StructureIncompatible(ByVal wbN As Workbook) As String

    Dim u As Variant, loL As ListObject, loN As ListObject, s As String, lc As ListColumn

    For Each u In ListeTables()
        Set loL = TrouverTable(ThisWorkbook, CStr(u))
        Set loN = TrouverTable(wbN, CStr(u))
        If (loL Is Nothing) <> (loN Is Nothing) Then
            s = s & "- " & T("tableau ") & u & T(" absent d'un des deux fichiers") & vbCrLf
        ElseIf Not loL Is Nothing Then
            For Each lc In loL.ListColumns
                If IndexColonne(loN, lc.Name) = 0 Then s = s & "- " & u & T(" : colonne ") & lc.Name & T(" absente du central") & vbCrLf
            Next lc
            For Each lc In loN.ListColumns
                If IndexColonne(loL, lc.Name) = 0 Then s = s & "- " & u & T(" : colonne ") & lc.Name & T(" absente du classeur local") & vbCrLf
            Next lc
        End If
    Next u
    For Each u In ListeFeuilles()
        If (TrouverFeuille(ThisWorkbook, CStr(u)) Is Nothing) <> (TrouverFeuille(wbN, CStr(u)) Is Nothing) Then
            s = s & "- " & T("feuille ") & u & T(" absente d'un des deux fichiers") & vbCrLf
        End If
    Next u
    For Each u In ListeNoms()
        If (PlageNom(ThisWorkbook, CStr(u)) Is Nothing) <> (PlageNom(wbN, CStr(u)) Is Nothing) Then
            s = s & "- " & T("cellule nomm{e2}e ") & u & T(" absente d'un des deux fichiers") & vbCrLf
        End If
    Next u
    StructureIncompatible = s

End Function

' =====================================================================================
' EMPREINTES DES DONNÉES
' Renvoie un dictionnaire : "T:<table>", "F:<feuille>", "N:<nom>" -> empreinte, et "TOTAL".
' =====================================================================================
Private Function CalculerEmpreintes(ByVal wb As Workbook) As Object

    Dim d As Object, u As Variant, lo As ListObject, ws As Worksheet, r As Range, k As Variant

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    For Each u In ListeTables()
        Set lo = TrouverTable(wb, CStr(u))
        If Not lo Is Nothing Then d.Add "T:" & u, EmpreinteTable(lo)
    Next u
    For Each u In ListeFeuilles()
        Set ws = TrouverFeuille(wb, CStr(u))
        If Not ws Is Nothing Then d.Add "F:" & u, EmpreinteFeuille(ws)
    Next u
    For Each u In ListeNoms()
        Set r = PlageNom(wb, CStr(u))
        If Not r Is Nothing Then d.Add "N:" & u, Etiquette(r.Value2)
    Next u

    HashReset
    For Each k In d.Keys
        HashAjouter CStr(k) & "=" & d(k) & "|"
    Next k
    d.Add "TOTAL", HashResultat()
    Set CalculerEmpreintes = d

End Function

' Liste (séparée par des virgules) des éléments qui diffèrent entre deux dictionnaires d'empreintes.
Private Function UnitesDifferentes(ByVal dA As Object, ByVal dB As Object) As String
    Dim k As Variant, s As String
    For Each k In dA.Keys
        If CStr(k) <> "TOTAL" Then
            If Not dB.Exists(k) Then
                s = s & ", " & Mid$(CStr(k), 3)
            ElseIf dA(k) <> dB(k) Then
                s = s & ", " & Mid$(CStr(k), 3)
            End If
        End If
    Next k
    For Each k In dB.Keys
        If CStr(k) <> "TOTAL" Then
            If Not dA.Exists(k) Then s = s & ", " & Mid$(CStr(k), 3)
        End If
    Next k
    If Len(s) > 0 Then s = Mid$(s, 3)
    UnitesDifferentes = s
End Function

Private Function EmpreinteTable(ByVal lo As ListObject) As String
    Dim noms() As String, n As Long, sig() As String, nEff As Long, i As Long, k As Long
    noms = ColonnesCopiables(lo, n)
    sig = SignaturesLignes(lo, noms, n, nEff)
    HashReset
    For k = 0 To n - 1
        HashAjouter noms(k) & ","
    Next k
    For i = 1 To nEff
        HashAjouter sig(i) & Chr$(30)
    Next i
    EmpreinteTable = HashResultat() & "/" & nEff
End Function

Private Function EmpreinteFeuille(ByVal ws As Worksheet) As String
    Dim nl As Long, nc As Long, v As Variant, i As Long, j As Long, e As String, nb As Long
    DimFeuille ws, nl, nc
    HashReset
    If nl > 0 And nc > 0 Then
        v = LireFeuille(ws, nl, nc)
        For i = 1 To nl
            For j = 1 To nc
                e = Etiquette(v(i, j))
                If Len(e) > 0 Then
                    HashAjouter i & "," & j & "=" & e & ";"
                    nb = nb + 1
                End If
            Next j
        Next i
    End If
    EmpreinteFeuille = HashResultat() & "/" & nb
End Function

' --- Hachage (deux cumuls 32 bits indépendants, en Double pour éviter tout débordement) ---
Private Sub HashReset()
    mH1 = 5381#
    mH2 = 52711#
End Sub

Private Sub HashAjouter(ByVal s As String)
    Dim b() As Byte, i As Long
    If LenB(s) = 0 Then Exit Sub
    b = s                                   ' octets UTF-16 de la chaîne
    For i = 0 To UBound(b)
        mH1 = mH1 * 131# + b(i)
        mH1 = mH1 - Int(mH1 / 4294967296#) * 4294967296#
        mH2 = mH2 * 257# + b(i) + 1#
        mH2 = mH2 - Int(mH2 / 4294967296#) * 4294967296#
    Next i
End Sub

Private Function HashResultat() As String
    HashResultat = Format$(mH1, "0") & "-" & Format$(mH2, "0")
End Function

' =====================================================================================
' LECTURE DES TABLEAUX ET FEUILLES
' =====================================================================================
Private Function TrouverTable(ByVal wb As Workbook, ByVal nom As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In wb.Worksheets
        Set lo = Nothing
        On Error Resume Next
        Set lo = ws.ListObjects(nom)
        On Error GoTo 0
        If Not lo Is Nothing Then Set TrouverTable = lo: Exit Function
    Next ws
End Function

Private Function TrouverFeuille(ByVal wb As Workbook, ByVal nom As String) As Worksheet
    On Error Resume Next
    Set TrouverFeuille = wb.Worksheets(nom)
    On Error GoTo 0
End Function

Private Function PlageNom(ByVal wb As Workbook, ByVal nom As String) As Range
    On Error Resume Next
    Set PlageNom = wb.Names(nom).RefersToRange
    On Error GoTo 0
End Function

Private Function IndexColonne(ByVal lo As ListObject, ByVal nom As String) As Long
    On Error Resume Next
    IndexColonne = lo.ListColumns(nom).Index
    On Error GoTo 0
End Function

' Une colonne "calculée" (formules) est ignorée : elle se recalcule seule dans chaque classeur.
Private Function EstColonneFormule(ByVal lc As ListColumn) As Boolean
    Dim v As Variant
    If lc.DataBodyRange Is Nothing Then Exit Function
    v = lc.DataBodyRange.HasFormula
    If VarType(v) = vbBoolean Then EstColonneFormule = v
End Function

' Noms des colonnes à synchroniser, triés par ordre alphabétique (l'ordre des colonnes
' dans le tableau n'a donc aucune importance). n = nombre de noms valides.
Private Function ColonnesCopiables(ByVal lo As ListObject, ByRef n As Long) As String()
    Dim t() As String, i As Long, j As Long, tmp As String
    n = 0
    ReDim t(0 To lo.ListColumns.Count)
    For i = 1 To lo.ListColumns.Count
        If Not EstColonneFormule(lo.ListColumns(i)) Then
            t(n) = lo.ListColumns(i).Name
            n = n + 1
        End If
    Next i
    For i = 1 To n - 1
        tmp = t(i)
        j = i - 1
        Do While j >= 0
            If StrComp(t(j), tmp, vbTextCompare) <= 0 Then Exit Do
            t(j + 1) = t(j)
            j = j - 1
        Loop
        t(j + 1) = tmp
    Next i
    ColonnesCopiables = t
End Function

' Valeurs du corps du tableau, toujours sous forme de tableau 2D (Empty s'il n'a aucune ligne).
Private Function CorpsTable(ByVal lo As ListObject) As Variant
    Dim v As Variant, a() As Variant
    If lo.DataBodyRange Is Nothing Then Exit Function
    v = lo.DataBodyRange.Value2
    If IsArray(v) Then
        CorpsTable = v
    Else
        ReDim a(1 To 1, 1 To 1)
        a(1, 1) = v
        CorpsTable = a
    End If
End Function

' Une "signature" par ligne (valeurs des colonnes retenues, typées, séparées par Chr(31)).
' Les lignes entièrement vides en fin de tableau sont ignorées : nEff = nombre de lignes retenues.
Private Function SignaturesLignes(ByVal lo As ListObject, ByRef noms() As String, ByVal n As Long, _
                                  ByRef nEff As Long) As String()
    Dim corps As Variant, idx() As Long, k As Long, i As Long, nl As Long, vide As Boolean
    Dim sig() As String, s As String

    nEff = 0
    corps = CorpsTable(lo)
    If IsEmpty(corps) Or n = 0 Then Exit Function

    ReDim idx(0 To n - 1)
    For k = 0 To n - 1
        idx(k) = IndexColonne(lo, noms(k))
    Next k

    nl = UBound(corps, 1)
    nEff = nl
    Do While nEff > 0
        vide = True
        For k = 0 To n - 1
            If idx(k) > 0 Then
                If Not EstVide(corps(nEff, idx(k))) Then vide = False: Exit For
            End If
        Next k
        If Not vide Then Exit Do
        nEff = nEff - 1
    Loop
    If nEff = 0 Then Exit Function

    ReDim sig(1 To nEff)
    For i = 1 To nEff
        s = ""
        For k = 0 To n - 1
            If idx(k) > 0 Then s = s & Etiquette(corps(i, idx(k)))
            s = s & Chr$(31)
        Next k
        sig(i) = s
    Next i
    SignaturesLignes = sig
End Function

' Compte, pour le résumé, les lignes nouvelles/modifiées et retirées/modifiées (comparaison de multi-ensembles).
Private Sub CompterDifferences(ByRef sigOld() As String, ByVal nOld As Long, ByRef sigNew() As String, _
                               ByVal nNew As Long, ByRef ajout As Long, ByRef retire As Long)
    Dim d As Object, i As Long, c As Long
    Set d = CreateObject("Scripting.Dictionary")
    For i = 1 To nOld
        If d.Exists(sigOld(i)) Then d(sigOld(i)) = d(sigOld(i)) + 1 Else d.Add sigOld(i), 1
    Next i
    ajout = 0
    For i = 1 To nNew
        If d.Exists(sigNew(i)) Then
            c = d(sigNew(i))
            If c > 0 Then d(sigNew(i)) = c - 1 Else ajout = ajout + 1
        Else
            ajout = ajout + 1
        End If
    Next i
    retire = nOld - (nNew - ajout)
End Sub

' Dernière ligne / colonne réellement remplies d'une feuille (0 si vide).
Private Sub DimFeuille(ByVal ws As Worksheet, ByRef nl As Long, ByRef nc As Long)
    Dim f As Range
    nl = 0: nc = 0
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    If f Is Nothing Then Exit Sub
    nl = f.Row
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    nc = f.Column
End Sub

' Valeurs de A1 à (nl, nc), les zones occupées par des tableaux étant mises à blanc.
Private Function LireFeuille(ByVal ws As Worksheet, ByVal nl As Long, ByVal nc As Long) As Variant
    Dim v As Variant, a() As Variant, lo As ListObject
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long, i As Long, j As Long

    v = ws.Range(ws.Cells(1, 1), ws.Cells(nl, nc)).Value2
    If Not IsArray(v) Then
        ReDim a(1 To 1, 1 To 1)
        a(1, 1) = v
        v = a
    End If
    For Each lo In ws.ListObjects
        r1 = lo.Range.Row
        r2 = r1 + lo.Range.Rows.Count - 1
        c1 = lo.Range.Column
        c2 = c1 + lo.Range.Columns.Count - 1
        If r2 > nl Then r2 = nl
        If c2 > nc Then c2 = nc
        For i = r1 To r2
            For j = c1 To c2
                v(i, j) = Empty
            Next j
        Next i
    Next lo
    LireFeuille = v
End Function

Private Function EstVide(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then
        EstVide = True
    ElseIf VarType(v) = vbString Then
        EstVide = (LenB(v) = 0)
    End If
End Function

' Représentation texte typée d'une valeur de cellule (indépendante des réglages régionaux).
Private Function Etiquette(ByVal v As Variant) As String
    If IsError(v) Then Etiquette = "e" & CStr(v): Exit Function
    Select Case VarType(v)
        Case vbEmpty, vbNull
            Etiquette = ""
        Case vbString
            If LenB(v) > 0 Then Etiquette = "s" & v
        Case vbBoolean
            Etiquette = IIf(v, "b1", "b0")
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal, vbDate
            Etiquette = "n" & Trim$(Str$(CDbl(v)))
        Case Else
            Etiquette = "x" & CStr(v)
    End Select
End Function

' =====================================================================================
' PROPRIÉTÉS DU CLASSEUR (empreinte de la dernière synchronisation)
' =====================================================================================
Private Function LirePropriete(ByVal nom As String) As String
    On Error Resume Next
    LirePropriete = CStr(ThisWorkbook.CustomDocumentProperties(nom).Value)
    On Error GoTo 0
End Function

Private Sub EcrirePropriete(ByVal nom As String, ByVal valeur As String)
    Dim p As Object
    On Error Resume Next
    Set p = ThisWorkbook.CustomDocumentProperties(nom)
    If p Is Nothing Then
        ThisWorkbook.CustomDocumentProperties.Add Name:=nom, LinkToContent:=False, Type:=4, Value:=valeur   ' 4 = texte
    Else
        p.Value = valeur
    End If
End Sub

' =====================================================================================
' OUTILS : fichiers, chemins, état d'Excel
' =====================================================================================
Private Function FichierExiste(ByVal p As String) As Boolean
    If Len(p) = 0 Then Exit Function
    On Error Resume Next
    FichierExiste = (Len(Dir$(p)) > 0)
    If Err.Number <> 0 Then FichierExiste = False
End Function

Private Function ExtensionDe(ByVal nom As String) As String
    Dim i As Long
    i = InStrRev(nom, ".")
    If i > InStrRev(nom, "\") Then ExtensionDe = Mid$(nom, i)
End Function

Private Function DossierDe(ByVal chemin As String) As String
    DossierDe = Left$(chemin, InStrRev(chemin, "\") - 1)
End Function

Private Sub SauverEtat()
    mEcran = Application.ScreenUpdating
    mEvt = Application.EnableEvents
    mCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub RestaurerEtat()
    Application.Calculation = mCalc
    Application.EnableEvents = mEvt
    Application.ScreenUpdating = mEcran
End Sub
