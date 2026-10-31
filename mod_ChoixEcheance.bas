Attribute VB_Name = "mod_ChoixEcheance"
Option Explicit

' =====================================================================================
' MODULE : mod_ChoixEcheance
'
' RÔLE (à lire en premier, même si vous débutez) :
'   Formulaire frm_ChoixEcheance : pendant le contrôle des catégories d'un import,
'   l'opérateur peut désigner À LA MAIN l'échéance de l'échéancier qui correspond à
'   l'opération affichée, quand la détection automatique ne l'a pas proposée (écart de
'   montant ou de date trop grand, Tiers différent...).
'
'   Ce module ne gère QUE l'écran (affichage, choix, boutons). Il ne touche ni à
'   l'échéancier ni à l'opération : la liste des échéances candidates lui est fournie par
'   mod_Echeancier.RapprocherManuellement, qui reçoit en retour le choix de l'opérateur.
'
' FONCTIONNEMENT DE L'ÉCRAN :
'   - Le haut rappelle l'opération importée (lecture seule).
'   - Le champ "Échéance" est une liste déroulante des échéances candidates, de la plus
'     proche à la plus éloignée de l'opération. Chaque entrée indique les écarts de date et
'     de montant.
'   - Les lignes en dessous détaillent l'échéance choisie et expliquent pourquoi la
'     détection automatique ne l'a pas retenue. Ce sont des FORMULES qui lisent les
'     colonnes techniques masquées (Z à AG) : aucun événement de feuille n'est nécessaire.
'   - Boutons : Rapprocher / Défaire le rapprochement / Annuler / Aide.
'
' ÉVÉNEMENT DE FEUILLE : seul le verrou "modal" (Worksheet_Deactivate) doit être collé dans
' la feuille, voir CodeBehind_frm_ChoixEcheance.txt.
'
' À PROPOS DES ACCENTS : textes affichés via mod_Display.FR() ; commentaires en UTF-8.
' =====================================================================================

' --- Adresses des champs (colonne B = étiquettes, C = valeurs) ---------------------------
Public Const CE_ADR_DATE As String = "C7"
Public Const CE_ADR_TIERS As String = "C8"
Public Const CE_ADR_MONTANT As String = "C9"
Public Const CE_ADR_CATEGORIE As String = "C10"
Public Const CE_ADR_CHEQUE As String = "C11"
Public Const CE_ADR_CHOIX As String = "C14"        ' liste déroulante (zone fusionnée C14:E14)
Public Const CE_ADR_MESSAGE As String = "B23"      ' zone fusionnée B23:E23
' Détail de l'échéance choisie : lignes 15 à 21, une colonne technique par ligne (voir CE_COL_LISTE).
Public Const CE_LIGNE_DETAIL As Long = 15
Public Const CE_NB_DETAILS As Long = 7

' --- Zone technique masquée : libellés des échéances (Z), puis un détail par colonne (AA...) ---
Public Const CE_COL_LISTE As Long = 26
Public Const CE_LIGNE_MAX As Long = 500
Public Const CE_NOM_LISTE As String = "ListeChoixEcheance"

' --- Résultat de l'écran ---------------------------------------------------------------------
Public Const CE_ANNULE As Long = 0
Public Const CE_RAPPROCHER As Long = 1
Public Const CE_DEFAIRE As Long = 2

' --- État du formulaire ----------------------------------------------------------------------
Public g_ChoixEnCours As Boolean        ' Vrai tant que le formulaire est ouvert (boucle d'attente)
Public g_ChoixVerrouActif As Boolean    ' Vrai tant que Worksheet_Deactivate doit réactiver la feuille
Private g_ChoixAction As Long
Private g_ChoixLibelle As String        ' libellé choisi dans la liste
Private g_ChoixCourant As String        ' libellé de l'échéance déjà rapprochée à la main ("" si aucune)
Private g_ChoixNomFeuillePrec As String


' =====================================================================================
' OUVERTURE DU FORMULAIRE (modal : le code reste bloqué ici jusqu'à la fermeture)
' =====================================================================================
' Paramètres :
'   opDate ... opCheque : texte de l'opération à afficher (lecture seule)
'   libelles    : libellé de chaque échéance candidate, indices 1..nbEch (sans * ? ~)
'   details     : details(k, 1..CE_NB_DETAILS) = Libellé, Tiers, Catégorie, Montant, Date prévue,
'                 Périodicité, Écarts de l'échéance k
'   libelleCourant : libellé de l'échéance déjà rapprochée à la main avec cette opération ("" sinon)
'   messageInfo : texte affiché sous le formulaire
'   libelleChoisi : en sortie, libellé choisi (résultat CE_RAPPROCHER)
' Renvoie CE_RAPPROCHER, CE_DEFAIRE ou CE_ANNULE.
Public Function OuvrirChoixEcheance(ByVal opDate As String, ByVal opTiers As String, ByVal opMontant As String, _
                                    ByVal opCategorie As String, ByVal opCheque As String, _
                                    ByRef libelles() As String, ByRef details() As String, ByVal nbEch As Long, _
                                    ByVal libelleCourant As String, ByVal messageInfo As String, _
                                    ByRef libelleChoisi As String) As Long

    Dim ws As Worksheet

    OuvrirChoixEcheance = CE_ANNULE
    libelleChoisi = ""

    On Error GoTo Erreur

    Set ws = mod_InstallCommun.TrouverFeuille(mod_VarGlobales.NOM_FEUILLE_CHOIX_ECHEANCE)
    If ws Is Nothing Then
        MsgBox mod_Display.FR("La feuille '") & mod_VarGlobales.NOM_FEUILLE_CHOIX_ECHEANCE & mod_Display.FR("' est introuvable.") & vbCrLf & _
               mod_Display.FR("Ex{e2}cutez d'abord la macro InstallerChoixEcheance."), vbExclamation
        Exit Function
    End If

    g_ChoixNomFeuillePrec = ActiveSheet.Name
    g_ChoixAction = CE_ANNULE
    g_ChoixLibelle = ""
    g_ChoixCourant = libelleCourant

    RemplirFormulaire ws, opDate, opTiers, opMontant, opCategorie, opCheque, libelles, details, nbEch, libelleCourant, messageInfo

    ws.Visible = xlSheetVisible
    ws.Activate
    mod_InstallCommun.MasquerQuadrillage
    ws.Range(CE_ADR_CHOIX).Select
    ' Liste posée une seconde fois, feuille affichée et active (comme frm_Echeance).
    mod_InstallCommun.PoserListeDeroulante ws.Range(CE_ADR_CHOIX), CE_NOM_LISTE, True

    Application.EnableEvents = True
    g_ChoixEnCours = True
    g_ChoixVerrouActif = True
    Do While g_ChoixEnCours
        DoEvents
    Loop

    On Error Resume Next
    mod_InstallCommun.ActiverFeuille g_ChoixNomFeuillePrec
    On Error GoTo 0

    OuvrirChoixEcheance = g_ChoixAction
    libelleChoisi = g_ChoixLibelle
    Exit Function

Erreur:
    Application.EnableEvents = True
    g_ChoixEnCours = False
    g_ChoixVerrouActif = False
    MsgBox mod_Display.FR("Erreur inattendue dans le choix de l'{e2}ch{e2}ance :") & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
    OuvrirChoixEcheance = CE_ANNULE

End Function

Private Sub RemplirFormulaire(ByVal ws As Worksheet, ByVal opDate As String, ByVal opTiers As String, _
                              ByVal opMontant As String, ByVal opCategorie As String, ByVal opCheque As String, _
                              ByRef libelles() As String, ByRef details() As String, ByVal nbEch As Long, _
                              ByVal libelleCourant As String, ByVal messageInfo As String)

    Dim evenementsAvant As Boolean
    Dim numErr As Long, descErr As String
    Dim sortie() As Variant
    Dim k As Long, j As Long

    evenementsAvant = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Sortie

    ' --- Opération importée (lecture seule) ---
    ws.Range(CE_ADR_DATE).Value = opDate
    ws.Range(CE_ADR_TIERS).Value = opTiers
    ws.Range(CE_ADR_MONTANT).Value = opMontant
    ws.Range(CE_ADR_CATEGORIE).Value = opCategorie
    ws.Range(CE_ADR_CHEQUE).Value = opCheque

    ' --- Zone technique : libellés (colonne Z) puis détails (colonnes suivantes) ---
    ws.Range(ws.Cells(2, CE_COL_LISTE), ws.Cells(CE_LIGNE_MAX, CE_COL_LISTE + CE_NB_DETAILS)).ClearContents
    If nbEch > 0 Then
        If nbEch > CE_LIGNE_MAX - 1 Then nbEch = CE_LIGNE_MAX - 1
        ReDim sortie(1 To nbEch, 1 To 1 + CE_NB_DETAILS)
        For k = 1 To nbEch
            sortie(k, 1) = libelles(k)
            For j = 1 To CE_NB_DETAILS
                sortie(k, 1 + j) = details(k, j)
            Next j
        Next k
        With ws.Cells(2, CE_COL_LISTE).Resize(nbEch, 1 + CE_NB_DETAILS)
            .NumberFormat = "@"
            .Value2 = sortie
        End With
    End If
    mod_InstallCommun.PoserNomListeDynamique ws, CE_NOM_LISTE, CE_COL_LISTE, 2, CE_LIGNE_MAX

    ' --- Choix : l'échéance déjà rapprochée à la main, sinon vide ---
    ws.Range(CE_ADR_CHOIX).Value = libelleCourant
    mod_InstallCommun.PoserListeDeroulante ws.Range(CE_ADR_CHOIX), CE_NOM_LISTE, True

    ws.Range(CE_ADR_MESSAGE).Value = messageInfo
    ws.Range(CE_ADR_MESSAGE).Font.Color = RGB(90, 90, 90)

Sortie:
    numErr = Err.Number
    descErr = Err.Description
    Application.EnableEvents = evenementsAvant
    If numErr <> 0 Then Err.Raise numErr, "RemplirFormulaire", descErr

End Sub


' =====================================================================================
' ACTIONS DES BOUTONS (macros appelées par les boutons de la feuille)
' =====================================================================================

' --- Bouton "Rapprocher" ---
Public Sub ChoixEcheanceRapprocher()

    Dim ws As Worksheet
    Dim choix As String
    Dim lignes As Variant
    Dim r As Long
    Dim trouve As Boolean

    If Not g_ChoixEnCours Then Exit Sub
    On Error GoTo Erreur
    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_CHOIX_ECHEANCE)

    choix = mod_DataStructure.CellText(ws.Range(CE_ADR_CHOIX).Value)
    If choix = "" Then
        AfficherMessage ws, "Choisissez une {e2}ch{e2}ance dans la liste."
        Exit Sub
    End If

    ' La saisie doit correspondre EXACTEMENT à une entrée de la liste (un collage peut contourner le menu).
    lignes = ws.Range(ws.Cells(2, CE_COL_LISTE), ws.Cells(CE_LIGNE_MAX, CE_COL_LISTE)).Value2
    For r = 1 To UBound(lignes, 1)
        If Not IsError(lignes(r, 1)) Then
            If CStr(lignes(r, 1)) = choix Then
                trouve = True
                Exit For
            End If
        End If
    Next r
    If Not trouve Then
        AfficherMessage ws, "Cette {e2}ch{e2}ance ne figure pas dans la liste : choisissez-la dans le menu d{e2}roulant."
        Exit Sub
    End If

    g_ChoixLibelle = choix
    g_ChoixAction = CE_RAPPROCHER
    Fermer ws
    Exit Sub

Erreur:
    Application.EnableEvents = True
    MsgBox mod_Display.FR("Erreur dans ChoixEcheanceRapprocher : ") & Err.Number & " - " & Err.Description, vbCritical

End Sub

' --- Bouton "Défaire le rapprochement" ---
Public Sub ChoixEcheanceDefaire()

    Dim ws As Worksheet

    If Not g_ChoixEnCours Then Exit Sub
    On Error GoTo Erreur
    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_CHOIX_ECHEANCE)

    If g_ChoixCourant = "" Then
        AfficherMessage ws, "Cette op{e2}ration n'est rapproch{e2}e {a2} la main d'aucune {e2}ch{e2}ance : il n'y a rien {a2} d{e2}faire."
        Exit Sub
    End If

    g_ChoixAction = CE_DEFAIRE
    Fermer ws
    Exit Sub

Erreur:
    Application.EnableEvents = True
    MsgBox mod_Display.FR("Erreur dans ChoixEcheanceDefaire : ") & Err.Number & " - " & Err.Description, vbCritical

End Sub

' --- Bouton "Annuler" ---
Public Sub ChoixEcheanceAnnuler()

    Dim ws As Worksheet

    If Not g_ChoixEnCours Then Exit Sub
    On Error GoTo Erreur
    Set ws = ThisWorkbook.Worksheets(mod_VarGlobales.NOM_FEUILLE_CHOIX_ECHEANCE)

    g_ChoixAction = CE_ANNULE
    Fermer ws
    Exit Sub

Erreur:
    Application.EnableEvents = True
    MsgBox mod_Display.FR("Erreur dans ChoixEcheanceAnnuler : ") & Err.Number & " - " & Err.Description, vbCritical

End Sub

' Verrou "modal" : tant que le formulaire est ouvert, impossible de le quitter en changeant d'onglet.
' Appelé par le code de la feuille (Worksheet_Deactivate).
Public Sub ChoixEcheanceVerrouiller(ByVal ws As Worksheet)
    If g_ChoixVerrouActif And Not mod_Aide.AideOuverte() Then ws.Activate
End Sub

Private Sub AfficherMessage(ByVal ws As Worksheet, ByVal texte As String)
    ws.Range(CE_ADR_MESSAGE).Value = mod_Display.FR(texte)
    ws.Range(CE_ADR_MESSAGE).Font.Color = RGB(192, 80, 0)
End Sub

' g_ChoixEnCours doit passer à False AVANT de masquer la feuille, sinon le verrou la réactiverait.
Private Sub Fermer(ByVal ws As Worksheet)
    g_ChoixEnCours = False
    g_ChoixVerrouActif = False
    ws.Range(ws.Cells(2, CE_COL_LISTE), ws.Cells(CE_LIGNE_MAX, CE_COL_LISTE + CE_NB_DETAILS)).ClearContents
    mod_Aide.MasquerFormulaire ws
End Sub
