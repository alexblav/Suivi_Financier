Attribute VB_Name = "mod_Aide"
Option Explicit

' =====================================================================================
' MODULE : mod_Aide
'
' RÔLE : bascule entre un formulaire de travail frm_Nomform et sa feuille d'aide
' frm_Nomform_aide (le bloc "Instructions", déplacé hors du formulaire pour libérer de la
' place ; voir mod_InstallCommun.EcrireInstructions, qui construit la feuille d'aide).
'
'   - bouton "Aide" du formulaire ..... AfficherAide : ouvre l'aide SANS masquer le formulaire
'   - bouton "Sortir" de l'aide ....... SortirAide   : ferme l'aide seule et revient au formulaire
'   - bouton "Sortir" du formulaire ... MasquerFormulaire : ferme les DEUX (formulaire et aide,
'                                       si elle est encore ouverte)
'
' FORMULAIRES MODAUX (Ventilation, Contrôle des catégories, Échéance, Relevé, Contrat,
' Nouvelle catégorie) : leur code ramène de force l'opérateur sur la feuille dès qu'elle
' est quittée (événement Worksheet_Deactivate). Pendant que l'aide est affichée, ce
' réflexe est suspendu : il consulte AideOuverte().
' =====================================================================================

Private m_AideOuverte As Boolean

' Vrai tant que la feuille d'aide d'un formulaire est affichée à sa place.
Public Function AideOuverte() As Boolean
    AideOuverte = m_AideOuverte
End Function

' Bouton "Aide" d'un formulaire : le formulaire est la feuille active.
Public Sub AfficherAide()

    Dim ws As Worksheet
    Dim wsAide As Worksheet

    Set ws = ActiveSheet
    Set wsAide = mod_InstallCommun.TrouverFeuille(mod_InstallCommun.NomFeuilleAide(ws.Name))

    If wsAide Is Nothing Then
        MsgBox mod_Display.FR("La feuille d'aide de ce formulaire n'existe pas encore : relancez la macro de construction du formulaire (Ctrl+G)."), vbExclamation
        Exit Sub
    End If

    m_AideOuverte = True            ' avant de quitter le formulaire : suspend son verrou
    wsAide.Visible = xlSheetVisible
    wsAide.Activate
    mod_InstallCommun.MasquerQuadrillage

End Sub

' Bouton "Sortir" de l'aide : ferme l'aide seule ; le formulaire reste ouvert.
Public Sub SortirAide()

    Dim wsAide As Worksheet
    Dim ws As Worksheet
    Dim nom As String

    Set wsAide = ActiveSheet
    nom = wsAide.Name
    If Right$(nom, 5) <> "_aide" Then Exit Sub

    Set ws = mod_InstallCommun.TrouverFeuille(Left$(nom, Len(nom) - 5))
    If ws Is Nothing Then Exit Sub

    ws.Visible = xlSheetVisible
    ws.Activate
    mod_InstallCommun.MasquerQuadrillage
    wsAide.Visible = xlSheetVeryHidden
    m_AideOuverte = False           ' après le retour : le verrou du formulaire reprend

End Sub

' Sortie d'un formulaire : si son aide est encore affichée, on la ferme aussi.
Public Sub FermerAide(ByVal ws As Worksheet)

    Dim wsAide As Worksheet

    Set wsAide = mod_InstallCommun.TrouverFeuille(mod_InstallCommun.NomFeuilleAide(ws.Name))
    If Not wsAide Is Nothing Then
        If wsAide.Visible = xlSheetVisible Then
            On Error Resume Next
            wsAide.Visible = xlSheetVeryHidden
            On Error GoTo 0
        End If
    End If
    m_AideOuverte = False

End Sub

' Remplace "ws.Visible = xlSheetVeryHidden" à la sortie d'un formulaire : masque le
' formulaire et son aide.
Public Sub MasquerFormulaire(ByVal ws As Worksheet)
    FermerAide ws
    ws.Visible = xlSheetVeryHidden
End Sub
