VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmResolutionCategories 
   Caption         =   "Resolution des categories ambigues"
   ClientHeight    =   5010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7215
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnTerminer 
      Caption         =   "Terminer et appliquer"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   4440
      Width           =   3255
   End
   Begin VB.CommandButton btnValider 
      Caption         =   "Valider ce cas"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   3255
   End
   Begin VB.ComboBox cboCategorie 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3960
      Width           =   6855
   End
   Begin VB.ListBox lstCas 
      Height          =   2400
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   6855
   End
   Begin VB.Label lblCategorie 
      Caption         =   "Categorie a retenir pour le cas selectionne :"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3660
      Width           =   4695
   End
   Begin VB.Label lblInstructions 
      Caption         =   "Certaines operations ont plusieurs categories possibles dans le fichier CSV. Selectionnez une ligne, choisissez la bonne categorie dans la liste deroulante, cliquez sur 'Valider ce cas'. Une fois tous les cas traites, cliquez sur 'Terminer et appliquer'."
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "frmResolutionCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
'  FORMULAIRE : frmResolutionCategories
'  RÔLE : Permettre à l'opérateur de choisir, pour chaque opération dont la
'         catégorie est ambiguë (plusieurs valeurs possibles trouvées dans le
'         CSV), la bonne valeur à retenir, sans jamais avoir à ouvrir ou
'         modifier le tableau data_import lui-même.
'
'  Ce formulaire lit et remplit les variables PUBLIQUES déclarées dans le
'  module mod_ImportOFX (g_NbCasAmbigus, g_CasTexte, g_CasCandidats, g_CasChoix).
'  C'est la macro principale qui les prépare avant d'ouvrir ce formulaire,
'  et qui relit g_CasChoix() une fois le formulaire fermé.
' ============================================================================

' Suivi local (dans le formulaire) de l'état "traité / pas traité" de chaque cas
Private m_Resolus() As Boolean


' Au chargement du formulaire : affichage de la liste des cas ambigus
Private Sub UserForm_Initialize()
    Dim i As Long

    lstCas.Clear
    cboCategorie.Clear

    If g_NbCasAmbigus > 0 Then
        ReDim m_Resolus(1 To g_NbCasAmbigus)
        For i = 1 To g_NbCasAmbigus
            lstCas.AddItem "[ ? ] " & g_CasTexte(i)
            m_Resolus(i) = False
        Next i
    End If

    AfficherCompteur
End Sub

' Quand l'opérateur clique sur un cas dans la liste : ses catégories
' candidates sont proposées dans la liste déroulante.
Private Sub lstCas_Click()
    Dim indexCas As Long
    Dim candidats() As String
    Dim i As Long

    If lstCas.ListIndex < 0 Then Exit Sub
    indexCas = lstCas.ListIndex + 1   ' les tableaux g_Cas* commencent à l'indice 1

    candidats = Split(g_CasCandidats(indexCas), ";")
    cboCategorie.Clear
    For i = LBound(candidats) To UBound(candidats)
        cboCategorie.AddItem candidats(i)
    Next i

    ' Si ce cas a déjà été résolu, on réaffiche le choix effectué.
    If g_CasChoix(indexCas) <> "" Then
        cboCategorie.Value = g_CasChoix(indexCas)
    Else
        cboCategorie.ListIndex = -1
    End If
End Sub

' Bouton "Valider ce cas" : enregistre le choix pour le cas sélectionné.
Private Sub btnValider_Click()
    Dim indexCas As Long

    If lstCas.ListIndex < 0 Then
        MsgBox "Selectionnez d'abord un cas dans la liste.", vbExclamation
        Exit Sub
    End If
    If cboCategorie.ListIndex < 0 Then
        MsgBox "Choisissez une categorie dans la liste deroulante.", vbExclamation
        Exit Sub
    End If

    indexCas = lstCas.ListIndex + 1
    g_CasChoix(indexCas) = cboCategorie.Value
    m_Resolus(indexCas) = True

    lstCas.List(lstCas.ListIndex) = "[OK] " & g_CasTexte(indexCas) & "  ->  " & cboCategorie.Value
    AfficherCompteur
End Sub

' Bouton "Terminer et appliquer" : ferme le formulaire. Les cas non traités
' conserveront une catégorie vide (modifiable ultérieurement).
Private Sub btnTerminer_Click()
    Dim nbRestants As Long, i As Long
    Dim reponse As VbMsgBoxResult

    nbRestants = 0
    For i = 1 To g_NbCasAmbigus
        If Not m_Resolus(i) Then nbRestants = nbRestants + 1
    Next i

    If nbRestants > 0 Then
        reponse = MsgBox(nbRestants & " cas n'ont pas ete traites. Leur categorie restera vide." & vbCrLf & _
                          "Voulez-vous fermer quand meme ?", vbYesNo + vbQuestion, "Cas non resolus")
        If reponse = vbNo Then Exit Sub
    End If

    Unload Me
End Sub

' Met à jour le titre de la fenêtre avec le nombre de cas restant à traiter.
Private Sub AfficherCompteur()
    Dim nbRestants As Long, i As Long
    nbRestants = 0
    For i = 1 To g_NbCasAmbigus
        If Not m_Resolus(i) Then nbRestants = nbRestants + 1
    Next i
    Me.Caption = "Resolution des categories ambigues (" & nbRestants & " restant(s) sur " & g_NbCasAmbigus & ")"
End Sub
