Option Explicit

' =====================================================================================
' MODULE : mod_InstallPassphrase  (ajout 11/10/2026)
'
' Cree le formulaire frm_Passphrase : une petite fenetre de saisie MASQUEE (caracteres *)
' utilisee par mod_Synchro pour demander la passphrase du fichier central. Sans ce
' formulaire, mod_Synchro retombe sur une InputBox (saisie visible).
'
' A lancer UNE SEULE FOIS (Alt+F8 > InstallerFormPassphrase), comme les autres modules
' mod_Install*. Necessite "Acces approuve au modele d'objet du projet VBA" (deja requis
' par ImportUTF8). Relancer la macro recree le formulaire.
' =====================================================================================

Public Sub InstallerFormPassphrase()

    Dim vbp As Object, comp As Object, c As Object

    On Error GoTo Echec
    Set vbp = ThisWorkbook.VBProject

    On Error Resume Next
    Set comp = vbp.VBComponents("frm_Passphrase")
    On Error GoTo Echec
    If Not comp Is Nothing Then vbp.VBComponents.Remove comp

    Set comp = vbp.VBComponents.Add(3)              ' 3 = vbext_ct_MSForm
    comp.Name = "frm_Passphrase"
    comp.Properties("Caption") = "Passphrase"
    comp.Properties("Width") = 300
    comp.Properties("Height") = 130

    Set c = comp.Designer.Controls.Add("Forms.Label.1", "lblMessage", True)
    c.Left = 12: c.Top = 10: c.Width = 270: c.Height = 40
    c.Caption = ""

    Set c = comp.Designer.Controls.Add("Forms.TextBox.1", "txtPass", True)
    c.Left = 12: c.Top = 56: c.Width = 270: c.Height = 18
    c.PasswordChar = "*"

    Set c = comp.Designer.Controls.Add("Forms.CommandButton.1", "btnOK", True)
    c.Left = 130: c.Top = 86: c.Width = 70: c.Height = 22
    c.Caption = "OK"
    c.Default = True

    Set c = comp.Designer.Controls.Add("Forms.CommandButton.1", "btnAnnuler", True)
    c.Left = 212: c.Top = 86: c.Width = 70: c.Height = 22
    c.Caption = "Annuler"
    c.Cancel = True

    comp.CodeModule.AddFromString CodeFormulaire()

    MsgBox "Le formulaire frm_Passphrase est installe." & vbCrLf & _
           "La passphrase sera desormais saisie de facon masquee.", vbInformation, "Passphrase"
    Exit Sub

Echec:
    MsgBox "Installation impossible : " & Err.Description & vbCrLf & vbCrLf & _
           "Verifiez Fichier > Options > Centre de gestion de la confidentialite > Parametres des macros :" & vbCrLf & _
           "cocher 'Acces approuve au modele d'objet du projet VBA'.", vbExclamation, "Passphrase"
End Sub

Private Function CodeFormulaire() As String
    ' Une instruction par ligne : VBA limite à 24 les continuations de ligne ( _ ) d'une instruction.
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "Private mValide As Boolean" & vbCrLf
    s = s & "Public Property Get Valide() As Boolean" & vbCrLf
    s = s & "    Valide = mValide" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "Public Property Get Saisie() As String" & vbCrLf
    s = s & "    Saisie = txtPass.Text" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "Public Sub Preparer(ByVal texte As String)" & vbCrLf
    s = s & "    lblMessage.Caption = texte" & vbCrLf
    s = s & "    txtPass.Text = """"" & vbCrLf
    s = s & "    mValide = False" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnOK_Click()" & vbCrLf
    s = s & "    mValide = True" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub btnAnnuler_Click()" & vbCrLf
    s = s & "    mValide = False" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    If CloseMode = vbFormControlMenu Then" & vbCrLf
    s = s & "        Cancel = True" & vbCrLf
    s = s & "        mValide = False" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodeFormulaire = s
End Function
