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

    Dim vbp As Object, comp As Object, c As Object, etape As String

    On Error GoTo Echec
    etape = "acces au projet VBA"
    Set vbp = ThisWorkbook.VBProject

    On Error Resume Next
    Set comp = vbp.VBComponents("frm_Passphrase")
    On Error GoTo Echec
    If Not comp Is Nothing Then
        etape = "suppression de l'ancien formulaire"
        vbp.VBComponents.Remove comp
        Set comp = Nothing
    End If

    etape = "creation du formulaire"
    Set comp = vbp.VBComponents.Add(3)              ' 3 = vbext_ct_MSForm
    comp.Name = "frm_Passphrase"
    comp.Properties("Caption") = "Passphrase"
    etape = "taille du formulaire"
    comp.Properties("Width") = 440
    comp.Properties("Height") = 200

    etape = "etiquette"
    Set c = comp.Designer.Controls.Add("Forms.Label.1", "lblMessage", True)
    c.Left = 12: c.Top = 10: c.Width = 410: c.Height = 52
    PoserPolice c, 12
    c.Caption = "Saisissez la passphrase."
    c.WordWrap = True

    etape = "champ de saisie"
    Set c = comp.Designer.Controls.Add("Forms.TextBox.1", "txtPass", True)
    c.Left = 12: c.Top = 70: c.Width = 410: c.Height = 28
    PoserPolice c, 14
    c.PasswordChar = "*"

    etape = "bouton OK"
    Set c = comp.Designer.Controls.Add("Forms.CommandButton.1", "btnOK", True)
    c.Left = 222: c.Top = 112: c.Width = 98: c.Height = 32
    PoserPolice c, 12
    c.Caption = "OK"
    c.Default = True

    etape = "bouton Annuler"
    Set c = comp.Designer.Controls.Add("Forms.CommandButton.1", "btnAnnuler", True)
    c.Left = 326: c.Top = 112: c.Width = 98: c.Height = 32
    PoserPolice c, 12
    c.Caption = "Annuler"
    c.Cancel = True

    etape = "code du formulaire"
    comp.CodeModule.AddFromString CodeFormulaire()

    MsgBox "Le formulaire frm_Passphrase est installe." & vbCrLf & _
           "La passphrase sera desormais saisie de facon masquee.", vbInformation, "Passphrase"
    Exit Sub

Echec:
    MsgBox "Installation impossible a l'etape : " & etape & vbCrLf & _
           "Erreur " & Err.Number & " : " & Err.Description & vbCrLf & vbCrLf & _
           "Si l'erreur concerne l'acces au projet VBA, verifiez Fichier > Options > Centre de gestion de la" & vbCrLf & _
           "confidentialite > Parametres des macros : 'Acces approuve au modele d'objet du projet VBA'.", vbExclamation, "Passphrase"
End Sub

' La police est cosmetique : si elle ne peut pas etre reglee, l'installation continue.
Private Sub PoserPolice(ByVal c As Object, ByVal taille As Single)
    On Error Resume Next
    c.Font.Size = taille
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
