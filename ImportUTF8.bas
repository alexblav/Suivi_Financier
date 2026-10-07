'===============================================================
' MODULE : ImportUTF8
' OBJET : Lire un fichier UTF-8, convertir en ANSI, normaliser LF,
'         enregistrer dans Téléchargements et remplacer un module VBA.
'===============================================================

Option Explicit

'---------------------------------------------------------------
' Fonction : Choisir un fichier via une boîte de dialogue
'---------------------------------------------------------------
Function SelectFileUTF8() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = "Sélectionnez un fichier UTF-8 (.bas)"
        .Filters.Clear
        .Filters.Add "Modules VBA", "*.bas"
        .AllowMultiSelect = False

        If .Show = -1 Then
            SelectFileUTF8 = .SelectedItems(1)
        Else
            SelectFileUTF8 = ""
        End If
    End With
End Function

'---------------------------------------------------------------
' Fonction : Lire un fichier UTF-8 correctement
'---------------------------------------------------------------
Function ReadUTF8(path As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")

    'Lecture binaire
    stream.Type = 1
    stream.Open
    stream.LoadFromFile path

    'Passage en mode texte UTF-8
    stream.Position = 0
    stream.Type = 2
    stream.Charset = "UTF-8"

    ReadUTF8 = stream.ReadText
    stream.Close
End Function

'---------------------------------------------------------------
' Fonction : Convertir LF → CRLF
'---------------------------------------------------------------
Function NormalizeLineEndings(txt As String) As String
    NormalizeLineEndings = Replace(txt, vbLf, vbCrLf)
End Function

'---------------------------------------------------------------
' Procédure : Enregistrer en ANSI dans Téléchargements
'---------------------------------------------------------------
Sub SaveANSIToDownloads(filename As String, content As String)
    Dim stream As Object
    Dim downloads As String

    downloads = Environ$("USERPROFILE") & "\Downloads\" & filename

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "Windows-1252"
    stream.Open
    stream.WriteText content
    stream.SaveToFile downloads, 2
    stream.Close

    MsgBox "Fichier converti enregistré dans : " & downloads, vbInformation
End Sub

'---------------------------------------------------------------
' Fonction : Trouver automatiquement le module à remplacer
' Règle : module portant le même nom que le fichier (sans extension)
'---------------------------------------------------------------
Function DetectModuleName(filePath As String) As String
    Dim f As String, m As VBIDE.VBComponent

    f = Mid(filePath, InStrRev(filePath, "\") + 1)
    f = Replace(f, ".bas", "", , , vbTextCompare)

    For Each m In ThisWorkbook.VBProject.VBComponents
        If StrComp(m.Name, f, vbTextCompare) = 0 Then
            DetectModuleName = m.Name
            Exit Function
        End If
    Next m

    DetectModuleName = "" 'non trouvé
End Function

'---------------------------------------------------------------
' Procédure : Remplacer le contenu d’un module VBA
'---------------------------------------------------------------
Sub ReplaceModuleContent(moduleName As String, newContent As String)
    Dim vbComp As VBIDE.VBComponent

    On Error Resume Next
    Set vbComp = ThisWorkbook.VBProject.VBComponents(moduleName)
    On Error GoTo 0

    If vbComp Is Nothing Then
        MsgBox "Module '" & moduleName & "' introuvable.", vbCritical
        Exit Sub
    End If

    vbComp.CodeModule.DeleteLines 1, vbComp.CodeModule.CountOfLines
    vbComp.CodeModule.AddFromString newContent

    MsgBox "Module '" & moduleName & "' mis à jour.", vbInformation
End Sub

'---------------------------------------------------------------
' Procédure principale : tout enchaîner
'---------------------------------------------------------------
Sub ImportUTF8ConvertAndReplace()
    Dim sourcePath As String
    Dim raw As String, normalized As String
    Dim moduleName As String

    '1. Sélection du fichier
    sourcePath = SelectFileUTF8()
    If sourcePath = "" Then
        MsgBox "Aucun fichier sélectionné.", vbExclamation
        Exit Sub
    End If

    '2. Lecture UTF-8
    raw = ReadUTF8(sourcePath)

    '3. Normalisation des fins de ligne
    normalized = NormalizeLineEndings(raw)

    '4. Enregistrement dans Téléchargements
    SaveANSIToDownloads "module_converti.bas", normalized

    '5. Détection automatique du module à remplacer
    moduleName = DetectModuleName(sourcePath)

    If moduleName = "" Then
        MsgBox "Aucun module correspondant au nom du fichier n'a été trouvé." & vbCrLf & _
               "Nom recherché : " & Replace(Mid(sourcePath, InStrRev(sourcePath, "\") + 1), ".bas", ""), vbExclamation
        Exit Sub
    End If

    '6. Remplacement du module
    ReplaceModuleContent moduleName, normalized
End Sub
