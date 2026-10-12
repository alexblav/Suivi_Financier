Attribute VB_Name = "mod_DataStructure"
Option Explicit

' =====================================================================================
' MODULE : mod_DataStructure

' Ce module regroupe toutes les fonctions de CONVERSION SECURISEE.
' "Sécurisée" veut dire : si la cellule contient une valeur inattendue
' (vide, texte, erreur Excel comme #N/A...), la fonction ne plante pas
' le programme. Elle renvoie une valeur par défaut (0, "" ou False) à la place.
' C'est le même principe partout : on préfère une valeur neutre à un plantage.
'
' Toutes ces fonctions sont regroupées ici (et non éparpillées dans les
' modules "métier") pour qu'il n'existe qu'UN SEUL endroit à modifier
' si on doit un jour changer la façon dont on lit une cellule Excel.

' Les macros de ce module ne sont pas visibles dans la liste "Macros" d'Excel
Option Private Module

' Bornes des dates acceptées en saisie (voir TryLireDate et PoserValidationDate).
Public Const DATE_MIN_SAISIE As Long = 36526   ' 01/01/2000
Public Const DATE_MAX_SAISIE As Long = 73415   ' 31/12/2100

' ----------------------------------------------------------------------
' CellText : transforme n'importe quelle valeur de cellule en texte "propre"
' ----------------------------------------------------------------------
' "Propre" signifie : jamais d'erreur, et sans espaces inutiles au début/fin
' (grâce à Trim$). Utilisée par exemple pour comparer un libellé ou un type
' d'opération ("Positif"/"Négatif") de façon fiable, même si la cellule
' contient des espaces superflus copiés depuis la banque.
Public Function CellText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        CellText = ""
    Else
        ' CStr convertit en texte, Trim$ retire les espaces avant/après.
        CellText = Trim$(CStr(value))
    End If
End Function

' ----------------------------------------------------------------------
' ToDouble : équivalent de ToLong, mais pour les nombres à virgule (Double)
' ----------------------------------------------------------------------
' Utilisée notamment pour lire des montants en euros, qui peuvent avoir
' des décimales (ex : 12.50).
Public Function ToDouble(ByVal value As Variant) As Double
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Or Not IsNumeric(value) Then
        ' Valeur invalide : on retourne 0 pour ne jamais fausser une somme.
        ToDouble = 0
    Else
        ToDouble = CDbl(value)
    End If
End Function

' ----------------------------------------------------------------------
' ToLong : version "simplifiée" de TryGetLong pour les cas où on veut juste
' un nombre, sans avoir besoin de savoir si la conversion a réussi.
' ----------------------------------------------------------------------
' Une valeur invalide devient simplement 0. Pratique pour lire un paramètre
' (comme B1, B2, B3 dans la feuille Synthese) sans avoir à écrire un test
' à chaque fois dans le code appelant.
Public Function ToLong(ByVal value As Variant) As Long
    Dim result As Long
    ' On réutilise TryGetLong : si elle réussit, on récupère sa valeur.
    ' Si elle échoue, "result" reste à 0 (valeur par défaut d'un Long en VBA),
    ' et ToLong renvoie donc 0 automatiquement (valeur par défaut d'une fonction).
    If TryGetLong(value, result) Then ToLong = result
End Function

' ----------------------------------------------------------------------
' LECTURE STRICTE DES SAISIES OPERATEUR (dates et montants)
' ----------------------------------------------------------------------
' Contrairement a IsDate / IsNumeric (trop permissifs : "15/03" ou "10:30" passent pour
' des dates, "1e3" ou "$5" pour des nombres), ces fonctions n'acceptent que :
'   - une vraie date Excel (ou un numero de serie plausible) ou un texte JJ/MM/AAAA ;
'   - un vrai nombre, ou un texte de chiffres avec virgule ou point decimal.
' Elles servent de seconde barriere derriere la validation de cellule (posee par
' mod_InstallCommun.PoserChampDate / PoserChampMontant), que le copier-coller contourne.
Public Function MessageDateAttendue() As String
    MessageDateAttendue = "Format de date attendu : JJ/MM/AAAA (exemple : 15/03/2027)."
End Function

Public Function MessageMontantAttendu(Optional ByVal maxDecimales As Long = 2) As String
    MessageMontantAttendu = "Saisissez un nombre sans symbole ni unit" & ChrW(233) & ", avec une virgule pour les d" & ChrW(233) & "cimales (exemple : 98,43)."
End Function

' Renvoie Vrai et le jour (heure ignoree) si v est une date valable.
Public Function TryLireDate(ByVal v As Variant, ByRef jour As Date) As Boolean

    Dim s As String
    Dim p() As String
    Dim sep As String
    Dim j As Long, m As Long, a As Long
    Dim d As Date

    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function

    Select Case VarType(v)
        Case vbDate
            d = v
        Case vbDouble, vbSingle, vbCurrency, vbInteger, vbLong
            d = CDate(Int(CDbl(v)))
        Case vbString
            s = Trim$(Replace(v, ChrW(160), " "))
            If s = "" Then Exit Function
            If InStr(s, "/") > 0 Then
                sep = "/"
            ElseIf InStr(s, "-") > 0 Then
                sep = "-"
            ElseIf InStr(s, ".") > 0 Then
                sep = "."
            Else
                Exit Function
            End If
            p = Split(s, sep)
            If UBound(p) <> 2 Then Exit Function
            If Len(p(0)) < 1 Or Len(p(0)) > 2 Or Len(p(1)) < 1 Or Len(p(1)) > 2 Or Len(p(2)) <> 4 Then Exit Function
            If Not SeulementChiffres(p(0)) Or Not SeulementChiffres(p(1)) Or Not SeulementChiffres(p(2)) Then Exit Function
            j = CLng(p(0)): m = CLng(p(1)): a = CLng(p(2))
            If m < 1 Or m > 12 Or j < 1 Or j > 31 Then Exit Function
            d = DateSerial(a, m, j)
            If Day(d) <> j Or Month(d) <> m Then Exit Function   ' ex : 31/02 deborde sur mars
        Case Else
            Exit Function
    End Select

    If CLng(Int(CDbl(d))) < DATE_MIN_SAISIE Or CLng(Int(CDbl(d))) > DATE_MAX_SAISIE Then Exit Function
    jour = d
    TryLireDate = True

End Function

' Renvoie Vrai et la valeur si v est un montant/nombre valable (au plus maxDecimales decimales
' quand il est saisi en texte).
Public Function TryLireMontant(ByVal v As Variant, ByRef montant As Double, _
                               Optional ByVal maxDecimales As Long = 2) As Boolean

    Dim s As String
    Dim neg As Boolean
    Dim posSep As Long
    Dim partieEnt As String, partieDec As String

    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function

    Select Case VarType(v)
        Case vbDouble, vbSingle, vbCurrency, vbInteger, vbLong
            montant = CDbl(v)
            TryLireMontant = True
        Case vbString
            s = Replace(Replace(Replace(Trim$(v), " ", ""), ChrW(160), ""), ChrW(8364), "")
            If s = "" Then Exit Function
            If Left$(s, 1) = "-" Then
                neg = True
                s = Mid$(s, 2)
            ElseIf Left$(s, 1) = "+" Then
                s = Mid$(s, 2)
            End If
            If s = "" Then Exit Function
            posSep = InStr(s, ",")
            If posSep = 0 Then posSep = InStr(s, ".")
            If posSep > 0 Then
                partieEnt = Left$(s, posSep - 1)
                partieDec = Mid$(s, posSep + 1)
                If partieDec = "" Or Len(partieDec) > maxDecimales Then Exit Function
                If Not SeulementChiffres(partieDec) Then Exit Function
            Else
                partieEnt = s
            End If
            If partieEnt = "" Or Len(partieEnt) > 12 Then Exit Function
            If Not SeulementChiffres(partieEnt) Then Exit Function
            montant = Val(partieEnt & IIf(partieDec <> "", "." & partieDec, ""))   ' Val : point decimal, independant de la langue
            If neg Then montant = -montant
            TryLireMontant = True
    End Select

End Function

Private Function SeulementChiffres(ByVal s As String) As Boolean
    Dim k As Long
    If s = "" Then Exit Function
    For k = 1 To Len(s)
        If Mid$(s, k, 1) < "0" Or Mid$(s, k, 1) > "9" Then Exit Function
    Next k
    SeulementChiffres = True
End Function
