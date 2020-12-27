Imports System.Text.RegularExpressions

Public Class clsRScript

    Public lstRStatements As List(Of clsRStatement)

    Private ReadOnly arrROperators() As String = {"::", "$", "@", "^", "%%", "%/%",
        "%*%", "%o%", "%x%", "%in%", "/", "*", "+", "-", ">>=", "<<=", "==", "!=", "!", "&",
        "&&", "|", "||", "~", "->", "->>", "<-", "<<-", "="}

    Private ReadOnly arrRPartialOperators() As String = {":", ">", ">>", "<", "<<", "->", "<", "<<"}

    Private ReadOnly arrRSeperators() As String = {",", ";", vbCr, vbLf, vbCrLf}

    Private ReadOnly arrRBrackets() As String = {"(", ")", "{", "}"}

    Private ReadOnly arrROperatorBrackets() As String = {"[", "]", "[[", "]]"}

    Public Sub New(strInput As String)
        If IsNothing(strInput) Then
            Exit Sub
        End If

        Dim lstLexemes As List(Of String) = GetLstLexemes(strInput)
    End Sub

    'TODO how should we deal with newlines?
    '  - assume that newlines always indicate the end of the current R statement (not always strictly true)?
    '  - Separate strInput into a list of newline terminated strings and call CreateTreeFromString for each?

    'Public Function GetLstLexemesAsStr(strInput As String) As String
    '    Dim lstLexemes As List(Of String) = GetLstLexemes(strInput)
    '    For Each strNew As String In lstLexemes

    '    Next
    '    Return GetLstLexemes(strInput)
    'End Function

    Public Function GetLstLexemes(strInput As String) As List(Of String)

        If (String.IsNullOrEmpty(strInput)) Then
            'TODO throw exception
            Return Nothing
        End If

        Dim lstLexemes = New List(Of String)
        Dim strTextNew As String = Nothing

        For Each chrNew As Char In strInput
            If IsValidLexeme(strTextNew & chrNew) Then
                strTextNew &= chrNew
            Else
                'add 'strTextNew' element to list
                lstLexemes.Add(strTextNew)
                strTextNew = chrNew
            End If
        Next

        ''if text remains that is not yet added to tree (occurs when string does not end in ')')
        'If Not IsNothing(strTextNew) Then
        lstLexemes.Add(strTextNew)
        'End If

        Return lstLexemes
    End Function

    Private Function IsReserved(strNew As String) As Boolean
        Return arrROperators.Contains(strNew) Or
               arrRSeperators.Contains(strNew) Or
               arrRBrackets.Contains(strNew) Or
               arrROperatorBrackets.Contains(strNew)
    End Function

    Private Function IsValidLexeme(strNew As String) As Boolean

        If (String.IsNullOrEmpty(strNew)) Then
            'TODO throw exception
            Return False
        End If

        If Not strNew = vbCrLf AndAlso Regex.IsMatch(strNew, ".+\n$") OrElse 'string is >1 char and ends in newline (useful to do this test first because it simplifies the regular expressions below)
                Regex.IsMatch(strNew, ".+\r$") OrElse                        'string is >1 char and ends in carriage return
                Regex.IsMatch(strNew, "^%.*%.+") OrElse                      'string is a user-defined operator followed by another character
                Regex.IsMatch(strNew, "^"".*"".+") Then                      'string is a quoted string followed by another character
            Return False
        End If

        'if string is a valid ...
        If Regex.IsMatch(strNew, "^[a-zA-Z0-9_\.]+$") OrElse 'variable name or keyword name (rules for variable names are actually stricter than this but this library assumes it is parsing valid R code)
                arrROperators.Contains(strNew) OrElse        'operator (e.g. '+')
                arrROperatorBrackets.Contains(strNew) OrElse 'bracket operator (e.g. '[')
                arrRPartialOperators.Contains(strNew) OrElse 'partial operator (e.g. ':')
                arrRSeperators.Contains(strNew) OrElse       'separator (e.g. ',')
                arrRBrackets.Contains(strNew) OrElse         'bracket (e.g. '{')
                Regex.IsMatch(strNew, "^ *$") OrElse         'sequence of spaces
                Regex.IsMatch(strNew, "^"".*") OrElse        'quoted string (starts with '"*')
                Regex.IsMatch(strNew, "^%.*") OrElse         'user-defined Operator (starts with '%*')
                Regex.IsMatch(strNew, "^#.*") Then           'comment (starts with '#*')
            Return True
        End If

        Return False
    End Function

End Class
