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
        Dim lstLexemes = New List(Of String)
        Dim strTextNew As String = Nothing

        For Each chrNew As Char In strInput

            'if there are an odd number of double quotes then it means we are inside a string literal (e.g. inside '" - "')
            'Note:  this check is needed to ensure that special characters (e.g. brackets or operator characters) inside string literals are ignored
            If Not IsNothing(strTextNew) AndAlso strTextNew.Where(Function(c) c = Chr(34)).Count Mod 2 Then
                'append characters in the literal up to and including the 2nd double quote
                strTextNew &= chrNew
                Continue For
            End If

            'if there are an odd number of % characters then it means we are inside a string literal (e.g. inside '" - "')
            'Note:  this check is needed to ensure that special characters (e.g. brackets or operator characters) inside string literals are ignored
            If Not IsNothing(strTextNew) AndAlso strTextNew.Where(Function(c) c = Chr(34)).Count Mod 2 Then
                'append characters in the literal up to and including the 2nd double quote
                strTextNew &= chrNew
                Continue For
            End If

            'if new char is a reserved single char lexeme but is not joined to previous text as part of a multi-char reserved lexeme
            'e.g. new char is '-' but is not part of '<-'
            'OR 
            '  new char is NOT a single char reserved lexeme and previous text is a valid lexeme
            'If (IsReserved(chrNew) AndAlso (Not IsReserved(strTextNew & chrNew))) Or
            '           ((Not IsReserved(chrNew)) AndAlso IsReserved(strTextNew)) Then
            If IsValidLexeme(strTextNew & chrNew) Then
                strTextNew &= chrNew
            Else
                'add 'strTextNew' element to list
                lstLexemes.Add(strTextNew)
                strTextNew = chrNew
            End If
        Next

        'if text remains that is not yet added to tree (occurs when string does not end in ')')
        If Not IsNothing(strTextNew) Then
            lstLexemes.Add(strTextNew)
        End If

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

        'if string is >1 char and ends in newline
        '(useful to do this test first because it simplifies the regular expressions below)
        If Not strNew = vbCrLf AndAlso Regex.IsMatch(strNew, ".+\n$") Then
            Return False
        End If

        'if string is a user-defined operator followed by another character
        If Regex.IsMatch(strNew, "^%.*%.+") Then
            Return False
        End If

        'if string is a valid ...
        If Regex.IsMatch(strNew, "^[a-zA-Z0-9_\.]+$") OrElse 'variable name (or keyword name)
                arrROperators.Contains(strNew) OrElse        'operator (e.g. '+')
                arrROperatorBrackets.Contains(strNew) OrElse 'bracket operator (e.g. '[')
                arrRPartialOperators.Contains(strNew) OrElse 'partial operator (e.g. ':')
                arrRSeperators.Contains(strNew) OrElse       'separator (e.g. ',')
                arrRBrackets.Contains(strNew) OrElse         'bracket (e.g. '{')
                Regex.IsMatch(strNew, "^ *$") OrElse         'sequence of spaces
                Regex.IsMatch(strNew, "^%.*%$") OrElse       'user-defined operator (starts and ends with '%', e.g. '%>%')
                Regex.IsMatch(strNew, "^%.*") Then           'partial user-defined Operator (contains only '%*')
            Return True
        End If

        Return False
    End Function

End Class
