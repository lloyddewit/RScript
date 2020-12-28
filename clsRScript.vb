Imports System.Text.RegularExpressions

Public Class clsRScript

    Public lstRStatements As List(Of clsRStatement)

    Private ReadOnly arrROperators() As String = {"::", ":::", "$", "@", "^", ":", "%%", "%/%",
        "%*%", "%o%", "%x%", "%in%", "/", "*", "+", "-", "<", ">", "<=", ">=", "==", "!=", "!", "&",
        "&&", "|", "||", "~", "->", "->>", "<-", "<<-", "="}
    Private ReadOnly arrROperatorBrackets() As String = {"[", "]", "[[", "]]"}
    Private ReadOnly arrRPartialOperators() As String = {"<<"}
    Private ReadOnly arrRBrackets() As String = {"(", ")", "{", "}"}
    Private ReadOnly arrRSeperators() As String = {",", ";", vbCr, vbLf, vbCrLf}

    Public Sub New(strInput As String)
        If IsNothing(strInput) Then
            Exit Sub
        End If

        Dim lstLexemes As List(Of String) = GetLstLexemes(strInput)
        Dim lstTokens As List(Of clsRToken) = GetLstTokens(lstLexemes)
    End Sub

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns <paramref name="strRScript"/> as a list of its constituent lexemes 
    '''             A lexeme is a string of characters that represent a valid R element (identifier, operator, keyword, seperator, bracket etc.)
    '''             The lexem is just a string, it does not include any type information
    '''             This function identifies lexemes using a technique known as 'longest match' or 'maximal munch'
    '''             It processes the input characters one at a time until it reaches a character that is not in the set of characters acceptable for that lexeme (i.e. a sequence of characters that are not a valid R element)
    '''             This function assumes that <paramref name="strRScript"/> is a syntactically correct R script.
    '''             This function does not attempt to validate <paramref name="strRScript"/>.
    '''             If <paramref name="strRScript"/> is not a syntactically correct R script then the values returned by this function will also be invalid.
    '''             </summary>
    '''
    ''' <param name="strRScript"> The input. </param>
    '''
    ''' <returns>   The list lexemes. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Function GetLstLexemes(strRScript As String) As List(Of String)

        If (String.IsNullOrEmpty(strRScript)) Then
            'TODO throw exception
            Return Nothing
        End If

        Dim lstLexemes = New List(Of String)
        Dim strTextNew As String = Nothing

        For Each chrNew As Char In strRScript
            If IsValidLexeme(strTextNew & chrNew) Then
                strTextNew &= chrNew
            Else
                'add 'strTextNew' element to list
                lstLexemes.Add(strTextNew)
                strTextNew = chrNew
            End If
        Next

        lstLexemes.Add(strTextNew)

        Return lstLexemes
    End Function

    Public Function GetLstTokens(lstLexemes As List(Of String)) As List(Of clsRToken)

        If lstLexemes Is Nothing OrElse lstLexemes.Count = 0 Then
            'TODO throw exception
            Return Nothing
        End If

        Dim lstRTokens = New List(Of clsRToken)
        Dim strLexemePrev As String = ""
        Dim strLexemeCurrent As String = ""
        Dim strLexemeNext As String

        For intPos As Integer = 0 To lstLexemes.Count - 1

            'store previous non-space lexeme
            If Not String.IsNullOrEmpty(strLexemeCurrent) AndAlso Not Regex.IsMatch(strLexemeCurrent, "^ *$") Then
                strLexemePrev = strLexemeCurrent
            End If

            strLexemeCurrent = lstLexemes.Item(intPos)

            'find next non-space lexeme
            strLexemeNext = Nothing
            Dim strLexemeTmp As String
            For intNextPos As Integer = intPos + 1 To lstLexemes.Count - 1
                strLexemeTmp = lstLexemes.Item(intNextPos)
                If Not String.IsNullOrEmpty(strLexemeTmp) AndAlso Not Regex.IsMatch(strLexemeTmp, "^ *$") Then
                    strLexemeNext = strLexemeTmp
                    Exit For
                End If
            Next intNextPos

            lstRTokens.Add(GetRToken(strLexemePrev, strLexemeCurrent, strLexemeNext))
        Next intPos

        Return lstRTokens
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Query if 'strNew' is valid lexeme.
    '''              </summary>
    '''
    ''' <param name="strNew">   A sequence of characters taken from a syntactically correct R script </param>
    '''
    ''' <returns>   True if valid lexeme, false if not. </returns>
    '''--------------------------------------------------------------------------------------------
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
        If Regex.IsMatch(strNew, "^[a-zA-Z0-9_\.]+$") OrElse 'syntactic name or keyword (rules for syntactic names are actually stricter than this but this library assumes it is parsing valid R code)
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

    Private Function GetRToken(strLexemePrev As String, strLexemeCurrent As String, strLexemeNext As String) As clsRToken

        If String.IsNullOrEmpty(strLexemeCurrent) Then
            'TODO throw exception
            Return Nothing
        End If

        Dim clsRTokenNew As New clsRToken With {
            .strText = strLexemeCurrent
        }

        If Regex.IsMatch(strLexemeCurrent, "^[a-zA-Z0-9_\.]+$") Then 'syntactic name or keyword
            clsRTokenNew.enuToken = clsRToken.typToken.RSyntacticName
        ElseIf Regex.IsMatch(strLexemeCurrent, "^#.*") Then          'comment (starts with '#*')
            clsRTokenNew.enuToken = clsRToken.typToken.RComment
        ElseIf Regex.IsMatch(strLexemeCurrent, "^"".*") Then         'string literal (starts with '"*')
            clsRTokenNew.enuToken = clsRToken.typToken.RStringLiteral
        ElseIf arrRSeperators.Contains(strLexemeCurrent) Then        'separator (e.g. ',')
            clsRTokenNew.enuToken = clsRToken.typToken.RSeparator
        ElseIf Regex.IsMatch(strLexemeCurrent, "^ *$") Then          'sequence of spaces (needs to be after separator check, 
            clsRTokenNew.enuToken = clsRToken.typToken.RSpace        '        else linefeed is recognised as space)
        ElseIf arrRBrackets.Contains(strLexemeCurrent) Then          'bracket (e.g. '{')
            clsRTokenNew.enuToken = clsRToken.typToken.RBracket
        ElseIf arrROperatorBrackets.Contains(strLexemeCurrent) Then  'bracket operator (e.g. '[')
            clsRTokenNew.enuToken = clsRToken.typToken.ROperatorBracket
        ElseIf String.IsNullOrEmpty(strLexemePrev) OrElse            'unary right operator (e.g. '!x')
                Not Regex.IsMatch(strLexemePrev, "[a-zA-Z0-9_\.)\]]$") Then
            clsRTokenNew.enuToken = clsRToken.typToken.ROperatorUnaryRight
        ElseIf String.IsNullOrEmpty(strLexemeNext) OrElse            'unary left operator (e.g. x~)
                Not Regex.IsMatch(strLexemeNext, "^[a-zA-Z0-9_\.(]") Then
            clsRTokenNew.enuToken = clsRToken.typToken.ROperatorUnaryLeft
        Else                                                         'binary operator (e.g. '+')
            clsRTokenNew.enuToken = clsRToken.typToken.ROperatorBinary
        End If

        Return clsRTokenNew
    End Function

End Class
