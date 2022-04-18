Imports System.Text.RegularExpressions

Public Class clsRToken
    ''' <summary>   The different types of R element (function name, key word, comment etc.) 
    '''             that the token may represent. </summary>
    Public Enum typToken
        RSyntacticName
        RFunctionName
        RKeyWord
        RConstantString
        RComment
        RSpace
        RBracket
        RSeparator
        REndStatement
        REndScript
        RNewLine
        ROperatorUnaryLeft
        ROperatorUnaryRight
        ROperatorBinary
        ROperatorBracket
        RPresentation
        RInvalid
    End Enum

    ''' <summary>   The lexeme associated with the token. </summary>
    Public strTxt As String

    ''' <summary>   The token type (function name, key word, comment etc.).  </summary>
    Public enuToken As typToken

    ''' <summary>   The token's children. </summary>
    Public lstTokens As New List(Of clsRToken)

    '''--------------------------------------------------------------------------------------------
    ''' <summary>
    '''     Constructs a new token with lexeme <paramref name="strTxtNew"/> and token type 
    '''     <paramref name="enuTokenNew"/>.
    '''     <para>
    '''     A token is a string of characters that represent a valid R element, plus meta data about
    '''     the token type (identifier, operator, keyword, bracket etc.).
    '''     </para>
    ''' </summary>
    '''
    ''' <param name="strTxtNew">    The lexeme to associate with the token. </param>
    ''' <param name="enuTokenNew">  The token type (function name, key word, comment etc.). </param>
    '''--------------------------------------------------------------------------------------------
    Public Sub New(strTxtNew As String, enuTokenNew As typToken)
        strTxt = strTxtNew
        enuToken = enuTokenNew
    End Sub


    '''--------------------------------------------------------------------------------------------
    ''' <summary>
    '''     Constructs a token from <paramref name="strLexemeCurrent"/>. 
    '''     <para>
    '''     A token is a string of characters that represent a valid R element, plus meta data about
    '''     the token type (identifier, operator, keyword, bracket etc.).
    '''     </para><para>
    '''     <paramref name="strLexemePrev"/> and <paramref name="strLexemeNext"/> are needed
    '''     to correctly identify if <paramref name="strLexemeCurrent"/> is a unary or binary
    '''     operator.</para>
    ''' </summary>
    '''
    ''' <param name="strLexemePrev">    The non-space lexeme immediately to the left of
    '''                                 <paramref name="strLexemeCurrent"/>. </param>
    ''' <param name="strLexemeCurrent"> The lexeme to convert to a token. </param>
    ''' <param name="strLexemeNext">    The non-space lexeme immediately to the right of
    '''                                 <paramref name="strLexemeCurrent"/>. </param>
    '''
    '''--------------------------------------------------------------------------------------------
    Public Sub New(strLexemePrev As String, strLexemeCurrent As String, strLexemeNext As String, bLexemeNextOnSameLine As Boolean)
        'TODO refactor so that strLexemePrev and strLexemeNext are booleans rather than strings?
        If String.IsNullOrEmpty(strLexemeCurrent) Then
            Exit Sub
        End If

        strTxt = strLexemeCurrent

        If IsKeyWord(strLexemeCurrent) Then                   'reserved key word (e.g. if, else etc.)
            enuToken = clsRToken.typToken.RKeyWord
        ElseIf IsSyntacticName(strLexemeCurrent) Then
            If strLexemeNext = "(" AndAlso bLexemeNextOnSameLine Then
                enuToken = clsRToken.typToken.RFunctionName   'function name
            Else
                enuToken = clsRToken.typToken.RSyntacticName  'syntactic name
            End If
        ElseIf IsComment(strLexemeCurrent) Then               'comment (starts with '#*')
            enuToken = clsRToken.typToken.RComment
        ElseIf IsConstantString(strLexemeCurrent) Then        'string literal (starts with single or double quote)
            enuToken = clsRToken.typToken.RConstantString
        ElseIf IsNewLine(strLexemeCurrent) Then               'new line (e.g. '\n')
            enuToken = clsRToken.typToken.RNewLine
        ElseIf strLexemeCurrent = ";" Then                    'end statement
            enuToken = clsRToken.typToken.REndStatement
        ElseIf strLexemeCurrent = "," Then                    'parameter separator
            enuToken = clsRToken.typToken.RSeparator
        ElseIf IsSequenceOfSpaces(strLexemeCurrent) Then      'sequence of spaces (needs to be after separator check, 
            enuToken = clsRToken.typToken.RSpace              '    else linefeed is recognised as space)
        ElseIf IsBracket(strLexemeCurrent) Then               'bracket (e.g. '{')
            If strLexemeCurrent = "}" Then
                enuToken = clsRToken.typToken.REndScript
            Else
                enuToken = clsRToken.typToken.RBracket
            End If
        ElseIf IsOperatorBrackets(strLexemeCurrent) Then      'bracket operator (e.g. '[')
            enuToken = clsRToken.typToken.ROperatorBracket
        ElseIf IsOperatorUnary(strLexemeCurrent) AndAlso      'unary right operator (e.g. '!x')
                (String.IsNullOrEmpty(strLexemePrev) OrElse
                Not Regex.IsMatch(strLexemePrev, "[a-zA-Z0-9_\.)\]]$")) Then
            enuToken = clsRToken.typToken.ROperatorUnaryRight
        ElseIf strLexemeCurrent = "~" AndAlso                 'unary left operator (e.g. x~)
                (String.IsNullOrEmpty(strLexemeNext) OrElse
                Not bLexemeNextOnSameLine OrElse
                Not Regex.IsMatch(strLexemeNext, "^[a-zA-Z0-9_\.(\+\-\!~]")) Then
            enuToken = clsRToken.typToken.ROperatorUnaryLeft
        ElseIf IsOperatorReserved(strLexemeCurrent) OrElse    'binary operator (e.g. '+')
                Regex.IsMatch(strLexemeCurrent, "^%.*%$") Then
            enuToken = clsRToken.typToken.ROperatorBinary
        Else
            enuToken = clsRToken.typToken.RInvalid
        End If

    End Sub

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Creates and returns a clone of this object. </summary>
    '''
    ''' <exception cref="Exception">    Thrown when the object has an empty child token. </exception>
    '''
    ''' <returns>   A clone of this object. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Function CloneMe() As clsRToken
        Dim clsToken = New clsRToken(strTxt, enuToken)

        For Each clsTokenChild As clsRToken In lstTokens
            If IsNothing(clsTokenChild) Then
                Throw New Exception("Token has illegal empty child.")
            End If
            clsToken.lstTokens.Add(clsTokenChild.CloneMe)
        Next

        Return clsToken
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns true if <paramref name="strTxt"/> is a valid lexeme (either partial or 
    '''             complete), else returns false.
    '''             </summary>
    '''
    ''' <param name="strTxt">   A sequence of characters from a syntactically correct R script </param>
    '''
    ''' <returns>   True if <paramref name="strTxt"/> is a valid lexeme, else  false. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Shared Function IsValidLexeme(strTxt As String) As Boolean

        If (String.IsNullOrEmpty(strTxt)) Then
            Return False
        End If

        ' if string constant (starts with single, double or backtick)
        '    Note: String constants are the only lexemes that can contain newlines.
        '          So if we process string constants first, then it makes checks below simpler.
        If IsConstantString(strTxt) Then
            ' if string constant is closed and followed by another character (e.g. '"hello"\n')
            If Regex.IsMatch(strTxt, strTxt(0) & "(.|\n)*" & strTxt(0) & "(.|\n)+") Then
                Return False
            End If
            Return True
        End If

        'if string is not a valid lexeme ...
        If (Regex.IsMatch(strTxt, ".+\n$") _                 'string is >1 char and ends in newline
                    AndAlso Not (strTxt = vbCrLf OrElse IsConstantString(strTxt))) _
                OrElse Regex.IsMatch(strTxt, ".+\r$") _      'string is >1 char and ends in carriage return
                OrElse Regex.IsMatch(strTxt, "^%.*%.+") Then 'string is a user-defined operator followed by another character
            Return False
        End If

        'if string is a valid lexeme ...
        If IsSyntacticName(strTxt) OrElse                    'syntactic name or reserved word
                IsOperatorReserved(strTxt) OrElse            'operator (e.g. '+')
                IsOperatorBrackets(strTxt) OrElse            'bracket operator (e.g. '[')
                strTxt = "<<" OrElse                         'partial operator (e.g. ':')
                IsNewLine(strTxt) OrElse                     'newlines (e.g. '\n')
                strTxt = "," OrElse strTxt = ";" OrElse      'parameter separator or end statement
                IsBracket(strTxt) OrElse                     'bracket (e.g. '{')
                IsSequenceOfSpaces(strTxt) OrElse            'sequence of spaces
                IsOperatorUserDefined(strTxt) OrElse         'user-defined operator (starts with '%*')
                IsComment(strTxt) Then                       'comment (starts with '#*')
            Return True
        End If

        'if the string is not covered by any of the checks above, 
        '       then we assume by default, that it's not a valid lexeme
        Return False
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns true if <paramref name="strTxt"/> is a complete or partial 
    '''             valid R syntactic name or key word, else returns false.<para>
    '''             Please note that the rules for syntactic names are actually stricter than 
    '''             the rules used in this function, but this library assumes it is parsing valid 
    '''             R code. </para></summary>
    '''
    ''' <param name="strTxt">   The text to check. </param>
    '''
    ''' <returns>   True if <paramref name="strTxt"/> is a valid R syntactic name or key word, 
    '''             else returns false.</returns>
    '''--------------------------------------------------------------------------------------------
    Private Shared Function IsSyntacticName(strTxt As String) As Boolean
        If String.IsNullOrEmpty(strTxt) Then
            Return False
        End If
        Return Regex.IsMatch(strTxt, "^[a-zA-Z0-9_\.]+$") OrElse
               Regex.IsMatch(strTxt, "^`.*")
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns true if <paramref name="strTxt"/> is a complete or partial string 
    '''             constant, else returns false.<para>
    '''             String constants are delimited by a pair of single (‘'’), double (‘"’)
    '''             or backtick ('`') quotes and can contain all other printable characters. 
    '''             Quotes and other special characters within strings are specified using escape 
    '''             sequences. </para></summary>
    '''
    ''' <param name="strTxt">   The text to check. </param>
    '''
    ''' <returns>   True if <paramref name="strTxt"/> is a complete or partial string constant,
    '''             else returns false.</returns>
    '''--------------------------------------------------------------------------------------------
    Private Shared Function IsConstantString(strTxt As String) As Boolean
        If Not String.IsNullOrEmpty(strTxt) AndAlso (
                Regex.IsMatch(strTxt, "^"".*") OrElse
                Regex.IsMatch(strTxt, "^'.*") OrElse
                Regex.IsMatch(strTxt, "^`.*")) Then
            Return True
        End If
        Return False
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns true if <paramref name="strTxt"/> is a comment, else returns false.
    '''             <para>
    '''             Any text from a # character to the end of the line is taken to be a comment,
    '''             unless the # character is inside a quoted string. </para></summary>
    '''
    ''' <param name="strTxt">   The text to check. </param>
    '''
    ''' <returns>   True if <paramref name="strTxt"/> is a comment, else returns false.</returns>
    '''--------------------------------------------------------------------------------------------
    Private Shared Function IsComment(strTxt As String) As Boolean
        If Not String.IsNullOrEmpty(strTxt) AndAlso Regex.IsMatch(strTxt, "^#.*") Then
            Return True
        End If
        Return False
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns true if <paramref name="strTxt"/> is sequence of spaces (and no other 
    '''             characters), else returns false. </summary>
    '''
    ''' <param name="strTxt">   The text to check . </param>
    '''
    ''' <returns>   True  if <paramref name="strTxt"/> is sequence of spaces (and no other 
    '''             characters), else returns false. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Shared Function IsSequenceOfSpaces(strTxt As String) As Boolean 'TODO make private?
        If Not String.IsNullOrEmpty(strTxt) AndAlso
                Not strTxt = vbLf AndAlso
                Regex.IsMatch(strTxt, "^ *$") Then
            Return True
        End If
        Return False
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns true if <paramref name="strTxt"/> is a functional R element 
    '''             (i.e. not empty, and not a space, comment or new line), else returns false. </summary>
    '''
    ''' <param name="strTxt">   The text to check . </param>
    '''
    ''' <returns>   True  if <paramref name="strTxt"/> is a functional R element
    '''             (i.e. not a space, comment or new line), else returns false. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Shared Function IsElement(strTxt As String) As Boolean 'TODO make private?
        If Not (String.IsNullOrEmpty(strTxt) OrElse
                IsNewLine(strTxt) OrElse
                IsSequenceOfSpaces(strTxt) OrElse
                IsComment(strTxt)) Then
            Return True
        End If
        Return False
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns true if <paramref name="strTxt"/> is a complete or partial  
    '''             user-defined operator, else returns false.</summary>
    '''
    ''' <param name="strTxt">   The text to check. </param>
    '''
    ''' <returns>   True if <paramref name="strTxt"/> is a complete or partial  
    '''             user-defined operator, else returns false.</returns>
    '''--------------------------------------------------------------------------------------------
    Public Shared Function IsOperatorUserDefined(strTxt As String) As Boolean 'TODO make private?
        If Not String.IsNullOrEmpty(strTxt) AndAlso Regex.IsMatch(strTxt, "^%.*") Then
            Return True
        End If
        Return False
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns true if <paramref name="strTxt"/> is a resrved operator, else returns 
    '''             false.</summary>
    '''
    ''' <param name="strTxt">   The text to check. </param>
    '''
    ''' <returns>   True if <paramref name="strTxt"/> is a reserved operator, else returns false.
    '''             </returns>
    '''--------------------------------------------------------------------------------------------
    Public Shared Function IsOperatorReserved(strTxt As String) As Boolean 'TODO make private?
        Dim arrROperators() As String = {"::", ":::", "$", "@", "^", ":", "%%", "%/%",
        "%*%", "%o%", "%x%", "%in%", "/", "*", "+", "-", "<", ">", "<=", ">=", "==", "!=", "!", "&",
        "&&", "|", "||", "~", "->", "->>", "<-", "<<-", "="}
        Return arrROperators.Contains(strTxt)
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns true if <paramref name="strTxt"/> is a bracket operator, else returns 
    '''             false.</summary>
    '''
    ''' <param name="strTxt">   The text to check. </param>
    '''
    ''' <returns>   True if <paramref name="strTxt"/> is a bracket operator, else returns false.
    '''             </returns>
    '''--------------------------------------------------------------------------------------------
    Private Shared Function IsOperatorBrackets(strTxt As String) As Boolean
        Dim arrROperatorBrackets() As String = {"[", "]", "[[", "]]"}
        Return arrROperatorBrackets.Contains(strTxt)
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns true if <paramref name="strTxt"/> is a unary operator, else returns 
    '''             false.</summary>
    '''
    ''' <param name="strTxt">   The text to check. </param>
    '''
    ''' <returns>   True if <paramref name="strTxt"/> is a unary operator, else returns false.
    '''             </returns>
    '''--------------------------------------------------------------------------------------------
    Private Shared Function IsOperatorUnary(strTxt As String) As Boolean
        Dim arrROperatorUnary() As String = {"+", "-", "!", "~"}
        Return arrROperatorUnary.Contains(strTxt)
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns true if <paramref name="strTxt"/> is a bracket, else returns 
    '''             false.</summary>
    '''
    ''' <param name="strTxt">   The text to check. </param>
    '''
    ''' <returns>   True if <paramref name="strTxt"/> is a bracket, else returns false.
    '''             </returns>
    '''--------------------------------------------------------------------------------------------
    Private Shared Function IsBracket(strTxt As String) As Boolean
        Dim arrRBrackets() As String = {"(", ")", "{", "}"}
        Return arrRBrackets.Contains(strTxt)
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns true if <paramref name="strTxt"/> is a new line, else returns 
    '''             false.</summary>
    '''
    ''' <param name="strTxt">   The text to check. </param>
    '''
    ''' <returns>   True if <paramref name="strTxt"/> is a new line, else returns false.
    '''             </returns>
    '''--------------------------------------------------------------------------------------------
    Private Shared Function IsNewLine(strTxt As String) As Boolean
        Dim arrRNewLines() As String = {vbCr, vbLf, vbCrLf}
        Return arrRNewLines.Contains(strTxt)
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns true if <paramref name="strTxt"/> is a key word, else returns 
    '''             false.</summary>
    '''
    ''' <param name="strTxt">   The text to check. </param>
    '''
    ''' <returns>   True if <paramref name="strTxt"/> is a key word, else returns false.
    '''             </returns>
    '''--------------------------------------------------------------------------------------------
    Private Shared Function IsKeyWord(strTxt As String) As Boolean
        Dim arrKeyWords() As String = {"if", "else", "repeat", "while", "function", "for", "in", "next", "break"}
        Return arrKeyWords.Contains(strTxt)
    End Function

End Class
