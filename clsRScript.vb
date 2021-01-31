Imports System.Text.RegularExpressions

'TODO There are five types of constants: integer, logical, numeric, complex and string
'TODO Private ReadOnly arrSpecialConstants() As String = {"NULL", "NA", "Inf", "NaN"}
'TODO handle '...' (used in function definition)
'TODO handle '.' normally just part of a syntactic name, but has a special meaning when in a function name, or when referring to data (represents no variable)
'TODO insert invisible '{' and '}' for key words?
'TODO is it possible for packages to be nested (e.g. 'p1::p1_1::f()')?
'TODO handle function names as string constants. E.g 'x + y can equivalently be written "+"(x, y). Notice that since '+' is a non-standard function name, it needs to be quoted (see https://cran.r-project.org/doc/manuals/r-release/R-lang.html#Writing-functions)'
'TODO currently all newlines (vbLf, vbCr and vbCrLf) are converted to vbLf. Is it important to remember what the original new line character was?
' 
' 03/11/20
' - Process multiline input. Newline indicates end of statement when: 
'     1. line ends in an operator (excluding spaces and comments)  
'     2. odd number of '()' or '[]' are open ('{}' are a special case related to key words)   
' - Use assigned statements to manage the state of the R environment (ask David for more details)  
'
' 17/11/20
' - if comment on own line then link it to the next statement, else link to prev element  
' - support all assignments including '=' and assign to right operators  
' - store state of R environment   
'       e.g. don't treat variables as text but store object they represent (e.g. value or function deinition)
'       RSyntax already does this
' - allow named operator params (R-Insta allows operator params to be named, but this infor is lost in script)

Public Class clsRScript

    Public lstRStatements As List(Of clsRStatement)

    ''' <summary>   The current state of the token parsing. </summary>
    Private Enum typTokenState
        WaitingForOpenCondition
        WaitingForCloseCondition
        WaitingForStartScript
        WaitingForEndScript
    End Enum

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Parses the R script in <paramref name="strInput"/> and populates the list of
    '''             R statements.
    '''             <para>
    '''             This subroutine will accept, and correctly process all valid R. However, this 
    '''             class does not attempt to validate <paramref name="strInput"/>. If it is not 
    '''             valid R then this subroutine may still process the script without throwing an 
    '''             exception. In this case, the list of R statements will be undefined.
    '''             </para><para>
    '''             In other words, this subroutine will never generate false negatives (reject 
    '''             valid R) but may generate false positives (accept invalid R).
    '''             </para></summary>
    '''
    ''' <param name="strInput"> The R script to parse. This must be valid R according to the 
    '''                         R language specification at 
    '''                         https://cran.r-project.org/doc/manuals/r-release/R-lang.html. 
    '''                         </param>
    '''--------------------------------------------------------------------------------------------
    Public Sub New(strInput As String)
        If IsNothing(strInput) Then
            Exit Sub
        End If

        Dim lstLexemes As List(Of String) = GetLstLexemes(strInput)
        Dim lstTokens As List(Of clsRToken) = GetLstTokens(lstLexemes)
        lstRStatements = GetLstRStatements(lstTokens)

    End Sub

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns <paramref name="strRScript"/> as a list of its constituent lexemes. 
    '''             A lexeme is a string of characters that represent a valid R element 
    '''             (identifier, operator, keyword, seperator, bracket etc.). A lexeme does not 
    '''             include any type information.
    '''             <para>
    '''             This function identifies lexemes using a technique known as 'longest match' 
    '''             or 'maximal munch'. It keeps adding characters to the lexeme one at a time 
    '''             until it reaches a character that is not in the set of characters acceptable 
    '''             for that lexeme.
    '''             </para></summary>
    '''
    ''' <param name="strRScript"> The R script to convert (must be syntactically correct R). </param>
    '''
    ''' <returns>   <paramref name="strRScript"/> as a list of its constituent lexemes. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Function GetLstLexemes(strRScript As String) As List(Of String)

        If (String.IsNullOrEmpty(strRScript)) Then
            Return Nothing
        End If

        Dim lstLexemes = New List(Of String)
        Dim strTextNew As String = Nothing
        Dim stkIsSingleBracket As Stack(Of Boolean) = New Stack(Of Boolean)

        For Each chrNew As Char In strRScript
            If IsValidLexeme(strTextNew & chrNew) AndAlso
                    Not ((strTextNew & chrNew) = "]]" AndAlso 'edge case for nested operator brackets
                    (stkIsSingleBracket.Count < 1 OrElse stkIsSingleBracket.Peek())) Then
                strTextNew &= chrNew
            Else
                'Edge case: We need to handle nested operator brackets e.g. 'k[[l[[m[6]]]]]'. 
                '           For the above example, we need to recognise that the ']' to the right 
                '           of '6' is a single ']' bracket and is not part of a double ']]' bracket.
                '           To achieve this, we push each open bracket to a stack so that we know 
                '           which type of closing bracket is expected for each open bracket.
                If IsOperatorBrackets(strTextNew) Then
                    Select Case strTextNew
                        Case "["
                            stkIsSingleBracket.Push(True)
                        Case "[["
                            stkIsSingleBracket.Push(False)
                        Case "]", "]]"
                            If stkIsSingleBracket.Count < 1 Then
                                Throw New Exception("Closing bracket detected ('" & strTextNew & "') with no corresponding open bracket.")
                            End If
                            stkIsSingleBracket.Pop()
                        Case Else
                            Throw New Exception("Unknown bracket detected ('" & strTextNew & "').")
                    End Select
                End If

                'add 'strTextNew' element to list
                lstLexemes.Add(strTextNew)
                strTextNew = chrNew
            End If
        Next

        lstLexemes.Add(strTextNew)

        Return lstLexemes
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns <paramref name="lstLexemes"/> as a list of tokens.<para>
    '''             A token is a string of characters that represent a valid R element, plus meta 
    '''             data about the token type (identifier, operator, keyword, bracket etc.). 
    '''             </para></summary>
    '''
    ''' <param name="lstLexemes">   The list of lexemes to convert to tokens. </param>
    '''
    ''' <returns>   <paramref name="lstLexemes"/> as a list of tokens. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Function GetLstTokens(lstLexemes As List(Of String)) As List(Of clsRToken)

        If lstLexemes Is Nothing OrElse lstLexemes.Count = 0 Then
            Return Nothing
        End If

        Dim lstRTokens = New List(Of clsRToken)
        Dim strLexemePrev As String = ""
        Dim strLexemeCurrent As String = ""
        Dim strLexemeNext As String
        Dim bStatementComplete As Boolean = False
        Dim clsToken As clsRToken

        Dim stkNumOpenBrackets As Stack(Of Integer) = New Stack(Of Integer)
        stkNumOpenBrackets.Push(0)

        Dim stkIsScriptEnclosedByCurlyBrackets As Stack(Of Boolean) = New Stack(Of Boolean)
        stkIsScriptEnclosedByCurlyBrackets.Push(True)

        Dim stkTokenState As Stack(Of typTokenState) = New Stack(Of typTokenState)
        stkTokenState.Push(typTokenState.WaitingForEndScript)

        For intPos As Integer = 0 To lstLexemes.Count - 1

            If stkNumOpenBrackets.Count < 1 Then
                Throw New Exception("The stack storing the number of open brackets must have at least one value.")
            ElseIf stkIsScriptEnclosedByCurlyBrackets.Count < 1 Then
                Throw New Exception("The stack storing the number of open curly brackets must have at least one value.")
            ElseIf stkTokenState.Count < 1 Then
                Throw New Exception("The stack storing the current state of the token parsing must have at least one value.")
            End If

            'store previous non-space lexeme
            If IsElement(strLexemeCurrent) Then
                strLexemePrev = strLexemeCurrent
            End If

            strLexemeCurrent = lstLexemes.Item(intPos)

            'find next non-space lexeme
            strLexemeNext = Nothing
            For intNextPos As Integer = intPos + 1 To lstLexemes.Count - 1
                If Not String.IsNullOrEmpty(lstLexemes.Item(intNextPos)) AndAlso
                        Not IsSequenceOfSpaces(lstLexemes.Item(intNextPos)) Then
                    strLexemeNext = lstLexemes.Item(intNextPos)
                    Exit For
                End If
            Next intNextPos

            'determine whether the current sequence of tokens makes a complete valid R statement
            '    This is needed to determine whether a newline marks the end of the statement
            '    or is just for presentation.
            '    The current sequence of tokens is considered a complete valid R statement if it 
            '    has no open brackets and it does not end in an operator.
            Select Case strLexemeCurrent
                Case "(", "[", "[["
                    stkNumOpenBrackets.Push(stkNumOpenBrackets.Pop() + 1)
                Case ")", "]", "]]"
                    stkNumOpenBrackets.Push(stkNumOpenBrackets.Pop() - 1)
                Case "if", "while", "for", "function"
                    stkTokenState.Push(typTokenState.WaitingForOpenCondition)
                    stkNumOpenBrackets.Push(0)
                Case "else", "repeat"
                    stkTokenState.Push(typTokenState.WaitingForCloseCondition) ''else' and 'repeat' keywords have no condition (e.g. 'if (x==1) y<-0 else y<-1'
                    stkNumOpenBrackets.Push(0)                                 '    after the keyword is processed, the state will automatically change to 'WaitingForEndScript'
            End Select

            'identify the token associated with the current lexeme and add the token to the list
            clsToken = GetRToken(strLexemePrev, strLexemeCurrent, strLexemeNext)

            'Process key words
            '    Determine whether the next end statement will also be the end of the current script.
            '    Normally, a '}' indicates the end of the current script. However, R allows single
            '    statement scripts, not enclosed with '{}' for selected key words. 
            '    The key words that allow this are: if, else, while, for and function.
            '    For example:
            '        if(x <= 0) y <- log(1+x) else y <- log(x)
            If clsToken.enuToken = clsRToken.typToken.RComment OrElse       'ignore comments, spaces and newlines (they don't affect key word processing)
                    clsToken.enuToken = clsRToken.typToken.RSpace Then
                '    clsToken.enuToken = clsRToken.typToken.RNewLine Then
                ' clsToken.enuToken = clsRToken.typToken.RKeyWord Then    'ignore keywords (already processed above)
                'do nothing
            Else
                Select Case stkTokenState.Peek()
                    Case typTokenState.WaitingForOpenCondition

                        If Not clsToken.enuToken = clsRToken.typToken.RNewLine Then
                            If clsToken.strTxt = "(" Then
                                stkTokenState.Pop()
                                stkTokenState.Push(typTokenState.WaitingForCloseCondition)
                            End If
                        End If

                    Case typTokenState.WaitingForCloseCondition

                        If stkNumOpenBrackets.Peek() = 0 Then
                            stkTokenState.Pop()
                            stkTokenState.Push(typTokenState.WaitingForStartScript)
                        End If

                    Case typTokenState.WaitingForStartScript

                        If Not clsToken.enuToken = clsRToken.typToken.RNewLine Then
                            stkTokenState.Pop()
                            stkTokenState.Push(typTokenState.WaitingForEndScript)
                            If clsToken.strTxt = "{" Then
                                stkIsScriptEnclosedByCurlyBrackets.Push(True)  'script will terminate with '}'
                            Else
                                stkIsScriptEnclosedByCurlyBrackets.Push(False) 'script will terminate with end statement
                            End If
                        End If

                    Case typTokenState.WaitingForEndScript

                        If clsToken.enuToken = clsRToken.typToken.RNewLine AndAlso
                                stkNumOpenBrackets.Peek() = 0 AndAlso               'if there are no open brackets
                                Not IsOperatorUserDefined(strLexemePrev) AndAlso    'if line doesn't end in a user-defined operator
                                Not (IsOperatorReserved(strLexemePrev) AndAlso  'if line doesn't end in a predefined operator
                                     Not strLexemePrev = "~") Then                  '    unless it's a tilda (the only operator that doesn't need a right-hand value)
                            clsToken.enuToken = clsRToken.typToken.REndStatement
                        End If

                        If clsToken.enuToken = clsRToken.typToken.REndStatement AndAlso
                                stkIsScriptEnclosedByCurlyBrackets.Peek() = False Then
                            clsToken.enuToken = clsRToken.typToken.REndScript
                        End If

                        If clsToken.enuToken = clsRToken.typToken.REndScript Then
                            stkIsScriptEnclosedByCurlyBrackets.Pop()
                            stkNumOpenBrackets.Pop()
                            stkTokenState.Pop()
                        End If

                    Case Else
                        Throw New Exception("The token is in an unknown state.")
                End Select
            End If

            'add new token to token list
            lstRTokens.Add(clsToken)
        Next intPos

        Return lstRTokens
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   TODO. </summary>
    '''
    ''' <param name="lstRTokens">   The list r tokens. </param>
    '''
    ''' <returns>   The list r statements. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Function GetLstRStatements(lstRTokens As List(Of clsRToken)) As List(Of clsRStatement)
        If lstRTokens Is Nothing OrElse lstRTokens.Count = 0 Then
            Return Nothing
        End If

        Dim lstRStatements = New List(Of clsRStatement)
        Dim intPos As Integer = 0

        While (intPos < lstRTokens.Count)
            lstRStatements.Add(New clsRStatement(lstRTokens, intPos))
        End While

        Return lstRStatements
    End Function

    Public Function GetAsExecutableScript() As String
        Dim strTxt As String = ""
        For Each clsStatement As clsRStatement In lstRStatements
            strTxt &= clsStatement.GetAsExecutableScript()
        Next

        Return strTxt
    End Function
    '''--------------------------------------------------------------------------------------------
    ''' <summary>   TODO. </summary>
    '''
    ''' <returns>   as debug string. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Function GetAsDebugString() As String
        Dim strTxt As String = "Script: " & vbLf
        For Each clsStatement As clsRStatement In lstRStatements
            strTxt &= clsStatement.GetAsDebugString()
        Next

        Return strTxt
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns true if <paramref name="strNew"/> is valid lexeme, else returns false.
    '''             </summary>
    '''
    ''' <param name="strNew">   A sequence of characters from a syntactically correct R script </param>
    '''
    ''' <returns>   True if <paramref name="strNew"/> is a valid lexeme, else returns false. </returns>
    '''--------------------------------------------------------------------------------------------
    Private Function IsValidLexeme(strNew As String) As Boolean

        If (String.IsNullOrEmpty(strNew)) Then
            Return False
        End If

        If Not strNew = vbCrLf AndAlso Regex.IsMatch(strNew, ".+\n$") OrElse 'string is >1 char and ends in newline (useful to do this test first because it simplifies the regular expressions below)
                Regex.IsMatch(strNew, ".+\r$") OrElse                        'string is >1 char and ends in carriage return
                Regex.IsMatch(strNew, "^%.*%.+") OrElse                      'string is a user-defined operator followed by another character
                Regex.IsMatch(strNew, "^'.*'.+") OrElse                      'string is a single quoted string followed by another character
                Regex.IsMatch(strNew, "^"".*"".+") OrElse                    'string is a double quoted string followed by another character
                Regex.IsMatch(strNew, "^`.*`.+") Then                        'string is a backtick quoted string followed by another character
            Return False
        End If

        'if string is a valid ...
        If IsSyntacticName(strNew) OrElse                    'syntactic name or reserved word
                IsOperatorReserved(strNew) OrElse            'operator (e.g. '+')
                IsOperatorBrackets(strNew) OrElse            'bracket operator (e.g. '[')
                strNew = "<<" OrElse                         'partial operator (e.g. ':')
                IsNewLine(strNew) OrElse                     'newlines (e.g. '\n')
                strNew = "," OrElse strNew = ";" OrElse      'parameter separator or end statement
                IsBracket(strNew) OrElse                     'bracket (e.g. '{')
                IsSequenceOfSpaces(strNew) OrElse            'sequence of spaces
                IsConstantString(strNew) OrElse              'string constant (starts with single or double)
                IsOperatorUserDefined(strNew) OrElse         'user-defined operator (starts with '%*')
                IsComment(strNew) Then                       'comment (starts with '#*')
            Return True
        End If

        Return False
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns <paramref name="strLexemeCurrent"/> as a token. <para>
    '''             A token is a string of characters that represent a valid R element, plus meta
    '''             data about the token type (identifier, operator, keyword, bracket etc.).
    '''             </para><para>
    '''             <paramref name="strLexemePrev"/> and <paramref name="strLexemeNext"/> are needed
    '''             to correctly identify if <paramref name="strLexemeCurrent"/> is a unary or 
    '''             binary operator.</para></summary>
    '''
    ''' <param name="strLexemePrev">    The non-space lexeme immediately to the left of 
    '''                                 <paramref name="strLexemeCurrent"/>. </param>
    ''' <param name="strLexemeCurrent"> The lexeme to convert to a token. </param>
    ''' <param name="strLexemeNext">    The non-space lexeme immediately to the right of
    '''                                 <paramref name="strLexemeCurrent"/>. </param>
    '''
    ''' <returns>   <paramref name="strLexemeCurrent"/> as a token. </returns>
    '''--------------------------------------------------------------------------------------------
    Private Function GetRToken(strLexemePrev As String, strLexemeCurrent As String, strLexemeNext As String) As clsRToken
        'TODO refactor so that strLexemePrev and strLexemeNext are booleans rather than strings?
        If String.IsNullOrEmpty(strLexemeCurrent) Then
            Return Nothing
        End If

        Dim clsRTokenNew As New clsRToken With {
            .strTxt = strLexemeCurrent
        }

        If IsKeyWord(strLexemeCurrent) Then                         'reserved key word (e.g. if, else etc.)
            clsRTokenNew.enuToken = clsRToken.typToken.RKeyWord
        ElseIf IsSyntacticName(strLexemeCurrent) Then
            If strLexemeNext = "(" Then
                clsRTokenNew.enuToken = clsRToken.typToken.RFunctionName  'function name
            Else
                clsRTokenNew.enuToken = clsRToken.typToken.RSyntacticName 'syntactic name
            End If
        ElseIf IsComment(strLexemeCurrent) Then                      'comment (starts with '#*')
            clsRTokenNew.enuToken = clsRToken.typToken.RComment
        ElseIf IsConstantString(strLexemeCurrent) Then               'string literal (starts with single or double quote)
            clsRTokenNew.enuToken = clsRToken.typToken.RConstantString
        ElseIf IsNewLine(strLexemeCurrent) Then                      'new line (e.g. '\n')
            clsRTokenNew.enuToken = clsRToken.typToken.RNewLine
        ElseIf strLexemeCurrent = ";" Then                           'end statement
            clsRTokenNew.enuToken = clsRToken.typToken.REndStatement
        ElseIf strLexemeCurrent = "," Then                           'parameter separator
            clsRTokenNew.enuToken = clsRToken.typToken.RSeparator
        ElseIf IsSequenceOfSpaces(strLexemeCurrent) Then             'sequence of spaces (needs to be after separator check, 
            clsRTokenNew.enuToken = clsRToken.typToken.RSpace        '        else linefeed is recognised as space)
        ElseIf IsBracket(strLexemeCurrent) Then                      'bracket (e.g. '{')
            If strLexemeCurrent = "}" Then
                clsRTokenNew.enuToken = clsRToken.typToken.REndScript
            Else
                clsRTokenNew.enuToken = clsRToken.typToken.RBracket
            End If
        ElseIf IsOperatorBrackets(strLexemeCurrent) Then             'bracket operator (e.g. '[')
            clsRTokenNew.enuToken = clsRToken.typToken.ROperatorBracket
        ElseIf IsOperatorUnary(strLexemeCurrent) AndAlso             'unary right operator (e.g. '!x')
                (String.IsNullOrEmpty(strLexemePrev) OrElse
                Not Regex.IsMatch(strLexemePrev, "[a-zA-Z0-9_\.)\]]$")) Then
            clsRTokenNew.enuToken = clsRToken.typToken.ROperatorUnaryRight
        ElseIf strLexemeCurrent = "~" AndAlso                        'unary left operator (e.g. x~)
                (String.IsNullOrEmpty(strLexemeNext) OrElse
                Not Regex.IsMatch(strLexemeNext, "^[a-zA-Z0-9_\.(\+\-\!~]")) Then
            clsRTokenNew.enuToken = clsRToken.typToken.ROperatorUnaryLeft
        ElseIf IsOperatorReserved(strLexemeCurrent) OrElse           'binary operator (e.g. '+')
                Regex.IsMatch(strLexemeCurrent, "^%.*%$") Then
            clsRTokenNew.enuToken = clsRToken.typToken.ROperatorBinary
        Else
            Throw New Exception("Cannot build a valid token from lexeme \'" & strLexemeCurrent & "\'.")
        End If

        Return clsRTokenNew
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
    Private Function IsSyntacticName(strTxt As String) As Boolean
        If String.IsNullOrEmpty(strTxt) Then
            Return False
        End If
        Return Regex.IsMatch(strTxt, "^[a-zA-Z0-9_\.]+$") OrElse
               Regex.IsMatch(strTxt, "^`.*")
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns true if <paramref name="strTxt"/> is a complete or partial string 
    '''             constant, else returns false.<para>
    '''             String constants are delimited by a pair of single (‘'’) or double (‘"’) quotes 
    '''             and can contain all other printable characters. Quotes and other special 
    '''             characters within strings are specified using escape sequences. </para></summary>
    '''
    ''' <param name="strTxt">   The text to check. </param>
    '''
    ''' <returns>   True if <paramref name="strTxt"/> is a complete or partial string constant,
    '''             else returns false.</returns>
    '''--------------------------------------------------------------------------------------------
    Private Function IsConstantString(strTxt As String) As Boolean
        If Not String.IsNullOrEmpty(strTxt) AndAlso
            (Regex.IsMatch(strTxt, "^"".*") OrElse (Regex.IsMatch(strTxt, "^'.*"))) Then
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
    Private Function IsComment(strTxt As String) As Boolean
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
    Private Function IsSequenceOfSpaces(strTxt As String) As Boolean
        If Not String.IsNullOrEmpty(strTxt) AndAlso
                Not strTxt = vbLf AndAlso
                Regex.IsMatch(strTxt, "^ *$") Then
            Return True
        End If
        Return False
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns true if <paramref name="strTxt"/> is a functional R element 
    '''             (i.e. not a space, comment or new line), else returns false. </summary>
    '''
    ''' <param name="strTxt">   The text to check . </param>
    '''
    ''' <returns>   True  if <paramref name="strTxt"/> is a functional R element
    '''             (i.e. not a space, comment or new line), else returns false. </returns>
    '''--------------------------------------------------------------------------------------------
    Private Function IsElement(strTxt As String) As Boolean
        If Not IsNewLine(strTxt) AndAlso
           Not IsSequenceOfSpaces(strTxt) AndAlso
           Not IsComment(strTxt) Then
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
    Private Function IsOperatorUserDefined(strTxt As String) As Boolean
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
    Private Function IsOperatorReserved(strTxt As String) As Boolean
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
    Private Function IsOperatorBrackets(strTxt As String) As Boolean
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
    Private Function IsOperatorUnary(strTxt As String) As Boolean
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
    Private Function IsBracket(strTxt As String) As Boolean
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
    Private Function IsNewLine(strTxt As String) As Boolean
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
    Private Function IsKeyWord(strTxt As String) As Boolean
        Dim arrKeyWords() As String = {"if", "else", "repeat", "while", "function", "for", "in", "next", "break"}
        Return arrKeyWords.Contains(strTxt)
    End Function


End Class
