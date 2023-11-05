Imports System.Collections.Specialized

'TODO Should we model constants differently to syntactic names? (there are five types of constants: integer, logical, numeric, complex and string)
'TODO Test special constants {"NULL", "NA", "Inf", "NaN"}
'TODO Test function names as string constants. E.g 'x + y can equivalently be written "+"(x, y). Notice that since '+' is a non-standard function name, it needs to be quoted (see https://cran.r-project.org/doc/manuals/r-release/R-lang.html#Writing-functions)'
'TODO handle '...' (used in function definition)
'TODO handle '.' normally just part of a syntactic name, but has a special meaning when in a function name, or when referring to data (represents no variable)
'TODO is it possible for packages to be nested (e.g. 'p1::p1_1::f()')?
'TODO currently all newlines (vbLf, vbCr and vbCrLf) are converted to vbLf. Is it important to remember what the original new line character was?
'TODO convert public data members to properties (all classes)
'TODO provide an option to get script with automatic indenting (specifiy num spaces for indent and max num Columns per line)
' 
' 17/11/20
' - allow named operator params (R-Instat allows operator params to be named, but this infor is lost in script)
'
' 01/03/21
' - how should bracket operator separators be modelled?
'   strInput = "df[1:2,]"
'   strInput = "df[,1:2]"
'   strInput = "df[1:2,1:2]"
'   strInput = "df[1:2,""x""]"
'   strInput = "df[1:2,c(""x"",""y"")]"
'

''' <summary>   TODO Add class summary. </summary>
Public Class clsRScript

    ''' <summary>   
    ''' The R statements in the script. The dictionary key is the start position of the statement 
    ''' in the script. The dictionary value is the statement itself. </summary>
    Public dctRStatements As New OrderedDictionary()

    ''' <summary>   The current state of the token parsing. </summary>
    Private Enum typTokenState
        WaitingForOpenCondition
        WaitingForCloseCondition
        WaitingForStartScript
        WaitingForEndScript
    End Enum

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Parses the R script in <paramref name="strInput"/> and populates the distionary
    '''             of R statements.
    '''             <para>
    '''             This subroutine will accept, and correctly process all valid R. However, this 
    '''             class does not attempt to validate <paramref name="strInput"/>. If it is not 
    '''             valid R then this subroutine may still process the script without throwing an 
    '''             exception. In this case, the list of R statements will be undefined.
    '''             </para><para>
    '''             In other words, this subroutine should not generate false negatives (reject 
    '''             valid R) but may generate false positives (accept invalid R).
    '''             </para></summary>
    '''
    ''' <param name="strInput"> The R script to parse. This must be valid R according to the 
    '''                         R language specification at 
    '''                         https://cran.r-project.org/doc/manuals/r-release/R-lang.html 
    '''                         (referenced 01 Feb 2021).</param>
    '''--------------------------------------------------------------------------------------------
    Public Sub New(strInput As String)
        If String.IsNullOrEmpty(strInput) Then
            Exit Sub
        End If

        Dim lstLexemes As List(Of String) = GetLstLexemes(strInput)
        Dim lstTokens As List(Of clsRToken) = GetLstTokens(lstLexemes)

        Dim iPos As Integer = 0
        Dim dctAssignments As New Dictionary(Of String, clsRStatement)
        While iPos < lstTokens.Count
            Dim iScriptPos As UInteger = lstTokens.Item(iPos).iScriptPos
            Dim clsStatement As clsRStatement = New clsRStatement(lstTokens, iPos, dctAssignments)
            dctRStatements.Add(iScriptPos, clsStatement)

            'if the value of an assigned element is new/updated
            If Not IsNothing(clsStatement.clsAssignment) Then
                'store the updated/new definition in the dictionary
                If dctAssignments.ContainsKey(clsStatement.clsAssignment.strTxt) Then
                    dctAssignments(clsStatement.clsAssignment.strTxt) = clsStatement
                Else
                    dctAssignments.Add(clsStatement.clsAssignment.strTxt, clsStatement)
                End If
            End If

        End While
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
        Dim strTxt As String = Nothing
        Dim stkIsSingleBracket As Stack(Of Boolean) = New Stack(Of Boolean)

        For Each chrNew As Char In strRScript
            'we keep adding characters to the lexeme, one at a time, until we reach a character that 
            '    would make the lexeme invalid
            If clsRToken.IsValidLexeme(strTxt & chrNew) AndAlso
                    Not ((strTxt & chrNew) = "]]" AndAlso 'edge case for nested operator brackets (see note below)
                    (stkIsSingleBracket.Count < 1 OrElse stkIsSingleBracket.Peek())) Then
                strTxt &= chrNew
                Continue For
            End If

            'Edge case: We need to handle nested operator brackets e.g. 'k[[l[[m[6]]]]]'. 
            '           For the above example, we need to recognise that the ']' to the right 
            '           of '6' is a single ']' bracket and is not part of a double ']]' bracket.
            '           To achieve this, we push each open bracket to a stack so that we know 
            '           which type of closing bracket is expected for each open bracket.
            Select Case strTxt
                Case "["
                    stkIsSingleBracket.Push(True)
                Case "[["
                    stkIsSingleBracket.Push(False)
                Case "]", "]]"
                    If stkIsSingleBracket.Count < 1 Then
                        Throw New Exception("Closing bracket detected ('" & strTxt & "') with no corresponding open bracket.")
                    End If
                    stkIsSingleBracket.Pop()
            End Select

            'adding the new char to the lexeme would make the lexeme invalid, 
            '       so we add the existing lexeme to the list and start a new lexeme
            lstLexemes.Add(strTxt)
            strTxt = chrNew
        Next

        lstLexemes.Add(strTxt)
        Return lstLexemes
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns <paramref name="lstLexemes"/> as a list of tokens.
    '''             <para>
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
        Dim bLexemePrevOnSameLine As Boolean = False
        Dim bLexemeNextOnSameLine As Boolean
        Dim bStatementContainsElement As Boolean = False
        Dim clsToken As clsRToken

        Dim stkNumOpenBrackets As Stack(Of Integer) = New Stack(Of Integer)
        stkNumOpenBrackets.Push(0)

        Dim stkIsScriptEnclosedByCurlyBrackets As Stack(Of Boolean) = New Stack(Of Boolean)
        stkIsScriptEnclosedByCurlyBrackets.Push(True)

        Dim stkTokenState As Stack(Of typTokenState) = New Stack(Of typTokenState)
        stkTokenState.Push(typTokenState.WaitingForStartScript)

        Dim iScriptPos As UInteger = 0

        For iPos As Integer = 0 To lstLexemes.Count - 1

            If stkNumOpenBrackets.Count < 1 Then
                Throw New Exception("The stack storing the number of open brackets must have at least one value.")
            ElseIf stkIsScriptEnclosedByCurlyBrackets.Count < 1 Then
                Throw New Exception("The stack storing the number of open curly brackets must have at least one value.")
            ElseIf stkTokenState.Count < 1 Then
                Throw New Exception("The stack storing the current state of the token parsing must have at least one value.")
            End If

            'store previous non-space lexeme
            If clsRToken.IsElement(strLexemeCurrent) Then
                strLexemePrev = strLexemeCurrent
                bLexemePrevOnSameLine = True
            ElseIf clsRToken.IsNewLine(strLexemeCurrent) Then
                bLexemePrevOnSameLine = False
            End If

            strLexemeCurrent = lstLexemes.Item(iPos)
            bStatementContainsElement = If(bStatementContainsElement, bStatementContainsElement, clsRToken.IsElement(strLexemeCurrent))

            'find next lexeme that represents an R element
            strLexemeNext = Nothing
            bLexemeNextOnSameLine = True
            For iNextPos As Integer = iPos + 1 To lstLexemes.Count - 1
                If clsRToken.IsElement(lstLexemes.Item(iNextPos)) Then
                    strLexemeNext = lstLexemes.Item(iNextPos)
                    Exit For
                End If
                Select Case lstLexemes.Item(iNextPos)
                    Case vbLf, vbCr, vbCr
                        bLexemeNextOnSameLine = False
                End Select
            Next iNextPos

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
            clsToken = New clsRToken(strLexemePrev, strLexemeCurrent, strLexemeNext, bLexemePrevOnSameLine, bLexemeNextOnSameLine, iScriptPos)
            iScriptPos += strLexemeCurrent.Length

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

                        If Not (clsToken.enuToken = clsRToken.typToken.RComment OrElse
                                clsToken.enuToken = clsRToken.typToken.RPresentation OrElse
                                clsToken.enuToken = clsRToken.typToken.RSpace OrElse
                                clsToken.enuToken = clsRToken.typToken.RNewLine) Then
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
                                bStatementContainsElement AndAlso                   'if statement contains at least one R element (i.e. not just spaces, comments, or newlines)
                                stkNumOpenBrackets.Peek() = 0 AndAlso               'if there are no open brackets
                                Not clsRToken.IsOperatorUserDefined(strLexemePrev) AndAlso    'if line doesn't end in a user-defined operator
                                Not (clsRToken.IsOperatorReserved(strLexemePrev) AndAlso      'if line doesn't end in a predefined operator
                                     Not strLexemePrev = "~") Then                  '    unless it's a tilda (the only operator that doesn't need a right-hand value)
                            clsToken.enuToken = clsRToken.typToken.REndStatement
                            bStatementContainsElement = False
                        End If

                        If clsToken.enuToken = clsRToken.typToken.REndStatement AndAlso
                                stkIsScriptEnclosedByCurlyBrackets.Peek() = False AndAlso
                                String.IsNullOrEmpty(strLexemeNext) Then
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

            'Edge case: if the script has ended and there are no more R elements to process, 
            'then ensure that only formatting lexemes (i.e. spaces, newlines or comments) follow
            'the script's final statement.
            If clsToken.enuToken = clsRToken.typToken.REndScript AndAlso
                    String.IsNullOrEmpty(strLexemeNext) Then

                For iNextPos As Integer = iPos + 1 To lstLexemes.Count - 1

                    strLexemeCurrent = lstLexemes.Item(iNextPos)

                    clsToken = New clsRToken("", strLexemeCurrent, "", False, False, iScriptPos)
                    iScriptPos += strLexemeCurrent.Length

                    Select Case clsToken.enuToken
                        Case clsRToken.typToken.RSpace, clsRToken.typToken.RNewLine, clsRToken.typToken.RComment
                        Case Else
                            Throw New Exception("Only spaces, newlines and comments are allowed after the script ends.")
                    End Select

                    'add new token to token list
                    lstRTokens.Add(clsToken)

                Next iNextPos
                Return lstRTokens
            End If
        Next iPos

        Return lstRTokens
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns this object as a valid, executable R script. </summary>
    '''
    ''' <param name="bIncludeFormatting">   If True, then include all formatting information in 
    '''     returned string (comments, indents, padding spaces, extra line breaks etc.). </param>
    '''
    ''' <returns>   The current state of this object as a valid, executable R script. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Function GetAsExecutableScript(Optional bIncludeFormatting As Boolean = True) As String
        Dim strTxt As String = ""
        For Each clsStatement In dctRStatements
            strTxt &= clsStatement.Value.GetAsExecutableScript(bIncludeFormatting) & If(bIncludeFormatting, "", vbLf)
        Next
        Return strTxt
    End Function

End Class
