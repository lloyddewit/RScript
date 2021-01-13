Public Class clsRStatement

    ''' <summary>
    ''' If true, then when this R statement is converted to a script, then it will be 
    '''             terminated with a newline (else if false then a semicolon)
    ''' </summary>
    Public bTerminateWithNewline As Boolean = True

    Public clsAssignment As clsRElement
    Public clsElement As clsRElement

    Private strDebug As String = ""

    Private ReadOnly arrOperatorPrecedence(17)() As String
    Private ReadOnly intOperatorsBrackets As Integer = 2
    Private ReadOnly intOperatorsUnaryOnly As Integer = 4
    Private ReadOnly intOperatorsUserDefined As Integer = 6
    Private ReadOnly intOperatorsTilda As Integer = 13


    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Constructor.
    '''             Creates an object representing a valid R statement.
    '''             processes the tokens from <paramref name="lstTokens"/> from position <paramref name="intPos"/> to the end of statement, end of script or end of list (whichever comes first)
    '''             TODO </summary>
    '''
    ''' <param name="lstTokens">   The list r tokens. </param>
    ''' <param name="intPos">       [in,out] The int position. </param>
    '''--------------------------------------------------------------------------------------------
    Public Sub New(lstTokens As List(Of clsRToken), ByRef intPos As Integer)

        arrOperatorPrecedence(0) = New String() {"::", ":::"}
        arrOperatorPrecedence(1) = New String() {"$", "@"}
        arrOperatorPrecedence(intOperatorsBrackets) = New String() {"[", "[["} 'bracket operators
        arrOperatorPrecedence(3) = New String() {"^"}                          'right to left precedence
        arrOperatorPrecedence(intOperatorsUnaryOnly) = New String() {"-", "+"} 'unary operarors
        arrOperatorPrecedence(5) = New String() {":"}
        arrOperatorPrecedence(intOperatorsUserDefined) = New String() {"%"}    'any operator that starts with '%' (including user-defined operators)
        arrOperatorPrecedence(7) = New String() {"*", "/"}
        arrOperatorPrecedence(8) = New String() {"+", "-"}
        arrOperatorPrecedence(9) = New String() {"<>", "<=", ">=", "==", "!="}
        arrOperatorPrecedence(10) = New String() {"!"}
        arrOperatorPrecedence(11) = New String() {"&", "&&"}
        arrOperatorPrecedence(12) = New String() {"|", "||"}
        arrOperatorPrecedence(intOperatorsTilda) = New String() {"~"}          'unary or binary
        arrOperatorPrecedence(14) = New String() {"->", "->>"}
        arrOperatorPrecedence(15) = New String() {"<-", "<<-"}
        arrOperatorPrecedence(16) = New String() {"="}

        'create list of tokens for this statement
        Dim lstStatementTokens As List(Of clsRToken) = New List(Of clsRToken)
        While (intPos < lstTokens.Count)
            lstStatementTokens.Add(lstTokens.Item(intPos))
            If lstTokens.Item(intPos).enuToken = clsRToken.typToken.REndStatement OrElse 'we don't add this termination condition to the while loop 
                 lstTokens.Item(intPos).enuToken = clsRToken.typToken.REndScript Then    '    because we also want the token that terminates the statement 
                intPos += 1                                                              '    to be part of the statement's list of tokens
                Exit While
            End If
            intPos += 1
        End While

        Dim lstTokenBrackets As List(Of clsRToken) = GetLstTokenBrackets(lstStatementTokens, 0)
        Dim lstTokenFunctionBrackets As List(Of clsRToken) = GetLstTokenFunctionBrackets(lstTokenBrackets)
        Dim lstTokenCommas As List(Of clsRToken) = GetLstTokenCommas(lstTokenFunctionBrackets, 0)
        Dim lstTokenOperators As List(Of clsRToken) = GetLstTokenOperators(lstTokenCommas)
        strDebug = GetLstTokensAsString(lstTokenOperators)

    End Sub

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   TODO. </summary>
    '''
    ''' <returns>   as debug string. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Function GetAsDebugString() As String
        Return strDebug
        'Return "Statement: " & vbLf &
        '        "bTerminateWithNewline: " & bTerminateWithNewline & vbLf &
        '        "Assignment: " & clsAssignment.GetAsDebugString() & vbLf &
        '        "Element: " & clsElement.GetAsDebugString() & vbLf
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   TODO
    '''             Iterates through the tokens in <paramref name="lstStatementTokens"/>.
    '''             If the token is a '(' then it makes everything inside the brackets a child of the '(' token.
    '''             if the '(' belongs to a function then makes the '(' a child of the function
    '''             Brackets may be nested. For example:
    '''             'myFunction(a*(b+c))' is structured as:
    '''             myFunction
    '''             ..(
    '''             ....a
    '''             ....*
    '''             ....(
    '''             ......b
    '''             ......+
    '''             ......c
    '''             ......)
    '''             ....)
    '''              </summary>
    '''
    ''' <param name="lstStatementTokens">   The list statement tokens. </param>
    '''
    ''' <returns>   The list token brackets. </returns>
    '''--------------------------------------------------------------------------------------------
    Private Function GetLstTokenBrackets(lstStatementTokens As List(Of clsRToken), ByRef intPos As Integer) As List(Of clsRToken)
        Dim lstTokens As List(Of clsRToken) = New List(Of clsRToken)
        Dim clsToken As clsRToken

        While (intPos < lstStatementTokens.Count)
            clsToken = lstStatementTokens.Item(intPos)
            intPos += 1
            Select Case clsToken.strTxt
                Case "("
                    clsToken.lstTokens = GetLstTokenBrackets(lstStatementTokens, intPos)
                Case ")"
                    lstTokens.Add(clsToken)
                    Return lstTokens
            End Select
            lstTokens.Add(clsToken)
        End While
        Return lstTokens
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>
    '''     TODO Traverses the tree of tokens in <paramref name="lstTokens"/>. If the token is a ','
    '''     then it makes everything up to the next ',' or ')' a child of the ',' token. Commas are
    '''     used to separate function parameters. Parameters between commas are optional. For
    '''     example: 'myFunction(a,,b)' is structured as: myFunction (
    '''     ..a
    '''     ..,
    '''     ..,
    '''     ....b
    '''     ....)
    ''' </summary>
    '''
    ''' <param name="lstTokens">    The list statement tokens. </param>
    '''
    ''' <returns>   The list token brackets. </returns> 
    '''--------------------------------------------------------------------------------------------
    Private Function GetLstTokenFunctionBrackets(lstTokens As List(Of clsRToken)) As List(Of clsRToken)
        'if nothing to process then return empty list
        If lstTokens.Count <= 0 Then
            Return New List(Of clsRToken)
        End If

        Dim lstTokensNew As List(Of clsRToken) = New List(Of clsRToken)
        Dim clsToken As clsRToken
        Dim intPos As Integer = 0

        While (intPos < lstTokens.Count)
            clsToken = lstTokens.Item(intPos)

            If clsToken.enuToken = clsRToken.typToken.RFunctionName Then
                'if function name already has children, or next steps will go out of bounds, then throw developer error
                If clsToken.lstTokens.Count > 0 OrElse
                        intPos >= lstTokens.Count - 1 OrElse
                        (lstTokens.Item(intPos + 1).enuToken = clsRToken.typToken.RSpace AndAlso
                        intPos >= lstTokens.Count - 2) Then
                    'TODO throw developer error
                    Exit While
                End If
                'make the function's open bracket (and any preceeding spaces) a child of the function name
                intPos += 1
                Dim clsTokenNext = lstTokens.Item(intPos)
                clsToken.lstTokens.Add(clsTokenNext.CloneMe)
                If clsTokenNext.enuToken = clsRToken.typToken.RSpace Then
                    intPos += 1
                    clsTokenNext = lstTokens.Item(intPos)
                    clsToken.lstTokens.Add(clsTokenNext.CloneMe)
                End If
            End If
            clsToken.lstTokens = GetLstTokenFunctionBrackets(clsToken.CloneMe.lstTokens)
            lstTokensNew.Add(clsToken.CloneMe)
            intPos += 1
        End While
        Return lstTokensNew
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   TODO
    '''             Traverses the tree of tokens in <paramref name="lstStatementTokens"/>.
    '''             If the token is a ',' then it makes everything up to the next ',' or ')' a child of the ',' token.
    '''             Commas are used to separate function parameters. Parameters between commas are optional. For example:
    '''             'myFunction(a,,b)' is structured as:
    '''             myFunction
    '''             (
    '''             ..a
    '''             ..,
    '''             ..,
    '''             ....b
    '''             ....)
    '''              </summary>
    '''
    ''' <param name="lstStatementTokens">   The list statement tokens. </param>
    '''
    ''' <returns>   The list token brackets. </returns>
    '''--------------------------------------------------------------------------------------------
    Private Function GetLstTokenCommas(lstStatementTokens As List(Of clsRToken),
                                       ByRef intPos As Integer,
                                       Optional bProcessingComma As Boolean = False) As List(Of clsRToken)
        Dim lstTokens As List(Of clsRToken) = New List(Of clsRToken)
        Dim clsToken As clsRToken

        While (intPos < lstStatementTokens.Count)
            clsToken = lstStatementTokens.Item(intPos)
            Select Case clsToken.strTxt
                Case ","
                    If bProcessingComma Then
                        intPos -= 1  'ensure this comma is processed in the level above
                        Return lstTokens
                    Else
                        intPos += 1
                        clsToken.lstTokens = GetLstTokenCommas(lstStatementTokens, intPos, True)
                    End If
                Case ")"
                    lstTokens.Add(clsToken)
                    Return lstTokens
                Case Else
                    If lstStatementTokens.Item(intPos).lstTokens.Count > 0 Then
                        clsToken.lstTokens = GetLstTokenCommas(lstStatementTokens.Item(intPos).lstTokens, 0)
                    End If
            End Select
            lstTokens.Add(clsToken)
            intPos += 1
        End While
        Return lstTokens
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   TODO
    '''             Iterates through all the possible operators in order of precedence.
    '''             For each operator, traverses the tree of tokens in <paramref name="lstTokens"/>.
    '''             If the operator is found then the operator's parameters (typically the tokens to the left and right of the operator) are made children of the operator.
    '''             For example:
    '''             'a*b+c' is structured as:
    '''             +
    '''             ..*
    '''             ....a
    '''             ....b
    '''             ..c
    '''             </summary>
    '''
    ''' <param name="lstTokens">   The list statement tokens. </param>
    '''
    ''' <returns>   The list token brackets. </returns>
    '''--------------------------------------------------------------------------------------------
    Private Function GetLstTokenOperators(lstTokens As List(Of clsRToken)) As List(Of clsRToken)
        'if nothing to process then return empty list
        If lstTokens.Count <= 0 Then
            Return New List(Of clsRToken)
        End If

        Dim lstTokensNew As List(Of clsRToken) = New List(Of clsRToken)
        For intPosOperators As Integer = 0 To UBound(arrOperatorPrecedence) - 1

            'restructure the tree for the next group of operators in the precedence list
            lstTokensNew = GetLstTokenOperatorGroup(lstTokens, intPosOperators)

            'clone the new tree before restructuring for the next operator
            lstTokens = New List(Of clsRToken)
            For Each clsTokenTmp As clsRToken In lstTokensNew
                lstTokens.Add(clsTokenTmp.CloneMe)
            Next
        Next

        Return lstTokensNew
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>
    '''     TODO
    '''     Traverses the tree of tokens in <paramref name="lstTokens"/> looking for occurences an operator in operator group <paramref name="intPosOperators"/>. 
    '''     If an operator is found then the operator's parameters (typically the tokens to the left and right of the
    '''     operator) are made children of the operator. For example: 'a+b' is structured as:
    '''     +
    '''     ..a
    '''     ..b
    '''     
    ''' </summary>
    '''
    ''' <param name="lstTokens">        The list statement tokens. </param>
    ''' <param name="intPosOperators">  The int position operators. </param>
    '''
    ''' <returns>   The list token brackets. </returns>
    '''--------------------------------------------------------------------------------------------
    Private Function GetLstTokenOperatorGroup(lstTokens As List(Of clsRToken), intPosOperators As Integer)
        'TODO process spaces, comments and newlines

        'if nothing to process then return empty list
        If lstTokens.Count <= 0 Then
            Return New List(Of clsRToken)
        End If

        Dim lstTokensNew As List(Of clsRToken) = New List(Of clsRToken)
        Dim clsToken As clsRToken
        Dim clsTokenPrev As clsRToken = Nothing
        Dim bPrevTokenProcessed As Boolean = False

        Dim intPosTokens As Integer = 0
        While (intPosTokens < lstTokens.Count)
            clsToken = lstTokens.Item(intPosTokens).CloneMe
            Select Case intPosOperators
                Case intOperatorsBrackets
                Case intOperatorsUnaryOnly
                Case intOperatorsUserDefined
                Case intOperatorsTilda
                Case Else
                    'if token is the operator we're looking for
                    If clsToken.enuToken = clsRToken.typToken.ROperatorBinary AndAlso
                            arrOperatorPrecedence(intPosOperators).Contains(clsToken.strTxt) Then
                        If clsToken.lstTokens.Count > 0 OrElse
                                IsNothing(clsTokenPrev) OrElse intPosTokens >= (lstTokens.Count - 1) Then
                            'TODO throw developer error: binary operator already has children or is missing a parameter
                            Exit While
                        End If
                        intPosTokens += 1
                        'process the next token's children
                        Dim clsTokenNext As clsRToken = lstTokens.Item(intPosTokens).CloneMe
                        clsTokenNext.lstTokens = GetLstTokenOperatorGroup(clsTokenNext.CloneMe.lstTokens, intPosOperators)

                        'make the previous and next tokens, the children of the current token
                        clsToken.lstTokens.Add(clsTokenPrev.CloneMe)
                        clsToken.lstTokens.Add(clsTokenNext.CloneMe)
                        bPrevTokenProcessed = True
                    End If
            End Select

            'if token was not the operator we were looking for
            If Not bPrevTokenProcessed Then
                'add the previous token to the tree
                If Not IsNothing(clsTokenPrev) Then
                    lstTokensNew.Add(clsTokenPrev)
                End If
                'process the current token's children
                clsToken.lstTokens = GetLstTokenOperatorGroup(clsToken.CloneMe.lstTokens, intPosOperators)
            End If

            clsTokenPrev = clsToken.CloneMe
            bPrevTokenProcessed = False
            intPosTokens += 1
        End While

        'if there is still an outstanding token that is not yet added to the new list, then add it
        If IsNothing(clsTokenPrev) Then
            'TODO throw developer error: There should always be at least one token still to add to the tree
            Return Nothing
        End If
        lstTokensNew.Add(clsTokenPrev.CloneMe)

        Return lstTokensNew
    End Function

    Private Function GetLstTokensAsString(lstTokens As List(Of clsRToken), Optional strIndent As String = "") As String
        If lstTokens Is Nothing OrElse lstTokens.Count = 0 Then
            'TODO throw exception
            Return Nothing
        End If
        Dim strTxt As String = strIndent

        For Each clsToken In lstTokens

            strTxt &= clsToken.strTxt & "("
            Select Case clsToken.enuToken
                Case RScript.clsRToken.typToken.RSyntacticName
                    strTxt &= "RSyntacticName"
                Case RScript.clsRToken.typToken.RKeyWord
                    strTxt &= "RKeyWord"
                Case RScript.clsRToken.typToken.RConstantString
                    strTxt &= "RStringLiteral"
                Case RScript.clsRToken.typToken.RComment
                    strTxt &= "RComment"
                Case RScript.clsRToken.typToken.RSpace
                    strTxt &= "RSpace"
                Case RScript.clsRToken.typToken.RBracket
                    strTxt &= "RBracket"
                Case RScript.clsRToken.typToken.RSeparator
                    strTxt &= "RSeparator"
                Case RScript.clsRToken.typToken.RNewLine
                    strTxt &= "RNewLine"
                Case RScript.clsRToken.typToken.REndStatement
                    strTxt &= "REndStatement"
                Case RScript.clsRToken.typToken.REndScript
                    strTxt &= "REndScript"
                Case RScript.clsRToken.typToken.ROperatorUnaryLeft
                    strTxt &= "ROperatorUnaryLeft"
                Case RScript.clsRToken.typToken.ROperatorUnaryRight
                    strTxt &= "ROperatorUnaryRight"
                Case RScript.clsRToken.typToken.ROperatorBinary
                    strTxt &= "ROperatorBinary"
                Case RScript.clsRToken.typToken.ROperatorBracket
                    strTxt &= "ROperatorBracket"
            End Select
            strTxt &= "), "

            If clsToken.lstTokens.Count > 0 Then
                strTxt &= vbLf & GetLstTokensAsString(clsToken.lstTokens, strIndent & "..") & vbLf & strIndent
            End If
        Next
        Return strTxt
    End Function

End Class
