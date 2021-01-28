Imports System.Text.RegularExpressions

Public Class clsRStatement

    ''' <summary>
    ''' If true, then when this R statement is converted to a script, then it will be 
    '''             terminated with a newline (else if false then a semicolon)
    ''' </summary>
    Public bTerminateWithNewline As Boolean = True

    Public strAssignmentOperator As String
    Public clsAssignment As clsRElement = Nothing
    Public clsElement As clsRElement

    Private strDebug As String = ""

    Private ReadOnly arrOperatorPrecedence(17)() As String
    Private ReadOnly intOperatorsBrackets As Integer = 2
    Private ReadOnly intOperatorsUnaryOnly As Integer = 4
    Private ReadOnly intOperatorsUserDefined As Integer = 6
    Private ReadOnly intOperatorsTilda As Integer = 13
    Private ReadOnly intOperatorsRightAssignment As Integer = 14
    Private ReadOnly intOperatorsLeftAssignment1 As Integer = 15
    Private ReadOnly intOperatorsLeftAssignment2 As Integer = 16


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

        'if nothing to process then exit
        If lstTokens.Count <= 0 Then
            Exit Sub
        End If

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
        arrOperatorPrecedence(intOperatorsRightAssignment) = New String() {"->", "->>"}
        arrOperatorPrecedence(intOperatorsLeftAssignment1) = New String() {"<-", "<<-"}
        arrOperatorPrecedence(intOperatorsLeftAssignment2) = New String() {"="}

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

        'restructure the list into a token tree
        Dim lstTokenBrackets As List(Of clsRToken) = GetLstTokenBrackets(lstStatementTokens, 0)
        Dim lstTokenFunctionBrackets As List(Of clsRToken) = GetLstTokenFunctionBrackets(lstTokenBrackets)
        Dim lstTokenCommas As List(Of clsRToken) = GetLstTokenCommas(lstTokenFunctionBrackets, 0)
        Dim lstTokenTree As List(Of clsRToken) = GetLstTokenOperators(lstTokenCommas)
        strDebug = GetLstTokensAsString(lstTokenTree) 'TODO just for debug, remove

        'if the tree does not include at least one token, then raise development error
        If lstTokenTree.Count < 1 Then
            Exit Sub 'TODO development error: token tree must contain at least one token
        End If

        'if the statement includes an assignment, construct assignment element
        If lstTokenTree.Item(0).enuToken = clsRToken.typToken.ROperatorBinary AndAlso
                lstTokenTree.Item(0).lstTokens.Count >= 2 Then
            'if the statement has a left assignment (e.g. 'x<-value', 'x<<-value' or 'x=value')
            If arrOperatorPrecedence(intOperatorsLeftAssignment1).Contains(lstTokenTree.Item(0).strTxt) OrElse
                arrOperatorPrecedence(intOperatorsLeftAssignment2).Contains(lstTokenTree.Item(0).strTxt) Then
                strAssignmentOperator = lstTokenTree.Item(0).strTxt
                clsAssignment = GetRElement(lstTokenTree.Item(0).lstTokens.Item(0))
                clsElement = GetRElement(lstTokenTree.Item(0).lstTokens.Item(1))
            ElseIf arrOperatorPrecedence(intOperatorsRightAssignment).Contains(lstTokenBrackets.Item(0).strTxt) Then
                'else if the statement has a right assignment (e.g. 'value->x' or 'value->>x')
                strAssignmentOperator = lstTokenTree.Item(0).strTxt
                clsAssignment = GetRElement(lstTokenTree.Item(0).lstTokens.Item(1))
                clsElement = GetRElement(lstTokenTree.Item(0).lstTokens.Item(0))
            End If
        End If

        'if there's no assignment, then build the main element from the token tree's top element
        If IsNothing(clsAssignment) Then
            clsElement = GetRElement(lstTokenTree.Item(0))
        End If

        'check if the statement is terminated with a semi-colon
        If lstTokenTree.Item(lstTokenTree.Count - 1).enuToken = clsRToken.typToken.REndStatement AndAlso
             lstTokenTree.Item(lstTokenTree.Count - 1).strTxt = ";" Then
            bTerminateWithNewline = False
        End If

    End Sub

    Public Function GetAsExecutableScript() As String
        Dim strScript As String
        Dim strElement As String = GetScriptElement(clsElement)
        'if there is no assignment, then just process the statement's element
        If IsNothing(clsAssignment) OrElse String.IsNullOrEmpty(strAssignmentOperator) Then
            strScript = strElement
        Else 'else if the statement has an assignment
            Dim strAssigment As String = GetScriptElement(clsAssignment)
            'if the statement has a left assignment (e.g. 'x<-value', 'x<<-value' or 'x=value')
            If arrOperatorPrecedence(intOperatorsLeftAssignment1).Contains(strAssignmentOperator) OrElse
                arrOperatorPrecedence(intOperatorsLeftAssignment2).Contains(strAssignmentOperator) Then
                strScript = strAssigment & strAssignmentOperator & strElement
            ElseIf arrOperatorPrecedence(intOperatorsRightAssignment).Contains(strAssignmentOperator) Then
                'else if the statement has a right assignment (e.g. 'value->x' or 'value->>x')
                strScript = strElement & strAssignmentOperator & strAssigment
            Else
                'TODO development error: unexpected assignment operator
                Return Nothing
            End If
        End If

        strScript &= If(bTerminateWithNewline, vbLf, ";")
        Return strScript
    End Function

    Private Function GetScriptElement(clsElement As Object) As String

        'Public strTxt As String
        'Public bBracketed As Boolean
        'Public clsPresentation As clsRElementPresentation

        If IsNothing(clsElement) Then
            Return Nothing
        End If

        Dim strScript As String = If(clsElement.bBracketed, "(", "")

        Select Case clsElement.GetType()
            Case GetType(clsRElementFunction)
                'Public lstParameters As New List(Of clsRParameterNamed)

                strScript &= GetScriptElementProperty(clsElement)
                strScript &= "("
                If Not IsNothing(clsElement.lstParameters) Then
                    Dim bPrefixComma As Boolean = False
                    For Each clsRParameter In clsElement.lstParameters
                        strScript &= If(bPrefixComma, ",", "")
                        bPrefixComma = True
                        strScript &= If(String.IsNullOrEmpty(clsRParameter.strArgumentName), "", clsRParameter.strArgumentName + "=")
                        strScript &= GetScriptElement(clsRParameter.clsArgValue)
                    Next
                End If
                strScript &= ")"
            Case GetType(clsRElementProperty)
                'Public strPackageName As String = "" 'only used for functions and variables (e.g. 'constants::syms$h')
                'Public lstObjects As New List(Of clsRElement) 'only used for functions and variables (e.g. 'constants::syms$h')

                strScript &= GetScriptElementProperty(clsElement)
            Case GetType(clsRElementOperator)
                'Public bFirstParamOnRight As Boolean = False
                'Public strTerminator As String = "" 'only used for '[' and '[[' operators
                'Public lstRParameters As New List(Of clsRParameter)

                Dim bPrefixOperator As Boolean = clsElement.bFirstParamOnRight
                For Each clsRParameter In clsElement.lstParameters
                    strScript &= If(bPrefixOperator, clsElement.strTxt, "")
                    bPrefixOperator = True
                    strScript &= GetScriptElement(clsRParameter.clsArgValue)
                Next
                strScript &= If(clsElement.lstParameters.Count = 1 AndAlso Not clsElement.bFirstParamOnRight, clsElement.strTxt, "")
            Case GetType(clsRElementKeyWord)
            Case GetType(clsRElement)
                strScript &= clsElement.strTxt
        End Select
        strScript &= If(clsElement.bBracketed, ")", "")
        Return strScript
    End Function

    Private Function GetScriptElementProperty(clsElement As clsRElementProperty) As String
        Dim strScript As String = If(String.IsNullOrEmpty(clsElement.strPackageName), "", clsElement.strPackageName + "::")
        If Not IsNothing(clsElement.lstObjects) AndAlso clsElement.lstObjects.Count > 0 Then
            For Each clsObject In clsElement.lstObjects
                strScript &= GetScriptElement(clsObject)
                strScript &= "$"
            Next
        End If
        strScript &= clsElement.strTxt
        Return strScript
    End Function

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
    '''             Iterates through the tokens in <paramref name="lstTokens"/>.
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
    ''' <param name="lstTokens">   The list statement tokens. </param>
    '''
    ''' <returns>   The list token brackets. </returns>
    '''--------------------------------------------------------------------------------------------
    Private Function GetLstTokenBrackets(lstTokens As List(Of clsRToken), ByRef intPos As Integer) As List(Of clsRToken)
        'if nothing to process then return empty list
        If lstTokens.Count <= 0 Then
            Return New List(Of clsRToken)
        End If

        Dim lstTokensNew As List(Of clsRToken) = New List(Of clsRToken)
        Dim clsToken As clsRToken

        While (intPos < lstTokens.Count)
            clsToken = lstTokens.Item(intPos)
            intPos += 1
            Select Case clsToken.strTxt
                Case "("
                    clsToken.lstTokens = GetLstTokenBrackets(lstTokens, intPos)
                Case ")"
                    lstTokensNew.Add(clsToken)
                    Return lstTokensNew
            End Select
            lstTokensNew.Add(clsToken)
        End While
        Return lstTokensNew
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
    '''     Note: this function cannot process the edge case where a binary operator is 
    '''     immediately followed by a unary operator with the same or a lower precedence 
    '''     (e.g. 'a^-b', 'a+~b', 'a~~b' etc.). This is because of the R default precedence rules. 
    '''     The workaround is to enclose the unary operator in brackets (e.g. 'a^(-b)', 'a+(~b)', 'a~(~b)' etc.).
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
                Case intOperatorsBrackets    'handles '[' and '[['
                    'TODO
                    'Case intOperatorsUnaryOnly   'handles '+' and '-' when they are unary operators (e.g. 'a * -b)
                    'TODO
                    'Case intOperatorsUserDefined 'handles all operators that start with '%'
                    'TODO
                Case Else 'handles all other operators including 'intOperatorsUnaryOnly', 'intOperatorsUserDefined' and 'intOperatorsTilda'
                    'if token is not the operator we're looking for, then exit select
                    If intPosOperators = intOperatorsUserDefined Then
                        If Not Regex.IsMatch(clsToken.strTxt, "^%.*%$") Then
                            Exit Select
                        End If
                    ElseIf Not arrOperatorPrecedence(intPosOperators).Contains(clsToken.strTxt) Then
                        Exit Select
                    End If
                    Select Case clsToken.enuToken
                        Case clsRToken.typToken.ROperatorBinary
                            If IsNothing(clsTokenPrev) Then
                                'TODO throw developer error: binary operator missing previous parameter
                                Exit Select
                            End If
                            'edge case: if we are looking for unary '+' or '-' and we found a binary '+' or '-'
                            If intPosOperators = intOperatorsUnaryOnly Then
                                'do not process (binary '+' and '-' have a lower precedence and will be processed later)
                                Exit Select
                            End If
                            'make the previous and next tokens, the children of the current token
                            clsToken.lstTokens.Add(clsTokenPrev.CloneMe)
                            bPrevTokenProcessed = True
                            clsToken.lstTokens.Add(GetNextToken(lstTokens, intPosTokens, intPosOperators))
                            intPosTokens += 1
                            'while next token is the same operator (e.g. 'a+b+c+d...'), 
                            '    then keep making the next token, the child of the current operator token
                            Dim clsTokenNext As clsRToken
                            While intPosTokens < lstTokens.Count - 1
                                clsTokenNext = GetNextToken(lstTokens, intPosTokens, intPosOperators)
                                If IsNothing(clsTokenNext) OrElse
                                        Not clsToken.enuToken = clsTokenNext.enuToken OrElse
                                        Not clsToken.strTxt = clsTokenNext.strTxt Then
                                    Exit While
                                End If

                                intPosTokens += 1
                                clsToken.lstTokens.Add(GetNextToken(lstTokens, intPosTokens, intPosOperators))
                                intPosTokens += 1
                            End While
                        Case clsRToken.typToken.ROperatorUnaryRight
                            'edge case: if we found a unary '+' or '-', but we are not currently processing the unary '+'and '-' operators
                            If arrOperatorPrecedence(intOperatorsUnaryOnly).Contains(clsToken.strTxt) AndAlso
                                Not intPosOperators = intOperatorsUnaryOnly Then
                                Exit Select
                            End If
                            'make the next token, the child of the current operator token
                            clsToken.lstTokens.Add(GetNextToken(lstTokens, intPosTokens, intPosOperators))
                            intPosTokens += 1
                        Case clsRToken.typToken.ROperatorUnaryLeft
                            If IsNothing(clsTokenPrev) OrElse Not intPosOperators = intOperatorsTilda Then
                                'TODO throw developer error: illegal unary left operator ('~' is the only valid unary left operator)
                                Exit Select
                            End If
                            'make the previous token, the child of the current operator token
                            clsToken.lstTokens.Add(clsTokenPrev.CloneMe)
                            bPrevTokenProcessed = True
                        Case Else
                            'TODO throw developer error: expecting an operator
                            Exit Select
                    End Select
            End Select

            'if token was not the operator we were looking for
            If Not bPrevTokenProcessed Then
                'add the previous token to the tree
                If Not IsNothing(clsTokenPrev) Then
                    lstTokensNew.Add(clsTokenPrev)
                End If
            End If

            'process the current token's children
            clsToken.lstTokens = GetLstTokenOperatorGroup(clsToken.CloneMe.lstTokens, intPosOperators)

            clsTokenPrev = clsToken.CloneMe
            bPrevTokenProcessed = False
            intPosTokens += 1
        End While

        If IsNothing(clsTokenPrev) Then
            'TODO throw developer error: There should always be at least one token still to add to the tree
            Return Nothing
        End If
        lstTokensNew.Add(clsTokenPrev.CloneMe)

        Return lstTokensNew
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns the next token in the <paramref name="lstTokens"/> list, after <paramref name="intPosTokens"/>. 
    '''             Processes the next token's children for the current operator.
    '''             If there is no next token then throws an error.</summary>
    '''
    ''' <param name="lstTokens">        The list of tokens. </param>
    ''' <param name="intPosTokens">     The position of the current token in the list. </param>
    ''' <param name="intPosOperators">  The group of operators currenty being processed. </param>
    '''
    ''' <returns>   The next token. </returns>
    '''--------------------------------------------------------------------------------------------
    Private Function GetNextToken(lstTokens As List(Of clsRToken), intPosTokens As Integer, intPosOperators As Integer) As clsRToken

        If intPosTokens >= (lstTokens.Count - 1) Then
            'TODO throw developer error: operator already has children or is missing a parameter
            Return Nothing
        End If

        'process the next token's children
        Dim clsTokenNext As clsRToken = lstTokens.Item(intPosTokens + 1).CloneMe
        'clsTokenNext.lstTokens = GetLstTokenOperatorGroup(clsTokenNext.CloneMe.lstTokens, intPosOperators) 'TODO is this needed?
        Return clsTokenNext
    End Function

    Private Function GetRElement(clsToken As clsRToken, Optional bBracketedNew As Boolean = False, Optional strPackageName As String = "", Optional lstObjects As List(Of clsRElement) = Nothing) As clsRElement
        If IsNothing(clsToken) Then
            Return Nothing
        End If

        Select Case clsToken.enuToken
            Case clsRToken.typToken.RBracket
                'if text is a round bracket, then return the bracket's child
                If clsToken.strTxt = "(" Then
                    Return GetRElement(clsToken.lstTokens.Item(0), True)
                End If

                Return New clsRElement(clsToken)

            Case clsRToken.typToken.RFunctionName
                Dim clsFunction As New clsRElementFunction(clsToken, bBracketedNew, strPackageName, lstObjects)
                'if function has at least one parameter
                'Note: Function tokens are structured as a tree.
                '      For example 'f(a,b,c=d)' is structured as:
                '    f
                '    ..(
                '    ....a
                '    ....,
                '    ......b 
                '    ....,
                '    ......=
                '    ........c
                '    ........d
                '    ........)    
                '    
                If clsToken.lstTokens.Count > 0 Then
                    'process each parameter
                    Dim bFirstParam As Boolean = True
                    For Each clsTokenParam In clsToken.lstTokens.Item(0).lstTokens
                        Dim clsParameter As clsRParameterNamed = GetRParameterNamed(clsTokenParam)
                        If Not IsNothing(clsParameter) Then
                            If bFirstParam AndAlso IsNothing(clsParameter.clsArgValue) Then
                                clsFunction.lstParameters.Add(clsParameter) 'add extra empty parameter for case 'f(,)'
                            End If
                            clsFunction.lstParameters.Add(clsParameter)
                        End If
                        bFirstParam = False
                    Next
                End If
                Return clsFunction

            Case clsRToken.typToken.ROperatorUnaryLeft
                Dim clsOperator As New clsRElementOperator(clsToken, bBracketedNew)
                If clsToken.lstTokens.Count < 1 Then
                    'TODO developer error: unary left operator must have at least one parameter
                    Return Nothing
                End If
                clsOperator.lstParameters.Add(GetRParameter(clsToken.lstTokens.Item(0)))
                Return clsOperator

            Case clsRToken.typToken.ROperatorUnaryRight
                Dim clsOperator As New clsRElementOperator(clsToken, bBracketedNew, True)
                If clsToken.lstTokens.Count < 1 Then
                    'TODO developer error: unary right operator must have at least one parameter
                    Return Nothing
                End If
                clsOperator.lstParameters.Add(GetRParameter(clsToken.lstTokens.Item(0)))
                Return clsOperator

            Case clsRToken.typToken.ROperatorBinary
                If clsToken.lstTokens.Count < 2 Then
                    'TODO developer error: binary operator must have at least 2 parameters
                    Return Nothing
                End If

                'if object operator
                Select Case clsToken.strTxt
                    Case "$"
                        Dim strPackageNameNew As String = ""
                        Dim lstObjectsNew As New List(Of clsRElement)

                        'add each object parameter to the object list (except last parameter)
                        For intPos As Integer = 0 To clsToken.lstTokens.Count - 2
                            'if the first parameter is a package operator ('::'), then make this the package name for the returned element
                            If intPos = 0 AndAlso
                                    clsToken.lstTokens.Item(intPos).enuToken = clsRToken.typToken.ROperatorBinary AndAlso
                                    clsToken.lstTokens.Item(intPos).strTxt = "::" Then
                                If clsToken.lstTokens.Item(intPos).lstTokens.Count < 2 Then
                                    'TODO developer error: package operator ('::') should have at least 2 parameters
                                    Return Nothing
                                End If
                                strPackageNameNew = clsToken.lstTokens.Item(intPos).lstTokens.Item(0).strTxt
                                lstObjectsNew.Add(GetRElement(clsToken.lstTokens.Item(intPos).lstTokens.Item(1)))
                                Continue For
                            End If
                            lstObjectsNew.Add(GetRElement(clsToken.lstTokens.Item(intPos)))
                        Next

                        'the last item in the parameter list is the element we need to return
                        Return GetRElement(clsToken.lstTokens.Item(clsToken.lstTokens.Count - 1), bBracketedNew, strPackageNameNew, lstObjectsNew)
                    Case "::"
                        'the '::' operator has two parameters, the first is the pacakage name, the second contains the element we need to return
                        Return GetRElement(clsToken.lstTokens.Item(1), bBracketedNew, clsToken.lstTokens.Item(0).strTxt)
                    Case Else 'else if not an object or package operator, then add each parameter to the operator
                        Dim clsOperator As New clsRElementOperator(clsToken, bBracketedNew)
                        For intPos As Integer = 0 To clsToken.lstTokens.Count - 1
                            clsOperator.lstParameters.Add(GetRParameter(clsToken.lstTokens.Item(intPos)))
                        Next
                        Return clsOperator
                End Select

            Case clsRToken.typToken.ROperatorBracket
                'Dim clsROperator As New clsRElementOperator
                'clsROperator.strTxt = clsToken.strText
                'clsROperator.bBracketed = bBracketedNew
                'clsROperator.clsAssignment = GetRElement(clsToken.clsAssign)
                'For Each clsTreeParam As clsRTreeElement In clsToken.lstTreeElements
                '    Dim clsRParameter As New clsRParameter
                '    clsRParameter.clsArgValue = GetRElement(clsTreeParam.lstTreeElements.Item(0))
                '    clsROperator.lstParameters.Add(clsRParameter)
                'Next
                'Select Case clsROperator.strTxt
                '    Case "["
                '        clsROperator.strTerminator = "]"
                '    Case "[["
                '        clsROperator.strTerminator = "]]"
                'End Select
                'Return clsROperator
            Case clsRToken.typToken.RSyntacticName, clsRToken.typToken.RConstantString
                'if element has a pacakage name or object list, then return a property element
                If Not String.IsNullOrEmpty(strPackageName) OrElse Not IsNothing(lstObjects) Then
                    Return New clsRElementProperty(clsToken, bBracketedNew, strPackageName, lstObjects)
                End If
                'else just return a regular element
                Return New clsRElement(clsToken, bBracketedNew)
            Case Else
                'TODO raise developer error: unknown token type
                Exit Select
        End Select
        'TODO developer error
        Return Nothing 'should never reach this point
    End Function

    Private Function GetRParameterNamed(clsToken As clsRToken) As clsRParameterNamed
        If IsNothing(clsToken) Then
            Return Nothing
        End If

        Select Case clsToken.strTxt
            Case "="
                If clsToken.lstTokens.Count = 2 Then
                    Dim clsParameterNamed As New clsRParameterNamed With {
                    .strArgumentName = clsToken.lstTokens.Item(0).strTxt}
                    clsParameterNamed.clsArgValue = GetRElement(clsToken.lstTokens.Item(1))
                    Return clsParameterNamed
                Else
                    'TODO developer error: '=' in a parmeter must always have an argument name and argument value
                    Return Nothing
                End If
            Case ","
                'if ',' is followed by a parameter name or value (e.g. 'fn(a,b)'), then return the parameter
                If clsToken.lstTokens.Count > 0 AndAlso Not clsToken.lstTokens.Item(0).strTxt = ")" Then
                    Return GetRParameterNamed(clsToken.lstTokens.Item(0))
                Else 'else return empty parameter (e.g. for cases like 'fn(a,)')
                    Return (New clsRParameterNamed)
                End If
            Case "("
                If clsToken.lstTokens.Count = 0 Then
                    'TODO developer error: '(' must have at least one child
                    Return Nothing
                End If
                Dim clsParameterNamed As New clsRParameterNamed With {
                    .clsArgValue = GetRElement(clsToken.lstTokens.Item(0), True)
                }
                Return clsParameterNamed
            Case ")"
                Return Nothing
            Case Else
                Dim clsParameterNamed As New clsRParameterNamed With {
                    .clsArgValue = GetRElement(clsToken)
                }
                Return clsParameterNamed
        End Select
    End Function

    Private Function GetRParameter(clsToken As clsRToken) As clsRParameter
        If IsNothing(clsToken) Then
            'TODO developer error: operator parameter is empty
            Return Nothing
        End If
        Return New clsRParameter With {.clsArgValue = GetRElement(clsToken)}
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
