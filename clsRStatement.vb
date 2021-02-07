Imports System.Text.RegularExpressions

Public Class clsRStatement

    ''' <summary>
    ''' If true, then when this R statement is converted to a script, then it will be 
    '''             terminated with a newline (else if false then a semicolon)
    ''' </summary>
    Public bTerminateWithNewline As Boolean = True

    ''' <summary>   The assignment operator used in this statement (e.g. '=' in the statement 'a=b').
    '''             If there is no assignment (e.g. as in 'myFunction(a)' then set to 'nothing'. </summary>
    Public strAssignmentOperator As String
    ''' <summary>   The element assigned to by the statement (e.g. 'a' in the statement 'a=b').
    '''             If there is no assignment (e.g. as in 'myFunction(a)' then set to 'nothing'. </summary>
    Public clsAssignment As clsRElement = Nothing
    ''' <summary>   The element to assigned in the statement (e.g. 'b' in the statement 'a=b').
    '''             If there is no assignment (e.g. as in 'myFunction(a)' then set to the top-
    '''             level element in the statement (e.g. 'myFunction'). </summary>
    Public clsElement As clsRElement

    Private strDebug As String = "" 'TODO delete when debugging done?

    ''' <summary>   The relative precedence of the R operators. This is a two-dimensional array 
    '''             because the operators are stored in groups together with operators that 
    '''             have the same precedence.</summary>
    Private ReadOnly arrOperatorPrecedence(17)() As String

    'Constants for operator precedence groups that have special characteristics (e.g. must be unary)
    Private ReadOnly intOperatorsBrackets As Integer = 2
    Private ReadOnly intOperatorsUnaryOnly As Integer = 4
    Private ReadOnly intOperatorsUserDefined As Integer = 6
    Private ReadOnly intOperatorsTilda As Integer = 13
    Private ReadOnly intOperatorsRightAssignment As Integer = 14
    Private ReadOnly intOperatorsLeftAssignment1 As Integer = 15
    Private ReadOnly intOperatorsLeftAssignment2 As Integer = 16

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Constructor an object representing a valid R statement.
    '''             <para>
    '''             Processes the tokens from <paramref name="lstTokens"/> from position <paramref name="intPos"/> 
    '''             to the end of statement, end of script or end of list (whichever comes first).
    '''             </para> </summary>
    '''
    ''' <param name="lstTokens">   The list of R tokens to process </param>
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
            If lstTokens.Item(intPos).enuToken = clsRToken.typToken.REndStatement OrElse 'we don't add this termination condition to the while statement 
                 lstTokens.Item(intPos).enuToken = clsRToken.typToken.REndScript Then    '    because we also want the token that terminates the statement 
                intPos += 1                                                              '    to be part of the statement's list of tokens
                Exit While
            End If
            intPos += 1
        End While

        'restructure the list into a token tree
        Dim lstTokenPresentation As List(Of clsRToken) = GetLstPresentation(lstStatementTokens, 0)
        Dim lstTokenBrackets As List(Of clsRToken) = GetLstTokenBrackets(lstTokenPresentation, 0)
        Dim lstTokenFunctionBrackets As List(Of clsRToken) = GetLstTokenFunctionBrackets(lstTokenBrackets)
        Dim lstTokenCommas As List(Of clsRToken) = GetLstTokenCommas(lstTokenFunctionBrackets, 0)
        Dim lstTokenTree As List(Of clsRToken) = GetLstTokenOperators(lstTokenCommas)
        strDebug = GetLstTokensAsString(lstTokenTree) 'TODO just for debug, remove

        'if the tree does not include at least one token, then raise development error
        If lstTokenTree.Count < 1 Then
            Throw New Exception("The token tree must contain at least one token.")
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
            ElseIf arrOperatorPrecedence(intOperatorsRightAssignment).Contains(lstTokenTree.Item(0).strTxt) Then
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

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns this object as a valid, executable R statement. </summary>
    '''
    ''' <returns>   The current state of this object as a valid, executable R statement. </returns>
    '''--------------------------------------------------------------------------------------------
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
                Throw New Exception("The statement's assignment operator is an unknown type.")
                Return Nothing
            End If
        End If

        strScript &= If(bTerminateWithNewline, vbLf, ";")
        Return strScript
    End Function

    Private Function GetScriptElement(clsElement As Object) As String

        If IsNothing(clsElement) Then
            Return Nothing
        End If

        Dim strScript As String = clsElement.clsPresentation.strPrefix
        strScript &= If(clsElement.bBracketed, "(", "") 'TODO also store presentation info for open and close brackets?

        Select Case clsElement.GetType()
            Case GetType(clsRElementFunction)
                strScript &= GetScriptElementProperty(clsElement)
                strScript &= "(" 'TODO also store presentation info for function brackets, commas and equals?
                If Not IsNothing(clsElement.lstParameters) Then
                    Dim bPrefixComma As Boolean = False
                    For Each clsRParameter In clsElement.lstParameters
                        strScript &= If(bPrefixComma, ",", "")
                        bPrefixComma = True
                        strScript &= If(String.IsNullOrEmpty(clsRParameter.strArgumentName), "", clsRParameter.clsPresentation.strPrefix & clsRParameter.strArgumentName + "=")
                        strScript &= GetScriptElement(clsRParameter.clsArgValue)
                    Next
                End If
                strScript &= ")"
            Case GetType(clsRElementProperty)
                strScript &= GetScriptElementProperty(clsElement)
            Case GetType(clsRElementOperator)
                Dim bPrefixOperator As Boolean = clsElement.bFirstParamOnRight
                For Each clsRParameter In clsElement.lstParameters
                    strScript &= If(bPrefixOperator, clsElement.strTxt, "")
                    bPrefixOperator = True
                    strScript &= GetScriptElement(clsRParameter.clsArgValue)
                Next
                strScript &= If(clsElement.lstParameters.Count = 1 AndAlso Not clsElement.bFirstParamOnRight, clsElement.strTxt, "")
            Case GetType(clsRElementKeyWord) 'TODO
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
    ''' <summary>   
    ''' Iterates through the tokens in <paramref name="lstTokens"/> and makes each presentation 
    ''' element a child of the next non-presentation element. 
    ''' <para>
    ''' A presentation element is an element that has no functionality and is only used to make 
    ''' the script easier to read. It may be a block of spaces, a comment or a newline that does
    ''' not end a statement.
    ''' </para><para>
    ''' For example, the list of tokens representing the following block of script:
    ''' </para><code>
    ''' # comment1 <para>
    ''' a =b # comment2 </para></code><para>
    ''' </para><para>
    ''' Will be structured as:</para><code><para>
    ''' a</para><para>
    ''' .."# comment1\n"</para><para>
    ''' =</para><para>
    ''' .." "</para><para>
    ''' b</para><para>
    ''' (endStatement)</para><para>
    ''' .." # comment2"</para><para>
    ''' </para></code></summary>
    ''' 
    ''' <param name="lstTokens">   The list of tokens to process. </param>
    '''
    ''' <returns>   A token tree where presentation information is stored as a child of the next 
    '''             non-presentation element. </returns>
    '''--------------------------------------------------------------------------------------------
    Private Function GetLstPresentation(lstTokens As List(Of clsRToken), intPos As Integer) As List(Of clsRToken)
        'if nothing to process then return empty list
        If lstTokens.Count < 1 Then
            Return New List(Of clsRToken)
        End If

        Dim lstTokensNew As List(Of clsRToken) = New List(Of clsRToken)
        Dim clsToken As clsRToken
        Dim strPrefix As String = ""

        While (intPos < lstTokens.Count)
            clsToken = lstTokens.Item(intPos)
            intPos += 1
            Select Case clsToken.enuToken
                Case clsRToken.typToken.RSpace, clsRToken.typToken.RComment, clsRToken.typToken.RNewLine
                    strPrefix &= clsToken.strTxt
                Case Else
                    If Not String.IsNullOrEmpty(strPrefix) Then
                        clsToken.lstTokens.Add(New clsRToken(strPrefix, clsRToken.typToken.RPresentation))
                    End If
                    lstTokensNew.Add(clsToken.CloneMe)
                    strPrefix = ""
            End Select
        End While

        'edge case: if there is still presentation information not yet added to a tree element
        '           (this may happen if the last statement in the script is not terminated 
        '           with a new line or '}')
        If Not String.IsNullOrEmpty(strPrefix) Then
            'add a new end statement token that contains the presentation information
            clsToken = New clsRToken(vbLf, clsRToken.typToken.REndStatement)
            clsToken.lstTokens.Add(New clsRToken(strPrefix, clsRToken.typToken.RPresentation))
            lstTokensNew.Add(clsToken)
        End If

        Return lstTokensNew
    End Function


    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Iterates through the tokens in <paramref name="lstTokens"/>.
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
                    Dim lstTokensTmp As List(Of clsRToken) = GetLstTokenBrackets(lstTokens, intPos)
                    For Each clsTokenChild As clsRToken In lstTokensTmp
                        If IsNothing(clsTokenChild) Then
                            Throw New Exception("Token has illegal empty child.")
                        End If
                        clsToken.lstTokens.Add(clsTokenChild.CloneMe)
                    Next
                Case ")"
                    lstTokensNew.Add(clsToken.CloneMe)
                    Return lstTokensNew
            End Select
            lstTokensNew.Add(clsToken.CloneMe)
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
                'if next steps will go out of bounds, then throw developer error
                If intPos > lstTokens.Count - 2 Then
                    Throw New Exception("The function's parameters have an unexpected format and cannot be processed.")
                    Exit While
                End If
                'make the function's open bracket a child of the function name
                intPos += 1
                Dim clsTokenNext = lstTokens.Item(intPos)
                clsToken.lstTokens.Add(clsTokenNext.CloneMe)
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
                        clsToken.lstTokens = clsToken.lstTokens.Concat(GetLstTokenCommas(lstStatementTokens, intPos, True)).ToList()
                    End If
                Case ")"
                    lstTokens.Add(clsToken)
                    Return lstTokens
                Case Else
                    If clsToken.lstTokens.Count > 0 Then
                        clsToken.lstTokens = GetLstTokenCommas(clsToken.CloneMe.lstTokens, 0)
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
        If lstTokens.Count < 1 Then
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
                        'edge case: if the operator already has (non-presentation) children then it means that it has already been processed.
                        '    This happens when the child is in the same precedence group as the parent but was processed first 
                        '    in accordance with the left to right rule (e.g. 'a/b*c').
                    ElseIf clsToken.lstTokens.Count > 1 OrElse
                            (clsToken.lstTokens.Count > 0 AndAlso
                            Not clsToken.lstTokens.Item(0).enuToken = clsRToken.typToken.RPresentation) Then
                        Exit Select
                    End If
                    Select Case clsToken.enuToken
                        Case clsRToken.typToken.ROperatorBinary
                            'edge case: if we are looking for unary '+' or '-' and we found a binary '+' or '-'
                            If intPosOperators = intOperatorsUnaryOnly Then
                                'do not process (binary '+' and '-' have a lower precedence and will be processed later)
                                Exit Select
                            ElseIf IsNothing(clsTokenPrev) Then
                                Throw New Exception("The binary operator has no parameter on its left.")
                            End If

                            'make the previous and next tokens, the children of the current token
                            clsToken.lstTokens.Add(clsTokenPrev.CloneMe)
                            bPrevTokenProcessed = True
                            clsToken.lstTokens.Add(GetNextToken(lstTokens, intPosTokens))
                            intPosTokens += 1
                            'while next token is the same operator (e.g. 'a+b+c+d...'), 
                            '    then keep making the next token, the child of the current operator token
                            Dim clsTokenNext As clsRToken
                            While intPosTokens < lstTokens.Count - 1
                                clsTokenNext = GetNextToken(lstTokens, intPosTokens)
                                If Not clsToken.enuToken = clsTokenNext.enuToken OrElse
                                        Not clsToken.strTxt = clsTokenNext.strTxt Then
                                    Exit While
                                End If

                                intPosTokens += 1
                                clsToken.lstTokens.Add(GetNextToken(lstTokens, intPosTokens))
                                intPosTokens += 1
                            End While
                        Case clsRToken.typToken.ROperatorUnaryRight
                            'edge case: if we found a unary '+' or '-', but we are not currently processing the unary '+'and '-' operators
                            If arrOperatorPrecedence(intOperatorsUnaryOnly).Contains(clsToken.strTxt) AndAlso
                                    Not intPosOperators = intOperatorsUnaryOnly Then
                                Exit Select
                            End If
                            'make the next token, the child of the current operator token
                            clsToken.lstTokens.Add(GetNextToken(lstTokens, intPosTokens))
                            intPosTokens += 1
                        Case clsRToken.typToken.ROperatorUnaryLeft
                            If IsNothing(clsTokenPrev) OrElse Not intPosOperators = intOperatorsTilda Then
                                Throw New Exception("Illegal unary left operator ('~' is the only valid unary left operator).")
                                Exit Select
                            End If
                            'make the previous token, the child of the current operator token
                            clsToken.lstTokens.Add(clsTokenPrev.CloneMe)
                            bPrevTokenProcessed = True
                        Case Else
                            Throw New Exception("The token has an unknown operator type.")
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
            Throw New Exception("Expected that there would still be a token to add to the tree.")
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
    '''
    ''' <returns>   The next token. </returns>
    '''--------------------------------------------------------------------------------------------
    Private Function GetNextToken(lstTokens As List(Of clsRToken), intPosTokens As Integer) As clsRToken

        If intPosTokens >= (lstTokens.Count - 1) Then
            Throw New Exception("Token list ended unexpectedly.")
        End If
        Dim clsTokenNext As clsRToken = lstTokens.Item(intPosTokens + 1).CloneMe
        Return clsTokenNext
    End Function

    Private Function GetRElement(clsToken As clsRToken,
                                 Optional bBracketedNew As Boolean = False,
                                 Optional strPackageName As String = "",
                                 Optional lstObjects As List(Of clsRElement) = Nothing) As clsRElement
        If IsNothing(clsToken) Then
            Return Nothing
        End If

        Select Case clsToken.enuToken
            Case clsRToken.typToken.RBracket
                'if text is a round bracket, then return the bracket's child
                If clsToken.strTxt = "(" Then
                    'an open bracket must have at least one child
                    If clsToken.lstTokens.Count < 1 OrElse clsToken.lstTokens.Count > 2 Then
                        Throw New Exception("Open bracket token has " & clsToken.lstTokens.Count &
                                            " children. An open bracket must have exactly one child (plus an Optional presentation child).")
                    End If
                    Return GetRElement(clsToken.lstTokens.Item(GetChildPosNonPresentation(clsToken)), True)
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
                    If clsToken.lstTokens.Count > 2 Then
                        Throw New Exception("Function token has " & clsToken.lstTokens.Count &
                                            " children. A function token may have 0 or 1 children (plus an Optional presentation child).")
                    End If

                    'process each parameter
                    Dim bFirstParam As Boolean = True
                    For Each clsTokenParam In clsToken.lstTokens.Item(clsToken.lstTokens.Count - 1).lstTokens
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
                If clsToken.lstTokens.Count < 1 OrElse clsToken.lstTokens.Count > 2 Then
                    Throw New Exception("Unary left operator token has " & clsToken.lstTokens.Count &
                                            " children. A Unary left operator must have 1 child (plus an Optional presentation child).")
                End If
                Dim clsOperator As New clsRElementOperator(clsToken, bBracketedNew)
                clsOperator.lstParameters.Add(GetRParameter(clsToken.lstTokens.Item(clsToken.lstTokens.Count - 1)))
                Return clsOperator

            Case clsRToken.typToken.ROperatorUnaryRight
                If clsToken.lstTokens.Count < 1 OrElse clsToken.lstTokens.Count > 2 Then
                    Throw New Exception("Unary right operator token has " & clsToken.lstTokens.Count &
                                            " children. A Unary right operator must have 1 child (plus an Optional presentation child).")
                End If
                Dim clsOperator As New clsRElementOperator(clsToken, bBracketedNew, True)
                clsOperator.lstParameters.Add(GetRParameter(clsToken.lstTokens.Item(clsToken.lstTokens.Count - 1)))
                Return clsOperator

            Case clsRToken.typToken.ROperatorBinary
                If clsToken.lstTokens.Count < 2 Then
                    Throw New Exception("Binary operator token has " & clsToken.lstTokens.Count &
                                            " children. A binary operator must have at least 2 children (plus an Optional presentation child).")
                End If

                'if object operator
                Select Case clsToken.strTxt
                    Case "$"
                        Dim strPackageNameNew As String = ""
                        Dim lstObjectsNew As New List(Of clsRElement)

                        'add each object parameter to the object list (except last parameter)
                        Dim startPos As Integer = If(clsToken.lstTokens.Item(0).enuToken = clsRToken.typToken.RPresentation, 1, 0)
                        For intPos As Integer = startPos To clsToken.lstTokens.Count - 2
                            'if the first parameter is a package operator ('::'), then make this the package name for the returned element
                            If intPos = startPos AndAlso
                                    clsToken.lstTokens.Item(intPos).enuToken = clsRToken.typToken.ROperatorBinary AndAlso
                                    clsToken.lstTokens.Item(intPos).strTxt = "::" Then
                                If clsToken.lstTokens.Item(intPos).lstTokens.Count < 2 Then
                                    Throw New Exception("Package operator token has " & clsToken.lstTokens.Item(intPos).lstTokens.Count &
                                                                " children. Package operator must have at least 2 children (plus an Optional presentation child).")
                                End If
                                strPackageNameNew = clsToken.lstTokens.Item(intPos).lstTokens.Item(clsToken.lstTokens.Item(intPos).lstTokens.Count - 2).strTxt
                                lstObjectsNew.Add(GetRElement(clsToken.lstTokens.Item(intPos).lstTokens.Item(clsToken.lstTokens.Item(intPos).lstTokens.Count - 1)))
                                Continue For
                            End If
                            lstObjectsNew.Add(GetRElement(clsToken.lstTokens.Item(intPos)))
                        Next

                        'the last item in the parameter list is the element we need to return
                        Return GetRElement(clsToken.lstTokens.Item(clsToken.lstTokens.Count - 1), bBracketedNew, strPackageNameNew, lstObjectsNew)
                    Case "::"
                        'the '::' operator has two parameters, the first is the package name, the second contains the element we need to return
                        Return GetRElement(clsToken.lstTokens.Item(clsToken.lstTokens.Count - 1), bBracketedNew, clsToken.lstTokens.Item(clsToken.lstTokens.Count - 2).strTxt)
                    Case Else 'else if not an object or package operator, then add each parameter to the operator
                        Dim clsOperator As New clsRElementOperator(clsToken, bBracketedNew)
                        Dim startPos As Integer = If(clsToken.lstTokens.Item(0).enuToken = clsRToken.typToken.RPresentation, 1, 0)
                        For intPos As Integer = startPos To clsToken.lstTokens.Count - 1
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
            Case clsRToken.typToken.RPresentation
                'presentation tokens should already have been processed by their parent token, so we can ignore
            Case Else
                Throw New Exception("The token has an unexpected type.")
        End Select
        Throw New Exception("It should be impossible for the code to reach this point.")
    End Function

    Private Function GetRParameterNamed(clsToken As clsRToken) As clsRParameterNamed
        If IsNothing(clsToken) Then
            Return Nothing
        End If

        Select Case clsToken.strTxt
            Case "="
                If clsToken.lstTokens.Count < 2 Then
                    Throw New Exception("Named parameter token has " & clsToken.lstTokens.Count &
                                        " children. Named parameter must have at least 2 children (plus an Optional presentation child).")
                End If

                Dim clsParameterNamed As New clsRParameterNamed With {
                    .strArgumentName = clsToken.lstTokens.Item(clsToken.lstTokens.Count - 2).strTxt}
                clsParameterNamed.clsArgValue = GetRElement(clsToken.lstTokens.Item(clsToken.lstTokens.Count - 1))
                clsParameterNamed.clsPresentation.strPrefix = If(clsToken.lstTokens.Item(0).enuToken = clsRToken.typToken.RPresentation, clsToken.lstTokens.Item(0).strTxt, "")
                Return clsParameterNamed
            Case ","
                ''if ',' is followed by a parameter name or value (e.g. 'fn(a,b)'), then return the parameter
                'If clsToken.lstTokens.Count > 0 AndAlso Not clsToken.lstTokens.Item(clsToken.lstTokens.Count - 1).strTxt = ")" Then
                '    Return GetRParameterNamed(clsToken.lstTokens.Item(GetChildPosNonPresentation(clsToken)))
                'Else 'else return empty parameter (e.g. for cases like 'fn(a,)')
                '    Return (New clsRParameterNamed)
                'End If
                'if ',' is followed by a parameter name or value (e.g. 'fn(a,b)'), then return the parameter
                Try
                    'throws exception if nonpresentation child not found
                    Return GetRParameterNamed(clsToken.lstTokens.Item(GetChildPosNonPresentation(clsToken)))
                Catch ex As Exception
                    'return empty parameter (e.g. for cases like 'fn(a,)')
                    Return (New clsRParameterNamed)
                End Try
            Case "("
                If clsToken.lstTokens.Count < 1 Then
                    Throw New Exception("Named parameter bracket token has " & clsToken.lstTokens.Count &
                                        " children. Named parameter must have at least 1 child (plus an Optional presentation child).")
                End If
                Dim clsParameterNamed As New clsRParameterNamed With {
                    .clsArgValue = GetRElement(clsToken.lstTokens.Item(GetChildPosNonPresentation(clsToken)), True)
                }
                clsParameterNamed.clsPresentation.strPrefix = If(clsToken.lstTokens.Item(0).enuToken = clsRToken.typToken.RPresentation, clsToken.lstTokens.Item(0).strTxt, "")
                Return clsParameterNamed
            Case ")"
                Return Nothing
            Case Else
                Dim clsParameterNamed As New clsRParameterNamed With {
                    .clsArgValue = GetRElement(clsToken)
                }
                clsParameterNamed.clsPresentation.strPrefix = If(clsToken.lstTokens.Count > 0 AndAlso clsToken.lstTokens.Item(0).enuToken = clsRToken.typToken.RPresentation, clsToken.lstTokens.Item(0).strTxt, "")
                Return clsParameterNamed
        End Select
    End Function

    Private Function GetChildPosNonPresentation(clsToken As clsRToken) As Integer
        For intPos As Integer = 0 To clsToken.lstTokens.Count - 1
            If Not clsToken.lstTokens.Item(intPos).enuToken = clsRToken.typToken.RPresentation AndAlso
                    Not (clsToken.lstTokens.Item(intPos).enuToken = clsRToken.typToken.RBracket AndAlso clsToken.lstTokens.Item(intPos).strTxt = ")") Then
                Return intPos
            End If
        Next
        Throw New Exception("Token must contain at least one non-presentation child.")
    End Function

    Private Function GetRParameter(clsToken As clsRToken) As clsRParameter
        If IsNothing(clsToken) Then
            Throw New ArgumentException("Cannot create a parameter from an empty token.")
        End If
        Dim clsParameter As clsRParameter = New clsRParameter With {.clsArgValue = GetRElement(clsToken)}
        clsParameter.clsPresentation.strPrefix = If(clsToken.lstTokens.Count > 0 AndAlso clsToken.lstTokens.Item(0).enuToken = clsRToken.typToken.RPresentation, clsToken.lstTokens.Item(0).strTxt, "")
        Return clsParameter
    End Function

    Private Function GetLstTokensAsString(lstTokens As List(Of clsRToken), Optional strIndent As String = "") As String
        If lstTokens Is Nothing OrElse lstTokens.Count = 0 Then
            Throw New ArgumentException("Cannot convert an empty token into a string.")
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
