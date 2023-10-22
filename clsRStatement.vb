Imports System.Text.RegularExpressions

''' <summary>   TODO Add class summary. </summary>
Public Class clsRStatement

    ''' <summary>   If true, then when this R statement is converted to a script, then it will be 
    '''             terminated with a newline (else if false then a semicolon)
    ''' </summary>
    Public bTerminateWithNewline As Boolean = True

    ''' <summary>   The assignment operator used in this statement (e.g. '=' in the statement 'a=b').
    '''             If there is no assignment (e.g. as in 'myFunction(a)' then set to 'nothing'. </summary>
    Public strAssignmentOperator As String

    ''' <summary>   If this R statement is converted to a script, then contains the formatting 
    '''             string that will prefix the assignment operator.
    '''             This is typically used to insert spaces before the assignment operator to line 
    '''             up the assignment operators in a list of assignments. For example:
    '''             <code>
    '''             shortName    = 1 <para>
    '''             veryLongName = 2 </para></code>
    '''             </summary>
    Public strAssignmentPrefix As String

    ''' <summary>   If this R statement is converted to a script, then contains the formatting 
    '''             string that will be placed at the end of the statement.
    '''             This is typically used to insert a comment at the end of the statement. 
    '''             For example:
    '''             <code>
    '''             a = b * 2 # comment1</code>
    '''             </summary>
    Public strSuffix As String

    ''' <summary>   The element assigned to by the statement (e.g. 'a' in the statement 'a=b').
    '''             If there is no assignment (e.g. as in 'myFunction(a)' then set to 'nothing'. </summary>
    Public clsAssignment As clsRElement = Nothing

    ''' <summary>   The element assigned in the statement (e.g. 'b' in the statement 'a=b').
    '''             If there is no assignment (e.g. as in 'myFunction(a)' then set to the top-
    '''             level element in the statement (e.g. 'myFunction'). </summary>
    Public clsElement As clsRElement

    ''' <summary>   The relative precedence of the R operators. This is a two-dimensional array 
    '''             because the operators are stored in groups together with operators that 
    '''             have the same precedence.</summary>
    Private ReadOnly arrOperatorPrecedence(18)() As String

    'Constants for operator precedence groups that have special characteristics (e.g. must be unary)
    Private ReadOnly intOperatorsBrackets As Integer = 2
    Private ReadOnly intOperatorsUnaryOnly As Integer = 4
    Private ReadOnly intOperatorsUserDefined As Integer = 6
    Private ReadOnly intOperatorsTilda As Integer = 14
    Private ReadOnly intOperatorsRightAssignment As Integer = 15
    Private ReadOnly intOperatorsLeftAssignment1 As Integer = 16
    Private ReadOnly intOperatorsLeftAssignment2 As Integer = 17

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   
    ''' Constructs an object representing a valid R statement.<para>
    ''' Processes the tokens from <paramref name="lstTokens"/> from position <paramref name="intPos"/> 
    ''' to the end of statement, end of script or end of list (whichever comes first).</para></summary>
    '''
    ''' <param name="lstTokens">   The list of R tokens to process </param>
    ''' <param name="intPos">      [in,out] The position in the list to start processing </param>
    '''--------------------------------------------------------------------------------------------
    Public Sub New(lstTokens As List(Of clsRToken), ByRef intPos As Integer, dctAssignments As Dictionary(Of String, clsRStatement))

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
        arrOperatorPrecedence(7) = New String() {"|>"}
        arrOperatorPrecedence(8) = New String() {"*", "/"}
        arrOperatorPrecedence(9) = New String() {"+", "-"}
        arrOperatorPrecedence(10) = New String() {"<", ">", "<>", "<=", ">=", "==", "!="}
        arrOperatorPrecedence(11) = New String() {"!"}
        arrOperatorPrecedence(12) = New String() {"&", "&&"}
        arrOperatorPrecedence(13) = New String() {"|", "||"}
        arrOperatorPrecedence(intOperatorsTilda) = New String() {"~"}          'unary or binary
        arrOperatorPrecedence(intOperatorsRightAssignment) = New String() {"->", "->>"}
        arrOperatorPrecedence(intOperatorsLeftAssignment1) = New String() {"<-", "<<-"}
        arrOperatorPrecedence(intOperatorsLeftAssignment2) = New String() {"="}

        'create list of tokens for this statement
        Dim lstStatementTokens As List(Of clsRToken) = New List(Of clsRToken)
        While intPos < lstTokens.Count
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
        Dim lstTokenFunctionCommas As List(Of clsRToken) = GetLstTokenFunctionCommas(lstTokenFunctionBrackets, 0)
        Dim lstTokenTree As List(Of clsRToken) = GetLstTokenOperators(lstTokenFunctionCommas)

        'if the tree does not include at least one token, then raise development error
        If lstTokenTree.Count < 1 Then
            Throw New Exception("The token tree must contain at least one token.")
        End If

        'if the statement includes an assignment, then construct the assignment element
        If lstTokenTree.Item(0).enuToken = clsRToken.typToken.ROperatorBinary AndAlso
                lstTokenTree.Item(0).lstTokens.Count > 1 Then

            Dim clsTokenChildLeft As clsRToken = lstTokenTree.Item(0).lstTokens.Item(lstTokenTree.Item(0).lstTokens.Count - 2)
            Dim clsTokenChildRight As clsRToken = lstTokenTree.Item(0).lstTokens.Item(lstTokenTree.Item(0).lstTokens.Count - 1)

            'if the statement has a left assignment (e.g. 'x<-value', 'x<<-value' or 'x=value')
            If arrOperatorPrecedence(intOperatorsLeftAssignment1).Contains(lstTokenTree.Item(0).strTxt) OrElse
                arrOperatorPrecedence(intOperatorsLeftAssignment2).Contains(lstTokenTree.Item(0).strTxt) Then
                clsAssignment = GetRElement(clsTokenChildLeft, dctAssignments)
                clsElement = GetRElement(clsTokenChildRight, dctAssignments)
            ElseIf arrOperatorPrecedence(intOperatorsRightAssignment).Contains(lstTokenTree.Item(0).strTxt) Then
                'else if the statement has a right assignment (e.g. 'value->x' or 'value->>x')
                clsAssignment = GetRElement(clsTokenChildRight, dctAssignments)
                clsElement = GetRElement(clsTokenChildLeft, dctAssignments)
            End If
        End If

        'if there was an assigment then set the assignment operator and its presentation information
        If Not IsNothing(clsAssignment) Then
            strAssignmentOperator = lstTokenTree.Item(0).strTxt
            strAssignmentPrefix = If(lstTokenTree.Item(0).lstTokens.Item(0).enuToken = clsRToken.typToken.RPresentation,
                                     lstTokenTree.Item(0).lstTokens.Item(0).strTxt, "")
        Else 'if there was no assignment, then build the main element from the token tree's top element
            clsElement = GetRElement(lstTokenTree.Item(0), dctAssignments)
        End If

        'check if the statement is terminated with a semi-colon
        Dim clsTokenEndStatement As clsRToken = lstTokenTree.Item(lstTokenTree.Count - 1)
        If clsTokenEndStatement.enuToken = clsRToken.typToken.REndStatement AndAlso
             clsTokenEndStatement.strTxt = ";" Then
            bTerminateWithNewline = False
        End If

        'store any remaining presentation data
        'TODO
        strSuffix = If(clsTokenEndStatement.lstTokens.Count > 0 AndAlso
                       clsTokenEndStatement.lstTokens.Item(0).enuToken = clsRToken.typToken.RPresentation,
                       clsTokenEndStatement.lstTokens.Item(0).strTxt, "")
        strSuffix = If(strSuffix.EndsWith(vbLf), strSuffix.Substring(0, strSuffix.Length - 1), strSuffix)
        'strSuffix = If(Not IsNothing(clsElement) AndAlso
        '               clsTokenEndStatement.lstTokens.Count > 0 AndAlso
        '               clsTokenEndStatement.lstTokens.Item(0).enuToken = clsRToken.typToken.RPresentation,
        '               clsTokenEndStatement.lstTokens.Item(0).strTxt, "")

    End Sub

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   
    ''' Returns this object as a valid, executable R statement. <para>
    ''' The script may contain formatting information such as spaces, comments and extra new lines.
    ''' If this object was created by analysing original R script, then the returned script's 
    ''' formatting will be as close as possible to the original.</para><para>
    ''' The script may vary slightly because some formatting information is lost in the object 
    ''' model. For lost formatting, the formatting will be done according to the guidelines in
    ''' https://style.tidyverse.org/syntax.html  </para><para>
    ''' The returned script will always show:</para><list type="bullet"><item>
    ''' No spaces before commas</item><item>
    ''' No spaces before brackets</item><item>
    ''' No spaces before package ('::') and object ('$') operators</item><item>
    ''' One space before parameter assignments ('=')</item><item>
    ''' For example,  'pkg ::obj1 $obj2$fn1 (a ,b=1,    c    = 2 )' will be returned as 
    '''                                                 'pkg::obj1$obj2$fn1(a, b =1, c = 2)'</item>
    ''' </list></summary>
    '''
    ''' <param name="bIncludeFormatting">   If True, then include all formatting information in 
    '''     returned string (comments, indents, padding spaces, extra line breaks etc.). </param>
    '''
    ''' <returns>   The current state of this object as a valid, executable R statement. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Function GetAsExecutableScript(Optional bIncludeFormatting As Boolean = True) As String
        Dim strScript As String
        Dim strElement As String = GetScriptElement(clsElement, bIncludeFormatting)
        'if there is no assignment, then just process the statement's element
        If IsNothing(clsAssignment) OrElse String.IsNullOrEmpty(strAssignmentOperator) Then
            strScript = strElement
        Else 'else if the statement has an assignment
            Dim strAssignment As String = GetScriptElement(clsAssignment, bIncludeFormatting)
            Dim strAssignmentPrefixTmp = If(bIncludeFormatting, strAssignmentPrefix, "")
            'if the statement has a left assignment (e.g. 'x<-value', 'x<<-value' or 'x=value')
            If arrOperatorPrecedence(intOperatorsLeftAssignment1).Contains(strAssignmentOperator) OrElse
                arrOperatorPrecedence(intOperatorsLeftAssignment2).Contains(strAssignmentOperator) Then
                strScript = strAssignment & strAssignmentPrefixTmp & strAssignmentOperator & strElement
            ElseIf arrOperatorPrecedence(intOperatorsRightAssignment).Contains(strAssignmentOperator) Then
                'else if the statement has a right assignment (e.g. 'value->x' or 'value->>x')
                strScript = strElement & strAssignmentPrefixTmp & strAssignmentOperator & strAssignment
            Else
                Throw New Exception("The statement's assignment operator is an unknown type.")
            End If
        End If

        If bIncludeFormatting Then
            strScript &= strSuffix
            strScript &= If(bTerminateWithNewline, vbLf, ";")
        End If

        Return strScript
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns <paramref name="clsElement"/> as an executable R script. </summary>
    '''
    ''' <param name="clsElement">   The R element to convert to an executable R script. 
    '''                             The R element may be a function, operator, constant, 
    '''                             syntactic name, key word etc. </param>
    '''
    ''' <param name="bIncludeFormatting">   If True, then include all formatting information in 
    '''     returned string (comments, indents, padding spaces, extra line breaks etc.). </param>
    '''
    ''' <returns>   <paramref name="clsElement"/> as an executable R script. </returns>
    '''--------------------------------------------------------------------------------------------
    Private Function GetScriptElement(clsElement As Object, Optional bIncludeFormatting As Boolean = True) As String

        If IsNothing(clsElement) Then
            Return ""
        End If

        Dim strScript As String = ""
        Dim strElementPrefix As String = If(bIncludeFormatting, clsElement.strPrefix, "")
        strScript &= If(clsElement.bBracketed, "(", "")

        Select Case clsElement.GetType()
            Case GetType(clsRElementFunction)
                strScript &= GetScriptElementProperty(clsElement, bIncludeFormatting)
                strScript &= "("
                If Not IsNothing(clsElement.lstParameters) Then
                    Dim bPrefixComma As Boolean = False
                    For Each clsRParameter In clsElement.lstParameters
                        strScript &= If(bPrefixComma, ",", "")
                        bPrefixComma = True
                        Dim strParameterPrefix As String = If(bIncludeFormatting, clsRParameter.strPrefix, "")
                        strScript &= If(String.IsNullOrEmpty(clsRParameter.strArgName), "", strParameterPrefix & clsRParameter.strArgName + " =")
                        strScript &= GetScriptElement(clsRParameter.clsArgValue, bIncludeFormatting)
                    Next
                End If
                strScript &= ")"
            Case GetType(clsRElementProperty)
                strScript &= GetScriptElementProperty(clsElement, bIncludeFormatting)
            Case GetType(clsRElementOperator)
                If clsElement.strTxt = "[" OrElse clsElement.strTxt = "[[" Then
                    Dim bOperatorAppended As Boolean = False
                    For Each clsRParameter In clsElement.lstParameters
                        strScript &= GetScriptElement(clsRParameter.clsArgValue, bIncludeFormatting)
                        strScript &= If(bOperatorAppended, "", strElementPrefix & clsElement.strTxt)
                        bOperatorAppended = True
                    Next

                    Select Case clsElement.strTxt
                        Case "["
                            strScript &= "]"
                        Case "[["
                            strScript &= "]]"
                    End Select
                Else
                    Dim bPrefixOperator As Boolean = clsElement.bFirstParamOnRight
                    For Each clsRParameter In clsElement.lstParameters
                        strScript &= If(bPrefixOperator, strElementPrefix & clsElement.strTxt, "")
                        bPrefixOperator = True
                        strScript &= GetScriptElement(clsRParameter.clsArgValue, bIncludeFormatting)
                    Next
                    strScript &= If(clsElement.lstParameters.Count = 1 AndAlso Not clsElement.bFirstParamOnRight, strElementPrefix & clsElement.strTxt, "")
                End If
            Case GetType(clsRElementKeyWord) 'TODO add key word functionality
            Case GetType(clsRElement), GetType(clsRElementAssignable)
                strScript &= strElementPrefix & clsElement.strTxt
        End Select
        strScript &= If(clsElement.bBracketed, ")", "")
        Return strScript
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns <paramref name="clsElement"/> as an executable R script. </summary>
    '''
    ''' <param name="clsElement">   The R element to convert to an executable R script. The R element
    '''                             may have an associated package name, and a list of associated 
    '''                             objects e.g. 'pkg::obj1$obj2$fn1(a)'. </param>
    '''
    ''' <param name="bIncludeFormatting">   If True, then include all formatting information in 
    '''     returned string (comments, indents, padding spaces, extra line breaks etc.). </param>
    '''
    ''' <returns>   <paramref name="clsElement"/> as an executable R script. </returns>
    '''--------------------------------------------------------------------------------------------
    Private Function GetScriptElementProperty(clsElement As clsRElementProperty, Optional bIncludeFormatting As Boolean = True) As String
        Dim strScript As String = If(bIncludeFormatting, clsElement.strPrefix, "") &
                                  If(String.IsNullOrEmpty(clsElement.strPackageName), "", clsElement.strPackageName & "::")
        If Not IsNothing(clsElement.lstObjects) AndAlso clsElement.lstObjects.Count > 0 Then
            For Each clsObject In clsElement.lstObjects
                strScript &= GetScriptElement(clsObject, bIncludeFormatting)
                strScript &= "$"
            Next
        End If
        strScript &= clsElement.strTxt
        Return strScript
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

        If lstTokens.Count < 1 Then
            Return New List(Of clsRToken)
        End If

        Dim lstTokensNew As List(Of clsRToken) = New List(Of clsRToken)
        Dim clsToken As clsRToken
        Dim strPrefix As String = ""

        While intPos < lstTokens.Count
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

        'Edge case: if there is still presentation information not yet added to a tree element
        '           (this may happen if the last statement in the script is not terminated 
        '           with a new line or '}')
        If Not String.IsNullOrEmpty(strPrefix) Then
            'add a new end statement token that contains the presentation information
            clsToken = New clsRToken("", clsRToken.typToken.REndStatement)
            clsToken.lstTokens.Add(New clsRToken(strPrefix, clsRToken.typToken.RPresentation))
            lstTokensNew.Add(clsToken)
        End If

        Return lstTokensNew
    End Function


    '''--------------------------------------------------------------------------------------------
    ''' <summary>   
    ''' Iterates through the tokens in <paramref name="lstTokens"/>.
    ''' If the token is a '(' then it makes everything inside the brackets a child of the '(' token.
    ''' If the '(' belongs to a function then makes the '(' a child of the function. Brackets may 
    ''' be nested. For example, '(a*(b+c))' is structured as:<code>
    '''   (<para>
    '''   ..a</para><para>
    '''   ..*</para><para>
    '''   ..(</para><para>
    '''   ....b</para><para>
    '''   ....+</para><para>
    '''   ....c</para><para>
    '''   ....)</para><para>
    '''   ..)</para></code></summary>
    '''
    ''' <param name="lstTokens">   The token tree to restructure. </param>
    '''
    ''' <returns>   A token tree restructured for round brackets. </returns>
    '''--------------------------------------------------------------------------------------------
    Private Function GetLstTokenBrackets(lstTokens As List(Of clsRToken), ByRef intPos As Integer) As List(Of clsRToken)

        If lstTokens.Count <= 0 Then
            Return New List(Of clsRToken)
        End If

        Dim lstTokensNew As List(Of clsRToken) = New List(Of clsRToken)
        Dim clsToken As clsRToken
        While intPos < lstTokens.Count
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
    ''' Traverses the tree of tokens in <paramref name="lstTokens"/>. If the token is a function name then it 
    ''' makes the subsequent '(' a child of the function name token. </summary>
    '''
    ''' <param name="lstTokens">   The token tree to restructure. </param>
    '''
    ''' <returns>   A token tree restructured for function names. </returns>
    '''--------------------------------------------------------------------------------------------
    Private Function GetLstTokenFunctionBrackets(lstTokens As List(Of clsRToken)) As List(Of clsRToken)

        If lstTokens.Count <= 0 Then
            Return New List(Of clsRToken)
        End If

        Dim lstTokensNew As List(Of clsRToken) = New List(Of clsRToken)
        Dim clsToken As clsRToken
        Dim intPos As Integer = 0
        While intPos < lstTokens.Count
            clsToken = lstTokens.Item(intPos)

            If clsToken.enuToken = clsRToken.typToken.RFunctionName Then
                'if next steps will go out of bounds, then throw developer error
                If intPos > lstTokens.Count - 2 Then
                    Throw New Exception("The function's parameters have an unexpected format and cannot be processed.")
                    Exit While
                End If
                'make the function's open bracket a child of the function name
                intPos += 1
                clsToken.lstTokens.Add(lstTokens.Item(intPos).CloneMe)
            End If
            clsToken.lstTokens = GetLstTokenFunctionBrackets(clsToken.CloneMe.lstTokens)
            lstTokensNew.Add(clsToken.CloneMe)
            intPos += 1
        End While
        Return lstTokensNew
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Traverses the tree of tokens in <paramref name="lstTokens"/>. If the token is a ',' that 
    ''' separates function parameters, then it makes everything up to the next ',' or ')' a child 
    ''' of the ',' token. Parameters between function commas are optional. For example, 
    ''' `myFunction(a,,b)` is structured as: <code>
    '''   myFunction<para>
    '''   ..(</para><para>
    '''   ....a</para><para>
    '''   ....,</para><para>
    '''   ....,</para><para>
    '''   ......b</para><para>
    '''   ......)</para></code>
    ''' Commas used within square brackets (e.g. `a[b,c]`, `a[b,]` etc.) are ignored.
    ''' </summary>
    '''
    ''' <param name="lstTokens">        The token tree to restructure. </param>
    ''' <param name="intPos">           [in,out] The position in the list to start processing. </param>
    ''' <param name="bProcessingComma"> (Optional) True if function called when already processing 
    '''     a comma (prevents commas being nested inside each other). </param>
    '''
    ''' <returns>   A token tree restructured for function commas. </returns>
    '''--------------------------------------------------------------------------------------------
    Private Function GetLstTokenFunctionCommas(lstTokens As List(Of clsRToken),
                                       ByRef intPos As Integer,
                                       Optional bProcessingComma As Boolean = False) As List(Of clsRToken)
        Dim lstTokensNew As List(Of clsRToken) = New List(Of clsRToken)
        Dim clsToken As clsRToken
        Dim lstOpenBrackets As New List(Of String) From {"[", "[["}
        Dim lstCloseBrackets As New List(Of String) From {"]", "]]"}
        Dim iNumOpenBrackets As Integer = 0

        While intPos < lstTokens.Count
            clsToken = lstTokens.Item(intPos)

            'only process commas that separate function parameters,
            '    ignore commas inside square bracket (e.g. `a[b,c]`)
            iNumOpenBrackets += If(lstOpenBrackets.Contains(clsToken.strTxt), 1, 0)
            iNumOpenBrackets -= If(lstCloseBrackets.Contains(clsToken.strTxt), 1, 0)
            If iNumOpenBrackets = 0 AndAlso clsToken.strTxt = "," Then
                If bProcessingComma Then
                    intPos -= 1  'ensure this comma is processed in the level above
                    Return lstTokensNew
                Else
                    intPos += 1
                    clsToken.lstTokens = clsToken.lstTokens.Concat(GetLstTokenFunctionCommas(lstTokens, intPos, True)).ToList()
                End If
            Else
                clsToken.lstTokens = GetLstTokenFunctionCommas(clsToken.CloneMe.lstTokens, 0)
            End If

            lstTokensNew.Add(clsToken)
            intPos += 1
        End While
        Return lstTokensNew
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary> 
    ''' Iterates through all the possible operators in order of precedence. For each operator, 
    ''' traverses the tree of tokens in <paramref name="lstTokens"/>. If the operator is found then 
    ''' the operator's parameters (typically the tokens to the left and right of the operator) are 
    ''' made children of the operator. For example, 'a*b+c' is structured as:<code>
    '''   +<para>
    '''   ..*</para><para>
    '''   ....a</para><para>
    '''   ....b</para><para>
    '''   ..c</para></code></summary>
    '''
    ''' <param name="lstTokens">   The token tree to restructure. </param>
    '''
    ''' <returns>   A token tree restructured for all the possible operators. </returns>
    '''--------------------------------------------------------------------------------------------
    Private Function GetLstTokenOperators(lstTokens As List(Of clsRToken)) As List(Of clsRToken)
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
    ''' Traverses the tree of tokens in <paramref name="lstTokens"/>. If one of the operators in 
    ''' the <paramref name="intPosOperators"/> group is found, then the operator's parameters 
    ''' (typically the tokens to the left and right of the operator) are made children of the 
    ''' operator. For example, 'a*b+c' is structured as:<code>
    '''   +<para>
    '''   ..*</para><para>
    '''   ....a</para><para>
    '''   ....b</para><para>
    '''   ..c</para></code>
    '''
    ''' Edge case: This function cannot process the  case where a binary operator is immediately 
    ''' followed by a unary operator with the same or a lower precedence (e.g. 'a^-b', 'a+~b', 
    ''' 'a~~b' etc.). This is because of the R default precedence rules. The workaround is to 
    ''' enclose the unary operator in brackets (e.g. 'a^(-b)', 'a+(~b)', 'a~(~b)' etc.).
    ''' </summary>
    ''' <param name="lstTokens">        The token tree to restructure. </param>
    ''' <param name="intPosOperators">  The group of operators to search for in the tree. </param>
    '''
    ''' <returns>   A token tree restructured for the specified group of operators. </returns>
    '''--------------------------------------------------------------------------------------------
    Private Function GetLstTokenOperatorGroup(lstTokens As List(Of clsRToken), intPosOperators As Integer) As List(Of clsRToken)

        If lstTokens.Count < 1 Then
            Return New List(Of clsRToken)
        End If

        Dim lstTokensNew As List(Of clsRToken) = New List(Of clsRToken)
        Dim clsToken As clsRToken
        Dim clsTokenPrev As clsRToken = Nothing
        Dim bPrevTokenProcessed As Boolean = False

        Dim intPosTokens As Integer = 0
        While intPosTokens < lstTokens.Count
            clsToken = lstTokens.Item(intPosTokens).CloneMe

            'if the token is the operator we are looking for and it has not been processed already
            'Edge case: if the operator already has (non-presentation) children then it means 
            '           that it has already been processed. This happens when the child is in the 
            '           same precedence group as the parent but was processed first in accordance 
            '           with the left to right rule (e.g. 'a/b*c').
            If (arrOperatorPrecedence(intPosOperators).Contains(clsToken.strTxt) OrElse
                    intPosOperators = intOperatorsUserDefined AndAlso Regex.IsMatch(clsToken.strTxt, "^%.*%$")) AndAlso
                    (clsToken.lstTokens.Count = 0 OrElse (clsToken.lstTokens.Count = 1 AndAlso
                    clsToken.lstTokens.Item(0).enuToken = clsRToken.typToken.RPresentation)) Then

                Select Case clsToken.enuToken
                    Case clsRToken.typToken.ROperatorBracket 'handles '[' and '[['
                        If intPosOperators <> intOperatorsBrackets Then
                            Exit Select
                        End If

                        'make the previous and next tokens (up to the corresponding close bracket), the children of the current token
                        clsToken.lstTokens.Add(clsTokenPrev.CloneMe)
                        bPrevTokenProcessed = True
                        intPosTokens += 1
                        Dim strCloseBracket = If(clsToken.strTxt = "[", "]", "]]")
                        Dim iNumOpenBrackets As Integer = 1
                        While intPosTokens < lstTokens.Count
                            iNumOpenBrackets += If(lstTokens.Item(intPosTokens).strTxt = clsToken.strTxt, 1, 0)
                            iNumOpenBrackets -= If(lstTokens.Item(intPosTokens).strTxt = strCloseBracket, 1, 0)
                            'discard the terminating cloe bracket
                            If iNumOpenBrackets = 0 Then
                                Exit While
                            End If
                            clsToken.lstTokens.Add(lstTokens.Item(intPosTokens).CloneMe)
                            intPosTokens += 1
                        End While

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
                        End If
                        'make the previous token, the child of the current operator token
                        clsToken.lstTokens.Add(clsTokenPrev.CloneMe)
                        bPrevTokenProcessed = True
                    Case Else
                        Throw New Exception("The token has an unknown operator type.")
                End Select
            End If

            'if token was not the operator we were looking for
            '    (or we were looking for a unary right operator)
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
    ''' <summary>   Returns a clone of the next token in the <paramref name="lstTokens"/> list, 
    '''             after <paramref name="intPosTokens"/>. If there is no next token then throws 
    '''             an error.</summary>
    '''
    ''' <param name="lstTokens">        The list of tokens. </param>
    ''' <param name="intPosTokens">     The position of the current token in the list. </param>
    '''
    ''' <returns>   A clone of the next token in the <paramref name="lstTokens"/> list. </returns>
    '''--------------------------------------------------------------------------------------------
    Private Function GetNextToken(lstTokens As List(Of clsRToken), intPosTokens As Integer) As clsRToken
        If intPosTokens >= (lstTokens.Count - 1) Then
            Throw New Exception("Token list ended unexpectedly.")
        End If
        Return lstTokens.Item(intPosTokens + 1).CloneMe
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns an R element object constructed from the <paramref name="clsToken"/> 
    '''             token. </summary>
    '''
    ''' <param name="clsToken">         The token to convert into an R element object. </param>
    ''' <param name="dctAssignments">   Dictionary containing all the current existing assignments. 
    '''                                 The key is the name of the variable. The value is a reference 
    '''                                 to the R statement that performed the assignment. </param>
    ''' <param name="bBracketedNew">    (Optional) True if the token is enclosed in brackets. </param>
    ''' <param name="strPackageName">   (Optional) The package name associated with the token. </param>
    ''' <param name="strPackagePrefix"> (Optional) The formatting string that prefixes the package 
    '''                                 name (e.g. spaces or comment lines). </param>
    ''' <param name="lstObjects">       (Optional) The list of objects associated with the token 
    '''                                 (e.g. 'obj1$obj2$myFn()'). </param>
    '''
    ''' <returns>   An R element object constructed from the <paramref name="clsToken"/>
    '''             token. </returns>
    '''--------------------------------------------------------------------------------------------
    Private Function GetRElement(clsToken As clsRToken,
                                 dctAssignments As Dictionary(Of String, clsRStatement),
                                 Optional bBracketedNew As Boolean = False,
                                 Optional strPackageName As String = "",
                                 Optional strPackagePrefix As String = "",
                                 Optional lstObjects As List(Of clsRElement) = Nothing) As clsRElement
        If IsNothing(clsToken) Then
            Throw New ArgumentException("Cannot create an R element from an empty token.")
        End If

        Select Case clsToken.enuToken
            Case clsRToken.typToken.RBracket
                'if text is a round bracket, then return the bracket's child
                If clsToken.strTxt = "(" Then
                    'an open bracket must have at least one child
                    If clsToken.lstTokens.Count < 1 OrElse clsToken.lstTokens.Count > 3 Then
                        Throw New Exception("Open bracket token has " & clsToken.lstTokens.Count &
                                " children. An open bracket must have exactly one child (plus an " &
                                "optional presentation child and/or an optional close bracket).")
                    End If
                    Return GetRElement(GetChildPosNonPresentation(clsToken), dctAssignments, True)
                End If
                Return New clsRElement(clsToken)

            Case clsRToken.typToken.RFunctionName
                Dim clsFunction As New clsRElementFunction(clsToken, bBracketedNew, strPackageName, strPackagePrefix, lstObjects)
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
                If clsToken.lstTokens.Count < 1 OrElse clsToken.lstTokens.Count > 2 Then
                    Throw New Exception("Function token has " & clsToken.lstTokens.Count &
                                            " children. A function token must have 1 child (plus an optional presentation child).")
                End If

                'process each parameter
                Dim bFirstParam As Boolean = True
                For Each clsTokenParam In clsToken.lstTokens.Item(clsToken.lstTokens.Count - 1).lstTokens
                    'if list item is a presentation element, then ignore it
                    If clsTokenParam.enuToken = clsRToken.typToken.RPresentation Then
                        If bFirstParam Then
                            Continue For
                        End If
                        Throw New Exception("Function parameter list contained an unexpected presentation element.")
                    End If

                    Dim clsParameter As clsRParameter = GetRParameterNamed(clsTokenParam, dctAssignments)
                    If Not IsNothing(clsParameter) Then
                        If bFirstParam AndAlso IsNothing(clsParameter.clsArgValue) Then
                            clsFunction.lstParameters.Add(clsParameter) 'add extra empty parameter for case 'f(,)'
                        End If
                        clsFunction.lstParameters.Add(clsParameter)
                    End If
                    bFirstParam = False
                Next
                Return clsFunction

            Case clsRToken.typToken.ROperatorUnaryLeft
                If clsToken.lstTokens.Count < 1 OrElse clsToken.lstTokens.Count > 2 Then
                    Throw New Exception("Unary left operator token has " & clsToken.lstTokens.Count &
                                            " children. A Unary left operator must have 1 child (plus an optional presentation child).")
                End If
                Dim clsOperator As New clsRElementOperator(clsToken, bBracketedNew)
                clsOperator.lstParameters.Add(GetRParameter(clsToken.lstTokens.Item(clsToken.lstTokens.Count - 1), dctAssignments))
                Return clsOperator

            Case clsRToken.typToken.ROperatorUnaryRight
                If clsToken.lstTokens.Count < 1 OrElse clsToken.lstTokens.Count > 2 Then
                    Throw New Exception("Unary right operator token has " & clsToken.lstTokens.Count &
                                            " children. A Unary right operator must have 1 child (plus an optional presentation child).")
                End If
                Dim clsOperator As New clsRElementOperator(clsToken, bBracketedNew, True)
                clsOperator.lstParameters.Add(GetRParameter(clsToken.lstTokens.Item(clsToken.lstTokens.Count - 1), dctAssignments))
                Return clsOperator

            Case clsRToken.typToken.ROperatorBinary
                If clsToken.lstTokens.Count < 2 Then
                    Throw New Exception("Binary operator token has " & clsToken.lstTokens.Count &
                                            " children. A binary operator must have at least 2 children (plus an optional presentation child).")
                End If

                'if object operator
                Select Case clsToken.strTxt
                    Case "$"
                        Dim strPackagePrefixNew As String = ""
                        Dim strPackageNameNew As String = ""
                        Dim lstObjectsNew As New List(Of clsRElement)

                        'add each object parameter to the object list (except last parameter)
                        Dim startPos As Integer = If(clsToken.lstTokens.Item(0).enuToken = clsRToken.typToken.RPresentation, 1, 0)
                        For intPos As Integer = startPos To clsToken.lstTokens.Count - 2
                            Dim clsTokenObject As clsRToken = clsToken.lstTokens.Item(intPos)
                            'if the first parameter is a package operator ('::'), then make this the package name for the returned element
                            If intPos = startPos AndAlso
                                    clsTokenObject.enuToken = clsRToken.typToken.ROperatorBinary AndAlso
                                    clsTokenObject.strTxt = "::" Then
                                'get the package name and any package presentation information
                                strPackageNameNew = GetTokenPackageName(clsTokenObject).strTxt
                                strPackagePrefixNew = GetPackagePrefix(clsTokenObject)
                                'get the object associated with the package, and add it to the object list
                                lstObjectsNew.Add(GetRElement(clsTokenObject.lstTokens.Item(clsTokenObject.lstTokens.Count - 1), dctAssignments))
                                Continue For
                            End If
                            lstObjectsNew.Add(GetRElement(clsTokenObject, dctAssignments))
                        Next
                        'the last item in the parameter list is the element we need to return
                        Return GetRElement(clsToken.lstTokens.Item(clsToken.lstTokens.Count - 1),
                                           dctAssignments, bBracketedNew, strPackageNameNew,
                                           strPackagePrefixNew, lstObjectsNew)

                    Case "::"
                        'the '::' operator parameter list contains:
                        ' - the presentation string (optional)
                        ' - the package name
                        ' - the element associated with the package
                        Return GetRElement(clsToken.lstTokens.Item(clsToken.lstTokens.Count - 1),
                                           dctAssignments, bBracketedNew,
                                           GetTokenPackageName(clsToken).strTxt,
                                           GetPackagePrefix(clsToken))

                    Case Else 'else if not an object or package operator, then add each parameter to the operator
                        Dim clsOperator As New clsRElementOperator(clsToken, bBracketedNew)
                        Dim startPos As Integer = If(clsToken.lstTokens.Item(0).enuToken = clsRToken.typToken.RPresentation, 1, 0)
                        For intPos As Integer = startPos To clsToken.lstTokens.Count - 1
                            clsOperator.lstParameters.Add(GetRParameter(clsToken.lstTokens.Item(intPos), dctAssignments))
                        Next
                        Return clsOperator
                End Select

            Case clsRToken.typToken.ROperatorBracket
                If clsToken.lstTokens.Count < 1 Then
                    Throw New Exception("Square bracket operator token has no children. A binary " _
                                        & "operator must have at least 1 child (plus an optional " _
                                        & "presentation child).")
                End If

                Dim clsBracketOperator As New clsRElementOperator(clsToken, bBracketedNew)
                Dim startPos As Integer = If(clsToken.lstTokens.Item(0).enuToken = clsRToken.typToken.RPresentation, 1, 0)
                For intPos As Integer = startPos To clsToken.lstTokens.Count - 1
                    clsBracketOperator.lstParameters.Add(GetRParameter(clsToken.lstTokens.Item(intPos), dctAssignments))
                Next
                Return clsBracketOperator

            Case clsRToken.typToken.RSyntacticName, clsRToken.typToken.RConstantString
                'if element has a package name or object list, then return a property element
                If Not String.IsNullOrEmpty(strPackageName) OrElse Not IsNothing(lstObjects) Then
                    Return New clsRElementProperty(clsToken, bBracketedNew, strPackageName, strPackagePrefix, lstObjects)
                End If

                'if element was assigned in a previous statement, then return an assigned element
                Dim clsStatement As clsRStatement = If(dctAssignments.ContainsKey(clsToken.strTxt), dctAssignments(clsToken.strTxt), Nothing)
                If Not IsNothing(clsStatement) Then
                    Return New clsRElementAssignable(clsToken, clsStatement, bBracketedNew)
                End If

                'else just return a regular element
                Return New clsRElement(clsToken, bBracketedNew)

            Case clsRToken.typToken.RSeparator 'a comma within a square bracket, e.g. `a[b,c]`
                'just return a regular element
                Return New clsRElement(clsToken, bBracketedNew)

            'Case clsRToken.typToken.RPresentation, clsRToken.typToken.REndStatement
            '    'if token can't be used to generate an R element then ignore
            '    Return Nothing

            Case clsRToken.typToken.REndStatement
                'TODO
                'If clsToken.lstTokens.Count > 0 AndAlso clsToken.lstTokens.Item(0).enuToken = clsRToken.typToken.RPresentation Then
                '    Return New clsRElement(clsToken.lstTokens.Item(0), bBracketedNew)
                'End If
                Return Nothing

            Case Else
                Throw New Exception("The token has an unexpected type.")
        End Select

        Throw New Exception("It should be impossible for the code to reach this point.")
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns the package name token associated with the <paramref name="clsToken"/> 
    '''             package operator. </summary>
    '''
    ''' <param name="clsToken"> Package operator ('::') token. </param>
    '''
    ''' <returns>   The package name associated with the <paramref name="clsToken"/> package 
    '''             operator. </returns>
    '''--------------------------------------------------------------------------------------------
    Private Function GetTokenPackageName(clsToken As clsRToken) As clsRToken
        If IsNothing(clsToken) Then
            Throw New ArgumentException("Cannot return a package name from an empty token.")
        End If

        If clsToken.lstTokens.Count < 2 OrElse clsToken.lstTokens.Count > 3 Then
            Throw New Exception("The package operator '::' has " & clsToken.lstTokens.Count &
                                " parameters. It must have 2 parameters (plus an optional presentation parameter).")
        End If
        Return clsToken.lstTokens.Item(clsToken.lstTokens.Count - 2)
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns the formatting prefix (spaces or comment lines) associated with the 
    '''             <paramref name="clsToken"/> package operator. If the package operator has no 
    '''             associated formatting, then returns an empty string.</summary>
    '''
    ''' <param name="clsToken"> Package operator ('::') token. </param>
    '''
    ''' <returns>   The formatting prefix (spaces or comment lines) associated with the
    '''             <paramref name="clsToken"/> package operator. </returns>
    '''--------------------------------------------------------------------------------------------
    Private Function GetPackagePrefix(clsToken As clsRToken) As String
        If IsNothing(clsToken) Then
            Throw New ArgumentException("Cannot return a package prefix from an empty token.")
        End If

        Dim clsTokenPackageName As clsRToken = GetTokenPackageName(clsToken)
        Return If(clsTokenPackageName.lstTokens.Count > 0 AndAlso clsTokenPackageName.lstTokens.Item(0).enuToken = clsRToken.typToken.RPresentation,
                                                              clsTokenPackageName.lstTokens.Item(0).strTxt, "")
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   
    ''' Returns a named parameter element constructed from the <paramref name="clsToken"/> token 
    ''' tree. The top-level element in the token tree may be:<list type="bullet"><item>
    ''' 'value' e.g. for fn(a)</item><item>
    ''' '=' e.g. for 'fn(a=1)'</item><item>
    ''' ',' e.g. for 'fn(a,b) or 'fn(a=1,b,,c,)'</item><item>
    ''' ')' indicates the end of the parameter list, returns nothing</item>
    ''' </list></summary>
    '''
    ''' <param name="clsToken">         The token tree to convert into a named parameter element. </param>
    ''' <param name="dctAssignments">   Dictionary containing all the current existing assignments.
    '''                                 The key is the name of the variable. The value is a reference
    '''                                 to the R statement that performed the assignment. </param>
    '''
    ''' <returns>   A named parameter element constructed from the <paramref name="clsToken"/> token
    '''             tree. </returns>
    '''--------------------------------------------------------------------------------------------
    Private Function GetRParameterNamed(clsToken As clsRToken, dctAssignments As Dictionary(Of String, clsRStatement)) As clsRParameter
        If IsNothing(clsToken) Then
            Throw New ArgumentException("Cannot create a named parameter from an empty token.")
        End If

        Select Case clsToken.strTxt
            Case "="
                If clsToken.lstTokens.Count < 2 Then
                    Throw New Exception("Named parameter token has " & clsToken.lstTokens.Count &
                                        " children. Named parameter must have at least 2 children (plus an optional presentation child).")
                End If

                Dim clsTokenArgumentName = clsToken.lstTokens.Item(clsToken.lstTokens.Count - 2)
                Dim clsParameter As New clsRParameter With {
                    .strArgName = clsTokenArgumentName.strTxt}
                clsParameter.clsArgValue = GetRElement(clsToken.lstTokens.Item(clsToken.lstTokens.Count - 1), dctAssignments)

                'set the parameter's formatting prefix to the prefix of the parameter name
                '    Note: if the equals sign has any formatting information then this information 
                '          will be lost.
                clsParameter.strPrefix =
                        If(clsTokenArgumentName.lstTokens.Count > 0 AndAlso
                                clsTokenArgumentName.lstTokens.Item(0).enuToken = clsRToken.typToken.RPresentation,
                                clsTokenArgumentName.lstTokens.Item(0).strTxt, "")

                Return clsParameter
            Case ","
                'if ',' is followed by a parameter name or value (e.g. 'fn(a,b)'), then return the parameter
                Try
                    'throws exception if nonpresentation child not found
                    Return GetRParameterNamed(GetChildPosNonPresentation(clsToken), dctAssignments)
                Catch ex As Exception
                    'return empty parameter (e.g. for cases like 'fn(a,)')
                    Return (New clsRParameter)
                End Try
            Case ")"
                Return Nothing
            Case Else
                Dim clsParameterNamed As New clsRParameter With {
                    .clsArgValue = GetRElement(clsToken, dctAssignments)
                }
                clsParameterNamed.strPrefix = If(clsToken.lstTokens.Count > 0 AndAlso clsToken.lstTokens.Item(0).enuToken = clsRToken.typToken.RPresentation, clsToken.lstTokens.Item(0).strTxt, "")
                Return clsParameterNamed
        End Select
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns the first child of <paramref name="clsToken"/> that is not a 
    '''             presentation token or a close bracket ')'. </summary>
    '''
    ''' <param name="clsToken"> The token tree to search for non-presentation children. </param>
    '''
    ''' <returns>   The first child of <paramref name="clsToken"/> that is not a presentation token 
    '''             or a close bracket ')'. </returns>
    '''--------------------------------------------------------------------------------------------
    Private Function GetChildPosNonPresentation(clsToken As clsRToken) As clsRToken
        If IsNothing(clsToken) Then
            Throw New ArgumentException("Cannot return a non-presentation child from an empty token.")
        End If

        'for each child token
        For Each clsTokenChild In clsToken.lstTokens
            'if token is not a presentation token or a close bracket ')', then return the token
            If Not clsTokenChild.enuToken = clsRToken.typToken.RPresentation AndAlso
                    Not (clsTokenChild.enuToken = clsRToken.typToken.RBracket AndAlso clsTokenChild.strTxt = ")") Then
                Return clsTokenChild
            End If
        Next
        Throw New Exception("Token must contain at least one non-presentation child.")
    End Function

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Returns a  parameter element constructed from the <paramref name="clsToken"/> 
    '''             token tree. </summary>
    '''
    ''' <param name="clsToken">         The token tree to convert into a parameter element. </param>
    ''' <param name="dctAssignments">   Dictionary containing all the current existing assignments.
    '''                                 The key is the name of the variable. The value is a reference
    '''                                 to the R statement that performed the assignment. </param>
    '''
    ''' <returns>   A parameter element constructed from the <paramref name="clsToken"/> token tree. </returns>
    '''--------------------------------------------------------------------------------------------
    Private Function GetRParameter(clsToken As clsRToken, dctAssignments As Dictionary(Of String, clsRStatement)) As clsRParameter
        If IsNothing(clsToken) Then
            Throw New ArgumentException("Cannot create a parameter from an empty token.")
        End If
        Return New clsRParameter With {
            .clsArgValue = GetRElement(clsToken, dctAssignments),
            .strPrefix = If(clsToken.lstTokens.Count > 0 AndAlso
                                clsToken.lstTokens.Item(0).enuToken = clsRToken.typToken.RPresentation,
                            clsToken.lstTokens.Item(0).strTxt, "")}
    End Function

End Class
