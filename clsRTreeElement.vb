Imports System.ComponentModel.Design

Public Class clsRTreeElement

    'TODO make public data members below read-only properties?
    ' 
    Public Enum typeTreeElement
        Text
        RFunction
        ROperatorUnaryLeft
        ROperatorUnaryRight
        ROperatorBinary
        ROperatorBracket
        RSeparator
        RParameter
    End Enum

    Public strText As String
    Public enuTreeElementType As typeTreeElement
    Public lstTreeElements As New List(Of clsRTreeElement)

    Public strPackageName As String
    Public lstObjects As New List(Of clsRTreeElement)
    Public clsAssign As clsRTreeElement

    'TODO missing unary '+' & '-' (should be just after '^') - make map of operator and bBinary?
    'TODO missing assignment '=' (should be at end just after '<<-')
    'TODO add package defined operators (e.g. '%>%)
    'TODO according to ref below, some operators group left to right, others group right to left
    'TODO handle CrLf (if statement complete then marks end of statement, else ignored) - do we need to remember to recreate exact appearance with 'ToScript?
    'TODO handle ';' (marks end of stetment, alternative to CrLf)
    'TODO is it possible for packages to be nested (e.g. 'p1::p1_1::f()')?
    'TODO handle function definitions and blocks ('{}'), handle special function argument '...'
    'TODO handle 'x + y can equivalently be written `+`(x, y). Notice that since ‘+’ is a non-standard function name, it needs to be quoted'
    'TODO keep comments for 'ToScript' or ignore comments? ' Any text from a # character to the end of the line is taken to be a comment, unless the # character is inside a quoted string'

    ''' <summary> A complete list of R operators in precedence order.
    ''' See https://cran.r-project.org/doc/manuals/r-release/R-lang.html#Operators
    ''' Note that '=' as a parameter assignment, is not strictly an R operator but this class 
    ''' treats it as an R operator. </summary>
    Private ReadOnly arrROperators() As String = {"::", "$", "@", "^", "%%", "%/%",
        "%*%", "%o%", "%x%", "%in%", "/", "*", "+", "-", ">>=", "<<=", "==", "!=", "!", "&",
        "&&", "|", "||", "~", "->", "->>", "<-", "<<-", "="}

    Private ReadOnly arrRSeperators() As String = {","}

    Private ReadOnly arrROperatorBrackets() As String = {"[", "]", "[[", "]]"}

    Private ReadOnly arrPartialROperators() As String = {":", "%", "%/",
        "%*", "%o", "%x", "%i", "%in", ">", ">>", "<", "<<", "->", "<", "<<"}

    Private Sub New(newEnumTreeType As typeTreeElement)
        enuTreeElementType = newEnumTreeType
    End Sub

    Private Sub New(strTextNew As String, strTextPrev As String, strTextNext As String)
        strText = strTextNew
        If Not isRecognized(strText) Then
            enuTreeElementType = typeTreeElement.Text
        ElseIf arrRSeperators.Contains(strText) Then
            enuTreeElementType = typeTreeElement.RSeparator
        ElseIf arrROperatorBrackets.Contains(strText) Then
            enuTreeElementType = typeTreeElement.ROperatorBracket
        ElseIf IsNothing(strTextPrev) OrElse arrROperators.Contains(strTextPrev) OrElse arrRSeperators.Contains(strTextPrev) Then
            enuTreeElementType = typeTreeElement.ROperatorUnaryRight
        ElseIf IsNothing(strTextNext) OrElse arrRSeperators.Contains(strTextNext) Then
            enuTreeElementType = typeTreeElement.ROperatorUnaryLeft
        Else
            enuTreeElementType = typeTreeElement.ROperatorBinary
        End If
    End Sub

    Public Sub New(strInput As String, Optional bFromRCode As Boolean = False)
        If IsNothing(strInput) Then
            Exit Sub
        ElseIf Not bFromRCode Then
            strText = strInput
            enuTreeElementType = If(isRecognized(strInput), typeTreeElement.ROperatorBinary, typeTreeElement.Text)
            Exit Sub
        End If

        'TODO how should we deal with newlines?
        '  - assume that newlines always indicate the end of the current R statement (not always strictly true)?
        '  - Separate strInput into a list of newline terminated strings and call CreateTreeFromString for each?
        CreateTreeFromString(strInput)

        'TODO debug only
        Console.WriteLine(vbCrLf + vbCrLf + "CreateTreeFromString")
        WriteTree(Me)
        Console.WriteLine("")

        RestructureTreeForBracket(Me.CloneMe)
        'TODO debug only
        Console.WriteLine(vbCrLf + vbCrLf + "RestructureTreeForBracket")
        WriteTree(Me)
        Console.WriteLine("")

        RestructureTreeForComma(Me.CloneMe)
        'TODO debug only
        Console.WriteLine(vbCrLf + vbCrLf + "RestructureTreeForComma")
        WriteTree(Me)
        Console.WriteLine("")

        For Each strROperator As String In arrROperators
            If strInput.Contains(strROperator) Then
                RestructureTreeForOperator(Me.CloneMe, strROperator)

                'TODO debug only
                Console.WriteLine(vbCrLf + vbCrLf + "RestructureTreeForOperator " & strROperator)
                WriteTree(Me)
                Console.WriteLine("")
            End If
        Next

        RestructureTreeForPackageName(Me.CloneMe)
        'TODO debug only
        Console.WriteLine(vbCrLf + vbCrLf + "RestructureTreeForPackageName")
        WriteTree(Me)
        Console.WriteLine("")

        RestructureTreeForDollar(Me.CloneMe)
        'TODO debug only
        Console.WriteLine(vbCrLf + vbCrLf + "RestructureTreeForDollar")
        WriteTree(Me)
        Console.WriteLine("")

        RestructureTreeForAssignment(Me.CloneMe)
        'TODO debug only
        Console.WriteLine(vbCrLf + vbCrLf + "RestructureTreeForAssignment")
        WriteTree(Me)
        Console.WriteLine("")

    End Sub

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   Creates tree from string.TODO </summary>
    '''
    ''' <param name="strInput"> The input. </param>
    ''' <param name="intPos">   [in,out] (Optional) The int position. </param>
    '''--------------------------------------------------------------------------------------------
    Private Sub CreateTreeFromString(strInput As String, Optional ByRef intPos As Integer = 0)
        Dim strTextPrev As String = Nothing
        Dim strTextNew As String = Nothing
        Dim strCharNew As String
        Do Until intPos = Len(strInput)
            strCharNew = strInput(intPos)
            intPos += 1

            'if there are an odd number of double quotes then it means we are inside a string literal (e.g. inside '" - "')
            'Note:  this check is needed to ensure that special characters (e.g. brackets or operator characters) inside string literals are ignored
            If Not IsNothing(strTextNew) AndAlso strTextNew.Where(Function(c) c = Chr(34)).Count Mod 2 Then
                'append characters in the literal up to and including the 2nd double quote
                strTextNew &= strCharNew
                Continue Do
            End If

            Select Case strCharNew
                Case "(" 'build new tree from next char up to next ')' (or string end)
                    'if there is no text preceding '('. This may happen if '(' is first char in string, or '(' is part of nested brackets
                    If IsNothing(strTextNew) Then
                        'add a new '(' element that contains the children TODO tidy up
                        Dim clsTempRTreeElement As New clsRTreeElement("(")
                        clsTempRTreeElement.CreateTreeFromString(strInput, intPos)
                        lstTreeElements.Add(clsTempRTreeElement)
                        'if text preceding '(' is not recognized text then text preceding '(' must be a function name
                    ElseIf Not isRecognized(strTextNew) Then
                        Dim clsTempRTreeElement As New clsRTreeElement(strTextNew)
                        clsTempRTreeElement.CreateTreeFromString(strInput, intPos)
                        clsTempRTreeElement.enuTreeElementType = typeTreeElement.RFunction
                        lstTreeElements.Add(clsTempRTreeElement)
                    Else 'else if '(' is preceded by an operator
                        ' add the operator element
                        Dim clsTempRTreeElementOperator As New clsRTreeElement(strTextNew, strTextPrev, strCharNew)
                        lstTreeElements.Add(clsTempRTreeElementOperator)

                        'add a new '(' element that contains the children
                        Dim clsTempRTreeElement = New clsRTreeElement("(")
                        clsTempRTreeElement.CreateTreeFromString(strInput, intPos)
                        lstTreeElements.Add(clsTempRTreeElement)
                    End If

                    strTextPrev = "(" 'remember that this was the last text processed (needed to detect if operator is unary or binary)
                    strTextNew = Nothing 'prevent text being added to tree more than once 
                Case ")"
                    If IsNothing(strTextNew) Then
                        Exit Sub
                    End If

                    lstTreeElements.Add(New clsRTreeElement(strTextNew, strTextPrev, strCharNew))
                    Exit Sub
                Case "[" 'build new tree from next char up to next ']' (or string end)
                    ' if left-hand operand of bracket exists, then add it to tree (e.g. for 'a[1]', add 'a')
                    If Not IsNothing(strTextNew) Then
                        Dim clsTempRTreeElementText As New clsRTreeElement(strTextNew, strTextPrev, strCharNew)
                        lstTreeElements.Add(clsTempRTreeElementText)
                    End If

                    Dim strBracketOperator As String = "["
                    'if bracket is part of double-bracket operator
                    If (intPos < Len(strInput)) AndAlso strInput(intPos) = "[" Then
                        strBracketOperator = "[["
                        intPos += 1
                    End If

                    'add a new '[' element that contains the children
                    Dim clsTempRTreeElement = New clsRTreeElement(strBracketOperator, Nothing, Nothing)
                    clsTempRTreeElement.CreateTreeFromString(strInput, intPos)
                    lstTreeElements.Add(clsTempRTreeElement)

                    strTextPrev = strBracketOperator 'remember that this was the last text processed (needed to detect if operator is unary or binary)
                    strTextNew = Nothing 'prevent text being added to tree more than once 
                Case "]"
                    'if bracket is part of double-bracket operator
                    If (intPos < Len(strInput)) AndAlso strInput(intPos) = "]" Then
                        intPos += 1
                    End If

                    If IsNothing(strTextNew) Then
                        Exit Sub
                    End If

                    lstTreeElements.Add(New clsRTreeElement(strTextNew, strTextPrev, strCharNew))
                    Exit Sub
                Case vbCr
                    If IsNothing(strTextNew) Then
                        Exit Sub
                    End If

                    lstTreeElements.Add(New clsRTreeElement(strTextNew, strTextPrev, strCharNew))
                    Exit Sub
                Case Else
                    'if new char is a standalone operator (i.e. it's not joined to previous text as part of a multi-char operator)
                    'e.g. new char is '-' but is not part of '<-'
                    'OR 
                    '  previous element is an operator, and new char is not an operator
                    If (isRecognized(strCharNew) AndAlso (Not isRecognized(strTextNew & strCharNew))) Or
                       ((Not isRecognized(strCharNew)) AndAlso isRecognized(strTextNew)) Then
                        'add 'strTextNew' element to list
                        lstTreeElements.Add(New clsRTreeElement(strTextNew, strTextPrev, strCharNew))
                        strTextPrev = strTextNew
                        strTextNew = strCharNew.Trim()
                    Else
                        strTextNew &= strCharNew.Trim()
                    End If
            End Select
        Loop
        'if text remains that is not yet added to tree (occurs when string does not end in ')')
        If Not IsNothing(strTextNew) Then
            lstTreeElements.Add(New clsRTreeElement(strTextNew, strTextPrev, Nothing))
        End If
    End Sub

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   TBD </summary>
    '''
    ''' <param name="clsInputTreeElement">  The cls input tree element. </param>
    ''' <param name="strROperator">            The operator. </param>
    '''--------------------------------------------------------------------------------------------
    Private Sub RestructureTreeForBracket(clsInputTreeElement As clsRTreeElement)
        Dim intPos As Integer = 0
        Dim clsTreeElementPrev As clsRTreeElement
        Dim clsTreeElementBracket As clsRTreeElement

        lstTreeElements = New List(Of clsRTreeElement)

        Do While intPos < clsInputTreeElement.lstTreeElements.Count
            'if current tree element not the open bracket then keep adding previous element to main list
            If Not arrROperatorBrackets.Contains(clsInputTreeElement.lstTreeElements.Item(intPos).strText) Then
                If Not IsNothing(clsTreeElementPrev) Then
                    lstTreeElements.Add(clsTreeElementPrev)
                End If
                clsTreeElementPrev = clsInputTreeElement.lstTreeElements.Item(intPos).CloneMe
                clsTreeElementPrev.RestructureTreeForBracket(clsTreeElementPrev.CloneMe)
                intPos += 1
                Continue Do
            End If
            'else current tree element is the bracket

            'create new bracket operator
            clsTreeElementBracket = New clsRTreeElement(clsInputTreeElement.lstTreeElements.Item(intPos).strText, Nothing, Nothing)

            'if the previous element not yet added
            If Not IsNothing(clsTreeElementPrev) Then
                'add it to a new param element
                Dim clsTreeElementParamLeft As New clsRTreeElement(typeTreeElement.RParameter)
                clsTreeElementParamLeft.lstTreeElements.Add(clsTreeElementPrev.CloneMe)
                'clsTreeElementPrev = Nothing

                'add the param element to bracket operator
                clsTreeElementBracket.lstTreeElements.Add(clsTreeElementParamLeft)
            End If

            'add the next elements as a new param for the bracket operator until close bracket or end of list
            Dim clsTreeElementParamRight As New clsRTreeElement(typeTreeElement.RParameter)
            Dim clsTreeElementTmp As clsRTreeElement = clsInputTreeElement.lstTreeElements.Item(intPos).CloneMe

            clsTreeElementTmp.RestructureTreeForBracket(clsTreeElementTmp.CloneMe)
            For Each clsTreeElementTmp2 As clsRTreeElement In clsTreeElementTmp.lstTreeElements
                clsTreeElementParamRight.lstTreeElements.Add(clsTreeElementTmp2)
            Next

            'add the param element to the bracket operator
            clsTreeElementBracket.lstTreeElements.Add(clsTreeElementParamRight)
            'add the bracket operator to the main list
            'lstTreeElements.Add(clsTreeElementBracket.CloneMe)
            clsTreeElementPrev = clsTreeElementBracket.CloneMe
            intPos += 1
        Loop

        'if the previous element not yet added to the main list, then add it
        If Not IsNothing(clsTreeElementPrev) Then
            lstTreeElements.Add(clsTreeElementPrev)
        End If

    End Sub

    Private Sub RestructureTreeForComma(clsInputTreeElement As clsRTreeElement)
        Dim intPos As Integer = 0
        Dim lstTreeElementsTemp As New List(Of clsRTreeElement)
        Dim bCommaFound As Boolean = False

        lstTreeElements = New List(Of clsRTreeElement)

        Do While intPos < clsInputTreeElement.lstTreeElements.Count
            'if current tree element not a comma then keep adding elements to temporary list
            If Not clsInputTreeElement.lstTreeElements.Item(intPos).strText = "," Then
                Dim clsNewTreeElement As clsRTreeElement = clsInputTreeElement.lstTreeElements.Item(intPos)
                clsNewTreeElement.RestructureTreeForComma(clsNewTreeElement.CloneMe)
                lstTreeElementsTemp.Add(clsNewTreeElement.CloneMe)
                intPos += 1
                Continue Do
            End If
            'else current tree element is a comma
            bCommaFound = True

            'create new param element
            Dim clsTreeElementParam As New clsRTreeElement(typeTreeElement.RParameter)

            'if the temp list contains any tree elements that haven't yet been added
            If lstTreeElementsTemp.Count > 0 Then
                'add elements in temp list to the param element
                For Each clsTreeElement As clsRTreeElement In lstTreeElementsTemp
                    clsTreeElementParam.lstTreeElements.Add(clsTreeElement)
                Next
            Else 'else if nothing precedes the comma (e.g. occurs with 'f1(,b)', 'f2(,,)' etc.)
                'add an empty element to the parameter
                clsTreeElementParam.lstTreeElements.Add(New clsRTreeElement(typeTreeElement.Text))
            End If

            'add the param element to the main element list
            lstTreeElements.Add(clsTreeElementParam)
            lstTreeElementsTemp = New List(Of clsRTreeElement)

            'if nothing follows the comma (e.g. occurs with 'f1(a,)', 'f2(,,)' etc.)
            If intPos = clsInputTreeElement.lstTreeElements.Count - 1 Then
                'add an empty element to the temporary list
                '(this will be added as a new param in the main list when this loop exits)
                lstTreeElementsTemp.Add(New clsRTreeElement(typeTreeElement.Text))
            End If

            intPos += 1
        Loop

        'if the temporary list contains any tree elements that still need to be added to a param element, then add them
        If bCommaFound Then
            Dim clsTreeElementParam As New clsRTreeElement(typeTreeElement.RParameter)
            For Each clsTreeElement As clsRTreeElement In lstTreeElementsTemp
                clsTreeElementParam.lstTreeElements.Add(clsTreeElement)
            Next
            lstTreeElements.Add(clsTreeElementParam)
            Exit Sub
        End If

        'if the temporary list still contains any tree elements that haven't yet been added, then add them
        For Each clsTreeElement As clsRTreeElement In lstTreeElementsTemp
            lstTreeElements.Add(clsTreeElement)
        Next

    End Sub

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   TBD </summary>
    '''
    ''' <param name="clsInputTreeElement">  The cls input tree element. </param>
    ''' <param name="strROperator">            The operator. </param>
    '''--------------------------------------------------------------------------------------------
    Private Sub RestructureTreeForOperator(clsInputTreeElement As clsRTreeElement, strROperator As String)
        Dim intPos As Integer = 0
        Dim clsTreeElementPrev As clsRTreeElement
        Dim clsTreeElementOp As clsRTreeElement

        lstTreeElements = New List(Of clsRTreeElement)

        Do While intPos < clsInputTreeElement.lstTreeElements.Count
            'if current tree element not the operator then keep adding previous element to main list
            If Not clsInputTreeElement.lstTreeElements.Item(intPos).strText = strROperator Then
                If Not IsNothing(clsTreeElementOp) Then
                    lstTreeElements.Add(clsTreeElementOp.CloneMe)
                    clsTreeElementOp = Nothing 'prevent operator being added to list twice
                End If
                If Not IsNothing(clsTreeElementPrev) Then
                    lstTreeElements.Add(clsTreeElementPrev)
                End If
                clsTreeElementPrev = clsInputTreeElement.lstTreeElements.Item(intPos).CloneMe
                clsTreeElementPrev.RestructureTreeForOperator(clsTreeElementPrev.CloneMe, strROperator)
                intPos += 1
                Continue Do
            End If
            'else current tree element is the operator

            'if first param is on right, then continue looping (e.g. occurs for '!' with f1(!a)', 'f2(b~!c)' etc.)
            If clsInputTreeElement.lstTreeElements.Item(intPos).enuTreeElementType = typeTreeElement.ROperatorUnaryRight Then
                clsTreeElementOp = clsInputTreeElement.lstTreeElements.Item(intPos).CloneMe

                'if the previous element not yet added to the main list, then add it
                If Not IsNothing(clsTreeElementPrev) Then
                    lstTreeElements.Add(clsTreeElementPrev.CloneMe)
                    clsTreeElementPrev = Nothing
                End If
            ElseIf Not IsNothing(clsTreeElementPrev) Then 'if the previous element not yet added
                'add it to a new param element
                Dim clsTreeElementParamLeft As New clsRTreeElement(typeTreeElement.RParameter)
                clsTreeElementParamLeft.lstTreeElements.Add(clsTreeElementPrev.CloneMe)
                clsTreeElementPrev = Nothing

                'add the param element to a new operator element
                clsTreeElementOp = clsInputTreeElement.lstTreeElements.Item(intPos).CloneMe
                clsTreeElementOp.lstTreeElements.Add(clsTreeElementParamLeft)
            End If

            'if right element also belongs to this operator
            If Not clsInputTreeElement.lstTreeElements.Item(intPos).enuTreeElementType = typeTreeElement.ROperatorUnaryLeft Then
                intPos += 1
                If intPos >= clsInputTreeElement.lstTreeElements.Count Then
                    Exit Do
                End If

                'add the next element to a new param element, and add this to the operator element also
                Dim clsTreeElementNext As clsRTreeElement = clsInputTreeElement.lstTreeElements.Item(intPos).CloneMe
                clsTreeElementNext.RestructureTreeForOperator(clsTreeElementNext.CloneMe, strROperator)

                Dim clsTreeElementParamRight As New clsRTreeElement(typeTreeElement.RParameter)
                clsTreeElementParamRight.lstTreeElements.Add(clsTreeElementNext)
                clsTreeElementOp.lstTreeElements.Add(clsTreeElementParamRight) 'add the param element to the operator element
            End If
            intPos += 1
        Loop

        If Not IsNothing(clsTreeElementOp) Then
            lstTreeElements.Add(clsTreeElementOp.CloneMe)
        End If

        'if the previous element not yet added to the main list, then add it
        If Not IsNothing(clsTreeElementPrev) Then
            lstTreeElements.Add(clsTreeElementPrev.CloneMe)
        End If

    End Sub

    Private Sub RestructureTreeForPackageName(clsInputTreeElement As clsRTreeElement)
        Dim intPos As Integer = 0
        Dim lstTreeElementsTemp As New List(Of clsRTreeElement)

        lstTreeElements = New List(Of clsRTreeElement)

        Do While intPos < clsInputTreeElement.lstTreeElements.Count
            Dim clsTreeElementNew As clsRTreeElement = clsInputTreeElement.lstTreeElements.Item(intPos)
            'if current tree element not a package operator then keep adding elements to temporary list
            If Not clsTreeElementNew.strText = "::" Then
                clsTreeElementNew.RestructureTreeForPackageName(clsTreeElementNew.CloneMe)
                lstTreeElements.Add(clsTreeElementNew.CloneMe)
                intPos += 1
                Continue Do
            End If
            'else current tree element is a package operator

            Dim strPackageNameNew As String = clsTreeElementNew.lstTreeElements.Item(0).lstTreeElements.Item(0).strText
            clsTreeElementNew = clsTreeElementNew.lstTreeElements.Item(1).lstTreeElements.Item(0).CloneMe
            clsTreeElementNew.strPackageName = strPackageNameNew
            clsTreeElementNew.RestructureTreeForPackageName(clsTreeElementNew.CloneMe)
            lstTreeElements.Add(clsTreeElementNew.CloneMe)
            intPos += 1
        Loop
    End Sub

    Private Sub RestructureTreeForDollar(clsInputTreeElement As clsRTreeElement)
        Dim intPos As Integer = 0
        Dim lstTreeElementsTemp As New List(Of clsRTreeElement)

        lstTreeElements = New List(Of clsRTreeElement)

        Do While intPos < clsInputTreeElement.lstTreeElements.Count
            Dim clsTreeElementNew As clsRTreeElement = clsInputTreeElement.lstTreeElements.Item(intPos)
            'if current tree element not a dollar operator then keep adding elements to temporary list
            If Not clsTreeElementNew.strText = "$" Then
                clsTreeElementNew.RestructureTreeForDollar(clsTreeElementNew.CloneMe)
                lstTreeElements.Add(clsTreeElementNew.CloneMe)
                intPos += 1
                Continue Do
            End If
            'else current tree element is a $ operator

            'get the last item in the $ operator's param list (this contains the variable preceded by the object(s))
            Dim clsTreeElementVar As clsRTreeElement = clsTreeElementNew.lstTreeElements.Item(clsTreeElementNew.lstTreeElements.Count - 1).lstTreeElements.Item(0).CloneMe

            'for each $ operator parameter
            For intParamPos As Integer = 0 To clsTreeElementNew.lstTreeElements.Count - 2
                clsTreeElementVar.lstObjects.Add(clsTreeElementNew.lstTreeElements.Item(intParamPos).lstTreeElements.Item(0).CloneMe)
            Next

            clsTreeElementNew.RestructureTreeForDollar(clsTreeElementVar.CloneMe)
            lstTreeElements.Add(clsTreeElementVar.CloneMe)
            intPos += 1
        Loop
    End Sub

    Private Sub RestructureTreeForAssignment(clsInputTreeElement As clsRTreeElement)
        Dim intPos As Integer = 0
        Dim lstTreeElementsTemp As New List(Of clsRTreeElement)

        lstTreeElements = New List(Of clsRTreeElement)

        Do While intPos < clsInputTreeElement.lstTreeElements.Count
            Dim clsTreeElementNew As clsRTreeElement = clsInputTreeElement.lstTreeElements.Item(intPos)
            'if current tree element not a package operator then keep adding elements to temporary list
            If Not clsTreeElementNew.strText = "<-" Then
                clsTreeElementNew.RestructureTreeForAssignment(clsTreeElementNew.CloneMe)
                lstTreeElements.Add(clsTreeElementNew.CloneMe)
                intPos += 1
                Continue Do
            End If
            'else current tree element is an assignment operator

            Dim clsTreeElementAssign As clsRTreeElement = clsTreeElementNew.lstTreeElements.Item(0).lstTreeElements.Item(0).CloneMe
            clsTreeElementAssign.RestructureTreeForAssignment(clsTreeElementAssign.CloneMe)

            clsTreeElementNew = clsTreeElementNew.lstTreeElements.Item(1).lstTreeElements.Item(0).CloneMe
            clsTreeElementNew.clsAssign = clsTreeElementAssign
            clsTreeElementNew.RestructureTreeForAssignment(clsTreeElementNew.CloneMe)
            lstTreeElements.Add(clsTreeElementNew.CloneMe)
            intPos += 1
        Loop
    End Sub

    Private Function isRecognized(strTextToCheck As String) As Boolean
        If arrROperators.Contains(strTextToCheck) OrElse
            arrROperatorBrackets.Contains(strTextToCheck) OrElse
            arrRSeperators.Contains(strTextToCheck) OrElse
            arrPartialROperators.Contains(strTextToCheck) Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function CloneMe() As clsRTreeElement
        Dim clsNewTreeElement = New clsRTreeElement(enuTreeElementType)
        clsNewTreeElement.strText = strText
        clsNewTreeElement.strPackageName = strPackageName
        clsNewTreeElement.clsAssign = If(IsNothing(clsAssign), Nothing, clsAssign.CloneMe)

        For Each clsChildTreeElement As clsRTreeElement In lstTreeElements
            clsNewTreeElement.lstTreeElements.Add(clsChildTreeElement.CloneMe)
        Next

        For Each clsChildTreeElement As clsRTreeElement In lstObjects
            clsNewTreeElement.lstObjects.Add(clsChildTreeElement.CloneMe)
        Next

        Return clsNewTreeElement
    End Function

    'TODO just for debugging
    Private Sub WriteTree(tree As clsRTreeElement, Optional indent As String = "")
        For Each element In tree.lstTreeElements
            Console.Write(indent + element.strText + " " + element.enuTreeElementType.ToString)
            Console.Write(If(IsNothing(element.strPackageName), "", " " + element.strPackageName + "::"))
            Console.Write(If(IsNothing(element.clsAssign), "", " " + element.clsAssign.strText + "<-"))
            For Each clsObject As clsRTreeElement In element.lstObjects
                Console.Write(" " + clsObject.strText + "$ " + clsObject.enuTreeElementType.ToString)
            Next
            Console.WriteLine()
            WriteTree(element, indent + "..")
        Next
    End Sub
End Class

