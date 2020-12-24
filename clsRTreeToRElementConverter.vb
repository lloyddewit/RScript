'''--------------------------------------------------------------------------------------------
''' <summary> This class provides functionality to convert an R lexical analysis tree to a list 
'''           of clsRElement objects. <para>
'''           This is a 'static' class. Therefore objects of this class cannot be instantiated, 
'''           and the class cannot be inherited from. </para><para>
''' 
''' Notes: </para><para>
''' 
''' 03/11/20 </para><para>
''' - The current class structure allows an operator to be assigned to (don't know if this is 
''' valid R) </para><para>
''' - Software shall not reject any valid R (i.e. false negatives are *not* allowed).
''' - However the software is not intended to check that text is valid R. 
''' If the software populates RElement objects from invalid R code without raising an error 
''' then that is acceptable (i.e. false positives are allowed). </para><para>
''' 
''' TODO: </para><para>
''' 
''' 03/11/20
''' - Process multiline input. Newline indicates end of statement when: </para><para>
'''     1. line ends in an operator (excluding spaces and comments)  
'''     2. odd number of '()' or '[]' are open ('{}' are a special case related to key words)   
''' - Store spaces, comments and formatting new lines (also tabs?)
''' - Create automatic unit tests  
''' - Remove clsRTreeElement and parse string directly into clsRElement structure 
''' - Process ';', indicates end of statement 
''' - Process key words. Some key words (e.g. if) have a statement that returns true/false,   
''' followed by  a statement to execute (may be bracketed in '{}'
''' - Use assigned statements to manage the state of the R environment (ask David for more details)  
'''</para><para>
'''
''' 17/11/20
''' - move clsAssignment to clsRStatement  
''' - function: if function parameter is single element then   
'''                 if function *definition* then element is name
'''                 else if function *call* then element is value
''' - add 'in' operator (distinguish from 'introvert')  
''' - add key words (if, else, repeat, while, function, for)  
''' - add '{'   
''' - if comment on own line then link it to the next statement, else link to prev element  
''' - support all assignments including '=' and assign to right operators  
''' - allow new operators (e.g. '%div%), make operator list dynamic  
''' - if a package is loaded that contains new operators then add the package's operators to the operator list  
''' - '@' operator (similar to '$', store at same level)  
''' - store state of R environment   
'''       e.g. don't treat variables as text but store object they represent (e.g. value or function deinition)
'''       RSyntax already does this
''' - allow named operator params (R-Insta allows operator params to be named, but this infor is lost in script)
'''           </summary>
'''--------------------------------------------------------------------------------------------
Public NotInheritable Class clsRTreeToRElementConverter
    ''' <summary> Constructor made private to prevent instantiation </summary>
    Private Sub New()
    End Sub

    Public Shared Function GetRElements(clsInputTreeElement As clsRTreeElement) As List(Of clsRElement)
        'TODO make top-level element a newline operator and iterate through each clsInputTreeElement.lstTreeElements item
        Dim lstRElements As New List(Of clsRElement)
        lstRElements.Add(GetRElement(clsInputTreeElement.lstTreeElements.Item(0)))

        Console.WriteLine(vbCrLf + vbCrLf + "ToScript")
        Console.WriteLine(ToScript(lstRElements.Item(0)))
        Return lstRElements
    End Function

    Private Shared Function GetRElement(clsRTreeElementNew As clsRTreeElement, Optional bBracketedNew As Boolean = False) As clsRElement
        If IsNothing(clsRTreeElementNew) Then
            Return Nothing
        End If

        Select Case clsRTreeElementNew.enuTreeElementType
            Case clsRTreeElement.typeTreeElement.Text
                'if text is a round bracket, then return the bracket's child
                If clsRTreeElementNew.strText = "(" Then
                    Return GetRElement(clsRTreeElementNew.lstTreeElements.Item(0), True)
                End If

                'if text has package name or object list, then return an element with properties
                If Not IsNothing(clsRTreeElementNew.strPackageName) OrElse
                (Not IsNothing(clsRTreeElementNew.lstObjects) AndAlso clsRTreeElementNew.lstObjects.Count > 0) Then
                    Dim clsRProperty As New clsRElementProperty
                    clsRProperty.strText = clsRTreeElementNew.strText
                    clsRProperty.bBracketed = bBracketedNew
                    clsRProperty.clsAssignment = GetRElement(clsRTreeElementNew.clsAssign)
                    clsRProperty.strPackageName = clsRTreeElementNew.strPackageName

                    For Each clsRObject As clsRTreeElement In clsRTreeElementNew.lstObjects
                        clsRProperty.lstRObjects.Add(GetRElement(clsRObject))
                    Next

                    Return clsRProperty
                End If

                'else return a simple text element
                Dim clsRText As New clsRElement
                clsRText.strText = clsRTreeElementNew.strText
                clsRText.bBracketed = bBracketedNew
                clsRText.clsAssignment = GetRElement(clsRTreeElementNew.clsAssign)

                Return clsRText

            Case clsRTreeElement.typeTreeElement.RFunction
                Dim clsRFunction As New clsRElementFunction
                clsRFunction.strText = clsRTreeElementNew.strText
                clsRFunction.bBracketed = bBracketedNew
                clsRFunction.clsAssignment = GetRElement(clsRTreeElementNew.clsAssign)
                clsRFunction.strPackageName = clsRTreeElementNew.strPackageName

                For Each clsRObject As clsRTreeElement In clsRTreeElementNew.lstObjects
                    clsRFunction.lstRObjects.Add(GetRElement(clsRObject))
                Next

                'if the function has only 1 parameter then it has no 'param' element 
                '    (this is because function 'param' elements are added when processing comma operators)
                If clsRTreeElementNew.lstTreeElements.Count = 1 Then
                    Dim clsRParameter As New clsRParameterNamed
                    If clsRTreeElementNew.lstTreeElements.Item(0).strText = "=" Then
                        clsRParameter.strArgumentName = clsRTreeElementNew.lstTreeElements.Item(0).lstTreeElements.Item(0).lstTreeElements.Item(0).strText
                        clsRParameter.clsArgValue = GetRElement(clsRTreeElementNew.lstTreeElements.Item(0).lstTreeElements.Item(1).lstTreeElements.Item(0))
                    Else
                        clsRParameter.clsArgValue = GetRElement(clsRTreeElementNew.lstTreeElements.Item(0))
                    End If
                    clsRFunction.lstRParameters.Add(clsRParameter)
                Else
                    For Each clsTreeParam As clsRTreeElement In clsRTreeElementNew.lstTreeElements
                        Dim clsRParameter As New clsRParameterNamed
                        If clsTreeParam.lstTreeElements.Item(0).strText = "=" Then
                            clsRParameter.strArgumentName = clsTreeParam.lstTreeElements.Item(0).lstTreeElements.Item(0).lstTreeElements.Item(0).strText
                            clsRParameter.clsArgValue = GetRElement(clsTreeParam.lstTreeElements.Item(0).lstTreeElements.Item(1).lstTreeElements.Item(0))
                        Else
                            clsRParameter.clsArgValue = GetRElement(clsTreeParam.lstTreeElements.Item(0))
                        End If
                        clsRFunction.lstRParameters.Add(clsRParameter)
                    Next
                End If
                Return clsRFunction

            Case clsRTreeElement.typeTreeElement.ROperatorUnaryLeft
                Dim clsROperator As New clsRElementOperator
                clsROperator.strText = clsRTreeElementNew.strText
                clsROperator.bBracketed = bBracketedNew
                clsROperator.clsAssignment = GetRElement(clsRTreeElementNew.clsAssign)
                For Each clsTreeParam As clsRTreeElement In clsRTreeElementNew.lstTreeElements
                    Dim clsRParameter As New clsRParameter
                    clsRParameter.clsArgValue = GetRElement(clsTreeParam.lstTreeElements.Item(0))
                    clsROperator.lstRParameters.Add(clsRParameter)
                Next
                Return clsROperator

            Case clsRTreeElement.typeTreeElement.ROperatorUnaryRight
                Dim clsROperator As New clsRElementOperator
                clsROperator.strText = clsRTreeElementNew.strText
                clsROperator.bBracketed = bBracketedNew
                clsROperator.clsAssignment = GetRElement(clsRTreeElementNew.clsAssign)
                For Each clsTreeParam As clsRTreeElement In clsRTreeElementNew.lstTreeElements
                    Dim clsRParameter As New clsRParameter
                    clsRParameter.clsArgValue = GetRElement(clsTreeParam.lstTreeElements.Item(0))
                    clsROperator.lstRParameters.Add(clsRParameter)
                Next
                clsROperator.bFirstParamOnRight = True
                Return clsROperator

            Case clsRTreeElement.typeTreeElement.ROperatorBinary
                Dim clsROperator As New clsRElementOperator
                clsROperator.strText = clsRTreeElementNew.strText
                clsROperator.bBracketed = bBracketedNew
                clsROperator.clsAssignment = GetRElement(clsRTreeElementNew.clsAssign)
                For Each clsTreeParam As clsRTreeElement In clsRTreeElementNew.lstTreeElements
                    Dim clsRParameter As New clsRParameter
                    clsRParameter.clsArgValue = GetRElement(clsTreeParam.lstTreeElements.Item(0))
                    clsROperator.lstRParameters.Add(clsRParameter)
                Next
                Return clsROperator

            Case clsRTreeElement.typeTreeElement.ROperatorBracket
                Dim clsROperator As New clsRElementOperator
                clsROperator.strText = clsRTreeElementNew.strText
                clsROperator.bBracketed = bBracketedNew
                clsROperator.clsAssignment = GetRElement(clsRTreeElementNew.clsAssign)
                For Each clsTreeParam As clsRTreeElement In clsRTreeElementNew.lstTreeElements
                    Dim clsRParameter As New clsRParameter
                    clsRParameter.clsArgValue = GetRElement(clsTreeParam.lstTreeElements.Item(0))
                    clsROperator.lstRParameters.Add(clsRParameter)
                Next
                Select Case clsROperator.strText
                    Case "["
                        clsROperator.strTerminator = "]"
                    Case "[["
                        clsROperator.strTerminator = "]]"
                End Select
                Return clsROperator
        End Select
        'TODO developer error
        Return Nothing 'should never reach this point
    End Function

    Private Shared Function ToScript(clsRElementNew As Object) As String
        If IsNothing(clsRElementNew) Then
            Return Nothing
        End If

        Dim strScript As String = If(clsRElementNew.bBracketed, "(", "")
        strScript &= If(IsNothing(clsRElementNew.clsAssignment), "", clsRElementNew.clsAssignment.strText + "<-")

        Select Case clsRElementNew.GetType()
            Case GetType(clsRElementFunction)
                strScript &= If(IsNothing(clsRElementNew.strPackageName), "", clsRElementNew.strPackageName + "::")
                If Not IsNothing(clsRElementNew.lstRObjects) AndAlso clsRElementNew.lstRObjects.Count > 0 Then
                    For Each clsRObject In clsRElementNew.lstRObjects
                        strScript &= ToScript(clsRObject)
                        strScript &= "$"
                    Next
                End If
                strScript &= clsRElementNew.strText
                If Not IsNothing(clsRElementNew.lstRParameters) Then
                    strScript &= "("
                    Dim bPrefixComma As Boolean = False
                    For Each clsRParameter In clsRElementNew.lstRParameters
                        strScript &= If(bPrefixComma, ",", "")
                        bPrefixComma = True
                        strScript &= If(IsNothing(clsRParameter.strArgumentName), "", clsRParameter.strArgumentName + "=")
                        strScript &= ToScript(clsRParameter.clsArgValue)
                    Next
                    strScript &= ")"
                End If
            Case GetType(clsRElementProperty)
                strScript &= If(IsNothing(clsRElementNew.strPackageName), "", clsRElementNew.strPackageName + "::")
                If Not IsNothing(clsRElementNew.lstRObjects) AndAlso clsRElementNew.lstRObjects.Count > 0 Then
                    For Each clsRObject In clsRElementNew.lstRObjects
                        strScript &= ToScript(clsRObject)
                        strScript &= "$"
                    Next
                End If
                strScript &= clsRElementNew.strText
            Case GetType(clsRElementOperator)
                Dim bPrefixOperator As Boolean = If(clsRElementNew.bFirstParamOnRight, True, False)
                For Each clsRParameter In clsRElementNew.lstRParameters
                    strScript &= If(bPrefixOperator, clsRElementNew.strText, "")
                    bPrefixOperator = True
                    strScript &= ToScript(clsRParameter.clsArgValue)
                Next
                strScript &= If(clsRElementNew.lstRParameters.Count = 1 AndAlso Not clsRElementNew.bFirstParamOnRight, clsRElementNew.strText, "")
                strScript &= clsRElementNew.strTerminator
            Case GetType(clsRElement)
                strScript &= clsRElementNew.strText
        End Select
        strScript &= If(clsRElementNew.bBracketed, ")", "")
        Return strScript
    End Function

End Class
