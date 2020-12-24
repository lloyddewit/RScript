'''--------------------------------------------------------------------------------------------
''' <summary> This class provides functionality to convert an R lexical analysis tree to a list 
'''           of RCodeStructure objects. <para>
'''           This is a 'static' class. Therefore objects of this class cannot be instantiated, 
'''           and the class cannot be inherited from. </para></summary>
'''--------------------------------------------------------------------------------------------
Public NotInheritable Class clsRTreeToRCodeConverter
    ''' <summary> Constructor made private to prevent instantiation </summary>
    Private Sub New()
    End Sub

    Public Shared Function GetListRCodeStructures(clsInputTreeElement As clsRTreeElement) As List(Of String)
        'Public Shared Function GetListRCodeStructures(clsRTreeElement As clsRTreeElement) As List(Of RCodeStructure)
        ' the most important data members we need to populate are:
        '
        '        RCodeStructure
        '        clsParameters 'The list of parameters associated with this R code
        '        I didn't look at the `RCodeStructure` assignment data members yet
        '        These will include strAssignTo and bToBeAssigned
        '        
        '        RFunction
        '        strRCommand, strPackageName
        '
        ' ROperator
        '        strOperation, bBrackets, bAllBrackets
        '
        ' RParameter
        '        strArgumentName, strArgumentValue, clsArgumentCodeStructure, bIsFunction, bIsOperator, bIsString, iPosition, bIncludeArgumentName
        '
        Dim lstRCodeStructures As New List(Of String)
        Dim intPos As Integer = 0

        Do While intPos < clsInputTreeElement.lstTreeElements.Count
            lstRCodeStructures.Add(clsInputTreeElement.lstTreeElements.Item(intPos).strText)
            intPos += 1
        Loop

        Return lstRCodeStructures
    End Function
End Class
