Public Class clsRElementOperator
    Inherits clsRElement
    Public clsRParameter As Boolean = False
    Public strTerminator 'only used for '[' and '[[' operators
    Public lstRParameters As New List(Of clsRParameter)

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   TODO. </summary>
    '''
    ''' <returns>   as debug string. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Overloads Function GetAsDebugString() As String
        Dim strTxt As String = "ElementOperator: " & vbLf &
                "clsRParameter: " & clsRParameter & vbLf &
                "strTerminator: " & strTerminator & vbLf

        For Each clsParameter As clsRParameter In lstRParameters
            strTxt &= clsParameter.GetAsDebugString()
        Next

        Return strTxt
    End Function

End Class
