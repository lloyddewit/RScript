Public Class clsRElementFunction
    Inherits clsRElementProperty

    Public lstRParameters As New List(Of clsRParameterNamed)

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   TODO. </summary>
    '''
    ''' <returns>   as debug string. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Overloads Function GetAsDebugString() As String
        Dim strTxt As String = "ElementFunction: " & vbLf

        For Each clsParameter As clsRParameterNamed In lstRParameters
            strTxt &= clsParameter.GetAsDebugString()
        Next

        Return strTxt
    End Function

End Class
