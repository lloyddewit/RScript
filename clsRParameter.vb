Public Class clsRParameter
    Public clsArgValue As clsRElement
    Public clsPresentation As clsRElementPresentation

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   TODO. </summary>
    '''
    ''' <returns>   as debug string. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Function GetAsDebugString() As String
        Return "Parameter: " & vbLf &
                "clsArgValue: " & clsArgValue.GetAsDebugString() & vbLf &
                "clsPresentation: " & clsPresentation.GetAsDebugString() & vbLf
    End Function

End Class
