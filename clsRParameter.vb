Public Class clsRParameter
    Public clsArgValue As clsRElement
    Public strPrefix As String = ""

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   TODO. </summary>
    '''
    ''' <returns>   as debug string. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Function GetAsDebugString() As String
        Return "Parameter: " & vbLf &
                "clsArgValue: " & clsArgValue.GetAsDebugString() & vbLf &
                "strPrefix: " & strPrefix & vbLf
    End Function

End Class
