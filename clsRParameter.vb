Public Class clsRParameter
    Public strArgName As String 'TODO spaces around '=' as option?
    Public clsArgValue As clsRElement
    Public clsArgValueDefault As clsRElement
    Public iArgPos As Integer
    Public iArgPosDefinition As Integer
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
