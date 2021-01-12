Public Class clsRParameterNamed
    Inherits clsRParameter
    Public strArgumentName As String 'TODO spaces around '=' as option?

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   TODO. </summary>
    '''
    ''' <returns>   as debug string. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Overloads Function GetAsDebugString() As String
        Return "ParameterNamed: " & vbLf &
                "strArgumentName: " & strArgumentName & vbLf
    End Function

End Class
