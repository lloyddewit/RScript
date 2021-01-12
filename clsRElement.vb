Public Class clsRElement

    Public strText As String
    Public bBracketed As Boolean
    Public clsPresentation As clsRElementPresentation

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   TODO. </summary>
    '''
    ''' <returns>   as debug string. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Function GetAsDebugString() As String
        Return "Element: " & vbLf &
                "strText: " & strText & vbLf &
                "bBracketed: " & bBracketed & vbLf &
                "clsPresentation: " & clsPresentation.GetAsDebugString() & vbLf
    End Function


End Class
