Public Class clsRElement

    Public strTxt As String
    Public bBracketed As Boolean
    Public clsPresentation As clsRElementPresentation

    Public Sub New(clsToken As clsRToken, Optional bBracketedNew As Boolean = False, Optional clsPresentationNew As clsRElementPresentation = Nothing)
        strTxt = clsToken.strTxt
        bBracketed = bBracketedNew
        clsPresentation = clsPresentationNew
    End Sub

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   TODO. </summary>
    '''
    ''' <returns>   as debug string. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Function GetAsDebugString() As String
        Return "Element: " & vbLf &
                "strTxt: " & strTxt & vbLf &
                "bBracketed: " & bBracketed & vbLf &
                "clsPresentation: " & clsPresentation.GetAsDebugString() & vbLf
    End Function


End Class
