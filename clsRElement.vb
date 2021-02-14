Public Class clsRElement

    Public strTxt As String
    Public bBracketed As Boolean
    Public clsPresentation As clsRElementPresentation

    Public Sub New(clsToken As clsRToken,
                   Optional bBracketedNew As Boolean = False,
                   Optional strPackagePrefix As String = "")
        strTxt = clsToken.strTxt
        bBracketed = bBracketedNew
        clsPresentation = New clsRElementPresentation With {
            .strPrefix = strPackagePrefix &
                         If(clsToken.lstTokens.Count > 0 AndAlso
                            clsToken.lstTokens.Item(0).enuToken = clsRToken.typToken.RPresentation,
                            clsToken.lstTokens.Item(0).strTxt, "")
        }
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
