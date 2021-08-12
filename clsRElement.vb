'---------------------------------------------------------------------------------------------------
' file:		clsRElement.vb
'
' summary:	TODO
'---------------------------------------------------------------------------------------------------
Public Class clsRElement
    ''' <summary> The text representation of the element (e.g. '+', '/', 'myFunction', 
    '''           '"my string constant"' etc.). </summary>
    Public strTxt As String

    ''' <summary> If true, then the element is surrounded by round brackets. For example, if the 
    '''           script is 'a*(b+c)', then the element representing the '+' operator will have 
    '''           'bBracketed' set to true. </summary>
    Public bBracketed As Boolean

    ''' <summary> 
    ''' Any formatting text that precedes the element. The formatting text may consist of spaces, 
    ''' comments and new lines to make the script more readable for humans. For example, in the 
    ''' example below, 'strprefix' for the 'myFunction' element shall be set to 
    ''' "#comment1\n  #comment2\n  ".<code>
    ''' 
    ''' #comment1<para>
    '''   #comment2</para><para>
    '''   myFunction()</para></code></summary>
    Public strPrefix As String = ""

    Public Sub New(clsToken As clsRToken,
                   Optional bBracketedNew As Boolean = False,
                   Optional strPackagePrefix As String = "")
        strTxt = clsToken.strTxt
        bBracketed = bBracketedNew
        strPrefix = strPackagePrefix &
                         If(clsToken.lstTokens.Count > 0 AndAlso
                            clsToken.lstTokens.Item(0).enuToken = clsRToken.typToken.RPresentation,
                            clsToken.lstTokens.Item(0).strTxt, "")

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
                "strPrefix: " & strPrefix & vbLf
    End Function


End Class
