Public Class clsRElementFunction
    Inherits clsRElementProperty

    Public lstRParameters As New List(Of clsRParameterNamed)

    Public Sub New(clsToken As clsRToken, Optional bBracketedNew As Boolean = False, Optional clsPresentationNew As clsRElementPresentation = Nothing)
        MyBase.New(clsToken, bBracketedNew, clsPresentationNew)
    End Sub



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
