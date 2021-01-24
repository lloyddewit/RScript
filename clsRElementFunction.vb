Public Class clsRElementFunction
    Inherits clsRElementProperty

    Public lstParameters As New List(Of clsRParameterNamed)

    Public Sub New(clsToken As clsRToken, Optional bBracketedNew As Boolean = False,
                   Optional lstObjectsNew As List(Of clsRElement) = Nothing,
                   Optional clsPresentationNew As clsRElementPresentation = Nothing)
        MyBase.New(clsToken, bBracketedNew, lstObjectsNew, clsPresentationNew)
    End Sub



    '''--------------------------------------------------------------------------------------------
    ''' <summary>   TODO. </summary>
    '''
    ''' <returns>   as debug string. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Overloads Function GetAsDebugString() As String
        Dim strTxt As String = "ElementFunction: " & vbLf

        For Each clsParameter As clsRParameterNamed In lstParameters
            strTxt &= clsParameter.GetAsDebugString()
        Next

        Return strTxt
    End Function

End Class
