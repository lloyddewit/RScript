Public Class clsRElementOperator
    Inherits clsRElement
    Public bFirstParamOnRight As Boolean = False
    Public strTerminator As String = "" 'only used for '[' and '[[' operators
    Public lstParameters As New List(Of clsRParameter)

    Public Sub New(clsToken As clsRToken, Optional bBracketedNew As Boolean = False,
                   Optional bFirstParamOnRightNew As Boolean = False)
        MyBase.New(clsToken, bBracketedNew)
        bFirstParamOnRight = bFirstParamOnRightNew
    End Sub

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   TODO. </summary>
    '''
    ''' <returns>   as debug string. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Overloads Function GetAsDebugString() As String
        Dim strTxt As String = "ElementOperator: " & vbLf &
                "bFirstParamOnRight: " & bFirstParamOnRight & vbLf &
                "strTerminator: " & strTerminator & vbLf &
                "lstRParameters" & vbLf

        For Each clsParameter As clsRParameter In lstParameters
            strTxt &= clsParameter.GetAsDebugString()
        Next

        Return strTxt
    End Function

End Class
