﻿Public Class clsRElementOperator
    Inherits clsRElement
    Public bFirstParamOnRight As Boolean = False
    Public strTerminator As String = "" 'only used for '[' and '[[' operators
    Public lstRParameters As New List(Of clsRParameter)

    Public Sub New(clsToken As clsRToken, Optional bBracketedNew As Boolean = False, Optional clsPresentationNew As clsRElementPresentation = Nothing)
        MyBase.New(clsToken, bBracketedNew, clsPresentationNew)
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

        For Each clsParameter As clsRParameter In lstRParameters
            strTxt &= clsParameter.GetAsDebugString()
        Next

        Return strTxt
    End Function

End Class
