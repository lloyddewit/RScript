Public Class clsRElementProperty
    Inherits clsRElement

    Public strPackageName As String = "" 'only used for functions and variables (e.g. 'constants::syms$h')
    Public lstObjects As New List(Of clsRElement) 'only used for functions and variables (e.g. 'constants::syms$h')

    Public Sub New(clsToken As clsRToken, Optional bBracketedNew As Boolean = False,
                   Optional strPackageNameNew As String = "",
                   Optional lstObjectsNew As List(Of clsRElement) = Nothing)
        MyBase.New(clsToken, bBracketedNew)
        strPackageName = strPackageNameNew
        lstObjects = lstObjectsNew
    End Sub


    '''--------------------------------------------------------------------------------------------
    ''' <summary>   TODO. </summary>
    '''
    ''' <returns>   as debug string. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Overloads Function GetAsDebugString() As String
        Dim strTxt As String = "ElementProperty: " & vbLf &
                "strPackageName: " & strPackageName & vbLf

        For Each clsElement As clsRElement In lstObjects
            strTxt &= clsElement.GetAsDebugString()
        Next

        Return strTxt
    End Function

End Class
