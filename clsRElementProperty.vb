Public Class clsRElementProperty
    Inherits clsRElement

    Public strPackageName As String 'only used for functions and variables (e.g. 'constants::syms$h')
    Public lstRObjects As New List(Of clsRElement) 'only used for functions and variables (e.g. 'constants::syms$h')

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   TODO. </summary>
    '''
    ''' <returns>   as debug string. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Overloads Function GetAsDebugString() As String
        Dim strTxt As String = "ElementProperty: " & vbLf &
                "strPackageName: " & strPackageName & vbLf

        For Each clsElement As clsRElement In lstRObjects
            strTxt &= clsElement.GetAsDebugString()
        Next

        Return strTxt
    End Function

End Class
