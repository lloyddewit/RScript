Public Class clsRElementProperty
    Inherits clsRElementAssignable

    Public strPackageName As String = "" 'only used for functions and variables (e.g. 'constants::syms$h')
    Public lstObjects As New List(Of clsRElement) 'only used for functions and variables (e.g. 'constants::syms$h')

    Public Sub New(clsToken As clsRToken, Optional bBracketedNew As Boolean = False,
                   Optional strPackageNameNew As String = "",
                   Optional strPackagePrefix As String = "",
                   Optional lstObjectsNew As List(Of clsRElement) = Nothing)
        MyBase.New(GetTokenCleanedPresentation(clsToken, strPackageNameNew, lstObjectsNew), Nothing, bBracketedNew, strPackagePrefix)
        strPackageName = strPackageNameNew
        lstObjects = lstObjectsNew
    End Sub

    Private Shared Function GetTokenCleanedPresentation(clsToken As clsRToken,
                                    strPackageNameNew As String,
                                    lstObjectsNew As List(Of clsRElement)) As clsRToken
        Dim clsTokenNew As clsRToken = clsToken.CloneMe

        'Edge case: if the object has a package name or an object list, and formatting information
        If (Not String.IsNullOrEmpty(strPackageNameNew) OrElse
                (Not IsNothing(lstObjectsNew) AndAlso lstObjectsNew.Count > 0)) AndAlso
                (Not IsNothing(clsToken.lstTokens) AndAlso
                 clsToken.lstTokens.Count > 0 AndAlso
                 clsToken.lstTokens.Item(0).enuToken = clsRToken.typToken.RPresentation) Then
            'remove any formatting information associated with the main element.
            '  This is needed to pass test cases such as:
            '  'pkg ::  obj1 $ obj2$ fn1 ()' should be displayed as 'pkg::obj1$obj2$fn1()'
            clsTokenNew.lstTokens.Item(0).strTxt = ""
        End If

        Return clsTokenNew
    End Function

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
