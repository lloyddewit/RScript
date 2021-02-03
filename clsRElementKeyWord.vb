Public Class clsRElementKeyWord
    Inherits clsRElement

    Public lstRParameters As New List(Of clsRParameterNamed)
    Public clsScript As clsRScript

    Public Sub New(clsToken As clsRToken, Optional bBracketedNew As Boolean = False)
        MyBase.New(clsToken, bBracketedNew)
    End Sub

    'Public clsObject As Object 'if statement part in '()' that returns true or false
    'fn: argument definition (also in '()')
    'else: ! of if?

End Class
