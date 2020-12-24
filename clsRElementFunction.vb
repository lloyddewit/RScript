Public Class clsRElementFunction
    Inherits clsRElementProperty

    Public lstRParameters As New List(Of clsRParameterNamed)
    'Public lstRParams As New List(Of clsRParameterFunction) 'need to override parameter list of parent

    'Private _lstRParams As New List(Of clsRParameterFunction)

    'Public Overrides Property LstRParams As List(Of clsRParameterFunction)
    '    Get
    '        Return _lstRParams
    '    End Get
    '    Set(value As List(Of clsRParameter))
    '        _lstRParams = value
    '    End Set
    'End Property

End Class
