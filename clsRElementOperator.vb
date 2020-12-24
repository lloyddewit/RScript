Public Class clsRElementOperator
    Inherits clsRElement
    Public bFirstParamOnRight As Boolean = False
    Public strTerminator 'only used for '[' and '[[' operators
    Public lstRParameters As New List(Of clsRParameter)
End Class
