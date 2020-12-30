Public Class clsRToken
    Public Enum typToken
        RSyntacticName
        RKeyWord
        RConstantString
        RComment
        RSpace
        RBracket
        RSeparator
        ROperatorUnaryLeft
        ROperatorUnaryRight
        ROperatorBinary
        ROperatorBracket
    End Enum

    Public strText As String
    Public enuToken As typToken
End Class
