'---------------------------------------------------------------------------------------------------
' file:		clsRToken.vb
'
' summary:	TBD
'---------------------------------------------------------------------------------------------------
Public Class clsRToken
    ''' <summary>   Values that represent TBD. </summary>
    Public Enum typToken
        RSyntacticName
        RKeyWord
        RConstantString
        RComment
        RSpace
        RBracket
        RSeparator
        REndStatement
        REndScript
        RNewLine
        ROperatorUnaryLeft
        ROperatorUnaryRight
        ROperatorBinary
        ROperatorBracket
        RInvalid
    End Enum

    ''' <summary>   TBD. </summary>
    Public strTxt As String
    ''' <summary>   TBD. </summary>
    Public enuToken As typToken
End Class
