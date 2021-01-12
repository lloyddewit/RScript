'---------------------------------------------------------------------------------------------------
' file:		clsRToken.vb
'
' summary:	TBD
'---------------------------------------------------------------------------------------------------
Public Class clsRToken
    ''' <summary>   Values that represent TBD. </summary>
    Public Enum typToken
        RSyntacticName
        RFunctionName
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
    ''' <summary>   TBD. </summary>
    Public lstTokens As New List(Of clsRToken)

    Public Function CloneMe() As clsRToken
        Dim clsToken = New clsRToken
        clsToken.strTxt = strTxt
        clsToken.enuToken = enuToken

        For Each clsTokenChild As clsRToken In lstTokens
            clsToken.lstTokens.Add(clsTokenChild.CloneMe)
        Next

        Return clsToken
    End Function
End Class
