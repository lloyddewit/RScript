Public Class clsRElementAssignable
    Inherits clsRElement

    ''' <summary>   
    ''' The statement where this element is assigned. For example, for the following R script, on the 2nd line, the statement associated with 'a' will be 'a=1'.
    ''' <code>
    ''' a=1<para>
    ''' b=a</para></code></summary>
    Public clsStatement As clsRStatement

    Public Sub New(clsToken As clsRToken, Optional clsStatementNew As clsRStatement = Nothing, Optional bBracketedNew As Boolean = False)
        MyBase.New(clsToken, bBracketedNew)
        clsStatement = clsStatementNew
    End Sub

End Class
