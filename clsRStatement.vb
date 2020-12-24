Public Class clsRStatement

    ''' <summary>
    ''' If true, then when this R statement is converted to a script, then it will be 
    '''             terminated with a newline (else if false then a semicolon)
    ''' </summary>
    Public bTerminateWithNewline As Boolean = True

    Public clsAssignment As clsRElement
    Public clsElement As clsRElement

    ''' <summary>   The list of R elements that make up this R statement </summary>
    'Public lstRElements As List(Of clsRElement)

End Class
