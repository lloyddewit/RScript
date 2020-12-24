Public Class clsRElement

    Public strText As String
    Public bBracketed As Boolean
    Public clsPresentation As clsRElementPresentation

    ''' <summary> The R element on the target (left-hand) side of the assignment </summary>
    Public clsAssignment As clsRElement 'TODO remove this


    'TODO also possible to have complex assignments (e.g. 'x[3:5] <- 13:15', 'names(x) <- c("a","b")', 
    ' 'names(x)[3] <- "Three"'). 
    ' Any expression is allowed also on the target side of an assignment, as far as the parser is 
    ' concerned
    ' 
End Class
