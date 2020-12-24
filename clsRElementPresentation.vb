Public Class clsRElementPresentation

    ''' <summary>   The number of spaces preceeding this R element </summary>
    Public intSpacesBefore As Integer = 0

    ''' <summary>   The comment to display after the R element.
    '''             This should include any preceeding spaces and the '#' character
    '''             (e.g. '  # x axis label').
    '''             If not null then a newline is appended to the end of the comment.
    '''             An empty string ("" *not* null) results in a newline with no visible comment</summary>
    Public strComment As String 'TODO make property so that `Set` always appends new line when not nothing


End Class
