Public Class clsRElementPresentation

    Public strPrefix As String = ""
    Public strSuffix As String = ""

    '''' <summary>   The number of spaces preceeding this R element </summary>
    'Public intSpacesBefore As Integer = 0

    '''' <summary>   The comment to display before/after the R element.
    ''''             This should include any preceeding spaces and the '#' character
    ''''             (e.g. '  # x axis label').
    ''''             There is no need to include a newline character at the end of the comment.
    ''''             If the comment is converted back into a script then the new line will be added automatically.</summary>
    'Public strComment As String 'TODO make property so that `Set` always appends new line when not nothing

    '''' <summary>   True if comment should be displayed after the R element.
    ''''             False if this comment should be displayed before the R element. In this case,
    ''''             the comment will be displayed on its own line.
    ''''             </summary>
    'Public bCommentAfterElement As Boolean = True

    '''--------------------------------------------------------------------------------------------
    ''' <summary>   TODO. </summary>
    '''
    ''' <returns>   as debug string. </returns>
    '''--------------------------------------------------------------------------------------------
    Public Function GetAsDebugString() As String
        'Return "ElementPresentation: " & vbLf &
        '        "intSpacesBefore: " & intSpacesBefore & vbLf &
        '        "strComment: " & strComment & vbLf &
        '        "bCommentAfterElement: " & bCommentAfterElement & vbLf
        Return Nothing
    End Function

End Class
