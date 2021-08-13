Public Class clsRControl

    Public lstParameters As New List(Of clsRParameter)

    ' TODO specify how each child is updated when this control's value changes
    '      OR should we manage the dependencies between the controls at the dialog level?
    Public lstChildControls As List(Of clsRControl)

    Public Sub AddParameter(clsParameter As clsRParameter)
        lstParameters.Add(clsParameter)
    End Sub

    Public Sub SetValue(strNewValue As String)
        For Each clsParamater As clsRParameter In lstParameters
            'clsParameter.value = strNewValue
        Next

        For Each clsControl As clsRControl In lstChildControls
            'determine new value of child control
            ' clsChildControl.SetValue(new child value)
        Next

        'raise event that this control's value has changed, UI control then updates itself based on the new value
        '     or we could pass a function reference when we create the clsRControl object, and then we could call this function
    End Sub
End Class
