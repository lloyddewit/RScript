Public Class clsRDialog

    Public clsScript As clsRScript

    'Public lstControls As New List(Of clsRControl)

    Public Sub New(strScript As String)
        clsScript = New clsRScript(strScript)
    End Sub

    'TODO add new parameter: event to raise when value changes (or function to call when value changes)
    Public Function GetControl(lstParameterSpecs As List(Of clsRParameterSpec), lstChildControls As List(Of clsRControl)) As clsRControl
        Dim clsControl As New clsRControl

        For Each clsParameterSpec As clsRParameterSpec In lstParameterSpecs
            Dim clsParamater As clsRParameter

            'search through clsScript until a param is found that matches the parameter specification
            ' clsParameter = <the matching parameter in clsScript>

            clsControl.AddParameter(clsParamater)
        Next

        Return clsControl
    End Function



End Class
