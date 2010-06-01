Imports Microsoft.Office.Tools 'toolbar import

'vb addin to create a speech pane for Office 2007
'code of interest primarily in SpeechControl.vb file, this file just declares toolbar
Public Class ThisAddIn

    'creates SpeechToolbox on word startup, sets visible
    Private speechTaskPane As CustomTaskPane
    Private Sub ThisAddIn_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup
        speechTaskPane = Me.CustomTaskPanes.Add(New SpeechControl, "Text to Speech")
        speechTaskPane.Visible = True
    End Sub

    'removes toolbox on word shutdown
    Private Sub ThisAddIn_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown

    End Sub

End Class
