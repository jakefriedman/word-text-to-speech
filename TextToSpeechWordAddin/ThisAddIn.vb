Option Explicit On
Imports Microsoft.Office.Tools 'toolbar import

'vb addin to create a speech pane for Office 2007
'code of interest primarily in SpeechControl.vb file, this file declares toolbar and handles UI and file changes
Public Class ThisAddIn


    Private speechTaskPane As CustomTaskPane
    Private speechDisplayed As Boolean
    Private show As Boolean

    'creates Speech taskpanes on word startup, sets visible
    Private Sub ThisAddIn_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup
        AddAllSpeechTaskPanes()
        show = True
    End Sub

    'Creates a task pane for every open document
    Public Sub AddAllSpeechTaskPanes()
        If Globals.ThisAddIn.Application.Documents.Count > 0 Then 'are docs open?
            If Me.Application.ShowWindowsInTaskbar Then
                For Each _doc As Word.Document In Me.Application.Documents
                    AddSpeechTaskPane(_doc)
                Next
            Else
                If Not speechDisplayed Then
                    AddSpeechTaskPane(Me.Application.ActiveDocument)
                End If
            End If
            speechDisplayed = True
        End If
    End Sub

    'Makes all taskpanes visible
    Public Sub ShowAllSpeechTaskPanes()
        Dim ctp As Microsoft.Office.Tools.CustomTaskPane
        For i As Integer = Me.CustomTaskPanes.Count - 1 To 0 Step -1
            ctp = Me.CustomTaskPanes.Item(i)
            If ctp.Title = "Text to Speech" Then
                ctp.Visible = True
            End If
        Next
    End Sub

    'Hides all Taskpanes
    Public Sub HideAllSpeechTaskPanes()
        Dim ctp As Microsoft.Office.Tools.CustomTaskPane
        For i As Integer = Me.CustomTaskPanes.Count - 1 To 0 Step -1
            ctp = Me.CustomTaskPanes.Item(i)
            If ctp.Title = "Text to Speech" Then
                ctp.Visible = False 'just hide taskpane
            End If
        Next
    End Sub

    'Creates a taskpane in active window
    Public Sub AddSpeechTaskPane(ByVal doc As Word.Document)
        speechTaskPane = Me.CustomTaskPanes.Add(New SpeechControl, "Text to Speech", doc.ActiveWindow)
        speechTaskPane.Visible = True
    End Sub

    'enables modified Ribbon XML
    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return New MyRibbon()
    End Function

    'show/hide button was pressed
    Public Sub Button()
        If show Then
            show = False
            Globals.ThisAddIn.HideAllSpeechTaskPanes()
        Else
            show = True
            Globals.ThisAddIn.ShowAllSpeechTaskPanes()
        End If
    End Sub

    'on new document, make new taskpane
    Private Sub Application_NewDocument(ByVal Doc As Word.Document) Handles Application.NewDocument
        If show And Me.Application.ShowWindowsInTaskbar Then
            AddSpeechTaskPane(Doc)
        End If
    End Sub

    'on opened document, make new taskpane
    Private Sub Application_DocumentOpen(ByVal Doc As Word.Document) Handles Application.DocumentOpen
        RemoveOrphanedTaskPanes()
        If show And Me.Application.ShowWindowsInTaskbar Then
            AddSpeechTaskPane(Doc) 'make it and make it visible
        Else
            speechTaskPane = Me.CustomTaskPanes.Add(New SpeechControl, "Text to Speech", Doc.ActiveWindow) 'only make it
        End If
    End Sub

    'documents closed, remove panes associated with no document
    Private Sub Application_DocumentChange() Handles Application.DocumentChange
        RemoveOrphanedTaskPanes()
    End Sub

    'removes old taskpanes (doc not open anymore)
    Private Sub RemoveOrphanedTaskPanes()
        Dim ctp As Microsoft.Office.Tools.CustomTaskPane
        For i As Integer = Me.CustomTaskPanes.Count - 1 To 2 Step -1
            ctp = Me.CustomTaskPanes.Item(i)
            If ctp.Window Is Nothing Then
                Me.CustomTaskPanes.Remove(ctp)
            End If
        Next
    End Sub

    'dont call remove panes on Shutdown!
    Private Sub ThisAddIn_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown
    End Sub



End Class
