Option Explicit On
Public Class ThisApplication
    Public WithEvents oApp As Word.Application

    Private Sub oApp_DocumentOpen(ByVal Doc As Microsoft.Office.Interop.Word.Document) Handles oApp.DocumentOpen

    End Sub


End Class
