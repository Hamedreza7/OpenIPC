Imports System.IO

Public Class Form3
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PDF1.src = Directory.GetCurrentDirectory() & "\Manual_OpenIPC.pdf"

    End Sub
End Class