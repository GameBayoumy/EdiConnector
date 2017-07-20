Imports System.IO

Public Class Form3

    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim fs As New StreamReader(Application.StartupPath & "\log.txt", True)
        TextBox1.Text = fs.ReadToEnd()
        fs.Close()

    End Sub

End Class