Imports System.IO

Public Class Form1


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Me.Cursor = Cursors.WaitCursor
        Connect_to_Sap()
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click

        Me.Cursor = Cursors.WaitCursor
        Disconnect_to_Sap()
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click

        Me.Cursor = Cursors.WaitCursor
        View_log()
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click

        Me.Cursor = Cursors.WaitCursor
        Split_Order()
        Read_SO_file()
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        Me.Cursor = Cursors.WaitCursor
        Disconnect_to_Sap()
        Me.Cursor = Cursors.Default

        End

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.Cursor = Cursors.WaitCursor
        ReadSettings()
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub Button8_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click

        Me.Cursor = Cursors.WaitCursor
        CreateUdfFiledsText()
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Me.Cursor = Cursors.WaitCursor
        CheckAndExport_delivery()
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Me.Cursor = Cursors.WaitCursor
        CheckAndExport_invoice()
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        Me.Cursor = Cursors.WaitCursor

        Try
            File.Delete(Application.StartupPath & "\log.txt")
        Catch ex As Exception

        End Try

        Me.Cursor = Cursors.Default

    End Sub

End Class
