Imports System.IO
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class Dialog1

    Private Sub Dialog1_Load(sender As Object, e As EventArgs) Handles Me.Load
        CheckedListBox1.Items.Clear()
        Dim path As String = "\\R0327234\mrit\EmailRecipients.txt"
        Dim applicationsToCheck As New List(Of String)

        Try
            applicationsToCheck = File.ReadAllLines(path).ToList()
        Catch ex As Exception
            MessageBox.Show("An error occurred: " & ex.Message)
        End Try

        For Each email As String In applicationsToCheck
            CheckedListBox1.Items.Add(email)
        Next
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

End Class
