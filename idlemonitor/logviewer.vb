Imports System.IO

Public Class logviewer
    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        OpenFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
        OpenFileDialog1.Filter = "Log Files (*.log)|*.log"

        Dim result As DialogResult = OpenFileDialog1.ShowDialog()

        If result = Windows.Forms.DialogResult.OK Then
            Dim path As String = OpenFileDialog1.FileName
            If File.Exists(path) Then
                Try
                    readLogFile(path)
                Catch ex As Exception
                    MsgBox("Error: " & ex.Message)
                End Try
            End If
        End If
    End Sub

    Private Sub readLogFile(ByVal logFile As String)
        TextBox1.Clear()
        Dim b As Byte()
        For Each line As String In File.ReadAllLines(logFile)
            b = System.Convert.FromBase64String(line)
            TextBox1.AppendText(System.Text.ASCIIEncoding.ASCII.GetString(b) & vbCrLf)
        Next
    End Sub

End Class