Public Class StartupWindow
    Private Sub StartupWindow_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If My.Settings.FirstTimeUser = True Then
            MsgBox("Warning! It is highly recommended to make a backup of your Bus files before using this program. It has a minumum amount of safety checks, and can break your SOURCE.DBF file if not careful.", MsgBoxStyle.OkOnly)
            My.Settings.FirstTimeUser = False
        End If
    End Sub

    Private Sub OpenSourceDBFWindow_Click(sender As Object, e As EventArgs) Handles OpenSourceDBFWindow.Click
        EditSourceDBFWindow.Show()
        Me.Hide()
        EditSourceDBFWindow.Focus()
        EditSourceDBFWindow.btnOpenFile.PerformClick()
    End Sub
End Class