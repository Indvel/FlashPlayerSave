Imports System
Imports System.IO
Imports System.IO.Compression

Public Class Form1

    Dim path As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Macromedia\Flash Player\#SharedObjects"

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim di As New System.IO.DirectoryInfo(path)

        For Each dir As System.IO.DirectoryInfo In di.GetDirectories()
            ListBox1.Items.Add(dir)
        Next

        If My.Computer.FileSystem.FileExists(Application.StartupPath & "\FlashSaves.zip") Then
            Label1.Text = "최근 백업일 : " & File.GetLastWriteTime(Application.StartupPath & "\FlashSaves.zip")
        End If
    End Sub

    Private Sub ListBox1_MouseDoubleClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles ListBox1.MouseDoubleClick
        Process.Start(path & "\" & ListBox1.SelectedItem.ToString)
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Button1.Text = "백업 중.."
        If My.Computer.FileSystem.FileExists(Application.StartupPath & "\FlashSaves.zip") Then
            System.IO.File.Delete(Application.StartupPath & "\FlashSaves.zip")
        End If
        ZipFile.CreateFromDirectory(path, "FlashSaves.zip")
        Button1.Text = "백업"
        MsgBox("백업이 완료되었습니다.", MsgBoxStyle.Information, "백업 완료")
        Label1.Text = "최근 백업일 : " & File.GetLastWriteTime(Application.StartupPath & "\FlashSaves.zip")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Button2.Text = "복구 중.."
        If My.Computer.FileSystem.FileExists(Application.StartupPath & "\FlashSaves.zip") Then
            If Directory.Exists(path) Then
                For Each file As String In Directory.GetDirectories(path, "*.*", SearchOption.TopDirectoryOnly)
                    System.IO.Directory.Delete(file, True)
                Next
            End If
        Else
            MsgBox("백업된 FlashSaves.zip 파일을 찾을 수 없습니다.", MsgBoxStyle.Critical, "복구 실패")
            Button2.Text = "복구"
        End If
        ZipFile.ExtractToDirectory(Application.StartupPath & "\FlashSaves.zip", path)
        MsgBox("복구가 완료되었습니다.", MsgBoxStyle.Information, "복구 완료")
        Button2.Text = "복구"
    End Sub
End Class
