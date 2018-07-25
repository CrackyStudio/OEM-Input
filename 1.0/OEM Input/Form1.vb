Imports System.Runtime.InteropServices
Imports Microsoft.Win32
Public Class Form1
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As Integer, <MarshalAs(UnmanagedType.LPWStr)> ByVal lParam As String) As Int32
    End Function
    Private LogoLocation As String = ""
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckBox1.Checked = True
        SendMessage(Me.TextBox1.Handle, &H1501, 0, "Manufacturer")
        SendMessage(Me.TextBox2.Handle, &H1501, 0, "Model")
        SendMessage(Me.TextBox3.Handle, &H1501, 0, "SupportURL")
        SendMessage(Me.TextBox4.Handle, &H1501, 0, "SupportPhone")
        SendMessage(Me.TextBox5.Handle, &H1501, 0, "SupportHours")
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim OEMInformation = My.Computer.Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation", True)
        If TextBox1.Text IsNot "" Then
            OEMInformation.SetValue("Manufacturer", TextBox1.Text)
        End If
        If TextBox2.Text IsNot "" Then
            OEMInformation.SetValue("Model", TextBox2.Text)
        End If
        If TextBox3.Text IsNot "" Then
            OEMInformation.SetValue("SupportURL", TextBox3.Text)
        End If
        If TextBox4.Text IsNot "" Then
            OEMInformation.SetValue("SupportPhone", TextBox4.Text)
        End If
        If TextBox5.Text IsNot "" Then
            OEMInformation.SetValue("SupportHours", TextBox5.Text)
        End If
        If LogoLocation IsNot "" Then
            OEMInformation.SetValue("Logo", LogoLocation)
        End If
        OEMInformation.Close()
        If CheckBox1.Checked = True Then
            System.Diagnostics.Process.Start("control", "system")
        End If
        Application.Exit()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        PictureBox1.Image = Nothing
        CheckBox1.Checked = True
    End Sub
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Dim open As New OpenFileDialog()
        open.Filter = "Image Files(*.bmp)|*.bmp"
        If open.ShowDialog() = DialogResult.OK Then
            Dim bmp = Image.FromFile(open.FileName)
            If bmp.Size.Width <= 120 AndAlso bmp.Size.Height <= 120 Then
                LogoLocation = System.IO.Path.GetFullPath(open.FileName)
                PictureBox1.Image = New Bitmap(open.FileName)
                Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
            Else
                MsgBox("Image must be 120x120 pixels max")
            End If
        End If
    End Sub
End Class
