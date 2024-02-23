Imports System.Data.Odbc
Public Class Form_Login
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call koneksi()
        cmd = New OdbcCommand("select * from user where id_user='" & TextBox1.Text & "' and password='" & TextBox2.Text & "'", conn)
        dr = cmd.ExecuteReader
        dr.Read()
        If dr.HasRows Then
            Me.Close()
            Call bukakunci()
        Else
            MsgBox("ID User atau Password Salah!!!")

        End If
    End Sub

    Sub bukakunci()
        Form1.LoginToolStripMenuItem.Enabled = False
        Form1.LogoutToolStripMenuItem.Enabled = True
        Form1.MasterToolStripMenuItem.Enabled = True
        Form1.TransaksiToolStripMenuItem.Enabled = True
        Form1.UtilityToolStripMenuItem.Enabled = True
        Form1.LaporanToolStripMenuItem.Enabled = True
    End Sub

    Private Sub Form_Login_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
        TextBox2.Text = "123"
        TextBox2.PasswordChar = "*"
        Call koneksi()
        lbltanggal.Text = Format(Today)




    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        lbljam.Text = Format(TimeOfDay)

    End Sub
End Class