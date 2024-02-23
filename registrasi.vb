Imports System.Data.Odbc
Public Class registrasi
    Private Sub registrasi_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call koneksi()
        teks = textrunning1.Text()
        lbltanggal.Text = Format(Today)
    End Sub
    Sub gambar()
        On Error Resume Next
        PictureBox3.Load(TextBox6.Text)
        PictureBox3.SizeMode = PictureBoxSizeMode.StretchImage
    End Sub
    Sub cek_email()
        Call koneksi()
        Cmd = New OdbcCommand("select * from user where status_user='Peminjam' and email='" & TextBox3.Text & "' or nama_lengkap='" & TextBox4.Text & "'", conn)
        dr = Cmd.ExecuteReader
        dr.Read()
    End Sub
    Sub hapus()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        ComboBox1.Text = ""
        PictureBox3.Refresh()
    End Sub
    Private Sub Label18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label18.Click
        login.Show()
        Me.Hide()
    End Sub
    Dim bergerak As Integer
    Dim teks As String
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        textrunning1.Text = bergerak
        teks = Microsoft.VisualBasic.Right(teks, Len(teks) - 3) & Microsoft.VisualBasic.Left(teks, 3)
        textrunning1.Text = teks
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        lbljam.Text = Format(TimeOfDay)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        OpenFileDialog1.ShowDialog()
        OpenFileDialog1.Filter = "(*.Jpg) | *.Jpg"
        PictureBox3.Text = PictureBox3.Text + "<img>" + OpenFileDialog1.FileName + "</img>"
        TextBox6.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged
        Call gambar()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or ComboBox1.Text = "" Then
            MsgBox("Data Belum lengkap Harap Isi Dengan Lengkap...!", MsgBoxStyle.Critical, "INFORMASI")
            Exit Sub
        Else
            Dim newtext As String
            Dim oldtext As String = TextBox6.Text
            If oldtext = "" Then
                newtext = ""
            Else
                newtext = oldtext.Replace("\", "\\")
            End If

            Call cek_email()
            If dr.HasRows Then
                MsgBox("Email atau Nama telah terdaftar!!", MsgBoxStyle.Information, "Informasi")
            Else
                'simpan
                Dim simpan As String = "insert into user values ('" & TextBox1.Text & "','" & TextBox2.Text & "', '" & TextBox3.Text & "', '" & ComboBox1.Text & "', '" & TextBox4.Text & "', '" & TextBox5.Text & "', '" & newtext & "')"
                Cmd = New OdbcCommand(simpan, conn)
                Cmd.ExecuteNonQuery()
                MsgBox("Registrasi Berhasil....", MsgBoxStyle.Information, "INFORMASI")
                login.Show()
                Me.Hide()
            End If
            Call hapus()
            TextBox1.Focus()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
    Private Sub PictureBox2_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox2.MouseDown
        TextBox2.UseSystemPasswordChar = False
    End Sub

    Private Sub PictureBox2_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox2.MouseUp
        TextBox2.UseSystemPasswordChar = True
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        TextBox1.MaxLength = 11
        If e.KeyChar = Chr(13) Then
            TextBox2.Focus()
        End If
    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        TextBox2.MaxLength = 50
        If e.KeyChar = Chr(13) Then
            TextBox3.Focus()
        End If
    End Sub

    Private Sub TextBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        TextBox3.MaxLength = 50
        If e.KeyChar = Chr(13) Then
            TextBox4.Focus()
        End If
    End Sub

    Private Sub TextBox4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox4.KeyPress
        TextBox4.MaxLength = 200
        If e.KeyChar = Chr(13) Then
            TextBox5.Focus()
        End If
    End Sub

    Private Sub TextBox5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox5.KeyPress
        TextBox5.MaxLength = 200
        If e.KeyChar = Chr(13) Then
            ComboBox1.Focus()
        End If
    End Sub
End Class