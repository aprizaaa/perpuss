Imports System.Data.Odbc
Public Class menu_utama
    Private Sub menu_utama_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call koneksi()
        Call biodata()
    End Sub
    Sub biodata()
        Call koneksi()
        Cmd = New OdbcCommand("select * from user where id_user='" & login.TextBox1.Text & "'", conn)
        dr = Cmd.ExecuteReader
        dr.Read()
        If dr.HasRows Then
            lblnama.Text = dr.Item("nama_lengkap")
            lblstatus.Text = dr.Item("status_user")
            TextBox1.Text = dr.Item("foto")
            Call Gambar()
        End If
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        If MessageBox.Show("Yakin Akan Keluar?", "WARNING", MessageBoxButtons.YesNo, MessageBoxIcon.Information) = Windows.Forms.DialogResult.Yes Then
            Me.Close()
            login.Show()
            login.kosongkan()
            Call simpanlog_logout()
        Else
        End If
    End Sub
    Sub Gambar()
        PictureBox1.Load(TextBox1.Text)
        PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
    End Sub
    Sub simpanlog_logout()
        Call koneksi()
        Dim simpanlog As String = "insert into log_aktivitas values('""','" & panel1.Text & "','logout','" & Format(DateValue(login.lbltanggal.Text), "yyyy-MM-dd") & "','" & login.lbljam.Text & "')"
        Cmd = New OdbcCommand(simpanlog, conn)
        Cmd.ExecuteNonQuery()
    End Sub

    Private Sub lbljam_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbljam.Click

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        lbljam.Text = (TimeOfDay)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        user.Show()
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        buku.Show()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        cari_buku.Show()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        peminjaman.Show()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        kartu.Show()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Call cek_pengembalian()
        If dr.HasRows Then
            pengembalian.Show()
        Else
            MsgBox("Anda Harus meminjam buku Dulu....", MsgBoxStyle.Information, "informasi")
        End If
    End Sub
End Class