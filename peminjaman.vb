Imports System.Data.Odbc
Public Class peminjaman
    Private Sub peminjaman_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call koneksi()
        Call tampil_nomor_peminjaman()
        Call tampil_buku()
        Call tanggal_pengembalian()
        lbltglpeminjaman.Text = Format(Today, "yyyy-MM-dd")
        lblstatus.Text = "Dipinjam"
    End Sub
    Sub kosongkan()
        cmbuku.Text = ""
        lbljudul.Text = ""
        lbltersedia.Text = ""
    End Sub
    Sub tampil_nomor_peminjaman()
        Call koneksi()
        Cmd = New OdbcCommand("select id_peminjaman from peminjaman order by id_peminjaman desc", conn)
        dr = Cmd.ExecuteReader
        dr.Read()
        If Not dr.HasRows Then
            lblpeminjaman.Text = Format(Today, "yyMMdd") + "001"
        Else
            If Microsoft.VisualBasic.Left(dr("id_peminjaman"), 6) = Format(Today, "yyMMdd") Then
                lblpeminjaman.Text = dr("id_peminjaman") + 1
            Else
                lblpeminjaman.Text = Format(Today, "yyMMdd") + "001"
            End If
        End If
    End Sub

    Private Sub cmbuku_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbuku.SelectedIndexChanged
        Call koneksi()
        Cmd = New OdbcCommand("select * from buku where id_buku='" & cmbuku.Text & "'", conn)
        dr = Cmd.ExecuteReader
        dr.Read()
        If dr.HasRows Then
            lbljudul.Text = dr.Item("judul")
            lbltersedia.Text = dr.Item("ketersediaan_buku")
        End If
    End Sub
    Sub tampil_buku()
        Call koneksi()
        Cmd = New OdbcCommand("select distinct id_buku from buku", conn)
        dr = Cmd.ExecuteReader
        cmbuku.Items.Clear()
        Do While dr.Read
            cmbuku.Items.Add(dr.Item("id_buku"))
        Loop
    End Sub
    Sub tanggal_pengembalian()
        Dim todayDate As DateTime = DateTime.Now
        Dim tomorrowDate As DateTime = todayDate.AddDays(7)
        lbltglpengembalian.Text = Format(tomorrowDate, "yyyy-MM-dd")
    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call kosongkan()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If cmbuku.Text = "" Or lbljudul.Text = "" Or lbltersedia.Text = "" Then
            MsgBox("Data Belum lengkap wir...!", MsgBoxStyle.Critical, "INFORMASI")
            Exit Sub
        Else
            If dr("ketersediaan_buku") = "Tersedia" Then
                Dim simpan As String = "insert into peminjaman values('" & lblpeminjaman.Text & "', '" & menu_utama.panel1.Text & "', '" & cmbuku.Text & "',  '" & lbltglpeminjaman.Text & "', '" & lbltglpengembalian.Text & "', '" & lblstatus.Text & "')"
                Cmd = New OdbcCommand(simpan, conn)
                Cmd.ExecuteNonQuery()
                MsgBox("Peminjaman Berhasil....", MsgBoxStyle.Information, "INFORMASI")
            Else
                MsgBox("Maaf,Buku Tidak Tersedia...", MsgBoxStyle.Information, "Informasi")
            End If
            Call kosongkan()
        End If
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
End Class