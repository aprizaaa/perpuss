Imports System.Data.Odbc
Public Class pengembalian
    Private Sub pengembalian_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call koneksi()
        Call tampil_nomor_peminjaman()
        lbltglpengembalian.Text = Format(Today, "yyyy-MM-dd")
        lblstatus.Text = "Dikembalikan"
    End Sub
    Private Sub cmpeminjaman_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmpeminjaman.SelectedIndexChanged
        Call koneksi()
        Cmd = New OdbcCommand("select * from peminjaman inner join buku on peminjaman.id_buku = buku.id_buku  where peminjaman.id_peminjaman='" & cmpeminjaman.Text & "'", conn)
        dr = Cmd.ExecuteReader
        dr.Read()
        If dr.HasRows Then
            lblkode.Text = dr.Item("id_buku")
            lbljudul.Text = dr.Item("judul")
            lbltglpeminjaman.Text = dr.Item("tanggal_peminjaman")
        End If
    End Sub
    Sub kosongkan()
        cmpeminjaman.Text = ""
        lbltglpengembalian.Text = ""
        lbljudul.Text = ""
    End Sub
    Sub tampil_nomor_peminjaman()
        Call koneksi()
        Cmd = New OdbcCommand("select * from peminjaman where id_user='" & menu_utama.panel1.Text & "'", conn)
        dr = Cmd.ExecuteReader
        cmpeminjaman.Items.Clear()
        Do While dr.Read
            cmpeminjaman.Items.Add(dr.Item("id_peminjaman"))
        Loop
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If cmpeminjaman.Text = "" Or lbltglpengembalian.Text = "" Or lbljudul.Text = "" Then
            MsgBox("Data Harus lengkap...!", MsgBoxStyle.Critical, "INFORMASI")
            Exit Sub
        Else
            Dim update As String = "update peminjaman set tanggal_pengembalian='" & Format(Today, "yyyy-MM-dd") & "', status_peminjaman='" & lblstatus.Text & "' where id_peminjaman='" & cmpeminjaman.Text & "'"
            Cmd = New OdbcCommand(update, conn)
            Cmd.ExecuteNonQuery()
            MsgBox("Buku Berhasil Dikembalikan....", MsgBoxStyle.Information, "INFORMASI")

            Call cek_pengembalian()
            If Not dr.HasRows Then
                MsgBox("Anda sudah tidak meminjam buku apapun", MsgBoxStyle.Information, "Informasi")
                Me.Close()
            End If
            Call kosongkan()
        End If
    End Sub
End Class