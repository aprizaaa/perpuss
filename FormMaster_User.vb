Imports System.Data.Odbc
Public Class FormMaster_User

    Sub kondisiAwal()
        Call koneksi()
        da = New OdbcDataAdapter("select * from user", conn)
        ds = New DataSet
        da.Fill(ds, "user")
        DataGridView1.DataSource = (ds.Tables("user"))
        ComboBox1.Items.Clear()
        ComboBox1.Items.Add("ADMINISTRATOR")
        ComboBox1.Items.Add("PETUGAS")
        ComboBox1.Items.Add("PEMINJAM")

    End Sub

    Private Sub FormMaster_User_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call kondisiAwal()
    End Sub

    ''untuk menyimpan data user''
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call koneksi()
        Dim simpan As String = "insert into user values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "','" & ComboBox1.Text & "')"
        cmd = New OdbcCommand(simpan, conn)
        cmd.ExecuteNonQuery()
        MsgBox("DATA BERHASIL DI SIMPAN")
        Call kondisiAwal()


    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()

    End Sub
    'untuk mengedit data user'
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call koneksi()
        Dim edit As String = "UPDATE user SET username='" & TextBox2.Text & "', password='" & TextBox3.Text & "', email='" & TextBox4.Text & "', nama_lengkap='" & TextBox5.Text & "', alamat='" & TextBox6.Text & "', level='" & ComboBox1.Text & "' WHERE id_user='" & TextBox1.Text & "'"
        cmd = New OdbcCommand(edit, conn)
        cmd.ExecuteNonQuery()
        MsgBox("Data berhasil di edit...", MsgBoxStyle.Information, "INFORMASI")
        Call kondisiAwal()

    End Sub
    'untuk menghapus data user'
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call koneksi()
        Dim hapus As String = "delete from user where id_user='" & TextBox1.Text & "'"
        cmd = New OdbcCommand(hapus, conn)
        cmd.ExecuteNonQuery()
        MsgBox("Data berhasil di hapus...", MsgBoxStyle.Information, "INFORMASI")
        Call kondisiAwal()
    End Sub
    'untuk mencari data user'
    Sub Caridata()
        Call koneksi()
        da = New OdbcDataAdapter("select * from user where id_user like '%" & TextBox7.Text & "' or username like '%" & TextBox7.Text & "%'", conn)
        ds = New DataSet
        da.Fill(ds)
        DataGridView1.DataSource = ds.Tables(0)
        DataGridView1.ReadOnly = True
    End Sub
    Sub Panggil_Kode()
        Call koneksi()
        cmd = New OdbcCommand("select * from user where id_user='" & TextBox1.Text & "'", conn)
        dr = cmd.ExecuteReader
        dr.Read()
    End Sub

    Sub Panggil_Data()
        On Error Resume Next
        TextBox2.Text = dr.Item(1) 'nama user
        TextBox3.Text = dr.Item(2) 'password
        TextBox4.Text = dr.Item(3) 'email
        TextBox5.Text = dr.Item(4) 'nama lengkap
        TextBox6.Text = dr.Item(5) 'alamat
        ComboBox1.Text = dr.Item(2) 'level
    End Sub

    Private Sub TextBox7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox7.TextChanged
        Call Caridata()

    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        On Error Resume Next
        TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        Call Panggil_Kode()
        If dr.HasRows Then
            Call Panggil_Data()
        End If
    End Sub
End Class