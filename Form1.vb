Public Class Form1

    Private Sub KeluarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KeluarToolStripMenuItem.Click
        End
    End Sub


    Sub terkunci()
        LoginToolStripMenuItem.Enabled = True
        LogoutToolStripMenuItem.Enabled = False
        MasterToolStripMenuItem.Enabled = False
        TransaksiToolStripMenuItem.Enabled = False
        UtilityToolStripMenuItem.Enabled = False
        LaporanToolStripMenuItem.Enabled = False

    End Sub


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call terkunci()

    End Sub

    Private Sub LoginToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoginToolStripMenuItem.Click
        Form_Login.Show()

    End Sub

    Private Sub LogoutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogoutToolStripMenuItem.Click
        Call terkunci()
    End Sub

    Private Sub PetugasToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PetugasToolStripMenuItem.Click
        FormMaster_User.ShowDialog()
    End Sub

    Private Sub AnggotaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        FormMasterAnggota.Show()

    End Sub
End Class
