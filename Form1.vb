
Public Class Form1

    Private Sub txtUTS_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUTS.TextChanged
        Const UTS As Single = 0.2
        Dim rata_uts As Single
        rata_uts = txtUTS.Text
        txtrataUTS.Text = UTS * rata_uts
    End Sub

    Private Sub txtUAS_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUAS.TextChanged
        Const UAS As Single = 0.2
        Dim rata_uas As Single
        rata_uas = txtUAS.Text
        txtrataUAS.Text = UAS * rata_uas
    End Sub

    Private Sub txtTP_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTP.TextChanged
        Const Project As Single = (25 / 100)
        Dim rata_project As Single
        rata_project = txtTP.Text
        txtrataTP.Text = Project * rata_project
    End Sub

    Private Sub txtPP_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPP.TextChanged
        Const Presentasi_Project As Single = (15 / 100)
        Dim rata_pp As Single
        rata_pp = txtPP.Text
        txtrataPP.Text = Presentasi_Project * rata_pp
    End Sub

    Private Sub txtaktif_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtaktif.TextChanged
        Const Keaktifan As Single = (10 / 100)
        Dim rata_keaktifan As Single
        rata_keaktifan = txtaktif.Text
        txtrataaktif.Text = Keaktifan * rata_keaktifan
    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtpresent.TextChanged
        Const Presentasi As Single = (10 / 100)
        Dim rata_presentasi As Single
        rata_presentasi = txtpresent.Text
        txtrataP.Text = Presentasi * rata_presentasi
    End Sub

    Private Sub cmdHASIL_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHASIL.Click
        txtRATA.Text = Val(txtrataUTS.Text) + Val(txtrataUAS.Text) + Val(txtrataTP.Text) + Val(txtrataPP.Text) + Val(txtrataaktif.Text) + Val(txtrataP.Text)
        If txtRATA.Text >= 85 Then
            txtHURUF.Text = "A"
        ElseIf txtRATA.Text >= 75 Then
            txtHURUF.Text = "B"
        ElseIf txtRATA.Text >= 60 Then
            txtHURUF.Text = "C"
        ElseIf txtRATA.Text >= 50 Then
            txtHURUF.Text = "D"
        Else
            txtHURUF.Text = "E"
        End If

    End Sub
End Class


