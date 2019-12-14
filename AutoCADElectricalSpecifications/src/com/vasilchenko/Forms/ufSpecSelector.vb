Namespace com.vasilchenko.Forms
    Public Class SpecSelector
        Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
            rbProjCreate.Checked = False
            rbProjUpdate.Checked = False
            rbSheetCreate.Checked = False
            rbSheetUpdate.Checked = False
            rbPageCreate.Checked = False
            rbPageUpdate.Checked = False
            Close()
        End Sub

        Private Sub btnOk_Click(sender As Object, e As EventArgs) Handles btnOk.Click
            Close()
        End Sub

        Private Sub rbProjCreate_Click(sender As Object, e As EventArgs) Handles rbProjCreate.Click
            rbProjUpdate.Checked = False
            rbSheetCreate.Checked = False
            rbSheetUpdate.Checked = False
            rbPageCreate.Checked = False
            rbPageUpdate.Checked = False
        End Sub
        Private Sub rbProjUpdate_Click(sender As Object, e As EventArgs) Handles rbProjUpdate.Click
            rbProjCreate.Checked = False
            rbSheetCreate.Checked = False
            rbSheetUpdate.Checked = False
            rbPageCreate.Checked = False
            rbPageUpdate.Checked = False
        End Sub
        Private Sub rbSheetCreate_Click(sender As Object, e As EventArgs) Handles rbSheetCreate.Click
            rbProjUpdate.Checked = False
            rbProjUpdate.Checked = False
            rbSheetUpdate.Checked = False
            rbPageCreate.Checked = False
            rbPageUpdate.Checked = False
        End Sub
        Private Sub rbSheetUpdate_Click(sender As Object, e As EventArgs) Handles rbSheetUpdate.Click
            rbProjUpdate.Checked = False
            rbSheetCreate.Checked = False
            rbProjCreate.Checked = False
            rbPageCreate.Checked = False
            rbPageUpdate.Checked = False
        End Sub

        Private Sub rbPageCreate_CheckedChanged(sender As Object, e As EventArgs) Handles rbPageCreate.CheckedChanged
            rbProjCreate.Checked = False
            rbProjUpdate.Checked = False
            rbSheetCreate.Checked = False
            rbSheetUpdate.Checked = False
            rbPageUpdate.Checked = False
        End Sub

        Private Sub rbPageUpdate_CheckedChanged(sender As Object, e As EventArgs) Handles rbPageUpdate.CheckedChanged
            rbProjCreate.Checked = False
            rbProjUpdate.Checked = False
            rbSheetCreate.Checked = False
            rbSheetUpdate.Checked = False
            rbPageCreate.Checked = False
        End Sub
    End Class
End Namespace