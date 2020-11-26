Public Class frmmenu

    Private Sub SalirToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SalirToolStripMenuItem1.Click

        If MsgBox("¿Desea salir de la aplicación?", MsgBoxStyle.YesNo,
"Cuidado") = MsgBoxResult.Yes Then Me.Close()

    End Sub

    Private Sub NominaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NominaToolStripMenuItem.Click

        frmnomina.Show()

    End Sub

    Private Sub AcercaDeToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles AcercaDeToolStripMenuItem1.Click

        frmacercade.Show()

    End Sub

    Private Sub frmmenu_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class