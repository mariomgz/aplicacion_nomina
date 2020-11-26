Public Class LoginForm1

    Dim i As Integer

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click


        ''While (1 < 3)

        If (txtusuario.Text = "nomina" Or txtusuario.Text = "NOMINA") And txtcontrasena.Text = "123" Then

            frmmenu.Show()
            ''MsgBox("BIENVENIDO :)")
            Me.Close()

        Else

            MsgBox("Clave Invalida")
            txtusuario.Text = ""
            txtcontrasena.Text = ""
            txtusuario.Focus()
            i = i + 1

            If i = 3 Then

                MsgBox("Detectado")
                Me.Close()

            End If

            ''Exit While

        End If

        ''End While

    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

End Class
