Public Class frmnomina

    'Función Limpiar Datos

    Function LimpiarDatosIngreso()

        Cboxcedula.Text = ""
        txtnombre.Text = ""
        txtsueldo.Text = ""
        cboxriesgo.Text = ""
        txtdias.Text = ""
        txthed.Text = 0
        txthen.Text = 0
        txthedd.Text = 0
        txthedn.Text = 0
        txthrn.Text = 0
        lblcedulap.Text = ""
        lblempleadop.Text = ""
        lblcedulap.Text = ""
        lblsueldop.Text = ""
        lblbasicop.Text = ""
        lblauxiliop.Text = ""
        lblextrasp.Text = ""
        lbldevengadop.Text = ""
        lblsaludp.Text = ""
        lblpensionp.Text = ""
        lblfondop.Text = ""
        lblretefuentep.Text = ""
        lbldeduccionesp.Text = ""
        lblnetop.Text = ""

    End Function

    ''Ingresar Cédula, Nombre y Sueldo de cada empleado desde un ComboBox

    Private Sub Cboxcedula_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cboxcedula.SelectedIndexChanged

        Dim cedula As String = Cboxcedula.Text

        Select Case cedula
            Case "41724736"
                txtnombre.Text = "Claudia Patricia"
                txtsueldo.Text = Format(10000000, "###,###.##")
            Case "52984576"
                txtnombre.Text = "Andrea Cruz "
                txtsueldo.Text = Format(950000, "###,###.##")
            Case "79899233"
                txtnombre.Text = "Alejandro Acevedo"
                txtsueldo.Text = Format(6000000, "###,###.##")
            Case "85890669"
                txtnombre.Text = "Camilo Rodríguez"
                txtsueldo.Text = Format(1200000, "###,###.##")
            Case "1000123456"
                txtnombre.Text = "Angélica Blanco"
                txtsueldo.Text = Format(4500000, "###,###.##")
            Case "1000147258"
                txtnombre.Text = "Diego Forero"
                txtsueldo.Text = Format(15000000, "###,###.##")
            Case "1000258369"
                txtnombre.Text = "Gabriel Nieto"
                txtsueldo.Text = Format(3800000, "###,###.##")
            Case "1000456789"
                txtnombre.Text = "Gloria Mendoza"
                txtsueldo.Text = Format(2100000, "###,###.##")
            Case "1000789123"
                txtnombre.Text = "Juliana Gaviria"
                txtsueldo.Text = Format(3000000, "###,###.##")
            Case "1000963852"
                txtnombre.Text = "Jorge Orozco"
                txtsueldo.Text = Format(2500000, "###,###.##")

        End Select


        cboxriesgo.Focus()

    End Sub

    Private Sub cboxriesgo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboxriesgo.SelectedIndexChanged

        txtdias.Focus()

    End Sub

    'Adicionar Datos en DataGridview

    Private n As Integer = 0

    Private Sub btnadicionarempleado_Click(sender As Object, e As EventArgs) Handles btnadicionarempleado.Click

        If Cboxcedula.Text.Trim = String.Empty Then
            MsgBox("Introduzca un número de cédula", MsgBoxStyle.OkOnly)
            Return
            Cboxcedula.Focus()

        ElseIf cboxriesgo.Text.Trim = String.Empty Then
            MsgBox("Introduzca el Nivel de Riesgo", MsgBoxStyle.OkOnly)
            Return
            cboxriesgo.Focus()

        ElseIf txtdias.Text.Trim = String.Empty Then
            MsgBox("Introduzca el Número de días", MsgBoxStyle.OkOnly)
            Return
            txtdias.Focus()
        
        ElseIf CboxMes.Text.Trim = String.Empty Then
            MsgBox("Introduzca el Mes a Liquidar", MsgBoxStyle.OkOnly)
            Return
            CboxMes.Focus()

        End If


        If Val(txtdias.Text) > 30 Then

            MsgBox("El número de días no puede ser mayor de 30", MsgBoxStyle.Information, "Error")
            txtdias.Clear()
            Return
        End If

        Dim basico As Integer
        Dim auxilio As Integer
        Dim sueldo As Integer = txtsueldo.Text
        Dim dias As Integer = Val(txtdias.Text)
        Dim vrhed As Integer = Val(txthed.Text)
        Dim vrhen As Integer = Val(txthen.Text)
        Dim vrhedd As Integer = Val(txthedd.Text)
        Dim vrhedn As Integer = Val(txthedn.Text)
        Dim vrhrn As Integer = Val(txthrn.Text)

        n = dtgrempleados.Rows.Add()

        dtgrempleados.Rows(n).Cells(0).Value = Cboxcedula.Text
        dtgrempleados.Rows(n).Cells(1).Value = txtnombre.Text
        dtgrempleados.Rows(n).Cells(2).Value = txtsueldo.Text
        dtgrempleados.Rows(n).Cells(3).Value = txtdias.Text
        dtgrempleados.Rows(n).Cells(4).Value = txthed.Text
        dtgrempleados.Rows(n).Cells(5).Value = txthen.Text
        dtgrempleados.Rows(n).Cells(6).Value = txthedd.Text
        dtgrempleados.Rows(n).Cells(7).Value = txthedn.Text
        dtgrempleados.Rows(n).Cells(8).Value = txthrn.Text

        'Devengado

        basico = sueldo / 30 * dias

        dtgrempleados.Rows(n).Cells(9).Value = basico

        If sueldo <= (2 * 877803) Then

            auxilio = 102854 / 30 * dias

        End If

        dtgrempleados.Rows(n).Cells(10).Value = auxilio

        Dim vrhora As Integer = sueldo / 240

        dtgrempleados.Rows(n).Cells(11).Value = vrhora

        Dim vrhed1 As Integer = 1.25 * vrhed * vrhora
        Dim vrhen1 As Integer = 1.75 * vrhen * vrhora
        Dim vrhedd1 As Integer = 2 * vrhedd * vrhora
        Dim vrhedn1 As Integer = 2.5 * vrhedn * vrhora
        Dim vrhrn1 As Integer = 1.35 * vrhrn * vrhora

        dtgrempleados.Rows(n).Cells(12).Value = vrhed1
        dtgrempleados.Rows(n).Cells(13).Value = vrhen1
        dtgrempleados.Rows(n).Cells(14).Value = vrhedd1
        dtgrempleados.Rows(n).Cells(15).Value = vrhedn1
        dtgrempleados.Rows(n).Cells(16).Value = vrhrn1

        Dim textras As Integer = vrhed1 + vrhen1 + vrhedd1 + vrhedn1 + vrhrn1
        Dim totaldevengado As Double = basico + auxilio + textras

        'Deducciones

        Dim ibc As Double = totaldevengado - auxilio
        Dim salud As Integer = ibc * 4 / 100
        Dim pension As Integer = ibc * 4 / 100
        Dim fondo As Double

        Select Case ibc
            Case 3511212 To 14044088
                fondo = Math.Round(ibc * 0.01)
            Case 14044089 To 14922651
                fondo = Math.Round(ibc * 0.012)
            Case 14922652 To 15800454
                fondo = Math.Round(ibc * 0.014)
            Case 15800455 To 16678257
                fondo = Math.Round(ibc * 0.016)
            Case 16678258 To 17556060
                fondo = Math.Round(ibc * 0.018)
            Case > 17556060
                fondo = Math.Round(ibc * 0.02)

        End Select

        dtgrempleados.Rows(n).Cells(17).Value = textras
        dtgrempleados.Rows(n).Cells(18).Value = totaldevengado
        dtgrempleados.Rows(n).Cells(19).Value = ibc
        dtgrempleados.Rows(n).Cells(20).Value = salud
        dtgrempleados.Rows(n).Cells(21).Value = pension
        dtgrempleados.Rows(n).Cells(22).Value = fondo

        Dim uvt As Double = 35607
        Dim retefuente As Integer
        Dim baseretefuente As Double = (totaldevengado - salud - pension - fondo) * 0.75
        Dim numuvt As Double = baseretefuente / uvt

        Select Case numuvt
            Case 0 To 95
                retefuente = 0
            Case 96 To 150
                retefuente = ((numuvt - 95) * 0.19) * uvt
            Case 151 To 360
                retefuente = ((numuvt - 150) * 0.28 + 10) * uvt
            Case 361 To 640
                retefuente = ((numuvt - 360) * 0.33 + 69) * uvt
            Case 641 To 945
                retefuente = ((numuvt - 640) * 0.35 + 162) * uvt
            Case 946 To 2300
                retefuente = ((numuvt - 945) * 0.37 + 268) * uvt
            Case > 2300
                retefuente = ((numuvt - 2300) * 0.39 + 770) * uvt

        End Select

        Dim totaldeducciones As Double = Math.Round(salud + pension + fondo + retefuente)
        Dim neto As Double = Math.Round(totaldevengado - totaldeducciones)

        dtgrempleados.Rows(n).Cells(23).Value = numuvt
        dtgrempleados.Rows(n).Cells(24).Value = retefuente
        dtgrempleados.Rows(n).Cells(25).Value = totaldeducciones
        dtgrempleados.Rows(n).Cells(26).Value = neto

        'Parafiscales

        n = dtgrempleados2.Rows.Add()

        Dim saludP As Integer = ibc * 8.5 / 100
        Dim pensionP As Integer = ibc * 12 / 100
        Dim riesgo As String = cboxriesgo.Text
        Dim arpp As Integer

        Select Case riesgo
            Case "Clase I"
                arpp = ibc * 0.522 / 100
            Case "Clase II"
                arpp = ibc * 1.044 / 100
            Case "Clase III"
                arpp = ibc * 2.436 / 100
            Case "Clase IV"
                arpp = ibc * 4.35 / 100
            Case "Clase V"
                arpp = ibc * 6.96 / 100

        End Select

        Dim senaP As Integer = ibc * 2 / 100
        Dim icbfP As Integer = ibc * 3 / 100
        Dim cajaP As Integer = ibc * 4 / 100
        Dim totalparaf As Integer = saludP + pensionP + arpp + senaP + icbfP + cajaP
        Dim prima As Integer = ibc * 8.33 / 100
        Dim vacaciones As Integer = ibc * 4.17 / 100
        Dim cesantias As Integer = ibc * 8.33 / 100
        Dim intcesantias As Integer = cesantias * 1 / 100
        Dim totalprest As Integer = prima + vacaciones + cesantias + intcesantias
        Dim totalnomina As Integer = neto + totalparaf + totalprest


        dtgrempleados2.Rows(n).Cells(0).Value = Cboxcedula.Text
        dtgrempleados2.Rows(n).Cells(1).Value = txtnombre.Text
        dtgrempleados2.Rows(n).Cells(2).Value = saludP
        dtgrempleados2.Rows(n).Cells(3).Value = pensionP
        dtgrempleados2.Rows(n).Cells(4).Value = arpp
        dtgrempleados2.Rows(n).Cells(5).Value = senaP
        dtgrempleados2.Rows(n).Cells(6).Value = icbfP
        dtgrempleados2.Rows(n).Cells(7).Value = cajaP
        dtgrempleados2.Rows(n).Cells(8).Value = totalparaf
        dtgrempleados2.Rows(n).Cells(9).Value = prima
        dtgrempleados2.Rows(n).Cells(10).Value = vacaciones
        dtgrempleados2.Rows(n).Cells(11).Value = cesantias
        dtgrempleados2.Rows(n).Cells(12).Value = intcesantias
        dtgrempleados2.Rows(n).Cells(13).Value = totalprest
        dtgrempleados2.Rows(n).Cells(14).Value = totalnomina

        LimpiarDatosIngreso()
        Cboxcedula.Focus()

        If Val(txtsueldo.Text) > 0 And Val(txtdias.Text) > 0 Then

            btnadicionarempleado.Enabled = True

        End If

        lblfecha.Text = CboxMes.Text

    End Sub


    Private Sub btnborrarregistro_Click(sender As Object, e As EventArgs) Handles btnborrarregistro.Click

        CboxMes.Text = ""
        LimpiarDatosIngreso()
        dtgrempleados.Rows.Clear()
        dtgrempleados2.Rows.Clear()
        lblfecha.Text = ""
        CboxMes.Enabled = True

    End Sub

    Private Sub frmnomina_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        txthed.Text = 0
        txthen.Text = 0
        txthedd.Text = 0
        txthedn.Text = 0
        txthrn.Text = 0

        Cboxcedula.Focus()

    End Sub

    'Comprobante de Pago

    Private Sub dtgrempleados_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dtgrempleados.CellClick

        lblempleadop.Text = dtgrempleados.CurrentRow.Cells(0).Value
        lblcedulap.Text = dtgrempleados.CurrentRow.Cells(1).Value
        lblsueldop.Text = dtgrempleados.CurrentRow.Cells(2).Value
        lblbasicop.Text = Format(dtgrempleados.CurrentRow.Cells(9).Value, "###,###.##")
        lblauxiliop.Text = Format(dtgrempleados.CurrentRow.Cells(10).Value, "###,###.##")
        lblextrasp.Text = Format(dtgrempleados.CurrentRow.Cells(17).Value, "###,###.##")
        lbldevengadop.Text = Format(dtgrempleados.CurrentRow.Cells(18).Value, "###,###.##")
        lblsaludp.Text = Format(dtgrempleados.CurrentRow.Cells(20).Value, "###,###.##")
        lblpensionp.Text = Format(dtgrempleados.CurrentRow.Cells(21).Value, "###,###.##")
        lblfondop.Text = Format(dtgrempleados.CurrentRow.Cells(22).Value, "###,###.##")
        lblretefuentep.Text = Format(dtgrempleados.CurrentRow.Cells(24).Value, "###,###.##")
        lbldeduccionesp.Text = Format(dtgrempleados.CurrentRow.Cells(25).Value, "###,###.##")
        lblnetop.Text = Format(dtgrempleados.CurrentRow.Cells(26).Value, "###,###.##")

    End Sub

    'Validación Datos

    Private Sub txtdias_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtdias.KeyPress

        If Char.IsNumber(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            MsgBox("Debe de ingresar sólo número.", MsgBoxStyle.Information, "Error")
            txtdias.Clear()
            e.Handled = True

        End If

    End Sub

    Private Sub txthed_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txthed.KeyPress

        If Char.IsNumber(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            MsgBox("Debe de ingresar sólo número.", MsgBoxStyle.Information, "Error")
            txthed.Clear()
            e.Handled = True

        End If

    End Sub

    Private Sub txthen_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txthen.KeyPress

        If Char.IsNumber(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            MsgBox("Debe de ingresar sólo número.", MsgBoxStyle.Information, "Error")
            txthen.Clear()
            e.Handled = True

        End If

    End Sub

    Private Sub txthedd_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txthedd.KeyPress

        If Char.IsNumber(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            MsgBox("Debe de ingresar sólo número.", MsgBoxStyle.Information, "Error")
            txthedd.Clear()
            e.Handled = True

        End If

    End Sub

    Private Sub txthedn_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txthedn.KeyPress

        If Char.IsNumber(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            MsgBox("Debe de ingresar sólo número.", MsgBoxStyle.Information, "Error")
            txthedn.Clear()
            e.Handled = True

        End If

    End Sub

    Private Sub txthrn_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txthrn.KeyPress

        If Char.IsNumber(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            MsgBox("Debe de ingresar sólo número.", MsgBoxStyle.Information, "Error")
            txthrn.Clear()
            e.Handled = True

        End If

    End Sub

    Dim resultado As MsgBoxResult

    Private Sub frmnomina_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        resultado = CType(MsgBox("¿Desea regresar al Menú Principal?", MsgBoxStyle.YesNo,
"Cuidado"), MsgBoxResult)

        If resultado = MsgBoxResult.No Then
            e.Cancel = True

        Else
            e.Cancel = False
            frmmenu.Show()

        End If

    End Sub

    Private Sub CboxMes_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboxMes.SelectedIndexChanged

        CboxMes.Enabled = False

    End Sub
End Class
