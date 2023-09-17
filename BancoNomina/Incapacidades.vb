Public Class Incapacidades

    Public vlrNomina As Double
    Dim vlrIncapacidad As Double = 0

    Private Sub DateTimePicker4_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker4.ValueChanged
        DateTimePicker1.MinDate = DateTimePicker4.Value
        Dim tiempo As Integer = calcularDias()
        If Not (tiempo = 0) Then
            Label23.Text = tiempo.ToString()
        Else
            Label23.Text = "..."
        End If
    End Sub

    'Guadar
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim tiempo As Integer = calcularDias()
        If tiempo > 3 Then
            vlrIncapacidad = vlrNomina * (66.66 / 100)
        End If

        Nomina.totalIncapacidades(Math.Round(vlrIncapacidad, 2))
        Me.Hide()
        Nomina.Show()
        Nomina.TabPage1.Show()
    End Sub

    'Método para cambiar a formato yyyy-MM-dd la fecha y poder almacenarla en la BD
    Function convertFecha(fecha As String)
        Dim vec() As String = Split(fecha, "/")
        Dim fec As String = vec(2) & "-" & vec(1) & "-" & vec(0)

        Return fec
    End Function

    Function calcularDias()
        Dim fechaIni As Date
        Dim fechaFin As Date
        fechaIni = DateTimePicker4.Value
        fechaFin = DateTimePicker1.Value

        Return (fechaFin - fechaIni.AddDays(-1)).Days
    End Function

    'Calcular los días
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Dim tiempo As Integer = calcularDias()
        If Not (tiempo = 0) Then
            Label23.Text = tiempo.ToString()
        Else
            Label23.Text = "..."
        End If
    End Sub
End Class