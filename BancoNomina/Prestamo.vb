Imports System.Data.SqlClient

Public Class Prestamo

    Dim cnx As New Conexion()

#Region "Eventos"

    'Se ejcuta al cargar la página por primera vez
    Private Sub Prestamo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DateTimePicker4.Value = Today
        DateTimePicker1.MinDate = Today
    End Sub

    'Calcular los días
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Dim fechaIni As Date = Today
        Dim fechaFin As Date
        fechaFin = DateTimePicker1.Value
        Dim tiempo As Integer = (fechaFin - fechaIni.AddDays(0)).Days
        If Not (tiempo = 0) Then
                Label23.Text = tiempo.ToString()
            Else
                Label23.Text = "..."
            End If
    End Sub

    'Guardar
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox3.Text = "" Then
            MsgBox("Debe ingresar el número de documento")
        ElseIf TextBox1.Text = "" Then
            MsgBox("Debe ingresar el valor del préstamo")
        ElseIf ComboBox1.Text = "" Then
            MsgBox("Debe seleccionar un método de pago")
        Else
            Try
                Dim cmm As New SqlCommand("sp_mant_prestamo", cnx.conexion)
                cmm.CommandType = CommandType.StoredProcedure

                Dim fecha As String = convertFecha(DateTimePicker1.Text)

                'Todos los datos son obligatorios
                cmm.Parameters.AddWithValue("@Num_documento", TextBox3.Text)
                cmm.Parameters.AddWithValue("@Fecha_fin_prestamo", fecha)
                cmm.Parameters.AddWithValue("@Metodo_pago", ComboBox1.SelectedItem.ToString())
                cmm.Parameters.AddWithValue("@Valor_prestamo", Math.Round(Convert.ToSingle(TextBox1.Text), 2))

                If cnx.ejecutarSP(cmm) Then
                    MsgBox("Préstamo agregado correctamente")
                Else
                    MsgBox("Hubo un error agregando los datos")
                End If
            Catch ex As Exception
                'MsgBox(ex.Message)
            End Try
        End If
    End Sub

    'Cancelar
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
        menum.Show()
    End Sub
#End Region

#Region "Métodos"

    'Método para cambiar a formato yyyy-MM-dd la fecha y poder almacenarla en la BD
    Function convertFecha(fecha As String)
        Dim vec() As String = Split(fecha, "/")
        Dim fec As String = vec(2) & "-" & vec(1) & "-" & vec(0)

        Return fec
    End Function

#End Region

End Class