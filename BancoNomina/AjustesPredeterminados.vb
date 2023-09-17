Imports System.Data.SqlClient

Public Class Ajustespredeterminados

    Dim cnx As New Conexion()

#Region "Eventos"

    'Se ejecuta al cargar la página por primera vez
    Private Sub Ajustespredeterminados_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox12.Text = Today.Year.ToString()
        cargarDatos()
    End Sub

    'Editar 
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        setReadOnly(False)
    End Sub

    'Guardar 
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MsgBox("Debe ingresar el salario mínimo")
        ElseIf TextBox2.Text = "" Then
            MsgBox("Debe ingresar el auxilio de transporte")
        ElseIf TextBox3.Text = "" Then
            MsgBox("Debe ingresar los días laborados")
        ElseIf TextBox4.Text = "" Then
            MsgBox("Debe ingresar el porcentaje de hora diurna")
        ElseIf TextBox5.Text = "" Then
            MsgBox("Debe ingresar el porcentaje de hora nocturna")
        ElseIf TextBox11.Text = "" Then
            MsgBox("Debe ingresar el porcentaje de hora dominical")
        ElseIf TextBox17.Text = "" Then
            MsgBox("Debe ingresar el porcentaje de hora festiva")
        ElseIf TextBox6.Text = "" Then
            MsgBox("Debe ingresar el porcentaje de recargo nocturno")
        ElseIf TextBox7.Text = "" Then
            MsgBox("Debe ingresar el porcentaje de recargo dominical")
        ElseIf TextBox8.Text = "" Then
            MsgBox("Debe ingresar el porcentaje de recargo festivo")
        ElseIf TextBox9.Text = "" Then
            MsgBox("Debe ingresar el porcentaje de salud")
        ElseIf TextBox10.Text = "" Then
            MsgBox("Debe ingresar el porcentaje de pensión")
        Else
            Try
                Dim cmm As New SqlCommand("sp_mant_valores_pred", cnx.conexion)
                cmm.CommandType = CommandType.StoredProcedure

                'Todos los datos son obligatorios
                cmm.Parameters.AddWithValue("@Anio", Math.Round(Convert.ToInt32(TextBox12.Text), 2))
                cmm.Parameters.AddWithValue("@Salario_minimo", Math.Round(Convert.ToInt32(TextBox1.Text), 2))
                cmm.Parameters.AddWithValue("@Aux_transporte", Math.Round(Convert.ToInt32(TextBox2.Text), 2))
                cmm.Parameters.AddWithValue("@Salario_dia", Math.Round(Convert.ToInt32(TextBox3.Text), 2))
                cmm.Parameters.AddWithValue("@Hora_diurna", Math.Round(Convert.ToInt32(TextBox4.Text), 2))
                cmm.Parameters.AddWithValue("@Hora_nocturna", Math.Round(Convert.ToInt32(TextBox5.Text), 2))
                cmm.Parameters.AddWithValue("@Hora_dominical", Math.Round(Convert.ToInt32(TextBox11.Text), 2))
                cmm.Parameters.AddWithValue("@Hora_festivo", Math.Round(Convert.ToInt32(TextBox17.Text), 2))
                cmm.Parameters.AddWithValue("@Recargo_nocturno", Math.Round(Convert.ToInt32(TextBox6.Text), 2))
                cmm.Parameters.AddWithValue("@Recargo_dominical", Math.Round(Convert.ToInt32(TextBox7.Text), 2))
                cmm.Parameters.AddWithValue("@Recargo_festivo", Math.Round(Convert.ToInt32(TextBox8.Text), 2))
                cmm.Parameters.AddWithValue("@Salud", Math.Round(Convert.ToInt32(TextBox9.Text), 2))
                cmm.Parameters.AddWithValue("@Pension", Math.Round(Convert.ToInt32(TextBox10.Text), 2))

                If cnx.ejecutarSP(cmm) Then
                    MsgBox("Valores agregados o modificados correctamente")
                    setReadOnly(True)
                Else
                    MsgBox("Hubo un error agregando o modificando los datos")
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    'Cancelar
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        setReadOnly(True)
        Me.Hide()
        menum.Show()
    End Sub

#End Region

#Region "Métodos"

    'Método para establecer propiedad readonly
    Private Sub setReadOnly(valor As Boolean)
        TextBox1.ReadOnly = valor
        TextBox2.ReadOnly = valor
        TextBox3.ReadOnly = valor
        TextBox4.ReadOnly = valor
        TextBox5.ReadOnly = valor
        TextBox11.ReadOnly = valor
        TextBox17.ReadOnly = valor
        TextBox6.ReadOnly = valor
        TextBox7.ReadOnly = valor
        TextBox8.ReadOnly = valor
        TextBox9.ReadOnly = valor
        TextBox10.ReadOnly = valor
    End Sub

    'Método para consultar datos y cargarlos
    Private Sub cargarDatos()
        Dim consulta As String = ("SELECT [Salario_minimo], [Aux_transporte], [Hora_diurna],
                [Hora_nocturna], [Hora_dominical], [Hora_festivo], [Recargo_nocturno], [Recargo_dominical],
                [Recargo_festivo], [Salud], [Pension] FROM ValoresxAnio WHERE [Anio] = " & TextBox12.Text)

        Dim valores As List(Of String) = cnx.execSelectVarios(consulta, False)

        If Not valores.Count() = 0 Then
            TextBox1.Text = valores.ElementAt(0)
            TextBox2.Text = valores.ElementAt(1)
            TextBox3.Text = valores.ElementAt(2)
            TextBox4.Text = valores.ElementAt(3)
            TextBox5.Text = valores.ElementAt(4)
            TextBox11.Text = valores.ElementAt(5)
            TextBox17.Text = valores.ElementAt(6)
            TextBox6.Text = valores.ElementAt(7)
            TextBox7.Text = valores.ElementAt(8)
            TextBox8.Text = valores.ElementAt(9)
            TextBox9.Text = valores.ElementAt(10)
            TextBox10.Text = valores.ElementAt(11)
        End If
    End Sub

#End Region

End Class