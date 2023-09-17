Imports System.Data.SqlClient

Public Class ActualizarDatospersonales

    Dim cnx As New Conexion()

    Public documento As String

#Region "Eventos"

    'Se ejecuta al cargar la página por primera vez
    Private Sub ActualizarDatospersonales_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DateTimePicker4.MaxDate = Today
        DateTimePicker1.MaxDate = Today.AddYears(-14)
        llenarComboDpto()
        llenarCombosCiudades()
    End Sub

    'Cargar ciudades dependiendo del dpto seleccionado
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        llenarComboCiudad(ComboBox3.SelectedIndex + 1)
    End Sub

    'Actualizar datos
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim fechaExp As String = convertFecha(DateTimePicker4.Text)
            Dim fechaNac As String = convertFecha(DateTimePicker1.Text)

            Dim cmm As New SqlCommand("sp_mant_hoja_vida", cnx.conexion)
            cmm.CommandType = CommandType.StoredProcedure

            'Todos los datos son obligatorios
            cmm.Parameters.AddWithValue("@Nombre", TextBox1.Text)
            cmm.Parameters.AddWithValue("@Apellido", TextBox2.Text)
            cmm.Parameters.AddWithValue("@Tipo_documento", ComboBox1.Text)
            cmm.Parameters.AddWithValue("@Numero_documento", TextBox3.Text)
            cmm.Parameters.AddWithValue("@Fecha_expedicion", fechaExp)
            cmm.Parameters.AddWithValue("@Lugar_expedicion", (cmbLugarExp.SelectedIndex + 1))
            cmm.Parameters.AddWithValue("@Fecha_nacimiento", fechaNac)
            cmm.Parameters.AddWithValue("@Lugar_nacimiento", (cmbLugarNac.SelectedIndex + 1))
            cmm.Parameters.AddWithValue("@Telefono", TextBox4.Text)
            cmm.Parameters.AddWithValue("@Celular", TextBox5.Text)
            cmm.Parameters.AddWithValue("@Correo_electronico", TextBox6.Text)
            cmm.Parameters.AddWithValue("@Estado_civil", ComboBox2.Text)
            cmm.Parameters.AddWithValue("@Ocupacion", TextBox7.Text)
            cmm.Parameters.AddWithValue("@Departamento", (ComboBox3.SelectedIndex + 1))
            cmm.Parameters.AddWithValue("@Ciudad", (ComboBox4.SelectedIndex + 1))
            cmm.Parameters.AddWithValue("@direccion", TextBox8.Text)
            cmm.Parameters.AddWithValue("@Numero_de_cuenta", TextBox9.Text)
            cmm.Parameters.AddWithValue("@Banco", TextBox10.Text)
            cmm.Parameters.AddWithValue("@Nombre_de_contacto", TextBox11.Text)
            cmm.Parameters.AddWithValue("@Telefono_contacto", TextBox12.Text)
            cmm.Parameters.AddWithValue("@EPS", TextBox16.Text)
            cmm.Parameters.AddWithValue("@AFP", TextBox34.Text)
            cmm.Parameters.AddWithValue("@ARL", TextBox35.Text)

            If cnx.ejecutarSP(cmm) Then
                MsgBox("Datos actualizados correctamente")
            Else
                MsgBox("Hubo un error actualizando los datos")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    'Eliminar datos
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim consulta As String = ("select id from Contrato where Nume_documento = " & documento)
        Dim contratos As List(Of String) = cnx.execSelectVarios(consulta, False)
        Dim res As Boolean = True

        If Not contratos.Count() = 0 Then
            For Each item As String In contratos
                Dim cmm As New SqlCommand("sp_eliminar_otrosi", cnx.conexion)
                cmm.CommandType = CommandType.StoredProcedure

                cmm.Parameters.AddWithValue("@id_contrato", item)
                res = cnx.ejecutarSP(cmm)
            Next
        End If

        If res Then
            Dim cm As New SqlCommand("sp_eliminar_hdv", cnx.conexion)
            cm.CommandType = CommandType.StoredProcedure
            cm.Parameters.AddWithValue("@Numero_documento", documento)

            If cnx.ejecutarSP(cm) Then
                MsgBox("El empleado fue eliminado correctamente")
            Else
                MsgBox("Hubo un error eliminando los datos")
            End If
        Else
            MsgBox("Hubo un error eliminando los datos")
        End If
    End Sub

    'Devolverse a gestión
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
        HojaVida.Show()
        HojaVida.TabPage6.Show()
    End Sub

#End Region

#Region "Métodos"

    'Método para cambiar a formato yyyy-MM-dd la fecha y poder almacenarla en la BD
    Function convertFecha(fecha As String)
        Dim vec() As String = Split(fecha, "/")
        Dim fec As String = vec(2) & "-" & vec(1) & "-" & vec(0)

        Return fec
    End Function

    'Método para llenar el combo de departamentos
    Private Sub llenarComboDpto()
        Dim consulta As String = ("select nombre from Departamentos")
        Dim departamentos As List(Of String) = cnx.execSelectVarios(consulta, False)

        If Not departamentos.Count() = 0 Then
            For Each item As String In departamentos
                ComboBox3.Items.Add(item)
            Next
        End If
    End Sub

    'Método para llenar el combo de ciudades, recibiendo el nombre del dpto
    Private Sub llenarComboCiudad(dpto As Integer)
        Dim consulta As String = ("select m.nombre from Municipios m inner join Departamentos d 
                on m.departamento_id = d.id where d.id = '" & dpto & "' order by m.nombre")
        Dim ciudades As List(Of String) = cnx.execSelectVarios(consulta, False)
        ComboBox4.Items.Clear()

        If Not ciudades.Count() = 0 Then
            For Each item As String In ciudades
                ComboBox4.Items.Add(item)
            Next
        End If
    End Sub

    'Método para llenar combos de ciudades. Lugar expedición y nacimiento
    Private Sub llenarCombosCiudades()
        Dim consulta As String = ("select nombre from Municipios order by nombre")
        Dim ciudades As List(Of String) = cnx.execSelectVarios(consulta, False)

        If Not ciudades.Count() = 0 Then
            For Each item As String In ciudades
                cmbLugarExp.Items.Add(item)
                cmbLugarNac.Items.Add(item)
            Next
        End If
    End Sub

    'Método para consultar datos y cargarlos
    Function cargarDatos()
        ActualizarDatospersonales_Load(Nothing, Nothing)
        Dim consulta As String = ("SELECT [Nombre], [Apellido], [Tipo_documento], [Numero_documento],
                [Fecha_expedicion], [Lugar_expedicion], [Fecha_nacimiento], [Lugar_nacimiento], [Telefono],
                [Celular], [Correo_electronico], [Estado_civil], [Ocupacion], [Departamento], [Ciudad],
                [direccion], [Numero_de_cuenta], [Banco], [Nombre_de_contacto], [Telefono_contacto],
                [EPS], [AFP], [ARL]
            FROM Hoja_de_vida
            WHERE Numero_documento =" & documento)

        Dim valores As List(Of String) = cnx.execSelectVarios(consulta, False)

        If Not valores.Count() = 0 Then
            TextBox1.Text = valores.ElementAt(0)
            TextBox2.Text = valores.ElementAt(1)
            ComboBox1.Text = valores.ElementAt(2)
            TextBox3.Text = valores.ElementAt(3)
            DateTimePicker4.Text = valores.ElementAt(4)
            cmbLugarExp.SelectedIndex = Convert.ToInt32(valores.ElementAt(5)) - 1
            DateTimePicker1.Text = valores.ElementAt(6)
            cmbLugarNac.SelectedIndex = Convert.ToInt32(valores.ElementAt(7)) - 1
            TextBox4.Text = valores.ElementAt(8)
            TextBox5.Text = valores.ElementAt(9)
            TextBox6.Text = valores.ElementAt(10)
            ComboBox2.Text = valores.ElementAt(11)
            TextBox7.Text = valores.ElementAt(12)
            ComboBox3.SelectedIndex = Convert.ToInt32(valores.ElementAt(13)) - 1
            ComboBox4.SelectedIndex = Convert.ToInt32(valores.ElementAt(14)) - 1
            TextBox8.Text = valores.ElementAt(15)
            TextBox9.Text = valores.ElementAt(16)
            TextBox10.Text = valores.ElementAt(17)
            TextBox11.Text = valores.ElementAt(18)
            TextBox12.Text = valores.ElementAt(19)
            TextBox16.Text = valores.ElementAt(20)
            TextBox34.Text = valores.ElementAt(21)
            TextBox35.Text = valores.ElementAt(22)
        End If
    End Function

#End Region

End Class