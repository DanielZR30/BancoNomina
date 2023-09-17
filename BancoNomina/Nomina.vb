Imports System.Data.SqlClient
Imports System.IO

Public Class Nomina

#Region "Variables"
    Dim cnx As New Conexion()
    Dim empColilla As List(Of String) = New List(Of String)
    Dim empCuenta As List(Of String) = New List(Of String)
    Dim empTodos As List(Of String) = New List(Of String)
    Dim documentos As List(Of Byte()) = New List(Of Byte())
    Dim auxTrans As Single
    Dim porSalud As Single
    Dim porPension As Single
    Public vrHrsDiurnas As Double
    Public vrHrsNocturnas As Double
    Public vrHrsDominicales As Double
    Public vrHrsFestivas As Double
    Dim j As Integer = 0
    Dim i As Integer = 0
#End Region

#Region "Eventos"

    'Se ejecuta cuando la página se carga por primera vez
    Private Sub Nomina_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DateTimePicker3.Value = Today.AddDays(16)
        DateTimePicker1.Value = Today.AddDays(31)
        calcularDias(1)
        calcularDias(2)
    End Sub

    'Redirigirse a plantilla horas extras
    Private Sub TextBox10_Click(sender As Object, e As EventArgs) Handles TextBox10.Click
        Me.Hide()
        Horasextras.vlrNomina = Convert.ToSingle(TextBox2.Text)
        Horasextras.Show()
    End Sub

    'Redirigirse a la plantilla recargos
    Private Sub TextBox9_Click(sender As Object, e As EventArgs) Handles TextBox9.Click
        Me.Hide()
        Recargos.vlrNomina = Convert.ToSingle(TextBox2.Text)
        Recargos.Show()
    End Sub

    'Redirigirse a la plantilla incapacidades
    Private Sub TextBox14_Click(sender As Object, e As EventArgs) Handles TextBox14.Click
        Me.Hide()
        Incapacidades.vlrNomina = (Convert.ToSingle(TextBox2.Text) / 30) * 15
        Incapacidades.Show()
    End Sub

    'Calcular días trabajados de acuerdo a las fechas de inicio y fin 
    Private Sub DateTimePicker3_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker3.ValueChanged
        calcularDias(1)
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        calcularDias(1)
    End Sub

    'Calcular días trabajados de acuerdo a las fechas de inicio y fin 
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        calcularDias(2)
    End Sub

    Private Sub DateTimePicker4_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker4.ValueChanged
        calcularDias(2)
    End Sub

    'Calcular días trabajados de acuerdo a las fechas de inicio y fin 
    Private Sub DateTimePicker5_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker5.ValueChanged
        calcularDias(3)
    End Sub

    Private Sub DateTimePicker6_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker6.ValueChanged
        calcularDias(3)
    End Sub

    'Tomar número documento al dar doble click en el grid
    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If e.RowIndex > -1 Then
            Dim numDoc As String = empColilla.ElementAt(e.RowIndex).ToString()
            limpiar(1)
            cargarDatosColilla(numDoc)
        End If
    End Sub

    'Tomar número documento al dar doble click en el grid
    Private Sub DataGridView2_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellDoubleClick
        If e.RowIndex > -1 Then
            Dim numDoc As String = empCuenta.ElementAt(e.RowIndex).ToString()
            limpiar(2)
            cargarDatosCuenta(numDoc)
        End If
    End Sub

    'Tomar número documento al dar doble click en el grid
    Private Sub DataGridView3_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellDoubleClick
        If e.RowIndex > -1 Then
            Dim numDoc As String = empTodos.ElementAt(e.RowIndex).ToString()
            limpiar(3)
            cargarDatosPrestaciones(numDoc)
        End If
    End Sub

    'Buscar colilla de pago
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        limpiar(1)
        cargarDatosColilla(TextBox1.Text)
    End Sub

    'Buscar cuenta de cobro
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        limpiar(2)
        cargarDatosCuenta(TextBox40.Text)
    End Sub

    'Buscar liquidaciones
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        limpiar(3)
        If Not cnx.execSelect("select numero_documento from hoja_de_vida where numero_documento = " & TextBox42.Text) = 0 Then
            cargarDatosPrestaciones(TextBox42.Text)
        Else
            MsgBox("No existe un empleado con el documento " & TextBox42.Text)
        End If
    End Sub

    'Calcular totales al entrar al formulario
    Private Sub TabPage5_Enter(sender As Object, e As EventArgs) Handles TabPage5.Enter
        calcularTotales()
    End Sub

    'Entrar al tabcontrol1 se cargan datos
    Private Sub TabControl1_Enter(sender As Object, e As EventArgs) Handles TabControl1.Enter
        empColilla.Clear()
        empCuenta.Clear()
        empTodos.Clear()
        consultaNominaColilla()
        consultaNominaCuenta()
        consultaLiquidaciones()
    End Sub

    'Combo liquidación cambia
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If Not ComboBox1.Text = "" Then
            Dim val As String = ComboBox1.SelectedIndex
            If Not Label62.Text = "..." Then
                If Not TextBox42.Text = "" Then
                    calcularPrestaciones(val)
                Else
                    MsgBox("El campo documento no puede estar vacío")
                End If
            Else
                MsgBox("Debe establecer una fecha inicial y una fecha final")
                ComboBox1.Text = ""
            End If
        End If
    End Sub

    'Cargar links de documentos al entrar a la plantilla
    Private Sub TabPage4_Enter(sender As Object, e As EventArgs) Handles TabPage4.Enter
        AxAcroPDF5.LoadFile("none")
        Button18.Enabled = False
        i = 0
        j = 0
        cargarLinksDocs("1999-09-19")
    End Sub

    'Filtrar links por fecha
    Private Sub DateTimePicker7_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker7.ValueChanged
        AxAcroPDF5.LoadFile("none")
        Button18.Enabled = False
        i = 0
        j = 0
        cargarLinksDocs(DateTimePicker7.Text)
    End Sub

#End Region

#Region "Métodos"

    'Método para regresar al menú
    Private Sub salir()
        Me.Hide()
        menum.Show()
    End Sub

    'Método para buscar valores predeterminados
    Private Sub buscarValores()
        Dim consulta As String = "select Aux_transporte, Salud, Pension from ValoresxAnio"
        Dim datos As List(Of String) = cnx.execSelectVarios(consulta, False)

        If Not datos.Count() = 0 Then
            auxTrans = Convert.ToSingle(datos.ElementAt(0).ToString())
            porSalud = Convert.ToSingle(datos.ElementAt(1).ToString())
            porPension = Convert.ToSingle(datos.ElementAt(2).ToString())
        End If
    End Sub

    'Método para calcular días
    Private Sub calcularDias(identificador As Integer)
        Dim fechaIni As Date
        Dim fechaFin As Date
        Dim tiempo As Integer

        If identificador = 1 Then
            fechaIni = DateTimePicker2.Value
            fechaFin = DateTimePicker3.Value
            tiempo = (fechaFin - fechaIni).Days
            If Not (tiempo = 0) Then
                Label23.Text = tiempo.ToString()
            Else
                Label23.Text = "..."
            End If
        ElseIf identificador = 2 Then
            fechaIni = DateTimePicker4.Value
            fechaFin = DateTimePicker1.Value
            tiempo = (fechaFin - fechaIni).Days
            If Not (tiempo = 0) Then
                Label34.Text = tiempo.ToString()
            Else
                Label34.Text = "..."
            End If
        Else
            fechaIni = DateTimePicker6.Value
            fechaFin = DateTimePicker5.Value
            tiempo = (fechaFin - fechaIni).Days
            If Not (tiempo = 0) Then
                Label62.Text = tiempo.ToString()
            Else
                Label62.Text = "..."
            End If
        End If
    End Sub

    'Método para asignar valor horas extras
    Function hrsExtras(Total As Double)
        TextBox10.Text = Total.ToString()
    End Function

    'Método para limpiar campos
    Private Sub limpiar(num As Integer)
        If num = 1 Then
            TextBox4.Text = ""
            TextBox5.Text = ""
            TextBox2.Text = ""
            Label31.Text = "..."
            TextBox3.Text = ""
            TextBox8.Text = ""
            TextBox10.Text = ""
            TextBox20.Text = ""
            TextBox21.Text = ""
            TextBox22.Text = ""
            TextBox19.Text = ""
            TextBox12.Text = ""
            TextBox13.Text = ""
            TextBox16.Text = ""
            TextBox7.Text = ""
            TextBox6.Text = ""
            TextBox15.Text = ""
            TextBox17.Text = ""
            TextBox14.Text = ""
            Horasextras.TextBox13.Text = ""
            Horasextras.TextBox16.Text = ""
            Horasextras.TextBox19.Text = ""
            Horasextras.TextBox21.Text = ""
            Horasextras.Label7.Text = ""
            Recargos.TextBox13.Text = ""
            Recargos.TextBox16.Text = ""
            Recargos.TextBox19.Text = ""
            Recargos.Label6.Text = ""
            TextBox9.Text = ""
        ElseIf num = 2 Then
            TextBox23.Text = ""
            TextBox18.Text = ""
            TextBox34.Text = ""
            Label38.Text = "..."
            TextBox35.Text = ""
            TextBox39.Text = ""
            TextBox38.Text = ""
            TextBox27.Text = ""
            TextBox25.Text = ""
            TextBox24.Text = ""
            TextBox26.Text = ""
            TextBox28.Text = ""
            TextBox29.Text = ""
        Else
            ComboBox1.Text = ""
            TextBox42.Text = ""
            TextBox30.Text = ""
            TextBox41.Text = ""
            TextBox11.Text = ""
            TextBox44.Text = ""
            TextBox46.Text = "0"
            TextBox45.Text = "0"
            TextBox47.Text = "0"
            TextBox48.Text = "0"
            TextBox43.Text = ""
            Label73.Text = "..."
        End If
    End Sub

    'Método para asignar valor recargos
    Function totalRecargos(Total As Double)
        TextBox9.Text = Total.ToString()
    End Function

    'Método para asignar valor incapacidades
    Function totalIncapacidades(Total As Double)
        TextBox14.Text = Total.ToString()
    End Function

    'Método para calcular los totales de nómina
    Private Sub calcularTotales()
        Dim consulta As String = "select * from totales_nomina"
        Dim datos As List(Of String) = cnx.execSelectVarios(consulta, False)

        If Not datos.Count() = 0 Then
            Label50.Text = datos.ElementAt(0)
            Label48.Text = datos.ElementAt(1)
            Label46.Text = datos.ElementAt(2)
        End If
    End Sub

    'Método para cambiar a formato yyyy-MM-dd la fecha y poder almacenarla en la BD
    Function convertFecha(fecha As String)
        Dim vec() As String = Split(fecha, "/")
        Dim fec As String = vec(2) & "-" & vec(1) & "-" & vec(0)

        Return fec
    End Function

    'Método para cargar datos en datagridview
    Private Sub cargarDatosGridColilla(num As Integer, documento As String)
        colillaLoad(num, documento)
    End Sub

    'Metodo para cargar datos de colilla 
    Private Sub colillaLoad(flag As Integer, documento As String)
        If flag = 1 Then
            'Datos de colilla antes de generar nómina
            Dim consulta As String = "select h.Nombre, h.Apellido, c.Cargo, c.salario_base
                From Hoja_de_vida h inner join Contrato c on c.Nume_documento = h.Numero_documento where c.Actividad = 'Activo'
                and c.Tipo_Contrato <> 'Prestación de servicios' and h.Numero_documento = " & documento
            Dim datos As List(Of String) = cnx.execSelectVarios(consulta, False)

            If Not datos.Count() = 0 Then
                empColilla.Add(documento)
                DataGridView1.Rows.Add(
                    documento,
                    datos.ElementAt(0).ToString(),
                    datos.ElementAt(1).ToString(),
                    datos.ElementAt(2).ToString(),
                    datos.ElementAt(3).ToString()
                )
            End If
        Else
            'Datos de colilla con nómina generada
            Dim consulta As String = "select co.Nombre, co.Apellido, c.Cargo, co.Neto_pagar from Colilla_pago co 
                inner join Contrato c on co.Numero_documento = c.Nume_documento 
                where co.Numero_documento = " & documento
            Dim datos As List(Of String) = cnx.execSelectVarios(consulta, False)

            If Not datos.Count() = 0 Then
                empColilla.Add(documento)
                DataGridView1.Rows.Add(
                    documento,
                    datos.ElementAt(0).ToString(),
                    datos.ElementAt(1).ToString(),
                    datos.ElementAt(2).ToString(),
                    datos.ElementAt(3).ToString()
                )
            End If
        End If

        For Each col As DataGridViewColumn In DataGridView1.Columns
            col.SortMode = DataGridViewColumnSortMode.NotSortable
        Next
    End Sub

    'Método para cargar datos básicos del empleado
    Private Sub cargarBasicos(documento As String, metodo As Integer)
        'Cunsultar valores básicos del empleado
        Dim consulta As String = "select h.nombre, h.apellido, c.salario_base from Hoja_de_vida h inner join Contrato c 
                        on c.Nume_documento = h.Numero_documento where c.Nume_documento = " & documento
        Dim datos As List(Of String) = cnx.execSelectVarios(consulta, False)

        If metodo = 1 Then 'Colilla
            If Not datos.Count() = 0 Then
                TextBox4.Text = datos.ElementAt(0).ToString()
                TextBox5.Text = datos.ElementAt(1).ToString()
                TextBox2.Text = datos.ElementAt(2).ToString()
                DateTimePicker3.Value = DateTimePicker2.Value.AddDays(15)
            End If
        ElseIf metodo = 2 Then 'Cuenta de cobro
            If Not datos.Count() = 0 Then
                TextBox23.Text = datos.ElementAt(0).ToString()
                TextBox18.Text = datos.ElementAt(1).ToString()
                TextBox34.Text = datos.ElementAt(2).ToString()
                DateTimePicker1.Value = DateTimePicker4.Value.AddDays(30)
            End If

            consulta = "select id, Por_concepto_de from Contrato where Nume_documento = " & documento
            datos = cnx.execSelectVarios(consulta, False)
            If Not datos.Count() = 0 Then
                TextBox39.Text = datos.ElementAt(0).ToString()
                TextBox38.Text = datos.ElementAt(1).ToString()
            End If
        Else
            consulta = "select Inicio_contrato from Contrato where Nume_documento = " & documento
            Dim fecha As Date = cnx.execSelect(consulta)
            DateTimePicker6.Format = DateTimePickerFormat.Short
            DateTimePicker6.Text = fecha
            DateTimePicker6.Format = DateTimePickerFormat.Long
            DateTimePicker6.Enabled = False
            TextBox41.Text = datos.ElementAt(0).ToString()
            TextBox11.Text = datos.ElementAt(1).ToString()
            Dim base As Double = Convert.ToDouble(datos.ElementAt(2))
            TextBox44.Text = ((base / 30) * 15)
            TextBox43.Text = auxTrans.ToString("F2")
        End If
    End Sub

    'Método para cargar datos a los campos de colilla de pago
    Private Sub cargarDatosColilla(documento As String)
        Dim consulta As String = "select top(1)id_colilla_empleado from Colilla_pago 
                                    where Numero_documento = " & documento & " order by id_colilla_empleado desc"
        Dim dat As List(Of String) = cnx.execSelectVarios(consulta, False)

        If dat.Count() = 0 Then 'No tiene ninguna colilla de pago
            consulta = "select Tipo_Contrato from Contrato where Nume_documento = " & documento
            Dim tipo As String = cnx.execSelect(consulta)
            If Not tipo = "0" Then
                If Not tipo = "Prestación de servicios" Then
                    TextBox1.Text = documento
                    cargarBasicos(documento, 1)
                    Dim basico As Double = Convert.ToDouble(TextBox2.Text)
                    TextBox3.Text = ((basico / 30) * Convert.ToInt16(Label23.Text)).ToString("F2")
                    TextBox8.Text = auxTrans.ToString("F2")
                    TextBox20.Text = (basico * (porSalud / 100)).ToString("F2")
                    TextBox19.Text = (basico * (porPension / 100)).ToString("F2")

                    consulta = "select sum(valor_prestamo) from Prestamos where Num_documento = " & documento &
                                    " and Metodo_pago = 'Deducción de nómina'"
                    TextBox17.Text = cnx.execSelect(consulta)
                End If
            Else
                MsgBox("Documento no encontrado")
            End If
        ElseIf dat.Count() = 1 Then 'Tiene una colilla solamente
            consulta = "if ((select DATEPART(month, Fecha_inicial) from Colilla_pago where id_colilla_empleado = " &
                                 dat.ElementAt(0) & " and Numero_documento = " & documento & ") = (select DATEPART(month, GETDATE())))
	                        select 1
                        else
	                        select 0"
            If cnx.execSelect(consulta) = 1 Then 'Si tiene la última colilla del mes
                consulta = "select * from Colilla_pago where id_colilla_empleado = " & dat.ElementAt(0)
                Dim emple As List(Of String) = cnx.execSelectVarios(consulta, False)

                If Not emple.Count() = 0 Then
                    'Pasar fecha a formato corto para almacenarlas y luego a largo otra vez
                    DateTimePicker2.Format = DateTimePickerFormat.Short
                    DateTimePicker3.Format = DateTimePickerFormat.Short

                    TextBox1.Text = emple.ElementAt(2)
                    TextBox4.Text = emple.ElementAt(3)
                    TextBox5.Text = emple.ElementAt(4)
                    TextBox2.Text = emple.ElementAt(5)
                    DateTimePicker2.Text = emple.ElementAt(6)
                    DateTimePicker3.Text = emple.ElementAt(7)
                    TextBox3.Text = emple.ElementAt(8)
                    TextBox8.Text = emple.ElementAt(9)
                    Horasextras.TextBox13.Text = emple.ElementAt(10)
                    Horasextras.TextBox16.Text = emple.ElementAt(11)
                    Horasextras.TextBox19.Text = emple.ElementAt(12)
                    Horasextras.TextBox21.Text = emple.ElementAt(13)
                    TextBox10.Text = emple.ElementAt(14)
                    Recargos.TextBox13.Text = emple.ElementAt(15)
                    Recargos.TextBox16.Text = emple.ElementAt(16)
                    Recargos.TextBox19.Text = emple.ElementAt(17)
                    TextBox9.Text = emple.ElementAt(18)
                    TextBox14.Text = emple.ElementAt(19)
                    TextBox13.Text = emple.ElementAt(20)
                    TextBox22.Text = emple.ElementAt(21)
                    TextBox21.Text = emple.ElementAt(22)
                    TextBox12.Text = emple.ElementAt(23)
                    TextBox20.Text = emple.ElementAt(24)
                    TextBox19.Text = emple.ElementAt(25)
                    TextBox7.Text = emple.ElementAt(26)
                    TextBox6.Text = emple.ElementAt(27)
                    TextBox15.Text = emple.ElementAt(28)
                    TextBox17.Text = emple.ElementAt(29)
                    TextBox16.Text = emple.ElementAt(30)
                    Label31.Text = emple.ElementAt(31)

                    DateTimePicker2.Format = DateTimePickerFormat.Long
                    DateTimePicker3.Format = DateTimePickerFormat.Long
                End If
            Else
                TextBox1.Text = documento
                cargarBasicos(documento, 1)
                Dim basico As Double = Convert.ToDouble(TextBox2.Text)
                TextBox3.Text = ((basico / 30) * Convert.ToInt16(Label23.Text)).ToString("F2")
                TextBox8.Text = auxTrans.ToString("F2")
                TextBox20.Text = (basico * (porSalud / 100)).ToString("F2")
                TextBox19.Text = (basico * (porPension / 100)).ToString("F2")

                consulta = "select sum(valor_prestamo) from Prestamos where Num_documento = " & documento
                TextBox17.Text = cnx.execSelect(consulta)
            End If
        End If
    End Sub

    'Método para conocer el último id de colilla de un empleado
    Function ultimaColilla(documento As String)
        Dim consulta As String = "select top(1)id_colilla_empleado from Colilla_pago 
                                        where Numero_documento = " & documento & " order by id_colilla_empleado desc"
        Dim id As String = cnx.execSelect(consulta)
        If Not id = "" Then
            Dim vl As Integer = Convert.ToInt32(id) + 1
            If vl < 10 Then
                Return "000" & vl
            ElseIf vl < 100 Then
                Return "00" & vl
            ElseIf vl < 1000 Then
                Return "0" & vl
            Else
                Return vl
            End If
        Else
            Return "0001"
        End If
    End Function

    'Método para consultar si ya se le realizó nómina a un empleado
    Function consultaNominaColilla()
        DataGridView1.Rows.Clear()
        buscarValores()
        Dim idcolilla1 As String
        Dim idcolilla2 As String
        Dim doc As String

        Dim consulta As String = "select h.Numero_documento from Hoja_de_vida h inner join contrato c on 
                                        h.Numero_documento = c.Nume_documento where c.Actividad = 'Activo'
                                        and c.Tipo_Contrato <> 'Prestación de servicios'"
        Dim datos As List(Of String) = cnx.execSelectVarios(consulta, False)

        If Not datos.Count() = 0 Then
            For i As Integer = 0 To datos.Count() - 1 Step 1
                doc = datos.ElementAt(i)
                consulta = "select top(2)id_colilla_empleado from Colilla_pago 
                                    where Numero_documento = " & doc & " order by id_colilla_empleado desc"
                Dim dat As List(Of String) = cnx.execSelectVarios(consulta, False)

                If dat.Count() = 0 Then 'No tiene ninguna colilla de pago
                    cargarDatosGridColilla(1, datos.ElementAt(i))

                ElseIf dat.Count() = 1 Then 'Tiene una colilla solamente
                    idcolilla1 = dat.ElementAt(0)
                    consulta = "if ((select DATEPART(month, Fecha_inicial) from Colilla_pago where 
                                    id_colilla_empleado = " & idcolilla1 & " and Numero_documento = " & doc & ") = (select DATEPART(month, GETDATE())))
                                    select 1
                                else
                                    select 0"
                    If cnx.execSelect(consulta) = 1 Then
                        If Today.Day < 25 Then
                            cargarDatosGridColilla(2, doc)
                        Else
                            cargarDatosGridColilla(1, doc)
                        End If
                    Else
                        cargarDatosGridColilla(1, doc)
                    End If
                Else 'Tiene más de una colilla
                    idcolilla1 = dat.ElementAt(0)
                    idcolilla2 = dat.ElementAt(1)

                    consulta = "if ((select DATEPART(month, Fecha_inicial) from Colilla_pago 
                                        where id_colilla_empleado = " & idcolilla1 & " and Numero_documento = " & doc & ")  = (select DATEPART(month, GETDATE())))
                                    select 1
                                else
                                    select 0"

                    Dim cons As String = "if ((select DATEPART(month, Fecha_inicial) from Colilla_pago where id_colilla_empleado = " &
                                                 idcolilla2 & " and Numero_documento = " & doc & ")  = (select DATEPART(month, GETDATE())))
	                                select 1
                                else
	                                select 0"
                    Dim res1 As String = cnx.execSelect(consulta)
                    Dim res2 As String = cnx.execSelect(cons)

                    If res1 = 1 And res2 = 1 Then 'Ya tiene las dos colillas del mes
                        cargarDatosGridColilla(2, doc)
                    End If
                End If
            Next
        End If
    End Function

    '----------------------------------------

    'Método para cargar datos en datagridview
    Private Sub cargarDatosGridCuenta(num As Integer, documento As String)
        cuentaLoad(num, documento)
    End Sub

    'Metodo para cargar datos de cuenta cobro al grid
    Private Sub cuentaLoad(flag As Integer, documento As String)
        If flag = 1 Then
            'Datos de cuenta de cobro antes de generar nómina
            Dim consulta As String = "select h.Numero_documento, h.Nombre, h.Apellido, c.Cargo, c.salario_base
            from Hoja_de_vida h inner join Contrato c on c.Nume_documento = h.Numero_documento where c.Actividad = 'Activo'
            and c.Tipo_Contrato = 'Prestación de servicios'"
            Dim data As List(Of String) = cnx.execSelectVarios(consulta, False)

            If Not data.Count() = 0 Then
                For i As Integer = 0 To data.Count() - 1 Step 5
                    empCuenta.Add(data.ElementAt(i).ToString())
                    DataGridView2.Rows.Add(
                    data.ElementAt(i).ToString(),
                    data.ElementAt(i + 1).ToString(),
                    data.ElementAt(i + 2).ToString(),
                    data.ElementAt(i + 3).ToString(),
                    data.ElementAt(i + 4).ToString()
                )
                Next
            End If
        Else
            'Datos de cuenta de cobro con nómina generada
            Dim consulta As String = "select cc.nombre, cc.apellido, c.cargo, cc.neto_pagar from Cuenta_cobro cc 
                inner join Contrato c on cc.Numero_documento = c.Nume_documento 
                where cc.Numero_documento = " & documento
            Dim datos As List(Of String) = cnx.execSelectVarios(consulta, False)

            If Not datos.Count() = 0 Then
                empCuenta.Add(documento)
                DataGridView2.Rows.Add(
                    documento,
                    datos.ElementAt(0).ToString(),
                    datos.ElementAt(1).ToString(),
                    datos.ElementAt(2).ToString(),
                    datos.ElementAt(3).ToString()
                )
            End If
        End If
        For Each col As DataGridViewColumn In DataGridView2.Columns
            col.SortMode = DataGridViewColumnSortMode.NotSortable
        Next
    End Sub

    'Método para cargar datos a los campos de cuenta cobro
    Private Sub cargarDatosCuenta(documento As String)
        Dim consulta As String = "select top(1)id from Cuenta_cobro where Numero_documento = " & documento & " order by id desc"
        Dim dat As String = cnx.execSelect(consulta)

        If dat = 0 Then 'No tiene ninguna cuenta de cobro
            consulta = "select Tipo_Contrato from Contrato where Nume_documento = " & documento
            Dim tipo As String = cnx.execSelect(consulta)
            If Not tipo = "0" Then
                If tipo = "Prestación de servicios" Then
                    TextBox40.Text = documento
                    cargarBasicos(documento, 2)
                    Dim basico As Double = Convert.ToDouble(TextBox34.Text)
                    TextBox35.Text = ((basico / 30) * Convert.ToInt16(Label34.Text)).ToString("F2")

                    consulta = "select sum(valor_prestamo) from Prestamos where Num_documento = " & documento &
                                    " and Metodo_pago = 'Deducción de nómina'"
                    TextBox29.Text = cnx.execSelect(consulta)
                End If
            Else
                MsgBox("Documento no encontrado")
            End If
        Else 'Tiene una cuenta de cobro
            consulta = "if ((select DATEPART(month, Fecha_inicial) from Cuenta_cobro where id = " & dat &
                                " and Numero_documento = " & documento & ") = (select DATEPART(month, GETDATE())))
	                        select 1
                        else
	                        select 0"
            If cnx.execSelect(consulta) = 1 Then 'Si tiene la última cuenta de cobro del mes
                consulta = "select * from Cuenta_cobro where id = " & dat
                Dim emple As List(Of String) = cnx.execSelectVarios(consulta, False)

                If Not emple.Count() = 0 Then
                    'Pasar fecha a formato corto para almacenarlas y luego a largo otra vez
                    DateTimePicker4.Format = DateTimePickerFormat.Short
                    DateTimePicker1.Format = DateTimePickerFormat.Short

                    TextBox40.Text = emple.ElementAt(1)
                    TextBox23.Text = emple.ElementAt(2)
                    TextBox18.Text = emple.ElementAt(3)
                    TextBox34.Text = emple.ElementAt(4)
                    DateTimePicker4.Text = emple.ElementAt(5)
                    DateTimePicker1.Text = emple.ElementAt(6)
                    TextBox35.Text = emple.ElementAt(7)
                    TextBox39.Text = emple.ElementAt(8)
                    TextBox38.Text = emple.ElementAt(9)
                    TextBox27.Text = emple.ElementAt(10)
                    TextBox25.Text = emple.ElementAt(11)
                    TextBox28.Text = emple.ElementAt(12)
                    TextBox24.Text = emple.ElementAt(13)
                    TextBox26.Text = emple.ElementAt(14)
                    TextBox29.Text = emple.ElementAt(15)
                    Label38.Text = emple.ElementAt(16)

                    DateTimePicker4.Format = DateTimePickerFormat.Long
                    DateTimePicker1.Format = DateTimePickerFormat.Long
                End If
            End If
        End If
    End Sub

    'Método para consultar si ya se le realizó nómina a un empleado
    Function consultaNominaCuenta()
        DataGridView2.Rows.Clear()
        Dim idcuenta As String
        Dim doc As String

        Dim consulta As String = "select Numero_documento from Hoja_de_vida h inner join contrato c on 
                                    h.Numero_documento = c.Nume_documento where c.Actividad = 'Activo'
                                    and c.Tipo_Contrato = 'Prestación de servicios'"
        Dim datos As List(Of String) = cnx.execSelectVarios(consulta, False)

        If Not datos.Count() = 0 Then
            For i As Integer = 0 To datos.Count() - 1 Step 1
                doc = datos.ElementAt(i)
                consulta = "select top(1)id from Cuenta_cobro where Numero_documento = " & doc & " order by id desc"
                Dim dat As List(Of String) = cnx.execSelectVarios(consulta, False)

                If dat.Count() = 0 Then 'No tiene ninguna cuenta de cobro
                    cargarDatosGridCuenta(1, datos.ElementAt(i))
                Else
                    'Tiene una cuenta de cobro
                    idcuenta = dat.ElementAt(0)
                    consulta = "if ((select DATEPART(month, Fecha_inicial) from Cuenta_cobro where id = " & idcuenta &
                                " and Numero_documento = " & doc & ")  = (select DATEPART(month, GETDATE())))
	                                select 1
                                else
	                                select 0"
                    If cnx.execSelect(consulta) = 1 Then
                        If Today.Day < 25 Then
                            cargarDatosGridCuenta(2, doc)
                        Else
                            cargarDatosGridCuenta(1, doc)
                        End If
                    Else
                        cargarDatosGridCuenta(1, doc)
                    End If
                End If
            Next
        End If
    End Function

    '----------------------------------------

    'Método para calcular promedio neto de un empleado
    Function calcularPromNeto(documento As String, tipo As Integer)
        If tipo = 1 Then 'Sin deducciones
            Return Convert.ToDecimal((cnx.execSelect("select format(avg(Neto_pagar), 'F0') from Colilla_pago where Numero_documento = " & documento)))
        Else
            Return (Convert.ToDecimal(cnx.execSelect("select format(((avg(Neto_pagar)) - Aux_transporte - Salud - Pension), 'F0') 
                from Colilla_pago where Numero_documento = " & documento & " group by Aux_transporte, Salud, Pension")))
        End If
    End Function

    'Método para cargar datos en datagridview
    Private Sub cargarDatosGridPrestaciones(documento As String)
        prestacionesLoad(documento)
    End Sub

    'Método para consultar si ya se liquidó
    Function consultaLiquidaciones()
        DataGridView3.Rows.Clear()
        Dim consulta As String = "select Numero_documento from Hoja_de_vida h inner join contrato c on 
                                    h.Numero_documento = c.Nume_documento where c.Actividad = 'Activo'
                                    and c.Tipo_Contrato <> 'Prestación de servicios'"
        Dim datos As List(Of String) = cnx.execSelectVarios(consulta, False)

        If Not datos.Count() = 0 Then
            For i As Integer = 0 To datos.Count() - 1 Step 1
                cargarDatosGridPrestaciones(datos.ElementAt(i))
            Next
        End If
    End Function

    'Metodo para cargar datos de prestaciones 
    Private Sub prestacionesLoad(documento As String)
        Dim consulta As String = "
            declare @total decimal(10, 2)
                set @total = (select sum(Total_liquidado) from liquidaciones where Numero_documento = " & documento & ")
                if (@total IS NOT NULL)
	                begin
		                select h.Nombre, h.Apellido, c.Cargo, @total total
                        From Hoja_de_vida h inner join Contrato c on c.Nume_documento = h.Numero_documento 
		                where c.Actividad = 'Activo'
                        and c.Tipo_Contrato <> 'Prestación de servicios' and Numero_documento = " & documento & "
	                end
                else
	                begin
		                select h.Nombre, h.Apellido, c.Cargo, 0 total
                        From Hoja_de_vida h inner join Contrato c on c.Nume_documento = h.Numero_documento 
		                where c.Actividad = 'Activo'
                        and c.Tipo_Contrato <> 'Prestación de servicios' and Numero_documento = " & documento & "
	                end"

        Dim datos As List(Of String) = cnx.execSelectVarios(consulta, False)

        If Not datos.Count() = 0 Then
            empTodos.Add(documento)
            DataGridView3.Rows.Add(
                documento,
                datos.ElementAt(0).ToString(),
                datos.ElementAt(1).ToString(),
                datos.ElementAt(2).ToString(),
                datos.ElementAt(3).ToString()
            )
        End If

        For Each col As DataGridViewColumn In DataGridView1.Columns
            col.SortMode = DataGridViewColumnSortMode.NotSortable
        Next
    End Sub

    'Método para saber total de los préstamos prestaciones
    Function prestamosPS(documento As String)
        Dim consulta As String =
            "declare @pres decimal(10, 2) 
            set @pres = (select sum(valor_prestamo) from Prestamos where Metodo_pago = 'Prestaciones sociales' or 
                Metodo_pago = 'Liquidación' and Num_documento = " & documento & ")
            if @pres is null
	            begin 
		            select 0
	            end
            else
	            begin
		            select @pres
	            end"
        Dim prest As Single = Convert.ToSingle(cnx.execSelect(consulta))
        Return prest
    End Function

    'Método para cargar datos a los campos de prestaciones
    Private Sub cargarDatosPrestaciones(documento As String)
        Dim vaca As Double
        Dim cesan As Double
        Dim int_cesan As Double
        Dim prima As Double

        TextBox42.Text = documento
        cargarBasicos(documento, 3)

        Dim consulta As String = "select fecha_inicial, fecha_final from liquidaciones where numero_documento = " & documento
        Dim datos As List(Of String) = cnx.execSelectVarios(consulta, False)
        If Not datos.Count() = 0 Then
            DateTimePicker6.Format = DateTimePickerFormat.Short
            DateTimePicker5.Format = DateTimePickerFormat.Short

            DateTimePicker6.Value = datos.ElementAt(0)
            DateTimePicker5.Value = datos.ElementAt(1)

            DateTimePicker6.Format = DateTimePickerFormat.Long
            DateTimePicker5.Format = DateTimePickerFormat.Long
        End If

        consulta = "select total_liquidado from liquidaciones where Descripcion_prestacion = 'Vacaciones' 
                           and (datepart(year, fecha_inicial) = datepart(year, getdate())) and numero_documento = " & documento
        Dim dat As String = cnx.execSelect(consulta)
        vaca = Convert.ToDouble(dat)
        TextBox47.Text = vaca.ToString("F2")

        consulta = "select total_liquidado from liquidaciones where Descripcion_prestacion = 'Cesantías' 
                           and (datepart(year, fecha_inicial) = datepart(year, getdate())) and numero_documento = " & documento
        dat = cnx.execSelect(consulta)
        cesan = Convert.ToDouble(dat)
        TextBox45.Text = cesan.ToString("F2")

        consulta = "select total_liquidado from liquidaciones where Descripcion_prestacion = 'Intereses a las cesantías' 
                           and (datepart(year, fecha_inicial) = datepart(year, getdate())) and numero_documento = " & documento
        dat = cnx.execSelect(consulta)
        int_cesan = Convert.ToDouble(dat)
        TextBox48.Text = int_cesan.ToString("F2")

        consulta = "select total_liquidado from liquidaciones where Descripcion_prestacion = 'Prima' 
                           and (datepart(year, fecha_inicial) = datepart(year, getdate())) and numero_documento = " & documento
        dat = cnx.execSelect(consulta)
        prima = Convert.ToDouble(dat)
        TextBox46.Text = prima.ToString("F2")

        consulta = "
            if not (select nombre from liquidaciones where Descripcion_prestacion = 'Liquidar todos' 
                    and (datepart(year, fecha_inicial) = datepart(year, getdate())) and numero_documento = " & documento & ") = '' 
	            begin
		            select 1
	            end
            else
	            begin
		            select 0
	            end"
        dat = cnx.execSelect(consulta)
        If Convert.ToInt16(dat) = 1 Then
            calcularPrestaciones(4)
        End If

        TextBox30.Text = prestamosPS(documento)
    End Sub

    'Método para calcular valores de prestaciones
    Private Sub calcularPrestaciones(val As Integer)
        Dim pres As Single = prestamosPS(TextBox42.Text)
        Dim promNeto As Decimal = calcularPromNeto(TextBox42.Text, 1)
        Dim promNeto2 As Decimal = calcularPromNeto(TextBox42.Text, 2)
        Dim cesantias As Double = ((promNeto * Convert.ToInt16(Label62.Text)) / 360)

        Select Case val
            Case 0
                TextBox46.Text = ((promNeto * Convert.ToInt16(Label62.Text)) / 360).ToString("F2")
            Case 1
                TextBox45.Text = ((promNeto * Convert.ToInt16(Label62.Text)) / 360).ToString("F2")
            Case 2
                TextBox48.Text = ((cesantias * Convert.ToInt16(Label62.Text) * 0.12) / 360).ToString("F2")
            Case 3
                TextBox47.Text = ((promNeto2 * Convert.ToInt16(Label62.Text)) / 720).ToString("F2")
            Case 4
                TextBox46.Text = ((promNeto * Convert.ToInt16(Label62.Text)) / 360).ToString("F2")
                TextBox45.Text = ((promNeto * Convert.ToInt16(Label62.Text)) / 360).ToString("F2")
                TextBox48.Text = ((cesantias * Convert.ToInt16(Label62.Text) * 0.12) / 360).ToString("F2")
                TextBox47.Text = ((promNeto2 * Convert.ToInt16(Label62.Text)) / 720).ToString("F2")
        End Select
        Label73.Text = (Convert.ToDouble(TextBox46.Text) + Convert.ToDouble(TextBox45.Text) +
            Convert.ToDouble(TextBox48.Text) + Convert.ToDouble(TextBox47.Text)) - prestamosPS(TextBox42.Text)
    End Sub

    'Método para cargar todos los links de los documentos
    Private Sub cargarLinksDocs(fecha As Date)
        Dim consulta As String
        Dim consulta2 As String
        Dim archi As List(Of Byte())
        Dim arch As List(Of String)
        Dim texto As String
        Panel1.Controls.Clear()
        documentos.Clear()

        If fecha = "1999-09-19" Then 'Traer todos
            consulta = "select Documento from documentos_nomina order by id"
            consulta2 = "select fecha, tipo_contrato from documentos_nomina order by id"
        Else
            consulta = "select Documento from documentos_nomina where fecha = '" & convertFecha(fecha) & "' order by id"
            consulta2 = "select fecha, tipo_contrato from documentos_nomina where fecha = '" & convertFecha(fecha) & "' order by id"
        End If

        archi = cnx.execSelectVarios(consulta, True)
        If Not archi.Count() = 0 Then
            For i As Integer = 0 To archi.Count() - 1 Step 1
                If Not UBound(archi(i)) = 0 Then
                    documentos.Add(archi(i))
                End If
            Next

            arch = cnx.execSelectVarios(consulta2, False)
            If Not arch.Count() = 0 Then
                For i As Integer = 0 To arch.Count() - 1 Step 2
                    texto = arch.ElementAt(i + 1) & " " & arch.ElementAt(i)
                    crearControl(texto, j.ToString())
                    j += 1
                Next
            End If
        Else
            MsgBox("No hay documentos generados en la fecha seleccionada")
        End If
    End Sub

    'Método para agregar links al form
    Private Sub crearControl(text As String, name As String)
        Dim link As New LinkLabel
        With link
            .Visible = True
            .LinkColor = Color.Green
            .Name = name
            .Text = text
            .Width = 200
            AddHandler link.Click, AddressOf eventoLink
        End With
        link.Location = New Point(0, (link.Top + link.Height + 10) * i)
        Panel1.Controls.Add(link)
        i += 1
    End Sub

    'Método para los eventos de click de los link de anexos
    Private Sub eventoLink(ByVal sender As System.Object, ByVal e As EventArgs)
        Dim link = TryCast(sender, LinkLabel)

        If link IsNot Nothing Then
            recuperarPDF(documentos.ElementAt(Convert.ToInt16(link.Name)), 1)
            Button18.Enabled = True
        End If
    End Sub

    'Método para recuperar pdf
    Private Sub recuperarPDF(byt As Byte(), iden As Integer)
        Try
            Dim directorioArchivo As String
            directorioArchivo = System.AppDomain.CurrentDomain.BaseDirectory() & "temp.pdf"

            If (BytesAArchivo(byt, directorioArchivo)) Then
                If iden = 1 Then
                    AxAcroPDF5.src = directorioArchivo
                    My.Computer.FileSystem.DeleteFile(directorioArchivo)
                Else
                    Imprimir.AxAcroPDF1.src = directorioArchivo
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    'Método para convertir bytes a pdf
    Function BytesAArchivo(ByVal bytes() As Byte, ByVal Path As String)
        Try
            Dim K As Long = UBound(bytes)
            If Not K = 0 Then
                Dim fs As New FileStream(Path, FileMode.OpenOrCreate, FileAccess.Write)
                fs.Write(bytes, 0, K)
                fs.Close()
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message, ex)
        End Try
    End Function

#End Region

#Region "Guardar"

    'Guardar colilla de pago
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim bool As Boolean = False
        Dim totalDeven As Double
        Dim totalDeduc As Double
        Dim netoPagar As Double
        Dim lic As Single
        Dim otrosDev As Single
        Dim desc1 As String
        Dim otrasDed As Single
        Dim desc2 As String
        Dim rete As Single
        Dim hrsext As Single
        Dim rec As Single
        Dim inca As Single
        Dim pres As Single

        If TextBox13.Text = "" Then
            lic = 0
        Else
            lic = Convert.ToSingle(TextBox13.Text)
        End If

        If TextBox22.Text = "" Then
            otrosDev = 0
        Else
            otrosDev = Convert.ToSingle(TextBox22.Text)
        End If

        If TextBox7.Text = "" Then
            otrasDed = 0
        Else
            otrasDed = Convert.ToSingle(TextBox7.Text)
        End If

        If TextBox15.Text = "" Then
            rete = 0
        Else
            rete = Convert.ToSingle(TextBox15.Text)
        End If

        If TextBox10.Text = "" Then
            hrsext = 0
        Else
            hrsext = Convert.ToSingle(TextBox10.Text)
        End If

        If TextBox14.Text = "" Then
            inca = 0
        Else
            inca = Convert.ToSingle(TextBox14.Text)
        End If

        If TextBox9.Text = "" Then
            rec = 0
        Else
            rec = Convert.ToSingle(TextBox9.Text)
        End If

        If TextBox17.Text = "" Then
            pres = 0
        Else
            pres = Convert.ToSingle(TextBox17.Text)
        End If

        If Not otrosDev = 0 Then
            If TextBox21.Text = "" Then
                MsgBox("Debe ingresar la descripción de otros devengados")
            Else
                If Not otrasDed = 0 Then
                    If TextBox21.Text = "" Then
                        MsgBox("Debe ingresar la descripción de otras deducciones")
                    Else
                        bool = True
                    End If
                Else
                    bool = True
                End If
            End If
        Else
            If Not otrasDed = 0 Then
                If TextBox21.Text = "" Then
                    MsgBox("Debe ingresar la descripción de otros devengados")
                Else
                    bool = True
                End If
            Else
                bool = True
            End If
        End If

        If bool Then 'Se puede guardar
            totalDeven = Convert.ToSingle(TextBox3.Text) + auxTrans + hrsext + rec + inca + lic + otrosDev
            totalDeven = Math.Round(totalDeven, 2)

            totalDeduc = Convert.ToSingle(TextBox20.Text) + Convert.ToSingle(TextBox19.Text) + otrasDed + rete + pres
            totalDeduc = Math.Round(totalDeduc, 2)

            netoPagar = totalDeven - totalDeduc

            TextBox12.Text = totalDeven.ToString("F2")
            TextBox16.Text = totalDeduc.ToString("F2")
            Label31.Text = netoPagar.ToString("F2")

            Try
                Dim idColilla As String = ultimaColilla(TextBox1.Text)

                'Pasar fecha a formato corto para almacenarlas y luego a largo otra vez
                DateTimePicker2.Format = DateTimePickerFormat.Short
                DateTimePicker3.Format = DateTimePickerFormat.Short

                Dim cmm As New SqlCommand("sp_mant_colilla_pago", cnx.conexion)
                cmm.CommandType = CommandType.StoredProcedure

                'Todos los datos son obligatorios en la BD
                cmm.Parameters.AddWithValue("@id_colilla_empleado", idColilla)
                cmm.Parameters.AddWithValue("@Numero_documento", TextBox1.Text)
                cmm.Parameters.AddWithValue("@Nombre", TextBox4.Text)
                cmm.Parameters.AddWithValue("@Apellido", TextBox5.Text)
                cmm.Parameters.AddWithValue("@Salario_base", Convert.ToDecimal(TextBox2.Text))
                cmm.Parameters.AddWithValue("@Fecha_inicial", convertFecha(DateTimePicker2.Text))
                cmm.Parameters.AddWithValue("@Fecha_final", convertFecha(DateTimePicker3.Text))
                cmm.Parameters.AddWithValue("@Basico", Convert.ToDecimal(TextBox3.Text))
                cmm.Parameters.AddWithValue("@Aux_transporte", auxTrans)
                cmm.Parameters.AddWithValue("@Hrs_diurnas", vrHrsDiurnas)
                cmm.Parameters.AddWithValue("@Hrs_nocturnas", vrHrsNocturnas)
                cmm.Parameters.AddWithValue("@Hrs_dominicales", vrHrsDominicales)
                cmm.Parameters.AddWithValue("@Hrs_festivas", vrHrsFestivas)
                cmm.Parameters.AddWithValue("@Total_Hrs", hrsext)
                cmm.Parameters.AddWithValue("@Rec_nocturno", Recargos.vrRecNocturno)
                cmm.Parameters.AddWithValue("@Rec_dominical", Recargos.vrRecDominical)
                cmm.Parameters.AddWithValue("@Rec_festivo", Recargos.vrRecFestivo)
                cmm.Parameters.AddWithValue("@Total_Rec", rec)
                cmm.Parameters.AddWithValue("@Incapacidades", inca)
                cmm.Parameters.AddWithValue("@Licencias", lic)
                cmm.Parameters.AddWithValue("@Otros_devengados", otrosDev)
                cmm.Parameters.AddWithValue("@Descripcion_otros_dev", TextBox21.Text)
                cmm.Parameters.AddWithValue("@Total_devengado", totalDeven)
                cmm.Parameters.AddWithValue("@Salud", Convert.ToDecimal(TextBox20.Text))
                cmm.Parameters.AddWithValue("@Pension", Convert.ToDecimal(TextBox19.Text))
                cmm.Parameters.AddWithValue("@Otras_deducciones", otrasDed)
                cmm.Parameters.AddWithValue("@Descripcion_otras_ded", TextBox7.Text)
                cmm.Parameters.AddWithValue("@Retencion", rete)
                cmm.Parameters.AddWithValue("@Prestamos", pres)
                cmm.Parameters.AddWithValue("@Total_deducciones", totalDeduc)
                cmm.Parameters.AddWithValue("@Neto_pagar", netoPagar)

                DateTimePicker2.Format = DateTimePickerFormat.Long
                DateTimePicker3.Format = DateTimePickerFormat.Long

                If cnx.ejecutarSP(cmm) Then
                    MsgBox("Datos agregados correctamente")
                    limpiar(1)
                    consultaNominaColilla()
                Else
                    MsgBox("Hubo un error agregando los datos")
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    'Guardar cuenta de cobro
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim bool As Boolean = False
        Dim netoPagar As Double
        Dim otrosDev As Single
        Dim desc1 As String
        Dim otrasDed As Single
        Dim desc2 As String
        Dim rete As Single
        Dim pres As Single

        If TextBox27.Text = "" Then
            otrosDev = 0
        Else
            otrosDev = Convert.ToSingle(TextBox27.Text)
        End If

        If TextBox28.Text = "" Then
            otrasDed = 0
        Else
            otrasDed = Convert.ToSingle(TextBox28.Text)
        End If

        If TextBox26.Text = "" Then
            rete = 0
        Else
            rete = Convert.ToSingle(TextBox26.Text)
        End If

        If TextBox29.Text = "" Then
            pres = 0
        Else
            pres = Convert.ToSingle(TextBox29.Text)
        End If

        If Not otrosDev = 0 Then
            If TextBox25.Text = "" Then
                MsgBox("Debe ingresar la descripción de otros devengados")
            Else
                If Not otrasDed = 0 Then
                    If TextBox24.Text = "" Then
                        MsgBox("Debe ingresar la descripción de otras deducciones")
                    Else
                        bool = True
                    End If
                Else
                    bool = True
                End If
            End If
        Else
            If Not otrasDed = 0 Then
                If TextBox25.Text = "" Then
                    MsgBox("Debe ingresar la descripción de otras deducciones")
                Else
                    bool = True
                End If
            Else
                bool = True
            End If
        End If

        If bool Then 'Se puede guardar
            netoPagar = Convert.ToSingle(TextBox35.Text) + otrosDev - otrasDed - rete - pres

            Label38.Text = netoPagar.ToString("F2")

            Try
                'Pasar fecha a formato corto para almacenarlas y luego a largo otra vez
                DateTimePicker1.Format = DateTimePickerFormat.Short
                DateTimePicker4.Format = DateTimePickerFormat.Short

                Dim cmm As New SqlCommand("sp_mant_cuenta_cobro", cnx.conexion)
                cmm.CommandType = CommandType.StoredProcedure

                'Todos los datos son obligatorios en la BD
                cmm.Parameters.AddWithValue("@Numero_documento", TextBox40.Text)
                cmm.Parameters.AddWithValue("@Nombre", TextBox23.Text)
                cmm.Parameters.AddWithValue("@Apellido", TextBox18.Text)
                cmm.Parameters.AddWithValue("@Salario_base", Convert.ToDecimal(TextBox34.Text))
                cmm.Parameters.AddWithValue("@Fecha_inicial", convertFecha(DateTimePicker4.Text))
                cmm.Parameters.AddWithValue("@Fecha_final", convertFecha(DateTimePicker1.Text))
                cmm.Parameters.AddWithValue("@Basico", Convert.ToDecimal(TextBox35.Text))
                cmm.Parameters.AddWithValue("@Contrato", Convert.ToInt16(TextBox39.Text))
                cmm.Parameters.AddWithValue("@Concepto_de", TextBox38.Text)
                cmm.Parameters.AddWithValue("@Otros_devengados", otrosDev)
                cmm.Parameters.AddWithValue("@Descripcion_otros_dev", TextBox25.Text)
                cmm.Parameters.AddWithValue("@Otras_deducciones", otrasDed)
                cmm.Parameters.AddWithValue("@Descripcion_otras_ded", TextBox24.Text)
                cmm.Parameters.AddWithValue("@Retencion", rete)
                cmm.Parameters.AddWithValue("@Prestamos", pres)
                cmm.Parameters.AddWithValue("@Neto_pagar", netoPagar)

                DateTimePicker1.Format = DateTimePickerFormat.Long
                DateTimePicker4.Format = DateTimePickerFormat.Long

                If cnx.ejecutarSP(cmm) Then
                    MsgBox("Datos agregados correctamente")
                    limpiar(2)
                    consultaNominaCuenta()
                Else
                    MsgBox("Hubo un error agregando los datos")
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    'Guardar prestaciones sociales
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Try
            DateTimePicker6.Format = DateTimePickerFormat.Short
            DateTimePicker5.Format = DateTimePickerFormat.Short

            Dim cmm As New SqlCommand("sp_mant_liquidaciones", cnx.conexion)
            cmm.CommandType = CommandType.StoredProcedure

            cmm.Parameters.AddWithValue("@Numero_documento", TextBox42.Text)
            cmm.Parameters.AddWithValue("@Nombre", TextBox41.Text)
            cmm.Parameters.AddWithValue("@Apellido", TextBox11.Text)
            cmm.Parameters.AddWithValue("@Fecha_inicial", convertFecha(DateTimePicker6.Text))
            cmm.Parameters.AddWithValue("@Fecha_final", convertFecha(DateTimePicker5.Text))
            cmm.Parameters.AddWithValue("@Basico", Convert.ToDecimal(TextBox44.Text))
            cmm.Parameters.AddWithValue("@Aux_transporte", Convert.ToDecimal(TextBox43.Text))
            cmm.Parameters.AddWithValue("@Descripcion_prestacion", ComboBox1.SelectedItem.ToString())
            cmm.Parameters.AddWithValue("@Prima", Convert.ToDecimal(TextBox46.Text))
            cmm.Parameters.AddWithValue("@Cesantias", Convert.ToDecimal(TextBox45.Text))
            cmm.Parameters.AddWithValue("@Int_cesantias", Convert.ToDecimal(TextBox48.Text))
            cmm.Parameters.AddWithValue("@Vacaciones", Convert.ToDecimal(TextBox47.Text))
            cmm.Parameters.AddWithValue("@Total_liquidado", Convert.ToDecimal(Label73.Text))

            If cnx.ejecutarSP(cmm) Then
                MsgBox("Datos agregados correctamente")
                limpiar(3)
                consultaLiquidaciones()
            Else
                MsgBox("Hubo un error agregando los datos")
            End If

            DateTimePicker6.Format = DateTimePickerFormat.Long
            DateTimePicker5.Format = DateTimePickerFormat.Long
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#End Region

#Region "Imprimir"

    'Imprimir colilla de pago
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        imprimirDoc("select top(1)documento from documentos_nomina where tipo_contrato = 'Colilla de pago' order by id desc")
    End Sub

    Private Sub imprimirDoc(consulta As String)
        Dim archi As List(Of Byte()) = cnx.execSelectVarios(consulta, True)
        If Not archi.Count() = 0 Then
            If Not UBound(archi.ElementAt(0)) = 0 Then
                Me.Hide()
                recuperarPDF(archi.ElementAt(0), 2)
                Imprimir.Show()
            End If
            Dim directorioArchivo As String = System.AppDomain.CurrentDomain.BaseDirectory() & "temp.pdf"
            My.Computer.FileSystem.DeleteFile(directorioArchivo)
        Else
            MsgBox("No se han generado documentos")
        End If
    End Sub

    'Imprimir cuenta de cobro
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        imprimirDoc("select top(1)documento from documentos_nomina where tipo_contrato = 'Cuenta de cobro' order by id desc")
    End Sub

    'Imprimir prestaciones sociales
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        imprimirDoc("select top(1)documento from documentos_nomina where tipo_contrato = 'Prestaciones sociales' order by id desc")
    End Sub

    'Imprimir consolidado
    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        If Not AxAcroPDF5.src = "none" Then
            AxAcroPDF5.Print()
        End If
    End Sub

#End Region

#Region "Exportar"

    'Exportar colilla de pago
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Hide()
        Exportar.plantilla = 1
        Exportar.Show()
    End Sub

    'Exportar cuenta de cobro
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Me.Hide()
        Exportar.plantilla = 2
        Exportar.Show()
    End Sub

    'Exportar prestaciones sociales
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Me.Hide()
        Exportar.plantilla = 3
        Exportar.Show()
    End Sub

#End Region

#Region "Cancelar"

    'Cancelar colilla de pago
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        salir()
    End Sub

    'Cancelar prestaciones sociales
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        salir()
    End Sub

    'Cancelar cuenta de cobro
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        salir()
    End Sub

    'Cancelar consolidado
    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        salir()
    End Sub

#End Region

End Class