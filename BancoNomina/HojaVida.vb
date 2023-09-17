Imports System.Data.SqlClient
Imports System.IO
Imports AxAcroPDFLib

Public Class HojaVida

#Region "Variables"
    Dim cnx As New Conexion()
    Dim i As Integer = 0 'Usada oara organizar los links en panel1
    Dim y As Integer = 0 'Usada oara organizar los links en panel2
    Dim anexos As List(Of Byte()) = New List(Of Byte()) 'ArrayList para los archivos de anexos
    Dim contrato As List(Of Byte()) = New List(Of Byte()) 'ArrayList para los archivos de contratos
    Dim otrosi As List(Of Byte()) = New List(Of Byte()) 'ArrayList para los archivos de otrosi de contrato
    Dim comunicados As List(Of Byte()) = New List(Of Byte()) 'ArrayList para los archivos de comunicados
    Dim certificados As List(Of Byte()) = New List(Of Byte()) 'ArrayList para los archivos de certificados laborales
#End Region

#Region "Eventos"

    'Se ejecuta al cargar la página por primera vez
    Private Sub HojaVida_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DateTimePicker4.MaxDate = Today
        DateTimePicker1.MaxDate = Today.AddYears(-14)
        DateTimePicker3.MinDate = DateTimePicker2.Value
        Label23.Text = "..."
        calcularEmp()
        llenarComboDpto()
        llenarCombosCiudades()
        llenarComboTC()
    End Sub

    'Cargar ciudades dependiendo del dpto seleccionado
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        llenarComboCiudad(ComboBox3.SelectedIndex + 1)
    End Sub

    'Establecer fecha mínima para la fecha fin de contrato
    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        DateTimePicker3.MinDate = DateTimePicker2.Value
    End Sub

    'Calcular días trabajados de acuerdo a las fechas de inicio y fin 
    Private Sub DateTimePicker3_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker3.ValueChanged
        Dim fechaIni As Date = DateTimePicker2.Value
        Dim fechaFin As Date = DateTimePicker3.Value

        Dim tiempo As Integer = (fechaFin - fechaIni.AddDays(-1)).Days
        If Not (tiempo = 0) Then
            Label23.Text = tiempo.ToString()
        Else
            Label23.Text = "..."
        End If
    End Sub

    'Buscar empleado gestión
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Dim documento As String = TextBox30.Text

        If Not documento = "" Then
            Dim id As Integer = Convert.ToInt32(cnx.execSelect("select id from Hoja_de_vida where Numero_documento = " & documento))
            If Not id = 0 Then
                habilitarLinks(True)
            Else
                habilitarLinks(False)
                MsgBox("No existe un empleado con el documento " & documento)
            End If
        Else
            MsgBox("Ingrese documento")
        End If
    End Sub

    'Buscar empleado descargar/imprimir
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Dim documento As String = TextBox31.Text
        limpiarTab7()
        TextBox31.Text = documento

        If Not documento = "" Then
            Dim id As Integer = cnx.execSelect("select id from Hoja_de_vida where Numero_documento = " & documento)
            If Not Convert.ToInt32(id) = 0 Then
                Button15.Enabled = True
                crearLinksAnexo(documento) 'Links anexos
                crearLinksContratos(documento, 0) 'Links contratos
                crearLinksComunicados(documento, 0) 'Links comunicados
                crearLinksCertificados(documento, 0) 'Links certificados
            Else
                MsgBox("No existen empleados con el número de documento " & documento)
            End If
        Else
            MsgBox("Ingrese documento")
        End If
    End Sub

    'Botón para descargar documentos. Sólo se habilita si hay resultados
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        If Not AxAcroPDF5.src = "none" Then
            AxAcroPDF5.Print()
        End If
    End Sub

    'Cuando se da click en la pestaña gestión para limpiar txt y desabilitar links
    Private Sub TabPage6_Enter(sender As Object, e As EventArgs) Handles TabPage6.Enter
        habilitarLinks(False)
        calcularEmp()
        TextBox30.Text = ""
    End Sub

    'Finalizar gestión
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        habilitarLinks(False)
        TextBox30.Text = ""
    End Sub

    'Poner fecha fin contrato inactiva
    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        If ComboBox5.Text = "Indefinido" Then
            DateTimePicker3.Enabled = False
        Else
            DateTimePicker3.Enabled = True
        End If
    End Sub

#End Region

#Region "Checks"
    'Checks de archivos unchecked
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        quitarArchivos(CheckBox1, TextBox17, AxAcroPDF1)
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        quitarArchivos(CheckBox2, TextBox18, AxAcroPDF1)
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        quitarArchivos(CheckBox3, TextBox19, AxAcroPDF1)
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        quitarArchivos(CheckBox4, TextBox20, AxAcroPDF1)
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        quitarArchivos(CheckBox5, TextBox21, AxAcroPDF1)
    End Sub

    Private Sub CheckBox10_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox10.CheckedChanged
        quitarArchivos(CheckBox10, TextBox22, AxAcroPDF1)
    End Sub

    Private Sub CheckBox9_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox9.CheckedChanged
        quitarArchivos(CheckBox9, TextBox23, AxAcroPDF1)
    End Sub

    Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox8.CheckedChanged
        quitarArchivos(CheckBox8, TextBox24, AxAcroPDF1)
    End Sub

    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        quitarArchivos(CheckBox7, TextBox25, AxAcroPDF1)
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        quitarArchivos(CheckBox6, TextBox26, AxAcroPDF1)
    End Sub

    Private Sub CheckBox12_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox12.CheckedChanged
        quitarArchivos(CheckBox12, TextBox27, AxAcroPDF1)
    End Sub

    Private Sub CheckBox11_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox11.CheckedChanged
        quitarArchivos(CheckBox11, TextBox28, AxAcroPDF1)
    End Sub

    Private Sub CheckBox13_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox13.CheckedChanged
        quitarArchivos(CheckBox13, TextBox38, AxAcroPDF2)
    End Sub

    Private Sub CheckBox14_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox14.CheckedChanged
        quitarArchivos(CheckBox14, TextBox39, AxAcroPDF2)
    End Sub

    Private Sub CheckBox15_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox15.CheckedChanged
        quitarArchivos(CheckBox15, TextBox29, AxAcroPDF3)
    End Sub

    Private Sub CheckBox16_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox16.CheckedChanged
        quitarArchivos(CheckBox16, TextBox32, AxAcroPDF4)
    End Sub
#End Region

#Region "Métodos"

    'Método para redirigirse al menú
    Private Sub salir()
        Me.Hide()
        menum.Show()
    End Sub

    'Método para calcular cantidad empleados
    Private Sub calcularEmp()
        Dim query As String = "SELECT COUNT(Numero_documento) cantidad 
                FROM Hoja_de_vida 
                WHERE Numero_documento IN (SELECT h.Numero_documento 
                                            FROM Hoja_de_vida h 
                                            INNER JOIN contrato c 
	                                        ON h.Numero_documento = c.Nume_documento 
                                            WHERE c.Actividad = 'Activo'
	                                        GROUP BY h.Numero_documento)"

        Label29.Text = cnx.execSelect(query)

        query = "SELECT COUNT(Numero_documento) cantidad 
                FROM Hoja_de_vida 
                WHERE Numero_documento IN (SELECT h.Numero_documento 
                                            FROM Hoja_de_vida h 
                                            INNER JOIN contrato c 
	                                        ON h.Numero_documento = c.Nume_documento 
                                            WHERE c.Actividad = 'Inactivo' 
                                            AND h.Numero_documento NOT IN (SELECT h.Numero_documento 
                                                                            FROM Hoja_de_vida h 
                                                                            INNER JOIN contrato c 
	                                                                        ON h.Numero_documento = c.Nume_documento 
                                                                            WHERE c.Actividad = 'Activo'
	                                                                        GROUP BY h.Numero_documento)
	                                        GROUP BY h.Numero_documento)"
        Label30.Text = cnx.execSelect(query)

        query = "SELECT COUNT(Numero_documento) cantidad FROM Hoja_de_vida"
        Label31.Text = cnx.execSelect(query)
    End Sub

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

    'Método para llenar combo de días, para opción de tiempo corte
    Private Sub llenarComboTC()
        For i As Integer = 1 To 31 Step 1
            ComboBox7.Items.Add(i.ToString())
        Next
    End Sub

    'Método que retorna un archivo en binario para poder almacenarlo
    Function pasarABinario(texto As String)
        If Not (texto = "") Then
            Dim path As String = texto
            Dim ruta As New FileStream(path, FileMode.Open, FileAccess.Read)
            Dim binario(ruta.Length) As Byte
            ruta.Read(binario, 0, ruta.Length) 'Leer archivo y pasarlo a binario
            ruta.Close()

            Return binario 'Retorna el archivo en binario
        Else
            Dim binario(0) As Byte
            Return binario 'Retorna un byte de tamaño 0 vacío, si no se cargó ningún documento
        End If
    End Function

    'Método para leer archivos pdf, escribir ruta y seleccionar check
    Private Sub leerPDF(AxAcroPDF As AxAcroPDF, campo As TextBox, chk As CheckBox)
        'abrir el explorador de archivos desde un link 
        Try
            Dim archivo As New OpenFileDialog
            archivo.Filter = "Archivo PDF|*.pdf"
            If archivo.ShowDialog = DialogResult.OK Then
                campo.Text = archivo.FileName
                AxAcroPDF.src = archivo.FileName
                chk.Checked = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    'Método para cambiar a formato yyyy-MM-dd la fecha y poder almacenarla en la BD
    Function convertFecha(fecha As String)
        Dim vec() As String = Split(fecha, "/")
        Dim fec As String = vec(2) & "-" & vec(1) & "-" & vec(0)

        Return fec
    End Function

    'Método para quitar archivos que ya estaban cargados
    Private Sub quitarArchivos(chk As CheckBox, txtbox As TextBox, Acro As AxAcroPDF)
        If Not (chk.Checked) Then
            txtbox.Text = ""
            Acro.LoadFile("none")
        End If
    End Sub

    'Método para habilitar o desabilitar links de gestión
    Private Sub habilitarLinks(estado As Boolean)
        LinkLabel21.Enabled = estado
        LinkLabel20.Enabled = estado
        LinkLabel19.Enabled = estado
        LinkLabel18.Enabled = estado
        LinkLabel17.Enabled = estado
    End Sub

    'Método para recuperar pdf
    Private Sub recuperarPDF(byt As Byte())
        Try
            Dim directorioArchivo As String
            directorioArchivo = System.AppDomain.CurrentDomain.BaseDirectory() & "temp.pdf"

            If (BytesAArchivo(byt, directorioArchivo)) Then
                AxAcroPDF5.src = directorioArchivo
                My.Computer.FileSystem.DeleteFile(directorioArchivo)
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

    'Método para reiniciar valores de página imprimir
    Private Sub limpiarTab7()
        Button15.Enabled = False
        Panel1.Controls.Clear()
        Panel2.Controls.Clear()
        TextBox31.Text = ""
        i = 0
        y = 0
        AxAcroPDF5.LoadFile("none")
    End Sub

    'Método para crear links de anexos con documento devuelto de la BD
    Private Sub crearLinksAnexo(documento As String)
        anexos.Clear()
        Dim consulta As String = "select documento, RUT, Tarjeta_profesional, Hoja_de_vida, Certificado_academico, 
                    Certificado_laboral, Certificado_bancario, Certificado_Eps, AFP, Antecedentes, Examen_ingreso, Otros
                    from Archivos where Num_documento = " & documento
        Dim archi As List(Of Byte()) = cnx.execSelectVarios(consulta, True)

        If Not archi.Count() = 0 Then
            For i As Integer = 0 To archi.Count() - 1 Step 1
                anexos.Add(archi(i))
                If Not UBound(archi(i)) = 0 Then
                    Select Case i
                        Case 0
                            crearControl("Documento")
                        Case 1
                            crearControl("RUT")
                        Case 2
                            crearControl("Tarjeta_profesional")
                        Case 3
                            crearControl("Hoja_de_vida")
                        Case 4
                            crearControl("Certificado_academico")
                        Case 5
                            crearControl("Certificado_laboral")
                        Case 6
                            crearControl("Certificado_bancario")
                        Case 7
                            crearControl("Certificado_Eps")
                        Case 8
                            crearControl("AFP")
                        Case 9
                            crearControl("Antecedentes")
                        Case 10
                            crearControl("Examen_ingreso")
                        Case 11
                            crearControl("Otros")
                    End Select
                End If
            Next
        End If
    End Sub

    'Método para crear links de contratos con documento devuelto de la BD
    Private Sub crearLinksContratos(documento As String, idArchivo As Integer)
        Dim j As Integer = 1
        Dim k As Integer = 1
        Dim consulta As String
        Dim archi
        contrato.Clear()
        otrosi.Clear()

        If idArchivo = 0 Then
            consulta = "select Contrato_laboral from Contrato where Nume_documento = " & documento
            archi = cnx.execSelectVarios(consulta, True)

            If Not archi.Count() = 0 Then
                For i As Integer = 0 To archi.Count() - 1 Step 1
                    If Not UBound(archi(i)) = 0 Then
                        contrato.Add(archi(i))
                        crearControl("Contrato_laboral" & j)
                        j += 1
                    End If
                Next
            End If
            j = 1

            consulta = "select o.otrosi from OtrosSi o inner join Contrato c on o.id_contrato = c.id 
                            where c.Nume_documento = " & documento
            If Not archi.Count() = 0 Then
                For i As Integer = 0 To archi.Count() - 1 Step 1
                    If Not UBound(archi(i)) = 0 Then
                        otrosi.Add(archi(i))
                        crearControl("Otrosi" & k)
                        k += 1
                    End If
                Next
            End If
            k = 1
        Else
            consulta = "select Contrato_laboral from Contrato where Nume_documento = " & documento &
                " and id = " & idArchivo
            archi = cnx.execSelectVarios(consulta, True)

            If Not archi.Count() = 0 Then
                For i As Integer = 0 To archi.Count() - 1 Step 1
                    If Not UBound(archi(i)) = 0 Then
                        contrato.Add(archi(i))
                        crearControl("Contrato_laboral" & j)
                        j += 1
                    End If
                Next
            End If
            j = 1

            consulta = "select o.otrosi from OtrosSi o inner join Contrato c on o.id_contrato = c.id 
                            where c.Nume_documento = " & documento & " and c.id = " & idArchivo & " order by c.id desc"
            If Not archi.Count() = 0 Then
                For i As Integer = 0 To archi.Count() - 1 Step 1
                    If Not UBound(archi(i)) = 0 Then
                        otrosi.Add(archi(i))
                        crearControl("Otrosi" & k)
                        k += 1
                    End If
                Next
            End If
            k = 1
        End If
    End Sub

    'Método para crear links de comunicados con documento devuelto de la BD
    Private Sub crearLinksComunicados(documento As String, idArchivo As Integer)
        Dim j As Integer = 1
        comunicados.Clear()

        Dim consulta As String
        If idArchivo = 0 Then
            consulta = "select Archivos from comunicados where Nume_documento = " & documento
        Else
            consulta = "select Archivos from comunicados where Nume_documento = " & documento &
                    " and id = " & idArchivo
        End If
        Dim archi As List(Of Byte()) = cnx.execSelectVarios(consulta, True)

        If Not archi.Count() = 0 Then
            For i As Integer = 0 To archi.Count() - 1 Step 1
                If Not UBound(archi(i)) = 0 Then
                    comunicados.Add(archi(i))
                    crearControl("Archivos" & j)
                    j += 1
                End If
            Next
        End If
        j = 1
    End Sub

    'Método para crear links de certificados con documento devuelto de la BD
    Private Sub crearLinksCertificados(documento As String, idArchivo As Integer)
        Dim j As Integer = 1
        certificados.Clear()

        Dim consulta As String
        If idArchivo = 0 Then
            consulta = "select Certificado_laboral from Certificado_laboral where Nume_documento = " & documento
        Else
            consulta = "select Certificado_laboral from Certificado_laboral where Nume_documento = " & documento &
                    " and id = " & idArchivo
        End If
        Dim archi As List(Of Byte()) = cnx.execSelectVarios(consulta, True)

        If Not archi.Count() = 0 Then
            For i As Integer = 0 To archi.Count() - 1 Step 1
                If Not UBound(archi(i)) = 0 Then
                    certificados.Add(archi(i))
                    crearControl("Certificado_labora" & j)
                    j += 1
                End If
            Next
        End If
        j = 1
    End Sub

    'Método para los eventos de click de los link de anexos
    Private Sub eventoLink(ByVal sender As System.Object, ByVal e As EventArgs)
        Dim link = TryCast(sender, LinkLabel)

        If link IsNot Nothing Then

            'Para buscar archivos de anexos
            If link.Name = "Documento" Then
                recuperarPDF(anexos.ElementAt(0))
            ElseIf link.Name = "RUT" Then
                recuperarPDF(anexos.ElementAt(1))
            ElseIf link.Name = "Tarjeta_profesional" Then
                recuperarPDF(anexos.ElementAt(2))
            ElseIf link.Name = "Hoja_de_vida" Then
                recuperarPDF(anexos.ElementAt(3))
            ElseIf link.Name = "Certificado_academico" Then
                recuperarPDF(anexos.ElementAt(4))
            ElseIf link.Name = "Certificado_laboral" Then
                recuperarPDF(anexos.ElementAt(5))
            ElseIf link.Name = "Certificado_bancario" Then
                recuperarPDF(anexos.ElementAt(6))
            ElseIf link.Name = "Certificado_Eps" Then
                recuperarPDF(anexos.ElementAt(7))
            ElseIf link.Name = "AFP" Then
                recuperarPDF(anexos.ElementAt(8))
            ElseIf link.Name = "Antecedentes" Then
                recuperarPDF(anexos.ElementAt(9))
            ElseIf link.Name = "Examen_ingreso" Then
                recuperarPDF(anexos.ElementAt(10))
            ElseIf link.Name = "Otros" Then
                recuperarPDF(anexos.ElementAt(11))
            End If

            'Para buscar archivos de contratos laborales
            For ini As Integer = 1 To contrato.Count() Step 1
                If link.Name = "Contrato_laboral" & ini Then
                    recuperarPDF(contrato.ElementAt(ini - 1))
                End If
            Next

            'Para buscar links de otrosi
            For ini As Integer = 1 To otrosi.Count() Step 1
                If link.Name = "Otrosi" & ini Then
                    recuperarPDF(otrosi.ElementAt(ini - 1))
                End If
            Next

            'Para buscar archivos de comunicados
            For ini As Integer = 1 To comunicados.Count() Step 1
                If link.Name = "Archivos" & ini Then
                    recuperarPDF(comunicados.ElementAt(ini - 1))
                End If
            Next

            'Para buscar archivos de certificados laborales
            For ini As Integer = 1 To certificados.Count() Step 1
                If link.Name = "Certificado_laboral" & ini Then
                    recuperarPDF(certificados.ElementAt(ini - 1))
                End If
            Next
        End If
    End Sub

    'Método para agregar links al form
    Private Sub crearControl(text As String)
        Button15.Enabled = True
        Dim link As New LinkLabel
        With link
            .Visible = True
            .LinkColor = Color.Green
            .Name = text
            .Text = text
            .Width = 200
            AddHandler link.Click, AddressOf eventoLink
        End With
        If i < 12 Then
            link.Location = New Point(5, (link.Top + link.Height + 10) * i)
            Panel1.Controls.Add(link)
            i += 1
        Else
            link.Location = New Point(5, (link.Top + link.Height + 10) * y)
            Panel2.Controls.Add(link)
            y += 1
        End If

    End Sub

    'Método para guardar hoja de vida
    Function guardarHDV()
        If TextBox1.Text = "" Then
            MsgBox("Debe ingresar el nombre")
        ElseIf TextBox2.Text = "" Then
            MsgBox("Debe ingresar el apellido")
        ElseIf ComboBox1.Text = "" Then
            MsgBox("Debe seleccionar el Tipo de documento")
        ElseIf TextBox3.Text = "" Then
            MsgBox("Debe ingresar el número de documento")
        ElseIf cmbLugarExp.Text = "" Then
            MsgBox("Debe seleccionar el lugar de expedición del documento")
        ElseIf cmbLugarNac.Text = "" Then
            MsgBox("Debe seleccionar el lugar de nacimiento")
        ElseIf TextBox5.Text = "" Then
            MsgBox("Debe ingresar el número de celular")

        ElseIf ComboBox2.Text = "" Then
            MsgBox("Debe seleccionar el estado civil")
        ElseIf TextBox7.Text = "" Then
            MsgBox("Debe ingresar la ocupación")
        ElseIf ComboBox3.Text = "" Then
            MsgBox("Debe seleccionar el departamento de residencia")
        ElseIf ComboBox4.Text = "" Then
            MsgBox("Debe seleccionar la ciudad de residencia")
        ElseIf TextBox8.Text = "" Then
            MsgBox("Debe ingresar la dirección")
        ElseIf TextBox9.Text = "" Then
            MsgBox("Debe ingresar el cuenta bancaria")
        ElseIf TextBox16.Text = "" Then
            MsgBox("Debe ingresar la EPS a la que pertenece")
        ElseIf TextBox34.Text = "" Then
            MsgBox("Debe ingresar la AFP a la que pertenece")
        ElseIf TextBox35.Text = "" Then
            MsgBox("Debe ingresar la ARL a la que pertenece")
        Else
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

                Return cnx.ejecutarSP(cmm)
            Catch ex As Exception
                MsgBox(ex.Message)
                Return False
            End Try
        End If
    End Function

    'Método para guardar anexos
    Function guardarAnexos()
        If (TextBox20.Text = "" Or TextBox3.Text = "") Then
            MsgBox("Alguno de los archivos obligatorios no fue cargado")
        Else
            Try
                Dim cmm As New SqlCommand("sp_mant_archivo", cnx.conexion)
                cmm.CommandType = CommandType.StoredProcedure

                'Todos los datos son obligatorios en la BD
                cmm.Parameters.AddWithValue("@Num_documento", TextBox3.Text)
                cmm.Parameters.AddWithValue("@documento", pasarABinario(TextBox17.Text))
                cmm.Parameters.AddWithValue("@RUT", pasarABinario(TextBox18.Text))
                cmm.Parameters.AddWithValue("@Tarjeta_profesional", pasarABinario(TextBox19.Text))
                cmm.Parameters.AddWithValue("@Hoja_de_vida", pasarABinario(TextBox20.Text))
                cmm.Parameters.AddWithValue("@Certificado_academico", pasarABinario(TextBox21.Text))
                cmm.Parameters.AddWithValue("@Certificado_laboral", pasarABinario(TextBox22.Text))
                cmm.Parameters.AddWithValue("@Certificado_bancario", pasarABinario(TextBox23.Text))
                cmm.Parameters.AddWithValue("@Certificado_Eps", pasarABinario(TextBox24.Text))
                cmm.Parameters.AddWithValue("@AFP", pasarABinario(TextBox25.Text))
                cmm.Parameters.AddWithValue("@Antecedentes", pasarABinario(TextBox26.Text))
                cmm.Parameters.AddWithValue("@Examen_ingreso", pasarABinario(TextBox27.Text))
                cmm.Parameters.AddWithValue("@Otros", pasarABinario(TextBox28.Text))

                Return cnx.ejecutarSP(cmm)
            Catch ex As Exception
                Return False
            End Try
        End If
    End Function

#End Region

#Region "Links"

    '------------Leer archivos PDF------------

    'Link archivo documento
    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        leerPDF(AxAcroPDF1, TextBox17, CheckBox1)
    End Sub

    'Link archivo RUT
    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        leerPDF(AxAcroPDF1, TextBox18, CheckBox2)
    End Sub

    'Link archivo tarjeta profesional
    Private Sub LinkLabel4_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel4.LinkClicked
        leerPDF(AxAcroPDF1, TextBox19, CheckBox3)
    End Sub

    'Link archivo hoja de vida
    Private Sub LinkLabel3_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        leerPDF(AxAcroPDF1, TextBox20, CheckBox4)
    End Sub

    'Link archivo certificado académico
    Private Sub LinkLabel8_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel8.LinkClicked
        leerPDF(AxAcroPDF1, TextBox21, CheckBox5)
    End Sub

    'Link archivo certificado laboral
    Private Sub LinkLabel0_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel0.LinkClicked
        leerPDF(AxAcroPDF1, TextBox22, CheckBox10)
    End Sub

    'Link archivo certificado bancario
    Private Sub LinkLabel6_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel6.LinkClicked
        leerPDF(AxAcroPDF1, TextBox23, CheckBox9)
    End Sub

    'Link archivo certificado EPS
    Private Sub LinkLabel5_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel5.LinkClicked
        leerPDF(AxAcroPDF1, TextBox24, CheckBox8)
    End Sub

    'Link archivo certificado fondo de pensiones
    Private Sub LinkLabel11_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel11.LinkClicked
        leerPDF(AxAcroPDF1, TextBox25, CheckBox7)
    End Sub

    'Link archivo antecedentes
    Private Sub LinkLabel10_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel10.LinkClicked
        leerPDF(AxAcroPDF1, TextBox26, CheckBox6)
    End Sub

    'Link archivo exámenes de ingreso
    Private Sub LinkLabel9_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel9.LinkClicked
        leerPDF(AxAcroPDF1, TextBox27, CheckBox12)
    End Sub

    'Link archivo otros
    Private Sub LinkLabel12_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel12.LinkClicked
        leerPDF(AxAcroPDF1, TextBox28, CheckBox11)
    End Sub

    'Link archivo anexar contrato laboral
    Private Sub LinkLabel13_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel13.LinkClicked
        leerPDF(AxAcroPDF2, TextBox38, CheckBox13)
    End Sub

    'Link archivo otrosí
    Private Sub LinkLabel14_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel14.LinkClicked
        leerPDF(AxAcroPDF2, TextBox39, CheckBox14)
    End Sub

    'Link archivo comunicados
    Private Sub LinkLabel15_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel15.LinkClicked
        leerPDF(AxAcroPDF3, TextBox29, CheckBox15)
    End Sub

    'Link archivo certificado laboral
    Private Sub LinkLabel16_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel16.LinkClicked
        leerPDF(AxAcroPDF4, TextBox32, CheckBox16)
    End Sub


    '------------Redirecciones forms de actualización------------

    'Link para redirigir a actualizar datos personales
    Private Sub LinkLabel21_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel21.LinkClicked
        Me.Hide()
        ActualizarDatospersonales.documento = TextBox30.Text
        ActualizarDatospersonales.cargarDatos()
        ActualizarDatospersonales.Show()
    End Sub

    'Link para redirigir a actualizar anexos 
    Private Sub LinkLabel20_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel20.LinkClicked
        Me.Hide()
        Actualizaranexo.documento = TextBox30.Text
        Actualizaranexo.cargarDatos()
        Actualizaranexo.Show()
    End Sub

    'Link para redirigir a actualizar contratos
    Private Sub LinkLabel19_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel19.LinkClicked
        Me.Hide()
        Actualizarcontrato.documento = TextBox30.Text
        Actualizarcontrato.llenarComboId()
        Actualizarcontrato.Show()
    End Sub

    'Link para redirigir a actualizar comunicados
    Private Sub LinkLabel18_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel18.LinkClicked
        Me.Hide()
        Actualizarcomunicados.documento = TextBox30.Text
        Actualizarcomunicados.llenarComboId()
        Actualizarcomunicados.Show()
    End Sub

    'Link para redirigir a actualizar certificado laboral
    Private Sub LinkLabel17_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel17.LinkClicked
        Me.Hide()
        ActualizarCertificadoLaboral.documento = TextBox30.Text
        ActualizarCertificadoLaboral.llenarComboId()
        ActualizarCertificadoLaboral.Show()
    End Sub

#End Region

#Region "Guardar"

    'Guardar contrato
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        TextBox33.Text = TextBox3.Text

        If TextBox3.Text = "" Then
            MsgBox("El campo número de documento de la pestaña 'Datos Personales' no puede estar vacío")
        ElseIf ComboBox5.Text = "" Then
            MsgBox("Debe seleccionar el tipo de contrato")
        ElseIf ComboBox6.Text = "" Then
            MsgBox("Debe seleccionar la actividad")
        ElseIf TextBox13.Text = "" Then
            MsgBox("Debe ingresar el cargo")
        ElseIf TextBox14.Text = "" Then
            MsgBox("Debe ingresar el proyecto")
        ElseIf TextBox15.Text = "" Then
            MsgBox("Debe ingresar el salario")
        ElseIf ComboBox7.Text = "" Then
            MsgBox("Debe seleccionar el día de corte")
        Else
            Try
                'Pasar fecha a formato corto para almacenarlas y luego a largo otra vez
                DateTimePicker2.Format = DateTimePickerFormat.Short
                Dim inicioCon As String = convertFecha(DateTimePicker2.Text)
                DateTimePicker2.Format = DateTimePickerFormat.Long

                Dim cmm As New SqlCommand("sp_mant_contrato", cnx.conexion)
                cmm.CommandType = CommandType.StoredProcedure

                'Todos los datos son obligatorios en la BD
                cmm.Parameters.AddWithValue("@id", 0) 'Cero porque es contrato nuevo
                If TextBox33.ReadOnly Then
                    cmm.Parameters.AddWithValue("@Nume_documento", TextBox33.Text)
                Else
                    cmm.Parameters.AddWithValue("@Nume_documento", TextBox3.Text)
                End If

                cmm.Parameters.AddWithValue("@Contrato_laboral", pasarABinario(TextBox38.Text))
                cmm.Parameters.AddWithValue("@Tipo_Contrato", ComboBox5.Text)
                cmm.Parameters.AddWithValue("@Actividad", ComboBox6.Text)
                cmm.Parameters.AddWithValue("@Cargo", TextBox13.Text)
                cmm.Parameters.AddWithValue("@Por_concepto_de", TextBox14.Text)
                cmm.Parameters.AddWithValue("@salario_base", Convert.ToDecimal(TextBox15.Text))
                cmm.Parameters.AddWithValue("@tiempo_corte", ComboBox7.Text)
                cmm.Parameters.AddWithValue("@Inicio_contrato", inicioCon)
                If DateTimePicker3.Enabled Then 'Si está inhabilitada el valor es null
                    DateTimePicker3.Format = DateTimePickerFormat.Short
                    cmm.Parameters.AddWithValue("@Fin_contrato", convertFecha(DateTimePicker3.Text))
                    DateTimePicker3.Format = DateTimePickerFormat.Long
                Else
                    cmm.Parameters.AddWithValue("@Fin_contrato", DBNull.Value)
                End If

                Dim id As Integer
                If TextBox33.ReadOnly Then 'Como está inhabilitado es porque se guardan todos los datos
                    If guardarHDV() Then 'Se guardaron los datos en hoja de vida
                        If guardarAnexos() Then 'Se guardaron los datos en anexos
                            crearLinksAnexo(TextBox33.Text)
                            If cnx.ejecutarSP(cmm) Then
                                id = Convert.ToInt16(cnx.execSelect("select top(1)id from Contrato order by id desc"))
                                If Not TextBox39.Text = "" Then
                                    Dim c As New SqlCommand("sp_mant_otrosi", cnx.conexion)
                                    c.CommandType = CommandType.StoredProcedure
                                    c.Parameters.AddWithValue("@Otrosi", pasarABinario(TextBox39.Text))
                                    c.Parameters.AddWithValue("@id_contrato", id)
                                    Dim ret As Boolean = cnx.ejecutarSP(c)
                                End If
                                MsgBox("Datos agregados correctamente")
                                crearLinksContratos(TextBox33.Text, id)
                            Else
                                'Eliminar registro almacenado
                                deleteHDV()
                            End If
                        Else
                            'Eliminar registro almacenado
                            deleteHDV()
                        End If
                    Else
                        MsgBox("No se guardó la hoja de vida")
                    End If
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        DateTimePicker3.Enabled = True

    End Sub

    Private Sub deleteHDV()
        Try
            Dim cm As New SqlCommand("sp_eliminar_hdv", cnx.conexion)
            cm.CommandType = CommandType.StoredProcedure
            cm.Parameters.AddWithValue("@Numero_documento", TextBox3.Text)
            Dim s As Boolean = cnx.ejecutarSP(cm)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    'Guardar comunicados
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If (TextBox29.Text = "") Then
            MsgBox("No se ha cargado ningún archivo")
        ElseIf TextBox40.Text = "" Then
            MsgBox("Debe ingresar el número de documento")
        Else
            Try
                Dim cmm As New SqlCommand("sp_mant_comunicado", cnx.conexion)
                cmm.CommandType = CommandType.StoredProcedure

                'Todos los datos son obligatorios en la BD
                cmm.Parameters.AddWithValue("@vlr", 0) 'Cero para insertar
                cmm.Parameters.AddWithValue("@id", 0) 'Cero porque no se necesita actualizar
                cmm.Parameters.AddWithValue("@Nume_documento", TextBox40.Text)
                cmm.Parameters.AddWithValue("@Archivos", pasarABinario(TextBox29.Text))

                If cnx.ejecutarSP(cmm) Then
                    MsgBox("Comunicado agregado correctamente")
                    Dim id As Integer
                    id = Convert.ToInt16(cnx.execSelect("select top(1)id from Comunicados order by id desc"))
                    crearLinksContratos(TextBox40.Text, id)
                Else
                    MsgBox("Hubo un error agregando los datos")
                End If

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    'Guardar certificado laboral
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        If (TextBox32.Text = "") Then
            MsgBox("No se ha cargado ningún archivo")
        ElseIf TextBox41.Text = "" Then
            MsgBox("Deb ingresar el número de documento")
        Else
            Try
                Dim cmm As New SqlCommand("sp_mant_certificado_laboral", cnx.conexion)
                cmm.CommandType = CommandType.StoredProcedure

                'Todos los datos son obligatorios en la BD
                cmm.Parameters.AddWithValue("@vlr", 0) '0 para insertar
                cmm.Parameters.AddWithValue("@id", 0) '0 porque no se necesita actualizar
                cmm.Parameters.AddWithValue("@Nume_documento", TextBox41.Text)
                cmm.Parameters.AddWithValue("@Certificado_laboral", pasarABinario(TextBox32.Text))

                If cnx.ejecutarSP(cmm) Then
                    MsgBox("Certificado agregado correctamente")
                    Dim id As Integer
                    id = Convert.ToInt16(cnx.execSelect("select top(1)id from Certificado_laboral order by id desc"))
                    crearLinksContratos(TextBox41.Text, id)
                Else
                    MsgBox("Hubo un error agregando los datos")
                End If

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

#End Region

#Region "Cancelar"

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        salir()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        salir()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        DateTimePicker3.Enabled = True
        salir()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        salir()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        salir()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        salir()
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        limpiarTab7()
        salir()
    End Sub

    'Siguiente
    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        TabControl1.SelectedIndex = TabControl1.SelectedIndex + 1
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        TabControl1.SelectedIndex = TabControl1.SelectedIndex + 1
        If Not TextBox3.Text = "" Then
            TextBox33.Text = TextBox3.Text
        End If
        Button8.Enabled = True
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        TabControl1.SelectedIndex = TabControl1.SelectedIndex - 1
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        TabControl1.SelectedIndex = TabControl1.SelectedIndex - 1
        Button8.Enabled = False
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged, TextBox1.TextChanged

    End Sub

    Private Sub TextBox33_TextChanged(sender As Object, e As EventArgs) Handles TextBox33.TextChanged

    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

#End Region

End Class