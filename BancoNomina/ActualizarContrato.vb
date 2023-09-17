Imports System.Data.SqlClient
Imports System.IO
Imports AxAcroPDFLib

Public Class Actualizarcontrato

    Dim cnx As New Conexion()

    Public documento As String
    Dim contrato As Byte()
    Dim otroSi As Byte()

#Region "Eventos"
    'Se ejecuta al cargar la página por primera vez
    Private Sub Actualizarcontrato_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        llenarComboTC()
        TextBox38.Text = "_"
        TextBox39.Text = "_"
        Label23.Text = "..."
    End Sub

    'Actualizar datos
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Try
            'Pasar fecha a formato corto para almacenarlas y luego a largo otra vez
            DateTimePicker2.Format = DateTimePickerFormat.Short
            Dim inicioCon As String = convertFecha(DateTimePicker2.Text)
            DateTimePicker2.Format = DateTimePickerFormat.Long

            Dim cmm As New SqlCommand("sp_mant_contrato", cnx.conexion)
            cmm.CommandType = CommandType.StoredProcedure

            'Todos los datos son obligatorios en la BD
            cmm.Parameters.AddWithValue("@id", Convert.ToInt16(ComboBox1.SelectedItem))
            cmm.Parameters.AddWithValue("@Nume_documento", documento)
            cmm.Parameters.AddWithValue("@Contrato_laboral", contrato)
            cmm.Parameters.AddWithValue("@Tipo_Contrato", ComboBox5.Text)
            cmm.Parameters.AddWithValue("@Actividad", ComboBox6.Text)
            cmm.Parameters.AddWithValue("@Cargo", TextBox13.Text)
            cmm.Parameters.AddWithValue("@Por_concepto_de", TextBox14.Text)
            cmm.Parameters.AddWithValue("@salario_base", Math.Round(Convert.ToDecimal(TextBox15.Text), 2))
            cmm.Parameters.AddWithValue("@tiempo_corte", ComboBox7.Text)
            cmm.Parameters.AddWithValue("@Inicio_contrato", inicioCon)

            If DateTimePicker3.Enabled Then 'Si está inhabilitada el valor es null
                DateTimePicker3.Format = DateTimePickerFormat.Short
                cmm.Parameters.AddWithValue("@Fin_contrato", convertFecha(DateTimePicker3.Text))
                DateTimePicker3.Format = DateTimePickerFormat.Long
            Else
                cmm.Parameters.AddWithValue("@Fin_contrato", DBNull.Value)
            End If

            If (cnx.ejecutarSP(cmm)) Then
                Dim valor As Integer = validarCampos(TextBox39, CheckBox14)
                If valor = 3 Then
                    Dim cm As New SqlCommand("sp_mant_otrosi", cnx.conexion)
                    cm.CommandType = CommandType.StoredProcedure
                    cm.Parameters.AddWithValue("@Otrosi", pasarABinario(TextBox39.Text))
                    cm.Parameters.AddWithValue("@id_contrato", ComboBox1.Text)
                    If Not cnx.ejecutarSP(cm) Then
                        MsgBox("No fue posible guardar el otro si")
                    End If
                End If
                MsgBox("Contrato actualizado correctamente")
                Button8.Enabled = True
                Button1.Enabled = True
            Else
                MsgBox("Hubo un error actualizando los datos")
            End If
        Catch ex As Exception
            'MsgBox(ex.Message)
        End Try
    End Sub

    'Devolverse a gestión
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Me.Hide()
        HojaVida.Show()
        HojaVida.TabPage6.Show()
    End Sub

    'Eliminar datos
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim cmm As New SqlCommand("sp_eliminar_contrato", cnx.conexion)
            cmm.CommandType = CommandType.StoredProcedure

            cmm.Parameters.AddWithValue("@id", Convert.ToInt16(ComboBox1.SelectedItem))
            cmm.Parameters.AddWithValue("@Nume_documento", documento)

            If cnx.ejecutarSP(cmm) Then
                MsgBox("El contrato fue eliminado correctamente")
            Else
                MsgBox("Hubo un error eliminando los datos")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    'Poner fecha fin contrato inactiva
    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        If ComboBox5.Text = "Indefinido" Then
            DateTimePicker3.Enabled = False
        Else
            DateTimePicker3.Enabled = True
        End If
    End Sub

    'Cambio en fecha fin
    Private Sub DateTimePicker3_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker3.ValueChanged
        calcularDias()
    End Sub

#End Region

#Region "Métodos"

    'Método para calcular días
    Private Sub calcularDias()
        Dim fechaIni As Date = DateTimePicker2.Value
        Dim fechaFin As Date = DateTimePicker3.Value

        Dim tiempo As Integer = (fechaFin - fechaIni.AddDays(0)).Days
        If Not (tiempo = 0) Then
            Label23.Text = tiempo.ToString()
        Else
            Label23.Text = "..."
        End If
    End Sub

    'Método para cargar ids de comunicados
    Function llenarComboId()
        ComboBox1.Items.Clear()
        LinkLabel14.Enabled = False
        TextBox14.Enabled = False
        TextBox15.Enabled = False
        TextBox13.Enabled = False
        Button1.Enabled = False
        Button8.Enabled = False
        AxAcroPDF2.LoadFile("none")

        Dim consulta As String = ("select id from contrato where Nume_documento = " & documento)
        Dim valores As List(Of String) = cnx.execSelectVarios(consulta, False)

        If Not valores.Count() = 0 Then
            For Each item As String In valores
                ComboBox1.Items.Add(item)
            Next
        Else
            MsgBox("No hay contratos relacionados al documento " & documento)
            LinkLabel14.Enabled = False
            Button1.Enabled = False
        End If
    End Function

    'Método para consultar datos y cargarlos
    Private Sub cargarDatos(id As String)
        Dim consulta As String = ("SELECT [Contrato_laboral] FROM contrato WHERE id = " & id)
        Dim valores As List(Of Byte()) = cnx.execSelectVarios(consulta, True)

        If Not valores.Count() = 0 Then
            recuperarPDF(valores.ElementAt(0), CheckBox13, "contrato")
        End If

        LinkLabel14.Enabled = True

        consulta = ("select top(1) o.otrosi from OtrosSi o inner join Contrato c on o.id_contrato = c.id 
                        where c.id = " & id & " order by c.id desc")
        Dim val As List(Of Byte()) = cnx.execSelectVarios(consulta, True)

        If Not val.Count() = 0 Then
            recuperarPDF(val.ElementAt(0), CheckBox14, "otrosi")
        End If

        If CheckBox14.Checked Then
            TextBox14.Enabled = True
            TextBox15.Enabled = True
            TextBox13.Enabled = True
        End If

        consulta = ("if (select fin_contrato from Contrato where id = " & id & ") is null
	                    begin
		                    SELECT [Tipo_Contrato], [Actividad], [Cargo], [Por_concepto_de], [salario_base], 
				                    [tiempo_corte], [Inicio_contrato], '' FROM contrato WHERE id = " & id & "
	                    end
                    else
	                    begin
		                    SELECT [Tipo_Contrato], [Actividad], [Cargo], [Por_concepto_de], [salario_base], 
			                    [tiempo_corte], [Inicio_contrato], [fin_contrato] FROM contrato WHERE id = " & id & "
	                    end")
        Dim vals As List(Of String) = cnx.execSelectVarios(consulta, False)

        'Pasar campos de fechas a formato corto para cargarlas
        DateTimePicker2.Format = DateTimePickerFormat.Short
        DateTimePicker3.Format = DateTimePickerFormat.Short

        If Not vals.Count() = 0 Then
            ComboBox5.Text = vals.ElementAt(0)
            ComboBox6.Text = vals.ElementAt(1)
            TextBox13.Text = vals.ElementAt(2)
            TextBox14.Text = vals.ElementAt(3)
            TextBox15.Text = vals.ElementAt(4)
            ComboBox7.SelectedIndex = Convert.ToInt32(vals.ElementAt(5)) - 1
            DateTimePicker2.Text = vals.ElementAt(6)
            If Not vals.ElementAt(7) = "" Then
                DateTimePicker3.Text = vals.ElementAt(7)
            End If
            Button1.Enabled = True
            Button8.Enabled = True
        End If

        'Pasar campos de fechas a formato largo otra vez
        DateTimePicker2.Format = DateTimePickerFormat.Long
        DateTimePicker3.Format = DateTimePickerFormat.Long
    End Sub

    'Método para llenar combo de días, para opción de tiempo corte
    Private Sub llenarComboTC()
        For i As Integer = 1 To 31 Step 1
            ComboBox7.Items.Add(i.ToString())
        Next
    End Sub

    'Método para leer archivos pdf, escribir ruta y seleccionar check
    Private Sub leerPDF(AxAcroPDF1 As AxAcroPDF, campo As TextBox, chk As CheckBox)
        'abrir el explorador de archivos desde un link 
        Try
            Dim archivo As New OpenFileDialog
            archivo.Filter = "Archivo PDF|*.pdf"
            If archivo.ShowDialog = DialogResult.OK Then
                campo.Text = archivo.FileName
                AxAcroPDF1.src = archivo.FileName
                chk.Checked = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
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

    'Método para quitar archivos que ya estaban cargados
    Private Sub quitarArchivos(chk As CheckBox, txtbox As TextBox, Acro As AxAcroPDF)
        If Not (chk.Checked) Then
            txtbox.Text = "."
            Acro.LoadFile("none")
            TextBox14.Enabled = False
            TextBox15.Enabled = False
            TextBox13.Enabled = False
            DateTimePicker3.Enabled = False
            Button8.Enabled = False
        Else
            TextBox14.Enabled = True
            TextBox15.Enabled = True
            TextBox13.Enabled = True
            DateTimePicker3.Enabled = True
            Button8.Enabled = True
        End If
    End Sub

    'Método para recuperar pdf
    Private Sub recuperarPDF(byt As Byte(), chk As CheckBox, campo As String)
        Try
            Dim directorioArchivo As String
            directorioArchivo = System.AppDomain.CurrentDomain.BaseDirectory() & "temp.pdf"

            Dim bytes() As Byte
            bytes = byt

            Dim res As Boolean = BytesAArchivo(bytes, directorioArchivo, chk, campo)
            If (res) Then
                AxAcroPDF2.src = directorioArchivo
                My.Computer.FileSystem.DeleteFile(directorioArchivo)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    'Método para convertir bytes a pdf
    Function BytesAArchivo(ByVal bytes() As Byte, ByVal Path As String, chk As CheckBox, campo As String)
        Try
            Dim K As Long = UBound(bytes)
            If Not K = 0 Then
                If (campo = "contrato") Then
                    contrato = bytes
                ElseIf (campo = "otrosi") Then
                    otroSi = bytes
                End If
                Dim fs As New FileStream(Path, FileMode.OpenOrCreate, FileAccess.Write)
                fs.Write(bytes, 0, K)
                fs.Close()
                chk.Checked = True
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message, ex)
        End Try
    End Function

    'Método para cambiar a formato yyyy-MM-dd la fecha y poder almacenarla en la BD
    Function convertFecha(fecha As String)
        Dim vec() As String = Split(fecha, "/")
        Dim fec As String = vec(2) & "-" & vec(1) & "-" & vec(0)

        Return fec
    End Function

    'Método para validar campos
    Function validarCampos(text As TextBox, chk As CheckBox)
        Try
            If (text.Text = ".") Then
                If Not (chk.Checked) Then
                    'Se modifica el archivo a 0
                    Return 1
                End If
            ElseIf text.Text = "_" Then
                If chk.Checked Then
                    'No se realiza ninguna acción
                    Return 2
                End If
            Else
                'Se modifica archivo por uno nuevo
                Return 3
            End If

            Return 1
        Catch ex As Exception
            Return 1
        End Try
    End Function

#End Region

#Region "Checks"
    'Checks de archivos unchecked
    Private Sub CheckBox14_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox14.CheckedChanged
        quitarArchivos(CheckBox14, TextBox39, AxAcroPDF2)
    End Sub

    Private Sub CheckBox13_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox13.CheckedChanged
        quitarArchivos(CheckBox13, TextBox38, AxAcroPDF2)
    End Sub
#End Region

#Region "Links"

    'Link archivo anexar contrato laboral
    Private Sub LinkLabel13_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel13.LinkClicked
        leerPDF(AxAcroPDF2, TextBox38, CheckBox13)
    End Sub

    'Link archivo otrosí
    Private Sub LinkLabel14_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel14.LinkClicked
        leerPDF(AxAcroPDF2, TextBox39, CheckBox14)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If Not ComboBox1.SelectedItem = "" Then
            cargarDatos(ComboBox1.SelectedItem.ToString())
        End If
    End Sub

#End Region

End Class