Imports System.Data.SqlClient
Imports System.IO
Imports AxAcroPDFLib

Public Class Actualizaranexo

    Dim cnx As New Conexion()

    Public documento As String

    Dim archivos As List(Of Byte()) = New List(Of Byte())

#Region "Eventos"

    'Devolverse a gestión
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Hide()
        HojaVida.Show()
        HojaVida.TabPage6.Show()
    End Sub

    'Actualizar datos
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            Dim cmm As New SqlCommand("sp_mant_archivo", cnx.conexion)
            cmm.CommandType = CommandType.StoredProcedure

            'Todos los datos son obligatorios en la BD
            cmm.Parameters.AddWithValue("@Num_documento", documento)

            Dim valor As Integer = validarCampos(TextBox17, CheckBox1)
            If valor = 1 Then
                cmm.Parameters.AddWithValue("@documento", pasarABinario(""))
            ElseIf valor = 2 Then
                cmm.Parameters.AddWithValue("@documento", archivos.ElementAt(0))
            ElseIf valor = 3 Then
                cmm.Parameters.AddWithValue("@documento", pasarABinario(TextBox17.Text))
            End If

            valor = validarCampos(TextBox18, CheckBox2)
            If valor = 1 Then
                cmm.Parameters.AddWithValue("@RUT", pasarABinario(""))
            ElseIf valor = 2 Then
                cmm.Parameters.AddWithValue("@RUT", archivos.ElementAt(1))
            ElseIf valor = 3 Then
                cmm.Parameters.AddWithValue("@RUT", pasarABinario(TextBox18.Text))
            End If

            valor = validarCampos(TextBox19, CheckBox3)
            If valor = 1 Then
                cmm.Parameters.AddWithValue("@Tarjeta_profesional", pasarABinario(""))
            ElseIf valor = 2 Then
                cmm.Parameters.AddWithValue("@Tarjeta_profesional", archivos.ElementAt(2))
            ElseIf valor = 3 Then
                cmm.Parameters.AddWithValue("@Tarjeta_profesional", pasarABinario(TextBox19.Text))
            End If

            valor = validarCampos(TextBox20, CheckBox4)
            If valor = 1 Then
                cmm.Parameters.AddWithValue("@Hoja_de_vida", pasarABinario(""))
            ElseIf valor = 2 Then
                cmm.Parameters.AddWithValue("@Hoja_de_vida", archivos.ElementAt(3))
            ElseIf valor = 3 Then
                cmm.Parameters.AddWithValue("@Hoja_de_vida", pasarABinario(TextBox20.Text))
            End If

            valor = validarCampos(TextBox21, CheckBox5)
            If valor = 1 Then
                cmm.Parameters.AddWithValue("@Certificado_academico", pasarABinario(""))
            ElseIf valor = 2 Then
                cmm.Parameters.AddWithValue("@Certificado_academico", archivos.ElementAt(4))
            ElseIf valor = 3 Then
                cmm.Parameters.AddWithValue("@Certificado_academico", pasarABinario(TextBox21.Text))
            End If

            valor = validarCampos(TextBox22, CheckBox10)
            If valor = 1 Then
                cmm.Parameters.AddWithValue("@Certificado_laboral", pasarABinario(""))
            ElseIf valor = 2 Then
                cmm.Parameters.AddWithValue("@Certificado_laboral", archivos.ElementAt(5))
            ElseIf valor = 3 Then
                cmm.Parameters.AddWithValue("@Certificado_laboral", pasarABinario(TextBox22.Text))
            End If

            valor = validarCampos(TextBox23, CheckBox9)
            If valor = 1 Then
                cmm.Parameters.AddWithValue("@Certificado_bancario", pasarABinario(""))
            ElseIf valor = 2 Then
                cmm.Parameters.AddWithValue("@Certificado_bancario", archivos.ElementAt(6))
            ElseIf valor = 3 Then
                cmm.Parameters.AddWithValue("@Certificado_bancario", pasarABinario(TextBox23.Text))
            End If

            valor = validarCampos(TextBox24, CheckBox8)
            If valor = 1 Then
                cmm.Parameters.AddWithValue("@Certificado_Eps", pasarABinario(""))
            ElseIf valor = 2 Then
                cmm.Parameters.AddWithValue("@Certificado_Eps", archivos.ElementAt(7))
            ElseIf valor = 3 Then
                cmm.Parameters.AddWithValue("@Certificado_Eps", pasarABinario(TextBox24.Text))
            End If

            valor = validarCampos(TextBox25, CheckBox7)
            If valor = 1 Then
                cmm.Parameters.AddWithValue("@AFP", pasarABinario(""))
            ElseIf valor = 2 Then
                cmm.Parameters.AddWithValue("@AFP", archivos.ElementAt(8))
            ElseIf valor = 3 Then
                cmm.Parameters.AddWithValue("@AFP", pasarABinario(TextBox25.Text))
            End If

            valor = validarCampos(TextBox26, CheckBox6)
            If valor = 1 Then
                cmm.Parameters.AddWithValue("@Antecedentes", pasarABinario(""))
            ElseIf valor = 2 Then
                cmm.Parameters.AddWithValue("@Antecedentes", archivos.ElementAt(9))
            ElseIf valor = 3 Then
                cmm.Parameters.AddWithValue("@Antecedentes", pasarABinario(TextBox26.Text))
            End If

            valor = validarCampos(TextBox27, CheckBox12)
            If valor = 1 Then
                cmm.Parameters.AddWithValue("@Examen_ingreso", pasarABinario(""))
            ElseIf valor = 2 Then
                cmm.Parameters.AddWithValue("@Examen_ingreso", archivos.ElementAt(10))
            ElseIf valor = 3 Then
                cmm.Parameters.AddWithValue("@Examen_ingreso", pasarABinario(TextBox27.Text))
            End If

            valor = validarCampos(TextBox28, CheckBox11)
            If valor = 1 Then
                cmm.Parameters.AddWithValue("@Otros", pasarABinario(""))
            ElseIf valor = 2 Then
                cmm.Parameters.AddWithValue("@Otros", archivos.ElementAt(11))
            ElseIf valor = 3 Then
                cmm.Parameters.AddWithValue("@Otros", pasarABinario(TextBox28.Text))
            End If

            If (cnx.ejecutarSP(cmm)) Then
                MsgBox("Anexos actualizados correctamente")
            Else
                MsgBox("Hubo un error actualizando los datos")
                AxAcroPDF1.Dispose()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    'Eliminar datos
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim cmm As New SqlCommand("sp_eliminar_archivo", cnx.conexion)
            cmm.CommandType = CommandType.StoredProcedure

            cmm.Parameters.AddWithValue("@Num_documento", documento)

            If cnx.ejecutarSP(cmm) Then
                MsgBox("El anexo fue eliminado correctamente")
            Else
                MsgBox("Hubo un error eliminando los datos")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#End Region

#Region "Checks"
    'Checks de archivos unchecked
    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        quitarArchivos(CheckBox2, TextBox18, AxAcroPDF1)
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        quitarArchivos(CheckBox3, TextBox19, AxAcroPDF1)
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
#End Region

#Region "Links"

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

#End Region

#Region "Métodos"

    'Método para validar campo y poder enviar valor al sp
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

    'Método para consultar datos y cargarlos
    Function cargarDatos()
        AxAcroPDF1.LoadFile("none")
        Dim consulta As String = ("SELECT [documento], [RUT], [Tarjeta_profesional], [Hoja_de_vida],
                [Certificado_academico], [Certificado_laboral], [Certificado_bancario], [Certificado_Eps], [AFP], 
                [Antecedentes], [Examen_ingreso], [Otros]
                FROM Archivos
                WHERE Num_documento =" & documento)
        Dim valores As List(Of Byte()) = cnx.execSelectVarios(consulta, True)

        If Not valores.Count() = 0 Then
            recuperarPDF(valores.ElementAt(0), CheckBox1)
            recuperarPDF(valores.ElementAt(1), CheckBox2)
            recuperarPDF(valores.ElementAt(2), CheckBox3)
            recuperarPDF(valores.ElementAt(3), CheckBox4)
            recuperarPDF(valores.ElementAt(4), CheckBox5)
            recuperarPDF(valores.ElementAt(5), CheckBox10)
            recuperarPDF(valores.ElementAt(6), CheckBox9)
            recuperarPDF(valores.ElementAt(7), CheckBox8)
            recuperarPDF(valores.ElementAt(8), CheckBox7)
            recuperarPDF(valores.ElementAt(9), CheckBox6)
            recuperarPDF(valores.ElementAt(10), CheckBox12)
            recuperarPDF(valores.ElementAt(11), CheckBox11)
        Else
            MsgBox("No hay anexos relacionados al documento " & documento)
            Button1.Enabled = False
            Button4.Enabled = False
        End If
        modificarCampos()
    End Function

    'Método para quitar archivos que ya estaban cargados
    Private Sub quitarArchivos(chk As CheckBox, txtbox As TextBox, Acro As AxAcroPDF)
        If Not (chk.Checked) Then
            txtbox.Text = "."
            Acro.LoadFile("none")
        End If
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

    'Método para recuperar pdf
    Private Sub recuperarPDF(byt As Byte(), chk As CheckBox)
        Try
            Dim directorioArchivo As String
            directorioArchivo = System.AppDomain.CurrentDomain.BaseDirectory() & "temp.pdf"

            Dim bytes() As Byte
            bytes = byt

            archivos.Add(bytes)

            Dim res As Boolean = BytesAArchivo(bytes, directorioArchivo, chk)
            If (res) Then
                AxAcroPDF1.src = directorioArchivo
                My.Computer.FileSystem.DeleteFile(directorioArchivo)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    'Método para convertir bytes a pdf
    Function BytesAArchivo(ByVal bytes() As Byte, ByVal Path As String, chk As CheckBox)
        Try
            Dim K As Long = UBound(bytes)
            If Not K = 0 Then
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

    'Método que pone todos los campos de archivos con _ para distinguir cuáles actualizar
    Private Sub modificarCampos()
        TextBox17.Text = "_"
        TextBox20.Text = "_"
        TextBox24.Text = "_"
        TextBox18.Text = "_"
        TextBox19.Text = "_"
        TextBox21.Text = "_"
        TextBox22.Text = "_"
        TextBox23.Text = "_"
        TextBox25.Text = "_"
        TextBox26.Text = "_"
        TextBox27.Text = "_"
        TextBox28.Text = "_"
    End Sub

    'Método para ejcutar SP anexos
    Function validarCampos(text As TextBox, chk As CheckBox, sp As String, parametro As String)
        Try
            If (text.Text = ".") Then
                If Not (chk.Checked) Then
                    'Se modifica el archivo a 0
                    Dim cmm As New SqlCommand(sp, cnx.conexion)
                    cmm.CommandType = CommandType.StoredProcedure
                    cmm.Parameters.AddWithValue("@Num_documento", documento)
                    cmm.Parameters.AddWithValue(parametro, pasarABinario(""))
                    Return cnx.ejecutarSP(cmm)
                End If
            ElseIf text.Text = "_" Then
                If chk.Checked Then
                    'No se realiza ninguna acción
                    Return True
                End If
            Else
                'Se modifica archivo por uno nuevo
                Dim cmm As New SqlCommand(sp, cnx.conexion)
                cmm.CommandType = CommandType.StoredProcedure
                cmm.Parameters.AddWithValue("@Num_documento", documento)
                cmm.Parameters.AddWithValue(parametro, pasarABinario(text.Text))
                Return cnx.ejecutarSP(cmm)
            End If

            Return False
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

    End Sub

#End Region

End Class