Imports System.Data.SqlClient

Public Class Registrarme

    Dim cnx As New Conexion()

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'cerrar la ventana actual y abrir el login
        Me.Hide()
        Login.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MsgBox("Debe ingresar el nombre completo")
        ElseIf TextBox2.Text = "" Then
            MsgBox("Debe ingresar el número de documento")
        ElseIf ComboBox1.Text = "" Then
            MsgBox("Debe seleccionar el tipo de documento")
        ElseIf TextBox3.Text = "" Then
            MsgBox("Debe ingresar la contraseña")
        ElseIf TextBox4.Text = "" Then
            MsgBox("Debe confirmar la contraseña")
        Else
            If TextBox3.Text = TextBox4.Text Then
                Try
                    Dim cmm As New SqlCommand("sp_mant_registro", cnx.conexion)
                    cmm.CommandType = CommandType.StoredProcedure

                    'Todos los datos son obligatorios en la BD
                    cmm.Parameters.AddWithValue("@Numero_Documento", TextBox2.Text)
                    cmm.Parameters.AddWithValue("@Tipo_Documento", ComboBox1.Text)
                    cmm.Parameters.AddWithValue("@Nombre", TextBox1.Text)
                    cmm.Parameters.AddWithValue("@Contrasena", TextBox3.Text)

                    If cnx.ejecutarSP(cmm) Then
                        MsgBox("Datos insertados correctamente")
                        Me.Hide()
                        Login.Show()
                    Else
                        MsgBox("Hubo un error agregando los datos")
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            Else
                MsgBox("La contraseña no es igual. Por favor verifique")
            End If
        End If
    End Sub
End Class