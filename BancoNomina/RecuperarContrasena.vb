Imports System.Data.SqlClient

Public Class RecuperarContrasena

    Dim cnx As New Conexion()

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim consulta As String = ("select Contrasena from Ingresar where Numero_documento = " + TextBox1.Text + "")
        Label6.Text = cnx.execSelect(consulta)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'cerrar ventana actual y abrir el login
        Me.Hide()
        Login.Show()
    End Sub
End Class