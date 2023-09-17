Imports System.Data.SqlClient
Public Class Login

    'Cadena conexión BD
    Dim conexion As String = "Data Source=MIPCESCRITORIO;Initial Catalog=BancoHV;Integrated Security=True"

    Dim cnx As New Conexion()

    'Boton login
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MsgBox("Debe ingresar número de documento")
        ElseIf TextBox1.Text = "" Then
            MsgBox("Debe ingresar contraseña")
        End If

        Dim consulta As String = ("select Contrasena from Ingresar where Numero_documento = " + TextBox1.Text + "")

        If (TextBox2.Text = cnx.execSelect(consulta)) Then
            menum.Show()
            Me.Hide()
        Else
            MsgBox("Datos son incorrectos")
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        'Abrir la ventana Registrarme
        Registrarme.Show()
        Me.Hide()
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        'Abrir la pestaña recuperar contraseña
        RecuperarContrasena.Show()
        Me.Hide()
    End Sub
End Class