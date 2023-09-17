Public Class menum
    Private Sub hojaDeVida_Click(sender As Object, e As EventArgs) Handles hojaDeVida.Click
        'Cerar ventana actual. Abrir ventana de hoja de vida
        Me.Hide()
        HojaVida.Show()
    End Sub

    Private Sub ajustesPre_Click(sender As Object, e As EventArgs) Handles ajustesPre.Click
        'Cerar ventana actual. Abrir ventana de ajustes predeterminados
        Me.Hide()
        Ajustespredeterminados.Show()
    End Sub

    Private Sub Salir_Click(sender As Object, e As EventArgs) Handles Salir.Click
        'Cerar ventana actual. ir a la ventana de login
        Me.Hide()
        Login.Show()
    End Sub

    Private Sub prestamos_Click(sender As Object, e As EventArgs) Handles prestamos.Click
        'Cerar ventana actual. Abrir ventana de préstamos
        Me.Hide()
        Prestamo.Show()
    End Sub

    Private Sub bNomina_Click(sender As Object, e As EventArgs) Handles bNomina.Click
        'Cerar ventana actual. Abrir ventana de nómina
        Me.Hide()
        Nomina.Show()
    End Sub
End Class