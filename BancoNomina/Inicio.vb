Public Class Inicio

    Dim contador As Byte = 0

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If ProgressBar1.Value < 100 Then
            contador += 25
            ProgressBar1.Value = contador
        Else
            Me.Hide()
            Login.Show()
            Timer1.Enabled = False
        End If
    End Sub
End Class