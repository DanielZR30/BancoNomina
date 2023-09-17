Public Class Imprimir

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Not AxAcroPDF1.src = "none" Then
            AxAcroPDF1.Print()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        AxAcroPDF1.LoadFile("none")
        Me.Hide()
        Nomina.Show()
        Nomina.TabPage1.Show()
    End Sub
End Class