Imports System.Data.SqlClient

Public Class Recargos

#Region "Variables"

    Dim cnx As New Conexion()

    Public vlrNomina As Double
    Public vrRecNocturno As Double
    Public vrRecDominical As Double
    Public vrRecFestivo As Double
    Dim vlrPagar As Double
    Dim vlrDia As Double
    Dim vlrRecargoNocturno As Single
    Dim vlrRecargoDominical As Single
    Dim vlrRecargoFestivo As Single
#End Region

#Region "Eventos"

    'Se ejecuta al cargar primera vez la página
    Private Sub Recargos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        consultarValores()
        vlrDia = vlrNomina / 30
    End Sub

    'Guardar
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If TextBox13.Text = "" Then 'Nocturno
            vrRecNocturno = 0
        Else
            vrRecNocturno = FormatNumber(Convert.ToDouble(TextBox13.Text) * (vlrDia * vlrRecargoNocturno), 0)
        End If

        If TextBox16.Text = "" Then 'Dominical
            vrRecDominical = 0
        Else
            vrRecDominical = Convert.ToDouble(TextBox16.Text) * (vlrDia * vlrRecargoDominical)
        End If

        If TextBox19.Text = "" Then 'Festivo
            vrRecFestivo = 0
        Else
            vrRecFestivo = Convert.ToDouble(TextBox19.Text) * (vlrDia * vlrRecargoFestivo)
        End If

        vlrPagar = vrRecNocturno + vrRecDominical + vrRecFestivo

        Label31.Text = vrRecNocturno.ToString("F2")
        Label3.Text = vrRecDominical.ToString("F2")
        Label4.Text = vrRecFestivo.ToString("F2")
        Label6.Text = vlrPagar.ToString("F2")
        Nomina.totalRecargos(Math.Round(vlrPagar, 2))
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Hide()
        Nomina.Show()
        Nomina.TabPage1.Show()
    End Sub

#End Region

#Region "Métodos"

    'Método para consultar los valores de recargos
    Private Sub consultarValores()
        Dim consulta As String = "select Recargo_nocturno, Recargo_dominical, Recargo_festivo 
                                            from ValoresxAnio where Anio = " & Today.Year.ToString()
        Dim valores As List(Of String) = cnx.execSelectVarios(consulta, False)

        If Not valores.Count() = 0 Then
            vlrRecargoNocturno = (Convert.ToSingle(valores.ElementAt(0)) / 100) + 1
            vlrRecargoDominical = (Convert.ToSingle(valores.ElementAt(1)) / 100) + 1
            vlrRecargoFestivo = (Convert.ToSingle(valores.ElementAt(2)) / 100) + 1
        End If
    End Sub
#End Region

End Class