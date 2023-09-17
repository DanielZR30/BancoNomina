Imports System.Data.SqlClient

Public Class Horasextras

#Region "Variables"

    Dim cnx As New Conexion()

    Public vlrNomina As Double
    Public vrHrsDiurnas As Double
    Public vrHrsNocturnas As Double
    Public vrHrsDominicales As Double
    Public vrHrsFestivas As Double
    Dim vlrPagar As Double
    Dim vlrHora As Single
    Dim vlrHoraDiurna As Single
    Dim vlrHoraNocturna As Single
    Dim vlrHoraDominical As Single
    Dim vlrHoraFestivo As Single

#End Region

#Region "Eventos"

    'Se ejecuta cuando la página se carga la primera vez
    Private Sub Horasextras_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        vlrHora = vlrNomina / 240
        consultarValores()
    End Sub

    'Regresar a nómina
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Hide()
        Nomina.Show()
        Nomina.TabPage1.Show()
    End Sub

    'Guardar
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim vrHrsDiurnas As Double
        Dim vrHrsNocturnas As Double
        Dim vrHrsDominicales As Double
        Dim vrHrsFestivas As Double
        Dim vlrPagar As Double

        If TextBox13.Text = "" Then 'Diurnas
            vrHrsDiurnas = 0
        Else
            vrHrsDiurnas = Convert.ToDouble(TextBox13.Text) * (vlrHora * vlrHoraDiurna)
        End If

        If TextBox16.Text = "" Then 'Nocturnas
            vrHrsNocturnas = 0
        Else
            vrHrsNocturnas = Convert.ToDouble(TextBox16.Text) * (vlrHora * vlrHoraNocturna)
        End If

        If TextBox19.Text = "" Then 'Dominicales
            vrHrsDominicales = 0
        Else
            vrHrsDominicales = Convert.ToDouble(TextBox19.Text) * (vlrHora * vlrHoraDominical)
        End If

        If TextBox21.Text = "" Then 'Festivas
            vrHrsFestivas = 0
        Else
            vrHrsFestivas = Convert.ToDouble(TextBox21.Text) * (vlrHora * vlrHoraFestivo)
        End If

        Nomina.vrHrsDiurnas = vrHrsDiurnas
        Nomina.vrHrsDominicales = vrHrsDominicales
        Nomina.vrHrsFestivas = vrHrsFestivas
        Nomina.vrHrsNocturnas = vrHrsNocturnas

        vlrPagar = vrHrsDiurnas + vrHrsNocturnas + vrHrsDominicales + vrHrsFestivas

        Label31.Text = vrHrsDiurnas.ToString("F2")
        Label3.Text = vrHrsNocturnas.ToString("F2")
        Label4.Text = vrHrsDominicales.ToString("F2")
        Label5.Text = vrHrsFestivas.ToString("F2")
        Label7.Text = vlrPagar.ToString("F2")
        Nomina.hrsExtras(Math.Round(vlrPagar, 2))
    End Sub

#End Region

#Region "Métodos"

    'Método para consultar los valores de recargos
    Private Sub consultarValores()
        Dim consulta As String = "select Hora_diurna, Hora_dominical, Hora_festivo, Hora_nocturna
                                            from ValoresxAnio where Anio = " & Today.Year.ToString()
        Dim valores As List(Of String) = cnx.execSelectVarios(consulta, False)

        If Not valores.Count() = 0 Then
            vlrHoraDiurna = (Convert.ToSingle(valores.ElementAt(0)) / 100) + 1
            vlrHoraDominical = (Convert.ToSingle(valores.ElementAt(1)) / 100) + 1
            vlrHoraFestivo = (Convert.ToSingle(valores.ElementAt(2)) / 100) + 1
            vlrHoraNocturna = (Convert.ToSingle(valores.ElementAt(3)) / 100) + 1
        End If
    End Sub

#End Region

End Class