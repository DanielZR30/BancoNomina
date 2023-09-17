Imports System.IO
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Interop

Public Class Exportar

    Dim cnx As New Conexion()
    Dim xlibro As Excel.Application
    Dim obword As Word.Application
    Dim wd As Document

    Public plantilla As Integer

    'Guardar
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If ComboBox1.Text = "" Or TextBox1.Text = "" Then
            MsgBox("Debe insertar los valores")
        Else
            Dim nombre As String = TextBox1.Text.Trim()
            Dim fecha As Date = InicialDateExportation.Text
            Dim seleccion As Integer = ComboBox1.SelectedIndex
            Select Case seleccion
                Case 0
                    'Excel
                    nombre = nombre
                    If plantilla = 1 Then
                        exportarExcelColilla(nombre, fecha)
                    ElseIf plantilla = 2 Then
                        exportarExcelCuenta(nombre)
                    Else
                        exportarExcelPrestaciones(nombre)
                    End If
                Case 1
                    'Word
                    exportarWord(nombre, plantilla)
            End Select
        End If
    End Sub


    'Método para exportar datos a excel de colilla de pago
    Private Sub exportarExcelColilla(nombre As String, fechaInicial As Date)
        Try
            Dim rutaAct As String = Path.Combine(Directory.GetCurrentDirectory(), "colillaDePago.xlsx")
            Dim ruta As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\" & nombre & ".xlsx"
            Dim ruta2 As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\" & nombre & ".PDF"

            My.Computer.FileSystem.CopyFile(rutaAct, ruta, True)

            xlibro = CreateObject("Excel.Application")
            xlibro.Workbooks.Open(ruta)
            xlibro.Visible = False

            Dim idcolilla As String
            Dim fecha As String

            Dim fechaFinal As Date = DateAdd("d", 15, fechaInicial)


            Dim consulta As String = "
                select id_colilla_empleado, Fecha_inicial, Fecha_final, (Nombre + ' ' + Apellido) 
                                    nombre, Numero_documento, Salario_base, DATEDIFF(day, Fecha_inicial, Fecha_final) dias, Basico, Aux_transporte, 
				                    Total_Hrs, Total_Rec, Incapacidades, Licencias, Otros_devengados, Total_devengado, Salud, Pension, Otras_deducciones,
                                    Retencion, Prestamos, Total_deducciones, Neto_pagar from Colilla_pago where Fecha_inicial >=
                                    '" + Format(fechaInicial, "yyyy-MM-dd").ToString() + "' and Fecha_final <= '" + Format(fechaFinal, "yyyy-MM-dd").ToString() + "'"

            Dim datos As List(Of String) = cnx.execSelectVarios(consulta, False)

            If Not datos.Count() = 0 Then
                Dim cantidad As Integer = datos.Count()
                If Convert.ToInt16(cantidad) > 1 Then
                    For i As Integer = 0 To cantidad - 2 Step 1
                        xlibro.Range("A" & (11 + i)).EntireRow.Insert()
                    Next
                End If

                Dim contC As Integer = 6
                Dim rangoA As Integer = 1
                Dim rangoE As Integer = 33
                Dim contador = 36
                Dim contD As Integer = 10
                Dim con As Integer = 0
                Dim ultimaFila As Integer = 10

                For i As Integer = 0 To datos.Count() - 1 Step 22
                    idcolilla = "No-" & datos.ElementAt(i)

                    xlibro.Sheets("NOMINA").Select()
                    If i = 0 Then
                        fecha = datos.ElementAt(i + 1) & " A " & datos.ElementAt(i + 2)
                        xlibro.Range("H4").Value = fecha
                    End If

                    xlibro.Range("B" & (10 + con)).Value = datos.ElementAt(i + 3)
                    xlibro.Range("C" & (10 + con)).Value = datos.ElementAt(i + 4)
                    xlibro.Range("D" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 5))
                    xlibro.Range("E" & (10 + con)).Value = Convert.ToInt16(datos.ElementAt(i + 6))
                    xlibro.Range("F" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 7))
                    xlibro.Range("G" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 8))
                    xlibro.Range("H" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 9))
                    xlibro.Range("I" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 10))
                    xlibro.Range("J" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 11))
                    xlibro.Range("K" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 12))
                    xlibro.Range("L" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 13))
                    xlibro.Range("M" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 14))
                    xlibro.Range("N" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 15))
                    xlibro.Range("O" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 16))
                    xlibro.Range("P" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 17))
                    xlibro.Range("Q" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 18))
                    xlibro.Range("R" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 19))
                    xlibro.Range("S" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 20))
                    xlibro.Range("T" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 21))
                    xlibro.Range("U" & (10 + con)).Value = "Efectivo"

                    xlibro.Sheets("RECIBOS").Select()
                    xlibro.Range("C" & contC).Value = idcolilla

                    If i > 0 Then
                        xlibro.Range("A1:E33").Copy()
                        xlibro.Range("A" & rangoA & ":E" & rangoE).PasteSpecial(XlPasteType.xlPasteAll)

                        xlibro.Range("E" & contC).Formula = "=+NOMINA!H4"
                        xlibro.Range("C" & (contC + 1)).Formula = "=+NOMINA!" & ("B" & (10 + con))
                        xlibro.Range("C" & (contC + 2)).Formula = "=+NOMINA!" & ("C" & (10 + con))
                        xlibro.Range("D" & contD).Formula = "=+NOMINA!" & ("E" & (10 + con))
                        xlibro.Range("D" & (contD + 1)).Formula = "=+NOMINA!" & ("F" & (10 + con))
                        xlibro.Range("D" & (contD + 2)).Formula = "=+NOMINA!" & ("G" & (10 + con))
                        xlibro.Range("D" & (contD + 3)).Formula = "=+NOMINA!" & ("H" & (10 + con))
                        xlibro.Range("D" & (contD + 4)).Formula = "=+NOMINA!" & ("I" & (10 + con))
                        xlibro.Range("D" & (contD + 5)).Formula = "=+NOMINA!" & ("J" & (10 + con))
                        xlibro.Range("D" & (contD + 6)).Formula = "=+NOMINA!" & ("K" & (10 + con))
                        xlibro.Range("D" & (contD + 7)).Formula = "=+NOMINA!" & ("L" & (10 + con))
                        xlibro.Range("D" & (contD + 10)).Formula = "=+NOMINA!" & ("N" & (10 + con))
                        xlibro.Range("D" & (contD + 11)).Formula = "=+NOMINA!" & ("O" & (10 + con))
                        xlibro.Range("D" & (contD + 12)).Formula = "=+NOMINA!" & ("P" & (10 + con))
                        xlibro.Range("D" & (contD + 13)).Formula = "=+NOMINA!" & ("Q" & (10 + con))
                        xlibro.Range("D" & (contD + 14)).Formula = "=+NOMINA!" & ("R" & (10 + con))
                    End If

                    contC += contador
                    contD += contador
                    rangoA += contador
                    rangoE += contador
                    con += 1
                    ultimaFila += 1
                Next
                xlibro.Sheets("NOMINA").Select()

                xlibro.Range("D" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("D10:D" & ultimaFila - 1))
                xlibro.Range("F" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("F10:F" & ultimaFila - 1))
                xlibro.Range("G" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("G10:G" & ultimaFila - 1))
                xlibro.Range("H" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("H10:H" & ultimaFila - 1))
                xlibro.Range("I" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("I10:I" & ultimaFila - 1))
                xlibro.Range("J" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("J10:J" & ultimaFila - 1))
                xlibro.Range("K" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("K10:K" & ultimaFila - 1))
                xlibro.Range("L" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("L10:L" & ultimaFila - 1))
                xlibro.Range("M" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("M10:M" & ultimaFila - 1))
                xlibro.Range("N" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("N10:N" & ultimaFila - 1))
                xlibro.Range("O" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("O10:O" & ultimaFila - 1))
                xlibro.Range("P" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("P10:P" & ultimaFila - 1))
                xlibro.Range("Q" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("Q10:Q" & ultimaFila - 1))
                xlibro.Range("R" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("R10:R" & ultimaFila - 1))
                xlibro.Range("S" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("S10:S" & ultimaFila - 1))
                xlibro.Range("T" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("T10:T" & ultimaFila - 1))

                xlibro.ActiveWorkbook.Save()
                guardarArchivo(ruta2, 1)
                MsgBox("El archivo fue almacenado exitosamente en la ruta " & ruta)
            Else
                MsgBox("No se han generado colillas de pago correspondientes a la fecha")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox("Posiblemente hay un archivo existente con el nombre " & nombre & " y se encuentra abierto, o no fue posible sobreescribirlo")
        Finally
            xlibro.DisplayAlerts() = False
            xlibro.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlibro)
        End Try
    End Sub

    'Método para exportar datos a excel de cuenta de cobro
    Private Sub exportarExcelCuenta(nombre As String)
        Try
            Dim rutaAct As String = Path.Combine(Directory.GetCurrentDirectory(), "cuentaDeCobro.xlsx")
            Dim ruta As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\" & nombre & ".xlsx"
            Dim ruta2 As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\" & nombre & ".PDF"

            My.Computer.FileSystem.CopyFile(rutaAct, ruta, True)

            xlibro = CreateObject("Excel.Application")
            xlibro.Workbooks.Open(ruta)
            xlibro.Visible = False

            Dim fecha As String
            Dim con As Integer = 1
            Dim sheet As String = "CUENTA"

            Dim consulta As String = "select distinct(Numero_documento) from Cuenta_cobro where 
                                            DATEPART(month, Fecha_final)  = DATEPART(month, GETDATE()) "
            Dim datos As List(Of String) = cnx.execSelectVarios(consulta, False)

            If Not datos.Count() = 0 Then
                xlibro.Sheets(sheet).select()
                For i As Integer = 0 To datos.Count() - 1 Step 1
                    If i > 0 Then
                        sheet = "CUENTA" & i.ToString()
                        xlibro.Sheets("CUENTA").range("A1:F30").select()
                        xlibro.Sheets("CUENTA").range("A1:F30").Copy()
                        xlibro.ActiveWorkbook.Sheets.Add().name = sheet
                        xlibro.Sheets(sheet).activate
                        xlibro.Sheets(sheet).range("A1:F30").select()
                        xlibro.Sheets(sheet).paste()
                    End If

                    consulta = "select cc.Fecha_final,
                        (h.nombre + ' ' + h.apellido) nombre, h.Numero_documento, h.direccion, h.Telefono, 
                        d.nombre dpto, m.nombre ciudad, cc.Concepto_de, convert(integer,cc.Basico) from Cuenta_cobro cc inner join
                        Hoja_de_vida h on cc.Numero_documento = h.Numero_documento inner join Contrato co on 
                        co.Nume_documento = h.Numero_documento inner join Municipios m on h.Ciudad = m.id inner 
                        join Departamentos d on m.departamento_id = d.id where DATEPART(day, GETDATE()) <= co.tiempo_corte 
                        and h.Numero_documento = " & datos.ElementAt(i) & " and co.Fin_contrato >= GETDATE() 
                        and co.Inicio_contrato <= GETDATE()"
                    Dim dat As List(Of String) = cnx.execSelectVarios(consulta, False)
                    fecha = "Fecha: " & dat.ElementAt(0)

                    xlibro.Range("C3").Value = fecha
                    xlibro.Range("A8").Value = dat.ElementAt(1)
                    xlibro.Range("D8").Value = dat.ElementAt(2)
                    xlibro.Range("A10").Value = dat.ElementAt(3)
                    xlibro.Range("C10").Value = dat.ElementAt(4)
                    xlibro.Range("D10").Value = dat.ElementAt(5)
                    xlibro.Range("F10").Value = dat.ElementAt(6)
                    xlibro.Range("A13").Value = dat.ElementAt(7)
                    xlibro.Range("A18").Value = Convert.ToDecimal(dat.ElementAt(8))
                    xlibro.Range("A19").Value = EnLetras(dat.ElementAt(8)).ToUpper()

                    xlibro.ActiveWorkbook.Save()
                    guardarArchivo(ruta2, 2)
                    MsgBox("El archivo del empleado con número documento " & dat.ElementAt(2) & " fue almacenado exitosamente en la ruta " & ruta)
                Next
            Else
                MsgBox("No se han generado cuentas de cobro correspondientes a la fecha")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox("Posiblemente hay un archivo existente con el nombre " & nombre & " y se encuentra abierto, o no fue posible sobreescribirlo")
        Finally
            xlibro.DisplayAlerts() = False
            xlibro.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlibro)
        End Try
    End Sub

    'Método para exportar datos a excel de prestaciones sociales
    Private Sub exportarExcelPrestaciones(nombre As String)
        Try
            Dim rutaAct As String = Path.Combine(Directory.GetCurrentDirectory(), "Liquidaciones.xlsx")
            Dim ruta As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\" & nombre & ".xlsx"
            Dim ruta2 As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\" & nombre & ".PDF"

            My.Computer.FileSystem.CopyFile(rutaAct, ruta, True)

            xlibro = CreateObject("Excel.Application")
            xlibro.Workbooks.Open(ruta)
            xlibro.Visible = False

            Dim cont As Integer = 0
            Dim contA As Integer
            Dim contB As Integer
            Dim contC As Integer
            Dim contador As Integer = 47

            Dim sheet As String = "LIQUIDACION"

            'Liquidar por persona
            Dim consulta As String = "select Numero_documento from Liquidaciones where DATEPART(year, Fecha_final) = DATEPART(year, getdate())"
            Dim datos As List(Of String) = cnx.execSelectVarios(consulta, False)

            If Not datos.Count = 0 Then
                xlibro.Sheets(sheet).select()
                For i As Integer = 0 To datos.Count() - 1 Step 1
                    contA = 28
                    contB = 7
                    contC = 14

                    If i > 0 Then
                        sheet = "LIQUIDACION" & i.ToString()
                        xlibro.Sheets("LIQUIDACION").range("A1:C44").select()
                        xlibro.Sheets("LIQUIDACION").range("A1:C443").Copy()
                        xlibro.ActiveWorkbook.Sheets.Add().name = sheet
                        xlibro.Sheets(sheet).activate
                        xlibro.Sheets(sheet).range("A1:C44").select()
                        xlibro.Sheets(sheet).paste()
                        xlibro.Sheets(sheet).select()
                    End If

                    Dim encabezado As String
                    Dim lapso As String
                    Dim dias As String
                    Dim mes As String
                    Dim documento As String = datos.ElementAt(i)
                    Dim rangoA As Integer = 1
                    Dim rangoC As Integer = 44

                    'Prima
                    For d As Integer = 1 To 2 Step 1
                        If d = 1 Then
                            consulta = "select (h.Nombre + ' ' + h.Apellido)nombre, c.Cargo, c.Inicio_contrato, c.salario_base, 
                            l.Aux_transporte, l.Fecha_inicial, l.Fecha_final, convert(integer,l.Total_liquidado), l.Descripcion_prestacion
                            from Hoja_de_vida h inner join Contrato c on h.Numero_documento = c.Nume_documento 
                            inner join Liquidaciones l on h.Numero_documento = l.Numero_documento where 
                            DATEPART(MONTH, Fecha_final) <= 6 and h.Numero_documento = " & documento & " and 
                            (Descripcion_prestacion = 'Prima'or Descripcion_prestacion = 'Liquidar todos')"
                            encabezado = "LIQUIDACIÓN PRIMA 1ER SEMESTRE " & Today.Year
                            lapso = "07 de enero de " & Today.Year & " a 30 de junio de " & Today.Year
                            dias = "treinta (30)"
                            mes = "junio de " & Today.Year
                        Else
                            consulta = "select (h.Nombre + ' ' + h.Apellido)nombre, c.Cargo, c.Inicio_contrato, c.salario_base, 
                            l.Aux_transporte, l.Fecha_inicial, l.Fecha_final, convert(integer,l.Total_liquidado), l.Descripcion_prestacion
                            from Hoja_de_vida h inner join Contrato c on h.Numero_documento = c.Nume_documento 
                            inner join Liquidaciones l on h.Numero_documento = l.Numero_documento where 
                            DATEPART(MONTH, Fecha_final) > 6 and h.Numero_documento = " & documento & " and 
                            (Descripcion_prestacion = 'Prima' or Descripcion_prestacion = 'Liquidar todos')"
                            encabezado = "LIQUIDACIÓN PRIMA 2DO SEMESTRE " & Today.Year
                            lapso = "01 de julio de " & Today.Year & " a 15 de diciembre de " & Today.Year
                            dias = "quince (15)"
                            mes = "diciembre de " & Today.Year
                        End If
                        Dim dat As List(Of String) = cnx.execSelectVarios(consulta, False)

                        If Not dat.Count() = 0 Then
                            If d = 2 Then
                                xlibro.Range("A1:C44").Copy()
                                xlibro.Range("A" & rangoA & ":C" & rangoC).PasteSpecial(XlPasteType.xlPasteAll)
                            End If
                            xlibro.Range("A" & contB).Value = encabezado
                            xlibro.Range("B" & (contB + 1)).Value = dat.ElementAt(0)
                            xlibro.Range("B" & (contB + 2)).Value = documento
                            xlibro.Range("B" & (contB + 5)).Value = dat.ElementAt(1)
                            xlibro.Range("C" & contC).Value = dat.ElementAt(2)
                            xlibro.Range("C" & (contC + 1)).Value = Convert.ToSingle(dat.ElementAt(3))
                            xlibro.Range("C" & (contC + 2)).Value = Convert.ToSingle(dat.ElementAt(4))
                            xlibro.Range("C" & (contC + 3)).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("C" & (contC + 1) & ":C" & (contC + 2)))
                            xlibro.Range("A" & (contC + 4)).Value = "PRIMA SERVICIOS"
                            xlibro.Range("A" & (contC + 6)).Value = "DIA PRIMA SERVICIOS"
                            xlibro.Range("C" & (contC + 5)).Value = dat.ElementAt(5) & " A " & dat.ElementAt(6)
                            Dim cantDias As Integer = calcularDias(dat.ElementAt(5), dat.ElementAt(6))
                            xlibro.Range("C" & (contC + 6)).Value = "( " & cantDias & ")"
                            Dim total As Single = Convert.ToSingle(dat.ElementAt(7))
                            If dat.ElementAt(8) = "Liquidar todos" Then
                                total = (calcularPromNeto(documento, 1) * cantDias) / 360
                            End If
                            xlibro.Range("C" & (contC + 10)).Value = total
                            xlibro.Range("B" & (contB + 19)).Value = dat.ElementAt(0)
                            xlibro.Range("A" & contA).Value = EnLetras(Convert.ToInt64(total)) & " pesos (" & Convert.ToInt64(total) & ")"
                            xlibro.Range("B" & (contA + 1)).Value = "prima de servicios"
                            xlibro.Range("B" & (contA + 2)).Value = lapso
                            xlibro.Range("A" & (contA + 5)).Value = dias
                            xlibro.Range("C" & (contA + 5)).Value = mes
                            xlibro.Range("A" & (contA + 14)).Value = dat.ElementAt(0)
                            xlibro.Range("B" & (contA + 15)).Value = documento

                            contA += contador
                            contB += contador
                            contC += contador
                            rangoA += contador
                            rangoC += contador
                        End If
                    Next

                    'Cesantías
                    Dim contr As String = cnx.execSelect("select Tipo_Contrato from Contrato where Nume_documento = " & documento)

                    If Today.Month = 12 Or contr = "Indefinido" Then
                        consulta = "select (h.Nombre + ' ' + h.Apellido)nombre, c.Cargo, c.Inicio_contrato, c.salario_base, 
                            l.Aux_transporte, l.Fecha_inicial, l.Fecha_final, convert(integer,l.Total_liquidado), l.Descripcion_prestacion
							from Hoja_de_vida h inner join Contrato c on h.Numero_documento = c.Nume_documento 
                            inner join Liquidaciones l on h.Numero_documento = l.Numero_documento where 
                            h.Numero_documento = " & documento & " and (Descripcion_prestacion = 'Cesantías' 
                            or Descripcion_prestacion = 'Liquidar todos') or c.Tipo_Contrato = 'Indefinido'"
                        encabezado = "LIQUIDACIÓN CESANTÍAS " & mesLetras(Today.Month).ToUpper() & " " & Today.Year
                        mes = " de " & Today.Year
                        Dim dat As List(Of String) = cnx.execSelectVarios(consulta, False)

                        If Not dat.Count() = 0 Then
                            xlibro.Range("A1:C44").Copy()
                            xlibro.Range("A" & rangoA & ":C" & rangoC).PasteSpecial(XlPasteType.xlPasteAll)

                            xlibro.Range("A" & contB).Value = encabezado
                            xlibro.Range("B" & (contB + 1)).Value = dat.ElementAt(0)
                            xlibro.Range("B" & (contB + 2)).Value = documento
                            xlibro.Range("B" & (contB + 5)).Value = dat.ElementAt(1)
                            xlibro.Range("C" & contC).Value = dat.ElementAt(2)
                            xlibro.Range("C" & (contC + 1)).Value = Convert.ToSingle(dat.ElementAt(3))
                            xlibro.Range("C" & (contC + 2)).Value = Convert.ToSingle(dat.ElementAt(4))
                            xlibro.Range("C" & (contC + 3)).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("C" & (contC + 1) & ":C" & (contC + 2)))
                            xlibro.Range("A" & (contC + 4)).Value = "CESANTÍAS"
                            xlibro.Range("A" & (contC + 6)).Value = "DIA CESANTÍAS"
                            xlibro.Range("C" & (contC + 5)).Value = dat.ElementAt(5) & " A " & dat.ElementAt(6)
                            Dim cantDias As Integer = calcularDias(dat.ElementAt(5), dat.ElementAt(6))
                            xlibro.Range("C" & (contC + 6)).Value = "( " & cantDias & ")"
                            Dim total As Single = Convert.ToSingle(dat.ElementAt(7))
                            If dat.ElementAt(8) = "Liquidar todos" Then
                                total = (calcularPromNeto(documento, 1) * cantDias) / 360
                            End If
                            xlibro.Range("C" & (contC + 10)).Value = total
                            xlibro.Range("B" & (contB + 19)).Value = dat.ElementAt(0)
                            xlibro.Range("A" & contA).Value = EnLetras(Convert.ToInt64(total)) & " pesos (" & Convert.ToInt64(total) & ")"
                            xlibro.Range("B" & (contA + 1)).Value = "cesantías"
                            xlibro.Range("B" & (contA + 2)).Value = dat.ElementAt(5) & " A " & dat.ElementAt(6)
                            xlibro.Range("A" & (contA + 5)).Value = EnLetras(Today.Day) & " (" & Today.Day & ")"
                            xlibro.Range("C" & (contA + 5)).Value = mesLetras(Today.Month) & " " & mes
                            xlibro.Range("A" & (contA + 14)).Value = dat.ElementAt(0)
                            xlibro.Range("B" & (contA + 15)).Value = documento

                            contA += contador
                            contB += contador
                            contC += contador
                            rangoA += contador
                            rangoC += contador
                        End If

                        'Intereses a las cesantías
                        consulta = "select (h.Nombre + ' ' + h.Apellido)nombre, c.Cargo, c.Inicio_contrato, c.salario_base, 
                            l.Aux_transporte, l.Fecha_inicial, l.Fecha_final, convert(integer,l.Total_liquidado), l.Descripcion_prestacion
							from Hoja_de_vida h inner join Contrato c on h.Numero_documento = c.Nume_documento 
                            inner join Liquidaciones l on h.Numero_documento = l.Numero_documento where 
                            h.Numero_documento = " & documento & " and (Descripcion_prestacion = 'Intereses a las cesantías'
                            or Descripcion_prestacion = 'Liquidar todos') or c.Tipo_Contrato = 'Indefinido'"
                        encabezado = "LIQUIDACIÓN INTERESES A LAS CESANTÍAS " & mesLetras(Today.Month).ToUpper() & " " & Today.Year
                        mes = " de " & Today.Year

                        dat = cnx.execSelectVarios(consulta, False)
                        If Not dat.Count() = 0 Then
                            xlibro.Range("A1:C44").Copy()
                            xlibro.Range("A" & rangoA & ":C" & rangoC).PasteSpecial(XlPasteType.xlPasteAll)

                            xlibro.Range("A" & contB).Value = encabezado
                            xlibro.Range("B" & (contB + 1)).Value = dat.ElementAt(0)
                            xlibro.Range("B" & (contB + 2)).Value = documento
                            xlibro.Range("B" & (contB + 5)).Value = dat.ElementAt(1)
                            xlibro.Range("C" & contC).Value = dat.ElementAt(2)
                            xlibro.Range("C" & (contC + 1)).Value = Convert.ToSingle(dat.ElementAt(3))
                            xlibro.Range("C" & (contC + 2)).Value = Convert.ToSingle(dat.ElementAt(4))
                            xlibro.Range("C" & (contC + 3)).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("C" & (contC + 1) & ":C" & (contC + 2)))
                            xlibro.Range("A" & (contC + 4)).Value = "INTERESES EN LAS CESANTÍAS"
                            xlibro.Range("A" & (contC + 6)).Value = "DIA INT. EN LAS CESANTÍAS"
                            xlibro.Range("C" & (contC + 5)).Value = dat.ElementAt(5) & " A " & dat.ElementAt(6)
                            Dim cantDias As Integer = calcularDias(dat.ElementAt(5), dat.ElementAt(6))
                            xlibro.Range("C" & (contC + 6)).Value = "( " & cantDias & ")"
                            Dim total As Single = Convert.ToSingle(dat.ElementAt(7))
                            If dat.ElementAt(8) = "Liquidar todos" Then
                                Dim neto As Single = ((calcularPromNeto(documento, 1) * cantDias) / 360)
                                total = Math.Round(((neto * cantDias * 0.12) / 360), 2)
                            End If
                            xlibro.Range("C" & (contC + 10)).Value = total
                            xlibro.Range("B" & (contB + 19)).Value = dat.ElementAt(0)
                            xlibro.Range("A" & contA).Value = EnLetras(Convert.ToInt64(total)) & " pesos (" & Convert.ToInt64(total) & ")"
                            xlibro.Range("B" & (contA + 1)).Value = "intereses a las cesantías"
                            xlibro.Range("B" & (contA + 2)).Value = dat.ElementAt(5) & " A " & dat.ElementAt(6)
                            xlibro.Range("A" & (contA + 5)).Value = EnLetras(Today.Day) & " (" & Today.Day & ")"
                            xlibro.Range("C" & (contA + 5)).Value = mesLetras(Today.Month) & " " & mes
                            xlibro.Range("A" & (contA + 14)).Value = dat.ElementAt(0)
                            xlibro.Range("B" & (contA + 15)).Value = documento

                            contA += contador
                            contB += contador
                            contC += contador
                            rangoA += contador
                            rangoC += contador
                        End If
                    End If

                    'Vacaciones
                    If Today.Month = 12 Then
                        consulta = "select (h.Nombre + ' ' + h.Apellido)nombre, c.Cargo, c.Inicio_contrato, c.salario_base, 
                            l.Aux_transporte, l.Fecha_inicial, l.Fecha_final, convert(integer,l.Total_liquidado), l.Descripcion_prestacion
							from Hoja_de_vida h inner join Contrato c on h.Numero_documento = c.Nume_documento 
                            inner join Liquidaciones l on h.Numero_documento = l.Numero_documento where 
                            h.Numero_documento = " & documento & " and (Descripcion_prestacion = 'Vacaciones'
                            or Descripcion_prestacion = 'Liquidar todos')"
                        encabezado = "LIQUIDACIÓN VACACIONES " & Today.Year
                        dias = "quince (15)"
                        mes = "diciembre de " & Today.Year

                        Dim dat As List(Of String) = cnx.execSelectVarios(consulta, False)
                        If Not dat.Count() = 0 Then
                            xlibro.Range("A1:C44").Copy()
                            xlibro.Range("A" & rangoA & ":C" & rangoC).PasteSpecial(XlPasteType.xlPasteAll)

                            xlibro.Range("A" & contB).Value = encabezado
                            xlibro.Range("B" & (contB + 1)).Value = dat.ElementAt(0)
                            xlibro.Range("B" & (contB + 2)).Value = documento
                            xlibro.Range("B" & (contB + 5)).Value = dat.ElementAt(1)
                            xlibro.Range("C" & contC).Value = dat.ElementAt(2)
                            xlibro.Range("C" & (contC + 1)).Value = Convert.ToSingle(dat.ElementAt(3))
                            xlibro.Range("C" & (contC + 2)).Value = Convert.ToSingle(dat.ElementAt(4))
                            xlibro.Range("C" & (contC + 3)).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("C" & (contC + 1) & ":C" & (contC + 2)))
                            xlibro.Range("A" & (contC + 4)).Value = "VACACIONES"
                            xlibro.Range("A" & (contC + 6)).Value = "DIA VACACIONES"
                            xlibro.Range("C" & (contC + 5)).Value = dat.ElementAt(5) & " A " & dat.ElementAt(6)
                            Dim cantDias As Integer = calcularDias(dat.ElementAt(5), dat.ElementAt(6))
                            xlibro.Range("C" & (contC + 6)).Value = "( " & cantDias & ")"
                            Dim total As Single = Convert.ToSingle(dat.ElementAt(7))
                            If dat.ElementAt(8) = "Liquidar todos" Then
                                total = Math.Round(((calcularPromNeto(documento, 2) * cantDias) / 720), 2)
                            End If
                            xlibro.Range("C" & (contC + 10)).Value = total
                            xlibro.Range("B" & (contB + 19)).Value = dat.ElementAt(0)
                            xlibro.Range("A" & contA).Value = EnLetras(Convert.ToInt64(total)) & " pesos (" & Convert.ToInt64(total) & ")"
                            xlibro.Range("B" & (contA + 1)).Value = "vacaciones"
                            xlibro.Range("B" & (contA + 2)).Value = dat.ElementAt(5) & " A " & dat.ElementAt(6)
                            xlibro.Range("A" & (contA + 5)).Value = dias
                            xlibro.Range("C" & (contA + 5)).Value = mes
                            xlibro.Range("A" & (contA + 14)).Value = dat.ElementAt(0)
                            xlibro.Range("B" & (contA + 15)).Value = documento

                            contA += contador
                            contB += contador
                            contC += contador
                            rangoA += contador
                            rangoC += contador
                        End If
                    End If
                Next

                xlibro.ActiveWorkbook.Save()
                guardarArchivo(ruta2, 3)
                MsgBox("Archivo  almacenado exitosamente en la ruta " & ruta)
            Else
                MsgBox("No se han generado liquidaciones correspondientes a la fecha")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox("Posiblemente hay un archivo existente con el nombre " & nombre & " y se encuentra abierto, o no fue posible sobreescribirlo")
        Finally
            xlibro.DisplayAlerts() = False
            xlibro.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlibro)
        End Try
    End Sub

    'Método para devolver el mes en texto
    Function mesLetras(mes As Integer) As String
        Select Case mes
            Case 1
                Return "enero"
            Case 2
                Return "febrero"
            Case 3
                Return "marzo"
            Case 4
                Return "abril"
            Case 5
                Return "mayo"
            Case 6
                Return "junio"
            Case 7
                Return "julio"
            Case 8
                Return "agosto"
            Case 9
                Return "septiembre"
            Case 10
                Return "octubre"
            Case 11
                Return "noviembre"
            Case 12
                Return "diciembre"
        End Select
    End Function

    'Método para calcular promedio neto de un empleado
    Function calcularPromNeto(documento As String, tipo As Integer)
        If tipo = 1 Then 'Sin deducciones
            Return Convert.ToDecimal((cnx.execSelect("select format(avg(Neto_pagar), 'F0') from Colilla_pago where Numero_documento = " & documento)))
        Else
            Return (Convert.ToDecimal(cnx.execSelect("select format(((avg(Neto_pagar)) - Aux_transporte - Salud - Pension), 'F0') 
                from Colilla_pago where Numero_documento = " & documento & " group by Aux_transporte, Salud, Pension")))
        End If
    End Function

    Function calcularDias(fechaIni As Date, fechaFin As Date)
        Dim tiempo As Integer
        tiempo = (fechaFin - fechaIni).Days
        If Not (tiempo = 0) Then
            Return tiempo
        End If
    End Function

    'Método para exportar datos a word
    Private Sub exportarWord(nombre As String, int As Integer)
        If int = 1 Then
            crearWordColilla(nombre)
        ElseIf int = 2 Then
            crearWordCuenta(nombre)
        Else
            crearWordPrestaciones(nombre)
        End If
    End Sub

    'Método para exportar datos a excel de colilla de pago
    Private Sub crearWordColilla(nombre As String)
        Try
            Dim rutaAct As String = Path.Combine(Directory.GetCurrentDirectory(), "colillaDePago.xlsx")
            Dim rutaAct2 As String = Path.Combine(Directory.GetCurrentDirectory(), "Documento.docx")
            Dim ruta As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\" & nombre & "ccc.xlsx"
            Dim ruta2 As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\" & nombre & ".docx"

            If File.Exists(ruta2) Then
                File.Delete(ruta2)
            End If

            My.Computer.FileSystem.CopyFile(rutaAct, ruta, True)
            My.Computer.FileSystem.CopyFile(rutaAct2, ruta2, True)

            xlibro = CreateObject("Excel.Application")
            xlibro.Workbooks.Open(ruta)
            xlibro.Visible = False

            Dim obword As Word.Application, wd As Word.Document
            obword = CreateObject("Word.Application")
            obword.DisplayAlerts = WdAlertLevel.wdAlertsNone
            obword.Visible = False
            wd = obword.Documents.Open(ruta2)
            obword.ActiveDocument.PageSetup.LeftMargin = 5
            obword.ActiveDocument.PageSetup.RightMargin = 52
            obword.ActiveDocument.PageSetup.PageWidth = 500
            obword.ActiveDocument.PageSetup.Orientation = WdOrientation.wdOrientLandscape

            Dim idcolilla As String
            Dim fecha As String

            Dim consulta As String = "
                if(select DATEPART(day, GETDATE())) <= 15
	                begin
		                select count(distinct(id)) from Colilla_pago where DATEPART(day, Fecha_inicial) <= 15 and 
                            DATEPART(month, Fecha_inicial) = DATEPART(month, GETDATE())
	                end
                else
	                begin
		                select count(distinct(id)) from Colilla_pago where DATEPART(day, Fecha_inicial) > 15 and 
                            DATEPART(month, Fecha_inicial) = DATEPART(month, GETDATE())
	                end"
            Dim datos As List(Of String) = cnx.execSelectVarios(consulta, False)

            If Not datos.Count() = 0 Then
                Dim cantidad As Integer = datos.ElementAt(0)
                If Convert.ToInt16(cantidad) > 1 Then
                    For i As Integer = 0 To cantidad - 2 Step 1
                        xlibro.Range("A" & (11 + i)).EntireRow.Insert()
                    Next
                End If

                consulta = "
                    if(select DATEPART(day, GETDATE())) < 15 --Primera quincena
	                    begin
		                    select id_colilla_empleado, Fecha_inicial, Fecha_final, (Nombre + ' ' + Apellido) 
                                    nombre, Numero_documento, Salario_base, DATEDIFF(day, Fecha_inicial, Fecha_final) dias, Basico, Aux_transporte, 
				                    Total_Hrs, Total_Rec, Incapacidades, Licencias, Otros_devengados, Total_devengado, Salud, Pension, Otras_deducciones,
                                    Retencion, Prestamos, Total_deducciones, Neto_pagar from Colilla_pago where DATEPART(day, Fecha_inicial) <= 15 and 
				                    DATEPART(month, Fecha_inicial) = DATEPART(month, GETDATE())
	                    end
                    else --Segunda quincena
	                    begin
		                    select id_colilla_empleado, Fecha_inicial, Fecha_final, (Nombre + ' ' + Apellido) 
                                    nombre, Numero_documento, Salario_base, DATEDIFF(day, Fecha_inicial, Fecha_final) dias, Basico, Aux_transporte, 
				                    Total_Hrs, Total_Rec, Incapacidades, Licencias, Otros_devengados, Total_devengado, Salud, Pension, Otras_deducciones,
                                    Retencion, Prestamos, Total_deducciones, Neto_pagar from Colilla_pago where DATEPART(day, Fecha_inicial) > 15 and 
				                    DATEPART(month, Fecha_inicial) = DATEPART(month, GETDATE())
	                    end"
                datos = cnx.execSelectVarios(consulta, False)

                Dim contC As Integer = 6
                Dim rangoA As Integer = 1
                Dim rangoE As Integer = 33
                Dim contador = 36
                Dim contD As Integer = 10
                Dim con As Integer = 0
                Dim ultimaFila As Integer = 10

                For i As Integer = 0 To datos.Count() - 1 Step 22
                    idcolilla = "No-" & datos.ElementAt(i)

                    xlibro.Sheets("NOMINA").Select()
                    If i = 0 Then
                        fecha = datos.ElementAt(i + 1) & " A " & datos.ElementAt(i + 2)
                        xlibro.Range("H4").Value = fecha
                    End If

                    xlibro.Range("B" & (10 + con)).Value = datos.ElementAt(i + 3)
                    xlibro.Range("C" & (10 + con)).Value = datos.ElementAt(i + 4)
                    xlibro.Range("D" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 5))
                    xlibro.Range("E" & (10 + con)).Value = Convert.ToInt16(datos.ElementAt(i + 6))
                    xlibro.Range("F" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 7))
                    xlibro.Range("G" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 8))
                    xlibro.Range("H" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 9))
                    xlibro.Range("I" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 10))
                    xlibro.Range("J" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 11))
                    xlibro.Range("K" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 12))
                    xlibro.Range("L" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 13))
                    xlibro.Range("M" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 14))
                    xlibro.Range("N" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 15))
                    xlibro.Range("O" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 16))
                    xlibro.Range("P" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 17))
                    xlibro.Range("Q" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 18))
                    xlibro.Range("R" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 19))
                    xlibro.Range("S" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 20))
                    xlibro.Range("T" & (10 + con)).Value = Convert.ToDecimal(datos.ElementAt(i + 21))
                    xlibro.Range("U" & (10 + con)).Value = "Efectivo"

                    xlibro.Sheets("RECIBOS").Select()
                    xlibro.Range("C" & contC).Value = idcolilla

                    If i > 0 Then
                        xlibro.Range("A1:E33").Copy()
                        xlibro.Range("A" & rangoA & ":E" & rangoE).PasteSpecial(XlPasteType.xlPasteAll)

                        xlibro.Range("E" & contC).Formula = "=+NOMINA!H4"
                        xlibro.Range("C" & (contC + 1)).Formula = "=+NOMINA!" & ("B" & (10 + con))
                        xlibro.Range("C" & (contC + 2)).Formula = "=+NOMINA!" & ("C" & (10 + con))
                        xlibro.Range("D" & contD).Formula = "=+NOMINA!" & ("E" & (10 + con))
                        xlibro.Range("D" & (contD + 1)).Formula = "=+NOMINA!" & ("F" & (10 + con))
                        xlibro.Range("D" & (contD + 2)).Formula = "=+NOMINA!" & ("G" & (10 + con))
                        xlibro.Range("D" & (contD + 3)).Formula = "=+NOMINA!" & ("H" & (10 + con))
                        xlibro.Range("D" & (contD + 4)).Formula = "=+NOMINA!" & ("I" & (10 + con))
                        xlibro.Range("D" & (contD + 5)).Formula = "=+NOMINA!" & ("J" & (10 + con))
                        xlibro.Range("D" & (contD + 6)).Formula = "=+NOMINA!" & ("K" & (10 + con))
                        xlibro.Range("D" & (contD + 7)).Formula = "=+NOMINA!" & ("L" & (10 + con))
                        xlibro.Range("D" & (contD + 10)).Formula = "=+NOMINA!" & ("N" & (10 + con))
                        xlibro.Range("D" & (contD + 11)).Formula = "=+NOMINA!" & ("O" & (10 + con))
                        xlibro.Range("D" & (contD + 12)).Formula = "=+NOMINA!" & ("P" & (10 + con))
                        xlibro.Range("D" & (contD + 13)).Formula = "=+NOMINA!" & ("Q" & (10 + con))
                        xlibro.Range("D" & (contD + 14)).Formula = "=+NOMINA!" & ("R" & (10 + con))
                    End If

                    contC += contador
                    contD += contador
                    rangoA += contador
                    rangoE += contador
                    con += 1
                    ultimaFila += 1
                Next
                xlibro.Sheets("NOMINA").Select()

                xlibro.Range("D" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("D10:D" & ultimaFila - 1))
                xlibro.Range("F" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("F10:F" & ultimaFila - 1))
                xlibro.Range("G" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("G10:G" & ultimaFila - 1))
                xlibro.Range("H" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("H10:H" & ultimaFila - 1))
                xlibro.Range("I" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("I10:I" & ultimaFila - 1))
                xlibro.Range("J" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("J10:J" & ultimaFila - 1))
                xlibro.Range("K" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("K10:K" & ultimaFila - 1))
                xlibro.Range("L" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("L10:L" & ultimaFila - 1))
                xlibro.Range("M" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("M10:M" & ultimaFila - 1))
                xlibro.Range("N" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("N10:N" & ultimaFila - 1))
                xlibro.Range("O" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("O10:O" & ultimaFila - 1))
                xlibro.Range("P" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("P10:P" & ultimaFila - 1))
                xlibro.Range("Q" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("Q10:Q" & ultimaFila - 1))
                xlibro.Range("R" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("R10:R" & ultimaFila - 1))
                xlibro.Range("S" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("S10:S" & ultimaFila - 1))
                xlibro.Range("T" & ultimaFila).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("T10:T" & ultimaFila - 1))

                xlibro.Range("B2:U" & ultimaFila).Copy()

                Dim ts As String = "[CAMPO]"
                obword.Selection.Move(6, -1)
                obword.Selection.Find.Execute(ts)
                If obword.Selection.Find.Found Then
                    obword.Selection.Paste()
                End If

                xlibro.Sheets("RECIBOS").Select()
                xlibro.Range("B2:E" & (rangoE - contador)).Copy()
                Dim st As String = "[RECIBO]"
                obword.Selection.Move(6, -1)
                obword.Selection.Find.Execute(st)
                If obword.Selection.Find.Found Then
                    obword.Selection.Paste()
                End If

                wd.Save()
                MsgBox("El archivo fue almacenado exitosamente en la ruta " & ruta2)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            xlibro.DisplayAlerts() = False
            xlibro.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlibro)
            If File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\" & nombre & "ccc.xlsx") Then
                File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\" & nombre & "ccc.xlsx")
            End If
        End Try
    End Sub

    'Método para exportar datos a excel de colilla de pago
    Private Sub crearWordPrestaciones(nombre As String)
        Try
            Dim rutaAct As String = Path.Combine(Directory.GetCurrentDirectory(), "Liquidaciones.xlsx")
            Dim rutaAct2 As String = Path.Combine(Directory.GetCurrentDirectory(), "Documento2.docx")
            Dim ruta As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\" & nombre & "ccc.xlsx"
            Dim ruta2 As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\" & nombre & ".docx"

            If File.Exists(ruta2) Then
                File.Delete(ruta2)
            End If

            My.Computer.FileSystem.CopyFile(rutaAct, ruta, True)
            My.Computer.FileSystem.CopyFile(rutaAct2, ruta2, True)

            xlibro = CreateObject("Excel.Application")
            xlibro.Workbooks.Open(ruta)
            xlibro.Visible = False

            Dim obword As Word.Application, wd As Word.Document
            obword = CreateObject("Word.Application")
            obword.DisplayAlerts = WdAlertLevel.wdAlertsNone
            obword.Visible = False
            wd = obword.Documents.Open(ruta2)
            obword.ActiveDocument.PageSetup.LeftMargin = 5
            obword.ActiveDocument.PageSetup.RightMargin = 5
            obword.ActiveDocument.PageSetup.PageWidth = 500
            obword.ActiveDocument.PageSetup.Orientation = WdOrientation.wdOrientLandscape

            Dim cont As Integer = 0
            Dim contA As Integer
            Dim contB As Integer
            Dim contC As Integer
            Dim contador As Integer = 47

            Dim sheet As String = "LIQUIDACION"

            'Liquidar por persona
            Dim consulta As String = "select Numero_documento from Liquidaciones where DATEPART(year, Fecha_final) = DATEPART(year, getdate())"
            Dim datos As List(Of String) = cnx.execSelectVarios(consulta, False)

            If Not datos.Count = 0 Then
                xlibro.Sheets(sheet).select()
                For i As Integer = 0 To datos.Count() - 1 Step 1
                    contA = 28
                    contB = 7
                    contC = 14
                    cont = 0

                    If i > 0 Then
                        sheet = "LIQUIDACION" & i.ToString()
                        xlibro.Sheets("LIQUIDACION").range("A1:C44").select()
                        xlibro.Sheets("LIQUIDACION").range("A1:C443").Copy()
                        xlibro.ActiveWorkbook.Sheets.Add().name = sheet
                        xlibro.Sheets(sheet).activate
                        xlibro.Sheets(sheet).range("A1:C44").select()
                        xlibro.Sheets(sheet).paste()
                        xlibro.Sheets(sheet).select()
                    End If

                    Dim encabezado As String
                    Dim lapso As String
                    Dim dias As String
                    Dim mes As String
                    Dim documento As String = datos.ElementAt(i)
                    Dim rangoA As Integer = 1
                    Dim rangoC As Integer = 44

                    'Prima
                    For d As Integer = 1 To 2 Step 1
                        If d = 1 Then
                            consulta = "select (h.Nombre + ' ' + h.Apellido)nombre, c.Cargo, c.Inicio_contrato, c.salario_base, 
                            l.Aux_transporte, l.Fecha_inicial, l.Fecha_final, convert(integer,l.Total_liquidado), l.Descripcion_prestacion
                            from Hoja_de_vida h inner join Contrato c on h.Numero_documento = c.Nume_documento 
                            inner join Liquidaciones l on h.Numero_documento = l.Numero_documento where 
                            DATEPART(MONTH, Fecha_final) <= 6 and h.Numero_documento = " & documento & " and 
                            (Descripcion_prestacion = 'Prima'or Descripcion_prestacion = 'Liquidar todos')"
                            encabezado = "LIQUIDACIÓN PRIMA 1ER SEMESTRE " & Today.Year
                            lapso = "07 de enero de " & Today.Year & " a 30 de junio de " & Today.Year
                            dias = "treinta (30)"
                            mes = "junio de " & Today.Year
                        Else
                            consulta = "select (h.Nombre + ' ' + h.Apellido)nombre, c.Cargo, c.Inicio_contrato, c.salario_base, 
                            l.Aux_transporte, l.Fecha_inicial, l.Fecha_final, convert(integer,l.Total_liquidado), l.Descripcion_prestacion
                            from Hoja_de_vida h inner join Contrato c on h.Numero_documento = c.Nume_documento 
                            inner join Liquidaciones l on h.Numero_documento = l.Numero_documento where 
                            DATEPART(MONTH, Fecha_final) > 6 and h.Numero_documento = " & documento & " and 
                            (Descripcion_prestacion = 'Prima' or Descripcion_prestacion = 'Liquidar todos')"
                            encabezado = "LIQUIDACIÓN PRIMA 2DO SEMESTRE " & Today.Year
                            lapso = "01 de julio de " & Today.Year & " a 15 de diciembre de " & Today.Year
                            dias = "quince (15)"
                            mes = "diciembre de " & Today.Year
                        End If
                        Dim dat As List(Of String) = cnx.execSelectVarios(consulta, False)

                        If Not dat.Count() = 0 Then
                            If d = 2 Then
                                xlibro.Range("A1:C44").Copy()
                                xlibro.Range("A" & rangoA & ":C" & rangoC).PasteSpecial(XlPasteType.xlPasteAll)
                            End If
                            xlibro.Range("A" & contB).Value = encabezado
                            xlibro.Range("B" & (contB + 1)).Value = dat.ElementAt(0)
                            xlibro.Range("B" & (contB + 2)).Value = documento
                            xlibro.Range("B" & (contB + 5)).Value = dat.ElementAt(1)
                            xlibro.Range("C" & contC).Value = dat.ElementAt(2)
                            xlibro.Range("C" & (contC + 1)).Value = Convert.ToSingle(dat.ElementAt(3))
                            xlibro.Range("C" & (contC + 2)).Value = Convert.ToSingle(dat.ElementAt(4))
                            xlibro.Range("C" & (contC + 3)).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("C" & (contC + 1) & ":C" & (contC + 2)))
                            xlibro.Range("A" & (contC + 4)).Value = "PRIMA SERVICIOS"
                            xlibro.Range("A" & (contC + 6)).Value = "DIA PRIMA SERVICIOS"
                            xlibro.Range("C" & (contC + 5)).Value = dat.ElementAt(5) & " A " & dat.ElementAt(6)
                            Dim cantDias As Integer = calcularDias(dat.ElementAt(5), dat.ElementAt(6))
                            xlibro.Range("C" & (contC + 6)).Value = "( " & cantDias & ")"
                            Dim total As Single = Convert.ToSingle(dat.ElementAt(7))
                            If dat.ElementAt(8) = "Liquidar todos" Then
                                total = (calcularPromNeto(documento, 1) * cantDias) / 360
                            End If
                            xlibro.Range("C" & (contC + 10)).Value = total
                            xlibro.Range("B" & (contB + 19)).Value = dat.ElementAt(0)
                            xlibro.Range("A" & contA).Value = EnLetras(Convert.ToInt64(total)) & " pesos (" & Convert.ToInt64(total) & ")"
                            xlibro.Range("B" & (contA + 1)).Value = "prima de servicios"
                            xlibro.Range("B" & (contA + 2)).Value = lapso
                            xlibro.Range("A" & (contA + 5)).Value = dias
                            xlibro.Range("C" & (contA + 5)).Value = mes
                            xlibro.Range("A" & (contA + 14)).Value = dat.ElementAt(0)
                            xlibro.Range("B" & (contA + 15)).Value = documento

                            contA += contador
                            contB += contador
                            contC += contador
                            rangoA += contador
                            rangoC += contador

                            xlibro.Range("A" & rangoA - contador & ":C" & rangoC - contador).Copy()

                            Dim ts As String = "[CAMPO]"
                            obword.Selection.Move(6, -1)
                            obword.Selection.Find.Execute(ts)
                            If obword.Selection.Find.Found Then
                                obword.Selection.Paste()
                                obword.Selection.Move(6, -1)
                                obword.Selection.Find.Execute(ts)
                            End If
                        End If
                    Next

                    'Cesantías
                    Dim contr As String = cnx.execSelect("select Tipo_Contrato from Contrato where Nume_documento = " & documento)

                    If Today.Month = 12 Or contr = "Indefinido" Then
                        consulta = "select (h.Nombre + ' ' + h.Apellido)nombre, c.Cargo, c.Inicio_contrato, c.salario_base, 
                            l.Aux_transporte, l.Fecha_inicial, l.Fecha_final, convert(integer,l.Total_liquidado), l.Descripcion_prestacion
							from Hoja_de_vida h inner join Contrato c on h.Numero_documento = c.Nume_documento 
                            inner join Liquidaciones l on h.Numero_documento = l.Numero_documento where 
                            h.Numero_documento = " & documento & " and (Descripcion_prestacion = 'Cesantías' 
                            or Descripcion_prestacion = 'Liquidar todos') or c.Tipo_Contrato = 'Indefinido'"
                        encabezado = "LIQUIDACIÓN CESANTÍAS " & mesLetras(Today.Month).ToUpper() & " " & Today.Year
                        mes = " de " & Today.Year
                        Dim dat As List(Of String) = cnx.execSelectVarios(consulta, False)

                        If Not dat.Count() = 0 Then
                            xlibro.Range("A1:C44").Copy()
                            xlibro.Range("A" & rangoA & ":C" & rangoC).PasteSpecial(XlPasteType.xlPasteAll)

                            xlibro.Range("A" & contB).Value = encabezado
                            xlibro.Range("B" & (contB + 1)).Value = dat.ElementAt(0)
                            xlibro.Range("B" & (contB + 2)).Value = documento
                            xlibro.Range("B" & (contB + 5)).Value = dat.ElementAt(1)
                            xlibro.Range("C" & contC).Value = dat.ElementAt(2)
                            xlibro.Range("C" & (contC + 1)).Value = Convert.ToSingle(dat.ElementAt(3))
                            xlibro.Range("C" & (contC + 2)).Value = Convert.ToSingle(dat.ElementAt(4))
                            xlibro.Range("C" & (contC + 3)).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("C" & (contC + 1) & ":C" & (contC + 2)))
                            xlibro.Range("A" & (contC + 4)).Value = "CESANTÍAS"
                            xlibro.Range("A" & (contC + 6)).Value = "DIA CESANTÍAS"
                            xlibro.Range("C" & (contC + 5)).Value = dat.ElementAt(5) & " A " & dat.ElementAt(6)
                            Dim cantDias As Integer = calcularDias(dat.ElementAt(5), dat.ElementAt(6))
                            xlibro.Range("C" & (contC + 6)).Value = "( " & cantDias & ")"
                            Dim total As Single = Convert.ToSingle(dat.ElementAt(7))
                            If dat.ElementAt(8) = "Liquidar todos" Then
                                total = (calcularPromNeto(documento, 1) * cantDias) / 360
                            End If
                            xlibro.Range("C" & (contC + 10)).Value = total
                            xlibro.Range("B" & (contB + 19)).Value = dat.ElementAt(0)
                            xlibro.Range("A" & contA).Value = EnLetras(Convert.ToInt64(total)) & " pesos (" & Convert.ToInt64(total) & ")"
                            xlibro.Range("B" & (contA + 1)).Value = "cesantías"
                            xlibro.Range("B" & (contA + 2)).Value = dat.ElementAt(5) & " A " & dat.ElementAt(6)
                            xlibro.Range("A" & (contA + 5)).Value = EnLetras(Today.Day) & " (" & Today.Day & ")"
                            xlibro.Range("C" & (contA + 5)).Value = mesLetras(Today.Month) & " " & mes
                            xlibro.Range("A" & (contA + 14)).Value = dat.ElementAt(0)
                            xlibro.Range("B" & (contA + 15)).Value = documento

                            contA += contador
                            contB += contador
                            contC += contador
                            rangoA += contador
                            rangoC += contador

                            xlibro.Range("A" & rangoA - contador & ":C" & rangoC - contador).Copy()

                            Dim ts As String = "[CAMPO]"
                            obword.Selection.Move(6, -1)
                            obword.Selection.Find.Execute(ts)
                            If obword.Selection.Find.Found Then
                                obword.Selection.Paste()
                                obword.Selection.Move(6, -1)
                                obword.Selection.Find.Execute(ts)
                            End If
                        End If

                        'Intereses a las cesantías
                        consulta = "select (h.Nombre + ' ' + h.Apellido)nombre, c.Cargo, c.Inicio_contrato, c.salario_base, 
                            l.Aux_transporte, l.Fecha_inicial, l.Fecha_final, convert(integer,l.Total_liquidado), l.Descripcion_prestacion
							from Hoja_de_vida h inner join Contrato c on h.Numero_documento = c.Nume_documento 
                            inner join Liquidaciones l on h.Numero_documento = l.Numero_documento where 
                            h.Numero_documento = " & documento & " and (Descripcion_prestacion = 'Intereses a las cesantías'
                            or Descripcion_prestacion = 'Liquidar todos') or c.Tipo_Contrato = 'Indefinido'"
                        encabezado = "LIQUIDACIÓN INTERESES A LAS CESANTÍAS " & mesLetras(Today.Month).ToUpper() & " " & Today.Year
                        mes = " de " & Today.Year

                        dat = cnx.execSelectVarios(consulta, False)
                        If Not dat.Count() = 0 Then
                            xlibro.Range("A1:C44").Copy()
                            xlibro.Range("A" & rangoA & ":C" & rangoC).PasteSpecial(XlPasteType.xlPasteAll)

                            xlibro.Range("A" & contB).Value = encabezado
                            xlibro.Range("B" & (contB + 1)).Value = dat.ElementAt(0)
                            xlibro.Range("B" & (contB + 2)).Value = documento
                            xlibro.Range("B" & (contB + 5)).Value = dat.ElementAt(1)
                            xlibro.Range("C" & contC).Value = dat.ElementAt(2)
                            xlibro.Range("C" & (contC + 1)).Value = Convert.ToSingle(dat.ElementAt(3))
                            xlibro.Range("C" & (contC + 2)).Value = Convert.ToSingle(dat.ElementAt(4))
                            xlibro.Range("C" & (contC + 3)).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("C" & (contC + 1) & ":C" & (contC + 2)))
                            xlibro.Range("A" & (contC + 4)).Value = "INTERESES EN LAS CESANTÍAS"
                            xlibro.Range("A" & (contC + 6)).Value = "DIA INT. EN LAS CESANTÍAS"
                            xlibro.Range("C" & (contC + 5)).Value = dat.ElementAt(5) & " A " & dat.ElementAt(6)
                            Dim cantDias As Integer = calcularDias(dat.ElementAt(5), dat.ElementAt(6))
                            xlibro.Range("C" & (contC + 6)).Value = "( " & cantDias & ")"
                            Dim total As Single = Convert.ToSingle(dat.ElementAt(7))
                            If dat.ElementAt(8) = "Liquidar todos" Then
                                Dim neto As Single = ((calcularPromNeto(documento, 1) * cantDias) / 360)
                                total = Math.Round(((neto * cantDias * 0.12) / 360), 2)
                            End If
                            xlibro.Range("C" & (contC + 10)).Value = total
                            xlibro.Range("B" & (contB + 19)).Value = dat.ElementAt(0)
                            xlibro.Range("A" & contA).Value = EnLetras(Convert.ToInt64(total)) & " pesos (" & Convert.ToInt64(total) & ")"
                            xlibro.Range("B" & (contA + 1)).Value = "intereses a las cesantías"
                            xlibro.Range("B" & (contA + 2)).Value = dat.ElementAt(5) & " A " & dat.ElementAt(6)
                            xlibro.Range("A" & (contA + 5)).Value = EnLetras(Today.Day) & " (" & Today.Day & ")"
                            xlibro.Range("C" & (contA + 5)).Value = mesLetras(Today.Month) & " " & mes
                            xlibro.Range("A" & (contA + 14)).Value = dat.ElementAt(0)
                            xlibro.Range("B" & (contA + 15)).Value = documento

                            contA += contador
                            contB += contador
                            contC += contador
                            rangoA += contador
                            rangoC += contador

                            xlibro.Range("A" & rangoA - contador & ":C" & rangoC - contador).Copy()

                            Dim ts As String = "[CAMPO]"
                            obword.Selection.Move(6, -1)
                            obword.Selection.Find.Execute(ts)
                            If obword.Selection.Find.Found Then
                                obword.Selection.Paste()
                                obword.Selection.Move(6, -1)
                                obword.Selection.Find.Execute(ts)
                            End If
                        End If
                    End If

                    'Vacaciones
                    If Today.Month = 12 Then
                        consulta = "select (h.Nombre + ' ' + h.Apellido)nombre, c.Cargo, c.Inicio_contrato, c.salario_base, 
                            l.Aux_transporte, l.Fecha_inicial, l.Fecha_final, convert(integer,l.Total_liquidado), l.Descripcion_prestacion
							from Hoja_de_vida h inner join Contrato c on h.Numero_documento = c.Nume_documento 
                            inner join Liquidaciones l on h.Numero_documento = l.Numero_documento where 
                            h.Numero_documento = " & documento & " and (Descripcion_prestacion = 'Vacaciones'
                            or Descripcion_prestacion = 'Liquidar todos')"
                        encabezado = "LIQUIDACIÓN VACACIONES " & Today.Year
                        dias = "quince (15)"
                        mes = "diciembre de " & Today.Year

                        Dim dat As List(Of String) = cnx.execSelectVarios(consulta, False)
                        If Not dat.Count() = 0 Then
                            xlibro.Range("A1:C44").Copy()
                            xlibro.Range("A" & rangoA & ":C" & rangoC).PasteSpecial(XlPasteType.xlPasteAll)

                            xlibro.Range("A" & contB).Value = encabezado
                            xlibro.Range("B" & (contB + 1)).Value = dat.ElementAt(0)
                            xlibro.Range("B" & (contB + 2)).Value = documento
                            xlibro.Range("B" & (contB + 5)).Value = dat.ElementAt(1)
                            xlibro.Range("C" & contC).Value = dat.ElementAt(2)
                            xlibro.Range("C" & (contC + 1)).Value = Convert.ToSingle(dat.ElementAt(3))
                            xlibro.Range("C" & (contC + 2)).Value = Convert.ToSingle(dat.ElementAt(4))
                            xlibro.Range("C" & (contC + 3)).Formula = xlibro.WorksheetFunction.Sum(xlibro.Range("C" & (contC + 1) & ":C" & (contC + 2)))
                            xlibro.Range("A" & (contC + 4)).Value = "VACACIONES"
                            xlibro.Range("A" & (contC + 6)).Value = "DIA VACACIONES"
                            xlibro.Range("C" & (contC + 5)).Value = dat.ElementAt(5) & " A " & dat.ElementAt(6)
                            Dim cantDias As Integer = calcularDias(dat.ElementAt(5), dat.ElementAt(6))
                            xlibro.Range("C" & (contC + 6)).Value = "( " & cantDias & ")"
                            Dim total As Single = Convert.ToSingle(dat.ElementAt(7))
                            If dat.ElementAt(8) = "Liquidar todos" Then
                                total = Math.Round(((calcularPromNeto(documento, 2) * cantDias) / 720), 2)
                            End If
                            xlibro.Range("C" & (contC + 10)).Value = total
                            xlibro.Range("B" & (contB + 19)).Value = dat.ElementAt(0)
                            xlibro.Range("A" & contA).Value = EnLetras(Convert.ToInt64(total)) & " pesos (" & Convert.ToInt64(total) & ")"
                            xlibro.Range("B" & (contA + 1)).Value = "vacaciones"
                            xlibro.Range("B" & (contA + 2)).Value = dat.ElementAt(5) & " A " & dat.ElementAt(6)
                            xlibro.Range("A" & (contA + 5)).Value = dias
                            xlibro.Range("C" & (contA + 5)).Value = mes
                            xlibro.Range("A" & (contA + 14)).Value = dat.ElementAt(0)
                            xlibro.Range("B" & (contA + 15)).Value = documento

                            contA += contador
                            contB += contador
                            contC += contador
                            rangoA += contador
                            rangoC += contador

                            xlibro.Range("A" & rangoA - contador & ":C" & rangoC - contador).Copy()

                            Dim ts As String = "[CAMPO]"
                            obword.Selection.Move(6, -1)
                            obword.Selection.Find.Execute(ts)
                            If obword.Selection.Find.Found Then
                                obword.Selection.Paste()
                                obword.Selection.Move(6, -1)
                                obword.Selection.Find.Execute(ts)
                            End If
                        End If
                    End If
                Next

                wd.Save()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            xlibro.DisplayAlerts() = False
            xlibro.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlibro)
            If File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\" & nombre & "ccc.xlsx") Then
                File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\" & nombre & "ccc.xlsx")
            End If
            MsgBox("El archivo fue almacenado exitosamente en la ruta " & Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\" & nombre & ".docx")
        End Try
    End Sub

    'Método para exportar datos a excel de cuenta de cobro
    Private Sub crearWordCuenta(nombre As String)
        Try
            Dim rutaAct As String = Path.Combine(Directory.GetCurrentDirectory(), "cuentaDeCobro.xlsx")
            Dim rutaAct2 As String = Path.Combine(Directory.GetCurrentDirectory(), "Documento2.docx")
            Dim ruta As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\" & nombre & "ccc.xlsx"
            Dim ruta2 As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\" & nombre & ".docx"

            If File.Exists(ruta2) Then
                File.Delete(ruta2)
            End If

            My.Computer.FileSystem.CopyFile(rutaAct, ruta, True)
            My.Computer.FileSystem.CopyFile(rutaAct2, ruta2, True)

            xlibro = CreateObject("Excel.Application")
            xlibro.Workbooks.Open(ruta)
            xlibro.Visible = False

            Dim obword As Word.Application, wd As Word.Document
            obword = CreateObject("Word.Application")
            obword.DisplayAlerts = WdAlertLevel.wdAlertsNone
            obword.Visible = False
            wd = obword.Documents.Open(ruta2)
            obword.ActiveDocument.PageSetup.LeftMargin = 5
            obword.ActiveDocument.PageSetup.RightMargin = 5
            obword.ActiveDocument.PageSetup.PageWidth = 500
            obword.ActiveDocument.PageSetup.Orientation = WdOrientation.wdOrientLandscape

            Dim fecha As String
            Dim con As Integer = 1
            Dim sheet As String = "CUENTA"

            Dim consulta As String = "select distinct(Numero_documento) from Cuenta_cobro where 
                                            DATEPART(month, Fecha_final)  = DATEPART(month, GETDATE()) "
            Dim datos As List(Of String) = cnx.execSelectVarios(consulta, False)

            If Not datos.Count() = 0 Then
                xlibro.Sheets(sheet).select()
                For i As Integer = 0 To datos.Count() - 1 Step 1
                    If i > 0 Then
                        sheet = "CUENTA" & i.ToString()
                        xlibro.Sheets("CUENTA").range("A1:F30").select()
                        xlibro.Sheets("CUENTA").range("A1:F30").Copy()
                        xlibro.ActiveWorkbook.Sheets.Add().name = sheet
                        xlibro.Sheets(sheet).activate
                        xlibro.Sheets(sheet).range("A1:F30").select()
                        xlibro.Sheets(sheet).paste()
                    End If

                    consulta = "select cc.Fecha_final,
                        (h.nombre + ' ' + h.apellido) nombre, h.Numero_documento, h.direccion, h.Telefono, 
                        d.nombre dpto, m.nombre ciudad, cc.Concepto_de, convert(integer,cc.Basico) from Cuenta_cobro cc inner join
                        Hoja_de_vida h on cc.Numero_documento = h.Numero_documento inner join Contrato co on 
                        co.Nume_documento = h.Numero_documento inner join Municipios m on h.Ciudad = m.id inner 
                        join Departamentos d on m.departamento_id = d.id where DATEPART(day, GETDATE()) <= co.tiempo_corte 
                        and h.Numero_documento = " & datos.ElementAt(i) & " and co.Fin_contrato >= GETDATE() 
                        and co.Inicio_contrato <= GETDATE()"
                    Dim dat As List(Of String) = cnx.execSelectVarios(consulta, False)
                    fecha = "Fecha: " & dat.ElementAt(0)

                    xlibro.Range("C3").Value = fecha
                    xlibro.Range("A8").Value = dat.ElementAt(1)
                    xlibro.Range("D8").Value = dat.ElementAt(2)
                    xlibro.Range("A10").Value = dat.ElementAt(3)
                    xlibro.Range("C10").Value = dat.ElementAt(4)
                    xlibro.Range("D10").Value = dat.ElementAt(5)
                    xlibro.Range("F10").Value = dat.ElementAt(6)
                    xlibro.Range("A13").Value = dat.ElementAt(7)
                    xlibro.Range("A18").Value = Convert.ToDecimal(dat.ElementAt(8))
                    xlibro.Range("A19").Value = EnLetras(dat.ElementAt(8)).ToUpper()

                    xlibro.Range("A1:F30").Copy()

                    Dim ts As String = "[CAMPO]"
                    obword.Selection.Move(6, -1)
                    obword.Selection.Find.Execute(ts)
                    If obword.Selection.Find.Found Then
                        obword.Selection.Paste()
                    End If

                    wd.Save()
                    MsgBox("El archivo fue almacenado exitosamente en la ruta " & ruta2)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            xlibro.DisplayAlerts() = False
            xlibro.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlibro)
            If File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\" & nombre & "ccc.xlsx") Then
                File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\" & nombre & "ccc.xlsx")
            End If
        End Try
    End Sub

    'Cancelar
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Hide()
        Nomina.Show()
        Nomina.TabPage1.Show()
    End Sub

    'Método para guardar archivos en la BD
    Private Sub guardarArchivo(ruta As String, iden As Integer)
        xlibro.ActiveWorkbook.ExportAsFixedFormat(0, ruta)
        Try
            Dim cmm As New SqlCommand("sp_mant_documentos", cnx.conexion)
            cmm.CommandType = CommandType.StoredProcedure

            cmm.Parameters.AddWithValue("@Documento", pasarABinario(ruta))
            cmm.Parameters.AddWithValue("@fecha", getFecha(iden))
            If iden = 1 Then
                cmm.Parameters.AddWithValue("@tipo_contrato", "Colilla de pago")
            ElseIf iden = 2 Then
                cmm.Parameters.AddWithValue("@tipo_contrato", "Cuenta de cobro")
            Else
                cmm.Parameters.AddWithValue("@tipo_contrato", "Prestaciones sociales")
            End If

            If cnx.ejecutarSP(cmm) Then
                'Se guardó documento en la bd
                If File.Exists(ruta) Then
                    File.Delete(ruta)
                End If
            End If
        Catch ex As Exception
            MsgBox("Hubo un error guardando el pdf")
        End Try
    End Sub

    'Método que retorna la fecha para pago
    Function getFecha(ident As Integer)
        If ident = 1 Then
            If Today.Day <= 15 Then
                Return Today.Year & "-" & Today.Month & "-15"
            Else
                Return Today.Year & "-" & Today.Month & "-30"
            End If
        Else
            Return Today.Year & "-" & Today.Month & "-30"
        End If
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

    'Metodo para pasar números a letras
    Public Function EnLetras(numero As String) As String
        Dim b, paso As Integer
        Dim expresion, entero, deci, flag As String

        flag = "N"
        For paso = 1 To Len(numero)
            If Mid(numero, paso, 1) = "." Then
                flag = "S"
            Else
                If flag = "N" Then
                    entero = entero + Mid(numero, paso, 1) 'Extae la parte entera del numero
                Else
                    deci = deci + Mid(numero, paso, 1) 'Extrae la parte decimal del numero
                End If
            End If
        Next paso

        If Len(deci) = 1 Then
            deci = deci & "0"
        End If

        flag = "N"
        If Val(numero) >= -999999999 And Val(numero) <= 999999999 Then 'si el numero esta dentro de 0 a 999.999.999
            For paso = Len(entero) To 1 Step -1
                b = Len(entero) - (paso - 1)
                Select Case paso
                    Case 3, 6, 9
                        Select Case Mid(entero, b, 1)
                            Case "1"
                                If Mid(entero, b + 1, 1) = "0" And Mid(entero, b + 2, 1) = "0" Then
                                    expresion = expresion & "cien "
                                Else
                                    expresion = expresion & "ciento "
                                End If
                            Case "2"
                                expresion = expresion & "doscientos "
                            Case "3"
                                expresion = expresion & "trescientos "
                            Case "4"
                                expresion = expresion & "cuatrocientos "
                            Case "5"
                                expresion = expresion & "quinientos "
                            Case "6"
                                expresion = expresion & "seiscientos "
                            Case "7"
                                expresion = expresion & "setecientos "
                            Case "8"
                                expresion = expresion & "ochocientos "
                            Case "9"
                                expresion = expresion & "novecientos "
                        End Select

                    Case 2, 5, 8
                        Select Case Mid(entero, b, 1)
                            Case "1"
                                If Mid(entero, b + 1, 1) = "0" Then
                                    flag = "S"
                                    expresion = expresion & "diez "
                                End If
                                If Mid(entero, b + 1, 1) = "1" Then
                                    flag = "S"
                                    expresion = expresion & "once "
                                End If
                                If Mid(entero, b + 1, 1) = "2" Then
                                    flag = "S"
                                    expresion = expresion & "doce "
                                End If
                                If Mid(entero, b + 1, 1) = "3" Then
                                    flag = "S"
                                    expresion = expresion & "trece "
                                End If
                                If Mid(entero, b + 1, 1) = "4" Then
                                    flag = "S"
                                    expresion = expresion & "catorce "
                                End If
                                If Mid(entero, b + 1, 1) = "5" Then
                                    flag = "S"
                                    expresion = expresion & "quince "
                                End If
                                If Mid(entero, b + 1, 1) > "5" Then
                                    flag = "N"
                                    expresion = expresion & "dieci"
                                End If

                            Case "2"
                                If Mid(entero, b + 1, 1) = "0" Then
                                    expresion = expresion & "veinte "
                                    flag = "S"
                                Else
                                    expresion = expresion & "veinti"
                                    flag = "N"
                                End If

                            Case "3"
                                If Mid(entero, b + 1, 1) = "0" Then
                                    expresion = expresion & "treinta "
                                    flag = "S"
                                Else
                                    expresion = expresion & "treinta y "
                                    flag = "N"
                                End If

                            Case "4"
                                If Mid(entero, b + 1, 1) = "0" Then
                                    expresion = expresion & "cuarenta "
                                    flag = "S"
                                Else
                                    expresion = expresion & "cuarenta y "
                                    flag = "N"
                                End If

                            Case "5"
                                If Mid(entero, b + 1, 1) = "0" Then
                                    expresion = expresion & "cincuenta "
                                    flag = "S"
                                Else
                                    expresion = expresion & "cincuenta y "
                                    flag = "N"
                                End If

                            Case "6"
                                If Mid(entero, b + 1, 1) = "0" Then
                                    expresion = expresion & "sesenta "
                                    flag = "S"
                                Else
                                    expresion = expresion & "sesenta y "
                                    flag = "N"
                                End If

                            Case "7"
                                If Mid(entero, b + 1, 1) = "0" Then
                                    expresion = expresion & "setenta "
                                    flag = "S"
                                Else
                                    expresion = expresion & "setenta y "
                                    flag = "N"
                                End If

                            Case "8"
                                If Mid(entero, b + 1, 1) = "0" Then
                                    expresion = expresion & "ochenta "
                                    flag = "S"
                                Else
                                    expresion = expresion & "ochenta y "
                                    flag = "N"
                                End If

                            Case "9"
                                If Mid(entero, b + 1, 1) = "0" Then
                                    expresion = expresion & "noventa "
                                    flag = "S"
                                Else
                                    expresion = expresion & "noventa y "
                                    flag = "N"
                                End If
                        End Select

                    Case 1, 4, 7
                        Select Case Mid(entero, b, 1)
                            Case "1"
                                If flag = "N" Then
                                    If paso = 1 Then
                                        expresion = expresion & "uno "
                                    Else
                                        expresion = expresion & "un "
                                    End If
                                End If
                            Case "2"
                                If flag = "N" Then
                                    expresion = expresion & "dos "
                                End If
                            Case "3"
                                If flag = "N" Then
                                    expresion = expresion & "tres "
                                End If
                            Case "4"
                                If flag = "N" Then
                                    expresion = expresion & "cuatro "
                                End If
                            Case "5"
                                If flag = "N" Then
                                    expresion = expresion & "cinco "
                                End If
                            Case "6"
                                If flag = "N" Then
                                    expresion = expresion & "seis "
                                End If
                            Case "7"
                                If flag = "N" Then
                                    expresion = expresion & "siete "
                                End If
                            Case "8"
                                If flag = "N" Then
                                    expresion = expresion & "ocho "
                                End If
                            Case "9"
                                If flag = "N" Then
                                    expresion = expresion & "nueve "
                                End If
                        End Select
                End Select

                If paso = 4 Then
                    If Mid(entero, 6, 1) <> "0" Or Mid(entero, 5, 1) <> "0" Or Mid(entero, 4, 1) <> "0" Or
                        (Mid(entero, 6, 1) = "0" And Mid(entero, 5, 1) = "0" And Mid(entero, 4, 1) = "0" And
                        Len(entero) <= 6) Then
                        expresion = expresion & "mil "
                    End If
                End If

                If paso = 7 Then
                    If Len(entero) = 7 And Mid(entero, 1, 1) = "1" Then
                        expresion = expresion & "millón "
                    Else
                        expresion = expresion & "millones "
                    End If
                End If
            Next paso

            If deci <> "" Then
                If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo
                    EnLetras = "menos " & expresion & "con " & deci ' & "/100"
                Else
                    EnLetras = expresion & "con " & deci ' & "/100"
                End If
            Else
                If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo
                    EnLetras = "menos " & expresion
                Else
                    EnLetras = expresion
                End If
            End If
        Else 'si el numero a convertir esta fuera del rango superior e inferior
            EnLetras = ""
        End If
    End Function

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles InicialDateExportation.ValueChanged

    End Sub
End Class