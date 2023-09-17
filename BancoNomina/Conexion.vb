Imports System.Data.SqlClient

Public Class Conexion

#Region "Variables"
    'Creación de variables para la conexion a la BD
    Dim cadena As String = "Data Source=MYDESKTOP;Initial Catalog=BancoHV;Integrated Security=True"
    Public conexion As New SqlConnection(cadena)
#End Region

#Region "Métodos"

    'Método para ejecutar select que retorna solamente 1 valor
    Function execSelect(consulta As String)
        Try
            conexion.Open()
            Using cad As New SqlConnection(cadena)
                cad.Open()

                Dim cmd As New SqlCommand(consulta, cad)
                Dim dr As SqlDataReader = cmd.ExecuteReader

                If dr.Read() Then
                    Return dr(0).ToString()
                End If
            End Using
            Return 0
        Catch ex As Exception
            Return 0
        Finally
            conexion.Close()
        End Try
    End Function

    'Método para ejecutar select que retorna más de 1 valor
    Function execSelectVarios(consulta As String, esByte As Boolean)
        Try
            conexion.Open()
            Dim elementos
            'Array donde se almacenarán los valores y se retorna
            If esByte Then
                elementos = New List(Of Byte())
            Else
                elementos = New List(Of String)
            End If

            Using cn As New SqlConnection(cadena)
                cn.Open()

                Dim cmd As New SqlCommand(consulta, cn)
                Dim dr As SqlDataReader = cmd.ExecuteReader

                If dr.HasRows Then
                    While dr.Read()
                        For i As Integer = 0 To dr.FieldCount - 1 Step 1
                            elementos.Add(dr(i))
                        Next
                    End While
                End If
            End Using
            Return elementos
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        Finally
            conexion.Close()
        End Try
    End Function

    'Método para ejecutar un procedimiento almacenado
    Function ejecutarSP(cmm As SqlCommand)
        Try
            conexion.Open()
            Using cone As New SqlConnection(cadena)
                cone.Open()
                If (cmm.ExecuteNonQuery() > 0) Then
                    Return True
                Else
                    Return False
                End If
            End Using
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        Finally
            cmm.Dispose()
            conexion.Close()
        End Try
    End Function
#End Region

End Class