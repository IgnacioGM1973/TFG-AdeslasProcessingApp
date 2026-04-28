Imports ClosedXML.Excel
Imports System.IO
Imports System.Data
Imports System.Data.OleDb

Module modLimpiezaExcel

    Public Function LimpiarExcel(rutaOriginal As String) As String

        Try
            Dim extension As String = Path.GetExtension(rutaOriginal).ToLower()

            Dim rutaLimpia As String = Path.Combine(
                Path.GetDirectoryName(rutaOriginal),
                Path.GetFileNameWithoutExtension(rutaOriginal) & "_LIMPIO.xlsx"
            )

            ' =====================================================
            ' CASO 1: XLSX / XLSM → ClosedXML (rápido y estable)
            ' =====================================================
            If extension = ".xlsx" OrElse extension = ".xlsm" Then

                Using wbOrigen As New XLWorkbook(rutaOriginal)
                    Dim wsOrigen = wbOrigen.Worksheet(1)

                    Using wbNuevo As New XLWorkbook()
                        Dim wsNuevo = wbNuevo.Worksheets.Add("Hoja1")

                        Dim rango = wsOrigen.RangeUsed()

                        If rango IsNot Nothing Then
                            ' 🔥 versión rápida (sin bucles)
                            wsNuevo.Cell(1, 1).InsertData(rango.Rows().Select(Function(r) r.Cells().Select(Function(c) c.Value)))
                        End If

                        wbNuevo.SaveAs(rutaLimpia)
                    End Using
                End Using

                ' =====================================================
                ' CASO 2: XLSB → OleDb (única forma fiable)
                ' =====================================================
            ElseIf extension = ".xlsb" Then

                Dim dt As New DataTable()

                Dim connString As String =
                    "Provider=Microsoft.ACE.OLEDB.12.0;" &
                    "Data Source=" & rutaOriginal & ";" &
                    "Extended Properties='Excel 12.0;HDR=YES;IMEX=1'"

                Using conn As New OleDbConnection(connString)
                    conn.Open()

                    Dim schema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
                    Dim sheetName As String = schema.Rows(0)("TABLE_NAME").ToString()

                    Dim cmd As New OleDbCommand("SELECT * FROM [" & sheetName & "]", conn)
                    Dim da As New OleDbDataAdapter(cmd)
                    da.Fill(dt)
                End Using

                ' Guardar con ClosedXML
                Using wbNuevo As New XLWorkbook()
                    Dim wsNuevo = wbNuevo.Worksheets.Add("Hoja1")
                    wsNuevo.Cell(1, 1).InsertTable(dt)
                    wbNuevo.SaveAs(rutaLimpia)
                End Using

            Else
                MessageBox.Show("Formato no soportado: " & extension)
                Return rutaOriginal
            End If

            Return rutaLimpia

        Catch ex As Exception
            MessageBox.Show("ERROR al limpiar Excel: " & vbCrLf & ex.Message)
            Return rutaOriginal
        End Try

    End Function

End Module