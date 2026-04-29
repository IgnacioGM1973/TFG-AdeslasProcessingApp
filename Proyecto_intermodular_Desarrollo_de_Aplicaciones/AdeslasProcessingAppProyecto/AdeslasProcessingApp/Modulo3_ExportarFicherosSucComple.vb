Imports System.Data.OleDb
Imports System.Text
Imports System.IO

Public Class Modulo3_ExportarFicherosSucComple
    'Modulo de Importación ficheros extras recibidos por parte de Adeslas Dept. Suscripción
    Public Shared Sub ImportarExcel_A_Access_MGA(ExcelFilePath As String, progressBar1 As ProgressBar)

        Try
            ' LIMPIAR EXCEL ANTES DE IMPORTAR
            ExcelFilePath = LimpiarExcel(ExcelFilePath)

            ' 🔥 Forzamos IMEX para que Excel no se coma registros
            Dim excelConnString As String =
        "Provider=Microsoft.ACE.OLEDB.16.0;" &
        "Data Source=" & ExcelFilePath & ";" &
        "Extended Properties='Excel 12.0 Xml;HDR=Yes;IMEX=1;TypeGuessRows=0;ImportMixedTypes=Text'"

            Dim leidos As Integer = 0
            Dim insertados As Integer = 0
            Dim fallidos As Integer = 0
            Dim totalRecords As Integer = 0

            Using excelConnection As New OleDb.OleDbConnection(excelConnString)
                excelConnection.Open()

                ' Obtener el nombre de la primera hoja
                Dim schemaTable As DataTable = excelConnection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Tables, Nothing)

                If schemaTable Is Nothing OrElse schemaTable.Rows.Count = 0 Then
                    MessageBox.Show("No se encontraron hojas en el archivo Excel.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If

                Dim sheetName As String = schemaTable.Rows(0)("TABLE_NAME").ToString()

                ' Contar registros del Excel
                Dim countQuery As String = $"SELECT COUNT(*) FROM [{sheetName}]"
                Dim countCommand As New OleDbCommand(countQuery, excelConnection)
                totalRecords = Convert.ToInt32(countCommand.ExecuteScalar())

                progressBar1.Minimum = 0
                progressBar1.Maximum = totalRecords
                progressBar1.Value = 0

                ' Leer datos
                Dim selectQuery As String = $"SELECT [POCENPOL], [POCECDCE], [DENTAL] FROM [{sheetName}]"
                Dim excelCommand As New OleDbCommand(selectQuery, excelConnection)

                Using reader As OleDbDataReader = excelCommand.ExecuteReader()

                    Using accessConnection As New OleDb.OleDbConnection(accessConnString)
                        accessConnection.Open()

                        Using transaction As OleDbTransaction = accessConnection.BeginTransaction()

                            ' Vaciar tabla MGA
                            Dim deleteQuery As String = "DELETE FROM MGA"
                            Using deleteCommand As New OleDbCommand(deleteQuery, accessConnection, transaction)
                                deleteCommand.ExecuteNonQuery()
                            End Using

                            ' Insertar en MGA
                            Dim insertQuery As String =
                        "INSERT INTO MGA (POCENPOL, POCECDCE, DENTAL, FECHA) " &
                        "VALUES (@POCENPOL, @POCECDCE, @DENTAL, @FECHA)"

                            Using accessCommand As New OleDbCommand(insertQuery, accessConnection, transaction)

                                accessCommand.Parameters.Add("@POCENPOL", OleDbType.VarWChar, 50)
                                accessCommand.Parameters.Add("@POCECDCE", OleDbType.VarWChar, 50)
                                accessCommand.Parameters.Add("@DENTAL", OleDbType.VarWChar, 10)
                                accessCommand.Parameters.Add("@FECHA", OleDbType.Date)

                                Dim recordCount As Integer = 0

                                While reader.Read()

                                    Dim pocenpol As String =
                                If(IsDBNull(reader("POCENPOL")), "", reader("POCENPOL").ToString().Trim())

                                    Dim pocecdce As String =
                                If(IsDBNull(reader("POCECDCE")), "", reader("POCECDCE").ToString().Trim())

                                    Dim dental As String =
                                If(IsDBNull(reader("DENTAL")), "", reader("DENTAL").ToString().Trim())

                                    leidos += 1

                                    If pocenpol <> "" AndAlso pocecdce <> "" Then

                                        accessCommand.Parameters("@POCENPOL").Value = pocenpol
                                        accessCommand.Parameters("@POCECDCE").Value = pocecdce
                                        accessCommand.Parameters("@DENTAL").Value = dental
                                        accessCommand.Parameters("@FECHA").Value = DateTime.Today

                                        Try
                                            accessCommand.ExecuteNonQuery()
                                            insertados += 1
                                        Catch ex As Exception
                                            fallidos += 1
                                        End Try

                                    Else
                                        fallidos += 1
                                    End If

                                    recordCount += 1

                                    If recordCount Mod 100 = 0 Then
                                        progressBar1.Value = Math.Min(recordCount, progressBar1.Maximum)
                                    End If

                                End While

                            End Using

                            transaction.Commit()

                        End Using
                    End Using
                End Using
            End Using

            MessageBox.Show(
        "IMPORTACIÓN FINALIZADA" & vbCrLf & vbCrLf &
        "Total en Excel: " & totalRecords & vbCrLf &
        "Leídos: " & leidos & vbCrLf &
        "Insertados: " & insertados & vbCrLf &
        "Fallidos: " & fallidos,
        "CONTROL REAL DE MGA",
        MessageBoxButtons.OK,
        MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error crítico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Public Shared Sub ImportarExcel_A_Access_SinAsistenciaViaje(ExcelFilePath As String, progressBar1 As ProgressBar)

        Try
            ' LIMPIAR EXCEL ANTES DE IMPORTAR
            ExcelFilePath = LimpiarExcel(ExcelFilePath)
            ' --- Conexión a Excel (forzando TEXTO y lectura completa) ---
            Dim excelConnString As String =
        "Provider=Microsoft.ACE.OLEDB.16.0;" &
        "Data Source=" & ExcelFilePath & ";" &
        "Extended Properties='Excel 12.0 Xml;HDR=Yes;IMEX=1;TypeGuessRows=0;ImportMixedTypes=Text'"
            Dim totalRecords As Integer
            Dim leidos As Integer = 0
            Dim insertados As Integer = 0
            Dim fallidos As Integer = 0



            Using excelConnection As New OleDb.OleDbConnection(excelConnString)
                excelConnection.Open()

                ' Obtener hojas del Excel
                Dim schemaTable As DataTable = excelConnection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Tables, Nothing)
                If schemaTable Is Nothing OrElse schemaTable.Rows.Count = 0 Then
                    MessageBox.Show("No se encontraron hojas en el archivo Excel.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If

                ' Primera hoja
                Dim sheetName As String = schemaTable.Rows(0)("TABLE_NAME").ToString()

                ' Contar registros
                Dim countQuery As String = $"SELECT COUNT(*) FROM [{sheetName}]"
                Dim countCommand As New OleDbCommand(countQuery, excelConnection)
                totalRecords = Convert.ToInt32(countCommand.ExecuteScalar())

                ' Configurar ProgressBar
                progressBar1.Minimum = 0
                progressBar1.Maximum = totalRecords
                progressBar1.Value = 0

                ' Leer datos
                Dim selectQuery As String = $"SELECT [POLINPOL], [POLIIDEX] FROM [{sheetName}]"
                Dim excelCommand As New OleDbCommand(selectQuery, excelConnection)
                Using reader As OleDbDataReader = excelCommand.ExecuteReader()

                    Using accessConnection As New OleDb.OleDbConnection(accessConnString)
                        accessConnection.Open()

                        Using transaction As OleDbTransaction = accessConnection.BeginTransaction()

                            ' Vaciar tabla
                            Dim deleteQuery As String = "DELETE FROM ColectivosSinAsistenciaViaje"
                            Using deleteCommand As New OleDbCommand(deleteQuery, accessConnection, transaction)
                                deleteCommand.ExecuteNonQuery()
                            End Using

                            ' Insert
                            Dim insertQuery As String =
                        "INSERT INTO ColectivosSinAsistenciaViaje (POLINPOL, POLIIDEX, FECHA) " &
                        "VALUES (@POLINPOL, @POLIIDEX, @FECHA)"

                            Using accessCommand As New OleDbCommand(insertQuery, accessConnection, transaction)

                                accessCommand.Parameters.Add("@POLINPOL", OleDbType.VarWChar, 50)
                                accessCommand.Parameters.Add("@POLIIDEX", OleDbType.VarWChar, 50)
                                accessCommand.Parameters.Add("@FECHA", OleDbType.Date)

                                Dim recordCount As Integer = 0

                                While reader.Read()

                                    Dim polinpol As String =
                                If(IsDBNull(reader("POLINPOL")), "", reader("POLINPOL").ToString().Trim())

                                    Dim poliidex As String =
                                If(IsDBNull(reader("POLIIDEX")), "", reader("POLIIDEX").ToString().Trim())

                                    leidos += 1

                                    ' Insertar solo si ambos campos tienen valor
                                    If polinpol <> "" AndAlso poliidex <> "" Then

                                        accessCommand.Parameters("@POLINPOL").Value = polinpol
                                        accessCommand.Parameters("@POLIIDEX").Value = poliidex
                                        accessCommand.Parameters("@FECHA").Value = DateTime.Today

                                        Try
                                            accessCommand.ExecuteNonQuery()
                                            insertados += 1
                                        Catch ex As Exception
                                            fallidos += 1
                                        End Try

                                    Else
                                        fallidos += 1
                                    End If

                                    recordCount += 1

                                    If recordCount Mod 100 = 0 Then
                                        progressBar1.Value = Math.Min(recordCount, progressBar1.Maximum)
                                    End If

                                End While

                            End Using

                            transaction.Commit()

                        End Using
                    End Using
                End Using
            End Using

            ' Resultado real
            MessageBox.Show(
        "IMPORTACIÓN FINALIZADA" & vbCrLf & vbCrLf &
        "Total en Excel: " & totalRecords & vbCrLf &
        "Leídos: " & leidos & vbCrLf &
        "Insertados: " & insertados & vbCrLf &
        "Fallidos: " & fallidos,
        "CONTROL REAL",
        MessageBoxButtons.OK,
        MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error crítico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub


    Public Shared Sub ImportarSenior(filePath As String)
        ' Conexión a la base de datos mediante la varable publica del modulo conexion accessConnString
        'Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & RutaPrincipal & "\bd\adeslas.accdb;"
        Using connection As New OleDb.OleDbConnection(accessConnString)
            connection.Open()

            ' Contador de registros importados
            Dim importedCount As Integer = 0

            ' Leer el archivo línea por línea
            Dim lines() As String = File.ReadAllLines(filePath, Encoding.Default)

            For Each line As String In lines
                ' Validar longitud mínima de línea
                If line.Length = 236 Then
                    ' Extraer datos según las posiciones
                    Dim TJASNTAR As String = line.Substring(0, 9).Trim()
                    Dim TJASCDDE As String = line.Substring(9, 3).Trim()
                    Dim TJASNPOL As String = line.Substring(12, 9).Trim()
                    Dim TJASCDCE As String = line.Substring(21, 9).Trim()
                    Dim TJASCDRE As String = line.Substring(30, 2).Trim()
                    Dim ASESOR As String = line.Substring(32, 41).Trim()
                    Dim DIR_ASESOR As String = line.Substring(73, 99).Trim()
                    Dim TEL_ASESOR As String = line.Substring(227, 9).Trim()
                    Dim CP_ASESOR As String = line.Substring(172, 35).Trim()
                    Dim POB_ASESOR As String = line.Substring(207, 20).Trim()

                    ' Verificar si el registro ya existe
                    Dim checkQuery As String = "SELECT COUNT(*) FROM AsesorSenior WHERE TJASNTAR = @TJASNTAR AND TJASCDDE = @TJASCDDE AND TJASNPOL = @TJASNPOL"
                    Using checkCommand As New OleDbCommand(checkQuery, connection)
                        checkCommand.Parameters.AddWithValue("@TJASNTAR", TJASNTAR)
                        checkCommand.Parameters.AddWithValue("@TJASCDDE", TJASCDDE)
                        checkCommand.Parameters.AddWithValue("@TJASNPOL", TJASNPOL)

                        Dim count As Integer = Convert.ToInt32(checkCommand.ExecuteScalar())

                        If count = 0 Then
                            ' Crear comando SQL de inserción
                            Dim insertQuery As String = "INSERT INTO AsesorSenior (TJASNTAR, TJASCDDE, TJASNPOL, TJASCDCE, TJASCDRE, ASESOR, DIR_ASESOR, TEL_ASESOR, CP_ASESOR, POB_ASESOR) " &
                                                        "VALUES (@TJASNTAR, @TJASCDDE, @TJASNPOL, @TJASCDCE, @TJASCDRE, @ASESOR, @DIR_ASESOR, @TEL_ASESOR, @CP_ASESOR, @POB_ASESOR)"

                            Using insertCommand As New OleDbCommand(insertQuery, connection)
                                ' Agregar parámetros
                                insertCommand.Parameters.AddWithValue("@TJASNTAR", TJASNTAR)
                                insertCommand.Parameters.AddWithValue("@TJASCDDE", TJASCDDE)
                                insertCommand.Parameters.AddWithValue("@TJASNPOL", TJASNPOL)
                                insertCommand.Parameters.AddWithValue("@TJASCDCE", TJASCDCE)
                                insertCommand.Parameters.AddWithValue("@TJASCDRE", TJASCDRE)
                                insertCommand.Parameters.AddWithValue("@ASESOR", ASESOR)
                                insertCommand.Parameters.AddWithValue("@DIR_ASESOR", DIR_ASESOR)
                                insertCommand.Parameters.AddWithValue("@TEL_ASESOR", TEL_ASESOR)
                                insertCommand.Parameters.AddWithValue("@CP_ASESOR", CP_ASESOR)
                                insertCommand.Parameters.AddWithValue("@POB_ASESOR", POB_ASESOR)

                                ' Ejecutar inserción
                                insertCommand.ExecuteNonQuery()
                                importedCount += 1 ' Incrementar contador de registros importados
                            End Using
                        Else
                            ' Mensaje de registro duplicado (opcional)
                            Debug.WriteLine($"Registro duplicado: TJASNTAR={TJASNTAR}, TJASCDDE={TJASCDDE}, TJASNPOL={TJASNPOL}")
                        End If
                    End Using
                Else
                    MessageBox.Show($"Línea ignorada por ser demasiado corta: {line}", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            Next

            ' Mostrar mensaje final
            If importedCount > 0 Then
                MessageBox.Show($"{importedCount} registro(s) importado(s) exitosamente.", "Importación completada", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("No hay registros nuevos para importar.", "Importación completada", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End Using
    End Sub
    Public Shared Sub ImportarExcel_A_Access_ColectivosConDireccionEnvio(ExcelFilePath As String, progressBar1 As ProgressBar)
        Try
            ' LIMPIAR EXCEL ANTES DE IMPORTAR
            ExcelFilePath = LimpiarExcel(ExcelFilePath)
            ' Conexión al archivo Excel
            Dim excelConnString As String = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & ExcelFilePath & ";Extended Properties='Excel 12.0 Xml;HDR=Yes;'"
            Dim errorCampo As Boolean = True

            Dim totalRecords As Integer
            Dim leidos As Integer = 0
            Dim insertados As Integer = 0
            Dim fallidos As Integer = 0
            Using excelConnection As New OleDb.OleDbConnection(excelConnString)
                excelConnection.Open()




                ' Obtener el nombre de la primera hoja
                Dim dtSchema As DataTable = excelConnection.GetSchema("Tables")
                If dtSchema.Rows.Count = 0 Then
                    MessageBox.Show("El archivo de Excel no contiene hojas.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If

                Dim firstSheetName As String = dtSchema.Rows(0)("TABLE_NAME").ToString()

                ' Contar el número total de registros en la hoja
                Dim countQuery As String = $"SELECT COUNT(*) FROM [{firstSheetName}]"
                Dim countCommand As New OleDbCommand(countQuery, excelConnection)
                totalRecords = Convert.ToInt32(countCommand.ExecuteScalar())

                ' Configurar la ProgressBar
                progressBar1.Minimum = 0
                progressBar1.Maximum = totalRecords
                progressBar1.Value = 0

                ' Leer los datos de la primera hoja de Excel
                Dim selectQuery As String = $"SELECT * FROM [{firstSheetName}]"
                Dim excelCommand As New OleDbCommand(selectQuery, excelConnection)
                Dim reader As OleDbDataReader = excelCommand.ExecuteReader()

                ' Conexión a la base de datos mediante la varable publica del modulo conexion accessConnString
                Using accessConnection As New OleDb.OleDbConnection(accessConnString)
                    accessConnection.Open()

                    ' Iniciar transacción
                    Using transaction As OleDbTransaction = accessConnection.BeginTransaction()

                        ' Eliminar registros existentes en la tabla
                        Dim deleteQuery As String = "DELETE FROM ColectivosConDireccionEnvio"
                        Using deleteCommand As New OleDbCommand(deleteQuery, accessConnection, transaction)
                            deleteCommand.ExecuteNonQuery()
                        End Using

                        ' Insertar datos en Access
                        Dim insertQuery As String = "
                    INSERT INTO ColectivosConDireccionEnvio 
                    (N_COLECTIVO, N_POLIZA, DESCRIPCION, DELEGACION, FECHA_EFECTO, REPONSABLE_COLECTIVO, TIPO_ENVIO, NOMBRE, PERSONA_CONTACTO, DIRECCION, CD_POSTAL, POBLACION, PROVINCIA, OBSERVACIONES, FECHA_MODIFICACION, F16, F17, FECHA) 
                    VALUES 
                    (@N_COLECTIVO, @N_POLIZA, @DESCRIPCION, @DELEGACION, @FECHA_EFECTO, @REPONSABLE_COLECTIVO, @TIPO_ENVIO, @NOMBRE, @PERSONA_CONTACTO, @DIRECCION, @CD_POSTAL, @POBLACION, @PROVINCIA, @OBSERVACIONES, @FECHA_MODIFICACION, @F16, @F17, @FECHA)"

                        Using accessCommand As New OleDbCommand(insertQuery, accessConnection, transaction)
                            ' Definir parámetros
                            accessCommand.Parameters.Add("@N_COLECTIVO", OleDbType.VarWChar)
                            accessCommand.Parameters.Add("@N_POLIZA", OleDbType.VarWChar)
                            accessCommand.Parameters.Add("@DESCRIPCION", OleDbType.VarWChar)
                            accessCommand.Parameters.Add("@DELEGACION", OleDbType.VarWChar)
                            accessCommand.Parameters.Add("@FECHA_EFECTO", OleDbType.VarWChar)
                            accessCommand.Parameters.Add("@REPONSABLE_COLECTIVO", OleDbType.VarWChar)
                            accessCommand.Parameters.Add("@TIPO_ENVIO", OleDbType.VarWChar)
                            accessCommand.Parameters.Add("@NOMBRE", OleDbType.VarWChar)
                            accessCommand.Parameters.Add("@PERSONA_CONTACTO", OleDbType.VarWChar)
                            accessCommand.Parameters.Add("@DIRECCION", OleDbType.VarWChar)
                            accessCommand.Parameters.Add("@CD_POSTAL", OleDbType.VarWChar)
                            accessCommand.Parameters.Add("@POBLACION", OleDbType.VarWChar)
                            accessCommand.Parameters.Add("@PROVINCIA", OleDbType.VarWChar)
                            accessCommand.Parameters.Add("@OBSERVACIONES", OleDbType.VarWChar)
                            accessCommand.Parameters.Add("@FECHA_MODIFICACION", OleDbType.VarWChar)
                            accessCommand.Parameters.Add("@F16", OleDbType.VarWChar)
                            accessCommand.Parameters.Add("@F17", OleDbType.VarWChar)
                            accessCommand.Parameters.Add("@FECHA", OleDbType.Date)

                            Dim recordCount As Integer = 0
                            While reader.Read()
                                Try
                                    accessCommand.Parameters("@N_COLECTIVO").Value = reader("Nº Colectivo").ToString()
                                    accessCommand.Parameters("@N_POLIZA").Value = reader("Nº Póliza").ToString()
                                    accessCommand.Parameters("@DESCRIPCION").Value = reader("Descripción").ToString()
                                    accessCommand.Parameters("@DELEGACION").Value = reader("Delegación").ToString()
                                    accessCommand.Parameters("@FECHA_EFECTO").Value = reader("Fecha Efecto").ToString().Split(" "c)(0).Trim()
                                    accessCommand.Parameters("@REPONSABLE_COLECTIVO").Value = reader("RESPONSABLE COLECTIVO").ToString()
                                    accessCommand.Parameters("@TIPO_ENVIO").Value = reader("TIPO ENVIO").ToString()
                                    accessCommand.Parameters("@NOMBRE").Value = reader("NOMBRE ").ToString()
                                    accessCommand.Parameters("@PERSONA_CONTACTO").Value = reader("PERSONA DE CONTACTO").ToString()
                                    accessCommand.Parameters("@DIRECCION").Value = reader("DIRECCION").ToString()
                                    accessCommand.Parameters("@CD_POSTAL").Value = reader("CODIGO POSTAL").ToString()
                                    accessCommand.Parameters("@POBLACION").Value = reader("POBLACION").ToString()
                                    accessCommand.Parameters("@PROVINCIA").Value = reader("PROVINCIA").ToString()
                                    accessCommand.Parameters("@OBSERVACIONES").Value = reader("OBSERVACIONES").ToString()
                                    accessCommand.Parameters("@FECHA_MODIFICACION").Value = reader("FECHA_MODIFICACION")
                                    accessCommand.Parameters("@F16").Value = reader("F16").ToString()
                                    accessCommand.Parameters("@F17").Value = reader("F17").ToString()
                                    accessCommand.Parameters("@FECHA").Value = DateTime.Today

                                    accessCommand.ExecuteNonQuery()

                                    ' Actualizar progreso

                                    recordCount += 1
                                    If recordCount Mod 100 = 0 Then
                                        progressBar1.Value = Math.Min(recordCount, progressBar1.Maximum)
                                    End If
                                    insertados += 1
                                Catch ex As Exception
                                    fallidos += 1
                                    MessageBox.Show($"Error: Campo Excel {ex.Message} incorrecto.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                    errorCampo = False
                                    'Exit While
                                End Try
                                leidos += 1
                            End While
                        End Using
                        transaction.Commit()
                    End Using
                End Using
            End Using

            '


            If errorCampo = True Then
                'Resultado real
                MessageBox.Show(
                    "IMPORTACIÓN FINALIZADA" & vbCrLf & vbCrLf &
                     "Total en Excel: " & totalRecords & vbCrLf &
                     "Leídos: " & leidos & vbCrLf &
                    "Insertados: " & insertados & vbCrLf &
                    "Fallidos: " & fallidos,
                    "CONTROL REAL",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)
            Else
                MessageBox.Show("Importación Fallida.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Shared Sub ImportarExcel_A_Access_TARJETAS_DENTALES_PYMES(ExcelFilePath As String, progressBar1 As ProgressBar)

        Try
            ' LIMPIAR EXCEL ANTES DE IMPORTAR
            ExcelFilePath = LimpiarExcel(ExcelFilePath)

            ' Conexión al archivo Excel
            Dim excelConnString As String =
            "Provider=Microsoft.ACE.OLEDB.16.0;" &
            "Data Source=" & ExcelFilePath & ";" &
            "Extended Properties='Excel 12.0 Xml;HDR=Yes;IMEX=1;TypeGuessRows=0;ImportMixedTypes=Text'"

            Dim errorCampo As Boolean = True

            Dim totalRecords As Integer = 0
            Dim leidos As Integer = 0
            Dim insertados As Integer = 0
            Dim fallidos As Integer = 0

            Using excelConnection As New OleDb.OleDbConnection(excelConnString)
                excelConnection.Open()

                ' Obtener hoja
                Dim dtSchema As DataTable = excelConnection.GetSchema("Tables")
                If dtSchema.Rows.Count = 0 Then
                    MessageBox.Show("El archivo de Excel no contiene hojas.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If

                Dim firstSheetName As String = dtSchema.Rows(0)("TABLE_NAME").ToString()

                ' Contar registros
                Dim countQuery As String = $"SELECT COUNT(*) FROM [{firstSheetName}]"
                Using countCommand As New OleDbCommand(countQuery, excelConnection)
                    totalRecords = Convert.ToInt32(countCommand.ExecuteScalar())
                End Using

                ' Configurar la ProgressBar
                progressBar1.Minimum = 0
                progressBar1.Maximum = totalRecords
                progressBar1.Value = 0

                ' Leer datos
                Dim selectQuery As String = $"SELECT * FROM [{firstSheetName}]"
                Using excelCommand As New OleDbCommand(selectQuery, excelConnection)
                    Using reader As OleDbDataReader = excelCommand.ExecuteReader()

                        Using accessConnection As New OleDb.OleDbConnection(accessConnString)
                            accessConnection.Open()

                            Using transaction As OleDbTransaction = accessConnection.BeginTransaction()

                                ' Vaciar tabla
                                Dim deleteQuery As String = "DELETE FROM TarjetasDentalesPymes"
                                Using deleteCommand As New OleDbCommand(deleteQuery, accessConnection, transaction)
                                    deleteCommand.ExecuteNonQuery()
                                End Using

                                ' Insert
                                Dim insertQuery As String =
                        "INSERT INTO TarjetasDentalesPymes " &
                        "(POLICDDE,POLINPOL,POLIFECA,POLIFECB,POLICDPT,PRODDSPT,POLIAGTA," &
                        "CLIEDOMI,CLIECDPS,TPOBPOBL,POCLCDPA,CLIENOMB,CLIEAPEL,CLIENIF,FECHA) " &
                        "VALUES (@POLICDDE,@POLINPOL,@POLIFECA,@POLIFECB,@POLICDPT,@PRODDSPT,@POLIAGTA," &
                        "@CLIEDOMI,@CLIECDPS,@TPOBPOBL,@POCLCDPA,@CLIENOMB,@CLIEAPEL,@CLIENIF,@FECHA)"

                                Using accessCommand As New OleDbCommand(insertQuery, accessConnection, transaction)

                                    ' Definir parámetros
                                    accessCommand.Parameters.Add("@POLICDDE", OleDbType.VarWChar)
                                    accessCommand.Parameters.Add("@POLINPOL", OleDbType.VarWChar)
                                    accessCommand.Parameters.Add("@POLIFECA", OleDbType.VarWChar)
                                    accessCommand.Parameters.Add("@POLIFECB", OleDbType.VarWChar)
                                    accessCommand.Parameters.Add("@POLICDPT", OleDbType.VarWChar)
                                    accessCommand.Parameters.Add("@PRODDSPT", OleDbType.VarWChar)
                                    accessCommand.Parameters.Add("@POLIAGTA", OleDbType.VarWChar)
                                    accessCommand.Parameters.Add("@CLIEDOMI", OleDbType.VarWChar)
                                    accessCommand.Parameters.Add("@CLIECDPS", OleDbType.VarWChar)
                                    accessCommand.Parameters.Add("@TPOBPOBL", OleDbType.VarWChar)
                                    accessCommand.Parameters.Add("@POCLCDPA", OleDbType.VarWChar)
                                    accessCommand.Parameters.Add("@CLIENOMB", OleDbType.VarWChar)
                                    accessCommand.Parameters.Add("@CLIEAPEL", OleDbType.VarWChar)
                                    accessCommand.Parameters.Add("@CLIENIF", OleDbType.VarWChar)
                                    accessCommand.Parameters.Add("@FECHA", OleDbType.Date)

                                    Dim recordCount As Integer = 0

                                    While reader.Read()

                                        leidos += 1  ' ✅ suma bien ahora

                                        Try
                                            accessCommand.Parameters("@POLICDDE").Value = reader("POLICDDE").ToString()
                                            accessCommand.Parameters("@POLINPOL").Value = reader("POLINPOL").ToString()
                                            accessCommand.Parameters("@POLIFECA").Value = reader("POLIFECA").ToString()
                                            accessCommand.Parameters("@POLIFECB").Value = reader("POLIFECB").ToString()
                                            accessCommand.Parameters("@POLICDPT").Value = reader("POLICDPT").ToString()
                                            accessCommand.Parameters("@PRODDSPT").Value = reader("PRODDSPT").ToString()
                                            accessCommand.Parameters("@POLIAGTA").Value = reader("POLIAGTA").ToString()
                                            accessCommand.Parameters("@CLIEDOMI").Value = reader("CLIEDOMI").ToString()
                                            accessCommand.Parameters("@CLIECDPS").Value = reader("CLIECDPS").ToString()
                                            accessCommand.Parameters("@TPOBPOBL").Value = reader("TPOBPOBL").ToString()
                                            accessCommand.Parameters("@POCLCDPA").Value = reader("POCLCDPA").ToString()
                                            accessCommand.Parameters("@CLIENOMB").Value = reader("CLIENOMB").ToString()
                                            accessCommand.Parameters("@CLIEAPEL").Value = reader("CLIEAPEL").ToString()
                                            accessCommand.Parameters("@CLIENIF").Value = reader("CLIENIF").ToString()
                                            accessCommand.Parameters("@FECHA").Value = DateTime.Today

                                            accessCommand.ExecuteNonQuery()

                                            insertados += 1
                                            recordCount += 1

                                            If recordCount Mod 100 = 0 Then
                                                progressBar1.Value = Math.Min(recordCount, progressBar1.Maximum)
                                            End If

                                        Catch ex As Exception
                                            fallidos += 1
                                            errorCampo = False
                                            Debug.WriteLine("Error fila " & leidos & ": " & ex.Message)
                                        End Try

                                    End While

                                End Using

                                transaction.Commit()

                            End Using
                        End Using
                    End Using
                End Using

            End Using

            ' Asegurar barra completa
            progressBar1.Value = progressBar1.Maximum

            ' Resultado
            MessageBox.Show(
            "IMPORTACIÓN FINALIZADA" & vbCrLf & vbCrLf &
             "Total en Excel: " & totalRecords & vbCrLf &
             "Leídos: " & leidos & vbCrLf &
             "Insertados: " & insertados & vbCrLf &
             "Fallidos: " & fallidos,
            "CONTROL REAL",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Public Shared Sub ImportarTarjetasNoProcesar(ExcelFilePath As String, progressBar1 As ProgressBar, lblInfo As Label)

        Dim errorCampo As Boolean = True

        Dim totalRecords As Integer = 0
        Dim leidos As Integer = 0
        Dim insertados As Integer = 0
        Dim fallidos As Integer = 0

        Try
            ' LIMPIAR EXCEL ANTES DE IMPORTAR
            ExcelFilePath = LimpiarExcel(ExcelFilePath)

            ' Conexión al archivo Excel
            Dim excelConnectionString As String =
            "Provider=Microsoft.ACE.OLEDB.16.0;" &
            "Data Source=" & ExcelFilePath & ";" &
            "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;TypeGuessRows=0;ImportMixedTypes=Text'"

            Using excelConnection As New OleDb.OleDbConnection(excelConnectionString)
                excelConnection.Open()

                ' Obtener el nombre de la primera hoja activa
                Dim dtSheets As DataTable = excelConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)

                If dtSheets Is Nothing OrElse dtSheets.Rows.Count = 0 Then
                    MsgBox("No se encontraron hojas en el archivo Excel.")
                    Exit Sub
                End If

                Dim firstSheetName As String = dtSheets.Rows(0)("TABLE_NAME").ToString()

                ' Contar total de registros en Excel
                Dim countQuery As String = $"SELECT COUNT(*) FROM [{firstSheetName}]"
                Using countCommand As New OleDbCommand(countQuery, excelConnection)
                    totalRecords = Convert.ToInt32(countCommand.ExecuteScalar())
                End Using

                ' Configurar ProgressBar
                progressBar1.Minimum = 0
                progressBar1.Maximum = totalRecords
                progressBar1.Value = 0

                ' Leer registros de Excel
                Dim query As String = $"SELECT * FROM [{firstSheetName}]"

                Using command As New OleDbCommand(query, excelConnection)
                    Using reader As OleDbDataReader = command.ExecuteReader()

                        Using connection As New OleDb.OleDbConnection(accessConnString)
                            connection.Open()

                            ' ===================================================
                            ' COMPROBAR SI YA HAY REGISTROS DEL DÍA DE HOY
                            ' ===================================================
                            Dim existeHoy As Boolean = False
                            Dim checkQuery As String = "SELECT COUNT(*) FROM TarjetasNoProcesar WHERE FECHA = Date()"

                            Using checkCommand As New OleDbCommand(checkQuery, connection)
                                Dim totalHoy As Integer = Convert.ToInt32(checkCommand.ExecuteScalar())
                                If totalHoy > 0 Then
                                    existeHoy = True
                                End If
                            End Using

                            ' ===================================================
                            ' BORRAR SOLO SI NO HAY REGISTROS DE HOY
                            ' ===================================================
                            If Not existeHoy Then
                                Dim deleteQuery As String = "DELETE FROM TarjetasNoProcesar"
                                Using deleteCommand As New OleDbCommand(deleteQuery, connection)
                                    deleteCommand.ExecuteNonQuery()
                                End Using

                                lblInfo.Text = "✅ Datos antiguos borrados (nuevo día)"
                                lblInfo.Refresh()
                            Else
                                lblInfo.Text = "ℹ Ya existen registros de hoy. No se borran"
                                lblInfo.Refresh()
                            End If

                            ' PREPARAR INSERT UNA SOLA VEZ (MUCHO MÁS RÁPIDO)
                            ' ✅ Añadido el campo nuor
                            Dim insertQuery As String =
                            "INSERT INTO TarjetasNoProcesar " &
                            "(DELEGACION, POLIZA, CERTIFICADO, nuor, f_baja, f_informe, fichero, FECHA) " &
                            "VALUES (?, ?, ?, ?, ?, ?, ?, ?)"

                            Using insertCommand As New OleDbCommand(insertQuery, connection)

                                Dim recordCount As Integer = 0

                                While reader.Read()

                                    leidos += 1

                                    Try
                                        Dim delegacion As String = reader("DELEGACION").ToString().Trim()
                                        Dim poliza As String = reader("POLIZA").ToString().Trim()
                                        Dim certificado As String = reader("CERTIFICADO").ToString().Trim()

                                        ' ✅ NUEVO: leer nuor (si no existe en algún Excel, no rompe)
                                        Dim nuor As String = ""
                                        Try
                                            nuor = reader("nuor").ToString().Trim()
                                        Catch
                                            nuor = ""
                                        End Try

                                        Dim fBaja As String = reader("f_baja").ToString().Trim()
                                        Dim fInforme As String = reader("f_informe").ToString().Trim()
                                        Dim fichero As String = reader("fichero").ToString().Trim()

                                        ' Validación para omitir registros completamente vacíos
                                        If String.IsNullOrWhiteSpace(delegacion) AndAlso
                                       String.IsNullOrWhiteSpace(poliza) AndAlso
                                       String.IsNullOrWhiteSpace(certificado) AndAlso
                                       String.IsNullOrWhiteSpace(nuor) AndAlso
                                       String.IsNullOrWhiteSpace(fBaja) AndAlso
                                       String.IsNullOrWhiteSpace(fInforme) AndAlso
                                       String.IsNullOrWhiteSpace(fichero) Then

                                            fallidos += 1
                                            Continue While
                                        End If

                                        insertCommand.Parameters.Clear()
                                        insertCommand.Parameters.AddWithValue("?", delegacion)
                                        insertCommand.Parameters.AddWithValue("?", poliza)
                                        insertCommand.Parameters.AddWithValue("?", certificado)
                                        insertCommand.Parameters.AddWithValue("?", nuor) ' ✅ NUEVO
                                        insertCommand.Parameters.AddWithValue("?", fBaja)
                                        insertCommand.Parameters.AddWithValue("?", fInforme)
                                        insertCommand.Parameters.AddWithValue("?", fichero)
                                        insertCommand.Parameters.AddWithValue("?", DateTime.Today)

                                        insertCommand.ExecuteNonQuery()
                                        insertados += 1

                                        ' Actualizar progreso
                                        recordCount += 1
                                        If recordCount Mod 100 = 0 Then
                                            progressBar1.Value = Math.Min(recordCount, progressBar1.Maximum)
                                        End If

                                    Catch ex As Exception
                                        fallidos += 1
                                        errorCampo = False
                                        Debug.WriteLine("❌ Error fila " & leidos & ": " & ex.Message)
                                    End Try

                                End While

                                ' FIX FINAL progress-bar
                                progressBar1.Value = progressBar1.Maximum

                            End Using
                        End Using
                    End Using
                End Using

            End Using

            ' MENSAJE FINAL DE CONTROL
            MessageBox.Show(
            "IMPORTACIÓN FINALIZADA" & vbCrLf & vbCrLf &
            "Total en Excel: " & totalRecords & vbCrLf &
            "Leídos: " & leidos & vbCrLf &
            "Insertados: " & insertados & vbCrLf &
            "Fallidos: " & fallidos,
            "CONTROL REAL",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub



    Public Shared Sub ImportarHispapost(ExcelFilePath As String, progressBar1 As ProgressBar)
        ' Conexión a la base de datos mediante la variable pública del módulo conexión accessConnString
        Using connection As New OleDb.OleDbConnection(accessConnString)
            connection.Open()

            ' Borrar los registros existentes en la tabla THPOST
            Dim deleteQuery As String = "DELETE FROM THPOST"
            Using deleteCommand As New OleDbCommand(deleteQuery, connection)
                deleteCommand.ExecuteNonQuery()
            End Using

            ' LIMPIAR EXCEL ANTES DE IMPORTAR
            ExcelFilePath = LimpiarExcel(ExcelFilePath)
            ' Conexión al archivo Excel usando OleDb
            Dim excelConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ExcelFilePath & ";Extended Properties=""Excel 8.0;HDR=YES;IMEX=1;"""

            Using excelConnection As New OleDb.OleDbConnection(excelConnectionString)
                excelConnection.Open()

                ' Obtener el nombre de la primera hoja activa
                Dim dtSheets As DataTable = excelConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
                If dtSheets Is Nothing OrElse dtSheets.Rows.Count = 0 Then
                    MsgBox("No se encontraron hojas en el archivo Excel.")
                    Exit Sub
                End If
                Dim firstSheetName As String = dtSheets.Rows(0)("TABLE_NAME").ToString()

                ' Contar el número total de registros en la hoja
                Dim countQuery As String = $"SELECT COUNT(*) FROM [{firstSheetName}]"
                Dim countCommand As New OleDbCommand(countQuery, excelConnection)
                Dim totalRecords As Integer = Convert.ToInt32(countCommand.ExecuteScalar())

                ' Consultar los datos de la primera hoja
                Dim query As String = $"SELECT * FROM [{firstSheetName}]"
                Using command As New OleDbCommand(query, excelConnection)
                    ' Configurar la ProgressBar
                    progressBar1.Minimum = 0
                    progressBar1.Maximum = totalRecords
                    progressBar1.Value = 0

                    Using reader As OleDbDataReader = command.ExecuteReader()
                        Dim recordCount As Integer = 0
                        ' Leer cada registro de la hoja Excel
                        While reader.Read()
                            ' Extraer los datos de cada columna de la hoja Excel
                            Dim cpostal As String = reader("CPOSTAL").ToString()
                            Dim poblacion As String = reader("POBLACION").ToString()
                            Dim zona As String = reader("ZONA").ToString()
                            Dim plataforma As String = reader("PLATAFORMA").ToString()

                            ' Comando SQL para insertar los datos en la tabla THPOST
                            Dim insertQuery As String = "INSERT INTO THPOST (CPOSTAL, POBLACION, ZONA, PLATAFORMA, FECHA) VALUES (?, ?, ?, ?, ?)"

                            Using insertCommand As New OleDbCommand(insertQuery, connection)
                                ' Agregar parámetros para evitar inyecciones SQL
                                insertCommand.Parameters.AddWithValue("?", cpostal)
                                insertCommand.Parameters.AddWithValue("?", poblacion)
                                insertCommand.Parameters.AddWithValue("?", zona)
                                insertCommand.Parameters.AddWithValue("?", plataforma)
                                insertCommand.Parameters.AddWithValue("?", DateTime.Today)

                                ' Ejecutar el comando
                                insertCommand.ExecuteNonQuery()
                            End Using

                            ' Actualizar progreso de la barra
                            recordCount += 1
                            If recordCount Mod 100 = 0 Then
                                ' Actualizar el valor de la ProgressBar en el hilo principal
                                progressBar1.Invoke(Sub()
                                                        progressBar1.Value = Math.Min(recordCount, progressBar1.Maximum)
                                                    End Sub)
                            End If
                        End While
                    End Using
                End Using
            End Using
        End Using

        ' Actualizar la barra de progreso al final (asegurándose que llegue al 100%)
        progressBar1.Invoke(Sub()
                                progressBar1.Value = progressBar1.Maximum
                            End Sub)

        ' Mensaje de confirmación
        MsgBox("Importación Completa")
    End Sub

End Class
