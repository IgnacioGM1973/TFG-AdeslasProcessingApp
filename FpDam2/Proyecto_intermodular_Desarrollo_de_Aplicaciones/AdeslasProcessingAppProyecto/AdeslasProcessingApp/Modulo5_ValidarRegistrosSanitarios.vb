Imports System.Data.OleDb
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

Public Class Modulo5_ValidarRegistrosSanitarios
    'Baja número de póliza certificado tabla  


    Public Shared Function BajaNumeroPolizaCertificadoorAccess(tarjetasSanitarias As List(Of String)) As List(Of String)

        Dim listaFiltrada As New List(Of String)()

        ' Validación inicial
        If tarjetasSanitarias Is Nothing OrElse tarjetasSanitarias.Count = 0 Then
            MessageBox.Show("La lista de registros está vacía o es nula.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return listaFiltrada ' 🔥 IMPORTANTE
        End If

        ' Ruta Access
        Dim rutaAccess As String = Path.Combine(RutaBD, "ADESLAS.accdb")

        Dim conexionStr As String = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={rutaAccess};Persist Security Info=False;"
        Dim clavesAccess As New HashSet(Of Tuple(Of String, String))()

        Try
            Using conexion As New OleDb.OleDbConnection(conexionStr)
                conexion.Open()

                Dim consulta As String = "SELECT POLIZA, NCERTIFICA FROM TCERTPROH"

                Using comando As New OleDbCommand(consulta, conexion)
                    Using lector As OleDbDataReader = comando.ExecuteReader()
                        While lector.Read()
                            Dim poliza As String = lector("POLIZA").ToString().Trim()
                            Dim ncertifica As String = lector("NCERTIFICA").ToString().Trim()

                            clavesAccess.Add(Tuple.Create(poliza, ncertifica))
                        End While
                    End Using
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show($"Error al leer la base de datos: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return listaFiltrada ' 🔥 IMPORTANTE
        End Try

        ' Procesar registros
        For Each registro As String In tarjetasSanitarias

            If registro.Length >= 21 Then
                Dim poliza As String = registro.Substring(3, 9).Trim()
                Dim certificado As String = registro.Substring(12, 9).Trim()

                If Not clavesAccess.Contains(Tuple.Create(poliza, certificado)) Then
                    listaFiltrada.Add(registro)
                End If
            End If

        Next

        ' Guardar fichero
        Try
            Dim carpetaDestino As String = Path.Combine(carpetaSalida, "reporte")

            If Not Directory.Exists(carpetaDestino) Then
                Directory.CreateDirectory(carpetaDestino)
            End If

            Dim rutaSalida As String = Path.Combine(carpetaDestino, "RegistrosFiltrados_" & NombreFichero)

            File.WriteAllLines(rutaSalida, listaFiltrada, System.Text.Encoding.Default)

            MessageBox.Show($"Registros filtrados guardados en: {rutaSalida}", "Proceso completado", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show($"Error al guardar el archivo: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' Lógica adicional
        RetirarNoProcesar(listaFiltrada)

        ' 🔥 DEVOLUCIÓN FINAL
        Return listaFiltrada

    End Function

    'Extraer del fichero Tarjetas para no procesar, mediante la tabla TarjetasNoProcesar.
    'Creamos un fichero con las que no tenemos que procesar
    'Extraer del fichero Tarjetas para no procesar, mediante la tabla TarjetasNoProcesar.
    'Creamos un fichero con las que no tenemos que procesar
    Public Shared Sub RetirarNoProcesar(listaRegistros As List(Of String))
        Dim outputFile As String = carpetaSalida & "reporte\RegistrosRetirados" & NombreFichero


        If Not Directory.Exists(Path.GetDirectoryName(outputFile)) Then
            Directory.CreateDirectory(Path.GetDirectoryName(outputFile))
        End If

        Dim connectionString As String =
        "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" &
        RutaBD & "\Adeslas.accdb"

        Dim registrosFiltrados As New List(Of String)
        Dim registrosRetirados As New List(Of String)
        Dim ContaRegExtraer As Long

        ' ==========================================================
        ' 1) Cargar TarjetasNoProcesar (clave: DELEGACION+POLIZA+CERT+NUOR)
        ' ==========================================================
        Dim registrosNoProcesar As New HashSet(Of String)(StringComparer.Ordinal)

        ' ==========================================================
        ' 2) Cargar MGA (clave: POLIZA|CERTIFICADO)  -> SOLO para regla 666300454
        ' ==========================================================
        Dim clavesMGA As New HashSet(Of String)(StringComparer.Ordinal)
        Dim polizaMGAControl As String = "666300454" ' <- la póliza que quieres controlar

        Try
            Using connection As New OleDb.OleDbConnection(connectionString)
                connection.Open()

                ' ---- 1) TarjetasNoProcesar
                Dim queryNP As String =
                "SELECT DELEGACION, POLIZA, CERTIFICADO, NUOR FROM TarjetasNoProcesar"

                Using command As New OleDb.OleDbCommand(queryNP, connection)
                    Using reader As OleDb.OleDbDataReader = command.ExecuteReader()
                        While reader.Read()

                            Dim delegacion As String = reader("DELEGACION").ToString().Trim().PadRight(3)
                            Dim poliza As String = reader("POLIZA").ToString().Trim().PadRight(9)
                            Dim certificado As String = reader("CERTIFICADO").ToString().Trim().PadRight(9)

                            ' ⚠️ OJO: aquí estabas usando 2 chars. Mantengo tu criterio (21,2)
                            Dim nuor As String = reader("NUOR").ToString().Trim()
                            If nuor.Length > 2 Then nuor = nuor.Substring(0, 2)
                            If nuor = "" Then nuor = " "

                            Dim claveNoProcesar As String = delegacion & poliza & certificado & nuor
                            registrosNoProcesar.Add(claveNoProcesar)
                        End While
                    End Using
                End Using

                ' ---- 2) MGA: solo nos interesa POCENPOL+POCECDCE
                Dim queryMGA As String = "SELECT POCENPOL, POCECDCE FROM MGA"
                Using command As New OleDb.OleDbCommand(queryMGA, connection)
                    Using reader As OleDb.OleDbDataReader = command.ExecuteReader()
                        While reader.Read()
                            Dim p As String = If(reader.IsDBNull(0), "", reader.GetValue(0).ToString().Trim())
                            Dim c As String = If(reader.IsDBNull(1), "", reader.GetValue(1).ToString().Trim())

                            If p <> "" AndAlso c <> "" Then
                                clavesMGA.Add(p & "|" & c)
                            End If
                        End While
                    End Using
                End Using

            End Using

        Catch ex As Exception
            MsgBox("Error al conectar con la base de datos: " & ex.Message)
            ' Si falla la BD, seguimos con lo que tengamos, pero ojo: puede cambiar el comportamiento
        End Try

        ' ==========================================================
        ' 3) FILTRAR REGISTROS DEL TXT: NoProcesar + Regla MGA
        ' ==========================================================
        For Each registro As String In listaRegistros

            If String.IsNullOrEmpty(registro) OrElse registro.Length < 23 Then
                ' Para Substring(21,2) necesito mínimo 23
                registrosFiltrados.Add(registro)
                Continue For
            End If

            Dim delegacion As String = registro.Substring(0, 3).PadRight(3)
            Dim poliza As String = registro.Substring(3, 9).PadRight(9)
            Dim certificado As String = registro.Substring(12, 9).PadRight(9)

            ' NUOR posición 21 longitud 2 (según tu implementación actual)
            Dim nuorRegistro As String = Trim(registro.Substring(21, 2))
            If nuorRegistro.Length = 0 Then nuorRegistro = " "

            Dim claveCompararNP As String = delegacion & poliza & certificado & nuorRegistro

            ' ---- A) Retirada por TarjetasNoProcesar
            If registrosNoProcesar.Contains(claveCompararNP) Then
                registrosRetirados.Add("NO_PROCESAR | " & registro)
                ContaRegExtraer += 1
                Continue For
            End If

            ' ---- B) Retirada por regla MGA (solo póliza 666300454)
            Dim polizaTrim As String = registro.Substring(3, 9).Trim()
            If polizaTrim = polizaMGAControl Then
                Dim certTrim As String = registro.Substring(12, 9).Trim()
                Dim kMGA As String = polizaTrim & "|" & certTrim

                If Not clavesMGA.Contains(kMGA) Then
                    registrosRetirados.Add("MGA_NO_COINCIDE | " & registro)
                    ContaRegExtraer += 1
                    Continue For
                End If
            End If

            ' ---- OK -> sigue el flujo normal
            registrosFiltrados.Add(registro)

        Next

        ' ==========================================================
        ' 4) Guardar retirados (incluye NoProcesar + MGA)
        ' ==========================================================
        Try
            File.WriteAllLines(outputFile, registrosRetirados, Encoding.Default)
        Catch ex As Exception
            MessageBox.Show("Error guardando retirados: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        MessageBox.Show($"Registros retirados (NoProcesar + MGA): {ContaRegExtraer}",
                    "Proceso Completado",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information)

        ' ==========================================================
        ' 5) Continuar pipeline SOLO con los filtrados
        ' ==========================================================
        ValidacionRegistro(registrosFiltrados)

    End Sub





    'Hace una validación de los registros
    Public Shared Function ValidacionRegistro(listaRegistrosSinNoProcesar As List(Of String)) As List(Of String)

        Dim registrosValidados As New List(Of String)()

        ' Carpeta salida
        Dim outputFolder As String = Path.Combine(carpetaSalida, "reporte")
        If Not Directory.Exists(outputFolder) Then Directory.CreateDirectory(outputFolder)

        Dim fragmentosValidados As New HashSet(Of String)()
        Dim registrosRepetidos As New List(Of String)()
        Dim registrosErroneos As New List(Of String)()

        Try
            Dim lineNumber As Integer = 0

            For Each line As String In listaRegistrosSinNoProcesar
                lineNumber += 1
                Dim lineNumberFormatted As String = lineNumber.ToString().PadLeft(8, " "c)

                ' 🔹 Validación de longitud
                If line.Length <> 781 And line.Length <> 537 And line.Length <> 538 Then
                    registrosErroneos.Add($"[Línea {lineNumberFormatted}] Longitud inválida: {line}")
                    Continue For
                End If

                ' 🔹 Seguridad antes de substrings
                If line.Length < 750 Then
                    registrosErroneos.Add($"[Línea {lineNumberFormatted}] Registro demasiado corto: {line}")
                    Continue For
                End If

                ' 🔹 Número tarjeta
                Dim ntarjeta As String = line.Substring(368, 8).Trim()
                If ntarjeta = "00000000" Then
                    registrosErroneos.Add($"[Línea {lineNumberFormatted}] Nº tarjeta inválido: {line}")
                    Continue For
                End If

                ' 🔹 Fragmento duplicado
                Dim fragmento As String = line.Substring(367, 9).Trim()

                If fragmento = "" AndAlso line.Substring(648, 2) <> "30" AndAlso line.Substring(648, 2) <> "32" Then
                    registrosErroneos.Add($"[Línea {lineNumberFormatted}] Fragmento vacío: {line}")
                    Continue For
                End If

                If fragmento <> "" AndAlso Not fragmentosValidados.Add(fragmento) Then
                    registrosRepetidos.Add($"[Línea {lineNumberFormatted}] Duplicado: {line}")
                    Continue For
                End If

                ' 🔹 Póliza
                Dim codigoPoliza As String = line.Substring(3, 9).Trim()
                If String.IsNullOrEmpty(codigoPoliza) Then
                    registrosErroneos.Add($"[Línea {lineNumberFormatted}] Póliza inválida: {line}")
                    Continue For
                End If

                ' 🔹 Nombre
                Dim nombre As String = line.Substring(53, 28).Trim()
                If String.IsNullOrEmpty(nombre) OrElse nombre.Length <= 4 OrElse Regex.IsMatch(nombre, "\d") Then
                    registrosErroneos.Add($"[Línea {lineNumberFormatted}] Nombre inválido: {line}")
                    Continue For
                End If

                If nombre.Contains("?") OrElse nombre.Contains("/") OrElse nombre.Contains("PRUEBA") Then
                    registrosErroneos.Add($"[Línea {lineNumberFormatted}] Nombre no permitido: {line}")
                    Continue For
                End If

                ' 🔹 Sexo
                Dim sexo As String = line.Substring(358, 1).Trim()
                If sexo <> "H" AndAlso sexo <> "M" Then
                    registrosErroneos.Add($"[Línea {lineNumberFormatted}] Sexo inválido: {line}")
                    Continue For
                End If

                ' 🔹 ISFAS
                If line.Substring(25, 5).Trim() = "ISFAS" AndAlso line.Substring(749, 16).Trim() = "" Then
                    registrosErroneos.Add($"[Línea {lineNumberFormatted}] DataMatrix vacío: {line}")
                    Continue For
                End If

                ' 🔹 Dirección envío
                Dim direccionEnvio As String = line.Substring(198, 40)
                If direccionEnvio.Trim() = "" Then
                    registrosErroneos.Add($"[Línea {lineNumberFormatted}] Dirección envío vacía: {line}")
                    Continue For
                End If

                ' 🔹 Código postal
                If String.IsNullOrEmpty(line.Substring(298, 5).Trim()) Then
                    registrosErroneos.Add($"[Línea {lineNumberFormatted}] CP inválido: {line}")
                    Continue For
                End If

                ' 🔹 Año nacimiento
                If String.IsNullOrEmpty(line.Substring(359, 8).Trim()) Then
                    registrosErroneos.Add($"[Línea {lineNumberFormatted}] Año nacimiento inválido: {line}")
                    Continue For
                End If

                ' 🔹 Correcciones texto
                line = line.Substring(0, 402) &
                   line.Substring(402, 20).Replace("ASITENCIA SANITARIA", "ASISTENCIA SANITARIA") &
                   line.Substring(422)

                line = line.Substring(0, 402) &
                   line.Substring(402, 20).Replace("Y PUS DENTAL", "Y PLUS DENTAL") &
                   line.Substring(422)

                ' 🔹 Modelo tarjeta
                If line.Substring(648, 2) = "69" Then
                    line = line.Substring(0, 648) & "00" & line.Substring(650)
                End If

                ' ✅ OK
                registrosValidados.Add(line)

            Next

            ' 🔹 Guardar resultados
            SaveResults(outputFolder, NombreFicheroCompleto.Replace(".txt", "") & "_Validados.txt", registrosValidados)
            SaveResults(outputFolder, NombreFicheroCompleto.Replace(".txt", "") & "_Repetidos.txt", registrosRepetidos)
            SaveResults(outputFolder, NombreFicheroCompleto.Replace(".txt", "") & "_Erroneos.txt", registrosErroneos)

            MessageBox.Show($"Proceso finalizado{vbNewLine}" &
                        $"✔ Válidos: {registrosValidados.Count}{vbNewLine}" &
                        $"⚠ Repetidos: {registrosRepetidos.Count}{vbNewLine}" &
                        $"✖ Erróneos: {registrosErroneos.Count}",
                        "Finalizado", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show($"Error procesando: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' 🔹 Ordenación final
        OrdenarFichero(registrosValidados)

        ' 🔥 RETURN FINAL (CLAVE)
        Return registrosValidados

    End Function
    ' Método auxiliar para guardar resultados
    Private Shared Sub SaveResults(outputFolder As String, fileName As String, data As IEnumerable(Of String))

        Dim fullPath As String = Path.Combine(outputFolder, fileName)
        'Cambio de Carpeta de salida
        fullPath = Replace(fullPath, "\Ficheros_Entrada\", "\Ficheros_Salida\Reporte\")
        Try
            File.WriteAllLines(fullPath, data, Encoding.Default)
        Catch ex As Exception
            Console.WriteLine($"Error al guardar {fileName}: {ex.Message}")
        End Try
    End Sub

    'Ordena por Cd_delegacion, Número Poliza, Número Certificado, Número orden, Código Rela
    Public Shared Sub OrdenarFichero(ListaRegistros As List(Of String))
        Dim lineasOrdenadas As New List(Of String)()
        Dim outputFile As String = carpetaSalida & "reporte\RegistrosOrdenados" & NombreFichero

        ' Crear la carpeta de salida si no existe
        If Not Directory.Exists(Path.GetDirectoryName(outputFile)) Then
            Directory.CreateDirectory(Path.GetDirectoryName(outputFile))
        End If

        Try
            Dim lineasProcesadas As New List(Of String)

            ' Procesar cada línea
            For Each linea As String In ListaRegistros
                If linea.Length >= 25 Then
                    Dim CodigoDelegacionFormateado As String = linea.Substring(0, 3).Trim().PadLeft(3, " "c)
                    Dim NumeroDePolizaFormateado As String = linea.Substring(3, 9).Trim().PadLeft(9, " "c)
                    Dim NumeroDeCertificadoFormateado As String = linea.Substring(12, 9).Trim().PadLeft(9, " "c)
                    Dim NumeroOrdenFormateado As String = linea.Substring(21, 2).Trim().PadLeft(2, " "c)
                    Dim CodigoRelacionFormateado As String = linea.Substring(23, 2).Trim().PadLeft(2, " "c)

                    Dim lineaFormateada As String = CodigoDelegacionFormateado &
                                                NumeroDePolizaFormateado &
                                                NumeroDeCertificadoFormateado &
                                                NumeroOrdenFormateado &
                                                CodigoRelacionFormateado

                    lineasProcesadas.Add(lineaFormateada & "-" & linea)
                Else
                    lineasProcesadas.Add(linea)
                End If
            Next

            ' Ordenar las líneas por los campos especificados
            lineasProcesadas.Sort(Function(x, y)
                                      Dim result = x.Substring(0, 3).CompareTo(y.Substring(0, 3))
                                      If result = 0 Then result = x.Substring(3, 9).CompareTo(y.Substring(3, 9))
                                      If result = 0 Then result = x.Substring(12, 9).CompareTo(y.Substring(12, 9))
                                      If result = 0 Then result = x.Substring(21, 2).CompareTo(y.Substring(21, 2))
                                      If result = 0 Then result = x.Substring(23, 2).CompareTo(y.Substring(23, 2))
                                      Return result
                                  End Function)

            ' Escribir en el archivo de salida
            Using sw As New StreamWriter(outputFile, False, Encoding.Default)
                For Each linea In lineasProcesadas
                    If linea.Length > 26 Then ' Evitar errores con líneas cortas
                        sw.WriteLine(linea.Substring(26))
                    Else
                        sw.WriteLine(linea)
                    End If
                    lineasOrdenadas.Add(linea.Substring(26))
                Next
            End Using

            ' Mensajes de éxito
            Console.WriteLine("Procesamiento completado. Archivo generado en: " & outputFile)
        Catch ex As Exception
            Console.WriteLine("Ocurrió un error: " & ex.Message)
        End Try
        CorregirDireccion(lineasOrdenadas)
        'Return lineasOrdenadas
        'SaveResults(outputFile, "RegistrosValidados.txt", lineasOrdenadas)
    End Sub

    'Corrige La direccón del encarte para que salga como quiere Adeslas con el nombre del tomador en las
    'tarjetas de los asegurados y no el del receptor
    Public Shared Sub CorregirDireccion(ListaRegistros As List(Of String))
        ' Directorio de salida
        Dim outputFile As String = carpetaSalida & "reporte\RegistrosOrdenadoCompletados" & NombreFichero
        If Not Directory.Exists(Path.GetDirectoryName(outputFile)) Then
            Directory.CreateDirectory(Path.GetDirectoryName(outputFile))
        End If

        ' Agrupar registros por la clave en las posiciones 0 a 21
        Dim registrosAgrupados As New Dictionary(Of String, List(Of String))()
        For Each line As String In ListaRegistros
            If line.Length < 781 Then Continue For ' Ignorar registros con longitud inválida
            Dim claveAgrupacion As String = line.Substring(0, 21).Trim()

            If Not registrosAgrupados.ContainsKey(claveAgrupacion) Then
                registrosAgrupados(claveAgrupacion) = New List(Of String)()
            End If

            registrosAgrupados(claveAgrupacion).Add(line)
        Next

        ' Actualizar registros según la lógica del menor valor en las posiciones 21,4
        For Each clave As String In registrosAgrupados.Keys
            Dim grupo As List(Of String) = registrosAgrupados(clave)

            ' Encontrar el registro con el menor valor en las posiciones 21,4
            Dim registroMenorValor As String = grupo.OrderBy(Function(registro) registro.Substring(21, 4).Trim()).FirstOrDefault()
            If registroMenorValor Is Nothing Then Continue For ' Si no hay registros válidos, continuar

            ' Extraer los valores del registro con el menor valor
            Dim nombreTomador As String = registroMenorValor.Substring(53, 28).Trim()
            Dim direccionTomador As String = registroMenorValor.Substring(198, 40).Trim()
            Dim cdPostalTomador As String = registroMenorValor.Substring(298, 5).Trim()
            Dim poblacionTomador As String = registroMenorValor.Substring(303, 25).Trim()
            Dim provinciaTomador As String = registroMenorValor.Substring(328, 30).Trim()

            Dim nombreReceptor As String = registroMenorValor.Substring(442, 55) '.Trim()
            Dim direccionReceptor As String = registroMenorValor.Substring(497, 40) '.Trim()
            Dim cdPostalReceptor As String = registroMenorValor.Substring(537, 5) '.Trim()
            Dim poblacionReceptor As String = registroMenorValor.Substring(542, 25) '.Trim()
            Dim provinciaReceptor As String = registroMenorValor.Substring(567, 30) '.Trim()


            ' Actualizar los registros del grupo con la información del registro menor 
            For i As Integer = 0 To grupo.Count - 1
                Dim registro As String = grupo(i)

                If registro.Length < 781 Then Continue For ' Ignorar registros con longitud inválida

                Dim valorActual442155 As String = registro.Substring(442, 155).Trim()
                If String.IsNullOrEmpty(valorActual442155) Then 'si solo quiero que se active cuando la direccion del receptor sea nula.
                    ' Concatenar y actualizar el registro
                    Dim parteInicial As String = registro.Substring(0, 442) ' Hasta la posición 442
                    Dim parteFinal As String = registro.Substring(597) ' Desde la posición 597 en adelante

                    ' Crear el registro actualizado
                    grupo(i) = parteInicial &
                           nombreTomador.PadRight(55) &
                           direccionTomador.PadRight(40) &
                           cdPostalTomador.PadRight(5) &
                           poblacionTomador.PadRight(25) &
                           provinciaTomador.PadRight(30) &
                           parteFinal
                Else
                    ' Concatenar y actualizar el registro
                    Dim parteInicial As String = registro.Substring(0, 442) ' Hasta la posición 442
                    Dim parteFinal As String = registro.Substring(597) ' Desde la posición 597 en adelante

                    ' Crear el registro actualizado
                    grupo(i) = parteInicial &
                           nombreReceptor.PadRight(55) &
                           direccionReceptor.PadRight(40) &
                           cdPostalReceptor.PadRight(5) &
                           poblacionReceptor.PadRight(25) &
                           provinciaReceptor.PadRight(30) &
                           parteFinal
                End If

            Next
        Next

        ' Guardar los registros actualizados en el archivo de salida
        Dim registrosActualizados As New List(Of String)
        For Each grupo As List(Of String) In registrosAgrupados.Values
            registrosActualizados.AddRange(grupo)
        Next

        File.WriteAllLines(outputFile, registrosActualizados, Encoding.Default)
        MessageBox.Show($"Procesamiento completado. Archivo guardado en: {outputFile}", "Finalización", MessageBoxButtons.OK, MessageBoxIcon.Information)
        'Return registrosActualizados
        OrdenarParaCorreos(registrosActualizados)


    End Sub




    'Ordenar los registros para correos    
    Public Shared Sub OrdenarParaCorreos(ListaRegistros As List(Of String))

        Dim outputFile As String = carpetaSalida & "reporte\RegistrosOrdenadosPostal" & NombreFichero

        ' Crear la carpeta de salida si no existe
        If Not Directory.Exists(Path.GetDirectoryName(outputFile)) Then
            Directory.CreateDirectory(Path.GetDirectoryName(outputFile))
        End If

        Try
            ' Proyectamos cada línea con sus claves de ordenación
            ' Añadimos también el índice original para garantizar estabilidad en empates
            Dim registrosOrdenados = ListaRegistros _
            .Select(Function(linea, idx)
                        Return New With {
                            .Idx = idx,
                            .Linea = linea,
                            .CP = ExtraerCampo(linea, 537, 5),
                            .Poliza = ExtraerCampo(linea, 3, 9),
                            .Certificado = ExtraerCampo(linea, 12, 9),
                            .NumOrden = ExtraerCampo(linea, 21, 2),
                            .CodRelacion = ExtraerCampo(linea, 23, 2)
                        }
                    End Function) _
            .OrderBy(Function(x) x.CP) _
            .ThenBy(Function(x) x.Poliza) _
            .ThenBy(Function(x) x.Certificado) _
            .ThenBy(Function(x) x.CodRelacion) _
            .ThenBy(Function(x) x.NumOrden) _
            .ThenBy(Function(x) x.Idx) _
            .Select(Function(x) x.Linea) _
            .ToList()

            ' Escribimos el fichero ya ordenado
            Using sw As New StreamWriter(outputFile, False, Encoding.Default)
                For Each linea In registrosOrdenados
                    sw.WriteLine(linea)
                Next
            End Using

            ' Devolvemos la misma lista ordenada al siguiente proceso
            AddAsesorSenior(registrosOrdenados)

            Console.WriteLine("Procesamiento completado. Archivo generado en: " & outputFile)

        Catch ex As Exception
            Console.WriteLine("Ocurrió un error: " & ex.Message)
        End Try

    End Sub

    ' ===========================================
    '  Helper: extrae campo de forma segura
    ' ===========================================
    Private Shared Function ExtraerCampo(linea As String, inicio As Integer, longitud As Integer) As String
        If String.IsNullOrEmpty(linea) Then Return "".PadLeft(longitud, "0"c)

        ' Si la línea es demasiado corta, evitamos excepción y devolvemos "0..."
        If linea.Length < inicio + longitud Then
            Return "".PadLeft(longitud, "0"c)
        End If

        Dim campo = linea.Substring(inicio, longitud).Trim()

        ' Rellenamos a la izquierda para que la comparación de texto se comporte como numérica
        Return campo.PadLeft(longitud, "0"c)
    End Function


    'Añadir nombre Asesor y teléfono 
    Public Shared Sub AddAsesorSenior(lineasOrdenadasCorreos As List(Of String))
        Dim asesor As String
        Dim dirAsesor As String
        Dim telAsesor As String

        ' Lista para almacenar los registros actualizados
        ' List to store the updated records
        Dim registrosActualizados As New List(Of String)

        ' Cadena de conexión a la base de datos
        ' Database connection string
        Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & RutaBD & "\adeslas.accdb;"

        ' Conexión a la base de datos
        ' Connection to the database
        Using connection As New OleDb.OleDbConnection(connectionString)
            Try
                connection.Open()

                ' Recorremos cada línea de la lista
                ' Iterate over each line in the list
                For Each linea In lineasOrdenadasCorreos
                    ' Obtenemos el campo de la posición 367 (9 caracteres)
                    ' Extract the field at position 367 (9 characters)
                    Dim tjasntar As String = linea.Substring(367, 9)

                    ' Consulta SQL para buscar el TJASNTAR
                    ' SQL query to search for TJASNTAR
                    Dim query As String = "SELECT ASESOR, DIR_ASESOR, TEL_ASESOR FROM AsesorSenior WHERE TJASNTAR = @TJASNTAR"
                    Using command As New OleDbCommand(query, connection)
                        ' Parámetro de búsqueda
                        ' Search parameter
                        command.Parameters.AddWithValue("@TJASNTAR", tjasntar)

                        ' Ejecutar la consulta
                        ' Execute the query
                        Using reader As OleDbDataReader = command.ExecuteReader()
                            If reader.Read() Then
                                ' Si hay coincidencia, obtenemos los valores de ASESOR, DIR_ASESOR y TEL_ASESOR
                                ' If there is a match, retrieve the values of ASESOR, DIR_ASESOR and TEL_ASESOR
                                asesor = reader("ASESOR").ToString().PadRight(41).Substring(0, 41)
                                dirAsesor = reader("DIR_ASESOR").ToString().PadRight(100).Substring(0, 100)
                                telAsesor = reader("TEL_ASESOR").ToString()

                                ' Añadimos los valores directamente al final del registro original
                                ' Add the values directly to the end of the original record
                                Dim registroActualizado As String = linea & asesor & dirAsesor & telAsesor
                                registrosActualizados.Add(registroActualizado)
                            Else
                                ' Si no hay coincidencia, añadimos el registro original
                                ' If there is no match, add the original record
                                asesor = Space(41)
                                dirAsesor = Space(100)
                                telAsesor = Space(9)
                                registrosActualizados.Add(linea & asesor & dirAsesor & telAsesor)
                            End If
                        End Using
                    End Using
                Next

            Catch ex As Exception
                ' Manejar errores
                ' Handle errors
                Throw New Exception($"Error al procesar los registros: {ex.Message}")
            Finally
                ' Aseguramos cerrar la conexión
                ' Ensure the connection is closed
                If connection.State = ConnectionState.Open Then
                    connection.Close()
                End If
            End Try
        End Using

        ' Devolvemos la lista actualizada
        ' Return the updated list
        'Return registrosActualizados
        TextoMGA(registrosActualizados)
    End Sub

    'Añade texto a las tarjetas MGA en función del listado Excel 300454.xlsx importado en la bd
    Public Shared Sub TextoMGA(lineasOrgenadasCorreos As List(Of String))
        Dim registrosMarcados As New List(Of String)()
        Dim nuevaLinea As String = String.Empty

        ' Ruta de la base de datos
        Dim rutaBD As String = RutaBase & "\AdeslasProcessingAppProyecto\AdeslasProcessingApp\bin\Debug\adeslas.accdb" ' Cambiar a la base de datos .accdb

        ' Diccionario para almacenar datos de la base de datos
        Dim datosBD As New List(Of Tuple(Of String, String, String))() ' Tuple: POCENPOL, POCECDCE, DENTAL

        ' Verificar si la fecha en la base de datos corresponde a la fecha del día
        Try
            Using conn As New OleDb.OleDbConnection($"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={rutaBD}") ' Usar el proveedor ACE para .accdb
                conn.Open()
                Dim queryFecha As String = "SELECT MIN(FECHA) FROM MGA" ' Seleccionar la fecha MENOS reciente de la tabla MGA
                Using cmd As New OleDb.OleDbCommand(queryFecha, conn)
                    Dim fechaBD As DateTime = Convert.ToDateTime(cmd.ExecuteScalar())
                    Dim fechaHoy As DateTime = DateTime.Today

                    ' Comparar la fecha de la base de datos con la fecha actual
                    If fechaBD <> fechaHoy Then
                        ' Preguntar al usuario si desea continuar
                        Dim resultado As DialogResult = MessageBox.Show("La fecha en la base de datos no corresponde al día de hoy. ¿Desea continuar?", "Fecha incorrecta", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                        If resultado = DialogResult.No Then
                            ' Si el usuario elige "No", cancelar la ejecución
                            'Return registrosMarcados
                        End If
                    End If
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error al acceder a la base de datos para verificar la fecha: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'Return registrosMarcados
        End Try

        ' Conectar con la base de datos y cargar los valores POCENPOL, POCECDCE y DENTAL
        Try
            Using conn As New OleDb.OleDbConnection($"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={rutaBD}") ' Usar el proveedor ACE para .accdb
                conn.Open()
                Dim query As String = "SELECT POCENPOL, POCECDCE, DENTAL FROM MGA"
                Using cmd As New OleDb.OleDbCommand(query, conn)
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim pocenpol As String = If(Not reader.IsDBNull(0), reader.GetValue(0).ToString().Trim(), String.Empty)
                            Dim pocecdce As String = If(Not reader.IsDBNull(1), reader.GetValue(1).ToString().Trim(), String.Empty)
                            Dim dental As String = If(Not reader.IsDBNull(2), reader.GetValue(2).ToString().Trim(), String.Empty)

                            ' Agregar al diccionario si los valores no están vacíos
                            If Not String.IsNullOrEmpty(pocenpol) AndAlso Not String.IsNullOrEmpty(pocecdce) Then
                                datosBD.Add(Tuple.Create(pocenpol, pocecdce, dental))
                            End If
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error al acceder a la base de datos: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'Return registrosMarcados
        End Try

        ' Procesar cada línea de la lista
        For Each linea As String In lineasOrgenadasCorreos
            If linea.Length < 21 Then ' Validar longitud mínima
                registrosMarcados.Add(linea)
                Continue For
            End If

            ' Extraer claves de las posiciones del registro
            Dim clavePocenpol As String = linea.Substring(3, 9).Trim() ' Clave en posición 3,9
            Dim clavePocecdce As String = linea.Substring(12, 9).Trim() ' Clave en posición 12,9
            Dim dentalValor As String = String.Empty
            Dim textoPersonalizado1 As String = String.Empty
            Dim textoPersonalizado2 As String = String.Empty

            ' Buscar coincidencias en la base de datos
            For Each registroBD In datosBD
                If registroBD.Item1 = clavePocenpol AndAlso registroBD.Item2 = clavePocecdce Then
                    If registroBD.Item3 = "SI" Then
                        dentalValor = "SI"
                        textoPersonalizado1 = "INCLUYE ASISTENCIA  "
                        textoPersonalizado2 = "SANITARIA Y DENTAL  "
                    Else
                        textoPersonalizado1 = Space(20)
                        textoPersonalizado2 = Space(20)
                    End If
                    Exit For
                End If
            Next

            ' Añadir el valor del campo DENTAL al final del registro si hay coincidencia
            If Not String.IsNullOrEmpty(dentalValor) Then
                ' Construir nueva línea con los textos personalizados
                nuevaLinea = linea.Substring(0, 402) & textoPersonalizado1 & textoPersonalizado2 & linea.Substring(442)
            Else
                nuevaLinea = linea
            End If

            registrosMarcados.Add(nuevaLinea)

        Next

        'Añadir los registros tarjetas sanitarias diario en la tabla TarjetasSanitariasDiario
        'pendiente añadir los tres campos extras del registros de los Agentes
        'SERGIO GIRON MORILLA                     SERVICIOS MEDICOS AUXILIARES, S.A. TORRENT DE LïOLLA 1                                              932705400"

        Mudulo6_AnexarRegistrosSanitarios_A_Tablas_TSD.CargaRegistrosTablaTarjetasSanitariasDiario(registrosMarcados)


    End Sub
End Class
