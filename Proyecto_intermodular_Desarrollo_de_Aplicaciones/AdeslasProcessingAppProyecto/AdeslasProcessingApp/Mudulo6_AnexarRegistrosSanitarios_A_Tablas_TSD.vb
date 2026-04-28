Public Class Mudulo6_AnexarRegistrosSanitarios_A_Tablas_TSD

    Public Shared Sub CargaRegistrosTablaTarjetasSanitariasDiario(ListaRegistros As List(Of String))
        If ListaRegistros Is Nothing OrElse ListaRegistros.Count = 0 Then
            MessageBox.Show("La lista de registros está vacía.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' Inicializar las listas
        Dim TarjetasSanitariasDiarioTag As New List(Of String)
        Dim TarjetasSanitariasDiarioTelemail As New List(Of String)
        Dim TarjetasSanitariasDiarioInteramit As New List(Of String)

        ' Clasificar los registros en las listas adecuadas
        For Each registro In ListaRegistros
            If registro.Length < 5 Then
                MessageBox.Show("Registro inválido: " & registro, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Continue For
            End If

            Select Case True
                ' QUITO LA SELECCION DE VC PARA QUE LO CARGUE TODO EN LA MISMA BD
                'Case registro.Substring(3, 2) = "88"
                '    TarjetasSanitariasDiarioTag.Add(registro)
                'Case registro.StartsWith("331") OrElse registro.StartsWith("481") OrElse registro.StartsWith("11201")
                Case registro.StartsWith("331") OrElse registro.StartsWith("11201")
                    TarjetasSanitariasDiarioTelemail.Add(registro)
                Case Else
                    TarjetasSanitariasDiarioInteramit.Add(registro)
            End Select
        Next

        ' Llamar a cargaDatosbd para cada lista con su respectiva tabla
        If TarjetasSanitariasDiarioTag.Count > 0 Then
            cargaDatosbd(TarjetasSanitariasDiarioTag, "TarjetasSanitariasDiarioTag")
        End If
        If TarjetasSanitariasDiarioTelemail.Count > 0 Then
            cargaDatosbd(TarjetasSanitariasDiarioTelemail, "TarjetasSanitariasDiarioTelemail")
        End If
        If TarjetasSanitariasDiarioInteramit.Count > 0 Then
            cargaDatosbd(TarjetasSanitariasDiarioInteramit, "TarjetasSanitariasDiarioInteramit")
        End If

    End Sub

    Public Shared Sub cargaDatosbd(ListaRegistros As List(Of String), nomTabla As String)

        Dim AccessDBPath As String = RutaBD & "\ADESLAS.accdb"
        Dim connectionString As String = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={AccessDBPath};Persist Security Info=False;"


        Using connection As New OleDb.OleDbConnection(connectionString)
            Try
                connection.Open()

                ' Paso 1: Borrar todos los datos de la tabla
                Dim deleteQuery As String = $"DELETE FROM {nomTabla}"
                Using deleteCommand As New OleDb.OleDbCommand(deleteQuery, connection)
                    deleteCommand.ExecuteNonQuery()
                End Using



                ' Paso 2: Insertar registros en la tabla
                Dim insertQuery As String = $"INSERT INTO {nomTabla} (CodigoDelegacion, NumeroPoliza, NumeroCertificado, NumeroOrden, CodigoRelacion, " &
                "ColectivoProducto, NombreApellidos, ProvinciaNacimiento, NIF, EstadoCivil, Profesion, TelefonoParticular, TelefonoTrabajo, Direccion, " &
                "CodigoPostal, Poblacion, Provincia, Sexo, AnoNacimiento, NumeroTarjeta, FechaAlta, Version, FechaInicioCarencia, FechaCaducidad, " &
                "MotivoPeticion, IndicadorPagoTalones, TipoContrato, DigitoControlProvincia, DigitoControlZ, IndicadorExtranjero, IndicadorExtraccion, " &
                "TextoPersonalizado1, TextoPersonalizado2, PersonaReceptora, DireccionEnvio, CodigoPostalEnvio, PoblacionEnvio, ProvinciaEnvio, " &
                "CodigoTarifa, CentroTrabajoCodigo, CentroTrabajoDescripcion, IndicadorCarrier, ModeloPlasticoCodigo, ModeloPlasticoDescripcion, " &
                "LogotipoDescripcion, ModeloCarrier, IndicadorIdioma, Marca, CIP_SNS, CIP_M) " &
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

                Using insertCommand As New OleDb.OleDbCommand(insertQuery, connection)
                    ' Crear los parámetros para la consulta
                    For i = 1 To 50
                        insertCommand.Parameters.Add(New OleDb.OleDbParameter($"Param{i}", OleDb.OleDbType.VarChar))
                    Next

                    ' Procesar cada registro de la lista
                    For Each registro In ListaRegistros
                        'REMPLAZAR CARACTERES ILEGALES
                        registro = Replace(registro, "Á", "A")
                        registro = Replace(registro, "É", "E")
                        registro = Replace(registro, "Í", "I")
                        registro = Replace(registro, "Ó", "O")
                        registro = Replace(registro, "Ú", "U")
                        registro = Replace(registro, "‡", "C")
                        registro = registro.Replace("à"c, "O"c)
                        registro = Replace(registro, "€", "C")





                        ' Formateamos el valor del código de plástico a 0 a la izquierda
                        If registro.Substring(649, 1) = " " Then
                            Dim caracteres() As Char = registro.ToCharArray()
                            caracteres(649) = caracteres(648)
                            caracteres(648) = "0"c
                            registro = New String(caracteres)
                        End If
                        ' Asignar valores a los parámetros
                        insertCommand.Parameters(0).Value = registro.Substring(0, 3)
                        insertCommand.Parameters(1).Value = registro.Substring(3, 9)
                        insertCommand.Parameters(2).Value = registro.Substring(12, 9)
                        insertCommand.Parameters(3).Value = registro.Substring(21, 2)
                        insertCommand.Parameters(4).Value = registro.Substring(23, 2)
                        insertCommand.Parameters(5).Value = registro.Substring(25, 28)
                        insertCommand.Parameters(6).Value = registro.Substring(53, 28)
                        insertCommand.Parameters(7).Value = registro.Substring(81, 30)
                        insertCommand.Parameters(8).Value = registro.Substring(111, 9)
                        insertCommand.Parameters(9).Value = registro.Substring(120, 30)
                        insertCommand.Parameters(10).Value = registro.Substring(150, 30)
                        insertCommand.Parameters(11).Value = registro.Substring(180, 9)
                        insertCommand.Parameters(12).Value = registro.Substring(189, 9)
                        insertCommand.Parameters(13).Value = registro.Substring(198, 100)
                        insertCommand.Parameters(14).Value = registro.Substring(298, 5)
                        insertCommand.Parameters(15).Value = registro.Substring(303, 25)
                        insertCommand.Parameters(16).Value = registro.Substring(328, 30)
                        insertCommand.Parameters(17).Value = registro.Substring(358, 1)
                        insertCommand.Parameters(18).Value = registro.Substring(359, 8)
                        insertCommand.Parameters(19).Value = registro.Substring(367, 9)
                        insertCommand.Parameters(20).Value = registro.Substring(376, 6)
                        insertCommand.Parameters(21).Value = registro.Substring(382, 1)
                        insertCommand.Parameters(22).Value = registro.Substring(383, 4)
                        insertCommand.Parameters(23).Value = registro.Substring(387, 4)
                        insertCommand.Parameters(24).Value = registro.Substring(391, 2)
                        insertCommand.Parameters(25).Value = registro.Substring(393, 1)
                        insertCommand.Parameters(26).Value = registro.Substring(394, 4)
                        insertCommand.Parameters(27).Value = registro.Substring(398, 1)
                        insertCommand.Parameters(28).Value = registro.Substring(399, 1)
                        insertCommand.Parameters(29).Value = registro.Substring(400, 1)
                        insertCommand.Parameters(30).Value = registro.Substring(401, 1)
                        insertCommand.Parameters(31).Value = registro.Substring(402, 20)
                        insertCommand.Parameters(32).Value = registro.Substring(422, 20)
                        insertCommand.Parameters(33).Value = registro.Substring(442, 55)
                        insertCommand.Parameters(34).Value = registro.Substring(497, 40)
                        insertCommand.Parameters(35).Value = registro.Substring(537, 5)
                        insertCommand.Parameters(36).Value = registro.Substring(542, 25)
                        insertCommand.Parameters(37).Value = registro.Substring(567, 30)
                        insertCommand.Parameters(38).Value = registro.Substring(597, 2)
                        insertCommand.Parameters(39).Value = registro.Substring(599, 8)
                        insertCommand.Parameters(40).Value = registro.Substring(607, 40)
                        insertCommand.Parameters(41).Value = registro.Substring(647, 1)
                        insertCommand.Parameters(42).Value = registro.Substring(648, 2)
                        insertCommand.Parameters(43).Value = registro.Substring(650, 40)
                        insertCommand.Parameters(44).Value = registro.Substring(690, 40)
                        insertCommand.Parameters(45).Value = registro.Substring(730, 15)
                        insertCommand.Parameters(46).Value = registro.Substring(745, 2)
                        insertCommand.Parameters(47).Value = registro.Substring(747, 2)
                        insertCommand.Parameters(48).Value = registro.Substring(749, 16)
                        insertCommand.Parameters(49).Value = registro.Substring(765, 16)

                        ' Ejecutar la consulta de inserción
                        insertCommand.ExecuteNonQuery()
                    Next

                End Using
                MessageBox.Show("Datos cargados exitosamente en la tabla TarjetasSanitariasDiario.", "Operación Exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception

                MessageBox.Show($"Error al cargar los registros: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using


    End Sub


End Class
