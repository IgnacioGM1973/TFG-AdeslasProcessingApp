'Realizado por: Ignacio Guijarro Melgar 
'En proceso. NO ESTA ENPRODUCCIÓN

Imports System.Data.OleDb
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports DocumentFormat.OpenXml.Bibliography

Public Class IncioForms

    Private selectedFiles As String()

    Private Sub ArchivoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ArchivoToolStripMenuItem.Click
        ProgressBar1.Value = 0
        ProgressBar1.Visible = False
        ProgressBar2.Value = 0
        ProgressBar2.Visible = True
    End Sub

    Private Sub IncioForms_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ProgressBar1.Visible = False
        ProgressBar2.Visible = False
    End Sub


    Private Sub MantenimientoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MantenimientoToolStripMenuItem.Click
        ProgressBar1.Value = 0
        ProgressBar1.Visible = True
        ProgressBar2.Value = 0
        ProgressBar2.Visible = False
    End Sub

    Private Sub ImportarDatosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportarDatosToolStripMenuItem.Click
        ProgressBar1.Value = 0
        ProgressBar1.Visible = True
        ProgressBar2.Value = 0
        ProgressBar2.Visible = False
    End Sub

    Private Sub SalirToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalirToolStripMenuItem.Click
        End
    End Sub

    Private Sub SeleccionarFicheroToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SeleccionarFicheroToolStripMenuItem.Click
        ' Crear una instancia de OpenFileDialog
        Dim openFileDialog As New OpenFileDialog()

        ' Configurar propiedades del OpenFileDialog
        openFileDialog.Title = "Seleccionar uno o más ficheros"
        openFileDialog.Filter = "Archivos de texto|*.txt;*.csv"
        openFileDialog.InitialDirectory = RutaPrincipal & "Entrada" ' Ruta inicial por defecto
        openFileDialog.Multiselect = True ' Habilitar selección múltiple

        ' Mostrar el cuadro de diálogo y verificar si el usuario seleccionó archivos
        If openFileDialog.ShowDialog() = DialogResult.OK Then
            ' Obtener las rutas de los archivos seleccionados
            Me.selectedFiles = openFileDialog.FileNames

            ' Mostrar solo los nombres de los archivos entre comillas en el TextBox
            Dim fileNames As String() = Me.selectedFiles.Select(Function(file) $" ({Path.GetFileName(file)}) ").ToArray()
            TextBox1.Text = String.Join(Environment.NewLine, fileNames)
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim validarNombre As Boolean = True
        ProgressBar2.Visible = True
        ' Verificar si se han seleccionado archivos
        If selectedFiles Is Nothing OrElse selectedFiles.Length = 0 Then
            MessageBox.Show("Por favor, seleccione uno o más ficheros antes de ejecutar el procesamiento.",
                    "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Procesar cada archivo seleccionado
        For Each filePath In selectedFiles
            'cargar el nombre del fichero para componer el fichero de salida


            NombreFichero = Path.GetFileNameWithoutExtension(filePath) & ".txt"


            If NombreFichero.Substring(0, 1) = "W" Then
                NombreFichero = NombreFichero.Replace(".txt", ".CSV")
            End If


            NombreFicheroCompleto = filePath
            ProgressBar2.Value = 0

            ' Controlar que los ficheros que manda el departamento de suscripcion del día esten importados.
            ' Si no estan, preguntar si se desea continuar con el proceso al departamanto de suscripción.
            ' Puede que no existan actualizaciones ese día y no se tenga que importar nada. Nos tienen que avisar o preguntar.
            '--------------------------------------------------------------------------------------------------------------------------------------------------------------------
            validarNombre = VerificarTablasSuscripcionActualizadas()
            If validarNombre = False Then
                ' Si la validación falla, saltar al siguiente fichero
                MessageBox.Show($"Procesamiento descartado. Proceda a la actualización de las tablas.",
                    "Finalización", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit For
            End If
            '--------------------------------------------------------------------------------------------------------------------------------------------------------------------
            'Validacion del nombre del fichero para que cumpla con el formato esperado
            '--------------------------------------------------------------------------------------------------------------------------------------------------------------------
            validarNombre = ValidarNombresFicheros(NombreFichero)
            If validarNombre = False Then
                ' Si la validación falla, saltar al siguiente fichero
                MessageBox.Show($"Procesamiento descartado para el archivo: {filePath}",
                    "Finalización", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Continue For
            End If
            '--------------------------------------------------------------------------------------------------------------------------------------------------------------------

            Modulo4_CargarRegistrosPorModelo.CargarRegistros(filePath)
            'Modulo7_SeparacionProductos_T_S_Adeslas.ProcesarDatosFuncionarios_ISFAS_MJ()


            MessageBox.Show($"Procesamiento completado para el archivo: {filePath}",
                    "Finalización", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Next
    End Sub
    'Comprueba que las tablas esten actualizadas
    Private Function VerificarTablasSuscripcionActualizadas() As Boolean
        ' Ruta de la base de datos
        'Dim dbPath As String = RutaPrincipal & "\bd\adeslas.accdb"
        'Dim connString As String = $"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={dbPath}"

        ' Fecha actual
        Dim fechaActual As DateTime = DateTime.Now.Date
        Dim tablasConFechaCorrecta As New List(Of String)
        ' Conexión a la base de datos mediante la varable publica del modulo conexion accessConnString
        Try
            Using connection As New OleDb.OleDbConnection(accessConnString)
                connection.Open()

                ' Consultas por cada tabla
                Dim querySelectColectivosConDireccionEnvio As String = "SELECT Fecha FROM ColectivosConDireccionEnvio WHERE Fecha = @FechaActual"
                Dim querySelectColectivosSinAsistenciaViaje As String = "SELECT Fecha FROM ColectivosSinAsistenciaViaje WHERE Fecha = @FechaActual"
                Dim querySelectMGA As String = "SELECT Fecha FROM MGA WHERE Fecha = @FechaActual"
                Dim querySelectColectivosSinAsistenciaViaje2 As String = "SELECT Fecha FROM TarjetasDentalesPymes WHERE Fecha = @FechaActual"
                Dim querySelectTarjetasNoProcesar As String = "SELECT Fecha FROM TarjetasNoProcesar WHERE Fecha = @FechaActual"

                ' Comprobar cada tabla
                If Not VerificarFechaEnTabla(connection, querySelectColectivosConDireccionEnvio, fechaActual) Then
                    tablasConFechaCorrecta.Add("ColectivosConDireccionEnvio")
                End If
                If Not VerificarFechaEnTabla(connection, querySelectColectivosSinAsistenciaViaje, fechaActual) Then
                    tablasConFechaCorrecta.Add("ColectivosSinAsistenciaViaje")
                End If
                If Not VerificarFechaEnTabla(connection, querySelectMGA, fechaActual) Then
                    tablasConFechaCorrecta.Add("MGA")
                End If
                If Not VerificarFechaEnTabla(connection, querySelectColectivosSinAsistenciaViaje2, fechaActual) Then
                    tablasConFechaCorrecta.Add("TarjetasDentalesPymes")
                End If
                If Not VerificarFechaEnTabla(connection, querySelectTarjetasNoProcesar, fechaActual) Then
                    tablasConFechaCorrecta.Add("TarjetasNoProcesar")
                End If


            End Using

            ' Si alguna tabla no tiene la fecha correcta, mostrar cuál es
            If tablasConFechaCorrecta.Any() Then
                Dim tablasErroneas As String = String.Join(", ", tablasConFechaCorrecta)
                Dim resultado As DialogResult = MessageBox.Show(
                $"Las Siguientes tablas no han sido actualizadas en el día de hoy: {tablasErroneas}. ¿Deseas continuar con el proceso o proceder a su actualización?",
                "Fecha Incorrecta",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning)

                ' Si el usuario elige "No", detener el proceso
                If resultado = DialogResult.No Then
                    Return False
                End If
            End If


            Return True

        Catch ex As Exception
            ' Manejo de errores
            MessageBox.Show($"Error al acceder a la base de datos: {ex.Message}",
                        "Error de Base de Datos",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)
            Return False
        End Try
    End Function
    'Comprueba que la fecha este actualizada
    Private Function VerificarFechaEnTabla(connection As OleDb.OleDbConnection, query As String, fecha As DateTime) As Boolean
        ' Formatear la fecha a dd/MM/yyyy
        Dim fechaFormateada As String = fecha.ToString("dd/MM/yyyy")

        Using cmdSelect As New OleDb.OleDbCommand(query, connection)
            ' Usar la fecha formateada como parámetro
            cmdSelect.Parameters.AddWithValue("@FechaActual", fechaFormateada)

            Using reader As OleDb.OleDbDataReader = cmdSelect.ExecuteReader()
                ' Si hay filas, la fecha es correcta en esta tabla
                Return reader.HasRows
            End Using
        End Using
    End Function
    'Valida el formato del nombre para los ficheros de entrada
    Private Function ValidarNombresFicheros(NombreFichero As String) As Boolean
        ' Definición de patrones según el tipo de archivo
        Dim patrones As New Dictionary(Of String, String) From {
        {"Sanitario", "^T\d{7}\.txt$"}, ' TYYMMDD.txt
        {"Dental", "^D\d{7}\.txt$"},    ' DYYMMDD.txt
        {"Cartera", "^W\d{7}W\.CSV$"},  ' WYYMMDDW.CSV
        {"Miracle", "^W\d{7}M\.CSV$"}   ' WYYMMDDM.CSV
    }

        ' Validar si el formato del fichero es correcto
        Dim tipoValido As Boolean = False
        For Each patron In patrones
            If Regex.IsMatch(NombreFichero, patron.Value, RegexOptions.IgnoreCase) Then
                tipoValido = True
                Exit For
            End If
        Next

        If Not tipoValido Then
            MessageBox.Show($"El fichero '{NombreFichero}' no cumple con el formato esperado.",
                    "Error de Validación",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)
            Return False
        End If

        ' Ruta de la base de datos
        'Dim dbPath As String = RutaPrincipal & "\bd\ADESLAS.accdb"
        'Dim connString As String = $"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={dbPath}"
        ' Conexión a la base de datos mediante la varable publica del modulo conexion accessConnString
        Try
            Using connection As New OleDb.OleDbConnection(accessConnString)
                connection.Open()

                ' Extraer el inicio del nombre del fichero (primeros 7 caracteres)
                Dim inicioNombre As String = NombreFichero.Substring(0, 7)
                Dim querySelect As String = "SELECT NombreFichero FROM HistoricoNombreFicheros WHERE LEFT(NombreFichero, 7) = @InicioNombre"

                Dim nombresEncontrados As New List(Of String)
                Using cmdSelect As New OleDb.OleDbCommand(querySelect, connection)
                    cmdSelect.Parameters.AddWithValue("@InicioNombre", inicioNombre)

                    Using reader As OleDb.OleDbDataReader = cmdSelect.ExecuteReader()
                        While reader.Read()
                            nombresEncontrados.Add(reader("NombreFichero").ToString())
                        End While
                    End Using
                End Using

                ' Verificar qué números entre 0 y 9 están disponibles
                Dim numerosDisponibles As New List(Of Integer)(Enumerable.Range(0, 10))
                For Each nombre In nombresEncontrados
                    Dim numeroFinal As String = nombre.Substring(7, 1)
                    If Integer.TryParse(numeroFinal, Nothing) Then
                        numerosDisponibles.Remove(Convert.ToInt32(numeroFinal))
                    End If
                Next

                If nombresEncontrados.Contains(NombreFichero) Then
                    If numerosDisponibles.Count > 0 Then
                        Dim opcionesDisponibles As String = String.Join(", ", numerosDisponibles.Select(Function(n) NombreFichero.Substring(0, 7) & n & ".txt"))
                        MessageBox.Show($"El fichero '{NombreFichero}' ha sido procesado." & vbCrLf &
                                $"Puedes renombrarlo con uno de estos nombres disponibles si quieres que tenga la misma fecha: {NombreFichero.Substring(5, 2) & "/" & NombreFichero.Substring(3, 2) & "/" & NombreFichero.Substring(1, 2)} -- {opcionesDisponibles}",
                                "Fichero Duplicado",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information)
                    Else
                        MessageBox.Show($"El fichero '{NombreFichero}' ya existe en la base de datos y no se pueden generar más opciones con la misma fecha '{NombreFichero.Substring(1, 6)}_'. Por favor, elija una fecha diferente y un número de secuencia único para renombrar el archivo entre 0 y 9.",
                                "Fichero Duplicado",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning)

                    End If
                    Return False
                End If

                '(activarlo cuando este opartivo)
                'Anexar el nombre del fichero en el HistoricoNombreFicheros
                ' Si no existe, anexa el nuevo nombre del fichero en la tabla HistoricoNombreFicheros 
                '--------------------------------------------------------------------------------------------------------------------------
                Dim queryInsert As String = "INSERT INTO HistoricoNombreFicheros (NombreFichero, Fecha) VALUES (@NombreFichero, @Fecha)"
                Using cmdInsert As New OleDb.OleDbCommand(queryInsert, connection)
                    cmdInsert.Parameters.AddWithValue("@NombreFichero", NombreFichero)

                    ' Formatear la fecha actual al formato ISO
                    Dim fechaFormateada As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                    cmdInsert.Parameters.AddWithValue("@Fecha", fechaFormateada)

                    cmdInsert.ExecuteNonQuery()
                End Using
                '--------------------------------------------------------------------------------------------------------------------------
            End Using

                ' Mostrar mensaje de éxito
                MessageBox.Show($"El fichero '{NombreFichero}' se ha registrado correctamente en la base de datos.",
                    "Registro Exitoso",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information)
            Return True
        Catch ex As Exception
            ' Manejo de errores
            MessageBox.Show($"Error al acceder a la base de datos: {ex.Message}",
                    "Error de Base de Datos",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)
            Return False
        End Try
    End Function
    'Opciones menu importación
    Private Sub SeleccionarFicheroToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SeleccionarFicheroToolStripMenuItem1.Click
        ' Crear una instancia de OpenFileDialog
        ' Configurar propiedades del OpenFileDialog
        Dim openFileDialog As New OpenFileDialog With {
            .Title = "Seleccionar un fichero",
            .Filter = "Archivos Excel|*300454.xlsx",
            .InitialDirectory = RutaPrincipal & "\Entrada"
        }

        ' Mostrar el cuadro de diálogo y verificar si el usuario seleccionó un archivo
        If openFileDialog.ShowDialog() = DialogResult.OK Then
            ' Obtener la ruta del archivo seleccionado
            Dim ExcelFilePath As String = openFileDialog.FileName
            Dim selectedFileName As String = Path.GetFileName(ExcelFilePath)

            ' Verificar si el nombre del archivo es "300454.xlsx"
            If Not selectedFileName.Equals("300454.xlsx", StringComparison.OrdinalIgnoreCase) Then
                MessageBox.Show("El nombre del archivo debe ser '300454.xlsx'. Por favor, seleccione el archivo correcto.",
                                "Error de archivo",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error)
                Return
            End If

            ' Mostrar el nombre del archivo en el TextBox
            TextBox1.Text = selectedFileName

            ' Mensaje de confirmación
            Dim result As DialogResult = MessageBox.Show("¿Deseas continuar con la importación?",
                                                     "Confirmación",
                                                     MessageBoxButtons.YesNo,
                                                     MessageBoxIcon.Question)
            ' Verificar la respuesta del usuario
            If result = DialogResult.Yes Then
                ' Continuar con la importación
                Modulo3_ExportarFicherosSucComple.ImportarExcel_A_Access_MGA(ExcelFilePath, ProgressBar1)
            Else
                ' Cancelar el proceso
                MessageBox.Show("El proceso ha sido cancelado.", "Cancelado", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End If
    End Sub
    Private Sub SinAsistenciaViajeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SinAsistenciaViajeToolStripMenuItem.Click
        ' Crear una instancia de OpenFileDialog
        ' Configurar propiedades del OpenFileDialog
        Dim openFileDialog As New OpenFileDialog With {
            .Title = "Seleccionar un fichero",
            .Filter = "Archivos Excel|COLECTIVOS_SINASISTENCIAVIAJE.xlsx",
            .InitialDirectory = RutaPrincipal & "\Entrada"
        }

        ' Mostrar el cuadro de diálogo y verificar si el usuario seleccionó un archivo
        If openFileDialog.ShowDialog() = DialogResult.OK Then
            ' Obtener la ruta del archivo seleccionado
            Dim ExcelFilePath As String = openFileDialog.FileName
            Dim selectedFileName As String = Path.GetFileName(ExcelFilePath)

            ' Verificar si el nombre del archivo es "COLECTIVOS_SINASISTENCIAVIAJE.xlsx"
            If Not selectedFileName.Equals("COLECTIVOS_SINASISTENCIAVIAJE.xlsx", StringComparison.OrdinalIgnoreCase) Then
                MessageBox.Show("El nombre del archivo debe ser 'COLECTIVOS_SINASISTENCIAVIAJE.xlsx'. Por favor, seleccione el archivo correcto.",
                            "Error de archivo",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error)
                Return
            End If

            ' Mostrar el nombre del archivo en el TextBox
            TextBox1.Text = Path.GetFileName(ExcelFilePath)

            ' Mensaje de confirmación
            Dim result As DialogResult = MessageBox.Show("¿Deseas continuar con la importación?",
                                                         "Confirmación",
                                                         MessageBoxButtons.YesNo,
                                                         MessageBoxIcon.Question)

            ' Verificar la respuesta del usuario
            If result = DialogResult.Yes Then
                ' Continuar con la importación

                Modulo3_ExportarFicherosSucComple.ImportarExcel_A_Access_SinAsistenciaViaje(ExcelFilePath, ProgressBar1)

            Else
                ' Cancelar el proceso
                MessageBox.Show("El proceso ha sido cancelado.", "Cancelado", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End If
    End Sub
    Private Sub ImportarSeniorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportarSeniorToolStripMenuItem.Click
        ' Crear OpenFileDialog para seleccionar múltiples archivos de texto
        Dim openFileDialog As New OpenFileDialog With {
        .Title = "Seleccionar archivos de texto",
        .Filter = "Archivos de texto|*Senior.txt",
        .InitialDirectory = RutaPrincipal & "\Entrada",
        .Multiselect = True ' Permitir seleccionar múltiples archivos
    }

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Dim filePaths As String() = openFileDialog.FileNames ' Obtener todos los archivos seleccionados

            ' Confirmación del usuario
            Dim result As DialogResult = MessageBox.Show($"¿Desea importar los registros de los {filePaths.Length} archivos seleccionados?",
                                                     "Confirmar Importación",
                                                     MessageBoxButtons.YesNo,
                                                     MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                Try
                    For Each filePath As String In filePaths
                        ' Llamar a la función de importación para cada archivo
                        Modulo3_ExportarFicherosSucComple.ImportarSenior(filePath)
                    Next

                    MessageBox.Show($"Archivos procesados: {filePaths.Length}",
                                "Éxito",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End If
    End Sub
    Private Sub ImnportarEnviosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImnportarEnviosToolStripMenuItem.Click

        Dim openFileDialog As New OpenFileDialog With {
        .Title = "Seleccionar un fichero",
        .Filter = "Archivos Excel|envio_*.xlsx",
        .InitialDirectory = RutaPrincipal & "\Entrada"
    }

        If openFileDialog.ShowDialog() = DialogResult.OK Then

            Dim ExcelFilePath As String = openFileDialog.FileName
            Dim selectedFileName As String = Path.GetFileName(ExcelFilePath)

            ' Validación correcta: debe comenzar por "envio_" y terminar en ".xlsx"
            If Not selectedFileName.StartsWith("envio_", StringComparison.OrdinalIgnoreCase) _
            OrElse Not selectedFileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) Then

                MessageBox.Show("El nombre del archivo debe ser 'envio_*.xlsx'. Por favor, seleccione el archivo correcto.",
                            "Error de archivo",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error)
                Return
            End If

            TextBox1.Text = selectedFileName

            Dim result As DialogResult = MessageBox.Show("¿Deseas continuar con la importación?",
                                                     "Confirmación",
                                                     MessageBoxButtons.YesNo,
                                                     MessageBoxIcon.Question)

            If result = DialogResult.Yes Then
                Modulo3_ExportarFicherosSucComple.ImportarExcel_A_Access_ColectivosConDireccionEnvio(ExcelFilePath, ProgressBar1)
            Else
                MessageBox.Show("El proceso ha sido cancelado.", "Cancelado", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

        End If
    End Sub

    Private Sub ImportarEmpresaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportarEmpresaToolStripMenuItem.Click
        ' Crear una instancia de OpenFileDialog
        Dim openFileDialog As New OpenFileDialog With {
        .Title = "Seleccionar un fichero",
        .Filter = "Archivos Excel (*.xlsb)|*_TARJETAS_DENTALES_PYMES.XLSB",
        .InitialDirectory = RutaPrincipal & "\Entrada"
    }

        ' Mostrar el cuadro de diálogo y verificar si el usuario seleccionó un archivo
        If openFileDialog.ShowDialog() = DialogResult.OK Then
            ' Obtener la ruta completa del archivo seleccionado
            Dim ExcelFilePath As String = openFileDialog.FileName
            Dim selectedFileName As String = Path.GetFileName(ExcelFilePath)

            ' Validar el nombre del archivo usando un filtro basado en el patrón esperado
            Dim regexPattern As String = "^\d{8}_TARJETAS_DENTALES_PYMES\.XLSB$" ' Formato esperado: yyyyMMdd_TARJETAS_DENTALES_PYMES.XLSB
            If Not System.Text.RegularExpressions.Regex.IsMatch(selectedFileName, regexPattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then
                MessageBox.Show("El nombre del archivo debe ser 'yyyymmdd_TARJETAS_DENTALES_PYMES.XLSB'. Por favor, seleccione el archivo correcto.",
                            "Error de archivo",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error)
                Return
            End If

            ' Mostrar el nombre del archivo en el TextBox
            TextBox1.Text = selectedFileName

            ' Mensaje de confirmación
            Dim result As DialogResult = MessageBox.Show("¿Deseas continuar con la importación?",
                                                     "Confirmación",
                                                     MessageBoxButtons.YesNo,
                                                     MessageBoxIcon.Question)

            ' Verificar la respuesta del usuario
            If result = DialogResult.Yes Then
                ' Llamar a la función de importación
                Modulo3_ExportarFicherosSucComple.ImportarExcel_A_Access_TARJETAS_DENTALES_PYMES(ExcelFilePath, ProgressBar1)
            Else
                ' Cancelar el proceso
                MessageBox.Show("El proceso ha sido cancelado.", "Cancelado", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End If
    End Sub
    Private Sub ImportarTarjetasNoProcesarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportarTaarjetasNoProcesarToolStripMenuItem.Click

        ' Cuadro de diálogo para seleccionar los Excel que empiezan por tarjetas_noprocesar
        Dim openFileDialog As New OpenFileDialog With {
        .Title = "Seleccionar un fichero",
        .Filter = "Archivos Excel|tarjetas_noprocesar*.xlsx",
        .InitialDirectory = RutaPrincipal & "\Entrada"
    }

        If openFileDialog.ShowDialog() = DialogResult.OK Then

            Dim ExcelFilePath As String = openFileDialog.FileName
            Dim selectedFileName As String = Path.GetFileName(ExcelFilePath)

            ' Validación: debe empezar por tarjetas_noprocesar y terminar en .xlsx
            If Not selectedFileName.StartsWith("tarjetas_noprocesar", StringComparison.OrdinalIgnoreCase) _
            OrElse Not selectedFileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) Then

                MessageBox.Show("El nombre del archivo debe ser 'tarjetas_noprocesar*.xlsx'. Selecciona un archivo válido.",
                            "Error de archivo",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error)
                Return
            End If

            TextBox1.Text = selectedFileName

            Dim result As DialogResult = MessageBox.Show("¿Deseas continuar con la importación?",
                                                     "Confirmación",
                                                     MessageBoxButtons.YesNo,
                                                     MessageBoxIcon.Question)

            If result = DialogResult.Yes Then
                Modulo3_ExportarFicherosSucComple.ImportarTarjetasNoProcesar(ExcelFilePath, ProgressBar1, eqT_Informativa)
            Else
                MessageBox.Show("El proceso ha sido cancelado.", "Cancelado", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End If

    End Sub

    Private Sub ImportarHispapostToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ImportarHispapostToolStripMenuItem1.Click
        ' Crear una instancia de OpenFileDialog

        Dim openFileDialog As New OpenFileDialog With {
        .Title = "Seleccionar un fichero",
        .Filter = "Archivos Excel (*.xls;*.xlsx)|*.xls;*.xlsx",
        .InitialDirectory = RutaPrincipal & "\Entrada"
    }

        ' Mostrar el cuadro de diálogo y verificar si el usuario seleccionó un archivo
        If openFileDialog.ShowDialog() = DialogResult.OK Then
            ' Obtener la ruta completa del archivo seleccionado
            Dim ExcelFilePath As String = openFileDialog.FileName
            Dim selectedFileName As String = Path.GetFileName(ExcelFilePath)

            ' Mostrar el nombre del archivo en el TextBox
            TextBox1.Text = selectedFileName

            ' Mensaje de confirmación
            Dim result As DialogResult = MessageBox.Show("¿Deseas continuar con la importación?",
                                                     "Confirmación",
                                                     MessageBoxButtons.YesNo,
                                                     MessageBoxIcon.Question)

            ' Verificar la respuesta del usuario
            If result = DialogResult.Yes Then
                ' Llamar a la función de importación
                Modulo3_ExportarFicherosSucComple.ImportarHispapost(ExcelFilePath, ProgressBar1)
            Else
                ' Cancelar el proceso
                MessageBox.Show("El proceso ha sido cancelado.", "Cancelado", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Modulo7_SeparacionProductos_T_S_Adeslas.NombreFicheroEntrada = TextBox1.Text.Trim()
        Modulo7_SeparacionProductos_T_S_Adeslas.EjecutarSeparacion()

        'Modulo7_SeparacionProductos_T_S_Adeslas.ProcesarDatosFuncionarios_ISFAS_MJ()
        'Modulo7_SeparacionProductos_T_S_Adeslas.ProcesarDatosLogos()
        'Modulo7_SeparacionProductos_T_S_Adeslas.MarcarPaquetizados()
        'Modulo7_SeparacionProductos_T_S_Adeslas.MarcarLogos()
        'Modulo7_SeparacionProductos_T_S_Adeslas.MarcarTipoContraPymes()
        'Modulo7_SeparacionProductos_T_S_Adeslas.SepararPorCodigoPlastico()
        'Modulo7_SeparacionProductos_T_S_Adeslas.SepararTarjeta00()

        MessageBox.Show("Proceso finalizado", "Finalizado", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub


End Class

Public Class Tarjetas
    Public Shared Sub TarjetasSanitariasAdeslas()

        Dim lis_Tar_Sanitarias_Adeslas As New List(Of String)


        MessageBox.Show($"Procesamiento Tarjetas sanitaria Adeslas",
                    "Finalización", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    Private Sub TarjetasSanitariasCaixa()
        MessageBox.Show($"Procesamiento Tarjetas sanitaria Vida Caixa",
                    "Finalización", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub TarjetasDentalesAdeslas()
        MessageBox.Show($"Procesamiento Tarjetas dentales Adeslas",
                    "Finalización", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub TarjetasDentalesVidaCixa()
        MessageBox.Show($"Procesamiento Tarjetas dentales VidaCaixa",
                    "Finalización", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub



End Class