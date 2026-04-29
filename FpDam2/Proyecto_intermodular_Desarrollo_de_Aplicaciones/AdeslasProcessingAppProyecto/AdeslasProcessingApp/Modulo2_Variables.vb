Imports System.IO
Imports System.Configuration

Module Variables

    ' 📍 Carpeta donde está el EXE (instalación)
    Public RutaPrincipal As String = Application.StartupPath

    ' 📁 Carpeta base (una superior al EXE)
    'Public RutaBase As String = Directory.GetParent(RutaPrincipal).FullName
    Public RutaBase As String = Directory.GetParent(RutaPrincipal).Parent.Parent.Parent.FullName
    ' 📤 Carpeta de salida
    'Public carpetaSalida As String = Path.Combine(RutaBase, "Ficheros_Salida")
    Public carpetaSalida As String = Path.Combine(RutaBase, "Ficheros_Salida")

    ' 📥 Carpeta BD (MISMA CARPETA QUE EL EXE)
    Public RutaBD As String = RutaPrincipal

    ' 🔹 Variables varias
    Public NombreFichero As String
    Public NombreFicheroCompleto As String
    Public nombreUsuario As String
    Public TextBox1 As String

    Public tarjetasSanitarias As New List(Of String)()
    Public tarjetasLaCaixa As New List(Of String)()
    Public tarjetasDentales As New List(Of String)()

    ' 🔹 Inicialización (crear carpetas automáticamente)
    Public Sub InicializarRutas()

        Try
            ' Crear carpeta salida si no existe
            If Not Directory.Exists(carpetaSalida) Then
                Directory.CreateDirectory(carpetaSalida)
            End If

            ' Crear subcarpeta reporte
            Dim rutaReporte As String = Path.Combine(carpetaSalida, "reporte")
            If Not Directory.Exists(rutaReporte) Then
                Directory.CreateDirectory(rutaReporte)
            End If

        Catch ex As Exception
            MessageBox.Show("Error creando carpetas: " & ex.Message,
                            "Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error)
        End Try

    End Sub

End Module