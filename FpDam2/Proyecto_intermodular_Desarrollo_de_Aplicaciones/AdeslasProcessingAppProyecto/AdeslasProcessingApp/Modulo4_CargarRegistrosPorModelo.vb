Imports System.Text
Imports System.IO

Public Class Modulo4_CargarRegistrosPorModelo
    Public Shared Sub CargarRegistros(filePath As String)
        'Se realiza un filtrado previo al la carga de los registros con los siguientes requerimientos:
        'tipocontra prohibidos
        '-------------------------

        '1) tipocontra >=  "4001" And tipocontra <= "4027" 
        ' 2) tipocontra  = 991, tipocontra  = 967, tipocontra  = 825 

        '------------------------------------------------------
        '3)NO PROCESAR si cumple estas condiciones (duplicados de pólizas privadas en fichero gral)

        'es fichero gral, ncertifica = 0, extrac = "W" 

        'tipocontra = 829
        'tipocontra = 828
        'tipocontra = 834
        'tipocontra = 835
        'tipocontra = 830
        'tipocontra = 2118
        'tipocontra = 2120
        'tipocontra = 2121
        'tipocontra = 2122
        'tipocontra = 2123
        'tipocontra = 1026
        'tipocontra = 1027
        'tipocontra = 1028
        'tipocontra = 1029
        'tipocontra = 831
        'tipocontra = 832
        'tipocontra = 833
        'tipocontra = 1092
        'tipocontra = 979
        'tipocontra = 1080
        'tipocontra = 1712
        'tipocontra = 1713
        'tipocontra = 1714
        'tipocontra = 1715
        'tipocontra = 2706
        'tipocontra = 2707
        'tipocontra = 2708
        'tipocontra = 2709

        ' Inicializar listas
        Dim tarjetasSanitarias As New List(Of String)()
        Dim tarjetasLaCaixa As New List(Of String)()
        Dim tarjetasDentales As New List(Of String)()

        Dim registrosMarcados2 As New List(Of String)()

        Dim swGeneral As Boolean = filePath.EndsWith("1.txt")

        Try
            For Each line As String In File.ReadLines(filePath, Encoding.Default)
                Dim longitud As Integer = line.Length
                Dim tipocontra As String = line.Substring(394, 4).Trim()

                ' Validaciones
                If tipocontra >= "4001" AndAlso tipocontra <= "4027" Then Continue For
                If {"0991", "0967", "0825"}.Contains(tipocontra) Then Continue For

                Dim duplicados As New HashSet(Of String) From {"0829", "0828", "0834", "0835", "0830", "2118", "2120", "2121", "2122", "2123", "1026", "1027", "1028", "1029", "0831", "0832", "0833", "01092", "0979", "1080", "1712", "1713", "1714", "1715", "2706", "2707", "2708", "2709"}
                If swGeneral AndAlso line.Substring(401, 1).Trim() = "W" AndAlso line.Substring(12, 9).Trim() = "0" AndAlso duplicados.Contains(tipocontra) Then Continue For

                ' Clasificación por longitud
                Select Case longitud
                    Case 781
                        tarjetasSanitarias.Add(line)
                    Case 538
                        tarjetasLaCaixa.Add(line)
                    Case 537
                        tarjetasDentales.Add(line)
                End Select
            Next
        Catch ex As Exception
            Console.WriteLine($"Error al cargar registros desde el archivo: {ex.Message}")
        End Try

        'Carga de registros Sanitarios
        If tarjetasSanitarias.Count > 0 Then

            registrosMarcados2 = Modulo5_ValidarRegistrosSanitarios.BajaNumeroPolizaCertificadoorAccess(tarjetasSanitarias)

        End If

        'Carga de registros LaCaixa pendiente

        'Carga de registros Dental pendiente


    End Sub
End Class
