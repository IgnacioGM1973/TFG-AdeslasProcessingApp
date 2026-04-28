Imports System.Data.OleDb
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Text.RegularExpressions
Imports ClosedXML.Excel
Imports DocumentFormat.OpenXml.Wordprocessing
Imports Microsoft.Win32
Imports SixLabors.Fonts.Tables

Public Module Modulo7_SeparacionProductos_T_S_Adeslas

    Public NombreFicheroEntrada As String = ""


    ' ============================================================
    ' CONEXIÓN
    ' ============================================================
    Private ReadOnly ConnectionString As String =
        "Provider=Microsoft.ACE.OLEDB.12.0;" &
        "Data Source=" & RutaBD & "ADESLAS.accdb;" &
        "Persist Security Info=False;"
    ' ============================================================
    ' Pólizas SIN asistencia viaje (en memoria)
    ' ============================================================
    Private polizasSinAsistenciaViaje As HashSet(Of String)


    ' ============================================================
    ' AUXILIARES
    ' ============================================================
    Private Function Texto(v As Object) As String
        If v Is Nothing OrElse IsDBNull(v) Then Return ""
        Return v.ToString().Trim().ToUpperInvariant()
    End Function
    Private Function TextoRaw(v As Object) As String
        If v Is Nothing OrElse IsDBNull(v) Then Return ""
        Return v.ToString().Trim()
    End Function


    Private Function LeerBool(v As Object) As Boolean
        If v Is Nothing OrElse IsDBNull(v) Then Return False

        Dim s As String = v.ToString().Trim().ToUpperInvariant()

        Return s = "X" OrElse s = "-1" OrElse s = "1" _
        OrElse s = "TRUE" OrElse s = "SI" OrElse s = "SÍ"
    End Function


    ' ============================================================
    ' CLASE REGLA
    ' ============================================================
    Private Class ReglaExtraccion
        Public Id As Integer
        Public Canal_Asegurador As String
        Public Tipo_Tarjeta As String
        Public Extraccion As String
        Public DescripcionTarjeta As String
        Public CodigoSalida As String
        Public ModeloPlasticoCodigo As String
        Public ColectivoProducto As String
        Public LogotipoDescripcion As String
        Public Direccion As String
        Public DireccionEnvio As String
        Public ModeloCarrier As String
        Public Paquetizado As String

        Public CodigoDelegacion As String
        Public NumeroPoliza As String

        Public BuscarLogo As Boolean
        Public OrdenacionEspecial As String
        Public SeparaCorreosHispapost As Boolean
        Public TipoContraPymes As Boolean
        Public IndicadorIdioma As String

        ' === NUEVOS CAMPOS DE SALIDA ===
        Public MODELO As String
        Public TOPPER As String
        Public LOGO As String
        Public ULTRAANV As String
        Public ULTRAREV As String
        Public CARRIER As String
        Public SOBRE As String
        Public TARJETA As String
        Public FOLLETO As String
        Public VC_PROD As String
        Public REV3 As String
        Public REV4 As String

        Public TipoContrato As String
        Public CentroTrabajoCodigo As String

    End Class



    ' ============================================================
    ' CLASE REGISTRO
    ' ============================================================
    Private Class RegistroTarjeta
        Public Id As Integer
        Public Canal_Asegurador As String
        Public NombreApellidos As String
        Public CodigoDelegacion As String
        Public NumeroPoliza As String
        Public NumeroCertificado As String
        Public CodigoRelacion As String
        Public NumeroOrden As String
        Public ModeloPlasticoCodigo As String
        Public ColectivoProducto As String
        Public Direccion As String
        Public CodigoPostal As String
        Public Poblacion As String
        Public Provincia As String
        Public TipoContraPymes As String
        Public LogotipoDescripcion As String
        Public HISPAPOST As String
        Public CodigoPostalEnvio As String
        Public PLATAFORMA As String
        Public ZONA As String
        Public DescripcionFinal As String
        Public CIP_SNS As String
        Public CIP_M As String
        Public Paquetizado As String
        Public Tipo_Tarjeta As String
        Public Extraccion As String
        Public AnoNacimiento As String
        Public Sexo As String
        Public FechaAlta As String
        Public NumeroTarjeta As String
        Public DigitoControlProvincia As String
        Public DigitoControlZ As String
        Public Version As String
        Public FechaCaducidad As String
        Public TipoContrato As String
        Public IndicadorIdioma As String
        Public IndicadorExtranjero As String
        Public FechaInicioCarencia As String
        Public PersonaReceptora As String
        Public DireccionEnvio As String
        Public PoblacionEnvio As String
        Public ProvinciaEnvio As String
        Public BARRAS_CONTROL As String
        Public ASESOR As String
        Public DIR_ASESOR As String
        Public TEL_ASESOR As String
        Public CP_ASESOR As String
        Public POB_ASESOR As String
        Public BENEF As String
        Public BEN_TARJ As String
        Public CPF As String
        Public IND1 As String
        Public IND2 As String
        Public TextoPersonalizado1 As String
        Public TextoPersonalizado2 As String
        Public CentroTrabajoCodigo As String




    End Class

    ' ============================================================
    ' LOGOS
    ' ============================================================
    Private Function CargarLogos() As Dictionary(Of String, String)

        Dim dic As New Dictionary(Of String, String)
        Dim rutaBDCompleta As String = Path.Combine(RutaBD, "ADESLAS.accdb")

        Dim connectionString As String =
            $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={rutaBDCompleta};Persist Security Info=False;"

        Using cn As New OleDbConnection(connectionString)

            cn.Open()

            Using cmd As New OleDbCommand(
                "SELECT POLIZA, CLAVE FROM Logos", cn)

                Using rd = cmd.ExecuteReader()
                    While rd.Read()
                        Dim poliza = Texto(rd("POLIZA"))
                        Dim clave = Texto(rd("CLAVE"))

                        If poliza <> "" AndAlso Not dic.ContainsKey(poliza) Then
                            dic.Add(poliza, clave)
                        End If
                    End While
                End Using
            End Using
        End Using

        Return dic
    End Function
    Private Function CargarLogosSalida() As Dictionary(Of String, ReglaExtraccion)

        Dim dic As New Dictionary(Of String, ReglaExtraccion)
        Dim rutaBDCompleta As String = Path.Combine(RutaBD, "ADESLAS.accdb")

        Dim connectionString As String =
            $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={rutaBDCompleta};Persist Security Info=False;"


        Using cn As New OleDbConnection(ConnectionString)
            cn.Open()

            Dim sql As String =
        "SELECT * FROM LogosSalida WHERE Activo = True"

            Using cmd As New OleDbCommand(sql, cn)
                Using rd = cmd.ExecuteReader()
                    While rd.Read()
                        Dim extraccion = Texto(rd("EXTRACCION"))
                        Dim claveLogo = Texto(rd("ClaveLogo"))
                        Dim modelo = Texto(rd("ModeloPlasticoCodigo")).PadLeft(2, "0"c)


                        ' Clave compuesta: LOGO|MODELO
                        Dim clave = extraccion & "|" & claveLogo & "|" & modelo

                        dic(clave) = New ReglaExtraccion With {
                        .MODELO = Texto(rd("MODELO")),
                        .TOPPER = Texto(rd("TOPPER")),
                        .LOGO = If(IsDBNull(rd("LOGO")), "", rd("LOGO").ToString().Trim()),
                        .ULTRAANV = Texto(rd("ULTRAANV")),
                        .ULTRAREV = Texto(rd("ULTRAREV")),
                        .CARRIER = Texto(rd("CARRIER")),
                        .SOBRE = Texto(rd("SOBRE")),
                        .TARJETA = Texto(rd("TARJETA")),
                        .FOLLETO = Texto(rd("FOLLETO")),
                        .VC_PROD = Texto(rd("VC_PROD")),
                        .REV3 = TextoRaw(rd("REV3")),
                        .REV4 = Texto(rd("REV4"))
                    }

                    End While
                End Using
            End Using
        End Using

        Return dic

    End Function


    Private Function CombinarReglaBaseConLogo(
    baseRegla As ReglaExtraccion,
    overrideLogo As ReglaExtraccion
) As ReglaExtraccion

        ' ✅ CLONAMOS la regla base (NO reutilizar referencia)
        Dim r As New ReglaExtraccion With {
.Id = baseRegla.Id,
        .Canal_Asegurador = baseRegla.Canal_Asegurador,
        .Tipo_Tarjeta = baseRegla.Tipo_Tarjeta,
        .Extraccion = baseRegla.Extraccion,
        .DescripcionTarjeta = baseRegla.DescripcionTarjeta,
        .CodigoSalida = baseRegla.CodigoSalida,
        .ModeloPlasticoCodigo = baseRegla.ModeloPlasticoCodigo,
        .ColectivoProducto = baseRegla.ColectivoProducto,
        .LogotipoDescripcion = baseRegla.LogotipoDescripcion,
        .Direccion = baseRegla.Direccion,
        .DireccionEnvio = baseRegla.DireccionEnvio,
        .ModeloCarrier = baseRegla.ModeloCarrier,
        .Paquetizado = baseRegla.Paquetizado,
        .CodigoDelegacion = baseRegla.CodigoDelegacion,
        .NumeroPoliza = baseRegla.NumeroPoliza,
        .BuscarLogo = baseRegla.BuscarLogo,
        .OrdenacionEspecial = baseRegla.OrdenacionEspecial,
        .SeparaCorreosHispapost = baseRegla.SeparaCorreosHispapost,
        .TipoContraPymes = baseRegla.TipoContraPymes,
        .TipoContrato = baseRegla.TipoContrato,
        .MODELO = baseRegla.MODELO,
        .TOPPER = baseRegla.TOPPER,
        .LOGO = baseRegla.LOGO,
        .ULTRAANV = baseRegla.ULTRAANV,
        .ULTRAREV = baseRegla.ULTRAREV,
        .CARRIER = baseRegla.CARRIER,
        .SOBRE = baseRegla.SOBRE,
        .TARJETA = baseRegla.TARJETA,
        .FOLLETO = baseRegla.FOLLETO,
        .VC_PROD = baseRegla.VC_PROD,
        .REV3 = baseRegla.REV3,
        .REV4 = baseRegla.REV4
    }

        If overrideLogo Is Nothing Then Return r

        ' Solo sobreescribimos si viene valor
        If overrideLogo.MODELO <> "" Then r.MODELO = overrideLogo.MODELO
        If overrideLogo.TOPPER <> "" Then r.TOPPER = overrideLogo.TOPPER
        If overrideLogo.LOGO <> "" Then r.LOGO = overrideLogo.LOGO
        If overrideLogo.ULTRAANV <> "" Then r.ULTRAANV = overrideLogo.ULTRAANV
        If overrideLogo.ULTRAREV <> "" Then r.ULTRAREV = overrideLogo.ULTRAREV
        If overrideLogo.CARRIER <> "" Then r.CARRIER = overrideLogo.CARRIER
        If overrideLogo.SOBRE <> "" Then r.SOBRE = overrideLogo.SOBRE
        If overrideLogo.TARJETA <> "" Then r.TARJETA = overrideLogo.TARJETA
        If overrideLogo.FOLLETO <> "" Then r.FOLLETO = overrideLogo.FOLLETO
        If overrideLogo.VC_PROD <> "" Then r.VC_PROD = overrideLogo.VC_PROD
        If overrideLogo.REV3 <> "" Then r.REV3 = overrideLogo.REV3
        If overrideLogo.REV4 <> "" Then r.REV4 = overrideLogo.REV4

        Return r
    End Function

    Public Sub ActualizarAsesorDesdeSenior()

        Dim AccessDBPath As String = RutaBD & "\ADESLAS.accdb"
        Dim cs As String = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={AccessDBPath};Persist Security Info=False;"

        Using conn As New OleDbConnection(cs)
            conn.Open()

            Dim sql As String =
            "UPDATE TarjetasSanitariasDiarioInteramit " &
            "INNER JOIN AsesorSenior " &
            "ON TarjetasSanitariasDiarioInteramit.NumeroPoliza = AsesorSenior.TJASNPOL " &
            "SET TarjetasSanitariasDiarioInteramit.ASESOR = AsesorSenior.ASESOR, " &
            "    TarjetasSanitariasDiarioInteramit.DIR_ASESOR = AsesorSenior.DIR_ASESOR, " &
            "    TarjetasSanitariasDiarioInteramit.TEL_ASESOR = AsesorSenior.TEL_ASESOR, " &
            "    TarjetasSanitariasDiarioInteramit.CP_ASESOR = AsesorSenior.CP_ASESOR, " &
            "    TarjetasSanitariasDiarioInteramit.POB_ASESOR = AsesorSenior.POB_ASESOR "

            Using cmd As New OleDbCommand(sql, conn)
                Dim afectados = cmd.ExecuteNonQuery()
                'MessageBox.Show($"Actualizados: {afectados}")
            End Using
        End Using

    End Sub



    Private Function DefinirComparaciones() As List(Of ComparacionCampo)

        Dim lista As New List(Of ComparacionCampo)
        Dim c0 As New ComparacionCampo
        c0.Nombre = "Canal_asegurador"
        c0.CampoRegla = Function(r) r.Canal_Asegurador
        c0.CampoRegistro = Function(t) t.Canal_Asegurador
        c0.Tipo = TipoComparacion.Igual
        lista.Add(c0)

        ' COLECTIVO
        Dim c1 As New ComparacionCampo
        c1.Nombre = "COLECTIVO"
        c1.CampoRegla = Function(r) r.ColectivoProducto
        c1.CampoRegistro = Function(t) t.ColectivoProducto
        c1.Tipo = TipoComparacion.Igual
        lista.Add(c1)

        ' LOGOTIPO
        Dim c2 As New ComparacionCampo
        c2.Nombre = "LOGOTIPO"
        c2.CampoRegla = Function(r) r.LogotipoDescripcion
        c2.CampoRegistro = Function(t) t.LogotipoDescripcion
        c2.Tipo = TipoComparacion.Contiene
        lista.Add(c2)

        ' DIRECCIÓN
        Dim c3 As New ComparacionCampo
        c3.CampoRegla = Function(r) r.Direccion
        c3.CampoRegistro = Function(t) t.Direccion
        c3.Tipo = TipoComparacion.Contiene
        lista.Add(c3)

        ' DIRECCIÓN ENVÍO
        Dim c3b As New ComparacionCampo
        c3b.Nombre = "DIRECCIONENVIO"
        c3b.CampoRegla = Function(r) r.DireccionEnvio
        c3b.CampoRegistro = Function(t) t.DireccionEnvio
        c3b.Tipo = TipoComparacion.Contiene
        lista.Add(c3b)

        ' TIPO CONTRA PYMES
        Dim c4 As New ComparacionCampo
        c4.CampoRegla = Function(r) If(r.TipoContraPymes, "X", "")
        c4.CampoRegistro = Function(t) t.TipoContraPymes
        c4.Tipo = TipoComparacion.Igual
        lista.Add(c4)

        ' BUSCAR LOGO
        Dim c5 As New ComparacionCampo
        c5.CampoRegla = Function(r) If(r.BuscarLogo, "X", "")
        c5.CampoRegistro = Function(t) t.NumeroPoliza
        c5.Tipo = TipoComparacion.ExisteEnLogos
        lista.Add(c5)

        ' TIPO CONTRATO
        Dim c6 As New ComparacionCampo
        c6.Nombre = "TIPOCONTRATO"
        c6.CampoRegla = Function(r) r.TipoContrato
        c6.CampoRegistro = Function(t) t.TipoContrato
        c6.Tipo = TipoComparacion.Igual
        lista.Add(c6)

        ' CÓDIGO DELEGACIÓN
        Dim c7 As New ComparacionCampo
        c7.Nombre = "CODIGODELEGACION"
        c7.CampoRegla = Function(r) r.CodigoDelegacion
        c7.CampoRegistro = Function(t) t.CodigoDelegacion
        c7.Tipo = TipoComparacion.Igual
        lista.Add(c7)

        ' Numero de poliza
        Dim c8 As New ComparacionCampo
        c8.Nombre = "NumeroPoliza"
        c8.CampoRegla = Function(r) r.NumeroPoliza
        c8.CampoRegistro = Function(t) t.NumeroPoliza
        c8.Tipo = TipoComparacion.Igual
        lista.Add(c8)

        ' Centro de trabajo (empieza por 001-)
        Dim c9 As New ComparacionCampo
        c9.Nombre = "CentroTrabajoCodigo"
        c9.CampoRegla = Function(r) Texto(r.CentroTrabajoCodigo)
        c9.CampoRegistro = Function(t) Texto(t.CentroTrabajoCodigo)
        c9.Tipo = TipoComparacion.EmpiezaPor
        lista.Add(c9)

        ' INDICADOR IDIOMA (normalizado 2 dígitos)
        Dim c10 As New ComparacionCampo
        c10.Nombre = "INDICADORIDIOMA"
        c10.CampoRegla = Function(r) NormalizarIdioma(r.IndicadorIdioma)
        c10.CampoRegistro = Function(t) NormalizarIdioma(t.IndicadorIdioma)
        c10.Tipo = TipoComparacion.Igual
        lista.Add(c10)



        Return lista

    End Function

    Private Function NormalizarIdioma(v As String) As String
        Dim s = Texto(v)
        If s = "" Then Return ""
        ' si viene "2" -> "02"; si viene "02" -> "02"
        Dim n As Integer
        If Integer.TryParse(Regex.Replace(s, "\D", ""), n) Then
            Return n.ToString("D2")
        End If
        Return s.PadLeft(2, "0"c)
    End Function


    ' ============================================================
    ' REGLAS 
    ' ============================================================
    Private Function CargarReglas() As List(Of ReglaExtraccion)

        Dim lista As New List(Of ReglaExtraccion)
        Dim rutaBDCompleta As String = Path.Combine(RutaBD, "ADESLAS.accdb")

        Dim connectionString As String =
            $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={rutaBDCompleta};Persist Security Info=False;"



        Using cn As New OleDbConnection(ConnectionString)
            cn.Open()

            Dim sql As String =
            "SELECT * FROM ExtracionModelosSanitarios " &
            "WHERE ExcepcionActivada=True"

            Using cmd As New OleDbCommand(sql, cn)
                Using rd = cmd.ExecuteReader()
                    While rd.Read()

                        lista.Add(New ReglaExtraccion With {
                        .Id = CInt(rd("Id")),
                        .Canal_Asegurador = Texto(rd("Canal_asegurador")),
                        .Tipo_Tarjeta = Texto(rd("Tipo_Tarjeta")),
                        .Extraccion = Texto(rd("Extraccion")),
                        .DescripcionTarjeta = Texto(rd("DescripcionTarjeta")),
                        .CodigoSalida = Texto(rd("CodigoSalida")),
                        .ModeloPlasticoCodigo = Texto(rd("ModeloPlasticoCodigo")),
                        .ColectivoProducto = Texto(rd("ColectivoProducto")),
                        .LogotipoDescripcion = Texto(rd("LogotipoDescripcion")),
                        .Direccion = Texto(rd("Direccion")),
                        .DireccionEnvio = Texto(rd("DireccionEnvio")),
                        .ModeloCarrier = Texto(rd("ModeloCarrier")),
                        .CodigoDelegacion = Texto(rd("CodigoDelegacion")),
                        .NumeroPoliza = Texto(rd("NumeroPoliza")),
                        .BuscarLogo = LeerBool(rd("BuscarLogo")),
                        .OrdenacionEspecial = Texto(rd("OrdenacionEspecial")),
                        .SeparaCorreosHispapost = LeerBool(rd("Separa_Correos_HISPAPOST")),
                        .Paquetizado = Texto(rd("PAQUETIZADO")),
                        .TipoContraPymes = LeerBool(rd("TipoContra")),
                        .MODELO = Texto(rd("MODELO")),
                        .TOPPER = Texto(rd("TOPPER")),
                        .LOGO = If(IsDBNull(rd("LOGO")), "", rd("LOGO").ToString().Trim()),
                        .ULTRAANV = Texto(rd("ULTRAANV")),
                        .ULTRAREV = Texto(rd("ULTRAREV")),
                        .CARRIER = Texto(rd("CARRIER")),
                        .SOBRE = Texto(rd("SOBRE")),
                        .TARJETA = Texto(rd("TARJETA")),
                        .FOLLETO = Texto(rd("FOLLETO")),
                        .VC_PROD = Texto(rd("VC_PROD")),
                        .TipoContrato = Texto(rd("TipoContrato")),
                        .CentroTrabajoCodigo = Texto(rd("CentroTrabajoCodigo")),
                        .IndicadorIdioma = Texto(rd("IndicadorIdioma"))
                    })

                    End While
                End Using
            End Using
        End Using

        Return lista

    End Function



    ' ============================================================
    ' REGISTROS
    ' ============================================================
    Private Function CargarRegistros() As List(Of RegistroTarjeta)

        Dim lista As New List(Of RegistroTarjeta)
        Dim rutaBDCompleta As String = Path.Combine(RutaBD, "ADESLAS.accdb")

        Dim connectionString As String =
            $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={rutaBDCompleta};Persist Security Info=False;"

        Using cn As New OleDbConnection(connectionString)

            cn.Open()

            Using cmd As New OleDbCommand(
                "SELECT Id_tarjetasSanitariasDiario, NombreApellidos, CodigoDelegacion,
                NumeroPoliza, NumeroCertificado, NumeroOrden,
                ModeloPlasticoCodigo, ColectivoProducto,
                Direccion, LogotipoDescripcion, TipoContraPymes,
                HISPAPOST, CodigoPostalEnvio, PLATAFORMA, ZONA, CodigoRelacion, CIP_SNS, CIP_M,
                PAQUETIZADO, AnoNacimiento, Sexo, FechaAlta, NumeroTarjeta, CodigoPostal, Poblacion, Provincia, DigitoControlProvincia, DigitoControlZ, 
                Version, FechaCaducidad, TipoContrato, IndicadorExtranjero, FechaInicioCarencia, PersonaReceptora, DireccionEnvio ,PoblacionEnvio, ProvinciaEnvio, CentroTrabajoCodigo,
                ASESOR, DIR_ASESOR, TEL_ASESOR, CP_ASESOR, POB_ASESOR, BENEF, BEN_TARJ, CPF, IND1, IND2, TextoPersonalizado1, TextoPersonalizado2, Canal_asegurador, IndicadorIdioma
                FROM TarjetasSanitariasDiarioInteramit", cn)


                Using rd = cmd.ExecuteReader()
                    While rd.Read()
                        lista.Add(New RegistroTarjeta With {
                            .Id = CInt(rd("Id_tarjetasSanitariasDiario")),
                            .Canal_Asegurador = Texto(rd("Canal_asegurador")),
                            .NombreApellidos = Texto(rd("NombreApellidos")),
                            .CodigoDelegacion = Texto(rd("CodigoDelegacion")),
                            .NumeroPoliza = Texto(rd("NumeroPoliza")),
                            .NumeroCertificado = Texto(rd("NumeroCertificado")),
                            .NumeroOrden = Texto(rd("NumeroOrden")),
                            .CodigoRelacion = Texto(rd("CodigoRelacion")),
                            .ModeloPlasticoCodigo = Texto(rd("ModeloPlasticoCodigo")),
                            .ColectivoProducto = Texto(rd("ColectivoProducto")),
                            .Direccion = Texto(rd("Direccion")),
                            .LogotipoDescripcion = Texto(rd("LogotipoDescripcion")),
                            .TipoContraPymes = Texto(rd("TipoContraPymes")),
                            .HISPAPOST = Texto(rd("HISPAPOST")),
                            .PLATAFORMA = Texto(rd("PLATAFORMA")),
                            .ZONA = Texto(rd("ZONA")),
                            .CIP_SNS = Texto(rd("CIP_SNS")),
                            .CIP_M = Texto(rd("CIP_M")),
                            .Paquetizado = Texto(rd("PAQUETIZADO")),
                            .AnoNacimiento = Texto(rd("AnoNacimiento")),
                            .Sexo = Texto(rd("Sexo")),
                            .FechaAlta = Texto(rd("FechaAlta")),
                            .NumeroTarjeta = Texto(rd("NumeroTarjeta")),
                            .CodigoPostal = Texto(rd("CodigoPostal")),
                            .Poblacion = Texto(rd("Poblacion")),
                            .Provincia = Texto(rd("Provincia")),
                            .DigitoControlProvincia = Texto(rd("DigitoControlProvincia")),
                            .DigitoControlZ = Texto(rd("DigitoControlZ")),
                            .Version = Texto(rd("Version")),
                            .FechaCaducidad = Texto(rd("FechaCaducidad")),
                            .TipoContrato = Texto(rd("TipoContrato")),
                            .IndicadorExtranjero = Texto(rd("IndicadorExtranjero")),
                            .IndicadorIdioma = Texto(rd("IndicadorIdioma")),
                            .FechaInicioCarencia = Texto(rd("FechaInicioCarencia")),
                            .PersonaReceptora = Texto(rd("PersonaReceptora")),
                            .DireccionEnvio = Texto(rd("DireccionEnvio")),
                            .CodigoPostalEnvio = Texto(rd("CodigoPostalEnvio")),
                            .PoblacionEnvio = Texto(rd("PoblacionEnvio")),
                            .ProvinciaEnvio = Texto(rd("ProvinciaEnvio")),
                            .CentroTrabajoCodigo = Texto(rd("CentroTrabajoCodigo")),
                            .ASESOR = Texto(rd("ASESOR")),
                            .DIR_ASESOR = Texto(rd("DIR_ASESOR")),
                            .TEL_ASESOR = Texto(rd("TEL_ASESOR")),
                            .CP_ASESOR = Texto(rd("CP_ASESOR")),
                            .POB_ASESOR = Texto(rd("POB_ASESOR")),
                            .BENEF = Texto(rd("BENEF")),
                            .BEN_TARJ = Texto(rd("BEN_TARJ")),
                            .CPF = Texto(rd("CPF")),
                            .IND1 = Texto(rd("IND1")),
                            .IND2 = Texto(rd("IND2")),
                            .TextoPersonalizado1 = Texto(rd("TextoPersonalizado1")),
                            .TextoPersonalizado2 = Texto(rd("TextoPersonalizado2"))
                            })

                    End While
                End Using
            End Using
        End Using

        Return lista
    End Function

    Private Function CumpleRegla(
    r As RegistroTarjeta,
    regla As ReglaExtraccion,
    logos As Dictionary(Of String, String)
) As Boolean

        ' ============================================================
        ' 1️⃣ MODELO PLÁSTICO (OBLIGATORIO)
        ' ============================================================
        If Texto(r.ModeloPlasticoCodigo).PadLeft(2, "0"c) <>
       Texto(regla.ModeloPlasticoCodigo).PadLeft(2, "0"c) Then
            Return False
        End If

        ' ============================================================
        ' 2️⃣ CAMPOS OPCIONALES (motor escalable)
        ' ============================================================
        For Each c In DefinirComparaciones()



            ' =========================================
            ' PRIORIDAD BuscarLogo
            ' =========================================
            If regla.BuscarLogo Then
                If c.Nombre = "COLECTIVO" OrElse c.Nombre = "LOGOTIPO" Then
                    Continue For
                End If
            End If

            Dim valorRegla As String = Texto(c.CampoRegla(regla))
            If valorRegla = "" Then Continue For

            Dim valorRegistro As String = Texto(c.CampoRegistro(r))

            Select Case c.Tipo
                Case TipoComparacion.Igual
                    If valorRegistro <> valorRegla Then Return False

                Case TipoComparacion.Contiene
                    If Not valorRegistro.Contains(valorRegla) Then Return False

                Case TipoComparacion.ExisteEnLogos
                    If Not logos.ContainsKey(valorRegistro) Then Return False

                Case TipoComparacion.EmpiezaPor
                    If Not valorRegistro.StartsWith(valorRegla) Then Return False

            End Select
        Next


        ' ============================================================
        ' 3️⃣ PAQUETIZADO (solo si la regla lo exige)
        ' ============================================================
        If regla.Paquetizado <> "" Then
            If Texto(r.Paquetizado) <> Texto(regla.Paquetizado) Then
                Return False
            End If
        End If

        ' ============================================================
        ' ✅ CUMPLE LA REGLA
        ' ============================================================
        Return True

    End Function

    Public Sub MarcarPaquetizados()

        Dim AccessDBPath As String = RutaBD & "\ADESLAS.accdb"
        Dim connectionString As String =
            $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={AccessDBPath};Persist Security Info=False;"

        Using conn As New OleDbConnection(connectionString)
            conn.Open()

            Dim sql As String =
                "UPDATE TarjetasSanitariasDiarioInteramit " &
                "SET PAQUETIZADO = 'X' " &
                "WHERE NumeroPoliza IN (SELECT N_POLIZA FROM ColectivosConDireccionEnvio WHERE N_POLIZA <> '') " &
                "AND IndicadorExtraccion NOT IN ('N','W','X')"

            Using cmd As New OleDbCommand(sql, conn)
                cmd.ExecuteNonQuery()
            End Using

        End Using

    End Sub

    ' ============================================================
    ' MARCAR HISPAPOST (X) Y CORREOS (C)
    ' ============================================================
    Public Sub MarcarRegistrosHispapost()

        Dim AccessDBPath As String = RutaBD & "\ADESLAS.accdb"
        Dim connectionString As String =
        $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={AccessDBPath};Persist Security Info=False;"

        Using conn As New OleDbConnection(connectionString)

            conn.Open()

            ' Primero todos = CORREOS
            Dim sqlC As String =
                "UPDATE TarjetasSanitariasDiarioInteramit " &
                "SET HISPAPOST = 'C', ZONA = '', PLATAFORMA = ''"

            Using cmd As New OleDbCommand(sqlC, conn)
                cmd.ExecuteNonQuery()
            End Using

            ' Después los HISPAPOST
            Dim sqlH As String =
                "UPDATE TarjetasSanitariasDiarioInteramit " &
                "INNER JOIN THPOST ON TarjetasSanitariasDiarioInteramit.CodigoPostalEnvio = THPOST.CPOSTAL " &
                "SET TarjetasSanitariasDiarioInteramit.HISPAPOST = 'X', " &
                "    TarjetasSanitariasDiarioInteramit.ZONA = THPOST.ZONA, " &
                "    TarjetasSanitariasDiarioInteramit.PLATAFORMA = THPOST.PLATAFORMA"

            Using cmd As New OleDbCommand(sqlH, conn)
                cmd.ExecuteNonQuery()
            End Using

        End Using

    End Sub

    ' ============================================================
    ' MARCAR TIPO CONTRA PYMES
    ' ============================================================
    Public Sub MarcarTipoContraPymes()

        Dim AccessDBPath As String = RutaBD & "\ADESLAS.accdb"
        Dim cs As String = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={AccessDBPath};Persist Security Info=False;"

        Using conn As New OleDbConnection(cs)
            conn.Open()

            Dim sql As String =
                "UPDATE TarjetasSanitariasDiarioInteramit " &
                "INNER JOIN TipoContra ON TarjetasSanitariasDiarioInteramit.TipoContrato = TipoContra.Cd_TipoContra " &
                "SET TarjetasSanitariasDiarioInteramit.TipoContraPymes = 'X'"

            Using cmd As New OleDbCommand(sql, conn)
                cmd.ExecuteNonQuery()
            End Using

        End Using

    End Sub


    ' ============================================================
    ' MÉTODO PRINCIPAL (AGRUPANDO POR NOMBRE DE FICHERO)
    ' ============================================================
    Public Sub EjecutarSeparacion()

        Dim sufijoEntrada As String = ExtraerSufijoDesdeEntrada(NombreFicheroEntrada)
        Dim listaFicherosGenerados As New List(Of Tuple(Of String, Integer))
        ' ============================================================
        ' PREPARACIÓN EN BBDD
        ' ============================================================
        MarcarCanalSanitario()
        MarcarCentroDeTrabajoCaixaOficinVirtual()
        ForzarVC_PL_ID_T1_T2_TEFex()
        MarcarPaquetizados()
        MarcarRegistrosHispapost()
        MarcarTipoContraPymes()
        CambioDeModelo_666_TipoContrato640_A_Modelo33()
        ActualizarAsesorDesdeSenior()
        CorrigeDireccionEnvio()
        RellenarProvinciaEnvioDesdeCP()

        ' ============================================================
        ' CARGA EN MEMORIA
        ' ============================================================
        Dim reglas = CargarReglas()

        ' ✅ ORDEN BASE ABSOLUTO POR ID (orden real de entrada)
        Dim registros = CargarRegistros().
        OrderBy(Function(r) r.Id).
        ToList()

        Dim logos = CargarLogos()
        Dim logosSalida = CargarLogosSalida()
        polizasSinAsistenciaViaje = CargarPolizasSinAsistenciaViaje()


        ' ============================================================
        ' AGRUPACIÓN POR FICHERO DE SALIDA
        ' ============================================================
        Dim ficheros As New Dictionary(Of String, List(Of RegistroTarjeta))
        Dim reglasPorFichero As New Dictionary(Of String, ReglaExtraccion)

        For Each r In registros

            Dim modeloRegistro As String =
            Texto(r.ModeloPlasticoCodigo).PadLeft(2, "0"c)
            Dim reglasModelo = reglas _
            .Where(Function(reg) Texto(reg.ModeloPlasticoCodigo).PadLeft(2, "0"c) = modeloRegistro) _
            .ToList()

            Dim reglaAplicada As ReglaExtraccion = Nothing

            If reglasModelo.Any() Then
                Dim candidatas = reglasModelo _
.Where(Function(reg) CumpleRegla(r, reg, logos)) _
.OrderByDescending(Function(reg) GradoEspecificidad(reg)) _
.ThenBy(Function(reg) reg.Id) _
                .ToList()

                If candidatas.Any() Then
                    reglaAplicada = candidatas.First()
                End If
            End If

            Dim nombreFichero As String
            Dim reglaFinal As ReglaExtraccion = reglaAplicada

            If reglaAplicada IsNot Nothing Then

                ' =====================================================
                ' ✅ DESCRIPCIÓN FINAL + COMBINACIÓN CON LOGO
                ' =====================================================

                Dim descripcionFinal As String = reglaAplicada.DescripcionTarjeta

                If reglaAplicada.BuscarLogo Then
                    Dim poliza As String = Texto(r.NumeroPoliza)

                    If logos.ContainsKey(poliza) Then

                        Dim claveLogo As String = logos(poliza)
                        descripcionFinal = claveLogo

                        Dim modelo As String =
                        Texto(reglaAplicada.ModeloPlasticoCodigo).PadLeft(2, "0"c)

                        Dim extraccion As String = reglaAplicada.Extraccion

                        Dim clave As String = extraccion & "|" & claveLogo & "|" & modelo

                        If logosSalida.ContainsKey(clave) Then
                            reglaFinal = CombinarReglaBaseConLogo(
                                reglaAplicada,
                                logosSalida(clave)
                            )
                        End If


                    End If
                End If

                ' =====================================================
                ' ✅ NOMBRE DE FICHERO
                ' =====================================================
                If reglaAplicada.SeparaCorreosHispapost Then

                    If r.HISPAPOST = "X" Then
                        nombreFichero =
                        reglaAplicada.Tipo_Tarjeta & "H" &
                        descripcionFinal &
                        reglaAplicada.CodigoSalida &
                        sufijoEntrada & ".TXT"

                    Else
                        nombreFichero =
                        reglaAplicada.Tipo_Tarjeta & "T" &
                        descripcionFinal &
                        reglaAplicada.CodigoSalida &
                        sufijoEntrada & ".TXT"

                    End If

                Else
                    nombreFichero =
                    reglaAplicada.Tipo_Tarjeta &
                    reglaAplicada.Extraccion &
                    descripcionFinal &
                    reglaAplicada.CodigoSalida &
                    sufijoEntrada & ".TXT"

                End If

            Else
                ' 🔒 Fallback seguro
                nombreFichero = "XX" & modeloRegistro & "X.txt"
            End If

            If Not ficheros.ContainsKey(nombreFichero) Then
                ficheros(nombreFichero) = New List(Of RegistroTarjeta)

                If reglaFinal IsNot Nothing Then
                    reglasPorFichero(nombreFichero) = reglaFinal
                End If
            End If

            ' ✅ SE MANTIENE ORDEN ORIGINAL (por ID)
            ficheros(nombreFichero).Add(r)

        Next

        ' ============================================================
        ' GENERACIÓN DE FICHEROS
        ' ============================================================
        'Dim carpetaSalida As String =
        '"C:\Adeslas\Adeslas_Tarjetas_Diario\salida\"

        Directory.CreateDirectory(carpetaSalida)

        For Each kvp In ficheros

            ' ✅ ORDEN BASE OBLIGATORIO POR ID
            Dim lista = kvp.Value _
            .OrderBy(Function(r) r.Id) _
.ToList()

            ' ✅ ORDENACIONES ESPECIALES
            If reglasPorFichero.ContainsKey(kvp.Key) Then
                Dim regla = reglasPorFichero(kvp.Key)

                If regla.SeparaCorreosHispapost AndAlso kvp.Key.Contains("H") Then
                    lista = AplicarOrdenacionHispapost(lista)

                ElseIf regla.OrdenacionEspecial <> "" Then
                    lista = AplicarOrdenacion(lista, regla.OrdenacionEspecial)
                End If
            End If

            'Dim rutaCompleta As String =
            'Path.Combine(carpetaSalida, kvp.Key)

            Dim nombreSalida As String =
            Path.GetFileNameWithoutExtension(kvp.Key) & "_.TXT"

            Dim rutaCompleta As String =
            Path.Combine(carpetaSalida, nombreSalida)


            Dim reglaSalida As ReglaExtraccion = Nothing
            If reglasPorFichero.ContainsKey(kvp.Key) Then
                reglaSalida = reglasPorFichero(kvp.Key)
            End If

            Dim listaAD = lista.Where(Function(t) Texto(t.Canal_Asegurador) = "AD").ToList()
            Dim listaVC = lista.Where(Function(t) Texto(t.Canal_Asegurador) = "VC").ToList()



            If listaAD.Count > 0 And Not nombreSalida.Contains("_LPM") Then
                CalcularCPF_IND_EN_LISTA_Y_GUARDAR(listaAD)
            Else
                CalcularVC_IND_Y_GUARDAR(listaAD)
            End If

            If listaVC.Count > 0 Then
                CalcularVC_IND_Y_GUARDAR(listaVC)
            End If

            GenerarTxt(rutaCompleta, lista, reglaSalida)

            listaFicherosGenerados.Add(New Tuple(Of String, Integer)(Path.GetFileName(rutaCompleta), lista.Count)
)
        Next



        ' ============================================================
        ' FIN
        ' ============================================================
        MessageBox.Show($"Separación finalizada{vbCrLf}" & $"Ficheros generados: {ficheros.Count}", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information)

        ' 🔥 GENERAR EXCEL LOG
        GenerarLogProcesoExcel(carpetaSalida, NombreFicheroEntrada, listaFicherosGenerados)

    End Sub
    Private Sub PonerDefaultsCPF(lista As List(Of RegistroTarjeta))
        For Each t In lista
            t.CPF = "1"
            t.IND1 = "1"
            t.IND2 = "L"
            t.BENEF = ""
            t.BEN_TARJ = ""
            t.BARRAS_CONTROL = "1"   ' 👈 CLAVE: siempre 1 para VC
        Next
    End Sub
    Private Sub CalcularVC_IND_Y_GUARDAR(lista As List(Of RegistroTarjeta))

        Using cn As New OleDbConnection(ConnectionString)
            cn.Open()

            Dim sql As String =
            "UPDATE TarjetasSanitariasDiarioInteramit " &
            "SET CPF=?, IND1=?, IND2=?, BENEF=?, BEN_TARJ=?, BARRAS_CONTROL=? " &
            "WHERE Id_tarjetasSanitariasDiario=?"

            Using cmd As New OleDbCommand(sql, cn)

                cmd.Parameters.Clear()
                cmd.Parameters.Add(New OleDbParameter With {.OleDbType = OleDbType.VarChar}) ' CPF
                cmd.Parameters.Add(New OleDbParameter With {.OleDbType = OleDbType.VarChar}) ' IND1
                cmd.Parameters.Add(New OleDbParameter With {.OleDbType = OleDbType.VarChar}) ' IND2
                cmd.Parameters.Add(New OleDbParameter With {.OleDbType = OleDbType.VarChar}) ' BENEF
                cmd.Parameters.Add(New OleDbParameter With {.OleDbType = OleDbType.VarChar}) ' BEN_TARJ
                cmd.Parameters.Add(New OleDbParameter With {.OleDbType = OleDbType.VarChar}) ' BARRAS_CONTROL
                cmd.Parameters.Add(New OleDbParameter With {.OleDbType = OleDbType.Integer}) ' ID

                ' Contador por grupo POLIZA+CERTIFICA
                Dim contadorPorGrupo As New Dictionary(Of String, Integer)

                For Each t In lista

                    ' Defaults VC
                    t.CPF = "1"
                    t.IND1 = "1"
                    t.IND2 = "L"
                    t.BENEF = ""
                    t.BEN_TARJ = ""

                    ' BARRAS_CONTROL incremental por grupo (misma lógica)
                    Dim claveGrupo As String = t.NumeroPoliza & "|" & t.NumeroCertificado

                    Dim indice As Integer
                    If Not contadorPorGrupo.ContainsKey(claveGrupo) Then
                        indice = 1
                    Else
                        indice = contadorPorGrupo(claveGrupo) + 1
                    End If
                    contadorPorGrupo(claveGrupo) = indice

                    t.BARRAS_CONTROL = indice.ToString()

                    GuardarRegistro(cmd, t)
                Next

            End Using
        End Using

    End Sub


    Private Sub CalcularCPF_IND_EN_LISTA_Y_GUARDAR(lista As List(Of RegistroTarjeta))
        Dim rutaBDCompleta As String = Path.Combine(RutaBD, "ADESLAS.accdb")

        Dim connectionString As String =
            $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={rutaBDCompleta};Persist Security Info=False;"


        Using cn As New OleDbConnection(ConnectionString)
            cn.Open()

            Dim sql As String =
            "UPDATE TarjetasSanitariasDiarioInteramit " &
            "SET CPF = ?, IND1 = ?, IND2 = ?, BENEF = ?, BEN_TARJ = ?, BARRAS_CONTROL = ? " &
            "WHERE Id_tarjetasSanitariasDiario = ?"

            Using cmd As New OleDbCommand(sql, cn)

                cmd.Parameters.Clear()
                ' IMPORTANTE: SOLO POR ORDEN, SIN NOMBRES
                cmd.Parameters.Add(New OleDbParameter With {.OleDbType = OleDbType.VarChar}) ' CPF
                cmd.Parameters.Add(New OleDbParameter With {.OleDbType = OleDbType.VarChar}) ' IND1
                cmd.Parameters.Add(New OleDbParameter With {.OleDbType = OleDbType.VarChar}) ' IND2
                cmd.Parameters.Add(New OleDbParameter With {.OleDbType = OleDbType.VarChar}) ' BENEF
                cmd.Parameters.Add(New OleDbParameter With {.OleDbType = OleDbType.VarChar}) ' BEN_TARJ
                cmd.Parameters.Add(New OleDbParameter With {.OleDbType = OleDbType.VarChar}) ' BARRAS_CONTROL
                cmd.Parameters.Add(New OleDbParameter With {.OleDbType = OleDbType.Integer}) ' ID

                ' Contador por grupo POLIZA+CERTIFICA
                Dim contadorPorGrupo As New Dictionary(Of String, Integer)

                Dim i As Integer = 0

                While i < lista.Count

                    Dim actual = lista(i)

                    ' ================================
                    ' VALORES POR DEFECTO (REGISTRO SUELTO)
                    ' ================================
                    actual.CPF = "1"
                    actual.IND1 = "1"
                    actual.IND2 = "L"
                    actual.BENEF = ""
                    actual.BEN_TARJ = ""

                    ' Clave del grupo
                    Dim claveGrupo As String = actual.NumeroPoliza & "|" & actual.NumeroCertificado

                    ' ================================
                    ' ¿HAY SIGUIENTE?
                    ' ================================
                    If i + 1 < lista.Count Then

                        Dim siguiente = lista(i + 1)

                        If actual.NumeroPoliza = siguiente.NumeroPoliza AndAlso
                       actual.NumeroCertificado = siguiente.NumeroCertificado Then

                            ' ---------- TITULAR ----------
                            actual.CPF = "2"
                            actual.IND1 = "1"
                            actual.IND2 = " "
                            actual.BENEF = siguiente.NombreApellidos
                            actual.BEN_TARJ = siguiente.NumeroTarjeta

                            ' ---------- BENEFICIARIO ----------
                            siguiente.CPF = "2"
                            siguiente.IND1 = "2"
                            siguiente.IND2 = "L"
                            siguiente.BENEF = ""
                            siguiente.BEN_TARJ = ""

                            ' =====================================
                            ' BARRAS_CONTROL PARA LA PAREJA (1,1) / (2,2) / (3,3) ...
                            ' =====================================
                            Dim indice As Integer
                            If Not contadorPorGrupo.ContainsKey(claveGrupo) Then
                                indice = 1
                            Else
                                indice = contadorPorGrupo(claveGrupo) + 1
                            End If
                            contadorPorGrupo(claveGrupo) = indice

                            actual.BARRAS_CONTROL = indice.ToString()
                            siguiente.BARRAS_CONTROL = indice.ToString()

                            ' Guardar ambos
                            GuardarRegistro(cmd, actual)
                            GuardarRegistro(cmd, siguiente)

                            i += 2
                            Continue While
                        End If
                    End If

                    ' ================================
                    ' REGISTRO SUELTO (no hay pareja)
                    ' ================================
                    Dim indiceSuelto As Integer
                    If Not contadorPorGrupo.ContainsKey(claveGrupo) Then
                        indiceSuelto = 1
                    Else
                        indiceSuelto = contadorPorGrupo(claveGrupo) + 1
                    End If
                    contadorPorGrupo(claveGrupo) = indiceSuelto

                    actual.BARRAS_CONTROL = indiceSuelto.ToString()

                    GuardarRegistro(cmd, actual)
                    i += 1

                End While
            End Using
        End Using

    End Sub


    Private Sub GuardarRegistro(cmd As OleDbCommand, r As RegistroTarjeta)

        cmd.Parameters(0).Value = If(String.IsNullOrWhiteSpace(r.CPF), "", r.CPF)
        cmd.Parameters(1).Value = If(String.IsNullOrWhiteSpace(r.IND1), "", r.IND1)
        cmd.Parameters(2).Value = If(String.IsNullOrWhiteSpace(r.IND2), "", r.IND2)
        cmd.Parameters(3).Value = If(String.IsNullOrWhiteSpace(r.BENEF), "", r.BENEF)
        cmd.Parameters(4).Value = If(String.IsNullOrWhiteSpace(r.BEN_TARJ), "", r.BEN_TARJ)
        cmd.Parameters(5).Value = If(String.IsNullOrWhiteSpace(r.BARRAS_CONTROL), "", r.BARRAS_CONTROL)

        cmd.Parameters(6).Value = r.Id

        cmd.ExecuteNonQuery()

    End Sub
    Private Function CargarPolizasSinAsistenciaViaje() As HashSet(Of String)

        Dim resultado As New HashSet(Of String)
        Dim rutaBDCompleta As String = Path.Combine(RutaBD, "ADESLAS.accdb")

        Dim connectionString As String =
            $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={rutaBDCompleta};Persist Security Info=False;"


        Using cn As New OleDbConnection(ConnectionString)
            cn.Open()

            Dim sql As String =
            "SELECT POLINPOL " &
            "FROM ColectivosSinAsistenciaViaje " &
            "WHERE POLINPOL Is Not Null AND Trim(POLINPOL)<>''"

            Using cmd As New OleDbCommand(sql, cn)
                Using rd = cmd.ExecuteReader()
                    While rd.Read()
                        Dim poliza As String = Texto(rd("POLINPOL"))
                        If poliza <> "" Then
                            resultado.Add(poliza)
                        End If
                    End While
                End Using
            End Using
        End Using

        Return resultado

    End Function






    ' ============================================================
    ' TXT – FORMATO FINAL CON CABECERA
    ' ============================================================
    Private Sub GenerarTxt(ruta As String, lista As List(Of RegistroTarjeta), regla As ReglaExtraccion)

        Dim reglasTextoPersonalizado = CargarReglasTextoPersonalizado()

        Using sw As New StreamWriter(ruta, False, Encoding.GetEncoding(1252))

            ' -------- CABECERA --------
            sw.WriteLine(
            "SECUENCIAL;PAN;EMB1;NOMBRE;COLEC;REV1;REV2;REV3;REV4;" &
            "ANV1;ANV2;NTARJETA;PISTA1;PISTA2;PRECEP;DIRECCION;" &
            "CPOSTAL;POBLACION;PROVINCIA;BENEF;BEN_TARJ;COD_BARRAS;" &
            "CENT_COD;CENT_DES;F_EFECTO;ASESOR;DIR_ASESOR;TEL_ASESOR;" &
            "POLIZA;CERTIFICA;EMPRESA;CPF;IND1;IND2;ARCHIVO;MODELO;" &
            "TOPPER;LOGO;ULTRAANV;ULTRAREV;CARRIER;SOBRE;TARJETA;" &
            "FOLLETO;VC_PROD;DMX;EMB0;PLATAFORMA;CIP_M"
        )

            Dim sec As Integer = 1

            For Each r In lista

                'CODIGO DE CAMPAÑA Lo asigna Asisa, por defecto 00
                Dim Cd_Campa = "00"

                If r.Canal_Asegurador = "AD" Then
                    Cd_Campa = "12"
                End If


                ' ✅ NOMBRE REAL DEL FICHERO DE SALIDA 
                Dim nombreFicheroSalida As String = Path.GetFileName(ruta).ToUpperInvariant()
                Dim NOMBRE As String = r.NombreApellidos
                Dim COD_BARRAS As String
                Dim campos As New List(Of String)
                Dim codDelegacionFormateado As String = r.CodigoDelegacion.Trim().PadLeft(3, "0"c)
                Dim DMX As String = " "
                Dim CIP_M As String = " "
                Dim PRECEP As String = r.PersonaReceptora
                Dim DireccionEnvio As String
                Dim EMPRESA As String
                Dim TextoPersonalizado1 As String = r.TextoPersonalizado1
                Dim TextoPersonalizado2 As String = r.TextoPersonalizado2
                Dim PoblacionEnvio As String = r.PoblacionEnvio
                Dim ProvinciaEnvio As String = r.ProvinciaEnvio
                ' ----------------------------------------------------------
                'ASESOR
                ' ----------------------------------------------------------
                Dim ASESOR As String = ""
                Dim DIR_ASESOR As String = ""
                Dim TEL_ASESOR As String = ""
                Dim CENT_COD As String = r.CentroTrabajoCodigo
                'VUELVO A PONER OFV A 001
                If Mid(CENT_COD, 1, 3) = "OFV" Then
                    Mid(CENT_COD, 1, 3) = "001"
                End If

                Dim CENT_DES As String = ""

                If r.Canal_Asegurador = "VC" Then

                    'PARA VC SENIOR

                    If nombreFicheroSalida.Contains("VL") Or nombreFicheroSalida.Contains("VJ") Or nombreFicheroSalida.Contains("SB") Then
                        'ASESOR = r.ASESOR.Split("("c)(0).Trim() 'SOLO CARGO HASTA QUE VEA UN "("

                        ASESOR = r.ASESOR
                        If ASESOR <> "" Then
                            DIR_ASESOR = r.DIR_ASESOR.Trim() & ". " & r.CP_ASESOR.Trim() & " - " & r.POB_ASESOR.Trim()
                            TEL_ASESOR = "Tlf. " & r.TEL_ASESOR.Trim()
                        End If
                        TextoPersonalizado1 = Replace(TextoPersonalizado1, "SALUD + DENTAL", "Salud + Dental")
                    End If

                    If r.CentroTrabajoCodigo.Contains("001-") Then
                        CENT_COD = r.CentroTrabajoCodigo
                    Else
                        CENT_COD = " "
                    End If
                    CENT_DES = " "

                End If



                ' ----------------------------------------------------------
                ' CORREGIR FECHA DE CARENCIA SI ESTÁ VACÍA O INVÁLIDA
                ' ----------------------------------------------------------
                If String.IsNullOrWhiteSpace(r.FechaInicioCarencia) Then
                    r.FechaInicioCarencia = "0000"
                ElseIf r.FechaInicioCarencia.Trim().Length < 4 Then
                    r.FechaInicioCarencia = "0000"
                End If
                ' ----------------------------------------------------------


                If nombreFicheroSalida.Contains("_TSL") Then
                    'ASESOR = r.ASESOR.Split("("c)(0).Trim() 'SOLO CARGO HASTA QUE VEA UN "("
                    ASESOR = r.ASESOR
                    DIR_ASESOR = r.DIR_ASESOR.Trim() & ". " & r.CP_ASESOR.Trim() & " - " & r.POB_ASESOR.Trim()
                    If r.TEL_ASESOR.Trim() = "" Then
                        TEL_ASESOR = ""
                    Else
                        TEL_ASESOR = "Tlf. " & r.TEL_ASESOR.Trim()
                    End If

                End If
                If nombreFicheroSalida.Contains("_TSR") Then
                    'ASESOR = r.ASESOR.Split("("c)(0).Trim() 'SOLO CARGO HASTA QUE VEA UN "("

                    ASESOR = r.ASESOR
                    If ASESOR <> "" Then
                        DIR_ASESOR = r.DIR_ASESOR.Trim() & ". " & r.CP_ASESOR.Trim() & " - " & r.POB_ASESOR.Trim()
                        TEL_ASESOR = "Tlf. " & r.TEL_ASESOR.Trim()
                    End If
                    TextoPersonalizado1 = Replace(TextoPersonalizado1, "SALUD + DENTAL", "Salud + Dental")
                End If

                ' ----------------------------------------------------------
                ' CODIGO DE BARRAS Adeslas
                ' ----------------------------------------------------------
                COD_BARRAS = Cd_Campa & codDelegacionFormateado & r.NumeroPoliza.Trim() & r.NumeroCertificado.Trim() & r.BARRAS_CONTROL & "-" & Replace(nombreFicheroSalida, "_.", ".")

                If nombreFicheroSalida.Contains("_UAB") Then
                    COD_BARRAS = Cd_Campa & codDelegacionFormateado & r.NumeroPoliza.Trim() & r.NumeroCertificado.Trim() & r.BARRAS_CONTROL & "-" & Replace(nombreFicheroSalida, "_.", ".")

                End If

                If nombreFicheroSalida.Contains("_L00") Then
                    COD_BARRAS = Cd_Campa & codDelegacionFormateado & r.NumeroPoliza.Trim() & r.NumeroCertificado.Trim() & r.NumeroOrden & r.BARRAS_CONTROL & "-" & Replace(nombreFicheroSalida, "_.", ".")

                End If

                If nombreFicheroSalida.Contains("_LIO") Then
                    COD_BARRAS = Cd_Campa & codDelegacionFormateado & r.NumeroPoliza.Trim() & r.NumeroCertificado.Trim() & r.NumeroOrden & r.BARRAS_CONTROL & "-" & Replace(nombreFicheroSalida, "_.", ".")

                End If

                If nombreFicheroSalida.Contains("_TGD") Then
                    COD_BARRAS = Cd_Campa & codDelegacionFormateado & r.NumeroPoliza.Trim() & r.NumeroCertificado.Trim() & r.NumeroOrden & r.BARRAS_CONTROL & "-" & Replace(nombreFicheroSalida, "_.", ".")

                End If

                If nombreFicheroSalida.Contains("_TMB") Then
                    COD_BARRAS = Cd_Campa & codDelegacionFormateado & r.NumeroPoliza.Trim() & r.NumeroCertificado.Trim() & r.NumeroOrden & r.BARRAS_CONTROL & "-" & Replace(nombreFicheroSalida, "_.", ".")

                End If

                If nombreFicheroSalida.Contains("_TTL") Then
                    COD_BARRAS = Cd_Campa & codDelegacionFormateado & r.NumeroPoliza.Trim() & r.NumeroCertificado.Trim() & r.BARRAS_CONTROL & "-" & Replace(nombreFicheroSalida, "_.", ".")

                End If




                ' ----------------------------------------------------------
                ' CODIGO DE BARRAS Vida Caixa
                ' ----------------------------------------------------------
                'PARA CORREOS
                If r.Canal_Asegurador.Contains("VC") Then
                    COD_BARRAS = Cd_Campa & codDelegacionFormateado & r.NumeroPoliza.Trim() & r.NumeroCertificado.Trim() & r.NumeroOrden & "1-" & Replace(nombreFicheroSalida, "_.", ".")
                    'PARA HISPAPOST
                    If nombreFicheroSalida.Contains("_H") Then
                        COD_BARRAS = Cd_Campa & codDelegacionFormateado & r.NumeroPoliza.Trim() & r.NumeroCertificado.Trim() & r.NumeroOrden & "1-" & Replace(nombreFicheroSalida, "_.", ".")
                    End If
                End If


                ' ----------------------------------------------------------
                ' PAN
                ' ----------------------------------------------------------
                Dim PAN As String = Mid(r.NumeroTarjeta, 1, 4) & " " &
                                Mid(r.NumeroTarjeta, 5, 4) & " " &
                                Mid(r.NumeroTarjeta, 9, 1) & "000 " &
                                Mid(codDelegacionFormateado, 1, 2) &
                                r.DigitoControlProvincia &
                                r.DigitoControlZ


                ' ----------------------------------------------------------
                ' EMB1
                ' ----------------------------------------------------------
                Dim EMB1 As String = Mid(r.AnoNacimiento, 3, 2) & Space(3) & (r.Sexo) & Space(10) & "AD" & Space(3) &
                                 (Mid(r.FechaAlta, 5, 2) & "/" & Mid(r.FechaAlta, 3, 2))

                ' ----------------------------------------------------------
                ' REV1 / REV2 + excepciones
                ' ----------------------------------------------------------
                Dim REV1 As String = ""
                Dim REV2 As String = ""
                Dim REV3 As String = If(regla IsNot Nothing, regla.REV3, "")
                Dim REV4 As String = If(regla IsNot Nothing, regla.REV4, "")
                Dim polizaActual As String = r.NumeroPoliza

                Dim PLATAFORMA = r.PLATAFORMA


                If polizasSinAsistenciaViaje Is Nothing _
                OrElse Not polizasSinAsistenciaViaje.Contains(polizaActual) Then

                    REV1 = "Tlf. Atención en Extranjero"
                    REV2 = "34-91-7453280"
                End If
                If nombreFicheroSalida.Contains("_UAB") Then
                    REV1 = ""
                    REV2 = ""
                End If
                If nombreFicheroSalida.Contains("_TCE") Then
                    REV1 = ""
                    REV2 = ""
                End If
                If nombreFicheroSalida.Contains("_MCE") Then
                    REV1 = ""
                    REV2 = ""
                End If
                If nombreFicheroSalida.Contains("_TMZ") Then
                    REV1 = ""
                    REV2 = ""
                End If
                If nombreFicheroSalida.Contains("_Y00") Then
                    REV1 = ""
                    REV2 = ""
                End If
                If nombreFicheroSalida.Contains("_TCP") Then
                    REV1 = "Tlf. Atención en Extranjero"
                    REV2 = "34-91-7453280"
                End If

                If nombreFicheroSalida.Contains("_TAR") Then
                    If polizaActual.Contains("6660") Then
                        REV1 = "Tlf. Atención en Extranjero"
                        REV2 = "34-91-7453280"
                    Else
                        REV1 = ""
                        REV2 = ""
                    End If
                End If
                If nombreFicheroSalida.Contains("_TIM") Then
                    REV1 = "Tlf. Atención en Extranjero"
                    REV2 = "34-91-7453280"
                    REV3 = "Tlf.Atención al Cliente en P.Vasco"
                    REV4 = "900-81-81-50"
                End If
                ' ----------------------------------------------------------
                'IndicadorExtranjero teléfono
                ' ----------------------------------------------------------
                Dim extranjeros As String = r.IndicadorExtranjero
                If extranjeros = "N" Then
                    REV1 = ""
                    REV2 = ""
                End If
                ' ----------------------------------------------------------
                ' VIDA CAIXA REV1 Y REV2 AÑADIR "Asistencia en viaje" "34.91.745.32.80"
                ' ----------------------------------------------------------
                If r.Canal_Asegurador = "VC" Then


                    If TextoPersonalizado1 = "ASSISTÈNCIA SANITÀRI" Then
                        TextoPersonalizado1 = "ASSISTÈNCIA SANITÀRIA"
                    End If
                    If REV1 <> "" Then
                        If nombreFicheroSalida.Contains("SB") Then
                            REV1 = "Asistencia en viaje"
                            REV2 = "34.91.745.32.80"
                        End If

                        If nombreFicheroSalida.Contains("VU") Then
                            REV1 = "Asistencia en viaje"
                            REV2 = "34.91.745.32.80"
                        End If
                    End If

                    'FORZAR TELEFONO
                    If nombreFicheroSalida.Contains("SB") Then
                        REV1 = "Asistencia en viaje"
                        REV2 = "34.91.745.32.80"
                    End If
                    If nombreFicheroSalida.Contains("MY") Then
                        REV1 = "Asistencia en viaje"
                        REV2 = "34.91.745.32.80"
                    End If
                    If nombreFicheroSalida.Contains("VC") Then
                        REV1 = "Asistencia en viaje"
                        REV2 = "34.91.745.32.80"
                    End If
                    If nombreFicheroSalida.Contains("VG") Then
                        REV1 = "Asistencia en viaje"
                        REV2 = "34.91.745.32.80"
                    End If
                    If nombreFicheroSalida.Contains("VL") Then

                        REV1 = "Asistencia en viaje"
                        REV2 = "34.91.745.32.80"
                    End If
                    If nombreFicheroSalida.Contains("BX") Then
                        REV1 = "Assistència en viatge"
                        REV2 = "34.91.745.32.80"
                    End If
                    If nombreFicheroSalida.Contains("VW") Then
                        REV1 = "Assistència en viatge"
                        REV2 = "34.91.745.32.80"
                    End If
                    If nombreFicheroSalida.Contains("VV") Then
                        REV1 = "Assistència en viatge"
                        REV2 = "34.91.745.32.80"
                    End If
                    If nombreFicheroSalida.Contains("VT") Then
                        REV1 = "Assistència en viatge"
                        REV2 = "34.91.745.32.80"
                    End If
                    If nombreFicheroSalida.Contains("VJ") Then
                        REV1 = "Assistència en viatge"
                        REV2 = "34.91.745.32.80"
                    End If
                    If nombreFicheroSalida.Contains("MX") Then
                        REV1 = "Asistencia en viaje"
                        REV2 = "34.91.745.32.80"
                    End If
                    If nombreFicheroSalida.Contains("VE") Then
                        REV1 = "Asistencia en viaje"
                        REV2 = "34.91.745.32.80"
                    End If
                    If nombreFicheroSalida.Contains("VU") Then
                        REV1 = "Asistencia en viaje"
                        REV2 = "34.91.745.32.80"
                    End If
                    If nombreFicheroSalida.Contains("EVL") Then
                        REV1 = ""
                        REV2 = ""
                    End If
                    If nombreFicheroSalida.Contains("VZ") Then
                        REV1 = ""
                        REV2 = ""
                    End If
                End If


                ' ----------------------------------------------------------
                ' TextoPersonalizado1/2 
                ' ----------------------------------------------------------


                'TextoPersonalizado1 DENT. COMPLEMENTARIO  a Y PLUS DENTAL
                TextoPersonalizado2 = Replace(TextoPersonalizado2, "DENT. COMPLEMENTARIO", "Y PLUS DENTAL")

                Dim reglaTP = reglasTextoPersonalizado.FirstOrDefault(Function(x) _
(x.NumeroPoliza <> "" AndAlso x.NumeroPoliza = r.NumeroPoliza) _
                Or (x.TipoContrato <> "" AndAlso x.TipoContrato = r.TipoContrato))

                If reglaTP IsNot Nothing Then
                    TextoPersonalizado1 = reglaTP.Texto1
                    TextoPersonalizado2 = reglaTP.Texto2
                End If

                If nombreFicheroSalida.Contains("_TTL") Then
                    TextoPersonalizado1 = "Salud + Dental"
                End If
                If nombreFicheroSalida.Contains("_TPI") Then
                    TextoPersonalizado1 = "Salud + Dental"
                End If
                If nombreFicheroSalida.Contains("_TPT") Then
                    If r.ColectivoProducto.Contains("ADESLAS PLENA TOTAL VITAL") Then
                        TextoPersonalizado1 = "Salud + Dental"

                    End If
                End If


                ' ----------------------------------------------------------
                ' PISTAS 
                ' ----------------------------------------------------------
                Dim PISTA1 As String = "B8034460000000" & Left(r.NumeroTarjeta, 8) & "00=" &
                                   (Left(r.NombreApellidos & Space(30), 30)) & "=" &
                                   Mid(r.TipoContrato, 2) & (r.Version) &
                                   Right(r.FechaCaducidad, 2) & Left(r.FechaCaducidad, 2) &
                                   Right(r.FechaAlta, 4)



                Dim PISTA2 As String = "8034460000000" & Left(r.NumeroTarjeta, 8) & "00" &
                                   Right(r.NumeroTarjeta, 1) & "=" &
                                   Mid(r.TipoContrato, 2) &
                                   Right(r.FechaCaducidad, 2) & Left(r.FechaCaducidad, 2) &
                                   r.Version &
                                   Right(r.FechaInicioCarencia, 2) & Left(r.FechaInicioCarencia, 2)

                PISTA1 = Replace(PISTA1, "¥", "N")
                PISTA1 = Replace(PISTA1, "§", ".")
                PISTA1 = Replace(PISTA1, ",", " ")
                PISTA1 = Replace(PISTA1, "`", " ")
                PISTA1 = Replace(PISTA1, "'", " ")
                PISTA1 = Replace(PISTA1, "Ÿ", " ")



                ' ----------------------------------------------------------
                ' ColectivoProducto 
                ' ----------------------------------------------------------
                Dim ColectivoProducto As String = r.ColectivoProducto



                If ColectivoProducto.StartsWith("-") Then ColectivoProducto = ColectivoProducto.Substring(1)
                If ColectivoProducto.StartsWith(".") Then ColectivoProducto = ColectivoProducto.Substring(1)
                If ColectivoProducto.StartsWith(",") Then ColectivoProducto = ColectivoProducto.Substring(1)


                ColectivoProducto = Replace(ColectivoProducto, ",", " ")
                ColectivoProducto = Replace(ColectivoProducto, Chr(34), "")
                ColectivoProducto = Replace(ColectivoProducto, "Š", "U")
                ColectivoProducto = Replace(ColectivoProducto, "(", " ")
                ColectivoProducto = Replace(ColectivoProducto, ")", " ")
                ColectivoProducto = Replace(ColectivoProducto, "§", ".")


                ColectivoProducto = Trim(ColectivoProducto)

                ' Correcion Vida Caixa ColectivoProducto
                If r.Canal_Asegurador = "VC" Then
                    If nombreFicheroSalida.Contains("VC") Then

                        'ColectivoProducto = Replace(ColectivoProducto, "ADES EXTRA EMPRESAS Y DENTAL", "ADESLAS EMPRESAS Y DENTAL")
                        'ColectivoProducto = Replace(ColectivoProducto, "ADESLAS NEGOCIOS CIF", "ADESLAS NEGOCIOS")
                        'ColectivoProducto = Replace(ColectivoProducto, "ADES EXTRA EMPRESAS Y DENTAL", "EXTRA EMPRESAS Y DENTAL")
                        'ColectivoProducto = Replace(ColectivoProducto, "ADESLAS PYMES Y DE", "ADESLAS EMPRESAS Y DENTAL")
                        'ColectivoProducto = Replace(ColectivoProducto, "ADES EXTR NEGOCI  Y DENT", "ADESLAS EXTRA NEGOCIOS")
                        ColectivoProducto = Replace(ColectivoProducto, "NIF", "")
                        ColectivoProducto = Replace(ColectivoProducto, "CIF", "")

                        ColectivoProducto = Trim(ColectivoProducto)
                    End If

                    If nombreFicheroSalida.Contains("VU") Then
                        ColectivoProducto = Replace(ColectivoProducto, "VCS ", "")
                        If ColectivoProducto <> "ADESLAS PLENA PLUS" Then
                            If ColectivoProducto = "ADESLAS PLENA PLU" Then
                                ColectivoProducto = "ADESLAS PLENA"
                            End If
                        End If
                    End If
                    If nombreFicheroSalida.Contains("VZ") Then

                        If r.TipoContrato = "0982" Then
                            ColectivoProducto = "ADESLAS BASICO"
                        End If
                        If r.TipoContrato = "0983" Then
                            ColectivoProducto = "ADESLAS BASICO"
                        End If
                        If r.TipoContrato = "0993" Then
                            ColectivoProducto = "ADESLAS BASICO FAMILIA VCS"
                        End If
                        If r.TipoContrato = "0836" Then
                            'ColectivoProducto = ""
                        End If
                        If r.TipoContrato = "0994" Then
                            'ColectivoProducto = ""
                        End If

                    End If

                End If


                ' ----------------------------------------------------------
                ' Dirección / PRECEP / EMPRESA 
                ' ----------------------------------------------------------
                DireccionEnvio = r.DireccionEnvio
                DireccionEnvio = Replace(DireccionEnvio, "¥", "Ñ")
                DireccionEnvio = Replace(DireccionEnvio, "§", ".")
                DireccionEnvio = Replace(DireccionEnvio, "¦", ".")
                DireccionEnvio = Replace(DireccionEnvio, "(", " ")
                DireccionEnvio = Replace(DireccionEnvio, ")", " ")
                DireccionEnvio = Replace(DireccionEnvio, ",", " ")
                DireccionEnvio = Replace(DireccionEnvio, "'", " ")
                DireccionEnvio = Replace(DireccionEnvio, "Ï", " ")
                DireccionEnvio = Replace(DireccionEnvio, "Š", "U")
                DireccionEnvio = Replace(DireccionEnvio, "<", " ")
                DireccionEnvio = Replace(DireccionEnvio, "Ú", "-")
                DireccionEnvio = Replace(DireccionEnvio, ";", " ")
                DireccionEnvio = Replace(DireccionEnvio, "`", " ")
                DireccionEnvio = Replace(DireccionEnvio, "*", " ")
                DireccionEnvio = Replace(DireccionEnvio, "Ÿ", " ")
                DireccionEnvio = Replace(DireccionEnvio, "`", " ")
                DireccionEnvio = Replace(DireccionEnvio, Chr(34), " ")
                DireccionEnvio = Replace(DireccionEnvio, "È", "E")
                DireccionEnvio = Replace(DireccionEnvio, "_", " ")
                DireccionEnvio = Replace(DireccionEnvio, "¨", "E")
                DireccionEnvio = Replace(DireccionEnvio, "	", " ")
                DireccionEnvio = Replace(DireccionEnvio, "À", "A")
                DireccionEnvio = Replace(DireccionEnvio, "¤", "Ñ")
                DireccionEnvio = Replace(DireccionEnvio, ">", " ")
                DireccionEnvio = Replace(DireccionEnvio, "Ì", "I")
                DireccionEnvio = Replace(DireccionEnvio, "£", "U")
                DireccionEnvio = Replace(DireccionEnvio, "¡", "I")
                DireccionEnvio = Replace(DireccionEnvio, "¢", "O")
                DireccionEnvio = Replace(DireccionEnvio, "~", " ")
                DireccionEnvio = Replace(DireccionEnvio, "Ã", "A")




                DireccionEnvio = Trim(DireccionEnvio)

                Dim CD_POSTA_ENVIO As String = r.CodigoPostalEnvio

                PRECEP = r.PersonaReceptora
                PRECEP = Replace(PRECEP, ",", " ")
                PRECEP = Replace(PRECEP, "¥", "Ñ")
                PRECEP = Replace(PRECEP, "§", ".")
                PRECEP = Replace(PRECEP, "¦", ".")
                PRECEP = Replace(PRECEP, "(", " ")
                PRECEP = Replace(PRECEP, ")", " ")
                PRECEP = Replace(PRECEP, "'", " ")
                PRECEP = Replace(PRECEP, "Ï", " ")
                PRECEP = Replace(PRECEP, "Ú", "-")
                PRECEP = Replace(PRECEP, "`", " ")
                PRECEP = Replace(PRECEP, "Ÿ", " ")

                PRECEP = Replace(PRECEP, "¦", "")
                PRECEP = Replace(PRECEP, "*", " ")
                PRECEP = Replace(PRECEP, "¤", "Ñ")


                PRECEP = Trim(PRECEP)

                EMPRESA = ""
                EMPRESA = Replace(EMPRESA, ",", " ")




                If nombreFicheroSalida.Contains("_LIO") Then
                    PRECEP = Left(r.NombreApellidos & Space(28), 28)
                    DireccionEnvio = Trim(DireccionEnvio)
                    EMPRESA = r.PersonaReceptora
                End If

                ' ----------------------------------------------------------
                ' Población / Provincia 
                ' ----------------------------------------------------------
                'forzar a que poblacion tenga datos
                If PoblacionEnvio = "" Then

                    PoblacionEnvio = ProvinciaEnvio

                End If
                PoblacionEnvio = Replace(PoblacionEnvio, "¥", "Ñ")
                PoblacionEnvio = Replace(PoblacionEnvio, "€", "C")
                PoblacionEnvio = Replace(PoblacionEnvio, ",", " ")
                PoblacionEnvio = Replace(PoblacionEnvio, "'", " ")
                PoblacionEnvio = Replace(PoblacionEnvio, "Ï", " ")
                PoblacionEnvio = Replace(PoblacionEnvio, "(", " ")
                PoblacionEnvio = Replace(PoblacionEnvio, ")", " ")
                PoblacionEnvio = Replace(PoblacionEnvio, "Š", "U")
                PoblacionEnvio = Replace(PoblacionEnvio, "`", " ")
                PoblacionEnvio = Replace(PoblacionEnvio, "È", "E")
                PoblacionEnvio = Replace(PoblacionEnvio, "À", "A")



                PoblacionEnvio = Trim(PoblacionEnvio)



                ProvinciaEnvio = Replace(ProvinciaEnvio, "¥", "Ñ")
                ProvinciaEnvio = Replace(ProvinciaEnvio, ",", " ")
                ProvinciaEnvio = Replace(ProvinciaEnvio, "'", " ")
                ProvinciaEnvio = Replace(ProvinciaEnvio, "Ï", " ")
                ProvinciaEnvio = Replace(ProvinciaEnvio, "(", " ")
                ProvinciaEnvio = Replace(ProvinciaEnvio, ")", " ")

                ProvinciaEnvio = Trim(ProvinciaEnvio)

                Dim CPF As String = r.CPF
                Dim IND1 As String = r.IND1
                Dim IND2 As String = r.IND2
                Dim BENEF As String = If(String.IsNullOrWhiteSpace(r.BENEF), "", Left(r.BENEF & Space(28), 28))

                BENEF = Replace(BENEF, "'", " ")

                Dim BEN_TARJ As String = Mid(r.BEN_TARJ, 1, 8)

                '-----------------------------------------------------------
                'CREAR FORMATO PARA L00
                '-----------------------------------------------------------                
                If nombreFicheroSalida.Contains("_L00") Then


                    PRECEP = Left(r.NombreApellidos & Space(28), 28)
                    DireccionEnvio = Trim(DireccionEnvio)
                    EMPRESA = r.PersonaReceptora
                    EMPRESA = Replace(EMPRESA, ",", " ")

                    'If CPF = "2" And IND1 = "2" And IND2 = "L" Then
                    'PRECEP = Left(NOMBRE & Space(28), 28)
                    'DireccionEnvio = Trim(DireccionEnvio)
                    'EMPRESA = r.PersonaReceptora
                    'EMPRESA = Replace(EMPRESA, ",", " ")
                    'End If

                    CPF = "1"
                    IND1 = "1"
                    IND2 = "L"
                    BENEF = ""
                    BEN_TARJ = ""



                    COD_BARRAS = Cd_Campa & codDelegacionFormateado & r.NumeroPoliza.Trim() & r.NumeroCertificado.Trim() & r.NumeroOrden & "1" & "-" & Replace(nombreFicheroSalida, "_.", ".")

                End If

                ' ----------------------------------------------------------
                ' Fecha efecto
                ' ----------------------------------------------------------
                Dim F_EFECTO As String = CalcularFechaEfecto(r.FechaAlta)
                'REGULARIZAR FECHA EFECTO
                If CPF = "2" And IND1 = "1" Then
                    If F_EFECTO = "      " Then
                        F_EFECTO = ""
                    End If
                End If

                ' ----------------------------------------------------------
                ' CREAR FORMATO PARA MUFACE
                ' ----------------------------------------------------------
                If nombreFicheroSalida.Contains("_TMF") Or nombreFicheroSalida.Contains("_YMF") Then

                    EMB1 = EMB1.Trim()
                    NOMBRE = NOMBRE.PadRight(28)
                    ColectivoProducto = r.CIP_SNS.PadRight(28)
                    BENEF = ""
                    BEN_TARJ = ""
                    DMX = "01" & r.CIP_M & "02" & r.CIP_SNS & "032104" & NOMBRE & "!05!06!0700120XXX!"
                    DMX = Replace(DMX, "'", " ")
                    DireccionEnvio = DireccionEnvio.Trim()
                    F_EFECTO = F_EFECTO.PadRight(6)


                    Dim PAN_MUFACE As String = Mid(r.NumeroTarjeta, 1, 4) & " " &
                                Mid(r.NumeroTarjeta, 5, 4) & " " &
                                Mid(r.NumeroTarjeta, 9, 1) & "000 " &
                                Mid(codDelegacionFormateado, 1, 2) &
                                r.DigitoControlProvincia &
                                r.DigitoControlZ

                    PAN = PAN_MUFACE

                End If

                ' ----------------------------------------------------------
                ' CREAR FORMATO PARA ISFAS
                ' ----------------------------------------------------------
                If nombreFicheroSalida.Contains("_TIF") Or nombreFicheroSalida.Contains("_YIF") Then

                    EMB1 = EMB1.Trim()
                    NOMBRE = NOMBRE.PadRight(28)
                    ColectivoProducto = "CIP-SNS " & r.CIP_SNS.PadRight(20)
                    BENEF = ""
                    BEN_TARJ = ""
                    DMX = "01" & r.CIP_M & "02" & r.CIP_SNS & "032304" & NOMBRE & "!05!06!0700120XXX!"
                    DireccionEnvio = DireccionEnvio.Trim()
                    F_EFECTO = F_EFECTO.PadRight(6)


                    Dim PAN_ISFAS As String = Mid(r.NumeroTarjeta, 1, 4) & " " &
                                Mid(r.NumeroTarjeta, 5, 4) & " " &
                                Mid(r.NumeroTarjeta, 9, 1) & "000 " &
                                Mid(codDelegacionFormateado, 1, 2) &
                                r.DigitoControlProvincia &
                                r.DigitoControlZ

                    PAN = PAN_ISFAS
                    CIP_M = r.CIP_M
                    CIP_M = "CIP-M: " & Mid(CIP_M, 1, 4) & " " & Mid(CIP_M, 5, 4) & " " & Mid(CIP_M, 9, 4) & " " & Mid(CIP_M, 13, 4)

                    'CAMBIAR DIRECION DE ENTREGA

                    If r.Direccion.Contains("AV. TENIENTE GENERAL GABELLA, 1 (ACADEMI") Then
                        DireccionEnvio = "AVDA. MADRID, 20 - BAJO"
                        CD_POSTA_ENVIO = "23003"
                        PoblacionEnvio = "JAEN"
                        ProvinciaEnvio = "JAEN"
                    End If
                    If r.Direccion.Contains("AV. TENIENTE GENERAL GABELLA, 1  BLQ-4") Then
                        DireccionEnvio = "AVDA. MADRID, 20 - BAJO"
                        CD_POSTA_ENVIO = "23003"
                        PoblacionEnvio = "JAEN"
                        ProvinciaEnvio = "JAEN"
                    End If
                    If r.Direccion.Contains("TENIENTE GENERAL GABELLA 1 ACADEMI") Then
                        DireccionEnvio = "AVDA. MADRID, 20 - BAJO"
                        CD_POSTA_ENVIO = "23003"
                        PoblacionEnvio = "JAEN"
                        ProvinciaEnvio = "JAEN"
                    End If
                    If r.Direccion.Contains("TENIENTE GENERAL GABELLA 40") Then
                        DireccionEnvio = "AVDA. MADRID, 20 - BAJO"
                        CD_POSTA_ENVIO = "23003"
                        PoblacionEnvio = "JAEN"
                        ProvinciaEnvio = "JAEN"
                    End If


                    If r.Direccion.Contains("CT. MERIDA CAMPAMENTO SANTA ANA, S/N (CE") Then
                        DireccionEnvio = "Campamento Santa Ana. Carretera N 630 Km 558 direccion Merida"
                        CD_POSTA_ENVIO = "10195"
                        PoblacionEnvio = "CACERES"
                        ProvinciaEnvio = "CACERES"
                    End If


                    ' ----------------------------------------------------------
                End If


                ' ----------------------------------------------------------
                ' CREAR FORMATO PARA ISFAS-C
                ' ----------------------------------------------------------
                If nombreFicheroSalida.Contains("_TIZ") Then

                    EMB1 = EMB1.Trim()
                    NOMBRE = NOMBRE.PadRight(28)
                    ColectivoProducto = "CIP-SNS " & r.CIP_SNS.PadRight(20)
                    BENEF = ""
                    BEN_TARJ = ""
                    DMX = "01" & r.CIP_M & "02" & r.CIP_SNS & "032304" & NOMBRE & "!05!06!0700120XXX!"
                    DireccionEnvio = DireccionEnvio.Trim()
                    F_EFECTO = F_EFECTO.PadRight(6)


                    Dim PAN_ISFAS As String = Mid(r.NumeroTarjeta, 1, 4) & " " &
                                Mid(r.NumeroTarjeta, 5, 4) & " " &
                                Mid(r.NumeroTarjeta, 9, 1) & "000 " &
                                Mid(codDelegacionFormateado, 1, 2) &
                                r.DigitoControlProvincia &
                                r.DigitoControlZ

                    ' ----------------------------------------------------------
                    ' PISTAS --cambia --- PARA ISFAS-C HAY QUE FORZAR EL TIPO CONTA A 007 REVISAR CUANDO SALGA ESTE MODELO
                    ' ----------------------------------------------------------
                    PISTA1 = "B8034460000000" & Left(r.NumeroTarjeta, 8) & "00=" &
                                   (Left(r.NombreApellidos & Space(30), 30)) & "=" &
                    "007" & (r.Version) &
                    Right(r.FechaCaducidad, 2) & Left(r.FechaCaducidad, 2) &
                    Right(r.FechaAlta, 4)



                    PISTA2 = "8034460000000" & Left(r.NumeroTarjeta, 8) & "00" &
                    Right(r.NumeroTarjeta, 1) & "=" &
                    "007" &
                    Right(r.FechaCaducidad, 2) & Left(r.FechaCaducidad, 2) &
                    r.Version &
                    Right(r.FechaInicioCarencia, 2) & Left(r.FechaInicioCarencia, 2)

                    PISTA1 = Replace(PISTA1, "¥", "N")

                    PAN = PAN_ISFAS
                    CIP_M = r.CIP_M
                    CIP_M = "CIP-M: " & Mid(CIP_M, 1, 4) & " " & Mid(CIP_M, 5, 4) & " " & Mid(CIP_M, 9, 4) & " " & Mid(CIP_M, 13, 4)

                    'CAMBIAR DIRECION DE ENTREGA

                    If r.Direccion.Contains("AV. TENIENTE GENERAL GABELLA, 1 (ACADEMI") Then
                        DireccionEnvio = "AVDA. MADRID, 20 - BAJO"
                        CD_POSTA_ENVIO = "23003"
                        PoblacionEnvio = "JAEN"
                        ProvinciaEnvio = "JAEN"
                    End If


                    ' ----------------------------------------------------------
                End If
                ' ----------------------------------------------------------
                ' CREAR FORMATO PARA MUTUALIDAD GENERAL J   
                ' ----------------------------------------------------------
                If nombreFicheroSalida.Contains("_TFU") Then

                    EMB1 = EMB1.Trim()
                    NOMBRE = NOMBRE.PadRight(28)
                    ColectivoProducto = ColectivoProducto.PadRight(28)
                    REV1 = "Tlf. Urgencias"
                    REV2 = "900 322 237"

                    DireccionEnvio = DireccionEnvio.Trim()
                    F_EFECTO = F_EFECTO.PadRight(6)

                    Dim PAN_MGJ As String = Mid(r.NumeroTarjeta, 1, 4) & " " &
                                Mid(r.NumeroTarjeta, 5, 4) & " " &
                                Mid(r.NumeroTarjeta, 9, 1) & "000 " &
                                Mid(codDelegacionFormateado, 1, 2) &
                                r.DigitoControlProvincia &
                                r.DigitoControlZ

                    PAN = PAN_MGJ

                End If

                ' ----------------------------------------------------------
                ' CREAR FORMATO PARA _TVI  
                ' ----------------------------------------------------------
                If nombreFicheroSalida.Contains("_TVI") Then
                    Dim archivo As String = Replace(nombreFicheroSalida, "_.TXT", ".TXT")
                    REV1 = "Asistencia en viaje"
                    REV2 = "34.91.745.32.80"
                    COD_BARRAS = Cd_Campa & codDelegacionFormateado & r.NumeroPoliza.Trim() & r.NumeroCertificado.Trim() & r.NumeroOrden & "1-" & archivo
                    CPF = "1"
                    IND1 = "1"
                    IND2 = "L"
                    BENEF = ""
                    BEN_TARJ = ""
                    PLATAFORMA = "1"

                    If regla.FOLLETO = "TRIPTICCUN" Then
                        regla.FOLLETO = "TripticCUN"
                    End If
                    If Mid(CENT_COD, 1, 3) = "001" Then
                        PLATAFORMA = " "
                    End If

                End If
                ' ----------------------------------------------------------
                ' CREAR FORMATO PARA _EVI  
                ' ----------------------------------------------------------
                If r.Canal_Asegurador = "AD" Then
                    If nombreFicheroSalida.Contains("_EVI") Then
                        Dim archivo As String = Replace(nombreFicheroSalida, "_.TXT", ".TXT")
                        REV1 = "Asistencia en viaje"
                        REV2 = "34.91.745.32.80"
                        DireccionEnvio = "OFICINA CAIXA NUM: " & CENT_COD
                        CD_POSTA_ENVIO = ""
                        PoblacionEnvio = ""
                        ProvinciaEnvio = ""
                        COD_BARRAS = "GS00" & Replace(CENT_COD, "-", "0")

                        'If r.TipoContrato = "0824" Then
                        'COD_BARRAS = "GQ00" & Replace(CENT_COD, "-", "0")
                        'End If


                        CPF = "1"
                        IND1 = "1"
                        IND2 = "L"
                        BENEF = ""
                        BEN_TARJ = ""
                        PLATAFORMA = "1"

                        If regla.FOLLETO = "TRIPTICCUN" Then
                            regla.FOLLETO = "TripticCUN"
                        End If
                    End If
                End If
                ' ----------------------------------------------------------
                ' CREAR FORMATO PARA _LPM  
                ' ----------------------------------------------------------
                'pendiente corregir nombre de la empresa al forzar que sena todos de 1 se queda con el nombre del paquete de 2
                If nombreFicheroSalida.Contains("_LPM") Then



                    PRECEP = Left(r.NombreApellidos & Space(28), 28)
                    DireccionEnvio = Trim(DireccionEnvio)
                    'EMPRESA = r.PersonaReceptora
                    EMPRESA = NOMBRE
                    EMPRESA = Replace(EMPRESA, ",", " ")
                    CPF = "1"
                    IND1 = "1"
                    IND2 = "L"
                    BENEF = ""
                    BEN_TARJ = ""

                    COD_BARRAS = Cd_Campa & codDelegacionFormateado & r.NumeroPoliza.Trim() & r.NumeroCertificado.Trim() & r.NumeroOrden & "1" & "-" & Replace(nombreFicheroSalida, "_.", ".")


                    TextoPersonalizado1 = Replace(TextoPersonalizado1, "SALUD+DENTAL", "Salud + Dental")

                End If

                ' ----------------------------------------------------------
                ' CREAR FORMATO PARA _TIV  (IMQ P. VASCO)
                ' ----------------------------------------------------------
                ' PENDIENTE DE PREGUNTAR EL FORMATO DEL PAN Y DE LA PISTA 1 Y PISTA 2
                If nombreFicheroSalida.Contains("_TIV") Then

                    'para generar el último dígito del PAN en este modelo
                    Dim ultimoDigitoPan1 As String
                    Dim ultimoDigitoPan2 As String
                    ultimoDigitoPan1 = r.NumeroTarjeta & 480444
                    ultimoDigitoPan2 = CalcularDigitoControl(ultimoDigitoPan1)
                    'PAN = Mid(r.NumeroTarjeta, 1, 4) & " " & Mid(r.NumeroTarjeta, 5, 4) & " " & Mid(r.NumeroTarjeta, 9, 1) & "480 444" & r.DigitoControlZ
                    PAN = Mid(r.NumeroTarjeta, 1, 4) & " " & Mid(r.NumeroTarjeta, 5, 4) & " " & Mid(r.NumeroTarjeta, 9, 1) & "480 444" & ultimoDigitoPan2
                    NOMBRE = NOMBRE.PadRight(28)
                    ColectivoProducto = ColectivoProducto.PadRight(28)

                    If Mid(CD_POSTA_ENVIO, 1, 2) = "01" Then
                        Mid(EMB1, 17, 2) = "IA"
                    End If
                    If Mid(CD_POSTA_ENVIO, 1, 2) = "20" Then
                        Mid(EMB1, 17, 2) = "IG"
                    End If
                    If Mid(CD_POSTA_ENVIO, 1, 2) = "48" Then
                        Mid(EMB1, 17, 2) = "IB"
                    End If

                    PISTA1 = "B8034440" & Left(r.NumeroTarjeta, 8) & "04804440" & "^" & NOMBRE & Space(2) & "^" & "003" & "1" & "6012" & Right(r.FechaAlta, 4)
                    PISTA1 = Replace(PISTA1, "¥", "N")

                    PISTA2 = "8034440" & Left(r.NumeroTarjeta, 8) & "04804440" & Right(r.NumeroTarjeta, 1) & "=" & "003" & "6012" & "1" & Right(r.FechaInicioCarencia, 2) & Left(r.FechaInicioCarencia, 2)


                    REV1 = "Tlf. Atención en Extranjero"
                    REV2 = "34-91-7453280"
                    REV3 = "Tlf.Atención al Cliente en P.Vasco"
                    REV4 = "900-81-81-50"
                    TextoPersonalizado1 = TextoPersonalizado1.PadRight(21)
                    TextoPersonalizado2 = TextoPersonalizado2.PadRight(21)
                End If
                ' ----------------------------------------------------------
                ' CREAR FORMATO PARA _TAW  
                ' ----------------------------------------------------------
                If nombreFicheroSalida.Contains("_TAW") Then
                    If r.TipoContrato = "0877" Then
                        REV1 = "REPATRIACION"
                        REV2 = "AL PAIS DE ORIGEN"
                        REV3 = "TELÉFONO:"
                        REV4 = "91 745 32 80"
                    End If
                End If
                ' ----------------------------------------------------------
                ' CREAR FORMATO PARA _EEX  
                ' ----------------------------------------------------------
                If nombreFicheroSalida.Contains("_EEX") Then
                    ColectivoProducto = ColectivoProducto & " ID." & Trim(r.NumeroCertificado)
                End If
                ' CARTERES ILEGALES ------------------------------------
                DireccionEnvio = Replace(DireccionEnvio, "¦", "")


                ASESOR = Convert.ToString(ASESOR)

                ASESOR = ASESOR.Replace("¦", "").Trim()

                If ASESOR.Length > 28 Then
                    ASESOR = ASESOR.Substring(0, 28)
                Else
                    ASESOR = ASESOR.PadRight(28)
                End If

                'campos.Add(ASESOR)


                DIR_ASESOR = Replace(DIR_ASESOR, "¦", "")
                DIR_ASESOR = Replace(DIR_ASESOR, "Ï", "ï")
                NOMBRE = Replace(NOMBRE, "§", ".")
                NOMBRE = Replace(NOMBRE, ",", "")
                NOMBRE = Replace(NOMBRE, "`", " ")
                NOMBRE = Replace(NOMBRE, "'", " ")
                NOMBRE = Replace(NOMBRE, "*", " ")


                EMPRESA = Replace(EMPRESA, "'", " ")

                ' ----------------------------------------------------------
                ' Dirección ENVIO VC OFICINAS EXTRACCION
                ' ----------------------------------------------------------

                If r.Canal_Asegurador = "VC" Then
                    If r.CentroTrabajoCodigo.Contains("001-") Then
                        DireccionEnvio = "OFICINA CAIXA NUM: " & r.CentroTrabajoCodigo
                        CD_POSTA_ENVIO = ""
                        PoblacionEnvio = ""
                        ProvinciaEnvio = ""
                        COD_BARRAS = "GQ00" & Replace(r.CentroTrabajoCodigo, "-", "0")
                        CENT_COD = " "
                    End If
                End If
                ' ----------------------------------------------------------

                '----------------------------------------------------------

                ' ----------------------------------------------------------
                ' ESCRITURA CAMPOS 
                ' ----------------------------------------------------------
                campos.Add(sec.ToString("D6"))          ' SECUENCIAL
                campos.Add(PAN)                         ' PAN
                campos.Add(EMB1)                        ' EMB1
                campos.Add(NOMBRE)                      ' NOMBRE
                campos.Add(ColectivoProducto)           ' COLEC
                campos.Add(REV1)                        ' REV1
                campos.Add(REV2)                        ' REV2
                campos.Add(REV3)                        ' REV3
                campos.Add(REV4)                        ' REV4
                campos.Add(TextoPersonalizado1)         ' ANV1
                campos.Add(TextoPersonalizado2)         ' ANV2
                campos.Add(r.NumeroTarjeta)             ' NTARJETA
                campos.Add(PISTA1)                      ' PISTA1 
                campos.Add(PISTA2)                      ' PISTA2 
                campos.Add(PRECEP)                      ' PRECEP
                campos.Add(Trim(DireccionEnvio))        ' DIRECCION
                campos.Add(CD_POSTA_ENVIO)              ' CPOSTAL
                campos.Add(PoblacionEnvio)              ' POBLACION
                campos.Add(ProvinciaEnvio)              ' PROVINCIA
                campos.Add(BENEF)                       ' BENEF
                campos.Add(BEN_TARJ)                    ' BEN_TARJ
                campos.Add(COD_BARRAS)                  ' COD_BARRAS
                campos.Add(CENT_COD)                    ' CENT_COD
                campos.Add(CENT_DES)                    ' CENT_DES
                campos.Add(F_EFECTO)                    ' F_EFECTO
                campos.Add(ASESOR.Trim())               ' ASESOR
                campos.Add(DIR_ASESOR)                  ' DIR_ASESOR
                campos.Add(TEL_ASESOR)                  ' TEL_ASESOR
                campos.Add(r.NumeroPoliza)              ' POLIZA
                campos.Add(r.NumeroCertificado)         ' CERTIFICA
                campos.Add(EMPRESA)                     ' EMPRESA
                campos.Add(CPF)                         ' CPF
                campos.Add(IND1)                        ' IND1
                campos.Add(IND2)                        ' IND2

                ' EXTENSION ARCHIVO = NOMBRE REAL DEL FICHERO DE SALIDA
                Dim quitar_txt As String = Replace(nombreFicheroSalida, "_.TXT", ".TXT")
                campos.Add(quitar_txt)         ' ARCHIVO 

                campos.Add(If(regla IsNot Nothing, regla.MODELO, ""))    ' MODELO
                campos.Add(If(regla IsNot Nothing, regla.TOPPER, ""))    ' TOPPER
                campos.Add(If(regla IsNot Nothing, regla.LOGO, ""))      ' LOGO
                campos.Add(If(regla IsNot Nothing, regla.ULTRAANV, ""))  ' ULTRAANV
                campos.Add(If(regla IsNot Nothing, regla.ULTRAREV, ""))  ' ULTRAREV
                campos.Add(If(regla IsNot Nothing, regla.CARRIER, ""))   ' CARRIER

                ' ===== SOBRE ===== Cambio de sobre si es HISPAPOST
                Dim sobreValor As String = ""
                If nombreFicheroSalida.Contains("_H") Then
                    sobreValor = "ADESLAS-HISPAPOST"
                Else
                    sobreValor = If(regla IsNot Nothing, regla.SOBRE, "")
                End If

                campos.Add(sobreValor)   ' SOBRE
                campos.Add(If(regla IsNot Nothing, regla.TARJETA, ""))   ' TARJETA
                campos.Add(If(regla IsNot Nothing, regla.FOLLETO, ""))   ' FOLLETO

                ' VC_PROD
                Dim vcProd As String = New String(" "c, 50)
                If regla IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(regla.VC_PROD) Then
                    vcProd = regla.VC_PROD
                End If
                campos.Add(vcProd)

                campos.Add(DMX) ' DMX

                ' EMB0
                If nombreFicheroSalida.Contains("_UAB") Then
                    campos.Add("NUM.AFIL." & r.NumeroCertificado.PadLeft(9, "0"c))
                Else
                    campos.Add(" ")
                End If

                ' PLATAFORMA
                If nombreFicheroSalida.Contains("_H") Or nombreFicheroSalida.Contains("_TVI") Then
                    campos.Add(If(String.IsNullOrWhiteSpace(PLATAFORMA), " ", "P-" & PLATAFORMA))
                Else
                    campos.Add(" ")
                End If

                campos.Add(If(String.IsNullOrWhiteSpace(CIP_M), "", CIP_M)) ' CIP_M

                sw.WriteLine(String.Join(";", campos))
                sec += 1

            Next

        End Using
    End Sub




    Private Class ReglaTextoPersonalizado
        Public NumeroPoliza As String
        Public TipoContrato As String
        Public Texto1 As String
        Public Texto2 As String
    End Class
    Private Function ExtraerSufijoDesdeEntrada(nombreEntrada As String) As String
        If String.IsNullOrWhiteSpace(nombreEntrada) Then Return ""

        Dim base As String = Path.GetFileNameWithoutExtension(nombreEntrada).ToUpperInvariant()

        ' Busca T + 7 u 8 dígitos
        Dim m As Match = Regex.Match(base, "T(\d{7,8})")

        If m.Success Then
            Return m.Groups(1).Value
        End If

        Return ""
    End Function

    Private Function CargarReglasTextoPersonalizado() As List(Of ReglaTextoPersonalizado)

        Dim lista As New List(Of ReglaTextoPersonalizado)
        Dim rutaBDCompleta As String = Path.Combine(RutaBD, "ADESLAS.accdb")

        Dim connectionString As String =
            $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={rutaBDCompleta};Persist Security Info=False;"


        Using cn As New OleDbConnection(ConnectionString)
            cn.Open()

            Dim sql As String =
            "SELECT NumeroPoliza, TipoContrato, Texto1, Texto2 " &
            "FROM TextoPesonalizado WHERE Activo=True"

            Using cmd As New OleDbCommand(sql, cn)
                Using rd = cmd.ExecuteReader()
                    While rd.Read()
                        lista.Add(New ReglaTextoPersonalizado With {
                        .NumeroPoliza = Texto(rd("NumeroPoliza")),
                        .TipoContrato = Texto(rd("TipoContrato")),
                        .Texto1 = Texto(rd("Texto1")),
                        .Texto2 = Texto(rd("Texto2"))
                    })
                    End While
                End Using
            End Using
        End Using

        Return lista
    End Function


    Private Function CalcularFechaEfecto(fechaAlta As String) As String

        ' Validación básica
        If String.IsNullOrWhiteSpace(fechaAlta) Then Return ""

        fechaAlta = fechaAlta.Trim()

        ' Debe venir como YYYYMM
        If fechaAlta.Length <> 6 OrElse Not IsNumeric(fechaAlta) Then
            Return ""
        End If

        Dim anio As Integer = CInt(fechaAlta.Substring(0, 4))
        Dim mes As Integer = CInt(fechaAlta.Substring(4, 2))

        ' Fecha calculada: día 01 del mes/año indicado
        Dim fechaEfecto As New Date(anio, mes, 1)

        ' Primer día del mes actual
        Dim hoy As Date = Date.Today
        Dim inicioMesActual As New Date(hoy.Year, hoy.Month, 1)

        ' 👉 DESDE el mes actual (incluido)
        If fechaEfecto >= inicioMesActual Then
            Return "el dia: " & fechaEfecto.ToString("dd/MM/yyyy")
        Else
            Return "      "
        End If

    End Function


    Private Function ObtenerBarrasControl(codBarras As String) As String

        If String.IsNullOrWhiteSpace(codBarras) Then Return ""

        codBarras = codBarras.Trim()

        If Not IsNumeric(codBarras) Then Return ""

        ' Último dígito
        Return codBarras.Substring(codBarras.Length - 1, 1)

    End Function


    Private Function GradoEspecificidad(regla As ReglaExtraccion) As Integer

        Dim puntos As Integer = 0

        If regla.ColectivoProducto <> "" Then puntos += 1
        If regla.LogotipoDescripcion <> "" Then puntos += 1
        If regla.Direccion <> "" Then puntos += 1
        If regla.DireccionEnvio <> "" Then puntos += 1
        If regla.ModeloCarrier <> "" Then puntos += 1
        If regla.Paquetizado <> "" Then puntos += 1
        If regla.TipoContraPymes Then puntos += 1
        If regla.CodigoDelegacion <> "" Then puntos += 1
        If regla.NumeroPoliza <> "" Then puntos += 1
        If regla.BuscarLogo Then puntos += 1


        If Texto(regla.TipoContrato) <> "" Then puntos += 1
        If Texto(regla.CentroTrabajoCodigo) <> "" Then puntos += 1
        If Texto(regla.IndicadorIdioma) <> "" Then puntos += 1

        Return puntos
    End Function


    ' ============================================================
    ' MOTOR DE COMPARACIÓN ESCALABLE
    ' ============================================================

    Private Enum TipoComparacion
        Igual
        Contiene
        ExisteEnLogos
        EmpiezaPor
    End Enum

    Private Class ComparacionCampo
        Public Nombre As String          ' ✅ NUEVO
        Public CampoRegla As Func(Of ReglaExtraccion, String)
        Public CampoRegistro As Func(Of RegistroTarjeta, String)
        Public Tipo As TipoComparacion
    End Class


    Private Function AplicarOrdenacion(
    lista As List(Of RegistroTarjeta),
    orden As String
) As List(Of RegistroTarjeta)

        If String.IsNullOrWhiteSpace(orden) Then Return lista

        Dim campos = orden.Split(","c).
            Select(Function(x) x.Trim()).
            Where(Function(x) x <> "").
            ToList()

        Dim ordenado As IOrderedEnumerable(Of RegistroTarjeta) = Nothing

        For i = 0 To campos.Count - 1

            Dim partes = campos(i).Split(" "c)
            Dim campo = partes(0).ToUpperInvariant()
            Dim desc = (partes.Length > 1 AndAlso partes(1).ToUpperInvariant() = "DESC")

            Dim clave As Func(Of RegistroTarjeta, Object) = Nothing

            Select Case campo
                Case "NUMEROPOLIZA"
                    clave = Function(r) Val(r.NumeroPoliza)
                Case "NUMEROCERTIFICADO"
                    clave = Function(r) Val(r.NumeroCertificado)
                Case "CODIGORELACION"
                    clave = Function(r) Val(r.CodigoRelacion)
                Case "NUMEROORDEN"
                    clave = Function(r) Val(r.NumeroOrden)
                Case "CENTROTRABAJOCODIGO", "CENTROTRABAJO", "CENTRO"
                    clave = Function(r) ObtenerCentroTrabajoOrdenable(r.CentroTrabajoCodigo)


            End Select


            If clave Is Nothing Then Continue For

            If ordenado Is Nothing Then
                ordenado = If(desc,
                    lista.OrderByDescending(clave),
                    lista.OrderBy(clave))
            Else
                ordenado = If(desc,
                    ordenado.ThenByDescending(clave),
                    ordenado.ThenBy(clave))
            End If

        Next

        Return If(ordenado Is Nothing, lista, ordenado.ToList())

    End Function
    Private Function ObtenerCentroTrabajoOrdenable(valor As String) As Long
        If String.IsNullOrWhiteSpace(valor) Then Return Long.MaxValue

        Dim v As String = valor.Trim().ToUpperInvariant()

        ' Queremos que los 001-xxxx vayan primero
        Dim prioridad As Long = If(v.StartsWith("001-") OrElse v.StartsWith("001"), 0L, 1L)

        ' Solo números: "001-7071" -> "0017071"
        Dim soloNumeros As String = Regex.Replace(v, "\D", "")
        If soloNumeros = "" Then Return Long.MaxValue

        Dim n As Long
        If Not Long.TryParse(soloNumeros, n) Then Return Long.MaxValue

        ' Clave combinada: prioridad + número (para que 001- vaya antes que otros)
        Return prioridad * 1000000000000L + n
    End Function


    Private Function AplicarOrdenacionHispapost(
    lista As List(Of RegistroTarjeta)
) As List(Of RegistroTarjeta)

        Return lista _
        .OrderBy(Function(r) ObtenerNumeroSeguro(r.PLATAFORMA)) _
        .ThenBy(Function(r) ObtenerNumeroZona(r.ZONA)) _
        .ThenBy(Function(r) ObtenerCodigoPostalSeguro(r.CodigoPostalEnvio)) _
        .ToList()
    End Function
    Private Function ObtenerNumeroSeguro(valor As String) As Integer
        If String.IsNullOrWhiteSpace(valor) Then Return 0
        Dim soloNumeros = Regex.Replace(valor, "\D", "")
        If soloNumeros = "" Then Return 0

        Return Integer.Parse(soloNumeros)
    End Function

    Private Function ObtenerCodigoPostalSeguro(cp As String) As Integer
        If String.IsNullOrWhiteSpace(cp) Then Return 0

        Dim soloNumeros = Regex.Replace(cp, "\D", "")
        If soloNumeros = "" Then Return 0

        Return Integer.Parse(soloNumeros)
    End Function



    Private Function ObtenerNumeroZona(zona As String) As Integer
        If String.IsNullOrWhiteSpace(zona) Then Return 0

        ' Quita cualquier letra: D1 → 1 | C12 → 12
        Dim soloNumeros As String = Regex.Replace(zona, "\D", "")
        If soloNumeros = "" Then Return 0

        Return Integer.Parse(soloNumeros)
    End Function


    Public Function CalcularDigitoControl(base As String) As Integer

        base = New String(base.Where(AddressOf Char.IsDigit).ToArray())

        Dim sumaPares As Integer = 0
        Dim sumaImpares As Integer = 0

        For i As Integer = 0 To base.Length - 1
            Dim d As Integer = CInt(base(i).ToString())
            Dim posBase As Integer = i + 1

            If posBase Mod 2 = 0 Then
                ' POSICIONES PARES: se suman tal cual
                sumaPares += d
            Else
                ' POSICIONES IMPARES: se multiplican x2 
                d *= 2

                sumaImpares += d
            End If
        Next

        'Suma total de pares e impares
        Dim sumaTotal As Integer = sumaPares + sumaImpares

        'Calcula el dígito de control (último dígito) para que el total final sea múltiplo de 10.
        Dim dc As Integer = (10 - (sumaTotal Mod 10)) Mod 10
        'Dim dc As Integer = 10 - (sumaTotal Mod 10)
        Return dc

    End Function


    ' ============================================================
    ' RELLENAR ProvinciaEnvio DESDE CP (con SELECT de control previo)
    ' ============================================================
    Public Sub RellenarProvinciaEnvioDesdeCP()

        Dim AccessDBPath As String = RutaBD & "\ADESLAS.accdb"
        Dim cs As String = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={AccessDBPath};Persist Security Info=False;"

        Using conn As New OleDbConnection(cs)
            conn.Open()

            ' ------------------------------------------------------------
            ' 1) SELECT DE CONTROL (cuántos se actualizarían)
            ' ------------------------------------------------------------
            Dim sqlSelect As String =
            "SELECT COUNT(*) " &
            "FROM TarjetasSanitariasDiarioInteramit " &
            "INNER JOIN Provincias " &
            "ON Left(TarjetasSanitariasDiarioInteramit.CodigoPostalEnvio,2) = Provincias.codProv " &
            "WHERE (TarjetasSanitariasDiarioInteramit.ProvinciaEnvio Is Null " &
            "       OR Trim(TarjetasSanitariasDiarioInteramit.ProvinciaEnvio)='') " &
            "AND TarjetasSanitariasDiarioInteramit.CodigoPostalEnvio Is Not Null " &
            "AND Len(Trim(TarjetasSanitariasDiarioInteramit.CodigoPostalEnvio))>=2"

            Dim totalPrevistos As Integer

            Using cmdSelect As New OleDbCommand(sqlSelect, conn)
                totalPrevistos = Convert.ToInt32(cmdSelect.ExecuteScalar())
            End Using

            ' ------------------------------------------------------------
            ' 2) UPDATE REAL
            ' ------------------------------------------------------------
            Dim sqlUpdate As String =
            "UPDATE TarjetasSanitariasDiarioInteramit " &
            "INNER JOIN Provincias " &
            "ON Left(TarjetasSanitariasDiarioInteramit.CodigoPostalEnvio,2) = Provincias.codProv " &
            "SET TarjetasSanitariasDiarioInteramit.ProvinciaEnvio = Provincias.NombreProv " &
            "WHERE (TarjetasSanitariasDiarioInteramit.ProvinciaEnvio Is Null " &
            "       OR Trim(TarjetasSanitariasDiarioInteramit.ProvinciaEnvio)='') " &
            "AND TarjetasSanitariasDiarioInteramit.CodigoPostalEnvio Is Not Null " &
            "AND Len(Trim(TarjetasSanitariasDiarioInteramit.CodigoPostalEnvio))>=2"

            Dim totalActualizados As Integer

            Using cmdUpdate As New OleDbCommand(sqlUpdate, conn)
                totalActualizados = cmdUpdate.ExecuteNonQuery()
            End Using

            ' ------------------------------------------------------------
            ' 3) INFORMACIÓN (opcional)
            ' ------------------------------------------------------------
            MessageBox.Show(
            $"ProvinciaEnvio completada desde CP." & vbCrLf &
            $"Previstos: {totalPrevistos}" & vbCrLf &
            $"Actualizados: {totalActualizados}",
            "Relleno de provincias",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        )

        End Using

    End Sub

    ' ============================================================
    ' CAMBIO DE MODELO:
    ' - NumeroPoliza empieza por 666
    ' - TipoContrato = 640 (acepta 0640, " 0640 ", etc.)
    ' => ModeloPlasticoCodigo = "33"
    ' ============================================================
    Public Sub CambioDeModelo_666_TipoContrato640_A_Modelo33()

        Dim AccessDBPath As String = RutaBD & "\ADESLAS.accdb"
        Dim cs As String = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={AccessDBPath};Persist Security Info=False;"

        Using conn As New OleDbConnection(cs)
            conn.Open()

            ' 1) SELECT de control (cuántos cumplen la condición)
            Dim sqlSelect As String =
                "SELECT COUNT(*) " &
                "FROM TarjetasSanitariasDiarioInteramit " &
                "WHERE Left(Trim(NumeroPoliza),3)='666' " &
                "AND Val(TipoContrato)=640 " &
                "AND (ModeloPlasticoCodigo Is Null OR Trim(ModeloPlasticoCodigo) <> '33')"

            Dim previstos As Integer
            Using cmdSelect As New OleDbCommand(sqlSelect, conn)
                previstos = Convert.ToInt32(cmdSelect.ExecuteScalar())
            End Using

            ' 2) UPDATE real
            Dim sqlUpdate As String =
                "UPDATE TarjetasSanitariasDiarioInteramit " &
                "SET ModeloPlasticoCodigo='33' " &
                "WHERE Left(Trim(NumeroPoliza),3)='666' " &
                "AND Val(TipoContrato)=640 " &
                "AND (ModeloPlasticoCodigo Is Null OR Trim(ModeloPlasticoCodigo) <> '33')"

            Dim actualizados As Integer
            Using cmdUpdate As New OleDbCommand(sqlUpdate, conn)
                actualizados = cmdUpdate.ExecuteNonQuery()
            End Using

            MessageBox.Show(
                $"Cambio de modelo (Poliza 666* + TipoContrato 640/0640 → Modelo 33)" & vbCrLf &
                $"Previstos: {previstos}" & vbCrLf &
                $"Actualizados: {actualizados}",
                "Cambio de modelo",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            )

        End Using

    End Sub

    Public Sub CorrigeDireccionEnvio()

        Dim AccessDBPath As String = RutaBD & "\ADESLAS.accdb"
        Dim cs As String = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={AccessDBPath};Persist Security Info=False;"

        Using conn As New OleDbConnection(cs)
            conn.Open()

            ' Condición: DireccionEnvio vacía (NULL o espacios)
            Dim whereCond As String =
                "(DireccionEnvio Is Null OR Trim(DireccionEnvio)='')"

            ' 1) SELECT de control
            Dim sqlSelect As String =
                "SELECT COUNT(*) " &
                "FROM TarjetasSanitariasDiarioInteramit " &
                "WHERE " & whereCond

            Dim previstos As Integer
            Using cmdSelect As New OleDbCommand(sqlSelect, conn)
                previstos = Convert.ToInt32(cmdSelect.ExecuteScalar())
            End Using

            ' 2) UPDATE real (un solo SET con comas) + WHERE
            Dim sqlUpdate As String =
                "UPDATE TarjetasSanitariasDiarioInteramit " &
                "SET " &
                "  PersonaReceptora = NombreApellidos, " &
                "  DireccionEnvio = Direccion, " &
                "  CodigoPostalEnvio = CodigoPostal, " &
                "  PoblacionEnvio = Poblacion, " &
                "  ProvinciaEnvio = Provincia " &
                "WHERE " & whereCond

            Dim actualizados As Integer
            Using cmdUpdate As New OleDbCommand(sqlUpdate, conn)
                actualizados = cmdUpdate.ExecuteNonQuery()
            End Using

            MessageBox.Show(
                $"Corrección DIRECCION ENVIO (rellenar desde datos normales)" & vbCrLf &
                $"Previstos: {previstos}" & vbCrLf &
                $"Actualizados: {actualizados}",
                "Corrección Dirección Envío",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            )

        End Using

    End Sub

    ' ============================================================
    ' MARCAR Canal_Sanitario: VC si NumeroPoliza empieza por 88, si no AD
    ' ============================================================
    Public Sub MarcarCanalSanitario()

        Dim AccessDBPath As String = RutaBD & "\ADESLAS.accdb"
        Dim cs As String = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={AccessDBPath};Persist Security Info=False;"

        Using conn As New OleDbConnection(cs)
            conn.Open()

            ' 1) SELECT de control (cuántos hay de cada tipo)
            Dim sqlVC As String =
            "SELECT COUNT(*) FROM TarjetasSanitariasDiarioInteramit " &
            "WHERE Left(Trim('' & NumeroPoliza),2)='88'"

            Dim sqlAD As String =
            "SELECT COUNT(*) FROM TarjetasSanitariasDiarioInteramit " &
            "WHERE Left(Trim('' & NumeroPoliza),2)<>'88'"


            Dim nVC As Integer
            Dim nAD As Integer

            Using cmd As New OleDbCommand(sqlVC, conn)
                nVC = Convert.ToInt32(cmd.ExecuteScalar())
            End Using
            Using cmd As New OleDbCommand(sqlAD, conn)
                nAD = Convert.ToInt32(cmd.ExecuteScalar())
            End Using

            ' 2) UPDATE real (en una sola sentencia con IIf)
            Dim sqlUpdate As String =
            "UPDATE TarjetasSanitariasDiarioInteramit " &
            "SET Canal_Asegurador = IIf(Left(Trim('' & NumeroPoliza),2)='88','VC','AD')"


            Dim actualizados As Integer
            Using cmdUpdate As New OleDbCommand(sqlUpdate, conn)
                actualizados = cmdUpdate.ExecuteNonQuery()
            End Using

            MessageBox.Show(
            $"Canal_Adegurador actualizado." & vbCrLf &
            $"VC (poliza 88*): {nVC}" & vbCrLf &
            $"AD (resto): {nAD}" & vbCrLf &
            $"Registros afectados: {actualizados}",
            "Canal_Adegurador",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        )

        End Using

    End Sub

    ' ============================================================
    ' FORZAR DATOS VC (por TipoContrato + IndicadorIdioma)
    ' Tabla origen: VC_FORZAR_PL_ID_T1_T2_TEFex
    ' Solo aplica si Canal_Asegurador = 'VC'
    ' ============================================================
    Public Sub ForzarVC_PL_ID_T1_T2_TEFex()

        Dim AccessDBPath As String = RutaBD & "\ADESLAS.accdb"
        Dim cs As String = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={AccessDBPath};Persist Security Info=False;"

        Using conn As New OleDbConnection(cs)
            conn.Open()

            Dim sqlSelect As String =
                "SELECT COUNT(*) " &
                "FROM TarjetasSanitariasDiarioInteramit AS T " &
                "INNER JOIN VC_FORZAR_PL_ID_T1_T2_TEFex AS F " &
                "  ON Val(T.TipoContrato)=Val(F.TipoContrato) " &
                " AND Val(T.IndicadorIdioma)=Val(F.IndicadorIdioma) " &
                "WHERE T.Canal_Asegurador='VC' " &
                "AND ( " &
                "  Trim('' & F.ModeloPlasticoCodigo)<>'' OR " &
                "  Trim('' & F.ColectivoProducto)<>'' OR " &
                "  Trim('' & F.TextoPersonalizado1)<>'' OR " &
                "  Trim('' & F.TextoPersonalizado2)<>'' OR " &
                "  Trim('' & F.IndicadorExtranjero)<>'' " &
                ")"

            Dim previstos As Integer
            Using cmdSelect As New OleDbCommand(sqlSelect, conn)
                previstos = Convert.ToInt32(cmdSelect.ExecuteScalar())
            End Using

            Dim sqlUpdate As String =
                "UPDATE TarjetasSanitariasDiarioInteramit AS T " &
                "INNER JOIN VC_FORZAR_PL_ID_T1_T2_TEFex AS F " &
                "  ON Val(T.TipoContrato)=Val(F.TipoContrato) " &
                " AND Val(T.IndicadorIdioma)=Val(F.IndicadorIdioma) " &
                "SET " &
                "  T.ModeloPlasticoCodigo = IIf(Trim('' & F.ModeloPlasticoCodigo)='', T.ModeloPlasticoCodigo, Left(Trim(F.ModeloPlasticoCodigo),2)), " &
                "  T.ColectivoProducto    = IIf(Trim('' & F.ColectivoProducto)='',    T.ColectivoProducto,    Left(Trim(F.ColectivoProducto),28)), " &
                "  T.TextoPersonalizado1  = IIf(Trim('' & F.TextoPersonalizado1)='',  T.TextoPersonalizado1,  Left(Trim(F.TextoPersonalizado1),20)), " &
                "  T.TextoPersonalizado2  = IIf(Trim('' & F.TextoPersonalizado2)='',  T.TextoPersonalizado2,  Left(Trim(F.TextoPersonalizado2),20)), " &
                "  T.IndicadorExtranjero  = IIf(Trim('' & F.IndicadorExtranjero)='',  T.IndicadorExtranjero,  Left(Trim(F.IndicadorExtranjero),1)) " &
                "WHERE T.Canal_Asegurador='VC'"

            Dim actualizados As Integer
            Using cmdUpdate As New OleDbCommand(sqlUpdate, conn)
                actualizados = cmdUpdate.ExecuteNonQuery()
            End Using

            MessageBox.Show(
                $"Forzar VC (TipoContrato + IndicadorIdioma) aplicado." & vbCrLf &
                $"Previstos: {previstos}" & vbCrLf &
                $"Actualizados: {actualizados}",
                "VC_FORZAR_PL_ID_T1_T2_TEFex",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            )

        End Using

    End Sub

    Public Sub MarcarCentroDeTrabajoCaixaOficinVirtual()

        Dim AccessDBPath As String = RutaBD & "\ADESLAS.accdb"
        Dim cs As String = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={AccessDBPath};Persist Security Info=False;"

        Using conn As New OleDbConnection(cs)
            conn.Open()

            ' Actualiza SOLO los registros cuyo CentroTrabajoCodigo coincide con CD_CENTRO_TRABAJO
            ' Sustituye los 3 primeros caracteres por "OFV" y mantiene el resto (desde el 4º char)
            Dim sql As String =
                "UPDATE TarjetasSanitariasDiarioInteramit AS T " &
                "INNER JOIN OficinasCaixaVirtualesDireCorreo AS O " &
                "ON Trim(T.CentroTrabajoCodigo) = Trim(O.CD_CENTRO_TRABAJO) " &
                "SET T.CentroTrabajoCodigo = IIF(Len(Trim('' & T.CentroTrabajoCodigo)) >= 4, " &
                "   'OFV' & Mid(Trim('' & T.CentroTrabajoCodigo), 4), " &
                "   'OFV') " &
                "WHERE Left(Trim('' & T.CentroTrabajoCodigo), 3) <> 'OFV';"

            Using cmd As New OleDbCommand(sql, conn)
                Dim afectados As Integer = cmd.ExecuteNonQuery()
                ' Si quieres, puedes dejar trazas:
                ' MessageBox.Show($"MarcarCentroDeTrabajoCaixaOficinVirtual: {afectados} registros actualizados.")
            End Using
        End Using

    End Sub

    Public Sub GenerarLogProcesoExcel(
    rutaSalida As String,
    nombreEntrada As String,
    listaFicherosSalida As List(Of Tuple(Of String, Integer))
)

        Try
            ' =====================================================
            ' NOMBRE FICHERO 
            ' =====================================================
            Dim nombreExcel As String = "Reporte_Tarjetas_Sanitarias" & Now.ToString("yyyyMMdd") & ".xlsx"
            Dim rutaExcel As String = Path.Combine(rutaSalida, nombreExcel)

            Dim wb As XLWorkbook

            If File.Exists(rutaExcel) Then
                wb = New XLWorkbook(rutaExcel)
            Else
                wb = New XLWorkbook()
            End If

            Dim ws As IXLWorksheet

            If wb.Worksheets.Count = 0 Then
                ws = wb.Worksheets.Add("Procesos")
            Else
                ws = wb.Worksheet(1)
            End If

            ' =====================================================
            ' 1) LEER LOG EXISTENTE
            ' =====================================================
            Dim registrosLog As New List(Of Tuple(Of DateTime, String, String, Integer))

            If ws.LastRowUsed() IsNot Nothing Then

                Dim ultimaFilaExistente As Integer = ws.LastRowUsed().RowNumber()
                Dim filaCabeceraLog As Integer = 0

                For i As Integer = 1 To ultimaFilaExistente
                    If ws.Cell(i, 1).GetString().Trim().ToUpper() = "FECHA" Then
                        filaCabeceraLog = i
                        Exit For
                    End If
                Next

                If filaCabeceraLog > 0 Then

                    Dim entradaActual As String = ""

                    For i As Integer = filaCabeceraLog + 1 To ultimaFilaExistente

                        Dim salida As String = ws.Cell(i, 3).GetString().Trim()

                        If salida = "" Then Continue For
                        If salida.ToUpper().StartsWith("TOTAL") Then Continue For
                        If salida.Contains("─") Then Continue For

                        Dim entrada As String = ws.Cell(i, 2).GetString().Trim()
                        If entrada <> "" Then entradaActual = entrada

                        If entradaActual = "" Then Continue For

                        Dim fecha As DateTime
                        DateTime.TryParse(ws.Cell(i, 1).GetFormattedString(), fecha)

                        Dim registros As Integer = 0
                        Integer.TryParse(ws.Cell(i, 4).GetFormattedString(), registros)

                        registrosLog.Add(New Tuple(Of DateTime, String, String, Integer)(
                        fecha, entradaActual, salida, registros))

                    Next
                End If
            End If

            ' =====================================================
            ' 2) AÑADIR PROCESO ACTUAL
            ' =====================================================
            Dim fechaProceso As DateTime = DateTime.Now

            For Each item In listaFicherosSalida
                registrosLog.Add(New Tuple(Of DateTime, String, String, Integer)(
                fechaProceso, nombreEntrada, item.Item1, item.Item2))
            Next

            ' =====================================================
            ' 3) LIMPIAR HOJA
            ' =====================================================
            ws.Clear()

            Dim colorCabecera = XLColor.FromHtml("#4472C4")
            Dim colorLog = XLColor.Gray
            Dim colorTotal = XLColor.FromHtml("#FFF2CC")
            Dim colorGlobal = XLColor.FromHtml("#D9E1F2")

            ' =====================================================
            ' 4) CABECERA INFORME
            ' =====================================================
            ws.Cell(1, 1).Value = "INFORME DE PROCESOS – GENERACIÓN DE FICHEROS DE TARJETAS"
            ws.Range(1, 1, 1, 4).Merge()

            ws.Cell(2, 1).Value = "Fecha de ejecución: " & Now.ToString("dd/MM/yyyy HH:mm")
            ws.Range(2, 1, 2, 4).Merge()

            With ws.Range(1, 1, 2, 4)
                .Style.Font.Bold = True
                .Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
                .Style.Fill.BackgroundColor = colorGlobal
            End With

            ws.Cell(1, 1).Style.Font.FontSize = 16

            ' =====================================================
            ' 5) RESUMEN
            ' =====================================================
            Dim filaCabeceraResumen As Integer = 4

            ws.Cell(filaCabeceraResumen, 1).Value = "Fichero Entrada"
            ws.Cell(filaCabeceraResumen, 2).Value = "Total Registros"
            ws.Cell(filaCabeceraResumen, 3).Value = "Num Ficheros"
            ws.Cell(filaCabeceraResumen, 4).Value = "Fecha"

            With ws.Range(filaCabeceraResumen, 1, filaCabeceraResumen, 4)
                .Style.Font.Bold = True
                .Style.Fill.BackgroundColor = colorCabecera
                .Style.Font.FontColor = XLColor.White
                .Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
            End With

            Dim grupos = registrosLog.
            GroupBy(Function(x) New With {
                Key .Entrada = x.Item2,
                Key .Fecha = x.Item1
            }).
            OrderBy(Function(g) g.Key.Fecha).
            ToList()

            Dim filaResumen As Integer = filaCabeceraResumen + 1
            Dim totalGlobal As Integer = 0

            For Each grupo In grupos

                Dim totalGrupo As Integer = grupo.Sum(Function(x) x.Item4)

                ws.Cell(filaResumen, 1).Value = grupo.Key.Entrada
                ws.Cell(filaResumen, 2).Value = totalGrupo
                ws.Cell(filaResumen, 3).Value = grupo.Count()
                ws.Cell(filaResumen, 4).Value = grupo.Key.Fecha
                ws.Cell(filaResumen, 4).Style.DateFormat.Format = "dd/MM/yyyy HH:mm"

                totalGlobal += totalGrupo
                filaResumen += 1

            Next

            ' TOTAL GLOBAL (UNA SOLA VEZ)
            ws.Cell(filaResumen, 1).Value = "TOTAL GLOBAL"
            ws.Cell(filaResumen, 2).Value = totalGlobal

            With ws.Range(filaResumen, 1, filaResumen, 4)
                .Style.Font.Bold = True
                .Style.Fill.BackgroundColor = colorGlobal
            End With

            ' =====================================================
            ' 6) LOG
            ' =====================================================
            Dim filaLogInicio As Integer = filaResumen + 3

            ws.Cell(filaLogInicio, 1).Value = "Fecha"
            ws.Cell(filaLogInicio, 2).Value = "Entrada"
            ws.Cell(filaLogInicio, 3).Value = "Salida"
            ws.Cell(filaLogInicio, 4).Value = "Registros"

            With ws.Range(filaLogInicio, 1, filaLogInicio, 4)
                .Style.Font.Bold = True
                .Style.Fill.BackgroundColor = colorLog
                .Style.Font.FontColor = XLColor.White
            End With

            Dim filaLog As Integer = filaLogInicio + 1

            For Each grupo In grupos

                Dim primera As Boolean = True

                For Each reg In grupo.OrderBy(Function(x) x.Item3)

                    ws.Cell(filaLog, 1).Value = reg.Item1
                    ws.Cell(filaLog, 1).Style.DateFormat.Format = "dd/MM/yyyy HH:mm:ss"

                    If primera Then
                        ws.Cell(filaLog, 2).Value = reg.Item2
                        ws.Cell(filaLog, 2).Style.Font.Bold = True
                        primera = False
                    End If

                    ws.Cell(filaLog, 3).Value = reg.Item3
                    ws.Cell(filaLog, 4).Value = reg.Item4

                    filaLog += 1
                Next

                ' TOTAL POR BLOQUE
                ws.Cell(filaLog, 3).Value = "TOTAL (" & grupo.Key.Entrada & ")"
                ws.Cell(filaLog, 4).Value = grupo.Sum(Function(x) x.Item4)

                With ws.Range(filaLog, 1, filaLog, 4)
                    .Style.Font.Bold = True
                    .Style.Fill.BackgroundColor = colorTotal
                End With

                filaLog += 1

            Next

            ' =====================================================
            ' 7) FORMATO FINAL
            ' =====================================================
            ws.Columns().AdjustToContents()
            ws.Column(2).Width = 25
            ws.Column(3).Width = 35

            ws.SheetView.FreezeRows(1)

            wb.SaveAs(rutaExcel)

            MessageBox.Show("✅ Reporte generado correctamente")

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try

    End Sub
End Module
