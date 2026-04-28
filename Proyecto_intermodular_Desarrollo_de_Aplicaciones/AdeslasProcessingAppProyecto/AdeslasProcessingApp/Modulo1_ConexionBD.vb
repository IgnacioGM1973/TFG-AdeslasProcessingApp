Imports System.IO

Module Modulo1_ConexionBD

    'Public accessConnString As String = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & RutaPrincipal & "\bd\ADESLAS.accdb"
    Public accessConnString As String = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & Path.Combine(RutaBD, "ADESLAS.accdb")

End Module
