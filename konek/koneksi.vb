﻿Imports System.Data.Odbc
Module koneksi
    Public conn As Odbc.OdbcConnection
    Public da As Odbc.OdbcDataAdapter
    Public dr As Odbc.OdbcDataReader
    Public cmd As Odbc.OdbcCommand

    Public ds As New DataSet

    Sub koneksinya()
        conn = New Odbc.OdbcConnection("Dsn=crud")
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If
    End Sub
End Module
