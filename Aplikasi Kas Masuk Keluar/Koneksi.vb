﻿Imports System.Data.Odbc
Module Koneksi
    Public CMD As OdbcCommand
    Public DS As New DataSet
    Public DA As OdbcDataAdapter
    Public RD As OdbcDataReader

    Public LokalData As String = "Driver={MySQL ODBC 3.51 Driver}; Database=db_praktikum_tugas; server=localhost; uid=root"
    Public CONN As OdbcConnection = New OdbcConnection(LokalData)

    Public Sub bukakoneksi()
        If CONN.State = ConnectionState.Closed Then
            Try
                CONN.Open()
            Catch ex As Exception
                MsgBox("Koneksi Gagal: " & ex.ToString)
            End Try
        End If
    End Sub

    Public Sub tutupkoneksi()
        If CONN.State = ConnectionState.Open Then
            Try
                CONN.Close()
            Catch ex As Exception
                MsgBox("Gagal menutup koneksi1: " & ex.ToString)
            End Try
        End If
    End Sub
End Module
