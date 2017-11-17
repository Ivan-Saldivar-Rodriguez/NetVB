Imports Microsoft.VisualBasic
Imports System.Data.SqlClient

Public Class Class_Servidor

    Public Conec As New SqlConnection

    Public Function Conectar(ByVal vConnectionString As String) As Boolean
        Try
            With Conec
                .ConnectionString = vConnectionString
                .Open()
            End With
            Conectar = True
        Catch ex As Exception
            Conectar = False
        End Try

    End Function
    Public Sub Desconectar()
        Try
            Conec.Close()
            Conec = Nothing
        Catch ex As Exception
            Conec = Nothing
        End Try
    End Sub
End Class