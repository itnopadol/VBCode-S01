Option Explicit On
Imports System.Data
Imports System.Data.SqlClient

Module Module1
    Public vConnectionString As String
    Public vConnection As New SqlConnection
    Public vUserLogIN As String
    Public vUserPosition As String

    Public Sub InitializeDatabase()

        If vConnection.State = 1 Then
            vConnection.Close()
        End If
        vConnectionString = "User ID = vbuser;Password = 132; Data Source=Nebula;Initial Catalog=BCNP"
        vConnection.ConnectionString = vConnectionString
        vConnection.Open()
    End Sub

End Module
