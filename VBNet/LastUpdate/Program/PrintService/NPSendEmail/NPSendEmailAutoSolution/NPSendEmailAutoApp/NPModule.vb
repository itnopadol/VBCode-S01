Option Explicit On
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic

Module NPModule
    Public vConnectionString As String
    Public vConnection As SqlConnection

    Public da As SqlDataAdapter
    Public ds As DataSet
    Public dt As New DataTable
    Public vUserID As String
    Public vPassword As String


    Public Sub InitializeDataBase()
        vConnectionString = "Persist Security Info = False;User ID='sa';Password='[ibdkifu';Data Source = Nebula;Initial Catalog = BCNP;Connect Timeout=8000"
        'vConnectionString = "Persist Security Info = False;User ID='vbuser';Password='132';Data Source = Nebula;Initial Catalog = BCNP"
        vConnection = New SqlConnection(vConnectionString)
        vConnection.Open()
    End Sub
End Module
