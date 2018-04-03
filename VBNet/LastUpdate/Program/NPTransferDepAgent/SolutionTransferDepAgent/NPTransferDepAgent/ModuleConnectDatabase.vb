Option Explicit On
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization
Imports System.Net.NetworkInformation

Module ModuleConnectDatabase
    'Public vConnStrNEBULA As String
    'Public vConnNEBULA As SqlConnection
    'Public daNEBULA As SqlDataAdapter
    'Public dsNEBULA As DataSet
    'Public dtNEBULA As DataTable
    'Public dvNEBULA As DataView
    'Public cmdNEBULA As SqlCommand

    'Public vConnStrS02DB As String
    'Public vConnS02DB As SqlConnection
    'Public daS02DB As SqlDataAdapter
    'Public dsS02DB As DataSet
    'Public dtS02DB As DataTable
    'Public dvS02DB As DataView
    'Public cmdS02DB As SqlCommand

    'Public vNEBULAServer As String
    'Public vNEBULADatabase As String
    'Public vNEBULAUserID As String
    'Public vNEBULAPassword As String

    'Public vS02DBServer As String
    'Public vS02DBDatabase As String
    'Public vS02DBUserID As String
    'Public vS02DBPassword As String

    Public vTrnState As Integer
    Public vIsprocess As Integer


    'Public Sub vConnctDataBaseNP()
    '    vConnStrNEBULA = "Persist Security Info =False;User ID ='sa';Password ='[ibdkifu';Max Pool Size = 70000;Min Pool Size = 5;Data Source =NEBULA;Initial Catalog =BCNP"
    '    vConnNEBULA = New SqlConnection(vConnStrNEBULA)
    '    vConnNEBULA.Open()
    'End Sub


    'Public Sub vConnctDataBaseS02()
    '    vConnStrS02DB = "Persist Security Info =False;User ID ='sa';Password ='[ibdkifu';Max Pool Size = 10000;Min Pool Size = 5;Data Source =192.168.2.2;Initial Catalog =BCNP"
    '    vConnS02DB = New SqlConnection(vConnStrS02DB)
    '    vConnS02DB.Open()
    'End Sub

End Module
