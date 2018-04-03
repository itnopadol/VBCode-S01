Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Collections
Imports System.Windows.Forms
Imports Microsoft.VisualBasic
Imports System.Net

'http://192.168.2.2/WebServiceCalc.asmx
'http://192.168.0.203/WebServiceCalc.asmx
'http://ws1.nopadol.com/WebServiceCalc.asmx


Module ModuleAddOn
    Public vConnectionString As String
    Public vConnection As SqlConnection
    Public vCon As SqlConnection
    Public vConnectionStringBPlus As String
    Public vConnectionBPlus As SqlConnection

    Public da As SqlDataAdapter
    Public pds As DataSet
    Public pds1 As DataSet
    Public pds2 As DataSet
    Public pds3 As DataSet
    Public pds4 As DataSet
    Public pds5 As DataSet
    Public pds6 As DataSet
    Public pds7 As DataSet
    Public pds8 As DataSet
    Public pds9 As DataSet

    Public eds As DataSet
    Public vdt As DataTable
    Public dt As New DataTable

    Public vUserID As String
    Public vPassword As String
    Public vMemProfit As String
    Public vUserName As String

    Public vMemReOrderIsOpen As Integer
    Public vMemInspectIsOpen As Integer
    Public vMemReqProIsOpen As Integer

    'Pickup================================================
    Public vCheckLogIn As String
    Public vConnectZone As String
    Public vSelectLineEdit As Integer
    Public vSelectCheckOutLine As Integer
    Public vIsOpen As Integer
    Public vIsCancel As Integer
    Public vIsconfirm As Integer
    Public vIsSendQue As Integer
    'Pickup================================================



    Public Sub InitializeDataNP()
        vConnectionString = "Persist Security Info = False;User ID='" & vUserID & "';Password='" & vPassword & "';Data Source = Nebula;Initial Catalog = BCNP"
        vConnection = New SqlConnection(vConnectionString)
        vConnection.Open()

    End Sub

    Public Sub InitializeDataS02()
        vConnectionString = "Persist Security Info = False;User ID='" & vUserID & "';Password='" & vPassword & "';Data Source = S02DB;Initial Catalog = BCNP"
        vConnection = New SqlConnection(vConnectionString)
        vConnection.Open()

    End Sub

    Public Sub vGetData(ByVal vProfit As String, ByVal vGetQuery As String)
        If vProfit = "S01" Then
            Dim vServiceS01 As New WebReferenceS01.WebServiceCalc
            pds = vServiceS01.vGetQueryAnlyzer(vGetQuery)
        End If

        If vProfit = "S02" Then
            Dim vServiceS02 As New WebReferenceS02.WebServiceCalc
            pds = vServiceS02.vGetQueryAnlyzer(vGetQuery)
        End If
    End Sub

    Public Sub vGetData1(ByVal vProfit As String, ByVal vGetQuery As String)
        If vProfit = "S01" Then
            Dim vServiceS01 As New WebReferenceS01.WebServiceCalc
            pds1 = vServiceS01.vGetQueryAnlyzer(vGetQuery)
        End If

        If vProfit = "S02" Then
            Dim vServiceS02 As New WebReferenceS02.WebServiceCalc
            pds1 = vServiceS02.vGetQueryAnlyzer(vGetQuery)
        End If
    End Sub

    Public Sub vGetData2(ByVal vProfit As String, ByVal vGetQuery As String)
        If vProfit = "S01" Then
            Dim vServiceS01 As New WebReferenceS01.WebServiceCalc
            pds2 = vServiceS01.vGetQueryAnlyzer(vGetQuery)
        End If

        If vProfit = "S02" Then
            Dim vServiceS02 As New WebReferenceS02.WebServiceCalc
            pds2 = vServiceS02.vGetQueryAnlyzer(vGetQuery)
        End If
    End Sub

    Public Sub vGetData3(ByVal vProfit As String, ByVal vGetQuery As String)
        If vProfit = "S01" Then
            Dim vServiceS01 As New WebReferenceS01.WebServiceCalc
            pds3 = vServiceS01.vGetQueryAnlyzer(vGetQuery)
        End If

        If vProfit = "S02" Then
            Dim vServiceS02 As New WebReferenceS02.WebServiceCalc
            pds3 = vServiceS02.vGetQueryAnlyzer(vGetQuery)
        End If
    End Sub

    Public Sub vGetData4(ByVal vProfit As String, ByVal vGetQuery As String)
        If vProfit = "S01" Then
            Dim vServiceS01 As New WebReferenceS01.WebServiceCalc
            pds4 = vServiceS01.vGetQueryAnlyzer(vGetQuery)
        End If

        If vProfit = "S02" Then
            Dim vServiceS02 As New WebReferenceS02.WebServiceCalc
            pds4 = vServiceS02.vGetQueryAnlyzer(vGetQuery)
        End If
    End Sub

    Public Sub vGetData5(ByVal vProfit As String, ByVal vGetQuery As String)
        If vProfit = "S01" Then
            Dim vServiceS01 As New WebReferenceS01.WebServiceCalc
            pds5 = vServiceS01.vGetQueryAnlyzer(vGetQuery)
        End If

        If vProfit = "S02" Then
            Dim vServiceS02 As New WebReferenceS02.WebServiceCalc
            pds5 = vServiceS02.vGetQueryAnlyzer(vGetQuery)
        End If
    End Sub

    Public Sub vGetData6(ByVal vProfit As String, ByVal vGetQuery As String)
        If vProfit = "S01" Then
            Dim vServiceS01 As New WebReferenceS01.WebServiceCalc
            pds6 = vServiceS01.vGetQueryAnlyzer(vGetQuery)
        End If

        If vProfit = "S02" Then
            Dim vServiceS02 As New WebReferenceS02.WebServiceCalc
            pds6 = vServiceS02.vGetQueryAnlyzer(vGetQuery)
        End If
    End Sub

    Public Sub vGetData7(ByVal vProfit As String, ByVal vGetQuery As String)
        If vProfit = "S01" Then
            Dim vServiceS01 As New WebReferenceS01.WebServiceCalc
            pds7 = vServiceS01.vGetQueryAnlyzer(vGetQuery)
        End If

        If vProfit = "S02" Then
            Dim vServiceS02 As New WebReferenceS02.WebServiceCalc
            pds7 = vServiceS02.vGetQueryAnlyzer(vGetQuery)
        End If
    End Sub

    Public Sub vGetData8(ByVal vProfit As String, ByVal vGetQuery As String)
        If vProfit = "S01" Then
            Dim vServiceS01 As New WebReferenceS01.WebServiceCalc
            pds8 = vServiceS01.vGetQueryAnlyzer(vGetQuery)
        End If

        If vProfit = "S02" Then
            Dim vServiceS02 As New WebReferenceS02.WebServiceCalc
            pds8 = vServiceS02.vGetQueryAnlyzer(vGetQuery)
        End If
    End Sub

    Public Sub vGetData9(ByVal vProfit As String, ByVal vGetQuery As String)
        If vProfit = "S01" Then
            Dim vServiceS01 As New WebReferenceS01.WebServiceCalc
            pds9 = vServiceS01.vGetQueryAnlyzer(vGetQuery)
        End If

        If vProfit = "S02" Then
            Dim vServiceS02 As New WebReferenceS02.WebServiceCalc
            pds9 = vServiceS02.vGetQueryAnlyzer(vGetQuery)
        End If
    End Sub

    Public Sub vExecData(ByVal vProfit As String, ByVal vGetQuery As String)
        If vProfit = "S01" Then
            Dim vServiceS01 As New WebReferenceS01.WebServiceCalc
            pds = vServiceS01.vGetQueryAnlyzer(vGetQuery)
        End If

        If vProfit = "S02" Then
            Dim vServiceS02 As New WebReferenceS02.WebServiceCalc
            pds = vServiceS02.vGetQueryAnlyzer(vGetQuery)
        End If
    End Sub
End Module
