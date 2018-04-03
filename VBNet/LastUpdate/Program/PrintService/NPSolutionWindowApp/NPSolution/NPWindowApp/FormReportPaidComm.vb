Option Explicit On
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization
Imports System.Drawing.Printing.PrintDocument
Imports System.Drawing.Printing

Public Class FormReportPaidComm
    Dim ds As DataSet
    Dim da As SqlDataAdapter
    Dim dt As DataTable
    Dim vQuery As String
    Dim cmd As SqlCommand
    Dim vReadQuery As SqlDataReader

    Dim vReportName As String
    Dim vReportPath As String



    Private Sub FormReportPaidComm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Width = Screen.PrimaryScreen.Bounds.Width
        Me.Height = Screen.PrimaryScreen.Bounds.Height
        Me.Left = 0
        Me.Top = 0
        Me.WindowState = FormWindowState.Maximized

        Me.PrintDialog1.UseEXDialog = True

        Call InitializeDataBase()
        Call PrintReport()
    End Sub

    Public Sub PrintReport1()
        Dim vDocNo As String
        Dim vProfitCenter As String

        vDocNo = vMemPaidCommNo
        vProfitCenter = vMemProfitCenter

        Dim rptDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument

        vQuery = "Exec dbo.USP_COM_PaidSearch2 '" & vDocNo & "','" & vProfitCenter & "'"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds)

        vReportName = "RP_COM_PaidComm"
        vReportPath = "W:\External\Reports\Commission" & "\" & vReportName & ".rpt"

        If Not IO.File.Exists(vReportPath) Then
            Throw (New Exception("Unable to locate report file:" & vbCrLf & vReportPath))
        End If

        rptDocument.Load(vReportPath)
        rptDocument.SetDataSource(ds.Tables(0))
        Crystal101.ShowRefreshButton = False
        Crystal101.ShowCloseButton = False
        Crystal101.ShowGroupTreeButton = False
        Crystal101.ReportSource = rptDocument

    End Sub

    Public Sub PrintReport()
        Dim vDocNo As String
        Dim strReportName As String
        Dim vProfitCenter As String

        vDocNo = vMemPaidCommNo
        vProfitCenter = vMemProfitCenter

        vQuery = "Exec dbo.USP_COM_PaidSearch2 '" & vDocNo & "','" & vProfitCenter & "'"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds)

        strReportName = "RP_COM_PaidComm"
        Dim rpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim frmObj As New FormReportPaidComm
        Dim FileName As New String("W:\External\Reports\Commission" & "\" & strReportName & ".rpt")
        rpt.Load(FileName)

        Dim myTableLogonInfos = New CrystalDecisions.Shared.TableLogOnInfos()
        Dim myTableLogonInfo = New CrystalDecisions.Shared.TableLogOnInfo()
        Dim myConnectionInfo = New CrystalDecisions.Shared.ConnectionInfo()

        With myConnectionInfo
            .ServerName = "NEBULA"
            .DatabaseName = "BCNP"
            .UserID = vUserID
            .Password = vPassword
        End With

        myTableLogonInfo.ConnectionInfo = myConnectionInfo
        myTableLogonInfo.TableName = "vConnect"
        myTableLogonInfos.Add(myTableLogonInfo)

        Crystal101.LogOnInfo = myTableLogonInfos

        Dim Params As New CrystalDecisions.Shared.ParameterField
        Dim ParamCollection As New CrystalDecisions.Shared.ParameterFields
        Dim ParamDisVal As New CrystalDecisions.Shared.ParameterDiscreteValue()

        Dim Params1 As New CrystalDecisions.Shared.ParameterField
        Dim ParamDisVal1 As New CrystalDecisions.Shared.ParameterDiscreteValue()

        Params.ParameterFieldName = "@Profitcenter"
        ParamDisVal.Value = vProfitCenter
        Params.CurrentValues.Add(ParamDisVal1)

        Params1.ParameterFieldName = "@Docno"
        ParamDisVal1.Value = vDocNo
        Params1.CurrentValues.Add(ParamDisVal1)

        ParamCollection.Add(Params)

        rpt.Load(FileName)

        rpt.SetDataSource(ds.Tables("Search"))
        rpt.SetParameterValue("@Profitcenter", ParamDisVal)
        rpt.SetParameterValue("@Docno", ParamDisVal1)

        Crystal101.ReportSource = rpt

        Dim printDlg As New PrintDialog()
        Dim printDoc As New PrintDocument()

        printDoc.DocumentName = rpt.Name.ToString
        printDlg.Document = printDoc
        printDlg.AllowSelection = True
        printDlg.AllowSomePages = True
        'If (printDlg.ShowDialog() = DialogResult.OK) Then
        '    printDoc.Print()
        'End If

    End Sub

End Class