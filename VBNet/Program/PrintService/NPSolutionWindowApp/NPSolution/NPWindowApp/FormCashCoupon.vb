Option Explicit On
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization
Public Class FormCashCoupon
    Dim ds As DataSet
    Dim da As SqlDataAdapter
    Dim dt As DataTable
    Dim vQuery As String
    Dim cmd As SqlCommand
    Dim vReadQuery As SqlDataReader

    Dim vReportName As String
    Dim vReportPath As String
    Private Sub FormCashCoupon_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Width = Screen.PrimaryScreen.Bounds.Width
        Me.Height = Screen.PrimaryScreen.Bounds.Height
        Me.Left = 0
        Me.Top = 0
        Me.WindowState = FormWindowState.Maximized

        Call InitializeDataBase()
        Call PrintReport()
    End Sub

    Public Sub PrintReport()
        Dim vDocNo As String

        vDocNo = FormSmartPoint.TBIssueNo.Text
        Dim rptDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument

        vQuery = "Exec dbo.USP_VP_WithdrawSearchSubNew '" & vDocNo & "'"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds)

        vReportName = "RP_NP_CashMemberCoupon"
        vReportPath = "C:\Program Files\BCAddOn_Extension.Net\Form" & "\" & vReportName & ".rpt"

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
End Class