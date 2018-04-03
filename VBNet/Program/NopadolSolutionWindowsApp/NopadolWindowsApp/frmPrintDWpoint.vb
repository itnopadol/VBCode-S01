Imports System.Windows.Forms
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.IO
Imports Microsoft.VisualBasic
Public Class frmPrintDWpoint
    Dim rptQry As String
    Dim pdwdono As String
    Dim vReportName As String
    Dim vReportPath As String

    Private Sub frmPrintDWpoint_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.CryRPTdw.RefreshReport()
        '-----------
        Call InitializeDataBase()
        Dim rptDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument

        rptQry = "exec dbo.USP_VP_WithDrawSearchSub '" & PdwDocNo & "'"
        da = New SqlDataAdapter(rptQry, vConnection)
        ds = New DataSet
        da.Fill(ds)

        vReportName = "rpt_drawpoint"
        ' vReportPath = "D:\NopadolSolutionWindowsApp" & "\" & vReportName & ".rpt"
        vReportPath = "V:\Reports" & "\" & vReportName & ".rpt"

        If Not IO.File.Exists(vReportPath) Then
            Throw (New Exception("Unable to locate report file:" & vbCrLf & vReportPath))
        End If

        rptDocument.Load(vReportPath)
        rptDocument.SetDataSource(ds.Tables(0))
        CryRPTdw.ShowRefreshButton = False
        CryRPTdw.ShowCloseButton = False
        CryRPTdw.ShowGroupTreeButton = False
        CryRPTdw.ReportSource = rptDocument
    End Sub

    Private Sub CryRPTdw_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CryRPTdw.Load

    End Sub
End Class