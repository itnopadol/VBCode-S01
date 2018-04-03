Imports CrystalDecisions
Imports CrystalDecisions.CrystalReports
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic
Public Class frmPrintVolumeSet

    Dim ds As DataSet
    Dim da As SqlDataAdapter
    Dim dt As DataTable
    Dim vQuery As String
    Dim vCMD As SqlCommand
    Dim vReportName As String
    Dim vReportPath As String
    Private Sub frmPrintVolumeSet_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.crtVW01.RefreshReport()
        '-----------
        Call InitializeDataBase()
        Dim rptDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument

        vQuery = "exec dbo.USP_PS_PriceVolumeSetSearchsub '" & publicFdocno & "'"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds)

        vReportName = "rpt_volumeset"
        ' vReportPath = "D:\NopadolSolutionWindowsApp" & "\" & vReportName & ".rpt"
        vReportPath = "W:\External\Reports\PriceStucture" & "\" & vReportName & ".rpt"

        If Not IO.File.Exists(vReportPath) Then
            Throw (New Exception("Unable to locate report file:" & vbCrLf & vReportPath))
        End If

        rptDocument.Load(vReportPath)
        rptDocument.SetDataSource(ds.Tables(0))
        crtVW01.ShowRefreshButton = False
        crtVW01.ShowCloseButton = False
        crtVW01.ShowGroupTreeButton = False
        crtVW01.ReportSource = rptDocument
    End Sub
    

    Private Sub crtVW01_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles crtVW01.Load

    End Sub
End Class