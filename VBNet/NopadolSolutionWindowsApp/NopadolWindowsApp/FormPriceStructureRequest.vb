Imports CrystalDecisions
Imports CrystalDecisions.CrystalReports
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic


Public Class FormPriceStructureRequest
    Dim ds As DataSet
    Dim da As SqlDataAdapter
    Dim dt As DataTable
    Dim vQuery As String
    Dim vCMD As SqlCommand
    Dim vReportName As String
    Dim vReportPath As String
    Private Sub FormPriceStructureRequest_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call InitializeDataBase()
        Dim rptDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument

        vQuery = "Exec dbo.USP_PS_RequestReport '" & vPriceStructureDocNo & "'"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds)

        vReportName = "ใบขออนุมัติโครงสร้างราคา"
        vReportPath = "W:\External\Reports\PriceStucture" & "\" & vReportName & ".rpt"

        If Not IO.File.Exists(vReportPath) Then
            Throw (New Exception("Unable to locate report file:" & vbCrLf & vReportPath))
        End If

        rptDocument.Load(vReportPath)
        rptDocument.SetDataSource(ds.Tables(0))
        CrystalReportViewer1.ShowRefreshButton = False
        CrystalReportViewer1.ShowCloseButton = False
        CrystalReportViewer1.ShowGroupTreeButton = False
        CrystalReportViewer1.ReportSource = rptDocument
    End Sub
End Class