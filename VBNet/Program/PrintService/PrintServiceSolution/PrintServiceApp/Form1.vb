Option Explicit On
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization

Public Class Form1
    Dim strReportName As String
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        strReportName = "Report1"
        'Get the Report Location
        Dim strReportPath As String = "C:\" & strReportName & ".rpt"
        'Check file exists
        If Not IO.File.Exists(strReportPath) Then
            Throw (New Exception("Unable to locate report file:" & vbCrLf & strReportPath))
        End If

        'Assign the datasource and set the properties for Report viewer
        Dim rptDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        rptDocument.Load(strReportPath)
        rptDocument.PrintOptions.PaperSource = CrystalDecisions.Shared.PaperSource.Lower
        rptDocument.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.PaperA4
        rptDocument.PrintOptions.PrinterName = "\\nova\LS-LQ590-01"
        CrystalReportViewer1.ReportSource = rptDocument
    End Sub
End Class
