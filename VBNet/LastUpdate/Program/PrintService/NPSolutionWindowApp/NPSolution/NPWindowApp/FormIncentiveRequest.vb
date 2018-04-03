Imports CrystalDecisions
Imports CrystalDecisions.CrystalReports
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic

Public Class FormIncentiveRequest
    Dim ds As DataSet
    Dim da As SqlDataAdapter
    Dim dt As DataTable
    Dim vQuery As String
    Dim vCMD As SqlCommand
    Dim vReportName As String
    Dim vReportPath As String
    Private Sub FormIncentiveRequest_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call InitializeDataBase()
        Dim rptDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument

        vQuery = "Exec dbo.USP_ICT_PaidRequestReport '" & vIncentiveDocNo & "'"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds)

        vReportName = "RP_CM_Request"
        vReportPath = "W:\External\Reports\Incentive" & "\" & vReportName & ".rpt"

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

    'Private Sub CrystalReportViewer1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CrystalReportViewer1.Load
    '    Call InitializeDataBase()

    '    Dim param1Feilds As New CrystalDecisions.Shared.ParameterFields
    '    Dim param1Feild As New CrystalDecisions.Shared.ParameterField
    '    Dim param1Range As New CrystalDecisions.Shared.ParameterDiscreteValue

    '    param1Feild.ParameterValueType = "@vUserID"
    '    param1Range.Value = vUserID
    '    param1Feild.CurrentValues.Add(param1Range)
    '    param1Feilds.Add(param1Feild)
    '    CrystalReportViewer1.ParameterFieldInfo = param1Feilds
    '    CrystalReportViewer1.ReportSource = "C:\Program Files\Sea & Hill co.,Ltd\LabelWizard\Form\LabelA51.rpt"
    'End Sub
End Class