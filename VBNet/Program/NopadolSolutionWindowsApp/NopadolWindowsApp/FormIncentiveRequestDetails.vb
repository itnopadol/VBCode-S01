Imports CrystalDecisions
Imports CrystalDecisions.CrystalReports
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic
Public Class FormIncentiveRequestDetails
    Dim ds As DataSet
    Dim da As SqlDataAdapter
    Dim dt As DataTable
    Dim vQuery As String
    Dim vCMD As SqlCommand
    Dim vReportName As String
    Dim vReportPath As String

    Private Sub FormIncentiveRequestDetails_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call InitializeDataBase()
        Dim rptDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument

        vQuery = "Exec dbo.USP_ICT_PaidRequestReport '" & vIncentiveDocNo & "'"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds)

        vReportName = "RP_CM_RequestDetail"
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



        'Dim rpt As New RP_CM_RequestDetail
        'Dim param1Feilds As New CrystalDecisions.Shared.ParameterFields
        'Dim param1Feild As New CrystalDecisions.Shared.ParameterField
        'Dim param1Range As New CrystalDecisions.Shared.ParameterDiscreteValue
        'Dim param2Feild As New CrystalDecisions.Shared.ParameterField
        'Dim param2Range As New CrystalDecisions.Shared.ParameterDiscreteValue

        'Dim myTableLogonInfos = New CrystalDecisions.Shared.TableLogOnInfos()
        'Dim myTableLogonInfo = New CrystalDecisions.Shared.TableLogOnInfo()
        'Dim myConnectionInfo = New CrystalDecisions.Shared.ConnectionInfo()

        'With myConnectionInfo
        '    .ServerName = "Nebula"
        '    .DatabaseName = "BCNP"
        '    .UserID = vUserID
        '    .Password = vPassword
        'End With

        'myTableLogonInfo.ConnectionInfo = myConnectionInfo
        'myTableLogonInfo.TableName = "vConnect"
        'myTableLogonInfos.Add(myTableLogonInfo)

        'CrystalReportViewer1.LogOnInfo = myTableLogonInfos
        'param1Feild.ParameterFieldName = "@Docno"
        'param1Range.Value = vIncentiveDocNo
        'param1Feild.CurrentValues.Add(param1Range)
        'param1Feilds.Add(param1Feild)


        'CrystalReportViewer1.ReportSource = rpt
    End Sub
End Class