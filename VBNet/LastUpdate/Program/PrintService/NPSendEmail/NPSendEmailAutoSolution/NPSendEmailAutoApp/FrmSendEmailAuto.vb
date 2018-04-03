Option Explicit On
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization

Imports CrystalDecisions.ReportSource
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine

Imports System.Net.Mail
Imports System.Net

Imports Ionic.Zip

'Imports System.Text
'Imports System.Text.RegularExpressions


Public Class FrmSendEmailAuto
    Dim ds As DataSet
    Dim da As SqlDataAdapter
    Dim dt As DataTable

    Dim ds1 As DataSet
    Dim da1 As SqlDataAdapter
    Dim dt1 As DataTable

    Dim vQuery As String
    Dim cmd As SqlCommand
    Dim vReadQuery As SqlDataReader

    Dim vReportName As String
    Dim vReportPath As String

    Dim vCountPO As Double

    Private Sub FrmSendEmailAuto_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

    End Sub

    Private Sub FrmSendEmailAuto_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim i As Integer
        Dim vGetReportID As String
        Dim vGetReportType As Integer
        Dim vGetProfitCenter As String
        Dim vGetExpertTeam As String
        Dim vGetSection As String
        Dim vGetDepartCode As String
        Dim vGetApCode As String
        Dim vGetDocNo As String
        Dim vGetEmail As String
        Dim vGetCC As String
        Dim vGetFromMail As String
        Dim vGetReportName As String
        Dim vGetPrintDateTime As String
        Dim vGetEmailAll As String
        Dim vGetContactName As String
        Dim vGetNopadolContactName As String

        Dim vListItem As ListViewItem
        Dim vFileName1 As String

        On Error Resume Next

        'Dim vMonth As String
        'Dim vLenMonth As Integer
        'Dim vDay As String
        'Dim vLenDay As Integer

        'vLenDay = vb6.Len(RTrim(DateAdd(DateInterval.Day, -1, Now).Day))
        'If vLenDay < 2 Then
        '    vDay = "0" & RTrim(DateAdd(DateInterval.Day, -1, Now).Day)
        'Else
        '    vDay = RTrim(DateAdd(DateInterval.Day, -1, Now).Day)
        'End If

        'vLenMonth = vb6.Len(RTrim(DateAdd(DateInterval.Day, -1, Now).Month))
        'If vLenMonth < 2 Then
        '    vMonth = "0" & RTrim(DateAdd(DateInterval.Day, -1, Now).Month)
        'Else
        '    vMonth = RTrim(DateAdd(DateInterval.Day, -1, Now).Month)
        'End If


        'vFileName1 = "\\dev\BCS\crm\3001278_MonthlyM-" & vday & "-" & vMonth & "-" & now.year & ".txt"

        'MsgBox(vFileName1)

        'On Error GoTo ErrDescription

        Call InitializeDataBase()

        vQuery = "exec dbo.USP_PR_SearchCheckPrintReOrderAuto_Cat 0"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "Docno")
        dt = ds.Tables("Docno")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                vGetReportID = dt.Rows(0).Item("reportid")
                If vGetReportID = "001" Then
                    vGetReportType = 2
                ElseIf vGetReportID = "011" Then
                    vGetReportType = 2
                ElseIf vGetReportID = "002" Then
                    vGetReportType = 1
                ElseIf vGetReportID = "012" Then
                    vGetReportType = 1
                ElseIf vGetReportID = "013" Then
                    vGetReportType = 1
                ElseIf vGetReportID = "014" Then
                    vGetReportType = 1
                ElseIf vGetReportID = "004" Then
                    vGetReportType = 1
                ElseIf vGetReportID = "005" Then 'CAT =< 100000
                    vGetReportType = 0
                ElseIf vGetReportID = "006" Then 'ALL=ทั้หมดและMG > 100000
                    vGetReportType = 1
                ElseIf vGetReportID = "999" Then 'ส่งเมลล์ปรับราคาให้น้ำหวานศูนย์สีไอซีไอ
                    vGetReportType = 1
                End If
                vGetProfitCenter = dt.Rows(0).Item("profitcenter")
                vGetExpertTeam = dt.Rows(0).Item("expertteam")
                vGetSection = dt.Rows(0).Item("sectionmanager")
                vGetDepartCode = dt.Rows(0).Item("department")
                vGetApCode = dt.Rows(0).Item("apcode")
                vGetDocNo = dt.Rows(0).Item("docno")
                vGetContactName = dt.Rows(0).Item("contactname")
                vGetNopadolContactName = dt.Rows(0).Item("nopadolcontact")

                vGetEmail = dt.Rows(0).Item("email")
                vGetCC = dt.Rows(0).Item("cc")
                vGetFromMail = dt.Rows(0).Item("fromemail")
                vGetEmailAll = vGetEmail & "," & vGetCC

                If vGetReportID = "001" Or vGetReportID = "002" Or vGetReportID = "004" Or vGetReportID = "011" Or vGetReportID = "012" Then

                    vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vGetReportID & "','" & vGetExpertTeam & "','" & vGetProfitCenter & "','" & vGetDepartCode & "','" & vGetSection & "','" & vGetApCode & "','" & vGetDocNo & "','" & vGetEmailAll & "'"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Call SendMailReport_Gmail(vGetReportID, vGetReportType, vGetProfitCenter, vGetExpertTeam, vGetSection, vGetDepartCode, vGetApCode, vGetDocNo, vGetEmail, vGetCC, vGetFromMail)

                ElseIf vGetReportID = "013" Or vGetReportID = "014" Then

                    vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vGetReportID & "','" & vGetExpertTeam & "','" & vGetProfitCenter & "','" & vGetDepartCode & "','" & vGetSection & "','" & vGetApCode & "','" & vGetDocNo & "','" & vGetEmailAll & "'"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Call SendMailItemTransfer_Gmail(vGetReportID, vGetReportType, vGetProfitCenter, vGetExpertTeam, vGetSection, vGetDepartCode, vGetApCode, vGetDocNo, vGetEmail, vGetCC, vGetFromMail)

                ElseIf vGetReportID = "003" Then

                    'vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vGetReportID & "','','','','','" & vGetApCode & "','" & vGetDocNo & "','" & vGetEmailAll & "'"
                    'cmd = New SqlCommand(vQuery, vConnection)
                    'cmd.ExecuteNonQuery()

                    Call SendMailPO_Gmail(vGetApCode, vGetDocNo, vGetEmail, vGetCC, vGetFromMail)

                ElseIf vGetReportID = "005" Then

                    vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vGetReportID & "','','','','','','','" & vGetEmailAll & "'"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Call SendMailPOApproveCAT_Gmail(vGetReportID, "CAT1", vGetEmail, vGetCC, vGetFromMail)

                ElseIf vGetReportID = "006" Then

                    vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vGetReportID & "','','','','','','','" & vGetEmailAll & "'"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Call SendMailPOApproveCAT_Gmail(vGetReportID, "CAT2", vGetEmail, vGetCC, vGetFromMail)

                ElseIf vGetReportID = "007" Then

                    vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vGetReportID & "','','','','','','','" & vGetEmailAll & "'"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Call SendMailPOApproveCAT_Gmail(vGetReportID, "CAT3", vGetEmail, vGetCC, vGetFromMail)

                ElseIf vGetReportID = "008" Then

                    vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vGetReportID & "','','','','','','','" & vGetEmailAll & "'"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Call SendMailPOApproveMG_Gmail(vGetEmail, vGetCC, vGetFromMail)

                ElseIf vGetReportID = "999" Then

                    vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vGetReportID & "','','','','','','','" & vGetEmailAll & "'"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Call SendMailChangePrice_Gmail(vGetReportID, vGetEmail, vGetCC, vGetFromMail)

                ElseIf vGetReportID = "SCG" Then
                    vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vGetReportID & "','','','','','','','" & vGetEmailAll & "'"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Call SendMailSCG_Gmail(vGetEmail, vGetCC, vGetFromMail)

                ElseIf vGetReportID = "SRF" Then

                    If vGetEmail = "" Then
                        vGetEmailAll = "ไม่ได้ระบุรหัสเจ้าหนี้"
                    End If

                    vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vGetReportID & "','','','','','" & vGetApCode & "','" & vGetDocNo & "','" & vGetEmailAll & "'"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Call SendMailReturn_Gmail(vGetApCode, vGetDocNo, vGetEmail, vGetCC, vGetFromMail, vGetContactName, vGetNopadolContactName)

                End If

                If vGetReportID <> "" Then
                    vQuery = "exec dbo.USP_PR_UpdatePrintReOrderComplete '" & vGetReportID & "','" & vGetExpertTeam & "','" & vGetProfitCenter & "','" & vGetDocNo & "'"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()
                End If
            Next
        Else
            vGetReportID = ""
            vGetReportType = 0
            vGetProfitCenter = ""
            vGetExpertTeam = ""
            vGetSection = ""
            vGetDepartCode = ""
            vGetApCode = ""
            vGetDocNo = ""

            vGetEmail = ""
            vGetCC = ""
        End If


        'ErrDescription:

        '        If Err.Description <> "" Then
        '            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        '            Exit Sub
        '        End If
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim i As Integer
        Dim vGetReportID As String
        Dim vGetReportType As Integer
        Dim vGetProfitCenter As String
        Dim vGetExpertTeam As String
        Dim vGetSection As String
        Dim vGetDepartCode As String
        Dim vGetApCode As String
        Dim vGetDocNo As String
        Dim vGetEmail As String
        Dim vGetCC As String
        Dim vGetFromMail As String
        Dim vGetEmailAll As String
        Dim vGetReportName As String
        Dim vGetPrintDateTime As String
        Dim vGetContactName As String
        Dim vGetNopadolContactName As String

        Dim vListItem As ListViewItem

        On Error Resume Next

        vQuery = "exec dbo.USP_PR_SearchCheckPrintReOrderAuto_Cat 0"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "Docno")
        dt = ds.Tables("Docno")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                vGetReportID = dt.Rows(0).Item("reportid")
                If vGetReportID = "001" Then
                    vGetReportType = 2
                ElseIf vGetReportID = "011" Then
                    vGetReportType = 2
                ElseIf vGetReportID = "002" Then
                    vGetReportType = 1
                ElseIf vGetReportID = "012" Then
                    vGetReportType = 1
                ElseIf vGetReportID = "013" Then
                    vGetReportType = 1
                ElseIf vGetReportID = "014" Then
                    vGetReportType = 1
                ElseIf vGetReportID = "004" Then
                    vGetReportType = 1
                ElseIf vGetReportID = "005" Then 'CAT =< 100000
                    vGetReportType = 0
                ElseIf vGetReportID = "006" Then 'ALL=ทั้หมดและMG > 100000
                    vGetReportType = 1
                ElseIf vGetReportID = "999" Then 'ส่งเมลล์ปรับราคาให้น้ำหวานศูนย์สีไอซีไอ
                    vGetReportType = 1
                End If
                vGetProfitCenter = dt.Rows(0).Item("profitcenter")
                vGetExpertTeam = dt.Rows(0).Item("expertteam")
                vGetSection = dt.Rows(0).Item("sectionmanager")
                vGetDepartCode = dt.Rows(0).Item("department")
                vGetApCode = dt.Rows(0).Item("apcode")
                vGetDocNo = dt.Rows(0).Item("docno")
                vGetContactName = dt.Rows(0).Item("contactname")
                vGetNopadolContactName = dt.Rows(0).Item("nopadolcontact")

                vGetEmail = dt.Rows(0).Item("email")
                vGetCC = dt.Rows(0).Item("cc")
                vGetFromMail = dt.Rows(0).Item("fromemail")
                vGetEmailAll = vGetEmail & "," & vGetCC

                If vGetReportID = "001" Or vGetReportID = "002" Or vGetReportID = "004" Or vGetReportID = "011" Or vGetReportID = "012" Then
                    'Call SendMailReport_Gmail(vGetReportID, vGetReportType, vGetProfitCenter, vGetExpertTeam, vGetSection, vGetDepartCode, vGetApCode, vGetDocNo, vGetEmail, vGetCC, vGetFromMail)

                    vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vGetReportID & "','" & vGetExpertTeam & "','" & vGetProfitCenter & "','" & vGetDepartCode & "','" & vGetSection & "','" & vGetApCode & "','" & vGetDocNo & "','" & vGetEmailAll & "'"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Call SendMailReport_Gmail(vGetReportID, vGetReportType, vGetProfitCenter, vGetExpertTeam, vGetSection, vGetDepartCode, vGetApCode, vGetDocNo, vGetEmail, vGetCC, vGetFromMail)

                ElseIf vGetReportID = "013" Or vGetReportID = "014" Then
                    'Call SendMailItemTransfer_Gmail(vGetReportID, vGetReportType, vGetProfitCenter, vGetExpertTeam, vGetSection, vGetDepartCode, vGetApCode, vGetDocNo, vGetEmail, vGetCC, vGetFromMail)

                    vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vGetReportID & "','" & vGetExpertTeam & "','" & vGetProfitCenter & "','" & vGetDepartCode & "','" & vGetSection & "','" & vGetApCode & "','" & vGetDocNo & "','" & vGetEmailAll & "'"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Call SendMailItemTransfer_Gmail(vGetReportID, vGetReportType, vGetProfitCenter, vGetExpertTeam, vGetSection, vGetDepartCode, vGetApCode, vGetDocNo, vGetEmail, vGetCC, vGetFromMail)

                ElseIf vGetReportID = "003" Then
                    Call SendMailPO_Gmail(vGetApCode, vGetDocNo, vGetEmail, vGetCC, vGetFromMail)

                    'vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vGetReportID & "','','','','','" & vGetApCode & "','" & vGetDocNo & "','" & vGetEmailAll & "'"
                    'cmd = New SqlCommand(vQuery, vConnection)
                    'cmd.ExecuteNonQuery()

                    'Call SendMailPO_Gmail(vGetApCode, vGetDocNo, vGetEmail, vGetCC, vGetFromMail)

                ElseIf vGetReportID = "005" Then

                    vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vGetReportID & "','','','','','','','" & vGetEmailAll & "'"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Call SendMailPOApproveCAT_Gmail(vGetReportID, "CAT1", vGetEmail, vGetCC, vGetFromMail)

                ElseIf vGetReportID = "006" Then

                    vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vGetReportID & "','','','','','','','" & vGetEmailAll & "'"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Call SendMailPOApproveCAT_Gmail(vGetReportID, "CAT2", vGetEmail, vGetCC, vGetFromMail)

                ElseIf vGetReportID = "007" Then

                    vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vGetReportID & "','','','','','','','" & vGetEmailAll & "'"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Call SendMailPOApproveCAT_Gmail(vGetReportID, "CAT3", vGetEmail, vGetCC, vGetFromMail)

                ElseIf vGetReportID = "008" Then

                    vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vGetReportID & "','','','','','','','" & vGetEmailAll & "'"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Call SendMailPOApproveMG_Gmail(vGetEmail, vGetCC, vGetFromMail)

                ElseIf vGetReportID = "999" Then

                    vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vGetReportID & "','','','','','','','" & vGetEmailAll & "'"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Call SendMailChangePrice_Gmail(vGetReportID, vGetEmail, vGetCC, vGetFromMail)

                ElseIf vGetReportID = "SCG" Then
                    vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vGetReportID & "','','','','','','','" & vGetEmailAll & "'"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Call SendMailSCG_Gmail(vGetEmail, vGetCC, vGetFromMail)
                ElseIf vGetReportID = "SRF" Then

                    If vGetEmail = "" Then
                        vGetEmailAll = "ไม่ได้ระบุรหัสเจ้าหนี้"
                    End If

                    vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vGetReportID & "','','','','','" & vGetApCode & "','" & vGetDocNo & "','" & vGetEmailAll & "'"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Call SendMailReturn_Gmail(vGetApCode, vGetDocNo, vGetEmail, vGetCC, vGetFromMail, vGetContactName, vGetNopadolContactName)

                End If

                If vGetReportID <> "" Then
                    vQuery = "exec dbo.USP_PR_UpdatePrintReOrderComplete '" & vGetReportID & "','" & vGetExpertTeam & "','" & vGetProfitCenter & "','" & vGetDocNo & "'"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()
                End If
            Next
        Else
            vGetReportID = ""
            vGetReportType = 0
            vGetProfitCenter = ""
            vGetExpertTeam = ""
            vGetSection = ""
            vGetDepartCode = ""
            vGetApCode = ""
            vGetDocNo = ""

            vGetEmail = ""
            vGetCC = ""
        End If

        Me.ListViewSendMail.Items.Clear()
        vQuery = "exec dbo.USP_PR_SearchCheckPrintReOrderAuto_Cat 1"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "Docno")
        dt = ds.Tables("Docno")
        If dt.Rows.Count > 0 Then

            vCountPO = dt.Rows(i).Item("vCount")

            For i = 0 To dt.Rows.Count - 1
                vGetReportID = dt.Rows(i).Item("reportid")
                If vGetReportID = "001" Then
                    vGetReportType = 0
                ElseIf vGetReportID = "002" Then
                    vGetReportType = 1
                ElseIf vGetReportID = "004" Then
                    vGetReportType = 1
                ElseIf vGetReportID = "003" Then
                    vGetReportType = 1
                Else
                    vReportName = ""
                End If
                vGetProfitCenter = dt.Rows(i).Item("profitcenter")
                vGetExpertTeam = dt.Rows(i).Item("expertteam")
                vGetSection = dt.Rows(i).Item("sectionmanager")
                vGetDepartCode = dt.Rows(i).Item("department")
                vGetApCode = dt.Rows(i).Item("apcode")
                vGetDocNo = dt.Rows(i).Item("docno")
                vGetReportName = dt.Rows(i).Item("reportname")

                vGetEmail = dt.Rows(i).Item("email")
                vGetCC = dt.Rows(i).Item("cc")
                vGetPrintDateTime = dt.Rows(i).Item("printdatetime")

                If vGetReportID <> "" Then
                    vListItem = Me.ListViewSendMail.Items.Add(vGetPrintDateTime)
                    vListItem.SubItems.Add(0).Text = vGetReportName
                    vListItem.SubItems.Add(1).Text = vGetProfitCenter
                    vListItem.SubItems.Add(2).Text = vGetExpertTeam
                    vListItem.SubItems.Add(3).Text = vGetDepartCode
                    vListItem.SubItems.Add(4).Text = vGetSection
                    vListItem.SubItems.Add(5).Text = vGetApCode
                    vListItem.SubItems.Add(6).Text = vGetDocNo
                    vListItem.SubItems.Add(7).Text = vGetEmail
                    vListItem.SubItems.Add(8).Text = vGetCC
                End If

            Next
        End If


        'ErrDescription:

        '        If Err.Description <> "" Then
        '            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        '            Exit Sub
        '        End If
    End Sub

    Public Sub SendEmail(ByVal vReportID As String, ByVal vReportType As Integer, ByVal vProfitCenter As String, ByVal vExpertTeam As String, ByVal vSection As String, ByVal vDepartCode As String, ByVal vApCode As String, ByVal vDocNo As String, ByVal vEmail As String, ByVal vCC As String, ByVal vFromMail As String)
        Dim vPDFName As String
        Dim vFileName As String
        Dim vSubject As String
        Dim vBody As String
        Dim vEmailAll As String
        Dim da1 As SqlDataAdapter
        Dim ds1 As DataSet
        Dim dt1 As New DataTable

        Dim i As Integer
        Dim vGetDocNo As String


        On Error Resume Next

        Dim myTableLogonInfo As CrystalDecisions.Shared.TableLogOnInfo
        Dim myConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo
        Dim lFileName As String

        Dim mail1 As New MailMessage()
        Dim smtp As New SmtpClient("192.168.0.169")

        'Me.Timer1.Enabled = False

        vEmailAll = vEmail & "," & vCC

        If vReportID = "001" Or vReportID = "002" Or vReportID = "004" Then
            If vReportID = "001" Then
                vReportName = "RP_PR_StockRequestReOrderDaily"
                vReportPath = Trim("W:\External\Reports\ReOrder\" & vReportName & ".rpt")

                vPDFName = Trim(vReportID & "-" & vExpertTeam & "-" & vb6.Year(Now) & "_" & vb6.Month(Now) & "_" & vb6.Day(Now))
                vFileName = Trim("C:\PDFDocApp\" & vPDFName & ".pdf")
                vSubject = vExpertTeam & "-" & "สินค้าที่เสนอซื้อประจำวันตามทีมขาย (ReOrder)" & vExpertTeam
                vBody = "เป็นรายงาน เสนอซื้อสินค้าของทีม" & vExpertTeam

            ElseIf vReportID = "002" Then
                vReportName = "RP_PR_StockRequestReOrderApproveDaily"
                vReportPath = Trim("W:\External\Reports\ReOrder\" & vReportName & ".rpt")

                vPDFName = Trim(vReportID & "-" & vExpertTeam & "-" & vb6.Year(Now) & "_" & vb6.Month(Now) & "_" & vb6.Day(Now))
                vFileName = Trim("C:\PDFDocApp\" & vPDFName & ".pdf")
                vSubject = vExpertTeam & "-" & "สินค้าที่เสนอซื้อที่อนุมัติแล้ว ประจำวันตามทีมขาย (ReOrder)" & vExpertTeam
                vBody = "เป็นรายงาน เสนอซื้อสินค้าที่อนุมัติจำนวนแล้ว ของทีม" & vExpertTeam

            ElseIf vReportID = "004" Then
                vReportName = "RP_PR_TeamExpertReOrderConfirmQty"
                vReportPath = Trim("W:\External\Reports\ReOrder\" & vReportName & ".rpt")

                vPDFName = Trim(vReportID & "-" & vExpertTeam & "-" & vb6.Year(Now) & "_" & vb6.Month(Now) & "_" & vb6.Day(Now))
                vFileName = Trim("C:\PDFDocApp\" & vPDFName & ".pdf")
                vSubject = vExpertTeam & "-" & "สินค้าที่เสนอซื้อที่ทางจัดซื้อได้พิจารณาแล้ว ประจำวันตามทีมขาย (ReOrder)เพื่อให้ทางผู้อำนวยการจัดซื้อ" & vExpertTeam
                vBody = "เป็นรายงาน สินค้าที่เสนอซื้อที่ทางจัดซื้อได้พิจารณาแล้ว ของทีม" & vExpertTeam & "เพื่อให้ทางผู้อำนวยการจัดซื้อพิจารณาตรวจสอบ"
            End If
        End If

        Dim myReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        myReport.Load(vReportPath)

        myConnectionInfo.ServerName = "Nebula"
        myConnectionInfo.DatabaseName = "BCNP"
        myConnectionInfo.UserID = "sa"
        myConnectionInfo.Password = "[ibdkifu"

        Dim myTables As CrystalDecisions.CrystalReports.Engine.Tables
        myTables = myReport.Database.Tables

        For Each myTable As CrystalDecisions.CrystalReports.Engine.Table In myTables
            myTableLogonInfo = myTable.LogOnInfo
            myTableLogonInfo.ConnectionInfo = myConnectionInfo
            myTable.ApplyLogOnInfo(myTableLogonInfo)
        Next

        lFileName = vFileName
        myReport.SetParameterValue("@vProfit", vProfitCenter)
        myReport.SetParameterValue("@vType", vReportType)
        myReport.SetParameterValue("@vExpertTeam", vExpertTeam)
        myReport.SetParameterValue("@vSectionID", vSection)
        myReport.SetParameterValue("@vDepartment", vDepartCode)
        myReport.SetParameterValue("@vDocNo", vDocNo)

        CrystalReportViewer1.ReportSource = myReport

        myReport.ExportToDisk(ExportFormatType.PortableDocFormat, lFileName)

        lFileName = vFileName
        Dim mail As New MailMessage()
        mail.From = New MailAddress(vFromMail)

        Dim strEmails() As String = vEmailAll.Split(",")
        For Each str As String In strEmails
            If str <> "" Then
                mail.To.Add(str)
            End If
        Next

        mail.Subject = vSubject
        mail.Body = vBody

        Dim att As New Attachment(lFileName)
        Dim att1 As New Attachment(lFileName)

        mail.Attachments.Add(att1)
        mail.Attachments.Add(att)

        smtp.Send(mail)

        vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vReportID & "','" & vExpertTeam & "','" & vProfitCenter & "','" & vDepartCode & "','" & vSection & "','" & vApCode & "','" & vDocNo & "','" & vEmailAll & "'"
        cmd = New SqlCommand(vQuery, vConnection)
        cmd.ExecuteNonQuery()

        Me.Timer1.Enabled = True

        'ErrDescription:

        '        If Err.Description <> "" Then
        '            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        '            Exit Sub
        '        End If
    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        NotifyIcon1.Visible = False
        Me.Visible = True
        Me.WindowState = FormWindowState.Normal
    End Sub

    Private Sub FrmSendEmailAuto_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged
        If (Me.WindowState = FormWindowState.Minimized) Then
            Me.Visible = False
            NotifyIcon1.Visible = True
        End If
    End Sub

    Public Sub SendMailPO(ByVal vApCode As String, ByVal vDocNo As String, ByVal vEmail As String, ByVal vCC As String, ByVal vFromMail As String)
        Dim myReport3 As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim i As Integer
        Dim vGetDocNo As String
        Dim vGetApName As String
        Dim vGetContactName As String
        Dim vPDFName As String
        Dim vFileName As String
        Dim vSubject As String
        Dim vBody As String
        Dim rptDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim vEmailAll As String
        Dim vReportID As String

        'On Error GoTo ErrDescription
        On Error Resume Next

        Dim myTables As CrystalDecisions.CrystalReports.Engine.Tables

        Dim myTableLogonInfo As CrystalDecisions.Shared.TableLogOnInfo
        Dim myConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo

        Dim lFileName As String
        Dim myReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument

        Dim mail1 As New MailMessage()
        Dim smtp As New SmtpClient("192.168.0.169")

        'Me.Timer1.Enabled = False

        vEmailAll = vEmail & "," & vCC
        vReportID = "003"

        vQuery = "exec dbo.USP_PR_SearchPOSendMail '" & vApCode & "'"
        da1 = New SqlDataAdapter(vQuery, vConnection)
        ds1 = New DataSet
        da1.Fill(ds1, "Docno1")
        dt1 = ds1.Tables("Docno1")
        If dt1.Rows.Count > 0 Then
            For i = 0 To dt1.Rows.Count - 1
                vGetDocNo = dt1.Rows(i).Item("docno")
                vGetApName = dt1.Rows(i).Item("apname")
                vGetContactName = dt1.Rows(i).Item("contactname")

                vQuery = "exec dbo.USP_PO_DeleteDiscountAuto '" & vGetDocNo & "' "
                cmd = New SqlCommand(vQuery, vConnection)
                cmd.ExecuteNonQuery()

                vReportName = "RP_PO_PurchaseOrderAuto"
                vReportPath = Trim("W:\External\Reports\ReOrder\" & vReportName & ".rpt")

                vPDFName = Trim(vGetDocNo)
                vFileName = Trim("C:\PDFDocApp\" & vPDFName & ".pdf")
                vSubject = Trim("เอกสารใบสั่งซื้อสินค้าอัตโนมัติ จาก บริษัท นพดลพานิช จำกัด ส่งถึง" & vGetApName & "-" & vb6.Year(Now) & "-" & vb6.Month(Now) & "-" & vb6.Day(Now))
                vBody = "เรียน    " & vGetContactName & "   " & vGetApName & "   ทางบริษัท นพดลพานิช จำกัด ได้แนบเอกสารใบสั่งซื้อสินค้า ให้ทางผู้แทนจำหน่ายดังนี้  (กรณีเมลล์ไม่มีเอกสาร ใบสั่งซื้อสินค้าแนบ กรุณาเมลล์กลับหรือโทรแจ้งให้ทางแผนกจัดซื้อทราบด้วย)"

                myReport3.Load(vReportPath)

                myConnectionInfo.ServerName = "Nebula"
                myConnectionInfo.DatabaseName = "BCNP"
                myConnectionInfo.UserID = "sa"
                myConnectionInfo.Password = "[ibdkifu"

                myTables = myReport3.Database.Tables

                For Each myTable As CrystalDecisions.CrystalReports.Engine.Table In myTables
                    myTableLogonInfo = myTable.LogOnInfo
                    myTableLogonInfo.ConnectionInfo = myConnectionInfo
                    myTable.ApplyLogOnInfo(myTableLogonInfo)
                Next

                lFileName = vFileName
                myReport3.SetParameterValue("@DocNo", vGetDocNo)

                'กำหนด Formula ให้กับฟอร์มรายงาน
                'Dim vComputerName As String
                'myReport3.DataDefinition.FormulaFields("ComName").Text = "'" & vComputerName & "'"

                CrystalReportViewer2.ReportSource = myReport3

                myReport3.ExportToDisk(ExportFormatType.PortableDocFormat, lFileName)

                lFileName = vFileName

                mail1.From = New MailAddress(vFromMail)

                Dim strEmails1() As String = vEmailAll.Split(",")
                For Each str As String In strEmails1
                    If str <> "" Then
                        mail1.To.Add(str)
                    End If
                Next

                mail1.Subject = vSubject
                mail1.Body = vBody

                Dim att As New Attachment(lFileName)
                mail1.Attachments.Add(att)

            Next

            mail1.Subject = vSubject
            mail1.Body = vBody

            smtp.Send(mail1)


            vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vReportID & "','','','','','" & vApCode & "','" & vGetDocNo & "','" & vEmailAll & "'"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            'Me.Timer1.Enabled = True
        End If

        'ErrDescription:

        '        If Err.Description <> "" Then
        '            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        '            Exit Sub
        '        End If
    End Sub

    Public Sub SendMailPOApproveCAT(ByVal vReportID As String, ByVal vCat As String, ByVal vEmail As String, ByVal vCC As String, ByVal vFromMail As String)
        Dim myReport3 As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim i As Integer
        Dim vGetDocNo As String
        Dim vGetApName As String
        Dim vGetContactName As String
        Dim vPDFName As String
        Dim vFileName As String
        Dim vSubject As String
        Dim vBody As String
        Dim rptDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim vEmailAll As String
        Dim vType As Integer


        On Error Resume Next

        Dim myTables As CrystalDecisions.CrystalReports.Engine.Tables

        Dim myTableLogonInfo As CrystalDecisions.Shared.TableLogOnInfo
        Dim myConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo

        Dim lFileName As String
        Dim myReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument

        Dim mail1 As New MailMessage()
        Dim smtp As New SmtpClient("192.168.0.169")

        vCC = "it@nopadol.com"

        vEmailAll = vEmail & "," & vCC

        vType = 1
        vReportName = "RP_PR_PurchaseOrderApproveDaily_Cat"
        vPDFName = Trim("PO-Approve" & vCat & "-" & vb6.Year(Now) & "_" & vb6.Month(Now) & "_" & vb6.Day(Now))

        vReportPath = Trim("W:\External\Reports\ReOrder\" & vReportName & ".rpt")

        vFileName = Trim("C:\PDFDocApp\" & vPDFName & ".pdf")
        vSubject = Trim(vCat & "เอกสารใบสั่งซื้อสินค้าที่ทางแต่ละ CAT จะต้องอนุมัติ ตามนโยบาย เอกสารใบสั่งซื้อที่มีมูลค่าไม่เกิน 100,000 บาท")
        vBody = Trim("เอกสารใบสั่งซื้อสินค้าที่ทางแต่ละ CAT จะต้องอนุมัติ ประจำวัน ของ " & vCat)

        myReport3.Load(vReportPath)

        myConnectionInfo.ServerName = "Nebula"
        myConnectionInfo.DatabaseName = "BCNP"
        myConnectionInfo.UserID = "sa"
        myConnectionInfo.Password = "[ibdkifu"

        myTables = myReport3.Database.Tables

        For Each myTable As CrystalDecisions.CrystalReports.Engine.Table In myTables
            myTableLogonInfo = myTable.LogOnInfo
            myTableLogonInfo.ConnectionInfo = myConnectionInfo
            myTable.ApplyLogOnInfo(myTableLogonInfo)
        Next

        lFileName = vFileName
        myReport3.SetParameterValue("@vType", vType)
        myReport3.SetParameterValue("@vCAT", vCat)

        'กำหนด Formula ให้กับฟอร์มรายงาน
        'Dim vComputerName As String
        'myReport3.DataDefinition.FormulaFields("ComName").Text = "'" & vComputerName & "'"

        CrystalReportViewer3.ReportSource = myReport3

        myReport3.ExportToDisk(ExportFormatType.PortableDocFormat, lFileName)

        lFileName = vFileName

        mail1.From = New MailAddress(vFromMail)

        Dim strEmails1() As String = vEmailAll.Split(",")
        For Each str As String In strEmails1
            If str <> "" Then
                mail1.To.Add(str)
            End If
        Next

        mail1.Subject = vSubject
        mail1.Body = vBody

        Dim att As New Attachment(lFileName)
        mail1.Attachments.Add(att)


        mail1.Subject = vSubject
        mail1.Body = vBody

        smtp.Send(mail1)


        vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vReportID & "','','','','','','','" & vEmailAll & "'"
        cmd = New SqlCommand(vQuery, vConnection)
        cmd.ExecuteNonQuery()

        'Me.Timer1.Enabled = True


        'ErrDescription:

        '        If Err.Description <> "" Then
        '            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        '            Exit Sub
        '        End If
    End Sub

    Public Sub SendMailPOApproveMG(ByVal vEmail As String, ByVal vCC As String, ByVal vFromMail As String)
        Dim myReport4 As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim i As Integer
        Dim vGetDocNo As String
        Dim vGetApName As String
        Dim vGetContactName As String
        Dim vPDFName As String
        Dim vFileName As String
        Dim vSubject As String
        Dim vBody As String
        Dim rptDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim vEmailAll As String
        Dim vReportID As String
        Dim vCAT As String
        Dim vType As Integer

        'On Error GoTo ErrDescription
        On Error Resume Next

        Dim myTables As CrystalDecisions.CrystalReports.Engine.Tables

        Dim myTableLogonInfo As CrystalDecisions.Shared.TableLogOnInfo
        Dim myConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo

        Dim lFileName As String
        Dim myReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument

        Dim mail1 As New MailMessage()
        Dim smtp As New SmtpClient("192.168.0.169")

        'Me.Timer1.Enabled = False

        vCC = "it@nopadol.com"

        vEmailAll = vEmail & "," & vCC

        For i = 1 To 2
            vReportID = "008"
            If i = 1 Then
                vType = 0
                vReportName = "RP_PR_PurchaseOrderApproveDaily"
                vPDFName = Trim("PO-ApproveAll" & "-" & vb6.Year(Now) & "_" & vb6.Month(Now) & "_" & vb6.Day(Now))
            ElseIf i = 2 Then
                vType = 0
                vReportName = "RP_PR_PurchaseOrderApproveDaily_MG"
                vPDFName = Trim("PO-ApproveMG" & "-" & vb6.Year(Now) & "_" & vb6.Month(Now) & "_" & vb6.Day(Now))
            End If

            vReportPath = Trim("W:\External\Reports\ReOrder\" & vReportName & ".rpt")

            vFileName = Trim("C:\PDFDocApp\" & vPDFName & ".pdf")
            vSubject = "เอกสารใบสั่งซื้อสินค้าที่จะต้องอนุมัติทั้งหมดและส่วนของผู้อำนวยการจัดซื้อ "
            vBody = "เอกสารใบสั่งซื้อสินค้าที่ทางแต่ละ CAT จะต้องอนุมัติ ประจำวัน"

            myReport4.Load(vReportPath)

            myConnectionInfo.ServerName = "Nebula"
            myConnectionInfo.DatabaseName = "BCNP"
            myConnectionInfo.UserID = "sa"
            myConnectionInfo.Password = "[ibdkifu"

            myTables = myReport4.Database.Tables

            For Each myTable As CrystalDecisions.CrystalReports.Engine.Table In myTables
                myTableLogonInfo = myTable.LogOnInfo
                myTableLogonInfo.ConnectionInfo = myConnectionInfo
                myTable.ApplyLogOnInfo(myTableLogonInfo)
            Next

            lFileName = vFileName
            myReport4.SetParameterValue("@vType", vType)
            myReport4.SetParameterValue("@vCAT", "")

            'กำหนด Formula ให้กับฟอร์มรายงาน
            'Dim vComputerName As String
            'myReport3.DataDefinition.FormulaFields("ComName").Text = "'" & vComputerName & "'"

            CrystalReportViewer4.ReportSource = myReport4

            myReport4.ExportToDisk(ExportFormatType.PortableDocFormat, lFileName)

            lFileName = vFileName

            mail1.From = New MailAddress(vFromMail)

            Dim strEmails1() As String = vEmailAll.Split(",")
            For Each str As String In strEmails1
                If str <> "" Then
                    mail1.To.Add(str)
                End If
            Next

            mail1.Subject = vSubject
            mail1.Body = vBody

            Dim att As New Attachment(lFileName)
            mail1.Attachments.Add(att)

        Next

        mail1.Subject = vSubject
        mail1.Body = vBody

        smtp.Send(mail1)


        vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vReportID & "','','','','','','','" & vEmailAll & "'"
        cmd = New SqlCommand(vQuery, vConnection)
        cmd.ExecuteNonQuery()

        'Me.Timer1.Enabled = True


        'ErrDescription:

        '        If Err.Description <> "" Then
        '            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        '            Exit Sub
        '        End If
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Me.LBLTime.Text = Now '.Hour & ":" & Now.Minute & ":" & Now.Second
    End Sub

    'Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
    '    Dim vFileName As String
    '    Dim vSubject As String
    '    Dim vBody As String
    '    Dim vEmailAll As String
    '    Dim vPDFName As String
    '    Dim lFileName As String

    '    Try

    '        vPDFName = "POV5503-0308"
    '        vFileName = Trim("C:\PDFDocApp\" & vPDFName & ".pdf")
    '        vSubject = "ทดสอบการส่งอีเมลล์จาก แผนกคอมพิวเตอร์"
    '        vBody = "เอกสารใบสั่งซื้อสินค้า"
    '        vEmailAll = "it@nopadol.com"

    '        Dim Mail As New MailMessage
    '        Mail.From = New MailAddress("mc2@nopadol.com")

    '        Dim strEmails1() As String = vEmailAll.Split(",")
    '        For Each str As String In strEmails1
    '            If str <> "" Then
    '                Mail.To.Add(str)
    '            End If
    '        Next

    '        lFileName = vFileName
    '        Mail.Subject = vSubject
    '        Mail.Body = vBody

    '        Dim att As New Attachment(lFileName)
    '        Mail.Attachments.Add(att)

    '        Dim smtp As New SmtpClient("smtp.gmail.com")
    '        smtp.Port = 587
    '        smtp.EnableSsl = True
    '        smtp.Credentials = New System.Net.NetworkCredential("mc2@nopadol.com", "212224236248")
    '        smtp.Send(Mail)
    '        MsgBox("ทำการส่ง E-Mail เรียบร้อย", vbInformation, "ขอบคุณ")
    '    Catch ex As Exception
    '        MsgBox(Err.Description)
    '    End Try
    'End Sub


    Public Sub SendMailReport_Gmail(ByVal vReportID As String, ByVal vReportType As Integer, ByVal vProfitCenter As String, ByVal vExpertTeam As String, ByVal vSection As String, ByVal vDepartCode As String, ByVal vApCode As String, ByVal vDocNo As String, ByVal vEmail As String, ByVal vCC As String, ByVal vFromMail As String)
        Dim vPDFName As String
        Dim vFileName As String
        Dim vSubject As String
        Dim vBody As String
        Dim vEmailAll As String
        Dim da1 As SqlDataAdapter
        Dim ds1 As DataSet
        Dim dt1 As New DataTable

        Dim i As Integer
        Dim vGetDocNo As String


        On Error Resume Next

        Dim myTableLogonInfo As CrystalDecisions.Shared.TableLogOnInfo
        Dim myConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo
        Dim lFileName As String

        Dim mail1 As New MailMessage()
        Dim smtp As New SmtpClient("smtp.gmail.com")
        smtp.Port = 587
        smtp.EnableSsl = True
        'smtp.Credentials = New System.Net.NetworkCredential("it@nopadol.com", "[vdw,jwfh")
        smtp.Credentials = New System.Net.NetworkCredential("nopadol_mailauto@nopadol.com", "[vdw,jwfh2012")

        Me.Timer1.Enabled = False
        Me.Timer2.Enabled = False

        vEmailAll = vEmail & "," & vCC

        'vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vReportID & "','" & vExpertTeam & "','" & vProfitCenter & "','" & vDepartCode & "','" & vSection & "','" & vApCode & "','" & vDocNo & "','" & vEmailAll & "'"
        'cmd = New SqlCommand(vQuery, vConnection)
        'cmd.ExecuteNonQuery()

        If vReportID = "001" Or vReportID = "002" Or vReportID = "004" Or vReportID = "011" Or vReportID = "012" Then
            If vReportID = "001" Then
                vReportName = "RP_PR_StockRequestReOrderDaily"
                vReportPath = Trim("W:\External\Reports\ReOrder\" & vReportName & ".rpt")

                vPDFName = Trim(vProfitCenter & "-" & vReportID & "-" & vExpertTeam & "-" & vb6.Year(Now) & "_" & vb6.Month(Now) & "_" & vb6.Day(Now))
                vFileName = Trim("C:\PDFDocApp\" & vPDFName & ".pdf")
                vSubject = vProfitCenter & "---" & vExpertTeam & "-" & "สินค้าที่เสนอซื้อประจำวันตามทีมขาย (ReOrder)" & vExpertTeam
                vBody = "เป็นรายงาน เสนอซื้อสินค้าของทีม" & vExpertTeam

            ElseIf vReportID = "011" Then
                vReportName = "RP_PR_StockRequestReOrderDaily"
                vReportPath = Trim("W:\External\Reports\ReOrder\" & vReportName & ".rpt")

                vPDFName = Trim(vProfitCenter & "-" & vReportID & "-" & vExpertTeam & "-" & vb6.Year(Now) & "_" & vb6.Month(Now) & "_" & vb6.Day(Now))
                vFileName = Trim("C:\PDFDocApp\" & vPDFName & ".pdf")
                vSubject = vProfitCenter & "---" & vExpertTeam & "-" & "สินค้าที่เสนอซื้อประจำวันตามทีมขาย (ReOrder)" & vExpertTeam & "สาขา สันกำแพง"
                vBody = "เป็นรายงาน เสนอซื้อสินค้าของทีม" & vExpertTeam

            ElseIf vReportID = "002" Then
                vReportName = "RP_PR_StockRequestReOrderApproveDaily"
                vReportPath = Trim("W:\External\Reports\ReOrder\" & vReportName & ".rpt")

                vPDFName = Trim(vProfitCenter & "-" & vReportID & "-" & vExpertTeam & "-" & vb6.Year(Now) & "_" & vb6.Month(Now) & "_" & vb6.Day(Now))
                vFileName = Trim("C:\PDFDocApp\" & vPDFName & ".pdf")
                vSubject = vProfitCenter & "---" & vExpertTeam & "-" & "สินค้าที่เสนอซื้อที่อนุมัติแล้ว ประจำวันตามทีมขาย (ReOrder)" & vExpertTeam
                vBody = "เป็นรายงาน เสนอซื้อสินค้าที่อนุมัติจำนวนแล้ว ของทีม" & vExpertTeam

            ElseIf vReportID = "012" Then
                vReportName = "RP_PR_StockRequestReOrderApproveDaily"
                vReportPath = Trim("W:\External\Reports\ReOrder\" & vReportName & ".rpt")

                vPDFName = Trim(vProfitCenter & "-" & vReportID & "-" & vExpertTeam & "-" & vb6.Year(Now) & "_" & vb6.Month(Now) & "_" & vb6.Day(Now))
                vFileName = Trim("C:\PDFDocApp\" & vPDFName & ".pdf")
                vSubject = vProfitCenter & "---" & vExpertTeam & "-" & "สินค้าที่เสนอซื้อที่อนุมัติแล้ว ประจำวันตามทีมขาย (ReOrder)" & vExpertTeam & "สาขา สันกำแพง"
                vBody = "เป็นรายงาน เสนอซื้อสินค้าที่อนุมัติจำนวนแล้ว ของทีม" & vExpertTeam

            ElseIf vReportID = "004" Then
                vReportName = "RP_PR_TeamExpertReOrderConfirmQty"
                vReportPath = Trim("W:\External\Reports\ReOrder\" & vReportName & ".rpt")

                vPDFName = Trim(vReportID & "-" & vExpertTeam & "-" & vb6.Year(Now) & "_" & vb6.Month(Now) & "_" & vb6.Day(Now))
                vFileName = Trim("C:\PDFDocApp\" & vPDFName & ".pdf")
                vSubject = vExpertTeam & "-" & "สินค้าที่เสนอซื้อที่ทางจัดซื้อได้พิจารณาแล้ว ประจำวันตามทีมขาย (ReOrder)เพื่อให้ทางผู้อำนวยการจัดซื้อ" & vExpertTeam
                vBody = "เป็นรายงาน สินค้าที่เสนอซื้อที่ทางจัดซื้อได้พิจารณาแล้ว ของทีม" & vExpertTeam & "เพื่อให้ทางผู้อำนวยการจัดซื้อพิจารณาตรวจสอบ"
            End If
        End If

        Dim myReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        myReport.Load(vReportPath)

        myConnectionInfo.ServerName = "Nebula"
        myConnectionInfo.DatabaseName = "BCNP"
        myConnectionInfo.UserID = "sa"
        myConnectionInfo.Password = "[ibdkifu"

        Dim myTables As CrystalDecisions.CrystalReports.Engine.Tables
        myTables = myReport.Database.Tables

        For Each myTable As CrystalDecisions.CrystalReports.Engine.Table In myTables
            myTableLogonInfo = myTable.LogOnInfo
            myTableLogonInfo.ConnectionInfo = myConnectionInfo
            myTable.ApplyLogOnInfo(myTableLogonInfo)
        Next

        lFileName = vFileName
        myReport.SetParameterValue("@vProfit", vProfitCenter)
        myReport.SetParameterValue("@vType", vReportType)
        myReport.SetParameterValue("@vExpertTeam", vExpertTeam)
        myReport.SetParameterValue("@vSectionID", vSection)
        myReport.SetParameterValue("@vDepartment", vDepartCode)
        myReport.SetParameterValue("@vDocNo", vDocNo)

        CrystalReportViewer1.ReportSource = myReport


        myReport.ExportToDisk(ExportFormatType.PortableDocFormat, lFileName)


        lFileName = vFileName
        Dim mail As New MailMessage()
        mail.From = New MailAddress(vFromMail)

        Dim strEmails() As String = vEmailAll.Split(",")
        For Each str As String In strEmails
            If str <> "" Then
                mail.To.Add(str)
            End If
        Next

        mail.Subject = vSubject
        mail.Body = vBody

        Dim att As New Attachment(lFileName)
        mail.Attachments.Add(att)

        smtp.Send(mail)

        Me.Timer1.Enabled = True
        Me.Timer2.Enabled = True
    End Sub

    Public Sub SendMailItemTransfer_Gmail(ByVal vReportID As String, ByVal vReportType As Integer, ByVal vProfitCenter As String, ByVal vExpertTeam As String, ByVal vSection As String, ByVal vDepartCode As String, ByVal vApCode As String, ByVal vDocNo As String, ByVal vEmail As String, ByVal vCC As String, ByVal vFromMail As String)
        Dim vPDFName As String
        Dim vFileName As String
        Dim vSubject As String
        Dim vBody As String
        Dim vEmailAll As String
        Dim da1 As SqlDataAdapter
        Dim ds1 As DataSet
        Dim dt1 As New DataTable

        Dim i As Integer
        Dim vGetDocNo As String


        On Error Resume Next

        Dim myTableLogonInfo As CrystalDecisions.Shared.TableLogOnInfo
        Dim myConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo
        Dim lFileName As String

        Dim mail1 As New MailMessage()
        Dim smtp As New SmtpClient("smtp.gmail.com")
        smtp.Port = 587
        smtp.EnableSsl = True
        'smtp.Credentials = New System.Net.NetworkCredential("it@nopadol.com", "[vdw,jwfh")
        smtp.Credentials = New System.Net.NetworkCredential("nopadol_mailauto@nopadol.com", "[vdw,jwfh2012")

        Me.Timer1.Enabled = False
        Me.Timer2.Enabled = False

        vEmailAll = vEmail & "," & vCC '"nopadol_mailauto@nopadol.com" '

        'vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vReportID & "','" & vExpertTeam & "','" & vProfitCenter & "','" & vDepartCode & "','" & vSection & "','" & vApCode & "','" & vDocNo & "','" & vEmailAll & "'"
        'cmd = New SqlCommand(vQuery, vConnection)
        'cmd.ExecuteNonQuery()

        If vReportID = "013" Or vReportID = "014" Then
            If vReportID = "013" Then
                vReportName = "RP_PR_ItemTransferBetweenProfit"
                vReportPath = Trim("W:\External\Reports\ReOrder\" & vReportName & ".rpt")

                vPDFName = Trim(vProfitCenter & "-" & vReportID & "-" & vExpertTeam & "-" & vb6.Year(Now) & "_" & vb6.Month(Now) & "_" & vb6.Day(Now))
                vFileName = Trim("C:\PDFDocApp\" & vPDFName & ".pdf")
                vSubject = vProfitCenter & "---" & "สินค้าที่เสนอซื้อประจำวันที่ควรจะโอนมาจากสาขา สันกำแพง"
                vBody = "สินค้าที่เสนอซื้อประจำวันที่ควรจะโอนมาจากสาขา สันกำแพง"

            ElseIf vReportID = "014" Then
                vReportName = "RP_PR_ItemTransferBetweenProfit"
                vReportPath = Trim("W:\External\Reports\ReOrder\" & vReportName & ".rpt")

                vPDFName = Trim(vProfitCenter & "-" & vReportID & "-" & vExpertTeam & "-" & vb6.Year(Now) & "_" & vb6.Month(Now) & "_" & vb6.Day(Now))
                vFileName = Trim("C:\PDFDocApp\" & vPDFName & ".pdf")
                vSubject = vProfitCenter & "---" & "สินค้าที่เสนอซื้อประจำวันที่ควรจะโอนมาจากสาขา สำนักงานใหญ่"
                vBody = "สินค้าที่เสนอซื้อประจำวันที่ควรจะโอนมาจากสาขา สำนักงานใหญ่"
            End If
        End If

        Dim myReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        myReport.Load(vReportPath)

        myConnectionInfo.ServerName = "Nebula"
        myConnectionInfo.DatabaseName = "BCNP"
        myConnectionInfo.UserID = "sa"
        myConnectionInfo.Password = "[ibdkifu"

        Dim myTables As CrystalDecisions.CrystalReports.Engine.Tables
        myTables = myReport.Database.Tables

        For Each myTable As CrystalDecisions.CrystalReports.Engine.Table In myTables
            myTableLogonInfo = myTable.LogOnInfo
            myTableLogonInfo.ConnectionInfo = myConnectionInfo
            myTable.ApplyLogOnInfo(myTableLogonInfo)
        Next

        lFileName = vFileName
        myReport.SetParameterValue("@vProfit", vProfitCenter)

        CrystalReportViewer1.ReportSource = myReport


        myReport.ExportToDisk(ExportFormatType.PortableDocFormat, lFileName)


        lFileName = vFileName
        Dim mail As New MailMessage()
        mail.From = New MailAddress(vFromMail)

        Dim strEmails() As String = vEmailAll.Split(",")
        For Each str As String In strEmails
            If str <> "" Then
                mail.To.Add(str)
            End If
        Next

        mail.Subject = vSubject
        mail.Body = vBody

        Dim att As New Attachment(lFileName)
        mail.Attachments.Add(att)

        smtp.Send(mail)

        Me.Timer1.Enabled = True
        Me.Timer2.Enabled = True
    End Sub

    Public Sub SendMailPO_Gmail(ByVal vApCode As String, ByVal vDocNo As String, ByVal vEmail As String, ByVal vCC As String, ByVal vFromMail As String)
        Dim myReport3 As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim i As Integer
        Dim vGetDocNo As String
        Dim vGetApName As String
        Dim vGetContactName As String
        Dim vPDFName As String
        Dim vFileName As String
        Dim vSubject As String
        Dim vBody As String
        Dim rptDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim vEmailAll As String
        Dim vReportID As String
        Dim vMydescription As String
        Dim vPicture1 As String
        Dim vPicture2 As String
        Dim vPicture3 As String

        Dim vLink As String
        Dim vDocDate As String
        Dim vLeadDate As String
        Dim vSendDate As String
        Dim vUserID As String
        Dim vCatCode As String

        On Error Resume Next

        Dim myTables As CrystalDecisions.CrystalReports.Engine.Tables
        Dim myTableLogonInfo As CrystalDecisions.Shared.TableLogOnInfo
        Dim myConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo

        Dim lFileName As String
        Dim myReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument

        Dim mail1 As New MailMessage()
        Dim smtp As New SmtpClient("smtp.gmail.com")

        smtp.Port = 587
        smtp.EnableSsl = True
        smtp.UseDefaultCredentials = False
        smtp.Credentials = New System.Net.NetworkCredential("nopadol_mailauto@nopadol.com", "[vdw,jwfh2012")

        Me.Timer1.Enabled = False
        Me.Timer2.Enabled = False

        vEmailAll = vEmail & "," & vCC & "," & vFromMail
        vReportID = "003"

        vQuery = "exec dbo.USP_PR_SearchPOSendMail '" & vApCode & "','" & vDocNo & "'"
        da1 = New SqlDataAdapter(vQuery, vConnection)
        ds1 = New DataSet
        da1.Fill(ds1, "Docno1")
        dt1 = ds1.Tables("Docno1")
        If dt1.Rows.Count > 0 Then

            vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vReportID & "','','','','','" & vApCode & "','" & vDocNo & "','" & vEmailAll & "'"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            For i = 0 To dt1.Rows.Count - 1
                vGetDocNo = dt1.Rows(i).Item("docno")
                vGetApName = dt1.Rows(i).Item("apname")
                vGetContactName = dt1.Rows(i).Item("contactname")
                vMydescription = dt1.Rows(i).Item("mydescription")
                vPicture1 = dt1.Rows(i).Item("picture1")
                vPicture2 = dt1.Rows(i).Item("picture2")
                vPicture3 = dt1.Rows(i).Item("picture3")

                vDocDate = dt1.Rows(i).Item("docdate")
                vLeadDate = dt1.Rows(i).Item("leaddate")
                vSendDate = dt1.Rows(i).Item("maildate")
                vUserID = dt1.Rows(i).Item("creatorcode")
                vCatCode = dt1.Rows(i).Item("expertcode")


                vQuery = "exec dbo.USP_PO_DeleteDiscountAuto '" & vGetDocNo & "' "
                cmd = New SqlCommand(vQuery, vConnection)
                cmd.ExecuteNonQuery()

                vReportName = "RP_PO_PurchaseOrderAuto"
                vReportPath = Trim("W:\External\Reports\ReOrder\" & vReportName & ".rpt")

                vPDFName = Trim(vGetDocNo)
                vFileName = Trim("C:\PDFDocApp\" & vPDFName & ".pdf")
                vSubject = Trim("บริษัท นพดลพานิช จำกัด จัดส่งเอกสารใบสั่งซื้อสินค้าอีเลคโทรนิคส์ ถึง" & vGetApName & "-" & vb6.Year(Now) & "-" & vb6.Month(Now) & "-" & vb6.Day(Now))
                vBody = "เรียน    " & vGetContactName & "   " & vGetApName & "   ทางบริษัท นพดลพานิช จำกัด ได้แนบเอกสารใบสั่งซื้อสินค้า ให้ทางผู้แทนจำหน่ายดังนี้  (กรณีเมลล์ไม่มีเอกสาร ใบสั่งซื้อสินค้าแนบ กรุณาเมลล์กลับหรือโทรแจ้งให้ทางแผนกจัดซื้อทราบด้วย และกรณีต้องการตอบกลับเมลล์ กรุณาเลือก ==ตอบทุกคน==)"

                myReport3.Load(vReportPath)

                myConnectionInfo.ServerName = "Nebula"
                myConnectionInfo.DatabaseName = "BCNP"
                myConnectionInfo.UserID = "sa"
                myConnectionInfo.Password = "[ibdkifu"

                myTables = myReport3.Database.Tables

                For Each myTable As CrystalDecisions.CrystalReports.Engine.Table In myTables
                    myTableLogonInfo = myTable.LogOnInfo
                    myTableLogonInfo.ConnectionInfo = myConnectionInfo
                    myTable.ApplyLogOnInfo(myTableLogonInfo)
                Next

                lFileName = vFileName
                myReport3.SetParameterValue("@DocNo", vGetDocNo)

                'กำหนด Formula ให้กับฟอร์มรายงาน
                'Dim vComputerName As String
                'myReport3.DataDefinition.FormulaFields("ComName").Text = "'" & vComputerName & "'"

                CrystalReportViewer2.ReportSource = myReport3

                myReport3.ExportToDisk(ExportFormatType.PortableDocFormat, lFileName)

                lFileName = vFileName

                mail1.From = New MailAddress(vFromMail)

                Dim strEmails1() As String = vEmailAll.Split(",")
                For Each str As String In strEmails1
                    If str <> "" Then
                        mail1.To.Add(str)
                    End If
                Next

                mail1.IsBodyHtml = True

                vLink = "<a href=""http://www.nopadol.com/reorder/frmupdatepo.php?gDocNo=" & vDocNo & "&gVendorCode=" & vApCode & "&gDocDate=" & vDocDate & "&gSendDate=" & vSendDate & "&gCat=" & vCatCode & "&gDueDate=" & vLeadDate & "&gMCName=" & vUserID & """>โปรดระบุ วันที่จัดส่งสินค้าให้กับทาง บริษัทนพดลพานิช จำกัด Click here</a>"


                mail1.Subject = vSubject
                If vMydescription <> "" Then
                    vBody = vBody & "<br>" & " ******************************************** ข้อความเพิ่มเติม ******************************************** " & "</br>" & "<br>" & "<font size=5 color=Red bold = 5>" & vMydescription & "</font>" & "</br>"
                Else
                    vBody = vBody
                End If

                vBody = vBody & "<br>" & " ******************************************** Link ระบุวันที่จัดส่งสินค้า ******************************************** " & "</br>" & "<br>" & "<font size=5 color=blue bold = 5>" & vLink & "</font>" & "</br>"

                mail1.Body = vBody
                Dim att As New Attachment(lFileName)
                mail1.Attachments.Add(att)

                If vPicture1 <> "" Then
                    Dim att1 As New Attachment(vPicture1)
                    mail1.Attachments.Add(att1)
                End If

                If vPicture2 <> "" Then
                    Dim att2 As New Attachment(vPicture2)
                    mail1.Attachments.Add(att2)
                End If

                If vPicture3 <> "" Then
                    Dim att3 As New Attachment(vPicture3)
                    mail1.Attachments.Add(att3)
                End If

            Next

            smtp.Send(mail1)

            Me.Timer1.Enabled = True
            Me.Timer2.Enabled = True

        End If


ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description & " " & vDocNo & " " & "ส่งไม่ผ่าน")


            Dim FILE_NAME As String = "C:\PDFDocApp\SendEmailLogs.txt"
            Dim errText As String

            errText = Now & " " & Err.Description & " " & "เจ้าหนี้รหัส" & " " & vApCode & "เลขที่ใบสั่งซื้อ" & vDocNo & " " & "โอนไม่ผ่าน มีปัญหาให้ปรับการส่งเมลล์ใหม่"

            Dim objWriter As New System.IO.StreamWriter(FILE_NAME, True)

            objWriter.WriteLine(errText)

            objWriter.Close()

            'vQuery = "exec dbo.USP_PR_UpdateSendPOEmailAgain '" & vReportID & "','" & vApCode & "','" & vDocNo & "'"
            'cmd = New SqlCommand(vQuery, vConnection)
            'cmd.ExecuteNonQuery()

            Exit Sub
        End If

    End Sub


    Public Sub SendMailReturn_Gmail(ByVal vApCode As String, ByVal vDocNo As String, ByVal vEmail As String, ByVal vCC As String, ByVal vFromMail As String, ByVal vGetContactName As String, ByVal vGetNopadolContactName As String)
        Dim myReport3 As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim i As Integer
        Dim vGetDocNo As String
        Dim vPDFName As String
        Dim vFileName As String
        Dim vSubject As String
        Dim vBody As String
        Dim rptDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim vEmailAll As String
        Dim vReportID As String
        Dim vMydescription As String
        Dim vPicture1 As String
        Dim vPicture2 As String
        Dim vPicture3 As String

        Dim vLink As String
        Dim vDocDate As String
        Dim vLeadDate As String
        Dim vSendDate As String
        Dim vUserID As String
        Dim vCatCode As String

        On Error Resume Next

        Dim myTables As CrystalDecisions.CrystalReports.Engine.Tables
        Dim myTableLogonInfo As CrystalDecisions.Shared.TableLogOnInfo
        Dim myConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo

        Dim lFileName As String
        Dim myReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument

        Dim mail1 As New MailMessage()
        Dim smtp As New SmtpClient("smtp.gmail.com")

        smtp.Port = 587
        smtp.EnableSsl = True
        smtp.UseDefaultCredentials = False
        smtp.Credentials = New System.Net.NetworkCredential("nopadol_mailauto@nopadol.com", "[vdw,jwfh2012")

        Me.Timer1.Enabled = False
        Me.Timer2.Enabled = False

        If vEmail <> "" Then
            vEmailAll = vEmail & "," & "it@nopadol.com,ch@nopadol.com,vilaivan@nopadol.com" 'vCC & "," & vFromMail 'vEmail & "," & vCC & "," & vFromMail
        Else
            vEmailAll = "it@nopadol.com,ch@nopadol.com,vilaivan@nopadol.com"
        End If

        vReportID = "SRF"

        vReportName = "RP_NP_StkRefundLetter"
        vReportPath = Trim("W:\External\Reports\ReOrder\" & vReportName & ".rpt")

        vPDFName = Trim(vDocNo)
        vFileName = Trim("C:\PDFDocApp\" & vPDFName & ".pdf")
        vSubject = Trim("บริษัท นพดลพานิช จำกัด จัดส่งเอกสารใบส่งคืนสินค้าอีเลคโทรนิคส์ ถึง    " & vGetContactName & "-" & vb6.Year(Now) & "-" & vb6.Month(Now) & "-" & vb6.Day(Now))
        vBody = "เรียน    " & vGetContactName & "    ทางบริษัท นพดลพานิช จำกัด ได้แนบเอกสารใบส่งคืนสินค้า ให้ทางผู้แทนจำหน่ายดังนี้  (กรณีเมลล์ไม่มีเอกสาร ใบส่งคืนสินค้าแนบ กรุณาเมลล์กลับหรือโทรแจ้งให้ทางแผนกบัญชีทราบด้วย และกรณีต้องการตอบกลับเมลล์ กรุณาเลือก ==ตอบทุกคน==)"

        myReport3.Load(vReportPath)

        myConnectionInfo.ServerName = "Nebula"
        myConnectionInfo.DatabaseName = "BCNP"
        myConnectionInfo.UserID = "sa"
        myConnectionInfo.Password = "[ibdkifu"

        myTables = myReport3.Database.Tables

        For Each myTable As CrystalDecisions.CrystalReports.Engine.Table In myTables
            myTableLogonInfo = myTable.LogOnInfo
            myTableLogonInfo.ConnectionInfo = myConnectionInfo
            myTable.ApplyLogOnInfo(myTableLogonInfo)
        Next

        lFileName = vFileName
        myReport3.SetParameterValue("@vDocNo", vDocNo)

        CrystalReportViewer2.ReportSource = myReport3

        myReport3.ExportToDisk(ExportFormatType.PortableDocFormat, lFileName)

        lFileName = vFileName

        mail1.From = New MailAddress(vFromMail)

        Dim strEmails1() As String = vEmailAll.Split(",")
        For Each str As String In strEmails1
            If str <> "" Then
                mail1.To.Add(str)
            End If
        Next

        mail1.IsBodyHtml = True

        mail1.Subject = vSubject
        vBody = vBody & "<br>" & " ******************************************** หากมีข้อสงสัย กรุณาติดต่อกลับ ******************************************** " & "</br>" & "<br>" & "<font size=5 color=blue bold = 5>" & vGetNopadolContactName & "</font>" & "</br>"

        mail1.Body = vBody
        Dim att As New Attachment(lFileName)
        mail1.Attachments.Add(att)

        smtp.Send(mail1)

        Me.Timer1.Enabled = True
        Me.Timer2.Enabled = True



ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description & " " & vDocNo & " " & "ส่งไม่ผ่าน")


            Dim FILE_NAME As String = "C:\PDFDocApp\SendEmailLogs.txt"
            Dim errText As String

            errText = Now & " " & Err.Description & " " & "เจ้าหนี้รหัส" & " " & vApCode & "เลขที่ใบส่งคืน" & vDocNo & " " & "โอนไม่ผ่าน มีปัญหาให้ปรับการส่งเมลล์ใหม่"
            Dim objWriter As New System.IO.StreamWriter(FILE_NAME, True)
            objWriter.WriteLine(errText)
            objWriter.Close()

            Exit Sub
        End If

    End Sub


    Public Sub SendMailSCG_Gmail(ByVal vEmail As String, ByVal vCC As String, ByVal vFromMail As String)
        Dim vSubject As String
        Dim vBody As String
        Dim vEmailAll As String
        Dim vReportID As String
        Dim vFileName1 As String
        Dim vFileName2 As String
        Dim vFileName3 As String
        Dim vFileName4 As String
        Dim vFileName5 As String
        Dim vZipName As String


        'On Error GoTo ErrDescription

        On Error Resume Next

        Dim mail1 As New MailMessage()
        Dim smtp As New SmtpClient("smtp.gmail.com")

        smtp.Port = 587
        smtp.EnableSsl = True
        smtp.UseDefaultCredentials = False
        smtp.Credentials = New System.Net.NetworkCredential("nopadol_mailauto@nopadol.com", "[vdw,jwfh2012")

        Me.Timer1.Enabled = False
        Me.Timer2.Enabled = False

        vEmailAll = vEmail & "," & vCC & "," & vFromMail
        vReportID = "SCG"

        vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vReportID & "','','','','','','','" & vEmailAll & "'"
        cmd = New SqlCommand(vQuery, vConnection)
        cmd.ExecuteNonQuery()

        vSubject = Trim("ข้อมูลบริษัท นพดลพานิช จำกัด ณ วันที่ " & vb6.Year(Now) & "-" & vb6.Month(Now) & "-" & vb6.Day(Now))
        vBody = "ข้อมูลบริษัท นพดลพานิช จำกัด  (กรณีเมลล์ไม่มีเอกสารแนบ กรุณาเมลล์กลับหรือโทรแจ้ง ทางแผนก IT ด้วยนะครับ"

        mail1.From = New MailAddress(vFromMail)

        Dim strEmails1() As String = vEmailAll.Split(",")
        For Each str As String In strEmails1
            If str <> "" Then
                mail1.To.Add(str)
            End If
        Next

        mail1.Subject = vSubject
        mail1.Body = vBody

        Dim vMonth As String
        Dim vLenMonth As Integer
        Dim vDay As String
        Dim vLenDay As Integer
        Dim vYear As String

        vLenDay = vb6.Len(RTrim(DateAdd(DateInterval.Day, -1, Now).Day))
        If vLenDay < 2 Then
            vDay = "0" & RTrim(DateAdd(DateInterval.Day, -1, Now).Day)
        Else
            vDay = RTrim(DateAdd(DateInterval.Day, -1, Now).Day)
        End If

        vLenMonth = vb6.Len(RTrim(DateAdd(DateInterval.Day, -1, Now).Month))
        If vLenMonth < 2 Then
            vMonth = "0" & RTrim(DateAdd(DateInterval.Day, -1, Now).Month)
        Else
            vMonth = RTrim(DateAdd(DateInterval.Day, -1, Now).Month)
        End If

        vYear = RTrim(DateAdd(DateInterval.Day, -1, Now).Year)

        vFileName1 = "\\nebula\BCS\crm\3001278_MonthlyM-" & vDay & "-" & vMonth & "-" & vYear & ".txt"
        vFileName2 = "\\nebula\BCS\crm\3001278_MonthlySTK-" & vDay & "-" & vMonth & "-" & vYear & ".txt"
        vFileName3 = "\\nebula\BCS\crm\3001278_WeeklyM-" & vDay & "-" & vMonth & "-" & vYear & ".txt"
        vFileName4 = "\\nebula\BCS\crm\3001278_WeeklyPO-" & vDay & "-" & vMonth & "-" & vYear & ".txt"
        vFileName5 = "\\nebula\BCS\crm\3001278_WeeklySO-" & vDay & "-" & vMonth & "-" & vYear & ".txt"

        vZipName = "\\nebula\BCS\crm\3001278_" & vDay & "-" & vMonth & "-" & vYear & ".zip"

        Using zip As ZipFile = New ZipFile
            zip.AddFile(vFileName1)
            zip.AddFile(vFileName2)
            zip.AddFile(vFileName3)
            zip.AddFile(vFileName4)
            zip.AddFile(vFileName5)
            zip.Save(vZipName)
        End Using


        Dim att1 As New Attachment(vZipName)
        mail1.Attachments.Add(att1)


        smtp.Send(mail1)

        Me.Timer1.Enabled = True
        Me.Timer2.Enabled = True

ErrDescription:
        If Err.Description <> "" Then

            Dim FILE_NAME As String = "C:\PDFDocApp\SendEmailLogs.txt"
            Dim errText As String

            errText = Now & " " & Err.Description & "โอนไม่ผ่าน มีปัญหาให้ปรับการส่งเมลล์ใหม่"

            Dim objWriter As New System.IO.StreamWriter(FILE_NAME, True)

            objWriter.WriteLine(errText)

            objWriter.Close()

            'vQuery = "exec dbo.USP_PR_UpdateSendPOEmailAgain '" & vReportID & "','" & vApCode & "','" & vDocNo & "'"
            'cmd = New SqlCommand(vQuery, vConnection)
            'cmd.ExecuteNonQuery()

            Exit Sub
        End If

    End Sub


    Public Sub SendMailPOApproveCAT_Gmail(ByVal vReportID As String, ByVal vCat As String, ByVal vEmail As String, ByVal vCC As String, ByVal vFromMail As String)
        Dim myReport3 As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim i As Integer
        Dim vGetDocNo As String
        Dim vGetApName As String
        Dim vGetContactName As String
        Dim vPDFName As String
        Dim vFileName As String
        Dim vSubject As String
        Dim vBody As String
        Dim rptDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim vEmailAll As String
        Dim vType As Integer


        On Error Resume Next

        Dim myTables As CrystalDecisions.CrystalReports.Engine.Tables

        Dim myTableLogonInfo As CrystalDecisions.Shared.TableLogOnInfo
        Dim myConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo

        Dim lFileName As String
        Dim myReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument

        Dim mail1 As New MailMessage()
        Dim smtp As New SmtpClient("smtp.gmail.com")
        smtp.Port = 587
        smtp.EnableSsl = True
        'smtp.Credentials = New System.Net.NetworkCredential("it@nopadol.com", "[vdw,jwfh")
        smtp.Credentials = New System.Net.NetworkCredential("nopadol_mailauto@nopadol.com", "[vdw,jwfh2012")

        Me.Timer1.Enabled = False
        Me.Timer2.Enabled = False

        vCC = "it@nopadol.com"
        vEmailAll = vEmail & "," & vCC

        'vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vReportID & "','','','','','','','" & vEmailAll & "'"
        'cmd = New SqlCommand(vQuery, vConnection)
        'cmd.ExecuteNonQuery()


        vType = 1
        vReportName = "RP_PR_PurchaseOrderApproveDaily_Cat"
        vPDFName = Trim("PO-Approve" & vCat & "-" & vb6.Year(Now) & "_" & vb6.Month(Now) & "_" & vb6.Day(Now))

        vReportPath = Trim("W:\External\Reports\ReOrder\" & vReportName & ".rpt")

        vFileName = Trim("C:\PDFDocApp\" & vPDFName & ".pdf")
        vSubject = Trim(vCat & "เอกสารใบสั่งซื้อสินค้าที่ทางแต่ละ CAT จะต้องอนุมัติ ตามนโยบาย เอกสารใบสั่งซื้อที่มีมูลค่าไม่เกิน 100,000 บาท")
        vBody = Trim("เอกสารใบสั่งซื้อสินค้าที่ทางแต่ละ CAT จะต้องอนุมัติ ประจำวัน ของ " & vCat)

        myReport3.Load(vReportPath)

        myConnectionInfo.ServerName = "Nebula"
        myConnectionInfo.DatabaseName = "BCNP"
        myConnectionInfo.UserID = "sa"
        myConnectionInfo.Password = "[ibdkifu"

        myTables = myReport3.Database.Tables

        For Each myTable As CrystalDecisions.CrystalReports.Engine.Table In myTables
            myTableLogonInfo = myTable.LogOnInfo
            myTableLogonInfo.ConnectionInfo = myConnectionInfo
            myTable.ApplyLogOnInfo(myTableLogonInfo)
        Next

        lFileName = vFileName
        myReport3.SetParameterValue("@vType", vType)
        myReport3.SetParameterValue("@vCAT", vCat)

        'กำหนด Formula ให้กับฟอร์มรายงาน
        'Dim vComputerName As String
        'myReport3.DataDefinition.FormulaFields("ComName").Text = "'" & vComputerName & "'"

        CrystalReportViewer3.ReportSource = myReport3

        myReport3.ExportToDisk(ExportFormatType.PortableDocFormat, lFileName)

        lFileName = vFileName

        mail1.From = New MailAddress(vFromMail)

        Dim strEmails1() As String = vEmailAll.Split(",")
        For Each str As String In strEmails1
            If str <> "" Then
                mail1.To.Add(str)
            End If
        Next

        mail1.Subject = vSubject
        mail1.Body = vBody

        Dim att As New Attachment(lFileName)
        mail1.Attachments.Add(att)


        mail1.Subject = vSubject
        mail1.Body = vBody

        smtp.Send(mail1)

        Me.Timer1.Enabled = True
        Me.Timer2.Enabled = True
    End Sub


    Public Sub SendMailPOApproveMG_Gmail(ByVal vEmail As String, ByVal vCC As String, ByVal vFromMail As String)
        Dim myReport4 As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim i As Integer
        Dim vGetDocNo As String
        Dim vGetApName As String
        Dim vGetContactName As String
        Dim vPDFName As String
        Dim vFileName As String
        Dim vSubject As String
        Dim vBody As String
        Dim rptDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim vEmailAll As String
        Dim vReportID As String
        Dim vCAT As String
        Dim vType As Integer

        On Error Resume Next

        Dim myTables As CrystalDecisions.CrystalReports.Engine.Tables

        Dim myTableLogonInfo As CrystalDecisions.Shared.TableLogOnInfo
        Dim myConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo

        Dim lFileName As String
        Dim myReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument

        Dim mail1 As New MailMessage()
        Dim smtp As New SmtpClient("smtp.gmail.com")
        smtp.Port = 587
        smtp.EnableSsl = True
        'smtp.Credentials = New System.Net.NetworkCredential("it@nopadol.com", "[vdw,jwfh")
        smtp.Credentials = New System.Net.NetworkCredential("nopadol_mailauto@nopadol.com", "[vdw,jwfh2012")

        Me.Timer1.Enabled = False
        Me.Timer2.Enabled = False

        vCC = "it@nopadol.com"

        vEmailAll = vEmail & "," & vCC
        vReportID = "008"

        'vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vReportID & "','','','','','','','" & vEmailAll & "'"
        'cmd = New SqlCommand(vQuery, vConnection)
        'cmd.ExecuteNonQuery()


        For i = 1 To 2
            If i = 1 Then
                vType = 0
                vReportName = "RP_PR_PurchaseOrderApproveDaily"
                vPDFName = Trim("PO-ApproveAll" & "-" & vb6.Year(Now) & "_" & vb6.Month(Now) & "_" & vb6.Day(Now))
            ElseIf i = 2 Then
                vType = 0
                vReportName = "RP_PR_PurchaseOrderApproveDaily_MG"
                vPDFName = Trim("PO-ApproveMG" & "-" & vb6.Year(Now) & "_" & vb6.Month(Now) & "_" & vb6.Day(Now))
            End If

            vReportPath = Trim("W:\External\Reports\ReOrder\" & vReportName & ".rpt")

            vFileName = Trim("C:\PDFDocApp\" & vPDFName & ".pdf")
            vSubject = "เอกสารใบสั่งซื้อสินค้าที่จะต้องอนุมัติทั้งหมดและส่วนของผู้อำนวยการจัดซื้อ "
            vBody = "เอกสารใบสั่งซื้อสินค้าที่ทางแต่ละ CAT จะต้องอนุมัติ ประจำวัน"

            myReport4.Load(vReportPath)

            myConnectionInfo.ServerName = "Nebula"
            myConnectionInfo.DatabaseName = "BCNP"
            myConnectionInfo.UserID = "sa"
            myConnectionInfo.Password = "[ibdkifu"

            myTables = myReport4.Database.Tables

            For Each myTable As CrystalDecisions.CrystalReports.Engine.Table In myTables
                myTableLogonInfo = myTable.LogOnInfo
                myTableLogonInfo.ConnectionInfo = myConnectionInfo
                myTable.ApplyLogOnInfo(myTableLogonInfo)
            Next

            lFileName = vFileName
            myReport4.SetParameterValue("@vType", vType)
            myReport4.SetParameterValue("@vCAT", "")

            'กำหนด Formula ให้กับฟอร์มรายงาน
            'Dim vComputerName As String
            'myReport3.DataDefinition.FormulaFields("ComName").Text = "'" & vComputerName & "'"

            CrystalReportViewer4.ReportSource = myReport4

            myReport4.ExportToDisk(ExportFormatType.PortableDocFormat, lFileName)

            lFileName = vFileName

            mail1.From = New MailAddress(vFromMail)

            Dim strEmails1() As String = vEmailAll.Split(",")
            For Each str As String In strEmails1
                If str <> "" Then
                    mail1.To.Add(str)
                End If
            Next

            mail1.Subject = vSubject
            mail1.Body = vBody

            Dim att As New Attachment(lFileName)
            mail1.Attachments.Add(att)

        Next

        mail1.Subject = vSubject
        mail1.Body = vBody

        smtp.Send(mail1)

        Me.Timer1.Enabled = True
        Me.Timer2.Enabled = True

    End Sub

    Public Sub SendMailChangePrice_Gmail(ByVal vReportID As String, ByVal vEmail As String, ByVal vCC As String, ByVal vFromMail As String)
        Dim myReport3 As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim i As Integer
        Dim vGetDocNo As String
        Dim vGetApName As String
        Dim vGetContactName As String
        Dim vPDFName As String
        Dim vFileName As String
        Dim vSubject As String
        Dim vBody As String
        Dim rptDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim vEmailAll As String
        Dim vType As Integer

        On Error Resume Next

        vCC = "it@nopadol.com"
        vEmailAll = vEmail & "," & vCC

        vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vReportID & "','','','','','','','" & vEmailAll & "'"
        cmd = New SqlCommand(vQuery, vConnection)
        cmd.ExecuteNonQuery()


        Dim myTables As CrystalDecisions.CrystalReports.Engine.Tables

        Dim myTableLogonInfo As CrystalDecisions.Shared.TableLogOnInfo
        Dim myConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo

        Dim lFileName As String
        Dim myReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument

        Dim mail1 As New MailMessage()
        Dim smtp As New SmtpClient("smtp.gmail.com")
        smtp.Port = 587
        smtp.EnableSsl = True
        'smtp.Credentials = New System.Net.NetworkCredential("it@nopadol.com", "[vdw,jwfh")
        smtp.Credentials = New System.Net.NetworkCredential("nopadol_mailauto@nopadol.com", "[vdw,jwfh2012")

        Me.Timer1.Enabled = False
        Me.Timer2.Enabled = False

        vType = 1
        vReportName = "RP_NP_ICIChangePriceDaily"
        vPDFName = Trim("ChangePrice" & "-" & vb6.Year(Now) & "_" & vb6.Month(Now) & "_" & vb6.Day(Now))

        vReportPath = Trim("W:\External\Reports\ReOrder\" & vReportName & ".rpt")

        vFileName = Trim("C:\PDFDocApp\" & vPDFName & ".pdf")
        vSubject = Trim("รายงาน การเปลี่ยนราคาสินค้าประจำวัน ของสินค้ายี่ห้อ ไอซีไอ")
        vBody = Trim("รายงาน การเปลี่ยนราคาสินค้าประจำวัน ของสินค้ายี่ห้อ ไอซีไอ")

        myReport3.Load(vReportPath)

        myConnectionInfo.ServerName = "Nebula"
        myConnectionInfo.DatabaseName = "BCNP"
        myConnectionInfo.UserID = "sa"
        myConnectionInfo.Password = "[ibdkifu"

        myTables = myReport3.Database.Tables

        For Each myTable As CrystalDecisions.CrystalReports.Engine.Table In myTables
            myTableLogonInfo = myTable.LogOnInfo
            myTableLogonInfo.ConnectionInfo = myConnectionInfo
            myTable.ApplyLogOnInfo(myTableLogonInfo)
        Next

        lFileName = vFileName

        CrystalReportViewer3.ReportSource = myReport3

        myReport3.ExportToDisk(ExportFormatType.PortableDocFormat, lFileName)

        lFileName = vFileName

        mail1.From = New MailAddress(vFromMail)

        Dim strEmails1() As String = vEmailAll.Split(",")
        For Each str As String In strEmails1
            If str <> "" Then
                mail1.To.Add(str)
            End If
        Next

        mail1.Subject = vSubject
        mail1.Body = vBody

        Dim att As New Attachment(lFileName)
        mail1.Attachments.Add(att)

        mail1.Subject = vSubject
        mail1.Body = vBody

        smtp.Send(mail1)

        Me.Timer1.Enabled = True
        Me.Timer2.Enabled = True
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Dim vPDFName As String
        'Dim vFileName As String
        'Dim vSubject As String
        'Dim vBody As String
        'Dim vEmailAll As String
        'Dim da1 As SqlDataAdapter
        'Dim ds1 As DataSet
        'Dim dt1 As New DataTable

        'Dim i As Integer
        'Dim vGetDocNo As String

        'On Error Resume Next

        'Dim vFileName1 As String

        'Dim mail1 As New MailMessage()
        'Dim smtp As New SmtpClient("192.168.0.169")

        ''Me.Timer1.Enabled = False

        'vEmailAll = "it@nopadol.com" & "," & "somrod@smartworks.in.th"


        'vFileName = Trim("E:\Picture\me.jpg")
        'vFileName1 = Trim("E:\Picture\mee.jpg")
        'vSubject = "test"
        'vBody = "เป็นรายงาน สินค้าที่เสนอซื้อที่ทางจัดซื้อได้พิจารณาแล้ว ของทีมเพื่อให้ทางผู้อำนวยการจัดซื้อพิจารณาตรวจสอบ"


        'Dim mail As New MailMessage()
        'mail.From = New MailAddress("it@nopadol.com")

        'Dim strEmails() As String = vEmailAll.Split(",")
        'For Each str As String In strEmails
        '    If str <> "" Then
        '        mail.To.Add(str)
        '    End If
        'Next

        'mail.Subject = vSubject
        'mail.Body = vBody

        'Dim att As New Attachment(vFileName)
        'Dim att1 As New Attachment(vFileName1)

        'mail.Attachments.Add(att1)
        'mail.Attachments.Add(att)

        'smtp.Send(mail)

        'Dim FILE_NAME As String = "C:\PDFDocApp\SendEmailLogs.txt"
        'Dim errText As String

        'errText = Now & " " & "เจ้าหนี้รหัส" & " " & "เลขที่ใบสั่งซื้อ" & " " & "โอนไม่ผ่าน มีปัญหา"

        'Dim objWriter As New System.IO.StreamWriter(FILE_NAME, True)

        'objWriter.WriteLine(errText)

        'objWriter.Close()
    End Sub

    Private Sub Timer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer3.Tick
        Me.LBLCheckTime.Text = Now

        Dim vCheckTime As Date
        Dim vTime As Date
        Dim vDiffTime As Integer

        On Error Resume Next

        vTime = Me.LBLTime.Text
        vCheckTime = Me.LBLCheckTime.Text
        vDiffTime = vb6.DateDiff(DateInterval.Minute, vTime, vCheckTime)

        'MsgBox(vTime & "   " & vCheckTime & "   " & vDiffTime)

        If vDiffTime >= 3 Then
            Me.Timer1.Enabled = True
            Me.Timer2.Enabled = True
        End If

    End Sub

    Private Sub Timer4_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer4.Tick
        Dim i As Integer
        Dim vGetReportID As String
        Dim vGetReportType As Integer
        Dim vGetProfitCenter As String
        Dim vGetExpertTeam As String
        Dim vGetSection As String
        Dim vGetDepartCode As String
        Dim vGetApCode As String
        Dim vGetDocNo As String
        Dim vGetEmail As String
        Dim vGetCC As String
        Dim vGetFromMail As String
        Dim vGetReportName As String
        Dim vGetPrintDateTime As String
        Dim vGetEmailAll As String

        Dim vListItem As ListViewItem

        On Error Resume Next

        Me.ListViewSendMail.Items.Clear()
        vQuery = "exec dbo.USP_PR_SearchCheckPrintReOrderAuto_Cat 1"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "Docno")
        dt = ds.Tables("Docno")
        If dt.Rows.Count > 0 Then

            vCountPO = dt.Rows(i).Item("vCount")

            For i = 0 To dt.Rows.Count - 1
                vGetReportID = dt.Rows(i).Item("reportid")
                If vGetReportID = "001" Then
                    vGetReportType = 0
                ElseIf vGetReportID = "002" Then
                    vGetReportType = 1
                ElseIf vGetReportID = "004" Then
                    vGetReportType = 1
                ElseIf vGetReportID = "003" Then
                    vGetReportType = 1
                Else
                    vReportName = ""
                End If
                vGetProfitCenter = dt.Rows(i).Item("profitcenter")
                vGetExpertTeam = dt.Rows(i).Item("expertteam")
                vGetSection = dt.Rows(i).Item("sectionmanager")
                vGetDepartCode = dt.Rows(i).Item("department")
                vGetApCode = dt.Rows(i).Item("apcode")
                vGetDocNo = dt.Rows(i).Item("docno")
                vGetReportName = dt.Rows(i).Item("reportname")

                vGetEmail = dt.Rows(i).Item("email")
                vGetCC = dt.Rows(i).Item("cc")
                vGetPrintDateTime = dt.Rows(i).Item("printdatetime")

                If vGetReportID <> "" Then
                    vListItem = Me.ListViewSendMail.Items.Add(vGetPrintDateTime)
                    vListItem.SubItems.Add(0).Text = vGetReportName
                    vListItem.SubItems.Add(1).Text = vGetProfitCenter
                    vListItem.SubItems.Add(2).Text = vGetExpertTeam
                    vListItem.SubItems.Add(3).Text = vGetDepartCode
                    vListItem.SubItems.Add(4).Text = vGetSection
                    vListItem.SubItems.Add(5).Text = vGetApCode
                    vListItem.SubItems.Add(6).Text = vGetDocNo
                    vListItem.SubItems.Add(7).Text = vGetEmail
                    vListItem.SubItems.Add(8).Text = vGetCC
                End If

            Next
        End If

        If vCountPO > 0 Then
            Me.Timer1.Enabled = False
            Me.Timer4.Enabled = True
        End If

        If vCountPO = 0 Then
            Me.Timer1.Enabled = True
            Me.Timer4.Enabled = False
        End If

        vQuery = "exec dbo.USP_PR_SearchCheckPrintReOrderAuto_Cat 0"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "Docno")
        dt = ds.Tables("Docno")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                vGetReportID = dt.Rows(0).Item("reportid")
                If vGetReportID = "001" Then
                    vGetReportType = 2
                ElseIf vGetReportID = "011" Then
                    vGetReportType = 2
                ElseIf vGetReportID = "002" Then
                    vGetReportType = 1
                ElseIf vGetReportID = "012" Then
                    vGetReportType = 1
                ElseIf vGetReportID = "013" Then
                    vGetReportType = 1
                ElseIf vGetReportID = "014" Then
                    vGetReportType = 1
                ElseIf vGetReportID = "004" Then
                    vGetReportType = 1
                ElseIf vGetReportID = "005" Then 'CAT =< 100000
                    vGetReportType = 0
                ElseIf vGetReportID = "006" Then 'ALL=ทั้หมดและMG > 100000
                    vGetReportType = 1
                ElseIf vGetReportID = "999" Then 'ส่งเมลล์ปรับราคาให้น้ำหวานศูนย์สีไอซีไอ
                    vGetReportType = 1
                End If
                vGetProfitCenter = dt.Rows(0).Item("profitcenter")
                vGetExpertTeam = dt.Rows(0).Item("expertteam")
                vGetSection = dt.Rows(0).Item("sectionmanager")
                vGetDepartCode = dt.Rows(0).Item("department")
                vGetApCode = dt.Rows(0).Item("apcode")
                vGetDocNo = dt.Rows(0).Item("docno")

                vGetEmail = dt.Rows(0).Item("email")
                vGetCC = dt.Rows(0).Item("cc")
                vGetFromMail = dt.Rows(0).Item("fromemail")
                vGetEmailAll = vGetEmail & "," & vGetCC

                If vGetReportID = "001" Or vGetReportID = "002" Or vGetReportID = "004" Or vGetReportID = "011" Or vGetReportID = "012" Then

                    vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vGetReportID & "','" & vGetExpertTeam & "','" & vGetProfitCenter & "','" & vGetDepartCode & "','" & vGetSection & "','" & vGetApCode & "','" & vGetDocNo & "','" & vGetEmailAll & "'"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Call SendMailReport_Gmail(vGetReportID, vGetReportType, vGetProfitCenter, vGetExpertTeam, vGetSection, vGetDepartCode, vGetApCode, vGetDocNo, vGetEmail, vGetCC, vGetFromMail)

                ElseIf vGetReportID = "013" Or vGetReportID = "014" Then

                    vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vGetReportID & "','" & vGetExpertTeam & "','" & vGetProfitCenter & "','" & vGetDepartCode & "','" & vGetSection & "','" & vGetApCode & "','" & vGetDocNo & "','" & vGetEmailAll & "'"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Call SendMailItemTransfer_Gmail(vGetReportID, vGetReportType, vGetProfitCenter, vGetExpertTeam, vGetSection, vGetDepartCode, vGetApCode, vGetDocNo, vGetEmail, vGetCC, vGetFromMail)

                ElseIf vGetReportID = "003" Then

                    Call SendMailPO_Gmail(vGetApCode, vGetDocNo, vGetEmail, vGetCC, vGetFromMail)

                ElseIf vGetReportID = "005" Then

                    vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vGetReportID & "','','','','','','','" & vGetEmailAll & "'"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Call SendMailPOApproveCAT_Gmail(vGetReportID, "CAT1", vGetEmail, vGetCC, vGetFromMail)

                ElseIf vGetReportID = "006" Then

                    vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vGetReportID & "','','','','','','','" & vGetEmailAll & "'"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Call SendMailPOApproveCAT_Gmail(vGetReportID, "CAT2", vGetEmail, vGetCC, vGetFromMail)

                ElseIf vGetReportID = "007" Then

                    vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vGetReportID & "','','','','','','','" & vGetEmailAll & "'"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Call SendMailPOApproveCAT_Gmail(vGetReportID, "CAT3", vGetEmail, vGetCC, vGetFromMail)

                ElseIf vGetReportID = "008" Then

                    vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vGetReportID & "','','','','','','','" & vGetEmailAll & "'"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Call SendMailPOApproveMG_Gmail(vGetEmail, vGetCC, vGetFromMail)

                ElseIf vGetReportID = "999" Then

                    vQuery = "exec dbo.USP_PR_UpdatePrintReOrderAuto '" & vGetReportID & "','','','','','','','" & vGetEmailAll & "'"
                    cmd = New SqlCommand(vQuery, vConnection)
                    cmd.ExecuteNonQuery()

                    Call SendMailChangePrice_Gmail(vGetReportID, vGetEmail, vGetCC, vGetFromMail)
                End If
            Next
        Else
            vGetReportID = ""
            vGetReportType = 0
            vGetProfitCenter = ""
            vGetExpertTeam = ""
            vGetSection = ""
            vGetDepartCode = ""
            vGetApCode = ""
            vGetDocNo = ""

            vGetEmail = ""
            vGetCC = ""
        End If

        'MsgBox("timer4")
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Call SendMailPOLetter_Gmail()
    End Sub

    Public Sub SendMailPOLetter_Gmail()
        Dim myReport3 As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim i As Integer
        Dim vGetApName As String
        Dim vGetContactName As String
        Dim vFileName As String
        Dim vSubject As String
        Dim vBody As String
        Dim vEmailAll As String
        Dim vFromMail As String
        Dim vMailID As Integer

        On Error Resume Next

        Dim myTables As CrystalDecisions.CrystalReports.Engine.Tables
        Dim myTableLogonInfo As CrystalDecisions.Shared.TableLogOnInfo
        Dim myConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo

        Dim lFileName As String
        Dim myReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument

        Dim mail1 As New MailMessage()
        Dim smtp As New SmtpClient("smtp.gmail.com")

        smtp.Port = 587
        smtp.EnableSsl = True
        smtp.UseDefaultCredentials = False
        smtp.Credentials = New System.Net.NetworkCredential("nopadol_mailauto@nopadol.com", "[vdw,jwfh2012")

        Me.Timer1.Enabled = False
        Me.Timer2.Enabled = False

        vQuery = "SELECT top 1 MailID,isnull(ContactName,'') as ContactName,isnull(Name1,'') as Name1,isnull(APCode,'') as APCode,isnull(Email,'') as Email,isnull(CC,'') as CC,isnull(FromEmail,'') as FromEmail FROM npmaster.dbo.TB_NP_EmailMaster WHERE ActiveStatus = 1 AND TypeID = 1 and isnull(brandcode,'') = ''  order by isnull(expertteam,'')"
        da1 = New SqlDataAdapter(vQuery, vConnection)
        ds1 = New DataSet
        da1.Fill(ds1, "Docno1")
        dt1 = ds1.Tables("Docno1")
        If dt1.Rows.Count > 0 Then
            vMailID = dt1.Rows(i).Item("MailID")
            vEmailAll = dt1.Rows(i).Item("Email") & "," & dt1.Rows(i).Item("CC")
            vFromMail = dt1.Rows(i).Item("FromEmail")
            vGetApName = dt1.Rows(i).Item("Name1")
            vGetContactName = dt1.Rows(i).Item("contactname")

            vFileName = Trim("C:\PDFDocApp\Letter_PO.JPG")
            vSubject = Trim("จดหมาย ขอสนันสนุนงานปีใหม่ ประจำปี 2557 ของ บริษัทนพดลพานิช จำกัด")
            vBody = "เรียนผู้แทนจำหน่ายทุกท่าน ทางบริษัทนพดลพานิช ได้แนบจดหมาย ของบสนันสนุนงานปีใหม่ ประจำปี 2557 ให้กับทางผู้แทนจำหน่าย"


            vBody = vBody & "<br>" & " " & "</br>" & "ติดต่อสอบถามรายละเอียดได้กับทางแผนกจัดซื้อ ของบริษัทนพดลพานิช จำกัด" & "</br>"


            vBody = vBody & "<br>" & " " & "</br>" & "ขอแสดงความขอบคุณล่วงหน้า มา ณ โอกาสนี้" & "</br>"

            vBody = vBody & "<br>" & " " & "</br>" & "แผนกจัดซื้อ บริษัทนพดลพานิช จำกัด " & "</br>"


            lFileName = vFileName

            mail1.From = New MailAddress(vFromMail)

            Dim strEmails1() As String = vEmailAll.Split(",")
            For Each str As String In strEmails1
                If str <> "" Then
                    mail1.To.Add(str)
                End If
            Next

            mail1.IsBodyHtml = True
            mail1.Subject = vSubject

            mail1.Body = vBody
            Dim att As New Attachment(lFileName)
            mail1.Attachments.Add(att)

            smtp.Send(mail1)

            vQuery = "update npmaster.dbo.TB_NP_EmailMaster set brandcode = 'send' where mailid = " & vMailID & ""
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

        End If


ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description)
            Exit Sub
        End If

    End Sub


    Public Sub SendMailPaybillCond_Gmail(ByVal vApCode As String, ByVal vEmail As String, ByVal vGetContactName As String, ByVal vGetNopadolContactName As String)
        Dim myReport3 As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim i As Integer
        Dim vGetDocNo As String
        Dim vPDFName As String
        Dim vFileName As String
        Dim vSubject As String
        Dim vBody As String
        Dim rptDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        Dim vEmailAll As String
        Dim vReportID As String
        Dim vMydescription As String
        Dim vPicture1 As String
        Dim vPicture2 As String
        Dim vPicture3 As String

        Dim vLink As String
        Dim vDocDate As String
        Dim vLeadDate As String
        Dim vSendDate As String
        Dim vUserID As String
        Dim vCatCode As String
        Dim vFromMail As String

        On Error GoTo ErrDescription

        Dim myTables As CrystalDecisions.CrystalReports.Engine.Tables
        Dim myTableLogonInfo As CrystalDecisions.Shared.TableLogOnInfo
        Dim myConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo

        Dim lFileName As String
        Dim myReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument

        Dim mail1 As New MailMessage()
        Dim smtp As New SmtpClient("smtp.gmail.com")

        smtp.Port = 587
        smtp.EnableSsl = True
        smtp.UseDefaultCredentials = False
        smtp.Credentials = New System.Net.NetworkCredential("nopadol_mailauto@nopadol.com", "[vdw,jwfh2012")

        Me.Timer1.Enabled = False
        Me.Timer2.Enabled = False

        If vEmail <> "" Then
            vEmailAll = vEmail & "," & "it@nopadol.com,ch@nopadol.com,vilaivan@nopadol.com"
        Else
            vEmailAll = "it@nopadol.com,ch@nopadol.com,vilaivan@nopadol.com"

        End If

        vReportID = "SRF"
        vFromMail = "it@nopadol.com,ch@nopadol.com,vilaivan@nopadol.com"

        vReportName = "RP_NP_VendorSendEmail"
        vReportPath = Trim("V:\Reports\" & vReportName & ".rpt")

        vPDFName = Trim(vApCode)
        vFileName = Trim("C:\PDFDocApp\" & vPDFName & ".pdf")
        vSubject = Trim("บริษัท นพดลพานิช จำกัด จัดส่งเอกสารอีเลคโทรนิคส์ แจ้งเรื่องระบบรับวางบิลและชำระเงิน")
        vBody = "เรียน ผู้จัดการ  และ   " & vGetContactName & "    ทางบริษัท นพดลพานิช จำกัด ได้ปรับปรุงระบบการรับวางบิลและชำระเงินให้กับทางเจ้าหนี้ ตามรายละเอียดที่แนบ"

        myReport3.Load(vReportPath)

        myConnectionInfo.ServerName = "Nebula"
        myConnectionInfo.DatabaseName = "BCNP"
        myConnectionInfo.UserID = "sa"
        myConnectionInfo.Password = "[ibdkifu"

        myTables = myReport3.Database.Tables

        For Each myTable As CrystalDecisions.CrystalReports.Engine.Table In myTables
            myTableLogonInfo = myTable.LogOnInfo
            myTableLogonInfo.ConnectionInfo = myConnectionInfo
            myTable.ApplyLogOnInfo(myTableLogonInfo)
        Next

        lFileName = vFileName
        myReport3.SetParameterValue("@vAPCode", vApCode)

        CrystalReportViewer2.ReportSource = myReport3

        myReport3.ExportToDisk(ExportFormatType.PortableDocFormat, lFileName)

        lFileName = vFileName

        mail1.From = New MailAddress(vFromMail)

        Dim strEmails1() As String = vEmailAll.Split(",")
        For Each str As String In strEmails1
            If str <> "" Then
                mail1.To.Add(str)
            End If
        Next

        mail1.IsBodyHtml = True

        mail1.Subject = vSubject
        vBody = vBody & "<br>" & " ******************************************** หากมีข้อสงสัย กรุณาติดต่อกลับ ******************************************** " & "</br>" & "<br>" & "<font size=5 color=blue bold = 5>" & "หากท่านมีข้อสงสัย กรุณาติดต่อคุณ กัญจิรา ภูมิทอง หัวหน้าแผนกการเงิน ได้ทางอีเมลล์ ch@nopadol.com หรือ เบอร์โทร 053-240377 ต่อ 515" & "</font>" & "</br>"

        mail1.Body = vBody
        Dim att As New Attachment(lFileName)
        mail1.Attachments.Add(att)

        smtp.Send(mail1)

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description & " " & vApCode & " " & "ส่งไม่ผ่าน")
            Application.Exit()
        End If

    End Sub

    Private Sub Timer6_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer6.Tick
        Dim vAPCode As String
        Dim vEmail As String
        Dim vContract As String
        Dim vNPContract As String

        On Error GoTo ErrDescription

        vQuery = "exec dbo.USP_AP_SearchSendPaybillMail"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "Docno")
        dt = ds.Tables("Docno")
        If dt.Rows.Count > 0 Then
            vAPCode = dt.Rows(0).Item("apcode")
            vEmail = dt.Rows(0).Item("email")
            vContract = dt.Rows(0).Item("contactname")

            vNPContract = ""
            Call SendMailPaybillCond_Gmail(vAPCode, vEmail, vContract, vNPContract)
        End If

        vQuery = "update npmaster.dbo.tb_np_emailmaster set issendmail = 1 where apcode = '" & vAPCode & "' "
        cmd = New SqlCommand(vQuery, vConnection)
        cmd.ExecuteNonQuery()

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description & " " & vAPCode & " " & "ส่งไม่ผ่าน")
            Application.Exit()
        End If
    End Sub
End Class
