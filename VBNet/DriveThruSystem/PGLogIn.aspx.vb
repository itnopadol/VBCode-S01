Imports System.data
Imports System.Data.SqlClient
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization
Imports System.Windows.Forms
Imports System.Web.SessionState
Imports System.Net
Imports System.Web

Partial Class PGLogIn
    Inherits System.Web.UI.Page

    Dim vQuery As String
    Dim vPointID As String
    Dim vUserID As String
    Dim vMemSaleName As String
    Dim vUserName As String
    Dim vDuty As String
    Dim vLevelID As Integer

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Call AddPointID()

        If Not Page.IsPostBack Then

            Dim vScript As String = "<SCRIPT language='javascript'>form1.TBUserID.focus();</SCRIPT>"
            Page.RegisterStartupScript("focus", vScript)

        End If

        'Me.TBUserID.Attributes.Add("onblur", Me.ClientScript.GetPostBackEventReference(Me.TBUserID, String.Empty))
        'Me.TBUserID.Attributes.Add("onkeydown", "fnTrapKD(" + TBUserID.ClientID + ",event)")
    End Sub

    Public Sub AddPointID()
        Me.DDLPoint.Items.Add("จุดที่ 1")
        Me.DDLPoint.Items.Add("จุดที่ 2")
        Me.DDLPoint.Items.Add("จุดที่ 3")
    End Sub

    Protected Sub TBUserID_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TBUserID.TextChanged
        Dim vCheckUserID As String

        If Me.TBUserID.Text <> "" Then
            vCheckUserID = Me.TBUserID.Text

            vQuery = "select * from npmaster.dbo.tb_np_bcuserid where comment = '" & vCheckUserID & "'"
            Dim vService1 As New WebReference.WebServiceCalc
            Dim ds1 As DataSet = vService1.vGetQueryAnlyzer(vQuery)

            If ds1.Tables(0).Rows.Count > 0 Then
                vUserID = ds1.Tables(0).Rows(0)("code").ToString
                vMemSaleName = ds1.Tables(0).Rows(0)("comment").ToString & "/" & ds1.Tables(0).Rows(0)("name").ToString
                vLevelID = 0
            Else
                vUserID = ""
                vLevelID = 0
                vMemSaleName = ""
            End If

            If Me.DDLPoint.Text = "จุดที่ 1" Then
                vPointID = 1
            End If
            If Me.DDLPoint.Text = "จุดที่ 2" Then
                vPointID = 2
            End If
            If Me.DDLPoint.Text = "จุดที่ 3" Then
                vPointID = 3
            End If

            If vUserID <> "" Then
                Me.LBLMessage.Text = ""
                Response.Redirect("PickupLicense.aspx?UserID=" & vUserID & "&SaleCode=" & vMemSaleName & "&PointID=" & vPointID)
            Else
                Me.LBLMessage.Text = "ไม่สามารถเข้าใช้งานได้ กรุณาตรวจสอบรหัสเข้าใช้งาน"
            End If
        End If
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim vCheckUserID As String

        If Me.TBUserID.Text <> "" Then
            vCheckUserID = Me.TBUserID.Text

            vQuery = "select * from npmaster.dbo.tb_np_bcuserid where comment = '" & vCheckUserID & "'"
            Dim vService1 As New WebReference.WebServiceCalc
            Dim ds1 As DataSet = vService1.vGetQueryAnlyzer(vQuery)

            If ds1.Tables(0).Rows.Count > 0 Then
                vUserID = ds1.Tables(0).Rows(0)("code").ToString
                vMemSaleName = ds1.Tables(0).Rows(0)("comment").ToString & "/" & ds1.Tables(0).Rows(0)("name").ToString
                vLevelID = 0
            Else
                vUserID = ""
                vLevelID = 0
                vMemSaleName = ""
            End If

            If Me.DDLPoint.Text = "จุดที่ 1" Then
                vPointID = 1
            End If
            If Me.DDLPoint.Text = "จุดที่ 2" Then
                vPointID = 2
            End If
            If Me.DDLPoint.Text = "จุดที่ 3" Then
                vPointID = 3
            End If

            If vUserID <> "" Then
                Me.LBLMessage.Text = ""
                Response.Redirect("PickupLicense.aspx?UserID=" & vUserID & "&SaleCode=" & vMemSaleName & "&PointID=" & vPointID)
            Else
                Me.LBLMessage.Text = "ไม่สามารถเข้าใช้งานได้ กรุณาตรวจสอบรหัสเข้าใช้งาน"
            End If
        End If
    End Sub
End Class
