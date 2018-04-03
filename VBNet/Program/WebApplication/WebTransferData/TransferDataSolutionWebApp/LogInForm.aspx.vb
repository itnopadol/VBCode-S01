Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.SqlServer
Imports Microsoft.VisualBasic
Imports ASP
Partial Class _Default
    Inherits System.Web.UI.Page
    Dim vConnectionString As String
    Dim vConnection As SqlConnection
    Dim da As SqlDataAdapter
    Dim ds As DataSet
    Dim dt As DataTable
    Dim vQuery As String
    Dim vCommand As SqlCommand

    Dim vServerName As String
    Dim vDatabaseName As String
    Dim vUserID As String
    Dim vPassword As String
    Dim vCheckConnect As Integer

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            Dim vScript As String = "<SCRIPT language='javascript'>form1.TBUserID.focus();</SCRIPT>"
            Page.RegisterStartupScript("focus", vScript)
        End If

    End Sub

    Protected Sub BTNLogIn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BTNLogIn.Click

        If Me.TBServer.Text = "" Then
            MsgBox("กรุณากรอก ชื่อ Server", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBServer.Focus()
            Exit Sub
        End If
        If Me.TBDatabase.Text = "" Then
            MsgBox("กรุณากรอก ชื่อ Database", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBDatabase.Focus()
            Exit Sub
        End If
        If Me.TBUserID.Text = "" Then
            MsgBox("กรุณากรอก ชื่อ UserID", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBUserID.Focus()
            Exit Sub
        End If


        vServerName = Me.TBServer.Text
        vDatabaseName = Me.TBDatabase.Text
        vUserID = Me.TBUserID.Text
        vPassword = Me.TBPassword.Text


        On Error GoTo ErrDescription
        vConnectionString = "Persist Security Info =False;User ID =" & vUserID & ";Password =" & vPassword & ";Data Source =" & vServerName & ";Initial Catalog =" & vDatabaseName & ""
        vConnection = New SqlConnection(vConnectionString)
        vConnection.Open()

        vCheckConnect = 1


ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            vCheckConnect = 0
            Me.TBPassword.Text = ""
            Me.TBUserID.Focus()
        End If


        If vCheckConnect = 1 Then

            Dim vServer As String = Me.TBServer.Text
            Dim vDataBase As String = Me.TBDatabase.Text
            Dim vUser As String = "sa"
            Dim vPassword As String = "[ibdkifu"
            Dim vTrnUserID As String = Me.TBUserID.Text

            Response.Redirect("TransferData.aspx?Server=" & vServer & "&DataBase=" & vDataBase & "&User=" & vUser & "&Password=" & vPassword & "&TrnUserID=" & vTrnUserID)
        End If
    End Sub

    Protected Sub TBPassword_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TBPassword.TextChanged

    End Sub
End Class
