Imports System.data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization
Imports System.Windows.Forms
Imports System.Web.SessionState
Imports System.Net
Imports System.Web

Partial Class PGCheckOut
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim dt1 As DataTable = Me.GetDataSourceVisible()
        Dim dr1 As DataRow = dt1.NewRow

        dr1("����") = ""
        dr1("����") = ""
        dr1("�ӹǹ") = ""
        dr1("˹���") = ""
        dr1("�Ҥ�") = ""
        dt1.Rows.Add(dr1)
        Me.BindGrid1(dt1)
    End Sub

    Private Function GetDataSourceVisible() As DataTable
        Dim dt1 As DataTable = Session("MyDataSource1")
        If dt1 Is Nothing Then
            dt1 = New DataTable()
            dt1.Columns.Add("����")
            dt1.Columns.Add("����")
            dt1.Columns.Add("�ӹǹ")
            dt1.Columns.Add("˹���")
            dt1.Columns.Add("�Ҥ�")
            Session("MyDataSource1") = dt1
        End If
        Return dt1
    End Function

    Private Function ClearSession()
        Me.GridView1.DataBind()
        'Me.GridView2.DataBind()
        'Me.LBLNetAmount.Text = ""
        Session("MyDataSource1") = Nothing
        Session("MyDataSource2") = Nothing
        Session.RemoveAll()
        Session.Remove("MyDataSource1")
        Session.Remove("MyDataSource2")
    End Function

    Private Function GetDataSourceHide() As DataTable
        Dim dt2 As DataTable = Session("MyDataSource2")
        If dt2 Is Nothing Then
            dt2 = New DataTable()
            dt2.Columns.Add("����")
            dt2.Columns.Add("����")
            dt2.Columns.Add("�ӹǹ")
            dt2.Columns.Add("˹���")
            dt2.Columns.Add("�Ҥ�")
            dt2.Columns.Add("���")
            dt2.Columns.Add("��ѧ")
            dt2.Columns.Add("�����")
            dt2.Columns.Add("����")
            dt2.Columns.Add("�����")
            dt2.Columns.Add("⫹")
            dt2.Columns.Add("�ش����")

            Session("MyDataSource2") = dt2
        End If
        Return dt2
    End Function

    Private Sub BindGrid1(ByVal dt1 As DataTable)
        'GridView2.DataSource = dt1
        'GridView2.DataBind()
    End Sub

    Private Sub BindGrid2(ByVal dt2 As DataTable)
        GridView1.DataSource = dt2
        GridView1.DataBind()
    End Sub
End Class
