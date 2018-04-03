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

        dr1("รหัส") = ""
        dr1("ชื่อ") = ""
        dr1("จำนวน") = ""
        dr1("หน่วย") = ""
        dr1("ราคา") = ""
        dt1.Rows.Add(dr1)
        Me.BindGrid1(dt1)
    End Sub

    Private Function GetDataSourceVisible() As DataTable
        Dim dt1 As DataTable = Session("MyDataSource1")
        If dt1 Is Nothing Then
            dt1 = New DataTable()
            dt1.Columns.Add("รหัส")
            dt1.Columns.Add("ชื่อ")
            dt1.Columns.Add("จำนวน")
            dt1.Columns.Add("หน่วย")
            dt1.Columns.Add("ราคา")
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
            dt2.Columns.Add("รหัส")
            dt2.Columns.Add("ชื่อ")
            dt2.Columns.Add("จำนวน")
            dt2.Columns.Add("หน่วย")
            dt2.Columns.Add("ราคา")
            dt2.Columns.Add("รวม")
            dt2.Columns.Add("คลัง")
            dt2.Columns.Add("ชั้นเก็บ")
            dt2.Columns.Add("บาร์")
            dt2.Columns.Add("ที่เก็บ")
            dt2.Columns.Add("โซน")
            dt2.Columns.Add("จุดจ่าย")

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
