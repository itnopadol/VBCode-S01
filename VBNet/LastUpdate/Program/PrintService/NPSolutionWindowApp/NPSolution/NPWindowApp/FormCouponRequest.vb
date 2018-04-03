Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic

Public Class FormCouponRequest
    Dim vQuery As String
    Dim ds As DataSet
    Dim da As SqlDataAdapter
    Dim dt As DataTable
    Dim vCMD As SqlCommand
    Dim vReadQuery As SqlDataReader

    Dim vCheckIndex As Integer

    Private Sub TextCPCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextCPCode.KeyDown
        Dim vCPCode As String
        Dim i As Integer
        Dim n As Integer
        Dim vListCoupon As ListViewItem

        If e.KeyCode = Keys.Enter Then
            vCPCode = Me.TextCPCode.Text
            vQuery = "exec dbo.USP_NP_CouponData '" & vCPCode & "'"
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "CPData")
            dt = ds.Tables("CPData")

            If dt.Rows.Count > 0 Then
                Me.TextCPName.Text = Trim(dt.Rows(i).Item("cpname"))

                Me.ListViewReqCP.Items.Clear()
                For i = 0 To dt.Rows.Count - 1
                    vListCoupon = Me.ListViewReqCP.Items.Add(dt.Rows(i).Item("cpheader"))
                    vListCoupon.SubItems.Add(1).Text = Format(dt.Rows(i).Item("cpvalue"), "##,##0.00")
                    vListCoupon.SubItems.Add(2).Text = Format(dt.Rows(i).Item("cpqty"), "##,##0.00")
                    vListCoupon.SubItems.Add(3).Text = Format(dt.Rows(i).Item("cpapprove"), "##,##0.00")
                    vListCoupon.SubItems.Add(4).Text = Format(dt.Rows(i).Item("cpremain"), "##,##0.00")
                    vListCoupon.SubItems.Add(5).Text = Format(0, "##,##0.00")

                    For n = 0 To ListViewReqCP.Items.Count - 1
                        If ListViewReqCP.Items.Count Mod 2 = 0 Then
                            ListViewReqCP.Items(ListViewReqCP.Items.Count - 1).BackColor = Color.PowderBlue
                        End If
                    Next
                Next
            End If
        End If
    End Sub

    Private Sub TextCPCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextCPCode.TextChanged

    End Sub

    Private Sub FormCouponRequest_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call InitializeDataBase()
    End Sub

    Private Sub TextDocNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextDocNo.TextChanged

    End Sub

    Private Sub ListViewReqCP_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles ListViewReqCP.ItemCheck
        If Me.ListViewReqCP.Items.Count > 0 Then
            vCheckIndex = e.Index
            If Me.ListViewReqCP.Items.Item(vCheckIndex).Checked = False Then
                Me.GB101.Visible = True
                Me.TextQTY.Focus()
            Else
                Me.ListViewReqCP.Items.Item(vCheckIndex).SubItems(5).Text = Format(0, "##,##0.00")
            End If
        End If
    End Sub

    Private Sub TextQTY_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextQTY.KeyDown
        Dim vIndex As Integer
        Dim vQTY As Integer

        If e.KeyCode = Keys.Enter Then
            If Me.TextQTY.Text <> 0 Then
                vIndex = vCheckIndex
                vQTY = Me.TextQTY.Text
                Me.ListViewReqCP.Items(vIndex).SubItems(5).Text = Format(vQTY, "##,##0.00")
                Me.TextQTY.Text = ""
                Me.GB101.Visible = False
            End If
        End If
    End Sub

    Private Sub TextQTY_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextQTY.KeyPress
        On Error GoTo ErrDescription

        Select Case Asc(e.KeyChar)
            Case 48 To 58, 8, 44, 46, 37
            Case Else
                e.Handled = True
        End Select

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error")
            Exit Sub
        End If
    End Sub

    Private Sub TextQTY_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextQTY.TextChanged

    End Sub

    Private Sub ListViewReqCP_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewReqCP.SelectedIndexChanged

    End Sub
End Class