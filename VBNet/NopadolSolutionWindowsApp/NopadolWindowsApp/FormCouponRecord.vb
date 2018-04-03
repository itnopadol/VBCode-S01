Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic

Public Class FormCouponRecord

    Dim vQuery As String
    Dim ds As DataSet
    Dim da As SqlDataAdapter
    Dim dt As DataTable
    Dim vCMD As SqlCommand
    Dim vReadQuery As SqlDataReader

    Private Sub BTNBasket_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNBasket.Click
        Dim vListCoupon As ListViewItem
        Dim vCouponHeader As String
        Dim vCouponValue As Double
        Dim vCouponQTY As Integer
        Dim i As Integer

        If Me.TextCouponCode.Text <> "" And Me.TextCouponName.Text <> "" And Me.TextCouponNo.Text <> "" And Me.TextCouponFormat.Text <> "" And Me.TextCountLenght.Text <> "" And Me.TextCouponQTY.Text <> "" And Me.TextCouponValue.Text <> "" Then
            vCouponHeader = Me.TextCouponNo.Text
            vCouponValue = Me.TextCouponValue.Text
            vCouponQTY = Me.TextCouponQTY.Text
            vListCoupon = Me.ListViewCoupon.Items.Add(vCouponHeader)
            vListCoupon.SubItems.Add(1).Text = vCouponValue
            vListCoupon.SubItems.Add(2).Text = vCouponQTY
            vListCoupon.SubItems.Add(3).Text = 0
            vListCoupon.SubItems.Add(4).Text = vCouponQTY

            For i = 0 To ListViewCoupon.Items.Count - 1
                If ListViewCoupon.Items.Count Mod 2 = 0 Then ListViewCoupon.Items(ListViewCoupon.Items.Count - 1).BackColor = Color.PowderBlue
            Next

            'For i = 0 To ListViewCoupon.Items.Count - 1
            'สลับแถวเลขคู่กับเลขคี่
            'If i Mod 2 = 0 Then
            '    ListViewCoupon.Items(i).BackColor = Color.Blue
            'Else
            '    ListViewCoupon.Items(i).BackColor = Color.Red
            'End If
            'Next

            Me.TextCouponNo.Text = ""
            Me.TextCouponValue.Text = ""
            Me.TextCouponQTY.Text = ""
            Me.TextCouponNo.Focus()
        Else
            MsgBox("ต้องกรอกข้อมูลให้ครบ ตามที่โปรแกรมต้องการ ถึงจะลงตารางข้างล่างได้", MsgBoxStyle.Critical, "Send Error Message")
        End If
    End Sub

    Private Sub RB101_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.TextCouponFormat.Text = 0
    End Sub

    Private Sub RB102_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.TextCouponFormat.Text = 1
    End Sub

    Private Sub RB103_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.TextCouponFormat.Text = 2
    End Sub

    Private Sub RadioButton1_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.TextCouponFormat.Text = 3
    End Sub


    Private Sub BTNSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSave.Click
        Dim vCPCode As String
        Dim vCPName As String
        Dim vCPMerge As Integer
        Dim vFromDate As String
        Dim vToDate As String
        Dim vFormatCode As Integer
        Dim vLenght As Integer
        Dim vMyDescription As String

        Dim vCPHeader As String
        Dim vCPQTY As Integer
        Dim vCPApprove As Integer
        Dim vCPRemain As Integer
        Dim vCPValue As Double
        Dim vLineNumber As Integer

        Dim i As Integer

        Try
            vQuery = "begin tran"
            vCMD = New SqlCommand(vQuery, vConnection)
            vCMD.ExecuteNonQuery()

            vCPCode = Me.TextCouponCode.Text
            vCPName = me.TextCouponName.Text  
            vFromDate = Me.DTPStartDate.Text
            vToDate = Me.DTPStopDate.Text
            If Me.CBMerge.Checked = False Then
                vCPMerge = 0
            Else
                vCPMerge = 1
            End If
            vFormatCode = Me.TextCouponFormat.Text
            vLenght = Me.TextCountLenght.Text
            vMyDescription = Me.TextMyDescription.Text

            vQuery = "exec dbo.USP_NP_InsertCouponMaster '" & vCPCode & "','" & vCPName & "','" & vFromDate & "','" & vToDate & "'," & vCPMerge & "," & vFormatCode & "," & vLenght & ",'" & vMyDescription & "','" & vUserID & "'"
            vCMD = New SqlCommand(vQuery, vConnection)
            vCMD.ExecuteNonQuery()

            For i = 0 To Me.ListViewCoupon.Items.Count - 1
                vCPHeader = Me.ListViewCoupon.Items(i).SubItems(0).Text
                vCPValue = Me.ListViewCoupon.Items(i).SubItems(1).Text
                vCPQTY = Me.ListViewCoupon.Items(i).SubItems(2).Text
                vCPApprove = Me.ListViewCoupon.Items(i).SubItems(3).Text
                vCPRemain = Me.ListViewCoupon.Items(i).SubItems(4).Text
                vLineNumber = i

                vQuery = "exec dbo.USP_NP_InsertCouponDetails '" & vCPCode & "','" & vCPHeader & "'," & vFormatCode & "," & vCPValue & "," & vCPQTY & "," & vcpapprove & "," & vcpremain & "," & vLineNumber & ""
                vCMD = New SqlCommand(vQuery, vConnection)
                vCMD.ExecuteNonQuery()
            Next

            vQuery = "commit tran"
            vCMD = New SqlCommand(vQuery, vConnection)
            vCMD.ExecuteNonQuery()

            ClearScreen()
        Catch ex As Exception
            vQuery = "rollback tran"
            vCMD = New SqlCommand(vQuery, vConnection)
            vCMD.ExecuteNonQuery()
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
        End Try


    End Sub

    Public Sub ClearScreen()
        Me.TextCouponNo.Text = ""
        Me.TextCouponName.Text = ""
        Me.DTPStartDate.Value = Now
        Me.DTPStopDate.Value = Now
        Me.TextCouponFormat.Text = ""
        Me.TextCountLenght.Text = ""
        Me.TextMyDescription.Text = ""
        Me.TextCouponCode.Text = ""
        Me.TextCouponValue.Text = ""
        Me.TextCouponQTY.Text = ""
    End Sub

    Private Sub FormCouponRecord_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call InitializeDataBase()
    End Sub

    Private Sub TextCouponCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextCouponCode.KeyDown
        Dim vCPCode As String
        Dim vCheckIsUsed As Integer
        Dim i As Integer
        Dim n As Integer
        Dim vListCoupon As ListViewItem

        If e.KeyCode = Keys.Enter Then
            vCPCode = Me.TextCouponCode.Text
            vQuery = "exec dbo.USP_NP_CouponData '" & vCPCode & "'"
            da = New SqlDataAdapter(vQuery, vConnection)
            ds = New DataSet
            da.Fill(ds, "IsUsed")
            dt = ds.Tables("IsUsed")

            If dt.Rows.Count > 0 Then
                vCheckIsUsed = dt.Rows.Item(0).Item("isused")
                Me.TextCountLenght.Text = dt.Rows.Item(0).Item("cplenght")
                Me.TextCouponFormat.Text = dt.Rows.Item(0).Item("cpformat")
                Me.TextCouponName.Text = dt.Rows.Item(0).Item("cpname")
                Me.DTPStartDate.Text = dt.Rows.Item(0).Item("fromdate")
                Me.DTPStopDate.Text = dt.Rows.Item(0).Item("todate")
                Me.TextMyDescription.Text = dt.Rows.Item(0).Item("mydescription")
                If vCheckIsUsed = 0 Then
                    Me.PB101.Visible = True
                    Me.PB102.Visible = False
                Else
                    Me.PB101.Visible = False
                    Me.PB102.Visible = True
                End If

                Me.ListViewCoupon.Items.Clear()
                For i = 0 To dt.Rows.Count - 1
                    vListCoupon = Me.ListViewCoupon.Items.Add(dt.Rows(i).Item("cpheader"))
                    vListCoupon.SubItems.Add(1).Text = Format(dt.Rows(i).Item("cpvalue"), "##,##0.00")
                    vListCoupon.SubItems.Add(2).Text = Format(dt.Rows(i).Item("cpqty"), "##,##0.00")
                    vListCoupon.SubItems.Add(3).Text = Format(dt.Rows(i).Item("cpapprove"), "##,##0.00")
                    vListCoupon.SubItems.Add(4).Text = Format(dt.Rows(i).Item("cpremain"), "##,##0.00")

                    For n = 0 To ListViewCoupon.Items.Count - 1
                        If ListViewCoupon.Items.Count Mod 2 = 0 Then
                            ListViewCoupon.Items(ListViewCoupon.Items.Count - 1).BackColor = Color.PowderBlue
                        End If
                    Next
                Next

            Else
                Me.PB101.Visible = True
                Me.PB102.Visible = False
            End If
        End If
    End Sub


    Private Sub TextCouponValue_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextCouponValue.KeyPress, TextCountLenght.KeyPress, TextCouponQTY.KeyPress
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

    Private Sub CMBFormat_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CMBFormat.Click
        Me.TextCouponFormat.Text = Me.CMBFormat.SelectedIndex
    End Sub

    'Private Sub ListViewCoupon_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewCoupon.SelectedIndexChanged
    '    listview1.EnsureVisible(listview1.Items.Count - 1)
    '    listview1.Items(listview1.Items.Count - 1).Focused = True
    '    listview1.Items(listview1.Items.Count - 1).Selected = True
    'End Sub

    Private Sub CMBFormat_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBFormat.SelectedIndexChanged

    End Sub

    Private Sub ListViewCoupon_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListViewCoupon.SelectedIndexChanged

    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub TextCouponCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextCouponCode.TextChanged

    End Sub

    Private Sub TextCouponNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextCouponNo.Leave

    End Sub

    Private Sub TextCouponNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextCouponNo.LostFocus

    End Sub

    Private Sub TextCouponNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextCouponNo.TextChanged

    End Sub
End Class