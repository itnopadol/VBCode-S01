Option Explicit On
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization
Public Class FormCEProgram
    Dim ds As DataSet
    Dim da As SqlDataAdapter
    Dim dt As DataTable
    Dim vQuery As String
    Dim cmd As SqlCommand
    Dim vReadQuery As SqlDataReader

    Dim vMemTimeIndex As Integer
    Dim vIsNumber As Integer

    Private Sub FormCEProgram_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Width = Screen.PrimaryScreen.Bounds.Width
        Me.Height = Screen.PrimaryScreen.Bounds.Height
        Me.Left = 0
        Me.Top = 0
        Me.WindowState = FormWindowState.Maximized

        Call InitializeDataBase()
        Call vGetBeginData()
        Me.CMBDocType.SelectedIndex = 0
        Me.DTPDocDate.Value = Now

        Call RetriveEntranceList()
        Me.CMBEntranceType.SelectedIndex = 1
        Call vSearchCEData()
    End Sub

    Public Sub RetriveEntranceList()
        Dim i As Integer

        On Error Resume Next

        Me.CMBEntranceType.Items.Clear()
        vQuery = "select * from Npmaster.dbo.TB_CE_Entrance where iscancel = 0"
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "SearchEntrance")
        dt = ds.Tables("SearchEntrance")
        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                Me.CMBEntranceType.Items.Add(dt.Rows(i).Item("name1"))
            Next
        End If

    End Sub

    Public Sub vGetBeginData()
        Dim i As Integer
        Dim n As Integer

        On Error Resume Next

        Me.DGVDataDetails.Rows.Add(24)
        For i = 0 To 24 - 1
            n = n + 1
            Me.DGVDataDetails.Item(0, i).Value = n
        Next

        Me.DGVDataDetails.CurrentCell = Me.DGVDataDetails.Item(1, 0)
    End Sub


    Public Sub vSearchCEData()
        Dim vDocDate As String
        Dim vTypeRec As Integer
        Dim vEntranceID As Integer
        Dim i As Integer

        On Error Resume Next

        If Me.CMBDocType.Text = "" Then
            Me.CMBDocType.Focus()
            Exit Sub
        End If

        If Me.CMBEntranceType.Text = "" Then
            Me.CMBEntranceType.Focus()
            Exit Sub
        End If

        vDocDate = vb6.Day(Me.DTPDocDate.Value) & "/" & vb6.Month(Me.DTPDocDate.Value) & "/" & vb6.Year(Me.DTPDocDate.Value)
        vTypeRec = Me.CMBDocType.SelectedIndex
        vEntranceID = Me.CMBEntranceType.SelectedIndex

        vQuery = "exec npmaster.dbo.USP_CE_TimeRecList '" & vDocDate & "'," & vTypeRec & "," & vEntranceID & " "
        da = New SqlDataAdapter(vQuery, vConnection)
        ds = New DataSet
        da.Fill(ds, "Data")
        dt = ds.Tables("Data")

        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                Me.DGVDataDetails.Item(0, i).Value = dt.Rows(i).Item("timetitle")

                If dt.Rows(i).Item("inrec") <> 0 Then
                    Me.DGVDataDetails.Item(1, i).Value = dt.Rows(i).Item("inrec")
                Else
                    Me.DGVDataDetails.Item(1, i).Value = ""
                End If
                If dt.Rows(i).Item("foreigner") <> 0 Then
                    Me.DGVDataDetails.Item(2, i).Value = dt.Rows(i).Item("foreigner")
                Else
                    Me.DGVDataDetails.Item(2, i).Value = ""
                End If
                If dt.Rows(i).Item("notbuy") <> 0 Then
                    Me.DGVDataDetails.Item(3, i).Value = dt.Rows(i).Item("notbuy")
                Else
                    Me.DGVDataDetails.Item(3, i).Value = ""
                End If

                If dt.Rows(i).Item("checkproblem") <> 0 Then
                    Me.DGVDataDetails.Item(4, i).Value = dt.Rows(i).Item("checkproblem")
                Else
                    Me.DGVDataDetails.Item(4, i).Value = ""
                End If

                Me.DGVDataDetails.Item(5, i).Value = dt.Rows(i).Item("mydescription")
            Next
        End If

    End Sub

    Private Sub CMBEntranceType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBEntranceType.SelectedIndexChanged
        Call vSearchCEData()
    End Sub

    Private Sub CMBDocType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBDocType.SelectedIndexChanged
        Call vSearchCEData()
    End Sub

    Private Sub DTPDocDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DTPDocDate.ValueChanged
        Call vSearchCEData()
    End Sub

    Private Sub BTNSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSave.Click
        Dim vDocType As Integer
        Dim vEntranceType As Integer
        Dim vDocDate As String
        Dim vTimeID As Integer
        Dim vInQty1 As String
        Dim vForeigner1 As String
        Dim vNotBuy1 As String
        Dim vHaveProblem1 As String

        Dim vInQty As Integer
        Dim vForeigner As Integer
        Dim vNotBuy As Integer
        Dim vHaveProblem As Integer

        Dim vMydescription As String

        Dim i As Integer


        If Me.CMBDocType.Text = "" Then
            MsgBox("ยังไม่ได้ระบุประเภทข้อมูลที่จะบันทึก กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If

        If Me.CMBEntranceType.Text = "" Then
            MsgBox("ยังไม่ได้ระบุช่องทางที่จะบันทึกข้อมูล กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If

        vDocType = Me.CMBDocType.SelectedIndex
        vEntranceType = Me.CMBEntranceType.SelectedIndex
        vDocDate = Day(Me.DTPDocDate.Value) & "/" & Month(Me.DTPDocDate.Value) & "/" & Year(Me.DTPDocDate.Value)

        Try

            vQuery = "begin tran"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            For i = 0 To Me.DGVDataDetails.Rows.Count - 1
                vTimeID = i + 1

                vInQty1 = Me.DGVDataDetails.Item(1, i).Value
                vForeigner1 = Me.DGVDataDetails.Item(2, i).Value
                vNotBuy1 = Me.DGVDataDetails.Item(3, i).Value
                vHaveProblem1 = Me.DGVDataDetails.Item(4, i).Value
                vMydescription = Me.DGVDataDetails.Item(5, i).Value

                If vInQty1 <> "" Then
                    vInQty = vInQty1
                Else
                    vInQty = 0
                End If

                If vForeigner1 <> "" Then
                    vForeigner = vForeigner1
                Else
                    vForeigner = 0
                End If

                If vNotBuy1 <> "" Then
                    vNotBuy = vNotBuy1
                Else
                    vNotBuy = 0
                End If

                If vHaveProblem1 <> "" Then
                    vHaveProblem = vHaveProblem1
                Else
                    vHaveProblem = 0
                End If

                vQuery = "exec npmaster.dbo.USP_CE_CustomerRecSaveData " & vDocType & ",'" & vDocDate & "'," & vEntranceType & "," & vTimeID & "," & vInQty & "," & vForeigner & "," & vNotBuy & "," & vHaveProblem & ",'" & vMydescription & "','" & vUserID & "' "
                cmd = New SqlCommand(vQuery, vConnection)
                cmd.ExecuteNonQuery()

            Next

            vQuery = "commit tran"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            Call vSearchCEData()
            MsgBox("บันทึกข้อมูลเรียบร้อยแล้ว กรุณาตรวจสอบ", MsgBoxStyle.Information, "Send Information Message")

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            vQuery = "rollback tran"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()
        End Try
    End Sub

    Public Sub SaveData()
        Dim vDocType As Integer
        Dim vEntranceType As Integer
        Dim vDocDate As String
        Dim vTimeID As Integer
        Dim vInQty1 As String
        Dim vForeigner1 As String
        Dim vNotBuy1 As String
        Dim vHaveProblem1 As String

        Dim vInQty As Integer
        Dim vForeigner As Integer
        Dim vNotBuy As Integer
        Dim vHaveProblem As Integer

        Dim vMydescription As String

        Dim i As Integer


        If Me.CMBDocType.Text = "" Then
            MsgBox("ยังไม่ได้ระบุประเภทข้อมูลที่จะบันทึก กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If

        If Me.CMBEntranceType.Text = "" Then
            MsgBox("ยังไม่ได้ระบุช่องทางที่จะบันทึกข้อมูล กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If

        vDocType = Me.CMBDocType.SelectedIndex
        vEntranceType = Me.CMBEntranceType.SelectedIndex
        vDocDate = Day(Me.DTPDocDate.Value) & "/" & Month(Me.DTPDocDate.Value) & "/" & Year(Me.DTPDocDate.Value)

        Try

            vQuery = "begin tran"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            For i = 0 To Me.DGVDataDetails.Rows.Count - 1
                vTimeID = i + 1

                vInQty1 = Me.DGVDataDetails.Item(1, i).Value
                vForeigner1 = Me.DGVDataDetails.Item(2, i).Value
                vNotBuy1 = Me.DGVDataDetails.Item(3, i).Value
                vHaveProblem1 = Me.DGVDataDetails.Item(4, i).Value
                vMydescription = Me.DGVDataDetails.Item(5, i).Value

                If vInQty1 <> "" Then
                    vInQty = vInQty1
                Else
                    vInQty = 0
                End If

                If vForeigner1 <> "" Then
                    vForeigner = vForeigner1
                Else
                    vForeigner = 0
                End If

                If vNotBuy1 <> "" Then
                    vNotBuy = vNotBuy1
                Else
                    vNotBuy = 0
                End If

                If vHaveProblem1 <> "" Then
                    vHaveProblem = vHaveProblem1
                Else
                    vHaveProblem = 0
                End If

                vQuery = "exec npmaster.dbo.USP_CE_CustomerRecSaveData " & vDocType & ",'" & vDocDate & "'," & vEntranceType & "," & vTimeID & "," & vInQty & "," & vForeigner & "," & vNotBuy & "," & vHaveProblem & ",'" & vMydescription & "','" & vUserID & "' "
                cmd = New SqlCommand(vQuery, vConnection)
                cmd.ExecuteNonQuery()

            Next

            vQuery = "commit tran"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()

            Call vSearchCEData()
            MsgBox("บันทึกข้อมูลเรียบร้อยแล้ว กรุณาตรวจสอบ", MsgBoxStyle.Information, "Send Information Message")

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            vQuery = "rollback tran"
            cmd = New SqlCommand(vQuery, vConnection)
            cmd.ExecuteNonQuery()
        End Try
    End Sub

    Private Sub BTNClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNClose.Click
        Me.Close()
    End Sub

    Private Sub BTNSave_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNSave.KeyDown, BTNClose.KeyDown, CMBDocType.KeyDown, CMBEntranceType.KeyDown, DTPDocDate.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.F5 Then
            Call SaveData()
        End If

        Dim vAnswer As Integer
        If e.KeyCode = Keys.Escape Then
            vAnswer = MsgBox("คุณต้องการออกโปรแกรมใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message")
            If vAnswer = 6 Then
                Me.Close()
            Else
                Exit Sub
            End If
        End If
    End Sub

    Private Sub DGVDataDetails_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVDataDetails.CellEndEdit
        Dim vCharStr As String

        On Error Resume Next

        If e.ColumnIndex = 1 Then
            vCharStr = Me.DGVDataDetails.Item(1, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVDataDetails.Item(1, e.RowIndex).Value = ""
                    MsgBox("ช่องจำนวนที่กำหนด ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                End If
            End If
        End If

        If e.ColumnIndex = 2 Then
            vCharStr = Me.DGVDataDetails.Item(2, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVDataDetails.Item(2, e.RowIndex).Value = ""
                    MsgBox("ช่องจำนวนที่กำหนด ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                End If
            End If
        End If

        If e.ColumnIndex = 3 Then
            vCharStr = Me.DGVDataDetails.Item(3, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVDataDetails.Item(3, e.RowIndex).Value = ""
                    MsgBox("ช่องจำนวนที่กำหนด ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                End If
            End If
        End If

        If e.ColumnIndex = 4 Then
            vCharStr = Me.DGVDataDetails.Item(4, e.RowIndex).Value
            If vCharStr <> "" Then
                Call vCheckNumber(vCharStr)
                If vIsNumber = 0 Then
                    Me.DGVDataDetails.Item(4, e.RowIndex).Value = ""
                    MsgBox("ช่องจำนวนที่กำหนด ไม่สามารถกรอกตัวหนังสือได้ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
                End If
            End If
        End If


        Dim vInQty1 As String
        Dim vForeigner1 As String
        Dim vNotBuy1 As String
        Dim vNotProblem1 As String
        Dim vHaveProblem1 As String

        Dim vInQty As Integer
        Dim vForeigner As Integer
        Dim vNotBuy As Integer
        Dim vNotProblem As Integer
        Dim vHaveProblem As Integer

        vInQty1 = Me.DGVDataDetails.Item(1, e.RowIndex).Value
        vForeigner1 = Me.DGVDataDetails.Item(2, e.RowIndex).Value
        vNotBuy1 = Me.DGVDataDetails.Item(3, e.RowIndex).Value
        vHaveProblem1 = Me.DGVDataDetails.Item(4, e.RowIndex).Value

        If vInQty1 <> "" Then
            vInQty = vInQty1
        Else
            vInQty = 0
        End If

        If vForeigner1 <> "" Then
            vForeigner = vForeigner1
        Else
            vForeigner = 0
        End If

        If vNotBuy1 <> "" Then
            vNotBuy = vNotBuy1
        Else
            vNotBuy = 0
        End If

        If vHaveProblem1 <> "" Then
            vHaveProblem = vHaveProblem1
        Else
            vHaveProblem = 0
        End If

        If vForeigner > vInQty Then
            MsgBox("ไม่สามารถกรอกจำนวน กลุ่มคนต่างชาติ มากกว่า จำนวนคนเข้าร้าน กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.DGVDataDetails.Item(2, e.RowIndex).Value = ""
        End If

        If vNotBuy > vInQty Then
            MsgBox("ไม่สามารถกรอกจำนวน ไม่ซื้อสินค้า มากกว่า จำนวนคนเข้าร้าน กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.DGVDataDetails.Item(3, e.RowIndex).Value = ""
        End If

        If vHaveProblem > vInQty Then
            MsgBox("ไม่สามารถกรอกจำนวน ตรวจสอบแล้วมีปัญหา มากกว่า จำนวนคนเข้าร้าน กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.DGVDataDetails.Item(4, e.RowIndex).Value = ""
        End If

        If vHaveProblem > (vInQty - vNotBuy) Then
            MsgBox("ไม่สามารถกรอกจำนวน ตรวจสอบแล้วมีปัญหา มากกว่า จำนวนคนซื้อ กรุณาตรวจสอบ", MsgBoxStyle.Critical, "Send Error Message")
            Me.DGVDataDetails.Item(4, e.RowIndex).Value = ""
        End If

    End Sub

    Public Sub vCheckNumber(ByVal vNumber As String)
        Dim vLen As Integer
        Dim vChar As String
        Dim i As Integer
        Dim vString As String

        On Error Resume Next

        vString = vNumber
        vLen = vb6.Len(vString)
        For i = 1 To vLen
            vChar = Mid(vString, i, 1)

            If vChar = "1" Or vChar = "2" Or vChar = "3" Or vChar = "4" Or vChar = "5" Or vChar = "6" Or vChar = "7" Or vChar = "8" Or vChar = "9" Or vChar = "0" Or vChar = "," Then
                vIsNumber = 1
            Else
                vIsNumber = 0
                GoTo Line1
            End If
        Next
Line1:

    End Sub

    Private Sub DGVDataDetails_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVDataDetails.CellContentClick

    End Sub

    Private Sub DGVDataDetails_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DGVDataDetails.KeyDown
        On Error Resume Next

        If e.KeyCode = Keys.F5 Then
            Call SaveData()
        End If

        Dim vAnswer As Integer
        If e.KeyCode = Keys.Escape Then
            vAnswer = MsgBox("คุณต้องการออกโปรแกรมใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message")
            If vAnswer = 6 Then
                Me.Close()
            Else
                Exit Sub
            End If
        End If
    End Sub
End Class