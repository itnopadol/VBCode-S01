Imports System.data
Imports Microsoft.VisualBasic
Imports System.IO
Imports vb6 = Microsoft.VisualBasic
Imports System.Globalization
Imports System.Windows.Forms


Public Class FrmMobileApp

    Private Sub TBPassword_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBPassword.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Up Then
            Me.TBUserID.Focus()
            Me.TBUserID.SelectAll()
        End If

        Dim vAnswer As Integer

        If e.KeyCode = Keys.Escape Then
            vAnswer = MsgBox("คุณต้องการออกจากโปรแกรมใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message?")
            If vAnswer = 6 Then
                Application.Exit()
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBPassword_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBPassword.TextChanged
        Dim vUserCode As String
        Dim vPassword As String
        Dim vLenPassword As Integer
        Dim vCheckTypeLogIn As String

        On Error GoTo ErrDescription

        vLenPassword = Len(Me.TBPassword.Text)
        If vLenPassword = 4 And Me.TBUserID.Text <> "" Then

            vUserCode = Me.TBUserID.Text
            vPassword = Me.TBPassword.Text

            Dim vService1 As New WebReference.WebServiceCalc
            Dim ds1 As DataSet = vService1.vLogIn(vUserCode, vPassword)

            If ds1.Tables(0).Rows.Count > 0 Then
                vCheckLogIn = ds1.Tables(0).Rows(0)("username").ToString
                vUserName = ds1.Tables(0).Rows(0)("username").ToString
                vDuty = ds1.Tables(0).Rows(0)("duty").ToString
                vLevelID = ds1.Tables(0).Rows(0)("levelid").ToString
                vPersonName = ds1.Tables(0).Rows(0)("username").ToString

                vMemUserID = vUserCode
                vMemPassword = vPassword
            Else
                vCheckLogIn = ""
                vUserName = ""
                vDuty = ""
                vLevelID = 0
                vPersonName = ""
                vMemUserID = vUserCode
                vMemPassword = vPassword
            End If


            If vCheckLogIn <> "" Then

                Me.PNLogIn.Visible = False
                Me.PNSelectJob.Visible = True
                Me.PNSelectJob.BringToFront()
                Me.RBJob1.Focus()
            Else
                MsgBox("ไม่สามารถเข้าใช้งานโปรแกรมได้ กรุณาตรวจสอบชื่อและรหัสผ่าน", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBPassword.Text = ""
                Me.TBPassword.Focus()
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub TBUserID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBUserID.KeyDown
        On Error GoTo ErrDescription

        If e.KeyCode = Keys.Enter Then
            Me.TBPassword.Focus()
            Me.TBPassword.SelectAll()
        End If

        Dim vAnswer As Integer

        If e.KeyCode = Keys.Escape Then
            vAnswer = MsgBox("คุณต้องการออกจากโปรแกรมใช่หรือไม่", MsgBoxStyle.YesNo, "Send Question Message?")
            If vAnswer = 6 Then
                Application.Exit()
            End If
        End If

ErrDescription:

        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error Message")
            Exit Sub
        End If
    End Sub

    Private Sub BTNSelectJob_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNSelectJob.Click
        On Error Resume Next

        If Me.RBJob1.Checked = True Then
            FrmItemData.Show()
            Me.Hide()
            FrmItemData.TBBarCode.Focus()
            FrmItemData.TBBarCode.SelectAll()
        End If

        If Me.RBJob2.Checked = True Then
            FrmAddItemShelf.Show()
            Me.Hide()
            FrmAddItemShelf.TBShelf.Focus()
            FrmAddItemShelf.TBShelf.SelectAll()
        End If

        If Me.RBJob3.Checked = True Then
            FrmCountStock.Show()
            Me.Hide()
            FrmCountStock.CMBReason.Focus()
        End If

        If Me.RBJob4.Checked = True Then
            FrmItemPrint.Show()
            Me.Hide()
            FrmItemPrint.TBBarCode.Focus()
        End If
    End Sub

    Private Sub RBJob1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RBJob1.KeyDown, RBJob2.KeyDown
        On Error Resume Next

        If e.KeyCode = 49 Then
            Me.RBJob1.Checked = True
            FrmItemData.Show()
            Me.Hide()
            FrmItemData.TBBarCode.Focus()
            FrmItemData.TBBarCode.SelectAll()
        End If

        If e.KeyCode = 50 Then
            Me.RBJob2.Checked = True
            FrmAddItemShelf.Show()
            Me.Hide()
            FrmAddItemShelf.TBShelf.Focus()
            FrmAddItemShelf.TBShelf.SelectAll()
        End If

        If e.KeyCode = 51 Then
            Me.RBJob3.Checked = True
            FrmCountStock.Show()
            Me.Hide()
            FrmCountStock.CMBReason.Focus()
        End If


        If e.KeyCode = 52 Then
            Me.RBJob4.Checked = True
            FrmItemPrint.Show()
            Me.Hide()
            FrmItemPrint.TBBarCode.Focus()
        End If

        If e.KeyCode = Keys.Escape Then
            Me.PNSelectJob.Visible = False
            Me.PNLogIn.Visible = True
            Me.PNLogIn.BringToFront()
            Me.TBPassword.Text = ""
            Me.TBUserID.Focus()
            Me.TBUserID.SelectAll()
        End If

    End Sub

    Private Sub BTNSelectJob_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNSelectJob.KeyDown
        On Error Resume Next

        If e.KeyCode = 49 Then
            Me.RBJob1.Checked = True
            FrmItemData.Show()
            Me.Hide()
            FrmItemData.TBBarCode.Focus()
            FrmItemData.TBBarCode.SelectAll()
        End If

        If e.KeyCode = 50 Then
            Me.RBJob2.Checked = True
            FrmAddItemShelf.Show()
            Me.Hide()
            FrmAddItemShelf.TBShelf.Focus()
            FrmAddItemShelf.TBShelf.SelectAll()
        End If

        If e.KeyCode = 51 Then
            Me.RBJob3.Checked = True
            FrmCountStock.Show()
            Me.Hide()
            FrmCountStock.CMBReason.Focus()
        End If

        If e.KeyCode = 52 Then
            Me.RBJob4.Checked = True
            FrmItemPrint.Show()
            Me.Hide()
            FrmItemPrint.TBBarCode.Focus()
        End If

        If e.KeyCode = Keys.Escape Then
            Me.PNSelectJob.Visible = False
            Me.PNLogIn.Visible = True
            Me.PNLogIn.BringToFront()
            Me.TBPassword.Text = ""
            Me.TBUserID.Focus()
            Me.TBUserID.SelectAll()
        End If
    End Sub

    Private Sub BTNLogIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNLogIn.Click
        On Error Resume Next

        Me.PNSelectJob.Visible = False
        Me.PNLogIn.Visible = True
        Me.PNLogIn.BringToFront()
        Me.TBPassword.Text = ""
        Me.TBUserID.Focus()
        Me.TBUserID.SelectAll()
    End Sub

    Private Sub RBJob1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBJob1.CheckedChanged

    End Sub

    Private Sub RBJob3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBJob3.CheckedChanged

    End Sub

    Private Sub RBJob3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RBJob3.KeyDown
        On Error Resume Next

        If e.KeyCode = 49 Then
            Me.RBJob1.Checked = True
            FrmItemData.Show()
            Me.Hide()
            FrmItemData.TBBarCode.Focus()
            FrmItemData.TBBarCode.SelectAll()
        End If

        If e.KeyCode = 50 Then
            Me.RBJob2.Checked = True
            FrmAddItemShelf.Show()
            Me.Hide()
            FrmAddItemShelf.TBShelf.Focus()
            FrmAddItemShelf.TBShelf.SelectAll()
        End If

        If e.KeyCode = 51 Then
            Me.RBJob3.Checked = True
            FrmCountStock.Show()
            Me.Hide()
            FrmCountStock.CMBReason.Focus()
        End If

        If e.KeyCode = 52 Then
            Me.RBJob4.Checked = True
            FrmItemPrint.Show()
            Me.Hide()
            FrmItemPrint.TBBarCode.Focus()
        End If

        If e.KeyCode = Keys.Escape Then
            Me.PNSelectJob.Visible = False
            Me.PNLogIn.Visible = True
            Me.PNLogIn.BringToFront()
            Me.TBPassword.Text = ""
            Me.TBUserID.Focus()
            Me.TBUserID.SelectAll()
        End If
    End Sub

    Private Sub RBJob4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBJob4.CheckedChanged

    End Sub
End Class
