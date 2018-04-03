Imports System.Data
Imports Microsoft.VisualBasic

Public Class FormLogIn

    Private Sub BTNLogIN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNLogIN.Click
        Dim vUserID As String
        Dim vPassWord As String

        vUserID = Me.TBUserID.Text
        vPassWord = Me.TBPassword.Text

        Dim vService As New WebReference.WebServiceCalc
        Dim vCheckLogIn As String = vService.vLogIn(vUserID, vPassWord)

        If vCheckLogIn <> "" Then

            If frmLogIn Is Nothing Then
                frmLogIn = New FormLogIn
            End If
            frmLogIn.Hide()

            If frmMain Is Nothing Then
                frmMain = New frmDriveIn
            End If
            frmMain.Show()
            frmMain.BringToFront()

            frmMain.TBUserID.Text = vCheckLogIn

        Else
            MsgBox("ไม่สามารถเข้าใช้งานโปรแกรมได้ กรุณาตรวจสอบชื่อและรหัสผ่าน", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBPassword.Text = ""
        End If
    End Sub
End Class