Imports System.IO
Imports Symbol
Imports Symbol.Barcode
Imports Symbol.Barcode.Reader
Imports System.Web
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Drawing
Imports System.Drawing.Bitmap
Imports System.ComponentModel
Imports System.Windows.Forms
Imports vb6 = Microsoft.VisualBasic

Public Class FormLogIn
    Dim vQuery As String

    Private Sub BTNExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNExit.Click
        Application.Exit()
    End Sub

    Private Sub BTNOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNOK.Click
        Dim vCheckAccess As Integer

        If Me.CMBProfit.Text = "" Then
            MsgBox("Please select profit", MsgBoxStyle.Critical, "Send Error Message")
            Me.CMBProfit.Focus()
            Exit Sub
        End If

        If Me.TBUserID.Text = "" Then
            MsgBox("Please insert userid", MsgBoxStyle.Critical, "Send Error Message")
            Me.TBUserID.Focus()
            Exit Sub
        End If

        vUserID = Me.TBUserID.Text
        vPassword = Me.TBPassword.Text
        vMemProfit = Me.CMBProfit.Text

        vUserID = Me.TBUserID.Text
        vPassword = Me.TBPassword.Text

        If vMemProfit = "S01" Then
            Dim vServiceS01 As New WebReferenceS01.WebServiceCalc
            Dim vAccessData As Integer = vServiceS01.vUserAccess(vUserID, vPassword)

            vCheckAccess = vAccessData

            If vAccessData = 0 Then
                MsgBox("User login fail.Check your userid and password", MsgBoxStyle.Critical, "Send Error Message")

                Me.TBPassword.Text = ""
                Me.TBUserID.Focus()
                Exit Sub
            End If
        End If


        If vMemProfit = "S02" Then
            Dim vServiceS02 As New WebReferenceS02.WebServiceCalc
            Dim vAccessData As Integer = vServiceS02.vUserAccess(vUserID, vPassword)

            vCheckAccess = vAccessData

            If vAccessData = 0 Then
                MsgBox("User login fail.Check your userid and password", MsgBoxStyle.Critical, "Send Error Message")

                Me.TBPassword.Text = ""
                Me.TBUserID.Focus()
                Exit Sub
            End If
        End If

        If vCheckAccess = 1 Then

            vQuery = "exec dbo.usp_hh_SearchPersonID '" & vUserID & "'"
            Call vGetData(vMemProfit, vQuery)
            vUserName = pds.Tables(0).Rows(0)("code").ToString

            FormMainApplication.Show()
            FormMainApplication.TBUserName.Text = vUserName
            Me.Hide()
        Else
            Me.TBPassword.Text = ""
            Me.TBUserID.Focus()
        End If
    End Sub

    Private Sub FormLogIn_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.CMBProfit.SelectedIndex = 0
    End Sub

    Private Sub CMBProfit_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CMBProfit.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.TBUserID.Focus()
        End If
    End Sub

    Private Sub CMBProfit_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMBProfit.SelectedIndexChanged
        If Me.CMBProfit.Text <> "" Then
            Me.TBUserID.Focus()
        End If
    End Sub

    Private Sub TBUserID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBUserID.KeyDown
        If e.KeyCode = Keys.Up Then
            Me.CMBProfit.Focus()
        End If

        If e.KeyCode = Keys.Enter Then
            Me.TBPassword.Focus()
        End If

        If e.KeyCode = Keys.Down Then
            Me.TBPassword.Focus()
        End If

        If e.KeyCode = Keys.Escape Then
            Me.TBUserID.Text = ""
            Me.TBUserID.Focus()
        End If
    End Sub

    Private Sub TBUserID_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBUserID.TextChanged

    End Sub

    Private Sub TBPassword_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TBPassword.KeyDown
        If e.KeyCode = Keys.Up Then
            Me.TBUserID.Focus()
        End If

        If e.KeyCode = Keys.Enter Then
            Me.BTNOK.Focus()
        End If

        If e.KeyCode = Keys.Down Then
            Me.BTNOK.Focus()
        End If

        If e.KeyCode = Keys.Escape Then
            Me.TBPassword.Text = ""
            Me.TBPassword.Focus()
        End If
    End Sub

    Private Sub TBPassword_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBPassword.TextChanged

    End Sub

    Private Sub BTNOK_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNOK.KeyDown
        Dim vCheckAccess As Integer

        If e.KeyCode = Keys.Enter Then

            If Me.CMBProfit.Text = "" Then
                MsgBox("Please select profit", MsgBoxStyle.Critical, "Send Error Message")
                Me.CMBProfit.Focus()
                Exit Sub
            End If

            If Me.TBUserID.Text = "" Then
                MsgBox("Please insert userid", MsgBoxStyle.Critical, "Send Error Message")
                Me.TBUserID.Focus()
                Exit Sub
            End If

            vUserID = Me.TBUserID.Text
            vPassword = Me.TBPassword.Text
            vMemProfit = Me.CMBProfit.Text

            vUserID = Me.TBUserID.Text
            vPassword = Me.TBPassword.Text

            If vMemProfit = "S01" Then
                Dim vServiceS01 As New WebReferenceS01.WebServiceCalc
                Dim vAccessData As Integer = vServiceS01.vUserAccess(vUserID, vPassword)

                vCheckAccess = vAccessData

                If vAccessData = 0 Then
                    MsgBox("User login fail.Check your userid and password", MsgBoxStyle.Critical, "Send Error Message")

                    Me.TBPassword.Text = ""
                    Me.TBUserID.Focus()
                End If
            End If


            If vMemProfit = "S02" Then
                Dim vServiceS02 As New WebReferenceS02.WebServiceCalc
                Dim vAccessData As Integer = vServiceS02.vUserAccess(vUserID, vPassword)

                vCheckAccess = vAccessData

                If vAccessData = 0 Then
                    MsgBox("User login fail.Check your userid and password", MsgBoxStyle.Critical, "Send Error Message")

                    Me.TBPassword.Text = ""
                    Me.TBUserID.Focus()
                End If
            End If

            If vCheckAccess = 1 Then

                vQuery = "exec dbo.usp_hh_SearchPersonID '" & vUserID & "'"
                Call vGetData(vMemProfit, vQuery)
                vUserName = pds.Tables(0).Rows(0)("code").ToString

                FormMainApplication.Show()
                FormMainApplication.TBUserName.Text = vUserName
                Me.Hide()
            Else
                Me.TBPassword.Text = ""
                Me.TBUserID.Focus()
            End If
        End If
    End Sub

    Private Sub BTNExit_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles BTNExit.KeyDown, BTNOK.KeyDown, CMBProfit.KeyDown
        If e.KeyCode = Keys.Escape Then
            Application.Exit()
        End If
    End Sub

    Private Sub Panel4_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Panel4.GotFocus

    End Sub
End Class