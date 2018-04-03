Option Explicit On
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic
Imports System.IO

Public Class FormLogIn
    Public vConnectionString As String
    Public vConnection As SqlConnection

    Private Sub BTNLogIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNLogIn.Click
        Dim vCheckLogIn As Integer

        If TextUserID.Text <> "" Then
            vUserID = Trim(TextUserID.Text)
            vPassword = Trim(TextPassword.Text)
            vCheckLogIn = CheckUserLogIn(vUserID, vPassword)
            If vCheckLogIn = 1 Then
                FormMain.Show()
                Me.Hide()

                FormMain.BuyMenu.Enabled = True
                FormMain.SaleMenu.Enabled = True
                FormMain.ItemStoreMenu.Enabled = True
                FormMain.VendorMenu.Enabled = True
                FormMain.CustomerMenu.Enabled = True
                FormMain.ItemMenu.Enabled = True
                FormMain.ProgramMenu.Enabled = True
                FormMain.AccountMenu.Enabled = True
                FormMain.AddOnMenu.Enabled = True
                FormMain.ManageMenu.Enabled = True
                FormMain.WindowsMenu.Enabled = True

                vComputerName = System.Environment.MachineName
                vWindowsLogIn = System.Environment.UserName
                vWindowsName = vComputerName & "/" & vWindowsLogIn

            End If
        Else
            MsgBox("กรุณากรอกชื่อเข้าใช้งานโปรแกรมด้วย", MsgBoxStyle.Critical, "Send Error")
        End If
    End Sub

    Public Function CheckUserLogIn(ByVal vUser As String, ByVal vPassword As String) As Integer
        On Error GoTo ErrDescription

        ''vConnectionString = "Persist Security Info = False;User ID='" & vUser & "';Password='" & vPassword & "';Max Pool Size = 10000;Min Pool Size = 5;Data Source = Nebula;Initial Catalog = BCNP"
        vConnectionString = "Persist Security Info = False;User ID='" & vUser & "';Password='" & vPassword & "';Max Pool Size = 10000;Min Pool Size = 5;Data Source = Nebula;Initial Catalog = BCNP"
        'vConnectionString = "Persist Security Info = False;User ID='" & vUser & "';Password='" & vPassword & "';Max Pool Size = 10000;Min Pool Size = 5;Data Source = S02DB;Initial Catalog = BCNP"
        vConnection = New SqlConnection(vConnectionString)
        vConnection.Open()
        CheckUserLogIn = 1

ErrDescription:
        If Err.Description <> "" Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Send Error LogIn Program")
            Me.TextPassword.Text = ""
            Me.TextPassword.Focus()
            CheckUserLogIn = 0
        End If

    End Function


    Private Sub BTNLogOff_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTNLogOff.Click
        Dim vAnswer As Integer
        vAnswer = MsgBox("คุณต้องการออกจากโปรแกรมใช่หรือไม่ ?", MsgBoxStyle.YesNo, "Send Question Message ?")
        If vAnswer = 6 Then
            Application.Exit()
        End If

    End Sub

    Private Sub TextPassword_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextPassword.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call BTNLogIn.Focus()
        End If
    End Sub

    Private Sub TextUserID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextUserID.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.TextPassword.Focus()
        End If
    End Sub


    Private Sub FormLogIn_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Dim vOpenFile As New OpenFileDialog
        'Dim vText As String
        'Dim wsho As String
        'Dim HKey As String

        'vOpenFile.FileName = "HKCU\Software\BanChiang Soft\LOGIN_DEFAULT"
        'Dim vString As New StreamReader(vOpenFile.FileName)

        'vText = Trim(vString.ReadToEnd)

        'HKey = Shell("CALC.EXE", AppWinStyle.NormalFocus)

        ''wsho = CreateObject("Wscript.Shell")
        ''HKey = wsho.RegRead("HKCU\Software\BanChiang Soft\LOGIN_DEFAULT")
        ''HKey = wsho.RegRead("HKCU\Software\BanChiang Soft\LOGIN_DEFAULT")

        'wsho = My.Computer.Name
        'MsgBox(wsho)
        'Me.TextUserID.Focus()

    End Sub
End Class