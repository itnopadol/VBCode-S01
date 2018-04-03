VERSION 5.00
Begin VB.Form Form001 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "LogIn Program"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form001.frx":0000
   ScaleHeight     =   7965
   ScaleWidth      =   10950
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7440
      Left            =   8775
      Picture         =   "Form001.frx":72FB
      ScaleHeight     =   7410
      ScaleWidth      =   2460
      TabIndex        =   8
      Top             =   1350
      Width           =   2490
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7440
      Left            =   0
      Picture         =   "Form001.frx":A58E
      ScaleHeight     =   7410
      ScaleWidth      =   2460
      TabIndex        =   7
      Top             =   1350
      Width           =   2490
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      ScaleHeight     =   285
      ScaleWidth      =   11610
      TabIndex        =   6
      Top             =   1050
      Width           =   11640
   End
   Begin VB.CheckBox Connect 
      BackColor       =   &H8000000E&
      Caption         =   "ทำงานแบบ OffLine"
      Height          =   240
      Left            =   2625
      TabIndex        =   5
      Top             =   1500
      Width           =   1740
   End
   Begin VB.CommandButton Cmd101 
      Caption         =   "ตกลง"
      Height          =   390
      Left            =   5475
      TabIndex        =   2
      Top             =   3900
      Width           =   1365
   End
   Begin VB.TextBox Text102 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   4275
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   3375
      Width           =   2565
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4275
      TabIndex        =   1
      Top             =   2850
      Width           =   2565
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      ForeColor       =   &H8000000E&
      Height          =   165
      Left            =   3375
      TabIndex        =   4
      Top             =   3375
      Width           =   765
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   3375
      TabIndex        =   3
      Top             =   2850
      Width           =   840
   End
End
Attribute VB_Name = "Form001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Cmd101_Click()
Dim NetWork
On Error GoTo CallError
vUserID = Trim(Text101.Text)
vPassword = Trim(Text102.Text)
If Connect.Value = 0 Then
    Call InitializeDataBase1
    vConnect = 0
Else
    vConnect = 1
End If
MDIForm1.Order1.Enabled = True
MDIForm1.Order2.Enabled = True
Unload Me

CallError:
If Err.Description <> "" Then
MsgBox Err.Description
Text102.Text = ""
End If
End Sub

Private Sub Form_Load()
Dim HKey, wsho, HKey1, wsho1
Dim vQuery As String
Dim vCompany1 As String
Dim vCompany As String
Dim vRecordset As New ADODB.Recordset
Dim vUserID1 As String

Call InitializeDatabase
On Error Resume Next
Set wsho = CreateObject("Wscript.Shell")
        HKey = wsho.RegRead("HKCU\Software\BanChiang Soft\LOGIN_DEFAULT")
        HKey = wsho.RegRead("HKCU\Software\BanChiang Soft\LOGIN_DEFAULT")
        vQuery = "select rtrim(SUBSTRING('" & Trim(HKey) & "', CHARINDEX('^','" & Trim(HKey) & "', 2)+1,100 )) as NameUser,rtrim(SUBSTRING('" & Trim(HKey) & "', 1,CHARINDEX('^','" & Trim(HKey) & "', 2)-1 )) as Company"
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
               vUserID1 = vRecordset.Fields("nameuser").Value
        End If
    vRecordset.Close

Me.Text101.Text = vUserID1
Me.Text102.SetFocus
End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text102.SetFocus
End If
End Sub

Private Sub Text102_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Cmd101_Click
End If
End Sub
