VERSION 5.00
Begin VB.Form FrmLogIN 
   Caption         =   "หน้า LogIN"
   ClientHeight    =   9915
   ClientLeft      =   3645
   ClientTop       =   1035
   ClientWidth     =   15375
   Icon            =   "FrmLogIN.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmLogIN.frx":08CA
   ScaleHeight     =   9915
   ScaleMode       =   0  'User
   ScaleWidth      =   15350
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMDLogIN 
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   5820
      TabIndex        =   1
      Top             =   4095
      Width           =   1140
   End
   Begin VB.TextBox TXTPassword 
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
      Left            =   4050
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   3450
      Width           =   1740
   End
   Begin VB.TextBox TXTUserID 
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
      Left            =   4050
      TabIndex        =   2
      Top             =   3000
      Width           =   1740
   End
   Begin VB.TextBox TXTCompany 
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
      Left            =   4050
      TabIndex        =   3
      Top             =   2475
      Width           =   1740
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "UserID และ Password เป็น UserID และ Password ที่เข้า BC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   315
      Left            =   1800
      TabIndex        =   8
      Top             =   1350
      Width           =   5715
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LogIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   390
      Left            =   2235
      TabIndex        =   7
      Top             =   1800
      Width           =   4725
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      X1              =   2066.634
      X2              =   7083.463
      Y1              =   4950
      Y2              =   4950
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      X1              =   2066.634
      X2              =   7083.463
      Y1              =   1725
      Y2              =   1725
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      X1              =   7083.463
      X2              =   7083.463
      Y1              =   1725
      Y2              =   4950
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   2066.634
      X2              =   2066.634
      Y1              =   1725
      Y2              =   4950
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PassWord"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   240
      Left            =   2925
      TabIndex        =   6
      Top             =   3450
      Width           =   990
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   240
      Left            =   3150
      TabIndex        =   5
      Top             =   3000
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Company"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   315
      Left            =   3075
      TabIndex        =   4
      Top             =   2475
      Width           =   840
   End
End
Attribute VB_Name = "FrmLogIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vDataBase As String, vServer As String
Dim vDocHeader As String

Private Sub CMDLogIN_Click()

On Error GoTo CallError

vPassword = Trim(TXTPassword.Text)
vCompany = Trim(TXTCompany.Text)
vUserID = Trim(TXTUserID.Text)

Call InitializeDataBase1
Call ConnectCompany
Call InitializeDataBase

MDIFrmProgramPrint.Order2.Enabled = True
MDIFrmProgramPrint.Order3.Enabled = True
MDIFrmProgramPrint.Order4.Enabled = True
MDIFrmProgramPrint.Order5.Enabled = True
MDIFrmProgramPrint.Order6.Enabled = True
MDIFrmProgramPrint.Order7.Enabled = True
MDIFrmProgramPrint.Order8.Enabled = True
MDIFrmProgramPrint.Order9.Enabled = True
MDIFrmProgramPrint.Order0.Enabled = True
MDIFrmProgramPrint.DO1.Enabled = True
MDIFrmProgramPrint.nWindows.Enabled = True
MDIFrmProgramPrint.Caption = "โปรแกรมพิมพ์เอกสาร BCAccount 5.5 Update 03.10.2008" & " User LogIN : " & vUserID
Unload Me

CallError:
If Err.Description <> "" Then
MsgBox Err.Description
TXTPassword.Text = ""
End If

End Sub
Public Function ConnectCompany()

Select Case UCase(vCompany)
    Case "NP"
        vDataBase = "BCNP"
        vServer = "Nebula"
    Case "CHAMP"
        vDataBase = "Champ"
        vServer = "Nebula"
    Case "BCVAT"
        vDataBase = "BCVAT"
        vServer = "Dev"
End Select

conn.Provider = "SQLOLEDB.1"
conn.Properties("Persist Security Info").Value = False
conn.Properties("User ID").Value = vUserID
conn.Properties("Password").Value = vPassword
conn.Properties("Initial Catalog").Value = vDataBase
conn.Properties("Data Source").Value = vServer
conn.CursorLocation = adUseClient
conn.Open
conn.Close


End Function

Private Sub Form_Load()
Dim HKey, wsho, HKey1, wsho1
Dim vQuery As String
Dim vCompany1 As String
Dim vCompany As String
Dim vRecordset As New ADODB.Recordset

Call InitializeDataBase
On Error Resume Next
TXTCompany.Text = "np"
Set wsho = CreateObject("Wscript.Shell")
        HKey = wsho.RegRead("HKCU\Software\BanChiang Soft\LOGIN_DEFAULT")
        HKey = wsho.RegRead("HKCU\Software\BanChiang Soft\LOGIN_DEFAULT")
        vQuery = "select rtrim(SUBSTRING('" & Trim(HKey) & "', CHARINDEX('^','" & Trim(HKey) & "', 2)+1,100 )) as NameUser,rtrim(SUBSTRING('" & Trim(HKey) & "', 1,CHARINDEX('^','" & Trim(HKey) & "', 2)-1 )) as Company"
        If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
            vCompany1 = vRecordset.Fields("company").Value
               vUserID1 = vRecordset.Fields("nameuser").Value
        End If
    vRecordset.Close

TXTCompany.Text = vCompany1
TXTUserID.Text = vUserID1
vPassword1 = Trim(TXTPassword.Text)

End Sub

Private Sub TXTPassword_KeyPress(KeyAscii As Integer)
On Error GoTo CallError

If KeyAscii = 13 Then
    Call CMDLogIN_Click
End If
CallError:
If Err.Description <> "" Then
MsgBox Err.Description
TXTPassword.Text = ""
End If
End Sub

Private Sub TXTUserID_KeyPress(KeyAscii As Integer)
On Error GoTo CallError
If KeyAscii = 13 Then
    Call InitializeDataBase1
    TXTPassword.SetFocus
End If
CallError:
If Err.Description <> "" Then
MsgBox Err.Description
TXTPassword.Text = ""
End If
End Sub
