VERSION 5.00
Begin VB.Form FormLogIN 
   Caption         =   "LogIN User"
   ClientHeight    =   8985
   ClientLeft      =   2025
   ClientTop       =   2010
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FormLogIN.frx":0000
   ScaleHeight     =   8985
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.TextBox TXTPassWord 
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
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   5325
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3975
      Width           =   1965
   End
   Begin VB.TextBox TXTUser 
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
      Height          =   390
      Left            =   5325
      TabIndex        =   2
      Top             =   3450
      Width           =   1965
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
      Height          =   390
      Left            =   5325
      TabIndex        =   1
      Top             =   2550
      Width           =   1965
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      X1              =   3525
      X2              =   7875
      Y1              =   5325
      Y2              =   5325
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      X1              =   7875
      X2              =   7875
      Y1              =   1950
      Y2              =   5325
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   3525
      X2              =   7875
      Y1              =   1950
      Y2              =   1950
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   3525
      X2              =   3525
      Y1              =   1950
      Y2              =   5325
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PassWord"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   4125
      TabIndex        =   5
      Top             =   3975
      Width           =   1065
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   4125
      TabIndex        =   4
      Top             =   3450
      Width           =   1140
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Company"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   4275
      TabIndex        =   0
      Top             =   2550
      Width           =   990
   End
End
Attribute VB_Name = "FormLogIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vDataBase As String, vServer As String

Private Sub Form_Load()
Dim HKey, wsho
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

Call InitializeDataBase
On Error GoTo Errdescription
TXTCompany.Text = "npvat"
Set wsho = CreateObject("Wscript.Shell")
        HKey = wsho.RegRead("HKCU\Software\BanChiang Soft\LOGIN_DEFAULT")
        vQuery = "select rtrim(SUBSTRING('" & Trim(HKey) & "', CHARINDEX('^','" & Trim(HKey) & "', 2)+1,100 )) as NameUser,rtrim(SUBSTRING('" & Trim(HKey) & "', 1,CHARINDEX('^','" & Trim(HKey) & "', 2)-1 )) as Company"
        If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
            vCompany1 = vRecordset.Fields("company").Value
            vUserID1 = vRecordset.Fields("nameuser").Value
        End If
    vRecordset.Close
'----------------------------------------------------------------------------------------------------------
If vCompany1 = "np" Then
vCompany1 = TXTCompany.Text
End If
'vUserID = TXTUserID.Text
TXTCompany.Text = vCompany1
TXTUser.Text = vUserID1
vPassword1 = Trim(TXTPassWord.Text)

Errdescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub TXTCompany_KeyPress(KeyAscii As Integer)
On Error GoTo CallError
If KeyAscii = 13 Then
    TXTUser.SetFocus
End If
CallError:
If Err.Description <> "" Then
MsgBox Err.Description
TXTPassWord.Text = ""
End If
End Sub

Private Sub TXTPassWord_KeyPress(KeyAscii As Integer)
On Error GoTo CallError
If KeyAscii = 13 Then

        vPassword = Trim(TXTPassWord.Text)
        vCompany = UCase(Trim(TXTCompany.Text))
        vUserID = Trim(TXTUser.Text)
        If vCompany = "NPVAT" Then
        vMemDatabase = "bcvat"
        vMemServer = "bi"
        ElseIf vCompany = "NPVAT53" Then
        vMemDatabase = "bcvat2010"
        vMemServer = "bi"
        End If
        Call InitializeDataBase
        Call InitializeDataBase1
        Call ConnectCompany
        
        AddOn_PrintBC.Order2.Enabled = True
        AddOn_PrintBC.Order3.Enabled = True
        AddOn_PrintBC.Order4.Enabled = True
        AddOn_PrintBC.Order5.Enabled = True
        AddOn_PrintBC.Order6.Enabled = True
        AddOn_PrintBC.Order7.Enabled = True
        AddOn_PrintBC.Order8.Enabled = True
        AddOn_PrintBC.Order9.Enabled = True
        AddOn_PrintBC.Order0.Enabled = True
        
        Unload Me

End If
CallError:
If Err.Description <> "" Then
MsgBox Err.Description
TXTPassWord.Text = ""
End If
End Sub

Private Sub TXTUser_KeyPress(KeyAscii As Integer)
On Error GoTo CallError
If KeyAscii = 13 Then
    'Call InitializeDataBase1
    TXTPassWord.SetFocus
End If
CallError:
If Err.Description <> "" Then
MsgBox Err.Description
TXTPassWord.Text = ""
End If
End Sub


Public Function ConnectCompany()

Select Case UCase(vCompany)
    Case "npvat"
        vMemDatabase = "BCVAT"
        vMemServer = "BI"
    Case "npvat47"
        vMemDatabase = "BCVAT47"
        vMemServer = "Solar"
    Case "npvat46"
        vMemDatabase = "BCVAT46A"
        vMemServer = "Solar"
End Select

vDataBase = "BCVAT"
vServer = "bi"
   
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

