VERSION 5.00
Begin VB.Form FrmLogIN 
   Caption         =   "เข้าสู่โปรแกรม"
   ClientHeight    =   7590
   ClientLeft      =   5325
   ClientTop       =   1980
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   7590
   ScaleWidth      =   11985
   WindowState     =   2  'Maximized
   Begin VB.ComboBox CMBZone 
      Height          =   315
      ItemData        =   "Form1.frx":D963
      Left            =   5490
      List            =   "Form1.frx":D965
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1485
      Width           =   1950
   End
   Begin VB.CommandButton CMDExit 
      Caption         =   "ยกเลิก"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6525
      TabIndex        =   3
      Top             =   3375
      Width           =   915
   End
   Begin VB.CommandButton CMDLogIN 
      Caption         =   "เข้าสู่ระบบ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5490
      TabIndex        =   2
      Top             =   3375
      Width           =   915
   End
   Begin VB.TextBox TextPassword 
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
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   5490
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2610
      Width           =   1950
   End
   Begin VB.TextBox TextUser 
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
      Height          =   330
      Left            =   5490
      TabIndex        =   0
      Top             =   2025
      Width           =   1950
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "โชนการทำงาน :"
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
      Height          =   285
      Left            =   4095
      TabIndex        =   6
      Top             =   1485
      Width           =   1320
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสผ่าน :"
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
      Height          =   330
      Left            =   3870
      TabIndex        =   5
      Top             =   2610
      Width           =   1545
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อเข้าโปรแกรม :"
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
      Height          =   375
      Left            =   3915
      TabIndex        =   4
      Top             =   2025
      Width           =   1500
   End
End
Attribute VB_Name = "FrmLogIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CMBZone_Click()
On Error Resume Next
Me.TextUser.SetFocus
End Sub

Private Sub CMDExit_Click()
MDIQueueManagement.Caption = Trim("โปรแกรม จัดการคิวจัดสินค้า")
Unload FrmLogIN
End Sub

Private Sub CMDLogIN_Click()
Dim Connection As New ADODB.Connection
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error GoTo ErrDescription

If TextUser.Text <> "" Then
vUserID = Trim(TextUser.Text)
vPassword = Trim(TextPassword.Text)
vQuery = "Provider = SQLOLEDB.1;Persist Security Info = False;User ID = " & vUserID & ";Password = " & vPassword & ";Data Source = NEBULA;Initial Catalog = BCNP"
Connection.Open (vQuery)


Select Case CMBZone.ListIndex
  Case 0:
    vSelectZoneID = 1
  Case 1:
    vSelectZoneID = 2
  Case 2:
   vSelectZoneID = 3
End Select

Call InitializeDataBase
Call InitializeDataBase1
Call InitializeDataBase2

Call ChekAuthorityAccess

If vUserAuthority = 0 Then
   MsgBox "คุณไม่สามารถเปิดใช้งานโปรแกรมได้ เนื่องจากไม่มีสิทธิ์ กรณีต้องการใช้กรุณาแจ้งห้องคอมพิวเตอร์", vbCritical, "Send Error Message"
   TextUser.Text = ""
   TextPassword.Text = ""
   TextUser.SetFocus
   Exit Sub
End If


Call UserCorrect
Unload FrmLogIN

Load FrmQueue
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    TextPassword.SetFocus
    TextPassword.Text = ""
End If
End Sub

Public Sub UserCorrect()
MDIQueueManagement.OrderProgram.Enabled = True
MDIQueueManagement.OrderReport.Enabled = True
MDIQueueManagement.OrderWindows.Enabled = True
MDIQueueManagement.Caption = MDIQueueManagement.Caption & " : " & vUserID
End Sub

Private Sub Form_Load()
CMBZone.AddItem Trim("สำนักงานใหญ่")
CMBZone.AddItem Trim("โกดังเหล็ก")
CMBZone.AddItem Trim("HomeMartMax")
CMBZone.Text = Trim("สำนักงานใหญ่")
End Sub

Private Sub TextPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If TextUser.Text <> "" Then
    Call CMDLogIN_Click
  Else
    TextUser.SetFocus
  End If
End If
End Sub

Private Sub TextUser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And TextUser.Text <> "" Then
  TextPassword.SetFocus
End If
End Sub
