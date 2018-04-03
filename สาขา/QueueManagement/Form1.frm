VERSION 5.00
Begin VB.Form FrmLogIN 
   Caption         =   "เข้าสู่โปรแกรม"
   ClientHeight    =   9600
   ClientLeft      =   2490
   ClientTop       =   1095
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   9600
   ScaleWidth      =   14880
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PICSelectPoint 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6180
      Left            =   270
      ScaleHeight     =   6150
      ScaleWidth      =   14295
      TabIndex        =   6
      Top             =   2115
      Visible         =   0   'False
      Width           =   14325
      Begin VB.CommandButton CMDPoint4 
         Caption         =   "จุดที่ 4 (D)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   10845
         TabIndex        =   11
         ToolTipText     =   "สินค้าโซน D"
         Top             =   1260
         Width           =   2355
      End
      Begin VB.CommandButton CMDPoint3 
         Caption         =   "จุดที่ 3 (C)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   7605
         TabIndex        =   10
         ToolTipText     =   "สินค้าโซน C"
         Top             =   1260
         Width           =   2355
      End
      Begin VB.CommandButton CMDPoint2 
         Caption         =   "จุดที่ 2 (B)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   4365
         TabIndex        =   9
         ToolTipText     =   "สินค้าโซน B"
         Top             =   1260
         Width           =   2355
      End
      Begin VB.CommandButton CMDPoint1 
         Caption         =   "จุดที่ 1 (A)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   1125
         TabIndex        =   8
         ToolTipText     =   "สินค้าโซน A"
         Top             =   1260
         Width           =   2355
      End
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
      Height          =   600
      Left            =   5130
      TabIndex        =   3
      Top             =   4230
      Width           =   1185
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
      Height          =   600
      Left            =   3780
      TabIndex        =   2
      Top             =   4230
      Width           =   1185
   End
   Begin VB.TextBox TextPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   3780
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3510
      Width           =   2535
   End
   Begin VB.TextBox TextUser 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3780
      TabIndex        =   0
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "กรอกรหัสเข้าใช้งานโปรแกรมและเลือกโซนการจัดสินค้า"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   222
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   90
      TabIndex        =   7
      Top             =   585
      Width           =   9555
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสผ่าน :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   465
      Left            =   2115
      TabIndex        =   5
      Top             =   3510
      Width           =   1545
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อเข้าโปรแกรม :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   1530
      TabIndex        =   4
      Top             =   2745
      Width           =   2130
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
vQuery = "Provider = SQLOLEDB.1;Persist Security Info = False;User ID = " & vUserID & ";Password = " & vPassword & ";Data Source = S02DB;Initial Catalog = BCNP"
Connection.Open (vQuery)


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

Me.PICSelectPoint.Visible = True
Me.CMDPoint1.SetFocus
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

Private Sub CMDPoint1_Click()
vSelectZoneID = 1
Call UserCorrect
Unload FrmLogIN
Load FrmQueue
End Sub

Private Sub CMDPoint1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 97 Then
'Call CMDPoint1_Click
End If

If KeyCode = 98 Then
Call CMDPoint2_Click
End If

If KeyCode = 99 Then
Call CMDPoint3_Click
End If

If KeyCode = 99 Then
Call CMDPoint4_Click
End If

If KeyCode = 27 Then
Me.PICSelectPoint.Visible = False
Me.TextUser.SetFocus
End If

End Sub

Private Sub CMDPoint2_Click()
vSelectZoneID = 2
Call UserCorrect
Unload FrmLogIN
Load FrmQueue
End Sub

Private Sub CMDPoint2_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 97 Then
'Call CMDPoint1_Click
End If

If KeyCode = 98 Then
Call CMDPoint2_Click
End If

If KeyCode = 99 Then
Call CMDPoint3_Click
End If

If KeyCode = 99 Then
Call CMDPoint4_Click
End If

If KeyCode = 27 Then
Me.PICSelectPoint.Visible = False
Me.TextUser.SetFocus
End If

End Sub

Private Sub CMDPoint3_Click()
vSelectZoneID = 3
Call UserCorrect
Unload FrmLogIN
Load FrmQueue
End Sub

Private Sub CMDPoint3_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 97 Then
'Call CMDPoint1_Click
End If

If KeyCode = 98 Then
Call CMDPoint2_Click
End If

If KeyCode = 99 Then
Call CMDPoint3_Click
End If

If KeyCode = 99 Then
Call CMDPoint4_Click
End If

If KeyCode = 27 Then
Me.PICSelectPoint.Visible = False
Me.TextUser.SetFocus
End If

End Sub

Private Sub CMDPoint4_Click()
vSelectZoneID = 4
Call UserCorrect
Unload FrmLogIN
Load FrmQueue
End Sub

Private Sub CMDPoint4_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 97 Then
'Call CMDPoint1_Click
End If

If KeyCode = 98 Then
Call CMDPoint2_Click
End If

If KeyCode = 99 Then
Call CMDPoint3_Click
End If

If KeyCode = 99 Then
Call CMDPoint4_Click
End If

If KeyCode = 27 Then
Me.PICSelectPoint.Visible = False
Me.TextUser.SetFocus
End If

End Sub

Private Sub PICSelectPoint_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 97 Then
'Call CMDPoint1_Click
End If

If KeyCode = 98 Then
Call CMDPoint2_Click
End If

If KeyCode = 99 Then
Call CMDPoint3_Click
End If

If KeyCode = 99 Then
Call CMDPoint4_Click
End If

If KeyCode = 27 Then
Me.PICSelectPoint.Visible = False
Me.TextUser.SetFocus
End If

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
