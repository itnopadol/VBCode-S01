VERSION 5.00
Begin VB.Form Form000 
   Caption         =   "LogIn User"
   ClientHeight    =   9000
   ClientLeft      =   1545
   ClientTop       =   1005
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text102 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   5325
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3450
      Width           =   1890
   End
   Begin VB.CommandButton Command101 
      Caption         =   "LogIn"
      Height          =   465
      Left            =   5925
      TabIndex        =   2
      Top             =   4200
      Width           =   1290
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
      Height          =   360
      Left            =   5325
      TabIndex        =   0
      Top             =   2925
      Width           =   1890
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   165
      Left            =   4425
      TabIndex        =   4
      Top             =   3525
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "UserName"
      Height          =   165
      Left            =   4425
      TabIndex        =   3
      Top             =   2925
      Width           =   840
   End
End
Attribute VB_Name = "Form000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command101_Click()
Dim Connection As New ADODB.Connection
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error GoTo ErrDescription
If Text101.Text <> "" Then
vUserID = Trim(Text101.Text)
vPassword = Trim(Text102.Text)
vQuery = "Provider = SQLOLEDB.1;Persist Security Info = False;User ID = " & vUserID & ";Password = " & vPassword & ";Data Source = S02DB;Initial Catalog = BCNP"
Connection.Open (vQuery)
Call InitializeDatabase
Call InitializeSendEmail
Call UserCorrect
End If


ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Text102.SetFocus
    Text102.Text = ""
End If
End Sub

Public Sub UserCorrect()
MDIForm1.Order0.Enabled = True
MDIForm1.Order1.Enabled = True
MDIForm1.Order2.Enabled = True
MDIForm1.Order3.Enabled = True
MDIForm1.Order4.Enabled = True
MDIForm1.Order9.Enabled = True
MDIForm1.Caption = MDIForm1.Caption & " : " & vUserID
Form000.Visible = False
End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text102.SetFocus
End If
End Sub

Private Sub Text102_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Call Command101_Click
End If
End Sub
