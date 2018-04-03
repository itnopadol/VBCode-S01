VERSION 5.00
Begin VB.Form Form002 
   Caption         =   "Form1"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form002.frx":0000
   ScaleHeight     =   8760
   ScaleWidth      =   10950
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "ไม่ติดต่อระบบ"
      Height          =   540
      Left            =   3225
      TabIndex        =   4
      Top             =   2775
      Width           =   1665
   End
   Begin VB.TextBox Text2 
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
      Left            =   2175
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2175
      Width           =   2715
   End
   Begin VB.TextBox Text1 
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
      Left            =   2175
      TabIndex        =   0
      Top             =   1500
      Width           =   2715
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   1125
      TabIndex        =   3
      Top             =   2175
      Width           =   990
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   1125
      TabIndex        =   2
      Top             =   1500
      Width           =   990
   End
End
Attribute VB_Name = "Form002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
vConnect = 1
MDIForm1.Order1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIForm1.Order001.Enabled = True
End Sub
