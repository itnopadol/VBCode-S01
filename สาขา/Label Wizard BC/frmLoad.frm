VERSION 5.00
Begin VB.Form frmLoad 
   Caption         =   "Load Form"
   ClientHeight    =   1350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5280
      Top             =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Please wait While now loading..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   4695
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim n As Integer

Private Sub Form_Load()
n = 0
End Sub

Private Sub Timer1_Timer()
n = n + 1
If n = 1 Then
        frmWizard.Show
        frmWizard.txtUsername.SetFocus
        Unload Me
End If
End Sub
