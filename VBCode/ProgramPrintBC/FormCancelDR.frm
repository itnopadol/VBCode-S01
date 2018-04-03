VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormCancelDR 
   Caption         =   "หน้าจอยกเลิกเอกสาร DR"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FormCancelDR.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7305
      Left            =   480
      ScaleHeight     =   7275
      ScaleWidth      =   14265
      TabIndex        =   6
      Top             =   1380
      Width           =   14295
      Begin VB.CommandButton CMDSearchDR 
         Caption         =   "Command1"
         Height          =   645
         Left            =   5190
         TabIndex        =   8
         Top             =   780
         Width           =   1035
      End
      Begin MSComctlLib.ListView ListView101 
         Height          =   2475
         Left            =   1950
         TabIndex        =   7
         Top             =   3180
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   4366
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   765
      Left            =   4590
      TabIndex        =   5
      Top             =   6330
      Width           =   1635
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   645
      Left            =   2340
      TabIndex        =   4
      Top             =   6420
      Width           =   1545
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2595
      Left            =   1410
      TabIndex        =   3
      Top             =   3090
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   4577
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox Text1 
      Height          =   435
      Left            =   2730
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1440
      Width           =   2445
   End
   Begin VB.CommandButton CMDSelectDR 
      Caption         =   "รายละเอียด"
      Height          =   525
      Left            =   5370
      TabIndex        =   0
      Top             =   1410
      Width           =   1485
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   435
      Left            =   780
      TabIndex        =   2
      Top             =   1470
      Width           =   1305
   End
End
Attribute VB_Name = "FormCancelDR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



