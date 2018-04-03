VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form513 
   Caption         =   "รายงาน ยอดลูกหนี้ประจำเดือน"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form513.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   11985
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   420
      Left            =   6390
      TabIndex        =   6
      Top             =   3645
      Width           =   1500
   End
   Begin MSComCtl2.DTPicker DTPicker101 
      Height          =   330
      Left            =   4230
      TabIndex        =   5
      Top             =   2790
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   582
      _Version        =   393216
      Format          =   16646145
      CurrentDate     =   38812
   End
   Begin VB.ComboBox CMB102 
      Height          =   315
      Left            =   4230
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2205
      Width           =   3705
   End
   Begin VB.ComboBox CMB101 
      Height          =   315
      Left            =   4230
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1665
      Width           =   3705
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ณ วันที่ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   2295
      TabIndex        =   4
      Top             =   2790
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทลูกค้า :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   2700
      TabIndex        =   2
      Top             =   2205
      Width           =   1410
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จากประเภทลูกค้า :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   2655
      TabIndex        =   0
      Top             =   1665
      Width           =   1455
   End
End
Attribute VB_Name = "Form513"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

DTPicker101 = Now

vQuery = "exec bcvat.dbo.USP_AR_CustType "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    CMB101.Text = Trim(vRecordset.Fields("custname").Value)
    CMB102.Text = Trim(vRecordset.Fields("custname").Value)
    While Not vRecordset.EOF
        CMB101.AddItem Trim(vRecordset.Fields("custname").Value)
        CMB102.AddItem Trim(vRecordset.Fields("custname").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close
End Sub
