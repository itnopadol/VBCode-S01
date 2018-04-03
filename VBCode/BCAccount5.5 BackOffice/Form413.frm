VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form413 
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form413.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker DTP102 
      Height          =   390
      Left            =   2475
      TabIndex        =   3
      Top             =   2625
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   688
      _Version        =   393216
      Format          =   16646145
      CurrentDate     =   38318
   End
   Begin MSComCtl2.DTPicker DTP101 
      Height          =   390
      Left            =   2475
      TabIndex        =   2
      Top             =   2100
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   688
      _Version        =   393216
      Format          =   16646145
      CurrentDate     =   38318
   End
   Begin VB.CommandButton Cmd101 
      Caption         =   "ดูรายงาน"
      Height          =   540
      Left            =   3900
      TabIndex        =   1
      Top             =   3300
      Width           =   1290
   End
   Begin VB.ComboBox Cmb101 
      Height          =   315
      Left            =   2475
      TabIndex        =   0
      Top             =   1425
      Width           =   2715
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงวันที่"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   1575
      TabIndex        =   6
      Top             =   2625
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ณ วันที่"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   1575
      TabIndex        =   5
      Top             =   2100
      Width           =   540
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสเจ้าหนี้"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   1575
      TabIndex        =   4
      Top             =   1425
      Width           =   840
   End
End
Attribute VB_Name = "Form413"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

vQuery = "select code+'-'+name1 as apname from dbo.bcap order by code"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        CMB101.AddItem Trim(vRecordset.Fields("apname").Value)
        vRecordset.MoveNext
    Wend
End If
vRecordset.Close

Me.DTP101.Value = Now
Me.DTP102.Value = Now
End Sub
