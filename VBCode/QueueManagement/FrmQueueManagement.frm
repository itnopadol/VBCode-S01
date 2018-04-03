VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmQueueManagement 
   Caption         =   "Form1"
   ClientHeight    =   7995
   ClientLeft      =   5325
   ClientTop       =   1605
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   7980
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   14076
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   706
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "คุมคิวจัดสินค้า"
      TabPicture(0)   =   "FrmQueueManagement.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Picture2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Picture3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "บันทึกผลการจัดสินค้า"
      TabPicture(1)   =   "FrmQueueManagement.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.PictureBox Picture4 
         Height          =   465
         Left            =   -75000
         Picture         =   "FrmQueueManagement.frx":0038
         ScaleHeight     =   405
         ScaleWidth      =   11835
         TabIndex        =   6
         Top             =   495
         Width           =   11895
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   45
         Picture         =   "FrmQueueManagement.frx":01C0
         ScaleHeight     =   390
         ScaleWidth      =   11865
         TabIndex        =   5
         Top             =   495
         Width           =   11895
      End
      Begin VB.PictureBox Picture2 
         Height          =   5325
         Left            =   4815
         ScaleHeight     =   5265
         ScaleWidth      =   6840
         TabIndex        =   2
         Top             =   1620
         Width           =   6900
         Begin MSComctlLib.ListView ListView1 
            Height          =   6045
            Left            =   90
            TabIndex        =   4
            Top             =   90
            Width           =   6675
            _ExtentX        =   11774
            _ExtentY        =   10663
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   5325
         Left            =   270
         ScaleHeight     =   5265
         ScaleWidth      =   4230
         TabIndex        =   1
         Top             =   1620
         Width           =   4290
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   3975
            Left            =   90
            TabIndex        =   3
            Top             =   810
            Width           =   4065
            _ExtentX        =   7170
            _ExtentY        =   7011
            _Version        =   393217
            Indentation     =   353
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            HotTracking     =   -1  'True
            Scroll          =   0   'False
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "FrmQueueManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim vFamily, vName, vDepartment As String
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vNode As Node

vFamily = "R-Queue"
vName = "คิวรอจัดสินค้า"
Set vNode = TreeView1.Nodes.Add(, , vFamily, vName)


vQuery = "select  distinct requestdate from npmaster.dbo.TB_NP_PickingQueueRequest "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    vFamily = "R-Queue"
    vName = Trim(vRecordset.Fields("requestdate").Value)
    vDepartment = Trim(vRecordset.Fields("requestdate").Value)
    Set vNode = TreeView1.Nodes.Add(vFamily, tvwChild, "R-" & vDepartment, vName)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close
End Sub
