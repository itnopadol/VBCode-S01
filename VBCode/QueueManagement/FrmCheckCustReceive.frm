VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCheckCustReceive 
   BackColor       =   &H00C0C0C0&
   Caption         =   "ตรวจสอบคิวที่ลูกค้ารับของกับใบจ่ายสินค้า"
   ClientHeight    =   8100
   ClientLeft      =   5130
   ClientTop       =   1605
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8100
   ScaleWidth      =   12090
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar PBarSearchQueue 
      Height          =   330
      Left            =   2655
      TabIndex        =   5
      Top             =   1395
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton CMDRefresh 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ฟื้นฟูข้อมูล"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1395
      Width           =   2265
   End
   Begin VB.CommandButton CMDCancel 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ออก"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10575
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6570
      Width           =   1140
   End
   Begin VB.CommandButton CMDSave 
      BackColor       =   &H00C0C0C0&
      Caption         =   "บันทึก"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9315
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6570
      Width           =   1140
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   4650
      Left            =   360
      TabIndex        =   0
      Top             =   1800
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   8202
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "เลขที่คิว"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "วันที่"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "เลขที่บิล"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ชื่อลูกค้า"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "อ้างอิง"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "คลัง"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่คิวที่ลูกค้ายังไม่ได้รับของ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   1
      Top             =   1035
      Width           =   2940
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   45
      Picture         =   "FrmCheckCustReceive.frx":0000
      Top             =   45
      Width           =   2160
   End
End
Attribute VB_Name = "FrmCheckCustReceive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vQuery As String

Private Sub CMDCancel_Click()
Unload FrmCheckCustReceive
End Sub

Private Sub CMDRefresh_Click()
 Call vGetQueueFinish
End Sub

Private Sub CMDSave_Click()
Dim i As Integer
Dim vQueueNo As String
Dim vSaleOrderNo As String
Dim vStatus As Integer
Dim vDescription As String

On Error GoTo ErrDescription

For i = 1 To Me.ListView101.ListItems.Count
   If Me.ListView101.ListItems.Item(i).Checked = True Then
      vQueueNo = Me.ListView101.ListItems(i).Text
      vSaleOrderNo = Me.ListView101.ListItems(i).SubItems(4)
      vStatus = 1
      vDescription = ""
      vQuery = "exec dbo.USP_NP_UpdateQueueReceivedStatus1 '" & vQueueNo & "','" & vSaleOrderNo & "'," & vStatus & ",'" & vDescription & "' "
      vConnection.Execute vQuery
   End If
Next i

Call vGetQueueFinish

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Form_Load()
Call vGetQueueFinish
End Sub

Public Sub vGetQueueFinish()
Dim vRecordset As New ADODB.Recordset
Dim vListQueue As ListItem

On Error Resume Next

Me.PBarSearchQueue.Value = 0
Me.ListView101.ListItems.Clear
vQuery = "exec dbo.USP_QM_SearchQueueFinishZone1 3"
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
   vRecordset.MoveFirst
   Me.PBarSearchQueue.Max = vRecordset.RecordCount
   While Not vRecordset.EOF
      Set vListQueue = Me.ListView101.ListItems.Add(, , vRecordset.Fields("docno").Value)
      vListQueue.SubItems(1) = Trim(vRecordset.Fields("docdate").Value)
      vListQueue.SubItems(2) = Trim(vRecordset.Fields("Invoiceno").Value)
      vListQueue.SubItems(3) = Trim(vRecordset.Fields("arname").Value)
      vListQueue.SubItems(4) = Trim(vRecordset.Fields("saleorderno").Value)
      vListQueue.SubItems(5) = Trim(vRecordset.Fields("whcode").Value)
      vRecordset.MoveNext
      Me.PBarSearchQueue.Value = Me.PBarSearchQueue.Value + 1
   Wend
End If
vRecordset.Close

End Sub
