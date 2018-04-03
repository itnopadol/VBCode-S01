VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCheckQTY 
   Caption         =   "ตรวจสอบสถานะการจัดสินค้า"
   ClientHeight    =   6945
   ClientLeft      =   5265
   ClientTop       =   3090
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmCheckQTY.frx":0000
   ScaleHeight     =   6945
   ScaleWidth      =   11970
   Begin VB.PictureBox Pic101 
      Height          =   1725
      Left            =   1350
      ScaleHeight     =   1665
      ScaleWidth      =   9090
      TabIndex        =   13
      Top             =   4590
      Visible         =   0   'False
      Width           =   9150
      Begin VB.CommandButton CMD102 
         Caption         =   "ยกเลิก"
         Height          =   375
         Left            =   7920
         TabIndex        =   23
         Top             =   1125
         Width           =   915
      End
      Begin VB.CommandButton CMD101 
         Caption         =   "ตกลง"
         Height          =   375
         Left            =   6795
         TabIndex        =   22
         Top             =   1125
         Width           =   915
      End
      Begin VB.TextBox TXTPicking 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3825
         TabIndex        =   21
         Top             =   585
         Width           =   1455
      End
      Begin VB.Label LBLItemName 
         Caption         =   "XXX"
         Height          =   330
         Left            =   1260
         TabIndex        =   18
         Top             =   90
         Width           =   7530
      End
      Begin VB.Label LBLUnitCode 
         Caption         =   "XXX"
         Height          =   285
         Left            =   6435
         TabIndex        =   20
         Top             =   585
         Width           =   2400
      End
      Begin VB.Label LBLQTY 
         Alignment       =   2  'Center
         Caption         =   "XXX"
         Height          =   285
         Left            =   1260
         TabIndex        =   19
         Top             =   585
         Width           =   1545
      End
      Begin VB.Label Label12 
         Caption         =   "หน่วย :"
         Height          =   240
         Left            =   5535
         TabIndex        =   17
         Top             =   585
         Width           =   645
      End
      Begin VB.Label Label3 
         Caption         =   "หยิบได้ :"
         Height          =   285
         Left            =   3060
         TabIndex        =   16
         Top             =   585
         Width           =   780
      End
      Begin VB.Label Label2 
         Caption         =   "ต้องการสินค้า :"
         Height          =   285
         Left            =   135
         TabIndex        =   15
         Top             =   585
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "ชื่อสินค้า :"
         Height          =   285
         Left            =   180
         TabIndex        =   14
         Top             =   90
         Width           =   780
      End
   End
   Begin VB.OptionButton OPT110 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "10.ปกติ"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1350
      TabIndex        =   37
      Top             =   5805
      Value           =   -1  'True
      Width           =   1680
   End
   Begin VB.OptionButton OPT109 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "9.รอรถโฟล์คลิฟ"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3240
      TabIndex        =   35
      Top             =   5400
      Width           =   1680
   End
   Begin VB.OptionButton OPT108 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "8.สินค้ามี 2 คลัง"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1350
      TabIndex        =   34
      Top             =   5400
      Width           =   1680
   End
   Begin VB.OptionButton OPT107 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "7.พนักงานไม่ได้กดคิวจัดสินค้า"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7245
      TabIndex        =   33
      Top             =   4995
      Width           =   3255
   End
   Begin VB.OptionButton OPT106 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "6.สินค้ามีหลายรายการ"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5130
      TabIndex        =   32
      Top             =   4995
      Width           =   1905
   End
   Begin VB.OptionButton OPT105 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "5.เครื่องพิมพ์ Error "
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3240
      TabIndex        =   31
      Top             =   4995
      Width           =   1680
   End
   Begin VB.OptionButton OPT104 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "4.เครื่องคอมฯ Error"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1350
      TabIndex        =   30
      Top             =   4995
      Width           =   1680
   End
   Begin VB.OptionButton OPT103 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "3.รอพนักงานประจำแผนกสินค้า"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7245
      TabIndex        =   29
      Top             =   4590
      Width           =   3255
   End
   Begin VB.OptionButton OPT102 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "2.จัดสินค้าพร้อมบิลอื่น"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5130
      TabIndex        =   28
      Top             =   4590
      Width           =   1905
   End
   Begin VB.OptionButton OPT101 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "1.สินค้าเป็นสีผสม"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3240
      TabIndex        =   27
      Top             =   4590
      Width           =   1680
   End
   Begin VB.TextBox TextDescription 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7245
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   5400
      Width           =   3255
   End
   Begin VB.CommandButton CMDCancel 
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
      Left            =   9585
      TabIndex        =   3
      Top             =   5895
      Width           =   915
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "ตกลง"
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
      Left            =   8550
      TabIndex        =   2
      Top             =   5895
      Width           =   915
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   2400
      Left            =   1350
      TabIndex        =   0
      Top             =   2115
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   4233
      View            =   3
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับ"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "รหัสสินค้า"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ชื่อสินค้า"
         Object.Width           =   7144
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "ต้องการ"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "หยิบได้"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "หน่วย"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "สถานะการจัด"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "คลัง"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เหตุผลอื่น ๆ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5130
      TabIndex        =   36
      Top             =   5400
      Width           =   1905
   End
   Begin VB.Label LBLDocDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5265
      TabIndex        =   26
      Top             =   900
      Width           =   1590
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่คิว :"
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
      Height          =   240
      Left            =   4365
      TabIndex        =   25
      Top             =   900
      Width           =   870
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "หมายเหตุการจัดสินค้า :"
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
      Left            =   1350
      TabIndex        =   24
      Top             =   4590
      Width           =   1770
   End
   Begin VB.Label LBLID 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2520
      TabIndex        =   12
      Top             =   1530
      Width           =   645
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "เอกสารชุดที่ :"
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
      Height          =   240
      Left            =   1350
      TabIndex        =   11
      Top             =   1530
      Width           =   1140
   End
   Begin VB.Label LBLArCode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2520
      TabIndex        =   10
      Top             =   1215
      Width           =   8025
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อลูกค้า :"
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
      Height          =   240
      Left            =   1665
      TabIndex        =   9
      Top             =   1215
      Width           =   825
   End
   Begin VB.Label LBLDocno2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   8550
      TabIndex        =   8
      Top             =   900
      Width           =   1995
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบสั่งขาย :"
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
      Height          =   240
      Left            =   7290
      TabIndex        =   7
      Top             =   900
      Width           =   1275
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "รายการสินค้า"
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
      Height          =   240
      Left            =   1350
      TabIndex        =   6
      Top             =   1890
      Width           =   1050
   End
   Begin VB.Label LBLDocno1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2520
      TabIndex        =   5
      Top             =   900
      Width           =   1725
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่คิว :"
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
      Height          =   240
      Left            =   1395
      TabIndex        =   4
      Top             =   900
      Width           =   1095
   End
End
Attribute VB_Name = "FrmCheckQTY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vQuery As String
Dim vIndex As Integer

Private Sub CMD101_Click()
Dim vPickQTY As String

On Error Resume Next

If LBLItemName.Caption <> "" And TXTPicking <> "" Then
  vPickQTY = CCur(TXTPicking.Text)
  vIndex = ListView101.SelectedItem.Index
  ListView101.ListItems.Item(vIndex).SubItems(4) = Format(vPickQTY, "##,##0.00")
  TXTPicking.Enabled = False
  LBLItemName.Caption = ""
  LBLQTY.Caption = ""
  LBLUnitCode.Caption = ""
  TXTPicking.Text = ""
  ListView101.SetFocus
  Pic101.Visible = False
Else
  MsgBox "กรุณากรอกจำนวนที่หยิบได้ด้วย", vbCritical, "Send Error"
End If

End Sub

Private Sub CMD101_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
  Call CMDCancel_Click
End If
If KeyCode = 116 Then
  Call CMDOK_Click
End If
End Sub

Private Sub CMD102_Click()
  Pic101.Visible = False
End Sub

Private Sub CMDCancel_Click()
FrmQueue.Text102.SetFocus
FrmQueue.Text102 = ""
FrmQueue.Text102.SetFocus
Call FrmQueue.StartTime
Unload FrmCheckQTY
End Sub

Private Sub CMDCancel_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
  Call CMDCancel_Click
End If
If KeyCode = 116 Then
  Call CMDOK_Click
End If
End Sub

Private Sub CMDOK_Click()
Dim vPickingNo As String
Dim vSaleOrderNo As String
Dim vItemCode As String
Dim vItemName As String
Dim vWHCode As String
Dim vQTY As Double
Dim vPickQTY As Double
Dim vUnitCode As String
Dim vPickItemStatus As String
Dim i As Integer
Dim vLineNumber As Integer
Dim vCheckPickQTY As Integer
Dim vDocDate As String
Dim vDescription As String
Dim vZoneID As String
Dim vRecordset As New ADODB.Recordset
Dim vPickReason As Integer


vPickingNo = Trim(LBLDocno1.Caption)
vSaleOrderNo = Trim(LBLDocno2.Caption)

vQuery = "select zoneid from npmaster.dbo.tb_np_queuemanagement where docno = '" & vPickingNo & "' order by docdate desc "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
vZoneID = vRecordset.Fields("zoneid")
End If
vRecordset.Close

If vPickingNo <> "" Then
  vCheckPickQTY = 1
  vDocDate = Me.LBLDocDate.Caption         'CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
  vDescription = TextDescription.Text
  
  
  If Me.OPT101.Value = True Then
    vPickReason = 1
  ElseIf Me.OPT102.Value = True Then
    vPickReason = 2
  ElseIf Me.OPT103.Value = True Then
    vPickReason = 3
  ElseIf Me.OPT104.Value = True Then
    vPickReason = 4
  ElseIf Me.OPT105.Value = True Then
    vPickReason = 5
  ElseIf Me.OPT106.Value = True Then
    vPickReason = 6
  ElseIf Me.OPT107.Value = True Then
    vPickReason = 7
  ElseIf Me.OPT108.Value = True Then
    vPickReason = 8
  ElseIf Me.OPT109.Value = True Then
    vPickReason = 9
    Else
      vPickReason = 0
  End If
  
  'On Error GoTo ErrDescription
  'vQuery = "begin tran"
  'vConnection.Execute vQuery
  
  vQuery = "exec dbo.USP_NP_UpdatePrintStatusQueueManagement1 '" & vPickingNo & "','" & vDocDate & "','',2," & vTimeID & "," & vCheckPickQTY & " "
  vConnection.Execute vQuery

  vQuery = "exec dbo.USP_NP_UpdateMydescriptionQueueManagement2 '" & vPickingNo & "','" & vDocDate & "'," & vTimeID & "," & vPickReason & ",'" & vDescription & "' "
  vConnection.Execute vQuery
  
  vLineNumber = -1
  
  vQuery = "exec dbo.USP_NP_DeleteQueueItemSub " & vPickingNo & ",'" & vDocDate & "'," & vTimeID & " "
  vConnection.Execute vQuery
  
  For i = 1 To ListView101.ListItems.Count
    vLineNumber = vLineNumber + 1
    vItemCode = Trim(ListView101.ListItems.Item(i).SubItems(1))
    vItemName = Trim(ListView101.ListItems.Item(i).SubItems(2))
    vQTY = Trim(ListView101.ListItems.Item(i).SubItems(3))
    vPickQTY = Trim(ListView101.ListItems.Item(i).SubItems(4))
    vUnitCode = Trim(ListView101.ListItems.Item(i).SubItems(5))
    vPickItemStatus = Trim(ListView101.ListItems.Item(i).SubItems(6))
    vWHCode = Trim(ListView101.ListItems.Item(i).SubItems(7))
    If vQTY - vPickQTY > 0 Then
      vCheckPickQTY = 2
    End If
    
    If vQTY - vPickQTY < 0 Then
      vCheckPickQTY = 3
    End If
    
    vQuery = "exec dbo.USP_NP_InsertQueueManagementSub1 '" & vPickingNo & "','" & vItemCode & "','" & vItemName & "','" & vWHCode & "'," & vQTY & "," & vPickQTY & ",'" & vUnitCode & "','" & vPickItemStatus & "'," & vLineNumber & "," & vTimeID & ",'" & vDocDate & "','" & vPickingNo & "' "
    vConnection.Execute vQuery
  Next i
  
  vQuery = "exec dbo.USP_NP_UpdatePrintStatusQueueManagement1 '" & vPickingNo & "','" & vDocDate & "','',2," & vTimeID & "," & vCheckPickQTY & " "
  vConnection.Execute vQuery
  
  vQuery = "exec dbo.USP_NP_InsertQueueSpeech " & vPickingNo & ",2," & vCheckPickQTY & ",'" & vZoneID & "' "
  vConnection.Execute vQuery
  
  'vQuery = "commit tran"
  'vConnection.Execute vQuery
  

'ErrDescription:
'If Err.Description <> "" Then
  'vQuery = "rollback tran"
  'vConnection.Execute vQuery
  'MsgBox Err.Description
  'Exit Sub
'End If

End If

Call RefreshQueuePicking
Call RefreshQueueFinish

FrmQueue.Text102.Text = ""
Unload FrmCheckQTY
Call FrmQueue.StartTime
FrmQueue.ListView104.SetFocus

End Sub


Private Sub CMDOK_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
  Call CMDCancel_Click
End If
If KeyCode = 116 Then
  Call CMDOK_Click
End If
End Sub

Private Sub Form_Load()
Call FrmQueue.StopTime
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call FrmQueue.StartTime
Unload FrmCheckQTY
End Sub

Private Sub ListView101_DblClick()
Dim vIndex As Integer

On Error Resume Next

If ListView101.ListItems.Count > 0 Then
  Pic101.Visible = True
  vIndex = ListView101.SelectedItem.Index
  LBLItemName.Caption = Trim(ListView101.ListItems.Item(vIndex).SubItems(2))
  LBLQTY.Caption = Trim(ListView101.ListItems.Item(vIndex).SubItems(3))
  LBLUnitCode.Caption = Trim(ListView101.ListItems.Item(vIndex).SubItems(5))
  TXTPicking.Enabled = True
  TXTPicking.SetFocus
End If
End Sub

Private Sub ListView101_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
  Call CMDCancel_Click
End If
If KeyCode = 116 Then
  Call CMDOK_Click
End If
End Sub

Private Sub ListView101_KeyPress(KeyAscii As Integer)
Dim vIndex As Integer

On Error Resume Next

If KeyAscii = 13 Then
  If ListView101.ListItems.Count > 0 Then
    Pic101.Visible = True
    vIndex = ListView101.SelectedItem.Index
    LBLItemName.Caption = Trim(ListView101.ListItems.Item(vIndex).SubItems(2))
    LBLQTY.Caption = Trim(ListView101.ListItems.Item(vIndex).SubItems(3))
    LBLUnitCode.Caption = Trim(ListView101.ListItems.Item(vIndex).SubItems(5))
    TXTPicking.Enabled = True
    TXTPicking.SetFocus
  End If
End If
End Sub


Private Sub TextDescription_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
  Call CMDCancel_Click
End If
If KeyCode = 116 Then
  Call CMDOK_Click
End If
End Sub

Private Sub TXTPicking_KeyPress(KeyAscii As Integer)
Dim vPickQTY As String

On Error Resume Next

If KeyAscii = 13 Then

If LBLItemName.Caption <> "" And TXTPicking <> "" Then
  vIndex = ListView101.SelectedItem.Index
  vPickQTY = CCur(TXTPicking.Text)
  ListView101.ListItems.Item(vIndex).SubItems(4) = Format(vPickQTY, "##,##0.00")
  TXTPicking.Enabled = False
  LBLItemName.Caption = ""
  LBLQTY.Caption = ""
  LBLUnitCode.Caption = ""
  TXTPicking.Text = ""
  ListView101.SetFocus
  Pic101.Visible = False
Else
  MsgBox "กรุณากรอกจำนวนที่หยิบได้ด้วย", vbCritical, "Send Error"
End If
End If
End Sub
