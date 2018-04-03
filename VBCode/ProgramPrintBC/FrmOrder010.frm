VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form FrmOrder010 
   Caption         =   "Form010 รายการจัดคิวขนส่ง"
   ClientHeight    =   7635
   ClientLeft      =   1935
   ClientTop       =   870
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmOrder010.frx":0000
   ScaleHeight     =   7635
   ScaleWidth      =   11955
   WindowState     =   2  'Maximized
   Begin VB.CheckBox CheckDueDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "วันที่นัดรับ :"
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
      Height          =   285
      Left            =   2205
      TabIndex        =   13
      Top             =   900
      Width           =   1320
   End
   Begin VB.CommandButton CMDCalcInvoice 
      Height          =   420
      Left            =   9045
      Picture         =   "FrmOrder010.frx":72FB
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1350
      Width           =   420
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
      Height          =   420
      Left            =   6660
      TabIndex        =   11
      Top             =   1350
      Width           =   2310
   End
   Begin VB.CommandButton CMD103 
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
      Height          =   600
      Left            =   10440
      TabIndex        =   8
      Top             =   6075
      Width           =   1095
   End
   Begin VB.CommandButton CMD101 
      Caption         =   "ดึงใบจัดคิว"
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
      Left            =   10215
      TabIndex        =   4
      Top             =   900
      Width           =   1320
   End
   Begin MSComCtl2.DTPicker DTPicker101 
      Height          =   330
      Left            =   3600
      TabIndex        =   2
      Top             =   900
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   64684033
      CurrentDate     =   38696
   End
   Begin VB.CommandButton CMD102 
      Caption         =   "เลือก"
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
      Left            =   8910
      TabIndex        =   7
      Top             =   6075
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   3705
      Left            =   315
      TabIndex        =   6
      Top             =   2205
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   6535
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับ"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "เลขที่ใบจัดคิว"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "รหัสสินค้า"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ชื่อสินค้า"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "จำนวน"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "หน่วย"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "วันที่นัดส่ง"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "เวลานัดส่ง"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "เส้นทาง"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "QueueID"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "InvoiceNo"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.CheckBox Check103 
      BackColor       =   &H8000000E&
      Caption         =   ": ไม่ใส่เงื่อนไข"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   315
      TabIndex        =   1
      Top             =   1260
      Width           =   1500
   End
   Begin VB.CheckBox Check102 
      BackColor       =   &H8000000E&
      Caption         =   ": Select All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   315
      TabIndex        =   5
      Top             =   1620
      Width           =   1500
   End
   Begin VB.CheckBox Check101 
      BackColor       =   &H8000000E&
      Caption         =   ": POS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   315
      TabIndex        =   0
      Top             =   900
      Width           =   1500
   End
   Begin VB.ComboBox CMB101 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6660
      TabIndex        =   3
      Top             =   900
      Width           =   3210
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่เอกสาร :"
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
      Height          =   285
      Left            =   5175
      TabIndex        =   10
      Top             =   1440
      Width           =   1320
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เส้นทางขนส่ง :"
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
      Left            =   5265
      TabIndex        =   9
      Top             =   900
      Width           =   1320
   End
End
Attribute VB_Name = "FrmOrder010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vSortResult As Integer
Dim vCheckSelect  As Integer
Dim vSelectOK As Integer

Private Sub Check102_Click()
Dim i As Integer

On Error GoTo ErrDescription

For i = 1 To ListView101.ListItems.Count
    If Check102.Value = 1 Then
        ListView101.ListItems.Item(i).Checked = True
    Else
        ListView101.ListItems.Item(i).Checked = False
    End If
Next i

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub Check103_Click()
On Error GoTo ErrDescription

If Check103.Value = 1 Then
    CMB101.Enabled = False
    DTPicker101.Enabled = False
ElseIf Check103.Value = 0 Then
    CMB101.Enabled = True
    DTPicker101.Enabled = True
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDueDate As String
Dim vRoute As String
Dim vRouteID As String
Dim vPosStatus As Integer
Dim vListSearch As ListItem
Dim vDate As String
Dim i As Integer
Dim vInvoiceNo As String

'On Error GoTo ErrDescription

Check102.Value = 0
If Check101.Value = 1 Then
    vPosStatus = 1
Else
    vPosStatus = 0
End If
If Check103.Value = 0 Then

If Me.CheckDueDate.Value = 1 Then
vDueDate = Trim(DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year)
Else
vDueDate = ""
End If

vInvoiceNo = Trim(Text101.Text)
    If CMB101.Text <> "" Then
        vRoute = Trim(CMB101.Text)
        vQuery = "select id from npmaster.dbo.tb_do_route where name1 = '" & vRoute & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRouteID = Trim(vRecordset.Fields("id").Value)
        Else
            MsgBox "ไม่มีข้อมูลเส้นทางขนส่ง " & vRoute & " นี้ในระบบ กรุณาตรวจสอบ", vbCritical, "Send Error"
        End If
        vRecordset.Close
        
        If vRouteID = "" Then
        Exit Sub
        End If
    
    Else
        vRouteID = "Null"
    End If
    vQuery = "exec bcnp.dbo.USP_DO_QueueList " & vPosStatus & ",'" & vDueDate & "'," & vRouteID & ",'" & vInvoiceNo & "' "
Else
    vRouteID = "Null"
    vDate = "Null"
    vQuery = "exec bcnp.dbo.USP_DO_QueueList " & vPosStatus & ",'" & vDate & "'," & vRouteID & ",'" & vInvoiceNo & "' "
End If
i = 1
ListView101.ListItems.Clear
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vRecordset.MoveFirst
While Not vRecordset.EOF
Set vListSearch = ListView101.ListItems.Add(, , Trim(i))
vListSearch.SubItems(1) = Trim(vRecordset.Fields("queuesubid").Value)
vListSearch.SubItems(2) = Trim(vRecordset.Fields("queuedocno").Value)
vListSearch.SubItems(3) = Trim(vRecordset.Fields("itemcode").Value)
vListSearch.SubItems(4) = Trim(vRecordset.Fields("itemname").Value)
vListSearch.SubItems(5) = Trim(vRecordset.Fields("remainqty").Value)
vListSearch.SubItems(6) = Trim(vRecordset.Fields("unitcode").Value)
vListSearch.SubItems(7) = Trim(vRecordset.Fields("duedate").Value)
vListSearch.SubItems(8) = Trim(vRecordset.Fields("duetime").Value)
vListSearch.SubItems(9) = Trim(vRecordset.Fields("routename").Value)
vListSearch.SubItems(10) = Trim(vRecordset.Fields("queueid").Value)
vListSearch.SubItems(11) = Trim(vRecordset.Fields("invoiceno").Value)
vRecordset.MoveNext
i = i + 1
Wend
Else
    MsgBox "ไม่มีข้อมูล ตามคำที่ใช้ค้นหา", vbInformation, "Send Information"
End If
vRecordset.Close

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD102_Click()
Dim i As Integer
Dim vListSelect As ListItem
Dim j As Integer
Dim m As Integer
Dim vCheckLineQueueID As String
Dim vSelectQueueID As String
Dim vCheckItemCode1 As String
Dim vCheckItemCode2 As String
Dim vCheckUnitCode1 As String
Dim vCheckUnitCode2 As String
Dim vCheckQTY1 As Currency
Dim vCheckQTY2 As Currency
Dim vCheckInvoice1 As String
Dim vCheckInvoice2 As String

On Error GoTo ErrDescription

If ListView101.ListItems.Count <> 0 Then
    m = FormDelivery.ListView101.ListItems.Count
    For i = 1 To ListView101.ListItems.Count
        If ListView101.ListItems.Item(i).Checked = True Then
            vCheckLineQueueID = Trim(ListView101.ListItems.Item(i).SubItems(1))
            vCheckItemCode1 = Trim(ListView101.ListItems.Item(i).SubItems(3))
            vCheckUnitCode1 = Trim(ListView101.ListItems.Item(i).SubItems(6))
            vCheckQTY1 = Trim(ListView101.ListItems.Item(i).SubItems(5))
            vCheckInvoice1 = Trim(ListView101.ListItems.Item(i).SubItems(11))
            For j = 1 To FormDelivery.ListView101.ListItems.Count
                vSelectQueueID = Trim(FormDelivery.ListView101.ListItems.Item(j).SubItems(1))
                vCheckItemCode2 = Trim(FormDelivery.ListView101.ListItems.Item(j).SubItems(2))
                vCheckUnitCode2 = Trim(FormDelivery.ListView101.ListItems.Item(j).SubItems(6))
                vCheckQTY2 = Trim(FormDelivery.ListView101.ListItems.Item(j).SubItems(4))
                vCheckInvoice2 = Trim(FormDelivery.ListView101.ListItems.Item(j).SubItems(9))
            If vCheckItemCode1 = vCheckItemCode2 Then
                If vCheckInvoice1 = vCheckInvoice2 Then
                    If vCheckUnitCode1 = vCheckUnitCode2 Then
                        If vCheckQTY1 = vCheckQTY2 Then
                            vSelectOK = 0
                            GoTo Line1
                        Else
                            vSelectOK = 1
                        End If
                    Else
                        vSelectOK = 1
                    End If
                Else
                    vSelectOK = 1
                End If
                
                'vSelectOK = 0
                'GoTo Line1
            Else
                vSelectOK = 1
            End If
            Next j
            If FormDelivery.ListView101.ListItems.Count = 0 Then
                vSelectOK = 1
            End If
    
        
Line1:
            If vSelectOK = 1 Then
                m = m + 1
                Set vListSelect = FormDelivery.ListView101.ListItems.Add(, , Trim(m))
                vListSelect.SubItems(1) = Trim(ListView101.ListItems.Item(i).SubItems(1))
                vListSelect.SubItems(2) = Trim(ListView101.ListItems.Item(i).SubItems(3))
                vListSelect.SubItems(3) = Trim(ListView101.ListItems.Item(i).SubItems(4))
                vListSelect.SubItems(4) = Trim(ListView101.ListItems.Item(i).SubItems(5))
                vListSelect.SubItems(5) = 0
                vListSelect.SubItems(6) = Trim(ListView101.ListItems.Item(i).SubItems(6))
                vListSelect.SubItems(8) = Trim(ListView101.ListItems.Item(i).SubItems(10))
                vListSelect.SubItems(9) = Trim(ListView101.ListItems.Item(i).SubItems(11))
            End If
        End If
    Next i
    Unload FrmOrder010
End If
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD103_Click()
Unload FrmOrder010
End Sub

Private Sub CMDCalcInvoice_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String


If Text101.Text <> "" Then
  vDocNo = Trim(Text101.Text)
  vQuery = "exec dbo.USP_DO_QueueRemain '" & vDocNo & "' "
  gConnection.Execute vQuery
  MsgBox "ได้ทำการคำณวนเอกสารเรียบร้อยแล้วครับ", vbInformation, "Send Information"
  Text101.Text = ""
Else
  MsgBox "กรุณากรอกเลขที่เอกสารด้วย", vbInformation, "Send Error"
End If

End Sub

Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error GoTo ErrDescription

DTPicker101 = Now
vQuery = "select  name1  from NPMaster.dbo.TB_DO_Route order by name1"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        CMB101.AddItem Trim(vRecordset.Fields("name1").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIFrmProgramPrint.DO1.Enabled = True
FormDelivery.Enabled = True
End Sub

Private Sub ListView101_DblClick()
On Error Resume Next

If ListView101.ListItems.Count <> 0 Then
    ListView101.ListItems.Item(ListView101.SelectedItem.Index).Checked = True
End If
End Sub

Private Sub ListView101_KeyPress(KeyAscii As Integer)
Dim i As Integer

On Error GoTo ErrDescription

If ListView101.ListItems.Count <> 0 Then
    If KeyAscii = 32 Then
        If vCheckSelect = 0 Then
            For i = ListView101.ListItems.Count To 1 Step -1
                If ListView101.ListItems(i).Selected = True Then
                    ListView101.ListItems(i).Checked = True
                End If
            Next i
            vCheckSelect = 1
        ElseIf vCheckSelect = 1 Then
            For i = ListView101.ListItems.Count To 1 Step -1
                If ListView101.ListItems(i).Selected = True Then
                    ListView101.ListItems(i).Checked = False
                End If
            Next i
            vCheckSelect = 0
        End If
    End If
End If
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Public Sub CheckSelectQueueID()
Dim vCheckLineQueueID As String
Dim vSelectQueueID As String
Dim i As Integer
Dim j As Integer

On Error GoTo ErrDescription

For i = 1 To ListView101.ListItems.Count
    If ListView101.ListItems.Item(i).Checked = True Then
        vCheckLineQueueID = Trim(ListView101.ListItems.Item(i).SubItems(5))
        For j = 1 To FormDelivery.ListView101.ListItems.Count
            vSelectQueueID = Trim(FormDelivery.ListView101.ListItems.Item(j).SubItems(1))
        Next j
        If vCheckLineQueueID = vSelectQueueID Then
            vSelectOK = 0
        Else
            vSelectOK = 1
        End If
    End If
Next i

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

