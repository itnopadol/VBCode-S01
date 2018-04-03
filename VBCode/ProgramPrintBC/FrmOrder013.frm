VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmOrder013 
   Caption         =   "Form013 : ค้นหาใบขนส่งสินค้า"
   ClientHeight    =   6795
   ClientLeft      =   6405
   ClientTop       =   1485
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmOrder013.frx":0000
   ScaleHeight     =   6795
   ScaleWidth      =   9765
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
      Height          =   420
      Left            =   8595
      TabIndex        =   5
      Top             =   5625
      Width           =   870
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
      Height          =   420
      Left            =   7470
      TabIndex        =   3
      Top             =   5625
      Width           =   870
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   3930
      Left            =   270
      TabIndex        =   2
      Top             =   1575
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   6932
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
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
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "เลขที่เอกสาร"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "วันที่เอกสาร"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "วันที่ส่ง"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "เวลาส่ง"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "วันที่กลับ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "เวลาที่กลับ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "กม. ที่เริ่ม"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "กม. ที่สิ้นสุด"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "คำอธิบาย"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "หมายเลขรถ"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "ยกเลิก"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.CommandButton CMD101 
      Height          =   330
      Left            =   3600
      Picture         =   "FrmOrder013.frx":72FB
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "ค้นหา ตามเงื่อนไขที่กรอก"
      Top             =   990
      Width           =   330
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1305
      TabIndex        =   0
      Top             =   990
      Width           =   2265
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ค้นหาเอกสาร"
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
      Left            =   270
      TabIndex        =   4
      Top             =   1035
      Width           =   1095
   End
End
Attribute VB_Name = "FrmOrder013"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vSearch As String
Dim vSearchList As ListItem
Dim i As Integer

On Error Resume Next

vSearch = Trim(Text101.Text)
ListView101.ListItems.Clear
vQuery = "exec bcnp.dbo.USP_DO_DeliverySearchHeader '" & vSearch & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    i = 1
    While Not vRecordset.EOF
        Set vSearchList = ListView101.ListItems.Add(, , Trim(i))
        vSearchList.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
        vSearchList.SubItems(2) = Trim(vRecordset.Fields("docdate").Value)
        vSearchList.SubItems(3) = Trim(vRecordset.Fields("DateSend").Value)
        vSearchList.SubItems(4) = Trim(vRecordset.Fields("TimeSend").Value)
        vSearchList.SubItems(5) = Trim(vRecordset.Fields("DateReturn").Value)
        vSearchList.SubItems(6) = Trim(vRecordset.Fields("TimeReturn").Value)
        vSearchList.SubItems(7) = Trim(vRecordset.Fields("MeasureStart").Value)
        vSearchList.SubItems(8) = Trim(vRecordset.Fields("MeasureStop").Value)
        vSearchList.SubItems(9) = Trim(vRecordset.Fields("MyDescription").Value)
        vSearchList.SubItems(10) = Trim(vRecordset.Fields("VehicalID").Value)
        vSearchList.SubItems(11) = Trim(vRecordset.Fields("iscancel").Value)
    vRecordset.MoveNext
    i = i + 1
    Wend
    Me.ListView101.SetFocus
    Else
        MsgBox "ไม่มีข้อมูล ตามคำที่ใช้ค้นหา", vbInformation, "Send Information"
    End If
    vRecordset.Close

End Sub

Private Sub CMD102_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vSearch As String
Dim vSearchList As ListItem
Dim vEmpListSelect  As ListItem
Dim i As Integer
Dim j As Integer

On Error GoTo ErrDescription

If ListView101.ListItems.Count <> 0 Then
    FormDelivery.ListView101.ListItems.Clear
    vIsOpen2 = 1
    vSearch = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
    vQuery = "exec bcnp.dbo.USP_DO_DeliverySearchDetail '" & vSearch & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        i = 1
        If Trim(vRecordset.Fields("iscancel").Value) = 1 Then
            FormDelivery.Image101.Visible = False
            FormDelivery.Image102.Visible = True
        Else
            FormDelivery.Image101.Visible = True
            FormDelivery.Image102.Visible = False
        End If
        FormDelivery.DTPicker101.Value = Trim(vRecordset.Fields("docdate").Value)
        FormDelivery.DTPicker102.Value = Trim(vRecordset.Fields("datesend").Value)
        FormDelivery.DTPicker103.Value = Now
        FormDelivery.Text101.Text = Trim(vRecordset.Fields("id").Value)
        FormDelivery.Text102.Text = Trim(vRecordset.Fields("docno").Value)
        FormDelivery.Text103.Text = Trim(vRecordset.Fields("VehicalID").Value)
        FormDelivery.Text104.Text = Trim(vRecordset.Fields("CarNo").Value)
        FormDelivery.Text105.Text = Trim(vRecordset.Fields("CarLicence").Value)
        FormDelivery.Text106.Text = Trim(vRecordset.Fields("MeasureStart").Value)
        FormDelivery.Text107.Text = Trim(vRecordset.Fields("MeasureStop").Value)
        FormDelivery.Text108.Text = Trim(vRecordset.Fields("MyDescription").Value)
        If Trim(vRecordset.Fields("timesend").Value) <> "__:__" And Trim(vRecordset.Fields("timesend").Value) <> "" Then
        FormDelivery.MaskEdBox101.Text = Trim(vRecordset.Fields("timesend").Value)
        Else
        FormDelivery.MaskEdBox101.Mask = "##:##"
        End If
        If Trim(vRecordset.Fields("timereturn").Value) <> "__:__" And Trim(vRecordset.Fields("timereturn").Value) <> "" Then
        FormDelivery.MaskEdBox102.Text = Trim(vRecordset.Fields("timereturn").Value)
        Else
        FormDelivery.MaskEdBox102.Mask = "##:##"
        End If
        If Trim(vRecordset.Fields("isreturn").Value) = 0 Then
            FormDelivery.Check101 = 0
        Else
            FormDelivery.Check101 = 1
        End If
        If vIsOpen2 = 1 And FormDelivery.Check101.Value = 1 Then
            FormDelivery.Check101.Enabled = False
            FormDelivery.CMD106.Enabled = False
            FormDelivery.CMD201.Enabled = False
        Else
            FormDelivery.Check101.Enabled = True
            FormDelivery.CMD106.Enabled = True
            FormDelivery.CMD201.Enabled = True
        End If
        Select Case Trim(vRecordset.Fields("sendresult").Value)
        Case 0:
            FormDelivery.Option201.Value = True
            FormDelivery.Option202.Value = False
            FormDelivery.Option203.Value = False
            FormDelivery.Option204.Value = False
            FormDelivery.Option205.Value = False
            FormDelivery.Option206.Value = False
        Case 1:
            FormDelivery.Option201.Value = False
            FormDelivery.Option202.Value = True
            FormDelivery.Option203.Value = False
            FormDelivery.Option204.Value = False
            FormDelivery.Option205.Value = False
            FormDelivery.Option206.Value = False
        Case 2:
            FormDelivery.Option201.Value = False
            FormDelivery.Option202.Value = False
            FormDelivery.Option203.Value = True
            FormDelivery.Option204.Value = False
            FormDelivery.Option205.Value = False
            FormDelivery.Option206.Value = False
        Case 3:
            FormDelivery.Option201.Value = False
            FormDelivery.Option202.Value = False
            FormDelivery.Option203.Value = False
            FormDelivery.Option204.Value = True
            FormDelivery.Option205.Value = False
            FormDelivery.Option206.Value = False
        Case 4:
            FormDelivery.Option201.Value = False
            FormDelivery.Option202.Value = False
            FormDelivery.Option203.Value = False
            FormDelivery.Option204.Value = False
            FormDelivery.Option205.Value = True
            FormDelivery.Option206.Value = False
        Case 5:
            FormDelivery.Option201.Value = False
            FormDelivery.Option202.Value = False
            FormDelivery.Option203.Value = False
            FormDelivery.Option204.Value = False
            FormDelivery.Option205.Value = False
            FormDelivery.Option206.Value = True
        End Select
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        Set vSearchList = FormDelivery.ListView101.ListItems.Add(, , Trim(i))
            vSearchList.SubItems(1) = Trim(vRecordset.Fields("QueuesubID").Value)
            vSearchList.SubItems(2) = Trim(vRecordset.Fields("itemcode").Value)
            vSearchList.SubItems(3) = Trim(vRecordset.Fields("itemname").Value)
            vSearchList.SubItems(4) = Trim(vRecordset.Fields("qty").Value)
            vSearchList.SubItems(5) = Trim(vRecordset.Fields("confirmqty").Value)
            vSearchList.SubItems(6) = Trim(vRecordset.Fields("unitcode").Value)
            vSearchList.SubItems(8) = Trim(vRecordset.Fields("queueid").Value)
            vSearchList.SubItems(9) = Trim(vRecordset.Fields("invoiceno").Value)
            
        vRecordset.MoveNext
        i = i + 1
        Wend
    End If
    vRecordset.Close
    
    FormDelivery.ListView102.ListItems.Clear
    vQuery = "exec bcnp.dbo.usp_DO_EmpDelivery_Search '" & vSearch & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        j = 1
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        Set vEmpListSelect = FormDelivery.ListView102.ListItems.Add(, , Trim(j))
        vEmpListSelect.SubItems(1) = Trim(vRecordset.Fields("empbplusid").Value)
        vEmpListSelect.SubItems(2) = Trim(vRecordset.Fields("code").Value)
        vEmpListSelect.SubItems(3) = Trim(vRecordset.Fields("name1").Value)
        vEmpListSelect.SubItems(4) = Trim(vRecordset.Fields("positionname").Value)
        vRecordset.MoveNext
        j = j + 1
        Wend
    End If
    vRecordset.Close
    'FormDelivery.Text106.Enabled = False
    'FormDelivery.DTPicker102.Enabled = False
    'FormDelivery.MaskEdBox101.Enabled = False
    'Unload FrmOrder013
    FormDelivery.Enabled = True
    FrmOrder013.Hide
End If


ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD103_Click()
'Unload FrmOrder013
FormDelivery.Enabled = True
FrmOrder013.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIFrmProgramPrint.DO1.Enabled = True
FormDelivery.Enabled = True
End Sub

Private Sub ListView101_DblClick()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vSearch As String
Dim vSearchList As ListItem
Dim vEmpListSelect  As ListItem
Dim i As Integer
Dim j As Integer

On Error GoTo ErrDescription

If ListView101.ListItems.Count <> 0 Then
    FormDelivery.ListView101.ListItems.Clear
    vIsOpen2 = 1
    vSearch = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
    vQuery = "exec bcnp.dbo.USP_DO_DeliverySearchDetail '" & vSearch & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        i = 1
        If Trim(vRecordset.Fields("iscancel").Value) = 1 Then
            FormDelivery.Image101.Visible = False
            FormDelivery.Image102.Visible = True
        Else
            FormDelivery.Image101.Visible = True
            FormDelivery.Image102.Visible = False
        End If
        FormDelivery.DTPicker101.Value = Trim(vRecordset.Fields("docdate").Value)
        FormDelivery.DTPicker102.Value = Trim(vRecordset.Fields("datesend").Value)
        FormDelivery.DTPicker103.Value = Now
        FormDelivery.Text101.Text = Trim(vRecordset.Fields("id").Value)
        FormDelivery.Text102.Text = Trim(vRecordset.Fields("docno").Value)
        FormDelivery.Text103.Text = Trim(vRecordset.Fields("VehicalID").Value)
        FormDelivery.Text104.Text = Trim(vRecordset.Fields("CarNo").Value)
        FormDelivery.Text105.Text = Trim(vRecordset.Fields("CarLicence").Value)
        FormDelivery.Text106.Text = Trim(vRecordset.Fields("MeasureStart").Value)
        FormDelivery.Text107.Text = Trim(vRecordset.Fields("MeasureStop").Value)
        FormDelivery.Text108.Text = Trim(vRecordset.Fields("MyDescription").Value)
        If Trim(vRecordset.Fields("timesend").Value) <> "__:__" And Trim(vRecordset.Fields("timesend").Value) <> "" Then
        FormDelivery.MaskEdBox101.Text = Trim(vRecordset.Fields("timesend").Value)
        Else
        FormDelivery.MaskEdBox101.Mask = "##:##"
        End If
        If Trim(vRecordset.Fields("timereturn").Value) <> "__:__" And Trim(vRecordset.Fields("timereturn").Value) <> "" Then
        FormDelivery.MaskEdBox102.Text = Trim(vRecordset.Fields("timereturn").Value)
        Else
        FormDelivery.MaskEdBox102.Mask = "##:##"
        End If
        If Trim(vRecordset.Fields("isreturn").Value) = 0 Then
            FormDelivery.Check101 = 0
        Else
            FormDelivery.Check101 = 1
        End If
        'If vIsOpen2 = 1 And FormDelivery.Check101.Value = 1 Then
            'FormDelivery.Check101.Enabled = False
            'FormDelivery.CMD106.Enabled = False
            'FormDelivery.CMD201.Enabled = False
        'Else
            FormDelivery.Check101.Enabled = True
            FormDelivery.CMD106.Enabled = True
            FormDelivery.CMD201.Enabled = True
        'End If
        Select Case Trim(vRecordset.Fields("sendresult").Value)
        Case 0:
            FormDelivery.Option201.Value = True
            FormDelivery.Option202.Value = False
            FormDelivery.Option203.Value = False
            FormDelivery.Option204.Value = False
            FormDelivery.Option205.Value = False
            FormDelivery.Option206.Value = False
        Case 1:
            FormDelivery.Option201.Value = False
            FormDelivery.Option202.Value = True
            FormDelivery.Option203.Value = False
            FormDelivery.Option204.Value = False
            FormDelivery.Option205.Value = False
            FormDelivery.Option206.Value = False
        Case 2:
            FormDelivery.Option201.Value = False
            FormDelivery.Option202.Value = False
            FormDelivery.Option203.Value = True
            FormDelivery.Option204.Value = False
            FormDelivery.Option205.Value = False
            FormDelivery.Option206.Value = False
        Case 3:
            FormDelivery.Option201.Value = False
            FormDelivery.Option202.Value = False
            FormDelivery.Option203.Value = False
            FormDelivery.Option204.Value = True
            FormDelivery.Option205.Value = False
            FormDelivery.Option206.Value = False
        Case 4:
            FormDelivery.Option201.Value = False
            FormDelivery.Option202.Value = False
            FormDelivery.Option203.Value = False
            FormDelivery.Option204.Value = False
            FormDelivery.Option205.Value = True
            FormDelivery.Option206.Value = False
        Case 5:
            FormDelivery.Option201.Value = False
            FormDelivery.Option202.Value = False
            FormDelivery.Option203.Value = False
            FormDelivery.Option204.Value = False
            FormDelivery.Option205.Value = False
            FormDelivery.Option206.Value = True
        End Select
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        Set vSearchList = FormDelivery.ListView101.ListItems.Add(, , Trim(i))
            vSearchList.SubItems(1) = Trim(vRecordset.Fields("QueuesubID").Value)
            vSearchList.SubItems(2) = Trim(vRecordset.Fields("itemcode").Value)
            vSearchList.SubItems(3) = Trim(vRecordset.Fields("itemname").Value)
            vSearchList.SubItems(4) = Trim(vRecordset.Fields("qty").Value)
            vSearchList.SubItems(5) = Trim(vRecordset.Fields("confirmqty").Value)
            vSearchList.SubItems(6) = Trim(vRecordset.Fields("unitcode").Value)
            vSearchList.SubItems(8) = Trim(vRecordset.Fields("queueid").Value)
            vSearchList.SubItems(9) = Trim(vRecordset.Fields("invoiceno").Value)
            
        vRecordset.MoveNext
        i = i + 1
        Wend
    End If
    vRecordset.Close
    
    FormDelivery.ListView102.ListItems.Clear
    vQuery = "exec bcnp.dbo.usp_DO_EmpDelivery_Search '" & vSearch & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        j = 1
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        Set vEmpListSelect = FormDelivery.ListView102.ListItems.Add(, , Trim(j))
        vEmpListSelect.SubItems(1) = Trim(vRecordset.Fields("empbplusid").Value)
        vEmpListSelect.SubItems(2) = Trim(vRecordset.Fields("code").Value)
        vEmpListSelect.SubItems(3) = Trim(vRecordset.Fields("name1").Value)
        vEmpListSelect.SubItems(4) = Trim(vRecordset.Fields("positionname").Value)
        vEmpListSelect.SubItems(5) = Format(Trim(vRecordset.Fields("empwages").Value), "##,##0.00")
        vRecordset.MoveNext
        j = j + 1
        Wend
    End If
    vRecordset.Close
    'FormDelivery.Text106.Enabled = False
    'FormDelivery.DTPicker102.Enabled = False
    'FormDelivery.MaskEdBox101.Enabled = False
    'Unload FrmOrder013
    FormDelivery.Enabled = True
    FrmOrder013.Hide
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub ListView101_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vSearch As String
Dim vSearchList As ListItem
Dim vEmpListSelect  As ListItem
Dim i As Integer
Dim j As Integer

On Error GoTo ErrDescription

If KeyAscii = 13 Then
If ListView101.ListItems.Count <> 0 Then
    FormDelivery.ListView101.ListItems.Clear
    vIsOpen2 = 1
    vSearch = Trim(ListView101.ListItems.Item(ListView101.SelectedItem.Index).SubItems(1))
    vQuery = "exec bcnp.dbo.USP_DO_DeliverySearchDetail '" & vSearch & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        i = 1
        If Trim(vRecordset.Fields("iscancel").Value) = 1 Then
            FormDelivery.Image101.Visible = False
            FormDelivery.Image102.Visible = True
        Else
            FormDelivery.Image101.Visible = True
            FormDelivery.Image102.Visible = False
        End If
        FormDelivery.DTPicker101.Value = Trim(vRecordset.Fields("docdate").Value)
        FormDelivery.DTPicker102.Value = Trim(vRecordset.Fields("datesend").Value)
        FormDelivery.DTPicker103.Value = Now
        FormDelivery.Text101.Text = Trim(vRecordset.Fields("id").Value)
        FormDelivery.Text102.Text = Trim(vRecordset.Fields("docno").Value)
        FormDelivery.Text103.Text = Trim(vRecordset.Fields("VehicalID").Value)
        FormDelivery.Text104.Text = Trim(vRecordset.Fields("CarNo").Value)
        FormDelivery.Text105.Text = Trim(vRecordset.Fields("CarLicence").Value)
        FormDelivery.Text106.Text = Trim(vRecordset.Fields("MeasureStart").Value)
        FormDelivery.Text107.Text = Trim(vRecordset.Fields("MeasureStop").Value)
        FormDelivery.Text108.Text = Trim(vRecordset.Fields("MyDescription").Value)
        If Trim(vRecordset.Fields("timesend").Value) <> "__:__" And Trim(vRecordset.Fields("timesend").Value) <> "" Then
        FormDelivery.MaskEdBox101.Text = Trim(vRecordset.Fields("timesend").Value)
        Else
        FormDelivery.MaskEdBox101.Mask = "##:##"
        End If
        If Trim(vRecordset.Fields("timereturn").Value) <> "__:__" And Trim(vRecordset.Fields("timereturn").Value) <> "" Then
        FormDelivery.MaskEdBox102.Text = Trim(vRecordset.Fields("timereturn").Value)
        Else
        FormDelivery.MaskEdBox102.Mask = "##:##"
        End If
        If Trim(vRecordset.Fields("isreturn").Value) = 0 Then
            FormDelivery.Check101 = 0
        Else
            FormDelivery.Check101 = 1
        End If
        'If vIsOpen2 = 1 And FormDelivery.Check101.Value = 1 Then
            'FormDelivery.Check101.Enabled = False
            'FormDelivery.CMD106.Enabled = False
            'FormDelivery.CMD201.Enabled = False
        'Else
            FormDelivery.Check101.Enabled = True
            FormDelivery.CMD106.Enabled = True
            FormDelivery.CMD201.Enabled = True
        'End If
        Select Case Trim(vRecordset.Fields("sendresult").Value)
        Case 0:
            FormDelivery.Option201.Value = True
            FormDelivery.Option202.Value = False
            FormDelivery.Option203.Value = False
            FormDelivery.Option204.Value = False
            FormDelivery.Option205.Value = False
            FormDelivery.Option206.Value = False
        Case 1:
            FormDelivery.Option201.Value = False
            FormDelivery.Option202.Value = True
            FormDelivery.Option203.Value = False
            FormDelivery.Option204.Value = False
            FormDelivery.Option205.Value = False
            FormDelivery.Option206.Value = False
        Case 2:
            FormDelivery.Option201.Value = False
            FormDelivery.Option202.Value = False
            FormDelivery.Option203.Value = True
            FormDelivery.Option204.Value = False
            FormDelivery.Option205.Value = False
            FormDelivery.Option206.Value = False
        Case 3:
            FormDelivery.Option201.Value = False
            FormDelivery.Option202.Value = False
            FormDelivery.Option203.Value = False
            FormDelivery.Option204.Value = True
            FormDelivery.Option205.Value = False
            FormDelivery.Option206.Value = False
        Case 4:
            FormDelivery.Option201.Value = False
            FormDelivery.Option202.Value = False
            FormDelivery.Option203.Value = False
            FormDelivery.Option204.Value = False
            FormDelivery.Option205.Value = True
            FormDelivery.Option206.Value = False
        Case 5:
            FormDelivery.Option201.Value = False
            FormDelivery.Option202.Value = False
            FormDelivery.Option203.Value = False
            FormDelivery.Option204.Value = False
            FormDelivery.Option205.Value = False
            FormDelivery.Option206.Value = True
        End Select
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        Set vSearchList = FormDelivery.ListView101.ListItems.Add(, , Trim(i))
            vSearchList.SubItems(1) = Trim(vRecordset.Fields("QueuesubID").Value)
            vSearchList.SubItems(2) = Trim(vRecordset.Fields("itemcode").Value)
            vSearchList.SubItems(3) = Trim(vRecordset.Fields("itemname").Value)
            vSearchList.SubItems(4) = Trim(vRecordset.Fields("qty").Value)
            vSearchList.SubItems(5) = Trim(vRecordset.Fields("confirmqty").Value)
            vSearchList.SubItems(6) = Trim(vRecordset.Fields("unitcode").Value)
            vSearchList.SubItems(8) = Trim(vRecordset.Fields("queueid").Value)
            vSearchList.SubItems(9) = Trim(vRecordset.Fields("invoiceno").Value)
            
        vRecordset.MoveNext
        i = i + 1
        Wend
    End If
    vRecordset.Close
    
    FormDelivery.ListView102.ListItems.Clear
    vQuery = "exec bcnp.dbo.usp_DO_EmpDelivery_Search '" & vSearch & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        j = 1
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        Set vEmpListSelect = FormDelivery.ListView102.ListItems.Add(, , Trim(j))
        vEmpListSelect.SubItems(1) = Trim(vRecordset.Fields("empbplusid").Value)
        vEmpListSelect.SubItems(2) = Trim(vRecordset.Fields("code").Value)
        vEmpListSelect.SubItems(3) = Trim(vRecordset.Fields("name1").Value)
        vEmpListSelect.SubItems(4) = Trim(vRecordset.Fields("positionname").Value)
        vEmpListSelect.SubItems(5) = Format(Trim(vRecordset.Fields("empwages").Value), "##,##0.00")
        vRecordset.MoveNext
        j = j + 1
        Wend
    End If
    vRecordset.Close
    'FormDelivery.Text106.Enabled = False
    'FormDelivery.DTPicker102.Enabled = False
    'FormDelivery.MaskEdBox101.Enabled = False
    'Unload FrmOrder013
    FormDelivery.Enabled = True
    FrmOrder013.Hide
End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vSearch As String
Dim vSearchList As ListItem
Dim i As Integer

On Error Resume Next

If KeyAscii = 13 Then
        vSearch = Trim(Text101.Text)
        ListView101.ListItems.Clear
        vQuery = "exec bcnp.dbo.USP_DO_DeliverySearchHeader '" & vSearch & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            i = 1
            While Not vRecordset.EOF
                Set vSearchList = ListView101.ListItems.Add(, , Trim(i))
                vSearchList.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
                vSearchList.SubItems(2) = Trim(vRecordset.Fields("docdate").Value)
                vSearchList.SubItems(3) = Trim(vRecordset.Fields("DateSend").Value)
                vSearchList.SubItems(4) = Trim(vRecordset.Fields("TimeSend").Value)
                vSearchList.SubItems(5) = Trim(vRecordset.Fields("DateReturn").Value)
                vSearchList.SubItems(6) = Trim(vRecordset.Fields("TimeReturn").Value)
                vSearchList.SubItems(7) = Trim(vRecordset.Fields("MeasureStart").Value)
                vSearchList.SubItems(8) = Trim(vRecordset.Fields("MeasureStop").Value)
                vSearchList.SubItems(9) = Trim(vRecordset.Fields("MyDescription").Value)
                vSearchList.SubItems(10) = Trim(vRecordset.Fields("VehicalID").Value)
                vSearchList.SubItems(11) = Trim(vRecordset.Fields("iscancel").Value)
                vRecordset.MoveNext
            i = i + 1
            Wend
            Me.ListView101.SetFocus
       Else
        MsgBox "ไม่มีข้อมูล ตามคำที่ใช้ค้นหา", vbInformation, "Send Information"
       End If
        vRecordset.Close
End If
End Sub

Private Sub Text1_Change()

End Sub
