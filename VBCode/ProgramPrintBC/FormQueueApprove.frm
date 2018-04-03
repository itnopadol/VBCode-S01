VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FormQueueApprove 
   Caption         =   "กำหนดเวลาขนส่งสินค้าของเอกสารขอเข้าคิวจัดส่งสินค้า"
   ClientHeight    =   11490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15375
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FormQueueApprove.frx":0000
   ScaleHeight     =   11490
   ScaleWidth      =   15375
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PICEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   10095
      Left            =   135
      ScaleHeight     =   10065
      ScaleWidth      =   14970
      TabIndex        =   4
      Top             =   1215
      Visible         =   0   'False
      Width           =   15000
      Begin MSMask.MaskEdBox MBReqTime 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "H:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1054
            SubFormatType   =   4
         EndProperty
         Height          =   285
         Left            =   4590
         TabIndex        =   27
         Top             =   2115
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12632319
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker DTPReqDate 
         Height          =   285
         Left            =   1800
         TabIndex        =   26
         Top             =   2115
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   503
         _Version        =   393216
         Format          =   64880641
         CurrentDate     =   40527
      End
      Begin VB.CommandButton CMDCloseEdit 
         Caption         =   "ปิด"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   9630
         TabIndex        =   25
         Top             =   7785
         Width           =   1815
      End
      Begin VB.CommandButton CMDEdit 
         BackColor       =   &H00C0C0C0&
         Caption         =   "แก้ไข"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   7695
         TabIndex        =   24
         Top             =   7785
         Width           =   1815
      End
      Begin MSComctlLib.ListView ListViewItem 
         Height          =   3120
         Left            =   135
         TabIndex        =   5
         Top             =   4410
         Width           =   11310
         _ExtentX        =   19950
         _ExtentY        =   5503
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ลำดับ"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "รหัสสินค้า"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ชื่อสินค้า"
            Object.Width           =   9701
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "จำนวน"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "หน่วย"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.Label LBLIndex 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   8145
         TabIndex        =   29
         Top             =   315
         Width           =   555
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "รายการสินค้า"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   135
         TabIndex        =   28
         Top             =   4095
         Width           =   1230
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เวลา :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3555
         TabIndex        =   23
         Top             =   2115
         Width           =   1005
      End
      Begin VB.Label LBLAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1800
         TabIndex        =   22
         Top             =   1215
         Width           =   9645
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ที่อยู่ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   540
         TabIndex        =   21
         Top             =   1215
         Width           =   1230
      End
      Begin VB.Label LBLDestination 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1800
         TabIndex        =   20
         Top             =   3015
         Width           =   9645
      End
      Begin VB.Label LBLDocDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6345
         TabIndex        =   19
         Top             =   315
         Width           =   1725
      End
      Begin VB.Label LBLMyDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1800
         TabIndex        =   18
         Top             =   2565
         Width           =   9645
      End
      Begin VB.Label LBLSaleName 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1800
         TabIndex        =   17
         Top             =   1665
         Width           =   4515
      End
      Begin VB.Label LBLStation 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1800
         TabIndex        =   16
         Top             =   3465
         Width           =   9645
      End
      Begin VB.Label LBLARName 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1800
         TabIndex        =   15
         Top             =   765
         Width           =   9645
      End
      Begin VB.Label LBLDocNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1800
         TabIndex        =   14
         Top             =   315
         Width           =   2490
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่ต้องการส่งสินค้า :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   13
         Top             =   2115
         Width           =   1680
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "สถานที่ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   945
         TabIndex        =   12
         Top             =   3465
         Width           =   825
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "รหัส/ชื่อ ลูกค้า :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   540
         TabIndex        =   11
         Top             =   765
         Width           =   1230
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "หมายเหตุ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   810
         TabIndex        =   10
         Top             =   2565
         Width           =   960
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "พนักงานขาย :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   540
         TabIndex        =   9
         Top             =   1665
         Width           =   1230
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เส้นทาง :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   990
         TabIndex        =   8
         Top             =   3015
         Width           =   780
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่เอกสาร :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5220
         TabIndex        =   7
         Top             =   315
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่ใบจัดส่งสินค้า :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   270
         TabIndex        =   6
         Top             =   315
         Width           =   1500
      End
   End
   Begin VB.CommandButton CMDSave 
      Caption         =   "บันทึกข้อมูล"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   13185
      TabIndex        =   3
      Top             =   8415
      Width           =   1905
   End
   Begin MSComctlLib.ListView ListViewQueue 
      Height          =   5775
      Left            =   135
      TabIndex        =   2
      Top             =   2520
      Width           =   14955
      _ExtentX        =   26379
      _ExtentY        =   10186
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับ"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "เลขที่"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "วันที่ขอส่ง"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "เวลาที่ขอส่ง"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "เปลี่ยนวัน"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "เปลี่ยนเวลา"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ชื่อลูกค้า"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "สถานที่"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "คำอธิบาย"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "ชื่อผู้ติดต่อ"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "เบอร์ติดต่อ"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "พนักงานขาย"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "สถานะยืนยัน"
         Object.Width           =   2
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPQueDate 
      Height          =   375
      Left            =   1710
      TabIndex        =   1
      Top             =   1485
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   64880641
      CurrentDate     =   40527
   End
   Begin VB.CheckBox CHKAll 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "เลือกทั้งหมด"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   135
      TabIndex        =   30
      Top             =   2160
      Width           =   1410
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่ขอส่งสินค้า :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   135
      TabIndex        =   0
      Top             =   1485
      Width           =   1635
   End
End
Attribute VB_Name = "FormQueueApprove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vQuery As String


Private Sub CHKAll_Click()
Dim i As Integer

On Error Resume Next

If Me.ListViewQueue.ListItems.Count > 0 Then
    If Me.CHKAll.Value = 1 Then
    For i = 1 To Me.ListViewQueue.ListItems.Count
    Me.ListViewQueue.ListItems(i).Checked = True
    Next i
    End If
    
    If Me.CHKAll.Value = 0 Then
    For i = 1 To Me.ListViewQueue.ListItems.Count
    Me.ListViewQueue.ListItems(i).Checked = False
    Next i
    End If
End If
End Sub

Private Sub CMDCloseEdit_Click()
Me.PICEdit.Visible = False
End Sub

Private Sub CMDEdit_Click()
Dim vIndex As Integer
Dim vCheckDate As Date
Dim vEditDate As Date
Dim vCheckHour As Integer
Dim vCheckMinute As Integer
Dim vEditHour As Integer
Dim vEditMinute As Integer
Dim vNow As Date

Dim vHour As Integer
Dim vMinute As Integer

Dim vCheckTime As String
Dim vCheckInstr As Integer

On Error Resume Next

vIndex = Me.LBLIndex.Caption

If Me.ListViewQueue.ListItems.Count > 0 Then
    If Me.MBReqTime.Text <> "" Then
    vCheckTime = Me.MBReqTime.Text
    
    vCheckInstr = InStr(1, vCheckTime, "_")
    
    If vCheckInstr > 0 Then
    MsgBox "กรุณากรอกเวลาให้ครบตาม รูปแบบของเวลา ดังตัวอย่าง 08:30,08:05,12:30 เป็นต้น"
    Me.MBReqTime.SetFocus
    Exit Sub
    End If
End If


If Me.ListViewQueue.ListItems(vIndex).SubItems(4) = "" And Me.ListViewQueue.ListItems(vIndex).SubItems(5) = "" Then
    vCheckDate = Me.ListViewQueue.ListItems(vIndex).SubItems(2)
    vCheckHour = Hour(Now) 'Left(Me.ListViewQueue.ListItems(vIndex).SubItems(3), 2)
    vCheckMinute = Minute(Now) 'Right(Me.ListViewQueue.ListItems(vIndex).SubItems(3), 2)
    
    vEditDate = Me.DTPReqDate.Value
    vEditHour = Left(Me.MBReqTime.Text, 2)
    vEditMinute = Right(Me.MBReqTime.Text, 2)
    
    vHour = Left(Me.ListViewQueue.ListItems(vIndex).SubItems(3), 2)
    vMinute = Right(Me.ListViewQueue.ListItems(vIndex).SubItems(3), 2)
    
    vNow = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
    
    If vCheckDate = vEditDate And vCheckDate = vNow Then
        If vEditHour = vCheckHour Then
            If vEditMinute = vCheckMinute Then
                Me.PICEdit.Visible = False
                Exit Sub
            ElseIf vEditMinute < vCheckMinute Then
                MsgBox "ไม่สามารถย้อนเวลาได้", vbCritical, "Send Error Message"
                Exit Sub
            End If
        ElseIf vEditHour < vCheckHour Then
            MsgBox "ไม่สามารถย้อนเวลาได้", vbCritical, "Send Error Message"
            Exit Sub
        End If
        
        If vEditHour = vHour Then
            If vEditMinute = vMinute Then
                Me.PICEdit.Visible = False
                Exit Sub
            End If
        End If
        
    ElseIf vEditDate < vCheckDate And vCheckDate = vNow Then
        MsgBox "ไม่สามารถย้อนเวลาได้", vbCritical, "Send Error Message"
        Exit Sub
    End If
End If

If Me.ListViewQueue.ListItems(vIndex).SubItems(4) <> "" And Me.ListViewQueue.ListItems(vIndex).SubItems(5) <> "" Then
    vCheckDate = Me.ListViewQueue.ListItems(vIndex).SubItems(4)
    vCheckHour = Hour(Now) 'Left(Me.ListViewQueue.ListItems(vIndex).SubItems(3), 2)
    vCheckMinute = Minute(Now) 'Right(Me.ListViewQueue.ListItems(vIndex).SubItems(3), 2)
    
    vHour = Left(Me.ListViewQueue.ListItems(vIndex).SubItems(5), 2)
    vMinute = Right(Me.ListViewQueue.ListItems(vIndex).SubItems(5), 2)
    
    vEditDate = Me.DTPReqDate.Value
    vEditHour = Left(Me.MBReqTime.Text, 2)
    vEditMinute = Right(Me.MBReqTime.Text, 2)
    
    vNow = Day(Now) & "/" & Month(Now) & "/" & Year(Now)
    
    If vCheckDate = vEditDate And vCheckDate = vNow Then
        If vEditHour = vCheckHour Then
            If vEditMinute = vCheckMinute Then
                Me.PICEdit.Visible = False
                Exit Sub
            ElseIf vEditMinute < vCheckMinute Then
                MsgBox "ไม่สามารถย้อนเวลาได้", vbCritical, "Send Error Message"
                Exit Sub
            End If
        ElseIf vEditHour < vCheckHour Then
            MsgBox "ไม่สามารถย้อนเวลาได้", vbCritical, "Send Error Message"
            Exit Sub
        End If
        
        If vEditHour = vHour Then
            If vEditMinute = vMinute Then
                Me.PICEdit.Visible = False
                Exit Sub
            End If
        End If
        
    ElseIf vEditDate < vCheckDate And vCheckDate = vNow Then
        MsgBox "ไม่สามารถย้อนเวลาได้", vbCritical, "Send Error Message"
        Exit Sub
    End If
End If

Me.ListViewQueue.ListItems(vIndex).SubItems(4) = Me.DTPReqDate.Value
Me.ListViewQueue.ListItems(vIndex).SubItems(5) = Me.MBReqTime.Text

Me.ListViewQueue.ListItems(vIndex).ForeColor = "&H000000FF"
Me.ListViewQueue.ListItems.Item(vIndex).ListSubItems(1).ForeColor = "&H000000FF"
Me.ListViewQueue.ListItems.Item(vIndex).ListSubItems(2).ForeColor = "&H000000FF"
Me.ListViewQueue.ListItems.Item(vIndex).ListSubItems(3).ForeColor = "&H000000FF"
Me.ListViewQueue.ListItems.Item(vIndex).ListSubItems(4).ForeColor = "&H000000FF"
Me.ListViewQueue.ListItems.Item(vIndex).ListSubItems(5).ForeColor = "&H000000FF"
Me.ListViewQueue.ListItems.Item(vIndex).ListSubItems(6).ForeColor = "&H000000FF"
Me.ListViewQueue.ListItems.Item(vIndex).ListSubItems(7).ForeColor = "&H000000FF"
Me.ListViewQueue.ListItems.Item(vIndex).ListSubItems(8).ForeColor = "&H000000FF"
Me.ListViewQueue.ListItems.Item(vIndex).ListSubItems(9).ForeColor = "&H000000FF"
Me.ListViewQueue.ListItems.Item(vIndex).ListSubItems(10).ForeColor = "&H000000FF"
Me.ListViewQueue.ListItems.Item(vIndex).ListSubItems(11).ForeColor = "&H000000FF"

Me.LBLIndex.Caption = ""
Me.PICEdit.Visible = False
End If
End Sub

Private Sub CMDSave_Click()
Dim i As Integer
Dim vDayConfirm As String
Dim vTimeConfirm As String
Dim vDocNo As String
Dim vCountSelect As Integer

On Error Resume Next

If Me.ListViewQueue.ListItems.Count > 0 Then
    For i = 1 To Me.ListViewQueue.ListItems.Count
    If Me.ListViewQueue.ListItems.Item(i).Checked = True Then
    vCountSelect = vCountSelect + 1
    End If
    Next

    If vCountSelect = 0 Then
        MsgBox "ไม่ได้เลือกเอกสารที่จะยืนยันการส่งสินค้า", vbCritical, "Send Error Message"
        Exit Sub
    End If
    
    For i = 1 To Me.ListViewQueue.ListItems.Count
    If Me.ListViewQueue.ListItems(i).Checked = True Then
        vDocNo = Me.ListViewQueue.ListItems.Item(i).SubItems(1)
        If Me.ListViewQueue.ListItems.Item(i).SubItems(4) = "" Then
        vDayConfirm = Me.ListViewQueue.ListItems.Item(i).SubItems(2)
        vTimeConfirm = Me.ListViewQueue.ListItems.Item(i).SubItems(3)
        Else
        vDayConfirm = Me.ListViewQueue.ListItems.Item(i).SubItems(4)
        vTimeConfirm = Me.ListViewQueue.ListItems.Item(i).SubItems(5)
        End If
        

        vQuery = "exec dbo.USP_DO_UpdateQueueSendConfirm '" & vDocNo & "','" & vUserID & "','" & vDayConfirm & "','" & vTimeConfirm & "' "
        gConnection.Execute vQuery
    End If
    Next
    
    MsgBox "บันทึกข้อมูลเรียบร้อยแล้ว", vbInformation, "Send Information Message"
    Call SearchQueDaily
End If

End Sub

Private Sub Form_Load()
Me.DTPQueDate.Value = Now
Me.DTPReqDate.Value = Now

Call SearchQueDaily
End Sub

Public Sub SearchQueDaily()
Dim vDocdate As String
Dim vRecordset As New ADODB.Recordset
Dim vListItemQueue As ListItem
Dim i As Integer

On Error Resume Next

vDocdate = Day(Me.DTPQueDate) & "/" & Month(Me.DTPQueDate) & "/" & Year(Me.DTPQueDate)

vQuery = "exec dbo.USP_DO_QueueConfirmSend '" & vDocdate & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    Me.ListViewQueue.ListItems.Clear
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    i = i + 1
    Set vListItemQueue = Me.ListViewQueue.ListItems.Add(, , i)
    vListItemQueue.SubItems(1) = vRecordset.Fields("docno").Value
    vListItemQueue.SubItems(2) = vRecordset.Fields("duedate").Value
    vListItemQueue.SubItems(3) = vRecordset.Fields("duetime").Value
    If vRecordset.Fields("timeconfirm").Value <> "" Then
    vListItemQueue.SubItems(4) = vRecordset.Fields("dayconfirm").Value
    vListItemQueue.SubItems(5) = vRecordset.Fields("timeconfirm").Value
    Else
    vListItemQueue.SubItems(4) = ""
    vListItemQueue.SubItems(5) = ""
    End If
    vListItemQueue.SubItems(6) = vRecordset.Fields("arname").Value
    vListItemQueue.SubItems(7) = "ต." & " " & vRecordset.Fields("district").Value & " อ." & " " & vRecordset.Fields("amphur").Value & " จ." & " " & vRecordset.Fields("province").Value
    vListItemQueue.SubItems(8) = vRecordset.Fields("transportlocation").Value
    vListItemQueue.SubItems(9) = vRecordset.Fields("receivename").Value
    vListItemQueue.SubItems(10) = vRecordset.Fields("receivetelhome").Value & "," & vRecordset.Fields("receivetelmobile").Value
    vListItemQueue.SubItems(11) = vRecordset.Fields("salename").Value
    vListItemQueue.SubItems(12) = vRecordset.Fields("sendapprove").Value
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

Dim n As Integer
Dim vQueDate As String
Dim vReqDate As String
Dim vSendApprove As Integer

vDocdate = Day(Me.DTPQueDate) & "/" & Month(Me.DTPQueDate) & "/" & Year(Me.DTPQueDate)
          
For n = 1 To Me.ListViewQueue.ListItems.Count

          vReqDate = Me.ListViewQueue.ListItems(n).SubItems(2)
          vSendApprove = Me.ListViewQueue.ListItems(n).SubItems(12)
          
          If vDocdate = vReqDate And vSendApprove = 0 Then
          Me.ListViewQueue.ListItems(n).ForeColor = "&H00008080"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(1).ForeColor = "&H00008080" 'yellow
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(2).ForeColor = "&H00008080"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(3).ForeColor = "&H00008080"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(4).ForeColor = "&H00008080"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(5).ForeColor = "&H00008080"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(6).ForeColor = "&H00008080"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(7).ForeColor = "&H00008080"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(8).ForeColor = "&H00008080"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(9).ForeColor = "&H00008080"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(10).ForeColor = "&H00008080"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(11).ForeColor = "&H00008080"
          ElseIf vSendApprove = 1 Then
          Me.ListViewQueue.ListItems(n).ForeColor = "&H00008000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(1).ForeColor = "&H00008000" 'yellow
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(2).ForeColor = "&H00008000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(3).ForeColor = "&H00008000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(4).ForeColor = "&H00008000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(5).ForeColor = "&H00008000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(6).ForeColor = "&H00008000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(7).ForeColor = "&H00008000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(8).ForeColor = "&H00008000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(9).ForeColor = "&H00008000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(10).ForeColor = "&H00008000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(11).ForeColor = "&H00008000"
          '&H00008000&
          End If
Next
End Sub


Private Sub DTPQueDate_Change()
Call SearchQueDaily
End Sub

Private Sub ListViewQueue_DblClick()
Dim vIndex As Integer
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim i As Integer
Dim vListItemItem As ListItem

On Error Resume Next

If Me.ListViewQueue.ListItems.Count > 0 Then
vIndex = Me.ListViewQueue.SelectedItem.Index

Me.LBLIndex.Caption = vIndex

vDocNo = Me.ListViewQueue.ListItems(vIndex).SubItems(1)

vQuery = "exec dbo.USP_DO_QueueDetails '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    Me.LBLDocNo.Caption = vRecordset.Fields("docno").Value
    Me.LBLDocDate.Caption = vRecordset.Fields("docdate").Value
    Me.LBLARName.Caption = vRecordset.Fields("arname").Value
    Me.LBLAddress.Caption = "ต." & " " & vRecordset.Fields("district").Value & " อ." & " " & vRecordset.Fields("amphur").Value & " จ." & " " & vRecordset.Fields("province").Value
    Me.LBLSaleName.Caption = vRecordset.Fields("salename").Value
    Me.LBLMyDescription.Caption = vRecordset.Fields("transportlocation").Value
    Me.LBLDestination.Caption = vRecordset.Fields("routename").Value
    Me.LBLStation.Caption = vRecordset.Fields("routedes").Value
    If Me.ListViewQueue.ListItems(vIndex).SubItems(4) <> "" Then
    Me.DTPReqDate.Value = Me.ListViewQueue.ListItems(vIndex).SubItems(4)
    Me.MBReqTime.Text = Me.ListViewQueue.ListItems(vIndex).SubItems(5)
    Else
    Me.DTPReqDate.Value = Me.ListViewQueue.ListItems(vIndex).SubItems(2)
    Me.MBReqTime.Text = Me.ListViewQueue.ListItems(vIndex).SubItems(3)
    End If
    
    Me.ListViewItem.ListItems.Clear
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    i = i + 1
    Set vListItemItem = Me.ListViewItem.ListItems.Add(, , i)
    vListItemItem.SubItems(1) = vRecordset.Fields("itemcode").Value
    vListItemItem.SubItems(2) = vRecordset.Fields("itemname").Value
    vListItemItem.SubItems(3) = Format(vRecordset.Fields("queremainqty").Value, "##,##0.00")
    vListItemItem.SubItems(4) = vRecordset.Fields("unitcode").Value
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

Me.PICEdit.Visible = True
Me.MBReqTime.SetFocus
End If
End Sub

Private Sub ListViewQueue_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vIndex As Integer
Dim n As Integer
Dim vQueDate As String
Dim vReqDate As String
Dim vDocdate  As String

On Error Resume Next

If KeyCode = 46 Then
vIndex = Me.ListViewQueue.SelectedItem.Index
Me.ListViewQueue.ListItems(vIndex).SubItems(4) = ""
Me.ListViewQueue.ListItems(vIndex).SubItems(5) = ""

vDocdate = Day(Me.DTPQueDate) & "/" & Month(Me.DTPQueDate) & "/" & Year(Me.DTPQueDate)
          
For n = 1 To Me.ListViewQueue.ListItems.Count

          vReqDate = Me.ListViewQueue.ListItems(n).SubItems(2)
          
          If vDocdate = vReqDate Then
          Me.ListViewQueue.ListItems(n).ForeColor = "&H00008000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(1).ForeColor = "&H00008000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(2).ForeColor = "&H00008000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(3).ForeColor = "&H00008000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(4).ForeColor = "&H00008000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(5).ForeColor = "&H00008000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(6).ForeColor = "&H00008000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(7).ForeColor = "&H00008000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(8).ForeColor = "&H00008000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(9).ForeColor = "&H00008000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(10).ForeColor = "&H00008000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(11).ForeColor = "&H00008000"
          Else
          Me.ListViewQueue.ListItems(n).ForeColor = "&H00000000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(1).ForeColor = "&H00000000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(2).ForeColor = "&H00000000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(3).ForeColor = "&H00000000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(4).ForeColor = "&H00000000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(5).ForeColor = "&H00000000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(6).ForeColor = "&H00000000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(7).ForeColor = "&H00000000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(8).ForeColor = "&H00000000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(9).ForeColor = "&H00000000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(10).ForeColor = "&H00000000"
          Me.ListViewQueue.ListItems.Item(n).ListSubItems(11).ForeColor = "&H00000000"
          End If
Next
End If

End Sub

Private Sub MBReqTime_LostFocus()
Dim vCheckTime As String
Dim vCheckInstr As Integer

On Error Resume Next

If Me.MBReqTime.Text <> "" Then
vCheckTime = Me.MBReqTime.Text

vCheckInstr = InStr(1, vCheckTime, "_")

If vCheckInstr > 0 Then
    MsgBox "กรุณากรอกเวลาให้ครบตาม รูปแบบของเวลา ดังตัวอย่าง 08:30,08:05,12:30 เป็นต้น"
    Me.MBReqTime.SetFocus
    Exit Sub
End If
End If
End Sub
