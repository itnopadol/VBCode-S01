VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form401 
   Caption         =   "ตรวจสอบ ใบเสนอสินค้าโปรโมชั่น"
   ClientHeight    =   9000
   ClientLeft      =   6000
   ClientTop       =   1485
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   14385
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab101 
      Height          =   8865
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   14250
      _ExtentX        =   25135
      _ExtentY        =   15637
      _Version        =   393216
      TabHeight       =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "รายการ ใบเสนอสินค้าโปรโมชั่น รอตรวจสอบ"
      TabPicture(0)   =   "Form401.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView101"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CMD101"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CMD103"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text106"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Crystal1011"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Crystal101"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "รายละเอียด ใบเสนอสินค้าโปรโมชั่น"
      TabPicture(1)   =   "Form401.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListView102"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "รายการ ใบเสนอสินค้าโปรโมชั่น ตรวจสอบแล้ว"
      TabPicture(2)   =   "Form401.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ListView103"
      Tab(2).Control(1)=   "Label7"
      Tab(2).ControlCount=   2
      Begin Crystal.CrystalReport Crystal101 
         Left            =   7695
         Top             =   7785
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.PictureBox Crystal1011 
         Height          =   480
         Left            =   4545
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   23
         Top             =   7830
         Width           =   1200
      End
      Begin VB.TextBox Text106 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1500
         TabIndex        =   22
         Top             =   2475
         Width           =   2415
      End
      Begin VB.CommandButton CMD103 
         Caption         =   "ดูรายละเอียดเอกสาร"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2400
         TabIndex        =   20
         Top             =   7350
         Width           =   1815
      End
      Begin MSComctlLib.ListView ListView103 
         Height          =   6240
         Left            =   -74475
         TabIndex        =   18
         Top             =   1200
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   11007
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "เลขที่เอกสาร"
            Object.Width           =   2734
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ผู้เสนอสินค้า"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "วันที่เอกสาร"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "วันที่ตรวจสอบ"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ตรวจสอบโดย"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "สำหรับโปรโมชั่น"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView ListView102 
         Height          =   5415
         Left            =   -74700
         TabIndex        =   8
         Top             =   2550
         Width           =   11265
         _ExtentX        =   19870
         _ExtentY        =   9551
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "รหัสสินค้า"
            Object.Width           =   2734
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ชื่อสินค้า"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ราคาปกติ"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ราคาโปรโมชั่น"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ลดราคา(บาท)"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "หน่วยนับ"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ยกเลิก"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Frame Frame2 
         Caption         =   "ข้อมูลแสดง รายละเอียดใบเสนอสินค้า"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1515
         Left            =   -74700
         TabIndex        =   7
         Top             =   750
         Width           =   11265
         Begin VB.TextBox Text105 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6675
            TabIndex        =   17
            Top             =   900
            Width           =   4440
         End
         Begin VB.TextBox Text104 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6675
            TabIndex        =   16
            Top             =   450
            Width           =   4440
         End
         Begin VB.TextBox Text103 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1575
            TabIndex        =   15
            Top             =   975
            Width           =   2040
         End
         Begin VB.CommandButton CMD102 
            Height          =   315
            Left            =   3675
            Picture         =   "Form401.frx":0054
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   450
            Width           =   315
         End
         Begin VB.TextBox Text102 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1425
            TabIndex        =   13
            Top             =   450
            Width           =   2190
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "โปรโมชั่น :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5700
            TabIndex        =   12
            Top             =   900
            Width           =   915
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "ชื่อผู้เสนอ :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5700
            TabIndex        =   11
            Top             =   450
            Width           =   990
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "วันที่ทำใบเสนอ :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   300
            TabIndex        =   10
            Top             =   975
            Width           =   1290
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "เลขที่ใบเสนอ :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   300
            TabIndex        =   9
            Top             =   450
            Width           =   1140
         End
      End
      Begin VB.CommandButton CMD101 
         Caption         =   "ตรวจสอบเอกสาร"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   375
         TabIndex        =   6
         Top             =   7350
         Width           =   1815
      End
      Begin MSComctlLib.ListView ListView101 
         Height          =   4215
         Left            =   375
         TabIndex        =   2
         Top             =   2925
         Width           =   11040
         _ExtentX        =   19473
         _ExtentY        =   7435
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "เลขที่เอกสาร"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ผู้เสนอสินค้า"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "วันที่ทำเอกสาร"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "วันที่แก้ไขล่าสุด"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "สำหรับโปรโมชั่น"
            Object.Width           =   6174
         EndProperty
      End
      Begin VB.Frame Frame1 
         Caption         =   "ข้อมูลการค้นหา รายการเสนอสินค้ารอตรวจสอบ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Left            =   375
         TabIndex        =   1
         Top             =   750
         Width           =   11040
         Begin VB.TextBox Text101 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   975
            TabIndex        =   4
            Top             =   600
            Width           =   5790
         End
         Begin VB.Label Label1 
            Caption         =   "ค้นหา"
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
            Left            =   450
            TabIndex        =   3
            Top             =   600
            Width           =   465
         End
      End
      Begin VB.Label Label8 
         Caption         =   "เลขที่เอกสาร"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   375
         TabIndex        =   21
         Top             =   2475
         Width           =   990
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "รายการ ใบเสนอสินค้าโปรโมชั่น"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74475
         TabIndex        =   19
         Top             =   975
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "รายการ ใบเสนอสินค้าโปรโมชั่น"
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
         Left            =   375
         TabIndex        =   5
         Top             =   2025
         Width           =   2565
      End
   End
End
Attribute VB_Name = "Form401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vSortResult      As Integer


Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim i As Integer
Dim vCheckDate As Date
Dim vDocno As String

On Error GoTo ErrDescription

vCheckDate = Trim(Now)
For i = ListView101.ListItems.Count To 1 Step -1
    If ListView101.ListItems.Item(i).Checked = True Then
        vDocno = ListView101.ListItems.Item(i).Text
        vQuery = "exec USP_PM_RequestCheck '" & vDocno & "','" & vUserID & "','" & vCheckDate & "' "
        gConnection.Execute vQuery
        ListView101.ListItems.Remove (i)
        Call InitializeSendEmail
        vQuery = "execute USP_PM_DeliverySendMail '" & vDocno & "' "
        vGetConnect.Execute vQuery
    End If
Next i

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD103_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vDocno As String

On Error GoTo ErrDescription

    vDocno = Trim(Text106.Text)
    If vDocno <> "" Then
        vQuery = "select reportname from bcnp.dbo.bcreportname where repid = 270 and reptype = 'PM' "
        If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
            vReportName = Trim(vRecordset.Fields("reportname").Value)
        End If
        vRecordset.Close
        
        With Crystal101
        .ReportFileName = vReportName & ".rpt"
        .ParameterFields(0) = "@vDocno;" & vDocno & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
        End With
        Text106.Text = ""
    Else
    MsgBox "กรุณาเลือก เลขที่เอกสารที่ต้องการดูรายละเอียดด้วยครับ", vbInformation, "ข้อความแจ้ง"
    Exit Sub
    End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Form_Load()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vListCheckPromotion As ListItem
Dim vListCheckedPromotion As ListItem
Dim vMemberDisc As Integer

'On Error Resume Next


Form401.ListView101.ListItems.Clear



    vQuery = "execute USP_PM_CheckOrConfirmSearch '0' "
    If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        Set vListCheckPromotion = Form401.ListView101.ListItems.Add(, , Trim(vRecordset.Fields("docno").Value))
        vListCheckPromotion.SubItems(1) = Trim(vRecordset.Fields("secmanname").Value)
        If IsNull(Trim(vRecordset.Fields("createdate").Value)) Then
        vListCheckPromotion.SubItems(2) = ""
        Else
            vListCheckPromotion.SubItems(2) = Trim(vRecordset.Fields("createdate").Value)
        End If
        If IsNull(Trim(vRecordset.Fields("editdate").Value)) Then
            vListCheckPromotion.SubItems(3) = ""
        Else
            vListCheckPromotion.SubItems(3) = Trim(vRecordset.Fields("editdate").Value)
        End If
        vListCheckPromotion.SubItems(4) = Trim(vRecordset.Fields("pmname").Value)
    vRecordset.MoveNext
    Wend
    End If
    vRecordset.Close
    
        Form401.ListView103.ListItems.Clear
        vQuery = "execute USP_PM_CheckedSearch"
        If OpenDatabaseBPlus(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                Set vListCheckedPromotion = Form401.ListView103.ListItems.Add(, , Trim(vRecordset.Fields("docno").Value))
                vListCheckedPromotion.SubItems(1) = Trim(vRecordset.Fields("secmanname").Value)
                If IsNull(Trim(vRecordset.Fields("createdate").Value)) Then
                vListCheckedPromotion.SubItems(2) = ""
                Else
                    vListCheckedPromotion.SubItems(2) = Trim(vRecordset.Fields("createdate").Value)
                End If
                If IsNull(Trim(vRecordset.Fields("checkdate").Value)) Then
                    vListCheckedPromotion.SubItems(3) = ""
                Else
                    vListCheckedPromotion.SubItems(3) = Trim(vRecordset.Fields("checkdate").Value)
                End If
                vListCheckedPromotion.SubItems(4) = Trim(vRecordset.Fields("checkername").Value)
                vListCheckedPromotion.SubItems(5) = Trim(vRecordset.Fields("pmname").Value)
            vRecordset.MoveNext
            Wend
        End If
        vRecordset.Close
End Sub

Private Sub ListView101_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

On Error GoTo ErrDescription

ListView101.Sorted = True
ListView101.SortKey = ColumnHeader.Index - 1
If vSortResult = 0 Then
    ListView101.SortOrder = lvwAscending
    vSortResult = 1
Else
    ListView101.SortOrder = lvwDescending
    vSortResult = 0
End If


ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub ListView101_DblClick()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocno As String
Dim vListItemSearch As ListItem
Dim i As Integer

On Error Resume Next

i = 0
vDocno = ListView101.SelectedItem
SSTab101.Tab = 1
ListView102.ListItems.Clear
    vQuery = "execute USP_PM_RequestSubSearch '" & vDocno & "' "
    If OpenDatabase(vGetConnect, vRecordset, vQuery) <> 0 Then
        Form401.Text102.Text = Trim(vRecordset.Fields("docno").Value)
        Form401.Text103.Text = Trim(vRecordset.Fields("docdate").Value)
        Form401.Text105.Text = Trim(vRecordset.Fields("pmname").Value)
        Form401.Text104.Text = Trim(vRecordset.Fields("secname").Value)
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            i = i + 1
            Set vListItemSearch = Form401.ListView102.ListItems.Add(, , Trim(vRecordset.Fields("itemcode").Value))
            vListItemSearch.SubItems(1) = Trim(vRecordset.Fields("itemname").Value)
            vListItemSearch.SubItems(2) = Trim(vRecordset.Fields("price").Value)
            vListItemSearch.SubItems(3) = Trim(vRecordset.Fields("promoprice").Value)
            vListItemSearch.SubItems(4) = Trim(vRecordset.Fields("discount").Value)
            vListItemSearch.SubItems(5) = Trim(vRecordset.Fields("unitcode").Value)
            vListItemSearch.SubItems(6) = Trim(vRecordset.Fields("iscancel").Value)
            If Trim(vRecordset.Fields("iscancel").Value) = 1 Then
                Form401.ListView102.ListItems(i).ForeColor = "&H000000FF"
                Form401.ListView102.ListItems.Item(i).ListSubItems(1).ForeColor = "&H000000FF"
                Form401.ListView102.ListItems.Item(i).ListSubItems(2).ForeColor = "&H000000FF"
                Form401.ListView102.ListItems.Item(i).ListSubItems(3).ForeColor = "&H000000FF"
                Form401.ListView102.ListItems.Item(i).ListSubItems(4).ForeColor = "&H000000FF"
                Form401.ListView102.ListItems.Item(i).ListSubItems(5).ForeColor = "&H000000FF"
                Form401.ListView102.ListItems.Item(i).ListSubItems(6).ForeColor = "&H000000FF"
            End If
            vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close

End Sub

Private Sub ListView101_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ErrDescription

Text106.Text = Trim(Item.Text)

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub ListView101_KeyPress(KeyAscii As Integer)
Dim i As Integer

On Error GoTo ErrDescription

If KeyAscii = 32 Then
    If vCheckSelect1 = 0 Then
        For i = ListView101.ListItems.Count To 1 Step -1
            If ListView101.ListItems(i).Selected = True Then
                ListView101.ListItems(i).Checked = True
            End If
        Next i
        vCheckSelect1 = 1
    Else
        For i = ListView101.ListItems.Count To 1 Step -1
            If ListView101.ListItems(i).Selected = True Then
            ListView101.ListItems(i).Checked = False
            End If
        Next i
        vCheckSelect1 = 0
    End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub SSTab101_Click(PreviousTab As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vListCheckedPromotion As ListItem

On Error GoTo ErrDescription

If SSTab101.Tab = 2 Then
    ListView103.ListItems.Clear
    vQuery = "execute USP_PM_CheckedSearch"
    If OpenDatabase(vGetConnect, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            Set vListCheckedPromotion = ListView103.ListItems.Add(, , Trim(vRecordset.Fields("docno").Value))
            vListCheckedPromotion.SubItems(1) = Trim(vRecordset.Fields("secmanname").Value)
            If IsNull(Trim(vRecordset.Fields("createdate").Value)) Then
            vListCheckedPromotion.SubItems(2) = ""
            Else
                vListCheckedPromotion.SubItems(2) = Trim(vRecordset.Fields("createdate").Value)
            End If
            If IsNull(Trim(vRecordset.Fields("checkdate").Value)) Then
                vListCheckedPromotion.SubItems(3) = ""
            Else
                vListCheckedPromotion.SubItems(3) = Trim(vRecordset.Fields("checkdate").Value)
            End If
            vListCheckedPromotion.SubItems(4) = Trim(vRecordset.Fields("checkername").Value)
            vListCheckedPromotion.SubItems(5) = Trim(vRecordset.Fields("pmname").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vSearch As String
Dim vListCheckPromotion As ListItem

On Error GoTo ErrDescription

If KeyAscii = 13 Then
    If Text101.Text <> "" Then
        vSearch = Trim(Text101.Text)
        ListView101.ListItems.Clear
        vQuery = "exec USP_PM_CheckOrConfirmSearch '0','" & vSearch & "' "
        If OpenDatabase(vGetConnect, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    Set vListCheckPromotion = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("docno").Value))
                    vListCheckPromotion.SubItems(1) = Trim(vRecordset.Fields("secmanname").Value)
                    If IsNull(Trim(vRecordset.Fields("createdate").Value)) Then
                        vListCheckPromotion.SubItems(2) = ""
                    Else
                        vListCheckPromotion.SubItems(2) = Trim(vRecordset.Fields("createdate").Value)
                    End If
                    If IsNull(Trim(vRecordset.Fields("editdate").Value)) Then
                        vListCheckPromotion.SubItems(3) = ""
                    Else
                        vListCheckPromotion.SubItems(3) = Trim(vRecordset.Fields("editdate").Value)
                    End If
                    vListCheckPromotion.SubItems(4) = Trim(vRecordset.Fields("pmname").Value)
            vRecordset.MoveNext
            Wend
        End If
        vRecordset.Close
        Text101.Text = ""
        ListView101.SetFocus
    End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

