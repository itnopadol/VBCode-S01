VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form301 
   Caption         =   "อนุมัติสินค้าโปรโมชั่น"
   ClientHeight    =   9000
   ClientLeft      =   6405
   ClientTop       =   1065
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   959
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab101 
      Height          =   8865
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   14250
      _ExtentX        =   25135
      _ExtentY        =   15637
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
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
      TabCaption(0)   =   "รายการ ใบเสนอสินค้าโปรโมชั่น รออนุมัติ"
      TabPicture(0)   =   "Form203.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView101"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CMD101"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CMD102"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text102"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Crystal1011"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Crystal101"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "ใบเสนอสินค้าโปรโมชั่น อนุมัติแล้ว"
      TabPicture(1)   =   "Form203.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "ListView102"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin Crystal.CrystalReport Crystal101 
         Left            =   -66900
         Top             =   7695
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.PictureBox Crystal1011 
         Height          =   480
         Left            =   -69555
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   12
         Top             =   7875
         Width           =   1200
      End
      Begin VB.TextBox Text102 
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
         Left            =   -73500
         TabIndex        =   11
         Top             =   2700
         Width           =   2190
      End
      Begin VB.CommandButton CMD102 
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
         Left            =   -72750
         TabIndex        =   9
         Top             =   7350
         Width           =   1740
      End
      Begin MSComctlLib.ListView ListView102 
         Height          =   6165
         Left            =   375
         TabIndex        =   7
         Top             =   1350
         Width           =   11040
         _ExtentX        =   19473
         _ExtentY        =   10874
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "เลขที่เอกสาร"
            Object.Width           =   2734
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ผู้เสนอ"
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
      End
      Begin VB.CommandButton CMD101 
         Caption         =   "อนุมัติ ใบเสนอสินค้า"
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
         Left            =   -74700
         TabIndex        =   6
         Top             =   7330
         Width           =   1740
      End
      Begin MSComctlLib.ListView ListView101 
         Height          =   3840
         Left            =   -74700
         TabIndex        =   2
         Top             =   3210
         Width           =   11265
         _ExtentX        =   19870
         _ExtentY        =   6773
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "เลขที่ใบเสนอสินค้าโปรโมชั่น"
            Object.Width           =   2734
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ผู้เสนอสินค้า"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "วันที่ทำใบเสนอ"
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
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Frame Frame1 
         Caption         =   "ค้นหา รายการเอกสารใบเสนอสินค้าโปรโมชั่น"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   -74700
         TabIndex        =   1
         Top             =   735
         Width           =   11265
         Begin VB.TextBox Text101 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1275
            TabIndex        =   4
            Top             =   675
            Width           =   4290
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
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
            Left            =   750
            TabIndex        =   3
            Top             =   675
            Width           =   540
         End
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
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
         Left            =   -74625
         TabIndex        =   10
         Top             =   2700
         Width           =   990
      End
      Begin VB.Label Label3 
         Caption         =   "เลขที่ใบเสนอสินค้าโปรโมชั่นที่อนุมัติเรียบร้อยแล้ว"
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
         TabIndex        =   8
         Top             =   1050
         Width           =   4215
      End
      Begin VB.Label Label2 
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
         Left            =   -74625
         TabIndex        =   5
         Top             =   2310
         Width           =   2415
      End
   End
End
Attribute VB_Name = "Form301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vSortResult As Integer

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim i As Integer
Dim vCheckDate As Date
Dim vDocno As String
Dim vMydescription As String

On Error GoTo ErrDescription


vCheckDate = Trim(Now)
For i = ListView101.ListItems.Count To 1 Step -1
    If ListView101.ListItems.Item(i).Checked = True Then
        vDocno = ListView101.ListItems.Item(i).Text
        vMydescription = "Promotion Confirm"
        vQuery = "exec USP_PM_RQConfirm  '" & vDocno & "','' "
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

Private Sub CMD102_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vDocno As String

On Error GoTo ErrDescription


    vDocno = Trim(Text102.Text)
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
        Text102.Text = ""
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

Private Sub ListView101_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ErrDescription

Text102.Text = Trim(Item.Text)

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
    If vCheckSelect2 = 0 Then
        For i = ListView101.ListItems.Count To 1 Step -1
            If ListView101.ListItems(i).Selected = True Then
                ListView101.ListItems(i).Checked = True
            End If
        Next i
        vCheckSelect2 = 1
    Else
        For i = ListView101.ListItems.Count To 1 Step -1
            If ListView101.ListItems(i).Selected = True Then
            ListView101.ListItems(i).Checked = False
            End If
        Next i
        vCheckSelect2 = 0
    End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub SSTab101_DblClick()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vListConfirmItem As ListItem

On Error GoTo ErrDescription

If SSTab101.Tab = 1 Then
    ListView102.ListItems.Clear
    vQuery = "exec USP_PM_CheckOrConfirmSearch '2' "
    If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            Set vListConfirmItem = ListView102.ListItems.Add(, , Trim(vRecordset.Fields("docno").Value))
            If IsNull(Trim(vRecordset.Fields("secmanname").Value)) Then
                vListConfirmItem.SubItems(1) = ""
             Else
                vListConfirmItem.SubItems(1) = Trim(vRecordset.Fields("secmanname").Value)
            End If
            If IsNull(Trim(vRecordset.Fields("createdate").Value)) Then
            vListConfirmItem.SubItems(2) = ""
            Else
                vListConfirmItem.SubItems(2) = Trim(vRecordset.Fields("createdate").Value)
            End If
            If IsNull(Trim(vRecordset.Fields("checkdate").Value)) Then
                vListConfirmItem.SubItems(3) = ""
            Else
                vListConfirmItem.SubItems(3) = Trim(vRecordset.Fields("checkdate").Value)
            End If
            If IsNull(Trim(vRecordset.Fields("checkername").Value)) Then
                vListConfirmItem.SubItems(4) = ""
            Else
                vListConfirmItem.SubItems(4) = Trim(vRecordset.Fields("checkername").Value)
            End If
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
Dim vListConfirmPromotion As ListItem

On Error GoTo ErrDescription

If KeyAscii = 13 Then
    If Text101.Text <> "" Then
        vSearch = Trim(Text101.Text)
        ListView101.ListItems.Clear
        vQuery = "exec USP_PM_CheckOrConfirmSearch '1','" & vSearch & "' "
        If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
                    Set vListConfirmPromotion = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("docno").Value))
                    vListConfirmPromotion.SubItems(1) = Trim(vRecordset.Fields("secmanname").Value)
                    If IsNull(Trim(vRecordset.Fields("createdate").Value)) Then
                    vListConfirmPromotion.SubItems(2) = ""
                    Else
                        vListConfirmPromotion.SubItems(2) = Trim(vRecordset.Fields("createdate").Value)
                    End If
                    If IsNull(Trim(vRecordset.Fields("checkdate").Value)) Then
                        vListConfirmPromotion.SubItems(3) = ""
                    Else
                        vListConfirmPromotion.SubItems(3) = Trim(vRecordset.Fields("checkdate").Value)
                    End If
                    vListConfirmPromotion.SubItems(4) = Trim(vRecordset.Fields("chekcername").Value)
                    vListConfirmPromotion.SubItems(5) = Trim(vRecordset.Fields("pmname").Value)
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
