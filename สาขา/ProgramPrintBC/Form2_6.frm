VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form2_6 
   Caption         =   "รวมสินค้าในใบ PR ที่อนุมัติแล้ว"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form2_6.frx":0000
   ScaleHeight     =   8115
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   4920
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6840
      Left            =   675
      TabIndex        =   3
      Top             =   900
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   12065
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   794
      WordWrap        =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "รวม PR ที่เสนอซื้อสินค้าเหมือนกัน"
      TabPicture(0)   =   "Form2_6.frx":72FB
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "ดูรายละเอียดการรวม PR"
      TabPicture(1)   =   "Form2_6.frx":7317
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         Caption         =   "Gen PQ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6015
         Left            =   720
         TabIndex        =   5
         Top             =   675
         Width           =   9165
         Begin VB.TextBox Text103 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6075
            TabIndex        =   16
            Top             =   4500
            Width           =   2565
         End
         Begin VB.CommandButton CMD102 
            Caption         =   "พิมพ์ใบเสนอซื้อ"
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
            Left            =   7275
            TabIndex        =   15
            Top             =   4950
            Width           =   1365
         End
         Begin VB.CommandButton CMD101 
            Caption         =   "รวม PR"
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
            Left            =   600
            TabIndex        =   2
            Top             =   4575
            Width           =   1365
         End
         Begin VB.TextBox Text101 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1875
            TabIndex        =   0
            Top             =   375
            Width           =   2190
         End
         Begin MSComctlLib.ListView ListView101 
            Height          =   2865
            Left            =   600
            TabIndex        =   1
            Top             =   1500
            Width           =   8040
            _ExtentX        =   14182
            _ExtentY        =   5054
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "เลขที่ใบเสนอซื้อสินค้า"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "วันที่ทำใบเสนอซื้อ"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "คำอธิบาย"
               Object.Width           =   7937
            EndProperty
         End
         Begin VB.Label Label3 
            Caption         =   "ได้เลขที่เอกสารเสนอซื้อ/เลขที่อนุมัติเลขที่"
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
            Left            =   2850
            TabIndex        =   11
            Top             =   4575
            Width           =   3165
         End
         Begin VB.Label Label102 
            BackColor       =   &H80000009&
            Height          =   315
            Left            =   5700
            TabIndex        =   10
            Top             =   900
            Width           =   2940
         End
         Begin VB.Label Label4 
            Caption         =   "วันที่ทำใบเสนอซื้อ"
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
            Left            =   4200
            TabIndex        =   9
            Top             =   900
            Width           =   1515
         End
         Begin VB.Label Label101 
            BackColor       =   &H80000009&
            Height          =   315
            Left            =   1875
            TabIndex        =   8
            Top             =   900
            Width           =   2190
         End
         Begin VB.Label Label2 
            Caption         =   "ผู้เสนอซื้อสินค้า"
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
            Left            =   600
            TabIndex        =   7
            Top             =   975
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "เลขที่ใบเสนอซื้อ"
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
            Left            =   600
            TabIndex        =   6
            Top             =   375
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "ดูรายละเอียดข้อมูล"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5190
         Left            =   -74280
         TabIndex        =   4
         Top             =   675
         Width           =   9165
         Begin MSComctlLib.ListView ListView102 
            Height          =   3240
            Left            =   375
            TabIndex        =   14
            Top             =   1125
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   5715
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   9
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "เลขที่ PR"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "เลขที่ PQ"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "เลขที่ AQ"
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
               Text            =   "จำนวนที่อนุมัติ"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "หน่วยนับ"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "ผู้ทำรายการ"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "วันที่ทำรายการ"
               Object.Width           =   2646
            EndProperty
         End
         Begin VB.TextBox Text102 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1500
            TabIndex        =   12
            Top             =   525
            Width           =   1815
         End
         Begin VB.Label Label5 
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
            Height          =   315
            Left            =   375
            TabIndex        =   13
            Top             =   525
            Width           =   1065
         End
      End
   End
End
Attribute VB_Name = "Form2_6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vGenCondition As String
Dim vCount As Integer
Dim i As Integer
Dim vCountPR As Integer
Dim vCountPRCheck As Integer
Dim vPrNO(100) As String
Dim vQuery As String
Dim vCheckDocNo As Integer
Dim vInsertRequestNo As Integer
Dim vGetPQ As String
Dim vCheckBoxValue As Integer

On Error GoTo ErrDescription
    
 If ListView101.ListItems.Count <> 0 Then
     vCount = ListView101.ListItems.Count
     vCountPR = 0
     vCountPRCheck = 0
    vCheckBoxValue = 0
     For i = 1 To vCount
        If ListView101.ListItems(i).Checked = True Then
        vCheckBoxValue = 1
        vGetPQ = ListView101.ListItems(i).Text
        vCountPR = vCountPR + 1
        vCountPRCheck = vCountPRCheck + 1
        vPrNO(vCountPRCheck) = ListView101.ListItems(i).Text
        
        vQuery = "select prno from npmaster.dbo.TB_PR_MergePR where prno = '" & vPrNO(vCountPRCheck) & "'  "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckDocNo = 1
            MsgBox "เอกสารใบเสนอซื้อสินค้า เลขที่ " & vPrNO(vCountPRCheck) & " ได้มีการรวมกับเอกสารอื่นแล้ว  กรุณาตรวจสอบด้วยครับ "
            Exit Sub
        Else
            vCheckDocNo = 0
        End If
        vRecordset.Close
    
        vQuery = "exec USP_PR_InsertTableMergeQTY '" & vPrNO(vCountPRCheck) & "','" & vUserID & "' "
        gConnection.Execute vQuery
        End If
        vInsertRequestNo = 1
     Next i
     
     If vCheckBoxValue = 1 Then
        If vInsertRequestNo = 1 Then
        vQuery = "exec USP_PR_ApprovePRGenPQ '" & vUserID & "' "
        gConnection.Execute vQuery
            vQuery = "select (mergeno+' / '+ mergeapprove) as Docno from npmaster.dbo.TB_PR_MergePR where prno = '" & vGetPQ & "' "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
               Text103.Text = Trim(vRecordset.Fields("Docno").Value)
            End If
            vRecordset.Close
        End If
        ListView101.ListItems.Clear
     Else
        MsgBox "กรุณาเลือก เลขที่ PR ที่จะรวมด้วยนะครับ"
     End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD102_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String, vReportName As String
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

vDocNo = Trim(Text103.Text)
If vDocNo <> "" Then

vRepID = 234
vRepType = "PR"
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = 234 and reptype = 'PR' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal101
.ReportFileName = vReportName & ".rpt"
.ParameterFields(0) = "@vDocno;" & vDocNo & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
Else
    MsgBox "กรอกข้อมูลที่จะพิมพ์ใบเสนอซื้อสินค้าด้วยนะครับ"
End If
Text103.Text = ""

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Form_Activate()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCheckUserID As Integer

'If vChkFrmActivate <> 1 Then
 '   vQuery = "select userid from npmaster.dbo.TB_CK_UserIDUseMergePR where userid = '" & vUserID & "' "
  '  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   '     vCheckUserID = 1
    '    MsgBox "มี UserID : " & vUserID & " เข้ามาใช้งานในหน้านี้แล้ว กรุณาตรวจสอบ ไม่งั้นจะไม่สามารถทำการรวม PR ได้ กรุณาติดต่อคอมฯนะครับ"
    'Else
     '   vCheckUserID = 0
    'End If
    'vRecordset.Close
    'If vCheckUserID = 0 Then
     '   Form2_6.Show
      '  Form2_6.SetFocus
       ' vQuery = "insert into npmaster.dbo.TB_CK_UserIDUseMergePR (UserID,ActiveStatus,ActivateDateTime) " _
                            & " values ('" & vUserID & "',1,getdate())"
        'gConnection.Execute (vQuery)
        
    'End If
'End If
End Sub

Private Sub Form_Deactivate()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

'vQuery = "delete npmaster.dbo.TB_CK_UserIDUseMergePR  where userid = '" & vUserID & "' "
'gConnection.Execute (vQuery)
End Sub

Private Sub Form_GotFocus()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCheckUserID As Integer

'vQuery = "select userid from npmaster.dbo.TB_CK_UserIDUseMergePR where userid = '" & vUserID & "' "
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 '   vCheckUserID = 1
  '  MsgBox "มี UserID : " & vUserID & " เข้ามาใช้งานในหน้านี้แล้ว กรุณาตรวจสอบ ไม่งั้นจะไม่สามารถทำการรวม PR ได้ กรุณาติดต่อคอมฯนะครับ"
'Else
 '   vCheckUserID = 0
'End If
'vRecordset.Close
'If vCheckUserID = 0 Then
 '   Form2_6.Show
  '  Form2_6.SetFocus
   ' vQuery = "insert into npmaster.dbo.TB_CK_UserIDUseMergePR (UserID,ActiveStatus,ActivateDateTime) " _
                        & " values ('" & vUserID & "',1,getdate())"
    'gConnection.Execute (vQuery)
    
'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error Resume Next
If vCheckDuplicate <> 1 Then
    vQuery = "delete npmaster.dbo.TB_CK_UserActivateProgram  where userid = '" & vUserID & "' and jobid = 1"
    gConnection.Execute (vQuery)
    vChkFrmActivate = 0
End If
End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vPrNO As String
Dim vItemList As ListItem

On Error GoTo ErrDescription

If KeyAscii = 13 Then
   If Text101.Text <> "" Then
        vPrNO = Trim(Text101.Text)
        vQuery = "exec USP_PR_MergePR '" & vPrNO & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            Label101.Caption = Trim(vRecordset.Fields("creatorcode"))
            Label102.Caption = Trim(vRecordset.Fields("createdatetime"))
            vRecordset.MoveFirst
            While Not vRecordset.EOF
            Set vItemList = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("docno").Value))
            vItemList.SubItems(1) = Trim(vRecordset.Fields("docdate").Value)
            vItemList.SubItems(2) = Trim(vRecordset.Fields("mydescription").Value)
            vRecordset.MoveNext
            Wend
        Else
            MsgBox "ไม่มีใบเสนอซื้อสินค้าเลขที่ " & vPrNO & " ที่มีสินค้าที่ได้ทำการอนุมัติ กรุณาตรวจสอบด้วยครับ"
            Text101.SetFocus
        End If
        vRecordset.Close
        Else
        MsgBox "กรุณา กรอกข้อมูลเลขที่ใบเสนอซื้อสินค้า"
        End If
        Text101.Text = ""
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Text102_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCheckTypeDocNo As Integer
Dim vDocNo As String
Dim vListPR As ListItem

On Error GoTo ErrDescription

If KeyAscii = 13 Then
    ListView102.ListItems.Clear
    vDocNo = Trim(Text102.Text)
    If UCase(Left(vDocNo, 2)) = "PR" Then
        vCheckTypeDocNo = 1
     ElseIf UCase(Left(vDocNo, 2)) = "PQ" Then
        vCheckTypeDocNo = 2
     ElseIf UCase(Left(vDocNo, 2)) = "AQ" Then
        vCheckTypeDocNo = 3
    End If
    
    vQuery = "exec USP_PR_SearchData '" & vCheckTypeDocNo & "','" & vDocNo & "'"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
            Set vListPR = ListView102.ListItems.Add(, , Trim(vRecordset.Fields("prno").Value))
            vListPR.SubItems(1) = Trim(vRecordset.Fields("mergeno").Value)
            vListPR.SubItems(2) = Trim(vRecordset.Fields("mergeapprove").Value)
            vListPR.SubItems(3) = Trim(vRecordset.Fields("itemcode").Value)
            vListPR.SubItems(4) = Trim(vRecordset.Fields("itemname").Value)
            vListPR.SubItems(5) = Trim(vRecordset.Fields("confirmqty").Value)
            vListPR.SubItems(6) = Trim(vRecordset.Fields("aqunitcode").Value)
            vListPR.SubItems(7) = Trim(vRecordset.Fields("userid").Value)
            vListPR.SubItems(8) = Trim(vRecordset.Fields("mergedatetime").Value)
            vRecordset.MoveNext
        Wend
    Else
        MsgBox "ไม่มีเอกสารเลขที่ '" & vDocNo & "' นี้ในระบบใบเสนอซื้อสินค้า "
        Text102.SetFocus
    End If
    vRecordset.Close
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub
