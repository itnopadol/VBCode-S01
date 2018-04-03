VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form30 
   Caption         =   "หน้าพิมพ์ใบเสนอราคา"
   ClientHeight    =   9000
   ClientLeft      =   2280
   ClientTop       =   1260
   ClientWidth     =   12000
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "Form3.frx":08CA
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal301 
      Left            =   2700
      Top             =   7695
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
   Begin VB.CheckBox CHK101 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ไม่โชว์ส่วนลด(เฉพาะขายโครงการ)"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   8010
      TabIndex        =   15
      Top             =   5895
      Width           =   3390
   End
   Begin MSComctlLib.ListView ListViewCondition 
      Height          =   2400
      Left            =   810
      TabIndex        =   14
      Top             =   4905
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   4233
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "เงื่อนไข"
         Object.Width           =   10231
      EndProperty
   End
   Begin Crystal.CrystalReport CrystalReport30 
      Left            =   11565
      Top             =   8505
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
   Begin VB.ComboBox CMBSale 
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
      Left            =   810
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1395
      Width           =   3315
   End
   Begin VB.CommandButton CMDClearDocument 
      Caption         =   "ยกเลิกการพิมพ์"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   9810
      TabIndex        =   10
      Top             =   6795
      Width           =   1560
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11430
      Top             =   7920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.CommandButton CMD302 
      Height          =   315
      Left            =   4185
      Picture         =   "Form3.frx":7BC5
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "ใช้ฟื้นฟูข้อมูล"
      Top             =   1395
      Width           =   315
   End
   Begin VB.Timer Timer301 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   11115
      Top             =   8505
   End
   Begin VB.CommandButton CMD301 
      Caption         =   "พิมพ์ใบเสนอราคา"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   8010
      TabIndex        =   6
      Top             =   6795
      Width           =   1515
   End
   Begin VB.TextBox TXT302 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9270
      TabIndex        =   3
      Top             =   5355
      Width           =   2145
   End
   Begin VB.TextBox TXT301 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9270
      TabIndex        =   2
      Top             =   4860
      Width           =   2145
   End
   Begin MSComctlLib.ListView ListView302 
      Height          =   2865
      Left            =   8865
      TabIndex        =   1
      Top             =   1875
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   5054
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ฟอร์มที่พิมพ์"
         Object.Width           =   4410
      EndProperty
   End
   Begin MSComctlLib.ListView ListView301 
      Height          =   2820
      Left            =   810
      TabIndex        =   0
      Top             =   1890
      Width           =   7830
      _ExtentX        =   13811
      _ExtentY        =   4974
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "เลขที่ใบเสนอราคา"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อลูกค้า"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "วันที่ทำเอกสาร"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ฟอร์มที่พิมพ์"
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
      Height          =   165
      Left            =   8865
      TabIndex        =   13
      Top             =   1665
      Width           =   1890
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "เลือกพนักงานขาย"
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
      Height          =   315
      Left            =   810
      TabIndex        =   12
      Top             =   1080
      Width           =   1440
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "พิมพ์ใบเสนอราคา"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   465
      Left            =   2550
      TabIndex        =   8
      Top             =   300
      Width           =   7365
   End
   Begin VB.Label LBLTime301 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   8700
      TabIndex        =   7
      Top             =   1125
      Width           =   1740
   End
   Begin VB.Label LBL302 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ฟอร์มที่พิมพ์ :"
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
      Height          =   315
      Left            =   7875
      TabIndex        =   5
      Top             =   5400
      Width           =   1260
   End
   Begin VB.Label LBL301 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบเสนอราคา :"
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
      Height          =   315
      Left            =   7560
      TabIndex        =   4
      Top             =   4860
      Width           =   1590
   End
End
Attribute VB_Name = "Form30"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vUserPrint As String

Private Sub CMBSale_Click()
Call RefreshData
TXT301.Text = ""
End Sub

Private Sub CMD301_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vPrint As Integer
Dim i As Integer

On Error GoTo ErrDescription

vDocNo = Trim(TXT301.Text)

Call GetComputerandUser

If TXT301.Text <> "" And TXT302.Text <> "" Then
    If TXT302.Text = "พิมพ์ใบเสนอราคา" Then
        Call PrintQuotation
    Else
    Call PrintQuotationWholeSale
    End If
    vQuery = "select printed from  bcnp.dbo.vw_sl_00001 where docno = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vPrint = Trim(vRecordset.Fields("printed").Value)
    End If
    vRecordset.Close
     If vPrint = 0 Then
        vQuery = "Update npmaster.dbo.npprintserver set Printed = 1 where Docno = '" & vDocNo & "' "
        gConnection.Execute vQuery
            ListView301.ListItems.Remove (ListView301.SelectedItem.Index)
    End If
    
Else
MsgBox "ไม่มีเลขที่ให้พิมพ์ หรือ ไม่ได้เลือกฟอร์มที่จะพิมพ์ หรือ ไม่ได้เลือกเลขที่เอกสาร", vbCritical, "ข้อความแจ้งเตือน"
End If

TXT301.Text = ""
TXT302.Text = ""
For i = 1 To Me.ListViewCondition.ListItems.Count
Me.ListViewCondition.ListItems(i).Checked = False
Next i

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD302_Click()
Call RefreshData
End Sub

Private Sub CMDClearDocument_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
On Error Resume Next

vDocNo = Trim(TXT301.Text)
vQuery = "Update npmaster.dbo.npprintserver set printed = 1 where docno = '" & vDocNo & "' "
gConnection.Execute vQuery
TXT301.Text = ""
MsgBox "กรุณา Click ที่รูปอีกครั้ง", vbInformation, "ข้อความแจ้งให้ทราบ"
End Sub

Private Sub Form_Load()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim QTListItems As ListItem
Dim QTItemforms As ListItem
Dim vTypeDoc As String
Dim vPicture As String
Dim ListX As ListItem
Dim x As ListImage
Dim i As Integer
Dim SOPListItems As ListItem

On Error GoTo ErrDescription

 vTypeDoc = "QT"
 '   vQuery = "select salecode,address from bcpicturesale " 'where description = 'ขายหน้าร้าน' "
  '  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   '     i = 1
    '    While Not vRecordset.EOF
     '       Set x = ImageList1.ListImages.Add(, , LoadPicture(Trim(vRecordset.Fields("address").Value)))
      '                     ImageList1.ListImages.Item(i).Tag = Trim(vRecordset.Fields("salecode").Value)
       '     Set ListX = ListView303.ListItems.Add(, , Trim(vRecordset.Fields("salecode").Value), ImageList1.ListImages(i).Index)
        '    vRecordset.MoveNext
         '   i = i + 1
          '  Wend
    'End If
    'vRecordset.Close
ListView301.ListItems.Clear
CMBSale.Clear
vQuery = "select * from vw_NP_SaleUserID "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    CMBSale.AddItem Trim(vRecordset.Fields("salename").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

 vQuery = "Select Docno,LastPrintDateTime  from NPMaster.dbo.NPPrintServer  where Printed = 0  " _
                        & " and LastPrintedUser = '" & vUserID & "' and DoctypeID = '" & vTypeDoc & "'  order by LastPrintDateTime "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            If Not vRecordset.EOF Then
                vRecordset.MoveFirst
                While Not vRecordset.EOF
                Set QTListItems = ListView301.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
                QTListItems.SubItems(1) = Trim(vRecordset.Fields("LastPrintDateTime").Value)
                vRecordset.MoveNext
                Wend
            End If
        End If
        vRecordset.Close
        '--------------------------------------------------------------------------------------------------------------------
        ListView302.ListItems.Clear
        vQuery = "select Name from npmaster.dbo.NPForm where ModuleID = '" & vTypeDoc & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
            Set QTItemforms = ListView302.ListItems.Add(, , Trim(vRecordset.Fields("Name").Value))
            vRecordset.MoveNext
            Wend
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------------------------------
        
        Me.ListViewCondition.ListItems.Clear
        vQuery = "select * from npmaster.dbo.TB_NP_QuotationTextCondition order by roworder"
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
            Set QTItemforms = ListViewCondition.ListItems.Add(, , Trim(vRecordset.Fields("textcondition").Value))
            vRecordset.MoveNext
            Wend
        End If
        vRecordset.Close
        
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub ListView301_ItemClick(ByVal Item As MSComctlLib.ListItem)
TXT301.Text = Item
End Sub

Private Sub ListView302_ItemClick(ByVal Item As MSComctlLib.ListItem)
TXT302.Text = Item
End Sub

Private Sub Timer301_Timer()
Dim vTime, vSubTime, vTotalTime
Dim vNumber As Integer

On Error GoTo ErrDescription

            If LBLTime301.Caption <> CStr(Time) Then
                 LBLTime301.Caption = Time
            End If
             vTime = Second(Time)
             vTotalTime = vTime Mod 5
             If vTotalTime = 0 Then
             Call RefreshData
             End If
    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Function RefreshData()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim aaa As Integer
Dim i As Integer
Dim DocListItem As ListItem
Dim vDocNo, vNewDoc As String
Dim vPrintStatus As Integer
Dim CountRecordset As Integer, CountList As Integer
Dim vDocHeader As String
Dim vUserprint1 As String

On Error Resume Next

        vDocHeader = "QT"
        vUserprint1 = Left(Trim(CMBSale.Text), InStr(Trim(CMBSale.Text), "-") - 1)
        vUserPrint = vUserprint1
        ListView301.ListItems.Clear
        vQuery = "Select Docno,name1,printed,lastprintdatetime  from BCNP.dbo.vw_sl_00003   where Printed = 0 " _
                            & " and salecode = '" & vUserprint1 & "' and DoctypeID = '" & vDocHeader & "' order by LastPrintDateTime "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    If Not vRecordset.EOF Then
                        CountRecordset = vRecordset.RecordCount
                        CountList = ListView301.ListItems.Count
                                If CountRecordset > CountList Then
                                vRecordset.MoveFirst
                                For i = 1 To CountRecordset
                                If i < CountRecordset Then
                                End If
                                vNewDoc = Trim(vRecordset.Fields("Docno").Value)
                                vPrintStatus = Trim(vRecordset.Fields("Printed").Value)
                                        If vDocNo <> vNewDoc Then
                                        Set DocListItem = ListView301.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
                                        DocListItem.SubItems(1) = Trim(vRecordset.Fields("name1").Value)
                                        DocListItem.SubItems(2) = Trim(vRecordset.Fields("lastprintdatetime").Value)
                                        End If
                                vRecordset.MoveNext
                            Next i
                            ElseIf CountRecordset < CountList Then
                            Call NewListItems
                            End If
                    End If
        End If
        vRecordset.Close
'ErrDescription:
'If Err.Description <> "" Then
'MsgBox Err.Description
'End If
End Function

Function NewListItems()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim INVItemLists As ListItem
Dim vTypeDoc As String

On Error GoTo ErrDescription
    vTypeDoc = "QT"
    ListView301.ListItems.Clear
    vQuery = "Select Docno,LastPrintDateTime  from NPMaster.dbo.NPPrintServer  where Printed = 0  " _
                        & " and LastPrintedUser = '" & vUserID & "' and DoctypeID = '" & vTypeDoc & "'  order by LastPrintDateTime "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            If Not vRecordset.EOF Then
                vRecordset.MoveFirst
                While Not vRecordset.EOF
                Set INVItemLists = ListView301.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
                INVItemLists.SubItems(1) = Trim(vRecordset.Fields("LastPrintDateTime").Value)
                vRecordset.MoveNext
                Wend
            End If
        End If
        vRecordset.Close

        '----------------------------------------------------------------------------------------------------------------------
        
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Function

Public Sub PrintQuotation()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vGroupDoc As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String
Dim vRunNumber As String
Dim vCheckBillType As Integer
Dim vHeaderType As Integer
Dim vNamePrint  As String
Dim vPicking  As String
Dim vHeader As String
Dim vDocNumber As String
Dim vDocuments As String
Dim vTypeNumber As Integer
Dim vCheckDocNo  As Integer

On Error GoTo ErrDescription

vDocNo = UCase(Trim(TXT301.Text))
vQuery = "select docno,runnumber  from npmaster.dbo.TB_DC_RunNumberDocumentLogs where docno = '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckDocNo = 1 'เคยพิมพ์แล้ว
    vRunNumber = Trim(vRecordset.Fields("runnumber").Value)
Else
    vCheckDocNo = 0 'ยังไม่ได้พิมพ์
End If
vRecordset.Close

If vCheckDocNo = 0 Then 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน
    vQuery = "select docno,billtype from bcnp.dbo.BCQuotation where docno = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vCheckBillType = Trim(vRecordset.Fields("billtype").Value)
    End If
    vRecordset.Close
       
    If vCheckBillType = 0 Then
        vHeaderType = 16
    ElseIf vCheckBillType = 1 Then
        vHeaderType = 17
    End If
    
    vNamePrint = Trim(vUserPrint)
    vQuery = "select header,autonumber,docnumber  from npmaster.dbo.NP_Generate_DocNo where headertype = " & vHeaderType & " "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vPicking = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
        vHeader = Trim(vRecordset.Fields("header").Value)
        vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
    End If
    vRecordset.Close
    vDocuments = vDocNumber & vHeader & "-" & vPicking
    
    vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = " & vHeaderType & " "
    gConnection.Execute vQuery
    
    vTypeNumber = 3
    vQuery = "exec bcnp.dbo.usp_DC_InsertRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "'," & vTypeNumber & " "
    gConnection.Execute vQuery
Else
    vDocuments = vRunNumber
    vQuery = "exec bcnp.dbo.usp_DC_PrintCopyRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "' "
    gConnection.Execute vQuery
End If 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน
        
        vQuery = "select groupdoc from  bcnp.dbo.vw_sl_00001 where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------
                                    
                            vRepType = "QT"
                            If UCase(vGroupDoc) = "QHV" Or vGroupDoc = "qhv" Then
                                vRepID = 72
                            ElseIf UCase(vGroupDoc) = "QHN" Or vGroupDoc = "qhn" Then
                                vRepID = 73
                            ElseIf UCase(vGroupDoc) = "QCV" Or vGroupDoc = "qcv" Then
                                vRepID = 72
                            ElseIf UCase(vGroupDoc) = "QCN" Or vGroupDoc = "qcn" Then
                                vRepID = 73
                            ElseIf UCase(vGroupDoc) = "QVD" Or vGroupDoc = "qvd" Then
                                vRepID = 72
                            ElseIf UCase(vGroupDoc) = "QVM" Or vGroupDoc = "qvm" Then
                                vRepID = 72
                            ElseIf UCase(vGroupDoc) = "QVN" Or vGroupDoc = "qvn" Then
                                vRepID = 73
                            ElseIf UCase(vGroupDoc) = "QAB" Or vGroupDoc = "qab" Then
                                vRepID = 72
                            Else
                            MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
                            Exit Sub
                            End If
                            
                            'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport30
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Formulas(0) = "computername='" & vComputerName1 & "' "
                                    .Formulas(1) = "username='" & vUserName1 & "' "
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            '---------------------------------------------------------------------------------------------------
                                                                                  
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Public Sub PrintQuotationA5()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vGroupDoc As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String
Dim vRunNumber As String
Dim vCheckBillType As Integer
Dim vHeaderType As Integer
Dim vNamePrint  As String
Dim vPicking  As String
Dim vHeader As String
Dim vDocNumber As String
Dim vDocuments As String
Dim vTypeNumber As Integer
Dim vCheckDocNo  As Integer

On Error GoTo ErrDescription

vDocNo = UCase(Trim(TXT301.Text))
vQuery = "select docno,runnumber  from npmaster.dbo.TB_DC_RunNumberDocumentLogs where docno = '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckDocNo = 1 'เคยพิมพ์แล้ว
    vRunNumber = Trim(vRecordset.Fields("runnumber").Value)
Else
    vCheckDocNo = 0 'ยังไม่ได้พิมพ์
End If
vRecordset.Close

If vCheckDocNo = 0 Then 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน
    vQuery = "select docno,billtype from bcnp.dbo.BCQuotation where docno = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vCheckBillType = Trim(vRecordset.Fields("billtype").Value)
    End If
    vRecordset.Close
       
    If vCheckBillType = 0 Then
        vHeaderType = 16
    ElseIf vCheckBillType = 1 Then
        vHeaderType = 17
    End If
    
    vNamePrint = Trim(vUserPrint)
    vQuery = "select header,autonumber,docnumber  from npmaster.dbo.NP_Generate_DocNo where headertype = " & vHeaderType & " "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vPicking = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
        vHeader = Trim(vRecordset.Fields("header").Value)
        vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
    End If
    vRecordset.Close
    vDocuments = vDocNumber & vHeader & "-" & vPicking
    
    vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = " & vHeaderType & " "
    gConnection.Execute vQuery
    
    vTypeNumber = 3
    vQuery = "exec bcnp.dbo.usp_DC_InsertRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "'," & vTypeNumber & " "
    gConnection.Execute vQuery
Else
    vDocuments = vRunNumber
    vQuery = "exec bcnp.dbo.usp_DC_PrintCopyRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "' "
    gConnection.Execute vQuery
End If 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน
        
        vQuery = "select groupdoc from  bcnp.dbo.vw_sl_00001 where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------
                                    
                            vRepType = "QT"
                            If UCase(vGroupDoc) = "QHV" Or vGroupDoc = "qhv" Then
                                vRepID = 394
                            ElseIf UCase(vGroupDoc) = "QHN" Or vGroupDoc = "qhn" Then
                                vRepID = 395
                            ElseIf UCase(vGroupDoc) = "QCV" Or vGroupDoc = "qcv" Then
                                vRepID = 394
                            ElseIf UCase(vGroupDoc) = "QCN" Or vGroupDoc = "qcn" Then
                                vRepID = 395
                            ElseIf UCase(vGroupDoc) = "QVD" Or vGroupDoc = "qvd" Then
                                vRepID = 394
                            ElseIf UCase(vGroupDoc) = "QVM" Or vGroupDoc = "qvm" Then
                                vRepID = 394
                            ElseIf UCase(vGroupDoc) = "QVN" Or vGroupDoc = "qvn" Then
                                vRepID = 395
                            ElseIf UCase(vGroupDoc) = "QAB" Or vGroupDoc = "qab" Then
                                vRepID = 394
                            Else
                            MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
                            Exit Sub
                            End If
                            
                            'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport30
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Formulas(0) = "computername='" & vComputerName1 & "' "
                                    .Formulas(1) = "username='" & vUserName1 & "' "
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            '---------------------------------------------------------------------------------------------------
                                                                                  
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Public Sub PrintQuotationWholeSale()
Dim vDocNo As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vGroupDoc As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String
Dim vRunNumber As String
Dim vCheckBillType As Integer
Dim vHeaderType As Integer
Dim vNamePrint  As String
Dim vPicking  As String
Dim vHeader As String
Dim vDocNumber As String
Dim vDocuments As String
Dim vTypeNumber As Integer
Dim vCheckDocNo  As Integer
Dim vCondition1 As String
Dim vCondition2 As String
Dim vCondition3 As String
Dim vCondition4 As String
Dim vCondition5 As String
Dim vIndexCondition1 As Integer
Dim vIndexCondition2 As Integer
Dim vIndexCondition3 As Integer
Dim vIndexCondition4 As Integer
Dim vIndexCondition5 As Integer
Dim i As Integer


vDocNo = UCase(Trim(TXT301.Text))
vQuery = "select docno,runnumber  from npmaster.dbo.TB_DC_RunNumberDocumentLogs where docno = '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckDocNo = 1 'เคยพิมพ์แล้ว
    vRunNumber = Trim(vRecordset.Fields("runnumber").Value)
Else
    vCheckDocNo = 0 'ยังไม่ได้พิมพ์
End If
vRecordset.Close

If vCheckDocNo = 0 Then 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน
    vQuery = "select docno,billtype from bcnp.dbo.BCQuotation where docno = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vCheckBillType = Trim(vRecordset.Fields("billtype").Value)
    End If
    vRecordset.Close
    If vCheckBillType = 0 Then
        vHeaderType = 16
    ElseIf vCheckBillType = 1 Then
        vHeaderType = 17
    End If
    
    
    vNamePrint = Trim(vUserPrint)
    vQuery = "select header,autonumber,docnumber  from npmaster.dbo.NP_Generate_DocNo where headertype = " & vHeaderType & " "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vPicking = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
        vHeader = Trim(vRecordset.Fields("header").Value)
        vDocNumber = Trim(vRecordset.Fields("docnumber").Value)
    End If
    vRecordset.Close
    vDocuments = vDocNumber & vHeader & "-" & vPicking
    
    vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = " & vHeaderType & " "
    gConnection.Execute vQuery
    
    vTypeNumber = 3
    vQuery = "exec bcnp.dbo.usp_DC_InsertRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "'," & vTypeNumber & " "
    gConnection.Execute vQuery
Else
    vDocuments = vRunNumber
    vQuery = "exec bcnp.dbo.usp_DC_PrintCopyRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "' "
    gConnection.Execute vQuery
End If 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน
        
        vQuery = "select groupdoc from  bcnp.dbo.vw_sl_00001 where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
        End If
        vRecordset.Close
        
        vRepType = "QT"
        
        If Me.CHK101.Value = 0 Then
           If UCase(vGroupDoc) = "QHV" Or vGroupDoc = "qhv" Then
               vRepID = 333
           ElseIf UCase(vGroupDoc) = "QHN" Or vGroupDoc = "qhn" Then
               vRepID = 334
           ElseIf UCase(vGroupDoc) = "QCV" Or vGroupDoc = "qcv" Then
               vRepID = 333
           ElseIf UCase(vGroupDoc) = "QCN" Or vGroupDoc = "qcn" Then
               vRepID = 334
           ElseIf UCase(vGroupDoc) = "QVD" Or vGroupDoc = "qvd" Then
               vRepID = 333
           ElseIf UCase(vGroupDoc) = "QVM" Or vGroupDoc = "qvm" Then
               vRepID = 333
           ElseIf UCase(vGroupDoc) = "QVN" Or vGroupDoc = "qvn" Then
               vRepID = 334
           ElseIf UCase(vGroupDoc) = "QAB" Or vGroupDoc = "qab" Then
               vRepID = 333
           Else
           MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
           Exit Sub
           End If
        Else
           If UCase(vGroupDoc) = "QHV" Or vGroupDoc = "qhv" Then
               vRepID = 463
           ElseIf UCase(vGroupDoc) = "QHN" Or vGroupDoc = "qhn" Then
               vRepID = 464
           ElseIf UCase(vGroupDoc) = "QCV" Or vGroupDoc = "qcv" Then
               vRepID = 463
           ElseIf UCase(vGroupDoc) = "QCN" Or vGroupDoc = "qcn" Then
               vRepID = 464
           ElseIf UCase(vGroupDoc) = "QVD" Or vGroupDoc = "qvd" Then
               vRepID = 463
           ElseIf UCase(vGroupDoc) = "QVM" Or vGroupDoc = "qvm" Then
               vRepID = 463
           ElseIf UCase(vGroupDoc) = "QVN" Or vGroupDoc = "qvn" Then
               vRepID = 464
           ElseIf UCase(vGroupDoc) = "QAB" Or vGroupDoc = "qab" Then
               vRepID = 463
           Else
           MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
           Exit Sub
           End If
        End If
For i = 1 To Me.ListViewCondition.ListItems.Count
    If Me.ListViewCondition.ListItems(i).Checked = True And vCondition1 = "" Then
          vCondition1 = Me.ListViewCondition.ListItems(i).Text
          vIndexCondition1 = i
    i = Me.ListViewCondition.ListItems.Count
    Else
    vIndexCondition1 = Me.ListViewCondition.ListItems.Count
    End If
Next i

If vIndexCondition1 <> Me.ListViewCondition.ListItems.Count And vIndexCondition1 <> 0 Then
For i = vIndexCondition1 + 1 To Me.ListViewCondition.ListItems.Count
    If Me.ListViewCondition.ListItems(i).Checked = True And vCondition2 = "" Then
          vCondition2 = Me.ListViewCondition.ListItems(i).Text
          vIndexCondition2 = i
    i = Me.ListViewCondition.ListItems.Count
    Else
    vIndexCondition2 = Me.ListViewCondition.ListItems.Count
    End If
Next i
End If

If vIndexCondition2 <> Me.ListViewCondition.ListItems.Count And vIndexCondition2 <> 0 Then
For i = vIndexCondition2 + 1 To Me.ListViewCondition.ListItems.Count
    If Me.ListViewCondition.ListItems(i).Checked = True And vCondition3 = "" Then
          vCondition3 = Me.ListViewCondition.ListItems(i).Text
          vIndexCondition3 = i
    i = Me.ListViewCondition.ListItems.Count
    Else
    vIndexCondition3 = Me.ListViewCondition.ListItems.Count
    End If
Next i
End If

If vIndexCondition3 <> Me.ListViewCondition.ListItems.Count And vIndexCondition3 <> 0 Then
For i = vIndexCondition3 + 1 To Me.ListViewCondition.ListItems.Count
    If Me.ListViewCondition.ListItems(i).Checked = True And vCondition4 = "" Then
          vCondition4 = Me.ListViewCondition.ListItems(i).Text
          vIndexCondition4 = i
    i = Me.ListViewCondition.ListItems.Count
    Else
    vIndexCondition4 = Me.ListViewCondition.ListItems.Count
    End If
Next i
End If

If vIndexCondition4 <> Me.ListViewCondition.ListItems.Count And vIndexCondition4 <> 0 Then
For i = vIndexCondition4 + 1 To Me.ListViewCondition.ListItems.Count
    If Me.ListViewCondition.ListItems(i).Checked = True And vCondition5 = "" Then
          vCondition5 = Me.ListViewCondition.ListItems(i).Text
    i = Me.ListViewCondition.ListItems.Count
    End If
Next i
End If

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = " & vRepID & " and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With Crystal301
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Formulas(0) = "Condition1='" & vCondition1 & "' "
        .Formulas(1) = "Condition2='" & vCondition2 & "' "
        .Formulas(2) = "Condition3='" & vCondition3 & "' "
        .Formulas(3) = "Condition4='" & vCondition4 & "' "
        .Formulas(4) = "Condition5='" & vCondition5 & "' "
        .Action = 1
    End With
End If
vRecordset.Close
End Sub

