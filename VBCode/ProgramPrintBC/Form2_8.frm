VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form2_8 
   Caption         =   "พิมพ์ใบเสนอซื้อสินค้า"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form2_8.frx":0000
   ScaleHeight     =   8415
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMD103 
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
      Height          =   375
      Left            =   6345
      TabIndex        =   10
      Top             =   4410
      Width           =   1590
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   2475
      Top             =   7020
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
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   4455
      Width           =   3750
   End
   Begin VB.CommandButton CMD101 
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
      Height          =   420
      Left            =   4365
      TabIndex        =   5
      Top             =   1035
      Width           =   1545
   End
   Begin VB.CommandButton CMD102 
      Caption         =   "พิมพ์"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4590
      TabIndex        =   4
      Top             =   5940
      Width           =   1320
   End
   Begin VB.TextBox Text102 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   3825
      TabIndex        =   3
      Top             =   5490
      Width           =   2085
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   3825
      TabIndex        =   2
      Top             =   5040
      Width           =   2085
   End
   Begin MSComctlLib.ListView ListView102 
      Height          =   2535
      Left            =   6345
      TabIndex        =   1
      Top             =   1575
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ชื่อฟอร์ม"
         Object.Width           =   6438
      EndProperty
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   2535
      Left            =   1080
      TabIndex        =   0
      Top             =   1575
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "เลขที่เอกสาร"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "วันที่เอกสาร"
         Object.Width           =   4904
      EndProperty
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสพนักงาน"
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
      Left            =   1080
      TabIndex        =   9
      Top             =   4410
      Width           =   1095
   End
   Begin VB.Label Label2 
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
      Height          =   240
      Left            =   2700
      TabIndex        =   7
      Top             =   5490
      Width           =   1140
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   2700
      TabIndex        =   6
      Top             =   5040
      Width           =   1140
   End
End
Attribute VB_Name = "Form2_8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vUserPrint As String

Private Sub CMB101_Click()
Call RefreshData
Text101.Text = ""
End Sub

Private Sub CMD101_Click()
If CMB101.Text <> "" Then
    Call RefreshData
Else
    MsgBox "กรุณาเลือก พนักงานก่อนแล้วค่อยกดฟื้นฟูข้อมูล", vbInformation, "Send Information"
    Exit Sub
End If
End Sub

Private Sub CMD102_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vPrint As Integer

On Error GoTo ErrDescription

vDocNo = Trim(Text101.Text)
Call GetComputerandUser
If Text101.Text <> "" And Text102.Text <> "" Then
    If Text102.Text = Trim("พิมพ์ใบเสนอซื้อสินค้า") Then
        Call PrintStockRequest
    End If
    vQuery = "select printed from  bcnp.dbo.vw_sl_00001 where docno = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vPrint = Trim(vRecordset.Fields("printed").Value)
    End If
    vRecordset.Close
     If vPrint = 0 Then
        vQuery = "Update npmaster.dbo.npprintserver set Printed = 1 where Docno = '" & vDocNo & "' "
        gConnection.Execute vQuery
            ListView101.ListItems.Remove (ListView101.SelectedItem.Index)
    End If
    
Else
MsgBox "ไม่มีเลขที่ให้พิมพ์ หรือ ไม่ได้เลือกฟอร์มที่จะพิมพ์ หรือ ไม่ได้เลือกเลขที่เอกสาร", vbCritical, "ข้อความแจ้งเตือน"
End If

Text101.Text = ""
Text102.Text = ""

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD103_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
On Error Resume Next

If Text101.Text <> "" Then
vDocNo = Trim(Text101.Text)
vQuery = "Update npmaster.dbo.npprintserver set printed = 1 where docno = '" & vDocNo & "' "
gConnection.Execute vQuery
Text101.Text = ""
MsgBox "กรุณา Click ปุ่มฟื้นฟูอีกครั้ง", vbInformation, "ข้อความแจ้งให้ทราบ"
End If
End Sub

Private Sub Form_Load()
Dim vRecordset As New Recordset
Dim vQuery As String
Dim vTypeDoc As String
Dim RQListItem As ListItem
Dim RQListForm As ListItem

ListView101.ListItems.Clear
CMB101.Clear
vQuery = "select * from vw_NP_SaleUserID "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    CMB101.AddItem Trim(vRecordset.Fields("salename").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

vTypeDoc = "RQ"
vQuery = "Select Docno,LastPrintDateTime  from NPMaster.dbo.NPPrintServer  where Printed = 0  " _
                       & " and LastPrintedUser = '" & vUserID & "' and DoctypeID = '" & vTypeDoc & "'  order by LastPrintDateTime "
       If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
           If Not vRecordset.EOF Then
               vRecordset.MoveFirst
               While Not vRecordset.EOF
               Set RQListItem = ListView101.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
               RQListItem.SubItems(1) = Trim(vRecordset.Fields("LastPrintDateTime").Value)
               vRecordset.MoveNext
               Wend
           End If
       End If
       vRecordset.Close
        
    ListView102.ListItems.Clear
    vQuery = "select Name from npmaster.dbo.NPForm where ModuleID = '" & vTypeDoc & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vRecordset.MoveFirst
        While Not vRecordset.EOF
        Set RQListForm = ListView102.ListItems.Add(, , Trim(vRecordset.Fields("Name").Value))
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close

End Sub

Private Sub ListView101_ItemClick(ByVal Item As MSComctlLib.ListItem)
Text101.Text = Trim(Item)
End Sub

Private Sub ListView102_ItemClick(ByVal Item As MSComctlLib.ListItem)
Text102.Text = Trim(Item)
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

        vDocHeader = "RQ"
        vUserprint1 = Left(Trim(CMB101.Text), InStr(Trim(CMB101.Text), "-") - 1)
        vUserPrint = vUserprint1
        ListView101.ListItems.Clear
        vQuery = "Select Docno,name1,printed,lastprintdatetime  from BCNP.dbo.vw_sl_00004   where Printed = 0 " _
                            & " and salecode = '" & vUserPrint & "' and DoctypeID = '" & vDocHeader & "' order by LastPrintDateTime "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    If Not vRecordset.EOF Then
                        CountRecordset = vRecordset.RecordCount
                        CountList = ListView101.ListItems.Count
                                If CountRecordset > CountList Then
                                vRecordset.MoveFirst
                                For i = 1 To CountRecordset
                                If i < CountRecordset Then
                                End If
                                vNewDoc = Trim(vRecordset.Fields("Docno").Value)
                                vPrintStatus = Trim(vRecordset.Fields("Printed").Value)
                                        If vDocNo <> vNewDoc Then
                                        Set DocListItem = ListView101.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
                                        DocListItem.SubItems(1) = Trim(vRecordset.Fields("lastprintdatetime").Value)
                                        End If
                                vRecordset.MoveNext
                            Next i
                            ElseIf CountRecordset < CountList Then
                            Call NewListItems
                            End If
                    End If
        End If
        vRecordset.Close

End Function

Function NewListItems()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim INVItemLists As ListItem
Dim vTypeDoc As String

On Error GoTo ErrDescription
    vTypeDoc = "RQ"
    ListView101.ListItems.Clear
    vQuery = "Select Docno,LastPrintDateTime  from NPMaster.dbo.NPPrintServer  where Printed = 0  " _
                        & " and LastPrintedUser = '" & vUserID & "' and DoctypeID = '" & vTypeDoc & "'  order by LastPrintDateTime "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            If Not vRecordset.EOF Then
                vRecordset.MoveFirst
                While Not vRecordset.EOF
                Set INVItemLists = ListView101.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
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

Public Sub PrintStockRequest()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vGroupDoc As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String
Dim vRunNumber As String
Dim vHeaderType As Integer
Dim vNamePrint  As String
Dim vPicking  As String
Dim vHeader As String
Dim vDocNumber As String
Dim vDocuments As String
Dim vTypeNumber As Integer
Dim vCheckDocNo  As Integer
Dim vCheckHeader As String

On Error GoTo ErrDescription

    vDocNo = UCase(Trim(Text101.Text))
    vCheckHeader = Left(vDocNo, 3)
    vQuery = "select docno,runnumber  from npmaster.dbo.TB_DC_RunNumberDocumentLogs where docno = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vCheckDocNo = 1 'เคยพิมพ์แล้ว
        vRunNumber = Trim(vRecordset.Fields("runnumber").Value)
    Else
        vCheckDocNo = 0 'ยังไม่ได้พิมพ์
    End If
    vRecordset.Close
    vHeaderType = 25
    vNamePrint = Trim(vUserPrint)
If vCheckDocNo = 0 Then
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
    
    vTypeNumber = 6
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
        '--------------------------------------------------------------------------------------------
        If vCheckHeader = "PRE" Then
         vRepID = 351
        Else
         vRepID = 293
        End If
        
        vRepType = "RQ"
        vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
        'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            With Crystal101
                .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                .ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
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


