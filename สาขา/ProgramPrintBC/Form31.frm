VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form31 
   Caption         =   "หน้าพิมพ์ใบ Back Order"
   ClientHeight    =   8355
   ClientLeft      =   2355
   ClientTop       =   420
   ClientWidth     =   12000
   Icon            =   "Form31.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form31.frx":08CA
   ScaleHeight     =   8355
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.CheckBox CHKReqPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ขอพิมพ์ฟอร์ม A4 กรณีพิมพ์กระดาษครึ่งหน้าไม่ได้"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7695
      TabIndex        =   14
      Top             =   5400
      Width           =   3930
   End
   Begin Crystal.CrystalReport CrystalReport311 
      Left            =   900
      Top             =   6600
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
      Left            =   1035
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1425
      Width           =   3165
   End
   Begin VB.CommandButton CMDClearDocuments 
      Caption         =   "Clear เอกสารที่ไม่ต้องพิมพ์"
      Height          =   390
      Left            =   5085
      TabIndex        =   10
      Top             =   6525
      Width           =   2190
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4185
      Top             =   6375
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.CommandButton CMD312 
      Height          =   315
      Left            =   4260
      Picture         =   "Form31.frx":7BC5
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1425
      Width           =   315
   End
   Begin VB.CommandButton CMD311 
      Caption         =   "พิมพ์ใบ Back Order"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      TabIndex        =   6
      Top             =   5895
      Width           =   2265
   End
   Begin VB.TextBox TXT312 
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
      Left            =   5085
      TabIndex        =   3
      Top             =   5925
      Width           =   2190
   End
   Begin VB.TextBox TXT311 
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
      Left            =   5085
      TabIndex        =   2
      Top             =   5400
      Width           =   2190
   End
   Begin VB.Timer Timer311 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   9750
      Top             =   7725
   End
   Begin MSComctlLib.ListView ListView312 
      Height          =   3240
      Left            =   7680
      TabIndex        =   1
      Top             =   1875
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   5715
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
         Object.Width           =   3704
      EndProperty
   End
   Begin MSComctlLib.ListView ListView311 
      Height          =   3240
      Left            =   1035
      TabIndex        =   0
      Top             =   1875
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   5715
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
         Text            =   "เลขที่ใบBack Order"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อลูกค้า"
         Object.Width           =   4340
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "วันที่เอกสาร"
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
      Height          =   240
      Left            =   7680
      TabIndex        =   13
      Top             =   1575
      Width           =   1590
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
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
      Height          =   240
      Left            =   1035
      TabIndex        =   11
      Top             =   1125
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "พิมพ์ใบ BackOrder"
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
   Begin VB.Label LBLTime311 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   7800
      TabIndex        =   7
      Top             =   0
      Width           =   2190
   End
   Begin VB.Label LBL312 
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
      Height          =   255
      Left            =   2595
      TabIndex        =   5
      Top             =   5925
      Width           =   2430
   End
   Begin VB.Label LBL311 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบBack Order :"
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
      Height          =   255
      Left            =   2910
      TabIndex        =   4
      Top             =   5400
      Width           =   2115
   End
End
Attribute VB_Name = "Form31"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vUserPrint As String

Private Sub CMBSale_Click()
Call RefreshData
TXT311.Text = ""
End Sub

Private Sub CMD311_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vPrint As Integer

On Error GoTo ErrDescription

'Call GeneratePR
vDocNo = Trim(TXT311.Text)
Call GetComputerandUser
If TXT311.Text <> "" And TXT312.Text <> "" Then
    If TXT312.Text = "พิมพ์ใบBackOrder" Then
        Call PrintBackOrder
    End If
    vQuery = "select printed from  bcnp.dbo.vw_sl_00001 where docno = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vPrint = Trim(vRecordset.Fields("printed").Value)
    End If
    vRecordset.Close
     If vPrint = 0 Then
        vQuery = "Update npmaster.dbo.npprintserver set Printed = 1 where Docno = '" & vDocNo & "' "
        gConnection.Execute vQuery
        ListView311.ListItems.Remove (ListView311.SelectedItem.Index)
    End If
Else
MsgBox "ไม่มีเลขที่ให้พิมพ์ หรือ ไม่ได้เลือกฟอร์มที่จะพิมพ์ หรือ ไม่ได้เลือกเลขที่เอกสาร", vbCritical, "ข้อความแจ้งเตือน"
End If

TXT311.Text = ""
TXT312.Text = ""

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD312_Click()
Call RefreshData
End Sub

Private Sub CMDClearDocuments_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
On Error Resume Next

vDocNo = Trim(TXT311.Text)
vQuery = "Update npmaster.dbo.npprintserver set printed = 1 where docno = '" & vDocNo & "' "
gConnection.Execute vQuery
TXT311.Text = ""
MsgBox "กรุณา Click ที่รูปอีกครั้ง", vbInformation, "ข้อความแจ้งให้ทราบ"
End Sub

Private Sub Form_Load()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim BOListItems As ListItem
Dim BOItemforms As ListItem
Dim vTypeDoc As String
Dim ListX As ListItem
Dim x As ListImage
Dim i As Integer
Dim SOPListItems As ListItem


On Error GoTo ErrDescription

 vTypeDoc = "BO"
ListView311.ListItems.Clear
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
                Set BOListItems = ListView311.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
                BOListItems.SubItems(1) = Trim(vRecordset.Fields("LastPrintDateTime").Value)
                vRecordset.MoveNext
                Wend
            End If
        End If
        vRecordset.Close
        '--------------------------------------------------------------------------------------------------------------------
        ListView312.ListItems.Clear
        vQuery = "select Name from npmaster.dbo.NPForm where ModuleID = '" & vTypeDoc & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
            Set BOItemforms = ListView312.ListItems.Add(, , Trim(vRecordset.Fields("Name").Value))
            vRecordset.MoveNext
            Wend
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------------------------------
        
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub ListView311_ItemClick(ByVal Item As MSComctlLib.ListItem)
TXT311.Text = Item
End Sub

Private Sub ListView312_ItemClick(ByVal Item As MSComctlLib.ListItem)
TXT312.Text = Item
End Sub

Private Sub Timer311_Timer()
Dim vTime, vSubTime, vTotalTime
Dim vNumber As Integer

On Error GoTo ErrDescription

            If LBLTime311.Caption <> CStr(Time) Then
                 LBLTime311.Caption = Time
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
vDocHeader = "BO"
vUserprint1 = Left(Trim(CMBSale.Text), InStr(Trim(CMBSale.Text), "-") - 1)
vUserPrint = vUserprint1
ListView311.ListItems.Clear

vQuery = "Select Docno,name1,printed,lastprintdatetime  from BCNP.dbo.vw_sl_00003   where Printed = 0 " _
                & " and salecode = '" & vUserprint1 & "' and DoctypeID = '" & vDocHeader & "' order by LastPrintDateTime "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    If Not vRecordset.EOF Then
                        CountRecordset = vRecordset.RecordCount
                        CountList = ListView311.ListItems.Count
                                If CountRecordset > CountList Then
                                vRecordset.MoveFirst
                                For i = 1 To CountRecordset
                                If i < CountRecordset Then
                                End If
                                vNewDoc = Trim(vRecordset.Fields("Docno").Value)
                                vPrintStatus = Trim(vRecordset.Fields("Printed").Value)
                                        If vDocNo <> vNewDoc Then
                                        Set DocListItem = ListView311.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
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
'------------------------------------------------------------------------------------------------

End Function

Function NewListItems()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim INVItemLists As ListItem
Dim vTypeDoc As String

On Error GoTo ErrDescription
    vTypeDoc = "BO"
    ListView311.ListItems.Clear
    vQuery = "Select Docno,LastPrintDateTime  from NPMaster.dbo.NPPrintServer  where Printed = 0  " _
                        & " and LastPrintedUser = '" & vUserID & "' and DoctypeID = '" & vTypeDoc & "'  order by LastPrintDateTime "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            If Not vRecordset.EOF Then
                vRecordset.MoveFirst
                While Not vRecordset.EOF
                Set INVItemLists = ListView311.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
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

Public Sub PrintBackOrder()
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

vDocNo = UCase(Trim(TXT311.Text))
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
                vHeaderType = 18
            ElseIf vCheckBillType = 1 Then
                vHeaderType = 19
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
            
            vTypeNumber = 4
            vQuery = "exec bcnp.dbo.usp_DC_InsertRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "'," & vTypeNumber & " "
            gConnection.Execute vQuery
        Else
            vDocuments = vRunNumber
            vQuery = "exec bcnp.dbo.usp_DC_PrintCopyRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "' "
            gConnection.Execute vQuery
    End If 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน
        

        vDocNo = Trim(TXT311.Text)
        vQuery = "select groupdoc from  bcnp.dbo.vw_sl_00001 where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------
        
        vRepType = "BO"
                            
        If (UCase(vDepartment) = "CH" Or UCase(vDepartment) = "CR" Or UCase(vDepartment) = "IS" Or UCase(vDepartment) = "PC") And Me.CHKReqPrint.Value = 0 Then
           If UCase(vGroupDoc) = "BHV" Or vGroupDoc = "bhv" Or UCase(vGroupDoc) = "BCV" Or vGroupDoc = "bcv" Or UCase(vGroupDoc) = "BVD" Or vGroupDoc = "bvd" Or UCase(vGroupDoc) = "BVM" Or vGroupDoc = "bvm" Or UCase(vGroupDoc) = "BAB" Or vGroupDoc = "bab" Then
               vRepID = 392
           ElseIf UCase(vGroupDoc) = "BHN" Or vGroupDoc = "bhn" Or UCase(vGroupDoc) = "BCN" Or vGroupDoc = "bcn" Or UCase(vGroupDoc) = "BVN" Or vGroupDoc = "bvn" Then
               vRepID = 393
           Else
              MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
              Exit Sub
           End If
        Else
           If UCase(vGroupDoc) = "BHV" Or vGroupDoc = "bhv" Or UCase(vGroupDoc) = "BCV" Or vGroupDoc = "bcv" Or UCase(vGroupDoc) = "BVD" Or vGroupDoc = "bvd" Or UCase(vGroupDoc) = "BVM" Or vGroupDoc = "bvm" Or UCase(vGroupDoc) = "BAB" Or vGroupDoc = "bab" Then
               vRepID = 74
           ElseIf UCase(vGroupDoc) = "BHN" Or vGroupDoc = "bhn" Or UCase(vGroupDoc) = "BCN" Or vGroupDoc = "bcn" Or UCase(vGroupDoc) = "BVN" Or vGroupDoc = "bvn" Then
               vRepID = 75
           Else
              MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
              Exit Sub
           End If
        End If

                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport311
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

Public Sub PrintBackOrderA5()
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

vDocNo = UCase(Trim(TXT311.Text))
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
                vHeaderType = 18
            ElseIf vCheckBillType = 1 Then
                vHeaderType = 19
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
            
            vTypeNumber = 4
            vQuery = "exec bcnp.dbo.usp_DC_InsertRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "'," & vTypeNumber & " "
            gConnection.Execute vQuery
        Else
            vDocuments = vRunNumber
            vQuery = "exec bcnp.dbo.usp_DC_PrintCopyRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "' "
            gConnection.Execute vQuery
    End If 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน
        

        vDocNo = Trim(TXT311.Text)
        vQuery = "select groupdoc from  bcnp.dbo.vw_sl_00001 where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------
        
        vRepType = "BO"
                            
        If (UCase(vDepartment) = "CH" Or UCase(vDepartment) = "CR" Or UCase(vDepartment) = "IS" Or UCase(vDepartment) = "PC") And Me.CHKReqPrint.Value = 0 Then
           If UCase(vGroupDoc) = "BHV" Or vGroupDoc = "bhv" Or UCase(vGroupDoc) = "BCV" Or vGroupDoc = "bcv" Or UCase(vGroupDoc) = "BVD" Or vGroupDoc = "bvd" Or UCase(vGroupDoc) = "BVM" Or vGroupDoc = "bvm" Or UCase(vGroupDoc) = "BAB" Or vGroupDoc = "bab" Then
               vRepID = 392
           ElseIf UCase(vGroupDoc) = "BHN" Or vGroupDoc = "bhn" Or UCase(vGroupDoc) = "BCN" Or vGroupDoc = "bcn" Or UCase(vGroupDoc) = "BVN" Or vGroupDoc = "bvn" Then
               vRepID = 393
           Else
              MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
              Exit Sub
           End If
        Else
           If UCase(vGroupDoc) = "BHV" Or vGroupDoc = "bhv" Or UCase(vGroupDoc) = "BCV" Or vGroupDoc = "bcv" Or UCase(vGroupDoc) = "BVD" Or vGroupDoc = "bvd" Or UCase(vGroupDoc) = "BVM" Or vGroupDoc = "bvm" Or UCase(vGroupDoc) = "BAB" Or vGroupDoc = "bab" Then
               vRepID = 74
           ElseIf UCase(vGroupDoc) = "BHN" Or vGroupDoc = "bhn" Or UCase(vGroupDoc) = "BCN" Or vGroupDoc = "bcn" Or UCase(vGroupDoc) = "BVN" Or vGroupDoc = "bvn" Then
               vRepID = 75
           Else
              MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
              Exit Sub
           End If
        End If

                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport311
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

Public Sub GeneratePR()
Dim vRecordset As New ADODB.Recordset
Dim vRecordset1 As New ADODB.Recordset
Dim vQuery As String
Dim vNewPR As String
Dim vLastPRNo, vLastPRNo1, vLastPRNo2, vLastPRNo3 As String
Dim vChkDoc, vChkDoc1, vChkDoc2, vChkDoc3 As String
Dim vAutoNumber, vCount, vChkYear As Integer
Dim vDocNo As String
Dim i, j, vLineNumber As Integer
Dim vApCode(10) As String
Dim vCheckDocExist As Integer


vDocNo = Trim(TXT311.Text)
i = 0
vCheckDocExist = 0
vQuery = "select * from vw_AP_BackOrderGenPR  where docno = '" & vDocNo & "'  "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        i = i + 1
        vApCode(i) = Trim(vRecordset.Fields("apcode").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close
 For j = 1 To i
Line1:
If vCheckDocExist = 1 Then
vCheckDocExist = 0
vRecordset.Close
End If
 vQuery = "select top 1 docno from dbo.bcstkrequest order by docno desc "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vLastPRNo = Trim(vRecordset.Fields("docno").Value)
End If
vRecordset.Close
vAutoNumber = Right(vLastPRNo, 4)
vLastPRNo1 = Right(vLastPRNo, Len(vLastPRNo) - 2)
vLastPRNo2 = UCase(Left(Right(vLastPRNo1, Len(vLastPRNo1) - InStr(vLastPRNo1, "-")), 2)) 'Left(vLastPRNo1, 2)
vLastPRNo3 = Mid(vLastPRNo1, 3, 2)
vChkDoc = Month(Now)
vChkDoc1 = Right(Year(Now), 2)
If Len(vChkDoc) = 1 Then
    vChkDoc = "0" & vChkDoc
End If
If vChkDoc1 < "48" Then
vChkYear = vChkDoc1 + 43
End If
vChkDoc3 = vChkYear

If vLastPRNo2 = vChkDoc3 Then
    If vLastPRNo3 = vChkDoc Then
        vNewPR = UCase(Trim("PR" & vChkYear & vChkDoc & "-" & Format(vAutoNumber + 1, "00000")))
    Else
        MsgBox "เลขที่เอกสารที่สร้างได้ไม่ถูกต้อง กรุณาแจ้งแผนกคอมพิวเตอร์"
    End If
Else
    vNewPR = UCase(Trim("PR" & vChkYear & vChkDoc & "-00001"))
End If

 vQuery = "select  * from vw_AP_BackOrderGenPR  where docno = '" & vDocNo & "' and apcode = '" & vApCode(i) & "' "
 If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 
 vQuery = "set dateformat dmy"
 gConnection.Execute vQuery
 
 vQuery = "select docno from dbo.bcstkrequest where docno = '" & vNewPR & "' "
 If OpenDataBase(vConnection, vRecordset1, vQuery) <> 0 Then
    vCheckDocExist = 1
 End If
 vRecordset1.Close
 
 If vCheckDocExist <> 1 Then
 vQuery = "insert into  dbo.bcstkrequest (DocNo,DocDate,TaxType,DepartCode,IsConfirm,MyDescription,WorkMan,BillStatus,IsCancel,CreatorCode,CreateDateTime)" _
                    & " values('" & vNewPR & "','" & vRecordset.Fields("docdate").Value & "',0 ,'CC' ,0 ,'" & vRecordset.Fields("MyDescription").Value & "' , 'somrod' ,0 ,0 ,'somrod' ,getdate() ) "
 gConnection.Execute vQuery
  Else
  GoTo Line1
 End If
 
 End If
 vRecordset.Close

 
 vLineNumber = 0
 vQuery = "select * from vw_AP_BackOrderGenPRSub where docno = '" & vDocNo & "' and apcode = '" & vApCode(i) & "' "
 If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    vLineNumber = vLineNumber + 1
    
    vQuery = "insert into dbo.bcstkrequestsub (DocNo,TaxType,ItemCode,DocDate,DepartCode,MyDescription,Qty, RemainQty,UnitCode,LineNumber,ItemName,packingrate1,packingrate2)" _
                    & " values('" & vNewPR & "',0,'" & vRecordset.Fields("itemcode").Value & "','" & vRecordset.Fields("docdate").Value & "','" & vRecordset.Fields("departcode").Value & "',''," & vRecordset.Fields("Qty").Value & "," & vRecordset.Fields("remainQty").Value & ",'" & vRecordset.Fields("unitcode").Value & "' ," & vLineNumber & ",'" & vRecordset.Fields("itemname").Value & "',1,1) "
    gConnection.Execute vQuery
    vRecordset.MoveNext
    Wend
 End If
 vRecordset.Close
 MsgBox " Back Order เลขที่ " & vDocNo & " Generate ใบเสนอซื้อได้หมายเลขที่ " & vNewPR & " "
 Next j

End Sub

