VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form20 
   Caption         =   "หน้าพิมพ์ใบสั่งซื้อ"
   ClientHeight    =   8280
   ClientLeft      =   2580
   ClientTop       =   1260
   ClientWidth     =   12000
   Icon            =   "Frm1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Frm1.frx":08CA
   ScaleHeight     =   8280
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport201 
      Left            =   5160
      Top             =   6480
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
   Begin MSComctlLib.ListView ListView201 
      Height          =   3090
      Left            =   2175
      TabIndex        =   9
      Top             =   1650
      Width           =   4890
      _ExtentX        =   8625
      _ExtentY        =   5450
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "เลขที่ใบสั่งซื้อ"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "วันที่ใบสั่งซื้อ"
         Object.Width           =   4958
      EndProperty
   End
   Begin MSComctlLib.ListView ListView202 
      Height          =   3090
      Left            =   7350
      TabIndex        =   8
      Top             =   1650
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   5450
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
         Text            =   "ฟอร์มที่พิมพ์ใบสั่งซื้อ"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.CommandButton CMD202 
      Caption         =   "RefreshData"
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
      Left            =   5550
      TabIndex        =   7
      Top             =   1200
      Width           =   1515
   End
   Begin VB.Timer Timer201 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   9825
      Top             =   7800
   End
   Begin VB.TextBox TXT202 
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
      Height          =   390
      Left            =   2175
      TabIndex        =   2
      Top             =   5775
      Width           =   2640
   End
   Begin VB.TextBox TXT201 
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
      Height          =   390
      Left            =   2175
      TabIndex        =   1
      Top             =   5100
      Width           =   2640
   End
   Begin VB.CommandButton CMD201 
      Caption         =   "พิมพ์ใบสั่งซื้อ"
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
      Left            =   8850
      TabIndex        =   0
      Top             =   5100
      Width           =   1590
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "พิมพ์ใบสั่งซื้อ/ใบตรวจรับสินค้า"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   540
      Left            =   2550
      TabIndex        =   6
      Top             =   300
      Width           =   7365
   End
   Begin VB.Label LBLTime201 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   8475
      TabIndex        =   5
      Top             =   975
      Width           =   1890
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
      Height          =   390
      Left            =   975
      TabIndex        =   4
      Top             =   5775
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบสั่งซื้อ"
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
      Height          =   390
      Left            =   975
      TabIndex        =   3
      Top             =   5100
      Width           =   1140
   End
End
Attribute VB_Name = "Form20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD201_Click()
Dim vDocNo As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vPrint As Integer

On Error Resume Next

vDocNo = Trim(TXT201.Text)
Call GetComputerandUser
If TXT201.Text <> "" And TXT202.Text <> "" Then

    If TXT202.Text = "พิมพ์ใบสั่งซื้อ" Then
            Call PrintPO
    ElseIf TXT202.Text = "พิมพ์ใบอนุมัติค่าใช้จ่าย(สำหรับบุคคล)" Then
            Call PrintPOExpense
    End If
    
     vQuery = "select printed from  bcnp.dbo.vw_sl_00001 where docno = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vPrint = Trim(vRecordset.Fields("printed").Value)
    End If
    vRecordset.Close
    
     If vPrint = 0 Then
        vQuery = "Update npmaster.dbo.npprintserver set Printed = 1 where Docno = '" & vDocNo & "' "
        gConnection.Execute vQuery
        ListView201.ListItems.Remove (ListView201.SelectedItem.Index)
    End If
Else
MsgBox "ไม่มีเลขที่ให้พิมพ์ หรือ ไม่ได้เลือกฟอร์มที่จะพิมพ์ หรือ ไม่ได้เลือกเลขที่เอกสาร", vbCritical, "ข้อความแจ้งเตือน"
End If

TXT201.Text = ""
TXT202.Text = ""

End Sub

Public Sub PrintPOExpense()
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
Dim vCheckList As Integer

On Error GoTo ErrDescription

vDocNo = UCase(Trim(TXT201.Text))
    vQuery = "select docno,runnumber  from npmaster.dbo.TB_DC_RunNumberDocumentLogs where docno = '" & vDocNo & "' and typenumber = 5"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckDocNo = 1 'เคยพิมพ์แล้ว
            vRunNumber = Trim(vRecordset.Fields("runnumber").Value)
        Else
            vCheckDocNo = 0 'ยังไม่ได้พิมพ์
        End If
        vRecordset.Close
        
        If vCheckDocNo = 0 Then 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน
            vHeaderType = 15
            vNamePrint = Trim(vUserID)
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
            
            vTypeNumber = 5
            vQuery = "exec dbo.usp_DC_InsertRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "'," & vTypeNumber & " "
            gConnection.Execute vQuery
        Else
            vDocuments = vRunNumber
            vQuery = "exec dbo.usp_DC_PrintCopyRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "' "
            gConnection.Execute vQuery
    End If 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน
    

        vDocNo = Trim(TXT201.Text)
        vQuery = "select groupdoc from  dbo.vw_sl_00001 where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------
                            vRepType = "PO"
                           If vGroupDoc = "POE" Or vGroupDoc = "poe" Then
                                vRepID = 418
                            Else
                               MsgBox "เลขที่เอกสารไม่สามารถพิมพ์ฟอร์มนี้ได้  กรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
                               Exit Sub
                            End If
                            
                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport201
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


Private Sub CMD202_Click()
Call RefreshData
End Sub

Private Sub Form_Load()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim POItemLists As ListItem
Dim POItemforms As ListItem
Dim vTypeDoc As String

On Error GoTo ErrDescription
 vTypeDoc = "PO"

ListView201.ListItems.Clear
 vQuery = "Select *  from dbo.vw_SL_00005  where  " _
                        & " LastPrintedUser = '" & vUserID & "' and DoctypeID = '" & vTypeDoc & "'  order by LastPrintDateTime "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            If Not vRecordset.EOF Then
                vRecordset.MoveFirst
                While Not vRecordset.EOF
                Set POItemLists = ListView201.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
                POItemLists.SubItems(1) = Trim(vRecordset.Fields("LastPrintDateTime").Value)
                vRecordset.MoveNext
                Wend
            End If
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------------------
        
        ListView202.ListItems.Clear
        vQuery = "select Name from npmaster.dbo.NPForm where ModuleID = '" & vTypeDoc & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
            Set POItemforms = ListView202.ListItems.Add(, , Trim(vRecordset.Fields("Name").Value))
            vRecordset.MoveNext
            Wend
        End If
        vRecordset.Close
        '--------------------------------------------------------------------------------------------------------

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub PrintPO()
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
Dim vCheckList As Integer

On Error GoTo ErrDescription

vDocNo = UCase(Trim(TXT201.Text))
    vQuery = "select docno,runnumber  from npmaster.dbo.TB_DC_RunNumberDocumentLogs where docno = '" & vDocNo & "' and typenumber = 5"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckDocNo = 1 'เคยพิมพ์แล้ว
            vRunNumber = Trim(vRecordset.Fields("runnumber").Value)
        Else
            vCheckDocNo = 0 'ยังไม่ได้พิมพ์
        End If
        vRecordset.Close
        
        If vCheckDocNo = 0 Then 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน
            vHeaderType = 15
            vNamePrint = Trim(vUserID)
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
            
            vTypeNumber = 5
            vQuery = "exec bcnp.dbo.usp_DC_InsertRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "'," & vTypeNumber & " "
            gConnection.Execute vQuery
        Else
            vDocuments = vRunNumber
            vQuery = "exec bcnp.dbo.usp_DC_PrintCopyRunNumberDocument '" & vDocNo & "','" & vDocuments & "','" & vUserID & "' "
            gConnection.Execute vQuery
    End If 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน
    

        vDocNo = Trim(TXT201.Text)
        vQuery = "select groupdoc from  bcnp.dbo.vw_sl_00001 where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------
                            vRepType = "PO"
                            If vGroupDoc = "POV" Or vGroupDoc = "pov" Or vGroupDoc = "POC" Or vGroupDoc = "poc" Then
                               vRepID = 59
                            ElseIf vGroupDoc = "POE" Or vGroupDoc = "poe" Then
                               vRepID = 352
                            ElseIf vGroupDoc = "PON" Or vGroupDoc = "pon" Then
                               vRepID = 60
                            Else
                            MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
                            Exit Sub
                            End If
                            
                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport201
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

Private Sub ListView201_ItemClick(ByVal Item As MSComctlLib.ListItem)
TXT201.Text = Item
End Sub

Private Sub ListView202_ItemClick(ByVal Item As MSComctlLib.ListItem)
TXT202.Text = Item
End Sub

Public Sub PrintPOCheck()
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

On Error Resume Next

    vDocNo = UCase(Trim(TXT201.Text))
    vQuery = "select docno,runnumber  from npmaster.dbo.TB_DC_RunNumberDocumentLogs where docno = '" & vDocNo & "' and typenumber = 6"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vCheckDocNo = 1 'เคยพิมพ์แล้ว
            vRunNumber = Trim(vRecordset.Fields("runnumber").Value)
        Else
            vCheckDocNo = 0 'ยังไม่ได้พิมพ์
        End If
        vRecordset.Close
        
        If vCheckDocNo = 0 Then 'เช็คว่าพิมพ์ครั้งแรกหรือ ทดแทน
            vHeaderType = 22
            vNamePrint = Trim(vUserID)
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
        '---------------------------------------------------------------------------------------------
        
                            vRepType = "PO"
                            If vGroupDoc = "POV" Or vGroupDoc = "pov" Or vGroupDoc = "POC" Or vGroupDoc = "poc" Or vGroupDoc = "POE" Or vGroupDoc = "poe" Then
                                vRepID = 61
                            ElseIf vGroupDoc = "PON" Or vGroupDoc = "pon" Then
                                vRepID = 61
                            Else
                            MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
                            Exit Sub
                            End If
                            
                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport201
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
End Sub

Private Sub Timer201_Timer()
Dim vTime, vSubTime, vTotalTime
Dim vNumber As Integer

On Error GoTo ErrDescription

            If LBLTime201.Caption <> CStr(Time) Then
                 LBLTime201.Caption = Time
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

On Error GoTo ErrDescription

            vDocHeader = "PO"

            vQuery = "Select  *  from dbo.vw_SL_00005   where Printed = 0 " _
                        & " and LastPrintedUser = '" & vUserID & "' and DoctypeID = '" & vDocHeader & "' order by LastPrintDateTime "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    If Not vRecordset.EOF Then
                        CountRecordset = vRecordset.RecordCount
                        CountList = ListView201.ListItems.Count
                                If CountRecordset > CountList Then
                                vRecordset.MoveFirst
                                For i = 1 To CountRecordset
                                If i < CountRecordset Then
                                vDocNo = ListView201.ListItems.Item(i).Text
                                End If
                                vNewDoc = Trim(vRecordset.Fields("Docno").Value)
                                vPrintStatus = Trim(vRecordset.Fields("Printed").Value)
                                        If vDocNo <> vNewDoc Then
                                        Set DocListItem = ListView201.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
                                        DocListItem.SubItems(1) = Trim(vRecordset.Fields("LastPrintDateTime").Value)
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

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Function

Function NewListItems()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim INVItemLists As ListItem
Dim vTypeDoc As String

On Error GoTo ErrDescription
    vTypeDoc = "PO"
    ListView201.ListItems.Clear
    vQuery = "Select Docno,LastPrintDateTime  from NPMaster.dbo.NPPrintServer  where Printed = 0  " _
                        & " and LastPrintedUser = '" & vUserID & "' and DoctypeID = '" & vTypeDoc & "'  order by LastPrintDateTime "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            If Not vRecordset.EOF Then
                vRecordset.MoveFirst
                While Not vRecordset.EOF
                Set INVItemLists = ListView201.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
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


