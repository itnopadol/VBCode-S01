VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form FrmCheckPrint 
   BackColor       =   &H00FF0000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "พิมพ์เอกสาร : 19012018"
   ClientHeight    =   1200
   ClientLeft      =   8475
   ClientTop       =   2055
   ClientWidth     =   2895
   Icon            =   "FrmCheckPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CMDExpand 
      Caption         =   "VVV"
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
      Left            =   1935
      TabIndex        =   7
      Top             =   135
      Width           =   780
   End
   Begin MSComctlLib.ListView ListViewPrintList 
      Height          =   4830
      Left            =   45
      TabIndex        =   6
      Top             =   1440
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   8520
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับ"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "เอกสาร"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Printer"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Timer TimerStart 
      Enabled         =   0   'False
      Left            =   315
      Top             =   3105
   End
   Begin VB.Timer TimerNow 
      Interval        =   65507
      Left            =   315
      Top             =   2655
   End
   Begin VB.PictureBox Crystal1011 
      Height          =   480
      Left            =   1350
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   3
      Top             =   4455
      Width           =   1200
   End
   Begin VB.PictureBox Crystal1031 
      Height          =   480
      Left            =   855
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   4
      Top             =   3060
      Width           =   1200
   End
   Begin VB.PictureBox Crystal1021 
      Height          =   480
      Left            =   225
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   5
      Top             =   3780
      Width           =   1200
   End
   Begin VB.Timer Timer3 
      Interval        =   3703
      Left            =   2025
      Top             =   2205
   End
   Begin VB.Timer Timer2 
      Interval        =   3303
      Left            =   1170
      Top             =   2205
   End
   Begin VB.Timer Timer1 
      Interval        =   3003
      Left            =   315
      Top             =   2205
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   690
      Left            =   720
      TabIndex        =   0
      Top             =   2025
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   1217
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "เลขที่"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "โซน"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ประเภท"
         Object.Width           =   2
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ครั้งที่"
         Object.Width           =   2
      EndProperty
   End
   Begin VB.Timer TimerCheckDocument 
      Interval        =   5000
      Left            =   270
      Top             =   1485
   End
   Begin Crystal.CrystalReport Crystal102 
      Left            =   1395
      Top             =   5445
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
   Begin Crystal.CrystalReport Crystal103 
      Left            =   540
      Top             =   5715
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
   Begin Crystal.CrystalReport Crystal101 
      Left            =   360
      Top             =   4995
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "รายการ พิมพ์"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   45
      TabIndex        =   8
      Top             =   1215
      Width           =   1365
   End
   Begin VB.Label LBLStatus 
      Alignment       =   2  'Center
      Caption         =   "กำลังทำงาน"
      Height          =   330
      Left            =   180
      TabIndex        =   2
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label LBLTime 
      Alignment       =   2  'Center
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
      Left            =   1170
      TabIndex        =   1
      Top             =   2700
      Width           =   1545
   End
   Begin VB.Image IM2 
      Height          =   480
      Left            =   1215
      Picture         =   "FrmCheckPrint.frx":1272
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IM3 
      Height          =   480
      Left            =   1215
      Picture         =   "FrmCheckPrint.frx":24E4
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IM1 
      Height          =   480
      Left            =   1215
      Picture         =   "FrmCheckPrint.frx":3756
      Top             =   90
      Width           =   480
   End
End
Attribute VB_Name = "FrmCheckPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vDocNo As String
Dim i As Integer

Private Sub CMDExpand_Click()

If Me.Height = 7200 Then
  Me.Height = 1665
Else
  Me.Height = 7200
End If

End Sub

Private Sub Form_Load()
Call InitializeDataBase
Me.ListViewPrintList.ListItems.Clear
End Sub

Private Sub Timer1_Timer()
Me.IM1.Visible = True
Me.IM2.Visible = False
Me.IM3.Visible = False
End Sub

Private Sub Timer2_Timer()
Me.IM1.Visible = False
Me.IM2.Visible = True
Me.IM3.Visible = False
End Sub

Private Sub Timer3_Timer()
Me.IM1.Visible = False
Me.IM2.Visible = False
Me.IM3.Visible = True
End Sub

Private Sub TimerCheckDocument_Timer()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vListItem As ListItem
Dim vListItemDoc As ListItem
Dim i As Integer
Dim vCount As Integer
Dim vDocType As Integer
Dim vTimeID  As Integer
Dim vZoneID As String
Dim vHourText As String
Dim vHourText1 As String
Dim vHour As Integer
Dim vSemi As Integer
Dim vLen As Integer
Dim vID As Integer
Dim vQueID As Integer
Dim vDocDate As String

On Error Resume Next

If Me.LBLTime.Caption <> "" Then
vHourText = Me.LBLTime.Caption
vLen = Len(vHourText)
vSemi = InStr(1, vHourText, ":")

If vLen = 3 Or vLen = 5 Then
vHourText1 = Left(vHourText, vLen - (vSemi))
ElseIf vLen = 4 And vSemi = 2 Then
vHourText1 = Left(vHourText, vLen - (vSemi + 1))
ElseIf vLen = 4 And vSemi = 3 Then
vHourText1 = Left(vHourText, vLen - (vSemi - 1))
End If

vHour = vHourText1

If vHour = 18 Then
Unload (FrmCheckPrint)
End If

End If

If ListView101.ListItems.Count <> 0 Then
  i = ListView101.ListItems.Count + 1
Else
  i = 1
End If
ListView101.ListItems.Clear
'vQuery = "exec dbo.USP_NP_SearchDocNopadolSystemAuto1"
vQuery = "exec dbo.USP_NP_SearchDocNopadolSystemAuto"
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
  While Not vRecordset.EOF
    If (vRecordset.Fields("doctype").Value = 2) Then
      Set vListItem = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("queid").Value))
      vListItem.SubItems(1) = Trim(vRecordset.Fields("zoneid").Value)
      vListItem.SubItems(2) = Trim(vRecordset.Fields("doctype").Value)
      vListItem.SubItems(3) = Trim(vRecordset.Fields("printtime").Value)
    Else
        Set vListItem = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("docno").Value))
        vListItem.SubItems(1) = Trim(vRecordset.Fields("zoneid").Value)
        vListItem.SubItems(2) = Trim(vRecordset.Fields("doctype").Value)
        vListItem.SubItems(3) = Trim(vRecordset.Fields("printtime").Value)
      End If
      
      i = i + 1
  vRecordset.MoveNext
  Wend
End If
vRecordset.Close

Dim m  As Double
Dim n As Double

vDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)

For vCount = 1 To ListView101.ListItems.Count
    vDocNo = Trim(ListView101.ListItems.Item(vCount).Text)
    vQueID = Trim(ListView101.ListItems.Item(vCount).Text)
    vZoneID = Trim(ListView101.ListItems.Item(vCount).SubItems(1))
    vDocType = Trim(ListView101.ListItems.Item(vCount).SubItems(2))
    vTimeID = Trim(ListView101.ListItems.Item(vCount).SubItems(3))
    
    
    If vDocNo <> "" And vDocType = 1 Then
        Call PrintRequestPickingSlip(vDocNo, vZoneID)
        vQuery = "exec dbo.USP_NP_UpdateQueuePrintStatus1 '" & vDocNo & "'," & vTimeID & " "
        vConnection.Execute vQuery
        
    ElseIf vDocNo <> "" And vDocType = 2 Then
        Call PrintSalePickingSlip(vQueID, vDocDate, vZoneID)
        vQuery = "exec dbo.USP_NP_UpdateQueuePrintStatus2 " & vQueID & "," & vTimeID & " "
        vConnection.Execute vQuery
    
    ElseIf vDocNo <> "" And vDocType = 3 And vZoneID <> "" Then
        'Call PrintDriveInDetails(vDocNo, vZoneID)
        Call PrintDriveInDetails(vDocNo, vZoneID)
        vQuery = "exec dbo.USP_NP_UpdateQueuePrintStatus1 '" & vDocNo & "'," & vTimeID & " "
        vConnection.Execute vQuery
        
        vID = Me.ListViewPrintList.ListItems.Count + 1
        
        Set vListItemDoc = Me.ListViewPrintList.ListItems.Add(, , vID)
        vListItemDoc.SubItems(1) = Trim(vDocNo)
        vListItemDoc.SubItems(2) = Trim(vMemPrinter)
        
    ElseIf vDocNo <> "" And vDocType = 6 Then
        'Call PrintReserveOrder(vDocNo)
        'vQuery = "exec dbo.USP_NP_UpdateQueuePrintStatus1 '" & vDocNo & "'," & vTimeID & " "
        'vConnection.Execute vQuery
    ElseIf vDocNo <> "" And vDocType = 7 Then
        'Call PrintSaleOrder(vDocNo)
        'vQuery = "exec dbo.USP_NP_UpdateQueuePrintStatus1 '" & vDocNo & "'," & vTimeID & " "
        'vConnection.Execute vQuery
    ElseIf vDocNo <> "" And vDocType = 8 Then
        'Call PrintQueueDelivery(vDocNo)
        'vQuery = "exec dbo.USP_NP_UpdateQueuePrintStatus1 '" & vDocNo & "'," & vTimeID & " "
        'vConnection.Execute vQuery
    ElseIf vDocNo <> "" And vDocType = 9 Then
        'Call PrintCheckOutHeader(vDocNo)
        'Call PrintCheckOutItem(vDocNo)
        'Call PrintHoldingBill(vDocNo)
        'Call PrintHoldingBillSub(vDocNo)
        vQuery = "exec dbo.USP_NP_UpdateQueuePrintStatus1 '" & vDocNo & "'," & vTimeID & " "
        vConnection.Execute vQuery
    ElseIf vDocNo <> "" And vDocType = 10 Then
        'Call PrintInspection(vDocNo)
        'vQuery = "exec dbo.USP_NP_UpdateQueuePrintStatus1 '" & vDocNo & "'," & vTimeID & " "
        'vConnection.Execute vQuery
    End If
Next vCount
      
ListView101.ListItems.Clear
End Sub

Public Sub PrintSalePickingSlip(vQueID As Integer, vQueDocDate As String, vZone As String)
Dim vPrinterName As String
Dim printerObj As Printer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim i As Integer
Dim prnPrinter As Printer
Dim n As Integer
Dim vQueNo As String
   

If vZone = "A" Or vZone = "X" Then
 vQuery = "exec dbo.USP_NP_SearchPrinter 2"
 If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
     vPrinterName = Trim(vRecordset.Fields("pathprinter").Value)
 End If
 vRecordset.Close
    
If vPrinterName <> "" Then
   For Each printerObj In Printers
   If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
   Set Printer = printerObj
   Set printerObj = Nothing
   
   'MsgBox vPrinterName
   
   Exit For
   End If
   Next
Else
   Exit Sub
End If
End If

If vZone = "B" Then
 vQuery = "exec dbo.USP_NP_SearchPrinter 3"
 If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
     vPrinterName = Trim(vRecordset.Fields("pathprinter").Value)
 End If
 vRecordset.Close
    
   If vPrinterName <> "" Then
      For Each printerObj In Printers
      If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
      Set Printer = printerObj
      Set printerObj = Nothing
      Exit For
      End If
      Next
   Else
      Exit Sub
   End If
End If


If vZone = "C" Then
 vQuery = "exec dbo.USP_NP_SearchPrinter 4"
 If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
     vPrinterName = Trim(vRecordset.Fields("pathprinter").Value)
 End If
 vRecordset.Close
    
   If vPrinterName <> "" Then
      For Each printerObj In Printers
      If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
      Set Printer = printerObj
      Set printerObj = Nothing
      Exit For
      End If
      Next
   Else
      Exit Sub
   End If
End If

        
vQuery = "exec dbo.USP_NP_SearchQueCenterDetails " & vQueID & ",'" & vQueDocDate & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then

vQueNo = Trim(vRecordset.Fields("queid").Value)

Printer.FontName = "AngsanaUPC"
Printer.Font.Size = 50
Printer.CurrentX = 1700
Printer.Print Trim(vRecordset.Fields("queid").Value)

'Printer.Font.Name = "Code128"
Printer.Font.Name = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.CurrentX = 1500
'Printer.Print "*" & Trim(vRecordset.Fields("queid").Value) & "*"
Printer.Print "*" & vQueNo & "*"

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 1000
Printer.FontBold = True
Printer.Print Trim("Picking Request Slip Details")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("วันที่คิว: ") & Trim(vRecordset.Fields("quedocdate").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("เลขที่เอกสาร : ") & Trim(vRecordset.Fields("docno").Value) & "          " & Trim("วันที่เอกสาร : ") & Trim(vRecordset.Fields("docdate").Value)

'MsgBox (vRecordset.Fields("docno").Value)

'Printer.Font.Name = "Code128"
Printer.Font.Name = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "*" & Trim(vRecordset.Fields("docno").Value) & "*"

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("รหัส/ชื่อลูกค้า : ") & Trim(vRecordset.Fields("arcode").Value) & "/" & Trim(vRecordset.Fields("arname").Value)

Printer.Font.Name = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "*" & Trim(vRecordset.Fields("arcode").Value) & "*"

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("ทะเบียนรถ : ") & Trim(vRecordset.Fields("refno").Value)

Printer.Font.Name = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "*" & Trim(vRecordset.Fields("refno").Value) & "*"


Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("พนักงานขาย : ") & Trim(vRecordset.Fields("salecode").Value) & "/" & Trim(vRecordset.Fields("salename").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("คลัง : ") & Trim(vRecordset.Fields("whcode").Value) & "/" & Trim(vRecordset.Fields("shelfcode").Value) & "     " & "โซนการจัด :" & Trim(vRecordset.Fields("quezone").Value)
vRecordset.MoveFirst

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("-----------------------------------------------------------------------------------------------")
vRecordset.MoveFirst
n = 1
While Not vRecordset.EOF

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print "ที่เก็บ :" & Trim(vRecordset.Fields("shelfid").Value) & "                 " & "  ยอดพอขายตามคลัง :  " & Trim(vRecordset.Fields("remainsale").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print n & ". " & "รหัสสินค้า :" & Trim(vRecordset.Fields("itemcode").Value) & "              " & " ขายชั้นเก็บ :" & Trim(vRecordset.Fields("shelfcode").Value)

Printer.FontName = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.FontBold = False
Printer.Print "*" & Trim(vRecordset.Fields("itemcode").Value) & "*"

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print "ชื่อสินค้า :" & Trim(vRecordset.Fields("itemname").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print "ต้องการ :" & Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00") & "    " & Trim(vRecordset.Fields("unitcode").Value) & "                         " & Trim("จัดได้ : ______________")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("-----------------------------------------------------------------------------------------------")
    
vRecordset.MoveNext
n = n + 1
Wend
End If
vRecordset.Close
    
Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("_______________________________________________________________________________________________")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "               ผู้จัดสินค้า                                             ผู้รับสินค้า"

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "         _____________                                    ______________"
      
Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("______________________________________________________________________________________________")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 10
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print Trim("วันที่พิมพ์ :") & Now

Printer.EndDoc
End Sub

Public Sub PrintInspection(vDocNo As String)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vRepID As Integer
Dim vRepType As String
Dim vReportName As String

vRepID = 213
vRepType = "IV"
vQuery = "select reportname from bcnp.dbo.bcreportname where  repid = " & vRepID & " and reptype = '" & vRepType & "'  "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal103
.ReportFileName = Trim(vReportName & ".rpt")
.ParameterFields(0) = "@Docno;" & vDocNo & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With

End Sub

Public Sub PrintDriveInDetails(vDocNo As String, vZoneID As String)
Dim vPrinterName As String
Dim printerObj As Printer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim i As Integer
Dim prnPrinter As Printer
Dim n As Integer
Dim vDIPoint As Integer

On Error Resume Next

'vDIPoint = vZoneID

'If vDIPoint = "1" Then
'For Each prnPrinter In Printers
 '  If prnPrinter.DeviceName = "\\hptc-5100418\SRP370A" Then
  '    Set Printer = prnPrinter
   '   Exit For
   'End If
'Next
'End If

'If vDIPoint = "2" Then
'For Each prnPrinter In Printers
 '  If prnPrinter.DeviceName = "\\hptc-5100421\SRP370B" Then
  '    Set Printer = prnPrinter
   '   Exit For
   'End If
'Next
'End If

'If vDIPoint = "3" Then
'For Each prnPrinter In Printers
 '  If prnPrinter.DeviceName = "\\Hptc-5100420\SRP370C" Then
  '    Set Printer = prnPrinter
   '   Exit For
   'End If
'Next
'End If

    vQuery = "exec dbo.USP_NP_SearchPrinter 1"
    If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
        vPrinterName = Trim(vRecordset.Fields("pathprinter").Value)
    End If
    vRecordset.Close

For Each prnPrinter In Printers
   'If prnPrinter.DeviceName = "\\hptc-checkout\SRP370-CheckOut" Then
   If prnPrinter.DeviceName = vPrinterName Then ' "\\s2dt1t\BIXOLON SRP-370" Then
      Set Printer = prnPrinter
      vMemPrinter = vPrinterName '"\\s2dt1t\BIXOLON SRP-370"
      Exit For
   Else
      vMemPrinter = "Default"
   End If
Next
        
vQuery = "exec dbo.USP_NP_SearchDriveInPickZoneDetails1 '" & vDocNo & "','" & vZoneID & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 1450
Printer.FontBold = True
Printer.Print Trim("DriveIn Slip Details")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("เลขที่เอกสาร : ") & Trim(vRecordset.Fields("docno").Value) & "          " & Trim("วันที่เอกสาร : ") & Trim(vRecordset.Fields("docdate").Value)

Printer.Font.Name = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "*" & Trim(vRecordset.Fields("docno").Value) & "*"

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("รหัส/ชื่อลูกค้า : ") & Trim(vRecordset.Fields("arcode").Value) & "/" & Trim(vRecordset.Fields("arname").Value)

Printer.Font.Name = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "*" & Trim(vRecordset.Fields("arcode").Value) & "*"

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("ทะเบียนรถ : ") & Trim(vRecordset.Fields("refno").Value)

Printer.Font.Name = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "*" & Trim(vRecordset.Fields("refno").Value) & "*"

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("พนักงานขาย : ") & Trim(vRecordset.Fields("salecode").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("คลัง : ") & Trim(vRecordset.Fields("whcode").Value) & "/" & Trim(vRecordset.Fields("shelfcode").Value) & "     " & "โซนการจัด :" & Trim(vRecordset.Fields("zoneid").Value)


Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("_______________________________________________________________________________________________")
vRecordset.MoveFirst
n = 1
While Not vRecordset.EOF

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print "ที่เก็บ :" & Trim(vRecordset.Fields("shelfid").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print n & ". " & "รหัสสินค้า :" & Trim(vRecordset.Fields("itemcode").Value) & "              " & " ขายชั้นเก็บ :" & Trim(vRecordset.Fields("shelfcode").Value)

Printer.FontName = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.FontBold = False
Printer.Print "*" & Trim(vRecordset.Fields("itemcode").Value) & "*"

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print "ชื่อสินค้า :" & Trim(vRecordset.Fields("itemname").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print "จ่าย:" & Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00") & "    " & Trim(vRecordset.Fields("unitcode").Value)

If i = vRecordset.RecordCount Then

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("_______________________________________________________________________________________________")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 20
Printer.CurrentX = 1000
Printer.FontBold = True
Printer.Print "มูลค่าสุทธิ :" & "       " & Format(Trim(vRecordset.Fields("totalnetamount").Value), "##,##0.00")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("_______________________________________________________________________________________________")

End If

vRecordset.MoveNext
n = n + 1
Wend

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "               ผู้จ่ายสินค้า                                             ผู้รับสินค้า"

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "         _____________                                    ______________"

     
Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("______________________________________________________________________________________________")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 10
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print Trim("วันที่พิมพ์ :") & Now
End If
vRecordset.Close

Printer.EndDoc
End Sub

Public Sub PrintDriveInDetails1(vDocNo As String, vZoneID As String)
            With Crystal101
                .ReportFileName = "V:\Reports\RP_NP_DriveInDetails.rpt"
                .ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
                .ParameterFields(1) = "@vPickZone;" & vZoneID & ";true"
                .Destination = crptToWindow
                .WindowState = crptMaximized
                .Action = 1
            End With
        
End Sub

Public Sub PrintHoldingBill(vDocNo As String)
Dim vPrinterName As String
Dim printerObj As Printer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim i As Integer
Dim prnPrinter As Printer
Dim n As Integer

On Error Resume Next
   
For Each prnPrinter In Printers
   If prnPrinter.DeviceName = "\\hptc-checkout\CheckOut" Then
      Set Printer = prnPrinter
      Exit For
   End If
Next
        
vQuery = "exec dbo.usp_np_SearchHoldingDetails1'" & vDocNo & "'  "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 20
Printer.CurrentX = 950
Printer.FontBold = True
Printer.Print Trim("CheckOut Master")

Printer.Font.Name = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.CurrentX = 0
Printer.Print "*" & Trim(vRecordset.Fields("docno").Value) & "*"

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 1550
Printer.FontBold = True
Printer.Print Trim(vRecordset.Fields("docno").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("เลขที่เอกสาร : ") & Trim(vRecordset.Fields("docno").Value) & "          " & Trim("วันที่เอกสาร : ") & Trim(vRecordset.Fields("docdate").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("รหัส/ชื่อลูกค้า : ") & Trim(vRecordset.Fields("arcode").Value) & "/" & Trim(vRecordset.Fields("arname").Value)

Printer.Font.Name = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "*" & Trim(vRecordset.Fields("arcode").Value) & "*"

If Trim(vRecordset.Fields("license").Value) <> "" Then
Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("ทะเบียนรถ : ") & Trim(vRecordset.Fields("license").Value)

Printer.Font.Name = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "*" & Trim(vRecordset.Fields("license").Value) & "*"
End If

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("จุดแคชเชียร์: ") & Trim(vRecordset.Fields("machineno").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("พนักงานขาย : ") & Trim(vRecordset.Fields("salecode").Value) & "/" & Trim(vRecordset.Fields("salename").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("คลัง : ") & Trim(vRecordset.Fields("whcode").Value) & "/" & Trim(vRecordset.Fields("shelfcode").Value)
    
Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("_______________________________________________________________________________________________")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 16
Printer.CurrentX = 1400
Printer.FontBold = True
Printer.Print "มูลค่าสินค้า :" & "     " & Format(Trim(vRecordset.Fields("sumofitemamount").Value), "##,##0.00")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 16
Printer.CurrentX = 1400
Printer.FontBold = True
Printer.Print "มูลค่าส่วนลด :" & "  " & Format(Trim(vRecordset.Fields("discountamount").Value), "##,##0.00")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 16
Printer.CurrentX = 1400
Printer.FontBold = True
Printer.Print "มูลค่าภาษี :" & "       " & Format(Trim(vRecordset.Fields("taxamount").Value), "##,##0.00")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 16
Printer.CurrentX = 1400
Printer.FontBold = True
Printer.Print "มูลค่าสุทธิ :" & "      " & Format(Trim(vRecordset.Fields("totalamount").Value), "##,##0.00")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("_______________________________________________________________________________________________")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 10
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print Trim("วันที่พิมพ์ :") & Now
End If
vRecordset.Close

Printer.EndDoc
End Sub


Public Sub PrintHoldingBillSub(vDocNo As String)
Dim vPrinterName As String
Dim printerObj As Printer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim i As Integer
Dim prnPrinter As Printer
Dim n As Integer
   
On Error Resume Next

For Each prnPrinter In Printers
   If prnPrinter.DeviceName = "\\hptc-checkout\SRP370-CheckOut" Then
      Set Printer = prnPrinter
      Exit For
   End If
Next
        
vQuery = "exec dbo.usp_np_SearchHoldingDetails1 '" & vDocNo & "'  "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 20
Printer.CurrentX = 950
Printer.FontBold = True
Printer.Print Trim("CheckOut Details")

Printer.Font.Name = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.CurrentX = 0
Printer.Print "*" & Trim(vRecordset.Fields("docno").Value) & "*"

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 1550
Printer.FontBold = True
Printer.Print Trim(vRecordset.Fields("docno").Value)


Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("เลขที่เอกสาร : ") & Trim(vRecordset.Fields("docno").Value) & "          " & Trim("วันที่เอกสาร : ") & Trim(vRecordset.Fields("docdate").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("รหัส/ชื่อลูกค้า : ") & Trim(vRecordset.Fields("arcode").Value) & "/" & Trim(vRecordset.Fields("arname").Value)

Printer.Font.Name = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "*" & Trim(vRecordset.Fields("arcode").Value) & "*"

If Trim(vRecordset.Fields("license").Value) <> "" Then
Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("ทะเบียนรถ : ") & Trim(vRecordset.Fields("license").Value)

Printer.Font.Name = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "*" & Trim(vRecordset.Fields("license").Value) & "*"
End If

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
'Printer.Print Trim("จุดแคชเชียร์: ") & Trim(vRecordset.Fields("machineno").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("พนักงานขาย : ") & Trim(vRecordset.Fields("salecode").Value) & "/" & Trim(vRecordset.Fields("salename").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("คลัง : ") & Trim(vRecordset.Fields("whcode").Value) & "/" & Trim(vRecordset.Fields("shelfcode").Value)
vRecordset.MoveFirst

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("-----------------------------------------------------------------------------------------------")
vRecordset.MoveFirst
n = 1
For i = 1 To vRecordset.RecordCount

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print n & ". " & "รหัสสินค้า :" & Trim(vRecordset.Fields("itemcode").Value) & "              " & " ขายชั้นเก็บ :" & Trim(vRecordset.Fields("shelfcode").Value)

Printer.FontName = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.FontBold = False
Printer.Print "*" & Trim(vRecordset.Fields("itemcode").Value) & "*"

If Len(Trim(vRecordset.Fields("itemname").Value)) <= 40 Then
Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print "ชื่อสินค้า :" & Trim(vRecordset.Fields("itemname").Value)

ElseIf Len(Trim(vRecordset.Fields("itemname").Value)) > 40 And Len(Trim(vRecordset.Fields("itemname").Value)) <= 80 Then
Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print "ชื่อสินค้า :" & Left(Trim(vRecordset.Fields("itemname").Value), 40)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print Mid(Trim(vRecordset.Fields("itemname").Value), 41, 40)

ElseIf Len(Trim(vRecordset.Fields("itemname").Value)) > 80 Then
Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print "ชื่อสินค้า :" & Left(Trim(vRecordset.Fields("itemname").Value), 40)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print Mid(Trim(vRecordset.Fields("itemname").Value), 41, 40)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print Mid(Trim(vRecordset.Fields("itemname").Value), 81, 40)

End If

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print "จำนวน :" & Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00") & "    " & Trim(vRecordset.Fields("unitcode").Value) & "                         " & Trim("ราคา : ") & Format(Trim(vRecordset.Fields("price").Value), "##,##0.00") & "    บาท"

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("-----------------------------------------------------------------------------------------------")
    
If i = vRecordset.RecordCount Then

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 16
Printer.CurrentX = 1400
Printer.FontBold = True
Printer.Print "มูลค่าสินค้า :" & "     " & Format(Trim(vRecordset.Fields("sumofitemamount").Value), "##,##0.00")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 16
Printer.CurrentX = 1400
Printer.FontBold = True
Printer.Print "มูลค่าส่วนลด :" & "  " & Format(Trim(vRecordset.Fields("discountamount").Value), "##,##0.00")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 16
Printer.CurrentX = 1400
Printer.FontBold = True
Printer.Print "มูลค่าภาษี :" & "       " & Format(Trim(vRecordset.Fields("taxamount").Value), "##,##0.00")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 16
Printer.CurrentX = 1400
Printer.FontBold = True
Printer.Print "มูลค่าสุทธิ :" & "      " & Format(Trim(vRecordset.Fields("totalamount").Value), "##,##0.00")

End If
vRecordset.MoveNext
n = n + 1
Next i
End If
vRecordset.Close

    
Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("_______________________________________________________________________________________________")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "               ผู้ตรวจสินค้า                                             ผู้รับสินค้า"

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "         _____________                                    ______________"
      
Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("______________________________________________________________________________________________")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 10
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print Trim("วันที่พิมพ์ :") & Now

Printer.EndDoc
End Sub

Public Sub PrintCheckOutItem(vDocNo As String)
Dim vPrinterName As String
Dim printerObj As Printer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim i As Integer
Dim prnPrinter As Printer
Dim n As Integer

On Error Resume Next
   
For Each prnPrinter In Printers
   'If prnPrinter.DeviceName = "\\hptc-5100418\SRP370CheckOut" Then
    If prnPrinter.DeviceName = "\\hptc-checkout\SRP370-CheckOut" Then
      Set Printer = prnPrinter
      Exit For
   End If
Next
        
vQuery = "exec dbo.USP_NP_SearchDriveInMergeTemp1 '" & vDocNo & "'  "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 20
Printer.CurrentX = 950
Printer.FontBold = True
Printer.Print Trim("CheckOut Details")

Printer.Font.Name = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.CurrentX = 0
Printer.Print "*" & Trim(vRecordset.Fields("docno").Value) & "*"

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 1550
Printer.FontBold = True
Printer.Print Trim(vRecordset.Fields("docno").Value)


Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("เลขที่เอกสาร : ") & Trim(vRecordset.Fields("docno").Value) & "          " & Trim("วันที่เอกสาร : ") & Trim(vRecordset.Fields("docdate").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("รหัส/ชื่อลูกค้า : ") & Trim(vRecordset.Fields("arcode").Value) & "/" & Trim(vRecordset.Fields("arname").Value)

Printer.Font.Name = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "*" & Trim(vRecordset.Fields("arcode").Value) & "*"

If Trim(vRecordset.Fields("carlicense").Value) <> "" Then
Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("ทะเบียนรถ : ") & Trim(vRecordset.Fields("carlicense").Value)

Printer.Font.Name = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "*" & Trim(vRecordset.Fields("carlicense").Value) & "*"
End If

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("พนักงานขาย : ") & Trim(vRecordset.Fields("salecode").Value) & "/" & Trim(vRecordset.Fields("salename").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("คลัง : ") & Trim(vRecordset.Fields("whcode").Value) & "/" & Trim(vRecordset.Fields("shelfcode").Value)
vRecordset.MoveFirst

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("-----------------------------------------------------------------------------------------------")
vRecordset.MoveFirst
n = 1
For i = 1 To vRecordset.RecordCount

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print n & ". " & "รหัสสินค้า :" & Trim(vRecordset.Fields("barcode").Value) & "              " & " ขายชั้นเก็บ :" & Trim(vRecordset.Fields("shelfcode").Value)

Printer.FontName = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.FontBold = False
Printer.Print "*" & Trim(vRecordset.Fields("barcode").Value) & "*"

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print "ชื่อสินค้า :" & Trim(vRecordset.Fields("itemname").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print "จำนวน :" & Format(Trim(vRecordset.Fields("invqty").Value), "##,##0.00") & "    " & Trim(vRecordset.Fields("unitcode").Value) & "                         " & Trim("นับได้ : ") & Format(Trim(vRecordset.Fields("invqty").Value), "##,##0.00") & "    " & Trim(vRecordset.Fields("unitcode").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("-----------------------------------------------------------------------------------------------")
vRecordset.MoveNext
n = n + 1
Next i
End If
vRecordset.Close

    
Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("_______________________________________________________________________________________________")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "               ผู้ตรวจสินค้า                                             ผู้รับสินค้า"

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "         _____________                                    ______________"
      
Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("______________________________________________________________________________________________")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 10
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print Trim("วันที่พิมพ์ :") & Now

Printer.EndDoc
End Sub

Public Sub PrintQueueDelivery(vDocNo As String)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String


vRepID = 466
vRepType = "DO"


vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal102
.ReportFileName = vReportName & ".rpt"
.ParameterFields(0) = "@QRDocno;" & vDocNo & ";true"
.Destination = crptToPrinter
.WindowState = crptMaximized
.Action = 1
End With

End Sub

Public Sub PrintSaleOrder(vDocNo As String)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String
Dim vComputerName As String
Dim vUserName As String

vComputerName = "สาขา"
vUserName = "พิมพ์อัตโนมัติ"

vRepID = 467
vRepType = "SO"

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close


With Crystal102
.ReportFileName = vReportName & ".rpt"
.ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
.Destination = crptToPrinter
.WindowState = crptMaximized
.Formulas(0) = "computername='" & vComputerName & "' "
.Formulas(1) = "username='" & vUserName & "' "
.Action = 1
End With
End Sub


Public Sub PrintReserveOrder(vDocNo As String)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String
Dim vComputerName As String
Dim vUserName As String

vComputerName = "สาขา"
vUserName = "พิมพ์อัตโนมัติ"

vRepID = 468
vRepType = "RO"

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal102
.ReportFileName = vReportName & ".rpt"
.ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
.Destination = crptToPrinter
.WindowState = crptMaximized
.Formulas(0) = "computername='" & vComputerName & "' "
.Formulas(1) = "username='" & vUserName & "' "
.Action = 1
End With
End Sub

Public Sub PrintRequestPickingSlip(vDocNo As String, vZone As String)
Dim vPrinterName As String
Dim printerObj As Printer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim i As Integer
Dim prnPrinter As Printer
Dim n As Integer
Dim vDocDate As String
   
On Error Resume Next

vDocDate = Day(Now) & "/" & Month(Now) & "/" & Year(Now)

If vZone = "01" Then
 vQuery = "exec dbo.USP_NP_SearchPrinter 3"
 If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
     vPrinterName = Trim(vRecordset.Fields("pathprinter").Value)
 End If
 vRecordset.Close
    
If vPrinterName <> "" Then
   For Each printerObj In Printers
   If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
   Set Printer = printerObj
   Set printerObj = Nothing
   Exit For
   End If
   Next
Else
   Exit Sub
End If
End If

If vZone = "02" Then
 vQuery = "exec dbo.USP_NP_SearchPrinter 3"
 If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
     vPrinterName = Trim(vRecordset.Fields("pathprinter").Value)
 End If
 vRecordset.Close
    
   If vPrinterName <> "" Then
      For Each printerObj In Printers
      If UCase(printerObj.DeviceName) = UCase(vPrinterName) Then
      Set Printer = printerObj
      Set printerObj = Nothing
      Exit For
      End If
      Next
   Else
      Exit Sub
   End If
End If

        
vQuery = "exec dbo.USP_NP_SearchPickingRequestDetails '" & vDocNo & "','" & vDocDate & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then

Printer.FontName = "AngsanaUPC"
Printer.Font.Size = 20
Printer.CurrentX = 1700
Printer.Print Trim(vRecordset.Fields("docno").Value)

Printer.Font.Name = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.CurrentX = 1500
Printer.Print "*" & vDocNo & "*"

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 1000
Printer.FontBold = True
Printer.Print Trim("Picking Request Slip Details")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("วันที่ : ") & Trim(vRecordset.Fields("docdate").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("เลขที่เอกสาร : ") & Trim(vRecordset.Fields("docno").Value) & "          " & Trim("วันที่เอกสาร : ") & Trim(vRecordset.Fields("docdate").Value)

Printer.Font.Name = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "*" & Trim(vRecordset.Fields("docno").Value) & "*"

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("รหัส/ชื่อลูกค้า : ") & Trim(vRecordset.Fields("arcode").Value) & "/" & Trim(vRecordset.Fields("arname").Value)

Printer.Font.Name = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "*" & Trim(vRecordset.Fields("arcode").Value) & "*"

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("ทะเบียนรถ : ") & Trim(vRecordset.Fields("refno").Value)

Printer.Font.Name = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "*" & Trim(vRecordset.Fields("refno").Value) & "*"


Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("พนักงานขาย : ") & Trim(vRecordset.Fields("salecode").Value) & "/" & Trim(vRecordset.Fields("salename").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("คลัง : ") & Trim(vRecordset.Fields("whcode").Value) & "/" & Trim(vRecordset.Fields("shelfcode").Value) & "     " & "โซนการจัด :" & Trim(vRecordset.Fields("pointid").Value)
vRecordset.MoveFirst

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("-----------------------------------------------------------------------------------------------")
vRecordset.MoveFirst
n = 1
While Not vRecordset.EOF

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print "ที่เก็บ :" & Trim(vRecordset.Fields("shelfid").Value) & "                 " & "  ยอดพอขายตามคลัง :  " & Trim(vRecordset.Fields("remainsale").Value)

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print n & ". " & "รหัสสินค้า :" & Trim(vRecordset.Fields("itemcode").Value) & "              " & " ขายชั้นเก็บ :" & Trim(vRecordset.Fields("shelfcode").Value)

Printer.FontName = "3 of 9 Barcode"
Printer.Font.Size = 20
Printer.FontBold = False
Printer.Print "*" & Trim(vRecordset.Fields("itemcode").Value) & "*"

If Len(Trim(vRecordset.Fields("itemname").Value)) <= 40 Then
Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print "ชื่อสินค้า :" & Trim(vRecordset.Fields("itemname").Value)

ElseIf Len(Trim(vRecordset.Fields("itemname").Value)) > 40 And Len(Trim(vRecordset.Fields("itemname").Value)) <= 80 Then
Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print "ชื่อสินค้า :" & Left(Trim(vRecordset.Fields("itemname").Value), 40) 'บรรทัดที่ 1

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print Mid(Trim(vRecordset.Fields("itemname").Value), 41, 40) 'บรรทัดที่ 2

ElseIf Len(Trim(vRecordset.Fields("itemname").Value)) > 80 Then
Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print "ชื่อสินค้า :" & Left(Trim(vRecordset.Fields("itemname").Value), 40) 'บรรทัดที่ 1

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print Mid(Trim(vRecordset.Fields("itemname").Value), 41, 40) 'บรรทัดที่ 2

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print Mid(Trim(vRecordset.Fields("itemname").Value), 81, 40) 'บรรทัดที่ 3
End If


Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print "ต้องการ :" & Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00") & "    " & Trim(vRecordset.Fields("unitcode").Value) & "                         " & Trim("จัดได้ : ______________")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("-----------------------------------------------------------------------------------------------")
    
vRecordset.MoveNext
n = n + 1
Wend
End If
vRecordset.Close
    
Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("_______________________________________________________________________________________________")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "               ผู้จัดสินค้า                                             ผู้รับสินค้า"

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print "         _____________                                    ______________"
      
Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.FontBold = True
Printer.Print Trim("______________________________________________________________________________________________")

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 10
Printer.CurrentX = 0
Printer.FontBold = False
Printer.Print Trim("วันที่พิมพ์ :") & Now

Printer.EndDoc
End Sub

Private Sub TimerNow_Timer()
Me.LBLTime.Caption = Hour(Now) & ":" & Minute(Now)
End Sub

Private Sub TimerStart_Timer()
Dim vHourText As String
Dim vHourText1 As String
Dim vHour As Integer
Dim vSemi As Integer
Dim vLen As Integer

On Error Resume Next

If Me.LBLTime.Caption <> "" Then
vHourText = Me.LBLTime.Caption
vLen = Len(vHourText)
vSemi = InStr(1, vHourText, ":")

If vLen = 3 Or vLen = 5 Then
vHourText1 = Left(vHourText, vLen - (vSemi))
ElseIf vLen = 4 And vSemi = 2 Then
vHourText1 = Left(vHourText, vLen - (vSemi + 1))
ElseIf vLen = 4 And vSemi = 3 Then
vHourText1 = Left(vHourText, vLen - (vSemi - 1))
End If
'vHourText1 = Right(vHourText, vLen - vSemi)

vHour = vHourText1

If vHour >= 8 And vHour <= 18 Then
Me.TimerCheckDocument.Enabled = True
Me.LBLStatus.Caption = "กำลังทำงาน"
Me.TimerStart.Enabled = False
Exit Sub
End If

End If
End Sub
