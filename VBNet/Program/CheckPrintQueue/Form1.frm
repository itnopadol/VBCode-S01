VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCheckPrintQueue 
   Caption         =   "พิมพ์เอกสารอัตโนมัติ"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   5430
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20426
            Text            =   "กำลังค้นหา เอกสารที่ยังไม่ได้พิมพ์"
            TextSave        =   "กำลังค้นหา เอกสารที่ยังไม่ได้พิมพ์"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   4200
      Left            =   135
      TabIndex        =   0
      Top             =   1215
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   7408
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับที่"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "เลขที่เอกสาร"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ฟอร์มที่พิมพ์"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "คนที่สั่งพิมพ์"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "โซนสินค้า"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ประเภทเอกสาร"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ครั้งที่พิมพ์"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   510
      Left            =   5400
      TabIndex        =   3
      Top             =   4635
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Timer TimerCheck 
      Interval        =   1
      Left            =   720
      Top             =   4950
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   225
      Top             =   4950
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Timer TimerCheckDocument 
      Interval        =   20000
      Left            =   1215
      Top             =   4950
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   135
      Picture         =   "Form1.frx":20082
      Top             =   45
      Width           =   2160
   End
   Begin VB.Label Label1 
      Caption         =   "รายการ พิมพ์เอกสาร"
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
      Left            =   135
      TabIndex        =   2
      Top             =   900
      Width           =   1770
   End
   Begin VB.Label LBLTime 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10485
      TabIndex        =   1
      Top             =   855
      Width           =   1230
   End
End
Attribute VB_Name = "FrmCheckPrintQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vDocNo As String
Dim vReportName As String
Dim vCrystalReport As CrystalReport
Dim i As Integer

 
Private Sub Form_Load()
On Error Resume Next

Call InitializeDatabase

    Me.Show
    Me.Refresh
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Your ToolTip" & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nid
    Me.Hide

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

'this procedure receives the callbacks from the System Tray icon.
Dim Result As Long
Dim msg As Long
    'the value of X will vary depending upon the scalemode setting
    If Me.ScaleMode = vbPixels Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If
    Select Case msg
        Case WM_LBUTTONUP        '514 restore form window
            Me.WindowState = vbNormal
            Result = SetForegroundWindow(Me.hwnd)
            Me.Show
        Case WM_LBUTTONDBLCLK    '515 restore form window
            Me.WindowState = vbNormal
            Result = SetForegroundWindow(Me.hwnd)
            Me.Show
        Case WM_RBUTTONUP        '517 display popup menu
            Result = SetForegroundWindow(Me.hwnd)
    End Select
End Sub

Private Sub Form_Resize()
On Error Resume Next

    'this is necessary to assure that the minimized window is hidden
    If Me.WindowState = vbMinimized Then Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    'this removes the icon from the system tray
    Shell_NotifyIcon NIM_DELETE, nid
End Sub
   
Private Sub mPopExit_Click()
On Error Resume Next

    'called when user clicks the popup menu Exit command
    Unload Me
End Sub
   
Private Sub mPopRestore_Click()
Dim Result As Long

On Error Resume Next

    'called when the user clicks the popup menu Restore command
    Me.WindowState = vbNormal
    Result = SetForegroundWindow(Me.hwnd)
    Me.Show
End Sub
'Private Sub Timer2_Timer()
'LBLTime = Trim(Hour(Now) & ":" & Minute(Now) & ":" & Second(Now))
'End Sub

Private Sub TimerCheck_Timer()
On Error Resume Next
LBLTime = Trim(Hour(Now) & ":" & Minute(Now) & ":" & Second(Now))
End Sub

Private Sub TimerCheckDocument_Timer()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vListItem As ListItem
Dim i As Integer
Dim vCount As Integer
Dim vDocType As Integer
Dim vTimeID  As Integer
Dim vZoneID As String
Dim n As Integer


'On Error Resume Next


If ListView101.ListItems.Count <> 0 Then
  i = ListView101.ListItems.Count + 1
Else
  i = 1
End If
ListView101.ListItems.Clear
vQuery = "exec dbo.USP_NP_SearchDocNopadolSystemAuto"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
  While Not vRecordset.EOF
      Set vListItem = ListView101.ListItems.Add(, , i)
      vListItem.SubItems(1) = Trim(vRecordset.Fields("docno").Value)
      vListItem.SubItems(2) = Trim(vRecordset.Fields("form").Value)
      vListItem.SubItems(3) = Trim(vRecordset.Fields("userprint").Value)
      vListItem.SubItems(4) = Trim(vRecordset.Fields("zoneid").Value)
      vListItem.SubItems(5) = Trim(vRecordset.Fields("doctype").Value)
      vListItem.SubItems(6) = Trim(vRecordset.Fields("printtime").Value)
      i = i + 1
  vRecordset.MoveNext
  Wend
End If
vRecordset.Close

For vCount = 1 To ListView101.ListItems.Count
    vDocNo = Trim(ListView101.ListItems.Item(vCount).SubItems(1))
    vZoneID = Trim(ListView101.ListItems.Item(vCount).SubItems(4))
    vDocType = Trim(ListView101.ListItems.Item(vCount).SubItems(5))
    vTimeID = Trim(ListView101.ListItems.Item(vCount).SubItems(6))
    
    If vDocNo <> "" And vDocType = 3 And vZoneID <> "" Then
        Call PrintDriveInDetails(vDocNo, vZoneID)
        vQuery = "exec dbo.USP_NP_UpdateQueuePrintStatus1 '" & vDocNo & "'," & vTimeID & " "
        gConnection.Execute vQuery
        'ListView101.ListItems.Clear
    End If
Next vCount

End Sub
Public Sub PrintDriveInDetails1(vDocNo As String, vZoneID As String)
'            With Crystal101
 '               .ReportFileName = "V:\Reports\RP_NP_DriveInDetails.rpt"
  '              .ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
   '             .ParameterFields(1) = "@vPickZone;" & vZoneID & ";true"
    '            .Destination = crptToWindow
     '           .WindowState = crptMaximized
      '          .Action = 1
       '     End With
        
End Sub

Public Sub PrintDriveInHeader(vDocNo As String, vZoneID As String)
Dim vPrinterName As String
Dim printerObj As Printer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim i As Integer
Dim prnPrinter As Printer
Dim n As Integer
Dim vDIPoint As Integer
   
vDIPoint = Me.LBLDI.Caption
If vDIPoint = "01" Then
For Each prnPrinter In Printers
   If prnPrinter.DeviceName = "\\hptc-5100418\SRP370A" Then
      Set Printer = prnPrinter
      Exit For
   End If
Next
End If

If vDIPoint = "02" Then
For Each prnPrinter In Printers
   If prnPrinter.DeviceName = "\\hptc-5100421\SRP370B" Then
      Set Printer = prnPrinter
      Exit For
   End If
Next
End If

If vDIPoint = "03" Then
For Each prnPrinter In Printers
   If prnPrinter.DeviceName = "\\HPTC-5100420\SRP370C" Then
      Set Printer = prnPrinter
      Exit For
   End If
Next
End If

vQuery = "exec dbo.USP_NP_SearchDriveInDetails '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then

Printer.Font.Name = "AngsanaUPC"
Printer.Font.Size = 12
Printer.CurrentX = 1450
Printer.FontBold = True
Printer.Print Trim("DriveIn Slip Master")

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
Printer.Print Trim("คลัง : ") & Trim(vRecordset.Fields("whcode").Value) & "/" & Trim(vRecordset.Fields("shelfcode").Value) & "     " & "โซนการจัด :" & Trim(vRecordset.Fields("zoneid").Value)

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

Public Sub PrintDriveInDetails(vDocNo As String, vZoneID As String)
Dim vPrinterName As String
Dim printerObj As Printer
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim i As Integer
Dim prnPrinter As Printer
Dim n As Integer
Dim vDIPoint As Integer

   
If vZoneID = "01" Then
For Each prnPrinter In Printers
   'If prnPrinter.DeviceName = "\\hptc-5100418\SRP370A" Then
      If prnPrinter.DeviceName = "\\Hptc-5100420\SRP370C" Then
      Set Printer = prnPrinter
      Exit For
   End If
Next
End If

If vZoneID = "02" Then
For Each prnPrinter In Printers
   If prnPrinter.DeviceName = "\\hptc-5100421\SRP370B" Then
      Set Printer = prnPrinter
      Exit For
   End If
Next
End If

If vZoneID = "03" Then
For Each prnPrinter In Printers
   If prnPrinter.DeviceName = "\\HPTC-5100420\SRP370C" Then
      Set Printer = prnPrinter
      Exit For
   End If
Next
End If
        
vQuery = "exec dbo.USP_NP_SearchDriveInPickZoneDetails1 '" & vDocNo & "','" & vZoneID & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then

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
