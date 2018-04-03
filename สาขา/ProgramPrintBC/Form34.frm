VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form34 
   Caption         =   "หน้าพิมพ์เอกสารขาย"
   ClientHeight    =   8340
   ClientLeft      =   2655
   ClientTop       =   1410
   ClientWidth     =   12000
   Icon            =   "Form34.frx":0000
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   Picture         =   "Form34.frx":08CA
   ScaleHeight     =   8340
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport342 
      Left            =   6840
      Top             =   7065
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
   Begin Crystal.CrystalReport CrystalReport341 
      Left            =   7320
      Top             =   7080
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
   Begin VB.OptionButton Option1 
      BackColor       =   &H8000000E&
      Caption         =   "ไม่โชว์ส่วนลดรายตัว"
      Height          =   315
      Left            =   2205
      TabIndex        =   10
      Top             =   5985
      Width           =   2595
   End
   Begin VB.CommandButton CMD342 
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
      Height          =   465
      Left            =   5175
      TabIndex        =   9
      Top             =   1200
      Width           =   1365
   End
   Begin VB.CommandButton CMD341 
      Caption         =   "&Print"
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
      Left            =   9495
      TabIndex        =   3
      Top             =   5445
      Width           =   1020
   End
   Begin VB.TextBox TXT342 
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
      Left            =   2160
      TabIndex        =   5
      Top             =   6480
      Width           =   2640
   End
   Begin VB.TextBox TXT341 
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
      Left            =   2175
      TabIndex        =   4
      Top             =   5475
      Width           =   2640
   End
   Begin VB.Timer Timer341 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   7830
      Top             =   7065
   End
   Begin MSComctlLib.ListView ListView342 
      Height          =   3240
      Left            =   6900
      TabIndex        =   1
      Top             =   1725
      Width           =   3615
      _ExtentX        =   6376
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
         Object.Width           =   6261
      EndProperty
   End
   Begin MSComctlLib.ListView ListView341 
      Height          =   3240
      Left            =   300
      TabIndex        =   0
      Top             =   1725
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
         Text            =   "เลขที่เอกสาร"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ฟอร์มที่พิมพ์"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ชื่อลูกค้า"
         Object.Width           =   4145
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "พิมพ์เอกสารสำคัญ/ใบจ่ายสินค้า"
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
      TabIndex        =   8
      Top             =   300
      Width           =   7515
   End
   Begin VB.Label LBL342 
      BackStyle       =   0  'Transparent
      Caption         =   "ฟอร์มที่พิมพ์"
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   1125
      TabIndex        =   7
      Top             =   6480
      Width           =   915
   End
   Begin VB.Label LBL341 
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่เอกสาร"
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   1125
      TabIndex        =   6
      Top             =   5475
      Width           =   1065
   End
   Begin VB.Label LBLTime341 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   8775
      TabIndex        =   2
      Top             =   1425
      Width           =   1665
   End
End
Attribute VB_Name = "Form34"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vBillType As Integer
Dim vAR As String

Private Sub CMD341_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vAutoNumber As String, vFormat As String
Dim vShelfCode() As String
Dim vWHCode() As String
Dim i As Integer, vPrint As Integer
Dim vCountShelfCode As Integer
Dim vCompleteSave As Integer
Dim n As Integer
Dim vCheckItemPriceOver As Integer
Dim vAnswer As Integer

If Trim(TXT341.Text) <> "" And Trim(TXT342.Text) <> "" Then

   vDocNo = Trim(TXT341.Text)
   vQuery = "select  docno,isnull(iscompletesave,0) as iscompletesave  from dbo.bcarinvoice where docno = '" & vDocNo & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vCompleteSave = Trim(vRecordset.Fields("iscompletesave").Value)
   End If
   vRecordset.Close

   If vCompleteSave = 0 Then
      MsgBox "ไม่สามารถพิมพ์เอกสารที่ยังบันทึกข้อมูลไม่สมบูรณ์ได้ กรุณารอสักครู่แล้วกดพิมพ์ใหม่อีกครั้ง", vbCritical, "Send Error Message"
      Exit Sub
   End If
   
   vQuery = "exec dbo.USP_INV_CheckItemPriceOver '" & vDocNo & "'"
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vCheckItemPriceOver = Trim(vRecordset.Fields("checkitempriceover").Value)
   End If
   vRecordset.Close
   
   If vCheckItemPriceOver > 0 Then
        vAnswer = MsgBox("เอกสารนี้ มีการขายเกินราคาปกติ ต้องการพิมพ์เอกสารหรือไม่ ?", vbYesNo, "Question Message")
        
        If vAnswer = 7 Then
                Exit Sub
        Else
        
        vQuery = "exec dbo.USP_INV_InsertPrintPriceOver '" & vDocNo & "','" & vUserID & "'"
        gConnection.Execute vQuery
        
        End If
   End If
   
    vQuery = "exec dbo.USP_NP_UpdateCreditBill '" & vDocNo & "' "
    gConnection.Execute vQuery

    If Trim(TXT342.Text) = Trim("พิมพ์ใบเสร็จ+ใบจ่ายสินค้า") Then
    
      'vQuery = "exec dbo.USP_INV_SearchShelfPrintSlip'" & vDocNo & "' "
      'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       '  n = vRecordset.RecordCount
        
        'ReDim vWHCode(n) As String
        'ReDim vShelfCode(n) As String
        
        'vRecordset.MoveFirst
        'vCountShelfCode = 0

        
        'For i = 1 To vRecordset.RecordCount
        'vWHCode(i) = Trim(vRecordset.Fields("whcode").Value)
        'vShelfCode(i) = Trim(vRecordset.Fields("shelfcode").Value)
        'vCountShelfCode = vCountShelfCode + 1
        'vRecordset.MoveNext
        'Next i
      'End If
      'vRecordset.Close
      
     
      Call PrintInvoice
      Call PrintBillPayItemZone(vDocNo)
      
      'For i = 1 To vCountShelfCode
      'If vShelfCode(i) = "AVL" Then
       '   Call PrintItem_AVL(vWHCode(i))
      'ElseIf vShelfCode(i) = "BK1" Then
       '   Call PrintItem_BK1(vWHCode(i))
      'ElseIf vShelfCode(i) = "BK2" Then
       '   Call PrintItem_BK2(vWHCode(i))
      'ElseIf vShelfCode(i) = "BK3" Then
       '   Call PrintItem_BK3(vWHCode(i))
      'ElseIf vShelfCode(i) = "SPO" Then
       '   Call PrintItem_SPO(vWHCode(i))
      'ElseIf vShelfCode(i) = "OFS" Then
       '   Call PrintItem_OFS(vWHCode(i))
      'ElseIf vShelfCode(i) = "SHW" Then
       '   Call PrintItem_SHW(vWHCode(i))
      'ElseIf vShelfCode(i) = "RSV" Then
       '   Call PrintItem_RSV(vWHCode(i))
      'ElseIf vShelfCode(i) = "DMG" Then
       '   Call PrintItem_DMG(vWHCode(i))
      'ElseIf vShelfCode(i) = "VND" Then
       '   Call PrintItem_VND(vWHCode(i))
      'ElseIf vShelfCode(i) = "PRO" Then
       '   Call PrintItem_PRO(vWHCode(i))
      'End If
      'Next i
      
    ElseIf Trim(TXT342.Text) = Trim("พิมพ์ใบจ่ายสินค้า") Then
      vDocNo = Trim(TXT341.Text)
      Call PrintBillPayItemZone(vDocNo)
            
      'vQuery = "exec dbo.USP_INV_SearchShelfPrintSlip'" & vDocNo & "' "
      'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      
       ' n = vRecordset.RecordCount
        
        'ReDim vWHCode(n) As String
        'ReDim vShelfCode(n) As String
        
        'vRecordset.MoveFirst
        'vCountShelfCode = 0
        'For i = 1 To vRecordset.RecordCount
        'vWHCode(i) = Trim(vRecordset.Fields("whcode").Value)
        'vShelfCode(i) = Trim(vRecordset.Fields("shelfcode").Value)
        'vCountShelfCode = vCountShelfCode + 1
        'vRecordset.MoveNext
        'Next i
      'End If
      'vRecordset.Close
            
      'For i = 1 To vCountShelfCode
      'If vShelfCode(i) = "AVL" Then
       '   Call PrintItem_AVL(vWHCode(i))
      'ElseIf vShelfCode(i) = "BK1" Then
       '   Call PrintItem_BK1(vWHCode(i))
      'ElseIf vShelfCode(i) = "BK2" Then
       '   Call PrintItem_BK2(vWHCode(i))
      'ElseIf vShelfCode(i) = "BK3" Then
       '   Call PrintItem_BK3(vWHCode(i))
      'ElseIf vShelfCode(i) = "SPO" Then
       '   Call PrintItem_SPO(vWHCode(i))
      'ElseIf vShelfCode(i) = "OFS" Then
       '   Call PrintItem_OFS(vWHCode(i))
      'ElseIf vShelfCode(i) = "SHW" Then
       '   Call PrintItem_SHW(vWHCode(i))
      'ElseIf vShelfCode(i) = "PRO" Then
       '   Call PrintItem_PRO(vWHCode(i))
      'End If
      'Next i
            
    ElseIf Trim(TXT342.Text) = Trim("พิมพ์ใบเสร็จ") Then
      Call PrintInvoice
    ElseIf Trim(TXT342.Text) = Trim("พิมพ์ใบเสร็จ เฉพาะกิจ") Then
      Call PrintInvoice_NoVat
    End If
    
    vQuery = "select printed from  bcnp.dbo.vw_sl_00001 where docno = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vPrint = Trim(vRecordset.Fields("printed").Value)
    End If
    vRecordset.Close
    If vPrint = 0 Then
      vQuery = "Update npmaster.dbo.npprintserver set Printed = 1 where Docno = '" & vDocNo & "' "
      gConnection.Execute vQuery
      ListView341.ListItems.Remove (ListView341.SelectedItem.Index)
    End If
    TXT341.Text = ""
    TXT342.Text = ""
    Option1.Value = False
Else
  MsgBox "ไม่มีเลขที่ให้พิมพ์ หรือ ไม่ได้เลือกฟอร์มที่จะพิมพ์ หรือ ไม่ได้เลือกเลขที่เอกสาร", vbCritical, "ข้อความแจ้งเตือน"
End If

ErrDescription:
If Err.Description <> "" Then
vQuery = "Update npmaster.dbo.npprintserver set Printed = 1 where Docno = '" & vDocNo & "' "
gConnection.Execute vQuery
MsgBox Err.Description
Exit Sub
End If
End Sub


Public Sub PrintBillPayItemZone(vDocNo As String)
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String

Dim vWHCode(20) As String
Dim vShelfCode(20) As String
Dim vPickZone(20) As String
Dim vZoneID(20) As String
Dim i As Integer, vPrint As Integer
Dim vCountShelfCode As Integer
Dim vCompleteSave As Integer


On Error GoTo ErrDescription


vDocNo = Trim(TXT341.Text)

 vQuery = "exec dbo.USP_INV_SearchShelfPrintSlip '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vRecordset.MoveFirst
vCountShelfCode = 0
For i = 1 To vRecordset.RecordCount
  vWHCode(i) = Trim(vRecordset.Fields("whcode").Value)
  vShelfCode(i) = Trim(vRecordset.Fields("shelfcode").Value)
  vCountShelfCode = vCountShelfCode + 1
vRecordset.MoveNext
Next i
End If
vRecordset.Close
      
      
For i = 1 To vCountShelfCode
If vWHCode(i) = "S02" Then
    If vShelfCode(i) = "AVL" Then
      Call InsertPayNumber(vDocNo, vWHCode(i), vShelfCode(i))
    ElseIf vShelfCode(i) = "BAK" Then
      Call InsertPayNumber(vDocNo, vWHCode(i), vShelfCode(i))
    ElseIf vShelfCode(i) = "SPO" Then
      Call InsertPayNumber(vDocNo, vWHCode(i), vShelfCode(i))
    ElseIf vShelfCode(i) = "OFS" Then
      Call InsertPayNumber(vDocNo, vWHCode(i), vShelfCode(i))
    ElseIf vShelfCode(i) = "SHW" Then
      Call InsertPayNumber(vDocNo, vWHCode(i), vShelfCode(i))
    ElseIf vShelfCode(i) = "RSV" Then
      Call InsertPayNumber(vDocNo, vWHCode(i), vShelfCode(i))
    ElseIf vShelfCode(i) = "DMG" Then
      Call InsertPayNumber(vDocNo, vWHCode(i), vShelfCode(i))
    ElseIf vShelfCode(i) = "VND" Then
      Call InsertPayNumber(vDocNo, vWHCode(i), vShelfCode(i))
    End If
Else
    If vWHCode(i) <> "S0-PASS" Then
      Call InsertPayNumber(vDocNo, vWHCode(i), vShelfCode(i))
    End If
End If

Next i


If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'Left(vDocNo, 3)
Else
vDocno1 = UCase(Left(vDocNo, 3))
End If

vQuery = "select upper('" & vDocno1 & "') as vDocGroup"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vDocGroup1 = UCase(Trim(vRecordset.Fields("vDocGroup").Value))
End If
vRecordset.Close

If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 484
vRepType = "INV"
Else
vRepID = 482
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With CrystalReport341
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub


Public Sub InsertPayNumber(vDocNo As String, vWHCode As String, vShelfCode As String)
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String

On Error Resume Next

vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
    vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close

vGenerateNumber = vHeader & "-" & vAutoNumber

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','','" & vUserID & "' "
gConnection.Execute vQuery
End Sub


Public Sub PrintItem_AVL(vWHCode As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vShelfCode As String

On Error GoTo ErrDescription

vDocNo = Trim(TXT341.Text)
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
    vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vShelfCode = "AVL"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','','" & vUserID & "' "
gConnection.Execute vQuery

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'Left(vDocNo, 3)
Else
vDocno1 = UCase(Left(vDocNo, 3))
End If

vQuery = "select upper('" & vDocno1 & "') as vDocGroup"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vDocGroup1 = UCase(Trim(vRecordset.Fields("vDocGroup").Value))
End If
vRecordset.Close

If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 425
vRepType = "INV"
Else
vRepID = 423
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With CrystalReport341
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
        .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem_PRO(vWHCode As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vShelfCode As String

On Error GoTo ErrDescription

vDocNo = Trim(TXT341.Text)
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
    vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vShelfCode = "PRO"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','','" & vUserID & "' "
gConnection.Execute vQuery

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'Left(vDocNo, 3)
Else
vDocno1 = UCase(Left(vDocNo, 3))
End If

vQuery = "select upper('" & vDocno1 & "') as vDocGroup"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vDocGroup1 = UCase(Trim(vRecordset.Fields("vDocGroup").Value))
End If
vRecordset.Close

If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 471
vRepType = "INV"
Else
vRepID = 472
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With CrystalReport341
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
        .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub


Public Sub PrintItem_BK3(vWHCode As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vShelfCode As String

On Error GoTo ErrDescription

vDocNo = Trim(TXT341.Text)
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
    vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vShelfCode = "BK3"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','','" & vUserID & "' "
gConnection.Execute vQuery

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'Left(vDocNo, 3)
Else
vDocno1 = UCase(Left(vDocNo, 3))
End If

vQuery = "select upper('" & vDocno1 & "') as vDocGroup"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vDocGroup1 = UCase(Trim(vRecordset.Fields("vDocGroup").Value))
End If
vRecordset.Close

If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 437
vRepType = "INV"
Else
vRepID = 435
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With CrystalReport341
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
        .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem_BK2(vWHCode As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vShelfCode As String

On Error GoTo ErrDescription

vDocNo = Trim(TXT341.Text)
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
    vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vShelfCode = "BK2"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','','" & vUserID & "' "
gConnection.Execute vQuery

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'Left(vDocNo, 3)
Else
vDocno1 = UCase(Left(vDocNo, 3))
End If

vQuery = "select upper('" & vDocno1 & "') as vDocGroup"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vDocGroup1 = UCase(Trim(vRecordset.Fields("vDocGroup").Value))
End If
vRecordset.Close

If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 433
vRepType = "INV"
Else
vRepID = 431
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With CrystalReport341
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
        .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem_BK1(vWHCode As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vShelfCode As String

On Error GoTo ErrDescription

vDocNo = Trim(TXT341.Text)
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
    vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vShelfCode = "BK1"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','','" & vUserID & "' "
gConnection.Execute vQuery

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'Left(vDocNo, 3)
Else
vDocno1 = UCase(Left(vDocNo, 3))
End If

vQuery = "select upper('" & vDocno1 & "') as vDocGroup"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vDocGroup1 = UCase(Trim(vRecordset.Fields("vDocGroup").Value))
End If
vRecordset.Close

If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 429
vRepType = "INV"
Else
vRepID = 427
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With CrystalReport341
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
        .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem_SPO(vWHCode As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vShelfCode As String

On Error GoTo ErrDescription

vDocNo = Trim(TXT341.Text)
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
    vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vShelfCode = "SPO"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','','" & vUserID & "' "
gConnection.Execute vQuery

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'Left(vDocNo, 3)
Else
vDocno1 = UCase(Left(vDocNo, 3))
End If

vQuery = "select upper('" & vDocno1 & "') as vDocGroup"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vDocGroup1 = UCase(Trim(vRecordset.Fields("vDocGroup").Value))
End If
vRecordset.Close

If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 441
vRepType = "INV"
Else
vRepID = 439
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With CrystalReport341
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
        .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem_DMG(vWHCode As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vShelfCode As String

On Error GoTo ErrDescription

vDocNo = Trim(TXT341.Text)
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
    vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vShelfCode = "DMG"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','','" & vUserID & "' "
gConnection.Execute vQuery

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'Left(vDocNo, 3)
Else
vDocno1 = UCase(Left(vDocNo, 3))
End If

vQuery = "select upper('" & vDocno1 & "') as vDocGroup"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vDocGroup1 = UCase(Trim(vRecordset.Fields("vDocGroup").Value))
End If
vRecordset.Close

If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 456
vRepType = "INV"
Else
vRepID = 455
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With CrystalReport341
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
        .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem_VND(vWHCode As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vShelfCode As String

On Error GoTo ErrDescription

vDocNo = Trim(TXT341.Text)
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
    vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vShelfCode = "VND"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','','" & vUserID & "' "
gConnection.Execute vQuery

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'Left(vDocNo, 3)
Else
vDocno1 = UCase(Left(vDocNo, 3))
End If

vQuery = "select upper('" & vDocno1 & "') as vDocGroup"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vDocGroup1 = UCase(Trim(vRecordset.Fields("vDocGroup").Value))
End If
vRecordset.Close

If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 460
vRepType = "INV"
Else
vRepID = 459
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With CrystalReport341
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
        .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem_OFS(vWHCode As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vShelfCode As String

On Error GoTo ErrDescription

vDocNo = Trim(TXT341.Text)
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
    vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vShelfCode = "OFS"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','','" & vUserID & "' "
gConnection.Execute vQuery

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'Left(vDocNo, 3)
Else
vDocno1 = UCase(Left(vDocNo, 3))
End If

vQuery = "select upper('" & vDocno1 & "') as vDocGroup"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vDocGroup1 = UCase(Trim(vRecordset.Fields("vDocGroup").Value))
End If
vRecordset.Close

If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 444
vRepType = "INV"
Else
vRepID = 443
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With CrystalReport341
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
        .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem_SHW(vWHCode As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vShelfCode As String

On Error GoTo ErrDescription

vDocNo = Trim(TXT341.Text)
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
    vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vShelfCode = "SHW"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','','" & vUserID & "' "
gConnection.Execute vQuery

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'Left(vDocNo, 3)
Else
vDocno1 = UCase(Left(vDocNo, 3))
End If

vQuery = "select upper('" & vDocno1 & "') as vDocGroup"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vDocGroup1 = UCase(Trim(vRecordset.Fields("vDocGroup").Value))
End If
vRecordset.Close

If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 448
vRepType = "INV"
Else
vRepID = 447
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With CrystalReport341
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
        .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem_RSV(vWHCode As String)
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String
Dim vShelfCode As String

On Error GoTo ErrDescription

vDocNo = Trim(TXT341.Text)
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
    vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber
vShelfCode = "RSV"

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 8 "
gConnection.Execute vQuery

vQuery = "exec dbo.USP_INV_InsertPrintSlip '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vShelfCode & "','','" & vUserID & "' "
gConnection.Execute vQuery

If UCase(Left(vDocNo, 3)) = "S01" Or UCase(Left(vDocNo, 3)) = "S02" Or UCase(Left(vDocNo, 3)) = "W01" Or UCase(Left(vDocNo, 3)) = "W02" Then
vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'Left(vDocNo, 3)
Else
vDocno1 = UCase(Left(vDocNo, 3))
End If

vQuery = "select upper('" & vDocno1 & "') as vDocGroup"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vDocGroup1 = UCase(Trim(vRecordset.Fields("vDocGroup").Value))
End If
vRecordset.Close

If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 452
vRepType = "INV"
Else
vRepID = 451
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With CrystalReport341
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
        .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub


'Private Sub CMD341_Click()
'Dim vQuery As String
'Dim vRecordset As New ADODB.Recordset
'Dim vDocNo As String
'Dim vAutoNumber As String, vFormat As String
'Dim vWHCode(10) As String
'Dim i As Integer, vPrint As Integer
'Dim vCountWHCode As Integer

'If Trim(TXT341.Text) <> "" And Trim(TXT342.Text) <> "" Then
 '   If Trim(TXT342.Text) = Trim("พิมพ์ใบเสร็จ") Then
  '    vDocNo = Trim(TXT341.Text)
   '
    '  vQuery = "select docno,billtype,arcode from dbo.bcarinvoice where docno = '" & vDocNo & "' "
     ' If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      '  vBillType = Trim(vRecordset.Fields("billtype").Value)
       ' vAR = Trim(vRecordset.Fields("arcode").Value)
      'End If
      'vRecordset.Close
      
      'vQuery = "select  docno,whcode from bcarinvoicesub where docno = '" & vDocNo & "' group by docno,whcode "
      'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
       ' vRecordset.MoveFirst
        'vCountWHCode = 0
        'For i = 1 To vRecordset.RecordCount
        'vWHCode(i) = Trim(vRecordset.Fields("whcode").Value)
        'vCountWHCode = vCountWHCode + 1
        'vRecordset.MoveNext
        'Next i
      'End If
      'vRecordset.Close
      'Call PrintInvoice
      'For i = 1 To vCountWHCode
      'If vWHCode(i) = "010" Then
       '   Call PrintItem010
      'ElseIf vWHCode(i) = "012" Then
       '   Call PrintItem012
      'ElseIf vWHCode(i) = "014" Then
       '   Call PrintItem014
      'ElseIf vWHCode(i) = "015" Then
       '   Call PrintItem015
      'ElseIf vWHCode(i) = "020" Then
       '   Call PrintItem020
      'ElseIf vWHCode(i) = "097" Then
       '   Call PrintItem097
      'ElseIf vWHCode(i) = "016" Then
       '   Call PrintItem016
      'ElseIf vWHCode(i) = "070" Then
       '   Call PrintItem070
      'End If
      'Next i
    'End If
    'vQuery = "select printed from  bcnp.dbo.vw_sl_00001 where docno = '" & vDocNo & "' "
    'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
     '   vPrint = Trim(vRecordset.Fields("printed").Value)
    'End If
    'vRecordset.Close
    'If vPrint = 0 Then
     ' vQuery = "Update npmaster.dbo.npprintserver set Printed = 1 where Docno = '" & vDocNo & "' "
      'gConnection.Execute vQuery
      'ListView341.ListItems.Remove (ListView341.SelectedItem.Index)
    'End If
    'TXT341.Text = ""
    'TXT342.Text = ""
    'Option1.Value = False
'Else
 ' MsgBox "ไม่มีเลขที่ให้พิมพ์ หรือ ไม่ได้เลือกฟอร์มที่จะพิมพ์ หรือ ไม่ได้เลือกเลขที่เอกสาร", vbCritical, "ข้อความแจ้งเตือน"
'End If

'ErrDescription:
'If Err.Description <> "" Then
'vQuery = "Update npmaster.dbo.npprintserver set Printed = 1 where Docno = '" & vDocNo & "' "
'gConnection.Execute vQuery
'MsgBox Err.Description
'Exit Sub
'End If
'End Sub

Private Sub CMD342_Click()
Call RefreshData
End Sub

Private Sub Form_Load()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim INVItemLists As ListItem
Dim INVItemforms As ListItem
Dim vTypeDoc As String
Dim vCheckDate As Date

On Error GoTo ErrDescription

 
 vQuery = "select * from npmaster.dbo.NP_Generate_DocNo where headertype = 8"
 If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckDate = Trim(vRecordset.Fields("dateupdate").Value)
 End If
 vRecordset.Close
 If Year(vCheckDate) <> Year(Now) Or Month(vCheckDate) <> Month(Now) Or Day(vCheckDate) <> Day(Now) Then
 Call GenHeadDocument
 vQuery = "Update npmaster.dbo.NP_Generate_DocNo set header = '" & vGenDocNo & "', dateupdate = getdate(),autonumber = 1 where headertype = 8 "
 gConnection.Execute vQuery
 End If
 
 vTypeDoc = "INV"

ListView341.ListItems.Clear
 vQuery = "Select Docno,LastPrintDateTime,name1,printed  from BCNP.dbo.vw_sl_00001 where Printed = 0  " _
                        & " and LastPrintedUser = '" & vUserID & "' and DoctypeID = '" & vTypeDoc & "'  order by LastPrintDateTime "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            If Not vRecordset.EOF Then
                vRecordset.MoveFirst
                While Not vRecordset.EOF
                Set INVItemLists = ListView341.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
                INVItemLists.SubItems(1) = Trim(vRecordset.Fields("LastPrintDateTime").Value)
                INVItemLists.SubItems(2) = Trim(vRecordset.Fields("name1").Value)
                vRecordset.MoveNext
                Wend
            End If
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------------------
        
        ListView342.ListItems.Clear
        vQuery = "select Name from npmaster.dbo.NPForm where ModuleID = '" & vTypeDoc & "' and typereport = 0 "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
            Set INVItemforms = ListView342.ListItems.Add(, , Trim(vRecordset.Fields("Name").Value))
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


Public Sub PrintInvoice_NoVat()
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vDocNo As String
        Dim vGroupDoc As String
        Dim vReportName As String
        Dim vName As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vCheck As Integer
        Dim vDueDate As String
        Dim vCredit As Integer

        vCredit = 0
        vDocNo = Trim(TXT341.Text)
        On Error GoTo PrintError
         vQuery = "select groupdoc from  bcnp.dbo.vw_sl_00001 where docno = '" & vDocNo & "' "
         If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                 vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
         End If
         vRecordset.Close
         
         vRepType = "INV"
         vRepID = 493
                 
 
          If vCredit = 1 Then
            vQuery = "exec dbo.USP_AR_SearchDueDateInvoice '" & vDocNo & "' "
            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
              vDueDate = Trim(vRecordset.Fields("conditpaycode").Value)
            End If
            vRecordset.Close
            
          Else
            vDueDate = ""
          End If
        
          vCheck = 1
          vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
          If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            With CrystalReport341
                .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                .ParameterFields(0) = "@vDocNo;" & vDocNo & ";true"
                .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
                .Destination = crptToWindow
                .WindowState = crptMaximized
                .Action = 1
            End With
          End If
          vRecordset.Close
                            
                            '---------------------------------------------------------------------------------------------------------
PrintError:
                If Err.Description <> "" Then
                MsgBox Err.Description
                End If


End Sub



Public Sub PrintInvoice()
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vDocNo As String
        Dim vGroupDoc As String
        Dim vReportName As String
        Dim vName As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vCheck As Integer
        Dim vDueDate As String
        Dim vCredit As Integer

        vCredit = 0
        vDocNo = Trim(TXT341.Text)
        On Error GoTo PrintError
         vQuery = "select groupdoc from  bcnp.dbo.vw_sl_00001 where docno = '" & vDocNo & "' "
         If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                 vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
         End If
         vRecordset.Close
         vRepType = "INV"
             If vGroupDoc = "IHV" Or vGroupDoc = "ihv" Then
                 If Option1.Value = False Then
                     vRepID = 50
                 Else
                     vRepID = 150
                 End If
             ElseIf vGroupDoc = "IHN" Or vGroupDoc = "ihn" Then
                 If Option1.Value = False Then
                     vRepID = 51
                 Else
                     vRepID = 151
                 End If
             ElseIf vGroupDoc = "ICV" Or vGroupDoc = "icv" Then
                 vCredit = 1
                 If Option1.Value = False Then
                     vRepID = 52
                 Else
                     vRepID = 152
                 End If
             ElseIf vGroupDoc = "ICN" Or vGroupDoc = "icn" Then
                 vCredit = 1
                 If Option1.Value = False Then
                     vRepID = 53
                 Else
                     vRepID = 153
                 End If
             ElseIf vGroupDoc = "IVD" Or vGroupDoc = "ivd" Then
             MsgBox "ไม่สามารถออกเอกสารเก็บเงินปลายทางได้", vbInformation + vbCritical, "ข้อความเตือน"
               'vCredit = 1
                ' If Option1.Value = False Then
                 '    vRepID = 52
                 'Else
                  '   vRepID = 152
                 'End If
             ElseIf vGroupDoc = "IVM" Or vGroupDoc = "ivm" Then
             MsgBox "ไม่สามารถออกเอกสารเก็บเงินปลายทางได้", vbInformation + vbCritical, "ข้อความเตือน"
               'vCredit = 1
                ' If Option1.Value = False Then
                 '    vRepID = 52
                 'Else
                  '   vRepID = 152
                 'End If
             ElseIf vGroupDoc = "IVN" Or vGroupDoc = "ivn" Then
             MsgBox "ไม่สามารถออกเอกสารเก็บเงินปลายทางได้", vbInformation + vbCritical, "ข้อความเตือน"
               'vCredit = 1
                ' If Option1.Value = False Then
                 '    vRepID = 53
                 'Else
                  '   vRepID = 153
                 'End If
             ElseIf vGroupDoc = "IAB" Or vGroupDoc = "iab" Then
                 If Option1.Value = False Then
                     vRepID = 54
                 Else
                     vRepID = 154
                 End If
             ElseIf vGroupDoc = "IVE" Or vGroupDoc = "ive" Then
                 vRepID = 55
             Else
             MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
             Exit Sub
             End If
 
             If vCredit = 1 Then
               vQuery = "exec dbo.USP_AR_SearchDueDateInvoice '" & vDocNo & "' "
               If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                 vDueDate = Trim(vRecordset.Fields("conditpaycode").Value)
               End If
               vRecordset.Close
               
             Else
               vDueDate = ""
             End If
           
             vCheck = 1
             vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
             'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
             If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                 With CrystalReport342
                     .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                     .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                     .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
                     .Formulas(0) = "CreditCondition='" & vDueDate & "' "
                     .Destination = crptToWindow
                     .WindowState = crptMaximized
                     .Action = 1
                 End With
             End If
             vRecordset.Close
                            
                            '---------------------------------------------------------------------------------------------------------
PrintError:
                If Err.Description <> "" Then
                MsgBox Err.Description
                End If


End Sub

Private Sub ListView341_ItemClick(ByVal Item As MSComctlLib.ListItem)
TXT341.Text = Item
End Sub

Private Sub ListView341_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Timer341.Enabled = False
End Sub

Private Sub ListView342_ItemClick(ByVal Item As MSComctlLib.ListItem)
TXT342.Text = Item
End Sub

Private Sub Timer341_Timer()
Dim vTime, vSubTime, vTotalTime
Dim vNumber As Integer

On Error GoTo ErrDescription

            If LBLTime341.Caption <> CStr(Time) Then
                 LBLTime341.Caption = Time
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

            vDocHeader = "INV"

            vQuery = "Select Docno,LastPrintDateTime,name1,printed  from BCNP.dbo.vw_sl_00001   where Printed = 0 " _
                        & " and LastPrintedUser = '" & vUserID & "' and DoctypeID = '" & vDocHeader & "' order by LastPrintDateTime "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    If Not vRecordset.EOF Then
                        CountRecordset = vRecordset.RecordCount
                        CountList = ListView341.ListItems.Count
                                If CountRecordset > CountList Then
                                vRecordset.MoveFirst
                                For i = 1 To CountRecordset
                                If i < CountRecordset Then
                                vDocNo = ListView341.ListItems.Item(i).Text
                                End If
                                vNewDoc = Trim(vRecordset.Fields("Docno").Value)
                                vPrintStatus = Trim(vRecordset.Fields("Printed").Value)
                                        If vDocNo <> vNewDoc Then
                                        Set DocListItem = ListView341.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
                                        DocListItem.SubItems(1) = Trim(vRecordset.Fields("LastPrintDateTime").Value)
                                        DocListItem.SubItems(2) = Trim(vRecordset.Fields("name1").Value)
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
    vTypeDoc = "INV"
    ListView341.ListItems.Clear
    vQuery = "Select Docno,LastPrintDateTime,name1,printed  from BCNP.dbo.vw_sl_00001  where Printed = 0  " _
                        & " and LastPrintedUser = '" & vUserID & "' and DoctypeID = '" & vTypeDoc & "'  order by LastPrintDateTime "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            If Not vRecordset.EOF Then
                vRecordset.MoveFirst
                While Not vRecordset.EOF
                Set INVItemLists = ListView341.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
                INVItemLists.SubItems(1) = Trim(vRecordset.Fields("LastPrintDateTime").Value)
                INVItemLists.SubItems(2) = Trim(vRecordset.Fields("name1").Value)
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

'Public Sub PrintItem010()
'Dim vDocno As String
'Dim vQuery As String
'Dim vRecordset As New ADODB.Recordset
'Dim vAutoNumber As String
'Dim vGenerateNumber As String
'Dim vHeader As String
'Dim vDocdate As Date
'Dim vSaleType As Integer
'Dim vARCode As String
'Dim vInvoiceNo As String
'Dim vWHCode As String
'Dim vZoneID As String


'On Error GoTo ErrDescription

'vDocno = Trim(TXT341.Text)
'vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 '   vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
  '  vHeader = Trim(vRecordset.Fields("header").Value)
'End If
'vRecordset.Close
'vGenerateNumber = vHeader & "-" & vAutoNumber

'vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
'gConnection.Execute vQuery

'vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
'vSaleType = vBillType
'vARCode = vAR
'vInvoiceNo = vDocno
'vWHCode = "010"
'vZoneID = "01"
    
'vQuery = "exec dbo.USP_QUE_InsertCustItemReceipt '" & vGenerateNumber & "','" & vDocdate & "'," & vSaleType & ",'" & vARCode & "','" & vDocno & "','" & vWHCode & "','" & vZoneID & "' "
'gConnection.Execute vQuery
        

'ErrDescription:
'If Err.Description <> "" Then
'MsgBox Err.Description
'End If
'End Sub

'Public Sub PrintItem012()
'Dim vDocno As String
'Dim vQuery As String
'Dim vRecordset As New ADODB.Recordset
'Dim vAutoNumber As String
'Dim vGenerateNumber As String
'Dim vHeader As String
'Dim vDocdate As Date
'Dim vSaleType As Integer
'Dim vARCode As String
'Dim vInvoiceNo As String
'Dim vWHCode As String
'Dim vZoneID As String


'On Error GoTo ErrDescription

'vDocno = Trim(TXT341.Text)
'vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 '   vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
  '  vHeader = Trim(vRecordset.Fields("header").Value)
'End If
'vRecordset.Close
'vGenerateNumber = vHeader & "-" & vAutoNumber

'vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 8 "
'gConnection.Execute vQuery

'vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
'vSaleType = vBillType
'vARCode = vAR
'vInvoiceNo = vDocno
'vWHCode = "012"
'vZoneID = "01"
    
'vQuery = "exec dbo.USP_QUE_InsertCustItemReceipt '" & vGenerateNumber & "','" & vDocdate & "'," & vSaleType & ",'" & vARCode & "','" & vDocno & "','" & vWHCode & "','" & vZoneID & "' "
'gConnection.Execute vQuery
        
                
'ErrDescription:
'If Err.Description <> "" Then
'MsgBox Err.Description
'End If
'End Sub

'Public Sub PrintItem015()
'Dim vDocno As String
'Dim vQuery As String
'Dim vRecordset As New ADODB.Recordset
'Dim vAutoNumber As String
'Dim vGenerateNumber As String
'Dim vHeader As String
'Dim vDocdate As Date
'Dim vSaleType As Integer
'Dim vARCode As String
'Dim vInvoiceNo As String
'Dim vWHCode As String
'Dim vZoneID As String

'On Error GoTo ErrDescription

'vDocno = Trim(TXT341.Text)
'vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 '   vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
  '  vHeader = Trim(vRecordset.Fields("header").Value)
'End If
'vRecordset.Close
'vGenerateNumber = vHeader & "-" & vAutoNumber

'vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 8 "
'gConnection.Execute vQuery

        
'vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
'vSaleType = vBillType
'vARCode = vAR
'vInvoiceNo = vDocno
'vWHCode = "015"
'vZoneID = "01"
    
'vQuery = "exec dbo.USP_QUE_InsertCustItemReceipt '" & vGenerateNumber & "','" & vDocdate & "'," & vSaleType & ",'" & vARCode & "','" & vDocno & "','" & vWHCode & "','" & vZoneID & "' "
'gConnection.Execute vQuery

                
'ErrDescription:
'If Err.Description <> "" Then
'MsgBox Err.Description
'End If
'End Sub

'Public Sub PrintItem097()
'Dim vDocno As String
'Dim vQuery As String
'Dim vRecordset As New ADODB.Recordset
'Dim vAutoNumber As String
'Dim vGenerateNumber As String
'Dim vHeader As String
'Dim vDocdate As Date
'Dim vSaleType As Integer
'Dim vARCode As String
'Dim vInvoiceNo As String
'Dim vWHCode As String
'Dim vZoneID As String
      
'On Error GoTo ErrDescription

'vDocno = Trim(TXT341.Text)
'vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 '   vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
  '  vHeader = Trim(vRecordset.Fields("header").Value)
'End If
'vRecordset.Close
'vGenerateNumber = vHeader & "-" & vAutoNumber

'vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 8 "
'gConnection.Execute vQuery

'vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
'vSaleType = vBillType
'vARCode = vAR
'vInvoiceNo = vDocno
'vWHCode = "097"
'vZoneID = "01"
    
'vQuery = "exec dbo.USP_QUE_InsertCustItemReceipt '" & vGenerateNumber & "','" & vDocdate & "'," & vSaleType & ",'" & vARCode & "','" & vDocno & "','" & vWHCode & "','" & vZoneID & "' "
'gConnection.Execute vQuery

                
'ErrDescription:
'If Err.Description <> "" Then
'MsgBox Err.Description
'End If
'End Sub


'Public Sub PrintItem014()
'Dim vDocno As String
'Dim vQuery As String
'Dim vRecordset As New ADODB.Recordset
'Dim vAutoNumber As String
'Dim vGenerateNumber As String
'Dim vHeader As String
'Dim vDocdate As Date
'Dim vSaleType As Integer
'Dim vARCode As String
'Dim vInvoiceNo As String
'Dim vWHCode As String
'Dim vZoneID As String
      
'On Error GoTo ErrDescription

'vDocno = Trim(TXT341.Text)
'vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 '   vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
  '  vHeader = Trim(vRecordset.Fields("header").Value)
'End If
'vRecordset.Close
'vGenerateNumber = vHeader & "-" & vAutoNumber

'vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 8 "
'gConnection.Execute vQuery
        
'vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
'vSaleType = vBillType
'vARCode = vAR
'vInvoiceNo = vDocno
'vWHCode = "014"
'vZoneID = "02"
    
'vQuery = "exec dbo.USP_QUE_InsertCustItemReceipt '" & vGenerateNumber & "','" & vDocdate & "'," & vSaleType & ",'" & vARCode & "','" & vDocno & "','" & vWHCode & "','" & vZoneID & "' "
'gConnection.Execute vQuery
        
                
'ErrDescription:
'If Err.Description <> "" Then
'MsgBox Err.Description
'End If
'End Sub

'Public Sub PrintItem020()
'Dim vDocno As String
'Dim vQuery As String
'Dim vRecordset As New ADODB.Recordset
'Dim vAutoNumber As String
'Dim vGenerateNumber As String
'Dim vHeader As String
'Dim vDocdate As Date
'Dim vSaleType As Integer
'Dim vARCode As String
'Dim vInvoiceNo As String
'Dim vWHCode As String
'Dim vZoneID As String
      
'On Error GoTo ErrDescription

'vDocno = Trim(TXT341.Text)
'vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 '   vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
  '  vHeader = Trim(vRecordset.Fields("header").Value)
'End If
'vRecordset.Close
'vGenerateNumber = vHeader & "-" & vAutoNumber

'vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 8 "
'gConnection.Execute vQuery
        
'vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
'vSaleType = vBillType
'vARCode = vAR
'vInvoiceNo = vDocno
'vWHCode = "020"
'vZoneID = "01"
    
'vQuery = "exec dbo.USP_QUE_InsertCustItemReceipt '" & vGenerateNumber & "','" & vDocdate & "'," & vSaleType & ",'" & vARCode & "','" & vDocno & "','" & vWHCode & "','" & vZoneID & "' "
'gConnection.Execute vQuery

                
'ErrDescription:
'If Err.Description <> "" Then
'MsgBox Err.Description
'End If
'End Sub

'Public Sub PrintItem016()
'Dim vDocno As String
'Dim vQuery As String
'Dim vRecordset As New ADODB.Recordset
'Dim vAutoNumber As String
'Dim vGenerateNumber As String
'Dim vHeader As String
'Dim vDocdate As Date
'Dim vSaleType As Integer
'Dim vARCode As String
'Dim vInvoiceNo As String
'Dim vWHCode As String
'Dim vZoneID As String
      
'On Error GoTo ErrDescription

'vDocno = Trim(TXT341.Text)
'vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 '   vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
  '  vHeader = Trim(vRecordset.Fields("header").Value)
'End If
'vRecordset.Close
'vGenerateNumber = vHeader & "-" & vAutoNumber

'vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 8 "
'gConnection.Execute vQuery
        
'vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
'vSaleType = vBillType
'vARCode = vAR
'vInvoiceNo = vDocno
'vWHCode = "016"
'vZoneID = "01"
    
'vQuery = "exec dbo.USP_QUE_InsertCustItemReceipt '" & vGenerateNumber & "','" & vDocdate & "'," & vSaleType & ",'" & vARCode & "','" & vDocno & "','" & vWHCode & "','" & vZoneID & "' "
'gConnection.Execute vQuery
        
                
'ErrDescription:
'If Err.Description <> "" Then
'MsgBox Err.Description
'End If
'End Sub

'Public Sub PrintItem070()
'Dim vDocno As String
'Dim vQuery As String
'Dim vRecordset As New ADODB.Recordset
'Dim vAutoNumber As String
'Dim vGenerateNumber As String
'Dim vHeader As String
'Dim vDocdate As Date
'Dim vSaleType As Integer
'Dim vARCode As String
'Dim vInvoiceNo As String
'Dim vWHCode As String
'Dim vZoneID As String
        
'On Error GoTo ErrDescription

'vDocno = Trim(TXT341.Text)
'vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
 '   vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
  '  vHeader = Trim(vRecordset.Fields("header").Value)
'End If
'vRecordset.Close
'vGenerateNumber = vHeader & "-" & vAutoNumber

'vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 8 "
'gConnection.Execute vQuery

        
'vDocdate = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now))
'vSaleType = vBillType
'vARCode = vAR
'vInvoiceNo = vDocno
'vWHCode = "070"
'vZoneID = "01"
    
'vQuery = "exec dbo.USP_QUE_InsertCustItemReceipt '" & vGenerateNumber & "','" & vDocdate & "'," & vSaleType & ",'" & vARCode & "','" & vDocno & "','" & vWHCode & "','" & vZoneID & "' "
'gConnection.Execute vQuery

                
'ErrDescription:
'If Err.Description <> "" Then
'MsgBox Err.Description
'End If
'End Sub


Public Sub PrintItem010()
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
        Dim vGenerateNumber, vHeader As String
        Dim vCheck As Integer
        Dim vCheckItemLocation015 As Integer
        Dim vCheckItemNotLocation015 As Integer
      
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT341.Text)
        'vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
        'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
         '   vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
          '  vHeader = Trim(vRecordset.Fields("header").Value)
        'End If
        'vRecordset.Close
        'vGenerateNumber = vHeader & "-" & vAutoNumber
        
        'vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
        'gConnection.Execute vQuery
        'vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked) " _
        '& " values ('" & vDocno & "','" & vGenerateNumber & "',getdate(),'010','" & vUserID & "',1,0)"
        'gConnection.Execute vQuery

        
        'vDocno1 = UCase(Left(vDocno, 3))
        'vQuery = "select upper('" & vDocno1 & "') as vDocGroup"
        'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
         '   vDocGroup1 = UCase(Trim(vRecordset.Fields("vDocGroup").Value))
        'End If
        'vRecordset.Close
        
        vQuery = "exec dbo.USP_INV_CheckItemNotLocation015 '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
          vCheckItemNotLocation015 = vRecordset.Fields("vcount").Value
        End If
        vRecordset.Close
        
        vQuery = "exec dbo.USP_INV_CheckItemLocation015 '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
          vCheckItemLocation015 = vRecordset.Fields("vcount").Value
        End If
        vRecordset.Close
        
        'vRepType = "INV"
        
        If vCheckItemNotLocation015 > 0 Then
          Call PrintItem010_Zone010
        End If
        If vCheckItemLocation015 > 0 Then
          Call PrintItem010_Zone015
        End If

'vCheck = 1
'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
'If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    'With CrystalReport341
       ' .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
      '  .ParameterFields(0) = "@DocNo;" & vDocno & ";true"
     '   .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
    '    .Destination = crptToWindow
   '     .WindowState = crptMaximized
  '      .Action = 1
 '   End With
'End If
'vRecordset.Close
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem010_Zone010()
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
        Dim vGenerateNumber, vHeader As String
        Dim vCheck As Integer
      
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT341.Text)
        vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
            vHeader = Trim(vRecordset.Fields("header").Value)
        End If
        vRecordset.Close
        vGenerateNumber = vHeader & "-" & vAutoNumber
        
        vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
        gConnection.Execute vQuery
        vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked) " _
        & " values ('" & vDocNo & "','" & vGenerateNumber & "',getdate(),'010','" & vUserID & "',1,0)"
        gConnection.Execute vQuery

        vDocno1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
        vDocGroup1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
                
        vRepType = "INV"
        
        If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
          vRepID = 345
        Else
          vRepID = 341
        End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With CrystalReport341
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
        .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem010_Zone015()
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
        Dim vGenerateNumber, vHeader As String
        Dim vCheck As Integer
      
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT341.Text)
        
        vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
            vHeader = Trim(vRecordset.Fields("header").Value)
        End If
        vRecordset.Close
        vGenerateNumber = vHeader & "-" & vAutoNumber
        
        vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1 where headertype = 8 "
        gConnection.Execute vQuery
        vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked) " _
        & " values ('" & vDocNo & "','" & vGenerateNumber & "',getdate(),'010C','" & vUserID & "',1,0)"
        gConnection.Execute vQuery
        
        vDocGroup1 = UCase(Left(Right(vDocNo, Len(vDocNo) - InStr(vDocNo, "-")), 3)) 'UCase(Left(vDocNo, 3))
        vRepType = "INV"
        
        If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
          vRepID = 347
        Else
          vRepID = 343
        End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With CrystalReport341
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
        .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem012()
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
        Dim vCheck As Integer
        Dim vGenerateNumber, vHeader As String
      
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT341.Text)
        vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
            vHeader = Trim(vRecordset.Fields("header").Value)
        End If
        vRecordset.Close
        vGenerateNumber = vHeader & "-" & vAutoNumber
        
        vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 8 "
        gConnection.Execute vQuery
        vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked) " _
        & " values ('" & vDocNo & "','" & vGenerateNumber & "',getdate(),'012','" & vUserID & "',1,0)"
        gConnection.Execute vQuery
        
        vDocno1 = Left(vDocNo, 3)
        vQuery = "select upper('" & vDocno1 & "') as vDocGroup"
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vDocGroup1 = UCase(Trim(vRecordset.Fields("vDocGroup").Value))
        End If
        vRecordset.Close
        
        If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
        vRepID = 129
        vRepType = "INV"
        Else
        vRepID = 91
        vRepType = "INV"
        End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With CrystalReport341
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
        .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem015()
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
        Dim vCheck As Integer
        Dim vGenerateNumber, vHeader As String
      
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT341.Text)
        vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
            vHeader = Trim(vRecordset.Fields("header").Value)
        End If
        vRecordset.Close
        vGenerateNumber = vHeader & "-" & vAutoNumber
        
        vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 8 "
        gConnection.Execute vQuery
        vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked) " _
        & " values ('" & vDocNo & "','" & vGenerateNumber & "',getdate(),'015','" & vUserID & "',1,0)"
        gConnection.Execute vQuery
        
        vDocno1 = Left(vDocNo, 3)
        vQuery = "select upper('" & vDocno1 & "') as vDocGroup"
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vDocGroup1 = UCase(Trim(vRecordset.Fields("vDocGroup").Value))
        End If
        vRecordset.Close
        
        If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
        vRepID = 131
        vRepType = "INV"
        Else
        vRepID = 92
        vRepType = "INV"
        End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With CrystalReport341
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
        .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem097()
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
        Dim vCheck As Integer
        Dim vGenerateNumber, vHeader As String
      
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT341.Text)
        vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
            vHeader = Trim(vRecordset.Fields("header").Value)
        End If
        vRecordset.Close
        vGenerateNumber = vHeader & "-" & vAutoNumber
        
        vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 8 "
        gConnection.Execute vQuery
        vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked) " _
        & " values ('" & vDocNo & "','" & vGenerateNumber & "',getdate(),'097','" & vUserID & "',1,0)"
        gConnection.Execute vQuery
        
        vDocno1 = Left(vDocNo, 3)
        vQuery = "select upper('" & vDocno1 & "') as vDocGroup"
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vDocGroup1 = UCase(Trim(vRecordset.Fields("vDocGroup").Value))
        End If
        vRecordset.Close
        
        If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
        vRepID = 133
        vRepType = "INV"
        Else
        vRepID = 97
        vRepType = "INV"
        End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With CrystalReport341
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
        .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub


Public Sub PrintItem014()
Dim vReportName As String
Dim vDocNo As String
Dim vRepID As Integer
Dim vRepType As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
Dim vCheck As Integer
Dim vGenerateNumber, vHeader As String

On Error GoTo ErrDescription

vDocNo = Trim(TXT341.Text)
vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
    vHeader = Trim(vRecordset.Fields("header").Value)
End If
vRecordset.Close
vGenerateNumber = vHeader & "-" & vAutoNumber

vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 8 "
gConnection.Execute vQuery
vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked) " _
& " values ('" & vDocNo & "','" & vGenerateNumber & "',getdate(),'014','" & vUserID & "',1,0)"
gConnection.Execute vQuery

vDocno1 = Left(vDocNo, 3)
vQuery = "select upper('" & vDocno1 & "') as vDocGroup"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vDocGroup1 = UCase(Trim(vRecordset.Fields("vDocGroup").Value))
End If
vRecordset.Close

If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 170
vRepType = "INV"
Else
vRepID = 169
vRepType = "INV"
End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With CrystalReport341
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
        .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem020()
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
        Dim vGenerateNumber, vHeader As String
        Dim vCheck As Integer
      
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT341.Text)
        vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
            vHeader = Trim(vRecordset.Fields("header").Value)
        End If
        vRecordset.Close
        vGenerateNumber = vHeader & "-" & vAutoNumber
        
        vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 8 "
        gConnection.Execute vQuery
        vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked) " _
        & " values ('" & vDocNo & "','" & vGenerateNumber & "',getdate(),'020','" & vUserID & "',1,0)"
        gConnection.Execute vQuery
        
        vDocno1 = Left(vDocNo, 3)
        vQuery = "select upper('" & vDocno1 & "') as vDocGroup"
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vDocGroup1 = UCase(Trim(vRecordset.Fields("vDocGroup").Value))
        End If
        vRecordset.Close
        
        If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
        vRepID = 194
        vRepType = "INV"
        Else
        vRepID = 193
        vRepType = "INV"
        End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
With CrystalReport341
    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
    .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
    .Destination = crptToWindow
    .WindowState = crptMaximized
    .Action = 1
End With
End If
vRecordset.Close
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem016()
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vCheck As Integer
        Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
        Dim vGenerateNumber, vHeader As String
      
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT341.Text)
        vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
            vHeader = Trim(vRecordset.Fields("header").Value)
        End If
        vRecordset.Close
        vGenerateNumber = vHeader & "-" & vAutoNumber
        
        vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 8 "
        gConnection.Execute vQuery
        vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked) " _
        & " values ('" & vDocNo & "','" & vGenerateNumber & "',getdate(),'016','" & vUserID & "',1,0)"
        gConnection.Execute vQuery
        
        vDocno1 = Left(vDocNo, 3)
        vQuery = "select upper('" & vDocno1 & "') as vDocGroup"
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vDocGroup1 = UCase(Trim(vRecordset.Fields("vDocGroup").Value))
        End If
        vRecordset.Close
        
        If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
        vRepID = 217
        vRepType = "INV"
        Else
        vRepID = 216
        vRepType = "INV"
        End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With CrystalReport341
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
        .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub

Public Sub PrintItem070()
        Dim vReportName As String
        Dim vDocNo As String
        Dim vRepID As Integer
        Dim vRepType As String
        Dim vQuery As String
        Dim vRecordset As New ADODB.Recordset
        Dim vAutoNumber As String, vDocno1 As String, vDocGroup1 As String
        Dim vGenerateNumber, vHeader As String
        Dim vCheck As Integer
        
        On Error GoTo ErrDescription
        
        vDocNo = Trim(TXT341.Text)
        vQuery = "select header,autonumber  from npmaster.dbo.NP_Generate_DocNo where headertype = 8 "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vAutoNumber = Format(Trim(vRecordset.Fields("autonumber").Value), "0000")
            vHeader = Trim(vRecordset.Fields("header").Value)
        End If
        vRecordset.Close
        vGenerateNumber = vHeader & "-" & vAutoNumber
        
        vQuery = "Update npmaster.dbo.NP_Generate_DocNo  set autonumber = autonumber + 1  where headertype = 8 "
        gConnection.Execute vQuery
        vQuery = "insert npmaster.dbo.np_paygoods (invoiceno,paynumber,paydatetime,whcode,UserPrint,LastPrintCount,Checked) " _
        & " values ('" & vDocNo & "','" & vGenerateNumber & "',getdate(),'070','" & vUserID & "',1,0)"
        gConnection.Execute vQuery
        
        vDocno1 = Left(vDocNo, 3)
        vQuery = "select upper('" & vDocno1 & "') as vDocGroup"
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vDocGroup1 = UCase(Trim(vRecordset.Fields("vDocGroup").Value))
        End If
        vRecordset.Close
        
        If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
        vRepID = 298
        vRepType = "INV"
        Else
        vRepID = 296
        vRepType = "INV"
        End If
vCheck = 1

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    With CrystalReport341
        .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
        .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
        .ParameterFields(1) = "@vCheck;" & vCheck & ";true"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
vRecordset.Close
                
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Sub


