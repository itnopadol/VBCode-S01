VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPayItemCust 
   BackColor       =   &H80000009&
   Caption         =   "ตรวจสอบการจ่ายสินค้ากับลูกค้า"
   ClientHeight    =   8070
   ClientLeft      =   5070
   ClientTop       =   3075
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8070
   ScaleWidth      =   12090
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMDPicking 
      Caption         =   "หน้าใบจัดสินค้า"
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
      Left            =   7380
      TabIndex        =   14
      Top             =   45
      Width           =   1365
   End
   Begin VB.Timer Timer103 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   10260
      Top             =   7515
   End
   Begin VB.Timer Timer102 
      Interval        =   1000
      Left            =   9810
      Top             =   7515
   End
   Begin VB.Timer Timer101 
      Interval        =   12000
      Left            =   9405
      Top             =   7515
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   5370
      Left            =   5940
      ScaleHeight     =   5310
      ScaleWidth      =   6075
      TabIndex        =   6
      Top             =   585
      Width           =   6135
      Begin MSComctlLib.ListView ListView102 
         Height          =   3210
         Left            =   180
         TabIndex        =   4
         Top             =   1485
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   5662
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "เลขที่ใบเหลือง"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "เลขที่บิล"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "คลัง"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ชื่อลูกค้า"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "จำนวนครั้งพิมพ์"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.TextBox TextChecker 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1170
         TabIndex        =   3
         Top             =   720
         Width           =   1950
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "การตรวจสอบการจ่ายสินค้า"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   465
         Left            =   315
         TabIndex        =   10
         Top             =   90
         Width           =   5370
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่ใบจ่าย :"
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   180
         TabIndex        =   8
         Top             =   765
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000D&
      Height          =   8160
      Left            =   0
      ScaleHeight     =   8100
      ScaleWidth      =   5850
      TabIndex        =   5
      Top             =   -45
      Width           =   5910
      Begin VB.CommandButton CMDPrint 
         Caption         =   "พิมพ์"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4275
         TabIndex        =   1
         Top             =   900
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSComctlLib.ListView ListView101 
         Height          =   5730
         Left            =   90
         TabIndex        =   2
         Top             =   1485
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   10107
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "เลขที่ใบเหลือง"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "เลขที่บิล"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "คลัง"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ชื่อลูกค้า"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ครั้งที่พิมพ์"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.TextBox Text101 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1170
         TabIndex        =   0
         Top             =   720
         Width           =   2040
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "การพิมพ์ใบเหลือง"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   420
         Left            =   180
         TabIndex        =   9
         Top             =   135
         Width           =   5370
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่ใบจ่าย :"
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   180
         TabIndex        =   7
         Top             =   765
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport Crystal101 
      Left            =   11610
      Top             =   7560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Image Image103 
      Height          =   300
      Left            =   6885
      Picture         =   "FrmPayItemCust.frx":0000
      Top             =   90
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Image102 
      Height          =   300
      Left            =   6480
      Picture         =   "FrmPayItemCust.frx":23C1
      Top             =   90
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Image101 
      Height          =   300
      Left            =   6075
      Picture         =   "FrmPayItemCust.frx":4782
      Top             =   90
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label LBLUserPick 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   13
      Top             =   7020
      Width           =   5730
   End
   Begin VB.Label LBLRefNo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   6525
      Width           =   5730
   End
   Begin VB.Label LBLARName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   11
      Top             =   6075
      Width           =   5730
   End
End
Attribute VB_Name = "FrmPayItemCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vQuery As String

Private Sub CMDPicking_Click()
Call StopTime
Load FrmQueue
Unload FrmPayItemCust
End Sub

Private Sub CMDPrint_Click()
Dim vDocno As String
Dim vCheckDocNo As String
Dim i As Integer
Dim vWHCode As String
Dim vRefNo As String
Dim vRecordset As New ADODB.Recordset
Dim vCheckExist As Integer
Dim vIsCancel As Integer

vDocno = UCase(Trim(Text101.Text))
vCheckExist = CheckItemReceipt(vDocno)
Call StopTime
If vCheckExist = 0 Then
  MsgBox "ไม่มีเลขที่บิล " & vDocno & " นี้ ไม่สามารถพิมพ์ใบจ่ายสินค้า(ใบเหลือง)ได้ ", vbCritical, "Send Error"
  Exit Sub
  Call StartTime
End If

vQuery = "exec dbo.USP_QUE_CheckIsCancelInvoice '" & vDocno & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
  vIsCancel = Trim(vRecordset.Fields("iscancel").Value)
End If
vRecordset.Close

If vIsCancel = 1 Then
  vQuery = "exec dbo.USP_QUE_UpdateCancelCustItemReceipt '" & vDocno & "' "
  vConnection.Execute vQuery
  MsgBox "ไม่สามารถพิมพ์ใบจ่ายสินค้า(ใบเหลือง)ได้ เนื่องจากบิลได้ถูกยกเลิกไปแล้ว", vbCritical, "ข้อความแจ้งเตือน"
  Call StartTime
  Exit Sub
End If

If vDocno <> "" And vCheckExist = 1 Then
  For i = 1 To ListView101.ListItems.Count
  vWHCode = Trim(ListView101.ListItems.Item(i).SubItems(2))
  vCheckDocNo = Trim(ListView101.ListItems.Item(i).SubItems(1))
  vRefNo = Trim(ListView101.ListItems.Item(i).Text)
  If vDocno = vCheckDocNo Then
    Select Case vWHCode
    Case "010"
    Call Print010(vDocno, vWHCode, vRefNo)
    Case "012"
    Call Print012(vDocno, vWHCode, vRefNo)
    Case "014"
    Call Print014(vDocno, vWHCode, vRefNo)
    Case "015"
    Call Print015(vDocno, vWHCode, vRefNo)
    Case "016"
    Call Print016(vDocno, vWHCode, vRefNo)
    Case "020"
    Call Print020(vDocno, vWHCode, vRefNo)
    Case "070"
    Call Print070(vDocno, vWHCode, vRefNo)
    Case "097"
    Call Print097(vDocno, vWHCode, vRefNo)
    End Select
  End If
  Next i
Else
  MsgBox "ไม่ได้กรอกเลขที่เอกสารที่จะพิมพ์ใบจ่ายสินค้า", vbCritical, "Send Error"
  Text101.SetFocus
End If
Call StartTime
Call SearchCustItemReceiptChecking
Text101.Text = ""
End Sub

Function CheckItemReceipt(vSearch As String) As Integer
Dim i As Integer
Dim vCheckExist As String
Dim vInvoiceNo As String

vInvoiceNo = UCase(vSearch)
If vInvoiceNo <> "" Then
  For i = 1 To ListView101.ListItems.Count
    vCheckExist = ListView101.ListItems.Item(i).SubItems(1)
    If vInvoiceNo = vCheckExist Then
    CheckItemReceipt = 1
    Exit Function
    Else
    CheckItemReceipt = 0
    End If
  Next i
End If
End Function

Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset

SearchCustItemReceiptBegin
SearchCustItemReceiptChecking
End Sub

Private Sub ListView101_ItemClick(ByVal Item As MSComctlLib.ListItem)
  Dim vARName As String
  Dim vRefNo As String
  Dim vPicker As String
  Dim vIndex As Integer
  
  Text101.SetFocus
  If ListView101.ListItems.Count > 0 Then
    vIndex = ListView101.SelectedItem.Index
    vARName = Trim(ListView101.ListItems.Item(vIndex).SubItems(3))
    vRefNo = Trim(ListView101.ListItems.Item(vIndex).SubItems(1))
    vPicker = "-"
    FrmPayItemCust.LBLARName.Caption = vARName
    FrmPayItemCust.LBLRefNo.Caption = vRefNo
    FrmPayItemCust.LBLUserPick.Caption = Trim("-")
  End If
End Sub


Private Sub ListView102_ItemCheck(ByVal Item As MSComctlLib.ListItem)
vQuery = "usp_np_CheckInvoicePicking"
End Sub


Private Sub ListView102_ItemClick(ByVal Item As MSComctlLib.ListItem)
  Dim vARName As String
  Dim vRefNo As String
  Dim vPicker As String
  Dim vIndex As Integer
  
  TextChecker.SetFocus
  If ListView102.ListItems.Count > 0 Then
    vIndex = ListView102.SelectedItem.Index
    vARName = Trim(ListView102.ListItems.Item(vIndex).SubItems(3))
    vRefNo = Trim(ListView102.ListItems.Item(vIndex).SubItems(1))
    vPicker = "-"
    FrmPayItemCust.LBLARName.Caption = vARName
    FrmPayItemCust.LBLRefNo.Caption = vRefNo
    FrmPayItemCust.LBLUserPick.Caption = Trim("-")
  End If
End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)
Dim vDocno As String
Dim vCheckDocNo As String
Dim i As Integer
Dim vExistDoc As Integer

If KeyAscii = 13 Then
  Call StopTime
  vDocno = UCase(Trim(Text101.Text))
  If vDocno <> "" Then
    For i = 1 To ListView101.ListItems.Count
      vCheckDocNo = UCase(Trim(ListView101.ListItems.Item(i).SubItems(1)))
      If vDocno = vCheckDocNo Then
      CMDPrint.Visible = True
      vExistDoc = 1
      Exit Sub
      Else
      vExistDoc = 0
      End If
    Next i
    If vExistDoc = 0 Then
      MsgBox "ไม่มีเลขที่บิล " & vDocno & " นี้ ไม่สามารถพิมพ์ใบจ่ายสินค้า(ใบเหลือง)ได้ ", vbCritical, "Send Error"
    End If
  Else
    MsgBox "ไม่ได้กรอกเลขที่เอกสารที่จะพิมพ์ใบจ่ายสินค้า", vbCritical, "Send Error"
    Text101.SetFocus
  End If
Call StartTime
End If
End Sub


Public Sub Print010(vInputDocNo As String, vInputWHCode As String, vInputRefNo As String)
Dim vReportName As String
Dim vDocno As String
Dim vRepID As Integer
Dim vRepType As String
Dim vRecordset As New ADODB.Recordset
Dim vDocGroup1 As String
Dim vGenerateNumber
Dim vCheck As Integer
Dim vWHCode As String


vDocno = Trim(vInputDocNo)
vGenerateNumber = vInputRefNo
vWHCode = vInputWHCode

'vQuery = "exec dbo.USP_NP_InsertPayGoods '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vUserID & "')"
'vConnection.Execute vQuery

vDocGroup1 = UCase(Left(vDocno, 3))

If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 125
vRepType = "INV"
Else
vRepID = 89
vRepType = "INV"
End If
vCheck = 1
vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
With Crystal101
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@DocNo;" & vDocno & ";true"
.ParameterFields(1) = "@vCheck;" & vCheck & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close

vQuery = "exec dbo.USP_NP_UpdateStatusCustItemReceipt 0,'" & vInputRefNo & "',1,'' "
vConnection.Execute vQuery

End Sub

Public Sub Print012(vInputDocNo As String, vInputWHCode As String, vInputRefNo As String)
Dim vReportName As String
Dim vDocno As String
Dim vRepID As Integer
Dim vRepType As String
Dim vRecordset As New ADODB.Recordset
Dim vDocGroup1 As String
Dim vGenerateNumber
Dim vCheck As Integer
Dim vWHCode As String


vDocno = vInputDocNo
vGenerateNumber = vInputRefNo
vWHCode = vInputWHCode

'vQuery = "exec dbo.USP_NP_InsertPayGoods '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vUserID & "')"
'vConnection.Execute vQuery

vDocGroup1 = UCase(Left(vDocno, 3))

If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 129
vRepType = "INV"
Else
vRepID = 91
vRepType = "INV"
End If
vCheck = 1
vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
With Crystal101
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@DocNo;" & vDocno & ";true"
.ParameterFields(1) = "@vCheck;" & vCheck & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close

vQuery = "exec dbo.USP_NP_UpdateStatusCustItemReceipt 0,'" & vInputRefNo & "',1,'' "
vConnection.Execute vQuery

End Sub

Public Sub Print014(vInputDocNo As String, vInputWHCode As String, vInputRefNo As String)
Dim vReportName As String
Dim vDocno As String
Dim vRepID As Integer
Dim vRepType As String
Dim vRecordset As New ADODB.Recordset
Dim vDocGroup1 As String
Dim vGenerateNumber
Dim vCheck As Integer
Dim vWHCode As String


vDocno = vInputDocNo
vGenerateNumber = vInputRefNo
vWHCode = vInputWHCode

'vQuery = "exec dbo.USP_NP_InsertPayGoods '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vUserID & "')"
'vConnection.Execute vQuery

vDocGroup1 = UCase(Left(vDocno, 3))

If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 170
vRepType = "INV"
Else
vRepID = 69
vRepType = "INV"
End If
vCheck = 1
vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
With Crystal101
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@DocNo;" & vDocno & ";true"
.ParameterFields(1) = "@vCheck;" & vCheck & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close

vQuery = "exec dbo.USP_NP_UpdateStatusCustItemReceipt 0,'" & vInputRefNo & "',1,'' "
vConnection.Execute vQuery
End Sub

Public Sub Print015(vInputDocNo As String, vInputWHCode As String, vInputRefNo As String)
Dim vReportName As String
Dim vDocno As String
Dim vRepID As Integer
Dim vRepType As String
Dim vRecordset As New ADODB.Recordset
Dim vDocGroup1 As String
Dim vGenerateNumber
Dim vCheck As Integer
Dim vWHCode As String


vDocno = vInputDocNo
vGenerateNumber = vInputRefNo
vWHCode = vInputWHCode

'vQuery = "exec dbo.USP_NP_InsertPayGoods '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vUserID & "')"
'vConnection.Execute vQuery

vDocGroup1 = UCase(Left(vDocno, 3))

If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 131
vRepType = "INV"
Else
vRepID = 92
vRepType = "INV"
End If
vCheck = 1
vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
With Crystal101
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@DocNo;" & vDocno & ";true"
.ParameterFields(1) = "@vCheck;" & vCheck & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close

vQuery = "exec dbo.USP_NP_UpdateStatusCustItemReceipt 0,'" & vInputRefNo & "',1,'' "
vConnection.Execute vQuery
End Sub


Public Sub Print016(vInputDocNo As String, vInputWHCode As String, vInputRefNo As String)
Dim vReportName As String
Dim vDocno As String
Dim vRepID As Integer
Dim vRepType As String
Dim vRecordset As New ADODB.Recordset
Dim vDocGroup1 As String
Dim vGenerateNumber
Dim vCheck As Integer
Dim vWHCode As String


vDocno = vInputDocNo
vGenerateNumber = vInputRefNo
vWHCode = vInputWHCode

'vQuery = "exec dbo.USP_NP_InsertPayGoods '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vUserID & "')"
'vConnection.Execute vQuery

vDocGroup1 = UCase(Left(vDocno, 3))

If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 217
vRepType = "INV"
Else
vRepID = 216
vRepType = "INV"
End If
vCheck = 1
vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
With Crystal101
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@DocNo;" & vDocno & ";true"
.ParameterFields(1) = "@vCheck;" & vCheck & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close

vQuery = "exec dbo.USP_NP_UpdateStatusCustItemReceipt 0,'" & vInputRefNo & "',1,'' "
vConnection.Execute vQuery
End Sub

Public Sub Print020(vInputDocNo As String, vInputWHCode As String, vInputRefNo As String)
Dim vReportName As String
Dim vDocno As String
Dim vRepID As Integer
Dim vRepType As String
Dim vRecordset As New ADODB.Recordset
Dim vDocGroup1 As String
Dim vGenerateNumber
Dim vCheck As Integer
Dim vWHCode As String


vDocno = vInputDocNo
vGenerateNumber = vInputRefNo
vWHCode = vInputWHCode

'vQuery = "exec dbo.USP_NP_InsertPayGoods '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vUserID & "')"
'vConnection.Execute vQuery

vDocGroup1 = UCase(Left(vDocno, 3))

If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 194
vRepType = "INV"
Else
vRepID = 193
vRepType = "INV"
End If
vCheck = 1
vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
With Crystal101
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@DocNo;" & vDocno & ";true"
.ParameterFields(1) = "@vCheck;" & vCheck & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close

vQuery = "exec dbo.USP_NP_UpdateStatusCustItemReceipt 0,'" & vInputRefNo & "',1,'' "
vConnection.Execute vQuery
End Sub

Public Sub Print070(vInputDocNo As String, vInputWHCode As String, vInputRefNo As String)
Dim vReportName As String
Dim vDocno As String
Dim vRepID As Integer
Dim vRepType As String
Dim vRecordset As New ADODB.Recordset
Dim vDocGroup1 As String
Dim vGenerateNumber
Dim vCheck As Integer
Dim vWHCode As String


vDocno = vInputDocNo
vGenerateNumber = vInputRefNo
vWHCode = vInputWHCode

'vQuery = "exec dbo.USP_NP_InsertPayGoods '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vUserID & "')"
'vConnection.Execute vQuery

vDocGroup1 = UCase(Left(vDocno, 3))

If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 298
vRepType = "INV"
Else
vRepID = 296
vRepType = "INV"
End If
vCheck = 1
vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
With Crystal101
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@DocNo;" & vDocno & ";true"
.ParameterFields(1) = "@vCheck;" & vCheck & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close

vQuery = "exec dbo.USP_NP_UpdateStatusCustItemReceipt 0,'" & vInputRefNo & "',1,'' "
vConnection.Execute vQuery
End Sub


Public Sub Print097(vInputDocNo As String, vInputWHCode As String, vInputRefNo As String)
Dim vReportName As String
Dim vDocno As String
Dim vRepID As Integer
Dim vRepType As String
Dim vRecordset As New ADODB.Recordset
Dim vDocGroup1 As String
Dim vGenerateNumber
Dim vCheck As Integer
Dim vWHCode As String


vDocno = vInputDocNo
vGenerateNumber = vInputRefNo
vWHCode = vInputWHCode

'vQuery = "exec dbo.USP_NP_InsertPayGoods '" & vDocNo & "','" & vGenerateNumber & "','" & vWHCode & "','" & vUserID & "')"
'vConnection.Execute vQuery

vDocGroup1 = UCase(Left(vDocno, 3))

If vDocGroup1 = "ICV" Or vDocGroup1 = "ICN" Or vDocGroup1 = "IVD" Or vDocGroup1 = "IVN" Or vDocGroup1 = "IVM" Then
vRepID = 133
vRepType = "INV"
Else
vRepID = 97
vRepType = "INV"
End If
vCheck = 1
vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
With Crystal101
.ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
.ParameterFields(0) = "@DocNo;" & vDocno & ";true"
.ParameterFields(1) = "@vCheck;" & vCheck & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End If
vRecordset.Close

vQuery = "exec dbo.USP_NP_UpdateStatusCustItemReceipt 0,'" & vInputRefNo & "',1,'' "
vConnection.Execute vQuery
End Sub

Private Sub TextChecker_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vRefNo As String
Dim vInvoiceNo As String
Dim vWHCode As String
Dim vCountPrint As Integer
Dim vListItem As ListItem
Dim n As Integer
Dim i As Integer
Dim j As Integer
Dim vCheckQueue As String
Dim vIndex As Integer

If KeyAscii = 13 Then
  Call StopTime
  n = ListView102.ListItems.Count
  vRefNo = Trim(TextChecker.Text)
  For j = 1 To n
    vCheckQueue = Trim(ListView102.ListItems.Item(j).Text)
    If vRefNo = vCheckQueue Then
      vIndex = j
    End If
  Next j


  If vIndex <> 0 Then
    vRefNoReceive = Trim(ListView102.ListItems.Item(vIndex).Text)
    vWHCodeReceive = Trim(ListView102.ListItems.Item(vIndex).SubItems(2))
    vInvoiceNoReceive = Trim(ListView102.ListItems.Item(vIndex).SubItems(1))
    vCountPrint = Trim(ListView102.ListItems.Item(vIndex).SubItems(4))
    FormCustReceiptItem.Show
    FormCustReceiptItem.LBLRefNo.Caption = vRefNoReceive
    FormCustReceiptItem.LBLInvoice.Caption = vInvoiceNoReceive
    FormCustReceiptItem.LBLWHCode.Caption = vWHCodeReceive
    FormCustReceiptItem.LBLCount.Caption = vCountPrint
    
    If vRefNoReceive <> "" And vInvoiceNoReceive <> "" Then
    i = 0
    FormCustReceiptItem.ListView101.ListItems.Clear
    vQuery = "exec dbo.USP_QUE_SearchWHCodeCustReceiptItem '" & vRefNoReceive & "','" & vWHCodeReceive & "' "
    If OpenDataBase2(qConnection, vRecordset, vQuery) <> 0 Then
      vRecordset.MoveFirst
      While Not vRecordset.EOF
        i = i + 1
        Set vListItem = FormCustReceiptItem.ListView101.ListItems.Add(, , i)
          vListItem.SubItems(1) = Trim(vRecordset.Fields("itemcode").Value)
          vListItem.SubItems(2) = Trim(vRecordset.Fields("itemname").Value)
          vListItem.SubItems(3) = Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00")
          vListItem.SubItems(4) = Format(Trim(vRecordset.Fields("qty").Value), "##,##0.00")
          vListItem.SubItems(5) = Trim(vRecordset.Fields("unitcode").Value)
          vListItem.SubItems(6) = Trim(vRecordset.Fields("shelfcode").Value)
      vRecordset.MoveNext
      Wend
    End If
    vRecordset.Close
    End If
  Else
    MsgBox "กรอกเลขที่คิวไม่ถูกต้อง กรุณาตรวจสอบ", vbCritical, "Send Error"
  End If
End If
End Sub

Private Sub Timer103_Timer()
If Image101.Visible = True Then
  Image101.Visible = False
  Image102.Visible = False
  Image103.Visible = False
Else
  Image101.Visible = True
  Image102.Visible = True
  Image103.Visible = True
End If
End Sub

Private Sub Timer101_Timer()
Dim vRecordset As New ADODB.Recordset
Dim vListItem As ListItem
Dim vCount As Integer
Dim vPrinted As Integer
Dim vARName  As String
Dim vRefNo As String
Dim vPicker As String
Dim vSaleCode As String

On Error Resume Next

ListView101.ListItems.Clear
vQuery = "exec dbo.USP_QUE_SearchCustItemReceipt " & vSelectZoneID & " "
If OpenDataBase1(sConnection, vRecordset, vQuery) <> 0 Then
  If vCheckClickListview = 1 Then
    vARName = Trim(vRecordset.Fields("arname").Value)
    vRefNo = Trim(vRecordset.Fields("invoiceno").Value)
    vPicker = Trim(vRecordset.Fields("checker").Value)
    FrmPayItemCust.LBLARName.Caption = vARName
    FrmPayItemCust.LBLRefNo.Caption = vRefNo
    FrmPayItemCust.LBLUserPick.Caption = Trim("-")
  End If
    While Not vRecordset.EOF
        Set vListItem = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("docno").Value))
        vListItem.SubItems(1) = Trim(vRecordset.Fields("invoiceno").Value)
        vListItem.SubItems(2) = Trim(vRecordset.Fields("whcode").Value)
        vListItem.SubItems(3) = Trim(vRecordset.Fields("arname").Value)
        vListItem.SubItems(4) = Trim(vRecordset.Fields("printcount").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close
End Sub

Private Sub Timer102_Timer()
Dim vRecordset As New ADODB.Recordset
Dim vCount As Integer

vQuery = "select count(docno) as vCount from npmaster.dbo.TB_NP_QueueManagement where status = 1 and  year(docdate) = year(getdate()) and month(docdate) = month(getdate())  and day(docdate) = day(getdate()) and zoneid in ('02','03') "
If OpenDataBase(vConnection, vRecordset, vQuery) <> 0 Then
  vCount = Trim(vRecordset.Fields("vcount").Value)
End If
vRecordset.Close

If vCount > 0 Then
  Timer103.Enabled = True
End If
End Sub

Public Sub SearchCustItemReceiptBegin()
Dim vRecordset As New ADODB.Recordset
Dim vListItem As ListItem
Dim vCount As Integer
Dim vPrinted As Integer
Dim vARName  As String
Dim vRefNo As String
Dim vPicker As String
Dim vSaleCode As String

On Error Resume Next

ListView101.ListItems.Clear
vQuery = "exec dbo.USP_QUE_SearchCustItemReceipt " & vSelectZoneID & " "
If OpenDataBase1(sConnection, vRecordset, vQuery) <> 0 Then
  If vCheckClickListview = 1 Then
    vARName = Trim(vRecordset.Fields("arname").Value)
    vRefNo = Trim(vRecordset.Fields("invoiceno").Value)
    vPicker = Trim(vRecordset.Fields("checker").Value)
    FrmPayItemCust.LBLARName.Caption = vARName
    FrmPayItemCust.LBLRefNo.Caption = vRefNo
    FrmPayItemCust.LBLUserPick.Caption = Trim("-")
  End If
    While Not vRecordset.EOF
        Set vListItem = ListView101.ListItems.Add(, , Trim(vRecordset.Fields("docno").Value))
        vListItem.SubItems(1) = Trim(vRecordset.Fields("invoiceno").Value)
        vListItem.SubItems(2) = Trim(vRecordset.Fields("whcode").Value)
        vListItem.SubItems(3) = Trim(vRecordset.Fields("arname").Value)
        vListItem.SubItems(4) = Trim(vRecordset.Fields("printcount").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close
End Sub

Public Sub SearchCustItemReceiptChecking()
Dim vRecordset As New ADODB.Recordset
Dim vListItem As ListItem
Dim vCount As Integer
Dim vPrinted As Integer
Dim vARName  As String
Dim vRefNo As String
Dim vPicker As String
Dim vSaleCode As String

On Error Resume Next

ListView102.ListItems.Clear
vQuery = "exec dbo.USP_QUE_SearchCustItemReceiptChecking " & vSelectZoneID & " "
If OpenDataBase1(sConnection, vRecordset, vQuery) <> 0 Then
  If vCheckClickListview = 1 Then
    vARName = Trim(vRecordset.Fields("arname").Value)
    vRefNo = Trim(vRecordset.Fields("invoiceno").Value)
    vPicker = Trim(vRecordset.Fields("checker").Value)
    FrmPayItemCust.LBLARName.Caption = vARName
    FrmPayItemCust.LBLRefNo.Caption = vRefNo
    FrmPayItemCust.LBLUserPick.Caption = Trim("-")
  End If
    While Not vRecordset.EOF
        Set vListItem = ListView102.ListItems.Add(, , Trim(vRecordset.Fields("docno").Value))
        vListItem.SubItems(1) = Trim(vRecordset.Fields("invoiceno").Value)
        vListItem.SubItems(2) = Trim(vRecordset.Fields("whcode").Value)
        vListItem.SubItems(3) = Trim(vRecordset.Fields("arname").Value)
        vListItem.SubItems(4) = Trim(vRecordset.Fields("printcount").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close
End Sub


Public Sub StartTime()
Timer101.Enabled = True
Timer102.Enabled = True
Timer103.Enabled = True
End Sub


Public Sub StopTime()
Timer101.Enabled = False
Timer102.Enabled = False
Timer103.Enabled = False
End Sub

