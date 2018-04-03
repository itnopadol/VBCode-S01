VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form87 
   Caption         =   "หน้าบันทึกเวลาจัดสินค้า"
   ClientHeight    =   8385
   ClientLeft      =   2550
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form87.frx":0000
   ScaleHeight     =   8385
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport871 
      Left            =   630
      Top             =   7380
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
   Begin Crystal.CrystalReport CrystalReport872 
      Left            =   135
      Top             =   7380
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
      Height          =   315
      Left            =   2430
      TabIndex        =   5
      Top             =   5265
      Width           =   1935
   End
   Begin VB.CommandButton CMDGraph 
      Caption         =   "กราฟ"
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
      Left            =   3555
      TabIndex        =   8
      Top             =   6660
      Width           =   810
   End
   Begin MSComCtl2.DTPicker DTP1 
      Height          =   315
      Left            =   2430
      TabIndex        =   6
      Top             =   5805
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   67108865
      CurrentDate     =   38208
   End
   Begin VB.TextBox TXT876 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1755
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   4275
      Width           =   8355
   End
   Begin VB.CommandButton CMDChangePicker 
      Caption         =   "บันทึกคนจัด"
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
      Height          =   315
      Left            =   8910
      TabIndex        =   9
      Top             =   3825
      Width           =   1155
   End
   Begin VB.ComboBox CMBPicker 
      Height          =   315
      Left            =   6570
      TabIndex        =   2
      Top             =   3285
      Width           =   3480
   End
   Begin VB.CommandButton CMDReport 
      Caption         =   "รายงาน"
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
      Left            =   3555
      TabIndex        =   7
      Top             =   6255
      Width           =   810
   End
   Begin VB.TextBox TXTStop 
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
      Height          =   300
      Left            =   2430
      TabIndex        =   16
      Top             =   3285
      Width           =   2355
   End
   Begin VB.TextBox TXTStart 
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
      Left            =   6570
      TabIndex        =   15
      Top             =   1755
      Width           =   2295
   End
   Begin VB.TextBox TXTPickingDate 
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
      Left            =   6570
      TabIndex        =   13
      Top             =   1260
      Width           =   2295
   End
   Begin VB.CommandButton CMDFinish 
      Caption         =   "ปุ่มจัดเสร็จ"
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
      Height          =   315
      Left            =   3690
      TabIndex        =   3
      Top             =   3780
      Width           =   1065
   End
   Begin VB.CommandButton CMDPrint 
      Caption         =   "ปุ่มเริ่มจัด"
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
      Height          =   315
      Left            =   3735
      TabIndex        =   1
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox TXTDocno1 
      Appearance      =   0  'Flat
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
      Left            =   2430
      TabIndex        =   0
      Top             =   1260
      Width           =   2295
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000004&
      X1              =   405
      X2              =   10305
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000004&
      X1              =   10305
      X2              =   10305
      Y1              =   3105
      Y2              =   5040
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000004&
      X1              =   405
      X2              =   405
      Y1              =   3105
      Y2              =   5040
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000004&
      X1              =   405
      X2              =   10305
      Y1              =   3105
      Y2              =   3105
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000004&
      X1              =   10305
      X2              =   10305
      Y1              =   990
      Y2              =   2835
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000004&
      X1              =   405
      X2              =   10305
      Y1              =   2835
      Y2              =   2835
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000004&
      X1              =   405
      X2              =   10305
      Y1              =   990
      Y2              =   990
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000004&
      X1              =   405
      X2              =   405
      Y1              =   990
      Y2              =   2835
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภท เอกสาร :"
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
      Left            =   720
      TabIndex        =   21
      Top             =   5265
      Width           =   1560
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่ดูรายงาน :"
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
      Left            =   720
      TabIndex        =   20
      Top             =   5805
      Width           =   1665
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "หมายเหตุ :"
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
      Height          =   300
      Left            =   765
      TabIndex        =   19
      Top             =   4275
      Width           =   915
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "เวลาจัดเสร็จ :"
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
      Left            =   1260
      TabIndex        =   18
      Top             =   3285
      Width           =   1170
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เวลาเริ่มจัด :"
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
      Left            =   5265
      TabIndex        =   17
      Top             =   1755
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่ใบจัดสินค้า : "
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
      Left            =   5040
      TabIndex        =   14
      Top             =   1260
      Width           =   1515
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "พนักงานจัดสินค้า :"
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
      Left            =   4950
      TabIndex        =   12
      Top             =   3285
      Width           =   1545
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบจัดสินค้า :"
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
      Left            =   990
      TabIndex        =   11
      Top             =   1305
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "บันทึกเวลาการจัดสินค้า"
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
      TabIndex        =   10
      Top             =   225
      Width           =   7800
   End
End
Attribute VB_Name = "Form87"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vDocNo, vDocno1 As String
Dim vCheckNumber As String

Private Sub CMDChangePicker_Click()
Dim vQuery As String
Dim vPicker As String

vPicker = Trim(CMBPicker.Text)
vQuery = "Update npmaster.dbo.TB_IV_PulseOfPicking  set picker = '" & vPicker & "' where pickingno = '" & vDocNo & "'  "
gConnection.Execute vQuery
End Sub

Private Sub CMDFinish_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocdate, vDateStart, vPicker As String
Dim vPickingNo, vDescription As String
Dim vPayNumber As String
Dim vWHCode As String
Dim vCheckDat As Integer

On Error GoTo ErrDescription

vPicker = Trim(CMBPicker.Text)
If vPicker <> "" Then
If vCheckNumber = "PK" Or Len(vDocNo) < 4 Then
    vPicker = Trim(CMBPicker.Text)
    vDescription = Trim(TXT876.Text)
    vQuery = "exec dbo.USP_IV_FinishPulseOfPicking  '" & vDocNo & "','" & vPicker & "' ,'" & vDescription & "' "
    gConnection.Execute vQuery
    CMDFinish.Enabled = False
    CMBPicker.Text = ""
    TXTPickingDate.Text = ""
    TXTStart.Text = ""
    TXTStop.Text = ""
'ElseIf vCheckNumber <> "PK" And Len(vDocno) > 8 Then
 '   vQuery = "SELECT WHCode FROM npmaster.dbo.NP_PayGoods where paynumber = '" & vDocno & "' "
  '  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   '     vWHCode = Trim(vRecordset.Fields("whcode").Value)
    'End If
    'vRecordset.Close
    'vPicker = Trim(CMBPicker.Text)
    'vDescription = Trim(TXT876.Text)
    'vPayNumber = Trim(TXT871.Text)
    'vQuery = "insert into  npmaster.dbo.TB_IV_CountTimePayNumber  (PayNumber,WHCode,PayFinish,UserPack,Mydescription)  values('" & vPayNumber & "','" & vWHCode & "',getdate(),'" & vPicker & "','" & vDescription & "')"
    'gConnection.Execute vQuery
    'CMDFinish.Enabled = False
    'CMDPrint.Enabled = True
    'CMBPicker.Text = ""
    'TXT873.Text = ""
    'TXT874.Text = ""
    'TXT875.Text = ""
    'TXT876.Text = ""
    TXTDocno1.SetFocus
End If
Else
    MsgBox "กรุณาใส่ข้อมูล พนักงานจัดสินค้าด้วย"
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CMDGraph_Click()
Call PrintGraph
End Sub

Private Sub CMDPrint_Click()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocdate, vDateStart, vDateEnd As String
Dim vPickingNo As String

On Error GoTo ErrDescription

vQuery = "exec dbo.USP_IV_StartPulseOfPicking  '" & vDocNo & "' "
gConnection.Execute vQuery
CMDPrint.Enabled = False
TXTDocno1.SetFocus
TXTPickingDate.Text = ""

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub CMDReport_Click()
If CMB101.Text = "ใบหยิบสินค้า" Then
    Call PrintPicking
ElseIf CMB101.Text = "ใบจ่ายสินค้า" Then
    Call PrintPayNumber
Else
    MsgBox "กรุณาเลือกประเภทของรายงานด้วย"
End If
'Call PrintGraph
End Sub



Private Sub Form_Load()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vConnectionString As String
Dim conn As New ADODB.Connection

On Error GoTo ErrDescription

DTP1.Value = Now
vConnectionString = "Provider = SQLOLEDB.1;Data Source = Nebula;Initial Catalog = BPLUS;User ID =VBUSER;PassWord = 132"
conn.Open vConnectionString
vQuery = "select  *  from bcnp.dbo.vw_HR_Checker"
vRecordset.Open vQuery, conn, adOpenDynamic, adLockOptimistic
    If Not vRecordset.EOF Then
    vRecordset.MoveFirst
        While Not vRecordset.EOF
            CMBPicker.AddItem Trim(vRecordset.Fields("picker").Value)
        vRecordset.MoveNext
        Wend
    End If
    vRecordset.Close
    
CMB101.AddItem Trim("ใบหยิบสินค้า")
CMB101.AddItem Trim("ใบจ่ายสินค้า")

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Private Sub TXTDocno1_KeyPress(KeyAscii As Integer)
Dim vQuery As String
Dim vRecordset1 As New ADODB.Recordset
Dim vRecordset As New ADODB.Recordset
Dim vDocdate, vDateStart, vDateEnd As Date
Dim vPickingNo As String
Dim vZoneLoc As String
Dim vWHCode As String
Dim vPickingType As Integer
Dim vShelfGroup As String

On Error GoTo ErrDescription

If KeyAscii = 13 And TXTDocno1.Text <> "" Then
  vDocNo = UCase(Trim(TXTDocno1.Text))
  vCheckNumber = UCase(Left(vDocNo, 2))
  If vCheckNumber = "PK" Then
    TXTStart.Text = ""
    TXTStop.Text = ""
    CMBPicker.Text = ""
    CMDChangePicker.Enabled = False
    vQuery = "select pickingno,pickingdate,whcode,shelfgroup from npmaster.dbo.NP_PickingSlip_Logs where pickingno = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
      vDocno1 = Trim(vRecordset.Fields("pickingno").Value)
      TXTPickingDate.Text = Trim(vRecordset.Fields("pickingdate").Value)
      vWHCode = Trim(vRecordset.Fields("whcode").Value)
      vShelfGroup = Trim(vRecordset.Fields("shelfgroup").Value)
      vQuery = "exec dbo.USP_NP_SearchPulsePicking  '" & vDocNo & "' "
      If OpenDataBase(gConnection, vRecordset1, vQuery) <> 0 Then
          If Not IsNull(vRecordset1.Fields("printed").Value) Then
            CMDPrint.Enabled = False
            CMDFinish.Enabled = True
            CMDChangePicker.Enabled = False
            TXTStart.Text = Trim(vRecordset1.Fields("printed").Value)
            CMBPicker.SetFocus
            TXTStop.Text = ""
            CMBPicker.Text = ""
          Else
            CMDPrint.Enabled = True
            CMDFinish.Enabled = False
            CMDChangePicker.Enabled = False
            CMDPrint.SetFocus
            TXTStop.Text = ""
            CMBPicker.Text = ""
            TXTStart.Text = ""
          End If
          If Not IsNull(vRecordset1.Fields("finish").Value) Then
            CMDPrint.Enabled = False
            CMDFinish.Enabled = False
            CMDChangePicker.Enabled = True
            TXTStop.Text = Trim(vRecordset1.Fields("finish").Value)
            CMBPicker.Text = Trim(vRecordset1.Fields("picker").Value)
            TXTDocno1.SetFocus
          End If
      Else
        If vShelfGroup <> "M" Then
          vZoneLoc = Trim("Nopadol")
          vPickingType = 1
        Else
          vZoneLoc = Trim("OutLet")
          vPickingType = 2
        End If
        vQuery = "exec dbo.USP_PK_InsertPulseOfPicking '" & vDocno1 & "','" & vZoneLoc & "','" & vWHCode & "'," & vPickingType & " "
        gConnection.Execute vQuery
        CMDPrint.Enabled = True
        CMDPrint.SetFocus
        TXTStop.Text = ""
        CMBPicker.Text = ""
        TXTStart.Text = ""
      End If
      vRecordset1.Close
    Else
      MsgBox "ไม่มีข้อมูลใบหยิบสินค้า เลขที่ " & vDocNo & " ในฐานข้อมูล"
    CMDPrint.Enabled = False
    CMDFinish.Enabled = False
    CMDChangePicker.Enabled = False
    TXTStop.Text = ""
    CMBPicker.Text = ""
    TXTStart.Text = ""
    TXTPickingDate.Text = ""
    End If
    vRecordset.Close
  ElseIf vCheckNumber <> "PK" Then
  vQuery = "exec dbo.USP_NP_SearchPickingReq '" & vDocNo & "' "
  If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vDocno1 = Trim(vRecordset.Fields("docno").Value)
    TXTPickingDate.Text = Now
    vWHCode = "014"
    vQuery = "exec dbo.USP_NP_SearchPulsePicking  '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset1, vQuery) <> 0 Then
          If Not IsNull(vRecordset1.Fields("printed").Value) Then
            CMDPrint.Enabled = False
            CMDFinish.Enabled = True
            CMDChangePicker.Enabled = False
            TXTStart.Text = Trim(vRecordset1.Fields("printed").Value)
            CMBPicker.SetFocus
            TXTStop.Text = ""
            CMBPicker.Text = ""
          Else
            CMDPrint.Enabled = True
            CMDFinish.Enabled = False
            CMDChangePicker.Enabled = False
            CMDPrint.SetFocus
            TXTStop.Text = ""
            CMBPicker.Text = ""
            TXTStart.Text = ""
          End If
          If Not IsNull(vRecordset1.Fields("finish").Value) Then
            CMDPrint.Enabled = False
            CMDFinish.Enabled = False
            CMDChangePicker.Enabled = True
            TXTStop.Text = Trim(vRecordset1.Fields("finish").Value)
            CMBPicker.Text = Trim(vRecordset1.Fields("picker").Value)
            TXTDocno1.SetFocus
          End If
    Else
        vZoneLoc = Trim("OutLet")
        vPickingType = 2
        vQuery = "exec dbo.USP_PK_InsertPulseOfPicking '" & vDocno1 & "','" & vZoneLoc & "','" & vWHCode & "'," & vPickingType & " "
        gConnection.Execute vQuery
        CMDPrint.Enabled = True
        CMDPrint.SetFocus
        TXTStop.Text = ""
        CMBPicker.Text = ""
        TXTStart.Text = ""
    End If
    vRecordset1.Close
  Else
    MsgBox "ไม่มีข้อมูลใบหยิบสินค้า เลขที่ " & vDocNo & " ในวันนี้"
    CMDPrint.Enabled = False
    CMDFinish.Enabled = False
    CMDChangePicker.Enabled = False
    TXTStop.Text = ""
    CMBPicker.Text = ""
    TXTStart.Text = ""
    TXTPickingDate.Text = ""
  End If
  vRecordset.Close
    
  End If
  End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub PrintPicking()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vNowDate As Date
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

vNowDate = DTP1.Day & "/" & DTP1.Month & "/" & DTP1.Year

vRepID = 164
vRepType = "IV"

vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname  from bcnp.dbo.bcreportname where reptype = 'IV' and repid = 164 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With CrystalReport871
.ReportFileName = vReportName & ".rpt"
.ParameterFields(0) = "@vDocdate;" & vNowDate & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Public Sub PrintPayNumber()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vNowDate As Date
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

vNowDate = DTP1.Day & "/" & DTP1.Month & "/" & DTP1.Year
vRepID = 254
vRepType = "IV"
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname  from bcnp.dbo.bcreportname where reptype = 'IV' and repid = 254 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With CrystalReport871
.ReportFileName = vReportName & ".rpt"
.ParameterFields(0) = "@vDocDate;" & vNowDate & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
    Exit Sub
End If
End Sub

Public Sub PrintGraph()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vNowDate1 As Date
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String


vNowDate1 = DTP1.Day & "/" & DTP1.Month & "/" & DTP1.Year

vRepID = 166
vRepType = "IV"
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname  from bcnp.dbo.bcreportname where reptype = 'IV' and repid = 166 "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With CrystalReport872
.ReportFileName = vReportName & ".rpt"
.ParameterFields(0) = "@vStartDate;" & vNowDate1 & ";true"
.ParameterFields(1) = "@vEndDate;" & vNowDate1 & ";true"
.Destination = crptToWindow
.WindowState = crptMaximized
.Action = 1
End With
End Sub

