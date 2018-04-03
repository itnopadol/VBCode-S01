VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form871 
   Caption         =   "บันทึกเวลาจัดสินค้าในโฮมมาร์ท"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form871.frx":0000
   ScaleHeight     =   7995
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   3840
      Top             =   6120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowNavigationCtls=   -1  'True
      WindowShowCancelBtn=   -1  'True
      WindowShowPrintBtn=   -1  'True
      WindowShowExportBtn=   -1  'True
      WindowShowZoomCtl=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowProgressCtls=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton CMD102 
      Caption         =   "ลบเอกสาร"
      Height          =   540
      Left            =   5850
      TabIndex        =   12
      Top             =   3375
      Width           =   1515
   End
   Begin VB.ComboBox CMB101 
      Height          =   315
      Left            =   2175
      TabIndex        =   11
      Top             =   2700
      Width           =   2490
   End
   Begin VB.TextBox TXT105 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   390
      Left            =   5850
      TabIndex        =   9
      Top             =   2025
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.TextBox TXT103 
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
      TabIndex        =   7
      Top             =   2025
      Width           =   2490
   End
   Begin VB.TextBox TXT102 
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
      Left            =   5850
      TabIndex        =   4
      Top             =   1425
      Width           =   2490
   End
   Begin VB.CommandButton CMD103 
      Caption         =   "พิมพ์รายงาน"
      Height          =   540
      Left            =   3150
      TabIndex        =   2
      Top             =   4200
      Width           =   1515
   End
   Begin VB.CommandButton CMD101 
      Caption         =   "เวลาที่จัดเสร็จ"
      Enabled         =   0   'False
      Height          =   540
      Left            =   3150
      TabIndex        =   1
      Top             =   3375
      Width           =   1515
   End
   Begin VB.TextBox TXT101 
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
      Height          =   390
      Left            =   2175
      TabIndex        =   0
      Top             =   1425
      Width           =   2490
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "เวลาที่เริ่มจัด"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   4800
      TabIndex        =   10
      Top             =   2025
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่บิล"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   1500
      TabIndex        =   8
      Top             =   2025
      Width           =   690
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "คนจัดสินค้า"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   1200
      TabIndex        =   6
      Top             =   2700
      Width           =   915
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่เอกสาร"
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   4875
      TabIndex        =   5
      Top             =   1425
      Width           =   915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบจ่ายสินค้า"
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   825
      TabIndex        =   3
      Top             =   1425
      Width           =   1215
   End
End
Attribute VB_Name = "Form871"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vCheckDoc1 As Integer
Dim vPicker As String

On Error GoTo ErrDescription

    If TXT101.Text <> "" Then
    If CMB101.Text <> "" Then
    vDocNo = Trim(TXT101.Text)
    vQuery = "select paynumber from npmaster.dbo.TB_IV_TimeStampReceivingSlip where  paynumber = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vCheckDoc1 = 1
    Else
        vCheckDoc1 = 0
    End If
    vRecordset.Close
    
    If vCheckDoc1 = 1 Then
        vPicker = Trim(CMB101.Text)
        vQuery = "update npmaster.dbo.TB_IV_TimeStampReceivingSlip  set EndDate = getdate(),Picker = '" & vPicker & "'  where paynumber = '" & vDocNo & "' "
        gConnection.Execute vQuery
    Else
        MsgBox "ยังไม่ได้มีการเก็บข้อมูล เลขที่ใบจ่ายเลขที่ " & vDocNo & " "
        Exit Sub
    End If
    
    TXT101.Text = ""
    TXT102.Text = ""
    TXT103.Text = ""
    TXT105.Text = ""
    TXT105.Visible = False
    Label5.Visible = False
    CMB101.Text = ""
    MsgBox "เก็บเวลาการจัดสินค้าของใบจ่ายเลขที่ " & vDocNo & " เรียบร้อย"
    TXT101.SetFocus
    CMD101.Enabled = False
    Else
    MsgBox "กรุณาใส่ชื่อคนจัดสินค้าด้วยครับ"
    End If
    End If


ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Private Sub CMD102_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vAnswer As Integer
Dim vCheckPayNumber As Integer

On Error GoTo ErrDescription

If TXT101.Text <> "" Then
    vDocNo = Trim(TXT101.Text)
    vQuery = "select  paynumber from npmaster.dbo.TB_IV_TimeStampReceivingSlip where paynumber = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vCheckPayNumber = 1
    Else
        vCheckPayNumber = 0
    End If
    vRecordset.Close
    
    If vCheckPayNumber = 1 Then
            vAnswer = MsgBox("คุณต้องการที่จะลบการบันทึกเวลาการจัดสินค้าของใบจ่ายเลขที่ " & vDocNo & "ใช่หรือไม่", vbYesNo, "ข้อความสอบถาม")
            If vAnswer = 6 Then
            vQuery = "Delete npmaster.dbo.TB_IV_TimeStampReceivingSlip where paynumber = '" & vDocNo & "' "
            gConnection.Execute vQuery
            Else
            Exit Sub
            End If
    Else
        MsgBox "ไม่มีเลขที่ใบจ่ายเลขที่ " & vDocNo & " ในตารางบันทึกเวลาจัดสินค้า "
    End If
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD103_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

vRepID = 207
vRepType = "IV"
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = 207 and reptype = 'IV' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vReportName = Trim(vRecordset.Fields("reportname").Value)
End If
vRecordset.Close

With Crystal101
.ReportFileName = vReportName & ".rpt"
.WindowState = crptMaximized
.Destination = crptToWindow
.Action = 1
End With
TXT101.SetFocus

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error GoTo ErrDescription

vQuery = "select  salename  from npmaster.dbo.BCSaleGroup where    picker = 1 order by salecode desc"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    CMB101.AddItem Trim(vRecordset.Fields("salename").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Private Sub TXT101_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vCheckDoc As Integer
Dim vCheckDoc1 As Integer

On Error GoTo ErrDescription

If KeyAscii = 13 Then
    If TXT101.Text <> "" Then
    vDocNo = Trim(TXT101.Text)
    
    vQuery = "select paynumber from  npmaster.dbo.tb_iv_timestampreceivingslip where paynumber = '" & vDocNo & "' and (enddate is not null or picker is  not null )"
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckDoc1 = 1
    Else
    vCheckDoc1 = 0
    End If
    vRecordset.Close
    
    If vCheckDoc1 = 0 Then
    vQuery = "select invoiceno,paydatetime from npmaster.dbo.np_paygoods where  whcode = '014'  and paynumber = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        TXT102.Text = Trim(vRecordset.Fields("paydatetime").Value)
        TXT103.Text = Trim(vRecordset.Fields("invoiceno").Value)
    Else
        MsgBox "ไม่มีเอกสาร ใบจ่ายสินค้าเลขที่ " & vDocNo & " ที่เป็นใบจ่ายสินค้าคลัง 014 "
        Exit Sub
    End If
    vRecordset.Close

    
    vQuery = "select paynumber from npmaster.dbo.TB_IV_TimeStampReceivingSlip where  paynumber = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vCheckDoc = 1
    Else
        vCheckDoc = 0
    End If
    vRecordset.Close
    
    If vCheckDoc = 0 Then
    vQuery = "insert into npmaster.dbo.TB_IV_TimeStampReceivingSlip (Paynumber,InvoiceNo,StartDate,UserStamp) values ('" & vDocNo & "','" & TXT103.Text & "',getdate(),upper('" & vUserID & "'))"
    gConnection.Execute vQuery
    MsgBox "ได้เก็บเวลาเริ่มจัดสินค้าแล้วครับ"
    TXT101.Text = ""
    TXT102.Text = ""
    TXT103.Text = ""
    TXT101.SetFocus
    ElseIf vCheckDoc = 1 Then
    CMD101.Enabled = True
    TXT105.Visible = True
    Label5.Visible = True
    vQuery = "select startdate from npmaster.dbo.TB_IV_TimeStampReceivingSlip where paynumber = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        TXT105.Text = Trim(vRecordset.Fields("startdate").Value)
    End If
    vRecordset.Close
    'Else
    'MsgBox "ไม่สามารถแก้ไข ข้อมูลที่บันทึกเสร็จไปแล้วได้"
    'End If
    End If
Else
MsgBox "ไม่สามารถแก้ไข ข้อมูลที่บันทึกเสร็จไปแล้วได้"
End If
End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If

End Sub
