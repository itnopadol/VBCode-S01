VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form Form97 
   Caption         =   "เปลี่ยนเลขที่เอกสาร/เลขที่ภาษี/วันที่เอกสาร"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form97.frx":0000
   ScaleHeight     =   7980
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Crystal101 
      Left            =   240
      Top             =   6720
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
   Begin VB.CheckBox CHK104 
      Caption         =   "Check1"
      Height          =   195
      Left            =   2250
      TabIndex        =   30
      Top             =   1950
      Width           =   165
   End
   Begin MSComCtl2.DTPicker DTP102 
      Height          =   390
      Left            =   8250
      TabIndex        =   28
      Top             =   4950
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   688
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   58916865
      CurrentDate     =   38399
   End
   Begin VB.TextBox TXT106 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2250
      TabIndex        =   26
      Top             =   4950
      Width           =   1740
   End
   Begin VB.ComboBox CMB101 
      Enabled         =   0   'False
      Height          =   315
      Left            =   8250
      TabIndex        =   24
      Top             =   2850
      Width           =   1740
   End
   Begin VB.CommandButton CMD103 
      Caption         =   "พิมพ์รายงาน"
      Height          =   540
      Left            =   7725
      TabIndex        =   23
      Top             =   6750
      Width           =   1740
   End
   Begin VB.CommandButton CMD102 
      Caption         =   "ล้างข้อมูล"
      Height          =   540
      Left            =   4500
      TabIndex        =   22
      Top             =   6750
      Width           =   1740
   End
   Begin VB.PictureBox Picture2 
      Height          =   3390
      Left            =   5250
      ScaleHeight     =   3330
      ScaleWidth      =   30
      TabIndex        =   21
      Top             =   2550
      Width           =   90
   End
   Begin VB.PictureBox Picture1 
      Height          =   90
      Index           =   1
      Left            =   75
      ScaleHeight     =   30
      ScaleWidth      =   10530
      TabIndex        =   20
      Top             =   6000
      Width           =   10590
   End
   Begin VB.PictureBox Picture1 
      Height          =   90
      Index           =   0
      Left            =   75
      ScaleHeight     =   30
      ScaleWidth      =   10530
      TabIndex        =   19
      Top             =   2400
      Width           =   10590
   End
   Begin VB.TextBox TXT105 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   8250
      TabIndex        =   16
      Top             =   4425
      Width           =   1740
   End
   Begin VB.TextBox TXT102 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2250
      TabIndex        =   15
      Top             =   3900
      Width           =   1740
   End
   Begin VB.CheckBox CHK101 
      Caption         =   "Check1"
      Height          =   195
      Left            =   2250
      TabIndex        =   11
      Top             =   1050
      Width           =   165
   End
   Begin VB.CheckBox CHK103 
      Caption         =   "Check3"
      Height          =   195
      Left            =   2250
      TabIndex        =   10
      Top             =   1350
      Width           =   165
   End
   Begin VB.CheckBox CHK102 
      Caption         =   "Check2"
      Height          =   195
      Left            =   2250
      TabIndex        =   9
      Top             =   1650
      Width           =   165
   End
   Begin MSComCtl2.DTPicker DTP101 
      Height          =   390
      Left            =   8250
      TabIndex        =   2
      Top             =   3900
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   688
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   58916865
      CurrentDate     =   38397
   End
   Begin VB.TextBox TXT103 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2250
      TabIndex        =   1
      Top             =   4425
      Width           =   1740
   End
   Begin VB.TextBox TXT101 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2250
      TabIndex        =   0
      Top             =   3375
      Width           =   1740
   End
   Begin VB.TextBox TXT104 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8250
      TabIndex        =   4
      Top             =   3375
      Width           =   1740
   End
   Begin VB.CommandButton CMD101 
      Caption         =   "ปรับปรุงเอกสาร"
      Height          =   540
      Left            =   1425
      TabIndex        =   3
      Top             =   6750
      Width           =   1740
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "เปลี่ยนวันที่ใบกำกับภาษี"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   2550
      TabIndex        =   31
      Top             =   1950
      Width           =   1890
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เปลี่ยนวันที่ใบกำกับภาษี"
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   5760
      TabIndex        =   29
      Top             =   4995
      Width           =   2415
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่ใบกำกับภาษี"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   525
      TabIndex        =   27
      Top             =   4995
      Width           =   1665
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลือกหัวเอกสาร"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   6435
      TabIndex        =   25
      Top             =   2880
      Width           =   1740
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เปลี่ยนเป็นวันที่เอกสาร"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   5985
      TabIndex        =   18
      Top             =   3960
      Width           =   2190
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เปลี่ยนเป็นใบกำกับภาษีเลขที่"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   5625
      TabIndex        =   17
      Top             =   4500
      Width           =   2565
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่เอกสาร"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1275
      TabIndex        =   14
      Top             =   3960
      Width           =   915
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบกำกับภาษี"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   375
      TabIndex        =   13
      Top             =   4500
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่เอกสาร"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   825
      TabIndex        =   12
      Top             =   3420
      Width           =   1365
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "เปลี่ยนวันที่เอกสาร"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   2550
      TabIndex        =   8
      Top             =   1350
      Width           =   1440
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เปลี่ยนเป็นเอกสารเลขที่"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   6435
      TabIndex        =   7
      Top             =   3420
      Width           =   1740
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "เปลี่ยนเลขที่เอกสาร"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   2550
      TabIndex        =   6
      Top             =   1050
      Width           =   1515
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "เปลี่ยนเลขที่ใบกำกับภาษี"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   2550
      TabIndex        =   5
      Top             =   1650
      Width           =   1965
   End
End
Attribute VB_Name = "Form97"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim vCountNum As Integer
Dim vChangeDoc1 As String

Private Sub CHK101_Click()
If CMB101.Enabled = False Then
    CMB101.Enabled = True
Else
    CMB101.Enabled = False
    TXT104.Text = ""
    CMB101.Text = ""
End If
TXT101.SetFocus
End Sub

Private Sub CHK102_Click()
If TXT105.Enabled = False Then
    TXT105.Enabled = True
Else
    TXT105.Enabled = False
End If
TXT101.SetFocus
End Sub

Private Sub CHK103_Click()
If DTP101.Enabled = False Then
    DTP101.Enabled = True
Else
    DTP101.Enabled = False
End If

TXT101.SetFocus
End Sub

Private Sub CHK104_Click()
If DTP102.Enabled = False Then
    DTP102.Enabled = True
Else
    DTP102.Enabled = False
End If

TXT101.SetFocus
End Sub


Private Sub CMB101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vChangeDoc As String
Dim vLenHeader As Integer
Dim vLenNumber As Integer
Dim vAnswer As Integer
Dim vLenNumber1 As String, vLenNumber2 As String, vLenNumber3 As String
Dim vLenNumber4 As String, vLenNumber5 As String, vLenNumber6 As String
Dim vSelectDoc As String
Dim vDocno1 As String

On Error GoTo ErrDescription

If TXT101.Text <> "" Then
vDocNo = Trim(TXT101.Text)
vLenNumber = (Len(vDocNo) - InStr(1, vDocNo, "-"))
vLenHeader = InStr(1, vDocNo, "-")
vLenNumber1 = Year(TXT102.Text)
If Year(TXT102.Text) < 2500 Then
vLenNumber1 = Year(TXT102.Text) + 543
Else
vLenNumber1 = Year(TXT102.Text)
End If
vLenNumber2 = Right(vLenNumber1, 2)
vLenNumber3 = Month(TXT102.Text)
vLenNumber4 = Len(vLenNumber3)
If vLenNumber4 < 2 Then
vLenNumber5 = "0" & vLenNumber3
Else
vLenNumber5 = vLenNumber3
End If
vLenNumber6 = vLenNumber2 & vLenNumber5
vSelectDoc = Trim(CMB101.Text)
vDocno1 = vLenNumber6 'Left(Right(vDocno, Len(vDocno) - vCountNum), vLenNumber)

vQuery = "select    top 1 right(docno," & vLenNumber & ")+1 as docno" _
                    & " from bcnp.dbo.bcapinvoice  " _
                    & " where ltrim(left(docno," & vCountNum & ")) =   '" & vSelectDoc & "' and " _
                    & " left(right(docno,len(docno)-" & vCountNum & ")," & vLenNumber & ") = left(right('" & vDocno1 & "',len('" & vDocNo & "')-" & vCountNum & ")," & vLenNumber & ")  order by docno desc "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vChangeDoc = Trim(vRecordset.Fields("docno").Value)
Else
    vChangeDoc = Format(1, "0000")
End If
vRecordset.Close

vChangeDoc1 = UCase(Format(vChangeDoc, "0000"))
vChangeDoc1 = vSelectDoc & vDocno1 & "-" & vChangeDoc1
TXT104.Text = vChangeDoc1
Else
MsgBox "กรุณาใส่เลขที่เอกสารที่ต้องการปรับปรุงก่อนนะครับ"
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD101_Click()

If CHK101.Value = 0 And CHK102.Value = 0 And CHK103.Value = 0 And CHK104.Value = 0 Then
MsgBox "กรุณาเลือกหัวข้อในการเปลี่ยนข้อมูลด้วยนะครับ"
Exit Sub
End If

If CHK101.Value = 1 And TXT104.Text = "" Then
    MsgBox "เลือกหัวเอกสารที่จะเปลี่ยนด้วยนะครับ"
    Exit Sub
End If

If CHK101.Value = 1 Then
        Call ChangeDocno
End If
If CHK102.Value = 1 Then
        If TXT105.Text <> "" Then
        Call ChangeTaxNo
        Call ChangeTaxNoBackOffice
        Else
        MsgBox "กรุณาใส่ข้อมูลในการเปลี่ยนเลขที่ใบกำกับภาษีด้วยครับ"
        End If
End If
If CHK103.Value = 1 Then
        If DTP101.Value <> Trim(TXT102.Text) Then
        Call ChangeDocDate
        Else
        MsgBox "กรุณาใส่ข้อมูลในการเปลี่ยนเอกสารด้วยครับ"
        End If
End If
If CHK104.Value = 1 Then
        If DTP102.Value <> Trim(TXT106.Text) Then
        Call ChangeTaxDate
        Call ChangeTaxDateBackOffice
        Else
        MsgBox "กรุณาใส่ข้อมูลในการเปลี่ยนวันที่ใบกำกับภาษีด้วยครับ"
        End If
End If

TXT101.Text = ""
TXT102.Text = ""
TXT103.Text = ""
TXT104.Text = ""
TXT105.Text = ""
TXT106.Text = ""
TXT101.SetFocus
CHK101.Value = 0
CHK102.Value = 0
CHK103.Value = 0
CHK104.Value = 0


End Sub

Public Sub CheckNumeric()
Dim vDocNo As String
Dim vText As String

On Error GoTo ErrDescription
vDocNo = Trim(TXT101.Text)

For i = 1 To Len(TXT101.Text)
    If Mid(vDocNo, i, 1) = 0 Or Mid(vDocNo, i, 1) = 1 Or Mid(vDocNo, i, 1) = 2 Or Mid(vDocNo, i, 1) = 3 Or Mid(vDocNo, i, 1) = 4 Or Mid(vDocNo, i, 1) = 5 Or Mid(vDocNo, i, 1) = 6 Or Mid(vDocNo, i, 1) = 7 Or Mid(vDocNo, i, 1) = 8 Or Mid(vDocNo, i, 1) = 9 Then
        vCheckValue = True
        vCountNum = i - 1
        Exit Sub
    Else
        vCheckValue = False
    End If
Next i

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub ChangeDocno()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vChangeDoc As String
Dim vLenHeader As Integer
Dim vLenNumber As Integer
Dim vAnswer As Integer
Dim vTaxNo As String
Dim vDocdate As String
Dim vTaxDate As String
Dim vCheckNewDocNo As Integer

On Error GoTo ErrDescription

vDocNo = Trim(TXT101.Text)
If TXT104.Text <> "" Then
vChangeDoc1 = UCase(TXT104.Text)
Else
MsgBox "เลือกหัวเอกสารด้วยนะครับ"
Exit Sub
End If

Line1:
vQuery = "select  *  from bcnp.dbo.bcapinvoice where docno = '" & vChangeDoc1 & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckNewDocNo = 1
    MsgBox "เลขที่เอกสาร " & vChangeDoc1 & " มีอยู่แล้ว ต้องเปลี่ยนเป็นเลขที่ใหม่"
Else
    vCheckNewDocNo = 0
End If
vRecordset.Close

If vCheckNewDocNo = 1 Then
Call CMB101_Click
vChangeDoc1 = UCase(TXT104.Text)
GoTo Line1
End If

vAnswer = MsgBox("คุณต้องเปลี่ยนเลขที่เอกสาร จากเลขที่ " & vDocNo & " เป็นเลขที่ " & vChangeDoc1 & " นี้ใช่หรือไม่", vbYesNo, "ข้อความสอบถาม")
If vAnswer = 6 Then
    vQuery = "Exec bcnp.dbo.USP_AP_ChangeDocno '" & vDocNo & "' ,'" & vChangeDoc1 & "' "
    gConnection.Execute vQuery
    MsgBox "ได้มีการเปลี่ยนเลขที่เอกสาร จากเลขที่ " & vDocNo & " เป็นเลขที่ " & vChangeDoc1 & " เรียบร้อยแล้ว"
Else
    Exit Sub
End If
vDocNo = UCase(vDocNo)
vTaxNo = Trim(TXT103.Text)
vDocdate = Trim(TXT102.Text)
vTaxDate = Trim(TXT106.Text)
vQuery = "insert into npmaster.dbo.TB_AP_ChangeDataPayment  (OldDocNo,OldDocDate,OldTaxNo,OldTaxDate,NewDocNo,NewDocDate,NewTaxNo,NewTaxDate,UserChange,DateChange) " _
                    & " values('" & vDocNo & "','" & vDocdate & "','" & vTaxNo & "','" & vTaxDate & "','" & vChangeDoc1 & "','','','','" & vUserID & "',getdate())"
gConnection.Execute vQuery

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub ChangeDocnoBackOffice()

End Sub

Private Sub CMD102_Click()
TXT101.Text = ""
TXT102.Text = ""
TXT103.Text = ""
TXT104.Text = ""
TXT105.Text = ""
TXT106.Text = ""
CMB101.Text = ""
CHK101.Value = 0
CHK102.Value = 0
CHK103.Value = 0
CHK104.Value = 0

End Sub

Private Sub CMD103_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

vRepID = 206
vRepType = "AP"
vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
'vQuery = "select reportname from bcnp.dbo.bcreportname where repid = 206 and reptype = 'AP' "
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
vQuery = "select distinct left(docno,2) as docno from bcnp.dbo.bcapinvoice where grbillstatus in (0,1) and grirbillstatus = 2 order by docno "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
    CMB101.AddItem Trim(vRecordset.Fields("docno").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

'Call InitializeDataBaseVat

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If

End Sub


Private Sub TXT101_KeyPress(KeyAscii As Integer)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String

On Error GoTo ErrDescription
If KeyAscii = 13 Then
    If TXT101.Text <> "" Then
        vDocNo = TXT101.Text
        Call CheckNumeric
        vQuery = "select a.docno,a.docdate,b.taxno,b.taxdate from bcnp.dbo.bcapinvoice a " _
        & " left join bcnp.dbo.bcinputtax b on a.docno = b.docno " _
        & " where a.docno = '" & vDocNo & "' and a.grbillstatus in (0,1) and a.grirbillstatus in (0,2) and a.iscancel = 0"
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            TXT102.Text = Trim(vRecordset.Fields("docdate").Value)
                If Not IsNull(Trim(vRecordset.Fields("taxno").Value)) Then
                    TXT103.Text = Trim(vRecordset.Fields("taxno").Value)
                Else
                TXT103.Text = "NoTaxNo"
                End If
                If Not IsNull(Trim(vRecordset.Fields("taxdate").Value)) Then
                TXT106.Text = Trim(vRecordset.Fields("taxdate").Value)
                Else
                TXT106.Text = Trim(vRecordset.Fields("docdate").Value)
                End If
            Else
            MsgBox "เอกสาร เลขที่ " & vDocNo & " ไม่สามารถเปลี่ยนเอกสารได้ โปรดตรวจสอบด้วยนะครับ"
            TXT104.Text = ""
            TXT105.Text = ""
            Exit Sub
        End If
        vRecordset.Close
    End If
    TXT104.Text = ""
    TXT105.Text = ""
    DTP101.Value = TXT102.Text
    DTP102.Value = TXT106.Text
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub ChangeTaxNo()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vTaxNo As String, vTaxNo1 As String
Dim vCheckTax As Integer
Dim vDocdate As String, vTaxDate As String

On Error GoTo ErrDescription
If CHK101.Value = 0 Then
    vDocNo = Trim(TXT101.Text)
Else
    vDocNo = vChangeDoc1
End If

vQuery = "select docno from bcnp.dbo.bcinputtax where docno = '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckTax = 1
Else
vCheckTax = 0
End If
vRecordset.Close

If vCheckTax = 1 Then
    vTaxNo = Trim(TXT105.Text)
    vTaxNo1 = Trim(TXT103.Text)
    If vTaxNo <> "" Then
        vQuery = "Exec bcnp.dbo.USP_AP_ChangeTaxno '" & vDocNo & "','" & vTaxNo & "' "
        gConnection.Execute vQuery
        MsgBox "เอกสารเลขที่ " & vDocNo & " ได้แก้ไข ใบกำกับภาษีเป็นเลขที่ " & vTaxNo & " เรียบร้อยแล้วครับ"
    Else
        MsgBox "ไม่ได้ใส่ข้อมูลการเปลี่ยนใบกำกับภาษี"
    End If
Else
    MsgBox "ไม่มีข้อมูลเลขที่เอกสาร " & vDocNo & " ในตาราง BCInputtax กรุณาตรวจสอบด้วยครับ"
End If

If CHK101.Value = 1 Then
    vQuery = "Update npmaster.dbo.TB_AP_ChangeDataPayment set OldTaxNo = '" & vTaxNo1 & "',NewTaxNo = '" & vTaxNo & "' where newdocno = '" & vDocNo & "' and oldtaxno = '" & vTaxNo1 & "' "
    gConnection.Execute vQuery
Else
    vDocNo = UCase(vDocNo)
    vDocdate = Trim(TXT102.Text)
    vTaxDate = Trim(TXT106.Text)
    vQuery = "insert into npmaster.dbo.TB_AP_ChangeDataPayment  (OldDocNo,OldDocDate,OldTaxNo,OldTaxDate,NewDocNo,NewDocDate,NewTaxNo,NewTaxDate,UserChange,DateChange) " _
    & " values('" & vDocNo & "','" & vDocdate & "','" & vTaxNo1 & "','" & vTaxDate & "','','','" & vTaxNo & "','','" & vUserID & "',getdate())"
    gConnection.Execute vQuery
End If

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub


Public Sub ChangeTaxNoBackOffice()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vTaxNo As String, vTaxNo1 As String
Dim vCheckTax As Integer
Dim vDocdate As String, vTaxDate As String
Dim vVatRecordset As New ADODB.Recordset


On Error GoTo ErrDescription

If CHK101.Value = 0 Then
    vDocNo = Trim(TXT101.Text)
Else
    vDocNo = vChangeDoc1
End If

'vQuery = "select docno from solar.bcvat.dbo.bcinputtax where docno = '" & vDocNo & "' "
'If OpenDataBaseBCVat(vVatConnection, vVatRecordset, vQuery) <> 0 Then
'vCheckTax = 1
'Else
'vCheckTax = 0
'End If
'vVatRecordset.Close

'If vCheckTax = 1 Then
 '   vTaxNo = Trim(TXT105.Text)
  '  vTaxNo1 = Trim(TXT103.Text)
   ' If vTaxNo <> "" Then
    '  vQuery = "set dateformat dmy "
     ' vVatConnection.Execute vQuery
      '
      'vQuery = "Update  solar.bcvat.dbo.bcapinvoice set taxno = '" & vTaxNo & "'  where iscancel = 0 and docno =  '" & vDocNo & "'"
      'vVatConnection.Execute vQuery
      '
      'vQuery = "Update  solar.bcvat.dbo.bcinputtax set taxno = '" & vTaxNo & "'  where iscancel = 0 and docno =  '" & vDocNo & "'"
      'vVatConnection.Execute vQuery
      '
      'vQuery = "Update  solar.bcvat.dbo.bcirsub set taxno = '" & vTaxNo & "'  where iscancel = 0 and docno =  '" & vDocNo & "'"
      'vVatConnection.Execute vQuery
    '
     ' MsgBox "เอกสารเลขที่ " & vDocNo & " ได้แก้ไข ใบกำกับภาษีเป็นเลขที่ " & vTaxNo & " เรียบร้อยแล้วครับ ที่ข้อมูลหลังร้าน"
      
    'Else
     '   MsgBox "ไม่ได้ใส่ข้อมูลการเปลี่ยนใบกำกับภาษี"
    'End If
'Else
 '   MsgBox "ไม่มีข้อมูลเลขที่เอกสาร " & vDocNo & " ในตาราง BCInputtax กรุณาตรวจสอบด้วยครับ"
'End If


vDocNo = UCase(vDocNo)
vDocdate = Trim(TXT102.Text)
vTaxDate = Trim(TXT106.Text)
vQuery = "Update npmaster.dbo.TB_AP_ChangeDataPayment   set mydescription = 'Update BCVat'  where OldDocNo = '" & vDocNo & "' and OldTaxNo = '" & vTaxNo1 & "' and NewTaxNo = '" & vTaxNo & "' "
gConnection.Execute vQuery


ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub


Public Sub ChangeDocDate()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vDocdate As String
Dim vDocDate1 As String
Dim vTaxNo As String
Dim vTaxDate As String

On Error GoTo ErrDescription
If CHK101.Value = 0 Then
    vDocNo = Trim(TXT101.Text)
Else
    vDocNo = vChangeDoc1
End If
vTaxNo = Trim(TXT103.Text)
vTaxDate = Trim(TXT106.Text)
vDocDate1 = Trim(TXT102.Text)
vDocdate = DTP101.Day & "/" & DTP101.Month & "/" & DTP101.Year
    If vDocdate <> vDocDate1 Then
        vQuery = "Exec bcnp.dbo.USP_AP_ChangeDocDate '" & vDocNo & "','" & vDocdate & "' "
        gConnection.Execute vQuery
        MsgBox "เลขที่เอกสาร " & vDocNo & " ได้แก้ไขวันที่เอกสารเป็นวันที่ " & vDocdate & " เรียบร้อยแล้ว  "
    Else
        MsgBox "ไม่ได้ใส่ข้อมูลการเปลี่ยนวันที่เอกสาร"
        Exit Sub
    End If

    If CHK101.Value = 1 Then
        vQuery = "Update npmaster.dbo.TB_AP_ChangeDataPayment set OldDocdate = '" & vDocDate1 & "',newdocdate = '" & vDocdate & "' where newdocno = '" & vDocNo & "' "
        gConnection.Execute vQuery
    ElseIf CHK102.Value = 1 Then
        vQuery = "Update npmaster.dbo.TB_AP_ChangeDataPayment set OldDocdate = '" & vDocDate1 & "',newdocdate = '" & vDocdate & "' where olddocno = '" & vDocNo & "' and oldtaxno = '" & vTaxNo & "' "
        gConnection.Execute vQuery
    Else
        vQuery = "insert into npmaster.dbo.TB_AP_ChangeDataPayment  (OldDocNo,OldDocDate,OldTaxNo,OldTaxDate,NewDocNo,NewDocDate,NewTaxNo,NewTaxDate,UserChange,DateChange) " _
        & " values('" & vDocNo & "','" & vDocDate1 & "','" & vTaxNo & "','" & vTaxDate & "','','" & vDocdate & "','','','" & vUserID & "',getdate())"
        gConnection.Execute vQuery
    End If
    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub ChangeDocDateBackOffice()

End Sub

Public Sub ChangeTaxDate()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vTaxDate As String
Dim vTaxDate1 As String
Dim vCheckTax As Integer
Dim vTaxNo As String, vDocdate As String, vDocDate1 As String

On Error GoTo ErrDescription
If CHK101.Value = 0 Then
    vDocNo = Trim(TXT101.Text)
Else
    vDocNo = vChangeDoc1
End If

If CHK102.Value = 0 Then
    vTaxNo = Trim(TXT103.Text)
Else
    vTaxNo = Trim(TXT105.Text)
End If

vQuery = "select docno from bcnp.dbo.bcinputtax where docno = '" & vDocNo & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
vCheckTax = 1
Else
vCheckTax = 0
End If
vRecordset.Close

If vCheckTax = 1 Then
    vTaxDate1 = Trim(TXT106.Text)
    vTaxDate = DTP102.Day & "/" & DTP102.Month & "/" & DTP102.Year
        If vTaxDate <> vTaxDate1 Then
            vQuery = "Exec bcnp.dbo.USP_AP_ChangeTaxDate '" & vDocNo & "','" & vTaxNo & "','" & vTaxDate & "' "
            gConnection.Execute vQuery
            MsgBox "เลขที่เอกสาร " & vDocNo & " ได้แก้ไขวันที่ใบกำกับภาษีเป็นวันที่ " & vTaxDate & " เรียบร้อยแล้ว  "
        Else
            MsgBox "ไม่ได้ใส่ข้อมูลการเปลี่ยนวันที่เอกสาร"
            Exit Sub
        End If
    
        If CHK101.Value = 1 And CHK102.Value = 0 Then
            vQuery = "Update npmaster.dbo.TB_AP_ChangeDataPayment set Oldtaxdate = '" & vTaxDate1 & "',newtaxdate ='" & vTaxDate & "' where newdocno = '" & vDocNo & "' and oldtaxno = '" & vTaxNo & "' "
            gConnection.Execute vQuery
        ElseIf CHK101.Value = 1 And CHK102.Value = 1 Then
            vQuery = "Update npmaster.dbo.TB_AP_ChangeDataPayment set Oldtaxdate = '" & vTaxDate1 & "',newtaxdate ='" & vTaxDate & "' where newdocno = '" & vDocNo & "' and newtaxno = '" & vTaxNo & "' "
            gConnection.Execute vQuery
        ElseIf CHK101.Value = 0 And CHK102.Value = 1 Then
            vQuery = "Update npmaster.dbo.TB_AP_ChangeDataPayment set Oldtaxdate = '" & vTaxDate1 & "',newtaxdate ='" & vTaxDate & "' where olddocno = '" & vDocNo & "' and newtaxno = '" & vTaxNo & "' "
            gConnection.Execute vQuery
        ElseIf CHK103.Value = 1 Then
            vDocDate1 = DTP101.Day & "/" & DTP101.Month & "/" & DTP101.Year
            vQuery = "Update npmaster.dbo.TB_AP_ChangeDataPayment set Oldtaxdate = '" & vTaxDate1 & "',newtaxdate ='" & vTaxDate & "' where olddocno = '" & vDocNo & "' and newdocdate = '" & vDocDate1 & "' "
            gConnection.Execute vQuery
        Else
            vDocNo = UCase(vDocNo)
            vDocdate = Trim(TXT102.Text)
            vTaxNo = Trim(TXT103.Text)
            vQuery = "insert into npmaster.dbo.TB_AP_ChangeDataPayment  (OldDocNo,OldDocDate,OldTaxNo,OldTaxDate,NewDocNo,NewDocDate,NewTaxNo,NewTaxDate,UserChange,DateChange) " _
            & " values('" & vDocNo & "','" & vDocdate & "','" & vTaxNo & "','" & vTaxDate1 & "','','','','" & vTaxDate & "','" & vUserID & "',getdate())"
            gConnection.Execute vQuery
        End If
    Else
    MsgBox "ไม่มีข้อมูลเลขที่เอกสาร " & vDocNo & " ในตาราง BCInputtax กรุณาตรวจสอบด้วยครับ"
    End If
    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub ChangeTaxDateBackOffice()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vTaxDate As String
Dim vTaxDate1 As String
Dim vCheckTax As Integer
Dim vTaxNo As String, vDocdate As String, vDocDate1 As String
Dim vVatRecordset As New ADODB.Recordset

On Error GoTo ErrDescription

If CHK101.Value = 0 Then
    vDocNo = Trim(TXT101.Text)
Else
    vDocNo = vChangeDoc1
End If

If CHK102.Value = 0 Then
    vTaxNo = Trim(TXT103.Text)
Else
    vTaxNo = Trim(TXT105.Text)
End If

'vQuery = "select docno from solar.bcvat.dbo.bcinputtax where docno = '" & vDocNo & "' "
'If OpenDataBaseBCVat(vVatConnection, vVatRecordset, vQuery) <> 0 Then
'vCheckTax = 1
'Else
'vCheckTax = 0
'End If
'vVatRecordset.Close

'If vCheckTax = 1 Then
 '   vTaxDate1 = Trim(TXT106.Text)
  '  vTaxDate = DTP102.Day & "/" & DTP102.Month & "/" & DTP102.Year
   '     If vTaxDate <> vTaxDate1 Then
    '
     '       vQuery = "set dateformat dmy"
      '      vVatConnection.Execute vQuery
       '
        '    vQuery = "update  solar.bcvat.dbo.bcinputtax set  taxdate = '" & vTaxDate & "' where  iscancel = 0 and docno = '" & vDocNo & "' and taxno = '" & vTaxNo & "' "
         '   vVatConnection.Execute vQuery
          '  MsgBox "เลขที่เอกสาร " & vDocNo & " ได้แก้ไขวันที่ใบกำกับภาษีเป็นวันที่ " & vTaxDate & " เรียบร้อยแล้ว  ที่ข้อมูลหลังร้าน"
        'Else
         '   MsgBox "ไม่ได้ใส่ข้อมูลการเปลี่ยนวันที่เอกสาร"
          '  Exit Sub
        'End If
    

vDocNo = UCase(vDocNo)
vDocdate = Trim(TXT102.Text)
vTaxNo = Trim(TXT103.Text)
vQuery = "Update npmaster.dbo.TB_AP_ChangeDataPayment  set mydescription = 'Update BCVat'   where  OldDocNo = '" & vDocNo & "'  and OldTaxNo = '" & vTaxNo & "' and oldTaxDate = '" & vTaxDate1 & "' and NewTaxDate = '" & vTaxDate & "' "
gConnection.Execute vQuery


    'Else
    'MsgBox "ไม่มีข้อมูลเลขที่เอกสาร " & vDocNo & " ในตาราง BCInputtax กรุณาตรวจสอบด้วยครับ"
    'End If
    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub
