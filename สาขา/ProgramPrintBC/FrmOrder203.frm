VERSION 5.00
Begin VB.Form FrmOrder203 
   Caption         =   "Form203 กำหนด สถานที่ขนส่ง"
   ClientHeight    =   8010
   ClientLeft      =   1845
   ClientTop       =   870
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmOrder203.frx":0000
   ScaleHeight     =   8010
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMD105 
      Height          =   330
      Left            =   5040
      Picture         =   "FrmOrder203.frx":72FB
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "ออก"
      Top             =   5400
      Width           =   330
   End
   Begin VB.CommandButton CMD104 
      Height          =   330
      Left            =   4590
      Picture         =   "FrmOrder203.frx":7665
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "ยกเลิก เอกสาร"
      Top             =   5400
      Width           =   330
   End
   Begin VB.CommandButton CMD103 
      Height          =   330
      Left            =   4140
      Picture         =   "FrmOrder203.frx":9803
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "ค้นหา เอกสาร"
      Top             =   5400
      Width           =   330
   End
   Begin VB.CommandButton CMD102 
      Height          =   330
      Left            =   3690
      Picture         =   "FrmOrder203.frx":9BD0
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "บันทึกและปรับปรุงข้อมูล"
      Top             =   5400
      Width           =   330
   End
   Begin VB.CommandButton CMD101 
      Height          =   330
      Left            =   3240
      Picture         =   "FrmOrder203.frx":9EF7
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "เคลียร์ หน้าจอ"
      Top             =   5400
      Width           =   330
   End
   Begin VB.TextBox Text102 
      Appearance      =   0  'Flat
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
      Left            =   3285
      TabIndex        =   1
      Top             =   1845
      Width           =   3300
   End
   Begin VB.ComboBox CMB102 
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
      Left            =   3285
      TabIndex        =   3
      Text            =   "Combo2"
      Top             =   2880
      Width           =   2130
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
      Left            =   3285
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   2385
      Width           =   2130
   End
   Begin VB.TextBox Text103 
      Appearance      =   0  'Flat
      Height          =   1635
      Left            =   3285
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3375
      Width           =   5415
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
      Left            =   3285
      TabIndex        =   0
      Top             =   1350
      Width           =   2400
   End
   Begin VB.Image Image102 
      Height          =   300
      Left            =   3285
      Picture         =   "FrmOrder203.frx":A2DC
      Top             =   810
      Width           =   570
   End
   Begin VB.Image Image101 
      Height          =   300
      Left            =   3285
      Picture         =   "FrmOrder203.frx":A818
      Top             =   810
      Width           =   570
   End
   Begin VB.Label Label5 
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
      Height          =   330
      Left            =   1530
      TabIndex        =   14
      Top             =   3375
      Width           =   1545
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "จังหวัด :"
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
      Height          =   330
      Left            =   1440
      TabIndex        =   13
      Top             =   2880
      Width           =   1635
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "อำเภอ :"
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
      Height          =   330
      Left            =   1485
      TabIndex        =   12
      Top             =   2385
      Width           =   1590
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ตำบล :"
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
      Height          =   330
      Left            =   1575
      TabIndex        =   11
      Top             =   1845
      Width           =   1545
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ID สถานที่ :"
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
      Height          =   330
      Left            =   1620
      TabIndex        =   10
      Top             =   1350
      Width           =   1545
   End
End
Attribute VB_Name = "FrmOrder203"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
On Error GoTo ErrDescription

Text101.Text = ""
Text102.Text = ""
Text103.Text = ""
CMB101.Text = ""
CMB102.Text = ""
Image101.Visible = True
Image102.Visible = False
vCheckPlaceOpen = 0
Text102.SetFocus

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD102_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vID As String
Dim vDistrict As String
Dim vAmphur As String
Dim vProvince As String
Dim vMydescription As String
Dim vActiveStatus As String
    
On Error GoTo ErrDescription

If Text102.Text <> "" And CMB101.Text <> "" And CMB102.Text <> "" Then
    vDistrict = Trim(Text102.Text)
    vAmphur = Trim(CMB101.Text)
    vProvince = Trim(CMB102.Text)
    vActiveStatus = 1
    vMydescription = Trim(Text103.Text)
    
    If vCheckPlaceOpen = 0 Then
        Call CheckIsPlace
        If vCheckIsPlace = 1 Then
            MsgBox "สถานที่ขนส่งนี้มีอยู่แล้วกรุณาตรวจสอบ", vbCritical, "Send Error"
            Exit Sub
        End If
        vID = "Null"
    Else
        vID = Trim(Text101.Text)
    End If
    vQuery = "exec bcnp.dbo.USP_DO_PlaceUpdate " & vID & ",'" & vDistrict & "','" & vAmphur & "'," _
                        & " '" & vProvince & "','" & vMydescription & "','" & vActiveStatus & "','" & vUserID & "'"
    gConnection.Execute vQuery
    If vCheckPlaceOpen = 0 Then
        MsgBox "บันทึกข้อมูลเส้นทางขนส่งเรียบร้อยแล้ว", vbInformation, "Send Message"
    ElseIf vCheckPlaceOpen = 1 Then
        MsgBox "บันทึกการแก้ไขข้อมูลเส้นทางขนส่งเรียบร้อยแล้ว", vbInformation, "Send Message"
    End If

        Text101.Text = ""
        Text102.Text = ""
        Text103.Text = ""
        CMB101.Text = ""
        CMB102.Text = ""
        Image101.Visible = True
        Image102.Visible = False
        vCheckPlaceOpen = 0
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD103_Click()
On Error GoTo ErrDescription

MDIFrmProgramPrint.Enabled = False
FrmOrder008.Show
vPlaceModule = 2

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD104_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vID As Integer
Dim vAnswer As Integer

On Error GoTo ErrDescription

If Text101.Text <> "" Then
    If vCheckPlaceOpen = 1 Then
    vAnswer = MsgBox("คุณต้องการยกเลิก สถานที่ขนส่งนี้ใช่หรือไม่", vbYesNo, "Question Respond")
        If vAnswer = 6 Then
            vID = Trim(Text101.Text)
            vQuery = "Update npmaster.dbo.TB_DO_Place set activestatus = 0 where id = " & vID & " "
            gConnection.Execute vQuery
            Text101.Text = ""
            Text102.Text = ""
            Text103.Text = ""
            CMB101.Text = ""
            CMB102.Text = ""
            Image101.Visible = True
            Image102.Visible = False
            vCheckPlaceOpen = 0
        Else
            Exit Sub
        End If
    End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD105_Click()
Unload FrmOrder203
End Sub

Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error GoTo ErrDescription

CMB101.Clear
vQuery = "select distinct isnull(amphur,'') as amphur from npmaster.dbo.TB_DO_Place  where amphur is not null order by amphur"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        CMB101.AddItem Trim(vRecordset.Fields("amphur").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

CMB102.Clear
vQuery = "select distinct isnull(province,'') as province  from npmaster.dbo.TB_DO_Place where province is not null  order by province"
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        CMB102.AddItem Trim(vRecordset.Fields("province").Value)
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close
Image101.Visible = True
Image102.Visible = False

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Public Sub CheckIsPlace()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDistrict  As String
Dim vAmphur As String
Dim vProvince As String

On Error GoTo ErrDescription

vDistrict = Trim(Text102.Text)
vAmphur = Trim(CMB101.Text)
vProvince = Trim(CMB102.Text)
vQuery = "select id from npmaster.dbo.TB_DO_Place Where District = '" & vDistrict & "' and Amphur = '" & vAmphur & "' and Province = '" & vProvince & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckIsPlace = 1
Else
    vCheckIsPlace = 0
End If
vRecordset.Close

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub



