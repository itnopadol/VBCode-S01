VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmOrder206 
   Caption         =   "Form206 กำหนดพนักงานขนส่ง"
   ClientHeight    =   8055
   ClientLeft      =   1950
   ClientTop       =   1035
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "FrmOrder206.frx":0000
   ScaleHeight     =   8055
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMD104 
      Height          =   330
      Left            =   4500
      Picture         =   "FrmOrder206.frx":9673
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "ออก"
      Top             =   6435
      Width           =   375
   End
   Begin VB.TextBox Text105 
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
      Left            =   3510
      TabIndex        =   21
      Top             =   3060
      Width           =   2175
   End
   Begin VB.CheckBox Check101 
      BackColor       =   &H8000000E&
      Caption         =   "ออก"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3510
      TabIndex        =   6
      Top             =   4590
      Width           =   1275
   End
   Begin MSComCtl2.DTPicker DTPicker101 
      Height          =   375
      Left            =   3510
      TabIndex        =   5
      Top             =   4050
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   20774913
      CurrentDate     =   38696
   End
   Begin VB.CommandButton CMD103 
      Caption         =   "ปรับปรุงพนักงาน"
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
      Left            =   3510
      TabIndex        =   11
      Top             =   495
      Width           =   1770
   End
   Begin VB.CommandButton CMD102 
      Height          =   330
      Left            =   4005
      Picture         =   "FrmOrder206.frx":99DD
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "ค้นหา ข้อมูล"
      Top             =   6435
      Width           =   375
   End
   Begin VB.CommandButton CMD101 
      Height          =   330
      Left            =   3510
      Picture         =   "FrmOrder206.frx":9DAA
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "บันทึกและปรับปรุงข้อมูล"
      Top             =   6435
      Width           =   375
   End
   Begin VB.TextBox Text107 
      Appearance      =   0  'Flat
      Height          =   1185
      Left            =   3510
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   5040
      Width           =   4965
   End
   Begin VB.TextBox Text106 
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
      Left            =   3510
      TabIndex        =   4
      Top             =   3555
      Width           =   2175
   End
   Begin VB.TextBox Text104 
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
      Left            =   3510
      TabIndex        =   3
      Top             =   2565
      Width           =   2175
   End
   Begin VB.TextBox Text103 
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
      Left            =   3510
      TabIndex        =   2
      Top             =   2070
      Width           =   3210
   End
   Begin VB.TextBox Text102 
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
      Left            =   3510
      TabIndex        =   1
      Top             =   1575
      Width           =   2175
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
      Left            =   3510
      TabIndex        =   0
      Top             =   1080
      Width           =   1320
   End
   Begin VB.Image Image102 
      Height          =   300
      Left            =   2790
      Picture         =   "FrmOrder206.frx":A0D1
      Top             =   495
      Width           =   570
   End
   Begin VB.Image Image101 
      Height          =   300
      Left            =   2790
      Picture         =   "FrmOrder206.frx":A60D
      Top             =   495
      Width           =   570
   End
   Begin VB.Label Label9 
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
      Height          =   375
      Left            =   990
      TabIndex        =   20
      Top             =   5040
      Width           =   2400
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "สถานภาพพนักงาน :"
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
      Height          =   285
      Left            =   900
      TabIndex        =   19
      Top             =   4590
      Width           =   2490
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่หมดอายุใบขับขี่ :"
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
      Left            =   1125
      TabIndex        =   18
      Top             =   4005
      Width           =   2265
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบขับขี่ :"
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
      Height          =   285
      Left            =   1980
      TabIndex        =   17
      Top             =   3555
      Width           =   1410
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ตำแหน่ง :"
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
      Left            =   2025
      TabIndex        =   16
      Top             =   3060
      Width           =   1365
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสพนักงาน :"
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
      Left            =   2070
      TabIndex        =   15
      Top             =   2565
      Width           =   1320
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อพนักงาน :"
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
      Height          =   285
      Left            =   2070
      TabIndex        =   14
      Top             =   2070
      Width           =   1320
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสพนักงาน :"
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
      TabIndex        =   13
      Top             =   1575
      Width           =   1905
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ID :"
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
      Left            =   2880
      TabIndex        =   12
      Top             =   1080
      Width           =   510
   End
End
Attribute VB_Name = "FrmOrder206"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vID As String
Dim vLicenceNumber As String
Dim vLicenceExpired As String
Dim vMydescription As String
    
On Error GoTo ErrDescription

If Text102.Text <> "" Then
    If vCheckEmpOpen = 1 Then
        vID = Trim(Text101.Text)
        vLicenceNumber = Trim(Text106.Text)
        vLicenceExpired = Trim(DTPicker101.Day & "/" & DTPicker101.Month & "/" & DTPicker101.Year)
        vMydescription = Trim(Text107.Text)
    
        vQuery = "exec bcnp.dbo.USP_DO_EmpBplusUpdate " & vID & ",'" & vLicenceNumber & "','" & vLicenceExpired & "'," _
                            & " '" & vMydescription & "','" & vUserID & "' "
        gConnection.Execute vQuery
        MsgBox "บันทึกการแก้ไขข้อมูลพนักงานขนส่งเรียบร้อยแล้ว", vbInformation, "Send Message"
        Text101.Text = ""
        Text102.Text = ""
        Text103.Text = ""
        Text104.Text = ""
        Text105.Text = ""
        Text106.Text = ""
        Text107.Text = ""
        DTPicker101 = Now
        Image101.Visible = True
        Image102.Visible = False
        Check101.Value = 0
        vCheckEmpOpen = 0
        
    End If
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD102_Click()
On Error GoTo ErrDescription

MDIFrmProgramPrint.Enabled = False
FrmOrder012.Show
vEmpModule = 2
FrmOrder012.ListView101.Checkboxes = False

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub CMD103_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

Call GetDataBPlus
vQuery = "exec bcnp.dbo.USP_DO_EmpBplusSynchronise"
gConnectionBPlus.Execute vQuery
MsgBox "ปรับปรุงข้อมูลพนักงานเรียบร้อย", vbInformation, "Send Message"
End Sub

Private Sub CMD104_Click()
    Unload FrmOrder206
End Sub

Private Sub Form_Load()
DTPicker101 = Now
Image101.Visible = True
Image102.Visible = False
End Sub

Public Sub CheckIsEmp()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vCode As String

vCode = Trim(Text102.Text)
vQuery = "select id from npmaster.dbo.TB_DO_EmpBplus Where code = '" & vCode & "'  "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckIsEmp = 1
Else
    vCheckIsEmp = 0
End If
vRecordset.Close
End Sub


