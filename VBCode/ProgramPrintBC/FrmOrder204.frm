VERSION 5.00
Begin VB.Form FrmOrder204 
   Caption         =   "Form204 กำหนด เส้นทางขนส่ง"
   ClientHeight    =   7725
   ClientLeft      =   1740
   ClientTop       =   645
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmOrder204.frx":0000
   ScaleHeight     =   7725
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMD105 
      Height          =   330
      Left            =   4995
      Picture         =   "FrmOrder204.frx":9673
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "ออก"
      Top             =   4500
      Width           =   330
   End
   Begin VB.CommandButton CMD104 
      Height          =   330
      Left            =   4545
      Picture         =   "FrmOrder204.frx":99DD
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "ยกเลิก ข้อมูล"
      Top             =   4500
      Width           =   330
   End
   Begin VB.CommandButton CMD103 
      Height          =   330
      Left            =   4095
      Picture         =   "FrmOrder204.frx":BB7B
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "ค้นหา ข้อมูล"
      Top             =   4500
      Width           =   330
   End
   Begin VB.CommandButton CMD102 
      Height          =   330
      Left            =   3645
      Picture         =   "FrmOrder204.frx":BF48
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "บันทึกและปรับปรุงข้อมูล"
      Top             =   4500
      Width           =   330
   End
   Begin VB.CommandButton CMD101 
      Height          =   330
      Left            =   3195
      Picture         =   "FrmOrder204.frx":C26F
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "เคลียร์ หน้าจอ"
      Top             =   4500
      Width           =   330
   End
   Begin VB.TextBox Text104 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000012&
      Height          =   1185
      Left            =   3240
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2835
      Width           =   5955
   End
   Begin VB.TextBox Text103 
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
      ForeColor       =   &H80000012&
      Height          =   330
      Left            =   3240
      TabIndex        =   2
      Top             =   2340
      Width           =   2850
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
      ForeColor       =   &H80000012&
      Height          =   330
      Left            =   3240
      TabIndex        =   1
      Top             =   1845
      Width           =   2850
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
      ForeColor       =   &H80000012&
      Height          =   330
      Left            =   3240
      TabIndex        =   0
      Top             =   1350
      Width           =   1455
   End
   Begin VB.Image Image102 
      Height          =   300
      Left            =   3240
      Picture         =   "FrmOrder204.frx":C654
      Top             =   810
      Width           =   570
   End
   Begin VB.Image Image101 
      Height          =   300
      Left            =   3240
      Picture         =   "FrmOrder204.frx":CB90
      Top             =   810
      Width           =   570
   End
   Begin VB.Label Label4 
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
      Height          =   285
      Left            =   2025
      TabIndex        =   12
      Top             =   2835
      Width           =   1140
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อ 2 :"
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
      Left            =   1755
      TabIndex        =   11
      Top             =   2340
      Width           =   1410
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อ 1 :"
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
      TabIndex        =   10
      Top             =   1845
      Width           =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ID เส้นทาง :"
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
      Left            =   1845
      TabIndex        =   9
      Top             =   1350
      Width           =   1320
   End
End
Attribute VB_Name = "FrmOrder204"
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
Text104.Text = ""
Image101.Visible = True
Image102.Visible = False
vCheckRouteOpen = 0

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
Dim vName1 As String
Dim vName2 As String
Dim vMydescription As String
Dim vActiveStatus As Integer
    
On Error GoTo ErrDescription

If Text102.Text <> "" Then
    vName1 = Trim(Text102.Text)
    vName2 = Trim(Text103.Text)
    vActiveStatus = 1
    vMydescription = Trim(Text104.Text)
    
    If vCheckRouteOpen = 0 Then
        Call CheckIsRoute
        If vCheckIsRoute = 1 Then
            MsgBox "สถานที่ขนส่งนี้มีอยู่แล้วกรุณาตรวจสอบ", vbCritical, "Send Error"
            Exit Sub
        End If
        vID = "Null"
    Else
        vID = Trim(Text101.Text)
    End If
    vQuery = "exec bcnp.dbo.USP_DO_RouteUpdate " & vID & ",'" & vName1 & "','" & vName2 & "'," _
                        & " '" & vMydescription & "','" & vActiveStatus & "','" & vUserID & "'"
    gConnection.Execute vQuery
    If vCheckRouteOpen = 0 Then
        MsgBox "บันทึกข้อมูลเส้นทางขนส่งเรียบร้อยแล้ว", vbInformation, "Send Message"
    ElseIf vCheckRouteOpen = 1 Then
        MsgBox "บันทึกการแก้ไขข้อมูลเส้นทางขนส่งเรียบร้อยแล้ว", vbInformation, "Send Message"
    End If

    Text101.Text = ""
    Text102.Text = ""
    Text103.Text = ""
    Text104.Text = ""
    Image101.Visible = True
    Image102.Visible = False
    vCheckRouteOpen = 0
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
FrmOrder009.Show
vRouteModule = 2

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
    If vCheckRouteOpen = 1 Then
    vAnswer = MsgBox("คุณต้องการยกเลิก เส้นทางขนส่งนี้ใช่หรือไม่", vbYesNo, "Question Respond")
        If vAnswer = 6 Then
            vID = Trim(Text101.Text)
            vQuery = "Update npmaster.dbo.TB_DO_Route set activestatus = 0 where id = " & vID & " "
            gConnection.Execute vQuery
            Text101.Text = ""
            Text102.Text = ""
            Text103.Text = ""
            Text104.Text = ""
            Image101.Visible = True
            Image102.Visible = False
            vCheckRouteOpen = 0
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
Unload FrmOrder204
End Sub

Private Sub Form_Load()
Image101.Visible = True
Image102.Visible = False
End Sub

Public Sub CheckIsRoute()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vName1 As String
Dim vName2 As String

On Error GoTo ErrDescription

vName1 = Trim(Text102.Text)
vName2 = Trim(Text103.Text)
vQuery = "select id from npmaster.dbo.TB_DO_Route Where name1 = '" & vName1 & "' and name2 = '" & vName2 & "' "
If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
    vCheckIsRoute = 1
Else
    vCheckIsRoute = 0
End If
vRecordset.Close

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub


