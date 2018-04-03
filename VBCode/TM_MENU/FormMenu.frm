VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "NP_Menu"
   ClientHeight    =   8145
   ClientLeft      =   4500
   ClientTop       =   1200
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   Picture         =   "FormMenu.frx":0000
   ScaleHeight     =   8145
   ScaleWidth      =   8505
   Begin VB.CommandButton btnRefreshMenu 
      Caption         =   "กดฟื้นฟูโปรแกรม"
      Height          =   345
      Left            =   6525
      TabIndex        =   2
      Top             =   375
      Width           =   1800
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "ออกจากโปรแกรม"
      Height          =   345
      Left            =   4425
      TabIndex        =   1
      Top             =   375
      Width           =   1800
   End
   Begin VB.CommandButton btnLink 
      Height          =   465
      Index           =   0
      Left            =   825
      TabIndex        =   0
      Top             =   1200
      Width           =   2655
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Initialize()
    Dim vQuery As String
    Dim vRecordset As New ADODB.Recordset
    Dim vCount As Integer
    Dim vTop(12) As Integer
    
    On Error Resume Next
    vQuery = "select * from TMMENU where PROGSERVER = 'Lunar' "
    gNumOfSoftware = OpenTable(gConnection, vRecordset, vQuery)
    If gNumOfSoftware > 0 Then
        ReDim gSoftwareLink(gNumOfSoftware)
        vRecordset.MoveFirst
        btnLink(0).Caption = GetField(vRecordset, "PROGNAME")
        gSoftwareLink(0) = GetField(vRecordset, "PROGLOCN")
        btnLink(0).Left = 600
        btnLink(0).Top = 1200
        btnLink(0).Width = 3000
        btnLink(0).Height = 400

        vRecordset.MoveNext
        vCount = 1
        While Not vRecordset.EOF
            If vCount < 12 Then
            Load btnLink(vCount)
            Set btnLink(vCount).Container = frmMenu
            btnLink(vCount).Visible = True
            btnLink(vCount).Top = btnLink(vCount - 1).Top + 550
            btnLink(vCount).Left = btnLink(0).Left
            btnLink(vCount).Width = btnLink(0).Width
            btnLink(vCount).Height = btnLink(0).Height
            btnLink(vCount).Caption = GetField(vRecordset, "PROGNAME")
            gSoftwareLink(vCount) = GetField(vRecordset, "PROGLOCN")
            vRecordset.MoveNext
            vCount = vCount + 1
            ElseIf vCount = 12 Then
            Load btnLink(vCount)
            Set btnLink(vCount).Container = frmMenu
            btnLink(vCount).Visible = True
            btnLink(vCount).Top = 1200 'vTop(1)
            btnLink(vCount).Left = 5000
            btnLink(vCount).Width = btnLink(1).Width
            btnLink(vCount).Height = btnLink(1).Height
            btnLink(vCount).Caption = GetField(vRecordset, "PROGNAME")
            gSoftwareLink(vCount) = GetField(vRecordset, "PROGLOCN")
            vRecordset.MoveNext
            vCount = vCount + 1
            Else
            Load btnLink(vCount)
            Set btnLink(vCount).Container = frmMenu
            btnLink(vCount).Visible = True
            btnLink(vCount).Top = btnLink(vCount - 1).Top + 550
            btnLink(vCount).Left = 5000
            btnLink(vCount).Width = btnLink(1).Width
            btnLink(vCount).Height = btnLink(1).Height
            btnLink(vCount).Caption = GetField(vRecordset, "PROGNAME")
            gSoftwareLink(vCount) = GetField(vRecordset, "PROGLOCN")
            vRecordset.MoveNext
            vCount = vCount + 1
            End If
        Wend
    End If
End Sub

Private Sub btnExit_Click()
    gConnection.Close
    End
End Sub

Private Sub btnLink_Click(Index As Integer)
    On Error GoTo CallError
    gConnection.Close
    Call Connect
    Shell gSoftwareLink(Index)
    frmMenu.WindowState = 0
    Exit Sub
CallError:
    MsgBox Err.Description
End Sub

Private Sub btnRefreshMenu_Click()
    Call Initialize
End Sub

Private Sub Form_Load()
    frmMenu.Width = 8690 '4365= 8730
    frmMenu.Height = 8650
    gConnection.CursorLocation = adUseClient
    gConnection.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;" & _
        "Persist Security Info=False;Initial Catalog=NPDEV;Data Source=Nebula"
    gConnection.Open
    Show
    Call Initialize
End Sub

Public Sub Connect()
    gConnection.CursorLocation = adUseClient
    gConnection.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;" & _
        "Persist Security Info=False;Initial Catalog=NPDEV;Data Source=Nebula"
    gConnection.Open
End Sub
