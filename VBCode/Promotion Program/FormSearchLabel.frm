VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormSearchLabel 
   Caption         =   "ค้นหา ทะเบียนฟอร์มป้ายราคาโปรโมชั่น"
   ClientHeight    =   4890
   ClientLeft      =   3570
   ClientTop       =   2310
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   9705
   Begin MSComctlLib.ListView ListView102 
      Height          =   2130
      Left            =   1080
      TabIndex        =   2
      Top             =   765
      Visible         =   0   'False
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   3757
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ชื่อฟอร์มป้ายราคา"
         Object.Width           =   12347
      EndProperty
   End
   Begin VB.Frame Frame101 
      Caption         =   "ขนาดป้าย "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   90
      TabIndex        =   4
      Top             =   1215
      Width           =   4875
      Begin VB.OptionButton Option103 
         Caption         =   "P3 "
         Height          =   285
         Left            =   2745
         TabIndex        =   8
         Top             =   450
         Width           =   1455
      End
      Begin VB.OptionButton Option102 
         Caption         =   "P2"
         Height          =   285
         Left            =   1125
         TabIndex        =   7
         Top             =   945
         Width           =   1140
      End
      Begin VB.OptionButton Option104 
         Caption         =   "P4 "
         Height          =   285
         Left            =   2745
         TabIndex        =   6
         Top             =   945
         Width           =   1500
      End
      Begin VB.OptionButton Option101 
         Caption         =   "P1 "
         Height          =   285
         Left            =   1125
         TabIndex        =   5
         Top             =   450
         Value           =   -1  'True
         Width           =   1635
      End
   End
   Begin VB.CommandButton CMD102 
      Caption         =   "ตกลง"
      Height          =   420
      Left            =   4005
      TabIndex        =   10
      Top             =   2970
      Width           =   960
   End
   Begin VB.CommandButton CMD101 
      Height          =   330
      Left            =   4635
      Picture         =   "FormSearchLabel.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   450
      Width           =   330
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   1080
      TabIndex        =   3
      Top             =   450
      Width           =   3525
   End
   Begin VB.Label Label3 
      Caption         =   "ชื่อฟอร์ม :"
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
      Left            =   225
      TabIndex        =   1
      Top             =   450
      Width           =   825
   End
   Begin VB.Label Label2 
      Caption         =   "รูปตัวอย่างป้าย"
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
      Left            =   5175
      TabIndex        =   0
      Top             =   150
      Width           =   1515
   End
   Begin VB.Image Image101 
      Height          =   2955
      Left            =   5175
      Top             =   450
      Width           =   4455
   End
End
Attribute VB_Name = "FormSearchLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vQuery As String


Private Sub CMD101_Click()
ListView102.Visible = True
End Sub

Private Sub CMD102_Click()
Dim vLabID As String
Dim vRecordset As New ADODB.Recordset
Dim vForm As String
Dim vSize As String

On Error Resume Next

If Option101.Value = True Then
  vSize = "P1"
ElseIf Option102.Value = True Then
  vSize = "P2"
ElseIf Option103.Value = True Then
  vSize = "P3"
ElseIf Option104.Value = True Then
  vSize = "P4"
End If

Form103.Text103.Text = Trim(Text101.Text)
vForm = Left(Trim(Text101.Text), InStr(Trim(Text101.Text), ":") - 1)

vQuery = "exec dbo.USP_PM_SearchLabelPromotionMaster '" & vForm & "','" & vSize & "' "
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
  vFormName = Trim(vRecordset.Fields("labid").Value)
End If
vRecordset.Close

MDIForm1.Enabled = True
Unload FormSearchLabel
End Sub

Private Sub Form_Load()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vListLabel As ListItem

On Error GoTo ErrDescription

vQuery = "execute USP_PM_Label "
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
    vRecordset.MoveFirst
    While Not vRecordset.EOF
        Set vListLabel = ListView102.ListItems.Add(, , Trim(vRecordset.Fields("labname").Value))
    vRecordset.MoveNext
    Wend
End If
vRecordset.Close

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIForm1.Enabled = True
End Sub


Private Sub ListView101_DblClick()
Dim vIndex As Integer

On Error Resume Next

    Form103.Text103.Text = Trim(Text101.Text)
    vFormName = Trim(ListView102.SelectedItem.Text)
    MDIForm1.Enabled = True
    Unload FormSearchLabel
    
End Sub

Private Sub ListView101_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

On Error GoTo ErrDescription

vQuery = "select * from npmaster.dbo.TB_PM_Label where labid = '" & Item.Text & "' "
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
    If Not IsNull(Trim(vRecordset.Fields("labelpict").Value)) Then
    Image101.Picture = LoadPicture(Trim(vRecordset.Fields("labelpict").Value))
    End If
End If
vRecordset.Close

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub ListView102_DblClick()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String

Text101.Text = Trim(ListView102.SelectedItem.Text)
ListView102.Visible = False

vQuery = "select * from npmaster.dbo.TB_PM_Label where  labname  like  '%" & Trim(ListView102.SelectedItem.Text) & "%' "
If OpenDatabase(gConnection, vRecordset, vQuery) <> 0 Then
    If Not IsNull(Trim(vRecordset.Fields("labelpict").Value)) Then
    Image101.Picture = LoadPicture(Trim(vRecordset.Fields("labelpict").Value))
    End If
End If
vRecordset.Close

Select Case ListView102.SelectedItem.Index
Case 1
  Option101.Enabled = True
  Option102.Enabled = True
  Option103.Enabled = True
  Option104.Enabled = True
Case 2
  Option101.Enabled = True
  Option102.Enabled = True
  Option103.Enabled = True
  Option104.Enabled = True
Case 3
  Option101.Enabled = True
  Option102.Enabled = True
  Option103.Enabled = False
  Option104.Enabled = False
Case 4
  Option101.Enabled = False
  Option102.Enabled = True
  Option103.Enabled = True
  Option104.Enabled = True
Case 5
  Option101.Enabled = False
  Option102.Enabled = True
  Option103.Enabled = True
  Option104.Enabled = True
Case 6
  Option101.Enabled = False
  Option102.Enabled = True
  Option103.Enabled = True
  Option104.Enabled = True
Case 7
  Option101.Enabled = False
  Option102.Enabled = True
  Option103.Enabled = True
  Option104.Enabled = True
End Select

End Sub

