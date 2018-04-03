VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormSendResult 
   Caption         =   "สรุปผลการจัดส่งสินค้า"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FormSendResult.frx":0000
   ScaleHeight     =   8055
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMD101 
      Caption         =   "ตกลง"
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
      Left            =   9180
      TabIndex        =   3
      Top             =   6525
      Width           =   1185
   End
   Begin VB.TextBox Text101 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1530
      TabIndex        =   2
      Top             =   1170
      Width           =   2175
   End
   Begin MSComctlLib.ListView ListView101 
      Height          =   4560
      Left            =   360
      TabIndex        =   0
      Top             =   1755
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   8043
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับ"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "เลขที่ใบจัดคิว"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "เงินที่ต้องเก็บ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "เงินที่เก็บได้"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ผลการเก็บเงิน"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "วันที่เก็บเงิน"
         Object.Width           =   3705
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบขนส่ง"
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
      Left            =   360
      TabIndex        =   1
      Top             =   1215
      Width           =   1140
   End
End
Attribute VB_Name = "FormSendResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD101_Click()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDeliveryNo As String
Dim i As Integer
Dim vSendResult As String
Dim vHoardNet As Currency
Dim vHoardResult As String

On Error GoTo ErrDescription

If ListView101.ListItems.Count <> 0 Then
    For i = 1 To ListView101.ListItems.Count
    vDeliveryNo = Trim(ListView101.ListItems.Item(i).SubItems(1))
    vSendResult = Trim(ListView101.ListItems.Item(i).SubItems(2))
    vHoardNet = Trim(ListView101.ListItems.Item(i).SubItems(3))
    vHoardResult = Trim(ListView101.ListItems.Item(i).SubItems(4))
    vQuery = "exec dbo.USP_DO_UpdateDeliveryResult '" & vDeliveryNo & "'," & vHoardNet & ",'" & vHoardResult & "' "
    gConnection.Execute vQuery
    Next i
    ListView101.ListItems.Clear
    Unload FormSendResult
End If

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
FormDelivery.Enabled = True
End Sub

Private Sub ListView101_DblClick()
Dim i As Integer
Dim vQueueID As String
Dim vHoardNet As Currency
Dim vHoardAmount As Currency
Dim vDeliveryNo As String
Dim vHoardNetCur As Currency
Dim vHoardDate As Date

On Error GoTo ErrDescription

i = ListView101.SelectedItem.Index
vQueueID = ListView101.ListItems.Item(i).SubItems(1)
vDeliveryNo = Trim(Text101.Text)
vHoardNet = ListView101.ListItems.Item(i).SubItems(3)
vHoardAmount = ListView101.ListItems.Item(i).SubItems(2)
FormInsertResult.Text101.Text = vDeliveryNo
FormInsertResult.Text102.Text = vQueueID
FormInsertResult.Text103.Text = vHoardAmount
FormInsertResult.Text104.Text = vHoardNet
If ListView101.ListItems.Item(i).SubItems(3) <> "" Then
    vHoardNetCur = ListView101.ListItems.Item(i).SubItems(2)
    FormInsertResult.Text104.Text = vHoardNetCur
    Select Case ListView101.ListItems.Item(i).SubItems(4)
    Case 0
        FormInsertResult.Option101.Value = True
        FormInsertResult.Option102.Value = False
        FormInsertResult.Option103.Value = False
        FormInsertResult.Option104.Value = False
        FormInsertResult.Option105.Value = False
    Case 1
        FormInsertResult.Option101.Value = False
        FormInsertResult.Option102.Value = True
        FormInsertResult.Option103.Value = False
        FormInsertResult.Option104.Value = False
        FormInsertResult.Option105.Value = False
    Case 2
        FormInsertResult.Option101.Value = False
        FormInsertResult.Option102.Value = False
        FormInsertResult.Option103.Value = True
        FormInsertResult.Option104.Value = False
        FormInsertResult.Option105.Value = False
    Case 3
        FormInsertResult.Option101.Value = False
        FormInsertResult.Option102.Value = False
        FormInsertResult.Option103.Value = False
        FormInsertResult.Option104.Value = True
        FormInsertResult.Option105.Value = False
    Case 4
        FormInsertResult.Option101.Value = False
        FormInsertResult.Option102.Value = False
        FormInsertResult.Option103.Value = False
        FormInsertResult.Option104.Value = False
        FormInsertResult.Option105.Value = True
    End Select
    'Select Case ListView101.ListItems.Item(i).SubItems(5)
    'Case 0
     '   FormInsertResult.Option201.Value = True
      '  FormInsertResult.Option202.Value = False
       ' FormInsertResult.Option203.Value = False
        'FormInsertResult.Option204.Value = False
        'FormInsertResult.Option205.Value = False
        'FormInsertResult.Option206.Value = False
    'Case 1
     '   FormInsertResult.Option201.Value = False
      '  FormInsertResult.Option202.Value = True
       ' FormInsertResult.Option203.Value = False
        'FormInsertResult.Option204.Value = False
        'FormInsertResult.Option205.Value = False
        'FormInsertResult.Option206.Value = False
    'Case 2
     '   FormInsertResult.Option201.Value = False
      '  FormInsertResult.Option202.Value = False
       ' FormInsertResult.Option203.Value = True
        'FormInsertResult.Option204.Value = False
        'FormInsertResult.Option205.Value = False
        'FormInsertResult.Option206.Value = False
    'Case 3
     '   FormInsertResult.Option201.Value = False
      '  FormInsertResult.Option202.Value = False
       ' FormInsertResult.Option203.Value = False
        'FormInsertResult.Option204.Value = True
        'FormInsertResult.Option205.Value = False
        'FormInsertResult.Option206.Value = False
    'Case 4
     '   FormInsertResult.Option201.Value = False
      '  FormInsertResult.Option202.Value = False
       ' FormInsertResult.Option203.Value = False
        'FormInsertResult.Option204.Value = False
        'FormInsertResult.Option205.Value = True
        'FormInsertResult.Option206.Value = False
    'Case 5
      '  FormInsertResult.Option201.Value = False
     ' '  FormInsertResult.Option202.Value = False
        'FormInsertResult.Option203.Value = False
        'FormInsertResult.Option204.Value = False
        'FormInsertResult.Option205.Value = False
        'FormInsertResult.Option206.Value = True
    'End Select
End If
FormInsertResult.Show

ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description, vbCritical, "Send Error"
    Exit Sub
End If
End Sub
