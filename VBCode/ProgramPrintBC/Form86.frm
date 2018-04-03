VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form86 
   Caption         =   "พิมพ์ใบขอโอนสินค้า"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form86.frx":0000
   ScaleHeight     =   8385
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport861 
      Left            =   6000
      Top             =   6960
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
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000009&
      Caption         =   "ออกกระดาษ A4"
      Height          =   390
      Left            =   2550
      TabIndex        =   6
      Top             =   1200
      Width           =   2565
   End
   Begin VB.TextBox Text1 
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
      Left            =   1950
      TabIndex        =   5
      Text            =   "พิมพ์ใบขอโอนสินค้า"
      Top             =   6450
      Width           =   2340
   End
   Begin VB.CommandButton CMD861 
      Caption         =   "Print"
      Height          =   690
      Left            =   8235
      TabIndex        =   4
      Top             =   5940
      Width           =   1815
   End
   Begin VB.TextBox TXT861 
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
      Left            =   1950
      TabIndex        =   3
      Top             =   5925
      Width           =   2340
   End
   Begin VB.CommandButton CMD862 
      Caption         =   "RefreshData"
      Height          =   390
      Left            =   975
      TabIndex        =   1
      Top             =   1200
      Width           =   1515
   End
   Begin MSComctlLib.ListView ListView861 
      Height          =   3840
      Left            =   975
      TabIndex        =   0
      Top             =   1725
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   6773
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "เลขที่ใบขอโอน"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "วันที่ทำเอกสาร"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ประเภทเอกสาร"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "หัวเอกสาร"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ผู้ทำรายการ"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "สถานะการพิมพ์"
         Object.Width           =   2205
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่เอกสาร"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   975
      TabIndex        =   2
      Top             =   5925
      Width           =   990
   End
End
Attribute VB_Name = "Form86"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD861_Click()

On Error GoTo ErrDescription

If TXT861.Text <> "" Then
    If Check1.Value = 0 Then
        Call PrintStockTransfer
    Else
        Call PrintStockTransfer_A4
    End If
Else
MsgBox "ไม่มีเลขที่ให้พิมพ์ หรือ ไม่ได้เลือกฟอร์มที่จะพิมพ์ หรือ ไม่ได้เลือกเลขที่เอกสาร", vbCritical, "ข้อความแจ้งเตือน"
End If
TXT861.Text = ""

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD862_Click()
Call RefreshData
End Sub

Private Sub ListView861_ItemClick(ByVal Item As MSComctlLib.ListItem)
TXT861.Text = Item
End Sub

Function RefreshData()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim aaa As Integer
Dim i As Integer
Dim DocListItem As ListItem
Dim vDocNo, vNewDoc As String
Dim vPrintStatus As Integer
Dim CountRecordset As Integer, CountList As Integer
Dim vDocHeader As String

On Error Resume Next

            vDocHeader = "TF"

            vQuery = "Select  Docno,LastPrintDateTime,printed from NPMaster.dbo.NPPrintServer   where Printed = 0 " _
                        & " and DoctypeID = '" & vDocHeader & "' order by LastPrintDateTime "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    If Not vRecordset.EOF Then
                        CountRecordset = vRecordset.RecordCount
                        CountList = ListView861.ListItems.Count
                                If CountRecordset > CountList Then
                                vRecordset.MoveFirst
                                For i = 1 To CountRecordset
                                If i < CountRecordset Then
                                vDocNo = ListView861.ListItems.Item(i).Text
                                End If
                                vNewDoc = Trim(vRecordset.Fields("Docno").Value)
                                vPrintStatus = Trim(vRecordset.Fields("Printed").Value)
                                        If vDocNo <> vNewDoc Then
                                        Set DocListItem = ListView861.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
                                        DocListItem.SubItems(1) = Trim(vRecordset.Fields("LastPrintDateTime").Value)
                                        DocListItem.SubItems(2) = Trim(vRecordset.Fields("DocTypeID").Value)
                                        DocListItem.SubItems(3) = Trim(vRecordset.Fields("GroupDoc").Value)
                                        DocListItem.SubItems(4) = Trim(vRecordset.Fields("Code").Value)
                                        DocListItem.SubItems(5) = Trim(vRecordset.Fields("Printed").Value)
                                        End If
                                vRecordset.MoveNext
                            Next i
                            ElseIf CountRecordset < CountList Then
                            Call NewListItems
                            End If
                    End If
        End If
        vRecordset.Close
'------------------------------------------------------------------------------------------------

'ErrDescription:
'If Err.Description <> "" Then
'MsgBox Err.Description
'End If
End Function

Function NewListItems()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim INVItemLists As ListItem
Dim vTypeDoc As String

On Error GoTo ErrDescription
    vTypeDoc = "TF"
    ListView861.ListItems.Clear
    vQuery = "Select Docno,LastPrintDateTime  from NPMaster.dbo.NPPrintServer  where Printed = 0  " _
                        & " and LastPrintedUser = '" & vUserID & "' and DoctypeID = '" & vTypeDoc & "'  order by LastPrintDateTime "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            If Not vRecordset.EOF Then
                vRecordset.MoveFirst
                While Not vRecordset.EOF
                Set INVItemLists = ListView861.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
                INVItemLists.SubItems(1) = Trim(vRecordset.Fields("LastPrintDateTime").Value)
                INVItemLists.SubItems(2) = Trim(vRecordset.Fields("DocTypeID").Value)
                INVItemLists.SubItems(3) = Trim(vRecordset.Fields("GroupDoc").Value)
                INVItemLists.SubItems(4) = Trim(vRecordset.Fields("Code").Value)
                INVItemLists.SubItems(5) = Trim(vRecordset.Fields("Printed").Value)
                vRecordset.MoveNext
                Wend
            End If
        End If
        vRecordset.Close

        '----------------------------------------------------------------------------------------------------------------------
        
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Function

Private Sub Form_Load()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim DBListItems As ListItem
Dim DBItemforms As ListItem
Dim vTypeDoc As String

On Error GoTo ErrDescription

 vTypeDoc = "TF"

ListView861.ListItems.Clear
 vQuery = "Select Docno,LastPrintDateTime,DocTypeID,GroupDoc,Code,Printed  from NPMaster.dbo.NPPrintServer  where Printed = 0  " _
                        & " and DoctypeID = '" & vTypeDoc & "'  order by LastPrintDateTime "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            If Not vRecordset.EOF Then
                vRecordset.MoveFirst
                While Not vRecordset.EOF
                Set DBListItems = ListView861.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
                DBListItems.SubItems(1) = Trim(vRecordset.Fields("LastPrintDateTime").Value)
                DBListItems.SubItems(2) = Trim(vRecordset.Fields("DocTypeID").Value)
                DBListItems.SubItems(3) = Trim(vRecordset.Fields("GroupDoc").Value)
                DBListItems.SubItems(4) = Trim(vRecordset.Fields("Code").Value)
                DBListItems.SubItems(5) = Trim(vRecordset.Fields("Printed").Value)
                vRecordset.MoveNext
                Wend
            End If
        End If
        vRecordset.Close
        '--------------------------------------------------------------------------------------------------------------------
   
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If

End Sub

Public Sub PrintStockTransfer()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vGroupDoc As String
Dim vReportName As String
Dim vPrint As Integer
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

        vDocNo = Trim(TXT861.Text)

        vQuery = "select groupdoc,printed from  npmaster.dbo.npprintserver where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
                vPrint = Trim(vRecordset.Fields("Printed").Value)
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------
        
                            vRepType = "TF"
                            vRepID = 155
                            
                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport861
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            '---------------------------------------------------------------------------------------------------
                            
                             If vPrint = 0 Then
                                            vQuery = "Update npmaster.dbo.npprintserver set Printed = 1 where Docno = '" & vDocNo & "' "
                                            gConnection.Execute vQuery
                                                ListView861.ListItems.Remove (ListView861.SelectedItem.Index)
                                    End If
                                             
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Public Sub PrintStockTransfer_A4()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vGroupDoc As String
Dim vReportName As String
Dim vPrint As Integer
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

        vDocNo = Trim(TXT861.Text)

        vQuery = "select groupdoc,printed from  npmaster.dbo.npprintserver where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
                vPrint = Trim(vRecordset.Fields("Printed").Value)
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------
        
                            vRepType = "TF"
                            vRepID = 167
                            
                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport861
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            '---------------------------------------------------------------------------------------------------
                            
                             If vPrint = 0 Then
                                            vQuery = "Update npmaster.dbo.npprintserver set Printed = 1 where Docno = '" & vDocNo & "' "
                                            gConnection.Execute vQuery
                                                ListView861.ListItems.Remove (ListView861.SelectedItem.Index)
                                    End If
                                             
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

