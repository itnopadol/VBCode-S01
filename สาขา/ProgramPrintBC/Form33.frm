VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form33 
   Caption         =   "หน้าพิมพ์ใบมัดจำ"
   ClientHeight    =   8325
   ClientLeft      =   2580
   ClientTop       =   1335
   ClientWidth     =   12000
   Icon            =   "Form33.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "Form33.frx":08CA
   ScaleHeight     =   8325
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport331 
      Left            =   5160
      Top             =   6240
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
   Begin VB.CommandButton CMD332 
      Caption         =   "RefreshData"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4305
      TabIndex        =   9
      Top             =   1065
      Width           =   1710
   End
   Begin VB.Timer Timer331 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   9750
      Top             =   7725
   End
   Begin VB.TextBox TXT332 
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
      Left            =   2100
      TabIndex        =   4
      Top             =   6225
      Width           =   2415
   End
   Begin VB.TextBox TXT331 
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
      Left            =   2100
      TabIndex        =   3
      Top             =   5550
      Width           =   2415
   End
   Begin VB.CommandButton CMD331 
      Caption         =   "พิมพ์ใบมัดจำ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   8775
      TabIndex        =   2
      Top             =   5535
      Width           =   1665
   End
   Begin MSComctlLib.ListView ListView332 
      Height          =   3240
      Left            =   6825
      TabIndex        =   1
      Top             =   1650
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   5715
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ฟอร์มที่พิมพ์"
         Object.Width           =   6227
      EndProperty
   End
   Begin MSComctlLib.ListView ListView331 
      Height          =   3240
      Left            =   1050
      TabIndex        =   0
      Top             =   1650
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   5715
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "เลขที่ใบมัดจำ"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "วันที่ทำเอกสาร"
         Object.Width           =   5115
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "พิมพ์ใบมัดจำ"
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
      TabIndex        =   8
      Top             =   300
      Width           =   7365
   End
   Begin VB.Label LBLTime331 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   9000
      TabIndex        =   7
      Top             =   1350
      Width           =   1365
   End
   Begin VB.Label LBL332 
      BackStyle       =   0  'Transparent
      Caption         =   "ฟอร์มที่พิมพ์"
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
      Height          =   390
      Left            =   975
      TabIndex        =   6
      Top             =   6225
      Width           =   1065
   End
   Begin VB.Label LBL331 
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบมัดจำ"
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
      Left            =   975
      TabIndex        =   5
      Top             =   5550
      Width           =   1065
   End
End
Attribute VB_Name = "Form33"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD331_Click()
Dim vQuery As String
Dim vDocNo As String
Dim vPrint As Integer
Dim vRecordset As New ADODB.Recordset

On Error GoTo ErrDescription

vDocNo = Trim(TXT331.Text)
If TXT331.Text <> "" And TXT332.Text <> "" Then
    If TXT332.Text = "พิมพ์ใบมัดจำ" Then
        Call PrintDeposit
    End If
    vQuery = "select printed from  bcnp.dbo.vw_sl_00001 where docno = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vPrint = Trim(vRecordset.Fields("printed").Value)
    End If
    vRecordset.Close
    
     If vPrint = 0 Then
        vQuery = "Update npmaster.dbo.npprintserver set Printed = 1 where Docno = '" & vDocNo & "' "
        gConnection.Execute vQuery
        ListView331.ListItems.Remove (ListView331.SelectedItem.Index)
    End If
    
Else
MsgBox "ไม่มีเลขที่ให้พิมพ์ หรือ ไม่ได้เลือกฟอร์มที่จะพิมพ์ หรือ ไม่ได้เลือกเลขที่เอกสาร", vbCritical, "ข้อความแจ้งเตือน"
End If

TXT331.Text = ""
TXT332.Text = ""



ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD332_Click()
Call RefreshData
End Sub

Private Sub Form_Load()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim DPListItems As ListItem
Dim DPItemforms As ListItem
Dim vTypeDoc As String

On Error GoTo ErrDescription



 vTypeDoc = "DP"

ListView331.ListItems.Clear
 vQuery = "Select Docno,LastPrintDateTime  from NPMaster.dbo.NPPrintServer  where Printed = 0  " _
                        & " and LastPrintedUser = '" & vUserID & "' and DoctypeID = '" & vTypeDoc & "'  order by LastPrintDateTime "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            If Not vRecordset.EOF Then
                vRecordset.MoveFirst
                While Not vRecordset.EOF
                Set DPListItems = ListView331.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
                DPListItems.SubItems(1) = Trim(vRecordset.Fields("LastPrintDateTime").Value)
                vRecordset.MoveNext
                Wend
            End If
        End If
        vRecordset.Close
        '--------------------------------------------------------------------------------------------------------
        
        ListView332.ListItems.Clear
        vQuery = "select Name from npmaster.dbo.NPForm where ModuleID = '" & vTypeDoc & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
            Set DPItemforms = ListView332.ListItems.Add(, , Trim(vRecordset.Fields("Name").Value))
            vRecordset.MoveNext
            Wend
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------------------------------
        
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If

End Sub


Private Sub ListView331_ItemClick(ByVal Item As MSComctlLib.ListItem)
TXT331.Text = Item
End Sub


Private Sub ListView332_ItemClick(ByVal Item As MSComctlLib.ListItem)
TXT332.Text = Item
End Sub

Private Sub Timer331_Timer()
Dim vTime, vSubTime, vTotalTime
Dim vNumber As Integer

On Error GoTo ErrDescription

            If LBLTime331.Caption <> CStr(Time) Then
                 LBLTime331.Caption = Time
            End If
             vTime = Second(Time)
             vTotalTime = vTime Mod 5
             If vTotalTime = 0 Then
             Call RefreshData
             End If

    
ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub


Function NewListItems()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim INVItemLists As ListItem
Dim vTypeDoc As String

On Error GoTo ErrDescription
    vTypeDoc = "DP"
    ListView331.ListItems.Clear
    vQuery = "Select Docno,LastPrintDateTime  from NPMaster.dbo.NPPrintServer  where Printed = 0  " _
                        & " and LastPrintedUser = '" & vUserID & "' and DoctypeID = '" & vTypeDoc & "'  order by LastPrintDateTime "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            If Not vRecordset.EOF Then
                vRecordset.MoveFirst
                While Not vRecordset.EOF
                Set INVItemLists = ListView331.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
                INVItemLists.SubItems(1) = Trim(vRecordset.Fields("LastPrintDateTime").Value)
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

On Error GoTo ErrDescription

            vDocHeader = "DP"

            vQuery = "Select  Docno,LastPrintDateTime,printed from NPMaster.dbo.NPPrintServer   where Printed = 0 " _
                        & " and LastPrintedUser = '" & vUserID & "' and DoctypeID = '" & vDocHeader & "' order by LastPrintDateTime "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    If Not vRecordset.EOF Then
                        CountRecordset = vRecordset.RecordCount
                        CountList = ListView331.ListItems.Count
                                If CountRecordset > CountList Then
                                vRecordset.MoveFirst
                                For i = 1 To CountRecordset
                                If i < CountRecordset Then
                                vDocNo = ListView331.ListItems.Item(i).Text
                                End If
                                vNewDoc = Trim(vRecordset.Fields("Docno").Value)
                                vPrintStatus = Trim(vRecordset.Fields("Printed").Value)
                                        If vDocNo <> vNewDoc Then
                                        Set DocListItem = ListView331.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
                                        DocListItem.SubItems(1) = Trim(vRecordset.Fields("LastPrintDateTime").Value)
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

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
End If
End Function


Public Sub PrintDeposit()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vGroupDoc As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

        vDocNo = Trim(TXT331.Text)
        vQuery = "select groupdoc,printed from  bcnp.dbo.vw_sl_00001 where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------
        
                            vRepType = "DP"
                            If UCase(vGroupDoc) = "IDV" Then
                                vRepID = 340
                            ElseIf UCase(vGroupDoc) = "IDM" Then
                                vRepID = 66
                            ElseIf UCase(vGroupDoc) = "IDN" Then
                                vRepID = 67
                            Else
                            MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
                            Exit Sub
                            End If
                            
                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport331
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true'"
                                    .Destination = crptToWindow
                                    .WindowState = crptMaximized
                                    .Action = 1
                                End With
                            End If
                            vRecordset.Close
                            '---------------------------------------------------------------------------------------------------
                                             
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub
