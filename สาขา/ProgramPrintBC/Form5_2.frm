VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form5_2 
   Caption         =   "หน้าพิมพ์ใบเสร็จรับชำระ"
   ClientHeight    =   8355
   ClientLeft      =   2205
   ClientTop       =   885
   ClientWidth     =   12000
   Icon            =   "Form5_2.frx":0000
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   Picture         =   "Form5_2.frx":08CA
   ScaleHeight     =   8355
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport5_21 
      Left            =   6600
      Top             =   6360
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
   Begin VB.CommandButton CMD5_22 
      Caption         =   "RefreshData"
      Height          =   540
      Left            =   4950
      TabIndex        =   10
      Top             =   975
      Width           =   1440
   End
   Begin VB.OptionButton Opt5_21 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "โชว์เช็ค"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2850
      TabIndex        =   8
      Top             =   5025
      Width           =   2865
   End
   Begin VB.Timer Timer5_21 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   9450
      Top             =   7650
   End
   Begin VB.TextBox TXT5_22 
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
      Left            =   2850
      TabIndex        =   4
      Top             =   6000
      Width           =   2865
   End
   Begin VB.TextBox TXT5_21 
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
      Left            =   2850
      TabIndex        =   3
      Top             =   5475
      Width           =   2865
   End
   Begin VB.CommandButton CMD5_21 
      Caption         =   "พิมพ์"
      Height          =   690
      Left            =   8550
      TabIndex        =   2
      Top             =   5025
      Width           =   1890
   End
   Begin MSComctlLib.ListView ListView5_22 
      Height          =   3090
      Left            =   7050
      TabIndex        =   1
      Top             =   1575
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   5450
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
         Object.Width           =   5822
      EndProperty
   End
   Begin MSComctlLib.ListView ListView5_21 
      Height          =   3090
      Left            =   450
      TabIndex        =   0
      Top             =   1575
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   5450
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "เลขที่ใบเสร็จรับชำระ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "วันที่ทำใบเสร็จ"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ชื่อลูกค้า"
         Object.Width           =   4146
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "พิมพ์ใบเสร็จรับชำระ"
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
      TabIndex        =   9
      Top             =   300
      Width           =   7365
   End
   Begin VB.Label LBLTime5_21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   8925
      TabIndex        =   7
      Top             =   1275
      Width           =   1440
   End
   Begin VB.Label LBL5_22 
      BackStyle       =   0  'Transparent
      Caption         =   "ฟอร์มที่พิมพ์"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   1875
      TabIndex        =   6
      Top             =   6000
      Width           =   990
   End
   Begin VB.Label LBL5_21 
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบเสร็จรับชำระ"
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   1275
      TabIndex        =   5
      Top             =   5475
      Width           =   1590
   End
End
Attribute VB_Name = "Form5_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD5_21_Click()
Dim vQuery As String, vDocNo As String
Dim vRecordset As New ADODB.Recordset

On Error GoTo ErrDescription

If TXT5_21.Text <> "" And TXT5_22.Text <> "" Then
    If TXT5_22.Text = "พิมพ์ใบเสร็จรับชำระ" Then
        Call PrintCashReceipt
    End If
Else
MsgBox "ไม่มีเลขที่ให้พิมพ์ หรือ ไม่ได้เลือกฟอร์มที่จะพิมพ์ หรือ ไม่ได้เลือกเลขที่เอกสาร", vbCritical, "ข้อความแจ้งเตือน"
End If


TXT5_21.Text = ""
TXT5_22.Text = ""
Opt5_21.Value = False

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Private Sub CMD5_22_Click()
Call RefreshData
End Sub

Private Sub Form_Load()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim DBListItems As ListItem
Dim DBItemforms As ListItem
Dim vTypeDoc As String

On Error GoTo ErrDescription

 vTypeDoc = "RE"

ListView5_21.ListItems.Clear
 vQuery = "Select Docno,LastPrintDateTime,arname  from NPMaster.dbo.vw_RE_CashReceipt  where Printed = 0  " _
                        & " and LastPrintedUser = '" & vUserID & "' and DoctypeID = '" & vTypeDoc & "'  order by LastPrintDateTime "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            If Not vRecordset.EOF Then
                vRecordset.MoveFirst
                While Not vRecordset.EOF
                Set DBListItems = ListView5_21.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
                DBListItems.SubItems(1) = Trim(vRecordset.Fields("LastPrintDateTime").Value)
                DBListItems.SubItems(2) = Trim(vRecordset.Fields("arname").Value)
                vRecordset.MoveNext
                Wend
            End If
        End If
        vRecordset.Close
        '--------------------------------------------------------------------------------------------------------------------
        ListView5_22.ListItems.Clear
        vQuery = "select Name from npmaster.dbo.NPForm where ModuleID = '" & vTypeDoc & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
            Set DBItemforms = ListView5_22.ListItems.Add(, , Trim(vRecordset.Fields("Name").Value))
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

Private Sub ListView5_21_ItemClick(ByVal Item As MSComctlLib.ListItem)
TXT5_21.Text = Item
End Sub

Private Sub ListView5_22_ItemClick(ByVal Item As MSComctlLib.ListItem)
TXT5_22.Text = Item
End Sub

Private Sub Timer5_21_Timer()
Dim vTime, vSubTime, vTotalTime
Dim vNumber As Integer

On Error GoTo ErrDescription

            If LBLTime5_21.Caption <> CStr(Time) Then
                 LBLTime5_21.Caption = Time
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
    vTypeDoc = "RE"
    ListView5_21.ListItems.Clear
    vQuery = "Select Docno,LastPrintDateTime  from NPMaster.dbo.vw_RE_CashReceipt  where Printed = 0  " _
                        & " and LastPrintedUser = '" & vUserID & "' and DoctypeID = '" & vTypeDoc & "'  order by LastPrintDateTime "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            If Not vRecordset.EOF Then
                vRecordset.MoveFirst
                While Not vRecordset.EOF
                Set INVItemLists = ListView5_21.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
                INVItemLists.SubItems(1) = Trim(vRecordset.Fields("LastPrintDateTime").Value)
                INVItemLists.SubItems(2) = Trim(vRecordset.Fields("arname").Value)
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

            vDocHeader = "RE"

        vQuery = "Select  Docno,LastPrintDateTime,printed from NPMaster.dbo.vw_RE_CashReceipt   where Printed = 0 " _
                            & " and LastPrintedUser = '" & vUserID & "' and DoctypeID = '" & vDocHeader & "' order by LastPrintDateTime "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    If Not vRecordset.EOF Then
                        CountRecordset = vRecordset.RecordCount
                        CountList = ListView5_21.ListItems.Count
                                If CountRecordset > CountList Then
                                vRecordset.MoveFirst
                                For i = 1 To CountRecordset
                                If i < CountRecordset Then
                                vDocNo = ListView5_21.ListItems.Item(i).Text
                                End If
                                vNewDoc = Trim(vRecordset.Fields("Docno").Value)
                                vPrintStatus = Trim(vRecordset.Fields("Printed").Value)
                                        If vDocNo <> vNewDoc Then
                                        Set DocListItem = ListView5_21.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
                                        DocListItem.SubItems(1) = Trim(vRecordset.Fields("LastPrintDateTime").Value)
                                        DocListItem.SubItems(2) = Trim(vRecordset.Fields("arname").Value)
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


Public Sub PrintCashReceipt()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vGroupDoc As String
Dim vReportName As String
Dim vPrint As Integer
Dim vRepID As Integer
Dim vRepType As String
Dim vChqStatus As String

On Error GoTo ErrDescription

        If Opt5_21.Value = False Then
            vChqStatus = 1
        Else
            vChqStatus = 0
        End If
        vDocNo = Trim(TXT5_21.Text)

        vQuery = "select groupdoc,printed from  npmaster.dbo.npprintserver where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
                vPrint = Trim(vRecordset.Fields("Printed").Value)
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------
        
                            vRepType = "RE"
                            If vGroupDoc = "RE" Or vGroupDoc = "re" Then
                                vRepID = 38
                            ElseIf vGroupDoc = "RN" Or vGroupDoc = "rn" Then
                                vRepID = 79
                            ElseIf vGroupDoc = "RD" Or vGroupDoc = "rd" Then
                                vRepID = 38
                            ElseIf vGroupDoc = "RC" Or vGroupDoc = "rc" Then
                                vRepID = 79
                            ElseIf vGroupDoc = "RQ" Or vGroupDoc = "rq" Then
                                vRepID = 79
                            Else
                            MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
                            Exit Sub
                            End If
                            
                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport5_21
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "DocNo;" & vDocNo & ";true"
                                    .ParameterFields(1) = "ChqStatus;" & vChqStatus & ";true"
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
                                                ListView5_21.ListItems.Remove (ListView5_21.SelectedItem.Index)
                                    End If
                            
                            
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

