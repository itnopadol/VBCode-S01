VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form Form36 
   Caption         =   "หน้าพิมพ์ใบเพิ่มหนี้"
   ClientHeight    =   8340
   ClientLeft      =   2505
   ClientTop       =   810
   ClientWidth     =   12000
   Icon            =   "Form36.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Form36.frx":08CA
   ScaleHeight     =   8340
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport361 
      Left            =   6360
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowNavigationCtls=   -1  'True
      WindowShowCancelBtn=   -1  'True
      WindowShowPrintBtn=   -1  'True
      WindowShowExportBtn=   -1  'True
      WindowShowZoomCtl=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowProgressCtls=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton CMD362 
      Caption         =   "RefreshData"
      Height          =   465
      Left            =   4875
      TabIndex        =   9
      Top             =   1050
      Width           =   1440
   End
   Begin VB.TextBox TXT362 
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
      Height          =   465
      Left            =   2550
      TabIndex        =   4
      Top             =   6375
      Width           =   2790
   End
   Begin VB.TextBox TXT361 
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
      Height          =   465
      Left            =   2550
      TabIndex        =   3
      Top             =   5550
      Width           =   2790
   End
   Begin VB.CommandButton CMD361 
      Caption         =   "พิมพ์ใบเพิ่มหนี้"
      Height          =   690
      Left            =   8625
      TabIndex        =   2
      Top             =   5550
      Width           =   1815
   End
   Begin MSComctlLib.ListView ListView362 
      Height          =   3540
      Left            =   7050
      TabIndex        =   1
      Top             =   1575
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   6244
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
         Object.Width           =   5821
      EndProperty
   End
   Begin MSComctlLib.ListView ListView361 
      Height          =   3540
      Left            =   1275
      TabIndex        =   0
      Top             =   1575
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   6244
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
         Text            =   "เลขที่ใบเพิ่มหนี้"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "วันที่ทำใบเพิ่มหนี้"
         Object.Width           =   5239
      EndProperty
   End
   Begin VB.Timer Timer361 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   9675
      Top             =   7725
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "พิมพ์ใบเพิ่มหนี้"
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
      Width           =   7440
   End
   Begin VB.Label LBLTime361 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   8700
      TabIndex        =   7
      Top             =   1200
      Width           =   1740
   End
   Begin VB.Label LBL362 
      BackStyle       =   0  'Transparent
      Caption         =   "ฟอร์มที่พิมพ์"
      ForeColor       =   &H8000000E&
      Height          =   465
      Left            =   1350
      TabIndex        =   6
      Top             =   6375
      Width           =   990
   End
   Begin VB.Label LBL361 
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบเพิ่มหนี้"
      ForeColor       =   &H8000000E&
      Height          =   465
      Left            =   1350
      TabIndex        =   5
      Top             =   5550
      Width           =   1215
   End
End
Attribute VB_Name = "Form36"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD361_Click()
Dim vQuery As String
Dim vDocNo As String
Dim vPrint As Integer
Dim vRecordset As New ADODB.Recordset
Dim vCompleteSave As Integer

On Error GoTo ErrDescription

vDocNo = Trim(TXT361.Text)
If TXT361.Text <> "" And TXT362.Text <> "" Then

   vQuery = "select  docno,isnull(iscompletesave,0) as iscompletesave  from dbo.bcdebitnote1 where docno = '" & vDocNo & "' "
   If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   vCompleteSave = Trim(vRecordset.Fields("iscompletesave").Value)
   End If
   vRecordset.Close

   If vCompleteSave = 0 Then
      MsgBox "ไม่สามารถพิมพ์เอกสารที่ยังบันทึกข้อมูลไม่สมบูรณ์ได้ กรุณารอสักครู่แล้วกดพิมพ์ใหม่อีกครั้ง", vbCritical, "Send Error Message"
      Exit Sub
   End If
   
    If TXT362.Text = "พิมพ์ใบเพิ่มหนี้" Then
        Call PrintDebit
    End If
    vQuery = "select printed from  bcnp.dbo.vw_sl_00001 where docno = '" & vDocNo & "' "
    If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
        vPrint = Trim(vRecordset.Fields("printed").Value)
    End If
    vRecordset.Close
    
     If vPrint = 0 Then
        vQuery = "Update npmaster.dbo.npprintserver set Printed = 1 where Docno = '" & vDocNo & "' "
        gConnection.Execute vQuery
        ListView361.ListItems.Remove (ListView361.SelectedItem.Index)
    End If
    
Else
MsgBox "ไม่มีเลขที่ให้พิมพ์ หรือ ไม่ได้เลือกฟอร์มที่จะพิมพ์ หรือ ไม่ได้เลือกเลขที่เอกสาร", vbCritical, "ข้อความแจ้งเตือน"
End If


TXT361.Text = ""
TXT362.Text = ""

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If

End Sub

Private Sub CMD362_Click()
Call RefreshData
End Sub

Private Sub Form_Load()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim DBListItems As ListItem
Dim DBItemforms As ListItem
Dim vTypeDoc As String

On Error GoTo ErrDescription

 vTypeDoc = "DB"

ListView361.ListItems.Clear
 vQuery = "Select Docno,LastPrintDateTime from npmaster.dbo.npprintserver where Printed = 0  " _
                        & " and LastPrintedUser = '" & vUserID & "' and DoctypeID = '" & vTypeDoc & "'  order by LastPrintDateTime "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            If Not vRecordset.EOF Then
                vRecordset.MoveFirst
                While Not vRecordset.EOF
                Set DBListItems = ListView361.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
                DBListItems.SubItems(1) = Trim(vRecordset.Fields("LastPrintDateTime").Value)
                vRecordset.MoveNext
                Wend
            End If
        End If
        vRecordset.Close
        '--------------------------------------------------------------------------------------------------------------------
        ListView362.ListItems.Clear
        vQuery = "select Name from npmaster.dbo.NPForm where ModuleID = '" & vTypeDoc & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
            Set DBItemforms = ListView362.ListItems.Add(, , Trim(vRecordset.Fields("Name").Value))
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

Private Sub ListView361_ItemClick(ByVal Item As MSComctlLib.ListItem)
TXT361.Text = Item
End Sub

Private Sub ListView362_ItemClick(ByVal Item As MSComctlLib.ListItem)
TXT362.Text = Item
End Sub

Private Sub Timer361_Timer()
Dim vTime, vSubTime, vTotalTime
Dim vNumber As Integer

On Error GoTo ErrDescription

            If LBLTime361.Caption <> CStr(Time) Then
                 LBLTime361.Caption = Time
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

Public Sub PrintDebit()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vGroupDoc As String
Dim vReportName As String
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

        vDocNo = Trim(TXT361.Text)
        vQuery = "select groupdoc from  npmaster.dbo.npprintserver where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------
        
                            vRepType = "DB"
                            If vGroupDoc = "DHV" Or vGroupDoc = "dhv" Then
                                vRepID = 62
                            ElseIf vGroupDoc = "DHN" Or vGroupDoc = "dhn" Then
                                vRepID = 63
                            ElseIf vGroupDoc = "DCV" Or vGroupDoc = "dcv" Then
                                vRepID = 476
                            ElseIf vGroupDoc = "DCN" Or vGroupDoc = "dcn" Then
                                vRepID = 477
                            Else
                            MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
                            Exit Sub
                            End If
                            
                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport361
                                    .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                                    .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
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

Function NewListItems()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim INVItemLists As ListItem
Dim vTypeDoc As String

On Error GoTo ErrDescription
    vTypeDoc = "DB"
    ListView361.ListItems.Clear
    vQuery = "Select Docno,LastPrintDateTime  from NPMaster.dbo.NPPrintServer  where Printed = 0  " _
                        & " and LastPrintedUser = '" & vUserID & "' and DoctypeID = '" & vTypeDoc & "'  order by LastPrintDateTime "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            If Not vRecordset.EOF Then
                vRecordset.MoveFirst
                While Not vRecordset.EOF
                Set INVItemLists = ListView361.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
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

            vDocHeader = "DB"

            vQuery = "Select  Docno,LastPrintDateTime,printed from NPMaster.dbo.NPPrintServer   where Printed = 0 " _
                        & " and LastPrintedUser = '" & vUserID & "' and DoctypeID = '" & vDocHeader & "' order by LastPrintDateTime "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    If Not vRecordset.EOF Then
                        CountRecordset = vRecordset.RecordCount
                        CountList = ListView361.ListItems.Count
                                If CountRecordset > CountList Then
                                vRecordset.MoveFirst
                                For i = 1 To CountRecordset
                                If i < CountRecordset Then
                                vDocNo = ListView361.ListItems.Item(i).Text
                                End If
                                vNewDoc = Trim(vRecordset.Fields("Docno").Value)
                                vPrintStatus = Trim(vRecordset.Fields("Printed").Value)
                                        If vDocNo <> vNewDoc Then
                                        Set DocListItem = ListView361.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
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

