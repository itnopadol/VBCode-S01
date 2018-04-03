VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form Form7_0 
   Caption         =   "หน้าพิมพ์ใบนำฝาก"
   ClientHeight    =   8235
   ClientLeft      =   2205
   ClientTop       =   885
   ClientWidth     =   12000
   Icon            =   "Form7_0.frx":0000
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   Picture         =   "Form7_0.frx":08CA
   ScaleHeight     =   8235
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport7_01 
      Left            =   2040
      Top             =   6840
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
   Begin VB.CommandButton CMD7_02 
      Caption         =   "RefreshData"
      Height          =   465
      Left            =   4575
      TabIndex        =   9
      Top             =   975
      Width           =   1590
   End
   Begin VB.TextBox TXT7_02 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   465
      Left            =   2850
      TabIndex        =   4
      Top             =   5325
      Width           =   2790
   End
   Begin VB.TextBox TXT7_01 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   465
      Left            =   2850
      TabIndex        =   3
      Top             =   4725
      Width           =   2790
   End
   Begin VB.CommandButton CMD7_01 
      Caption         =   "พิมพ์"
      Height          =   765
      Left            =   8625
      TabIndex        =   2
      Top             =   4725
      Width           =   1815
   End
   Begin VB.Timer Timer7_01 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   9225
      Top             =   7650
   End
   Begin MSComctlLib.ListView ListView7_02 
      Height          =   2490
      Left            =   6825
      TabIndex        =   1
      Top             =   1500
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4392
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
   Begin MSComctlLib.ListView ListView7_01 
      Height          =   2490
      Left            =   1650
      TabIndex        =   0
      Top             =   1500
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   4392
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
         Text            =   "เลขที่ใบนำฝาก"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "วันที่ทำใบนำฝาก"
         Object.Width           =   5204
      EndProperty
   End
   Begin VB.Label LBL7_03 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "พิมพ์ใบนำฝาก"
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
   Begin VB.Label LBLTime7_01 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   8700
      TabIndex        =   7
      Top             =   1200
      Width           =   1440
   End
   Begin VB.Label LBL7_02 
      BackStyle       =   0  'Transparent
      Caption         =   "ฟอร์มที่พิมพ์"
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   1650
      TabIndex        =   6
      Top             =   5325
      Width           =   990
   End
   Begin VB.Label LBL7_01 
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบนำฝาก"
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   1650
      TabIndex        =   5
      Top             =   4725
      Width           =   1140
   End
End
Attribute VB_Name = "Form7_0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMD7_01_Click()
Dim vDocNo, vGroupDoc As String
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset

On Error GoTo ErrDescription

If TXT7_01.Text <> "" And TXT7_02.Text <> "" Then
        vDocNo = Trim(TXT7_01.Text)
        vQuery = "select groupdoc,printed from  npmaster.dbo.npprintserver where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
        End If
        vRecordset.Close
    If TXT7_02.Text = Trim("พิมพ์ใบนำฝากธนาคารกรุงเทพฯ") Then
            Call PrintBankDeposit_BBL
    ElseIf TXT7_02.Text = Trim("พิมพ์ใบนำฝากธนาคารเอเชีย") Then
            Call PrintBankDeposit_BOA
    ElseIf TXT7_02.Text = Trim("พิมพ์ใบนำฝากธนาคารไทยพาณิชย์") Or TXT7_02.Text = Trim("พิมพ์ใบนำฝากธนาคารกรุงไทย") Or TXT7_02.Text = Trim("พิมพ์ใบนำฝากธนาคารกสิกรไทย") Or TXT7_02.Text = Trim("พิมพ์ใบนำฝากธนาคารกรุงศรีอยุธยา") Then
            Call PrintBankDeposit_Others
    ElseIf TXT7_02.Text = Trim("พิมพ์ใบนำฝากธนาคารทหารไทย") Then
    Call PrintBankDeposit_TMB
    End If

Else
MsgBox "ไม่มีเลขที่ให้พิมพ์ หรือ ไม่ได้เลือกฟอร์มที่จะพิมพ์ หรือ ไม่ได้เลือกเลขที่เอกสาร", vbCritical, "ข้อความแจ้งเตือน"
End If
TXT7_01.Text = ""
TXT7_02.Text = ""

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If

End Sub

Private Sub CMD7_02_Click()
Call RefreshData
End Sub

Private Sub Form_Load()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim DBListItems As ListItem
Dim DBItemforms As ListItem
Dim vTypeDoc As String

On Error GoTo ErrDescription

 vTypeDoc = "BD"

ListView7_01.ListItems.Clear
 vQuery = "Select Docno,LastPrintDateTime  from NPMaster.dbo.NPPrintServer  where Printed = 0  " _
                        & " and LastPrintedUser = '" & vUserID & "' and DoctypeID = '" & vTypeDoc & "'  order by LastPrintDateTime "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            If Not vRecordset.EOF Then
                vRecordset.MoveFirst
                While Not vRecordset.EOF
                Set DBListItems = ListView7_01.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
                DBListItems.SubItems(1) = Trim(vRecordset.Fields("LastPrintDateTime").Value)
                vRecordset.MoveNext
                Wend
            End If
        End If
        vRecordset.Close
        '--------------------------------------------------------------------------------------------------------------------
        ListView7_02.ListItems.Clear
        vQuery = "select Name from npmaster.dbo.NPForm where ModuleID = '" & vTypeDoc & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            vRecordset.MoveFirst
            While Not vRecordset.EOF
            Set DBItemforms = ListView7_02.ListItems.Add(, , Trim(vRecordset.Fields("Name").Value))
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

            vDocHeader = "BD"

            vQuery = "Select  Docno,LastPrintDateTime,printed from NPMaster.dbo.NPPrintServer   where Printed = 0 " _
                        & " and LastPrintedUser = '" & vUserID & "' and DoctypeID = '" & vDocHeader & "' order by LastPrintDateTime "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                    If Not vRecordset.EOF Then
                        CountRecordset = vRecordset.RecordCount
                        CountList = ListView7_01.ListItems.Count
                                If CountRecordset > CountList Then
                                vRecordset.MoveFirst
                                For i = 1 To CountRecordset
                                If i < CountRecordset Then
                                vDocNo = ListView7_01.ListItems.Item(i).Text
                                End If
                                vNewDoc = Trim(vRecordset.Fields("Docno").Value)
                                vPrintStatus = Trim(vRecordset.Fields("Printed").Value)
                                        If vDocNo <> vNewDoc Then
                                        Set DocListItem = ListView7_01.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
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

Function NewListItems()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim INVItemLists As ListItem
Dim vTypeDoc As String

On Error GoTo ErrDescription
    vTypeDoc = "BD"
    ListView7_01.ListItems.Clear
    vQuery = "Select Docno,LastPrintDateTime  from NPMaster.dbo.NPPrintServer  where Printed = 0  " _
                        & " and LastPrintedUser = '" & vUserID & "' and DoctypeID = '" & vTypeDoc & "'  order by LastPrintDateTime "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
            If Not vRecordset.EOF Then
                vRecordset.MoveFirst
                While Not vRecordset.EOF
                Set INVItemLists = ListView7_01.ListItems.Add(, , vRecordset.Fields("DOCNO").Value)
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

'Public Sub PrintBankDeposit()
'Dim vQuery As String
'Dim vRecordset As New ADODB.Recordset
'Dim vDocNo As String
'Dim vGroupDoc As String
'Dim vReportName As String
'Dim vPrint As Integer
'Dim vRepId As Integer
'Dim vRepType As String
'
'On Error GoTo ErrDescription
'
 '       vDocNo = Trim(TXT7_01.Text)
'
 '       vQuery = "select groupdoc,printed from  npmaster.dbo.npprintserver where docno = '" & vDocNo & "' "
  '      If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
   '             vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
    '            vPrint = Trim(vRecordset.Fields("Printed").Value)
     '   End If
      '  vRecordset.Close
        '---------------------------------------------------------------------------------------------
        
       '                     vRepType = "BD"
        '                    If vGroupDoc = "BBL" Or vGroupDoc = "bbl" Then
         '                       vRepId = 39
          '                  ElseIf vGroupDoc = "BOA" Or vGroupDoc = "boa" Then
           '                     vRepId = 40
                '            ElseIf vGroupDoc = "BBT" Or vGroupDoc = "bbt" Then
            '                    vRepId = 39
             '               ElseIf vGroupDoc = "BOT" Or vGroupDoc = "bot" Then
              '                  vRepId = 39
               '             ElseIf vGroupDoc = "BON" Or vGroupDoc = "bon" Then
                 '               vRepId = 39
                  '          Else
                   '         MsgBox "เลขที่เอกสารไม่ถูกต้องกรุณาเลือกเอกสารใหม่ด้วยนะครับ", vbInformation + vbCritical, "ข้อความเตือน"
                    '        Exit Sub
                     '       End If
                            
                      '      vQuery = "select reportname from bcreportname where repid = '" & vRepId & "'  and reptype = '" & vRepType & "' "
                       '     If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                        '        With CrystalReport7_01
                         '           .ReportFileName = Trim(vRecordset.Fields("reportname").Value) & ".rpt"
                          '          .ParameterFields(0) = "@DocNo;" & vDocNo & ";true"
                           '         .Destination = crptToWindow
                            '        .WindowState = crptMaximized
                             '       .Action = 1
                              '  End With
                            'End If
                            'vRecordset.Close
                            '---------------------------------------------------------------------------------------------------
                            
                             'If vPrint = 0 Then
                              '              vQuery = "Update npmaster.dbo.npprintserver set Printed = 1 where Docno = '" & vDocNo & "' "
                               '             gConnection.Execute vQuery
                                '                ListView7_01.ListItems.Remove (ListView7_01.SelectedItem.Index)
                                 '   End If
                            
                            
'ErrDescription:
'If Err.Description <> "" Then
 '   MsgBox Err.Description
'End If
'End Sub

Private Sub ListView7_01_ItemClick(ByVal Item As MSComctlLib.ListItem)
TXT7_01.Text = Item
End Sub

Private Sub ListView7_02_ItemClick(ByVal Item As MSComctlLib.ListItem)
TXT7_02.Text = Item
End Sub

Private Sub Timer7_01_Timer()
Dim vTime, vSubTime, vTotalTime
Dim vNumber As Integer

On Error GoTo ErrDescription

            If LBLTime7_01.Caption <> CStr(Time) Then
                 LBLTime7_01.Caption = Time
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

Public Sub PrintBankDeposit_BBL()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vGroupDoc As String
Dim vReportName As String
Dim vPrint As Integer
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

        vDocNo = Trim(TXT7_01.Text)

        vQuery = "select groupdoc,printed from  npmaster.dbo.npprintserver where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
                vPrint = Trim(vRecordset.Fields("Printed").Value)
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------
        
                            vRepType = "BD"
                            vRepID = 39
                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport7_01
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
                                                ListView7_01.ListItems.Remove (ListView7_01.SelectedItem.Index)
                                    End If
                            
                            
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub
Public Sub PrintBankDeposit_BOA()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vGroupDoc As String
Dim vReportName As String
Dim vPrint As Integer
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

        vDocNo = Trim(TXT7_01.Text)

        vQuery = "select groupdoc,printed from  npmaster.dbo.npprintserver where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
                vPrint = Trim(vRecordset.Fields("Printed").Value)
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------
        
                            vRepType = "BD"
                            vRepID = 40
                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport7_01
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
                                                ListView7_01.ListItems.Remove (ListView7_01.SelectedItem.Index)
                                    End If
                            
                            
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Public Sub PrintBankDeposit_Others()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vGroupDoc As String
Dim vReportName As String
Dim vPrint As Integer
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

        vDocNo = Trim(TXT7_01.Text)

        vQuery = "select groupdoc,printed from  npmaster.dbo.npprintserver where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
                vPrint = Trim(vRecordset.Fields("Printed").Value)
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------
        
                            vRepType = "BD"
                            If Trim(TXT7_02.Text) = "พิมพ์ใบนำฝากธนาคารไทยพาณิชย์" Then
                                vRepID = 229
                            ElseIf Trim(TXT7_02.Text) = "พิมพ์ใบนำฝากธนาคารกรุงศรีอยุธยา" Then
                                vRepID = 230
                            ElseIf Trim(TXT7_02.Text) = "พิมพ์ใบนำฝากธนาคารกรุงไทย" Then
                                vRepID = 228
                            ElseIf Trim(TXT7_02.Text) = "พิมพ์ใบนำฝากธนาคารกสิกรไทย" Then
                                vRepID = 231
                            Else
                                MsgBox "เลือกฟอร์มพิมพ์ใบนำฝากไม่ถูกต้อง"
                            End If
                            
                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport7_01
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
                                                ListView7_01.ListItems.Remove (ListView7_01.SelectedItem.Index)
                                    End If
                            
                            
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub

Public Sub PrintBankDeposit_TMB()
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vDocNo As String
Dim vGroupDoc As String
Dim vReportName As String
Dim vPrint As Integer
Dim vRepID As Integer
Dim vRepType As String

On Error GoTo ErrDescription

        vDocNo = Trim(TXT7_01.Text)

        vQuery = "select groupdoc,printed from  npmaster.dbo.npprintserver where docno = '" & vDocNo & "' "
        If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                vGroupDoc = Trim(vRecordset.Fields("groupdoc").Value)
                vPrint = Trim(vRecordset.Fields("Printed").Value)
        End If
        vRecordset.Close
        '---------------------------------------------------------------------------------------------
        
                            vRepType = "BD"
                            vRepID = 332
                            vQuery = "exec dbo.USP_NP_SelectReportName " & vRepID & ",'" & vRepType & "' "
                            'vQuery = "select reportname from bcreportname where repid = '" & vRepID & "'  and reptype = '" & vRepType & "' "
                            If OpenDataBase(gConnection, vRecordset, vQuery) <> 0 Then
                                With CrystalReport7_01
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
                                                ListView7_01.ListItems.Remove (ListView7_01.SelectedItem.Index)
                                    End If
                            
                            
ErrDescription:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
End Sub
