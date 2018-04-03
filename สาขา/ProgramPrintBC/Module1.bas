Attribute VB_Name = "Module1"
Option Explicit
Global ConnectionString As String
Global VatConnectionString As String
Global gConnection As New ADODB.Connection
Global vConnection As New ADODB.Connection
Global vVatConnection As New ADODB.Connection
Global conn As New ADODB.Connection
Global rs As New ADODB.Recordset
Global Sqlstr As String
Global vCompany1 As String, vUserID1 As String, vPassword1 As String
Global vCompany As String, vUserID As String, vPassword As String
Global vCheckValue As Boolean
Global vCheckPercent As Boolean
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Global vUserName1 As String
Global vComputerName1 As String
Global vGenDocNo, vGenDocno1 As String
Global vChkFrmActivate As Integer
Global vClosePR As Integer
Global vCheckDuplicate As Integer
Global ConnectionStringBPlus As String
Global vBPlusConnect As New ADODB.Connection
Global vFormActivate As String
Global vAccess As Integer
Global vCheckButton As Integer
Global vZoneUser As Integer
'------------------------------------------------------------------
Global gConnectionBPlus As New ADODB.Connection
Global vConnectionString As String
Global vIsOpen1 As Integer
Global vIsOpen2 As Integer
Global vCheckDocnoExist1 As Integer
Global vCheckDocnoExist2 As Integer
Global vCheckIsCancel1 As Integer
Global vCheckIsCancel2 As Integer
Global vCheckIsConfirm1 As Integer
Global vPriorityModule As Integer
Global vReceiveModule As Integer
Global vCheckReceiveOpen As Integer
Global vCheckIsReceive As Integer
Global vCheckChageDataReceive As Integer
Global vCheckPriorityOpen As Integer
Global vCheckPlaceOpen As Integer
Global vPlaceModule As Integer
Global vCheckIsPlace As Integer
Global vCheckRouteOpen As Integer
Global vRouteModule As Integer
Global vCheckIsRoute As Integer
Global vVehicalModule As Integer
Global vCheckIsVehical As Integer
Global vCheckVehicalOpen As Integer
Global vEmpModule As Integer
Global vCheckIsEmp As Integer
Global vCheckEmpOpen As Integer
Global vExistNumber As Integer
Global vCheckTextBox As Integer
Global vCheckSale As Integer
Global vCheckPosition As Integer
Global vCheckAddReceiver As Integer
Global vPrintNo As String
Global vPrintForm As String
Global vWindows As Integer
Global vPrinter As Integer
Global vPrintClick As Integer
Global vShowDiscountLine As Integer

Global vOpenHoldBill    As Integer

Global vDepartment As String
Global vLevelID As Integer
Global vFormName As String
Global vUserAuthority As Integer
'--------------------------------------------------------------

Public Const vbViolet = &HFF8080
Public Const vbVioletBright = &HFFC0C0
Public Const vbForestGreen = &H228B22
Public Const vbGray = &HE0E0E0
Public Const vbLightBlue = &HFFD3A4
Public Const vbLightGreen = &HABFCBD
Public Const vbGreenLemon = &HB3FFBE
Public Const vbYellowBright = &HC0FFFF
Public Const vbOrange = &H2CCDFC
Public Const vbVeryLightBlue = &HFFFFC0
Public Const vbVeryLightGreen = &HC0FFC0
'--------------------------------------------------------------

Public Sub InitializeDataBase()
If vConnection.State <> 0 Then
vConnection.Close
ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info =false;User ID=VBUser;Password='132';Data Source = S02DB;Initial Catalog = BCNP"
vConnection.Open (ConnectionString)
Else
ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info =false;User ID=VBUser;Password='132';Data Source = S02DB;Initial Catalog = BCNP"
vConnection.Open (ConnectionString)
End If
End Sub

'Public Sub InitializeDataBaseVat()
'If vVatConnection.State <> 0 Then
'vVatConnection.Close
'VatConnectionString = "Provider=SQLOLEDB.1;Persist Security Info =false;User ID=sa;Password='';Data Source = Solar;Initial Catalog = BCVAT"
'vVatConnection.Open (ConnectionString)
'Else
'VatConnectionString = "Provider=SQLOLEDB.1;Persist Security Info =false;User ID=sa;Password='';Data Source = Solar;Initial Catalog = BCVAT"
'vVatConnection.Open (ConnectionString)
'End If
'End Sub

Public Function OpenDataBase(gConnection As ADODB.Connection, gRecordset As ADODB.Recordset, vQuery As String) As String
gConnection.CursorLocation = adUseClient
gRecordset.Open vQuery, gConnection, adOpenDynamic, adLockOptimistic
OpenDataBase = gRecordset.RecordCount
End Function

'Public Function OpenDataBaseBCVat(VatConnection As ADODB.Connection, gVatRecordset As ADODB.Recordset, vQuery As String) As String
'vVatConnection.CursorLocation = adUseClient
'gVatRecordset.Open vQuery, vVatConnection, adOpenDynamic, adLockOptimistic
'OpenDataBaseBCVat = gVatRecordset.RecordCount
'End Function

Public Function OpenBPlus(vConnection As ADODB.Connection, vRecordset As ADODB.Recordset, vQuery As String) As String
vConnection.CursorLocation = adUseClient
vRecordset.Open vQuery, vConnection, adOpenDynamic, adLockOptimistic
OpenBPlus = vRecordset.RecordCount
End Function
Public Sub GetConnect()
If conn.State = 0 Then
    conn.Provider = "SQLOLEDB.1"
    conn.Properties("Persist Security Info") = False
    conn.Properties("User ID") = "vbuser"
    conn.Properties("Password") = "132"
    conn.Properties("Data Source") = "S02DB"
    conn.Properties("Initial Catalog") = "BCNP"
    conn.CursorLocation = adUseClient
    conn.Open
Else
    conn.Close
    conn.Provider = "SQLOLEDB.1"
    conn.Properties("Persist Security Info") = False
    conn.Properties("User ID") = "vbuser"
    conn.Properties("Password") = "132"
    conn.Properties("Data Source") = "S02DB"
    conn.Properties("Initial Catalog") = "BCNP"
    conn.CursorLocation = adUseClient
    conn.Open

End If
End Sub

Public Sub InitializeDataBase1()
If gConnection.State <> 0 Then
gConnection.Close
ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info =false;User ID=" & vUserID & ";Password=" & vPassword & ";Data Source = S02DB;Initial Catalog = BCNP"
gConnection.Open (ConnectionString)
Else
ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info =false;User ID=" & vUserID & ";Password=" & vPassword & ";Data Source = S02DB;Initial Catalog = BCNP"
gConnection.Open (ConnectionString)
End If
End Sub

Public Function CheckNumeric(ChqNumber As String) As String
Dim i As Integer
Dim vQuery As String
Dim vRecordset As New ADODB.Recordset
Dim vCheckText As String

For i = 1 To 7
    vCheckText = Mid(ChqNumber, i, 1)
    If Mid(ChqNumber, i, 1) = 0 Or Mid(ChqNumber, i, 1) = 1 Or Mid(ChqNumber, i, 1) = 2 Or Mid(ChqNumber, i, 1) = 3 Or Mid(ChqNumber, i, 1) = 4 Or Mid(ChqNumber, i, 1) = 5 Or Mid(ChqNumber, i, 1) = 6 Or Mid(ChqNumber, i, 1) = 7 Or Mid(ChqNumber, i, 1) = 8 Or Mid(ChqNumber, i, 1) = 9 Then
        vCheckValue = True
    Else
        vCheckValue = False
        Exit Function
    End If
Next i

End Function

Public Sub GetComputerandUser()
Dim vReturnStatus As Long
Dim vComputerName As String
Dim vUserName As String
Dim vReturnStatus1 As Long

vUserName = Space(250)
vComputerName = Space(250)
vReturnStatus = GetComputerName(vComputerName, Len(vComputerName) - 1)
vComputerName = Trim(vComputerName)
vReturnStatus1 = GetUserName(vUserName, Len(vUserName))
vUserName = Trim(vUserName)
vComputerName1 = Left(vComputerName, Len(vComputerName) - 1)
vUserName1 = Left(vUserName, Len(vUserName) - 1)
End Sub

Public Sub GenHeadDocument()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String
Dim vDocNo As String
Dim vYear1, vYear2 As Integer
Dim vYear3, vMonth1, vMonth2, vDay1, vDay2, vDate As String

vYear1 = Right(Year(Now), 2)
If vYear1 < 48 Then
    vYear2 = vYear1 + 43
Else
    vYear2 = vYear1
End If
    vYear3 = vYear2

vMonth1 = Month(Now)
If vMonth1 < 10 Then
vMonth2 = "0" & vMonth1
Else
vMonth2 = vMonth1
End If

vDay1 = Day(Now)
If vDay1 < 10 Then
vDay2 = "0" & vDay1
Else
vDay2 = vDay1
End If

vGenDocNo = Trim(vYear3 & vMonth2 & vDay2)

End Sub

Public Sub ChekAuthorityAccess()
Dim vRecordset As New ADODB.Recordset
Dim vQuery As String


vQuery = "select department ,levelid from bcnp.dbo.vw_np_UserAutorityProgram where code = '" & vUserID & "' "
If OpenBPlus(vConnection, vRecordset, vQuery) <> 0 Then
    vDepartment = Trim(vRecordset.Fields("department").Value)
    vLevelID = Trim(vRecordset.Fields("levelid").Value)
    vUserAuthority = 1
Else
    vUserAuthority = 0
End If
vRecordset.Close

If vUserAuthority = 1 Then
vFormName = vFormActivate
vQuery = "select pagestatus from vw_np_AccessProgram where departmentcode = '" & vDepartment & "' and levelid = " & vLevelID & " and pagename = '" & vFormName & "' "
If OpenBPlus(vConnection, vRecordset, vQuery) <> 0 Then
    vAccess = Trim(vRecordset.Fields("pagestatus").Value)
Else
    vAccess = 0
End If
vRecordset.Close
End If
End Sub

Public Sub GetDataBPlus()
If gConnectionBPlus.State = 0 Then
vConnectionString = Trim("Provider = SQLOLEDB.1;Persist security Info = False;User ID='vbuser';Password = '132';Data Source = S02DB;Initial Catalog = BCNP")
gConnectionBPlus.Open (vConnectionString)
Else
gConnectionBPlus.Close
vConnectionString = Trim("Provider = SQLOLEDB.1;Persist security Info = False;User ID='vbuser';Password = '132';Data Source = S02DB;Initial Catalog = BCNP")
gConnectionBPlus.Open (vConnectionString)
End If
End Sub

Public Sub CheckNumber(vData As String)
Dim vDocNo As String
Dim vText As String
Dim i As Integer

On Error GoTo ErrDescription

vDocNo = Trim(vData)

For i = 1 To Len(vData)
    If Mid(vDocNo, i, 1) = 0 Or Mid(vDocNo, i, 1) = 1 Or Mid(vDocNo, i, 1) = 2 Or Mid(vDocNo, i, 1) = 3 Or Mid(vDocNo, i, 1) = 4 Or Mid(vDocNo, i, 1) = 5 Or Mid(vDocNo, i, 1) = 6 Or Mid(vDocNo, i, 1) = 7 Or Mid(vDocNo, i, 1) = 8 Or Mid(vDocNo, i, 1) = 9 Or Mid(vDocNo, i, 1) = "." Or Mid(vDocNo, i, 1) = "," Then
        vCheckValueNumber = True
    Else
        vCheckValueNumber = False
        Exit Sub
    End If
Next i

ErrDescription:
If Err.Description <> "" Then
MsgBox Err.Description
Exit Sub
End If
End Sub

Public Sub ReportSetLocation(pCrystalReport As CrystalReport, rDataBase As String)
    Dim vStop As Integer
    Dim i As Integer
    Dim gDefaultDatabaseReport As String
    Dim gCurrentDatabaseName As String
    Dim vName As String
     
    vStop = pCrystalReport.RetrieveDataFiles - 1
    gDefaultDatabaseReport = "BCNP"
    gCurrentDatabaseName = rDataBase
    vName = pCrystalReport.DataFiles(0)
    For i = 0 To vStop
        pCrystalReport.DataFiles(i) = ReplaceString(pCrystalReport.DataFiles(i), _
        gDefaultDatabaseReport, gCurrentDatabaseName)
    Next
End Sub

Public Function ReplaceString(pSourceString As String, pReplaceString As String, pWithString As String) As String
    Dim vPosition As Integer
    Dim vReturnValue As String
    Dim vTemp As String
    Dim vLengthReplaceString As Integer
    
    vLengthReplaceString = Len(pReplaceString)
    vTemp = pSourceString
    vReturnValue = ""
    Do Until vTemp = ""
        vPosition = InStr(1, vTemp, pReplaceString, vbTextCompare)
        If vPosition = 0 Then
            vReturnValue = vReturnValue & vTemp
            vTemp = ""
        Else
            vReturnValue = vReturnValue & Left(vTemp, vPosition - 1) & pWithString
            vTemp = Right(vTemp, Len(vTemp) - (vPosition + vLengthReplaceString - 1))
        End If
    Loop
    ReplaceString = vReturnValue
End Function


Public Sub SetListViewColor(pCtrlListView As ListView, _
                            pCtrlPictureBox As PictureBox, _
                            Color1 As Long, _
                            Color2 As Long)

On Error GoTo SetListViewColor_Error

    Dim iLineHeight As Long
    Dim iBarHeight  As Long
    Dim lBarWidth   As Long
    Dim lColor1     As Long
    Dim lColor2     As Long
 
    lColor1 = Color1
    lColor2 = Color2
    
    If pCtrlListView.View = lvwReport Then
        pCtrlListView.Picture = LoadPicture("")
        pCtrlListView.Refresh
        pCtrlPictureBox.Cls
        
        pCtrlPictureBox.AutoRedraw = True
        pCtrlPictureBox.BorderStyle = vbBSNone
        pCtrlPictureBox.ScaleMode = vbTwips
        pCtrlPictureBox.Visible = False
        
        pCtrlListView.PictureAlignment = lvwTile
        pCtrlPictureBox.Font = pCtrlListView.Font
        pCtrlPictureBox.Top = pCtrlListView.Top
        pCtrlPictureBox.Font = pCtrlListView.Font
        With pCtrlPictureBox.Font
            .Size = pCtrlListView.Font.Size '+ 2.75
            .Bold = pCtrlListView.Font.Bold
            .Charset = pCtrlListView.Font.Charset
            .Italic = pCtrlListView.Font.Italic
            .Name = pCtrlListView.Font.Name
            .Strikethrough = pCtrlListView.Font.Strikethrough
            .Underline = pCtrlListView.Font.Underline
            .Weight = pCtrlListView.Font.Weight
        End With
        pCtrlPictureBox.Refresh
        iLineHeight = pCtrlPictureBox.TextHeight("W") + Screen.TwipsPerPixelY
    
        iBarHeight = (iLineHeight * 1)
        lBarWidth = pCtrlListView.Width
    
        pCtrlPictureBox.Height = iBarHeight * 2
        pCtrlPictureBox.Width = lBarWidth
    
        pCtrlPictureBox.Line (0, 0)-(lBarWidth, iBarHeight), lColor1, BF
        pCtrlPictureBox.Line (0, iBarHeight)-(lBarWidth, iBarHeight * 2), lColor2, BF
    
        pCtrlPictureBox.AutoSize = True
        pCtrlListView.Picture = pCtrlPictureBox.Image
    Else
        pCtrlListView.Picture = LoadPicture("")
    End If
    
    pCtrlListView.Refresh
    Exit Sub
SetListViewColor_Error:
    pCtrlListView.Picture = LoadPicture("")
    pCtrlListView.Refresh
End Sub
