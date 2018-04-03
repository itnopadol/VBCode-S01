Attribute VB_Name = "Module1"
Option Explicit
Global ConnectionString As String
Global gConnection As New ADODB.Connection
Global vConnection As New ADODB.Connection
Global vCompany As String
Global vUserID As String
Global vPassword As String

'user defined type required by Shell_NotifyIcon API call
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type


'constants required by Shell_NotifyIcon API call:
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_LBUTTONUP = &H202       'Button up
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down
Public Const WM_RBUTTONUP = &H205       'Button up
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public nid As NOTIFYICONDATA

'Public Sub InitializeDatabase()
'If vConnection.State <> 0 Then
'vConnection.Close
'ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info =false;User ID=vbuser;Password=132;Data Source = Nebula;Initial Catalog = BCNP"
'vConnection.Open (ConnectionString)
'Else
'ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info =false;User ID=vbuser;Password=132;Data Source = Nebula;Initial Catalog = BCNP"
'vConnection.Open (ConnectionString)
'End If
'End Sub
 
'Public Function OpenDataBase(gConnection As ADODB.Connection, gRecordset As ADODB.Recordset, vQuery As String) As String
'gConnection.CursorLocation = adUseClient
'gRecordset.Open vQuery, vConnection, adOpenDynamic, adLockOptimistic
'OpenDataBase = gRecordset.RecordCount
'End Function


Public Sub InitializeDatabase()
Dim vQuery As String

'On Error Resume Next

If gConnection.State <> 0 Then
   gConnection.Close
   ConnectionString = "Provider = SQLOLEDB.1;Persist Security Info =False;User ID = vbuser;Password=132;Data Source  =S02DB;Initial Catalog = BCNP "
   gConnection.Open (ConnectionString)
Else
   ConnectionString = "Provider = SQLOLEDB.1;Persist Security Info =False;User ID = vbuser;Password=132;Data Source  =S02DB;Initial Catalog = BCNP "
   gConnection.Open (ConnectionString)
End If
End Sub


Public Function OpenDataBase(gConnection As ADODB.Connection, vRecordset As ADODB.Recordset, vQuery As String) As String
'On Error Resume Next

gConnection.CursorLocation = adUseClient
vRecordset.Open vQuery, gConnection, adOpenDynamic, adLockOptimistic
OpenDataBase = vRecordset.RecordCount
End Function

'Public Sub IconTray(DataIcon As PictureBox, _
 '                         ByVal zTip As String, remove As Boolean)
  '  With iData
   '     .cbSize = Len(iData)
    '    .hwnd = Form1.hwnd
     '   .uId = vbNull
      '  .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
       ' .uCallBackMessage = WM_MOUSEMOVE
        '.hIcon = DataIcon
        '.szTip = zTip & vbNullChar
    'End With
    'If Not remove Then
     '  Shell_NotifyIcon NIM_ADD, iData
       
    'Else
     '  Shell_NotifyIcon NIM_DELETE, iData
    'End If
'End Sub
