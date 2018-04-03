Attribute VB_Name = "Module3"
Option Explicit

'***** Module *****

'//This is hacked together from various sources, most notably www.vbAccelerator.com

  Public Enum PrinterOrientationConstants
    ntPortrait = 1
    ntLandscape = 2
  End Enum
  
  Public Enum PrinterPaperSize
    ntLetter = 1
    ntLegal = 5
  End Enum
    
  Public Enum ntOperatingSystem
      ntNTWorkStation4 = 0
      ntNTServer4 = 1
      ntWin2000 = 2
      ntWin2000Server = 3
      ntWinXPHome = 4
      ntWinXPProf = 5
  End Enum
  
  Private Type DEVMODE
    dmDeviceName       As String * 32
    dmSpecVersion      As Integer
    dmDriverVersion    As Integer
    dmSize             As Integer
    dmDriverExtra      As Integer
    dmFields           As Long
    dmOrientation      As Integer
    dmPaperSize        As Integer
    dmPaperLength      As Integer
    dmPaperWidth       As Integer
    dmScale            As Integer
    dmCopies           As Integer
    dmDefaultSource    As Integer
    dmPrintQuality     As Integer
    dmColor            As Integer
    dmDuplex           As Integer
    dmYResolution      As Integer
    dmTTOption         As Integer
    dmCollate          As Integer
    dmFormName         As String * 32
    dmUnusedPadding    As Integer
    dmBitsPerPel       As Integer
    dmPelsWidth        As Long
    dmPelsHeight       As Long
    dmDisplayFlags     As Long
    dmDisplayFrequency As Long
  End Type
  
  Private Type PRINTER_DEFAULTS
    pDataType     As String
    pDevMode      As Long
    DesiredAccess As Long
  End Type
  
  Private Type PRINTER_INFO_2
    pServerName         As Long
    pPrinterName        As Long
    pShareName          As Long
    pPortName           As Long
    pDriverName         As Long
    pComment            As Long
    pLocation           As Long
    pDevMode            As Long
    pSepFile            As Long
    pPrintProcessor     As Long
    pDataType           As Long
    pParameters         As Long
    pSecurityDescriptor As Long
    Attributes          As Long
    Priority            As Long
    DefaultPriority     As Long
    StartTime           As Long
    UntilTime           As Long
    Status              As Long
    cJobs               As Long
    AveragePPM          As Long
  End Type

  Private Const DM_IN_BUFFER = 8
  Private Const DM_OUT_BUFFER = 2
  Private Const DM_ORIENTATION = &H1
  Private Const DM_PAPERSIZE = &H2

  Private Const PRINTER_ACCESS_ADMINISTER = &H4
  Private Const PRINTER_ACCESS_USE = &H8
  Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
  Private Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or _
      PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)
   
 Public Enum ntAccessLevel
      ntNone = 0&
      ntUser = PRINTER_ACCESS_USE
      ntStdRights = STANDARD_RIGHTS_REQUIRED
      ntAdmin = PRINTER_ACCESS_ADMINISTER
      ntALL = PRINTER_ALL_ACCESS
  End Enum
         
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
      (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

  Private Declare Function OpenPrinter Lib "winspool.drv" Alias _
      "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As _
      Long, pDefault As Any) As Long

  Private Declare Function ClosePrinter Lib "winspool.drv" _
      (ByVal hPrinter As Long) As Long

  Private Declare Function DocumentProperties Lib "winspool.drv" _
      Alias "DocumentPropertiesA" (ByVal hWnd As Long, ByVal hPrinter As Long, _
      ByVal pDeviceName As String, pDevModeOutput As Any, pDevModeInput As Any, _
      ByVal fMode As Long) As Long

  Private Declare Function GetPrinter Lib "winspool.drv" _
      Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, _
      pPrinter As Any, ByVal cbBuf As Long, pcbNeeded As Long) As Long

  Private Declare Function SetPrinter Lib "winspool.drv" _
      Alias "SetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, _
      pPrinter As Any, ByVal Command As Long) As Long
      
Private m_AccessLevel As ntAccessLevel
Private m_OS As ntOperatingSystem
Private m_PrinterName As String
Private m_PrevOrientation As Integer
Private m_PrevPaperSize As Integer

Public Property Let AccessLevel(ByVal iLevel As ntAccessLevel)
   m_AccessLevel = iLevel
End Property
      
Public Property Let OperatingSystem(ByVal iOS As ntOperatingSystem)
   m_OS = iOS
End Property

Public Sub ResetDeFaults()
   Dim p As Printer
   
   If m_PrinterName = "" Then Exit Sub
   If m_PrevOrientation = -1 Then Exit Sub
   If m_PrevPaperSize = -1 Then Exit Sub
           
   For Each p In Printers
      If p.DeviceName = m_PrinterName Then
         Call SetPrinterOrientation(m_PrevOrientation, m_PrevPaperSize)
         Exit For
      End If
   Next
   
   m_PrinterName = vbNullString
   m_PrevOrientation = -1
   m_PrevPaperSize = -1

End Sub

Public Function SetPrinterOrientation(ByVal eOrientation As PrinterOrientationConstants, _
      ByVal ePaperSize As PrinterPaperSize) As Boolean
   
   Dim bDevMode() As Byte
   Dim bPrinterInfo2() As Byte
   Dim hPrinter As Long
   Dim lResult As Long
   Dim nSize As Long
   Dim sPrnName As String
   Dim dm As DEVMODE
   Dim PD As PRINTER_DEFAULTS
   Dim pi2 As PRINTER_INFO_2
   
   m_PrinterName = Printer.DeviceName
   m_PrevOrientation = Printer.Orientation
   m_PrevPaperSize = Printer.PaperSize
   
   'Get device name of default printer
   sPrnName = Printer.DeviceName
  
   'PRINTER_ALL_ACCESS required under NT, because we're going to call SetPrinter
   PD.DesiredAccess = m_AccessLevel
  
   'Get a handle to the printer.
   If OpenPrinter(sPrnName, hPrinter, PD) Then
      'Get number of bytes requires for PRINTER_INFO_2 structure
      Call GetPrinter(hPrinter, 2&, 0&, 0&, nSize)
      'Create a buffer of the required size
      ReDim bPrinterInfo2(1 To nSize) As Byte
      'Fill buffer with structure
      lResult = GetPrinter(hPrinter, 2, bPrinterInfo2(1), nSize, nSize)
      'Copy fixed portion of structure into VB Type variable
      Call CopyMemory(pi2, bPrinterInfo2(1), Len(pi2))
      'Get number of bytes requires for DEVMODE structure
      nSize = DocumentProperties(0&, hPrinter, sPrnName, 0&, 0&, 0)
      'Create a buffer of the required size
      ReDim bDevMode(1 To nSize)
      'If PRINTER_INFO_2 points to a DEVMODE structure, copy it into our Buffer
      If pi2.pDevMode Then
        Call CopyMemory(bDevMode(1), ByVal pi2.pDevMode, Len(dm))
      Else
        'Otherwise, call DocumentProperties to get a DEVMODE structure
        Call DocumentProperties(0&, hPrinter, sPrnName, bDevMode(1), 0&, DM_OUT_BUFFER)
      End If
      'Copy fixed portion of structure into VB Type variable
      Call CopyMemory(dm, bDevMode(1), Len(dm))
      With dm
        ' Set new orientation
        .dmOrientation = eOrientation
        .dmPaperSize = ePaperSize
        .dmFields = DM_ORIENTATION Or DM_PAPERSIZE
      End With
      'Copy our Type back into buffer
      Call CopyMemory(bDevMode(1), dm, Len(dm))
      'Set new orientation
      Call DocumentProperties(0&, hPrinter, sPrnName, bDevMode(1), bDevMode(1), DM_IN_BUFFER Or DM_OUT_BUFFER)
      'Point PRINTER_INFO_2 at our modified DEVMODE
      pi2.pDevMode = VarPtr(bDevMode(1))
      'Set new orientation system-wide
      lResult = SetPrinter(hPrinter, 2, pi2, 0&)
      
      ' Clean up and exit
      Call ClosePrinter(hPrinter)
      SetPrinterOrientation = True
      Printer.PaperSize = ePaperSize
      Printer.Orientation = eOrientation
   Else
      SetPrinterOrientation = False
      m_PrinterName = vbNullString
      m_PrevOrientation = -1
      m_PrevPaperSize = -1
   End If
   
End Function

Public Property Get PaperSize() As Long
   
   Dim bDevMode() As Byte
   Dim bPrinterInfo2() As Byte
   Dim hPrinter As Long
   Dim lResult As Long
   Dim nSize As Long
   Dim sPrnName As String
   Dim dm As DEVMODE
   Dim PD As PRINTER_DEFAULTS
   Dim pi2 As PRINTER_INFO_2
         
   'Get device name of default printer
   sPrnName = Printer.DeviceName
  
   'PRINTER_ALL_ACCESS required under NT, because we're going to call SetPrinter
   PD.DesiredAccess = m_AccessLevel
  
   'Get a handle to the printer.
   If OpenPrinter(sPrnName, hPrinter, PD) Then
      'Get number of bytes requires for PRINTER_INFO_2 structure
      Call GetPrinter(hPrinter, 2&, 0&, 0&, nSize)
      'Create a buffer of the required size
      ReDim bPrinterInfo2(1 To nSize) As Byte
      'Fill buffer with structure
      lResult = GetPrinter(hPrinter, 2, bPrinterInfo2(1), nSize, nSize)
      'Copy fixed portion of structure into VB Type variable
      Call CopyMemory(pi2, bPrinterInfo2(1), Len(pi2))
      'Get number of bytes requires for DEVMODE structure
      nSize = DocumentProperties(0&, hPrinter, sPrnName, 0&, 0&, 0)
      'Create a buffer of the required size
      ReDim bDevMode(1 To nSize)
      'If PRINTER_INFO_2 points to a DEVMODE structure, copy it into our Buffer
      If pi2.pDevMode Then
        Call CopyMemory(bDevMode(1), ByVal pi2.pDevMode, Len(dm))
      Else
        'Otherwise, call DocumentProperties to get a DEVMODE structure
        Call DocumentProperties(0&, hPrinter, sPrnName, bDevMode(1), 0&, DM_OUT_BUFFER)
      End If
      'Copy fixed portion of structure into VB Type variable
      Call CopyMemory(dm, bDevMode(1), Len(dm))
      PaperSize = dm.dmPaperSize
      ' Clean up and exit
      Call ClosePrinter(hPrinter)
   Else
      PaperSize = -1
      m_PrinterName = vbNullString
   End If
   
End Property

Public Property Get PrinterOrientation() As Long
   
   Dim bDevMode() As Byte
   Dim bPrinterInfo2() As Byte
   Dim hPrinter As Long
   Dim lResult As Long
   Dim nSize As Long
   Dim sPrnName As String
   Dim dm As DEVMODE
   Dim PD As PRINTER_DEFAULTS
   Dim pi2 As PRINTER_INFO_2
      
   'Get device name of default printer
   sPrnName = Printer.DeviceName
  
   'PRINTER_ALL_ACCESS required under NT, because we're going to call SetPrinter
   PD.DesiredAccess = m_AccessLevel
  
   'Get a handle to the printer.
   If OpenPrinter(sPrnName, hPrinter, PD) Then
      'Get number of bytes requires for PRINTER_INFO_2 structure
      Call GetPrinter(hPrinter, 2&, 0&, 0&, nSize)
      'Create a buffer of the required size
      ReDim bPrinterInfo2(1 To nSize) As Byte
      'Fill buffer with structure
      lResult = GetPrinter(hPrinter, 2, bPrinterInfo2(1), nSize, nSize)
      'Copy fixed portion of structure into VB Type variable
      Call CopyMemory(pi2, bPrinterInfo2(1), Len(pi2))
      'Get number of bytes requires for DEVMODE structure
      nSize = DocumentProperties(0&, hPrinter, sPrnName, 0&, 0&, 0)
      'Create a buffer of the required size
      ReDim bDevMode(1 To nSize)
      'If PRINTER_INFO_2 points to a DEVMODE structure, copy it into our Buffer
      If pi2.pDevMode Then
        Call CopyMemory(bDevMode(1), ByVal pi2.pDevMode, Len(dm))
      Else
        'Otherwise, call DocumentProperties to get a DEVMODE structure
        Call DocumentProperties(0&, hPrinter, sPrnName, bDevMode(1), 0&, DM_OUT_BUFFER)
      End If
      'Copy fixed portion of structure into VB Type variable
      Call CopyMemory(dm, bDevMode(1), Len(dm))
      ' Set new orientation
      PrinterOrientation = dm.dmOrientation
      ' Clean up and exit
      Call ClosePrinter(hPrinter)
   Else
      PrinterOrientation = -1
   End If
   
End Property

Private Sub Class_Initialize()
   m_AccessLevel = 0&
   m_PrinterName = vbNullString
   m_PrevOrientation = -1
   m_PrevPaperSize = -1
End Sub


