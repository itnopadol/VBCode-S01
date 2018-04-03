Attribute VB_Name = "MDDrawingUtils"
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long



Public Sub DrawGradient( _
      ByVal hdc As Long, _
      ByRef rct As RECT, _
      ByVal lEndColour As Long, _
      ByVal lStartColour As Long, _
      ByVal bVertical As Boolean _
   )
Dim lStep As Long
Dim lPos As Long, lSize As Long
Dim bRGB(1 To 3) As Integer
Dim bRGBStart(1 To 3) As Integer
Dim dR(1 To 3) As Double
Dim dPos As Double, d As Double
Dim hBr As Long
Dim tR As RECT
   
   LSet tR = rct
   If bVertical Then
      lSize = (tR.Bottom - tR.Top)
   Else
      lSize = (tR.Right - tR.Left)
   End If
   lStep = lSize \ 255
   If (lStep < 3) Then
       lStep = 3
   End If
       
   bRGB(1) = lStartColour And &HFF&
   bRGB(2) = (lStartColour And &HFF00&) \ &H100&
   bRGB(3) = (lStartColour And &HFF0000) \ &H10000
   bRGBStart(1) = bRGB(1): bRGBStart(2) = bRGB(2): bRGBStart(3) = bRGB(3)
   dR(1) = (lEndColour And &HFF&) - bRGB(1)
   dR(2) = ((lEndColour And &HFF00&) \ &H100&) - bRGB(2)
   dR(3) = ((lEndColour And &HFF0000) \ &H10000) - bRGB(3)
        
   For lPos = lSize To 0 Step -lStep
      ' Draw bar:
      If bVertical Then
         tR.Top = tR.Bottom - lStep
      Else
         tR.Left = tR.Right - lStep
      End If
      If tR.Top < rct.Top Then
         tR.Top = rct.Top
      End If
      If tR.Left < rct.Left Then
         tR.Left = rct.Left
      End If
      
      'Debug.Print tR.Right, tR.left, (bRGB(3) * &H10000 + bRGB(2) * &H100& + bRGB(1))
      hBr = CreateSolidBrush((bRGB(3) * &H10000 + bRGB(2) * &H100& + bRGB(1)))
      FillRect hdc, tR, hBr
      DeleteObject hBr
            
      ' Adjust colour:
      dPos = ((lSize - lPos) / lSize)
      If bVertical Then
         tR.Bottom = tR.Top
         bRGB(1) = bRGBStart(1) + dR(1) * dPos
         bRGB(2) = bRGBStart(2) + dR(2) * dPos
         bRGB(3) = bRGBStart(3) + dR(3) * dPos
      Else
         tR.Right = tR.Left
         bRGB(1) = bRGBStart(1) + dR(1) * dPos
         bRGB(2) = bRGBStart(2) + dR(2) * dPos
         bRGB(3) = bRGBStart(3) + dR(3) * dPos
      End If
      
   Next lPos

End Sub


