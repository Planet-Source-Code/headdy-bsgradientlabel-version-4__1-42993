Attribute VB_Name = "modEdges"
'---------------------------------------------------------------------------------------
' Module    : modEdges
' DateTime  : 29/10/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Used for drawing rectangular borders of various styles, similar
'             to system borders but with the added bonus of being able to
'             change the colours.
'---------------------------------------------------------------------------------------

Option Explicit
Private Declare Function FillRect Lib "User32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, col As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private Const CLR_INVALID = -1
Private Const PS_SOLID = 0

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Type edgesRECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type POINTAPI
   X As Long
   Y As Long
End Type


'---------------------------------------------------------------------------------------
' Procedure : modEdges.TranslateColour
' DateTime  : 02/11/2003
' Author    : Drew (aka The Bad One)
' Purpose   : Converts an Automation colour to a Windows long colour.
'---------------------------------------------------------------------------------------
'
Private Function TranslateColour(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long
   If TranslateColor(oClr, hPal, TranslateColour) Then
       TranslateColour = CLR_INVALID
   End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : modEdges.FlatEdge
' DateTime  : 02/04/2004 19:14
' Author    : Drew
' Purpose   : Draws a flat border around an edgesRECT definition.
'---------------------------------------------------------------------------------------
' 02/04/2003   Updated with DeleteObject SelectObject(hdc, lPen).

Sub FlatEdge(ByVal hDC As Long, Box As edgesRECT, lColour As Long)
   Dim lPen As Long
   lPen = CreatePen(PS_SOLID, 1, TranslateColour(lColour))
   DeleteObject SelectObject(hDC, lPen)
   Rectangle hDC, Box.Left, Box.Top, Box.Right, Box.Bottom
   DeleteObject lPen
End Sub

' EtchedEdge()
' -----------------------------
' I WAS going to use the DrawFrame API, except it only uses system
' colours. But there's a really easy way to obtain the same effect.

Sub EtchedEdge(ByVal hDC As Long, Box As RECT, hdColour As Long, _
   sdColour As Long)
   
   Dim lPen As Long
   
   lPen = CreatePen(PS_SOLID, 1, TranslateColour(hdColour))
   SelectObject hDC, lPen
   Rectangle hDC, Box.Left + 1, Box.Top + 1, Box.Right, Box.Bottom
   DeleteObject lPen
   
   lPen = CreatePen(PS_SOLID, 1, TranslateColour(sdColour))
   SelectObject hDC, lPen
   Rectangle hDC, Box.Left, Box.Top, Box.Right - 1, Box.Bottom - 1
   DeleteObject lPen
   
   SetPixel hDC, Box.Right - 1, Box.Top, TranslateColour(hdColour)
   SetPixel hDC, Box.Left, Box.Bottom - 1, TranslateColour(hdColour)
End Sub

' ThinEdge()
' -----------------------------
' This time there's no choice - we have to use lines.
Sub ThinEdge(ByVal hDC As Long, Box As RECT, lightColour As Long, _
   darkColour As Long)
   
   Dim lPen As Long
   Dim oldPoint As POINTAPI
   
   lPen = CreatePen(PS_SOLID, 1, TranslateColour(lightColour))
   SelectObject hDC, lPen
   MoveToEx hDC, Box.Right, Box.Top, oldPoint
   LineTo hDC, Box.Left, Box.Top
   LineTo hDC, Box.Left, Box.Bottom - 1
   DeleteObject lPen
   
   lPen = CreatePen(PS_SOLID, 1, TranslateColour(darkColour))
   SelectObject hDC, lPen
   LineTo hDC, Box.Right - 1, Box.Bottom - 1
   LineTo hDC, Box.Right - 1, Box.Top
   DeleteObject lPen
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ThickEdge
' DateTime  : 29/10/2003 08:35
' Author    : Drew (aka The Bad One)
' Purpose   : Draws those normally nasty thick 3D borders. If no colours are
'             specified they default to the system ones.
'---------------------------------------------------------------------------------------
'
Sub ThickEdge(ByVal hDC As Long, rctBox As RECT, Optional lLightestColour As Long, _
   Optional lLightColour As Long, Optional lDarkColour As Long, _
   Optional lDarkestColour As Long, Optional lFillColour As Long = -1)
   
   Dim hPen As Long, hBrush As Long
   Dim oldPoint As POINTAPI
   
   ' Default colours
   ' ------------------------------------------------------------------
   If IsMissing(lLightestColour) Then
      lLightestColour = vb3DLight
   End If
   If IsMissing(lLightColour) Then
      lLightColour = vb3DHighlight
   End If
   If IsMissing(lDarkColour) Then
      lDarkColour = vb3DShadow
   End If
   If IsMissing(lDarkestColour) Then
      lDarkestColour = vb3DDKShadow
   End If
   
   If lFillColour <> -1 Then
      hBrush = CreateSolidBrush(TranslateColour(lFillColour))
      DeleteObject SelectObject(hDC, hBrush)
      FillRect hDC, rctBox, hBrush
      DeleteObject hBrush
   End If
   
   ' Draw the borders
   ' ------------------------------------------------------------------
   hPen = CreatePen(PS_SOLID, 1, TranslateColour(lLightestColour))
   DeleteObject SelectObject(hDC, hPen)
   MoveToEx hDC, rctBox.Right, rctBox.Top, oldPoint
   LineTo hDC, rctBox.Left, rctBox.Top
   LineTo hDC, rctBox.Left, rctBox.Bottom - 1
   DeleteObject hPen
   
   hPen = CreatePen(PS_SOLID, 1, TranslateColour(lDarkestColour))
   DeleteObject SelectObject(hDC, hPen)
   LineTo hDC, rctBox.Right - 1, rctBox.Bottom - 1
   LineTo hDC, rctBox.Right - 1, rctBox.Top
   DeleteObject hPen
   
   hPen = CreatePen(PS_SOLID, 1, TranslateColour(lLightColour))
   DeleteObject SelectObject(hDC, hPen)
   MoveToEx hDC, rctBox.Right - 2, rctBox.Top + 1, oldPoint
   LineTo hDC, rctBox.Left + 1, rctBox.Top + 1
   LineTo hDC, rctBox.Left + 1, rctBox.Bottom - 2
   DeleteObject hPen
   
   hPen = CreatePen(PS_SOLID, 1, TranslateColour(lDarkColour))
   DeleteObject SelectObject(hDC, hPen)
   LineTo hDC, rctBox.Right - 2, rctBox.Bottom - 2
   LineTo hDC, rctBox.Right - 2, rctBox.Top + 1
   DeleteObject hPen
End Sub

