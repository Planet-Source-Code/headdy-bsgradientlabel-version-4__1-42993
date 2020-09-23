VERSION 5.00
Begin VB.UserControl bsGradientLabel 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   PropertyPages   =   "bsGradientLabel.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "bsGradientLabel.ctx":000C
End
Attribute VB_Name = "bsGradientLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'---------------------------------------------------------------------------------------
' Module    : bsGradientLabel
' DateTime  : 02/04/2004 17:18
' Author    : Drew (aka Headdy / The Bad One)
'             ©2002-2004 BadSoft Entertainment, all rights reserved.
'             http://www.badsoft.co.uk
'             drew@badsoft.co.uk
' Version   : 4.0
'---------------------------------------------------------------------------------------
' NOTE      The spelling of 'fount' is intended - I have an English English thing
'           going here. Not many people know that 'font' is actually an American
'           spelling. Just as 'colour' is 'color', and 'flavour' is 'flavor'.
'---------------------------------------------------------------------------------------

Option Explicit

'Default Property Values:
Const m_def_GradientAngle = 0
Const m_def_Version = ""
Const m_def_NonTTError = True
Const m_def_Offset = 6
Const m_def_WordWrap = 0
Const m_def_TextShadowYOffset = 2
Const m_def_TextShadowXOffset = 2
Const m_def_BorderStyle = 0
Const m_def_HighlightColour = vb3DHighlight
Const m_def_HighlightDKColour = vb3DLight
Const m_def_ShadowColour = vb3DShadow
Const m_def_ShadowDKColour = vb3DDKShadow
Const m_def_FlatBorderColour = vbBlack
Const m_def_TextShadowColour = vbBlack
Const m_def_TextShadow = False
Const m_def_LabelType = 0
Const m_def_CaptionAlignment = 0
Const m_def_Colour1 = 0
Const m_def_Colour2 = vbBlue
Const m_def_Colour3 = vbYellow
Const m_def_Colour4 = vbRed
Const m_def_CaptionColour = vbWhite
Const m_def_GradientType = 1

'Property Variables:
Dim m_GradientAngle As Single
Dim m_Version As String
Dim m_NonTTError As Boolean
Dim m_Offset As Integer
Dim m_WordWrap As Boolean
Dim m_TextShadowYOffset As Integer
Dim m_TextShadowXOffset As Integer
Dim m_BorderStyle As bsBorderStyle
Dim m_HighlightColour As OLE_COLOR
Dim m_HighlightDKColour As OLE_COLOR
Dim m_ShadowColour As OLE_COLOR
Dim m_ShadowDKColour As OLE_COLOR
Dim m_FlatBorderColour As OLE_COLOR
Dim m_TextShadowColour As OLE_COLOR
Dim m_TextShadow As Boolean
Dim m_LabelType As bsLabelType
Dim m_CaptionAlignment As bsCaptionAlign
Dim m_Colour1 As OLE_COLOR
Dim m_Colour2 As OLE_COLOR
Dim m_Colour3 As OLE_COLOR
Dim m_Colour4 As OLE_COLOR
Dim m_CaptionColour As OLE_COLOR
Dim m_GradientType As bsGradient
Dim m_Caption As String
Dim m_Fount As StdFont


' API CALLS
'-------------------------------------
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, col As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextAlign Lib "gdi32" (ByVal hDC As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawText Lib "User32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpString As String, ByVal nCount As Long, lpRect As RECT, ByVal uFormat As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hDC As Long, lpMetrics As TEXTMETRIC) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long


' CONSTANTS

' CreateFontIndirect()
'-------------------------------------
Private Const PROOF_QUALITY = 2

' DrawText()
'-------------------------------------
Private Const TA_CENTER = 6
Private Const TA_LEFT = 0
Private Const TA_RIGHT = 2
Private Const DT_CALCRECT = &H400
Private Const DT_WORDBREAK = &H10
Private Const DT_NOCLIP = &H100

' GetTextMetrics()
'-------------------------------------
Private Const TMPF_TRUETYPE = &H4

' CreateFontIndirect()
'-------------------------------------
Private Const LF_FACESIZE = 32

' CreatePen()
'-------------------------------------
Private Const PS_NULL = 5


' TYPES
Private Type POINTAPI
   X As Long
   Y As Long
End Type

Public Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To LF_FACESIZE) As Byte
End Type

Public Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type


' ENUMS
Enum bsCaptionAlign
   [AlignLeft]
   [AlignCentre]
   [AlignRight]
End Enum

Enum bsGradient
   [Plain]
   [2 Way]
   [4 Way]
End Enum

Enum bsLabelType
   [Horizontal] = 0
   [Vertical 90°] = 1
   [Vertical 270°] = 2
End Enum

Enum bsBorderStyle
   [None]
   [Flat]
   [Raised Thin]
   [Raised 3D]
   [Sunken Thin]
   [Sunken 3D]
   [Etched]
   [Bump]
End Enum
'Event Declarations:
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Attribute MouseUp.VB_UserMemId = -607
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Attribute MouseMove.VB_UserMemId = -606
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Attribute MouseDown.VB_UserMemId = -605
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Attribute DblClick.VB_UserMemId = -601
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600


'---------------------------------------------------------------------------------------
' Procedure : bsGradientLabel.DrawLabel
' DateTime  : 03/04/2004 23:40
' Author    : Drew
' Purpose   : This sub draws the background of the label first, then calls
'             other routines to do the text and border.
'---------------------------------------------------------------------------------------
' 03/04/2004   Changed the gradient drawing routines to use out new class and
'              module.
'---------------------------------------------------------------------------------------

Private Sub DrawLabel()

   Dim cGradient As clsGradient
   
   Cls
   ScaleMode = vbPixels
   AutoRedraw = True
   
' ----------------------------
' Drawing the gradient
' ----------------------------

   Set cGradient = New clsGradient
   
   Select Case m_GradientType
      Case [Plain]
         Dim hBrush As Long, hPen As Long
         hBrush = CreateSolidBrush(TranslateColour(m_Colour1))
         hPen = CreatePen(PS_NULL, 1, vbBlack)
         DeleteObject SelectObject(UserControl.hDC, hPen)
         DeleteObject SelectObject(UserControl.hDC, hBrush)
         Rectangle UserControl.hDC, 0, 0, ScaleWidth, ScaleHeight
         DeleteObject hPen
         DeleteObject hBrush
      Case [2 Way]
         With cGradient
            .Angle = m_GradientAngle
            .Color1 = TranslateColour(m_Colour1)
            .Color2 = TranslateColour(m_Colour2)
            .Draw UserControl.hWnd, UserControl.hDC
         End With
      Case [4 Way]
         ' Note how colours 3 and 4 are swapped around - this is the way
         ' the four-way gradient module works.
         modGradient.DrawGradient UserControl.hDC, 0, 0, ScaleWidth, ScaleHeight, _
            TranslateColour(m_Colour1), TranslateColour(m_Colour2), _
            TranslateColour(m_Colour4), TranslateColour(m_Colour3)
   End Select
   
   Set cGradient = Nothing
   Refresh
   
 
' ----------------------------
' Draw the text
' ----------------------------
   Select Case m_LabelType
      Case [Horizontal]
         DrawLabelText 0
      Case [Vertical 90°]
         DrawLabelText 90
      Case [Vertical 270°]
         DrawLabelText 270
   End Select
   
' ----------------------------
' ... and the edges
' ----------------------------
   DrawEdges
   
End Sub


'---------------------------------------------------------------------------------------
' Procedure : bsGLabel.DrawLabelText
' DateTime  : 02/04/2004 17:29
' Author    : Drew
' Purpose   : Draws the text for the control.
'---------------------------------------------------------------------------------------
' 06/04/2004   Adjusted a little piece of code so that the Offset applies to two
'              sides instead of just one.
'              Finished vertical multiline text support, and reduced the amount
'              of code needed as well. Also nipped a rather embarassing bug
'              concerning single-line text and text shadows.
' 04/04/2004   04:04:04 on 04/04/04. Figured out just hours before anyone else.
'              A major rewrite is needed concerning vertical text, as it no longer
'              works for some reason. We're also aiming to get rid of code
'              redundancy.
' 02/04/2004   Had to rewrite the code because we used a PictureBox
'              instead of a UserControl, and rewrote it again after rewriting
'              some of the clsGradient code. @:)
'---------------------------------------------------------------------------------------

Private Sub DrawLabelText(ByVal Angle As Integer)

   On Error GoTo GetOut
   
   Dim lfFount As LOGFONT, hPrevFount As Long, hFount As Long
   Dim lColour As Long
   Dim tmpRect As RECT
   Dim iBlockHeight As Integer, iLineHeight As Integer
   Dim XStart As Integer, YStart As Integer
   Dim I As Integer, N As Integer, iMaximumLines As Integer
   Dim tmpArray() As Byte
   Dim tmpCaption As String
   Dim MLines() As String
   Dim MLAlign As Long
   Dim RectWidth As Integer
   
' ----------------------------
' Check for no caption!
' ----------------------------
   If m_Caption = "" Then Exit Sub
   
' ----------------------------
' Set up fount
' ----------------------------
' To get the height of the fount (in pixels) using the TextHeight
' method, we need to set the UserControl fount to the one the user
' specified.
' The fount name is converted to a byte array for API reasons, and
' then attached to the target device context (DC). We also fix the
' text alignment setting.

   ScaleMode = vbPixels
   FontName = m_Fount.Name
   FontSize = m_Fount.Size
   
   On Error GoTo 0
   tmpArray = StrConv(m_Fount.Name & vbNullString, vbFromUnicode)
   For I = 0 To UBound(tmpArray)
       lfFount.lfFaceName(I + 1) = tmpArray(I)
   Next
   
   With lfFount
      .lfEscapement = 10 * Angle
      .lfHeight = (m_Fount.Size * -20) / Screen.TwipsPerPixelY
      .lfItalic = m_Fount.Italic
      .lfUnderline = m_Fount.Underline
      .lfWeight = IIf(m_Fount.Bold, 700, 0)
      .lfQuality = PROOF_QUALITY
   End With
   
   hFount = CreateFontIndirect(lfFount)
   hPrevFount = SelectObject(hDC, hFount)
      
' ----------------------------
' Get text height
' ----------------------------
' To get the height of the characters in the fount, we can use the UserControl's
' TextHeight method. To get the height of the block of text, we use the DrawText
' API (with the DT_CALCRECT flag set).
   
   If m_LabelType = [Horizontal] Then
      tmpRect.Left = m_Offset
      tmpRect.Right = ScaleWidth - m_Offset
   Else
      tmpRect.Bottom = ScaleWidth
   End If
   DrawText UserControl.hDC, m_Caption, Len(m_Caption), tmpRect, _
      IIf(m_WordWrap, DT_WORDBREAK, 0) + DT_CALCRECT
      
   iBlockHeight = tmpRect.Bottom
   iLineHeight = UserControl.TextHeight(" ")
       
   If m_LabelType = [Horizontal] Then
      Select Case m_CaptionAlignment
         Case [AlignLeft]
            UserControl.CurrentX = m_Offset
         Case [AlignRight]
            UserControl.CurrentX = ScaleWidth - m_Offset
         Case [AlignCentre]
            UserControl.CurrentX = ScaleWidth / 2
      End Select
      UserControl.CurrentY = (ScaleHeight - iBlockHeight) / 2
      
   Else
      Select Case m_CaptionAlignment
         Case [AlignLeft]
            UserControl.CurrentY = ScaleHeight - m_Offset
         Case [AlignRight]
            UserControl.CurrentY = m_Offset
         Case [AlignCentre]
            UserControl.CurrentY = ScaleHeight / 2
      End Select
      UserControl.CurrentX = (ScaleWidth - iBlockHeight) / 2
   End If
   
   
' ----------------------------
' Draw text + text shadows
' ----------------------------
' We need to use three different methods for drawing the text,
' depending on WordWrap and LabelType.
   
' Our job is made infinitely easy if the label is a non-
' wordwrapped one. We just use a single Print statement,
' regardless of whether it's horizontal or vertical.

' The variables px and py are needed because after each Print
' command the UserControl's CurrentX and CurrentY completely
' reset themselves.

' For a horizontal wordwrapped label, we can use the DrawText
' API call easily.

' But for vertical wordwrapped labels, we have to do the
' word wrapping ourselves! I tried to use the DrawText API call,
' but the lines aligned themselves to the left of the rect and
' consequently drew themselves over each other. So, we go
' through the whole caption and pick out the lines based on
' spaces and carriage returns. This took some doing, so please
' show your appreciation and leave feedback.

   If WordWrap = False Then
      ' ----------------------------
      ' Non-wordwrapped label
      ' ----------------------------
      Select Case m_CaptionAlignment
         Case [AlignCentre]
            SetTextAlign UserControl.hDC, TA_CENTER
         Case [AlignLeft]
            SetTextAlign UserControl.hDC, TA_LEFT
         Case [AlignRight]
            SetTextAlign UserControl.hDC, TA_RIGHT
      End Select
      
      With UserControl
         ' Caption shadow
         If m_TextShadow = True Then
            XStart = .CurrentX: YStart = .CurrentY
            SetTextColor UserControl.hDC, TranslateColour(m_TextShadowColour)
            With UserControl
               If m_LabelType = [Vertical 270°] Then
                  .CurrentX = .CurrentX + UserControl.TextHeight(m_Caption)
                  .CurrentY = ScaleHeight - CurrentY
               End If
            End With
            .CurrentX = .CurrentX + m_TextShadowXOffset
            .CurrentY = .CurrentY + m_TextShadowYOffset
            UserControl.Print m_Caption
            .CurrentX = XStart: .CurrentY = YStart
         End If
         ' Caption
         With UserControl
            If m_LabelType = [Vertical 270°] Then
               .CurrentX = .CurrentX + UserControl.TextHeight(m_Caption)
               .CurrentY = ScaleHeight - CurrentY
            End If
         End With
         SetTextColor .hDC, TranslateColour(m_CaptionColour)
         UserControl.Print m_Caption
      End With
      
   ElseIf m_LabelType = [Horizontal] Then
      ' ----------------------------
      ' Horizontal word-wrapped/multiline label
      ' ----------------------------
      ShiftRect tmpRect, 0, CurrentY
      Select Case m_CaptionAlignment
         Case AlignLeft
            MLAlign = TA_LEFT
         Case AlignRight
            MLAlign = TA_RIGHT
            ShiftRect tmpRect, ScaleWidth - m_Offset * 2, 0
         Case AlignCentre
            MLAlign = TA_CENTER
            ShiftRect tmpRect, ScaleWidth / 2 - m_Offset, 0
      End Select
      SetTextAlign UserControl.hDC, MLAlign
      
      ' Caption shadow
      If m_TextShadow = True Then
         ShiftRect tmpRect, m_TextShadowXOffset, m_TextShadowYOffset
         SetTextColor UserControl.hDC, TranslateColour(m_TextShadowColour)
         DrawText UserControl.hDC, m_Caption, Len(m_Caption), _
            tmpRect, DT_WORDBREAK + DT_NOCLIP
         ShiftRect tmpRect, -m_TextShadowXOffset, -m_TextShadowYOffset
      End If
      ' Caption
      SetTextColor UserControl.hDC, TranslateColour(m_CaptionColour)
      DrawText UserControl.hDC, m_Caption, Len(m_Caption), _
         tmpRect, DT_WORDBREAK + DT_NOCLIP
            
   Else
      ' ----------------------------
      ' Vertical word-wrapped/multiline label
      ' ----------------------------
      ' The most confusing and complicated part of the control.
      
      RectWidth = ScaleHeight - m_Offset * 2
      tmpCaption = m_Caption
      I = 1

      ' First of all we need to divide the caption up into separate lines.
      ' This bit parses the caption for spaces and carriage returns, and stores
      ' the lines in a string array.
      Dim iSpacePos As Integer
      While I < Len(tmpCaption)
         If Mid(tmpCaption, I, 1) = vbCr Then
            ' Split the line at the carriage return.
            N = N + 1
            ReDim Preserve MLines(1 To N)
            MLines(N) = Left(tmpCaption, I - 1)
            tmpCaption = Right(tmpCaption, Len(tmpCaption) - I - 1)
            I = 1
         ElseIf Mid(tmpCaption, I, 1) = " " Then
            ' Make a note of the position of the space.
            iSpacePos = I
            I = I + 1
         ElseIf TextWidth(Left(tmpCaption, I)) > RectWidth Then
            ' Text is too long for the rectangle - if there is a space in
            ' the preceding text split the line at that point, otherwise
            ' split the line at the last character.
            N = N + 1
            ReDim Preserve MLines(1 To N)
            If iSpacePos > 0 Then
               MLines(N) = Trim(Left(tmpCaption, iSpacePos - 1))
               tmpCaption = Trim(Right(tmpCaption, Len(tmpCaption) - iSpacePos))
               iSpacePos = 0
            Else
               MLines(N) = Trim(Left(tmpCaption, I - 1))
               tmpCaption = Trim(Right(tmpCaption, Len(tmpCaption) - I))
            End If
            I = 1
         Else
            ' Keep going...
            I = I + 1
         End If
      Wend
      ' If there's any text left over, put it into a last line.
      If Len(tmpCaption) > 0 Then
         N = N + 1
         ReDim Preserve MLines(1 To N)
         MLines(N) = tmpCaption
      End If

      iMaximumLines = ScaleWidth / iLineHeight
      iMaximumLines = IIf(iMaximumLines > UBound(MLines()), _
         UBound(MLines()), iMaximumLines)

      
      For I = 1 To UBound(MLines)
         Select Case m_CaptionAlignment
            Case [AlignCentre]
               SetTextAlign UserControl.hDC, TA_CENTER
               CurrentY = ScaleHeight / 2 'true for both
            Case [AlignLeft]
               SetTextAlign UserControl.hDC, TA_LEFT
               CurrentY = IIf(m_LabelType = [Vertical 90°], _
                  ScaleHeight - m_Offset, m_Offset)
            Case [AlignRight]
               SetTextAlign UserControl.hDC, TA_RIGHT
               CurrentY = IIf(m_LabelType = [Vertical 90°], _
                  m_Offset, ScaleHeight - m_Offset)
         End Select

         CurrentX = (ScaleWidth - iLineHeight * UBound(MLines)) / 2 + (I - 1) * iLineHeight
         If m_LabelType = [Vertical 270°] Then
            CurrentX = ScaleHeight - CurrentX
         End If
         XStart = CurrentX
         YStart = CurrentY
         
         ' Draw caption shadow (if applicable) and caption. The annoying thing
         ' is, because this is being done manually, CurrentX and CurrentY are
         ' reset after each Print statement. This bastard annoyance led to the
         ' code redundancy in the first place.
         If m_TextShadow Then
            CurrentX = CurrentX + m_TextShadowXOffset
            CurrentY = CurrentY + m_TextShadowYOffset
            SetTextColor UserControl.hDC, TranslateColour(m_TextShadowColour)
            Print MLines(I)
            CurrentX = XStart
            CurrentY = YStart
         End If
         SetTextColor UserControl.hDC, TranslateColour(m_CaptionColour)
         Print MLines(I)
      Next
   End If
   
   
' ----------------------------
' Clean up and restore original fount.
' ----------------------------
   hFount = SelectObject(UserControl.hDC, hPrevFount)
   DeleteObject hFount
   
   UserControl.Refresh
   
GetOut:
   Exit Sub

End Sub

' DrawEdges()
' ------------------------------
' Alongside many BadSoft controls, the bsGradientLabel has 7
' colour-customisable edge styles.

Sub DrawEdges()

   Dim lPen As Long
   Dim rctControl As edgesRECT
   Dim CurPos As POINTAPI
   
   If m_BorderStyle = None Then Exit Sub
   
   With rctControl
      .Right = ScaleWidth
      .Bottom = ScaleHeight
   End With
   
   Select Case m_BorderStyle
      Case [Flat]
         modEdges.FlatEdge UserControl.hDC, rctControl, m_FlatBorderColour
'         lPen = CreatePen(0, 0, TranslateColour(m_FlatBorderColour))
'         DeleteObject SelectObject(hdc, lPen)
'         Rectangle UserControl.hdc, 0, 0, ScaleWidth, ScaleHeight
'         DeleteObject lPen
      
      Case [Raised Thin]
         MoveToEx UserControl.hDC, ScaleWidth, 0, CurPos
         lPen = CreatePen(0, 0, TranslateColour(m_HighlightColour))
         DeleteObject SelectObject(hDC, lPen)
         LineTo UserControl.hDC, 0, 0
         LineTo UserControl.hDC, 0, ScaleHeight - 1
         DeleteObject lPen
         lPen = CreatePen(0, 0, TranslateColour(m_ShadowColour))
         DeleteObject SelectObject(hDC, lPen)
         LineTo UserControl.hDC, ScaleWidth - 1, ScaleHeight - 1
         LineTo UserControl.hDC, ScaleWidth - 1, 0
         DeleteObject lPen
         
      Case [Sunken Thin]
         MoveToEx UserControl.hDC, ScaleWidth, 0, CurPos
         lPen = CreatePen(0, 0, TranslateColour(m_ShadowColour))
         DeleteObject SelectObject(hDC, lPen)
         LineTo UserControl.hDC, 0, 0
         LineTo UserControl.hDC, 0, ScaleHeight - 1
         DeleteObject lPen
         lPen = CreatePen(0, 0, TranslateColour(m_HighlightColour))
         DeleteObject SelectObject(hDC, lPen)
         LineTo UserControl.hDC, ScaleWidth - 1, ScaleHeight - 1
         LineTo UserControl.hDC, ScaleWidth - 1, 0
         DeleteObject lPen
   
      Case [Raised 3D]
         MoveToEx UserControl.hDC, ScaleWidth, 0, CurPos
         lPen = CreatePen(0, 0, TranslateColour(m_HighlightColour))
         DeleteObject SelectObject(hDC, lPen)
         LineTo UserControl.hDC, 0, 0
         LineTo UserControl.hDC, 0, ScaleHeight - 1
         DeleteObject lPen
         lPen = CreatePen(0, 0, TranslateColour(m_ShadowDKColour))
         DeleteObject SelectObject(hDC, lPen)
         LineTo UserControl.hDC, ScaleWidth - 1, ScaleHeight - 1
         LineTo UserControl.hDC, ScaleWidth - 1, -1
         DeleteObject lPen
         MoveToEx UserControl.hDC, ScaleWidth - 2, 1, CurPos
         lPen = CreatePen(0, 0, TranslateColour(m_HighlightDKColour))
         DeleteObject SelectObject(hDC, lPen)
         LineTo UserControl.hDC, 1, 1
         LineTo UserControl.hDC, 1, ScaleHeight - 2
         DeleteObject lPen
         lPen = CreatePen(0, 0, TranslateColour(m_ShadowColour))
         DeleteObject SelectObject(hDC, lPen)
         LineTo UserControl.hDC, ScaleWidth - 2, ScaleHeight - 2
         LineTo UserControl.hDC, ScaleWidth - 2, 0
         DeleteObject lPen
   
      Case [Sunken 3D]
         MoveToEx UserControl.hDC, ScaleWidth, 0, CurPos
         lPen = CreatePen(0, 0, TranslateColour(m_ShadowDKColour))
         DeleteObject SelectObject(hDC, lPen)
         LineTo UserControl.hDC, 0, 0
         LineTo UserControl.hDC, 0, ScaleHeight - 1
         DeleteObject lPen
         lPen = CreatePen(0, 0, TranslateColour(m_HighlightColour))
         DeleteObject SelectObject(hDC, lPen)
         LineTo UserControl.hDC, ScaleWidth - 1, ScaleHeight - 1
         LineTo UserControl.hDC, ScaleWidth - 1, -1
         DeleteObject lPen
         MoveToEx UserControl.hDC, ScaleWidth - 2, 1, CurPos
         lPen = CreatePen(0, 0, TranslateColour(m_ShadowColour))
         DeleteObject SelectObject(hDC, lPen)
         LineTo UserControl.hDC, 1, 1
         LineTo UserControl.hDC, 1, ScaleHeight - 2
         DeleteObject lPen
         lPen = CreatePen(0, 0, TranslateColour(m_HighlightDKColour))
         DeleteObject SelectObject(hDC, lPen)
         LineTo UserControl.hDC, ScaleWidth - 2, ScaleHeight - 2
         LineTo UserControl.hDC, ScaleWidth - 2, 0
         DeleteObject lPen
   
      Case [Etched]
         MoveToEx UserControl.hDC, ScaleWidth, 0, CurPos
         lPen = CreatePen(0, 0, TranslateColour(m_ShadowColour))
         DeleteObject SelectObject(hDC, lPen)
         LineTo UserControl.hDC, 0, 0
         LineTo UserControl.hDC, 0, ScaleHeight - 1
         DeleteObject lPen
         lPen = CreatePen(0, 0, TranslateColour(m_HighlightColour))
         DeleteObject SelectObject(hDC, lPen)
         LineTo UserControl.hDC, ScaleWidth - 1, ScaleHeight - 1
         LineTo UserControl.hDC, ScaleWidth - 1, -1
         DeleteObject lPen
         MoveToEx UserControl.hDC, ScaleWidth - 2, 1, CurPos
         lPen = CreatePen(0, 0, TranslateColour(m_HighlightColour))
         DeleteObject SelectObject(hDC, lPen)
         LineTo UserControl.hDC, 1, 1
         LineTo UserControl.hDC, 1, ScaleHeight - 2
         DeleteObject lPen
         lPen = CreatePen(0, 0, TranslateColour(m_ShadowColour))
         DeleteObject SelectObject(hDC, lPen)
         LineTo UserControl.hDC, ScaleWidth - 2, ScaleHeight - 2
         LineTo UserControl.hDC, ScaleWidth - 2, 0
         DeleteObject lPen
   
      Case [Bump]
         MoveToEx UserControl.hDC, ScaleWidth, 0, CurPos
         lPen = CreatePen(0, 0, TranslateColour(m_HighlightColour))
         DeleteObject SelectObject(hDC, lPen)
         LineTo UserControl.hDC, 0, 0
         LineTo UserControl.hDC, 0, ScaleHeight - 1
         DeleteObject lPen
         lPen = CreatePen(0, 0, TranslateColour(m_ShadowColour))
         DeleteObject SelectObject(hDC, lPen)
         LineTo UserControl.hDC, ScaleWidth - 1, ScaleHeight - 1
         LineTo UserControl.hDC, ScaleWidth - 1, -1
         DeleteObject lPen
         MoveToEx UserControl.hDC, ScaleWidth - 2, 1, CurPos
         lPen = CreatePen(0, 0, TranslateColour(m_ShadowColour))
         DeleteObject SelectObject(hDC, lPen)
         LineTo UserControl.hDC, 1, 1
         LineTo UserControl.hDC, 1, ScaleHeight - 2
         DeleteObject lPen
         lPen = CreatePen(0, 0, TranslateColour(m_HighlightColour))
         DeleteObject SelectObject(hDC, lPen)
         LineTo UserControl.hDC, ScaleWidth - 2, ScaleHeight - 2
         LineTo UserControl.hDC, ScaleWidth - 2, 0
         DeleteObject lPen
   End Select
   Refresh
End Sub

' ShiftRect()
' ------------------------------
' A sub for quickly shifting a rect by a certain amount in
' either direction.

Private Sub ShiftRect(ByRef whichRect As RECT, X As Integer, Y As Integer)
   whichRect.Top = whichRect.Top + Y
   whichRect.Bottom = whichRect.Bottom + Y
   whichRect.Right = whichRect.Right + X
   whichRect.Left = whichRect.Left + X
End Sub

' TranslateColour()
' ------------------------------
' This translates any long value into an RGB colour, for use
' with drawing functions. I object to being forced to use
' American words so I renamed it myself.

Function TranslateColour(lColour As Long) As Long
   TranslateColor lColour, 0, TranslateColour
End Function

' ShowAbout()
' ------------------------------
' A small sub for showing the About screen.

Sub ShowAbout()
   frmAbout.Show vbModal
End Sub


' IsFountTrueType()
' ------------------------------
' At last, a way of telling if a font is TrueType or not. This
' came from James Crowley.

Public Function IsFountTrueType(sFontName As String) As Boolean
    Dim lf As LOGFONT
    Dim tm As TEXTMETRIC
    Dim oldfount As Long, newfount As Long
    Dim tmpArray() As Byte
    Dim dummy As Long
    Dim I As Integer
    
    tmpArray = StrConv(sFontName & vbNullString, vbFromUnicode)
    For I = 0 To UBound(tmpArray)
        lf.lfFaceName(I + 1) = tmpArray(I)
    Next
    
    newfount = CreateFontIndirect(lf)
    oldfount = SelectObject(UserControl.hDC, newfount)
    dummy = GetTextMetrics(UserControl.hDC, tm)
    IsFountTrueType = (tm.tmPitchAndFamily And TMPF_TRUETYPE)
    dummy = SelectObject(UserControl.hDC, oldfount)
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get GradientType() As bsGradient
Attribute GradientType.VB_Description = "The direction the gradient follows."
Attribute GradientType.VB_ProcData.VB_Invoke_Property = ";Appearance"
   GradientType = m_GradientType
End Property

Public Property Let GradientType(ByVal New_GradientType As bsGradient)
   m_GradientType = New_GradientType
   PropertyChanged "GradientType"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get Caption() As String
Attribute Caption.VB_Description = "The text the GradientLabel contains."
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Caption.VB_MemberFlags = "200"
   Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
   m_Caption = New_Caption
   PropertyChanged "Caption"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0

' Fount()
' -----------------------------
' A check is made when setting the fount to see if the user has
' selected a Vertical type label and a non-TrueType font.
Public Property Get Fount() As StdFont
Attribute Fount.VB_Description = "The fount used by the Caption property."
Attribute Fount.VB_ProcData.VB_Invoke_Property = ";Font"
   Set Fount = m_Fount
End Property

Public Property Set Fount(ByVal New_Fount As StdFont)
   Set m_Fount = New_Fount

   If m_LabelType <> [Horizontal] And IsFountTrueType(New_Fount.Name) = False Then
      If m_NonTTError Then
         MsgBox "The LabelType property can only be Vertical when the Fount is a TrueType fount.", vbExclamation
      End If
      LabelType = [Horizontal]
   End If
   
   PropertyChanged "Fount"
   DrawLabel
End Property

Private Sub UserControl_Click()
   RaiseEvent Click
End Sub


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
   m_GradientType = m_def_GradientType
   m_Caption = UserControl.Extender.Name
   Set m_Fount = Ambient.Font
   m_CaptionColour = m_def_CaptionColour
   m_Colour1 = m_def_Colour1
   m_Colour2 = m_def_Colour2
   m_Colour3 = m_def_Colour3
   m_Colour4 = m_def_Colour4
   m_LabelType = m_def_LabelType
   m_CaptionAlignment = m_def_CaptionAlignment
   m_BorderStyle = m_def_BorderStyle
   m_HighlightColour = m_def_HighlightColour
   m_HighlightDKColour = m_def_HighlightDKColour
   m_ShadowColour = m_def_ShadowColour
   m_ShadowDKColour = m_def_ShadowDKColour
   m_FlatBorderColour = m_def_FlatBorderColour
   m_TextShadowColour = m_def_TextShadowColour
   m_TextShadow = m_def_TextShadow
   m_TextShadowYOffset = m_def_TextShadowYOffset
   m_TextShadowXOffset = m_def_TextShadowXOffset
   m_WordWrap = m_def_WordWrap
   m_Offset = m_def_Offset
   m_NonTTError = m_def_NonTTError
   m_Version = m_def_Version
   m_GradientAngle = m_def_GradientAngle
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   m_GradientType = PropBag.ReadProperty("GradientType", m_def_GradientType)
   m_Caption = PropBag.ReadProperty("Caption", UserControl.Extender.Name)
   Set m_Fount = PropBag.ReadProperty("Fount", Ambient.Font)
   m_CaptionColour = PropBag.ReadProperty("CaptionColour", m_def_CaptionColour)
   m_Colour1 = PropBag.ReadProperty("Colour1", m_def_Colour1)
   m_Colour2 = PropBag.ReadProperty("Colour2", m_def_Colour2)
   m_Colour3 = PropBag.ReadProperty("Colour3", m_def_Colour3)
   m_Colour4 = PropBag.ReadProperty("Colour4", m_def_Colour4)
   m_LabelType = PropBag.ReadProperty("LabelType", m_def_LabelType)
   m_CaptionAlignment = PropBag.ReadProperty("CaptionAlignment", m_def_CaptionAlignment)
   m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
   m_HighlightColour = PropBag.ReadProperty("HighlightColour", m_def_HighlightColour)
   m_HighlightDKColour = PropBag.ReadProperty("HighlightDKColour", m_def_HighlightDKColour)
   m_ShadowColour = PropBag.ReadProperty("ShadowColour", m_def_ShadowColour)
   m_ShadowDKColour = PropBag.ReadProperty("ShadowDKColour", m_def_ShadowDKColour)
   m_FlatBorderColour = PropBag.ReadProperty("FlatBorderColour", m_def_FlatBorderColour)
   m_TextShadowColour = PropBag.ReadProperty("TextShadowColour", m_def_TextShadowColour)
   m_TextShadow = PropBag.ReadProperty("TextShadow", m_def_TextShadow)
   m_TextShadowYOffset = PropBag.ReadProperty("TextShadowYOffset", m_def_TextShadowYOffset)
   m_TextShadowXOffset = PropBag.ReadProperty("TextShadowXOffset", m_def_TextShadowXOffset)
   m_WordWrap = PropBag.ReadProperty("WordWrap", m_def_WordWrap)
   m_Offset = PropBag.ReadProperty("Offset", m_def_Offset)
   m_NonTTError = PropBag.ReadProperty("NonTTError", m_def_NonTTError)
   m_Version = PropBag.ReadProperty("Version", m_def_Version)
   m_GradientAngle = PropBag.ReadProperty("GradientAngle", m_def_GradientAngle)
End Sub

Private Sub UserControl_Resize()
   DrawLabel
End Sub

Private Sub UserControl_Show()
   DrawLabel
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   Call PropBag.WriteProperty("GradientType", m_GradientType, m_def_GradientType)
   Call PropBag.WriteProperty("Caption", m_Caption, UserControl.Extender.Name)
   Call PropBag.WriteProperty("Fount", m_Fount, Ambient.Font)
   Call PropBag.WriteProperty("CaptionColour", m_CaptionColour, m_def_CaptionColour)
   Call PropBag.WriteProperty("Colour1", m_Colour1, m_def_Colour1)
   Call PropBag.WriteProperty("Colour2", m_Colour2, m_def_Colour2)
   Call PropBag.WriteProperty("Colour3", m_Colour3, m_def_Colour3)
   Call PropBag.WriteProperty("Colour4", m_Colour4, m_def_Colour4)
   Call PropBag.WriteProperty("LabelType", m_LabelType, m_def_LabelType)
   Call PropBag.WriteProperty("CaptionAlignment", m_CaptionAlignment, m_def_CaptionAlignment)
   Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
   Call PropBag.WriteProperty("HighlightColour", m_HighlightColour, m_def_HighlightColour)
   Call PropBag.WriteProperty("HighlightDKColour", m_HighlightDKColour, m_def_HighlightDKColour)
   Call PropBag.WriteProperty("ShadowColour", m_ShadowColour, m_def_ShadowColour)
   Call PropBag.WriteProperty("ShadowDKColour", m_ShadowDKColour, m_def_ShadowDKColour)
   Call PropBag.WriteProperty("FlatBorderColour", m_FlatBorderColour, m_def_FlatBorderColour)
   Call PropBag.WriteProperty("TextShadowColour", m_TextShadowColour, m_def_TextShadowColour)
   Call PropBag.WriteProperty("TextShadow", m_TextShadow, m_def_TextShadow)
   Call PropBag.WriteProperty("TextShadowYOffset", m_TextShadowYOffset, m_def_TextShadowYOffset)
   Call PropBag.WriteProperty("TextShadowXOffset", m_TextShadowXOffset, m_def_TextShadowXOffset)
   Call PropBag.WriteProperty("WordWrap", m_WordWrap, m_def_WordWrap)
   Call PropBag.WriteProperty("Offset", m_Offset, m_def_Offset)
   Call PropBag.WriteProperty("NonTTError", m_NonTTError, m_def_NonTTError)
   Call PropBag.WriteProperty("Version", m_Version, m_def_Version)
   Call PropBag.WriteProperty("GradientAngle", m_GradientAngle, m_def_GradientAngle)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get CaptionColour() As OLE_COLOR
Attribute CaptionColour.VB_Description = "The colour of the Caption text."
Attribute CaptionColour.VB_ProcData.VB_Invoke_Property = ";Colours"
   CaptionColour = m_CaptionColour
End Property

Public Property Let CaptionColour(ByVal New_CaptionColour As OLE_COLOR)
   m_CaptionColour = New_CaptionColour
   PropertyChanged "CaptionColour"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get Colour1() As OLE_COLOR
Attribute Colour1.VB_Description = "The first gradient colour."
Attribute Colour1.VB_ProcData.VB_Invoke_Property = ";Colours"
   Colour1 = m_Colour1
End Property

Public Property Let Colour1(ByVal New_Colour1 As OLE_COLOR)
   m_Colour1 = New_Colour1
   PropertyChanged "Colour1"
   UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get Colour2() As OLE_COLOR
Attribute Colour2.VB_Description = "The second gradient colour."
Attribute Colour2.VB_ProcData.VB_Invoke_Property = ";Colours"
   Colour2 = m_Colour2
End Property

Public Property Let Colour2(ByVal New_Colour2 As OLE_COLOR)
   m_Colour2 = New_Colour2
   PropertyChanged "Colour2"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get Colour3() As OLE_COLOR
Attribute Colour3.VB_Description = "The third gradient colour."
Attribute Colour3.VB_ProcData.VB_Invoke_Property = ";Colours"
   Colour3 = m_Colour3
End Property

Public Property Let Colour3(ByVal New_Colour3 As OLE_COLOR)
   m_Colour3 = New_Colour3
   PropertyChanged "Colour3"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get Colour4() As OLE_COLOR
Attribute Colour4.VB_Description = "The fourth gradient colour."
Attribute Colour4.VB_ProcData.VB_Invoke_Property = ";Colours"
   Colour4 = m_Colour4
End Property

Public Property Let Colour4(ByVal New_Colour4 As OLE_COLOR)
   m_Colour4 = New_Colour4
   PropertyChanged "Colour4"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=21,0,0,0
Public Property Get LabelType() As bsLabelType
Attribute LabelType.VB_Description = "The alignment of the Caption."
Attribute LabelType.VB_ProcData.VB_Invoke_Property = ";Appearance"
   LabelType = m_LabelType
End Property

Public Property Let LabelType(ByVal New_LabelType As bsLabelType)
   m_LabelType = New_LabelType
   
   If (m_LabelType = [Vertical 90°] Or m_LabelType = [Vertical 270°]) And IsFountTrueType(m_Fount.Name) = False Then
      If m_NonTTError Then
         MsgBox "The LabelType property can only be Vertical when the Fount is a TrueType fount.", vbExclamation
      End If
      m_LabelType = [Horizontal]
   End If
   
   PropertyChanged "LabelType"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=1,0,0,0
Public Property Get CaptionAlignment() As bsCaptionAlign
Attribute CaptionAlignment.VB_ProcData.VB_Invoke_Property = ";Appearance"
   CaptionAlignment = m_CaptionAlignment
End Property

Public Property Let CaptionAlignment(ByVal New_CaptionAlignment As bsCaptionAlign)
   m_CaptionAlignment = New_CaptionAlignment
   PropertyChanged "CaptionAlignment"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get BorderStyle() As bsBorderStyle
   BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As bsBorderStyle)
   m_BorderStyle = New_BorderStyle
   PropertyChanged "BorderStyle"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get HighlightColour() As OLE_COLOR
   HighlightColour = m_HighlightColour
End Property

Public Property Let HighlightColour(ByVal New_HighlightColour As OLE_COLOR)
   m_HighlightColour = New_HighlightColour
   PropertyChanged "HighlightColour"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get HighlightDKColour() As OLE_COLOR
   HighlightDKColour = m_HighlightDKColour
End Property

Public Property Let HighlightDKColour(ByVal New_HighlightDKColour As OLE_COLOR)
   m_HighlightDKColour = New_HighlightDKColour
   PropertyChanged "HighlightDKColour"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ShadowColour() As OLE_COLOR
   ShadowColour = m_ShadowColour
End Property

Public Property Let ShadowColour(ByVal New_ShadowColour As OLE_COLOR)
   m_ShadowColour = New_ShadowColour
   PropertyChanged "ShadowColour"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ShadowDKColour() As OLE_COLOR
   ShadowDKColour = m_ShadowDKColour
End Property

Public Property Let ShadowDKColour(ByVal New_ShadowDKColour As OLE_COLOR)
   m_ShadowDKColour = New_ShadowDKColour
   PropertyChanged "ShadowDKColour"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get FlatBorderColour() As OLE_COLOR
   FlatBorderColour = m_FlatBorderColour
End Property

Public Property Let FlatBorderColour(ByVal New_FlatBorderColour As OLE_COLOR)
   m_FlatBorderColour = New_FlatBorderColour
   PropertyChanged "FlatBorderColour"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get TextShadowColour() As OLE_COLOR
Attribute TextShadowColour.VB_Description = "The colour of the shadow under the text when TextShadow is set to True."
   TextShadowColour = m_TextShadowColour
End Property

Public Property Let TextShadowColour(ByVal New_TextShadowColour As OLE_COLOR)
   m_TextShadowColour = New_TextShadowColour
   PropertyChanged "TextShadowColour"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get TextShadow() As Boolean
Attribute TextShadow.VB_Description = "Determines whether or not a shadow is drawn under the caption."
   TextShadow = m_TextShadow
End Property

Public Property Let TextShadow(ByVal New_TextShadow As Boolean)
   m_TextShadow = New_TextShadow
   PropertyChanged "TextShadow"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,2
Public Property Get TextShadowYOffset() As Integer
Attribute TextShadowYOffset.VB_Description = "The distance between the text shadow and the Caption vertically."
   TextShadowYOffset = m_TextShadowYOffset
End Property

Public Property Let TextShadowYOffset(ByVal New_TextShadowYOffset As Integer)
   m_TextShadowYOffset = New_TextShadowYOffset
   PropertyChanged "TextShadowYOffset"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,2
Public Property Get TextShadowXOffset() As Integer
Attribute TextShadowXOffset.VB_Description = "The distance between the text shadow and the Caption horizontally."
   TextShadowXOffset = m_TextShadowXOffset
End Property

Public Property Let TextShadowXOffset(ByVal New_TextShadowXOffset As Integer)
   m_TextShadowXOffset = New_TextShadowXOffset
   PropertyChanged "TextShadowXOffset"
   DrawLabel
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get WordWrap() As Boolean
Attribute WordWrap.VB_Description = "Enables and disabled multiple label lines."
   WordWrap = m_WordWrap
End Property

Public Property Let WordWrap(ByVal New_WordWrap As Boolean)
   m_WordWrap = New_WordWrap
   PropertyChanged "WordWrap"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,6
Public Property Get Offset() As Integer
Attribute Offset.VB_Description = "The text offset from the left."
   Offset = m_Offset
End Property

Public Property Let Offset(ByVal New_Offset As Integer)
   m_Offset = New_Offset
   PropertyChanged "Offset"
   DrawLabel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get NonTTError() As Boolean
Attribute NonTTError.VB_Description = "Decides whether or not to warn the user that a non-TrueType font cannot be rotated."
   NonTTError = m_NonTTError
End Property

Public Property Let NonTTError(ByVal New_NonTTError As Boolean)
   m_NonTTError = New_NonTTError
   PropertyChanged "NonTTError"
End Property

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_DblClick()
   RaiseEvent DblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,1,1,
Public Property Get Version() As String
Attribute Version.VB_Description = "Returns the control version."
   Version = App.Major & "." & App.Minor & " build " & App.Revision
End Property

Public Property Let Version(ByVal New_Version As String)
   If Ambient.UserMode = False Then Err.Raise 387
   If Ambient.UserMode Then Err.Raise 382
   m_Version = New_Version
   PropertyChanged "Version"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get GradientAngle() As Single
Attribute GradientAngle.VB_Description = "Controls the angle of the gradient when GradientType is set to 2 Way."
   GradientAngle = m_GradientAngle
End Property

Public Property Let GradientAngle(ByVal New_GradientAngle As Single)
   m_GradientAngle = New_GradientAngle
   PropertyChanged "GradientAngle"
   UserControl_Resize
End Property


Private Sub TextDraw_NonWrap()
End Sub
Public Sub TextDraw_VWrap(ByVal sArray, iOffset As Integer)
   
   Dim px As Integer
   Dim LineHeight As Integer, LineCount As Integer
   Dim I As Integer
   Dim iStartingOffset As Integer
   
   iStartingOffset = CurrentY
   px = CurrentX
   LineHeight = TextHeight(" ")
   
   ' number of lines that can fit into the UserControl area
   LineCount = ScaleWidth / LineHeight
   LineCount = IIf(ScaleWidth / LineHeight > UBound(sArray), UBound(sArray), _
      ScaleWidth / LineHeight)
         
   For I = 1 To LineCount
      CurrentX = px + iOffset + m_Offset
      Print sArray(I)
      px = px + LineHeight
      
      CurrentY = iStartingOffset
         
   Next
   
End Sub
