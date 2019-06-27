VERSION 5.00
Begin VB.UserControl StylerButton 
   ClientHeight    =   2010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   ScaleHeight     =   134
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   200
   Begin VB.PictureBox imgDis 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   960
      ScaleHeight     =   1
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.PictureBox imgIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   120
      ScaleHeight     =   1
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Timer tmrCheck 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   855
      Top             =   480
   End
End
Attribute VB_Name = "StylerButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'EVENTS.
Public Event Click()
Public Event DoubleClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseEnters(ByVal x As Long, ByVal y As Long)
Public Event MouseLeaves(ByVal x As Long, ByVal y As Long)

'===========================================
'===========================================
'===========================================

'CAPTION.
Private strCaption As String 'Caption text.
Private CapDis As OLE_COLOR  'Caption Disabled colour.
Private oleForeColor As OLE_COLOR 'Caption text color.
Private udtCaptionAlign As CaptionAlignmentS  'Caption Alignment.
Private fntFont As Font 'Caption font.
Private CEC As OLE_COLOR 'Caption Effect Colour.
Private CTE As CaptionTextEffects 'Caption Effect.
Private COX As Integer 'Caption Offset X.
Private COY As Integer 'Caption Offset Y.
Private SOX As Integer 'Caption Shadow Offset X.
Private SOY As Integer 'Caption Shadow Offset Y.
'===========================================
'===========================================
'===========================================



'MOUSE DIRECTION.
Private udtPoint As POINTAPI 'Current mouse position (for checking if mouse is over button).

'===========================================
'===========================================
'===========================================

'CHECK PROPERTY.
Private bolMouseDown As Boolean 'Mouse currently down?
Private bolMouseOver As Boolean 'Mouse currently over button?
Private bolHasFocus As Boolean 'Currently has focus?
Private bolEnabled As Boolean 'Enabled?

'===========================================
'===========================================
'===========================================

'FOCUS DOT RECT.
Private bolFocusDottedRect As Boolean 'Draw focus dotted rect?
'===========================================
'===========================================
'===========================================

'ROUNDED CORNER.
Private lonRoundValue As Long 'Rounded corners value.

'===========================================
'===========================================
'===========================================

'PRIVATE/PUBLIC TYPES.
Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type TRIVERTEX
   x As Long
   y As Long
   Red As Integer
   Green As Integer
   Blue As Integer
   Alpha As Integer
End Type

Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Type cRGB
    Blue As Byte
    Green As Byte
    Red As Byte
End Type

'===========================================
'===========================================
'===========================================

'PUBLIC/PRIVATE ENUM.
Public Enum CaptionAlignmentS
    [Left Top] = 1
    [Left Middle] = 2
    [Left Bottom] = 3
    [Center Top] = 4
    [Center Middle] = 5
    [Center Bottom] = 6
    [Right Top] = 7
    [Right Middle] = 8
    [Right Bottom] = 9
End Enum

Public Enum CaptionTextEffects
    [Normal] = 1
    [Embossed] = 2
    [Engraved] = 3
    [outline] = 4
    [Shadow] = 5
End Enum



Private Enum GRADIENT_DIRECT
    [Left to Right] = &H0
    [Top to Bottom] = &H1
End Enum

'===========================================
'===========================================
'===========================================

'FUNCTION DECLARE.
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal x3 As Long, ByVal y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal x3 As Long, ByVal y3 As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dY As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

'===========================================
'===========================================
'===========================================

'PRIVATE CONSTANT.
Private udtRect As RECT
Private Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Private Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0



'Print aligned text to the button (caption).
Private Sub PrintText(ByVal TextString As String, ByVal Alignment As CaptionAlignmentS)
Dim lonStartWidth As Long, lonStartHeight As Long

If Alignment = 1 Then
    lonStartWidth = 1 + CByte(COX)
    lonStartHeight = 0 + CByte(COY)
ElseIf Alignment = 2 Then
    lonStartWidth = 1 + CByte(COX)
    lonStartHeight = (UserControl.ScaleHeight / 2) - (UserControl.TextHeight(TextString) / 2) - 1 + CByte(COY)
ElseIf Alignment = 3 Then
    lonStartWidth = 1 + CByte(COX)
    lonStartHeight = (UserControl.ScaleHeight - UserControl.TextHeight(TextString)) - 1 + CByte(COY)


ElseIf Alignment = 4 Then
    lonStartWidth = (UserControl.ScaleWidth / 2) - (UserControl.TextWidth(TextString) / 2) - 1 + CByte(COX)
    lonStartHeight = 0 + CByte(COY)
ElseIf Alignment = 5 Then
    lonStartWidth = (UserControl.ScaleWidth / 2) - (UserControl.TextWidth(TextString) / 2) - 1 + CByte(COX)
    lonStartHeight = (UserControl.ScaleHeight / 2) - (UserControl.TextHeight(TextString) / 2) - 1 + CByte(COY)
ElseIf Alignment = 6 Then
    lonStartWidth = (UserControl.ScaleWidth / 2) - (UserControl.TextWidth(TextString) / 2) - 1 + CByte(COX)
    lonStartHeight = (UserControl.ScaleHeight - UserControl.TextHeight(TextString)) - 1 + CByte(COY)


ElseIf Alignment = 7 Then
    lonStartWidth = (UserControl.ScaleWidth - UserControl.TextWidth(TextString)) - 3 + CByte(COX)
    lonStartHeight = 0 + CByte(COY)
ElseIf Alignment = 8 Then
    lonStartWidth = (UserControl.ScaleWidth - UserControl.TextWidth(TextString)) - 3 + CByte(COX)
    lonStartHeight = (UserControl.ScaleHeight / 2) - (UserControl.TextHeight(TextString) / 2) - 1 + CByte(COY)
ElseIf Alignment = 9 Then
    lonStartWidth = (UserControl.ScaleWidth - UserControl.TextWidth(TextString)) - 3 + CByte(COX)
    lonStartHeight = (UserControl.ScaleHeight - UserControl.TextHeight(TextString)) - 1 + CByte(COY)
End If


If bolEnabled = False Then
    UserControl.CurrentX = lonStartWidth
    UserControl.CurrentY = lonStartHeight
    UserControl.Print TextString

Else
    If CTE = Normal Then
        UserControl.CurrentX = lonStartWidth
        UserControl.CurrentY = lonStartHeight
        UserControl.Print TextString
    ElseIf CTE = Engraved Then
        UserControl.ForeColor = CEC
        UserControl.CurrentX = lonStartWidth + 1
        UserControl.CurrentY = lonStartHeight + 1
        UserControl.Print TextString
        UserControl.ForeColor = RGB(128, 128, 128)
        UserControl.CurrentX = lonStartWidth - 1
        UserControl.CurrentY = lonStartHeight
        UserControl.Print TextString
        UserControl.ForeColor = oleForeColor
        UserControl.CurrentX = lonStartWidth
        UserControl.CurrentY = lonStartHeight
        UserControl.Print TextString
        
    ElseIf CTE = Embossed Then
        UserControl.ForeColor = CEC
        UserControl.CurrentX = lonStartWidth - 1
        UserControl.CurrentY = lonStartHeight - 1
        UserControl.Print TextString
        UserControl.ForeColor = RGB(128, 128, 128)
        UserControl.CurrentX = lonStartWidth + 1
        UserControl.CurrentY = lonStartHeight + 1
        UserControl.Print TextString
        UserControl.ForeColor = oleForeColor
        UserControl.CurrentX = lonStartWidth
        UserControl.CurrentY = lonStartHeight
        UserControl.Print TextString
    ElseIf CTE = outline Then
        UserControl.ForeColor = CEC
        UserControl.CurrentX = lonStartWidth + 1
        UserControl.CurrentY = lonStartHeight
        UserControl.Print TextString
        UserControl.CurrentX = lonStartWidth - 1
        UserControl.CurrentY = lonStartHeight
        UserControl.Print TextString
        UserControl.CurrentY = lonStartHeight - 1
        UserControl.CurrentX = lonStartWidth
        UserControl.Print TextString
        UserControl.CurrentY = lonStartHeight + 1
        UserControl.CurrentX = lonStartWidth
        UserControl.Print TextString
        UserControl.ForeColor = oleForeColor
        UserControl.CurrentX = lonStartWidth
        UserControl.CurrentY = lonStartHeight
        UserControl.Print TextString
        
    ElseIf CTE = Shadow Then
        UserControl.ForeColor = CEC
        UserControl.CurrentX = lonStartWidth + CByte(SOX)
        UserControl.CurrentY = lonStartHeight + CByte(SOY)
        UserControl.Print TextString
        UserControl.ForeColor = oleForeColor
        UserControl.CurrentX = lonStartWidth
        UserControl.CurrentY = lonStartHeight
        UserControl.Print TextString
    End If
End If

    'UserControl.CurrentX = lonStartWidth
    'UserControl.CurrentY = lonStartHeight
    'UserControl.Print TextString





End Sub


'Draw the dotted focus rect on the button.
Private Sub DrawDottedFocusRect()
Dim lonLoop As Long

    'Draw the top focus dotted line.
For lonLoop = 3 To (UserControl.ScaleWidth - 5) Step 2
    UserControl.PSet (lonLoop, 2), 0
Next lonLoop

    'Draw the left focus dotted line.
For lonLoop = 4 To (UserControl.ScaleHeight - 4) Step 2
    UserControl.PSet (2, lonLoop), 0
Next lonLoop

    'Draw the bottom focus dotted line.
For lonLoop = 3 To (UserControl.ScaleWidth - 5) Step 2
    UserControl.PSet (lonLoop, ScaleHeight - 4), 0
Next lonLoop

    'Draw the right focus dotted line.
For lonLoop = 4 To (UserControl.ScaleHeight - 4) Step 2
    UserControl.PSet (ScaleWidth - 4, lonLoop), 0
Next lonLoop


End Sub

'Draw the control.
Private Sub PaintControl()
On Error Resume Next

Dim lonRect As Long
Dim strName As String

'Shape control.
If lonRoundValue <= 0 Then
    lonRoundValue = 1
End If
lonRect = CreateRoundRectRgn(0, 0, ScaleWidth, ScaleHeight, lonRoundValue - 1, lonRoundValue - 1)
SetWindowRgn UserControl.hWnd, lonRect, True

strName = fntFont.Name

If Err = 0 Then
    Set UserControl.Font = fntFont
End If
    CapDis = RGB(183, 182, 186)
    If bolEnabled = False Then
    DefineRect 0, 0, ScaleWidth, ScaleHeight
    DrawGradientTwoColour UserControl.hdc, [Top to Bottom], RGB(238, 238, 238), RGB(238, 238, 238)
    UserControl.ForeColor = RGB(178, 178, 178)
    RoundRect UserControl.hdc, 0, 0, ScaleWidth - 1, ScaleHeight - 1, lonRoundValue, lonRoundValue
    UserControl.ForeColor = RGB(250, 250, 250)
    RoundRect UserControl.hdc, 1, 1, ScaleWidth - 2, ScaleHeight - 2, lonRoundValue, lonRoundValue

        
        
        
        UserControl.ForeColor = CapDis
        PrintText strCaption, udtCaptionAlign
        
        Exit Sub
    End If
    
    DefineRect 0, 0, ScaleWidth, ScaleHeight / 2 - 1
    DrawGradientTwoColour UserControl.hdc, [Top to Bottom], RGB(250, 250, 250), RGB(224, 225, 225)
    DefineRect 0, ScaleHeight / 2 - 1, ScaleWidth, ScaleHeight
    DrawGradientTwoColour UserControl.hdc, [Top to Bottom], RGB(184, 191, 200), RGB(231, 234, 236)
    DefineRect 0, 0, 2, ScaleHeight / 2 - 1
    DrawGradientTwoColour UserControl.hdc, [Top to Bottom], RGB(247, 249, 250), RGB(236, 238, 240)
    DefineRect 0, ScaleHeight / 2 - 1, 2, ScaleHeight
    DrawGradientTwoColour UserControl.hdc, [Top to Bottom], RGB(220, 226, 232), RGB(237, 240, 242)
    DefineRect ScaleWidth - 3, 0, ScaleWidth, ScaleHeight / 2 - 1
    DrawGradientTwoColour UserControl.hdc, [Top to Bottom], RGB(247, 249, 250), RGB(236, 238, 240)
    DefineRect ScaleWidth - 3, ScaleHeight / 2 - 1, ScaleWidth, ScaleHeight
    DrawGradientTwoColour UserControl.hdc, [Top to Bottom], RGB(220, 226, 232), RGB(237, 240, 242)
    UserControl.ForeColor = RGB(75, 86, 93)
    RoundRect UserControl.hdc, 0, 0, ScaleWidth - 1, ScaleHeight - 1, lonRoundValue, lonRoundValue


    If bolMouseOver = True And bolMouseDown = False Then
        DefineRect 0, 0, ScaleWidth, ScaleHeight / 2 - 1
        DrawGradientTwoColour UserControl.hdc, [Top to Bottom], RGB(252, 255, 255), RGB(207, 247, 255)
        DefineRect 0, ScaleHeight / 2 - 1, ScaleWidth - 1, ScaleHeight
        DrawGradientTwoColour UserControl.hdc, [Top to Bottom], RGB(138, 215, 255), RGB(220, 255, 255)
        DefineRect 0, 0, 2, ScaleHeight / 2 - 1
        DrawGradientTwoColour UserControl.hdc, [Top to Bottom], RGB(255, 255, 255), RGB(239, 255, 255)
        DefineRect 0, ScaleHeight / 2 - 1, 2, ScaleHeight
        DrawGradientTwoColour UserControl.hdc, [Top to Bottom], RGB(216, 251, 255), RGB(235, 255, 255)
        DefineRect ScaleWidth - 3, 0, ScaleWidth, ScaleHeight / 2 - 1
        DrawGradientTwoColour UserControl.hdc, [Top to Bottom], RGB(255, 255, 255), RGB(239, 255, 255)
        DefineRect ScaleWidth - 3, ScaleHeight / 2 - 1, ScaleWidth, ScaleHeight
        DrawGradientTwoColour UserControl.hdc, [Top to Bottom], RGB(216, 251, 255), RGB(235, 255, 255)
        UserControl.Line (0, ScaleHeight - 3)-(ScaleWidth, ScaleHeight - 3), RGB(191, 255, 255)
        UserControl.ForeColor = RGB(31, 87, 168)
        RoundRect UserControl.hdc, 0, 0, ScaleWidth - 1, ScaleHeight - 1, lonRoundValue, lonRoundValue


        GoTo Vista2Done
    End If
    
    If bolHasFocus = True And bolMouseDown = False Then
        DefineRect 0, 0, ScaleWidth, ScaleHeight / 2 - 1
        DrawGradientTwoColour UserControl.hdc, [Top to Bottom], RGB(254, 255, 255), RGB(207, 231, 245)
        DefineRect 0, ScaleHeight / 2 - 1, ScaleWidth - 1, ScaleHeight
        DrawGradientTwoColour UserControl.hdc, [Top to Bottom], RGB(144, 197, 241), RGB(206, 244, 246)
        DefineRect 0, 0, 2, ScaleHeight / 2 - 1
        DrawGradientTwoColour UserControl.hdc, [Top to Bottom], RGB(239, 247, 252), RGB(226, 240, 248)
        DefineRect 0, ScaleHeight / 2 - 1, 2, ScaleHeight
        DrawGradientTwoColour UserControl.hdc, [Top to Bottom], RGB(206, 230, 246), RGB(221, 242, 248)
        DefineRect ScaleWidth - 3, 0, ScaleWidth, ScaleHeight / 2 - 1
        DrawGradientTwoColour UserControl.hdc, [Top to Bottom], RGB(247, 249, 250), RGB(236, 238, 240)
        DefineRect ScaleWidth - 3, ScaleHeight / 2 - 1, ScaleWidth, ScaleHeight
        DrawGradientTwoColour UserControl.hdc, [Top to Bottom], RGB(220, 226, 232), RGB(237, 240, 242)
        UserControl.Line (0, ScaleHeight - 3)-(ScaleWidth, ScaleHeight - 3), RGB(185, 253, 247)
        UserControl.ForeColor = RGB(31, 87, 168)
        RoundRect UserControl.hdc, 0, 0, ScaleWidth - 1, ScaleHeight - 1, lonRoundValue, lonRoundValue
        If bolFocusDottedRect = True Then
            'Draw dotted focus rect.
            DrawDottedFocusRect
        End If
        GoTo Vista2Done
    
    End If
    
    If bolMouseDown = True Then
        DrawGradientFourColour UserControl.hdc, 1, 1, ScaleWidth / 4, ScaleHeight / 2 - 1, RGB(222, 245, 255), RGB(243, 255, 255), RGB(138, 204, 237), RGB(155, 226, 255)
        DrawGradientFourColour UserControl.hdc, ScaleWidth / 4, 1, ScaleWidth / 2, ScaleHeight / 2 - 1, RGB(243, 255, 255), RGB(243, 255, 255), RGB(155, 226, 255), RGB(155, 226, 255)
        DrawGradientFourColour UserControl.hdc, ScaleWidth - (ScaleWidth / 4) - 1, 1, ScaleWidth / 4, ScaleHeight / 2 - 1, RGB(243, 255, 255), RGB(222, 245, 255), RGB(155, 226, 255), RGB(138, 204, 237)
        DrawGradientFourColour UserControl.hdc, 1, ScaleHeight / 2 - 1, ScaleWidth / 4, ScaleHeight / 2 - 2, RGB(79, 161, 210), RGB(93, 181, 226), RGB(160, 214, 245), RGB(177, 233, 255)
        DrawGradientFourColour UserControl.hdc, ScaleWidth / 4, ScaleHeight / 2 - 1, ScaleWidth / 2, ScaleHeight / 2 - 2, RGB(93, 181, 226), RGB(93, 181, 226), RGB(177, 233, 255), RGB(177, 233, 255)
        DrawGradientFourColour UserControl.hdc, ScaleWidth - (ScaleWidth / 4) - 1, ScaleHeight / 2 - 1, ScaleWidth / 4, ScaleHeight / 2 - 2, RGB(93, 181, 226), RGB(79, 161, 210), RGB(177, 233, 255), RGB(160, 214, 245)
        DefineRect 0, 0, 2, ScaleHeight / 2 - 1
        DrawGradientTwoColour UserControl.hdc, [Top to Bottom], RGB(199, 223, 242), RGB(127, 190, 226)
        DefineRect 0, ScaleHeight / 2 - 1, 2, ScaleHeight
        DrawGradientTwoColour UserControl.hdc, [Top to Bottom], RGB(67, 148, 200), RGB(138, 193, 227)
        DefineRect ScaleWidth - 3, 0, ScaleWidth, ScaleHeight / 2 - 1
        DrawGradientTwoColour UserControl.hdc, [Top to Bottom], RGB(199, 223, 242), RGB(127, 190, 226)
        DefineRect ScaleWidth - 3, ScaleHeight / 2 - 1, ScaleWidth, ScaleHeight
        DrawGradientTwoColour UserControl.hdc, [Top to Bottom], RGB(67, 148, 200), RGB(138, 193, 227)
        UserControl.Line (0, 1)-(ScaleWidth, 1), RGB(145, 192, 223)
        UserControl.Line (0, ScaleHeight - 3)-(ScaleWidth, ScaleHeight - 3), RGB(145, 192, 223)
        UserControl.ForeColor = RGB(31, 87, 168)
        RoundRect UserControl.hdc, 0, 0, ScaleWidth - 1, ScaleHeight - 1, lonRoundValue, lonRoundValue



        If bolHasFocus = True And bolFocusDottedRect = True Then
            DrawDottedFocusRect
        End If
        
        GoTo Vista2Done
    End If
Vista2Done:
    UserControl.ForeColor = oleForeColor
    PrintText strCaption, udtCaptionAlign
    
    Exit Sub
If bolMouseOver = True And bolFocusDottedRect = True Then DrawDottedFocusRect
End Sub

Public Property Get CaptionAlignment() As CaptionAlignmentS
Attribute CaptionAlignment.VB_Description = "Button Caption Alignment."
CaptionAlignment = udtCaptionAlign
End Property

Public Property Let CaptionAlignment(ByVal NewValue As CaptionAlignmentS)
COX = 0
COY = 0
udtCaptionAlign = NewValue
PropertyChanged "CaptionAlignment"
PaintControl
End Property


Public Property Get Caption() As String
Attribute Caption.VB_Description = "Button Caption."
Caption = strCaption
End Property

Public Property Let Caption(ByVal NewValue As String)
strCaption = NewValue
PropertyChanged "Caption"
PaintControl
End Property
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Button Caption ForeColour."
ForeColor = oleForeColor
End Property

Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
oleForeColor = NewValue
PropertyChanged "ForeColor"
PaintControl
End Property


Public Property Get CaptionDisableColor() As OLE_COLOR
Attribute CaptionDisableColor.VB_Description = "Button Disabled Caption colour."
CaptionDisableColor = CapDis
End Property

Public Property Let CaptionDisableColor(ByVal NewValue As OLE_COLOR)
CapDis = NewValue
PropertyChanged "CaptionDisableColor"
PaintControl
End Property

Public Property Get CaptionEffectColor() As OLE_COLOR
Attribute CaptionEffectColor.VB_Description = "If You Apply Caption Effect(Shadow,Engraved,etc).This Colour used in Effect."
CaptionEffectColor = CEC
End Property

Public Property Let CaptionEffectColor(ByVal NewValue As OLE_COLOR)
CEC = NewValue
PropertyChanged "CaptionEffectColor"
PaintControl
End Property

Public Property Get CaptionEffect() As CaptionTextEffects
Attribute CaptionEffect.VB_Description = "Caption Text Effects (Shadow,Engraved,Outline,Embossed.etc)."
CaptionEffect = CTE
End Property

Public Property Let CaptionEffect(ByVal NewValue As CaptionTextEffects)
CTE = NewValue
PropertyChanged "CaptionEffect"
PaintControl
End Property

Public Property Get CaptionOffsetX() As Integer
Attribute CaptionOffsetX.VB_Description = "Caption Offset X."
CaptionOffsetX = COX
End Property

Public Property Let CaptionOffsetX(ByVal NewValue As Integer)
COX = NewValue
PropertyChanged "CaptionOffsetX"
PaintControl
End Property
Public Property Get CaptionOffsetY() As Integer
Attribute CaptionOffsetY.VB_Description = "Caption Offset Y."
CaptionOffsetY = COY
End Property

Public Property Let CaptionOffsetY(ByVal NewValue As Integer)
COY = NewValue
PropertyChanged "CaptionOffsetY"
PaintControl
End Property








Public Property Get ShadowOffsetX() As Integer
Attribute ShadowOffsetX.VB_Description = "If You apply Shadow in Captioneffects.Shadow Offset X."
ShadowOffsetX = SOX
End Property

Public Property Let ShadowOffsetX(ByVal NewValue As Integer)
SOX = NewValue
PropertyChanged "ShadowOffsetX"
PaintControl
End Property
Public Property Get ShadowOffsetY() As Integer
Attribute ShadowOffsetY.VB_Description = "If You apply Shadow in Captioneffects.Shadow Offset Y."
ShadowOffsetY = SOY
End Property

Public Property Let ShadowOffsetY(ByVal NewValue As Integer)
SOY = NewValue
PropertyChanged "ShadowOffsetY"
PaintControl
End Property
Public Property Get FocusDottedRect() As Boolean
Attribute FocusDottedRect.VB_Description = "Create Button Focus Dotted Rect."
FocusDottedRect = bolFocusDottedRect
End Property

Public Property Let FocusDottedRect(ByVal NewValue As Boolean)
bolFocusDottedRect = NewValue
PropertyChanged "FocusDottedRect"
PaintControl
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Button Enabled/Disable."
Enabled = bolEnabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
bolEnabled = NewValue
PropertyChanged "Enabled"
PaintControl
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Button Caption Font."
Set Font = fntFont
End Property

Public Property Set Font(ByVal NewValue As Font)
Set fntFont = NewValue
Set UserControl.Font = NewValue
PropertyChanged "Font"
PaintControl
End Property
Public Property Get RoundedValue() As Long
Attribute RoundedValue.VB_Description = "Button Border Rounded Value."
RoundedValue = lonRoundValue
End Property

Public Property Let RoundedValue(ByVal NewValue As Long)
lonRoundValue = NewValue
PropertyChanged "RoundedValue"
PaintControl
End Property

Private Sub tmrCheck_Timer()
If bolEnabled = False Then Exit Sub

Dim lonPosRet As Long, lonCurHWND As Long

tmrCheck.Enabled = False

lonPosRet = GetCursorPos(udtPoint)
lonCurHWND = WindowFromPoint(udtPoint.x, udtPoint.y)

If bolMouseOver = False Then
    
    If lonCurHWND = UserControl.hWnd Then
        bolMouseOver = True
        PaintControl
        RaiseEvent MouseEnters(udtPoint.x, udtPoint.y)
    End If

Else
    
    If lonCurHWND <> UserControl.hWnd Then
        bolMouseOver = False
        PaintControl
        RaiseEvent MouseLeaves(udtPoint.x, udtPoint.y)
    End If

End If

tmrCheck.Enabled = True
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    MsgBox KeyAscii
End Sub

Private Sub UserControl_Click()
    If bolEnabled = True Then RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
If bolEnabled = True Then RaiseEvent DoubleClick
End Sub

Private Sub UserControl_GotFocus()
If bolEnabled = True Then
    bolHasFocus = True
    PaintControl
End If
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
If bolEnabled = True Then
    RaiseEvent KeyDown(KeyCode, Shift)
    
    If KeyCode = 32 Then
        bolMouseDown = True
        PaintControl
    End If

End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
If bolEnabled = True Then
    RaiseEvent KeyPress(KeyAscii)
    
    If KeyAscii = 13 Then
        RaiseEvent Click
    End If

End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
If bolEnabled = True Then
    RaiseEvent KeyUp(KeyCode, Shift)
    
    If KeyCode = 32 Then
        bolMouseDown = False
        PaintControl
    End If

End If
End Sub

Private Sub UserControl_LostFocus()
If bolEnabled = True Then
    bolHasFocus = False
    PaintControl
End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If bolEnabled = True Then
    RaiseEvent MouseDown(Button, Shift, x, y)
    
    If Button = 1 Then
        bolMouseDown = True
        PaintControl
    End If

End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If bolEnabled = True Then
    RaiseEvent MouseMove(Button, Shift, x, y)
End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If bolEnabled = True Then
    RaiseEvent MouseUp(Button, Shift, x, y)
    If Button = 1 Then
        bolMouseDown = False
        PaintControl
    End If
End If
End Sub

Private Sub UserControl_Paint()
UserControl.Cls
PaintControl
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    Let Caption = .ReadProperty("Caption", "")
    Let CaptionDisableColor = .ReadProperty("CaptionDisableColor", RGB(212, 212, 212))
    Let CaptionEffectColor = .ReadProperty("CaptionEffectColor", RGB(200, 200, 200))
    Let CaptionEffect = .ReadProperty("CaptionEffect", 1)
    Let CaptionOffsetY = .ReadProperty("CaptionOffsetY", 0)
    Let CaptionOffsetX = .ReadProperty("CaptionOffsetX", 0)
    Let ShadowOffsetY = .ReadProperty("ShadowOffsetY", 2)
    Let ShadowOffsetX = .ReadProperty("ShadowOffsetX", 2)
    Let ForeColor = .ReadProperty("ForeColor", 0)
    Let FocusDottedRect = .ReadProperty("FocusDottedRect", False)
    Let Enabled = .ReadProperty("Enabled", True)
    Set Font = .ReadProperty("Font", Ambient.Font)
    Let RoundedValue = .ReadProperty("RoundedValue", 5)
    Let CaptionAlignment = .ReadProperty("CaptionAlignment", 5)
End With
tmrCheck.Enabled = Ambient.UserMode
End Sub



Private Sub UserControl_Terminate()
tmrCheck.Enabled = False
bolMouseDown = False
bolMouseOver = False
bolHasFocus = False
UserControl.Cls
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "Caption", strCaption, ""
    .WriteProperty "ForeColor", oleForeColor, 0
    .WriteProperty "CaptionDisableColor", CapDis, RGB(212, 212, 212)
    .WriteProperty "CaptionEffectColor", CEC, RGB(200, 200, 200)
    .WriteProperty "CaptionEffect", CTE, 1
    .WriteProperty "CaptionOffsetX", COX, 0
    .WriteProperty "CaptionOffsetY", COY, 0
    .WriteProperty "ShadowOffsetX", SOX, 2
    .WriteProperty "ShadowOffsetY", SOY, 2
    .WriteProperty "FocusDottedRect", bolFocusDottedRect, True
    .WriteProperty "Enabled", bolEnabled, True
    .WriteProperty "Font", fntFont, Ambient.Font
    .WriteProperty "RoundedValue", lonRoundValue, 5
    .WriteProperty "CaptionAlignment", udtCaptionAlign, 5
End With
End Sub

Private Sub UserControl_InitProperties()
Let Caption = Ambient.DisplayName
Let ForeColor = 0
Let FocusDottedRect = True
Let Enabled = True
Set Font = Ambient.Font
Let RoundedValue = 5
Let CaptionAlignment = 5
Let CaptionOffsetX = 0
Let CaptionOffsetY = 0
Let ShadowOffsetX = 2
Let ShadowOffsetY = 2
Let CaptionEffectColor = vbWhite
Let CaptionEffect = 1
Let CaptionDisableColor = RGB(212, 212, 212)
tmrCheck.Enabled = Ambient.UserMode
End Sub



'Invert a color; get the opposite color for another color (i.e: white = black).
Private Function InvertColor(ByVal RValue As Integer, ByVal GValue As Integer, ByVal BValue As Integer) As Long
Dim intR As Integer, intG As Integer, intB As Integer

intR = Abs(255 - RValue)
intG = Abs(255 - GValue)
intB = Abs(255 - BValue)

InvertColor = RGB(intR, intG, intB)
End Function

'Convert a long color value to an RGB value.
Private Sub LongToRGB(ByRef RValue As Integer, ByRef GValue As Integer, ByRef BValue As Integer, ByVal ColorValue As Long)
Dim intR As Integer, intG As Integer, intB As Integer

intR = ColorValue Mod 256
intG = ((ColorValue And &HFF00) / 256&) Mod 256&
intB = (ColorValue And &HFF0000) / 65536

RValue = intR
GValue = intG
BValue = intB
End Sub

'Lightens a color judging by the offset value.
Private Function LightenColor(ByVal RValue As Integer, ByVal GValue As Integer, ByVal BValue As Integer, Optional ByVal offset As Long = 1) As Long
Dim intR As Integer, intG As Integer, intB As Integer

intR = Abs(RValue + offset)
intG = Abs(GValue + offset)
intB = Abs(BValue + offset)

LightenColor = RGB(intR, intG, intB)
End Function

'Darkens a color judging by the offset value.
Private Function DarkenColor(ByVal RValue As Integer, ByVal GValue As Integer, ByVal BValue As Integer, Optional ByVal offset As Long = 1) As Long
Dim intR As Integer, intG As Integer, intB As Integer

intR = Abs(RValue - offset)
intG = Abs(GValue - offset)
intB = Abs(BValue - offset)

DarkenColor = RGB(intR, intG, intB)
End Function

'Replace one color with another color.
Private Sub ReplaceColor(PictureObject As PictureBox, ColorValue As Long, ReplaceWith As Long)
Dim lonSW As Long, lonSH As Long
Dim lonLoopW As Long, lonLoopH As Long

PictureObject.ScaleMode = vbPixels
lonSW = PictureObject.ScaleWidth
lonSH = PictureObject.ScaleHeight

For lonLoopW = 0 To lonSW
    
    For lonLoopH = 0 To lonSH
        
        If PictureObject.Point(lonLoopW, lonLoopH) = ColorValue Then
            PictureObject.PSet (lonLoopW, lonLoopH), ReplaceWith
        End If
    
    Next lonLoopH

Next lonLoopW
End Sub
Private Sub CreatePictureMask(nPictureBoxname As PictureBox, nTraansparentColor As OLE_COLOR, nMaskColor)
nPictureBoxname.AutoSize = True
Dim x As Long, y As Long
Dim SW As Long, sh As Long
sh = nPictureBoxname.ScaleHeight
SW = nPictureBoxname.ScaleWidth

For x = 0 To SW
    For y = 0 To sh
        If nPictureBoxname.Point(x, y) = nTraansparentColor Then
        Else
            nPictureBoxname.PSet (x, y), nMaskColor
        End If
    Next
Next
End Sub
Private Function LongToSignedShort(ByVal Unsigned As Long) As Integer
If Unsigned < 32768 Then
    LongToSignedShort = CInt(Unsigned)
Else
    LongToSignedShort = CInt(Unsigned - &H10000)
End If
End Function

Private Sub DefineRect(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
SetRect udtRect, x1, y1, x2, y2
End Sub

Private Sub DrawGradientTwoColour(ByVal hdc As Long, Direction As GRADIENT_DIRECT, ByVal StartColor As Long, ByVal EndColor As Long)
Dim udtVert(1) As TRIVERTEX, udtGRect As GRADIENT_RECT
With udtVert(0)
    .x = udtRect.Left
    .y = udtRect.Top
    .Red = LongToSignedShort(CLng((StartColor And &HFF&) * 256))
    .Green = LongToSignedShort(CLng(((StartColor And &HFF00&) \ &H100&) * 256))
    .Blue = LongToSignedShort(CLng(((StartColor And &HFF0000) \ &H10000) * 256))
    .Alpha = 0&
End With

With udtVert(1)
    .x = udtRect.Right
    .y = udtRect.Bottom
    .Red = LongToSignedShort(CLng((EndColor And &HFF&) * 256))
    .Green = LongToSignedShort(CLng(((EndColor And &HFF00&) \ &H100&) * 256))
    .Blue = LongToSignedShort(CLng(((EndColor And &HFF0000) \ &H10000) * 256))
    .Alpha = 0&
End With

udtGRect.UpperLeft = 0
udtGRect.LowerRight = 1

GradientFillRect hdc, udtVert(0), 2, udtGRect, 1, Direction
End Sub

Public Function DrawGradientFourColour(ObjectHDC As Long, Left As Long, Top As Long, Width As Long, Height As Long, TopLeftColour As Long, TopRightColour As Long, BottomLeftColour As Long, BottomRightColour As Long)
    Dim bi24BitInfo     As BITMAPINFO
    Dim bBytes()        As Byte
    Dim LeftGrads()     As cRGB
    Dim RightGrads()    As cRGB
    Dim MiddleGrads()   As cRGB
    Dim TopLeft         As cRGB
    Dim TopRight        As cRGB
    Dim BottomLeft      As cRGB
    Dim BottomRight     As cRGB
    Dim iLoop           As Long
    Dim bytesWidth      As Long
    
    With TopLeft
        .Red = Red(TopLeftColour)
        .Green = Green(TopLeftColour)
        .Blue = Blue(TopLeftColour)
    End With
    
    With TopRight
        .Red = Red(TopRightColour)
        .Green = Green(TopRightColour)
        .Blue = Blue(TopRightColour)
    End With
    
    With BottomLeft
        .Red = Red(BottomLeftColour)
        .Green = Green(BottomLeftColour)
        .Blue = Blue(BottomLeftColour)
    End With
    
    With BottomRight
        .Red = Red(BottomRightColour)
        .Green = Green(BottomRightColour)
        .Blue = Blue(BottomRightColour)
    End With
    
    GradateColours LeftGrads, Height, TopLeft, BottomLeft
    GradateColours RightGrads, Height, TopRight, BottomRight
    
    With bi24BitInfo.bmiHeader
        .biBitCount = 24
        .biCompression = BI_RGB
        .biPlanes = 1
        .biSize = Len(bi24BitInfo.bmiHeader)
        .biWidth = Width
        .biHeight = 1
    End With
    
    ReDim bBytes(1 To bi24BitInfo.bmiHeader.biWidth * bi24BitInfo.bmiHeader.biHeight * 3) As Byte
    
    bytesWidth = (Width) * 3
    
    For iLoop = 0 To Height - 1
        GradateColours MiddleGrads, Width, LeftGrads(iLoop), RightGrads(iLoop)
        CopyMemory bBytes(1), MiddleGrads(0), bytesWidth
        SetDIBitsToDevice ObjectHDC, Left, Top + iLoop, bi24BitInfo.bmiHeader.biWidth, bi24BitInfo.bmiHeader.biHeight, 0, 0, 0, bi24BitInfo.bmiHeader.biHeight, bBytes(1), bi24BitInfo, DIB_RGB_COLORS
    Next iLoop
    
    
End Function

'*******************************************************************************
' GradateColours (FUNCTION)
'
' DESCRIPTION:
' This function is to blend colour1 to colour2
'*******************************************************************************
Private Function GradateColours(cResults() As cRGB, Length As Long, Colour1 As cRGB, Colour2 As cRGB)
    Dim fromR   As Integer
    Dim toR     As Integer
    Dim fromG   As Integer
    Dim toG     As Integer
    Dim fromB   As Integer
    Dim toB     As Integer
    Dim stepR   As Single
    Dim stepG   As Single
    Dim stepB   As Single
    Dim iLoop   As Long
    
    ReDim cResults(0 To Length)
    
    fromR = Colour1.Red
    fromG = Colour1.Green
    fromB = Colour1.Blue
    
    toR = Colour2.Red
    toG = Colour2.Green
    toB = Colour2.Blue
    
    stepR = Divide(toR - fromR, Length)
    stepG = Divide(toG - fromG, Length)
    stepB = Divide(toB - fromB, Length)
    
    For iLoop = 0 To Length
        cResults(iLoop).Red = fromR + (stepR * iLoop)
        cResults(iLoop).Green = fromG + (stepG * iLoop)
        cResults(iLoop).Blue = fromB + (stepB * iLoop)
    Next iLoop
End Function

'*******************************************************************************
' Blue (FUNCTION)
'
' DESCRIPTION:
' Retrieve Blue from Long
'*******************************************************************************
Private Function Blue(Colour As Long) As Long
    Blue = (Colour And &HFF0000) / &H10000
End Function

'*******************************************************************************
' Green (FUNCTION)
'
' DESCRIPTION:
' Retrieve Green as long
'*******************************************************************************
Private Function Green(Colour As Long) As Long
    Green = (Colour And &HFF00&) / &H100
End Function

'*******************************************************************************
' Red (FUNCTION)
'
' DESCRIPTION:
' Retrieve Red from Long
'*******************************************************************************
Private Function Red(Colour As Long) As Long
    Red = (Colour And &HFF&)
End Function

'*******************************************************************************
' Divide (FUNCTION)
'
' DESCRIPTION:
' Division function to avoid division by 0 error
'*******************************************************************************
Private Function Divide(Numerator, Denominator) As Single
    If Numerator = 0 Or Denominator = 0 Then
        Divide = 0
    Else
        Divide = Numerator / Denominator
    End If
End Function

