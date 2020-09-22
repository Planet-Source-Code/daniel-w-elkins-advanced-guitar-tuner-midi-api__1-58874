VERSION 5.00
Begin VB.UserControl XPButton 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Timer tmrCheck 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   360
      Top             =   120
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
      Height          =   135
      Left            =   0
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox imgMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   0
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   135
   End
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
      Height          =   135
      Left            =   840
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "XPButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'An XP button I made.

Public Enum GRADIENT_DIRECT
    [Left to Right] = &H0
    [Top to Bottom] = &H1
End Enum

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type TRIVERTEX
   X As Long
   Y As Long
   Red As Integer
   Green As Integer
   Blue As Integer
   Alpha As Integer
End Type

Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Public Enum COLOR_STYLE
    [XP Blue] = 1
    [XP Silver] = 2
    [XP Olive Green] = 3
End Enum

Public Enum PICTURE_ALIGN
    [Left Justify] = 1
    [Right Justify] = 2
End Enum

Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function GradientFill Lib "msimg32" (ByVal hDC As Long, pVertex As Any, ByVal dwNumVertex As Long, pMesh As Any, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private udtRect As RECT

Private Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Private Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest

Public Event Click()
Public Event DoubleClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseEnters(ByVal X As Long, ByVal Y As Long)
Public Event MouseLeaves(ByVal X As Long, ByVal Y As Long)

Private strCaption As String 'Caption text.

Private oleForeColor As OLE_COLOR 'Caption text color.

Private udtColorStyle As COLOR_STYLE 'Color style of button.
Private udtCaptionAlign As AlignmentConstants 'Alignment for caption.
Private udtIconAlign As PICTURE_ALIGN 'Alignment for icon.
Private udtPoint As POINTAPI 'Current mouse position (for checking if mouse is over button).

Private bolMouseDown As Boolean 'Mouse currently down?
Private bolMouseOver As Boolean 'Mouse currently over button?
Private bolHasFocus As Boolean 'Currently has focus?
Private bolFocusDottedRect As Boolean 'Draw focus dotted rect?
Private bolEnabled As Boolean 'Enabled?

Private lonRoundValue As Long 'Rounded corners value.

Private fntFont As Font 'Caption font.

Private picIcon As Picture 'Small icon picture.
Private picIconMask As Picture 'Small icon mask picture (for transparency).

'Draw the icon on to the button.
Private Sub DrawIcon()
On Error Resume Next

Dim lonHeight As Long, lonLeft As Long

If bolEnabled = True Then
    
    If imgMask.Picture.Handle <> 0 And imgIcon.Picture.Handle <> 0 Then
        lonHeight = (UserControl.ScaleHeight * 0.5) - (imgIcon.ScaleHeight * 0.5)
        
        If udtIconAlign = [Left Justify] Then
            'Draw the icon on the left side.
            'Draw the mask.
            BitBlt UserControl.hDC, 5, lonHeight, imgMask.ScaleWidth, imgMask.ScaleHeight, imgMask.hDC, 0, 0, SRCAND
            'Overlay the mask with the icon.
            BitBlt UserControl.hDC, 5, lonHeight, imgIcon.ScaleWidth, imgIcon.ScaleHeight, imgIcon.hDC, 0, 0, SRCPAINT
        ElseIf udtIconAlign = [Right Justify] Then
            'Draw the icon on the right side.
            'Draw the mask.
            lonLeft = (UserControl.ScaleWidth - imgIcon.ScaleWidth)
            BitBlt UserControl.hDC, lonLeft - 5, lonHeight, imgMask.ScaleWidth, imgMask.ScaleHeight, imgMask.hDC, 0, 0, SRCAND
            BitBlt UserControl.hDC, lonLeft - 5, lonHeight, imgIcon.ScaleWidth, imgIcon.ScaleHeight, imgIcon.hDC, 0, 0, SRCPAINT
        End If
    
    End If

Else
    
    If imgMask.Picture.Handle <> 0 And imgIcon.Picture.Handle <> 0 Then
        lonHeight = (UserControl.ScaleHeight * 0.5) - (imgIcon.ScaleHeight * 0.5)
        Set imgDis.Picture = imgMask.Picture
        ReplaceColor imgDis, 0, 10070188
        
        If udtIconAlign = [Left Justify] Then
            BitBlt UserControl.hDC, 5, lonHeight, imgDis.ScaleWidth, imgDis.ScaleHeight, imgDis.hDC, 0, 0, SRCAND
            BitBlt UserControl.hDC, 5, lonHeight, imgDis.ScaleWidth, imgDis.ScaleHeight, imgDis.hDC, 0, 0, SRCPAINT
        ElseIf udtIconAlign = [Right Justify] Then
            lonLeft = (UserControl.ScaleWidth - imgDis.ScaleWidth)
            BitBlt UserControl.hDC, lonLeft - 5, lonHeight, imgDis.ScaleWidth, imgDis.ScaleHeight, imgDis.hDC, 0, 0, SRCAND
            BitBlt UserControl.hDC, lonLeft - 5, lonHeight - 5, imgDis.ScaleWidth, imgDis.ScaleHeight, imgDis.hDC, 0, 0, SRCPAINT
        End If
    
    End If

End If
End Sub

'Print aligned text to the button (caption).
Private Sub PrintText(ByVal TextString As String, ByVal Alignment As AlignmentConstants)
Dim lonSW As Long, lonSH As Long
Dim lonStartWidth As Long, lonStartHeight As Long

UserControl.ScaleMode = vbTwips
lonSW = UserControl.Width
lonSH = UserControl.Height

If Alignment = vbCenter Then
    lonStartWidth = (UserControl.Width * 0.5) - (UserControl.TextWidth(TextString) * 0.5)
    lonStartHeight = (UserControl.Height * 0.5) - ((UserControl.TextHeight(TextString) * 0.5) + 20)
    UserControl.CurrentX = lonStartWidth
    UserControl.CurrentY = lonStartHeight
    UserControl.Print TextString
ElseIf Alignment = vbLeftJustify Then
    lonStartWidth = 100
    lonStartHeight = (UserControl.Height * 0.5) - ((UserControl.TextHeight(TextString) * 0.5) + 20)
    UserControl.CurrentX = lonStartWidth
    UserControl.CurrentY = lonStartHeight
    UserControl.Print TextString
ElseIf Alignment = vbRightJustify Then
    lonStartWidth = (UserControl.Width - UserControl.TextWidth(TextString)) - 100
    lonStartHeight = (UserControl.Height * 0.5) - ((UserControl.TextHeight(TextString) * 0.5) + 20)
    UserControl.CurrentX = lonStartWidth
    UserControl.CurrentY = lonStartHeight
    UserControl.Print TextString
End If

UserControl.ScaleMode = vbPixels
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
lonRect = CreateRoundRectRgn(0, 0, ScaleWidth, ScaleHeight, lonRoundValue, lonRoundValue)
SetWindowRgn UserControl.hWnd, lonRect, True

strName = fntFont.Name

If Err = 0 Then
    Set UserControl.Font = fntFont
End If

'Check what style we should be using.
If udtColorStyle = [XP Blue] Then
    'Draw XP blue button.
    
    If bolEnabled = False Then
DrawDisabled:
        'Button is disabled.
        'Draw gradient background.
        DefineRect 0, 0, ScaleWidth, ScaleHeight
        DrawGradient UserControl.hDC, [Top to Bottom], 15398133, 15398133
        
        'Draw main border.
        UserControl.ForeColor = 12240841
        RoundRect UserControl.hDC, 0, 0, ScaleWidth - 1, ScaleHeight - 1, lonRoundValue, lonRoundValue
        'Draw icon.
        DrawIcon
        
        'Draw caption.
        UserControl.ForeColor = 12240841
        PrintText strCaption, udtCaptionAlign
        
        Exit Sub 'Done.
    End If
    
    'Draw gradient background.
    DefineRect 0, 0, ScaleWidth, ScaleHeight
    DrawGradient UserControl.hDC, [Top to Bottom], 16514300, 15397104
        
    'Draw main border.
    UserControl.ForeColor = 7617536
    RoundRect UserControl.hDC, 0, 0, ScaleWidth - 1, ScaleHeight - 1, lonRoundValue, lonRoundValue
    
    If bolMouseOver = True And bolMouseDown = False Then
        'Draw mouse over lines.
        'Top line.
        UserControl.Line (2, 1)-(ScaleWidth - 3, 1), 13627647
        'Left line.
        UserControl.Line (1, 2)-(1, ScaleHeight - 3), 5817338
        'Draw bottom line.
        UserControl.Line (2, ScaleHeight - 3)-(ScaleWidth - 3, ScaleHeight - 3), 38885
        'Draw right line.
        UserControl.Line (ScaleWidth - 3, 2)-(ScaleWidth - 3, ScaleHeight - 3), 6933244
        
        'Draw inner lines.
        'Top inner line.
        UserControl.Line (2, 2)-(ScaleWidth - 3, 2), 9033981
        'Left inner line.
        UserControl.Line (2, 2)-(2, ScaleHeight - 3), 6342907
        'Bottom inner line.
        UserControl.Line (2, ScaleHeight - 4)-(ScaleWidth - 3, ScaleHeight - 4), 3191800
        'Right inner line.
        UserControl.Line (ScaleWidth - 4, 2)-(ScaleWidth - 4, ScaleHeight - 3), 6408186
        
        GoTo XPBlueDone
    End If
    
    If bolHasFocus = True And bolMouseDown = False Then
        'Draw has focus lines.
        'Top line.
        UserControl.Line (2, 1)-(ScaleWidth - 3, 1), 16771022
        'Left line.
        UserControl.Line (1, 2)-(1, ScaleHeight - 3), 15383452
        'Draw bottom line.
        UserControl.Line (2, ScaleHeight - 3)-(ScaleWidth - 3, ScaleHeight - 3), 15630953
        'Draw right line.
        UserControl.Line (ScaleWidth - 3, 2)-(ScaleWidth - 3, ScaleHeight - 3), 15448988
        
        'Draw inner lines.
        'Top inner line.
        UserControl.Line (2, 2)-(ScaleWidth - 3, 2), 16176316
        'Left inner line.
        UserControl.Line (2, 2)-(2, ScaleHeight - 3), 15449245
        'Bottom inner line.
        UserControl.Line (2, ScaleHeight - 4)-(ScaleWidth - 3, ScaleHeight - 4), 14986633
        'Right inner line.
        UserControl.Line (ScaleWidth - 4, 2)-(ScaleWidth - 4, ScaleHeight - 3), 15448989
        
        If bolFocusDottedRect = True Then
            'Draw dotted focus rect.
            DrawDottedFocusRect
        End If
        
        GoTo XPBlueDone
    
    End If
    
    If bolMouseDown = True Then
        'Draw gradient for mouse down.
        DefineRect 0, 0, ScaleWidth, ScaleHeight
        DrawGradient UserControl.hDC, [Top to Bottom], 14542053, 14344930
        
        'Draw main border.
        UserControl.ForeColor = 7617536
        RoundRect UserControl.hDC, 0, 0, ScaleWidth - 1, ScaleHeight - 1, lonRoundValue, lonRoundValue
        
        'Draw button in mouse down state.
        'Top line.
        UserControl.Line (2, 1)-(ScaleWidth - 3, 1), 12700881
        'Left line.
        UserControl.Line (1, 2)-(1, ScaleHeight - 3), 13358295
        'Draw bottom line.
        UserControl.Line (2, ScaleHeight - 3)-(ScaleWidth - 3, ScaleHeight - 3), 15659506
        'Draw right line.
        UserControl.Line (ScaleWidth - 3, 2)-(ScaleWidth - 3, ScaleHeight - 3), 14410724
        
        'Draw inner lines.
        'Top inner line.
        UserControl.Line (2, 2)-(ScaleWidth - 3, 2), 13621468
        'Left inner line.
        UserControl.Line (2, 2)-(2, ScaleHeight - 3), 13884381
        'Bottom inner line.
        UserControl.Line (2, ScaleHeight - 4)-(ScaleWidth - 3, ScaleHeight - 4), 14936554
        'Right inner line.
        UserControl.Line (ScaleWidth - 4, 2)-(ScaleWidth - 4, ScaleHeight - 3), 14410467
        
        If bolHasFocus = True And bolFocusDottedRect = True Then
            DrawDottedFocusRect
        End If
        
        GoTo XPBlueDone
    End If
    
XPBlueDone:
    'Drawing complete, now we just need to draw the Icon and caption.
    'Draw icon.
    DrawIcon
    'Draw caption.
    UserControl.ForeColor = oleForeColor
    PrintText strCaption, udtCaptionAlign
    
    Exit Sub 'All done, stop here (all other statements (code) are omitted and not executed).

ElseIf udtColorStyle = [XP Olive Green] Then
    'Draw XP olive green button.
    
    If bolEnabled = False Then
        GoTo DrawDisabled
    End If
    
    'Draw gradient background.
    DefineRect 0, 0, ScaleWidth, ScaleHeight
    DrawGradient UserControl.hDC, [Top to Bottom], 15925246, 14413555
        
    'Draw main border.
    UserControl.ForeColor = 418359
    RoundRect UserControl.hDC, 0, 0, ScaleWidth - 1, ScaleHeight - 1, lonRoundValue, lonRoundValue
    
    If bolMouseOver = True And bolMouseDown = False Then
        'Draw mouse over lines.
        'Top line.
        UserControl.Line (2, 1)-(ScaleWidth - 3, 1), 9815548
        'Left line.
        UserControl.Line (1, 2)-(1, ScaleHeight - 3), 7777000
        'Draw bottom line.
        UserControl.Line (2, ScaleHeight - 3)-(ScaleWidth - 3, ScaleHeight - 3), 2454223
        'Draw right line.
        UserControl.Line (ScaleWidth - 3, 2)-(ScaleWidth - 3, ScaleHeight - 3), 7317223
        
        'Draw inner lines.
        'Top inner line.
        UserControl.Line (2, 2)-(ScaleWidth - 3, 2), 9879277
        'Left inner line.
        UserControl.Line (2, 2)-(2, ScaleHeight - 3), 7317223
        'Bottom inner line.
        UserControl.Line (2, ScaleHeight - 4)-(ScaleWidth - 3, ScaleHeight - 4), 5214691
        'Right inner line.
        UserControl.Line (ScaleWidth - 4, 2)-(ScaleWidth - 4, ScaleHeight - 3), 7842791
        
        GoTo XPOliveDone
    End If
    
    If bolHasFocus = True And bolMouseDown = False Then
        'Draw has focus lines.
        'Top line.
        UserControl.Line (2, 1)-(ScaleWidth - 3, 1), 9425346
        'Left line.
        UserControl.Line (1, 2)-(1, ScaleHeight - 3), 6801312
        'Draw bottom line.
        UserControl.Line (2, ScaleHeight - 3)-(ScaleWidth - 3, ScaleHeight - 3), 6727592
        'Draw right line.
        UserControl.Line (ScaleWidth - 3, 2)-(ScaleWidth - 3, ScaleHeight - 3), 6866593
        
        'Draw inner lines.
        'Top inner line.
        UserControl.Line (2, 2)-(ScaleWidth - 3, 2), 8440753
        'Left inner line.
        UserControl.Line (2, 2)-(2, ScaleHeight - 3), 6276251
        'Bottom inner line.
        UserControl.Line (2, ScaleHeight - 4)-(ScaleWidth - 3, ScaleHeight - 4), 5554576
        'Right inner line.
        UserControl.Line (ScaleWidth - 4, 2)-(ScaleWidth - 4, ScaleHeight - 3), 6144920
        
        If bolFocusDottedRect = True Then
            'Draw dotted focus rect.
            DrawDottedFocusRect
        End If
        
        GoTo XPOliveDone
    
    End If
    
    If bolMouseDown = True Then
        'Draw gradient for mouse down.
        DefineRect 0, 0, ScaleWidth, ScaleHeight
        DrawGradient UserControl.hDC, [Top to Bottom], 13821678, 13559020
        
        'Draw main border.
        UserControl.ForeColor = 418359
        RoundRect UserControl.hDC, 0, 0, ScaleWidth - 1, ScaleHeight - 1, lonRoundValue, lonRoundValue
        
        'Draw button in mouse down state.
        'Top line.
        UserControl.Line (2, 1)-(ScaleWidth - 3, 1), 11849183
        'Left line.
        UserControl.Line (1, 2)-(1, ScaleHeight - 3), 12571875
        'Draw bottom line.
        UserControl.Line (2, ScaleHeight - 3)-(ScaleWidth - 3, ScaleHeight - 3), 15004920
        'Draw right line.
        UserControl.Line (ScaleWidth - 3, 2)-(ScaleWidth - 3, ScaleHeight - 3), 13624814
        
        'Draw inner lines.
        'Top inner line.
        UserControl.Line (2, 2)-(ScaleWidth - 3, 2), 12835303
        'Left inner line.
        UserControl.Line (2, 2)-(2, ScaleHeight - 3), 13032680
        'Bottom inner line.
        UserControl.Line (2, ScaleHeight - 4)-(ScaleWidth - 3, ScaleHeight - 4), 14216434
        'Right inner line.
        UserControl.Line (ScaleWidth - 4, 2)-(ScaleWidth - 4, ScaleHeight - 3), 13624557
        
        If bolHasFocus = True And bolFocusDottedRect = True Then
            DrawDottedFocusRect
        End If
        
        GoTo XPOliveDone
    End If
    
XPOliveDone:
    'Drawing complete, now we just need to draw the Icon and caption.
    'Draw icon.
    DrawIcon
    'Draw caption.
    UserControl.ForeColor = oleForeColor
    PrintText strCaption, udtCaptionAlign
    
    Exit Sub 'All done, stop here (all other statements (code) are omitted and not executed).

ElseIf udtColorStyle = [XP Silver] Then
        'Draw XP blue button.
    
    If bolEnabled = False Then
        GoTo DrawDisabled
    End If
    
    'Draw gradient background.
    DefineRect 0, 0, ScaleWidth, ScaleHeight
    DrawGradient UserControl.hDC, [Top to Bottom], 16777215, 14140870
        
    'Draw main border.
    UserControl.ForeColor = 7617536
    RoundRect UserControl.hDC, 0, 0, ScaleWidth - 1, ScaleHeight - 1, lonRoundValue, lonRoundValue
    
    If bolMouseOver = True And bolMouseDown = False Then
        'Draw mouse over lines.
        'Top line.
        UserControl.Line (2, 1)-(ScaleWidth - 3, 1), 13627647
        'Left line.
        UserControl.Line (1, 2)-(1, ScaleHeight - 3), 5817338
        'Draw bottom line.
        UserControl.Line (2, ScaleHeight - 3)-(ScaleWidth - 3, ScaleHeight - 3), 38885
        'Draw right line.
        UserControl.Line (ScaleWidth - 3, 2)-(ScaleWidth - 3, ScaleHeight - 3), 6933244
        
        'Draw inner lines.
        'Top inner line.
        UserControl.Line (2, 2)-(ScaleWidth - 3, 2), 9033981
        'Left inner line.
        UserControl.Line (2, 2)-(2, ScaleHeight - 3), 6342907
        'Bottom inner line.
        UserControl.Line (2, ScaleHeight - 4)-(ScaleWidth - 3, ScaleHeight - 4), 3191800
        'Right inner line.
        UserControl.Line (ScaleWidth - 4, 2)-(ScaleWidth - 4, ScaleHeight - 3), 6408186
        
        GoTo XPSilverDone
    End If
    
    If bolHasFocus = True And bolMouseDown = False Then
        'Draw has focus lines.
        'Top line.
        UserControl.Line (2, 1)-(ScaleWidth - 3, 1), 16771022
        'Left line.
        UserControl.Line (1, 2)-(1, ScaleHeight - 3), 15515296
        'Draw bottom line.
        UserControl.Line (2, ScaleHeight - 3)-(ScaleWidth - 3, ScaleHeight - 3), 15630953
        'Draw right line.
        UserControl.Line (ScaleWidth - 3, 2)-(ScaleWidth - 3, ScaleHeight - 3), 15448988
        
        'Draw inner lines.
        'Top inner line.
        UserControl.Line (2, 2)-(ScaleWidth - 3, 2), 16176316
        'Left inner line.
        UserControl.Line (2, 2)-(2, ScaleHeight - 3), 16777215
        'Bottom inner line.
        UserControl.Line (2, ScaleHeight - 4)-(ScaleWidth - 3, ScaleHeight - 4), 14986633
        'Right inner line.
        UserControl.Line (ScaleWidth - 4, 2)-(ScaleWidth - 4, ScaleHeight - 3), 16777215
        
        If bolFocusDottedRect = True Then
            'Draw dotted focus rect.
            DrawDottedFocusRect
        End If
        
        GoTo XPSilverDone
    
    End If
    
    If bolMouseDown = True Then
        'Draw gradient for mouse down.
        DefineRect 0, 0, ScaleWidth, ScaleHeight
        DrawGradient UserControl.hDC, [Top to Bottom], 12430252, 16777215
        
        'Draw main border.
        UserControl.ForeColor = 7617536
        RoundRect UserControl.hDC, 0, 0, ScaleWidth - 1, ScaleHeight - 1, lonRoundValue, lonRoundValue
        
        'Draw button in mouse down state.
        'Top line.
        UserControl.Line (2, 1)-(ScaleWidth - 3, 1), 16777215
        'Left line.
        UserControl.Line (1, 2)-(1, ScaleHeight - 3), 16777215
        'Draw bottom line.
        UserControl.Line (2, ScaleHeight - 3)-(ScaleWidth - 3, ScaleHeight - 3), 16777215
        'Draw right line.
        UserControl.Line (ScaleWidth - 3, 2)-(ScaleWidth - 3, ScaleHeight - 3), 16777215
        
        If bolHasFocus = True And bolFocusDottedRect = True Then
            DrawDottedFocusRect
        End If
        
        GoTo XPSilverDone
    End If
    
XPSilverDone:
    'Drawing complete, now we just need to draw the Icon and caption.
    'Draw icon.
    DrawIcon
    'Draw caption.
    UserControl.ForeColor = oleForeColor
    PrintText strCaption, udtCaptionAlign
    
    Exit Sub 'All done, stop here (all other statements (code) are omitted and not executed).

End If

If bolMouseOver = True And bolFocusDottedRect = True Then DrawDottedFocusRect
End Sub

Public Property Get CaptionAlignment() As AlignmentConstants
CaptionAlignment = udtCaptionAlign
End Property

Public Property Let CaptionAlignment(ByVal NewValue As AlignmentConstants)
udtCaptionAlign = NewValue
PropertyChanged "CaptionAlignment"
PaintControl
End Property

Public Property Get IconAlignment() As PICTURE_ALIGN
IconAlignment = udtIconAlign
End Property

Public Property Let IconAlignment(ByVal NewValue As PICTURE_ALIGN)
udtIconAlign = NewValue
PropertyChanged "IconAlignment"
PaintControl
End Property

Public Property Get Caption() As String
Caption = strCaption
End Property

Public Property Let Caption(ByVal NewValue As String)
strCaption = NewValue
PropertyChanged "Caption"
PaintControl
End Property

Public Property Get ForeColor() As OLE_COLOR
ForeColor = oleForeColor
End Property

Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
oleForeColor = NewValue
PropertyChanged "ForeColor"
PaintControl
End Property

Public Property Get ColorStyle() As COLOR_STYLE
ColorStyle = udtColorStyle
End Property

Public Property Let ColorStyle(ByVal NewValue As COLOR_STYLE)
udtColorStyle = NewValue
PropertyChanged "ColorStyle"
PaintControl
End Property

Public Property Get FocusDottedRect() As Boolean
FocusDottedRect = bolFocusDottedRect
End Property

Public Property Let FocusDottedRect(ByVal NewValue As Boolean)
bolFocusDottedRect = NewValue
PropertyChanged "FocusDottedRect"
PaintControl
End Property

Public Property Get Enabled() As Boolean
Enabled = bolEnabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
bolEnabled = NewValue
PropertyChanged "Enabled"
PaintControl
End Property

Public Property Get FontType() As Font
Set FontType = fntFont
End Property

Public Property Set FontType(ByVal NewValue As Font)
Set fntFont = NewValue
Set UserControl.Font = NewValue
PropertyChanged "FontType"
PaintControl
End Property

Public Property Get IconMask() As Picture
Set IconMask = picIconMask
End Property

Public Property Set IconMask(ByVal NewValue As Picture)
Set picIconMask = NewValue
Set imgMask.Picture = NewValue
PropertyChanged "IconMask"
PaintControl
End Property

Public Property Get Icon() As Picture
Set Icon = picIcon
End Property

Public Property Set Icon(ByVal NewValue As Picture)
Set picIcon = NewValue
Set imgIcon.Picture = NewValue
PropertyChanged "Icon"
PaintControl
End Property

Public Property Get RoundedValue() As Long
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
lonCurHWND = WindowFromPoint(udtPoint.X, udtPoint.Y)

If bolMouseOver = False Then
    
    If lonCurHWND = UserControl.hWnd Then
        bolMouseOver = True
        PaintControl
        RaiseEvent MouseEnters(udtPoint.X, udtPoint.Y)
    End If

Else
    
    If lonCurHWND <> UserControl.hWnd Then
        bolMouseOver = False
        PaintControl
        RaiseEvent MouseLeaves(udtPoint.X, udtPoint.Y)
    End If

End If

tmrCheck.Enabled = True
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

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If bolEnabled = True Then
    RaiseEvent MouseDown(Button, Shift, X, Y)
    
    If Button = 1 Then
        bolMouseDown = True
        PaintControl
    End If

End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If bolEnabled = True Then
    RaiseEvent MouseMove(Button, Shift, X, Y)
End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If bolEnabled = True Then
    RaiseEvent MouseUp(Button, Shift, X, Y)
    
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
    Let ForeColor = .ReadProperty("ForeColor", 0)
    Let ColorStyle = .ReadProperty("ColorStyle", [XP Blue])
    Let FocusDottedRect = .ReadProperty("FocusDottedRect", True)
    Let Enabled = .ReadProperty("Enabled", True)
    Set FontType = .ReadProperty("FontType", Ambient.Font)
    Set Icon = .ReadProperty("Icon", Nothing)
    Set IconMask = .ReadProperty("IconMask", Nothing)
    Let RoundedValue = .ReadProperty("RoundedValue", 5)
    Let CaptionAlignment = .ReadProperty("CaptionAlignment", vbCenter)
    Let IconAlignment = .ReadProperty("IconAlignment", [Left Justify])
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
    .WriteProperty "ColorStyle", udtColorStyle, [XP Blue]
    .WriteProperty "FocusDottedRect", bolFocusDottedRect, True
    .WriteProperty "Enabled", bolEnabled, True
    .WriteProperty "FontType", fntFont, Ambient.Font
    .WriteProperty "Icon", picIcon, Nothing
    .WriteProperty "IconMask", picIconMask, Nothing
    .WriteProperty "RoundedValue", lonRoundValue, 5
    .WriteProperty "CaptionAlignment", udtCaptionAlign, vbCenter
    .WriteProperty "IconAlignment", udtIconAlign, [Left Justify]
End With
End Sub

Private Sub UserControl_InitProperties()
Let Caption = Ambient.DisplayName
Let ForeColor = 0
Let ColorStyle = [XP Blue]
Let FocusDottedRect = True
Let Enabled = True
Set Icon = Nothing
Set IconMask = Nothing
Let RoundedValue = 5
Let CaptionAlignment = vbCenter
Let IconAlignment = [Left Justify]
Set FontType = Ambient.Font
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
Private Function LightenColor(ByVal RValue As Integer, ByVal GValue As Integer, ByVal BValue As Integer, Optional ByVal OffSet As Long = 1) As Long
Dim intR As Integer, intG As Integer, intB As Integer

intR = Abs(RValue + OffSet)
intG = Abs(GValue + OffSet)
intB = Abs(BValue + OffSet)

LightenColor = RGB(intR, intG, intB)
End Function

'Darkens a color judging by the offset value.
Private Function DarkenColor(ByVal RValue As Integer, ByVal GValue As Integer, ByVal BValue As Integer, Optional ByVal OffSet As Long = 1) As Long
Dim intR As Integer, intG As Integer, intB As Integer

intR = Abs(RValue - OffSet)
intG = Abs(GValue - OffSet)
intB = Abs(BValue - OffSet)

DarkenColor = RGB(intR, intG, intB)
End Function

'Replace one color with another color.
Private Sub ReplaceColor(PictureObject As PictureBox, ColorValue As Long, ReplaceWith As Long)
Dim lonSW As Long, lonSH As Long
Dim lonLoopW As Long, lonLoopH As Long

PictureObject.ScaleMode = vbPixels
lonSW = PictureObject.ScaleWidth
lonSH = PictureObject.ScaleHeight

For lonLoopW = 1 To lonSW
    
    For lonLoopH = 1 To lonSH
        
        If PictureObject.Point(lonLoopW, lonLoopH) = ColorValue Then
            PictureObject.PSet (lonLoopW, lonLoopH), ReplaceWith
        End If
    
    Next lonLoopH

Next lonLoopW
End Sub

Private Function LongToSignedShort(ByVal Unsigned As Long) As Integer
If Unsigned < 32768 Then
    LongToSignedShort = CInt(Unsigned)
Else
    LongToSignedShort = CInt(Unsigned - &H10000)
End If
End Function

Private Sub DefineRect(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
SetRect udtRect, X1, Y1, X2, Y2
End Sub

Private Sub DrawGradient(ByVal hDC As Long, Direction As GRADIENT_DIRECT, ByVal StartColor As Long, ByVal EndColor As Long)
Dim udtVert(1) As TRIVERTEX, udtGRect As GRADIENT_RECT

With udtVert(0)
    .X = udtRect.Left
    .Y = udtRect.Top
    .Red = LongToSignedShort(CLng((StartColor And &HFF&) * 256))
    .Green = LongToSignedShort(CLng(((StartColor And &HFF00&) \ &H100&) * 256))
    .Blue = LongToSignedShort(CLng(((StartColor And &HFF0000) \ &H10000) * 256))
    .Alpha = 0&
End With

With udtVert(1)
    .X = udtRect.Right
    .Y = udtRect.Bottom
    .Red = LongToSignedShort(CLng((EndColor And &HFF&) * 256))
    .Green = LongToSignedShort(CLng(((EndColor And &HFF00&) \ &H100&) * 256))
    .Blue = LongToSignedShort(CLng(((EndColor And &HFF0000) \ &H10000) * 256))
    .Alpha = 0&
End With

udtGRect.UpperLeft = 0
udtGRect.LowerRight = 1

GradientFillRect hDC, udtVert(0), 2, udtGRect, 1, Direction
End Sub


