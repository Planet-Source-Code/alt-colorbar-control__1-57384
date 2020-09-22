VERSION 5.00
Begin VB.UserControl ColorBar 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2010
   FillStyle       =   0  'Solid
   ScaleHeight     =   9
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   134
   ToolboxBitmap   =   "ColorBar.ctx":0000
End
Attribute VB_Name = "ColorBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************************
' ColorBar.ctl
' The purpose of this control is to be used as a seekbar, level meter, progress bar, etc.
'**************************************************************************************************
'  Copyright © 2004, Alan Tucker, All Rights Reserved
'  Contact alan_usa@hotmail.com for usage restrictions.  With that said, for those
'  people who would use this code for personal gain, a pox on you.  It is for all
'  honest coders out there to which I provide this code with the hope that any
'  enhancements or improvements will come back to me so I can grow my programming
'  knowledge.  So, if any improvements or enhancements are made or, if you have
'  invented a wheel better than mine, show me so I can "on error, resume next." ;-)
'**************************************************************************************************
Option Explicit

'**************************************************************************************************
' Control Constant Declares
'**************************************************************************************************
Private Const COL_CNT = 2

'**************************************************************************************************
' Control Enums\Structs
'**************************************************************************************************
Public Enum eBorderStyle
     [None]
     [Fixed Single]
End Enum ' eBorderStyle

Public Enum eGradientType
     [Linear]
     [Rectangular]
End Enum ' eGradientType

' Provides colors for easy access
Public Enum ePC ' PeakColor
     DEF_BLUE = &HFF0000
     DEF_CYAN = &HFFFF00
     DEF_GREEN = &HFF00&
     DEF_ORANGE = &H80FF&
     DEF_RED = &HFF&
     DEF_WHITE = &HFFFFFF
     DEF_YELLOW = &HFFFF&
End Enum ' ePC

Public Enum eOrientation
     [Horizontal]
     [Vertical]
End Enum ' eOrientation

' Provides RGB values
Private Enum eRGB
     Red = &HFF&
     green = &HFF00&
     Blue = &HFF0000
End Enum ' eRGB

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
End Type ' BITMAPINFOHEADER


Private Type POINTAPI
     X As Long
     Y As Long
End Type ' POINTAPI

Private Type RECT
     Left As Long
     Top As Long
     Right As Long
     Bottom As Long
End Type

'**************************************************************************************************
' Control API Declares
'**************************************************************************************************
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, _
     ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, _
     ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, _
     ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, _
     ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
'**************************************************************************************************
' Control Default Properties
'**************************************************************************************************
Private Const m_def_BorderStyle = False
Private Const m_def_ForeColor = &HFFFF00
Private Const m_def_GradientEndColor = &HFF&
Private Const m_def_GradientMidColor = &HFFFF&
Private Const m_def_GradientStartColor = &HFF00&
Private Const m_def_GradientType = False
Private Const m_def_Locked = True
Private Const m_def_Orientation = False
Private Const m_def_PeakColor = &HFF&
Private Const m_def_PeakValue = False
Private Const m_def_Segmented = False
Private Const m_def_UseGradient = False
Private Const m_def_UsePeaks = False
Private Const m_def_Value = False

'**************************************************************************************************
' Control Module-Level Variables
'**************************************************************************************************
Private m_Decay As Long
Private m_Level As Long
Private m_PrevLevel As Long
Private m_PrevPeak As Long

'**************************************************************************************************
' Control Events
'**************************************************************************************************
Public Event Click(ByVal xLoc As Long)
Attribute Click.VB_Description = "Event notification that the control has been clicked.  Returns X and Y position of mouseclick."
Public Event ValueChange(lValue As Long)
Attribute ValueChange.VB_Description = "Event notification of when the value of the ColorBar changes."

'**************************************************************************************************
' Control Property Variables
'**************************************************************************************************
Dim m_ForeColor As OLE_COLOR
Dim m_GradientEndColor As OLE_COLOR
Dim m_GradientMidColor As OLE_COLOR
Dim m_GradientStartColor As OLE_COLOR
Dim m_GradientType As eGradientType
Dim m_Locked As Boolean
Dim m_Orientation As eOrientation
Dim m_PeakColor As OLE_COLOR
Dim m_PeakValue As Long
Dim m_Segmented As Boolean
Dim m_UseGradient As Boolean
Dim m_UsePeaks As Boolean
Dim m_Value As Long

'**************************************************************************************************
' Control Property Pairs
'**************************************************************************************************
Public Property Get BorderStyle() As eBorderStyle
Attribute BorderStyle.VB_Description = "Returns/Sets the ColorBar border style."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BorderStyle.VB_UserMemId = -504
     BorderStyle = UserControl.BorderStyle()
End Property ' Get BorderStyle

Public Property Let BorderStyle(New_BorderStyle As eBorderStyle)
     UserControl.BorderStyle = New_BorderStyle
     PropertyChanged "BorderStyle"
End Property ' Let BorderStyle

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/Sets the solid ColorBar color when not using a gradient"
Attribute ForeColor.VB_UserMemId = -513
     ForeColor = UserControl.ForeColor()
End Property ' GetForeColor

Public Property Let ForeColor(New_ForeColor As OLE_COLOR)
     Dim lLoop As Long
     UserControl.ForeColor() = New_ForeColor
     UserControl.FillColor() = New_ForeColor
     m_ForeColor = New_ForeColor
     DrawBar
     PropertyChanged "ForeColor"
End Property ' Let ForeColor

Public Property Get GradientEndColor() As OLE_COLOR
Attribute GradientEndColor.VB_Description = "Returns/Sets the color of the final gradient segment of the ColorBar when using gradients."
     GradientEndColor = m_GradientEndColor
End Property ' Get GradientEndColor

Public Property Let GradientEndColor(New_GradientEndColor As OLE_COLOR)
     m_GradientEndColor = New_GradientEndColor
     DrawBar
     PropertyChanged "GradientEndColor"
End Property ' Let GradientEndColor

Public Property Get GradientMidColor() As OLE_COLOR
Attribute GradientMidColor.VB_Description = "Returns/Sets the color of the middle gradient segment of the ColorBar when using gradients."
     GradientMidColor = m_GradientMidColor
End Property ' Get GradientMidColor

Public Property Let GradientMidColor(New_GradientMidColor As OLE_COLOR)
     m_GradientMidColor = New_GradientMidColor
     DrawBar
     PropertyChanged "GradientMidColor"
End Property ' Let GradientMidColor

Public Property Get GradientStartColor() As OLE_COLOR
Attribute GradientStartColor.VB_Description = "Returns/Sets the color of the initial gradient segment of the ColorBar when using gradients."
     GradientStartColor = m_GradientStartColor
End Property ' Get GradientStartColor

Public Property Let GradientStartColor(New_GradientStartColor As OLE_COLOR)
     m_GradientStartColor = New_GradientStartColor
     DrawBar
     PropertyChanged "GradientStartColor"
End Property ' Let GradientStartColor

Public Property Get GradientType() As eGradientType
     GradientType = m_GradientType
End Property ' Get GradientType

Public Property Let GradientType(New_GradientType As eGradientType)
     m_GradientType = New_GradientType
     PropertyChanged "GradientType"
End Property ' Let GradientType

Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Returns/Sets whether the value of the ColorBar can be changed using the mousepointer."
     Locked = m_Locked
End Property ' Get Locked

Public Property Let Locked(New_Locked As Boolean)
     m_Locked = New_Locked
     PropertyChanged "Locked"
End Property ' Let Locked

Public Property Get Orientation() As eOrientation
Attribute Orientation.VB_Description = "Returns/Sets whether the ColorBar is displayed horizontally or vertically."
     Orientation = m_Orientation
End Property ' Get Orientation

Public Property Let Orientation(New_Orientation As eOrientation)
     m_Orientation = New_Orientation
     Select Case m_Orientation
          Case 0
               UserControl.Height = 180
               UserControl.Width = 1200
          Case 1
               UserControl.Height = 1200
               UserControl.Width = 180
     End Select
     DrawBar
     PropertyChanged "Orientation"
End Property ' Let Orientation

Public Property Get PeakColor() As OLE_COLOR
Attribute PeakColor.VB_Description = "Returns/Sets the color of the peak bars shown when UsePeaks is enabled."
     PeakColor = m_PeakColor
End Property ' Get PeakColor

Public Property Let PeakColor(New_PeakColor As OLE_COLOR)
     m_PeakColor = New_PeakColor
     PropertyChanged "PeakColor"
End Property ' Let PeakColor

Private Property Get PeakValue() As Long
     PeakValue = m_PeakValue
End Property ' Get PeakValue

Private Property Let PeakValue(New_PeakValue As Long)
Attribute PeakValue.VB_Description = "Returns/Sets the value of the Peak bar's position."
     m_PeakValue = New_PeakValue
     PropertyChanged "PeakValue"
End Property ' Let PeakValue

Public Property Get Segmented() As Boolean
Attribute Segmented.VB_Description = "Returns/Sets whether the ColorBar appearance is solid or broken into segments."
     Segmented = m_Segmented
End Property ' Get Segmented

Public Property Let Segmented(New_Segmented As Boolean)
     m_Segmented = New_Segmented
     DrawBar
     PropertyChanged "Segmented"
End Property ' Let Segmented

Public Property Get UseGradient() As Boolean
Attribute UseGradient.VB_Description = "Returns/Sets whether the ColorBar employs gradient colors or a solid color."
     UseGradient = m_UseGradient
End Property ' Get UseGradient

Public Property Let UseGradient(New_UseGradient As Boolean)
     m_UseGradient = New_UseGradient
     DrawBar
     PropertyChanged "UseGradient"
End Property ' Let UseGradient

Public Property Get UsePeaks() As Boolean
Attribute UsePeaks.VB_Description = "Returns/Sets whether peak bars will be shown.  Should only be used if the ColorBar's value changes at rapid intervals."
     UsePeaks = m_UsePeaks
End Property ' Get UsePeaks

Public Property Let UsePeaks(New_UsePeaks As Boolean)
     m_UsePeaks = New_UsePeaks
     If New_UsePeaks = False Then
          m_PrevPeak = False
          m_PeakValue = False
     End If
     DrawBar
     PropertyChanged "UsePeaks"
End Property ' Let Peaks

Public Property Get Value() As Long
Attribute Value.VB_Description = "Returns/Sets the value of the ColorBar.  The value determines how much of the ColorBar is painted with the selected solid or gradient color."
     Value = m_Value
End Property ' Get Value

Public Property Let Value(New_Value As Long)
     If New_Value > 0 And New_Value <= 100 Then
          m_Value = New_Value
          ' Draw the bar
          DrawBar
          PropertyChanged "Value"
          RaiseEvent ValueChange(New_Value)
     ElseIf New_Value <= 0 Then
          m_Value = False
          m_Level = False
          m_PrevPeak = False
          m_Decay = False
          PeakValue = False
          UserControl.Cls
          MaskPicture = Image
     End If
End Property ' Let Value

'**************************************************************************************************
' UserControl Intrinisc Methods
'**************************************************************************************************
Private Sub UserControl_AmbientChanged(PropertyName As String)
     If PropertyName = "BackColor" Then _
          UserControl.BackColor = Ambient.BackColor
     DrawBar
End Sub ' UserControl_AmbientChanged

Private Sub UserControl_InitProperties()
     BorderStyle = m_def_BorderStyle
     ForeColor = m_def_ForeColor
     GradientEndColor = m_def_GradientEndColor
     GradientMidColor = m_def_GradientMidColor
     GradientStartColor = m_def_GradientStartColor
     GradientType = m_def_GradientType
     PeakColor = m_def_PeakColor
End Sub ' UserControl_InitProperties

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Dim lVal As Long
     ' make sure it is the left button
     If Button = 1 Then
          ' if control locked, don't want anyone messing with it
          If Not Locked Then
               ' not locked...are we horizontal or vertical
               Select Case m_Orientation
                    Case 0 ' horizontal
                         ' Calculate position
                         Value = (X / ScaleWidth) * 100
                    Case 1 ' vertical
                         ' calculate position and get the
                         ' inverse since we are using top-left coords
                         Value = 100 - (Y / ScaleHeight * 100)
                         ' If y is negative
                         If Y < 0 Then Value = 100
                         If Y > ScaleHeight Then Value = 0
               End Select
          End If
     End If
End Sub ' UserControl_MouseMove

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Dim lVal As Long
     If Button = 1 Then
          ' if control locked, don't want anyone messing with it
          If Not Locked Then
               Select Case m_Orientation
                    Case 0 ' horizontal
                         Value = (X / ScaleWidth) * 100
                    Case 1
                         Value = 100 - (Y / ScaleHeight * 100)
               End Select
          End If
     End If
End Sub ' UserControl_MouseUp

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
     With PropBag
          BorderStyle = .ReadProperty("BorderStyle", m_def_BorderStyle)
          ForeColor = .ReadProperty("ForeColor", m_def_ForeColor)
          GradientEndColor = .ReadProperty("GradientEndColor", m_def_GradientEndColor)
          GradientMidColor = .ReadProperty("GradientMidColor", m_def_GradientMidColor)
          GradientStartColor = .ReadProperty("GradientStartColor", m_def_GradientStartColor)
          GradientType = .ReadProperty("GradientType", m_def_GradientType)
          Locked = .ReadProperty("Locked", m_def_Locked)
          Orientation = .ReadProperty("Orientation", m_def_Orientation)
          PeakColor = .ReadProperty("PeakColor", m_def_PeakColor)
          Segmented = .ReadProperty("Segmented", m_def_Segmented)
          UseGradient = .ReadProperty("UseGradient", m_def_UseGradient)
          UsePeaks = .ReadProperty("UsePeaks", m_def_UsePeaks)
          Value = .ReadProperty("Value", m_def_Value)
     End With
     PeakValue = m_def_PeakValue
End Sub ' UserControl_ReadProperties

Private Sub UserControl_Resize()
     DrawBar
End Sub ' UserControl_Resize

Private Sub UserControl_Show()
     UserControl.BackColor = Extender.Container.BackColor
     DrawBar
End Sub ' UserControl_Show

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
     With PropBag
          .WriteProperty "BorderStyle", UserControl.BorderStyle, m_def_BorderStyle
          .WriteProperty "ForeColor", UserControl.ForeColor, m_def_ForeColor
          .WriteProperty "GradientEndColor", m_GradientEndColor, m_def_GradientEndColor
          .WriteProperty "GradientMidColor", m_GradientMidColor, m_def_GradientMidColor
          .WriteProperty "GradientStartColor", m_GradientStartColor, m_def_GradientStartColor
          .WriteProperty "GradientType", m_GradientType, m_def_GradientType
          .WriteProperty "Locked", m_Locked, m_def_Locked
          .WriteProperty "Orientation", m_Orientation, m_def_Orientation
          .WriteProperty "PeakColor", m_PeakColor, m_def_PeakColor
          .WriteProperty "Segmented", m_Segmented, m_def_Segmented
          .WriteProperty "UseGradient", m_UseGradient, m_def_UseGradient
          .WriteProperty "UsePeaks", m_UsePeaks, m_def_UsePeaks
          .WriteProperty "Value", m_Value, m_def_Value
     End With
End Sub ' UserControl_WriteProperties

'**************************************************************************************************
' UserControl Methods
'**************************************************************************************************
Private Sub DrawBar()
     Dim lLimit As Long
     Dim lLoop As Long
     Dim lRtn As Long
     Dim lIdx As Long
     Dim lCur As Long
     Dim lSegment As Long
     Dim lRed As Long
     Dim lGreen As Long
     Dim lBlue As Long
     Dim sglRed As Single
     Dim sglGreen As Single
     Dim sglBlue As Single
     Dim lColors As Variant
     Dim lclLevel As Long
     Dim pt As POINTAPI
     ' convert value to level
     Select Case m_Orientation
          Case 0
               m_Level = ScaleWidth * (m_Value / 100)
          Case 1
               m_Level = ScaleHeight * (m_Value / 100)
     End Select
     ' pass this value to peaklevel
     PeakValue = m_Level
     ' No colors passed so set the default colors
     If m_UseGradient Then
          lColors = Array(m_GradientStartColor, m_GradientMidColor, m_GradientEndColor)
     Else
          lColors = Array(m_ForeColor, m_ForeColor, m_ForeColor)
     End If
     ' Get our segments sizes for each color
     If m_Orientation = 0 Then
          lLimit = ScaleWidth
     Else
          lLimit = ScaleHeight
     End If
     ' Get our segments sizes for each color
     lSegment = lLimit \ COL_CNT
     ' Dimension segment array and store segments
     If lSegment <= 2 Then
          ' Not enough  real estate to draw a proper gradient
          Exit Sub
     Else
          ' Size segments array to color count and store segment sizes
          ReDim sglSegments(1 To COL_CNT)
          ' Now determine if the color count divides
          ' evenly with the scale height.  If not add
          ' remainder to the first segment
          lRtn = lLimit Mod lSegment
          ' Loop through and add segments to segment array
          For lLoop = 1 To COL_CNT
               If lLoop = 1 Then
                    ' add remainder to first segment
                    sglSegments(lLoop) = lSegment + lRtn
               Else
                    sglSegments(lLoop) = lSegment
               End If
          Next
     End If
     ' Index for ColorArray tracking
     lCur = 1
     ' Dimension color array t
     ReDim lColorArray(1 To lLimit)
     ' Loop and blend the colors stopping at the next to last color
     ' always loop 1 less than color count
    For lLoop = 1 To COL_CNT
          'Extract Red, Blue and Green values from the loop - 1 color
          lRed = (lColors(lLoop - 1) And eRGB.Red)
          lGreen = (lColors(lLoop - 1) And eRGB.green) / &H100&
          lBlue = (lColors(lLoop - 1) And eRGB.Blue) / &H10000
          'Find the range of change from one color to another
          sglRed = ColorDivide(CSng((lColors(lLoop) And eRGB.Red) - lRed), _
               sglSegments(lLoop))
          sglGreen = ColorDivide(CSng(((lColors(lLoop) And eRGB.green) / &H100&) - lGreen), _
               sglSegments(lLoop))
          sglBlue = ColorDivide(CSng(((lColors(lLoop) And eRGB.Blue) / &H10000) - lBlue), _
               sglSegments(lLoop))
          ' Create the gradients and add colors to array
          For lIdx = 1 To sglSegments(lLoop)
               lColorArray(lCur) = CLng(lRed + (sglRed * lIdx)) + (CLng(lGreen + _
                    (sglGreen * lIdx)) * &H100&) + (CLng(lBlue + (sglBlue * lIdx)) * &H10000)
               lCur = lCur + 1
          Next
     Next     ' clean the canvas
     UserControl.Cls
     ' are we horizontal or vertical
     Select Case m_Orientation
          Case 0
               ' output in segments?
               If m_Segmented Then
                    ' Loop through and output gradient stopping at level
                    For lIdx = 1 To m_Level Step 2
                         ' Set the forecolor so the right color line is drawn
                         UserControl.ForeColor = lColorArray(lIdx)
                         If ScaleWidth > 2 And (m_Level - lIdx) > 1 Then
                              SetPixel hdc, lIdx + 1, 0, lColorArray(lIdx)
                              SetPixel hdc, lIdx + 1, ScaleHeight - 1, lColorArray(lIdx)
                         End If
                         ' move the starting point of the line
                         MoveToEx hdc, lIdx, 0, pt
                         ' draw the line
                         LineTo hdc, lIdx, ScaleHeight
                    Next
               Else ' no segments
                    For lIdx = 1 To m_Level
                         ' Set the forecolor so the right color line is drawn
                         UserControl.ForeColor = lColorArray(lIdx)
                         ' Move the starting point of the line
                         MoveToEx hdc, lIdx, 0, pt
                         ' draw the line
                         LineTo hdc, lIdx, ScaleHeight
                    Next
               End If
          Case 1
               ' output in segments?
               If m_Segmented Then
                    ' Loop through and output gradient stopping at level
                    For lIdx = 1 To m_Level Step 2
                         ' Set the forecolor so the right color line is drawn
                         UserControl.ForeColor = lColorArray(lIdx)
                         If ScaleWidth > 2 And (m_Level - lIdx) > 1 Then
                              SetPixel hdc, 0, ScaleHeight - (lIdx + 1), lColorArray(lIdx)
                              SetPixel hdc, ScaleWidth - 1, ScaleHeight - (lIdx + 1), lColorArray(lIdx)
                         End If
                         ' Move the starting point of the line
                         MoveToEx hdc, 0, ScaleHeight - lIdx, pt
                         ' draw the line
                         LineTo hdc, ScaleWidth, ScaleHeight - lIdx
                    Next
               Else
                    For lIdx = 1 To m_Level
                         ' Set the forecolor so the right color line is drawn
                         UserControl.ForeColor = lColorArray(lIdx)
                         ' Move the starting point of the line
                         MoveToEx hdc, 0, ScaleHeight - lIdx, pt
                         ' draw the line
                         LineTo hdc, ScaleWidth, ScaleHeight - lIdx
                    Next
               End If
     End Select
     DoEvents
     ' reset forecolor
     UserControl.ForeColor = m_ForeColor
     ' Draw peaks
     If m_UsePeaks Then
          If PeakValue > m_PrevPeak Then
               m_PrevPeak = m_PeakValue + 1
               m_Decay = 0
          Else
               m_PrevPeak = m_PrevPeak - m_Decay
               m_Decay = m_Decay + 1
               If m_Decay = 2 Then m_Decay = m_Decay - 1
          End If
          ' If we hit zero, stop drawing peak
          If m_PrevPeak > 0 Then
               ' Are we horizontal or vertical
               If m_Orientation = 0 Then
                    If m_PrevPeak > ScaleWidth Then m_PrevPeak = ScaleWidth - 1
                    ' Set forecolor to peakcolor
                    UserControl.ForeColor = m_PeakColor
                    Rectangle hdc, m_PrevPeak, 0, m_PrevPeak + 1, ScaleHeight
               Else
                    If m_PrevPeak > ScaleHeight Then m_PrevPeak = ScaleHeight - 1
                    ' Set forecolor to peakcolor
                    UserControl.ForeColor = m_PeakColor
                    Rectangle hdc, 0, ScaleHeight - m_PrevPeak, _
                         ScaleWidth, (ScaleHeight - m_PrevPeak) - 1
               End If
          End If
     End If
     UserControl.ForeColor = m_ForeColor
     ' output drawing
     MaskPicture = Image
End Sub ' DrawBar

Private Function ColorDivide(ByVal dblNumerator As Double, _
    ByVal dblDenominator As Double) As Double
     ' Divides dblNumerator by dblDenominator if dblDenominator
     ' <> 0 to eliminate 'Division By Zero' error.
     If dblDenominator = False Then Exit Function
     ColorDivide = dblNumerator / dblDenominator
End Function ' ColorDivide

