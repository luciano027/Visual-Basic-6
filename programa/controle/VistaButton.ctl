VERSION 5.00
Begin VB.UserControl VistaButton 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1185
   ScaleHeight     =   1125
   ScaleWidth      =   1185
End
Attribute VB_Name = "VistaButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
' VistaButton (P) 2012 BobbySteels
'******************************************************************************

'******************************************************************************
' @file VistaButton.ctl
' @brief Vista style button
' @author BobbySteels (bobbysteels@web.de)
' @version 1.01
'
' History: 02/18/2012 Initial version
'          02/19/2012 Create Timer dynamically
'          02/20/2012 Added NoBackground and PictureOffset property
'          03/22/2012 Added hWnd property
'          03/27/2012 Fixed bug with focus
'
'******************************************************************************

Option Explicit

Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal HDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal HDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal HDC As Long, ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal HDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Sub BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long)
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hDCDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hDCSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal HDC As Long) As Long

Private Const DT_CALCRECT As Long = &H400
Private Const DT_WORDBREAK As Long = &H10
Private Const DT_CENTER As Long = &H1
Private Const RGN_DIFF As Long = 4
Private Const GRADIENT_FILL_RECT_V = &H1

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type TRIVERTEX
    X As Long
    Y As Long
    red As Integer
    green As Integer
    blue As Integer
    Alpha As Integer
End Type

Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Private Type RGBCOLOR
    red As Integer
    green As Integer
    blue As Integer
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Enum ButtonStateEnum
    btsNormal = 0
    btsHover = 1
    btsDown = 2
    btsFocus = 3
    btsDisabled = 4
End Enum

Private dCaption As String
Private dForeColor As OLE_COLOR
Private dPicture As StdPicture
Private dPictures As Integer
Private dUseMaskColor As Boolean
Private dMaskColor As OLE_COLOR
Private dNoBackground As Boolean
Private dPictureOffset As Integer

Dim dHDC As Long
Dim dPictureW As Integer
Dim dPictureH As Integer
Dim dFocus As Boolean
Dim dEnabled As Boolean
Dim dTabStop As Boolean
Dim dTabStopChanged As Boolean
Dim currentState As ButtonStateEnum
Dim mDown As Boolean
Dim oldWidth As Integer
Dim oldHeight As Integer

Event Click()
Event OLECompleteDrag(Effect As Long)
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Event OLESetData(Data As DataObject, DataFormat As Integer)
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseIn()
Event MouseOut()

Private WithEvents Hover As VB.Timer
Attribute Hover.VB_VarHelpID = -1

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get PictureOffset() As Integer
    PictureOffset = dPictureOffset
End Property

Public Property Let PictureOffset(nValue As Integer)
    dPictureOffset = nValue
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(nValue As OLE_COLOR)
    UserControl.BackColor = nValue
    UserControl_Resize
End Property

Public Property Get NoBackground() As Boolean
    NoBackground = dNoBackground
End Property

Public Property Let NoBackground(nValue As Boolean)
    dNoBackground = nValue
    UserControl_Resize
End Property

Private Sub InitHover()
    If Not Hover Is Nothing Then Exit Sub
    Set Hover = Controls.Add("VB.Timer", "Hover")
    Hover.Interval = 10
End Sub

Private Sub DestroyHover()
    If Hover Is Nothing Then Exit Sub
    Controls.Remove "Hover"

    If dEnabled And Not mDown Then switchState IIf(dFocus, btsFocus, btsNormal)
    Set Hover = Nothing
    RaiseEvent MouseOut
End Sub

Private Sub Hover_Timer()
    If Not isMouseOver Then DestroyHover
End Sub

Public Property Get Enabled() As Boolean
    Enabled = dEnabled
End Property

Public Property Let Enabled(nValue As Boolean)
    dEnabled = nValue
    currentState = IIf(dEnabled, btsNormal, btsDisabled)
    modifyTabStop
    UserControl_Resize
End Property

Public Property Get MaskColor() As OLE_COLOR
    MaskColor = dMaskColor
End Property

Public Property Let MaskColor(nValue As OLE_COLOR)
    dMaskColor = nValue
    UserControl_Resize
End Property

Public Property Get UseMaskColor() As Boolean
    UseMaskColor = dUseMaskColor
End Property

Public Property Let UseMaskColor(nValue As Boolean)
    dUseMaskColor = nValue
    UserControl_Resize
End Property

Public Property Get Pictures() As Integer
    Pictures = dPictures
End Property

Public Property Let Pictures(nValue As Integer)
    dPictures = nValue
    If dPictures < 0 Then dPictures = 1
    If dPictures > 2 Then dPictures = 2
    UserControl_Resize
End Property

Public Property Get Picture() As StdPicture
    Set Picture = dPicture
End Property

Public Property Set Picture(nValue As StdPicture)
    Set dPicture = nValue
    pictureToHDC
    UserControl_Resize
End Property

Private Sub pictureToHDC()

    If Not dPicture Is Nothing Then
        If dPicture.Width > 0 And dPicture.Height > 0 Then
            dHDC = CreateCompatibleDC(UserControl.HDC)
            DeleteObject SelectObject(dHDC, dPicture.Handle)
            dPictureW = Round(UserControl.ScaleX(dPicture.Width, vbHimetric, vbPixels))
            dPictureH = Round(UserControl.ScaleY(dPicture.Height, vbHimetric, vbPixels))
        Else
            dHDC = False
        End If
    Else
        dHDC = False
    End If

End Sub

Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property

Public Property Set Font(nValue As StdFont)
    Set UserControl.Font = nValue
    UserControl_Resize
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = dForeColor
End Property

Public Property Let ForeColor(nValue As OLE_COLOR)
    dForeColor = nValue
    UserControl_Resize
End Property

Public Property Get Caption() As String
    Caption = dCaption
End Property

Public Property Let Caption(nValue As String)
    dCaption = nValue
    UserControl_Resize
End Property

Private Function switchState(NewState As ButtonStateEnum) As Boolean

    Dim erg As Boolean
    erg = Not (NewState = currentState)
    currentState = NewState

    If erg Then UserControl_Resize
    switchState = erg

End Function

Private Function CCol(ByVal Col As Byte) As Integer
    Dim erg As Long
    erg = Col * &H100&
    If Col > &H7F Then erg = erg - &H10000
    CCol = Int(erg)
End Function

Private Sub UserControl_Click()
    If dEnabled Then RaiseEvent Click
End Sub

Private Sub UserControl_GotFocus()
    dFocus = True
    If dEnabled And Not mDown And Not isMouseOver Then
        switchState btsFocus
    End If
End Sub

Private Sub UserControl_Initialize()
    With UserControl
        .ScaleMode = vbPixels
        .AutoRedraw = True
    End With
    currentState = btsNormal
    mDown = False
    UserControl.OLEDropMode = 0
End Sub

Private Sub UserControl_InitProperties()
    dForeColor = RGB(0, 0, 0)
    dCaption = Ambient.DisplayName
    UserControl.FontName = "Arial"
    dPictures = 2
    UseMaskColor = True
    dMaskColor = vbGreen
    dEnabled = True
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 Then UserControl_MouseDown 0, 0, 0, 0
    If dEnabled Then RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    If dEnabled Then RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 Then
        UserControl_MouseUp 0, 0, IIf(isMouseOver, 0, -1), 0
        If dEnabled Then RaiseEvent Click
    End If
    If dEnabled Then RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_LostFocus()
    dFocus = False
    If dEnabled And Not isMouseOver Then switchState btsNormal
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If dEnabled Then
        RaiseEvent MouseDown(Button, Shift, X, Y)
        mDown = True
        switchState btsDown
        UserControl.Refresh
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If dEnabled Then
        RaiseEvent MouseMove(Button, Shift, X, Y)

        If X >= 0 And Y >= 0 And X < UserControl.ScaleWidth And _
         Y < UserControl.ScaleHeight Then
            RaiseEvent MouseIn
            If switchState(IIf(mDown, btsDown, btsHover)) Then
                InitHover
            End If
        Else
            switchState btsHover
        End If
    End If

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If dEnabled Then
        RaiseEvent MouseUp(Button, Shift, X, Y)
        mDown = False
        If X >= 0 And Y >= 0 And X < UserControl.ScaleWidth And _
         Y < UserControl.ScaleHeight Then
            switchState btsHover

        Else
            switchState IIf(dFocus, btsFocus, btsNormal)

        End If
    End If
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
    If dEnabled Then RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If dEnabled Then RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    If dEnabled Then RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    If dEnabled Then RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    If dEnabled Then RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    If dEnabled Then RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        dCaption = .ReadProperty("Caption", dCaption)
        dForeColor = .ReadProperty("ForeColor", dForeColor)
        Set UserControl.Font = .ReadProperty("Font", UserControl.Font)
        Set dPicture = .ReadProperty("Picture", UserControl.Picture)
        dPictures = .ReadProperty("Pictures", dPictures)
        dUseMaskColor = .ReadProperty("UseMaskColor", dUseMaskColor)
        dMaskColor = .ReadProperty("MaskColor", dMaskColor)
        dEnabled = .ReadProperty("Enabled", dEnabled)
        dNoBackground = .ReadProperty("NoBackground", dNoBackground)
        UserControl.BackColor = .ReadProperty("BackColor", vbButtonFace)
        dPictureOffset = .ReadProperty("PictureOffset", dPictureOffset)
        pictureToHDC
    End With
    If Not dEnabled Then currentState = btsDisabled
    modifyTabStop
    UserControl_Resize
End Sub

Private Sub modifyTabStop()
    If Not Ambient.UserMode Then Exit Sub
    With Extender
        If dEnabled Then
            If Not dTabStop = .TabStop And dTabStopChanged Then
                .TabStop = dTabStop
            End If
        Else
            If .TabStop Then
                dTabStop = .TabStop
                .TabStop = False
                dTabStopChanged = True
            End If
        End If
    End With
End Sub

Private Sub UserControl_Resize()
    drawButton currentState
    With UserControl
        If Not (.ScaleWidth = oldWidth And _
                .ScaleHeight = oldHeight) Then
            MakeRegion
            oldWidth = .ScaleWidth
            oldHeight = .ScaleHeight
        End If
    End With
    UserControl.Refresh
End Sub

Private Sub drawButton(State As ButtonStateEnum)

    If dNoBackground And (State = btsNormal Or State = btsFocus _
                          Or State = btsDisabled) Then
        UserControl.Cls
        DrawCaption
        Exit Sub
    End If

    Dim colorsGr(7) As RGBCOLOR
    Dim colors(20) As OLE_COLOR

    Select Case State

    Case btsNormal
        colorsGr(0) = setRGB(242, 242, 242): colorsGr(1) = setRGB(235, 235, 235)
        colorsGr(2) = setRGB(221, 221, 221): colorsGr(3) = setRGB(207, 207, 207)
        colorsGr(4) = setRGB(251, 251, 251): colorsGr(5) = setRGB(250, 250, 250)
        colorsGr(6) = setRGB(246, 246, 246): colorsGr(7) = setRGB(243, 243, 243)
        colors(0) = RGB(112, 112, 112): colors(1) = RGB(252, 252, 252)
        colors(2) = RGB(144, 144, 144): colors(3) = RGB(119, 119, 119)
        colors(4) = RGB(135, 136, 136): colors(5) = RGB(145, 145, 145)
        colors(6) = RGB(232, 232, 232): colors(7) = RGB(117, 117, 117)
        colors(8) = RGB(234, 234, 234): colors(9) = RGB(250, 250, 250)
        colors(10) = RGB(112, 112, 112): colors(11) = RGB(252, 252, 252)
        colors(12) = RGB(134, 134, 134): colors(13) = RGB(117, 117, 117)
        colors(14) = RGB(136, 136, 136): colors(15) = RGB(143, 143, 143)
        colors(16) = RGB(228, 228, 228): colors(17) = RGB(117, 117, 117)
        colors(18) = RGB(226, 226, 226): colors(19) = RGB(236, 236, 236)
        colors(20) = RGB(243, 243, 243)

    Case btsHover
        colorsGr(0) = setRGB(234, 246, 253): colorsGr(1) = setRGB(217, 240, 252)
        colorsGr(2) = setRGB(190, 230, 253): colorsGr(3) = setRGB(167, 217, 245)
        colorsGr(4) = setRGB(249, 253, 254): colorsGr(5) = setRGB(245, 251, 254)
        colorsGr(6) = setRGB(239, 249, 254): colorsGr(7) = setRGB(233, 246, 253)
        colors(20) = RGB(232, 245, 252)
        colors(0) = RGB(60, 127, 177): colors(1) = RGB(250, 253, 254)
        colors(2) = RGB(102, 147, 181): colors(3) = RGB(66, 131, 179)
        colors(4) = RGB(88, 138, 175): colors(5) = RGB(105, 157, 195)
        colors(6) = RGB(223, 235, 243): colors(7) = RGB(62, 128, 178)
        colors(8) = RGB(225, 237, 244): colors(9) = RGB(247, 251, 254)
        colors(10) = RGB(62, 128, 178): colors(11) = RGB(249, 253, 254)
        colors(12) = RGB(87, 138, 175): colors(13) = RGB(62, 128, 178)
        colors(14) = RGB(90, 139, 176): colors(15) = RGB(100, 155, 195)
        colors(16) = RGB(212, 231, 243): colors(17) = RGB(62, 129, 178)
        colors(18) = RGB(211, 230, 242): colors(19) = RGB(220, 240, 251)

    Case btsDown
        colorsGr(0) = setRGB(228, 243, 251): colorsGr(1) = setRGB(196, 229, 246)
        colorsGr(2) = setRGB(152, 209, 239): colorsGr(3) = setRGB(104, 179, 219)
        colorsGr(4) = setRGB(176, 195, 206): colorsGr(5) = setRGB(154, 186, 203)
        colorsGr(6) = setRGB(120, 170, 197): colorsGr(7) = setRGB(85, 146, 181)
        colors(0) = RGB(44, 98, 139): colors(1) = RGB(158, 176, 186)
        colors(2) = RGB(57, 89, 113): colors(3) = RGB(46, 96, 134)
        colors(4) = RGB(51, 86, 113): colors(5) = RGB(58, 104, 138)
        colors(6) = RGB(134, 157, 171): colors(7) = RGB(44, 95, 134)
        colors(8) = RGB(134, 157, 171): colors(9) = RGB(186, 203, 213)
        colors(10) = RGB(44, 98, 139): colors(11) = RGB(170, 188, 199)
        colors(12) = RGB(75, 117, 147): colors(13) = RGB(46, 100, 140)
        colors(14) = RGB(71, 110, 140): colors(15) = RGB(47, 100, 138)
        colors(16) = RGB(99, 172, 211): colors(17) = RGB(45, 97, 136)
        colors(18) = RGB(85, 146, 181): colors(19) = RGB(109, 181, 220)
        colors(20) = RGB(104, 179, 219)

    Case btsDisabled
        colorsGr(0) = setRGB(244, 244, 244): colorsGr(1) = setRGB(244, 244, 244)
        colorsGr(2) = setRGB(244, 244, 244): colorsGr(3) = setRGB(244, 244, 244)
        colorsGr(4) = setRGB(252, 252, 252): colorsGr(5) = setRGB(252, 252, 252)
        colorsGr(6) = setRGB(252, 252, 252): colorsGr(7) = setRGB(252, 252, 252)
        colors(0) = RGB(173, 178, 181): colors(1) = RGB(252, 252, 252)
        colors(2) = RGB(183, 187, 190): colors(3) = RGB(175, 180, 183)
        colors(4) = RGB(177, 181, 184): colors(5) = RGB(192, 195, 198)
        colors(6) = RGB(241, 242, 242): colors(7) = RGB(174, 179, 182)
        colors(8) = RGB(242, 242, 243): colors(9) = RGB(251, 251, 251)
        colors(10) = RGB(173, 178, 181): colors(11) = RGB(252, 252, 252)
        colors(12) = RGB(176, 181, 184): colors(13) = RGB(174, 179, 182)
        colors(14) = RGB(177, 182, 185): colors(15) = RGB(192, 195, 198)
        colors(16) = RGB(243, 243, 244): colors(17) = RGB(174, 179, 182)
        colors(18) = RGB(242, 242, 243): colors(19) = RGB(251, 251, 251)
        colors(20) = RGB(252, 252, 252)

    Case btsFocus
        colorsGr(0) = setRGB(236, 245, 250): colorsGr(1) = setRGB(230, 242, 248)
        colorsGr(2) = setRGB(206, 231, 244): colorsGr(3) = setRGB(187, 219, 237)
        colorsGr(4) = setRGB(52, 213, 254): colorsGr(5) = setRGB(51, 212, 253)
        colorsGr(6) = setRGB(56, 210, 252): colorsGr(7) = setRGB(42, 207, 250)
        colors(0) = RGB(60, 127, 177): colors(1) = RGB(52, 213, 254)
        colors(2) = RGB(160, 190, 212): colors(3) = RGB(86, 143, 186)
        colors(4) = RGB(158, 188, 211): colors(5) = RGB(50, 146, 195)
        colors(6) = RGB(44, 199, 243): colors(7) = RGB(83, 142, 185)
        colors(8) = RGB(45, 200, 244): colors(9) = RGB(107, 223, 253)
        colors(10) = RGB(60, 127, 177): colors(11) = RGB(52, 213, 254)
        colors(12) = RGB(155, 187, 210): colors(13) = RGB(81, 140, 184)
        colors(14) = RGB(158, 188, 211): colors(15) = RGB(50, 145, 194)
        colors(16) = RGB(37, 197, 242): colors(17) = RGB(83, 142, 185)
        colors(18) = RGB(37, 196, 241): colors(19) = RGB(84, 210, 247)
        colors(20) = RGB(41, 206, 250)

    End Select

    With UserControl
        .Cls

        drawGradient 2, 2, .ScaleWidth - 2, .ScaleHeight / 2.2, _
                     colorsGr(0), colorsGr(1), .HDC

        drawGradient 2, .ScaleHeight / 2.2, .ScaleWidth - 2, .ScaleHeight - 2, _
                     colorsGr(2), colorsGr(3), .HDC

        drawGradient 1, 3, 2, .ScaleHeight / 2.2, _
                     colorsGr(4), colorsGr(5), .HDC

        drawGradient .ScaleWidth - 2, 3, .ScaleWidth - 1, .ScaleHeight / 2.2, _
                     colorsGr(4), colorsGr(5), .HDC

        drawGradient 1, .ScaleHeight / 2.2, 2, .ScaleHeight - 3, _
                     colorsGr(6), colorsGr(7), .HDC

        drawGradient .ScaleWidth - 1, .ScaleHeight / 2.2, .ScaleWidth - 2, .ScaleHeight - 3, _
                     colorsGr(6), colorsGr(7), .HDC

        If State = btsDown Then
            UserControl.Line (3, 1)-(6, 1), RGB(151, 170, 180), B
            UserControl.Line (3, 2)-(6, 2), RGB(220, 237, 246), B
            UserControl.Line (7, 1)-(10, 1), RGB(157, 175, 185), B

            UserControl.Line (.ScaleWidth - 4, 1)-(.ScaleWidth - 7, 1), RGB(151, 170, 180), B
            UserControl.Line (.ScaleWidth - 4, 2)-(.ScaleWidth - 7, 2), RGB(220, 237, 246), B
            UserControl.Line (.ScaleWidth - 8, 1)-(.ScaleWidth - 11, 1), RGB(157, 175, 185), B

            UserControl.Line (3, .ScaleHeight - 2)-(6, .ScaleHeight - 2), RGB(99, 172, 211), B
            UserControl.Line (3, .ScaleHeight - 3)-(6, .ScaleHeight - 3), RGB(109, 181, 220), B
            UserControl.Line (7, .ScaleHeight - 2)-(10, .ScaleHeight - 2), RGB(104, 178, 218), B

            UserControl.PSet (2, 3), RGB(220, 237, 245)
            UserControl.PSet (.ScaleWidth - 3, 3), RGB(220, 237, 245)
        End If

        If State = btsFocus Then
            UserControl.Line (3, 2)-(.ScaleWidth - 4, 2), RGB(166, 233, 251), B
            UserControl.Line (3, .ScaleHeight - 3)-(.ScaleWidth - 4, .ScaleHeight - 3), RGB(132, 213, 239), B

            drawGradient 2, 3, 3, .ScaleHeight / 2.2, setRGB(165, 232, 251), setRGB(161, 230, 249), .HDC
            drawGradient .ScaleWidth - 2, 3, .ScaleWidth - 3, .ScaleHeight / 2.2, setRGB(165, 232, 251), setRGB(161, 230, 249), .HDC
            drawGradient 2, .ScaleHeight / 2.2, 3, .ScaleHeight - 3, setRGB(146, 222, 245), setRGB(133, 213, 240), .HDC
            drawGradient .ScaleWidth - 2, .ScaleHeight / 2.2, .ScaleWidth - 3, .ScaleHeight - 3, setRGB(146, 222, 245), setRGB(133, 213, 240), .HDC
        End If

        DrawCaption

        UserControl.Line (0, 0)-(.ScaleWidth - 1, .ScaleHeight - 1), colors(0), B
        UserControl.Line (3, 1)-(.ScaleWidth - 4, 1), colors(1), B
        UserControl.Line (3, .ScaleHeight - 2)-(.ScaleWidth - 4, .ScaleHeight - 2), colors(20), B

        UserControl.PSet (1, 0), colors(2)
        UserControl.PSet (2, 0), colors(3)
        UserControl.PSet (0, 1), colors(4)
        UserControl.PSet (1, 1), colors(5)
        UserControl.PSet (2, 1), colors(6)
        UserControl.PSet (0, 2), colors(7)
        UserControl.PSet (1, 2), colors(8)
        UserControl.PSet (2, 2), colors(9)
        UserControl.PSet (0, 3), colors(10)
        UserControl.PSet (1, 3), colors(11)

        UserControl.PSet (.ScaleWidth - 2, 0), colors(2)
        UserControl.PSet (.ScaleWidth - 3, 0), colors(3)
        UserControl.PSet (.ScaleWidth - 1, 1), colors(4)
        UserControl.PSet (.ScaleWidth - 2, 1), colors(5)
        UserControl.PSet (.ScaleWidth - 3, 1), colors(6)
        UserControl.PSet (.ScaleWidth - 1, 2), colors(7)
        UserControl.PSet (.ScaleWidth - 2, 2), colors(8)
        UserControl.PSet (.ScaleWidth - 3, 2), colors(9)
        UserControl.PSet (.ScaleWidth - 1, 3), colors(10)
        UserControl.PSet (.ScaleWidth - 2, 3), colors(11)

        UserControl.PSet (1, .ScaleHeight - 1), colors(12)
        UserControl.PSet (2, .ScaleHeight - 1), colors(13)
        UserControl.PSet (0, .ScaleHeight - 2), colors(14)
        UserControl.PSet (1, .ScaleHeight - 2), colors(15)
        UserControl.PSet (2, .ScaleHeight - 2), colors(16)
        UserControl.PSet (0, .ScaleHeight - 3), colors(17)
        UserControl.PSet (1, .ScaleHeight - 3), colors(18)
        UserControl.PSet (2, .ScaleHeight - 3), colors(19)

        UserControl.PSet (.ScaleWidth - 2, .ScaleHeight - 1), colors(12)
        UserControl.PSet (.ScaleWidth - 3, .ScaleHeight - 1), colors(13)
        UserControl.PSet (.ScaleWidth - 1, .ScaleHeight - 2), colors(14)
        UserControl.PSet (.ScaleWidth - 2, .ScaleHeight - 2), colors(15)
        UserControl.PSet (.ScaleWidth - 3, .ScaleHeight - 2), colors(16)
        UserControl.PSet (.ScaleWidth - 1, .ScaleHeight - 3), colors(17)
        UserControl.PSet (.ScaleWidth - 2, .ScaleHeight - 3), colors(18)
        UserControl.PSet (.ScaleWidth - 3, .ScaleHeight - 3), colors(19)

    End With

    Erase colors
    Erase colorsGr

End Sub

Private Function drawGradient(X1 As Integer, Y1 As Integer, X2 As Integer, _
                              Y2 As Integer, color1 As RGBCOLOR, _
                              color2 As RGBCOLOR, HDC As Long)

    Dim Vertex(0 To 1) As TRIVERTEX

    Dim RECT As GRADIENT_RECT

    With Vertex(0)
        .X = X1: .Y = Y1: .Alpha = 0
        .red = color1.red: .green = color1.green: .blue = color1.blue
    End With

    With Vertex(1)
        .X = X2: .Y = Y2: .Alpha = 0
        .red = color2.red: .green = color2.green: .blue = color2.blue
    End With

    With RECT
        .UpperLeft = 0: .LowerRight = 1
    End With

    GradientFillRect HDC, Vertex(0), 2, RECT, 1, GRADIENT_FILL_RECT_V

    Erase Vertex

End Function

Private Function setRGB(red As Integer, green As Integer, blue As Integer) As RGBCOLOR
    Dim erg As RGBCOLOR
    With erg
        .red = CCol(red)
        .green = CCol(green)
        .blue = CCol(blue)
    End With
    setRGB = erg
End Function

Private Sub MakeRegion()

    Dim rgn1 As Long
    Dim rgn2 As Long

    With UserControl
        DeleteObject rgn1
        rgn1 = CreateRectRgn(0, 0, .ScaleWidth, .ScaleHeight)
        rgn2 = CreateRectRgn(0, 0, 1, 1)
        CombineRgn rgn1, rgn1, rgn2, RGN_DIFF
        DeleteObject rgn2
        rgn2 = CreateRectRgn(0, .ScaleHeight - 1, 1, .ScaleHeight)
        CombineRgn rgn1, rgn1, rgn2, RGN_DIFF
        DeleteObject rgn2
        rgn2 = CreateRectRgn(.ScaleWidth - 1, 0, .ScaleWidth, 1)
        CombineRgn rgn1, rgn1, rgn2, RGN_DIFF
        DeleteObject rgn2
        rgn2 = CreateRectRgn(.ScaleWidth - 1, .ScaleHeight - 1, .ScaleWidth, .ScaleHeight)
        CombineRgn rgn1, rgn1, rgn2, RGN_DIFF
        DeleteObject rgn2
        SetWindowRgn .hWnd, rgn1, True
    End With

End Sub

Private Function isMouseOver() As Boolean
    Dim PT As POINTAPI
    GetCursorPos PT
    isMouseOver = (WindowFromPoint(PT.X, PT.Y) = UserControl.hWnd)
End Function

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Caption", dCaption
        .WriteProperty "ForeColor", dForeColor
        .WriteProperty "Font", UserControl.Font
        .WriteProperty "Picture", dPicture
        .WriteProperty "Pictures", dPictures
        .WriteProperty "UseMaskColor", dUseMaskColor
        .WriteProperty "MaskColor", dMaskColor
        .WriteProperty "Enabled", dEnabled
        .WriteProperty "NoBackground", dNoBackground
        .WriteProperty "BackColor", UserControl.BackColor
        .WriteProperty "PictureOffset", dPictureOffset
    End With
End Sub

Private Sub DrawCaption()

    Dim tmpH As Long
    Dim r As RECT
    Dim ax As Integer

    With r
        .Left = 0
        .Top = 0
        .Right = UserControl.ScaleWidth
        .Bottom = UserControl.ScaleHeight
    End With

    With UserControl
        SetTextColor .HDC, IIf(dEnabled, dForeColor, RGB(131, 131, 131))
        tmpH = DrawText(.HDC, dCaption, Len(dCaption), _
                        r, DT_CALCRECT Or DT_CENTER Or DT_WORDBREAK)

        ax = drawPicture(r.Right)
        r.Left = 3
        If dHDC Then r.Left = 8 + dPictureW / dPictures - ax
        r.Top = (.ScaleHeight - tmpH) / 2
        r.Right = .ScaleWidth - 3
        r.Bottom = (.ScaleHeight + tmpH) / 2
        DrawText .HDC, dCaption, Len(dCaption), r, IIf(ax > 0, 0, DT_CENTER) Or DT_WORDBREAK
    End With

End Sub

Private Function drawPicture(textWidth As Long) As Integer

    Dim r As RECT
    Dim dPic As Integer
    Dim dx, dy As Integer
    Dim ax, ay, ar, ab As Integer
    Dim db As Integer

    If dHDC Then
        If Not dCaption = "" Then
            dx = (UserControl.ScaleWidth - textWidth - dPictureW / dPictures - 5) / 2
        Else
            dx = (UserControl.ScaleWidth - dPictureW / dPictures) / 2
        End If
        dy = (UserControl.ScaleHeight - dPictureH) / 2

        If mDown Then
            dx = dx + dPictureOffset
            dy = dy + dPictureOffset
        End If

        ar = 0
        If dx + dPictureW / dPictures > UserControl.ScaleWidth - 3 Then
            ar = UserControl.ScaleWidth - 3 - dx - dPictureW / dPictures
        End If

        ax = 0: If dx < 3 Then ax = 3 - dx: dx = 3

        ab = 0
        If dy + dPictureH > UserControl.ScaleHeight - 3 Then
            ab = UserControl.ScaleHeight - 3 - dy - dPictureH
        End If

        ay = 0: If dy < 3 Then ay = 3 - dy: dy = 3

        dPic = 0
        If dPictures > 1 And Not dEnabled Then dPic = 1
        If dPictures = 3 And dEnabled And mDown Then dPic = 2

        db = 0
        If dCaption = "" Then db = 8

        With UserControl
            '            If ax < dPictureW / dPictures - db And ay < dPictureH - db Then
            If dUseMaskColor Then
                TransparentBlt .HDC, dx, dy, dPictureW / dPictures - ax + ar, _
                               dPictureH - ay + ab, dHDC, dPictureW / dPictures * dPic + ax, ay, _
                               dPictureW / dPictures - ax + ar, dPictureH - ay + ab, dMaskColor
            Else
                BitBlt .HDC, dx, dy, dPictureW / dPictures - ax + ar, _
                       dPictureH - ay + ab, dHDC, dPictureW / dPictures * dPic + ax, _
                       ay, vbSrcCopy
            End If
            '           End If
        End With
    End If

    drawPicture = ax

End Function
