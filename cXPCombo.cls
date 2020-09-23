VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cXPCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'This is not all my code
'Most of it (The hard part) is by the legend himself
'Steve McMahon at www.vbaccelerator.com
'so i will not take credit for the sub classing or message hooks
'The variables to some of the code have been re-named
'only to suit my variable naming convention

Implements ISubclass

Private Enum eDrawState
    eDrawNormal = 0
    eDrawRaised = 1
    eDrawPressed = 2
    eDrawDisabled = 3
End Enum

Private Const WM_COMMAND = &H111
Private Const WM_PAINT = &HF
Private Const WM_SETFOCUS = &H7
Private Const WM_KILLFOCUS = &H8
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_TIMER = &H113
Private Const WM_ENABLE = &HA

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetPixel Lib "gdi32" Alias "SetPixelV" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Const CLR_INVALID = -1

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CXHSCROLL = 21
Private Const SM_CXHTHUMB = 10
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Const PS_SOLID = 0
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const CBS_DROPDOWN = &H2&
Private Const CBS_DROPDOWNLIST = &H3&
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, lpsz2 As Any) As Long
Private Const CBN_CLOSEUP = 8
Private Const CB_GETDROPPEDSTATE = &H157
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Const GW_CHILD = 5
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long

'The colors and thumb pictures.
Private oFrameColorHot          As OLE_COLOR
Private oFrameColorNormal       As OLE_COLOR
Private oFrameColorDiss         As OLE_COLOR
Private oBackColor              As OLE_COLOR
Private oArrowColor             As OLE_COLOR
Private oArrowColorDiss         As OLE_COLOR
Private pNormalThumb            As StdPicture
Private pHotThumb               As StdPicture
Private pDropedThumb            As StdPicture
Private pDissThumb              As StdPicture
'//
'Just in case you dont like the default arrow include your own
'in the custom thumb pictures.
Private bShowArrow              As Boolean
'//
Private m_hpalHalftone As Long
Private cbHwnd As Long
Private cbEditHwnd As Long
Private cbParentHwnd As Long
Private bSubClass As Boolean
Private bMouseOver As Boolean
Private bMouseDown As Boolean
Private bGotFocus As Boolean
Private bDisabled  As Boolean

Public Sub MakeXP(ByRef cb As ComboBox)
Dim ihWnd As Long
    'www.vbaccelerator.com
    pRelease
    '//
    
    'Get a Handle and find the parents handle
    ihWnd = cb.hwnd
    cbParentHwnd = GetParent(ihWnd)
    '//
    
    'www.vbaccelerator.com
    pAttach ihWnd
    '//
End Sub

Private Sub pAttach(ByRef hWndA As Long)
'www.vbaccelerator.com
Dim lStyle As Long
    
    cbHwnd = hWndA
    If (cbHwnd <> 0) Then
        lStyle = GetWindowLong(cbHwnd, GWL_STYLE)
        If (lStyle And CBS_DROPDOWN) = CBS_DROPDOWN Then
            cbEditHwnd = GetWindow(cbHwnd, GW_CHILD)
        End If
        AttachMessage Me, cbHwnd, WM_PAINT
        AttachMessage Me, cbHwnd, WM_SETFOCUS
        AttachMessage Me, cbHwnd, WM_KILLFOCUS
        AttachMessage Me, cbHwnd, WM_MOUSEMOVE
        AttachMessage Me, cbHwnd, WM_LBUTTONDOWN
        AttachMessage Me, cbHwnd, WM_LBUTTONUP
        AttachMessage Me, cbHwnd, WM_TIMER
        AttachMessage Me, cbHwnd, WM_ENABLE
        If (cbEditHwnd <> 0) Then
            AttachMessage Me, cbEditHwnd, WM_SETFOCUS
            AttachMessage Me, cbEditHwnd, WM_KILLFOCUS
            AttachMessage Me, cbEditHwnd, WM_MOUSEMOVE
        End If
        AttachMessage Me, cbParentHwnd, WM_COMMAND
        bSubClass = True
    End If
'//
End Sub

Private Sub pRelease()
'www.vbaccelerator.com
    If bSubClass Then
        DetachMessage Me, cbHwnd, WM_PAINT
        DetachMessage Me, cbHwnd, WM_SETFOCUS
        DetachMessage Me, cbHwnd, WM_KILLFOCUS
        DetachMessage Me, cbHwnd, WM_MOUSEMOVE
        DetachMessage Me, cbHwnd, WM_LBUTTONDOWN
        DetachMessage Me, cbHwnd, WM_LBUTTONUP
        DetachMessage Me, cbHwnd, WM_TIMER
        DetachMessage Me, cbHwnd, WM_ENABLE
        If (cbEditHwnd <> 0) Then
            DetachMessage Me, cbEditHwnd, WM_SETFOCUS
            DetachMessage Me, cbEditHwnd, WM_KILLFOCUS
            DetachMessage Me, cbEditHwnd, WM_MOUSEMOVE
        End If
        DetachMessage Me, cbParentHwnd, WM_COMMAND
    End If
    cbHwnd = 0
    cbEditHwnd = 0
    cbParentHwnd = 0
'//
End Sub

Private Sub OnTimer(ByVal bCheckMouse As Boolean)
'www.vbaccelerator.com
    Dim bOver As Boolean
    Dim rcItem As RECT
    Dim pt As POINTAPI
    Dim lhWnd As Long
    
    If bCheckMouse Then
        bOver = True
        GetCursorPos pt
        lhWnd = WindowFromPoint(pt.X, pt.Y)
        If lhWnd <> cbHwnd And lhWnd <> cbEditHwnd Then
            bOver = False
        End If
    End If
    
    If Not bOver Then
        KillTimer cbHwnd, 1
        bMouseOver = False
        DrawComboState
    End If
'//
End Sub

Private Sub Class_Initialize()
Dim ScreenDC As Long
    ScreenDC = GetDC(0&)
    m_hpalHalftone = CreateHalftonePalette(ScreenDC)
    ReleaseDC 0&, ScreenDC
End Sub

Private Sub Class_Terminate()
'www.vbaccelerator.com
    pRelease
'//
    DeleteObject m_hpalHalftone
    Set pNormalThumb = Nothing
    Set pHotThumb = Nothing
    Set pDropedThumb = Nothing
    Set pDissThumb = Nothing
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer6.EMsgResponse)
    '
End Property

Private Property Get ISubclass_MsgResponse() As SSubTimer6.EMsgResponse
    If (CurrentMessage = WM_TIMER) Then
        ISubclass_MsgResponse = emrPostProcess
    Else
        ISubclass_MsgResponse = emrPreprocess
    End If
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'www.vbaccelerator.com
    Select Case iMsg
    Case WM_COMMAND
        If (cbHwnd = lParam) Then
            Select Case wParam \ &H10000
            Case CBN_CLOSEUP
                DrawComboState
            End Select
        End If
    Case WM_PAINT
        DrawComboState
    Case WM_ENABLE
        bDisabled = (IsWindowEnabled(cbHwnd) = 0)
        DrawComboState
    Case WM_SETFOCUS
        bGotFocus = True
        DrawComboState
    Case WM_KILLFOCUS
        bGotFocus = False
        DrawComboState
    Case WM_MOUSEMOVE
        If Not bMouseOver Then
            bMouseOver = True
            DrawComboState
            SetTimer cbHwnd, 1, 10, 0
        End If
    Case WM_LBUTTONDOWN
        bMouseDown = True
        DrawComboState
    Case WM_LBUTTONUP
        bMouseDown = False
        DrawComboState
    Case WM_TIMER
        OnTimer True
    End Select
'//
End Function

Private Sub DrawComboState()
Dim DrawStyle As eDrawState
        If bDisabled Then
            DrawStyle = eDrawDisabled
        ElseIf SendMessageLong(cbHwnd, CB_GETDROPPEDSTATE, 0, 0) <> 0 Then
            DrawStyle = eDrawPressed
        ElseIf bGotFocus Or bMouseOver Then
            DrawStyle = eDrawRaised
        Else
            DrawStyle = eDrawNormal
        End If
        DrawCombo DrawStyle
End Sub

Private Sub DrawCombo(ByVal DrawStyle As eDrawState)
Dim rc                      As RECT
Dim DestDC                  As Long
Dim i                       As Long
Dim iCl(0 To 1)             As Single
Dim ScreenDC                As Long
Dim ThumbDC                 As Long
Dim OldMemSrc               As Long
Dim OldPallet               As Long
Dim hPal                    As Long
Dim PalHalftone             As Long
Dim pTempPic                As Picture
Dim oNewArrowColor          As OLE_COLOR

    'Get the size and hDC of the ComboBox
    GetClientRect cbHwnd, rc
    DestDC = GetDC(cbHwnd)
    '//
    
    'Draw the frame portion of the ComboBox
    Select Case DrawStyle
        Case eDrawDisabled
            'Disabled
            DrawASquare DestDC, rc, FrameColorDiss
            InflateRect rc, -1, -1
            DrawASquare DestDC, rc, CboBackColor
            oNewArrowColor = ArrowDisabledColor
        Case eDrawNormal
            'As is
            DrawASquare DestDC, rc, FrameColorNormal
            InflateRect rc, -1, -1
            DrawASquare DestDC, rc, CboBackColor
            oNewArrowColor = ArrowEnabledColor
        Case Else
            'Hot/GotFocus
            DrawASquare DestDC, rc, FrameColorHot
            InflateRect rc, -1, -1
            DrawASquare DestDC, rc, CboBackColor
            oNewArrowColor = ArrowEnabledColor
    End Select
    '//
    
    InflateRect rc, -1, -1

    'Get the size of the Thumb
    rc.Left = rc.Right - GetSystemMetrics(SM_CXHTHUMB)
    i = GetSystemMetrics(SM_CXHTHUMB)
    '//

    'Set our Thumb pictures
    Select Case DrawStyle
        Case eDrawDisabled
            If pDissThumb Is Nothing Then
                Set pDissThumb = LoadResPicture(1400, vbResBitmap)
            End If
            Set pTempPic = pDissThumb
        Case eDrawNormal
            If pNormalThumb Is Nothing Then
                Set pNormalThumb = LoadResPicture(1100, vbResBitmap)
            End If
            Set pTempPic = pNormalThumb
        Case eDrawRaised
            If pHotThumb Is Nothing Then
                Set pHotThumb = LoadResPicture(1200, vbResBitmap)
            End If
            Set pTempPic = pHotThumb
        Case eDrawPressed
            If pDropedThumb Is Nothing Then
                Set pDropedThumb = LoadResPicture(1300, vbResBitmap)
            End If
            Set pTempPic = pDropedThumb
    End Select
    '//
    
    'Validate palette
    If hPal = 0 Then
        hPal = m_hpalHalftone
    End If
    '//
    
    'Select passed picture into a hDC
    ScreenDC = GetDC(0&)
    ThumbDC = CreateCompatibleDC(ScreenDC)
    OldMemSrc = SelectObject(ThumbDC, pTempPic.Handle)
    OldPallet = SelectPalette(ThumbDC, hPal, True)
    RealizePalette ThumbDC
    '//
    
    'Draw the bitmap
    StretchBlt DestDC, rc.Left, rc.Top, i, rc.Bottom - rc.Top, ThumbDC, 0, 0, 14, 17, vbSrcCopy
    '//
    
    'Clean up
    SelectObject ThumbDC, OldMemSrc
    SelectPalette ThumbDC, OldPallet, True
    RealizePalette ThumbDC
    DeleteDC ThumbDC
    ReleaseDC 0&, ScreenDC
    Set pTempPic = Nothing
    '//
    
    'Find the center of the Thumb
    iCl(0) = rc.Left + Format(i / 2, "#")
    iCl(1) = Format(((rc.Bottom - rc.Top) + 3) / 2, "#")
    '//
    
    'Now draw the arrow if the user wants one.
    If bShowArrow = True Then
        For i = 0 To 2
            DrawALine DestDC, iCl(0), iCl(1) + i, iCl(0) - 4, (iCl(1) - 4) + i, oNewArrowColor
        Next i
        For i = 0 To 2
            DrawALine DestDC, iCl(0), iCl(1) + i, iCl(0) + 4, (iCl(1) - 4) + i, oNewArrowColor
        Next i
        DrawADot DestDC, iCl(0) - 4, iCl(1) - 2, oNewArrowColor
        DrawADot DestDC, iCl(0) + 4, iCl(1) - 2, oNewArrowColor
    End If
    ' //
    
    'Clean up
    Erase iCl
    DeleteDC DestDC
    '//
End Sub

Private Sub DrawASquare(DestDC As Long, rc As RECT, oColor As OLE_COLOR, Optional bFillRect As Boolean)
Dim iBrush As Long
Dim i(0 To 3) As Long
    iBrush = CreateSolidBrush(oColor)
    If bFillRect = True Then
        FillRect DestDC, rc, iBrush
    Else
        FrameRect DestDC, rc, iBrush
    End If
    DeleteObject iBrush
End Sub

Private Sub DrawALine(DestDC As Long, X As Single, Y As Single, X1 As Single, Y1 As Single, oColor As OLE_COLOR, Optional iWidth As Long = 1)
Dim pt As POINTAPI
Dim iPen As Long
Dim iPen1 As Long

    iPen = CreatePen(PS_SOLID, iWidth, oColor)
    iPen1 = SelectObject(DestDC, iPen)
    
    MoveToEx DestDC, X, Y, pt
    LineTo DestDC, X1, Y1

    SelectObject DestDC, iPen1
    DeleteObject iPen
End Sub

Private Sub DrawADot(DestDC As Long, X As Long, Y As Long, oColor As OLE_COLOR)
    SetPixel DestDC, X, Y, oColor
End Sub

Private Function TransColor(ByVal oClr As OLE_COLOR) As OLE_COLOR
Dim iColor As Long
Dim hPal As Long
    If OleTranslateColor(oClr, hPal, iColor) Then
        TransColor = CLR_INVALID
    Else
        TransColor = iColor
    End If
End Function

'From here down are all the properties.

Public Property Get FrameColorNormal() As OLE_COLOR
    If oFrameColorNormal = -1 Then
        oFrameColorNormal = RGB(48, 72, 144)
    End If
    FrameColorNormal = TransColor(oFrameColorNormal)
End Property

Public Property Let FrameColorNormal(ByVal NewFrameColorNormal As OLE_COLOR)
    oFrameColorNormal = TransColor(NewFrameColorNormal)
End Property

Public Property Get FrameColorHot() As OLE_COLOR
    If oFrameColorHot = -1 Then
        oFrameColorHot = RGB(176, 208, 240)
    End If
    FrameColorHot = TransColor(oFrameColorHot)
End Property

Public Property Let FrameColorHot(ByVal NewFrameColorHot As OLE_COLOR)
    oFrameColorHot = TransColor(NewFrameColorHot)
End Property

Public Property Get CboBackColor() As OLE_COLOR
    CboBackColor = TransColor(oBackColor)
End Property

Public Property Let CboBackColor(ByVal NewBackColor As OLE_COLOR)
    oBackColor = TransColor(NewBackColor)
End Property

Public Property Get FrameColorDiss() As OLE_COLOR
    If oFrameColorDiss = -1 Then
        oFrameColorDiss = RGB(192, 168, 136)
    End If
    FrameColorDiss = TransColor(oFrameColorDiss)
End Property

Public Property Let FrameColorDiss(ByVal NewDissColor As OLE_COLOR)
    oFrameColorDiss = TransColor(NewDissColor)
End Property

Public Property Get ArrowEnabledColor() As OLE_COLOR
    If oArrowColor = -1 Then
        oArrowColor = RGB(80, 104, 128)
    End If
    ArrowEnabledColor = TransColor(oArrowColor)
End Property

Public Property Let ArrowEnabledColor(ByVal NewArrowColor As OLE_COLOR)
    oArrowColor = TransColor(NewArrowColor)
End Property

Public Property Get ArrowDisabledColor() As OLE_COLOR
    If oArrowColorDiss = -1 Then
        oArrowColorDiss = RGB(192, 168, 136)
    End If
    ArrowDisabledColor = TransColor(oArrowColorDiss)
End Property

Public Property Let ArrowDisabledColor(ByVal NewArrowColorDiss As OLE_COLOR)
    oArrowColorDiss = TransColor(NewArrowColorDiss)
End Property

Public Property Get PictureNormalThumb() As StdPicture
    Set PictureNormalThumb = pNormalThumb
End Property

Public Property Set PictureNormalThumb(ByVal NewNormalThumb As StdPicture)
    Set pNormalThumb = NewNormalThumb
End Property

Public Property Get PictureHotThumb() As StdPicture
    Set PictureHotThumb = pHotThumb
End Property

Public Property Set PictureHotThumb(ByVal NewHotThumb As StdPicture)
    Set pHotThumb = NewHotThumb
End Property

Public Property Get PictureDropedThumb() As StdPicture
    Set PictureDropedThumb = pDropedThumb
End Property

Public Property Set PictureDropedThumb(ByVal NewDropedThumb As StdPicture)
    Set pDropedThumb = NewDropedThumb
End Property

Public Property Get PictureDissThumb() As StdPicture
    Set PictureDissThumb = pDissThumb
End Property

Public Property Set PictureDissThumb(ByVal NewDissThumb As StdPicture)
    Set pDissThumb = NewDissThumb
End Property

Public Property Get ShowArrow() As Boolean
    ShowArrow = bShowArrow
End Property

Public Property Let ShowArrow(ByVal NewArrow As Boolean)
    bShowArrow = NewArrow
End Property
