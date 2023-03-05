VERSION 5.00
Begin VB.UserControl ColorSelector 
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   2880
   ScaleWidth      =   3840
   ToolboxBitmap   =   "ctlColorSelector.ctx":0000
   Begin VB.PictureBox picSlider 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2124
      Left            =   2904
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   1
      Top             =   120
      Width           =   132
   End
   Begin VB.Timer tmrFixSize 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   984
      Top             =   2400
   End
   Begin VB.PictureBox picColorSystem 
      BorderStyle     =   0  'None
      Height          =   228
      Left            =   72
      ScaleHeight     =   228
      ScaleWidth      =   780
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2256
      Width           =   780
      Begin ColorControls.LabelW lblMode 
         Height          =   180
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   408
         _ExtentX        =   720
         _ExtentY        =   318
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   7.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   0
         Caption         =   "Mode:"
         AutoSize        =   -1  'True
      End
   End
   Begin ColorControls.ComboBoxW cboColorSystem 
      Height          =   336
      Left            =   72
      TabIndex        =   6
      Top             =   2496
      Width           =   700
      _ExtentX        =   0
      _ExtentY        =   0
      Style           =   2
   End
   Begin VB.Timer tmrDraw 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1392
      Top             =   2400
   End
   Begin ColorControls.ComboBoxW cboSliderParameter 
      Height          =   336
      Left            =   2328
      TabIndex        =   3
      Top             =   2328
      Visible         =   0   'False
      Width           =   900
      _ExtentX        =   0
      _ExtentY        =   0
      Style           =   2
   End
   Begin ColorControls.CheckBoxW chkFixedPalette 
      Height          =   348
      Left            =   96
      TabIndex        =   7
      Top             =   72
      Visible         =   0   'False
      Width           =   600
      _ExtentX        =   0
      _ExtentY        =   0
      Value           =   1
      Caption         =   "Fixed"
      Style           =   1
   End
   Begin VB.PictureBox picAux 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   720
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   27
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   -48
      Visible         =   0   'False
      Width           =   324
   End
   Begin VB.PictureBox picShades 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2124
      Left            =   2496
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   27
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   144
      Width           =   324
   End
   Begin VB.PictureBox picPalette 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2124
      Left            =   144
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   187
      TabIndex        =   0
      Top             =   168
      Width           =   2244
      Begin VB.Line linPointer 
         BorderWidth     =   3
         DrawMode        =   5  'Not Copy Pen
         Index           =   0
         X1              =   10
         X2              =   24
         Y1              =   74
         Y2              =   74
      End
      Begin VB.Line linPointer 
         BorderWidth     =   3
         DrawMode        =   5  'Not Copy Pen
         Index           =   1
         X1              =   50
         X2              =   64
         Y1              =   74
         Y2              =   74
      End
      Begin VB.Line linPointer 
         BorderWidth     =   3
         DrawMode        =   5  'Not Copy Pen
         Index           =   3
         X1              =   36
         X2              =   36
         Y1              =   82
         Y2              =   96
      End
      Begin VB.Line linPointer 
         BorderWidth     =   3
         DrawMode        =   5  'Not Copy Pen
         Index           =   2
         X1              =   36
         X2              =   36
         Y1              =   54
         Y2              =   68
      End
   End
   Begin ColorControls.ToolTipHandler ToolTipHandler1 
      Left            =   3240
      Top             =   1620
      _ExtentX        =   720
      _ExtentY        =   720
   End
End
Attribute VB_Name = "ColorSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Control to select colors from a palette."
Option Explicit

Implements IBSSubclass

Private Const cSUBCLASS_IN_IDE As Boolean = True

Public Event Change()
Attribute Change.VB_Description = "Occurs when the selected color changes."
Public Event SliderParameterChange()
Attribute SliderParameterChange.VB_Description = "Ocurs when the parameter that is controlled by the slider cotnrol changes."
Public Event FixedPaletteChange()
Attribute FixedPaletteChange.VB_Description = "Occurs when the user changes the palette from fixed/dynamic."
Public Event ColorSystemChange()
Attribute ColorSystemChange.VB_Description = "Ocurrs when the color system changes."
Public Event MouseWheelScroll(Axis As CDMouseWheelScrollConstants)
Attribute MouseWheelScroll.VB_Description = "Occurs when the user scrolls the control using the mouse wheel. The parameter indicates what is being scrolled."
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while the control has the focus."
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while the control has the focus.\n"
Public Event DblClickOnColor()
Attribute DblClickOnColor.VB_Description = "Occurs when the user performs a double click on a color."
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over the control."
Public Event GetLocalizedText(ByVal LanguageID As Long, ByVal SubLanguageID As Long, ByVal TextID As Long, ByRef Text As String)
Attribute GetLocalizedText.VB_Description = "Occurs when setting a text for the IU. It allows to customize captions and texts."

Private Enum CDCaptionsIDConstants
    cdCWCaptionHue ' Hue
    cdCWCaptionLum ' Lum
    cdCWCaptionSat ' Sat
    cdCWCaptionRed ' Red
    cdCWCaptionGreen ' Green
    cdCWCaptionBlue ' Blue
    cdCWCaptionVal ' Val.
    cdCWCaptionFixed ' Fixed
    cdCWCaptionFixedToolTipText ' Reflects color changes visually or not
    cdCWCaptionSliderParameterToolTipText ' Select parameter
    cdCWCaptionMode ' Mode
End Enum

Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_SYSCOLORCHANGE As Long = &H15
Private Const WM_THEMECHANGED As Long = &H31A

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer

Private Type RGBQuad
    R As Byte
    G As Byte
    B As Byte
    a As Byte
End Type

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

Private Declare Function Polygon Lib "gdi32" (ByVal HDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Sub ClipCursor Lib "user32" (lpRect As Any)
Private Declare Sub GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT)
Private Declare Sub ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI)
Private Declare Sub OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long)
Private Declare Sub ColorRGBToHLS Lib "shlwapi" (ByVal clrRGB As Long, ByRef pwHue As Integer, ByRef pwLuminance As Integer, ByRef pwSaturation As Integer)
Private Declare Function ColorHLSToRGB Lib "shlwapi" (ByVal wHue As Integer, ByVal wLuminance As Integer, ByVal wSaturation As Integer) As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFOHEADER, ByVal wUsage As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dX As Long, ByVal dY As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long) As Long

Private WithEvents mForm As Form
Attribute mForm.VB_VarHelpID = -1

Private Const DIB_RGB_COLORS    As Long = 0

Private Const Pi = 3.14159265358979

Private mCx As Long
Private mCy As Long
Private mDiameter As Long
Private mRadius As Long
Private mPaletteWidth As Long
Private mPaletteHeight As Long
Private mBMPiH As BITMAPINFOHEADER
Private mPixelsBytes() As Byte
Private mPixelsBytes2() As Byte
Private mBorderPixels() As Long
Private mBorderPixels_Alpha() As Byte
Private mPixelsAreInPalette() As Boolean
Private mPixelsAngleOrX() As Double
Private mPixelsRadiusOrY() As Double
Private mBytesStride As Long
Private mBytesCount As Long
Private mBMPHeight As Long
Private mBMPWidth As Long
Private mPureColor As Long
Private mSettingColor As Boolean
Private mChangingShade As Boolean
Private mChangingHue As Boolean
Private mChangingLuminance As Boolean
Private mChangingSaturation As Boolean
Private mSelectingColor As Boolean
Private mClickingPalette As Boolean
Private mPointerX As Single
Private mPointerY As Single
Private mUserControlShown As Boolean
Private mChangingColorSystemOrInitializing As Boolean
Private mCaptionLum As String
Private mCaptionVal As String
Private mAmbientUserMode As Boolean
Private mDrawEnabled As Boolean
Private mSettingSlider As Boolean
Private mRedraw As Boolean
Private mDrawPending As Boolean
Private mPaletteColorsStored As Boolean
Private mNewHeight As Long
Private mNewWidth As Long
Private mPropertiesAreSet As Boolean
Private mChangingParameter As Boolean
Private mRaiseEvents As Boolean
Private mInitialized As Boolean
Private mUserControlHwnd As Long
Private mRadialParameter As CDSliderParameterConstants ' this parameter increases its value from the center of the wheel to the periphery (or Y position in case of Box Style)
Private mAxialParameter As CDSliderParameterConstants ' this parameter changes its value with the different angles to the center of the wheel  (or X position in case of Box Style)
Private mParametersCaptions(5) As String
Private mSelectionFromOutside As Boolean
Private Const cNarrowPicShadesWidth As Long = 280
Private Const cWidePicShadesWidth As Long = 440
Private mPicShadesWidth  As Long
Private mMouseDown As Boolean
Private mStyleBox As Boolean
Private mSubclassed As Boolean
Private mFormHwnd As Long

Private Const cPropDefault_Color As Long = &H808080
Private Const cPropDefault_SliderOptionsAvailable As Long = cdSliderOptionsNone
Private Const cPropDefault_FixedPaletteControlVisible As Boolean = False
Private Const cPropDefault_ColorSystemControlVisible As Boolean = False
Private Const cPropDefault_FixedPalette As Boolean = True
Private Const cPropDefault_SliderParameter As Long = cdParameterLuminance
Private Const cPropDefault_ColorSystem As Long = cdColorSystemHSV
Private Const cPropDefault_BackColor As Long = vbButtonFace
Private Const cPropDefault_SliderWide As Long = cdYNAuto
Private Const cPropDefault_HideLabels As Boolean = False
Private Const cPropDefault_Style As Long = cdStyleWheel
Private Const cPropDefault_RoundedBoxes As Boolean = True
Private Const cPropDefault_SliderParameterComboWidth As Long = 900

Private mColor As Long
Private mSliderOptionsAvailable As CDSliderOptionsAvailableConstants
Private mFixedPaletteControlVisible As Boolean
Private mColorSystemControlVisible As Boolean
Private mFixedPalette As Boolean
Private mSliderParameter As CDSliderParameterConstants
Private mColorSystem As CDColorSystemConstants
Private mBackColor As Long
Private mSliderWide As CDYesNoAutoConstants
Private mHideLabels As Boolean
Private mStyle As CDStyleConstants
Private mRoundedBoxes As Boolean
Private mSliderParameterComboWidth As Long

Private mH As Double
Private mL As Double
Private mS As Double
Private mR As Long
Private mG As Long
Private mB As Long

Private mH_Max As Long
Private mL_Max As Long
Private mS_Max As Long
Private mH_Fixed As Double
Private mL_Fixed As Double
Private mS_Fixed As Double

' Slider control
Private mSliderMin As Long
Private mSliderMax As Long
Private mSliderValue As Long
Private mGripLenght As Long
Private mGripWidth As Long

Private Sub cboColorSystem_Click()
    ColorSystem = cboColorSystem.ListIndex
End Sub

Private Sub cboSliderParameter_Click()
    If cboSliderParameter.ListIndex > -1 Then
        SliderParameter = cboSliderParameter.ItemData(cboSliderParameter.ListIndex)
    End If
End Sub

Private Sub chkFixedPalette_Click()
    FixedPalette = (chkFixedPalette.Value = 1)
End Sub

Private Function IBSSubclass_MsgResponse(ByVal hWnd As Long, ByVal iMsg As Long) As Long
    IBSSubclass_MsgResponse = emrPostProcess
End Function

Private Sub IBSSubclass_UnsubclassIt()
    DoUnsubclass
End Sub

Private Function IBSSubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, wParam As Long, lParam As Long, bConsume As Boolean) As Long
    Dim iPt As POINTAPI
    Dim iLng As Long
    Dim iShiftPressed As Boolean
    Dim iDelta As Integer
    Dim iHandle As Boolean
    
    Select Case iMsg
        Case WM_MOUSEWHEEL
            If (wParam And 128) = 0 Then ' if not already handled
                GetCursorPos iPt
                ScreenToClient mUserControlHwnd, iPt
                iShiftPressed = (GetAsyncKeyState(vbKeyShift) < 0)
                If (Sqr((iPt.X - mCx) ^ 2 + (iPt.Y - mCy) ^ 2) <= mRadius) Or (mStyleBox And iPt.X < mPaletteWidth) Then  ' if inside the wheel
                    iDelta = WordHi(wParam)
                    If (GetAsyncKeyState(vbKeyControl) < 0) Then
                        If iDelta > 1 Then
                            iLng = RadialValue + IIf(iShiftPressed, 1, RadialMax / 15)
                            If iLng > RadialMax Then iLng = RadialMax
                            RadialValue = iLng
                        Else
                            iLng = RadialValue - IIf(iShiftPressed, 1, RadialMax / 15)
                            If iLng < 0 Then iLng = 0
                            RadialValue = iLng
                        End If
                        RaiseEvent MouseWheelScroll(cdMouseWheelNavigatingRadial)
                    Else
                        If iDelta > 1 Then
                            iLng = AxialValue + IIf(iShiftPressed, 1, AxialMax / 30)
                            If iLng > AxialMax Then iLng = iLng - AxialMax
                            AxialValue = iLng
                        Else
                            iLng = AxialValue - IIf(iShiftPressed, 1, AxialMax / 30)
                            If iLng < 0 Then iLng = iLng + AxialMax
                            AxialValue = iLng
                        End If
                        RaiseEvent MouseWheelScroll(cdMouseWheelNavigatingAxial)
                    End If
                    wParam = wParam Or 128
                Else
                    iHandle = False
                    If (iPt.X >= (picShades.Left - picShades.Width) / Screen.TwipsPerPixelX) Then 'And (iPt.X <= (picSlider.Left + picSlider.Width + picShades.Width) / Screen.TwipsPerPixelX) Then
                        If (iPt.Y >= picShades.Top / Screen.TwipsPerPixelY) And (iPt.Y <= (picShades.Top + picShades.Height) / Screen.TwipsPerPixelY) Then
                            iHandle = True
                        End If
                    End If
                    If iHandle Then ' if inside or near the slider
                        iDelta = WordHi(wParam)
                        If iDelta > 1 Then
                            SliderValue = SliderValue - IIf(iShiftPressed, 1, mSliderMax / 30)
                        Else
                            SliderValue = SliderValue + IIf(iShiftPressed, 1, mSliderMax / 30)
                        End If
                        RaiseEvent MouseWheelScroll(cdMouseWheelNavigatingSlider)
                        wParam = wParam Or 128
                    End If
                End If
            End If
        Case WM_SYSCOLORCHANGE, WM_THEMECHANGED
            picPalette.Cls
            InitPalette
            StorePaletteColors
            DrawPalette
            picSlider.Cls
            DrawSliderGrip
    End Select
End Function

Private Function WordHi(ByVal LongIn As Long) As Integer
    ' Mask off low word then do integer divide to
    ' shift right by 16.
    
    WordHi = (LongIn And &HFFFF0000) \ &H10000
End Function

Private Sub mForm_Load()
    mDrawEnabled = True
    StorePaletteColors
    DrawPalette
    DrawShades
    ShowSelectedColor
End Sub

Private Sub picShades_DblClick()
    RaiseEvent DblClick
    RaiseEvent DblClickOnColor
End Sub

Private Sub picShades_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iRect As RECT
    Dim iPt As POINTAPI
    
    If Button = 1 Then
        GetClientRect picShades.hWnd, iRect
        iPt.X = iRect.Left
        iPt.Y = iRect.Top
        ClientToScreen picShades.hWnd, iPt
        OffsetRect iRect, iPt.X, iPt.Y
        ClipCursor iRect
        picShades_MouseMove Button, Shift, X, Y
    End If
End Sub

Private Sub picShades_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        SliderValue = mSliderMax / (picShades.ScaleHeight - 1) * (Y - 1)
    End If
End Sub

Private Sub picShades_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClipCursor ByVal 0&
End Sub

Private Sub picPalette_DblClick()
    Dim iPt As POINTAPI
    
    RaiseEvent DblClick
    GetCursorPos iPt
    ScreenToClient picPalette.hWnd, iPt
    If PixelIsInPalette(iPt.X, iPt.Y) Then
        RaiseEvent DblClickOnColor
    End If
End Sub

Private Sub picPalette_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iRect As RECT
    Dim iPt As POINTAPI
    
    If Button = 1 Then
        If PixelIsInPalette(X, Y) Then
            GetClientRect picPalette.hWnd, iRect
            iPt.X = iRect.Left
            iPt.Y = iRect.Top
            ClientToScreen picPalette.hWnd, iPt
            OffsetRect iRect, iPt.X, iPt.Y
            ClipCursor iRect
            PointerVisible = False
            mSelectingColor = True
            picPalette_MouseMove Button, Shift, X, Y
        Else
            RaiseEvent MouseDown(Button, Shift, X, Y)
            mMouseDown = True
        End If
    End If
End Sub

Private Sub DoSliderChange()
    If mSettingColor Then Exit Sub
    If mSettingSlider Then Exit Sub
    If Not mDrawEnabled Then Exit Sub
    
    mChangingShade = True
    
    If mSliderParameter = cdParameterLuminance Then
        mChangingLuminance = True
    ElseIf mSliderParameter = cdParameterHue Then
        mChangingHue = True
    ElseIf mSliderParameter = cdParameterSaturation Then
        mChangingSaturation = True
    End If
    
    If mSliderParameter = cdParameterLuminance Then
        mL = mSliderMax - SliderValue
    ElseIf mSliderParameter = cdParameterHue Then
        mH = mSliderMax - SliderValue
        If mH = mH_Max Then mH = 0
    ElseIf mSliderParameter = cdParameterSaturation Then
        mS = mSliderMax - SliderValue
    ElseIf mSliderParameter = cdParameterRed Then
        mR = mSliderMax - SliderValue
    ElseIf mSliderParameter = cdParameterGreen Then
        mG = mSliderMax - SliderValue
    ElseIf mSliderParameter = cdParameterBlue Then
        mB = mSliderMax - SliderValue
    End If
    
    If (mSliderParameter = cdParameterLuminance) Or (mSliderParameter = cdParameterSaturation) Then
        If (Not mFixedPalette) Then
            tmrDraw.Enabled = True
        End If
    Else
        tmrDraw.Enabled = True
    End If
        
    If Not SetColor(GetShadedColor) Then
        RaiseEvent Change
    End If
    
    If mSliderParameter = cdParameterLuminance Then
        mChangingLuminance = False
    ElseIf mSliderParameter = cdParameterHue Then
        mChangingHue = False
    ElseIf mSliderParameter = cdParameterSaturation Then
        mChangingSaturation = False
    End If
    
    mChangingShade = False
    
End Sub

Private Sub picSlider_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub tmrDraw_Timer()
    tmrDraw.Enabled = False
    DrawPalette
End Sub

Private Sub tmrFixSize_Timer()
    tmrFixSize.Enabled = False
    UserControl.Size mNewWidth, mNewHeight
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    mRedraw = True
    mDrawEnabled = False
    SetCaptions
    mRadialParameter = cdParameterSaturation
    mAxialParameter = cdParameterHue
End Sub

Private Sub UserControl_InitProperties()
    mRoundedBoxes = cPropDefault_RoundedBoxes
    mStyle = cPropDefault_Style
    mStyleBox = mStyle = cdStyleBox
    mHideLabels = cPropDefault_HideLabels
    mSliderWide = cPropDefault_SliderWide
    mColor = cPropDefault_Color
    mSliderOptionsAvailable = cPropDefault_SliderOptionsAvailable
    mFixedPaletteControlVisible = cPropDefault_FixedPaletteControlVisible
    mColorSystemControlVisible = cPropDefault_ColorSystemControlVisible
    mFixedPalette = cPropDefault_FixedPalette
    mSliderParameter = cPropDefault_SliderParameter
    mSliderParameterComboWidth = cPropDefault_SliderParameterComboWidth
    LoadcboSliderParameter
    If Not SelectInListByItemData(cboSliderParameter, mSliderParameter) Then
        If mSliderOptionsAvailable = cdSliderOptionsNone Then
            cboSliderParameter.ListIndex = cboSliderParameter.ListCount - 1
        Else
            cboSliderParameter.ListIndex = 0
        End If
    End If
    mColorSystem = cPropDefault_ColorSystem
    cboColorSystem.ListIndex = mColorSystem
    mBackColor = cPropDefault_BackColor
    SetBackColor
    
    On Error Resume Next
    mAmbientUserMode = Ambient.UserMode
    mUserControlHwnd = UserControl.hWnd
    If mAmbientUserMode Then
        If TypeOf Parent Is Form Then
            Set mForm = Parent
        End If
    End If
    If mForm Is Nothing Then mDrawEnabled = True
    mPropertiesAreSet = True
    init
    mRaiseEvents = True
    On Error Resume Next
    If mAmbientUserMode Then mFormHwnd = UserControl.Parent.hWnd
    On Error GoTo 0
    DoSubclass
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iPointerX As Single
    Dim iPointerY As Single
    
    If picSlider.Visible Then
        If (Shift And vbCtrlMask) Or (UserControl.ActiveControl Is picShades) Or (UserControl.ActiveControl Is picSlider) Then
            If KeyCode = vbKeyUp Then
                If SliderValue > mSliderMin Then
                    SliderValue = SliderValue - 1
                End If
            ElseIf KeyCode = vbKeyDown Then
                If SliderValue < mSliderMax Then
                    SliderValue = SliderValue + 1
                End If
            End If
        Else
            Select Case KeyCode
                Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight
                    iPointerX = mPointerX
                    iPointerY = mPointerY
                    If KeyCode = vbKeyUp Then
                        iPointerY = iPointerY - 1
                    ElseIf KeyCode = vbKeyDown Then
                        iPointerY = iPointerY + 1
                    ElseIf KeyCode = vbKeyLeft Then
                        iPointerX = iPointerX - 1
                    ElseIf KeyCode = vbKeyRight Then
                        iPointerX = iPointerX + 1
                    End If
                    If iPointerY < 0 Then iPointerY = 0
                    If iPointerX < 0 Then iPointerX = 0
                    If iPointerX > mBMPWidth Then iPointerX = mBMPWidth
                    If iPointerY > mBMPHeight Then iPointerY = mBMPHeight
                    If PixelIsInPalette(iPointerX, iPointerY) Then
                        mSelectingColor = True
                        picPalette_MouseUp vbLeftButton, 0, iPointerX, iPointerY
                        mSelectingColor = False
                    End If
            End Select
        End If
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    mMouseDown = True
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mMouseDown Then
        mMouseDown = False
        RaiseEvent MouseUp(Button, Shift, X, Y)
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mRoundedBoxes = PropBag.ReadProperty("RoundedBoxes", cPropDefault_RoundedBoxes)
    mStyle = PropBag.ReadProperty("Style", cPropDefault_Style)
    mStyleBox = mStyle = cdStyleBox
    mHideLabels = PropBag.ReadProperty("HideLabels", cPropDefault_HideLabels)
    mSliderWide = PropBag.ReadProperty("SliderWide", cPropDefault_SliderWide)
    mColor = PropBag.ReadProperty("Color", cPropDefault_Color)
    mSliderOptionsAvailable = PropBag.ReadProperty("SliderOptionsAvailable", cPropDefault_SliderOptionsAvailable)
    mFixedPaletteControlVisible = PropBag.ReadProperty("FixedPaletteControlVisible", cPropDefault_FixedPaletteControlVisible)
    mColorSystemControlVisible = PropBag.ReadProperty("ColorSystemControlVisible", cPropDefault_ColorSystemControlVisible)
    mFixedPalette = PropBag.ReadProperty("FixedPalette", cPropDefault_FixedPalette)
    mSliderParameter = PropBag.ReadProperty("SliderParameter", cPropDefault_SliderParameter)
    mSliderParameterComboWidth = PropBag.ReadProperty("SliderParameterComboWidth", cPropDefault_SliderParameterComboWidth)
    LoadcboSliderParameter
    If Not SelectInListByItemData(cboSliderParameter, mSliderParameter) Then
        If mSliderOptionsAvailable = cdSliderOptionsNone Then
            cboSliderParameter.ListIndex = cboSliderParameter.ListCount - 1
        Else
            cboSliderParameter.ListIndex = 0
        End If
    End If
    mColorSystem = PropBag.ReadProperty("ColorSystem", cPropDefault_ColorSystem)
    cboColorSystem.ListIndex = mColorSystem
    mBackColor = PropBag.ReadProperty("BackColor", cPropDefault_BackColor)
    SetBackColor
    On Error Resume Next
    mAmbientUserMode = Ambient.UserMode
    mUserControlHwnd = UserControl.hWnd
    If mAmbientUserMode Then
        If TypeOf Parent Is Form Then
            Set mForm = Parent
        End If
    End If
    If mForm Is Nothing Then mDrawEnabled = True
    On Error GoTo 0
    mPropertiesAreSet = True
    init
    mRaiseEvents = True
    On Error Resume Next
    If mAmbientUserMode Then mFormHwnd = UserControl.Parent.hWnd
    On Error GoTo 0
    DoSubclass
End Sub

Private Sub DoSubclass()
    Dim iInIDE As Boolean
    
    Debug.Assert MakeTrue(iInIDE)
    
    If (Not iInIDE) Or cSUBCLASS_IN_IDE Then
        If mUserControlHwnd <> 0 Then
            AttachMessage Me, mUserControlHwnd, WM_MOUSEWHEEL
            If mAmbientUserMode And (mFormHwnd <> 0) Then
                AttachMessage Me, mFormHwnd, WM_SYSCOLORCHANGE
                AttachMessage Me, mFormHwnd, WM_THEMECHANGED
            End If
            mSubclassed = True
        End If
    End If
End Sub

Private Function MakeTrue(Value As Boolean) As Boolean
    MakeTrue = True
    Value = True
End Function

Private Sub UserControl_Resize()
    Dim iAdditionalWidth As Long
    Dim ipicShadesWidth As Long
    Static sInside As Long
    Dim iNewHeight As Long
    Dim iNewWidth As Long
    Dim iRgn As Long
    
    sInside = sInside + 1
    
    If mSliderWide Then
        mPicShadesWidth = cWidePicShadesWidth
    Else
        mPicShadesWidth = cNarrowPicShadesWidth
    End If
    
    If mSliderWide = cdYNAuto Then
        If UserControl.Width > (2700 - IIf(mPicShadesWidth = cWidePicShadesWidth, (cWidePicShadesWidth - cNarrowPicShadesWidth), 0)) Then
            mPicShadesWidth = cWidePicShadesWidth
        Else
            mPicShadesWidth = cNarrowPicShadesWidth
        End If
    End If
    If (UserControl.Height >= 2400) Or mStyleBox Then
        ipicShadesWidth = mPicShadesWidth
        iAdditionalWidth = ipicShadesWidth + 150 + 15 + picSlider.Width
    Else
        ipicShadesWidth = mPicShadesWidth - 80
        iAdditionalWidth = ipicShadesWidth + 30 + 15 + picSlider.Width
    End If
    
    If (UserControl.Height + Screen.TwipsPerPixelX) < (IIf(iNewWidth <> 0, iNewWidth, UserControl.Width) - iAdditionalWidth) Then
        iNewHeight = UserControl.Height ' (IIf(iNewWidth <> 0, iNewWidth, UserControl.Width) - iAdditionalWidth)
    End If
    
    If (IIf(iNewHeight <> 0, iNewHeight, UserControl.ScaleHeight) / Screen.TwipsPerPixelY) Mod 2 <> 0 Then
        iNewHeight = Round(IIf(iNewHeight <> 0, iNewHeight, UserControl.Height) / Screen.TwipsPerPixelY / 2) * Screen.TwipsPerPixelY * 2
    End If
    
    If (mSliderOptionsAvailable <> cdSliderOptionsNone) Or mFixedPaletteControlVisible Or mColorSystemControlVisible Then
        If (IIf(iNewHeight <> 0, iNewHeight, UserControl.ScaleHeight) + Screen.TwipsPerPixelY) < 3700 Then
            iNewHeight = 3700
        End If
    End If
    
    If IIf(iNewHeight <> 0, iNewHeight, UserControl.Height) < iAdditionalWidth + 300 Then
        iNewHeight = iAdditionalWidth + 300
    End If
    If (IIf(iNewWidth <> 0, iNewWidth, UserControl.Width) + Screen.TwipsPerPixelX) < (IIf(iNewHeight <> 0, iNewHeight, UserControl.Height) + iAdditionalWidth) Then
        iNewWidth = (IIf(iNewHeight <> 0, iNewHeight, UserControl.Height) + iAdditionalWidth)
    End If
    If Abs(IIf(iNewWidth <> 0, iNewWidth, UserControl.Width) - (UserControl.Height + iAdditionalWidth)) > (Screen.TwipsPerPixelX * 1.3) Then
        iNewWidth = UserControl.Height + iAdditionalWidth
    End If
    If ((iNewHeight <> 0) Or (iNewWidth <> 0)) Then
        If iNewHeight = 0 Then iNewHeight = iNewWidth - iAdditionalWidth
        If iNewWidth = 0 Then iNewWidth = iNewHeight + iAdditionalWidth
        mNewWidth = iNewWidth
        mNewHeight = iNewHeight
        UserControl.Size iNewWidth, iNewHeight
        tmrFixSize.Enabled = True
    End If
    
    picPalette.Move 0, 0, UserControl.ScaleHeight, UserControl.ScaleHeight
    If mRoundedBoxes Then
        iRgn = CreateRoundRectRgn(0, 0, UserControl.ScaleX(picPalette.Width, UserControl.ScaleMode, vbPixels), UserControl.ScaleY(picPalette.Height, UserControl.ScaleMode, vbPixels), 6, 4)
        SetWindowRgn picPalette.hWnd, iRgn, True
        DeleteObject iRgn
    Else
        SetWindowRgn picPalette.hWnd, 0, True
    End If
    
    SetPicShades
    cboColorSystem.Move chkFixedPalette.Left, cboSliderParameter.Top
    lblMode.Move 0, 0
    lblMode.Visible = Not mHideLabels
    lblMode.AutoSize = False
    lblMode.AutoSize = True
    picColorSystem.Move chkFixedPalette.Left, cboSliderParameter.Top - lblMode.Height - 45, lblMode.Width, lblMode.Height
    chkFixedPalette.Visible = mFixedPaletteControlVisible And (Not mStyleBox)
    picColorSystem.Visible = mColorSystemControlVisible And (Not mStyleBox)
    cboColorSystem.Visible = mColorSystemControlVisible And (Not mStyleBox)
    
    sInside = sInside - 1

    If (sInside = 0) And mPropertiesAreSet Then
        InitPalette
        DrawPalette
        DrawShades
        ShowSelectedColor
        picSlider.Width = (mGripWidth + 8) * Screen.TwipsPerPixelX
        DrawSliderGrip
    End If
End Sub

Private Sub SetPicShades()
    Dim ipicShadesHeight As Long
    Dim ipicShadesWidth As Long
    Dim iRgn As Long
    
    If mSliderWide Then
        mPicShadesWidth = cWidePicShadesWidth
    Else
        mPicShadesWidth = cNarrowPicShadesWidth
    End If
    
    If (UserControl.Height >= 2400) Or mStyleBox Then
        ipicShadesWidth = mPicShadesWidth
    Else
        ipicShadesWidth = mPicShadesWidth - 80
    End If
    If mSliderOptionsAvailable <> cdSliderOptionsNone Then
        ipicShadesHeight = UserControl.ScaleHeight - mGripLenght * Screen.TwipsPerPixelY - cboSliderParameter.Height - 45
        chkFixedPalette.Move 90, 90
    Else
        ipicShadesHeight = UserControl.ScaleHeight - mGripLenght * Screen.TwipsPerPixelY
    End If
    picShades.Move picPalette.Width + IIf((UserControl.Height >= 2400 Or mStyleBox), 150, 30), mGripLenght / 2 * Screen.TwipsPerPixelY, ipicShadesWidth, ipicShadesHeight
    picSlider.Move picShades.Left + picShades.Width + 25, 0, picSlider.Width, picShades.Height + (mGripLenght - 1) * Screen.TwipsPerPixelY
    If mStyleBox Then
        On Error Resume Next
        cboSliderParameter.Width = UserControl.ScaleWidth - picShades.Left - 4
        On Error GoTo 0
        cboSliderParameter.Move picShades.Left, UserControl.ScaleHeight - cboSliderParameter.Height
    Else
        cboSliderParameter.Width = mSliderParameterComboWidth
        cboSliderParameter.Move picShades.Left + picShades.Width - cboSliderParameter.Width, UserControl.ScaleHeight - cboSliderParameter.Height
    End If
    cboSliderParameter.Visible = mSliderOptionsAvailable <> cdSliderOptionsNone
    
    If mRoundedBoxes Then
        iRgn = CreateRoundRectRgn(0, 0, UserControl.ScaleX(picShades.Width, UserControl.ScaleMode, vbPixels), UserControl.ScaleY(picShades.Height, UserControl.ScaleMode, vbPixels), 6, 4)
        SetWindowRgn picShades.hWnd, iRgn, True
        DeleteObject iRgn
    Else
        SetWindowRgn picShades.hWnd, 0, True
    End If
End Sub

Private Sub init()
    Dim iColor As Long
    Dim iRedrawPrev As Boolean
    
    mInitialized = False
    TranslateColor mColor, 0, iColor
    
    iRedrawPrev = Redraw
    Redraw = False
    InitSlider
    InitPalette
    SetMaxAndFixedvalues
    SetSliderParameter
    
    mR = iColor And 255
    mG = (iColor \ 256) And 255
    mB = (iColor \ 65536) And 255
    ColorRGBToCurrentColorSystem iColor, mH, mL, mS
    SetColor iColor
    
    ShowSelectedColor
    
    SetPicShades
    picColorSystem.Visible = mColorSystemControlVisible And (Not mStyleBox)
    cboColorSystem.Visible = mColorSystemControlVisible And (Not mStyleBox)
    chkFixedPalette.Visible = mFixedPaletteControlVisible And (Not mStyleBox)
    chkFixedPalette.Value = Abs(CLng(mFixedPalette))
    
    iColor = mColor
    mColor = -1
    mChangingColorSystemOrInitializing = True
    SetColor iColor
    mChangingColorSystemOrInitializing = False
    Redraw = iRedrawPrev
    mInitialized = True
End Sub

Private Sub InitPalette()
    Dim iBMP As BITMAP
    Dim IPic As StdPicture
    Dim c As Long
    Dim iUb As Long
    Dim iX As Long
    Dim iY As Long
    Dim iNear As Boolean
    Dim iX2 As Long
    Dim iY2 As Long
    Dim i As Long
    Dim iB1 As Long
    Dim iB2 As Long
    Dim iBackColor As Long
    Dim iBMPiH As BITMAPINFOHEADER
    Dim iPixelsBytes() As Byte
    Dim iPix_XY()  As POINTAPI
    Dim iPixToCheck() As Long
    Dim iUbp As Long
    Dim iIndexPixBorder As Long
    Dim iIndexPixToCheck As Long
    Dim iPixToCheckCount As Long
    Dim iCenterX As Long
    Dim iCenterY As Long
    Dim iRadius As Long
    Dim iDistanceToCircumference As Long
    
    If mDiameter = picPalette.ScaleWidth - 16 Then Exit Sub
    
    mCx = picPalette.ScaleWidth / 2
    mCy = picPalette.ScaleHeight / 2
    mDiameter = picPalette.ScaleWidth - 8
    mRadius = mDiameter / 2
    mPaletteWidth = picPalette.ScaleWidth
    mPaletteHeight = picPalette.ScaleHeight
    
    picAux.Move picAux.Left, picAux.Top, picPalette.Width, picPalette.Height
    picAux.FillStyle = vbFSSolid
    picAux.DrawWidth = 1
    picAux.BackColor = UserControl.BackColor
    iBackColor = picAux.BackColor
    TranslateColor iBackColor, 0&, iBackColor
    iB1 = (iBackColor \ 65536) And 255
    If iB1 = 255 Then
        iB2 = 200
    Else
        iB2 = 255
    End If
    picAux.FillColor = RGB(255, 255, iB2)
    picAux.Cls
    picAux.Circle (picAux.ScaleWidth / 2 - 1, picAux.ScaleHeight / 2 - 1), mRadius, RGB(255, 255, iB2)
    Set IPic = picAux.Image
    picAux.Cls
    
    GetObject IPic.Handle, Len(iBMP), iBMP
    With mBMPiH
        .biSize = Len(mBMPiH)
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = iBMP.bmWidth
        .biHeight = iBMP.bmHeight
        .biSizeImage = ((.biWidth * 4 + 3) And &HFFFFFFFC) * .biHeight
    End With
    ReDim mPixelsBytes(Len(mBMPiH) + mBMPiH.biSizeImage)
    GetDIBits picAux.HDC, IPic.Handle, 0, iBMP.bmHeight, mPixelsBytes(0), mBMPiH, DIB_RGB_COLORS
    
    mBMPHeight = iBMP.bmHeight
    mBMPWidth = iBMP.bmWidth
    ReDim mPixelsAreInPalette(UBound(mPixelsBytes) / 4)
    ReDim mPixelsAngleOrX(UBound(mPixelsAreInPalette))
    ReDim mPixelsRadiusOrY(UBound(mPixelsAreInPalette))
    ReDim mPixelsBytes2(UBound(mPixelsBytes))
    
    mBytesStride = mBMPWidth * 4
    mBytesCount = mBMPiH.biSizeImage - 1
    
    If mStyleBox Then
        For c = 0 To mBytesCount - 4 Step 4
            mPixelsAreInPalette(c / 4) = True
       Next c
    Else
        For c = 0 To mBytesCount - 4 Step 4
            If mPixelsBytes(c) = iB2 Then
                mPixelsAreInPalette(c / 4) = True
            End If
        Next c
    End If
    mPaletteColorsStored = False
    
    ' Prepare the info to add an anti-alias border
    picAux.Cls
    picAux.FillStyle = vbFSTransparent
    picAux.DrawWidth = 5
    picAux.Circle (picAux.ScaleWidth / 2 - 1, picAux.ScaleHeight / 2 - 1), mRadius, RGB(255, 255, iB2)
    Set IPic = picAux.Image
    picAux.Cls
    
    GetObject IPic.Handle, Len(iBMP), iBMP
    With iBMPiH
        .biSize = Len(mBMPiH)
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = iBMP.bmWidth
        .biHeight = iBMP.bmHeight
        .biSizeImage = ((.biWidth * 4 + 3) And &HFFFFFFFC) * .biHeight
    End With
    ReDim iPixelsBytes(Len(iBMPiH) + iBMPiH.biSizeImage)
    GetDIBits picAux.HDC, IPic.Handle, 0, iBMP.bmHeight, iPixelsBytes(0), iBMPiH, DIB_RGB_COLORS
    
    iCenterX = mCx * 100 - 50
    iCenterY = mCy * 100 - 50
    iRadius = mRadius * 100 + 55
    ' find one of the pixels that belong to the border
    iX = 0
    iY = 0
    i = (iY + 1) * mBMPHeight
    i = i + iX
    i = i * 4
    Do Until iPixelsBytes(i) = iB2
        iX = iX + 1
        iY = iY + 1
        i = (iY + 1) * mBMPHeight
        i = i + iX
        i = i * 4
        If i > mBytesCount Then Exit Sub
    Loop
    iUb = 1000
    If Not mStyleBox Then
        ReDim mBorderPixels(iUb)
        ReDim mBorderPixels_Alpha(iUb)
        iIndexPixBorder = 0
        If (iX > 0) And (iY) > 0 Then
            ReDim iPix_XY(mBytesCount)
            iUbp = 1000
            ReDim iPixToCheck(iUbp)
            iIndexPixToCheck = 1
            iPixToCheckCount = 1
            iPix_XY(i).X = iX
            iPix_XY(i).Y = iY
            iPixToCheck(iIndexPixToCheck) = i
            iDistanceToCircumference = Sqr((iCenterX - iX * 100 - 50) ^ 2 + (iCenterY - iY * 100 - 50) ^ 2) - iRadius
            iDistanceToCircumference = iDistanceToCircumference
            If iDistanceToCircumference < 0 Then iDistanceToCircumference = 0
            If iDistanceToCircumference > 255 Then iDistanceToCircumference = 255
            mBorderPixels_Alpha(iIndexPixBorder) = 255 - iDistanceToCircumference
            ' find the pixels of the border and save then in a vector (without scanning the whole picture)
            ' for each pixel, check all nearby whether they belong to the circunsference border or not
            Do While iIndexPixToCheck <= iPixToCheckCount
                iX = iPix_XY(iPixToCheck(iIndexPixToCheck)).X
                iY = iPix_XY(iPixToCheck(iIndexPixToCheck)).Y
                For iX2 = iX - 1 To iX + 1
                    For iY2 = iY - 1 To iY + 1
                        i = (iY2 + 1) * mBMPHeight
                        i = i + iX2
                        i = i * 4
                        If (i >= 0) And (i <= mBytesCount) Then
                            If iPixelsBytes(i) = iB2 Then ' it is part of the border
                                If iPix_XY(i).X = 0 Then ' if it is not already added
                                    If Not mPixelsAreInPalette(i / 4) Then ' if it is not inside the wheel
                                        iPixToCheckCount = iPixToCheckCount + 1
                                        If iPixToCheckCount > iUbp Then
                                            iUbp = iUbp + 1000
                                            ReDim Preserve iPixToCheck(iUbp)
                                        End If
                                        iPix_XY(i).X = iX2
                                        iPix_XY(i).Y = iY2
                                        iPixToCheck(iPixToCheckCount) = i
                                        iIndexPixBorder = iIndexPixBorder + 1
                                        If iIndexPixBorder > iUb Then
                                            iUb = iUb + 1000
                                            ReDim Preserve mBorderPixels(iUb)
                                            ReDim Preserve mBorderPixels_Alpha(iUb)
                                        End If
                                        mBorderPixels(iIndexPixBorder) = i
                                        iDistanceToCircumference = Sqr((iCenterX - iX2 * 100 - 50) ^ 2 + (iCenterY - iY2 * 100 - 50) ^ 2) - iRadius
                                        iDistanceToCircumference = iDistanceToCircumference * 1.5
                                        If iDistanceToCircumference < 0 Then iDistanceToCircumference = 0
                                        If iDistanceToCircumference > 255 Then iDistanceToCircumference = 255
                                        mBorderPixels_Alpha(iIndexPixBorder) = 255 - iDistanceToCircumference
                                    End If
                                End If
                            End If
                        End If
                    Next iY2
                Next iX2
                iIndexPixToCheck = iIndexPixToCheck + 1
            Loop
        End If
        ReDim Preserve mBorderPixels(iIndexPixBorder)
    End If
End Sub

Private Sub StorePaletteColors()
    Dim c As Long
    Dim iX As Long
    Dim iY As Long
    Dim iHorz As Long
    Dim iVert As Long
    Dim iAngle As Single
    Dim iRadius As Single
    Dim iColor As Long
    Dim iL1 As Long
    Dim i As Long
    Dim iP1 As Long
    Dim iP2 As Long
    Dim iRGB As RGBQuad
    
    If Not mDrawEnabled Then Exit Sub
    
    If mColorSystem = cdColorSystemHSV Then
        For c = 0 To mBytesCount - 4 Step 4
            iX = (c Mod mBytesStride) / 4
            iY = mBMPHeight - (c / mBMPHeight / 4 - 0.4999) - 1
            If iY < 0 Then iY = 0
            i = c / 4
            If mStyleBox Then
                If mSliderParameter = cdParameterLuminance Then
                    mPixelsAngleOrX(i) = mH_Max / mPaletteWidth * iX
                    mPixelsRadiusOrY(i) = mS_Max - mS_Max / mPaletteHeight * iY
                    iColor = ColorHSVToRGB(mPixelsAngleOrX(i), mL_Fixed, mPixelsRadiusOrY(i))
                ElseIf mSliderParameter = cdParameterHue Then
                    mPixelsAngleOrX(i) = mL_Max / mPaletteWidth * iX
                    mPixelsRadiusOrY(i) = mS_Max - mS_Max / mPaletteHeight * iY
                    iColor = ColorHSVToRGB(mH_Fixed, mPixelsAngleOrX(i), mPixelsRadiusOrY(i))
                ElseIf mSliderParameter = cdParameterSaturation Then
                    mPixelsAngleOrX(i) = mH_Max / mPaletteWidth * iX
                    mPixelsRadiusOrY(i) = mL_Max - mL_Max / mPaletteHeight * iY
                    iColor = ColorHSVToRGB(mPixelsAngleOrX(i), mPixelsRadiusOrY(i), mS_Fixed)
                ElseIf mSliderParameter = cdParameterRed Then
                    mPixelsAngleOrX(i) = 255 / mPaletteWidth * iX
                    mPixelsRadiusOrY(i) = 255 - 255 / mPaletteHeight * iY
                    iColor = RGB(0, mPixelsAngleOrX(i), mPixelsRadiusOrY(i))
                ElseIf mSliderParameter = cdParameterGreen Then
                    mPixelsAngleOrX(i) = 255 / mPaletteWidth * iX
                    mPixelsRadiusOrY(i) = 255 - 255 / mPaletteHeight * iY
                    iColor = RGB(mPixelsAngleOrX(i), 0, mPixelsRadiusOrY(i))
                ElseIf mSliderParameter = cdParameterBlue Then
                    mPixelsAngleOrX(i) = 255 / mPaletteWidth * iX
                    mPixelsRadiusOrY(i) = 255 - 255 / mPaletteHeight * iY
                    iColor = RGB(mPixelsAngleOrX(i), mPixelsRadiusOrY(i), 0)
                End If
                    
                CopyMemory iRGB, iColor, 4
                
                mPixelsBytes(c + 2) = iRGB.R
                mPixelsBytes(c + 1) = iRGB.G
                mPixelsBytes(c) = iRGB.B
                
            Else
                If mPixelsAreInPalette(i) Then
                    iHorz = iX - mCx
                    iVert = mCy - iY
                    If iHorz = 0 Then
                        iAngle = 90 * Pi / 180 ' angle is hue
                    Else
                        iAngle = Atn(iVert / iHorz)
                    End If
                    iAngle = 180 * iAngle / Pi
                    
                    If (iHorz >= 0) And (iVert >= 0) Then
                        ' ok
                    ElseIf (iHorz < 0) And (iVert >= 0) Then
                        iAngle = 180 - iAngle * -1
                    ElseIf (iHorz <= 0) And (iVert < 0) Then
                        iAngle = iAngle + 180
                    Else
                        iAngle = iAngle + 360
                    End If
                    
                    iRadius = Sqr(iHorz ^ 2 + iVert ^ 2) ' iRadius is saturation
                    
                    If mSliderParameter = cdParameterLuminance Then
                        iAngle = iAngle / 360 * mH_Max
                        mPixelsAngleOrX(i) = iAngle
                        iRadius = Int(iRadius / mDiameter * mS_Max * 2)
                        If iRadius > mS_Max Then iRadius = mS_Max
                        If iRadius < 1 Then iRadius = 1
                        mPixelsRadiusOrY(i) = iRadius
                        iColor = ColorHSVToRGB(mPixelsAngleOrX(i), mL_Fixed, mPixelsRadiusOrY(i))
                    ElseIf mSliderParameter = cdParameterHue Then
                        iAngle = iAngle / 360 * mL_Max
                        mPixelsAngleOrX(i) = iAngle
                        iRadius = Int(iRadius / mDiameter * mS_Max * 2)
                        If iRadius > mS_Max Then iRadius = mS_Max
                        If iRadius < 1 Then iRadius = 1
                        mPixelsRadiusOrY(i) = iRadius
                        iColor = ColorHSVToRGB(mH_Fixed, mPixelsAngleOrX(i), mPixelsRadiusOrY(i))
                    ElseIf mSliderParameter = cdParameterSaturation Then
                        iAngle = iAngle / 360 * mH_Max
                        mPixelsAngleOrX(i) = iAngle
                        iRadius = Int(iRadius / mDiameter * mL_Max * 2)
                        If iRadius > mL_Max Then iRadius = mL_Max
                        If iRadius < 1 Then iRadius = 1
                        mPixelsRadiusOrY(i) = iRadius
                        iColor = ColorHSVToRGB(mPixelsAngleOrX(i), mPixelsRadiusOrY(i), mS_Fixed)
                    ElseIf mSliderParameter = cdParameterRed Then
                        iAngle = iAngle / 360 * 255
                        mPixelsAngleOrX(i) = iAngle
                        iRadius = Int(iRadius / mDiameter * 255 * 2)
                        If iRadius > 255 Then iRadius = 255
                        If iRadius < 1 Then iRadius = 1
                        mPixelsRadiusOrY(i) = iRadius
                        iP1 = mPixelsAngleOrX(i): If iP1 > 255 Then iP1 = 255
                        iP2 = mPixelsRadiusOrY(i): If iP2 > 255 Then iP2 = 255
                        iColor = RGB(0, iP1, iP2)
                    ElseIf mSliderParameter = cdParameterGreen Then
                        iAngle = iAngle / 360 * 255
                        mPixelsAngleOrX(i) = iAngle
                        iRadius = Int(iRadius / mDiameter * 255 * 2)
                        If iRadius > 255 Then iRadius = 255
                        If iRadius < 1 Then iRadius = 1
                        mPixelsRadiusOrY(i) = iRadius
                        iP1 = mPixelsAngleOrX(i): If iP1 > 255 Then iP1 = 255
                        iP2 = mPixelsRadiusOrY(i): If iP2 > 255 Then iP2 = 255
                        iColor = RGB(iP1, 0, iP2)
                    ElseIf mSliderParameter = cdParameterBlue Then
                        iAngle = iAngle / 360 * 255
                        mPixelsAngleOrX(i) = iAngle
                        iRadius = Int(iRadius / mDiameter * 255 * 2)
                        If iRadius > 255 Then iRadius = 255
                        If iRadius < 1 Then iRadius = 1
                        mPixelsRadiusOrY(i) = iRadius
                        iP1 = mPixelsAngleOrX(i): If iP1 > 255 Then iP1 = 255
                        iP2 = mPixelsRadiusOrY(i): If iP2 > 255 Then iP2 = 255
                        iColor = RGB(iP1, iP2, 0)
                    End If
                    
                    CopyMemory iRGB, iColor, 4
                    
                    mPixelsBytes(c + 2) = iRGB.R
                    mPixelsBytes(c + 1) = iRGB.G
                    mPixelsBytes(c) = iRGB.B
                    
                Else
                    mPixelsBytes2(c) = mPixelsBytes(c)
                    mPixelsBytes2(c + 1) = mPixelsBytes(c + 1)
                    mPixelsBytes2(c + 2) = mPixelsBytes(c + 2)
                End If
            End If
        Next c
    Else
        For c = 0 To mBytesCount - 4 Step 4
            iX = (c Mod mBytesStride) / 4
            iY = mBMPHeight - (c / mBMPHeight / 4 - 0.4999) - 1
            i = c / 4
            If mStyleBox Then
                If mSliderParameter = cdParameterLuminance Then
                    mPixelsAngleOrX(i) = mH_Max / mPaletteWidth * iX
                    mPixelsRadiusOrY(i) = mS_Max - mS_Max / mPaletteHeight * iY
                    iColor = ColorHLSToRGB(mPixelsAngleOrX(i), mL_Fixed, mPixelsRadiusOrY(i))
                ElseIf mSliderParameter = cdParameterHue Then
                    mPixelsAngleOrX(i) = mL_Max / mPaletteWidth * iX
                    mPixelsRadiusOrY(i) = mS_Max - mS_Max / mPaletteHeight * iY
                    iColor = ColorHLSToRGB(mH_Fixed, mPixelsAngleOrX(i), mPixelsRadiusOrY(i))
                ElseIf mSliderParameter = cdParameterSaturation Then
                    mPixelsAngleOrX(i) = mH_Max / mPaletteWidth * iX
                    mPixelsRadiusOrY(i) = mL_Max - mL_Max / mPaletteHeight * iY
                    iColor = ColorHLSToRGB(mPixelsAngleOrX(i), mPixelsRadiusOrY(i), mS_Fixed)
                ElseIf mSliderParameter = cdParameterRed Then
                    mPixelsAngleOrX(i) = 255 / mPaletteWidth * iX
                    mPixelsRadiusOrY(i) = 255 - 255 / mPaletteHeight * iY
                    iColor = RGB(0, mPixelsAngleOrX(i), mPixelsRadiusOrY(i))
                ElseIf mSliderParameter = cdParameterGreen Then
                    mPixelsAngleOrX(i) = 255 / mPaletteWidth * iX
                    mPixelsRadiusOrY(i) = 255 - 255 / mPaletteHeight * iY
                    iColor = RGB(mPixelsAngleOrX(i), 0, mPixelsRadiusOrY(i))
                ElseIf mSliderParameter = cdParameterBlue Then
                    mPixelsAngleOrX(i) = 255 / mPaletteWidth * iX
                    mPixelsRadiusOrY(i) = 255 - 255 / mPaletteHeight * iY
                    iColor = RGB(mPixelsAngleOrX(i), mPixelsRadiusOrY(i), 0)
                End If
                    
                CopyMemory iRGB, iColor, 4
                
                mPixelsBytes(c + 2) = iRGB.R
                mPixelsBytes(c + 1) = iRGB.G
                mPixelsBytes(c) = iRGB.B
                
            Else
                If mPixelsAreInPalette(i) Then
                    iHorz = iX - mCx
                    iVert = mCy - iY
                    If iHorz = 0 Then
                        iAngle = 90 * Pi / 180 ' angle is hue
                    Else
                        iAngle = Atn(iVert / iHorz)
                    End If
                    iAngle = 180 * iAngle / Pi
                    
                    If (iHorz >= 0) And (iVert >= 0) Then
                        ' ok
                    ElseIf (iHorz < 0) And (iVert >= 0) Then
                        iAngle = 180 - iAngle * -1
                    ElseIf (iHorz <= 0) And (iVert < 0) Then
                        iAngle = iAngle + 180
                    Else
                        iAngle = iAngle + 360
                    End If
                    
                    iRadius = Sqr(iHorz ^ 2 + iVert ^ 2) ' iRadius is saturation
                    
                    If mSliderParameter = cdParameterLuminance Then
                        iAngle = iAngle / 360 * mH_Max
                        mPixelsAngleOrX(i) = iAngle
                        iRadius = Int(iRadius / mDiameter * mS_Max * 2)
                        If iRadius > mS_Max Then iRadius = mS_Max
                        If iRadius < 1 Then iRadius = 1
                        mPixelsRadiusOrY(i) = iRadius
                        iColor = ColorHLSToRGB(mPixelsAngleOrX(i), mL_Fixed, mPixelsRadiusOrY(i))
                    ElseIf mSliderParameter = cdParameterHue Then
                        iAngle = iAngle / 360 * mL_Max
                        mPixelsAngleOrX(i) = iAngle
                        iRadius = Int(iRadius / mDiameter * mS_Max * 2)
                        If iRadius > mS_Max Then iRadius = mS_Max
                        If iRadius < 1 Then iRadius = 1
                        mPixelsRadiusOrY(i) = iRadius
                        iColor = ColorHLSToRGB(mH_Fixed, mPixelsAngleOrX(i), mPixelsRadiusOrY(i))
                    ElseIf mSliderParameter = cdParameterSaturation Then
                        iAngle = iAngle / 360 * mH_Max
                        mPixelsAngleOrX(i) = iAngle
                        iRadius = Int(iRadius / mDiameter * mL_Max * 2)
                        If iRadius > mL_Max Then iRadius = mL_Max
                        If iRadius < 1 Then iRadius = 1
                        mPixelsRadiusOrY(i) = iRadius
                        iColor = ColorHLSToRGB(mPixelsAngleOrX(i), mPixelsRadiusOrY(i), mS_Fixed)
                    ElseIf mSliderParameter = cdParameterRed Then
                        iAngle = iAngle / 360 * 255
                        mPixelsAngleOrX(i) = iAngle
                        iRadius = Int(iRadius / mDiameter * 255 * 2)
                        If iRadius > 255 Then iRadius = 255
                        If iRadius < 1 Then iRadius = 1
                        mPixelsRadiusOrY(i) = iRadius
                        iP1 = mPixelsAngleOrX(i): If iP1 > 255 Then iP1 = 255
                        iP2 = mPixelsRadiusOrY(i): If iP2 > 255 Then iP2 = 255
                        iColor = RGB(0, iP1, iP2)
                    ElseIf mSliderParameter = cdParameterGreen Then
                        iAngle = iAngle / 360 * 255
                        mPixelsAngleOrX(i) = iAngle
                        iRadius = Int(iRadius / mDiameter * 255 * 2)
                        If iRadius > 255 Then iRadius = 255
                        If iRadius < 1 Then iRadius = 1
                        mPixelsRadiusOrY(i) = iRadius
                        iP1 = mPixelsAngleOrX(i): If iP1 > 255 Then iP1 = 255
                        iP2 = mPixelsRadiusOrY(i): If iP2 > 255 Then iP2 = 255
                        iColor = RGB(iP1, 0, iP2)
                    ElseIf mSliderParameter = cdParameterBlue Then
                        iAngle = iAngle / 360 * 255
                        mPixelsAngleOrX(i) = iAngle
                        iRadius = Int(iRadius / mDiameter * 255 * 2)
                        If iRadius > 255 Then iRadius = 255
                        If iRadius < 1 Then iRadius = 1
                        mPixelsRadiusOrY(i) = iRadius
                        iP1 = mPixelsAngleOrX(i): If iP1 > 255 Then iP1 = 255
                        iP2 = mPixelsRadiusOrY(i): If iP2 > 255 Then iP2 = 255
                        iColor = RGB(iP1, iP2, 0)
                    End If
                    
                    CopyMemory iRGB, iColor, 4
                    
                    mPixelsBytes(c + 2) = iRGB.R
                    mPixelsBytes(c + 1) = iRGB.G
                    mPixelsBytes(c) = iRGB.B
                    
                Else
                    mPixelsBytes2(c) = mPixelsBytes(c)
                    mPixelsBytes2(c + 1) = mPixelsBytes(c + 1)
                    mPixelsBytes2(c + 2) = mPixelsBytes(c + 2)
                End If
            End If
        Next c
    End If
    mPaletteColorsStored = True
    
End Sub

Private Sub DrawPalette()
    Dim c As Long
    Dim iColor As Long
    Dim iP1 As Long
    Dim i As Long
    Dim iRGB As RGBQuad
    Dim iP1b As Double
    Dim iR1 As Long
    Dim iG1 As Long
    Dim iB1 As Long
    Dim iX As Long
    Dim iY As Long
    Dim iX2 As Long
    Dim iY2 As Long
    Dim i2 As Long
    Dim c2 As Single
    Dim iDo As Boolean
    Dim iBackColor As RGBQuad
    Dim iLng As Long
    Dim t As Long
    
    If Not mRedraw Then
        mDrawPending = True
        Exit Sub
    End If
    If Not mDrawEnabled Then Exit Sub
    
    tmrDraw.Enabled = False
    mDrawPending = False
    If mChangingColorSystemOrInitializing Then Exit Sub
    If Not mPaletteColorsStored Then
        StorePaletteColors
    End If
    
    If mSliderParameter = cdParameterLuminance Then
        If mColorSystem = cdColorSystemHSV Then
            If mFixedPalette Then
                iP1b = mL_Fixed
            Else
                iP1b = mL
            End If
            For c = 0 To mBytesCount - 4 Step 4
                i = c / 4
                If mPixelsAreInPalette(i) Then
                    iColor = ColorHSVToRGB(mPixelsAngleOrX(i), iP1b, mPixelsRadiusOrY(i))
                    CopyMemory iRGB, iColor, 4
                    mPixelsBytes2(c + 2) = iRGB.R
                    mPixelsBytes2(c + 1) = iRGB.G
                    mPixelsBytes2(c) = iRGB.B
                End If
            Next c
        Else
            If mFixedPalette Then
                iP1 = mL_Fixed
            Else
                iP1 = mL
            End If
            For c = 0 To mBytesCount - 4 Step 4
                i = c / 4
                If mPixelsAreInPalette(i) Then
                    iColor = ColorHLSToRGB(mPixelsAngleOrX(i), iP1, mPixelsRadiusOrY(i))
                    CopyMemory iRGB, iColor, 4
                    mPixelsBytes2(c + 2) = iRGB.R
                    mPixelsBytes2(c + 1) = iRGB.G
                    mPixelsBytes2(c) = iRGB.B
                End If
            Next c
        End If
    ElseIf mSliderParameter = cdParameterHue Then
        If mColorSystem = cdColorSystemHSV Then
            iP1b = mH
            For c = 0 To mBytesCount - 4 Step 4
                i = c / 4
                If mPixelsAreInPalette(i) Then
                    iColor = ColorHSVToRGB(iP1b, mPixelsAngleOrX(i), mPixelsRadiusOrY(i))
                    CopyMemory iRGB, iColor, 4
                    mPixelsBytes2(c + 2) = iRGB.R
                    mPixelsBytes2(c + 1) = iRGB.G
                    mPixelsBytes2(c) = iRGB.B
                End If
            Next c
        Else
            iP1 = mH
            For c = 0 To mBytesCount - 4 Step 4
                i = c / 4
                If mPixelsAreInPalette(i) Then
                    iColor = ColorHLSToRGB(iP1, mPixelsAngleOrX(i), mPixelsRadiusOrY(i))
                    CopyMemory iRGB, iColor, 4
                    mPixelsBytes2(c + 2) = iRGB.R
                    mPixelsBytes2(c + 1) = iRGB.G
                    mPixelsBytes2(c) = iRGB.B
                End If
            Next c
        End If
    ElseIf mSliderParameter = cdParameterSaturation Then
        If mColorSystem = cdColorSystemHSV Then
            If mFixedPalette Then
                iP1b = mS_Fixed
            Else
                iP1b = mS
            End If
            For c = 0 To mBytesCount - 4 Step 4
                i = c / 4
                If mPixelsAreInPalette(i) Then
                    iColor = ColorHSVToRGB(mPixelsAngleOrX(i), mPixelsRadiusOrY(i), iP1b)
                    CopyMemory iRGB, iColor, 4
                    mPixelsBytes2(c + 2) = iRGB.R
                    mPixelsBytes2(c + 1) = iRGB.G
                    mPixelsBytes2(c) = iRGB.B
                End If
            Next c
        Else
            If mFixedPalette Then
                iP1 = mS_Fixed
            Else
                iP1 = mS
            End If
            For c = 0 To mBytesCount - 4 Step 4
                i = c / 4
                If mPixelsAreInPalette(i) Then
                    If (Round(mPixelsAngleOrX(i)) = 160) And (iP1 = 0) Then
                        iColor = 0
                    Else
                        iColor = ColorHLSToRGB(mPixelsAngleOrX(i), mPixelsRadiusOrY(i), iP1)
                    End If
                    CopyMemory iRGB, iColor, 4
                    mPixelsBytes2(c + 2) = iRGB.R
                    mPixelsBytes2(c + 1) = iRGB.G
                    mPixelsBytes2(c) = iRGB.B
                End If
            Next c
        End If
    ElseIf mSliderParameter = cdParameterRed Then
        For c = 0 To mBytesCount - 4 Step 4
            i = c / 4
            If mPixelsAreInPalette(i) Then
                If mFixedPalette Then
                    iP1 = 0
                Else
                    iP1 = mR
                End If
                mPixelsBytes2(c + 2) = iP1 'R
                mPixelsBytes2(c + 1) = mPixelsAngleOrX(i) 'G
                mPixelsBytes2(c) = mPixelsRadiusOrY(i) 'B
            End If
        Next c
    ElseIf mSliderParameter = cdParameterGreen Then
        For c = 0 To mBytesCount - 4 Step 4
            i = c / 4
            If mPixelsAreInPalette(i) Then
                If mFixedPalette Then
                    iP1 = 0
                Else
                    iP1 = mG
                End If
                mPixelsBytes2(c + 2) = mPixelsAngleOrX(i) 'R
                mPixelsBytes2(c + 1) = iP1 'G
                mPixelsBytes2(c) = mPixelsRadiusOrY(i) 'B
            End If
        Next c
    ElseIf mSliderParameter = cdParameterBlue Then
        For c = 0 To mBytesCount - 4 Step 4
            i = c / 4
            If mPixelsAreInPalette(i) Then
                If mFixedPalette Then
                    iP1 = 0
                Else
                    iP1 = mB
                End If
                mPixelsBytes2(c + 2) = mPixelsAngleOrX(i) 'R
                mPixelsBytes2(c + 1) = mPixelsRadiusOrY(i) 'G
                mPixelsBytes2(c) = iP1 'B
            End If
        Next c
    End If
    
    If Not mStyleBox Then
        ' Add a anti-aliased border
        Call OleTranslateColor(UserControl.BackColor, 0, iBackColor)
        
        For t = 1 To 3
            For c = 1 To UBound(mBorderPixels)
                i = mBorderPixels(c)
                iX = (i Mod mBytesStride) / 4
                iY = i / mBMPHeight / 4 - 0.4999 - 1
                iR1 = 0
                iG1 = 0
                iB1 = 0
                c2 = 0
                For iX2 = iX - 1 To iX + 1
                    For iY2 = iY - 1 To iY + 1
                        i2 = (iY2 + 1) * mBMPHeight
                        i2 = i2 + iX2
                        i2 = i2 * 4
                        If (i2 > 0) And (i2 <= (mBytesCount - 4)) Then
                            If t = 1 Then
                                iDo = mPixelsAreInPalette(i2 / 4)
                            Else
                                iDo = True
                            End If
                            If iDo Then
                                If (iX2 = iX) Or (iY2 = iY) Then
                                    c2 = c2 + 0.7
                                    iR1 = iR1 + mPixelsBytes2(i2 + 2) * 0.7
                                    iG1 = iG1 + mPixelsBytes2(i2 + 1) * 0.7
                                    iB1 = iB1 + mPixelsBytes2(i2) * 0.7
                                Else
                                    iR1 = iR1 + mPixelsBytes2(i2 + 2)
                                    iG1 = iG1 + mPixelsBytes2(i2 + 1)
                                    iB1 = iB1 + mPixelsBytes2(i2)
                                    c2 = c2 + 1
                                End If
                            End If
                        End If
                    Next
                Next
                
                If c2 > 0 Then
                    iLng = iR1 / c2 * mBorderPixels_Alpha(c) / 255 + CLng(iBackColor.R) * (255 - mBorderPixels_Alpha(c)) / 255
                    If iLng > 255 Then iLng = 255
                    mPixelsBytes2(i + 2) = iLng
                    iLng = iG1 / c2 * mBorderPixels_Alpha(c) / 255 + CLng(iBackColor.G) * (255 - mBorderPixels_Alpha(c)) / 255
                    If iLng > 255 Then iLng = 255
                    mPixelsBytes2(i + 1) = iLng
                    iLng = iB1 / c2 * mBorderPixels_Alpha(c) / 255 + CLng(iBackColor.B) * (255 - mBorderPixels_Alpha(c)) / 255
                    If iLng > 255 Then iLng = 255
                    mPixelsBytes2(i) = iLng
                End If
            Next c
        Next t
    End If
    
    SetDIBitsToDevice picPalette.HDC, 0, 0, mBMPWidth, mBMPHeight, 0, 0, 0, mBMPHeight, mPixelsBytes2(0), mBMPiH, DIB_RGB_COLORS
    picPalette.Refresh
    SetPointer
End Sub

Private Function PixelIsInPalette(ByVal X As Single, ByVal Y As Single) As Boolean
    Dim i As Long
    
    If mStyleBox Then
        If (X >= 0) And (X < mBMPWidth) And (Y >= 0) And (Y < mBMPHeight) Then
            PixelIsInPalette = True
        End If
    Else
        X = Int(X)
        Y = Int(Y)
        If (X >= 0) And (X < mBMPWidth) And (Y >= 0) And (Y < mBMPHeight) Then
            i = (Y + 1) * mBMPHeight
            i = i + X
            If (i >= 0) And (i <= UBound(mPixelsAreInPalette)) Then
                PixelIsInPalette = mPixelsAreInPalette(i)
            End If
        End If
    End If
End Function

Private Sub UserControl_Terminate()
    DoUnsubclass
End Sub

Private Sub DoUnsubclass()
    If mSubclassed Then
        DetachMessage Me, mUserControlHwnd, WM_MOUSEWHEEL
        If mAmbientUserMode And (mFormHwnd <> 0) Then
            DetachMessage Me, mFormHwnd, WM_SYSCOLORCHANGE
            DetachMessage Me, mFormHwnd, WM_THEMECHANGED
        End If
        mSubclassed = False
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "RoundedBoxes", mRoundedBoxes, cPropDefault_RoundedBoxes
    PropBag.WriteProperty "Style", mStyle, cPropDefault_Style
    PropBag.WriteProperty "HideLabels", mHideLabels, cPropDefault_HideLabels
    PropBag.WriteProperty "SliderWide", mSliderWide, cPropDefault_SliderWide
    PropBag.WriteProperty "Color", mColor, cPropDefault_Color
    PropBag.WriteProperty "SliderOptionsAvailable", mSliderOptionsAvailable, cPropDefault_SliderOptionsAvailable
    PropBag.WriteProperty "FixedPaletteControlVisible", mFixedPaletteControlVisible, cPropDefault_FixedPaletteControlVisible
    PropBag.WriteProperty "ColorSystemControlVisible", mColorSystemControlVisible, cPropDefault_ColorSystemControlVisible
    PropBag.WriteProperty "FixedPalette", mFixedPalette, cPropDefault_FixedPalette
    PropBag.WriteProperty "SliderParameter", mSliderParameter, cPropDefault_SliderParameter
    PropBag.WriteProperty "SliderParameterComboWidth", mSliderParameterComboWidth, cPropDefault_SliderParameterComboWidth
    PropBag.WriteProperty "ColorSystem", mColorSystem, cPropDefault_ColorSystem
    PropBag.WriteProperty "BackColor", mBackColor, cPropDefault_BackColor
End Sub

Private Sub picPalette_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mSelectingColor Then
        If PixelIsInPalette(X, Y) Then
            mPureColor = EnsurePrimary(GetPaletteColor(X, Y))
            If mSliderParameter = cdParameterLuminance Then
                ColorRGBToCurrentColorSystem mPureColor, mH, 0&, mS
            ElseIf mSliderParameter = cdParameterHue Then
                ColorRGBToCurrentColorSystem mPureColor, 0&, mL, mS
            ElseIf mSliderParameter = cdParameterSaturation Then
                ColorRGBToCurrentColorSystem mPureColor, mH, mL, 0&
            ElseIf mSliderParameter = cdParameterRed Then
                mG = (mPureColor \ 256) And 255
                mB = (mPureColor \ 65536) And 255
            ElseIf mSliderParameter = cdParameterGreen Then
                mR = mPureColor And 255
                mB = (mPureColor \ 65536) And 255
            ElseIf mSliderParameter = cdParameterBlue Then
                mR = mPureColor And 255
                mG = (mPureColor \ 256) And 255
            End If
            mClickingPalette = True
            SetColor GetShadedColor
            mClickingPalette = False
            PointerVisible = False
        Else
            mSelectionFromOutside = True
            GetXYSameAngleInsideCircle X, Y
            mPureColor = EnsurePrimary(GetPaletteColor(X, Y))
            If mSliderParameter = cdParameterLuminance Then
                ColorRGBToCurrentColorSystem mPureColor, mH, 0&, mS
            ElseIf mSliderParameter = cdParameterHue Then
                ColorRGBToCurrentColorSystem mPureColor, 0&, mL, mS
            ElseIf mSliderParameter = cdParameterSaturation Then
                ColorRGBToCurrentColorSystem mPureColor, mH, mL, 0&
            ElseIf mSliderParameter = cdParameterRed Then
                mG = (mPureColor \ 256) And 255
                mB = (mPureColor \ 65536) And 255
            ElseIf mSliderParameter = cdParameterGreen Then
                mR = mPureColor And 255
                mB = (mPureColor \ 65536) And 255
            ElseIf mSliderParameter = cdParameterBlue Then
                mR = mPureColor And 255
                mG = (mPureColor \ 256) And 255
            End If
            mClickingPalette = True
            SetColor GetShadedColor
            mClickingPalette = False
            SetPointer X, Y
            PointerVisible = True
            mSelectionFromOutside = False
        End If
    End If
End Sub

Private Sub picPalette_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mMouseDown Then
        mMouseDown = False
        RaiseEvent MouseUp(Button, Shift, X, Y)
    End If
    If mSelectingColor Then
        picPalette_MouseMove Button, Shift, X, Y
        ClipCursor ByVal 0&
        mSelectingColor = False
        If PixelIsInPalette(X, Y) Then SetPointer X, Y
        PointerVisible = True
    End If
End Sub


Public Property Get Color() As OLE_COLOR
Attribute Color.VB_Description = "Returns/sets the current color."
Attribute Color.VB_UserMemId = 0
    Color = mColor
End Property

Public Property Let Color(ByVal nValue As OLE_COLOR)
    If Not IsValidOLE_COLOR(nValue) Then Err.Raise 380: Exit Property
    mChangingParameter = True
    TranslateColor nValue, 0, nValue
    SetColor nValue
    mChangingParameter = False
End Property

Private Function SetColor(Value As Long) As Boolean
    Dim iPrev As Long
    Dim iColor As Long
    Dim iH1 As Double
    Dim iL1 As Double
    Dim iS1 As Double
    Dim iRGB As RGBQuad
    
    If Value = -1 Then Exit Function
    
    iPrev = mColor
    mColor = Value
    If (mColor <> iPrev) Or mChangingColorSystemOrInitializing Then
        mSettingColor = True
        TranslateColor mColor, 0, iColor
        CopyMemory iRGB, iColor, 4
        If (Not mSliderParameter = cdParameterRed) Or mChangingParameter Then
            mR = iRGB.R
        End If
        If (Not mSliderParameter = cdParameterGreen) Or mChangingParameter Then
            mG = iRGB.G
        End If
        If (Not mSliderParameter = cdParameterBlue) Or mChangingParameter Then
            mB = iRGB.B
        End If
        ColorRGBToCurrentColorSystem iColor, iH1, iL1, iS1
        If (ColorCurrentColorSystemToRGB(iH1, iL1, iS1) <> iPrev) Or mChangingColorSystemOrInitializing Then
            If Not (mChangingHue Or mChangingLuminance Or mChangingSaturation) Then
                If (Not mSliderParameter = cdParameterSaturation) Or mChangingColorSystemOrInitializing Or mChangingParameter Then
                    mS = iS1
                End If
                'If Not ((iH1 = 160) And (mS = 0)) Then
                If Not (mSliderParameter = cdParameterHue) Or mChangingColorSystemOrInitializing Or mChangingParameter Then
                    mH = iH1
                End If
                If mH = mH_Max Then mH = 0
                If Not (mSliderParameter = cdParameterLuminance) Or mChangingColorSystemOrInitializing Or mChangingParameter Then
                    If Not (mChangingShade And ((mSliderParameter = cdParameterRed) Or (mSliderParameter = cdParameterGreen) Or (mSliderParameter = cdParameterBlue)) And (mColorSystem = cdColorSystemHSV)) Then
                        mL = iL1
                    End If
                End If
                'End If
            End If
            If (Not mChangingShade) And (Not mClickingPalette) Then
                If mSliderParameter = cdParameterLuminance Then
                    SliderValue = mL_Max - mL
                    mPureColor = ColorCurrentColorSystemToRGB(mH, mL_Fixed, mS)
                ElseIf mSliderParameter = cdParameterHue Then
                    SliderValue = mH_Max - mH
                    mPureColor = ColorCurrentColorSystemToRGB(mH_Fixed, mL, mS)
                ElseIf mSliderParameter = cdParameterSaturation Then
                    SliderValue = mS_Max - mS
                    mPureColor = ColorCurrentColorSystemToRGB(mH, mL, mS_Fixed)
                Else
                    If mSliderParameter = cdParameterRed Then
                        SliderValue = 255 - mR
                        mPureColor = RGB(0, mG, mB)
                    ElseIf mSliderParameter = cdParameterGreen Then
                        SliderValue = 255 - mG
                        mPureColor = RGB(mR, 0, mB)
                    ElseIf mSliderParameter = cdParameterBlue Then
                        SliderValue = 255 - mB
                        mPureColor = RGB(mR, mG, 0)
                    End If
                End If
            End If
            DrawShades
            ShowSelectedColor
        End If
        If Not tmrDraw.Enabled Then
            If Not mClickingPalette Then
                If (mSliderParameter = cdParameterLuminance) Or (mSliderParameter = cdParameterSaturation) Then
                    If (Not mFixedPalette) Then
                        DrawPalette
                    End If
                Else
                    DrawPalette
            End If
            End If
        End If
        mSettingColor = False
        If mInitialized Then
            If mRaiseEvents Then RaiseEvent Change
        End If
        If mInitialized Then PropertyChanged "Color"
    End If
End Function

Private Property Get SliderValue() As Long
    SliderValue = mSliderValue
End Property

Private Property Let SliderValue(ByVal nValue As Long)
    If mSliderValue <> nValue Then
        If nValue > mSliderMax Then nValue = mSliderMax
        If nValue < mSliderMin Then nValue = mSliderMin
        If mSliderValue <> nValue Then
            mSliderValue = nValue
            DrawSliderGrip
            DoSliderChange
        End If
    End If
End Property

Public Property Get H() As Integer
Attribute H.VB_Description = "Returns/sets the 'Hue' component of the color."
    H = mH
End Property

Public Property Let H(ByVal nValue As Integer)
    If nValue <> mH Then
        If (nValue < 0) Or (nValue > mH_Max) Then
            Err.Raise 380, TypeName(Me)
            Exit Property
        End If
        mH = nValue
        If mH = mH_Max Then mH = 0
        mChangingHue = True
        mChangingParameter = True
        If Not SetColor(ColorCurrentColorSystemToRGB(mH, mL, mS)) Then
            If mSliderParameter = cdParameterHue Then
                SliderValue = mH_Max - mH
            End If
            DrawPalette
            DrawShades
            ShowSelectedColor
            RaiseEvent Change
        End If
        mChangingParameter = False
        mChangingHue = False
    End If
End Property


Public Property Get L() As Integer
Attribute L.VB_Description = "Returns/sets the 'Luminance' or 'Value' (depending on the color system , HSL or HSV) component of the color."
    L = mL
End Property

Public Property Let L(ByVal nValue As Integer)
    If nValue <> mL Then
        If (nValue < 0) Or (nValue > mL_Max) Then
            Err.Raise 380, TypeName(Me)
            Exit Property
        End If
        mL = nValue
        mChangingLuminance = True
        mChangingParameter = True
        If Not SetColor(ColorCurrentColorSystemToRGB(mH, mL, mS)) Then
            If mSliderParameter = cdParameterLuminance Then
                SliderValue = mL_Max - mL
            End If
            DrawPalette
            DrawShades
            ShowSelectedColor
            RaiseEvent Change
        End If
        mChangingParameter = False
        mChangingLuminance = False
    End If
End Property


Public Property Get V() As Integer
Attribute V.VB_Description = "Alias for 'L' property."
    V = mL
End Property

Public Property Let V(ByVal nValue As Integer)
    L = nValue
End Property


Public Property Get S() As Integer
Attribute S.VB_Description = "Returns/sets the 'Saturation' component of the color."
    S = mS
End Property

Public Property Let S(ByVal nValue As Integer)
    If nValue <> mS Then
        If (nValue < 0) Or (nValue > mS_Max) Then
            Err.Raise 380, TypeName(Me)
            Exit Property
        End If
        mS = nValue
        mChangingSaturation = True
        mChangingParameter = True
        If Not SetColor(ColorCurrentColorSystemToRGB(mH, mL, mS)) Then
            If mSliderParameter = cdParameterSaturation Then
                SliderValue = mS_Max - mS
            End If
            DrawPalette
            DrawShades
            ShowSelectedColor
            RaiseEvent Change
        End If
        mChangingParameter = False
        mChangingSaturation = False
    End If
End Property


Public Property Get R() As Integer
Attribute R.VB_Description = "Returns/sets the 'Red' component of the color."
    R = mR
End Property

Public Property Let R(ByVal nValue As Integer)
    If nValue <> mR Then
        If (nValue < 0) Or (nValue > 255) Then
            Err.Raise 380, TypeName(Me)
            Exit Property
        End If
        mR = nValue
        mChangingParameter = True
        SetColor RGB(mR, mG, mB)
        mChangingParameter = False
        DrawShades
    End If
End Property


Public Property Get G() As Integer
Attribute G.VB_Description = "Returns/sets the 'Green' component of the color."
    G = mG
End Property

Public Property Let G(ByVal nValue As Integer)
    If nValue <> mG Then
        If (nValue < 0) Or (nValue > 255) Then
            Err.Raise 380, TypeName(Me)
            Exit Property
        End If
        mG = nValue
        mChangingParameter = True
        SetColor RGB(mR, mG, mB)
        mChangingParameter = False
        DrawShades
    End If
End Property


Public Property Get B() As Integer
Attribute B.VB_Description = "Returns/sets the 'Blue' component of the color."
    B = mB
End Property

Public Property Let B(ByVal nValue As Integer)
    If nValue <> mB Then
        If (nValue < 0) Or (nValue > 255) Then
            Err.Raise 380, TypeName(Me)
            Exit Property
        End If
        mB = nValue
        mChangingParameter = True
        SetColor RGB(mR, mG, mB)
        mChangingParameter = False
        DrawShades
    End If
End Property

Private Sub DrawShades()
    Dim iY As Long
    Dim iColor As Long
    Dim iH1 As Double
    Dim iS1 As Double
    Dim iL1 As Double
    Dim iHeight As Long
    Dim iWidth As Long
    Dim iLng As Long
    
    If mDrawPending Then Exit Sub
    If Not mDrawEnabled Then Exit Sub
    If mChangingColorSystemOrInitializing Then Exit Sub
    
    picShades.Cls
    iHeight = picShades.ScaleHeight - 2
    iWidth = picShades.ScaleWidth - 1
    ColorRGBToCurrentColorSystem mPureColor, iH1, iL1, iS1
    
    If mSliderParameter = cdParameterLuminance Then
        If mColorSystem = cdColorSystemHSV Then
            For iY = 0 To iHeight - IIf(mRoundedBoxes, 0, 1)
                iColor = ColorHSVToRGB(iH1, iY / iHeight * mL_Max, iS1)
                picShades.Line (IIf(mRoundedBoxes, 0, 1), iHeight - iY)-(iWidth, iHeight - iY), iColor
            Next iY
        Else
            For iY = 0 To iHeight - IIf(mRoundedBoxes, 0, 1)
                iColor = ColorHLSToRGB(iH1, iY / iHeight * mL_Max, iS1)
                picShades.Line (IIf(mRoundedBoxes, 0, 1), iHeight - iY)-(iWidth, iHeight - iY), iColor
            Next iY
        End If
    ElseIf mSliderParameter = cdParameterHue Then
        If mColorSystem = cdColorSystemHSV Then
            For iY = 0 To iHeight - IIf(mRoundedBoxes, 0, 1)
                If mFixedPalette Then
                    iColor = ColorHSVToRGB(iY / iHeight * mH_Max, CDbl(mL_Fixed), CDbl(mS_Fixed))
                Else
                    iColor = ColorHSVToRGB(iY / iHeight * mH_Max, iL1, iS1)
                End If
                picShades.Line (IIf(mRoundedBoxes, 0, 1), iHeight - iY)-(iWidth, iHeight - iY), iColor
            Next iY
        Else
            For iY = 0 To iHeight - IIf(mRoundedBoxes, 0, 1)
                If mFixedPalette Then
                    iColor = ColorHLSToRGB(iY / iHeight * mH_Max, mL_Fixed, mS_Fixed)
                Else
                    If Round(iY / iHeight * mH_Max) = 160 And (iS1 = 0) Then
                        iColor = 0
                    Else
                        iColor = ColorHLSToRGB(iY / iHeight * mH_Max, iL1, iS1) 'aca
                    End If
                End If
                picShades.Line (IIf(mRoundedBoxes, 0, 1), iHeight - iY)-(iWidth, iHeight - iY), iColor
            Next iY
        End If
    ElseIf mSliderParameter = cdParameterSaturation Then
        If mColorSystem = cdColorSystemHSV Then
            For iY = 0 To iHeight - IIf(mRoundedBoxes, 0, 1)
                iColor = ColorHSVToRGB(iH1, iL1, iY / iHeight * mS_Max)
                picShades.Line (IIf(mRoundedBoxes, 0, 1), iHeight - iY)-(iWidth, iHeight - iY), iColor
            Next iY
        Else
            For iY = 0 To iHeight - IIf(mRoundedBoxes, 0, 1)
                If iY = 0 Then
                    iLng = 1
                Else
                    iLng = iY
                End If
                iColor = ColorHLSToRGB(iH1, iL1, iLng / iHeight * mS_Max)
                picShades.Line (IIf(mRoundedBoxes, 0, 1), iHeight - iY)-(iWidth, iHeight - iY), iColor
            Next iY
        End If
    ElseIf mSliderParameter = cdParameterRed Then
        For iY = 0 To iHeight - IIf(mRoundedBoxes, 0, 1)
            If mFixedPalette Then
                iColor = RGB(iY / iHeight * 255, mG, mB)
            Else
                iColor = RGB(iY / iHeight * 255, 0, 0)
            End If
            picShades.Line (IIf(mRoundedBoxes, 0, 1), iHeight - iY)-(iWidth, iHeight - iY), iColor
        Next iY
    ElseIf mSliderParameter = cdParameterGreen Then
        For iY = 0 To iHeight - IIf(mRoundedBoxes, 0, 1)
            If mFixedPalette Then
                iColor = RGB(mR, iY / iHeight * 255, mB)
            Else
                iColor = RGB(0, iY / iHeight * 255, 0)
            End If
            picShades.Line (IIf(mRoundedBoxes, 0, 1), iHeight - iY)-(iWidth, iHeight - iY), iColor
        Next iY
    ElseIf mSliderParameter = cdParameterBlue Then
        For iY = 0 To iHeight - IIf(mRoundedBoxes, 0, 1)
            If mFixedPalette Then
                iColor = RGB(mR, mG, iY / iHeight * 255)
            Else
                iColor = RGB(0, 0, iY / iHeight * 255)
            End If
            picShades.Line (IIf(mRoundedBoxes, 0, 1), iHeight - iY)-(iWidth, iHeight - iY), iColor
        Next iY
    End If
    picShades.Refresh
End Sub

Private Function GetShadedColor() As Long
    If (mSliderParameter = cdParameterLuminance) Or (mSliderParameter = cdParameterHue) Or (mSliderParameter = cdParameterSaturation) Then
        GetShadedColor = ColorCurrentColorSystemToRGB(mH, mL, mS)
    Else
        GetShadedColor = RGB(mR, mG, mB)
    End If
End Function

Private Function GetPaletteColor(ByVal X As Single, ByVal Y As Single) As Long
    Dim i As Long
    
    'Y = Y + 1
    If Y >= mBMPHeight Then Y = mBMPHeight - 1
    If X >= mBMPWidth Then Y = mBMPWidth - 1
    X = Int(X)
    Y = Int(Y)
    i = (mBMPHeight - Y - 1) * mBMPHeight
    i = i + X
    i = i * 4
    
    GetPaletteColor = RGB(mPixelsBytes(i + 2), mPixelsBytes(i + 1), mPixelsBytes(i))
End Function

Private Sub ShowSelectedColor()
    Dim iX As Single
    Dim iY As Single
    Dim iX2 As Long
    Dim iY2 As Long
    Dim iFound As Boolean
    Dim iR1 As Long
    Dim iG1 As Long
    Dim iB1 As Long
    Dim iColor As Long
    Dim iTolerance  As Long
    Dim c As Long
    Dim iP1 As Long
    Dim iP2 As Long
    Dim iP1Max As Long
    Dim iP2Max As Long
    Dim iRGB As RGBQuad
    Dim iListOfPossible_X() As Long
    Dim iListOfPossible_Y() As Long
    Dim iUb As Long
    Dim iCount As Long
    Dim iNearest_Index As Long
    Dim iNearest_Distance As Single
    Dim iDistance As Single
    
    If mDrawPending Then Exit Sub
    If Not mDrawEnabled Then Exit Sub
    If mClickingPalette Or mChangingShade Then Exit Sub
    
    If mH_Max = 0 Then SetMaxAndFixedvalues
    If mSliderParameter = cdParameterLuminance Then
        iP1 = mS ' iP1 is radius
        iP2 = mH ' iP2 is angle
        iP1Max = mS_Max
        iP2Max = mH_Max
    ElseIf mSliderParameter = cdParameterHue Then
        iP1 = mS
        iP2 = mL
        iP1Max = mS_Max
        iP2Max = mL_Max
    ElseIf mSliderParameter = cdParameterSaturation Then
        iP1 = mL
        iP2 = mH
        iP1Max = mL_Max
        iP2Max = mH_Max
    ElseIf mSliderParameter = cdParameterRed Then
        iP1 = mB
        iP2 = mG
        iP1Max = 255
        iP2Max = 255
    ElseIf mSliderParameter = cdParameterGreen Then
        iP1 = mB
        iP2 = mR
        iP1Max = 255
        iP2Max = 255
    ElseIf mSliderParameter = cdParameterBlue Then
        iP1 = mG
        iP2 = mR
        iP1Max = 255
        iP2Max = 255
    End If
    If mStyleBox Then
        SetPointer CSng(mPaletteWidth / iP2Max * iP2), CSng(mPaletteHeight - mPaletteHeight / iP1Max * iP1)
    Else
        iX = (iP1 * Cos(Pi / 180 * (1 - iP2 / iP2Max) * 360)) / iP1Max * mRadius + mCx
        iY = (iP1 * sIn(Pi / 180 * (1 - iP2 / iP2Max) * 360)) / iP1Max * mRadius + mCy
        If mColor = cPropDefault_Color Then
            iX = mCx
            iY = mCy
            SetPointer iX, iY
            Exit Sub
        End If
        iX = Int(iX)
        iY = Int(iY)
        If Not PixelIsInPalette(iX, iY) Then
            iX = iX + 1
            iY = iY + 1
            If Not PixelIsInPalette(iX, iY) Then
                iX = iX + 1
                iY = iY + 1
            End If
        End If
        If Not PixelIsInPalette(iX, iY) Then
            iX = (iP1 * Cos(Pi * 2 / 360 * (360 - iP2 / iP2Max * 360))) / iP1Max * mDiameter * 0.99 / 2 + mCx
            iY = (iP1 * sIn(Pi * 2 / 360 * (360 - iP2 / iP2Max * 360))) / iP1Max * mDiameter * 0.99 / 2 + mCy
            If Not PixelIsInPalette(iX, iY) Then
                iX = (iP1 * Cos(Pi * 2 / 360 * (360 - iP2 / iP2Max * 360))) / iP1Max * mDiameter * 0.98 / 2 + mCx
                iY = (iP1 * sIn(Pi * 2 / 360 * (360 - iP2 / iP2Max * 360))) / iP1Max * mDiameter * 0.98 / 2 + mCy
            End If
        End If
        If PixelIsInPalette(iX, iY) Then
    '        For iX2 = -10 To 10
    '            For iY2 = -10 To 10
    '                If PixelIsInPalette(iX + iX2, iY + iY2) Then
    '                    If GetPaletteColor(iX, iY) = mPureColor Then
    '                        iFound = True
    '                        Exit For
    '                    End If
    '                End If
    '            Next iY2
    '            If iFound Then Exit For
    '        Next iX2
    '        If Not iFound Then
            For iTolerance = 0 To 10
                iUb = 100
                ReDim iListOfPossible_X(iUb)
                ReDim iListOfPossible_Y(iUb)
                iCount = -1
                For iX2 = -10 To 10
                    For iY2 = -10 To 10
                        If PixelIsInPalette(iX + iX2, iY + iY2) Then
                            iColor = GetPaletteColor(iX + iX2, iY + iY2)
                            CopyMemory iRGB, iColor, 4
                            If (Abs(mR - iRGB.R) + Abs(mG - iRGB.G) + Abs(mB - iRGB.B)) <= iTolerance Then
                                iFound = True
                                iCount = iCount + 1
                                If iCount > iUb Then
                                    iUb = iUb + 100
                                    ReDim Preserve iListOfPossible_X(iUb)
                                    ReDim Preserve iListOfPossible_Y(iUb)
                                End If
                                iListOfPossible_X(iCount) = iX2
                                iListOfPossible_Y(iCount) = iY2
                                'Exit For
                            End If
                        End If
                    Next iY2
                    'If iFound Then Exit For
                Next iX2
                If iFound Then
                    iNearest_Distance = mRadius
                    For c = 0 To iCount
                        iDistance = Sqr(iListOfPossible_X(c) ^ 2 + iListOfPossible_Y(c) ^ 2)
                        If iDistance < iNearest_Distance Then
                            iNearest_Distance = iDistance
                            iNearest_Index = c
                        End If
                    Next c
                    iX2 = iListOfPossible_X(iNearest_Index)
                    iY2 = iListOfPossible_Y(iNearest_Index)
                    Exit For
                End If
            Next iTolerance
    '        End If
            
            If iFound Then
                iX = iX + iX2
                iY = iY + iY2
            End If
            
            SetPointer iX, iY
        End If
    End If
End Sub

Private Sub GetXYSameAngleInsideCircle(X As Single, Y As Single)
    Dim H As Single
    Dim a As Single
    Dim B As Single
    Dim iAngle As Single
    Dim iSin As Single
    Dim iHorz As Single
    Dim iVert As Single
    Dim iX2 As Single
    Dim iY2 As Single
    Dim R As Single
    Dim iRadius As Single
    
    iHorz = X - mCx
    iVert = Y - mCy
    If iHorz = 0 Then
        
        If Y < mCy Then
            iAngle = 270 * Pi / 180 ' angle is hue
        Else
            iAngle = 90 * Pi / 180
        End If
    Else
        iAngle = Atn(iVert / iHorz)
    End If
    
    iRadius = mRadius * 1.02
    iX2 = Cos(iAngle) * iRadius + mCx
    iY2 = sIn(iAngle) * iRadius + mCy
    If iHorz < 0 Then
        iX2 = mCx - (iX2 - mCx)
        iY2 = mCy - (iY2 - mCy)
    End If
    
    Do Until PixelIsInPalette(iX2, iY2)
        iRadius = iRadius * 0.9999
        iX2 = Cos(iAngle) * iRadius + CSng(mCx)
        iY2 = sIn(iAngle) * iRadius + CSng(mCy)
        If iHorz < 0 Then
            iX2 = mCx - (iX2 - mCx)
            iY2 = mCy - (iY2 - mCy)
        End If
    Loop
    
    X = iX2
    Y = iY2
End Sub

Public Property Get ColorHex() As String
Attribute ColorHex.VB_Description = "Returns/sets the current color en hexadecimal format. The format is in VB format for haxadecimal colors values."
    ColorHex = Hex(mColor)
    If Len(ColorHex) < 6 Then ColorHex = String$(6 - Len(ColorHex), "0") & ColorHex
    ColorHex = "&H" & ColorHex & "&"
End Property

Private Function DistanceFromCenter(X As Single, Y As Single) As Single
    DistanceFromCenter = Sqr((X - mCx) ^ 2 + (Y - mCy) ^ 2)
End Function

Private Property Let PointerVisible(ByVal nValue As Boolean)
    linPointer(0).Visible = nValue
    linPointer(1).Visible = nValue
    linPointer(2).Visible = nValue
    linPointer(3).Visible = nValue
End Property

Private Sub SetPointer(Optional X As Single = -1, Optional Y As Single)
    Dim c As Long
    Dim iPointerColor As Long
    Dim iX2 As Single
    Dim iY2 As Single
    Dim iDrawMode As Long
    Dim iColorBrightness As Long
    Dim iColor As Long
    
    If X <> -1 Then
        mPointerX = X
        mPointerY = Y
    End If
    
    iX2 = mPointerX
    iY2 = mPointerY
'    If iX2 < mCx Then
'        iX2 = iX2 + 1
'    Else
'        iX2 = iX2 - 1
'    End If
'    If iY2 < mCy Then
'        iY2 = iY2 + 1
'    Else
'        iY2 = iY2 - 1
'    End If
    
    If PixelIsInPalette(iX2, iY2) Then
        iColor = picPalette.Point(iX2, iY2)
    Else
        iColor = mColor
    End If

    iColorBrightness = GetColorBrightness(iColor)
    If iColorBrightness > 110 Then
        If (iColorBrightness > 200) Then
            iPointerColor = vbWhite
            iDrawMode = vbMaskPenNot
        Else
            iPointerColor = vbBlack
            iDrawMode = vbCopyPen
        End If
    Else
        iPointerColor = vbWhite
        If (mRadius - DistanceFromCenter(iX2, iY2) < ((linPointer(0).X2 - linPointer(0).X1) * 1.5)) And (Not (mStyleBox)) Then
            iDrawMode = vbMaskPenNot
        Else
            If (iColorBrightness < 50) Then
                iDrawMode = vbMaskPenNot
            Else
                iDrawMode = vbCopyPen
            End If
        End If
    End If
    
    For c = 0 To 3
        linPointer(c).BorderColor = iPointerColor
        linPointer(c).DrawMode = iDrawMode
    Next c
    
    linPointer(0).X1 = mPointerX - 14 * 15 / Screen.TwipsPerPixelX - 0.5
    linPointer(0).X2 = linPointer(0).X1 + 8 * 15 / Screen.TwipsPerPixelX
    linPointer(0).Y1 = mPointerY - 0.5
    linPointer(0).Y2 = mPointerY - 0.5

    linPointer(1).X1 = mPointerX + 14 * 15 / Screen.TwipsPerPixelX - 0.5
    linPointer(1).X2 = linPointer(1).X1 - 8 * 15 / Screen.TwipsPerPixelX
    linPointer(1).Y1 = mPointerY - 0.5
    linPointer(1).Y2 = mPointerY - 0.5

    linPointer(2).Y1 = mPointerY - 14 * 15 / Screen.TwipsPerPixelY - 0.5
    linPointer(2).Y2 = linPointer(2).Y1 + 8 * 15 / Screen.TwipsPerPixelY
    linPointer(2).X1 = mPointerX - 0.5
    linPointer(2).X2 = mPointerX - 0.5

    linPointer(3).Y1 = mPointerY + 14 * 15 / Screen.TwipsPerPixelY - 0.5
    linPointer(3).Y2 = linPointer(3).Y1 - 8 * 15 / Screen.TwipsPerPixelY
    linPointer(3).X1 = mPointerX - 0.5
    linPointer(3).X2 = mPointerX - 0.5

End Sub


Public Property Get SliderOptionsAvailable() As CDSliderOptionsAvailableConstants
Attribute SliderOptionsAvailable.VB_Description = "Returns/sets a value that determines which parameters for the slider the user will have available to choose from."
    SliderOptionsAvailable = mSliderOptionsAvailable
End Property

Public Property Let SliderOptionsAvailable(ByVal nValue As CDSliderOptionsAvailableConstants)
    Dim iPrev As CDSliderOptionsAvailableConstants
    
    If nValue <> mSliderOptionsAvailable Then
        If (nValue < cdSliderOptionsNone) Or (nValue > cdSliderOptionsAll) Then
            Err.Raise 380, TypeName(Me)
            Exit Property
        End If
        iPrev = mSliderOptionsAvailable
        mSliderOptionsAvailable = nValue
        If mInitialized Then PropertyChanged "SliderOptionsAvailable"
        If (mSliderOptionsAvailable <> cdSliderOptionsNone) And (iPrev = cdSliderOptionsNone) Or (mSliderOptionsAvailable = cdSliderOptionsNone) And (iPrev <> cdSliderOptionsNone) Then
            LoadcboSliderParameter
            UserControl_Resize
        Else
            LoadcboSliderParameter
        End If
        SetPicShades
    End If
End Property


Public Property Get FixedPaletteControlVisible() As Boolean
Attribute FixedPaletteControlVisible.VB_Description = "Returns/sets a value that determines if the control for changing the fixed setting of the palette is available."
    FixedPaletteControlVisible = mFixedPaletteControlVisible
End Property

Public Property Let FixedPaletteControlVisible(ByVal nValue As Boolean)
    If nValue <> mFixedPaletteControlVisible Then
        mFixedPaletteControlVisible = nValue
        chkFixedPalette.Visible = mFixedPaletteControlVisible And (Not mStyleBox)
        If mInitialized Then PropertyChanged "FixedPaletteControlVisible"
        UserControl_Resize
    End If
End Property


Public Property Get ColorSystemControlVisible() As Boolean
Attribute ColorSystemControlVisible.VB_Description = "Returns/sets a value that determines if the user is able to change the color system."
    ColorSystemControlVisible = mColorSystemControlVisible
End Property

Public Property Let ColorSystemControlVisible(ByVal nValue As Boolean)
    If nValue <> mColorSystemControlVisible Then
        mColorSystemControlVisible = nValue
        UserControl_Resize
        picColorSystem.Visible = mColorSystemControlVisible And (Not mStyleBox)
        cboColorSystem.Visible = mColorSystemControlVisible And (Not mStyleBox)
        If mInitialized Then PropertyChanged "ColorSystemControlVisible"
    End If
End Property


Public Property Get FixedPalette() As Boolean
Attribute FixedPalette.VB_Description = "Returns/sets a value that determines if the color palette (or the slidder in some configurations) colors change with the setting of the third partameter."
    FixedPalette = mFixedPalette
End Property

Public Property Let FixedPalette(ByVal nValue As Boolean)
    If nValue <> mFixedPalette Then
        mFixedPalette = nValue
        If mInitialized Then PropertyChanged "FixedPalette"
        chkFixedPalette.Value = Abs(CLng(mFixedPalette))
        DrawPalette
        DrawShades
        If mRaiseEvents Then RaiseEvent FixedPaletteChange
    End If
End Property


Public Property Get SliderParameter() As CDSliderParameterConstants
Attribute SliderParameter.VB_Description = "Returns/sets a value that determines the slidder parameter."
    SliderParameter = mSliderParameter
End Property

Public Property Let SliderParameter(ByVal nValue As CDSliderParameterConstants)
    If (nValue <> mSliderParameter) Then
        mSliderParameter = nValue
        If mInitialized Then PropertyChanged "SliderParameter"
        If SliderOptionsAvailable <> cdSliderOptionsNone Then
            If Not IsItemDataInList(cboSliderParameter, mSliderParameter) Then
                If mSliderParameter = cdParameterHue Then
                    SliderOptionsAvailable = cdSliderOptionsHueLumAndSat
                Else
                    SliderOptionsAvailable = cdSliderOptionsAll
                End If
            End If
        End If
        SetSliderParameter
        If mRaiseEvents Then RaiseEvent SliderParameterChange
    ElseIf (cboSliderParameter.ListIndex = -1) Then
        SetSliderParameter
    End If
End Property

Private Sub SetSliderParameter()
    mSettingSlider = True
    SelectInListByItemData cboSliderParameter, mSliderParameter
    StorePaletteColors
    If mSliderParameter = cdParameterLuminance Then
        mSliderMax = mL_Max
        SliderValue = mL_Max - mL
        mPureColor = ColorCurrentColorSystemToRGB(mH, mL_Fixed, mS)
        mRadialParameter = cdParameterSaturation
        mAxialParameter = cdParameterHue
    ElseIf mSliderParameter = cdParameterHue Then
        mSliderMax = mH_Max
        SliderValue = mH_Max - mH
        mPureColor = ColorCurrentColorSystemToRGB(mH_Fixed, mL, mS)
        mRadialParameter = cdParameterSaturation
        mAxialParameter = cdParameterLuminance
    ElseIf mSliderParameter = cdParameterSaturation Then
        mSliderMax = mS_Max
        SliderValue = mS_Max - mS
        mPureColor = ColorCurrentColorSystemToRGB(mH, mL, mS_Fixed)
        mRadialParameter = cdParameterLuminance
        mAxialParameter = cdParameterHue
    ElseIf mSliderParameter = cdParameterRed Then
        mSliderMax = 255
        SliderValue = 255 - mR
        mPureColor = RGB(0, mG, mB)
        mRadialParameter = cdParameterBlue
        mAxialParameter = cdParameterGreen
    ElseIf mSliderParameter = cdParameterGreen Then
        mSliderMax = 255
        SliderValue = 255 - mG
        mPureColor = RGB(mR, 0, mB)
        mRadialParameter = cdParameterBlue
        mAxialParameter = cdParameterRed
    Else ' cdParameterBlue
        mSliderMax = 255
        SliderValue = 255 - mB
        mPureColor = RGB(mR, mG, 0)
        mRadialParameter = cdParameterGreen
        mAxialParameter = cdParameterRed
    End If
    mSettingSlider = False
    DrawPalette
    DrawShades
    ShowSelectedColor
End Sub


Public Property Get ColorSystem() As CDColorSystemConstants
Attribute ColorSystem.VB_Description = "Returns/sets the color sytem, HSV or HSL."
    ColorSystem = mColorSystem
End Property

Public Property Let ColorSystem(ByVal nValue As CDColorSystemConstants)
    Dim iColor As Long
    Dim c As Long
    
    If nValue <> mColorSystem Then
        iColor = mColor
        mColor = -1
        mColorSystem = nValue
        ColorRGBToCurrentColorSystem iColor, mH, mL, mS
        SetMaxAndFixedvalues
        StorePaletteColors
        mChangingColorSystemOrInitializing = True
        For c = 0 To cboSliderParameter.ListCount - 1
            If cboSliderParameter.ItemData(c) = cdParameterLuminance Then
                If mColorSystem = cdColorSystemHSV Then
                    If mStyleBox Then
                        cboSliderParameter.List(c) = Left$(mCaptionVal, 1)
                    Else
                        cboSliderParameter.List(c) = mCaptionVal
                    End If
                Else
                    If mStyleBox Then
                        cboSliderParameter.List(c) = Left$(mCaptionLum, 1)
                    Else
                        cboSliderParameter.List(c) = mCaptionLum
                    End If
                End If
                Exit For
            End If
        Next c
        SetSliderParameter
        SetColor iColor
        cboColorSystem.ListIndex = mColorSystem
        mChangingColorSystemOrInitializing = False
        DrawPalette
        DrawShades
        If mInitialized Then PropertyChanged "ColorSystem"
        If mRaiseEvents Then RaiseEvent ColorSystemChange
    End If
End Property


Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color."
    BackColor = mBackColor
End Property

Public Property Let BackColor(ByVal nValue As OLE_COLOR)
    If nValue <> mBackColor Then
        If Not IsValidOLE_COLOR(nValue) Then Err.Raise 380: Exit Property
        mBackColor = nValue
        SetBackColor
        mDiameter = 0
        init
    End If
End Property


Public Property Get SliderWide() As CDYesNoAutoConstants
Attribute SliderWide.VB_Description = "Returns/sets a value that determines the width of the slider control between wide or narrow."
    SliderWide = mSliderWide
End Property

Public Property Let SliderWide(ByVal nValue As CDYesNoAutoConstants)
    If nValue <> mSliderWide Then
        mSliderWide = nValue
        picSlider.Cls
        picShades.Cls
        UserControl_Resize
        PropertyChanged "SliderWide"
    End If
End Property


Public Property Get HideLabels() As Boolean
Attribute HideLabels.VB_Description = "Returns/sets a value that determines if the labels are visible."
    HideLabels = mHideLabels
End Property

Public Property Let HideLabels(ByVal nValue As Boolean)
    If nValue <> mHideLabels Then
        mHideLabels = nValue
        UserControl_Resize
        PropertyChanged "HideLabels"
    End If
End Property


Public Property Get RoundedBoxes() As Boolean
Attribute RoundedBoxes.VB_Description = "Returns/sets a value that determines if some control borders are rounded."
    RoundedBoxes = mRoundedBoxes
End Property

Public Property Let RoundedBoxes(ByVal nValue As Boolean)
    If nValue <> mRoundedBoxes Then
        mRoundedBoxes = nValue
        UserControl_Resize
        PropertyChanged "RoundedBoxes"
    End If
End Property


Public Property Get Style() As CDStyleConstants
Attribute Style.VB_Description = "Returns/sets a value that determines the color palette style, between wheel or box."
Attribute Style.VB_MemberFlags = "200"
    Style = mStyle
End Property

Public Property Let Style(ByVal nValue As CDStyleConstants)
    If nValue <> mStyle Then
        mStyle = nValue
        mStyleBox = (mStyle = cdStyleBox)
        mPaletteColorsStored = False
        LoadcboSliderParameter
        UserControl_Resize
        PropertyChanged "Style"
    End If
End Property


Public Property Get SliderParameterComboWidth() As Long
    SliderParameterComboWidth = mSliderParameterComboWidth
End Property

Public Property Let SliderParameterComboWidth(ByVal nValue As Long)
    If nValue <> mSliderParameterComboWidth Then
        mSliderParameterComboWidth = nValue
        UserControl_Resize
        PropertyChanged "SliderParameterComboWidth"
    End If
End Property


Private Sub SetBackColor()
    UserControl.BackColor = mBackColor
    picSlider.BackColor = mBackColor
    picColorSystem.BackColor = mBackColor
    If GetColorBrightness(UserControl.BackColor) > 170 Then
        lblMode.ForeColor = vbWindowText
    Else
        lblMode.ForeColor = vbWindowBackground
    End If
End Sub


Private Sub SetMaxAndFixedvalues()
    If mColorSystem = cdColorSystemHSV Then
        mH_Max = 359
        mL_Max = 100
        mS_Max = 100
        mH_Fixed = 240
        mL_Fixed = mL_Max
        mS_Fixed = mS_Max
    Else
        mH_Max = 240
        mL_Max = 240
        mS_Max = 240
        mH_Fixed = 160
        mL_Fixed = 120
        mS_Fixed = mS_Max
    End If
End Sub

Private Function GetColorBrightness(ByVal nColor As Long) As Long
    Dim iRGB As RGBQuad
    
    TranslateColor nColor, 0&, nColor
    CopyMemory iRGB, nColor, 4
    GetColorBrightness = (0.2125 * iRGB.R + 0.7154 * iRGB.G + 0.0721 * iRGB.B)
End Function

Private Sub LoadcboSliderParameter()
    Dim c As Long
    Dim iCurrent As Long
    Dim iFrom As Long
    Dim iTo As Long
    
    If mSliderOptionsAvailable = cdSliderOptionsNone Then Exit Sub
    
    If cboSliderParameter.ListIndex > -1 Then
        iCurrent = cboSliderParameter.ItemData(cboSliderParameter.ListIndex)
    Else
        iCurrent = -1
    End If
    cboSliderParameter.Clear
    If mSliderOptionsAvailable = cdSliderOptionsAll Then
        iFrom = 0
        iTo = 5
    ElseIf mSliderOptionsAvailable = cdSliderOptionsHueLumAndSat Then
        iFrom = 0
        iTo = 2
    ElseIf mSliderOptionsAvailable = cdSliderOptionsLumAndSat Then
        iFrom = 1
        iTo = 2
    End If
    For c = iFrom To iTo
        If c = cdParameterLuminance Then
            If mColorSystem = cdColorSystemHSV Then
                If mStyleBox Then
                    cboSliderParameter.AddItem Left$(mCaptionVal, 1)
                Else
                    cboSliderParameter.AddItem mCaptionVal
                End If
            Else
                If mStyleBox Then
                    cboSliderParameter.AddItem Left$(mCaptionLum, 1)
                Else
                    cboSliderParameter.AddItem mCaptionLum
                End If
            End If
        Else
            If mStyleBox Then
                cboSliderParameter.AddItem Left$(mParametersCaptions(c), 1)
            Else
                cboSliderParameter.AddItem mParametersCaptions(c)
            End If
        End If
        cboSliderParameter.ItemData(cboSliderParameter.NewIndex) = c
    Next c
    If iCurrent > -1 Then
        If Not SelectInListByItemData(cboSliderParameter, iCurrent) Then
            cboSliderParameter.ListIndex = 0
        End If
    ElseIf mSliderOptionsAvailable <> cdSliderOptionsNone Then
        SetSliderParameter
    End If
End Sub

Private Function SelectInListByItemData(nListControl As Object, nItemData As Long) As Boolean
    Dim c As Long
    
    For c = 0 To nListControl.ListCount - 1
        If nListControl.ItemData(c) = nItemData Then
            nListControl.ListIndex = c
            SelectInListByItemData = True
            Exit For
        End If
    Next c
End Function

Private Function IsItemDataInList(nListControl As Object, nItemData As Long) As Boolean
    Dim c As Long
    
    For c = 0 To nListControl.ListCount - 1
        If nListControl.ItemData(c) = nItemData Then
            IsItemDataInList = True
            Exit For
        End If
    Next c
End Function

Public Function GetCaption(nCaptionID As Long) As String
Attribute GetCaption.VB_Description = "Returns the caption pointed by the parameter."
    If nCaptionID = cdCWCaptionFixed Then
        GetCaption = chkFixedPalette.Caption
    ElseIf nCaptionID = cdCWCaptionFixedToolTipText Then
        GetCaption = chkFixedPalette.ToolTipText
    ElseIf nCaptionID = cdCWCaptionSliderParameterToolTipText Then
        GetCaption = cboSliderParameter.ToolTipText
    ElseIf nCaptionID = cdCWCaptionLum Then
        GetCaption = mCaptionLum
    ElseIf nCaptionID = cdCWCaptionVal Then
        GetCaption = mCaptionVal
    ElseIf nCaptionID = cdCWCaptionMode Then
        GetCaption = lblMode.Caption
    Else
        GetCaption = mParametersCaptions(nCaptionID)
    End If
End Function
    
Private Sub ColorRGBToCurrentColorSystem(nColorRGB As Long, nHue As Double, nLuminance As Double, nSaturation As Double)
    If mColorSystem = cdColorSystemHSL Then
        Dim iH1 As Integer
        Dim iL1 As Integer
        Dim iS1 As Integer
        
        ColorRGBToHLS nColorRGB, iH1, iL1, iS1
        nHue = CDbl(iH1)
        nLuminance = CDbl(iL1)
        nSaturation = CDbl(iS1)
        If mSelectionFromOutside Then
            If mSliderParameter = cdParameterSaturation Then
                nLuminance = 240
                nColorRGB = ColorHLSToRGB(nHue, nLuminance, nSaturation)
            End If
        End If
    Else
        ColorRGBToHSV nColorRGB, nHue, nLuminance, nSaturation
        If mSelectionFromOutside Then
            If mSliderParameter = cdParameterLuminance Then
                If nSaturation <> 100 Then
                    nSaturation = 100
                    nColorRGB = ColorHSVToRGB(nHue, nLuminance, nSaturation)
                End If
            ElseIf mSliderParameter = cdParameterHue Then
                If nSaturation <> 100 Then
                    nSaturation = 100
                    nColorRGB = ColorHSVToRGB(nHue, nLuminance, nSaturation)
                End If
            Else
                If nLuminance <> 100 Then
                    nLuminance = 100
                    nColorRGB = ColorHSVToRGB(nHue, nLuminance, nSaturation)
                End If
            End If
        End If
    End If
End Sub
    
Private Sub ColorRGBToHSV(ByVal nColorRGB As Long, nHue As Double, nValue As Double, nSaturation As Double)
'--- based on wqweto (Vlad Vissoultchev)'S code  from http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=36529
'--- based on *cool* code by Branco Medeiros (http://www.myrealbox.com/branco_medeiros)
'--- Converts an RGB value to the HSB color model. Adapted from Java.awt.pvColor.java
    Dim nTemp           As Double
    Dim lMin            As Long
    Dim LMax            As Long
    Dim lDelta          As Long
    Dim rgbValue        As RGBQuad
    
    Call OleTranslateColor(nColorRGB, 0, rgbValue)
    LMax = IIf(rgbValue.R > rgbValue.G, IIf(rgbValue.R > rgbValue.B, rgbValue.R, rgbValue.B), IIf(rgbValue.G > rgbValue.B, rgbValue.G, rgbValue.B))
    lMin = IIf(rgbValue.R < rgbValue.G, IIf(rgbValue.R < rgbValue.B, rgbValue.R, rgbValue.B), IIf(rgbValue.G < rgbValue.B, rgbValue.G, rgbValue.B))
    lDelta = LMax - lMin
    nValue = (LMax * 100) / 255
    If LMax > 0 Then
        nSaturation = (lDelta / LMax) * 100
        If lDelta > 0 Then
            If LMax = rgbValue.R Then
                nTemp = (CLng(rgbValue.G) - rgbValue.B) / lDelta
            ElseIf LMax = rgbValue.G Then
                nTemp = 2 + (CLng(rgbValue.B) - rgbValue.R) / lDelta
            Else
                nTemp = 4 + (CLng(rgbValue.R) - rgbValue.G) / lDelta
            End If
            nHue = nTemp * 60
            If nHue < 0 Then
                nHue = nHue + 360
            End If
        End If
    End If
End Sub
    
Private Function ColorCurrentColorSystemToRGB(nHue As Double, nLuminance As Double, nSaturation As Double) As Long
    If mColorSystem = cdColorSystemHSL Then
        Dim iH1 As Long
        Dim iL1 As Long
        Dim iS1 As Long
        
        iH1 = CLng(nHue)
        iL1 = CLng(nLuminance)
        iS1 = CLng(nSaturation)
        
        ColorCurrentColorSystemToRGB = ColorHLSToRGB(iH1, iL1, iS1)
    Else
        ColorCurrentColorSystemToRGB = ColorHSVToRGB(nHue, nLuminance, nSaturation)
    End If
End Function

Private Function ColorHSVToRGB(nHue As Double, nValue As Double, nSaturation As Double) As Long
'--- based on wqweto (Vlad Vissoultchev)'S code  from http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=36529
'--- based on *cool* code by Branco Medeiros (http://www.myrealbox.com/branco_medeiros)
'--- Converts an HSB value to the RGB color model. Adapted from Java.awt.pvColor.java
    Dim nH              As Double
    Dim nS              As Double
    Dim nL              As Double
    Dim nF              As Double
    Dim nP              As Double
    Dim nQ              As Double
    Dim nT              As Double
    Dim lH              As Long
    Dim clrConv         As RGBQuad
    
    With clrConv
        If nSaturation > 0 Then
            nH = nHue / 60
            nL = nValue / 100
            nS = nSaturation / 100
            lH = Int(nH)
            nF = nH - lH
            nP = nL * (1 - nS)
            nQ = nL * (1 - nS * nF)
            nT = nL * (1 - nS * (1 - nF))
            Select Case lH
            Case 0
                .R = nL * 255
                .G = nT * 255
                .B = nP * 255
            Case 1
                .R = nQ * 255
                .G = nL * 255
                .B = nP * 255
            Case 2
                .R = nP * 255
                .G = nL * 255
                .B = nT * 255
            Case 3
                .R = nP * 255
                .G = nQ * 255
                .B = nL * 255
            Case 4
                .R = nT * 255
                .G = nP * 255
                .B = nL * 255
            Case 5
                .R = nL * 255
                .G = nP * 255
                .B = nQ * 255
            End Select
        Else
            .R = (nValue * 255) / 100
            .G = .R
            .B = .R
        End If
    End With
    '--- return long
    CopyMemory lH, clrConv, 4
    ColorHSVToRGB = lH
End Function

Public Property Get HMax() As Long
Attribute HMax.VB_Description = "Returns the maximum value that the 'Hue' can have."
    HMax = mH_Max
End Property

Public Property Get LMax() As Long
Attribute LMax.VB_Description = "Returns the maximum value that the Luminance (or Value) can have."
    LMax = mL_Max
End Property

Public Property Get SMax() As Long
Attribute SMax.VB_Description = "Returns the maximum value that the 'Saturation' can have."
    SMax = mS_Max
End Property


Public Property Get Redraw() As Boolean
    Redraw = mRedraw
End Property

Public Property Let Redraw(ByVal nValue As Boolean)
Attribute Redraw.VB_Description = "Enables or disables redrawing of the control."
Attribute Redraw.VB_MemberFlags = "400"
    If nValue <> mRedraw Then
        mRedraw = nValue
        If mRedraw Then
            If mDrawPending Then
                mDrawPending = False
                DrawPalette
                DrawShades
                ShowSelectedColor
            End If
        End If
    End If
End Property

Public Property Get SliderControlLeft() As Single
Attribute SliderControlLeft.VB_Description = "Returns the Left position of the slider control in units of the container."
    SliderControlLeft = Round(UserControl.ScaleX(picShades.Left, UserControl.ScaleMode, vbContainerPosition))
End Property

Public Property Get SliderControlWidth() As Single
Attribute SliderControlWidth.VB_Description = "Returns the Width of the slider control in units of the container."
    SliderControlWidth = Round(UserControl.ScaleX(picShades.Width, UserControl.ScaleMode, vbContainerSize))
End Property

Public Property Get SliderParameterControlLeft() As Single
Attribute SliderParameterControlLeft.VB_Description = "Returns the Left position of the control for changing the slider parameter, in units of the container."
    SliderParameterControlLeft = Round(UserControl.ScaleX(cboSliderParameter.Left, UserControl.ScaleMode, vbContainerPosition))
End Property

Public Property Get SliderParameterControlWidth() As Single
Attribute SliderParameterControlWidth.VB_Description = "Returns the Width of the control for changing the slider parameter, in units of the container."
    SliderParameterControlWidth = Round(UserControl.ScaleX(cboSliderParameter.Width, UserControl.ScaleMode, vbContainerSize))
End Property

Public Property Get PaletteCenterX() As Single
Attribute PaletteCenterX.VB_Description = "Returns the horizontal position of the palette center in units of the container."
    PaletteCenterX = Round(UserControl.ScaleX(mCx, vbPixels, vbContainerPosition))
End Property
    
Public Property Get PaletteCenterY() As Single
Attribute PaletteCenterY.VB_Description = "Returns the vertical position of the palette center in units of the container."
    PaletteCenterY = Round(UserControl.ScaleY(mCy, vbPixels, vbContainerPosition))
End Property
    
' Slider control
Private Sub picSlider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picSlider_MouseMove Button, Shift, X, Y
End Sub

Private Sub picSlider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        SliderValue = Y / picSlider.ScaleHeight * (mSliderMax - mSliderMin) + mSliderMin
        If SliderValue > mSliderMax Then SliderValue = mSliderMax
        If SliderValue < mSliderMin Then SliderValue = mSliderMin
    End If
End Sub

Private Sub DrawSliderGrip()
    Dim iPoints() As POINTAPI
    
    If mSliderMax <= mSliderMin Then mSliderMax = mSliderMin + 1
    ReDim iPoints(2)
    iPoints(0).Y = (picSlider.ScaleHeight - mGripLenght) / (mSliderMax - mSliderMin) * (mSliderValue - mSliderMin) + mGripLenght / 2
    iPoints(0).X = 0
    iPoints(1).Y = iPoints(0).Y - mGripLenght / 2
    iPoints(1).X = mGripWidth
    iPoints(2).Y = iPoints(0).Y + mGripLenght / 2
    iPoints(2).X = mGripWidth
    
    picSlider.ForeColor = &HFF0000
    picSlider.FillColor = &H6E6E6E
    picSlider.FillStyle = vbFSSolid
    picSlider.DrawStyle = vbSolid
    picSlider.DrawWidth = 1
    picSlider.Cls
    Polygon picSlider.HDC, iPoints(0), UBound(iPoints) + 1
    picSlider.Refresh
End Sub

' Slider
Private Sub InitSlider()
    mSliderMin = 0
    mSliderMax = 100
    SliderValue = 50
    picSlider.BackColor = mBackColor
    mGripLenght = 18 * 15 / Screen.TwipsPerPixelY
    mGripWidth = 9 * 15 / Screen.TwipsPerPixelY
    picSlider.AutoRedraw = True
    picSlider.Width = (mGripWidth + 8) * Screen.TwipsPerPixelX
    DrawSliderGrip
End Sub


Private Property Get RadialValue() As Double
    If mRadialParameter = cdParameterHue Then
        RadialValue = mH
    ElseIf mRadialParameter = cdParameterLuminance Then
        RadialValue = mL
    ElseIf mRadialParameter = cdParameterSaturation Then
        RadialValue = mS
    ElseIf mRadialParameter = cdParameterRed Then
        RadialValue = mR
    ElseIf mRadialParameter = cdParameterGreen Then
        RadialValue = mG
    ElseIf mRadialParameter = cdParameterBlue Then
        RadialValue = mB
    End If
End Property

Private Property Let RadialValue(ByVal nValue As Double)
    If mRadialParameter = cdParameterHue Then
        If nValue < 0 Then nValue = 0
        If nValue > mH_Max Then nValue = mH_Max
        H = nValue
    ElseIf mRadialParameter = cdParameterLuminance Then
        If nValue < 0 Then nValue = 0
        If nValue > mL_Max Then nValue = mL_Max
        L = nValue
    ElseIf mRadialParameter = cdParameterSaturation Then
        If nValue < 0 Then nValue = 0
        If nValue > mS_Max Then nValue = mS_Max
        S = nValue
    ElseIf mRadialParameter = cdParameterRed Then
        If nValue < 0 Then nValue = 0
        If nValue > 255 Then nValue = 255
        R = nValue
    ElseIf mRadialParameter = cdParameterGreen Then
        If nValue < 0 Then nValue = 0
        If nValue > 255 Then nValue = 255
        G = nValue
    ElseIf mRadialParameter = cdParameterBlue Then
        If nValue < 0 Then nValue = 0
        If nValue > 255 Then nValue = 255
        B = nValue
    End If
End Property


Private Property Get RadialMax() As Long
    If mRadialParameter = cdParameterHue Then
        RadialMax = mH_Max
    ElseIf mRadialParameter = cdParameterLuminance Then
        RadialMax = mL_Max
    ElseIf mRadialParameter = cdParameterSaturation Then
        RadialMax = mS_Max
    ElseIf mRadialParameter = cdParameterRed Then
        RadialMax = 255
    ElseIf mRadialParameter = cdParameterGreen Then
        RadialMax = 255
    ElseIf mRadialParameter = cdParameterBlue Then
        RadialMax = 255
    End If
End Property


Private Property Get AxialValue() As Double
    If mAxialParameter = cdParameterHue Then
        AxialValue = mH
    ElseIf mAxialParameter = cdParameterLuminance Then
        AxialValue = mL
    ElseIf mAxialParameter = cdParameterSaturation Then
        AxialValue = mS
    ElseIf mAxialParameter = cdParameterRed Then
        AxialValue = mR
    ElseIf mAxialParameter = cdParameterGreen Then
        AxialValue = mG
    ElseIf mAxialParameter = cdParameterBlue Then
        AxialValue = mB
    End If
End Property

Private Property Let AxialValue(ByVal nValue As Double)
    If mAxialParameter = cdParameterHue Then
        If nValue < 0 Then nValue = 0
        If nValue > mH_Max Then nValue = mH_Max
        H = nValue
    ElseIf mAxialParameter = cdParameterLuminance Then
        If nValue < 0 Then nValue = 0
        If nValue > mL_Max Then nValue = mL_Max
        L = nValue
    ElseIf mAxialParameter = cdParameterSaturation Then
        If nValue < 0 Then nValue = 0
        If nValue > mS_Max Then nValue = mS_Max
        S = nValue
    ElseIf mAxialParameter = cdParameterRed Then
        If nValue < 0 Then nValue = 0
        If nValue > 255 Then nValue = 255
        R = nValue
    ElseIf mAxialParameter = cdParameterGreen Then
        If nValue < 0 Then nValue = 0
        If nValue > 255 Then nValue = 255
        G = nValue
    ElseIf mAxialParameter = cdParameterBlue Then
        If nValue < 0 Then nValue = 0
        If nValue > 255 Then nValue = 255
        B = nValue
    End If
End Property


Private Property Get AxialMax() As Long
    If mAxialParameter = cdParameterHue Then
        AxialMax = mH_Max
    ElseIf mAxialParameter = cdParameterLuminance Then
        AxialMax = mL_Max
    ElseIf mAxialParameter = cdParameterSaturation Then
        AxialMax = mS_Max
    ElseIf mAxialParameter = cdParameterRed Then
        AxialMax = 255
    ElseIf mAxialParameter = cdParameterGreen Then
        AxialMax = 255
    ElseIf mAxialParameter = cdParameterBlue Then
        AxialMax = 255
    End If
End Property

Public Property Get RadialParameter() As CDSliderParameterConstants
Attribute RadialParameter.VB_Description = "Returns a value indicating the radial parameter."
    RadialParameter = mRadialParameter
End Property

Public Property Get AxialParameter() As CDSliderParameterConstants
Attribute AxialParameter.VB_Description = "Returns a value indicating the axial parameter."
    AxialParameter = mAxialParameter
End Property


Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns the hWnd of the dialog window."
    hWnd = UserControl.hWnd
End Property

Private Function EnsurePrimary(nColor As Long) As Long
    Dim iRGB As RGBQuad
    Dim iHight As Byte
    Dim iLow As Byte
    Const cTolerance As Byte = 6
    
    iHight = 255 - cTolerance
    iLow = cTolerance
    
    CopyMemory iRGB, nColor, 4
    If ((iRGB.R >= iHight) Or (iRGB.R <= iLow)) And ((iRGB.G >= iHight) Or (iRGB.G <= iLow)) And ((iRGB.B >= iHight) Or (iRGB.B <= iLow)) Then
        If iRGB.R >= iHight Then
            iRGB.R = 255
        Else
            iRGB.R = 0
        End If
        If iRGB.G >= iHight Then
            iRGB.G = 255
        Else
            iRGB.G = 0
        End If
        If iRGB.B >= iHight Then
            iRGB.B = 255
        Else
            iRGB.B = 0
        End If
        CopyMemory EnsurePrimary, iRGB, 4
    Else
        EnsurePrimary = nColor
    End If
End Function

Private Sub SetCaptions()
    chkFixedPalette.Caption = GetLocalizedString1(cdUIT_ColorSelector_chkFixedPalette_Caption)
    'chkFixedPalette.ToolTipText = GetLocalizedString1(cdUIT_ColorSelector_chkFixedPalette_ToolTipText)
    'cboSliderParameter.ToolTipText = GetLocalizedString1(cdUIT_ColorSelector_cboSliderParameter_ToolTipText)
    ToolTipHandler1.Add "chkFixedPalette", GetLocalizedString1(cdUIT_ColorSelector_chkFixedPalette_ToolTipText)
    ToolTipHandler1.Add "cboSliderParameter", GetLocalizedString1(cdUIT_ColorSelector_cboSliderParameter_ToolTipText)
    lblMode.Caption = GetLocalizedString1(cdUIT_ColorSelector_lblMode_Caption)
    
    mParametersCaptions(0) = GetLocalizedString1(cdUIT_ColorSelector_cboSliderParameter_ListItem_Hue)
    mCaptionLum = GetLocalizedString1(cdUIT_ColorSelector_cboSliderParameter_ListItem_Luminance)
    mCaptionVal = GetLocalizedString1(cdUIT_ColorSelector_cboSliderParameter_ListItem_Value)
    mParametersCaptions(2) = GetLocalizedString1(cdUIT_ColorSelector_cboSliderParameter_ListItem_Saturation)
    mParametersCaptions(3) = GetLocalizedString1(cdUIT_ColorSelector_cboSliderParameter_ListItem_Red)
    mParametersCaptions(4) = GetLocalizedString1(cdUIT_ColorSelector_cboSliderParameter_ListItem_Green)
    mParametersCaptions(5) = GetLocalizedString1(cdUIT_ColorSelector_cboSliderParameter_ListItem_Blue)
    
    cboColorSystem.Clear
    cboColorSystem.AddItem GetLocalizedString1(cdUIT_ColorSelector_cboColorSystem_ListItem_HSV)
    cboColorSystem.AddItem GetLocalizedString1(cdUIT_ColorSelector_cboColorSystem_ListItem_HSL)
End Sub

Private Function GetLocalizedString1(nTextID As CDUserInterfaceTextIDConstants) As String
    GetLocalizedString1 = GetLocalizedString(nTextID)
    RaiseEvent GetLocalizedText(LanguageWindowsUI, SubLanguageWindowsUI, nTextID, GetLocalizedString1)
End Function

Public Property Get Controls() As Object
Attribute Controls.VB_MemberFlags = "40"
    Set Controls = UserControl.Controls
End Property

' Extender properties and methods
Public Property Get Name() As String
    Name = Ambient.DisplayName
End Property

Public Property Get Tag() As String
    Tag = Extender.Tag
End Property

Public Property Let Tag(ByVal Value As String)
    Extender.Tag = Value
End Property

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property

Public Property Get Container() As Object
    Set Container = Extender.Container
End Property

Public Property Set Container(ByVal Value As Object)
    Set Extender.Container = Value
End Property

Public Property Get Left() As Single
    Left = Extender.Left
End Property

Public Property Let Left(ByVal Value As Single)
    Extender.Left = Value
End Property

Public Property Get Top() As Single
    Top = Extender.Top
End Property

Public Property Let Top(ByVal Value As Single)
    Extender.Top = Value
End Property

Public Property Get Width() As Single
    Width = Extender.Width
End Property

Public Property Let Width(ByVal Value As Single)
    Extender.Width = Value
End Property

Public Property Get Height() As Single
    Height = Extender.Height
End Property

Public Property Let Height(ByVal Value As Single)
    Extender.Height = Value
End Property

Public Property Get Visible() As Boolean
    Visible = Extender.Visible
End Property

Public Property Let Visible(ByVal Value As Boolean)
    Extender.Visible = Value
End Property

Public Property Get ToolTipText() As String
    ToolTipText = Extender.ToolTipText
End Property

Public Property Let ToolTipText(ByVal Value As String)
    Extender.ToolTipText = Value
End Property

Public Property Get HelpContextID() As Long
    HelpContextID = Extender.HelpContextID
End Property

Public Property Let HelpContextID(ByVal Value As Long)
    Extender.HelpContextID = Value
End Property

Public Property Get WhatsThisHelpID() As Long
    WhatsThisHelpID = Extender.WhatsThisHelpID
End Property

Public Property Let WhatsThisHelpID(ByVal Value As Long)
    Extender.WhatsThisHelpID = Value
End Property

Public Property Get DragIcon() As IPictureDisp
    Set DragIcon = Extender.DragIcon
End Property

Public Property Let DragIcon(ByVal Value As IPictureDisp)
    Extender.DragIcon = Value
End Property

Public Property Set DragIcon(ByVal Value As IPictureDisp)
    Set Extender.DragIcon = Value
End Property

Public Property Get DragMode() As Integer
    DragMode = Extender.DragMode
End Property

Public Property Let DragMode(ByVal Value As Integer)
    Extender.DragMode = Value
End Property

Public Sub Drag(Optional ByRef Action As Variant)
    If IsMissing(Action) Then Extender.Drag Else Extender.Drag Action
End Sub

Public Sub SetFocus()
    Extender.SetFocus
End Sub

Public Sub ZOrder(Optional ByRef Position As Variant)
    If IsMissing(Position) Then Extender.ZOrder Else Extender.ZOrder Position
End Sub

