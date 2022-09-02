VERSION 5.00
Begin VB.Form frmColorDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Color selection"
   ClientHeight    =   6324
   ClientLeft      =   5796
   ClientTop       =   2076
   ClientWidth     =   7668
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6324
   ScaleWidth      =   7668
   ShowInTaskbar   =   0   'False
   Begin ColorControls.EyeDropper EyeDropper1 
      Left            =   5640
      Top             =   4440
      _ExtentX        =   762
      _ExtentY        =   762
   End
   Begin VB.PictureBox picEyeDropper 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   492
      Left            =   4860
      ScaleHeight     =   492
      ScaleWidth      =   492
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   4380
      Visible         =   0   'False
      Width           =   492
      Begin VB.PictureBox picEyeDropperIcon 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   384
         Left            =   48
         Picture         =   "frmColorDialog.frx":0000
         ScaleHeight     =   384
         ScaleWidth      =   384
         TabIndex        =   36
         Top             =   48
         Width           =   384
      End
   End
   Begin VB.PictureBox picBasicColorsContainer 
      BorderStyle     =   0  'None
      Height          =   3900
      Left            =   5000
      ScaleHeight     =   3900
      ScaleWidth      =   1704
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   60
      Visible         =   0   'False
      Width           =   1700
      Begin VB.PictureBox picBasicColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   60
         ScaleHeight     =   240
         ScaleWidth      =   324
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   240
         Width           =   324
      End
      Begin VB.Label lblBasicColors 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Basic colors:"
         Height          =   192
         Left            =   180
         TabIndex        =   34
         Top             =   0
         Width           =   864
      End
   End
   Begin VB.ComboBox cboPalette 
      Height          =   288
      ItemData        =   "frmColorDialog.frx":10CA
      Left            =   4740
      List            =   "frmColorDialog.frx":10DA
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   5352
      Visible         =   0   'False
      Width           =   1400
   End
   Begin VB.ComboBox cboColorSystem 
      Height          =   288
      ItemData        =   "frmColorDialog.frx":1112
      Left            =   2640
      List            =   "frmColorDialog.frx":111C
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   5340
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.Timer tmrCheckToDrag 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1360
      Top             =   5616
   End
   Begin VB.Timer tmrDoNotShowTT 
      Enabled         =   0   'False
      Interval        =   59000
      Left            =   960
      Top             =   5616
   End
   Begin VB.Timer tmrHideTT 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   576
      Top             =   5616
   End
   Begin VB.PictureBox picParameterLabel 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   230
      Left            =   3864
      ScaleHeight     =   228
      ScaleWidth      =   732
      TabIndex        =   29
      Top             =   3900
      Visible         =   0   'False
      Width           =   732
      Begin VB.Label lblParameter 
         AutoSize        =   -1  'True
         Caption         =   "Lum."
         Height          =   192
         Left            =   180
         TabIndex        =   30
         Top             =   0
         Width           =   348
      End
   End
   Begin VB.PictureBox picColorValuesSection 
      BorderStyle     =   0  'None
      Height          =   1040
      Left            =   120
      ScaleHeight     =   1044
      ScaleWidth      =   3000
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4272
      Width           =   3000
      Begin VB.TextBox txtLum 
         Height          =   300
         Left            =   2184
         MaxLength       =   3
         TabIndex        =   12
         Top             =   720
         Width           =   800
      End
      Begin VB.TextBox txtSat 
         Height          =   300
         Left            =   2184
         MaxLength       =   3
         TabIndex        =   8
         Top             =   360
         Width           =   800
      End
      Begin VB.TextBox txtHue 
         Height          =   300
         Left            =   2184
         MaxLength       =   3
         TabIndex        =   4
         Top             =   0
         Width           =   800
      End
      Begin VB.TextBox txtBlue 
         Height          =   300
         Left            =   624
         MaxLength       =   3
         TabIndex        =   10
         Top             =   720
         Width           =   800
      End
      Begin VB.TextBox txtGreen 
         Height          =   300
         Left            =   624
         MaxLength       =   3
         TabIndex        =   6
         Top             =   360
         Width           =   800
      End
      Begin VB.TextBox txtRed 
         Height          =   300
         Left            =   624
         MaxLength       =   3
         TabIndex        =   2
         Top             =   0
         Width           =   800
      End
      Begin VB.Label lblLum 
         Alignment       =   1  'Right Justify
         Caption         =   "Lum.:"
         Height          =   300
         Left            =   1536
         TabIndex        =   11
         Top             =   768
         Width           =   588
      End
      Begin VB.Label lblSat 
         Alignment       =   1  'Right Justify
         Caption         =   "Sat.:"
         Height          =   300
         Left            =   1536
         TabIndex        =   7
         Top             =   408
         Width           =   588
      End
      Begin VB.Label lblHue 
         Alignment       =   1  'Right Justify
         Caption         =   "Hue:"
         Height          =   300
         Left            =   1536
         TabIndex        =   3
         Top             =   48
         Width           =   588
      End
      Begin VB.Label lblBlue 
         Alignment       =   1  'Right Justify
         Caption         =   "Blue:"
         Height          =   300
         Left            =   0
         TabIndex        =   9
         Top             =   768
         Width           =   588
      End
      Begin VB.Label lblGreen 
         Alignment       =   1  'Right Justify
         Caption         =   "Green:"
         Height          =   300
         Left            =   0
         TabIndex        =   5
         Top             =   408
         Width           =   588
      End
      Begin VB.Label lblRed 
         Alignment       =   1  'Right Justify
         Caption         =   "Red:"
         Height          =   300
         Left            =   0
         TabIndex        =   1
         Top             =   48
         Width           =   588
      End
   End
   Begin VB.PictureBox picRecentContainer 
      BorderStyle     =   0  'None
      Height          =   3900
      Left            =   6800
      ScaleHeight     =   3900
      ScaleWidth      =   780
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   60
      Width           =   780
      Begin VB.PictureBox picRecent 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   324
         Index           =   0
         Left            =   168
         ScaleHeight     =   324
         ScaleWidth      =   444
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   240
         Width           =   444
      End
      Begin VB.Label lblRecent 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Recent:"
         Height          =   192
         Left            =   132
         TabIndex        =   27
         Top             =   0
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   372
      Left            =   3888
      TabIndex        =   20
      Top             =   5700
      Width           =   1284
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   372
      Left            =   2496
      TabIndex        =   19
      Top             =   5700
      Width           =   1284
   End
   Begin VB.PictureBox picSelection 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   3888
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   4320
      Width           =   800
   End
   Begin VB.Timer tmrHexChange 
      Interval        =   3000
      Left            =   144
      Top             =   5616
   End
   Begin VB.TextBox txtHex 
      Height          =   300
      Left            =   744
      MaxLength       =   11
      TabIndex        =   14
      Top             =   5352
      Width           =   800
   End
   Begin ColorControls.ColorSelector ColorSelector1 
      Height          =   3888
      Left            =   120
      TabIndex        =   0
      Top             =   144
      Width           =   4716
      _ExtentX        =   7874
      _ExtentY        =   6858
   End
   Begin VB.Label lblTT2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Press Escape key to cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   168
      Left            =   108
      TabIndex        =   35
      Top             =   5760
      Visible         =   0   'False
      Width           =   2496
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblPalette 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Palette:"
      Height          =   192
      Left            =   3780
      TabIndex        =   17
      Top             =   5400
      Visible         =   0   'False
      Width           =   528
   End
   Begin VB.Label lblColorSystem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mode:"
      Height          =   192
      Left            =   1680
      TabIndex        =   15
      Top             =   5400
      Visible         =   0   'False
      Width           =   432
   End
   Begin VB.Label lblTT 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Hold the Control key down to navigate Saturation with the mouse wheel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   420
      Left            =   168
      TabIndex        =   31
      Top             =   5736
      Visible         =   0   'False
      Width           =   2532
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblPrevious 
      Alignment       =   2  'Center
      Caption         =   "previous"
      Height          =   228
      Left            =   3888
      TabIndex        =   23
      Top             =   5110
      Width           =   854
   End
   Begin VB.Label lblNew 
      Alignment       =   2  'Center
      Caption         =   "new"
      Height          =   228
      Left            =   3888
      TabIndex        =   22
      Top             =   4104
      Width           =   854
   End
   Begin VB.Label lblHex 
      Alignment       =   1  'Right Justify
      Caption         =   "Hex:"
      Height          =   228
      Left            =   120
      TabIndex        =   13
      Top             =   5400
      Width           =   588
   End
   Begin VB.Menu mnuPopupRecent 
      Caption         =   "mnuPopupRecent"
      Visible         =   0   'False
      Begin VB.Menu mnuForgetRecent 
         Caption         =   "Forget"
      End
      Begin VB.Menu mnuClearAllRecent 
         Caption         =   "Clear recent colors"
      End
   End
End
Attribute VB_Name = "frmColorDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IBSSubclass

Public Event Change()
Public Event ColorSet()
Public Event Hided()
Public Event GetLocalizedText(ByVal LanguageID As Long, ByVal SubLanguageID As Long, ByVal TextID As Long, ByRef Text As String)

Private Const WM_NCACTIVATE As Long = &H86

Private Type RGBQuad
    R As Byte
    G As Byte
    B As Byte
    a As Byte
End Type

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
Private Declare Function HashData Lib "shlwapi" (ByVal pbData As Long, ByVal cbData As Long, ByRef pbHash As Any, ByVal cbHash As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000
Private Const WS_THICKFRAME = &H40000

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
 
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long

Private Enum CDCaptionsIDConstants
    cdCWCaptionHue = 0 ' Hue
    cdCWCaptionLum = 1 ' Lum
    cdCWCaptionSat = 2 ' Sat
    cdCWCaptionRed = 3 ' Red
    cdCWCaptionGreen = 4 ' Green
    cdCWCaptionBlue = 5 ' Blue
    cdCWCaptionVal = 6 ' Val.
    cdCWCaptionFixed = 7 ' Fixed
    cdCWCaptionFixedToolTipText = 8 ' Reflects color changes visually or not
    cdCWCaptionSliderParameterToolTipText = 9 ' Select parameter
    cdCWCaptionMode = 10 ' Mode
End Enum

Private mPreviousColor As Long
Private mSelectedColor As Long
Private mSettingCurrent As Boolean
Private mOKPressed As Boolean
Private mIndexRecentToRemove As Long
Private mPreviousColorSet As Boolean
Private mContext As String
Private mHexControlVisible As Boolean
Private mHexFormatVB As Boolean
Private mColorValuesSectionVisible As Boolean
Private mSizeBig As Boolean
Private mSliderOptionsAvailable As CDSliderOptionsAvailableConstants
Private mPaletteTypeControlVisible As Boolean
Private mColorSystemControlVisible As Boolean
Private mFixedPalette As Boolean
Private mColorSelectionBoxVisible As Boolean
Private mColorSystem As CDColorSystemConstants
Private mSliderParameter As CDSliderParameterConstants
Private mSelectionDrawHorizontal As Boolean
Private mInvalidColorMessage As String
Private mCaptionColor As String
Private mCaptionColorSet As Boolean
Private mSettingParameters As Boolean
Private mNavigatedRadially As Boolean
Private mToolTipMouseWheelFirstPart As String
Private mToolTipMouseWheelLastPart As String
Private mDialogCaptionVisible As Boolean
Private mConfirmationButtonsVisible As Boolean
Private mSliderWide As CDYesNoAutoConstants
Private mHideLabels As Boolean
Private mBackColor As Long
Private mRememberPosition As Boolean
Private mActiveFormName As String
Private mModeless As Boolean
Private mPossibleDragStart As POINTAPI
Private mStyle As CDStyleConstants
Private mRoundedBoxes As Boolean
Private mRecentColorsColumns As Long
Private mBasicColorsVisible As Boolean
Private mEyeDropperVisible As Boolean

Private mSubclassed As Boolean
Private mFormHwnd As Long
Private mEyeDropping As Boolean

Private Sub cboColorSystem_Click()
    ColorSelector1.ColorSystem = cboColorSystem.ListIndex
    mColorSystem = ColorSelector1.ColorSystem
End Sub

Private Sub cboPalette_Click()
    If cboPalette.ListIndex Mod 2 = 1 Then
        ColorSelector1.Style = cdStyleBox
    Else
        ColorSelector1.Style = cdStyleWheel
    End If
    ColorSelector1.FixedPalette = cboPalette.ListIndex < 2
    mFixedPalette = ColorSelector1.FixedPalette
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    RaiseEvent ColorSet
    SaveRecentColors
    mOKPressed = True
    If Not mModeless Then
        Unload Me
    Else
        LoadRecentColors
    End If
End Sub

Private Sub ColorSelector1_Change()
    RaiseEvent Change
End Sub

Private Sub ColorSelector1_DblClickOnColor()
    cmdOK_Click
End Sub

Private Sub ColorSelector1_GetLocalizedText(ByVal LanguageID As Long, ByVal SubLanguageID As Long, ByVal TextID As Long, Text As String)
    RaiseEvent GetLocalizedText(LanguageID, SubLanguageID, TextID, Text)
End Sub

Private Sub ColorSelector1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        CheckPossibleDrag
    End If
End Sub

Private Sub ColorSelector1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrCheckToDrag.Enabled = False
End Sub

Private Sub ColorSelector1_MouseWheelScroll(Axis As CDMouseWheelScrollConstants)
    If Axis = cdMouseWheelNavigatingAxial Then
        If Not mNavigatedRadially Then
            If Not tmrDoNotShowTT.Enabled Then
                If Not lblTT.Visible Then
                    SetMouseWheelTTText
                    lblTT.Visible = True
                    tmrHideTT.Enabled = True
                End If
            End If
        End If
    ElseIf Axis = cdMouseWheelNavigatingRadial Then
        mNavigatedRadially = True
    End If
End Sub

Private Sub SetMouseWheelTTText()
    Dim iCaptionID As Long
    Dim iMove As Boolean
    
    If mToolTipMouseWheelFirstPart = "" Then
        mToolTipMouseWheelFirstPart = GetLocalizedString1(cdUIT_frmColorDialog_MouseWheel_ToolTipStart)
    End If
    If mToolTipMouseWheelLastPart = "" Then
        mToolTipMouseWheelLastPart = GetLocalizedString1(cdUIT_frmColorDialog_MouseWheel_ToolTipEnd)
    End If
    
    If ColorSelector1.RadialParameter = cdParameterLuminance Then
        If ColorSelector1.ColorSystem = cdColorSystemHSV Then
            iCaptionID = cdCWCaptionVal
        Else
            iCaptionID = cdCWCaptionLum
        End If
    Else
        iCaptionID = ColorSelector1.RadialParameter
    End If
    
    lblTT.Caption = Trim$(mToolTipMouseWheelFirstPart) & " " & GetParameterFullName(iCaptionID) & " " & mToolTipMouseWheelLastPart
    
    lblTT.AutoSize = False
    lblTT.AutoSize = True
    If txtHex.Visible And (mRecentColorsColumns = 0) Then
        lblTT.Width = cmdOK.Left - 150
    Else
        lblTT.Width = cmdOK.Left - 300
    End If
    lblTT.AutoSize = False
    lblTT.AutoSize = True
    If txtHex.Visible Then
        If (lblTT.Height + 120) > (Me.ScaleHeight - (txtHex.Top + txtHex.Height + 30)) Then
            lblTT.Width = cmdOK.Left - 150
            lblTT.FontSize = 6.5
            lblTT.AutoSize = False
            lblTT.AutoSize = True
            iMove = True
        End If
    ElseIf mSelectionDrawHorizontal Then
        If (lblTT.Height + 120) > (Me.ScaleHeight - (picSelection.Top + picSelection.Height + 30)) Then
            lblTT.Width = cmdOK.Left - 150
            lblTT.FontSize = 6.5
            lblTT.AutoSize = False
            lblTT.AutoSize = True
            iMove = True
        End If
    End If
    If iMove Then
        lblTT.Move 90, Me.ScaleHeight - lblTT.Height - 90
    Else
        lblTT.Move 120, Me.ScaleHeight - lblTT.Height - 120
    End If
End Sub

Private Sub ColorSelector1_SliderChange()
    Dim iStr As String
    Dim iSS As Long
    Dim iSL As Long
        
    mSettingParameters = True
    mSelectedColor = ColorSelector1.Color
    
    iSS = txtHex.SelStart
    iSL = txtHex.SelLength
    
    txtHex.Text = GetHexColor(ColorSelector1.Color)
    
    On Error Resume Next
    txtHex.SelStart = iSS
    txtHex.SelLength = iSL
    On Error GoTo 0
    
    txtRed.Text = ColorSelector1.R
    txtGreen.Text = ColorSelector1.G
    txtBlue.Text = ColorSelector1.B
    txtHue.Text = Round(ColorSelector1.H)
    txtSat.Text = Round(ColorSelector1.S)
    txtLum.Text = Round(ColorSelector1.L)
    
    mSettingParameters = False
    
    If Not mSettingCurrent Then
        ShowSelection
    End If
End Sub

Private Sub ColorSelector1_ColorSystemChange()
    lblLum.Caption = EnsureEnding(IIf(ColorSelector1.ColorSystem = cdColorSystemHSV, GetLocalizedString1(cdUIT_frmColorDialog_Value_Caption), GetLocalizedString1(cdUIT_frmColorDialog_Luminance_Caption)), ":")
    
    If picParameterLabel.Visible Then
        If ColorSelector1.SliderParameter = cdParameterLuminance Then
            lblParameter.Caption = LCase$(lblLum.Caption)
            PositionlblParameter
        End If
    End If
    lblTT.Visible = False
    tmrHideTT.Enabled = False
    tmrDoNotShowTT.Enabled = False
End Sub

Private Sub ColorSelector1_SliderParameterChange()
    lblTT.Visible = False
    tmrHideTT.Enabled = False
    tmrDoNotShowTT.Enabled = False
End Sub

Private Sub EyeDropper1_UnderMouseColor(nColor As Long)
    picEyeDropper.BackColor = nColor
    picEyeDropperIcon.BackColor = picEyeDropper.BackColor
    If Not mRoundedBoxes Then
        picEyeDropper.Line (0, 0)-(picEyeDropper.ScaleWidth, picEyeDropper.ScaleHeight), vbActiveBorder, B
    End If
End Sub

Private Sub Form_Initialize()
    mBackColor = -1
    mConfirmationButtonsVisible = cPropDefault_ColorDialog_ConfirmationButtonsVisible
    mColorSelectionBoxVisible = cPropDefault_ColorDialog_ColorSelectionBoxVisible
    mHexControlVisible = cPropDefault_ColorDialog_HexControlVisible
    mHexFormatVB = cPropDefault_ColorDialog_HexFormatVB
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cmdOK.Visible = False Then
            cmdOK_Click
        End If
    ElseIf KeyAscii = vbKeyEscape Then
        If cmdCancel.Visible = False Then
            If Not mModeless Then
                cmdCancel_Click
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim iLeft As Single
    
    Set Me.Icon = Nothing
    SetBackColor
    SetCaptions
    If (Not mDialogCaptionVisible) Then SetWindowLong Me.hWnd, GWL_STYLE, GetWindowLong(Me.hWnd, GWL_STYLE) And Not (WS_CAPTION Or WS_THICKFRAME)
    If Not mConfirmationButtonsVisible Then
        cmdOK.Visible = False
        cmdCancel.Visible = False
    ElseIf mModeless Then
        cmdOK.Visible = False
        cmdCancel.Caption = "Close"
    End If
    ColorSelector1.Redraw = False
    ColorSelector1.SliderOptionsAvailable = mSliderOptionsAvailable
    ColorSelector1.FixedPaletteControlVisible = mPaletteTypeControlVisible
    ColorSelector1.ColorSystemControlVisible = mColorSystemControlVisible
    ColorSelector1.FixedPalette = mFixedPalette
    ColorSelector1.ColorSystem = mColorSystem
    ColorSelector1.Style = mStyle
    ColorSelector1.SliderParameter = mSliderParameter
    ColorSelector1_ColorSystemChange
    ColorSelector1.RoundedBoxes = mRoundedBoxes
    ColorSelector1.HideLabels = mHideLabels
    If mHideLabels Then
        lblNew.Visible = False
        lblPrevious.Visible = False
        lblParameter.Visible = False
        lblRecent.Visible = False
        lblBasicColors.Visible = False
    End If
    ColorSelector1.SliderWide = mSliderWide
    LoadBasicColors
    PositionControls
    ColorSelector1_SliderChange
    mOKPressed = False
    LoadRecentColors
    'If mRememberPosition And (mActiveFormName <> "") Then
    If mRememberPosition Then
        iLeft = Val(GetSetting(RegKey, "WindowPos", "Left", "-1"))
        If iLeft = -1 Then
            PositionForm
        Else
            Me.Move iLeft, Val(GetSetting(RegKey, "WindowPos", "Top", Me.Top))
        End If
    Else
        PositionForm
    End If
    
    ColorSelector1.Redraw = True
    
    mFormHwnd = Me.hWnd
    AttachMessage Me, mFormHwnd, WM_NCACTIVATE
    mSubclassed = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iIgnore As Boolean
    
    If Button = vbLeftButton Then
        If X >= (ColorSelector1.Left + ColorSelector1.Width) Then
            If X <= picBasicColorsContainer.Left Then
                If Y <= (ColorSelector1.Top + ColorSelector1.Height) Then
                    iIgnore = True
                End If
            End If
        End If
        If Not iIgnore Then CheckPossibleDrag
    End If
End Sub

Private Sub CheckPossibleDrag()
    GetCursorPos mPossibleDragStart
    tmrCheckToDrag.Enabled = True
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrCheckToDrag.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mOKPressed Then
        If tmrHexChange.Enabled Then tmrHexChange_Timer
        If (mSliderOptionsAvailable <> cdSliderOptionsNone) Or mPaletteTypeControlVisible Or mColorSystemControlVisible Then
            If mOKPressed Then
                SaveSetting RegKey, "Initialize", "LastColor", mSelectedColor
            End If
            If (mSliderOptionsAvailable <> cdSliderOptionsNone) Then
                SaveSetting RegKey, "Initialize", "SliderParameter", CStr(ColorSelector1.SliderParameter)
            End If
            If mPaletteTypeControlVisible Then
                SaveSetting RegKey, "Initialize", "PaletteType", CStr(CLng(cboPalette.ListIndex))
            End If
            If mColorSystemControlVisible Then
                SaveSetting RegKey, "Initialize", "ColorSystem", CStr(CLng(ColorSelector1.ColorSystem))
            End If
        End If
    End If
End Sub

Public Sub SetBackColor()
    Dim ctl As Control
    Dim iPrev As Long
    Dim iForeColor As Long
    
    If mBackColor = -1 Then mBackColor = Me.BackColor
    iPrev = Me.BackColor
    Me.BackColor = mBackColor
    picRecentContainer.BackColor = Me.BackColor
    picBasicColorsContainer.BackColor = Me.BackColor
    picParameterLabel.BackColor = Me.BackColor
    ColorSelector1.BackColor = Me.BackColor
    picColorValuesSection.BackColor = Me.BackColor
    
    If GetColorBrightness(Me.BackColor) > 170 Then
        iForeColor = vbWindowText
    Else
        iForeColor = vbWindowBackground
    End If
    
    For Each ctl In Me.Controls
        If TypeOf ctl Is Label Then
            If ctl.BackColor = iPrev Then
                ctl.BackColor = Me.BackColor
                ctl.ForeColor = iForeColor
            End If
        End If
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'If mRememberPosition And (mActiveFormName <> "") Then
    DoUnsubclass
    If mRememberPosition Then
        SaveSetting RegKey, "WindowPos", "Left", Me.Left
        SaveSetting RegKey, "WindowPos", "Top", Me.Top
    End If
    RaiseEvent Hided
End Sub

Private Sub picEyeDropperIcon_Click()
    StartDropper
End Sub

Private Sub StartDropper()
    Dim iLng As Long
    
    If cmdOK.Left > 2700 Then
        lblTT2.Width = 2500
    Else
        lblTT2.Width = cmdOK.Left - 200
    End If
    mEyeDropping = True
    lblTT2.Caption = GetLocalizedString1(cdUIT_frmColorDialog_EyeDropper_ToolTip)
    
    lblTT2.Move 90, Me.ScaleHeight - lblTT2.Height - 120
    lblTT2.Visible = True
    picEyeDropper.Tag = picEyeDropper.BackColor
    If EyeDropper1.StartDropper Then
        ColorSelector1.Color = EyeDropper1.Color
    End If
    If picEyeDropper.Tag <> "" Then
        picEyeDropper.BackColor = Val(picEyeDropper.Tag)
        picEyeDropperIcon.BackColor = picEyeDropper.BackColor
    End If
    lblTT2.Visible = False
    mEyeDropping = False
End Sub
    
Private Sub mnuClearAllRecent_Click()
    ClearRecent
End Sub

Private Sub mnuForgetRecent_Click()
    picRecent(mIndexRecentToRemove).BackColor = vbWindowBackground
    picRecent(mIndexRecentToRemove).Tag = ""

    If GetSetting(RegKey, "RecentColors", CStr(mIndexRecentToRemove + 1), "-") <> "-" Then
        DeleteSetting RegKey, "RecentColors", CStr(mIndexRecentToRemove + 1)
    End If
End Sub

Private Sub picBasicColor_DblClick(Index As Integer)
    cmdOK_Click
End Sub

Private Sub picBasicColor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        CheckPossibleDrag
    End If
End Sub

Private Sub picBasicColor_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrCheckToDrag.Enabled = False
    If Button = 1 Then
        If picBasicColor(Index).Tag <> "" Then
            ColorSelector1.Color = Val(picBasicColor(Index).Tag)
        End If
    End If
End Sub

Private Sub picBasicColorsContainer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        CheckPossibleDrag
    End If
End Sub

Private Sub picBasicColorsContainer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrCheckToDrag.Enabled = False
End Sub

Private Sub picColorValuesSection_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        CheckPossibleDrag
    End If
End Sub

Private Sub picColorValuesSection_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrCheckToDrag.Enabled = False
End Sub

Private Sub picEyeDropper_Click()
    StartDropper
End Sub

Private Sub picRecent_DblClick(Index As Integer)
    If picRecent(Index).Tag <> "" Then
        cmdOK_Click
    End If
End Sub

Private Sub picRecent_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        CheckPossibleDrag
    End If
End Sub

Private Sub picRecent_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrCheckToDrag.Enabled = False
    If Button = 1 Then
        If picRecent(Index).Tag <> "" Then
            ColorSelector1.Color = Val(picRecent(Index).Tag)
        End If
    ElseIf Button = vbRightButton Then
        mnuForgetRecent.Visible = picRecent(Index).Tag <> ""
        mIndexRecentToRemove = Index
        PopupMenu mnuPopupRecent
    End If
End Sub

Private Sub picRecentContainer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        CheckPossibleDrag
    End If
End Sub

Private Sub picRecentContainer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrCheckToDrag.Enabled = False
End Sub

Private Sub picSelection_DblClick()
    cmdOK_Click
End Sub

Private Sub picSelection_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ColorSelector1.Color = picSelection.Point(X, Y)
    End If
End Sub

Private Sub tmrCheckToDrag_Timer()
    Dim iPt As POINTAPI
    If (GetAsyncKeyState(vbKeyLButton) = 0) Then
        tmrCheckToDrag.Enabled = False
    Else
        GetCursorPos iPt
        If (Abs(mPossibleDragStart.X - iPt.X) > 5) Or (Abs(mPossibleDragStart.Y - iPt.Y) > 5) Then
            ReleaseCapture
            SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
        End If
    End If
End Sub

Private Sub tmrDoNotShowTT_Timer()
    tmrDoNotShowTT.Enabled = False
End Sub

Private Sub tmrHexChange_Timer()
    Dim iStr As String
    
    tmrHexChange.Enabled = False
    iStr = Replace(Replace(Replace(UCase(txtHex.Text), "#", ""), "&", ""), "H", "")
    If Not mHexFormatVB Then
        If Len(iStr) < 6 Then
            iStr = String$(6 - Len(iStr), "0") & iStr
        End If
        iStr = Mid$(iStr, 5, 2) & Mid$(iStr, 3, 2) & Mid$(iStr, 1, 2)
    End If
    iStr = iStr & "&"
    iStr = "&H" & iStr
    If IsValidOLE_COLOR(Val(iStr)) Then
        ColorSelector1.Color = Val(iStr)
    End If
End Sub

Private Sub tmrHideTT_Timer()
    tmrDoNotShowTT.Enabled = True
    tmrHideTT.Enabled = False
    lblTT.Visible = False
End Sub

Private Sub txtBlue_GotFocus()
    SelectTxtOnGotFocus txtBlue
End Sub

Private Sub txtHex_Change()
    tmrHexChange.Enabled = False
    tmrHexChange.Enabled = True
End Sub

Private Sub txtHex_GotFocus()
    If txtHex.SelStart = 0 Then
        txtHex.SelStart = Len(txtHex.Text)
    End If
End Sub

Private Sub txtHex_KeyPress(KeyAscii As Integer)
    Dim iStr As String
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        iStr = UCase(txtHex.Text)
        If Right$(iStr, 1) <> "&" Then
            iStr = iStr & "&"
        End If
        If Left$(iStr, 2) <> "&H" Then
            iStr = "&H" & iStr
        End If
        If IsValidOLE_COLOR(Val(iStr)) Then
            ColorSelector1.Color = Val(iStr)
            If Hex(ColorSelector1.Color) <> Hex(Val(iStr)) Then
                If mInvalidColorMessage = "" Then
                    mInvalidColorMessage = GetLocalizedString1(cdUIT_frmColorDialog_InvalidColorMessage)
                End If
                MsgBox mInvalidColorMessage, vbExclamation
                txtHex.Text = GetHexColor(ColorSelector1.Color)
            End If
        Else
            If mInvalidColorMessage = "" Then
                mInvalidColorMessage = GetLocalizedString1(cdUIT_frmColorDialog_InvalidColorMessage)
            End If
            MsgBox mInvalidColorMessage, vbExclamation
            txtHex.Text = GetHexColor(ColorSelector1.Color)
        End If
        
        tmrHexChange_Timer
    Else
        iStr = UCase(Chr$(KeyAscii))
        If InStr("0123456789ABCDEF&H#", iStr) = 0 Then
            Select Case KeyAscii
                Case 3, 24, 22, 26 ' Control+C/V/X/Z
                    '
                Case Else
                    KeyAscii = 0
            End Select
        Else
            KeyAscii = Asc(iStr)
        End If
    End If
End Sub

Private Function GetHexColor(nColor As Long) As String
    GetHexColor = Hex(nColor)
    If mHexFormatVB Then
        GetHexColor = String$(6 - Len(GetHexColor), "0") & GetHexColor
        GetHexColor = "&H" & GetHexColor
    Else
        If Len(GetHexColor) < 6 Then
            GetHexColor = String$(6 - Len(GetHexColor), "0") & GetHexColor
        End If
        GetHexColor = Mid$(GetHexColor, 5, 2) & Mid$(GetHexColor, 3, 2) & Mid$(GetHexColor, 1, 2)
        GetHexColor = "#" & UCase$(GetHexColor)
    End If
End Function

Private Sub txtGreen_GotFocus()
    SelectTxtOnGotFocus txtGreen
End Sub

Private Sub txtHex_LostFocus()
    If tmrHexChange.Enabled Then tmrHexChange_Timer
End Sub

Private Sub txtHue_GotFocus()
    SelectTxtOnGotFocus txtHue
End Sub

Private Sub txtLum_GotFocus()
    SelectTxtOnGotFocus txtLum
End Sub

Private Sub txtRed_Change()
    Dim iVal As Long
    
    iVal = Val(txtRed.Text)
    If iVal < 0 Then iVal = 0
    If iVal > 255 Then iVal = 255
    ColorSelector1.R = iVal
End Sub

Private Sub txtGreen_Change()
    Dim iVal As Long
    
    iVal = Val(txtGreen.Text)
    If iVal < 0 Then iVal = 0
    If iVal > 255 Then iVal = 255
    ColorSelector1.G = iVal
End Sub

Private Sub txtBlue_Change()
    Dim iVal As Long
    
    iVal = Val(txtBlue.Text)
    If iVal < 0 Then iVal = 0
    If iVal > 255 Then iVal = 255
    ColorSelector1.B = iVal
End Sub

Private Sub txtHue_Change()
    Dim iVal As Long
    
    If mSettingParameters Then Exit Sub
    
    iVal = Val(txtHue.Text)
    If iVal < 0 Then iVal = 0
    If iVal > ColorSelector1.HMax Then iVal = ColorSelector1.HMax
    ColorSelector1.H = iVal
End Sub

Private Sub txtLum_Change()
    Dim iVal As Long
    
    If mSettingParameters Then Exit Sub
    
    iVal = Val(txtLum.Text)
    If iVal < 0 Then iVal = 0
    If iVal > ColorSelector1.LMax Then iVal = ColorSelector1.LMax
    ColorSelector1.L = iVal
End Sub

Private Sub txtRed_GotFocus()
    SelectTxtOnGotFocus txtRed
End Sub

Private Sub txtSat_Change()
    Dim iVal As Long
    
    If mSettingParameters Then Exit Sub
    
    iVal = Val(txtSat.Text)
    If iVal < 0 Then iVal = 0
    If iVal > ColorSelector1.SMax Then iVal = ColorSelector1.SMax
    ColorSelector1.S = iVal
End Sub

Private Function IsValidOLE_COLOR(nColor As Long) As Boolean
    Dim iLng As Long
    
    IsValidOLE_COLOR = True
    If nColor > &H100FFFF Then
        IsValidOLE_COLOR = False
    ElseIf nColor < 0 Then
        If (nColor And &HFF000000) = &H80000000 Then
            iLng = nColor And &HFFFF
            If iLng > 18 Then
                IsValidOLE_COLOR = False
            End If
        Else
            IsValidOLE_COLOR = False
        End If
    End If
End Function

Private Sub SelectTxtOnGotFocus(nTextBox As Control)
    If nTextBox.SelStart = 0 Then
        If nTextBox.SelLength = 0 Then
            nTextBox.SelLength = Len(nTextBox.Text)
        End If
    End If
End Sub

Private Sub txtSat_GotFocus()
    SelectTxtOnGotFocus txtSat
End Sub

Public Property Let CurrentColor(nColor As Long)
    Dim iStr As String
    Dim iLng As Long
    
    If Not IsValidOLE_COLOR(nColor) Then
        Err.Raise 1234, "ColorWheelDialog", "Invalid OLE color."
        Exit Property
    End If
    mSettingCurrent = True
    ColorSelector1.Redraw = False
    mPreviousColor = nColor
    ColorSelector1.Color = mPreviousColor
    mSettingCurrent = False
    mPreviousColorSet = True
    ShowSelection
    
    If (mSliderOptionsAvailable <> cdSliderOptionsNone) Or mPaletteTypeControlVisible Or mColorSystemControlVisible Then
        iStr = GetSetting(RegKey, "Initialize", "LastColor", "-")
        If iStr <> "-" Then
            If Val(iStr) = mPreviousColor Then
                If mPaletteTypeControlVisible Then
                    iLng = Val(GetSetting(RegKey, "Initialize", "PaletteType", CStr(CLng(mFixedPalette))))
                    If (iLng > -1) And (iLng < cboPalette.ListCount) Then
                        cboPalette.ListIndex = iLng
                    End If
                End If
                If (mSliderOptionsAvailable <> cdSliderOptionsNone) Then
                    iLng = GetSetting(RegKey, "Initialize", "SliderParameter", CStr(mSliderParameter))
                    If (iLng >= cdParameterHue) And (iLng <= cdParameterBlue) Then
                        ColorSelector1.SliderParameter = iLng
                        mSliderParameter = iLng
                    End If
                End If
                If mColorSystemControlVisible Then
                    iLng = GetSetting(RegKey, "Initialize", "ColorSystem", CStr(mColorSystem))
                    If (iLng = cdColorSystemHSV) Or (iLng = cdColorSystemHSL) Then
                        cboColorSystem.ListIndex = iLng
                        mColorSystem = iLng
                    End If
                End If
            End If
        End If
        SaveSetting RegKey, "Initialize", "LastColor", mPreviousColor
    End If
    ColorSelector1.Redraw = True
End Property

Public Property Let Context(nContext As String)
    mContext = nContext
End Property

Private Sub ShowSelection()
    picSelection.Cls

    If mPreviousColorSet And (Not mModeless) Then
        If mSelectionDrawHorizontal Then
            picSelection.Line (0, 0)-(picSelection.ScaleWidth / 2, picSelection.ScaleHeight), mPreviousColor, BF
            picSelection.Line (picSelection.ScaleWidth / 2, 0)-(picSelection.ScaleWidth, picSelection.ScaleHeight), mSelectedColor, BF
        Else
            picSelection.Line (0, 0)-(picSelection.ScaleWidth, picSelection.ScaleHeight / 2), mSelectedColor, BF
            picSelection.Line (0, picSelection.ScaleHeight / 2)-(picSelection.ScaleWidth, picSelection.ScaleHeight), mPreviousColor, BF
        End If
        If mSelectedColor = mPreviousColor Then
            If lblPrevious.Visible Then
                If mCaptionColor = "" Then
                    mCaptionColor = GetLocalizedString1(cdUIT_frmColorDialog_Color_Caption)
                End If
                If mSelectionDrawHorizontal Then
                    lblNew.Caption = ""
                    lblPrevious.Caption = mCaptionColor
                Else
                    lblNew.Caption = mCaptionColor
                    lblPrevious.Caption = ""
                End If
                mCaptionColorSet = True
            End If
        Else
            If lblPrevious.Visible Then
                lblPrevious.Caption = GetLocalizedString1(cdUIT_frmColorDialog_ColorPrevious_Caption)
                lblNew.Caption = GetLocalizedString1(cdUIT_frmColorDialog_ColorNew_Caption)
            End If
        End If
    Else
        If mCaptionColor = "" Then
            mCaptionColor = GetLocalizedString1(cdUIT_frmColorDialog_Color_Caption)
        End If
        lblNew.Caption = mCaptionColor
        mCaptionColorSet = True
        lblPrevious.Visible = False
        picSelection.Line (0, 0)-(picSelection.ScaleWidth, picSelection.ScaleHeight), mSelectedColor, BF
    End If
    If Not mRoundedBoxes Then
        picSelection.Line (0, 0)-(picSelection.ScaleWidth, picSelection.ScaleHeight), vbActiveBorder, B
    End If
End Sub

Private Sub LoadRecentColors()
    Dim c As Long
    Dim iStr As String
    Dim c2 As Long
    Dim iStep As Long
    Dim iRecentColorsPerColum As Long
    Dim iRecentColorsBackColor As Long
    Dim iRgn As Long
    Dim iCol As Long
    
    If mRecentColorsColumns = 0 Then Exit Sub
    
    iStep = IIf(mSizeBig, 398, 405)
    iRecentColorsPerColum = IIf(mSizeBig, 12, 9)
    If Abs(GetColorBrightness(vbWindowBackground) - GetColorBrightness(Me.BackColor)) > 10 Then
        iRecentColorsBackColor = vbWindowBackground
    Else
        iRecentColorsBackColor = &HF0F0F0
    End If
    picRecent(0).BackColor = iRecentColorsBackColor
    For c = 1 To iRecentColorsPerColum - 1
        If c > picRecent.UBound Then
            Load picRecent(c)
        End If
        picRecent(c).Top = picRecent(c - 1).Top + iStep
        picRecent(c).BackColor = iRecentColorsBackColor
        picRecent(c).Visible = True
    Next c
    If mRecentColorsColumns > 1 Then
        For iCol = 2 To mRecentColorsColumns
            For c = picRecent.Count To iRecentColorsPerColum * iCol - 1
                Load picRecent(c)
                picRecent(c).Move picRecent(c - iRecentColorsPerColum).Left + picRecent(c - iRecentColorsPerColum).Width + 120, picRecent(c - iRecentColorsPerColum).Top
                picRecent(c).BackColor = iRecentColorsBackColor
                picRecent(c).Visible = True
            Next c
        Next
    End If
    If mRoundedBoxes Then
        For c = 0 To picRecent.UBound
            iRgn = CreateRoundRectRgn(0, 0, picRecentContainer.ScaleX(picRecent(0).Width, vbTwips, vbPixels), picRecentContainer.ScaleY(picRecent(0).Height, vbTwips, vbPixels), 6, 4)
            SetWindowRgn picRecent(c).hWnd, iRgn, True
            DeleteObject iRgn
        Next
    End If
    picRecentContainer.Height = picRecent(picRecent.UBound).Top + picRecent(picRecent.UBound).Height + 30
    
    For c = 1 To picRecent.Count
        iStr = GetSetting(RegKey, "RecentColors", CStr(c), "-")
        If iStr <> "-" Then
            On Error Resume Next
            picRecent(c2).BackColor = Val(iStr)
            picRecent(c2).Tag = picRecent(c2).BackColor
            On Error GoTo 0
            c2 = c2 + 1
        End If
    Next c
End Sub

Private Sub SaveRecentColors()
    Dim c As Long
    Dim iList() As Long
    Dim c2 As Long
    
    If mRecentColorsColumns = 0 Then Exit Sub
    
    ReDim iList(picRecent.Count)
    
    iList(0) = mSelectedColor
    For c = 1 To picRecent.Count
        iList(c) = -1
        If picRecent(c - 1).Tag <> "" Then
            iList(c) = Val(picRecent(c - 1).Tag)
        End If
    Next c
    For c = 1 To picRecent.Count
        If iList(c) <> -1 Then
            For c2 = 0 To c - 1
                If iList(c2) = iList(c) Then
                    iList(c) = -1
                    Exit For
                End If
            Next c2
        End If
    Next c
    c2 = 0
    For c = 0 To picRecent.Count
        If iList(c) <> -1 Then
            c2 = c2 + 1
            SaveSetting RegKey, "RecentColors", CStr(c2), CStr(iList(c))
        End If
    Next c
    For c = c2 + 1 To picRecent.Count
        If GetSetting(RegKey, "RecentColors", CStr(c), "-") <> "-" Then
            DeleteSetting RegKey, "RecentColors", CStr(c)
        End If
    Next
End Sub

Private Sub ClearRecent()
    Dim c As Long
    
    For c = 1 To picRecent.UBound + 1
        If GetSetting(RegKey, "RecentColors", CStr(c), "-") <> "-" Then
            DeleteSetting RegKey, "RecentColors", CStr(c)
        End If
        picRecent(c - 1).BackColor = vbWindowBackground
        picRecent(c - 1).Tag = ""
    Next c
End Sub

Public Property Get OKPressed() As Boolean
    OKPressed = mOKPressed
End Property

Public Property Get SelectedColor() As Long
    SelectedColor = mSelectedColor
End Property


Public Property Let HexControlVisible(nValue As Boolean)
    mHexControlVisible = nValue
End Property

Public Property Let HexFormatVB(nValue As Boolean)
    mHexFormatVB = nValue
End Property

Public Property Let ColorValuesSectionVisible(nValue As Boolean)
    mColorValuesSectionVisible = nValue
End Property

Public Property Let RecentColorsPerColumn(nValue As Long)
    mRecentColorsColumns = nValue
End Property

Public Property Let SizeBig(nValue As Boolean)
    mSizeBig = nValue
End Property

Public Property Let SliderOptionsAvailable(nValue As CDSliderOptionsAvailableConstants)
    mSliderOptionsAvailable = nValue
End Property

Public Property Let PaletteTypeControlVisible(nValue As Boolean)
    mPaletteTypeControlVisible = nValue
End Property

Public Property Let ColorSystemControlVisible(nValue As Boolean)
    mColorSystemControlVisible = nValue
End Property

Public Property Let FixedPalette(nValue As Boolean)
    mFixedPalette = nValue
End Property

Public Property Let ColorSystem(nValue As CDColorSystemConstants)
    mColorSystem = nValue
End Property

Public Property Let SliderParameter(nValue As CDSliderParameterConstants)
    mSliderParameter = nValue
End Property

Public Property Let ConfirmationButtonsVisible(nValue As Boolean)
    mConfirmationButtonsVisible = nValue
End Property

Public Property Let ColorSelectionBoxVisible(nValue As Boolean)
    mColorSelectionBoxVisible = nValue
End Property

Public Property Let DialogCaptionVisible(nValue As Boolean)
    mDialogCaptionVisible = nValue
End Property

Public Property Let SliderWide(nValue As CDYesNoAutoConstants)
    mSliderWide = nValue
End Property

Public Property Let HideLabels(nValue As Boolean)
    mHideLabels = nValue
End Property

Public Property Let BackColor2(nValue As OLE_COLOR)
    mBackColor = nValue
End Property

Public Property Let RememberPosition(nValue As Boolean)
    mRememberPosition = nValue
End Property

Public Property Let ActiveFormName(nValue As String)
    mActiveFormName = nValue
End Property

Public Property Let Modeless(nValue As Boolean)
    mModeless = nValue
End Property

Public Property Let Style(nValue As CDStyleConstants)
    mStyle = nValue
End Property
    
Public Property Let RoundedBoxes(nValue As Boolean)
    mRoundedBoxes = nValue
End Property

Public Property Let RecentColorsColumns(nValue As Long)
    mRecentColorsColumns = nValue
End Property

Public Property Let BasicColorsVisible(nValue As Boolean)
    mBasicColorsVisible = nValue
End Property

Public Property Let EyeDropperVisible(nValue As Boolean)
    mEyeDropperVisible = nValue
End Property


Private Sub PositionControls()
    Dim iRgn As Long
    Dim iFormWidth As Long
    Dim iFormHeight As Long
    Dim iCboPaletteVisible As Boolean
    Dim iTitleBarHeight As Long
    Dim iLastTop As Long
    Const SM_CYCAPTION = 4
    
    lblBasicColors.Left = (picBasicColorsContainer.ScaleWidth - lblBasicColors.Width) / 2
    If mRecentColorsColumns > 1 Then
        picRecentContainer.Width = 780 + 570 * (mRecentColorsColumns - 1)
        lblRecent.Left = (picRecentContainer.ScaleWidth - lblRecent.Width) / 2
    End If
    If mSizeBig Then
        ColorSelector1.Height = 5000
        picColorValuesSection.Top = ColorSelector1.Height + 420
        txtHex.Top = picColorValuesSection.Top + picColorValuesSection.Height + 40
        lblHex.Top = txtHex.Top + 50
    End If
    cboColorSystem.Top = txtHex.Top + Screen.TwipsPerPixelY
    lblColorSystem.Top = lblHex.Top
    
    iFormHeight = Me.Height
    If mDialogCaptionVisible Then
        iTitleBarHeight = GetSystemMetrics(SM_CYCAPTION) * Screen.TwipsPerPixelY  'iFormHeight - Me.ScaleHeight
    Else
        iTitleBarHeight = 0
    End If
    iFormHeight = ColorSelector1.Height + 2500 + iTitleBarHeight + IIf(mHexControlVisible Or mColorSystemControlVisible Or mPaletteTypeControlVisible, IIf((CLng(mPaletteTypeControlVisible) + CLng(mHexControlVisible) + CLng(mColorSystemControlVisible) < -1) And Not (mHexControlVisible And mColorSystemControlVisible And (Not mPaletteTypeControlVisible)), 400, 0), -200)
    iFormWidth = ColorSelector1.Width + 320 ' IIf(mSizeBig, 320, 260) ' IIf(ColorSelector1.Style = cdStyleBox, 320, 220)
    iLastTop = ColorSelector1.Top + ColorSelector1.Height
    
    If mBasicColorsVisible Then
        If mSizeBig Then
            If mBasicColorsVisible Then
                picSelection.Left = ColorSelector1.Left + ColorSelector1.SliderControlLeft + (ColorSelector1.SliderControlWidth - picSelection.Width) / 2
            Else
                picSelection.Left = 3900
            End If
        Else
            picSelection.Left = ColorSelector1.Left + ColorSelector1.SliderControlLeft + (ColorSelector1.SliderControlWidth - picSelection.Width) / 2
        End If
    Else
        If (mSliderOptionsAvailable <> cdSliderOptionsNone) Then
            If ColorSelector1.Style = cdStyleWheel Then
                picSelection.Left = ColorSelector1.Left + ColorSelector1.SliderParameterControlLeft + ColorSelector1.SliderParameterControlWidth / 2 - ColorSelector1.SliderParameterControlWidth / 2
                picSelection.Width = ColorSelector1.SliderParameterControlWidth
            Else
                picSelection.Left = ColorSelector1.Left + ColorSelector1.SliderControlLeft + ColorSelector1.SliderControlWidth / 2 - picSelection.Width / 2
            End If
        Else
            picSelection.Left = ColorSelector1.Left + ColorSelector1.SliderControlLeft + ColorSelector1.SliderControlWidth / 2 - picSelection.Width / 2
        End If
    End If
    
    If mBasicColorsVisible Then
        picBasicColorsContainer.Left = ColorSelector1.Width + 170
        picBasicColorsContainer.Visible = True
        If mRecentColorsColumns = 0 Then
            iFormWidth = picBasicColorsContainer.Left + picBasicColorsContainer.Width + 220
        Else
            picRecentContainer.Left = picBasicColorsContainer.Left + picBasicColorsContainer.Width + 144
            iFormWidth = picRecentContainer.Left + picRecentContainer.Width + 120
            picRecentContainer.Visible = True
        End If
    Else
        If mRecentColorsColumns > 0 Then
            picRecentContainer.Left = ColorSelector1.Width + 170
            picRecentContainer.Visible = True
            iFormWidth = picRecentContainer.Left + picRecentContainer.Width + 120
        End If
    End If
    
    lblNew.Left = picSelection.Left
    lblPrevious.Left = picSelection.Left
    
    If Not mColorValuesSectionVisible Then
        picColorValuesSection.Visible = False
        mHexControlVisible = False
        mColorSystemControlVisible = False
        mPaletteTypeControlVisible = False
        'picSelection.Move ColorSelector1.Left + ColorSelector1.PaletteCenterX - 1000 / 2, ColorSelector1.Height + 350, 1000, 440
        picSelection.Move ColorSelector1.Left + 100 + ColorSelector1.PaletteCenterX - 1000 / 2, ColorSelector1.Height + 350, 1000, 440
        lblNew.Alignment = vbLeftJustify
        lblPrevious.Alignment = vbRightJustify
        lblPrevious.Move picSelection.Left - lblNew.Width - 60, picSelection.Top + picSelection.Height / 2 - lblNew.Height / 2
        lblNew.Move picSelection.Left + picSelection.Width + 60, lblPrevious.Top
        mSelectionDrawHorizontal = True
    Else
        picSelection.Top = (iFormHeight - iTitleBarHeight) - 1932 + IIf(mHexControlVisible Or mColorSystemControlVisible Or mPaletteTypeControlVisible, 0, 200)
        lblNew.Top = picSelection.Top - 220
        lblPrevious.Top = picSelection.Top + picSelection.Height + 10
    End If
    If Not mColorSelectionBoxVisible Then
        picSelection.Visible = False
        lblNew.Visible = False
        lblPrevious.Visible = False
    ElseIf mModeless Then
        lblPrevious.Visible = False
    End If
    If Not mHexControlVisible Then
        lblHex.Visible = False
        txtHex.Visible = False
        If mColorSystemControlVisible Or mPaletteTypeControlVisible Then
            If mColorSystemControlVisible Then
                cboColorSystem.Left = txtHex.Left
                lblColorSystem.Left = cboColorSystem.Left - lblColorSystem.Width - 60
                cboColorSystem.ListIndex = mColorSystem
                lblColorSystem.Visible = True
                cboColorSystem.Visible = True
                ColorSelector1.ColorSystemControlVisible = False
                If mPaletteTypeControlVisible Then
                    cboPalette.Top = cboColorSystem.Top + cboColorSystem.Height + 75
                    lblPalette.Top = lblColorSystem.Top + cboColorSystem.Height + 75
                    cboPalette.Left = txtHex.Left
                    lblPalette.Left = cboPalette.Left - lblPalette.Width - 60
                    iCboPaletteVisible = True
                End If
            ElseIf mPaletteTypeControlVisible Then
                cboPalette.Top = cboColorSystem.Top
                lblPalette.Top = lblColorSystem.Top
                cboPalette.Left = txtHex.Left ' lblPalette.Left + lblPalette.Width + 60
                lblPalette.Left = cboPalette.Left - lblPalette.Width - 60
                iCboPaletteVisible = True
            End If
        End If
    ElseIf mColorSystemControlVisible Or mPaletteTypeControlVisible Then
        If mColorSystemControlVisible Then
            cboColorSystem.Left = picColorValuesSection.Left + txtLum.Left ' lblColorSystem.Left + lblColorSystem.Width + 60
            lblColorSystem.Left = cboColorSystem.Left - lblColorSystem.Width - 60
            cboColorSystem.ListIndex = mColorSystem
            lblColorSystem.Visible = True
            cboColorSystem.Visible = True
            ColorSelector1.ColorSystemControlVisible = False
            If mPaletteTypeControlVisible Then
                cboPalette.Top = cboColorSystem.Top + cboColorSystem.Height + 75
                lblPalette.Top = lblColorSystem.Top + cboColorSystem.Height + 75
                cboPalette.Left = txtHex.Left ' lblPalette.Left + lblPalette.Width + 60
                lblPalette.Left = cboPalette.Left - lblPalette.Width - 60
                iCboPaletteVisible = True
            End If
        ElseIf mPaletteTypeControlVisible Then
            cboPalette.Top = cboColorSystem.Top + cboColorSystem.Height + 75
            lblPalette.Top = lblColorSystem.Top + cboColorSystem.Height + 75
            cboPalette.Left = txtHex.Left ' lblPalette.Left + lblPalette.Width + 60
            lblPalette.Left = cboPalette.Left - lblPalette.Width - 60
            iCboPaletteVisible = True
        End If
    End If
    
    If iCboPaletteVisible Then
        cboPalette.Visible = True
        lblPalette.Visible = True
        cboPalette.ListIndex = mStyle + (CLng(Not mFixedPalette) * -1) * 2
        ColorSelector1.FixedPaletteControlVisible = False
    End If
    If (mSliderOptionsAvailable = cdSliderOptionsNone) Then
        lblParameter.Move 0, 0
        PositionlblParameter
        picParameterLabel.Visible = True
    End If
    
    If mEyeDropperVisible Then
        If mBasicColorsVisible Or (mRecentColorsColumns > 0) Then
            If (mRecentColorsColumns > 0) Then
'                If mSizeBig Then
 '                   picEyeDropper.Move ColorSelector1.Left + ColorSelector1.SliderControlLeft + (ColorSelector1.SliderControlWidth - picEyeDropper.Width) / 2 + Screen.TwipsPerPixelX, picSelection.Top + (picSelection.Height - picEyeDropper.Height) / 2
  '              Else
                
   '             End If
                If mColorSelectionBoxVisible Then
                    picEyeDropper.Move picRecentContainer.Left + (picRecentContainer.Width - picEyeDropper.Width) / 2, picSelection.Top + (picSelection.Height - picEyeDropper.Height) / 2
                Else
                    If mColorValuesSectionVisible Then
                        picEyeDropper.Move picRecentContainer.Left + (picRecentContainer.Width - picEyeDropper.Width) / 2, picColorValuesSection.Top + (picColorValuesSection.Height - picEyeDropper.Height) / 2
                    Else
                        picEyeDropper.Move picRecentContainer.Left + (picRecentContainer.Width - picEyeDropper.Width) / 2, ColorSelector1.Top + ColorSelector1.Height + 120
                    End If
                End If
            Else
                picEyeDropper.Move picBasicColorsContainer.Left + (picBasicColorsContainer.Width - picEyeDropper.Width) / 2, picSelection.Top + (picSelection.Height - picEyeDropper.Height) / 2
            End If
        Else
            If mColorValuesSectionVisible Then
                If mSizeBig Then
                    picEyeDropper.Move (picColorValuesSection.Left + picColorValuesSection.Width + picSelection.Left - picEyeDropper.Width) / 2, picSelection.Top + (picSelection.Height - picEyeDropper.Height) / 2
                Else
                    picSelection.Top = picSelection.Top + 700
                    lblNew.Top = picSelection.Top - 220
                    lblPrevious.Top = picSelection.Top + picSelection.Height + 10
                    picEyeDropper.Move picSelection.Left + (picSelection.Width - picEyeDropper.Width) / 2, (ColorSelector1.Top + ColorSelector1.Height + picSelection.Top - picEyeDropper.Height) / 2
                End If
            Else
                picEyeDropper.Move ColorSelector1.Left + ColorSelector1.SliderControlLeft + (ColorSelector1.SliderControlWidth - picEyeDropper.Width) / 2 + Screen.TwipsPerPixelX, picSelection.Top + (picSelection.Height - picEyeDropper.Height) / 2
            End If
        End If
        picEyeDropper.Visible = True
    End If
    
    If mRoundedBoxes Then
        iRgn = CreateRoundRectRgn(0, 0, Me.ScaleX(picSelection.Width, vbTwips, vbPixels), Me.ScaleY(picSelection.Height, vbTwips, vbPixels), 6, 4)
        SetWindowRgn picSelection.hWnd, iRgn, True
        DeleteObject iRgn
        iRgn = CreateRoundRectRgn(0, 0, Me.ScaleX(picEyeDropper.Width, vbTwips, vbPixels), Me.ScaleY(picEyeDropper.Height, vbTwips, vbPixels), 6, 4)
        SetWindowRgn picEyeDropper.hWnd, iRgn, True
        DeleteObject iRgn
    End If
    
    If mColorValuesSectionVisible Then HandleLastTopValue iLastTop, picColorValuesSection
    If mHexControlVisible Then HandleLastTopValue iLastTop, txtHex
    If mColorSystemControlVisible Then HandleLastTopValue iLastTop, cboColorSystem
    If mPaletteTypeControlVisible Then HandleLastTopValue iLastTop, cboPalette
    If mEyeDropperVisible Then HandleLastTopValue iLastTop, picEyeDropper
    If mColorSelectionBoxVisible Then
        HandleLastTopValue iLastTop, picSelection
        HandleLastTopValue iLastTop, lblPrevious
    Else
        If Not mColorValuesSectionVisible Then
            iLastTop = iLastTop + 120
        End If
    End If
    If mConfirmationButtonsVisible Then
        If mDialogCaptionVisible Then
            iLastTop = iLastTop + 1120
        Else
            iLastTop = iLastTop + 720
        End If
        If Not mColorValuesSectionVisible Then
            iLastTop = iLastTop + 120
        End If
    Else
        If mDialogCaptionVisible Then
            iLastTop = iLastTop + 700
        Else
            iLastTop = iLastTop + 200
        End If
    End If
    
    If mConfirmationButtonsVisible Then
        Me.Move Me.Left, Me.Top, iFormWidth, iLastTop
        If (Not mSizeBig) And (mRecentColorsColumns = 0) And mColorValuesSectionVisible And mHexControlVisible Then
            cmdOK.Left = Me.ScaleWidth - 3030 + 120
        Else
            cmdOK.Left = Me.ScaleWidth - 3030
        End If
        cmdOK.Top = (iLastTop - iTitleBarHeight) - cmdOK.Height - 200
        cmdCancel.Move cmdOK.Left + 1400, cmdOK.Top
    Else
        Me.Move Me.Left, Me.Top, iFormWidth, iLastTop
    End If
End Sub

Private Sub HandleLastTopValue(ByRef nLT As Long, nCtl As Control)
    Dim t As Long
    
    t = nCtl.Top + nCtl.Height
    If t > nLT Then nLT = t
End Sub
    
Private Sub PositionlblParameter()
    If ColorSelector1.SliderParameter = cdParameterLuminance Then
        If ColorSelector1.ColorSystem = cdColorSystemHSV Then
            lblParameter.Caption = LCase$(WithoutEnding(ColorSelector1.GetCaption(cdCWCaptionVal), ":"))
        Else
            lblParameter.Caption = LCase$(WithoutEnding(ColorSelector1.GetCaption(cdCWCaptionLum), ":"))
        End If
    Else
        lblParameter.Caption = LCase$(WithoutEnding(ColorSelector1.GetCaption(ColorSelector1.SliderParameter), ":"))
    End If
    lblParameter.AutoSize = False
    lblParameter.AutoSize = True
    picParameterLabel.Move ColorSelector1.Left + ColorSelector1.SliderControlLeft + ColorSelector1.SliderControlWidth / 2 - lblParameter.Width / 2, ColorSelector1.Height + 60, lblParameter.Width
End Sub

Private Function EnsureEnding(nText As Variant, nEnding As String)
    EnsureEnding = nText
    If Right$(EnsureEnding, Len(nEnding)) <> nEnding Then
        EnsureEnding = EnsureEnding & nEnding
    End If
End Function

Private Function WithoutEnding(nText As Variant, nEnding As String)
    WithoutEnding = nText
    If Right$(WithoutEnding, Len(nEnding)) = nEnding Then
        WithoutEnding = Left$(WithoutEnding, Len(WithoutEnding) - 1)
    End If
End Function

Private Sub PositionForm()
    Dim iAFHwnd As Long
    Dim iRc As RECT
    Dim iPt As POINTAPI
    Dim iShift As Long
    
    iAFHwnd = GetActiveWindow
    If iAFHwnd <> 0 Then
        GetWindowRect iAFHwnd, iRc
        If iRc.Top < (Screen.Height / Screen.TwipsPerPixelY) And iRc.Left < (Screen.Width / Screen.TwipsPerPixelX) Then
            If (iRc.Top + 100 + Me.Height / Screen.TwipsPerPixelY) > (Screen.Height / Screen.TwipsPerPixelY - 100) Then
                 iRc.Top = (Screen.Height / Screen.TwipsPerPixelY - 100) - Me.Height / Screen.TwipsPerPixelY - 100
            End If
            If (iRc.Left + 150 + Me.Width / Screen.TwipsPerPixelX) > (Screen.Width / Screen.TwipsPerPixelX) Then
                iRc.Left = Screen.Width / Screen.TwipsPerPixelX - Me.Width / Screen.TwipsPerPixelX - 150
            End If
        End If
        Me.Move ScaleX(iRc.Left + 100, vbPixels, vbTwips), ScaleY(iRc.Top + 100, vbPixels, vbTwips)
    Else
        GetCursorPos iPt
        iPt.X = iPt.X - 15
        If iPt.X < 10 Then iPt.X = 10
        iPt.Y = iPt.Y + 20
        
        If iPt.Y < (Screen.Height / Screen.TwipsPerPixelY) And iPt.X < (Screen.Width / Screen.TwipsPerPixelX) Then
            If (iPt.Y + Me.Height / Screen.TwipsPerPixelY) > (Screen.Height / Screen.TwipsPerPixelY - 100) Then
                 iPt.Y = (Screen.Height / Screen.TwipsPerPixelY - 100) - Me.Height / Screen.TwipsPerPixelY
            End If
            If (iPt.X + 50 + Me.Width / Screen.TwipsPerPixelX) > (Screen.Width / Screen.TwipsPerPixelX) Then
                iPt.X = Screen.Width / Screen.TwipsPerPixelX - Me.Width / Screen.TwipsPerPixelX - 50
            End If
        End If
        Me.Move ScaleX(iPt.X, vbPixels, vbTwips), ScaleY(iPt.Y, vbPixels, vbTwips)
    End If
End Sub

Private Function RegKey() As String
    Static sValue As String
    
    If sValue = "" Then
        sValue = ClientProductName & "\ColorDialog"
        If (mActiveFormName <> "") Or (mContext <> "") Then
            sValue = sValue & "\" & SimpleHash(mActiveFormName & ":" & mContext)
        End If
    End If
    RegKey = sValue
End Function

Private Function SimpleHash(ByVal nData As Variant, Optional pNumberOfHasCharacters_MustBeEvenAndLessThan514 As Long = 8) As String
    Dim iHashBytes() As Byte
    Dim c As Long
    Dim n As Long
    Dim iStr As String
    Dim iVarType As Long
    Dim iDataBytes() As Byte
    
    n = (pNumberOfHasCharacters_MustBeEvenAndLessThan514 / 2)
    ReDim iHashBytes(n - 1)
    iVarType = VarType(nData)
    If iVarType = vbString Then
        iStr = nData
        HashData StrPtr(iStr), 2 * Len(iStr), iHashBytes(0), n
    Else
        If iVarType <> vbArray + vbByte Then
            Err.Raise 2345, , "Invalid data type"
            Exit Function
        Else
            iDataBytes = nData
            HashData VarPtr(iDataBytes(0)), UBound(iDataBytes) + 1, iHashBytes(0), n
        End If
    End If
    For c = 0 To UBound(iHashBytes)
        iStr = Hex$(iHashBytes(c))
        If Len(iStr) = 1 Then
            iStr = "0" & iStr
        End If
        SimpleHash = SimpleHash & iStr
    Next c
End Function

Private Function GetColorBrightness(ByVal nColor As Long) As Long
    Dim iRGB As RGBQuad
    
    TranslateColor nColor, 0&, nColor
    CopyMemory iRGB, nColor, 4
    GetColorBrightness = (0.2125 * iRGB.R + 0.7154 * iRGB.G + 0.0721 * iRGB.B)
End Function

Private Sub LoadBasicColors()
    Dim c As Long
    Dim iBasicColorsBackColor As Long
    Dim iStep As Long
    Dim iCol As Long
    Dim iRgn As Long
    Dim iColor(47) As Long
    
    If Abs(GetColorBrightness(vbWindowBackground) - GetColorBrightness(Me.BackColor)) > 10 Then
        iBasicColorsBackColor = vbWindowBackground
    Else
        iBasicColorsBackColor = &HF0F0F0
    End If
    
    If mSizeBig Then
        picBasicColor(0).Height = 324
    End If
    iStep = picBasicColor(0).Height + 62
    If mSizeBig Then iStep = iStep + 12
    picBasicColor(0).BackColor = iBasicColorsBackColor
    For c = 1 To 11
        If c > picBasicColor.UBound Then
            Load picBasicColor(c)
        End If
        picBasicColor(c).Top = picBasicColor(c - 1).Top + iStep
        picBasicColor(c).BackColor = iBasicColorsBackColor
        picBasicColor(c).Visible = True
    Next c
    For iCol = 2 To 4
        For c = picBasicColor.Count To 12 * iCol - 1
            Load picBasicColor(c)
            picBasicColor(c).Move picBasicColor(c - 12).Left + picBasicColor(c - 12).Width + 64, picBasicColor(c - 12).Top
            picBasicColor(c).BackColor = iBasicColorsBackColor
            picBasicColor(c).Visible = True
        Next c
    Next
    If mRoundedBoxes Then
        For c = 0 To picBasicColor.UBound
            iRgn = CreateRoundRectRgn(0, 0, picBasicColorsContainer.ScaleX(picBasicColor(0).Width, vbTwips, vbPixels), picBasicColorsContainer.ScaleY(picBasicColor(0).Height, vbTwips, vbPixels), 6, 4)
            SetWindowRgn picBasicColor(c).hWnd, iRgn, True
            DeleteObject iRgn
        Next
    End If
    picBasicColorsContainer.Height = picBasicColor(picBasicColor.UBound).Top + picBasicColor(picBasicColor.UBound).Height + 30
    picBasicColorsContainer.Width = picBasicColor(picBasicColor.UBound).Left + picBasicColor(picBasicColor.UBound).Width + 64
    
    iColor(0) = 8421631:    iColor(1) = 64:    iColor(2) = 4227327:    iColor(3) = 8454016:    iColor(4) = 16384:    iColor(5) = 8421376:    iColor(6) = 16777088:    iColor(7) = 8388608:    iColor(8) = 16744576:    iColor(9) = 12615935
    iColor(10) = 4194368:    iColor(11) = 8388863:    iColor(12) = 255:    iColor(13) = 0:    iColor(14) = 33023:    iColor(15) = 65408:    iColor(16) = 4227200:    iColor(17) = 4227072:    iColor(18) = 16776960:    iColor(19) = 8421440
    iColor(20) = 10485760:    iColor(21) = 12615808:    iColor(22) = 4194368:    iColor(23) = 16711808:    iColor(24) = 4210816:    iColor(25) = 8454143:    iColor(26) = 16512:    iColor(27) = 65280:    iColor(28) = 8453888:    iColor(29) = 4210688
    iColor(30) = 8404992:    iColor(31) = 16744448:    iColor(32) = 4194304:    iColor(33) = 4194432:    iColor(34) = 16744703:    iColor(35) = 8388672:    iColor(36) = 128:    iColor(37) = 65535:    iColor(38) = 32896:    iColor(39) = 32768
    iColor(40) = 4259584:    iColor(41) = 8421504:    iColor(42) = 16711680:    iColor(43) = 12615680:    iColor(44) = 12632256:    iColor(45) = 8388736:    iColor(46) = 16711935:    iColor(47) = 16777215
    
    For c = 0 To UBound(iColor)
        picBasicColor(c).BackColor = iColor(c)
        picBasicColor(c).Tag = iColor(c)
    Next
End Sub

Private Function IBSSubclass_MsgResponse(ByVal hWnd As Long, ByVal iMsg As Long) As Long
    IBSSubclass_MsgResponse = emrPostProcess
End Function

Private Sub IBSSubclass_UnsubclassIt()
    DoUnsubclass
End Sub

Private Function IBSSubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, wParam As Long, lParam As Long, bConsume As Boolean) As Long
    If iMsg = WM_NCACTIVATE Then
        If wParam = 0 Then ' deactivate
            If Not mConfirmationButtonsVisible Then
                If Not mEyeDropping Then
                    cmdOK_Click
                End If
            End If
        End If
    End If
End Function
 
Private Sub DoUnsubclass()
    If mSubclassed Then
        mSubclassed = False
        DetachMessage Me, mFormHwnd, WM_NCACTIVATE
    End If
End Sub

Private Sub SetCaptions()
    Me.Caption = GetLocalizedString1(cdUIT_frmColorDialog_Form_Caption)
    lblBasicColors.Caption = GetLocalizedString1(cdUIT_frmColorDialog_lblBasicColors_Caption)
    lblRecent.Caption = GetLocalizedString1(cdUIT_frmColorDialog_lblRecent_Caption)
    lblRed.Caption = GetLocalizedString1(cdUIT_frmColorDialog_lblRed_Caption)
    lblGreen.Caption = GetLocalizedString1(cdUIT_frmColorDialog_lblGreen_Caption)
    lblBlue.Caption = GetLocalizedString1(cdUIT_frmColorDialog_lblBlue_Caption)
    lblHex.Caption = GetLocalizedString1(cdUIT_frmColorDialog_lblHex_Caption)
    lblHue.Caption = GetLocalizedString1(cdUIT_frmColorDialog_lblHue_Caption)
    lblSat.Caption = GetLocalizedString1(cdUIT_frmColorDialog_lblSaturation_Caption)
    lblColorSystem.Caption = GetLocalizedString1(cdUIT_frmColorDialog_lblColorSystem_Caption)
    lblPalette.Caption = GetLocalizedString1(cdUIT_frmColorDialog_lblPalette_Caption)
    cboPalette.Clear
    cboPalette.AddItem GetLocalizedString1(cdUIT_frmColorDialog_cboPalette_ListItem1)
    cboPalette.AddItem GetLocalizedString1(cdUIT_frmColorDialog_cboPalette_ListItem2)
    cboPalette.AddItem GetLocalizedString1(cdUIT_frmColorDialog_cboPalette_ListItem3)
    cboPalette.AddItem GetLocalizedString1(cdUIT_frmColorDialog_cboPalette_ListItem4)
End Sub

Private Function GetParameterFullName(nID As Long) As String
    GetParameterFullName = "'" & GetLocalizedString1(cdUIT_frmColorDialog_ParameterFullName_Hue + nID) & "'"
End Function

Private Function GetLocalizedString1(nTextID As CDUserInterfaceTextIDConstants) As String
    GetLocalizedString1 = GetLocalizedString(nTextID)
    RaiseEvent GetLocalizedText(UILanguage, UISubLanguage, nTextID, GetLocalizedString1)
End Function
