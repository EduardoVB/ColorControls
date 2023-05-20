VERSION 5.00
Object = "*\A..\control-source\ClrDlg.vbp"
Begin VB.Form frmTestProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test properties"
   ClientHeight    =   11136
   ClientLeft      =   2460
   ClientTop       =   504
   ClientWidth     =   11976
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11136
   ScaleWidth      =   11976
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboPointerType 
      Height          =   336
      ItemData        =   "frmTestProperties.frx":0000
      Left            =   2760
      List            =   "frmTestProperties.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   41
      Top             =   9912
      Width           =   1752
   End
   Begin VB.CommandButton cmdDocs 
      Caption         =   "Documentation"
      Height          =   492
      Left            =   5580
      TabIndex        =   49
      Top             =   10440
      Width           =   1512
   End
   Begin VB.CheckBox chkHexFormatVB 
      Caption         =   "HexFormatVB"
      Height          =   372
      Left            =   2760
      TabIndex        =   20
      Top             =   5004
      Width           =   2592
   End
   Begin VB.CheckBox chkEyeDropperVisible 
      Caption         =   "EyeDropperVisible"
      Height          =   372
      Left            =   2760
      TabIndex        =   17
      Top             =   4044
      Width           =   2592
   End
   Begin ColorControls.EyeDropper EyeDropper1 
      Left            =   7620
      Top             =   5160
      _ExtentX        =   762
      _ExtentY        =   762
   End
   Begin VB.CommandButton cmdTestEyeDropper 
      Caption         =   "Pick a color from the screen"
      Height          =   492
      Left            =   9180
      TabIndex        =   46
      Top             =   5100
      Width           =   2592
   End
   Begin VB.PictureBox picEyeDropper 
      BorderStyle     =   0  'None
      FillStyle       =   5  'Downward Diagonal
      Height          =   372
      Left            =   8400
      ScaleHeight     =   372
      ScaleWidth      =   552
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   5160
      Width           =   552
   End
   Begin VB.PictureBox picColor 
      BorderStyle     =   0  'None
      FillStyle       =   5  'Downward Diagonal
      Height          =   372
      Left            =   2760
      ScaleHeight     =   372
      ScaleWidth      =   552
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1140
      Width           =   552
   End
   Begin VB.CommandButton cmdChangeColor 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3420
      TabIndex        =   7
      Top             =   1140
      Width           =   492
   End
   Begin VB.ComboBox cboSliderWide 
      Height          =   336
      ItemData        =   "frmTestProperties.frx":002A
      Left            =   2760
      List            =   "frmTestProperties.frx":0037
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   9072
      Width           =   1752
   End
   Begin ColorControls.ColorSelector ColorSelector1 
      Height          =   3696
      Left            =   6840
      TabIndex        =   43
      Top             =   480
      Width           =   4524
      _ExtentX        =   7535
      _ExtentY        =   6519
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set..."
      Height          =   492
      Left            =   2220
      TabIndex        =   47
      Top             =   10452
      Width           =   1512
   End
   Begin VB.ComboBox cboStyle 
      Height          =   336
      ItemData        =   "frmTestProperties.frx":004A
      Left            =   2760
      List            =   "frmTestProperties.frx":0054
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   9492
      Width           =   1752
   End
   Begin VB.CheckBox chkSizeBig 
      Caption         =   "SizeBig"
      Height          =   372
      Left            =   2760
      TabIndex        =   31
      Top             =   7848
      Width           =   2592
   End
   Begin VB.ComboBox cboSliderOptionsAvailable 
      Height          =   336
      ItemData        =   "frmTestProperties.frx":0064
      Left            =   2760
      List            =   "frmTestProperties.frx":0074
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   8652
      Width           =   1752
   End
   Begin VB.ComboBox cboSliderParameter 
      Height          =   336
      ItemData        =   "frmTestProperties.frx":009C
      Left            =   2760
      List            =   "frmTestProperties.frx":00B2
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   8232
      Width           =   1752
   End
   Begin VB.CheckBox chkColorSelectionBoxVisible 
      Caption         =   "ColorSelectionBoxVisible"
      Height          =   372
      Left            =   2760
      TabIndex        =   10
      Top             =   1932
      Width           =   2592
   End
   Begin VB.CheckBox chkRoundedBoxes 
      Caption         =   "RoundedBoxes(*)"
      Height          =   372
      Left            =   2760
      TabIndex        =   30
      Top             =   7524
      Width           =   2592
   End
   Begin VB.CheckBox chkRememberPosition 
      Caption         =   "RememberPosition"
      Height          =   372
      Left            =   2760
      TabIndex        =   29
      Top             =   7224
      Width           =   2592
   End
   Begin VB.ComboBox cboRecentColorsColumns 
      Height          =   336
      ItemData        =   "frmTestProperties.frx":00E0
      Left            =   2760
      List            =   "frmTestProperties.frx":00E2
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   6828
      Width           =   1752
   End
   Begin VB.CheckBox chkPaletteTypeControlVisible 
      Caption         =   "PaletteTypeControlVisible"
      Height          =   372
      Left            =   2760
      TabIndex        =   23
      Top             =   6012
      Width           =   2592
   End
   Begin VB.TextBox txtPositionTop 
      Height          =   360
      Left            =   3660
      TabIndex        =   26
      Text            =   "[Default]"
      Top             =   6396
      Width           =   800
   End
   Begin VB.TextBox txtPositionLeft 
      Height          =   360
      Left            =   2760
      TabIndex        =   25
      Text            =   "[Default]"
      Top             =   6396
      Width           =   800
   End
   Begin VB.CheckBox chkHideLabels 
      Caption         =   "HideLabels(*)"
      Height          =   372
      Left            =   2760
      TabIndex        =   21
      Top             =   5340
      Width           =   2592
   End
   Begin VB.CheckBox chkModeless 
      Caption         =   "Modeless"
      Height          =   372
      Left            =   2760
      TabIndex        =   22
      Top             =   5664
      Width           =   2592
   End
   Begin VB.CheckBox chkFixedPalette 
      Caption         =   "FixedPalette(*)"
      Height          =   372
      Left            =   2760
      TabIndex        =   18
      Top             =   4368
      Width           =   2592
   End
   Begin VB.CheckBox chkHexControlVisible 
      Caption         =   "HexControlVisible"
      Height          =   372
      Left            =   2760
      TabIndex        =   19
      Top             =   4692
      Width           =   2592
   End
   Begin VB.TextBox txtDialogCaption 
      Height          =   360
      Left            =   2760
      TabIndex        =   15
      Text            =   "[Default]"
      Top             =   3312
      Width           =   1752
   End
   Begin VB.CheckBox chkConfirmationButtonsVisible 
      Caption         =   "ConfirmationButtonsVisible"
      Height          =   372
      Left            =   2760
      TabIndex        =   13
      Top             =   2904
      Width           =   2592
   End
   Begin VB.CheckBox chkColorValuesSectionVisible 
      Caption         =   "ColorValuesSectionVisible"
      Height          =   372
      Left            =   2760
      TabIndex        =   12
      Top             =   2580
      Width           =   2592
   End
   Begin VB.CheckBox chkColorSystemControlVisible 
      Caption         =   "ColorSystemControlVisible(*)"
      Height          =   372
      Left            =   2760
      TabIndex        =   11
      Top             =   2256
      Width           =   2592
   End
   Begin VB.ComboBox cboColorSystem 
      Height          =   336
      ItemData        =   "frmTestProperties.frx":00E4
      Left            =   2760
      List            =   "frmTestProperties.frx":00EE
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1572
      Width           =   1752
   End
   Begin VB.CheckBox chkDialogCaptionVisible 
      Caption         =   "DialogCaptionVisible"
      Height          =   372
      Left            =   2760
      TabIndex        =   16
      Top             =   3720
      Width           =   2592
   End
   Begin VB.CheckBox chkBasicColorsVisible 
      Caption         =   "BasicColorsVisible"
      Height          =   372
      Left            =   2760
      TabIndex        =   4
      Top             =   792
      Width           =   2592
   End
   Begin ColorControls.ColorDlg ColorDlg1 
      Left            =   240
      Top             =   180
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show dialog"
      Height          =   492
      Left            =   3900
      TabIndex        =   48
      Top             =   10452
      Width           =   1512
   End
   Begin VB.CommandButton cmdChangeBackColor 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3420
      TabIndex        =   3
      Top             =   420
      Width           =   492
   End
   Begin VB.PictureBox picBackColor 
      BorderStyle     =   0  'None
      FillStyle       =   5  'Downward Diagonal
      Height          =   372
      Left            =   2760
      ScaleHeight     =   372
      ScaleWidth      =   552
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   552
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "PointerType(*):"
      Height          =   252
      Left            =   1260
      TabIndex        =   40
      Top             =   9972
      Width           =   1392
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "Test ColorEyeDropper control:"
      ForeColor       =   &H00008000&
      Height          =   252
      Left            =   6540
      TabIndex        =   44
      Top             =   4560
      Width           =   5196
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Color(*):"
      Height          =   252
      Left            =   1260
      TabIndex        =   5
      Top             =   1200
      Width           =   1392
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "SliderWide(*):"
      Height          =   252
      Left            =   1260
      TabIndex        =   36
      Top             =   9132
      Width           =   1392
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "ColorSelector control. Properties with (*) apply to both controls:"
      ForeColor       =   &H00008000&
      Height          =   252
      Left            =   6480
      TabIndex        =   42
      Top             =   120
      Width           =   5196
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Style(*):"
      Height          =   252
      Left            =   1260
      TabIndex        =   38
      Top             =   9552
      Width           =   1392
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "SliderOptionsAvailable(*):"
      Height          =   252
      Left            =   60
      TabIndex        =   34
      Top             =   8712
      Width           =   2592
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "SliderParameter(*)"
      Height          =   252
      Left            =   780
      TabIndex        =   32
      Top             =   8292
      Width           =   1872
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "RecentColorsColumns:"
      Height          =   252
      Left            =   840
      TabIndex        =   27
      Top             =   6876
      Width           =   1812
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "PositionLeft, Top:"
      Height          =   252
      Left            =   780
      TabIndex        =   24
      Top             =   6444
      Width           =   1872
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "DialogCaption:"
      Height          =   252
      Left            =   1260
      TabIndex        =   14
      Top             =   3360
      Width           =   1392
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "ColorDlg control properties (in alphabetical order):"
      ForeColor       =   &H00008000&
      Height          =   252
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   4392
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "ColorSystem(*):"
      Height          =   252
      Left            =   1260
      TabIndex        =   8
      Top             =   1632
      Width           =   1392
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "BackColor(*):"
      Height          =   252
      Left            =   1260
      TabIndex        =   1
      Top             =   480
      Width           =   1392
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuSetComplete 
         Caption         =   "Complete"
      End
      Begin VB.Menu mnuSetCompact 
         Caption         =   "Compact"
      End
      Begin VB.Menu mnuSetSimple 
         Caption         =   "Simple"
      End
   End
End
Attribute VB_Name = "frmTestProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IEyeDropperNotification

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Sub cboColorSystem_Click()
    ColorSelector1.ColorSystem = cboColorSystem.ListIndex
End Sub

Private Sub cboPointerType_Click()
    ColorSelector1.PointerType = cboPointerType.ListIndex
End Sub

Private Sub cboSliderOptionsAvailable_Click()
    ColorSelector1.SliderOptionsAvailable = cboSliderOptionsAvailable.ListIndex
End Sub

Private Sub cboSliderParameter_Click()
    ColorSelector1.SliderParameter = cboSliderParameter.ListIndex
End Sub

Private Sub cboSliderWide_Click()
    ColorSelector1.SliderWide = cboSliderWide.ListIndex
End Sub

Private Sub cboStyle_Click()
    ColorSelector1.Style = cboStyle.ListIndex
End Sub

Private Sub chkColorSystemControlVisible_Click()
    ColorSelector1.ColorSystemControlVisible = (chkColorSystemControlVisible.Value = 1)
End Sub

Private Sub chkFixedPalette_Click()
    ColorSelector1.FixedPalette = (chkFixedPalette.Value = 1)
End Sub

Private Sub chkHideLabels_Click()
    ColorSelector1.HideLabels = (chkHideLabels.Value = 1)
End Sub

Private Sub chkRoundedBoxes_Click()
    ColorSelector1.RoundedBoxes = (chkRoundedBoxes.Value = 1)
End Sub

Private Sub cmdChangeBackColor_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.Color = picBackColor.BackColor
    If oDlg.Show Then
        picBackColor.BackColor = oDlg.Color
        ColorSelector1.BackColor = picBackColor.BackColor
    End If
End Sub

Private Sub cmdChangeColor_Click()
    Dim iFrm As New frmDropDown
    
    iFrm.Color = picColor.BackColor
    iFrm.SetTransparency
    iFrm.Move Me.Left + cmdChangeColor.Left, Me.Top + cmdChangeColor.Top + cmdChangeColor.Height + (Me.Height - Me.ScaleHeight)
    iFrm.Show , Me
    Do While IsFormLoaded(iFrm)
        DoEvents
    Loop
    If iFrm.ColorSet Then
        picColor.BackColor = iFrm.Color
        ColorSelector1.Color = picColor.BackColor
    End If
End Sub

Private Function IsFormLoaded(nForm As Form) As Boolean
    Dim frm As Form
    
    For Each frm In Forms
        If frm Is nForm Then
            IsFormLoaded = True
            Exit For
        End If
    Next
End Function

Private Sub cmdDocs_Click()
    Const SW_SHOWMAXIMIZED = 3
    ShellExecute 0&, "OPEN", App.Path & "\..\docs\ColorControls_reference.html", "", "", SW_SHOWMAXIMIZED
End Sub

Private Sub cmdTestEyeDropper_Click()
    picEyeDropper.Tag = picEyeDropper.BackColor
    If EyeDropper1.StartDropper Then
        picEyeDropper.BackColor = EyeDropper1.Color
    Else
        picEyeDropper.BackColor = picEyeDropper.Tag
        picEyeDropper.Refresh
    End If
End Sub

Private Sub IEyeDropperNotification_ColorUnderMouseChange(ByVal nColor As Long)
    picEyeDropper.BackColor = nColor
End Sub

Private Sub Command1_Click()
    PopupMenu mnuPopup
End Sub

Private Sub Form_Load()
    Dim ctl As Control
    Dim iRgn As Long
    Dim c As Long
    
    For Each ctl In Me.Controls
        If TypeOf ctl Is PictureBox Then
            iRgn = CreateRoundRectRgn(0, 0, Me.ScaleX(ctl.Width, Me.ScaleMode, vbPixels), Me.ScaleY(ctl.Height, Me.ScaleMode, vbPixels), 12, 12)
            SetWindowRgn ctl.hWnd, iRgn, True
            DeleteObject iRgn
        End If
    Next
    cboColorSystem.ListIndex = 0
    
    For c = 0 To 10
        cboRecentColorsColumns.AddItem c
    Next
    SetConrolsFromProperties
End Sub

Private Sub cmdShow_Click()
    If picBackColor.BackColor <> vbButtonFace Then
        ColorDlg1.BackColor = picBackColor.BackColor
    End If
    ColorDlg1.BasicColorsVisible = (chkBasicColorsVisible.Value = 1)
    ColorDlg1.DialogCaptionVisible = (chkDialogCaptionVisible.Value = 1)
    ColorDlg1.Color = picColor.BackColor
    ColorDlg1.ColorSystem = cboColorSystem.ListIndex
    ColorDlg1.ColorSystemControlVisible = (chkColorSystemControlVisible.Value = 1)
    ColorDlg1.ColorValuesSectionVisible = (chkColorValuesSectionVisible.Value = 1)
    ColorDlg1.ConfirmationButtonsVisible = (chkConfirmationButtonsVisible.Value = 1)
    If txtDialogCaption.Text <> "[Default]" Then
        ColorDlg1.DialogCaption = txtDialogCaption.Text
    End If
    ColorDlg1.EyeDropperVisible = (chkEyeDropperVisible.Value = 1)
    ColorDlg1.FixedPalette = (chkFixedPalette.Value = 1)
    ColorDlg1.HexControlVisible = (chkHexControlVisible.Value = 1)
    ColorDlg1.HexFormatVB = (chkHexFormatVB.Value = 1)
    ColorDlg1.HideLabels = (chkHideLabels.Value = 1)
    ColorDlg1.PointerType = cboPointerType.ListIndex
    ColorDlg1.Modeless = (chkModeless.Value = 1)
    ColorDlg1.PaletteTypeControlVisible = (chkPaletteTypeControlVisible.Value = 1)
    If txtPositionLeft.Text <> "[Default]" Then
        ColorDlg1.PositionLeft = Val(txtPositionLeft.Text)
    End If
    If txtPositionTop.Text <> "[Default]" Then
        ColorDlg1.PositionTop = Val(txtPositionTop.Text)
    End If
    ColorDlg1.RecentColorsColumns = Val(cboRecentColorsColumns.Text)
    ColorDlg1.RememberPosition = (chkRememberPosition.Value = 1)
    ColorDlg1.RoundedBoxes = (chkRoundedBoxes.Value = 1)
    ColorDlg1.ColorSelectionBoxVisible = (chkColorSelectionBoxVisible.Value = 1)
    ColorDlg1.SizeBig = (chkSizeBig.Value = 1)
    ColorDlg1.SliderParameter = cboSliderParameter.ListIndex
    ColorDlg1.SliderOptionsAvailable = cboSliderOptionsAvailable.ListIndex
    ColorDlg1.SliderWide = cboSliderWide.ListIndex
    ColorDlg1.Style = cboStyle.ListIndex
    
    If ColorDlg1.Show Then
        picColor.BackColor = ColorDlg1.Color
        ColorSelector1.Color = picColor.BackColor
    End If
End Sub

Private Sub mnuSetCompact_Click()
    ColorDlg1.SetCompact
    SetConrolsFromProperties
End Sub

Private Sub mnuSetComplete_Click()
    ColorDlg1.SetComplete
    SetConrolsFromProperties
End Sub

Private Sub mnuSetSimple_Click()
    ColorDlg1.SetSimple
    SetConrolsFromProperties
End Sub

Private Sub picBackColor_Paint()
    If picBackColor.BackColor = vbButtonFace Then
        picBackColor.Line (0, 0)-(picBackColor.ScaleWidth - 1, picBackColor.ScaleHeight - 1), vbButtonFace, B
    End If
End Sub

Private Sub SetConrolsFromProperties()
    picBackColor.BackColor = ColorDlg1.BackColor: ColorSelector1.BackColor = picBackColor.BackColor
    chkBasicColorsVisible.Value = Abs(ColorDlg1.BasicColorsVisible)
    chkDialogCaptionVisible.Value = Abs(ColorDlg1.DialogCaptionVisible)
    picColor.BackColor = ColorDlg1.Color: ColorSelector1.Color = picColor.BackColor
    cboColorSystem.ListIndex = ColorDlg1.ColorSystem
    chkColorSystemControlVisible.Value = Abs(ColorDlg1.ColorSystemControlVisible)
    chkColorValuesSectionVisible.Value = Abs(ColorDlg1.ColorValuesSectionVisible)
    chkConfirmationButtonsVisible.Value = Abs(ColorDlg1.ConfirmationButtonsVisible)
    If ColorDlg1.DialogCaption = "" Then
        txtDialogCaption.Text = "[Default]"
    Else
        txtDialogCaption.Text = ColorDlg1.DialogCaption
    End If
    chkEyeDropperVisible.Value = Abs(ColorDlg1.EyeDropperVisible)
    chkFixedPalette.Value = Abs(ColorDlg1.FixedPalette)
    chkHexControlVisible.Value = Abs(ColorDlg1.HexControlVisible)
    chkHexFormatVB.Value = Abs(ColorDlg1.HexFormatVB)
    chkHideLabels.Value = Abs(ColorDlg1.HideLabels)
    chkModeless.Value = Abs(ColorDlg1.Modeless)
    chkPaletteTypeControlVisible.Value = Abs(ColorDlg1.PaletteTypeControlVisible)
    If ColorDlg1.PositionLeft = 0 Then
        txtPositionLeft.Text = "[Default]"
    Else
        txtPositionLeft.Text = ColorDlg1.PositionLeft
    End If
    If ColorDlg1.PositionTop = 0 Then
        txtPositionTop.Text = "[Default]"
    Else
        txtPositionTop.Text = ColorDlg1.PositionTop
    End If
    cboRecentColorsColumns.Text = ColorDlg1.RecentColorsColumns
    chkRememberPosition.Value = Abs(ColorDlg1.RememberPosition)
    chkRoundedBoxes.Value = Abs(ColorDlg1.RoundedBoxes)
    chkColorSelectionBoxVisible.Value = Abs(ColorDlg1.ColorSelectionBoxVisible)
    chkSizeBig.Value = Abs(ColorDlg1.SizeBig)
    cboSliderParameter.ListIndex = ColorDlg1.SliderParameter
    cboSliderOptionsAvailable.ListIndex = ColorDlg1.SliderOptionsAvailable
    cboSliderWide.ListIndex = ColorDlg1.SliderWide
    cboStyle.ListIndex = ColorDlg1.Style
    cboPointerType.ListIndex = ColorDlg1.PointerType
End Sub

Private Sub picEyeDropper_Paint()
    If picEyeDropper.BackColor = vbButtonFace Then
        picEyeDropper.Line (0, 0)-(picEyeDropper.ScaleWidth - 1, picEyeDropper.ScaleHeight - 1), vbButtonFace, B
    End If
End Sub
