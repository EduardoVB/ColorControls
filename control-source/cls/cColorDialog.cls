VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColorDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Color dialog replacement."
Option Explicit

Public Event Change()
Attribute Change.VB_Description = "Occurs when the selected color changes in the dialog (the color does not need to be accepted)."
Public Event ColorSet()
Attribute ColorSet.VB_Description = "Occurs when a color change was accepted."
Public Event Hided()
Attribute Hided.VB_Description = "Occurs when the dialog was hided (typically used in modeless situations)."
Public Event GetLocalizedText(ByVal LanguageID As Long, ByVal SubLanguageID As Long, ByVal TextID As Long, ByRef Text As String)
Attribute GetLocalizedText.VB_Description = "Occurs when setting a text for the IU. It allows to customize captions and texts."

Private Enum CDCaptionsIDConstants
    cdCaptionHue ' Hue
    cdCaptionLum ' Lum
    cdCaptionSat ' Sat
    cdCaptionRed ' Red
    cdCaptionGreen ' Green
    cdCaptionBlue ' Blue
    cdCaptionVal ' Val
    cdCaptionFixed ' Fixed
    cdCaptionFixedToolTipText ' Reflects color changes visually or not
    cdCaptionSliderParameterToolTipText ' Select parameter
    cdCaptionRecent ' recent
    cdCaptionNew     ' new
    cdCaptionPrevious  ' previous
    cdCaptionColor ' Color
    cdCaptionHex ' Hex
    cdCaptionInvalidColorMessage ' The Hex color is not valid.
    cdCaptionOK ' OK
    cdCaptionCancel ' Cancel
    cdCaptionMenuForgetRecent ' Forget
    cdCaptionMenuClearAllRecent ' Clear all recent
    cdCaptionMode ' Mode
    cdCaptionToolTipMouseWheelBeginning ' Hold the Control key down to navigate
    cdCaptionToolTipMouseWheelEnding ' with the mouse wheel
End Enum

Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long

Private WithEvents mFrm As frmColorDialog
Attribute mFrm.VB_VarHelpID = -1

Private mActiveForm As Object

' properties
Private mBackColor As Long
Private mBasicColorsVisible As Boolean
Private mDialogCaptionVisible As Boolean
Private mColor As Long
Private mColorSelectionBoxVisible As Boolean
Private mColorSystem As CDColorSystemConstants
Private mColorSystemControlVisible As Boolean
Private mColorValuesSectionVisible As Boolean
Private mConfirmationButtonsVisible As Boolean
Private mContext As String
Private mDialogCaption As String
Private mEyeDropperVisible As Boolean
Private mFixedPalette As Boolean
Private mPaletteTypeControlVisible As Boolean
Private mHexControlVisible As Boolean
Private mHexFormatVB As Boolean
Private mHideLabels As Boolean
Private mModeless As Boolean
Private mRecentColorsColumns As Long
Private mRememberPosition As Boolean
Private mRoundedBoxes As Boolean
Private mSliderParameter As CDSliderParameterConstants
Private mSliderOptionsAvailable As CDSliderOptionsAvailableConstants
Private mSizeBig As Boolean
Private mSliderWide As CDYesNoAutoConstants
Private mStyle As CDStyleConstants
Private mPositionLeft As Single
Private mPositionTop As Single
' read only (return values)
Private mCanceled As Boolean
Private mChanged As Boolean

Private mColorSet As Boolean
Private mTag As String

Public Function Show(Optional ByVal nStyleBox As Variant) As Boolean
Attribute Show.VB_Description = "Shows the dialog."
    Dim iColorPrev As Long
    Dim iActiveWindowHwnd As Long
    Dim iActiveFormHwnd As Long
    Static sShowing As Boolean
    
    If sShowing And (Not mFrm Is Nothing) Then
        On Error Resume Next
        mFrm.Show
        mFrm.ZOrder
        Exit Function
    End If
    
    If Not IsMissing(nStyleBox) Then
        If VarType(nStyleBox) = vbBoolean Then
            If nStyleBox Then
                mStyle = cdStyleBox
            End If
        End If
    End If
    mCanceled = False
    mChanged = False
    Set mFrm = New frmColorDialog
    
    'Set mActiveForm = Screen.ActiveForm
    mFrm.RememberPosition = mRememberPosition
    mFrm.Modeless = mModeless
    mFrm.Style = mStyle
    If Not mActiveForm Is Nothing Then
        mFrm.ActiveFormName = mActiveForm.Name
        iActiveFormHwnd = mActiveForm.hWnd
    End If
    
    mFrm.HexControlVisible = mHexControlVisible
    mFrm.HexFormatVB = mHexFormatVB
    mFrm.ColorValuesSectionVisible = mColorValuesSectionVisible
    mFrm.RecentColorsColumns = mRecentColorsColumns
    mFrm.SizeBig = mSizeBig
    mFrm.SliderOptionsAvailable = mSliderOptionsAvailable
    mFrm.PaletteTypeControlVisible = mPaletteTypeControlVisible
    mFrm.ColorSystemControlVisible = mColorSystemControlVisible
    mFrm.FixedPalette = mFixedPalette
    mFrm.ColorSystem = mColorSystem
    mFrm.SliderParameter = mSliderParameter
    mFrm.Context = mContext
    mFrm.DialogCaptionVisible = mDialogCaptionVisible
    mFrm.ColorSelectionBoxVisible = mColorSelectionBoxVisible
    mFrm.SliderWide = mSliderWide
    mFrm.HideLabels = mHideLabels
    mFrm.RoundedBoxes = mRoundedBoxes
    mFrm.BasicColorsVisible = mBasicColorsVisible
    mFrm.EyeDropperVisible = mEyeDropperVisible
    mFrm.ConfirmationButtonsVisible = mConfirmationButtonsVisible
    If mPositionLeft <> 0 Then mFrm.Left = mPositionLeft
    If mPositionTop <> 0 Then mFrm.Top = mPositionTop
    If (Not mDialogCaptionVisible) And (mBackColor = vbButtonFace) Then
        mFrm.BackColor2 = vbWhite
    Else
        mFrm.BackColor2 = mBackColor
    End If
    If mColorSet Then
        mFrm.CurrentColor = mColor
    End If
    If mDialogCaption <> "" Then
        mFrm.Caption = " " & mDialogCaption
    End If
    iColorPrev = mColor
    mFrm.SetBackColor
    If mModeless Then
        On Error Resume Next
        If mActiveForm Is Nothing Then
            iActiveWindowHwnd = GetActiveWindow
            If iActiveWindowHwnd <> 0 Then
                SetOwner mFrm.hWnd, iActiveWindowHwnd
            End If
            mFrm.Show
        Else
            mFrm.Show , mActiveForm
            If Err.Number Then
                mFrm.Show
            End If
        End If
        Exit Function
    ElseIf Not mConfirmationButtonsVisible Then
        On Error Resume Next
        If mActiveForm Is Nothing Then
            iActiveWindowHwnd = GetActiveWindow
            If iActiveWindowHwnd <> 0 Then
                SetOwner mFrm.hWnd, iActiveWindowHwnd
            End If
            mFrm.Show
        Else
            mFrm.Show , mActiveForm
        End If
        If Err.Number Then
            On Error GoTo 0
            mFrm.Show vbModal
        End If
        On Error GoTo 0
    Else
        mFrm.Show vbModal
    End If
    sShowing = True
    If Not mConfirmationButtonsVisible Then
        If Not mActiveForm Is Nothing Then
            Do While IsFormLoaded(mFrm) And IsWindowVisible(iActiveFormHwnd) <> 0
                DoEvents
                If IsFormLoaded(mActiveForm) Then
                If mActiveForm.Visible = False Then Exit Do
                End If
            Loop
        Else
            Do While IsFormLoaded(mFrm)
                DoEvents
            Loop
        End If
    End If
    sShowing = False
    If mFrm Is Nothing Then Exit Function
    mCanceled = Not mFrm.OKPressed
    If Not mCanceled Then
        If mColorSet Then
            mChanged = mFrm.SelectedColor <> iColorPrev
        Else
            mChanged = True
        End If
        If mChanged Then
            mColor = mFrm.SelectedColor
        End If
    End If
    Show = mChanged
    Set mFrm = Nothing
End Function

Public Sub Hide()
Attribute Hide.VB_Description = "Hides the dialog when it is displayed modeless."
    If Not mFrm Is Nothing Then
        mFrm.Modeless = False
        Unload mFrm
        Set mFrm = Nothing
    End If
End Sub

Private Function IsFormLoaded(nFrm As Object) As Boolean
    Dim f As Form
    
    For Each f In Forms
        If f Is nFrm Then
            IsFormLoaded = True
        End If
    Next
End Function

Public Property Get Color() As OLE_COLOR
Attribute Color.VB_Description = "Returns/sets the current color."
    Color = mColor
End Property

Public Property Let Color(ByVal nValue As OLE_COLOR)
    If Not IsValidOLE_COLOR(nValue) Then
        Err.Raise 380, TypeName(Me)
        Exit Property
    End If
    mColor = nValue
    mColorSet = True
End Property


Public Property Get DialogCaption() As String
Attribute DialogCaption.VB_Description = "Returns/sets the caption of the dialog's title bar."
    DialogCaption = mDialogCaption
End Property

Public Property Let DialogCaption(nValue As String)
    mDialogCaption = nValue
End Property


Public Property Get Canceled() As Boolean
Attribute Canceled.VB_Description = "Returns True if the dialog was canceled."
    Canceled = mCanceled
End Property


Public Property Get Changed() As Boolean
Attribute Changed.VB_Description = "Returns true if the color was changed."
    Changed = mChanged
End Property


Public Property Get HexControlVisible() As Boolean
Attribute HexControlVisible.VB_Description = "Returns/sets a value that determines if the hexadecimal text box is visible."
    HexControlVisible = mHexControlVisible
End Property

Public Property Let HexControlVisible(ByVal nValue As Boolean)
    mHexControlVisible = nValue
End Property


Public Property Get HexFormatVB() As Boolean
Attribute HexFormatVB.VB_Description = "Returns/sets a value that determines the format of the hexadecimal text box, for web or VB."
    HexFormatVB = mHexFormatVB
End Property

Public Property Let HexFormatVB(ByVal nValue As Boolean)
    mHexFormatVB = nValue
End Property


Public Property Get ColorValuesSectionVisible() As Boolean
Attribute ColorValuesSectionVisible.VB_Description = "Returns/sets a value that determines if the boxes to enter the color values by hand are visible or not."
    ColorValuesSectionVisible = mColorValuesSectionVisible
End Property

Public Property Let ColorValuesSectionVisible(ByVal nValue As Boolean)
    mColorValuesSectionVisible = nValue
End Property


Public Property Get RecentColorsColumns() As Long
Attribute RecentColorsColumns.VB_Description = "Returns/sets a value that determines the number of columns of the 'Recent colors'. Set to 0 for none."
    RecentColorsColumns = mRecentColorsColumns
End Property

Public Property Let RecentColorsColumns(ByVal nValue As Long)
    If (nValue < -1) Or (nValue > 10) Then
        Err.Raise 380, TypeName(Me)
        Exit Property
    End If
    mRecentColorsColumns = nValue
End Property


Public Property Get SizeBig() As Boolean
Attribute SizeBig.VB_Description = "Returns/sets a value that determines the size of the dialog, big or normal."
    SizeBig = mSizeBig
End Property

Public Property Let SizeBig(ByVal nValue As Boolean)
    mSizeBig = nValue
End Property


Public Property Get SliderOptionsAvailable() As CDSliderOptionsAvailableConstants
Attribute SliderOptionsAvailable.VB_Description = "Returns/sets a value that determines which parameters for the slider the user will have available to choose from."
    SliderOptionsAvailable = mSliderOptionsAvailable
End Property

Public Property Let SliderOptionsAvailable(ByVal nValue As CDSliderOptionsAvailableConstants)
    mSliderOptionsAvailable = nValue
End Property


Public Property Get PaletteTypeControlVisible() As Boolean
Attribute PaletteTypeControlVisible.VB_Description = "Returns/sets a value that determines if the control for changing the palette type is visible or not."
    PaletteTypeControlVisible = mPaletteTypeControlVisible
End Property

Public Property Let PaletteTypeControlVisible(ByVal nValue As Boolean)
    mPaletteTypeControlVisible = nValue
End Property


Public Property Get ColorSystemControlVisible() As Boolean
Attribute ColorSystemControlVisible.VB_Description = "Returns/sets a value that determines if the user is able to change the color system."
    ColorSystemControlVisible = mColorSystemControlVisible
End Property

Public Property Let ColorSystemControlVisible(ByVal nValue As Boolean)
    mColorSystemControlVisible = nValue
End Property


Public Property Get FixedPalette() As Boolean
Attribute FixedPalette.VB_Description = "Returns/sets a value that determines if the color palette (or the slidder in some configurations) colors change with the setting of the third partameter."
    FixedPalette = mFixedPalette
End Property

Public Property Let FixedPalette(ByVal nValue As Boolean)
    mFixedPalette = nValue
End Property


Public Property Get ColorSystem() As CDColorSystemConstants
Attribute ColorSystem.VB_Description = "Returns/sets a value that determines the color system: HSV or HSL."
    ColorSystem = mColorSystem
End Property

Public Property Let ColorSystem(ByVal nValue As CDColorSystemConstants)
    mColorSystem = nValue
End Property


Public Property Get SliderParameter() As CDSliderParameterConstants
Attribute SliderParameter.VB_Description = "Returns/sets a value that determines the slidder parameter."
    SliderParameter = mSliderParameter
End Property

Public Property Let SliderParameter(ByVal nValue As CDSliderParameterConstants)
    mSliderParameter = nValue
End Property


Public Property Get DialogCaptionVisible() As Boolean
Attribute DialogCaptionVisible.VB_Description = "Returns/sets a value that determines if the dialog's title bar is visible. When True, the default BackColor is changed to white."
    DialogCaptionVisible = mDialogCaptionVisible
End Property

Public Property Let DialogCaptionVisible(ByVal nValue As Boolean)
    mDialogCaptionVisible = nValue
End Property


Public Property Get ConfirmationButtonsVisible() As Boolean
Attribute ConfirmationButtonsVisible.VB_Description = "Returns/sets a value that determines if the OK and Cancel buttons are visible."
    ConfirmationButtonsVisible = mConfirmationButtonsVisible
End Property

Public Property Let ConfirmationButtonsVisible(ByVal nValue As Boolean)
    mConfirmationButtonsVisible = nValue
End Property


Public Property Get ColorSelectionBoxVisible() As Boolean
Attribute ColorSelectionBoxVisible.VB_Description = "Returns/sets a value that determines if the box showing the selected color is visible."
    ColorSelectionBoxVisible = mColorSelectionBoxVisible
End Property

Public Property Let ColorSelectionBoxVisible(ByVal nValue As Boolean)
    mColorSelectionBoxVisible = nValue
End Property


Public Property Get SliderWide() As CDYesNoAutoConstants
Attribute SliderWide.VB_Description = "Returns/sets a value that determines the width of the slider control between wide or narrow."
    SliderWide = mSliderWide
End Property

Public Property Let SliderWide(ByVal nValue As CDYesNoAutoConstants)
    mSliderWide = nValue
End Property


Public Property Get HideLabels() As Boolean
Attribute HideLabels.VB_Description = "Returns/sets a value that determines if the labels are visible."
    HideLabels = mHideLabels
End Property

Public Property Let HideLabels(ByVal nValue As Boolean)
    mHideLabels = nValue
End Property


Public Property Get Tag() As String
Attribute Tag.VB_Description = "Stores any extra data needed for your program."
    Tag = mTag
End Property

Public Property Let Tag(nValue As String)
    mTag = nValue
End Property


Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color."
    BackColor = mBackColor
End Property

Public Property Let BackColor(ByVal nValue As OLE_COLOR)
    If nValue <> mBackColor Then
        mBackColor = nValue
    End If
End Property


Public Property Get Context() As String
Attribute Context.VB_Description = "Returns/sets a context for saving the user settings and states."
    Context = mContext
End Property

Public Property Let Context(nValue As String)
    mContext = nValue
End Property


Public Property Get RememberPosition() As Boolean
Attribute RememberPosition.VB_Description = "Returns/sets a value that determines whether the position of the dialog will be remembered next time."
    RememberPosition = mRememberPosition
End Property

Public Property Let RememberPosition(ByVal nValue As Boolean)
    mRememberPosition = nValue
End Property


Public Property Get Modeless() As Boolean
Attribute Modeless.VB_Description = "Returns/sets a value that determines if the dialog is displayed modally or not."
    Modeless = mModeless
End Property

Public Property Let Modeless(ByVal nValue As Boolean)
    mModeless = nValue
End Property


Public Property Get Style() As CDStyleConstants
Attribute Style.VB_Description = "Returns/sets a value that determines the color palette style, between wheel or box."
    Style = mStyle
End Property

Public Property Let Style(ByVal nValue As CDStyleConstants)
    mStyle = nValue
End Property


Public Property Get RoundedBoxes() As Boolean
Attribute RoundedBoxes.VB_Description = "Returns/sets a value that determines if some control borders are rounded."
    RoundedBoxes = mRoundedBoxes
End Property

Public Property Let RoundedBoxes(ByVal nValue As Boolean)
    mRoundedBoxes = nValue
End Property


Public Property Get BasicColorsVisible() As Boolean
Attribute BasicColorsVisible.VB_Description = "Determines whether the ""Basic colors"" section is visible or not."
    BasicColorsVisible = mBasicColorsVisible
End Property

Public Property Let BasicColorsVisible(ByVal nValue As Boolean)
    mBasicColorsVisible = nValue
End Property


Public Property Get EyeDropperVisible() As Boolean
Attribute EyeDropperVisible.VB_Description = "Returns/sets a value that determines if the 'eye dropper' feature used to pick a color from the entire screen is available."
    EyeDropperVisible = mEyeDropperVisible
End Property

Public Property Let EyeDropperVisible(ByVal nValue As Boolean)
    mEyeDropperVisible = nValue
End Property


Public Property Get PositionLeft() As Single
Attribute PositionLeft.VB_Description = "Returns/sets a value that determines the Left position of the dialog."
    PositionLeft = mPositionLeft
End Property

Public Property Let PositionLeft(ByVal nValue As Single)
    If nValue <> mPositionLeft Then
        mPositionLeft = nValue
    End If
End Property


Public Property Get PositionTop() As Single
Attribute PositionTop.VB_Description = "Returns/sets a value that determines the Top position of the dialog."
    PositionTop = mPositionTop
End Property

Public Property Let PositionTop(ByVal nValue As Single)
    If nValue <> mPositionTop Then
        mPositionTop = nValue
    End If
End Property



Public Sub Move(ByVal nPositionLeft As Single, ByVal nPositionTop As Single)
Attribute Move.VB_Description = "Sets the initial position of the dialog."
    mPositionLeft = nPositionLeft
    mPositionTop = nPositionTop
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

Private Sub Class_Initialize()
    mBackColor = cPropDefault_ColorDialog_BackColor
    mBasicColorsVisible = cPropDefault_ColorDialog_BasicColorsVisible
    mColor = cPropDefault_ColorDialog_Color
    mColorSelectionBoxVisible = cPropDefault_ColorDialog_ColorSelectionBoxVisible
    mColorSystem = cPropDefault_ColorDialog_ColorSystem
    mColorSystemControlVisible = cPropDefault_ColorDialog_ColorSystemControlVisible
    mColorValuesSectionVisible = cPropDefault_ColorDialog_ColorValuesSectionVisible
    mConfirmationButtonsVisible = cPropDefault_ColorDialog_ConfirmationButtonsVisible
    mDialogCaptionVisible = cPropDefault_ColorDialog_DialogCaptionVisible
    mEyeDropperVisible = cPropDefault_ColorDialog_EyeDropperVisible
    mFixedPalette = cPropDefault_ColorDialog_FixedPalette
    mHexControlVisible = cPropDefault_ColorDialog_HexControlVisible
    mHexFormatVB = cPropDefault_ColorDialog_HexFormatVB
    mHideLabels = cPropDefault_ColorDialog_HideLabels
    mModeless = cPropDefault_ColorDialog_Modeless
    mRecentColorsColumns = cPropDefault_ColorDialog_RecentColorsColumns
    mRememberPosition = cPropDefault_ColorDialog_RememberPosition
    mRoundedBoxes = cPropDefault_ColorDialog_RoundedBoxes
    mPaletteTypeControlVisible = cPropDefault_ColorDialog_PaletteTypeControlVisible
    mSliderParameter = cPropDefault_ColorDialog_SliderParameter
    mSliderOptionsAvailable = cPropDefault_ColorDialog_SliderOptionsAvailable
    mSizeBig = cPropDefault_ColorDialog_SizeBig
    mSliderWide = cPropDefault_ColorDialog_SliderWide
    mStyle = cPropDefault_ColorDialog_Style
    mPositionLeft = 0
    mPositionTop = 0
End Sub

Public Sub SetComplete(Optional SizeBig As Boolean = False, Optional BasicColorsVisible As Boolean = True, Optional nRecentColorsColumns As Long = 2, Optional SliderOptionsAvailable As CDSliderOptionsAvailableConstants = cdSliderOptionsAll, Optional PaletteTypeControlVisible As Boolean = True, Optional ColorSystemControlVisible As Boolean = True, Optional EyeDropperVisible As Boolean = True)
Attribute SetComplete.VB_Description = "Helper method to set several properties to get a 'Compplete' configuration in one shot."
    mSizeBig = SizeBig
    mBasicColorsVisible = BasicColorsVisible
    mRecentColorsColumns = nRecentColorsColumns
    mEyeDropperVisible = EyeDropperVisible
    mSliderOptionsAvailable = SliderOptionsAvailable
    mPaletteTypeControlVisible = PaletteTypeControlVisible
    mColorSystemControlVisible = ColorSystemControlVisible
    mColorValuesSectionVisible = True
    mHexControlVisible = True
End Sub

Public Sub SetCompact(Optional nRecentColorsColumns As Long = 0, Optional ColorValuesSectionVisible As Boolean = False, Optional SliderOptionsAvailable As CDSliderOptionsAvailableConstants = cdSliderOptionsNone, Optional nDialogCaptionVisible As Boolean = True, Optional nConfirmationButtonsVisible As Boolean = True, Optional nColorSelectionBoxVisible As Boolean = True, Optional nSliderWide As CDYesNoAutoConstants = cdYNAuto, Optional nHideLabels As Boolean = False)
Attribute SetCompact.VB_Description = "Helper method to set several properties to get a 'Compact' configuration in one shot."
    mColorValuesSectionVisible = ColorValuesSectionVisible
    mSliderOptionsAvailable = SliderOptionsAvailable
    mDialogCaptionVisible = nDialogCaptionVisible
    mConfirmationButtonsVisible = nConfirmationButtonsVisible
    mColorSelectionBoxVisible = nColorSelectionBoxVisible
    mSliderWide = nSliderWide
    mRecentColorsColumns = nRecentColorsColumns
    mHideLabels = nHideLabels
    mBasicColorsVisible = False
    mEyeDropperVisible = False
    mColorSystemControlVisible = False
    mPaletteTypeControlVisible = False
    mHexControlVisible = False
End Sub

Public Sub SetSimple(Optional nRecentColorsColumns As Long = 2, Optional nDialogCaptionVisible As Boolean = False, Optional nConfirmationButtonsVisible As Boolean, Optional nColorSelectionBoxVisible As Boolean, Optional nSliderWide As CDYesNoAutoConstants = cdYNAuto, Optional nHideLabels As Boolean = False)
Attribute SetSimple.VB_Description = "Helper method to set several properties to get a 'Simple' configuration in one shot."
    mColorValuesSectionVisible = False
    mSliderOptionsAvailable = cdSliderOptionsNone
    mDialogCaptionVisible = nDialogCaptionVisible
    mConfirmationButtonsVisible = nConfirmationButtonsVisible
    mColorSelectionBoxVisible = nColorSelectionBoxVisible
    mSliderWide = nSliderWide
    mRecentColorsColumns = nRecentColorsColumns
    mHideLabels = nHideLabels
    mBasicColorsVisible = False
    mEyeDropperVisible = False
    mColorSystemControlVisible = False
    mPaletteTypeControlVisible = False
    mHexControlVisible = False
End Sub

Private Sub Class_Terminate()
    Set mActiveForm = Nothing
End Sub

Private Sub mFrm_Change()
    mColor = mFrm.ColorSelector1.Color
    RaiseEvent Change
End Sub

Private Sub mFrm_ColorSet()
    mColor = mFrm.ColorSelector1.Color
    RaiseEvent ColorSet
End Sub

Private Sub mFrm_GetLocalizedText(ByVal LanguageID As Long, ByVal SubLanguageID As Long, ByVal TextID As Long, Text As String)
    RaiseEvent GetLocalizedText(LanguageID, SubLanguageID, TextID, Text)
End Sub

Private Sub mFrm_Hided()
    RaiseEvent Hided
End Sub

Private Function SetOwner(ByVal HwndWindow As Long, ByVal hwndOwner As Long) As Long
    Const GWL_HWNDPARENT As Long = (-8)
    
    SetOwner = SetWindowLong(HwndWindow, GWL_HWNDPARENT, hwndOwner)
End Function

Public Property Set ActiveForm(nForm As Object)
Attribute ActiveForm.VB_Description = "Sets the active form"
    Set mActiveForm = nForm
End Property
