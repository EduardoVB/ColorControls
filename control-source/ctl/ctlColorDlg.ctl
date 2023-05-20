VERSION 5.00
Begin VB.UserControl ColorDlg 
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ctlColorDlg.ctx":0000
   ScaleHeight     =   2880
   ScaleWidth      =   3840
   ToolboxBitmap   =   "ctlColorDlg.ctx":182A
End
Attribute VB_Name = "ColorDlg"
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

Private WithEvents mDlg As ColorDialog
Attribute mDlg.VB_VarHelpID = -1

Private Sub mDlg_Change()
    RaiseEvent Change
End Sub

Private Sub mDlg_ColorSet()
    RaiseEvent ColorSet
End Sub

Private Sub mDlg_GetLocalizedText(ByVal LanguageID As Long, ByVal SubLanguageID As Long, ByVal TextID As Long, Text As String)
    RaiseEvent GetLocalizedText(LanguageID, SubLanguageID, TextID, Text)
End Sub

Private Sub mDlg_Hided()
    RaiseEvent Hided
End Sub

Private Sub UserControl_Initialize()
    Set mDlg = New ColorDialog
End Sub

Private Sub UserControl_InitProperties()
    mDlg.BackColor = cPropDefault_ColorDialog_BackColor
    mDlg.BasicColorsVisible = cPropDefault_ColorDialog_BasicColorsVisible
    mDlg.Color = cPropDefault_ColorDialog_Color
    mDlg.ColorSelectionBoxVisible = cPropDefault_ColorDialog_ColorSelectionBoxVisible
    mDlg.ColorSystem = cPropDefault_ColorDialog_ColorSystem
    mDlg.ColorSystemControlVisible = cPropDefault_ColorDialog_ColorSystemControlVisible
    mDlg.ColorValuesSectionVisible = cPropDefault_ColorDialog_ColorValuesSectionVisible
    mDlg.ConfirmationButtonsVisible = cPropDefault_ColorDialog_ConfirmationButtonsVisible
    mDlg.DialogCaptionVisible = cPropDefault_ColorDialog_DialogCaptionVisible
    mDlg.EyeDropperVisible = cPropDefault_ColorDialog_EyeDropperVisible
    mDlg.FixedPalette = cPropDefault_ColorDialog_FixedPalette
    mDlg.HexControlVisible = cPropDefault_ColorDialog_HexControlVisible
    mDlg.HexFormatVB = cPropDefault_ColorDialog_HexFormatVB
    mDlg.HideLabels = cPropDefault_ColorDialog_HideLabels
    mDlg.PointerType = cPropDefault_ColorDialog_PointerType
    mDlg.Modeless = cPropDefault_ColorDialog_Modeless
    mDlg.RecentColorsColumns = cPropDefault_ColorDialog_RecentColorsColumns
    mDlg.RememberPosition = cPropDefault_ColorDialog_RememberPosition
    mDlg.RoundedBoxes = cPropDefault_ColorDialog_RoundedBoxes
    mDlg.PaletteTypeControlVisible = cPropDefault_ColorDialog_PaletteTypeControlVisible
    mDlg.SliderParameter = cPropDefault_ColorDialog_SliderParameter
    mDlg.SliderOptionsAvailable = cPropDefault_ColorDialog_SliderOptionsAvailable
    mDlg.SizeBig = cPropDefault_ColorDialog_SizeBig
    mDlg.SliderWide = cPropDefault_ColorDialog_SliderWide
    mDlg.Style = cPropDefault_ColorDialog_Style
    If Ambient.UserMode Then
        On Error Resume Next
        Set mDlg.ActiveForm = UserControl.Parent
        On Error GoTo 0
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mDlg.BackColor = PropBag.ReadProperty("BackColor", cPropDefault_ColorDialog_BackColor)
    mDlg.BasicColorsVisible = PropBag.ReadProperty("BasicColorsVisible", cPropDefault_ColorDialog_BasicColorsVisible)
    mDlg.Color = PropBag.ReadProperty("Color", cPropDefault_ColorDialog_Color)
    mDlg.ColorSelectionBoxVisible = PropBag.ReadProperty("ColorSelectionBoxVisible", cPropDefault_ColorDialog_ColorSelectionBoxVisible)
    mDlg.ColorSystem = PropBag.ReadProperty("ColorSystem", cPropDefault_ColorDialog_ColorSystem)
    mDlg.ColorSystemControlVisible = PropBag.ReadProperty("ColorSystemControlVisible", cPropDefault_ColorDialog_ColorSystemControlVisible)
    mDlg.ColorValuesSectionVisible = PropBag.ReadProperty("ColorValuesSectionVisible", cPropDefault_ColorDialog_ColorValuesSectionVisible)
    mDlg.ConfirmationButtonsVisible = PropBag.ReadProperty("ConfirmationButtonsVisible", cPropDefault_ColorDialog_ConfirmationButtonsVisible)
    mDlg.Context = PropBag.ReadProperty("Context", "")
    mDlg.DialogCaption = PropBag.ReadProperty("DialogCaption", "")
    mDlg.DialogCaptionVisible = PropBag.ReadProperty("DialogCaptionVisible", cPropDefault_ColorDialog_DialogCaptionVisible)
    mDlg.EyeDropperVisible = PropBag.ReadProperty("EyeDropperVisible", cPropDefault_ColorDialog_EyeDropperVisible)
    mDlg.FixedPalette = PropBag.ReadProperty("FixedPalette", cPropDefault_ColorDialog_FixedPalette)
    mDlg.HexControlVisible = PropBag.ReadProperty("HexControlVisible", cPropDefault_ColorDialog_HexControlVisible)
    mDlg.HexFormatVB = PropBag.ReadProperty("HexFormatVB", cPropDefault_ColorDialog_HexFormatVB)
    mDlg.HideLabels = PropBag.ReadProperty("HideLabels", cPropDefault_ColorDialog_HideLabels)
    mDlg.PointerType = PropBag.ReadProperty("PointerType", cPropDefault_ColorDialog_PointerType)
    mDlg.Modeless = PropBag.ReadProperty("Modeless", cPropDefault_ColorDialog_Modeless)
    mDlg.RecentColorsColumns = PropBag.ReadProperty("RecentColorsColumns", cPropDefault_ColorDialog_RecentColorsColumns)
    mDlg.RememberPosition = PropBag.ReadProperty("RememberPosition", cPropDefault_ColorDialog_RememberPosition)
    mDlg.RoundedBoxes = PropBag.ReadProperty("RoundedBoxes", cPropDefault_ColorDialog_RoundedBoxes)
    mDlg.PaletteTypeControlVisible = PropBag.ReadProperty("PaletteTypeControlVisible", cPropDefault_ColorDialog_PaletteTypeControlVisible)
    mDlg.SliderParameter = PropBag.ReadProperty("SliderParameter", cPropDefault_ColorDialog_SliderParameter)
    mDlg.SliderOptionsAvailable = PropBag.ReadProperty("SliderOptionsAvailable", cPropDefault_ColorDialog_SliderOptionsAvailable)
    mDlg.SizeBig = PropBag.ReadProperty("SizeBig", cPropDefault_ColorDialog_SizeBig)
    mDlg.SliderWide = PropBag.ReadProperty("SliderWide", cPropDefault_ColorDialog_SliderWide)
    mDlg.Style = PropBag.ReadProperty("Style", cPropDefault_ColorDialog_Style)
    mDlg.PositionLeft = PropBag.ReadProperty("PositionLeft", 0)
    mDlg.PositionTop = PropBag.ReadProperty("PositionTop", 0)
    If Ambient.UserMode Then
        On Error Resume Next
        Set mDlg.ActiveForm = UserControl.Parent
        On Error GoTo 0
    End If
End Sub

Private Sub UserControl_Resize()
    Dim iH As Long
    Dim iW As Long
    Const cSize As Long = 45
    
    iH = UserControl.ScaleY(UserControl.ScaleHeight, vbTwips, vbPixels)
    iW = UserControl.ScaleX(UserControl.ScaleWidth, vbTwips, vbPixels)
    
    If (iH <> cSize) Or (iW <> cSize) Then
        If (iH <> cSize) Then
            iH = cSize
        End If
        If (iW <> cSize) Then
            iW = cSize
        End If
        UserControl.Size UserControl.ScaleX(iW, vbPixels, vbTwips), UserControl.ScaleY(iH, vbPixels, vbTwips)
    End If
End Sub

Private Sub UserControl_Terminate()
    Set mDlg = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BackColor", mDlg.BackColor, cPropDefault_ColorDialog_BackColor
    PropBag.WriteProperty "BasicColorsVisible", mDlg.BasicColorsVisible, cPropDefault_ColorDialog_BasicColorsVisible
    PropBag.WriteProperty "Color", mDlg.Color, cPropDefault_ColorDialog_Color
    PropBag.WriteProperty "ColorSelectionBoxVisible", mDlg.ColorSelectionBoxVisible, cPropDefault_ColorDialog_ColorSelectionBoxVisible
    PropBag.WriteProperty "ColorSystem", mDlg.ColorSystem, cPropDefault_ColorDialog_ColorSystem
    PropBag.WriteProperty "ColorSystemControlVisible", mDlg.ColorSystemControlVisible, cPropDefault_ColorDialog_ColorSystemControlVisible
    PropBag.WriteProperty "ColorValuesSectionVisible", mDlg.ColorValuesSectionVisible, cPropDefault_ColorDialog_ColorValuesSectionVisible
    PropBag.WriteProperty "ConfirmationButtonsVisible", mDlg.ConfirmationButtonsVisible, cPropDefault_ColorDialog_ConfirmationButtonsVisible
    PropBag.WriteProperty "Context", mDlg.Context, ""
    PropBag.WriteProperty "DialogCaption", mDlg.DialogCaption, ""
    PropBag.WriteProperty "DialogCaptionVisible", mDlg.DialogCaptionVisible, cPropDefault_ColorDialog_DialogCaptionVisible
    PropBag.WriteProperty "EyeDropperVisible", mDlg.EyeDropperVisible, cPropDefault_ColorDialog_EyeDropperVisible
    PropBag.WriteProperty "FixedPalette", mDlg.FixedPalette, cPropDefault_ColorDialog_FixedPalette
    PropBag.WriteProperty "HexControlVisible", mDlg.HexControlVisible, cPropDefault_ColorDialog_HexControlVisible
    PropBag.WriteProperty "HexFormatVB", mDlg.HexFormatVB, cPropDefault_ColorDialog_HexFormatVB
    PropBag.WriteProperty "HideLabels", mDlg.HideLabels, cPropDefault_ColorDialog_HideLabels
    PropBag.WriteProperty "PointerType", mDlg.PointerType, cPropDefault_ColorDialog_PointerType
    PropBag.WriteProperty "Modeless", mDlg.Modeless, cPropDefault_ColorDialog_Modeless
    PropBag.WriteProperty "RecentColorsColumns", mDlg.RecentColorsColumns, cPropDefault_ColorDialog_RecentColorsColumns
    PropBag.WriteProperty "RememberPosition", mDlg.RememberPosition, cPropDefault_ColorDialog_RememberPosition
    PropBag.WriteProperty "RoundedBoxes", mDlg.RoundedBoxes, cPropDefault_ColorDialog_RoundedBoxes
    PropBag.WriteProperty "PaletteTypeControlVisible", mDlg.PaletteTypeControlVisible, cPropDefault_ColorDialog_PaletteTypeControlVisible
    PropBag.WriteProperty "SliderParameter", mDlg.SliderParameter, cPropDefault_ColorDialog_SliderParameter
    PropBag.WriteProperty "SliderOptionsAvailable", mDlg.SliderOptionsAvailable, cPropDefault_ColorDialog_SliderOptionsAvailable
    PropBag.WriteProperty "SizeBig", mDlg.SizeBig, cPropDefault_ColorDialog_SizeBig
    PropBag.WriteProperty "SliderWide", mDlg.SliderWide, cPropDefault_ColorDialog_SliderWide
    PropBag.WriteProperty "Style", mDlg.Style, cPropDefault_ColorDialog_Style
    PropBag.WriteProperty "PositionLeft", mDlg.PositionLeft, 0
    PropBag.WriteProperty "PositionTop", mDlg.PositionTop, 0
End Sub


Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color."
    BackColor = mDlg.BackColor
End Property

Public Property Let BackColor(ByVal nValue As OLE_COLOR)
    If nValue <> mDlg.BackColor Then
        If Not IsValidOLE_COLOR(nValue) Then Err.Raise 380: Exit Property
        mDlg.BackColor = nValue
        PropertyChanged "BackColor"
    End If
End Property


Public Property Get BasicColorsVisible() As Boolean
Attribute BasicColorsVisible.VB_Description = "Determines whether the ""Basic colors"" section is visible or not."
    BasicColorsVisible = mDlg.BasicColorsVisible
End Property

Public Property Let BasicColorsVisible(ByVal nValue As Boolean)
    If nValue <> mDlg.BasicColorsVisible Then
        mDlg.BasicColorsVisible = nValue
        PropertyChanged "BasicColorsVisible"
    End If
End Property


Public Property Get DialogCaption() As String
Attribute DialogCaption.VB_Description = "Returns/sets the caption of the dialog's title bar."
    DialogCaption = mDlg.DialogCaption
End Property

Public Property Let DialogCaption(nValue As String)
    If nValue <> mDlg.DialogCaption Then
        mDlg.DialogCaption = nValue
        PropertyChanged "DialogCaption"
    End If
End Property


Public Property Get DialogCaptionVisible() As Boolean
Attribute DialogCaptionVisible.VB_Description = "Returns/sets a value that determines if the dialog's title bar is visible. When True, the default BackColor is changed to white."
    DialogCaptionVisible = mDlg.DialogCaptionVisible
End Property

Public Property Let DialogCaptionVisible(ByVal nValue As Boolean)
    If nValue <> mDlg.DialogCaptionVisible Then
        mDlg.DialogCaptionVisible = nValue
        PropertyChanged "DialogCaptionVisible"
    End If
End Property


Public Property Get EyeDropperVisible() As Boolean
Attribute EyeDropperVisible.VB_Description = "Returns/sets a value that determines if the 'eye dropper' feature used to pick a color from the entire screen is available."
    EyeDropperVisible = mDlg.EyeDropperVisible
End Property

Public Property Let EyeDropperVisible(ByVal nValue As Boolean)
    If nValue <> mDlg.EyeDropperVisible Then
        mDlg.EyeDropperVisible = nValue
        PropertyChanged "EyeDropperVisible"
    End If
End Property


Public Property Get Color() As OLE_COLOR
Attribute Color.VB_Description = "Returns/sets the current color value."
Attribute Color.VB_UserMemId = 0
    Color = mDlg.Color
End Property

Public Property Let Color(ByVal nValue As OLE_COLOR)
    If nValue <> mDlg.Color Then
        mDlg.Color = nValue
        PropertyChanged "Color"
    End If
End Property


Public Property Get ColorSelectionBoxVisible() As Boolean
Attribute ColorSelectionBoxVisible.VB_Description = "Returns/sets a value that determines if the box showing the selected color is visible."
    ColorSelectionBoxVisible = mDlg.ColorSelectionBoxVisible
End Property

Public Property Let ColorSelectionBoxVisible(ByVal nValue As Boolean)
    If nValue <> mDlg.ColorSelectionBoxVisible Then
        mDlg.ColorSelectionBoxVisible = nValue
        PropertyChanged "ColorSelectionBoxVisible"
    End If
End Property


Public Property Get ColorSystem() As CDColorSystemConstants
Attribute ColorSystem.VB_Description = "Returns/sets a value that determines the color system: HSV or HSL."
    ColorSystem = mDlg.ColorSystem
End Property

Public Property Let ColorSystem(ByVal nValue As CDColorSystemConstants)
    If nValue <> mDlg.ColorSystem Then
        mDlg.ColorSystem = nValue
        PropertyChanged "ColorSystem"
    End If
End Property


Public Property Get ColorSystemControlVisible() As Boolean
Attribute ColorSystemControlVisible.VB_Description = "Returns/sets a value that determines if the user is able to change the color system."
    ColorSystemControlVisible = mDlg.ColorSystemControlVisible
End Property

Public Property Let ColorSystemControlVisible(ByVal nValue As Boolean)
    If nValue <> mDlg.ColorSystemControlVisible Then
        mDlg.ColorSystemControlVisible = nValue
        PropertyChanged "ColorSystemControlVisible"
    End If
End Property


Public Property Get ColorValuesSectionVisible() As Boolean
Attribute ColorValuesSectionVisible.VB_Description = "Returns/sets a value that determines if the boxes to enter the color values by hand are visible or not."
    ColorValuesSectionVisible = mDlg.ColorValuesSectionVisible
End Property

Public Property Let ColorValuesSectionVisible(ByVal nValue As Boolean)
    If nValue <> mDlg.ColorValuesSectionVisible Then
        mDlg.ColorValuesSectionVisible = nValue
        PropertyChanged "ColorValuesSectionVisible"
    End If
End Property


Public Property Get ConfirmationButtonsVisible() As Boolean
Attribute ConfirmationButtonsVisible.VB_Description = "Returns/sets a value that determines if the OK and Cancel buttons are visible."
    ConfirmationButtonsVisible = mDlg.ConfirmationButtonsVisible
End Property

Public Property Let ConfirmationButtonsVisible(ByVal nValue As Boolean)
    If nValue <> mDlg.ConfirmationButtonsVisible Then
        mDlg.ConfirmationButtonsVisible = nValue
        PropertyChanged "ConfirmationButtonsVisible"
    End If
End Property


Public Property Get Context() As String
Attribute Context.VB_Description = "Returns/sets a context for saving the user settings and states."
    Context = mDlg.Context
End Property

Public Property Let Context(nValue As String)
    If nValue <> mDlg.Context Then
        mDlg.Context = nValue
        PropertyChanged "Context"
    End If
End Property


Public Property Get FixedPalette() As Boolean
Attribute FixedPalette.VB_Description = "Returns/sets a value that determines if the color palette (or the slidder in some configurations) colors change with the setting of the third partameter."
    FixedPalette = mDlg.FixedPalette
End Property

Public Property Let FixedPalette(ByVal nValue As Boolean)
    If nValue <> mDlg.FixedPalette Then
        mDlg.FixedPalette = nValue
        PropertyChanged "FixedPalette"
    End If
End Property


Public Property Get HexControlVisible() As Boolean
Attribute HexControlVisible.VB_Description = "Returns/sets a value that determines if the hexadecimal text box is visible."
    HexControlVisible = mDlg.HexControlVisible
End Property

Public Property Let HexControlVisible(ByVal nValue As Boolean)
    If nValue <> mDlg.HexControlVisible Then
        mDlg.HexControlVisible = nValue
        PropertyChanged "HexControlVisible"
    End If
End Property


Public Property Get HexFormatVB() As Boolean
Attribute HexFormatVB.VB_Description = "Returns/sets a value that determines the format of the hexadecimal text box, for web or VB."
    HexFormatVB = mDlg.HexFormatVB
End Property

Public Property Let HexFormatVB(ByVal nValue As Boolean)
    If nValue <> mDlg.HexFormatVB Then
        mDlg.HexFormatVB = nValue
        PropertyChanged "HexFormatVB"
    End If
End Property


Public Property Get HideLabels() As Boolean
Attribute HideLabels.VB_Description = "Returns/sets a value that determines if the labels are visible."
    HideLabels = mDlg.HideLabels
End Property

Public Property Let HideLabels(ByVal nValue As Boolean)
    If nValue <> mDlg.HideLabels Then
        mDlg.HideLabels = nValue
        PropertyChanged "HideLabels"
    End If
End Property


Public Property Get PointerType() As CDPointerTypeConstants
Attribute PointerType.VB_Description = "Returns/sets the pointer type that will indicate the selected color in the palette."
    PointerType = mDlg.PointerType
End Property

Public Property Let PointerType(ByVal nValue As CDPointerTypeConstants)
    If nValue <> mDlg.PointerType Then
        mDlg.PointerType = nValue
        PropertyChanged "PointerType"
    End If
End Property


Public Property Get Modeless() As Boolean
Attribute Modeless.VB_Description = "Returns/sets a value that determines if the dialog is displayed modally or not."
    Modeless = mDlg.Modeless
End Property

Public Property Let Modeless(ByVal nValue As Boolean)
    If nValue <> mDlg.Modeless Then
        mDlg.Modeless = nValue
        PropertyChanged "Modeless"
    End If
End Property


Public Property Get RecentColorsColumns() As Long
Attribute RecentColorsColumns.VB_Description = "Returns/sets a value that determines the number of columns of the 'Recent colors'. Set to 0 for none."
    RecentColorsColumns = mDlg.RecentColorsColumns
End Property

Public Property Let RecentColorsColumns(ByVal nValue As Long)
    If nValue <> mDlg.RecentColorsColumns Then
        mDlg.RecentColorsColumns = nValue
        PropertyChanged "RecentColorsColumns"
    End If
End Property


Public Property Get RememberPosition() As Boolean
Attribute RememberPosition.VB_Description = "Returns/sets a value that determines whether the position of the dialog will be remembered next time."
    RememberPosition = mDlg.RememberPosition
End Property

Public Property Let RememberPosition(ByVal nValue As Boolean)
    If nValue <> mDlg.RememberPosition Then
        mDlg.RememberPosition = nValue
        PropertyChanged "RememberPosition"
    End If
End Property


Public Property Get RoundedBoxes() As Boolean
Attribute RoundedBoxes.VB_Description = "Returns/sets a value that determines if some control borders are rounded."
    RoundedBoxes = mDlg.RoundedBoxes
End Property

Public Property Let RoundedBoxes(ByVal nValue As Boolean)
    If nValue <> mDlg.RoundedBoxes Then
        mDlg.RoundedBoxes = nValue
        PropertyChanged "RoundedBoxes"
    End If
End Property


Public Property Get PaletteTypeControlVisible() As Boolean
Attribute PaletteTypeControlVisible.VB_Description = "Returns/sets a value that determines if the control for changing the palette type is visible or not."
    PaletteTypeControlVisible = mDlg.PaletteTypeControlVisible
End Property

Public Property Let PaletteTypeControlVisible(ByVal nValue As Boolean)
    If nValue <> mDlg.PaletteTypeControlVisible Then
        mDlg.PaletteTypeControlVisible = nValue
        PropertyChanged "PaletteTypeControlVisible"
    End If
End Property


Public Property Get SliderParameter() As CDSliderParameterConstants
Attribute SliderParameter.VB_Description = "Returns/sets a value that determines the slidder parameter."
    SliderParameter = mDlg.SliderParameter
End Property

Public Property Let SliderParameter(ByVal nValue As CDSliderParameterConstants)
    If nValue <> mDlg.SliderParameter Then
        mDlg.SliderParameter = nValue
        PropertyChanged "SliderParameter"
    End If
End Property


Public Property Get SliderOptionsAvailable() As CDSliderOptionsAvailableConstants
Attribute SliderOptionsAvailable.VB_Description = "Returns/sets a value that determines which parameters for the slider the user will have available to choose from."
    SliderOptionsAvailable = mDlg.SliderOptionsAvailable
End Property

Public Property Let SliderOptionsAvailable(ByVal nValue As CDSliderOptionsAvailableConstants)
    If nValue <> mDlg.SliderOptionsAvailable Then
        mDlg.SliderOptionsAvailable = nValue
        PropertyChanged "SliderOptionsAvailable"
    End If
End Property


Public Property Get SizeBig() As Boolean
Attribute SizeBig.VB_Description = "Returns/sets a value that determines the size of the dialog, big or normal."
    SizeBig = mDlg.SizeBig
End Property

Public Property Let SizeBig(ByVal nValue As Boolean)
    If nValue <> mDlg.SizeBig Then
        mDlg.SizeBig = nValue
        PropertyChanged "SizeBig"
    End If
End Property


Public Property Get SliderWide() As CDYesNoAutoConstants
Attribute SliderWide.VB_Description = "Returns/sets a value that determines the width of the slider control between wide or narrow."
    SliderWide = mDlg.SliderWide
End Property

Public Property Let SliderWide(ByVal nValue As CDYesNoAutoConstants)
    If nValue <> mDlg.SliderWide Then
        mDlg.SliderWide = nValue
        PropertyChanged "SliderWide"
    End If
End Property


Public Property Get Style() As CDStyleConstants
Attribute Style.VB_Description = "Returns/sets a value that determines the color palette style, between wheel or box."
Attribute Style.VB_MemberFlags = "200"
    Style = mDlg.Style
End Property

Public Property Let Style(ByVal nValue As CDStyleConstants)
    If nValue <> mDlg.Style Then
        mDlg.Style = nValue
        PropertyChanged "Style"
    End If
End Property


Public Property Get PositionLeft() As Single
Attribute PositionLeft.VB_Description = "Returns/sets a value that determines the Left position of the dialog."
    PositionLeft = mDlg.PositionLeft
End Property

Public Property Let PositionLeft(ByVal nValue As Single)
    If nValue <> mDlg.PositionLeft Then
        mDlg.PositionLeft = nValue
        PropertyChanged "PositionLeft"
    End If
End Property


Public Property Get PositionTop() As Single
Attribute PositionTop.VB_Description = "Returns/sets a value that determines the Top position of the dialog."
    PositionTop = mDlg.PositionTop
End Property

Public Property Let PositionTop(ByVal nValue As Single)
    If nValue <> mDlg.PositionTop Then
        mDlg.PositionTop = nValue
        PropertyChanged "PositionTop"
    End If
End Property


Public Property Get Canceled() As Boolean
Attribute Canceled.VB_Description = "Returns True if the dialog was canceled."
    Canceled = mDlg.Canceled
End Property


Public Property Get Changed() As Boolean
Attribute Changed.VB_Description = "Returns true if the color was changed."
    Changed = mDlg.Changed
End Property


Public Function Show(Optional ByVal nStyleBox As Variant) As Boolean
Attribute Show.VB_Description = "Shows the dialog."
    Show = mDlg.Show(nStyleBox)
End Function

Public Sub Hide()
Attribute Hide.VB_Description = "Hides the dialog when it is displayed modeless."
    mDlg.Hide
End Sub

Public Sub Move(ByVal nPositionLeft As Single, ByVal nPositionTop As Single)
Attribute Move.VB_Description = "Sets the initial position of the dialog."
    mDlg.Move nPositionLeft, nPositionTop
End Sub

Public Sub SetCompact(Optional RecentColorsColumns As Long = 0, Optional ColorValuesSectionVisible As Boolean = False, Optional SliderOptionsAvailable As CDSliderOptionsAvailableConstants = cdSliderOptionsNone, Optional nDialogCaptionVisible As Boolean = True, Optional nConfirmationButtonsVisible As Boolean = True, Optional nColorSelectionBoxVisible As Boolean = True, Optional nSliderWide As CDYesNoAutoConstants = cdYNAuto, Optional nHideLabels As Boolean = False)
Attribute SetCompact.VB_Description = "Helper method to set several properties to get a 'Compact' configuration in one shot."
    mDlg.SetCompact RecentColorsColumns, ColorValuesSectionVisible, SliderOptionsAvailable, nDialogCaptionVisible, nConfirmationButtonsVisible, nColorSelectionBoxVisible, nSliderWide, nHideLabels
End Sub

Public Sub SetComplete(Optional SizeBig As Boolean = False, Optional BasicColorsVisible As Boolean = True, Optional nRecentColorsColumns As Long = 2, Optional SliderOptionsAvailable As CDSliderOptionsAvailableConstants = cdSliderOptionsAll, Optional PaletteTypeControlVisible As Boolean = True, Optional ColorSystemControlVisible As Boolean = True, Optional EyeDropperVisible As Boolean = True)
Attribute SetComplete.VB_Description = "Helper method to set several properties to get a 'Complete' configuration in one shot."
    mDlg.SetComplete SizeBig, BasicColorsVisible, nRecentColorsColumns, SliderOptionsAvailable, PaletteTypeControlVisible, ColorSystemControlVisible
End Sub

Public Sub SetSimple(Optional RecentColorsColumns As Long = 2, Optional nDialogCaptionVisible As Boolean = False, Optional nConfirmationButtonsVisible As Boolean, Optional nColorSelectionBoxVisible As Boolean, Optional nSliderWide As CDYesNoAutoConstants = cdYNAuto, Optional nHideLabels As Boolean = False)
Attribute SetSimple.VB_Description = "Helper method to set several properties to get a 'Simple' configuration in one shot."
    mDlg.SetSimple RecentColorsColumns, nDialogCaptionVisible, nConfirmationButtonsVisible, nColorSelectionBoxVisible, nSliderWide, nHideLabels
End Sub

