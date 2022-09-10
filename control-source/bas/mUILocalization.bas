Attribute VB_Name = "mUILocalization"
Option Explicit

Private Type TLOCALESIGNATURE
    lsUsb(0 To 15) As Byte
    lsCsbDefault(0 To 1) As Long
    lsCsbSupported(0 To 1) As Long
End Type

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetLocaleInfoW Lib "kernel32" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long
'Private Declare Function GetThreadLocale Lib "kernel32" () As Long
Private Declare Function GetUserDefaultUILanguage Lib "kernel32.dll" () As Long
Private Declare Function LocaleNameToLCID Lib "kernel32" (ByVal lpName As Long, ByVal dwFlags As Long) As Long

'Private Const SUBLANG_ENGLISH_US As Long = &H1
Private Const SUBLANG_ENGLISH_UK  As Long = &H2
Private Const SUBLANG_PORTUGUESE_BRAZILIAN As Long = &H1

Public Enum BSLanguageIDConstants
    bsLang_AUTO_SELECT = 0
    bsLang_CHINESE_SIMPLIFIED = 4
    bsLang_GERMAN = 7
    bsLang_GREEK = 8
    bsLang_ENGLISH = 9
    bsLang_SPANISH = 10
    bsLang_FRENCH = 12
    bsLang_ITALIAN = 16
    
' Full list:

'    bsLang_AUTO_SELECT = 0
'    bsLang_ARABIC = 1
'    bsLang_BULGARIAN = 2
'    bsLang_CATALAN = 3
'    bsLang_CHINESE_SIMPLIFIED = 4
'    bsLang_CZECH = 5
'    bsLang_DANISH = 6
'    bsLang_GERMAN = 7
'    bsLang_GREEK = 8
'    bsLang_ENGLISH = 9
'    bsLang_SPANISH = 10
'    bsLang_FINNISH = 11
'    bsLang_FRENCH = 12
'    bsLang_HEBREW = 13
'    bsLang_HUNGARIAN = 14
'    bsLang_ICELANDIC = 15
'    bsLang_ITALIAN = 16
'    bsLang_JAPANESE = 17
'    bsLang_KOREAN = 18
'    bsLang_DUTCH = 19
'    bsLang_NORWEGIAN = 20
'    bsLang_POLISH = 21
'    bsLang_PORTUGUESE = 22
'    bsLang_ROMANSH = 23
'    bsLang_ROMANIAN = 24
'    bsLang_RUSSIAN = 25
'    bsLang_BOSNIAN_SERBIAN_CROATIAN = 26
'    bsLang_SLOVAK = 27
'    bsLang_ALBANIAN = 28
'    bsLang_SWEDISH = 29
'    bsLang_THAI = 30
'    bsLang_TURKISH = 31
'    bsLang_URDU = 32
'    bsLang_INDONESIAN = 33
'    bsLang_UKRAINIAN = 34
'    bsLang_BELARUSIAN = 35
'    bsLang_SLOVENIAN = 36
'    bsLang_ESTONIAN = 37
'    bsLang_LATVIAN = 38
'    bsLang_LITHUANIAN = 39
'    bsLang_TAJIK = 40
'    bsLang_FARSI_PERSIAN = 41
'    bsLang_VIETNAMESE = 42
'    bsLang_ARMENIAN = 43
'    bsLang_AZERI = 44
'    bsLang_BASQUE = 45
'    bsLang_SORBIAN = 46
'    bsLang_MACEDONIAN = 47
'    bsLang_TSWANA = 50
'    bsLang_XHOSA = 52
'    bsLang_ZULU = 53
'    bsLang_AFRIKAANS = 54
'    bsLang_GEORGIAN = 55
'    bsLang_FAEROESE = 56
'    bsLang_HINDI = 57
'    bsLang_MALTESE = 58
'    bsLang_SAMI = 59
'    bsLang_IRISH = 60
'    bsLang_MALAY = 62
'    bsLang_KAZAK = 63
'    bsLang_KYRGYZ = 64
'    bsLang_SWAHILI = 65
'    bsLang_TURKMEN = 66
'    bsLang_UZBEK = 67
'    bsLang_TATAR = 68
'    bsLang_BENGALI = 69
'    bsLang_PUNJABI = 70
'    bsLang_GUJARATI = 71
'    bsLang_ORIYA = 72
'    bsLang_TAMIL = 73
'    bsLang_TELUGU = 74
'    bsLang_KANNADA = 75
'    bsLang_MALAYALAM = 76
'    bsLang_ASSAMESE = 77
'    bsLang_MARATHI = 78
'    bsLang_SANSKRIT = 79
'    bsLang_MONGOLIAN = 80
'    bsLang_TIBETAN = 81
'    bsLang_WELSH = 82
'    bsLang_KHMER = 83
'    bsLang_LAO = 84
'    bsLang_GALICIAN = 86
'    bsLang_KONKANI = 87
'    bsLang_MANIPURI = 88
'    bsLang_SINDHI = 89
'    bsLang_SYRIAC = 90
'    bsLang_SINHALESE = 91
'    bsLang_INUKTITUT = 93
'    bsLang_AMHARIC = 94
'    bsLang_TAMAZIGHT = 95
'    bsLang_KASHMIRI = 96
'    bsLang_NEPALI = 97
'    bsLang_FRISIAN = 98
'    bsLang_PASHTO = 99
'    bsLang_FILIPINO = 100
'    bsLang_DIVEHI = 101
'    bsLang_HAUSA = 104
'    bsLang_YORUBA = 106
'    bsLang_QUECHUA = 107
'    bsLang_SOTHO = 108
'    bsLang_BASHKIR = 109
'    bsLang_LUXEMBOURGISH = 110
'    bsLang_GREENLANDIC = 111
'    bsLang_IGBO = 112
'    bsLang_TIGRIGNA = 115
'    bsLang_YI = 120
'    bsLang_MAPUDUNGUN = 122
'    bsLang_MOHAWK = 124
'    bsLang_BRETON = 126
'    bsLang_UIGHUR = 128
'    bsLang_MAORI = 129
'    bsLang_OCCITAN = 130
'    bsLang_CORSICAN = 131
'    bsLang_ALSATIAN = 132
'    bsLang_YAKUT = 133
'    bsLang_KICHE = 134
'    bsLang_KINYARWANDA = 135
'    bsLang_WOLOF = 136
'    bsLang_DARI = 140
'    bsLang_SCOTTISH_GAELIC = 145
'    bsLang_BOSNIAN_NEUTRAL = 30746
'    bsLang_CHINESE_TRADITIONAL = 31748
'    bsLang_SERBIAN_NEUTRAL = 31770

End Enum

Public Enum CDUserInterfaceTextIDConstants
    
    
    ' ColorSelector
    cdUIT_ColorSelector_chkFixedPalette_Caption = 2000
    cdUIT_ColorSelector_chkFixedPalette_ToolTipText = 2100
    cdUIT_ColorSelector_cboSliderParameter_ToolTipText = 2200
    cdUIT_ColorSelector_lblMode_Caption = 2300
    cdUIT_ColorSelector_cboSliderParameter_ListItem_Hue = 2400
    cdUIT_ColorSelector_cboSliderParameter_ListItem_Luminance = 2500
    cdUIT_ColorSelector_cboSliderParameter_ListItem_Value = 2600
    cdUIT_ColorSelector_cboSliderParameter_ListItem_Saturation = 2700
    cdUIT_ColorSelector_cboSliderParameter_ListItem_Red = 2800
    cdUIT_ColorSelector_cboSliderParameter_ListItem_Green = 2900
    cdUIT_ColorSelector_cboSliderParameter_ListItem_Blue = 3000
    cdUIT_ColorSelector_cboColorSystem_ListItem_HSV = 3100
    cdUIT_ColorSelector_cboColorSystem_ListItem_HSL = 3200
    
    ' frmColorDialog
    cdUIT_frmColorDialog_Form_Caption = 4000
    cdUIT_frmColorDialog_lblBasicColors_Caption = 4100
    cdUIT_frmColorDialog_lblRecent_Caption = 4200
    cdUIT_frmColorDialog_Color_Caption = 4300
    cdUIT_frmColorDialog_ColorNew_Caption = 4400
    cdUIT_frmColorDialog_ColorPrevious_Caption = 4500
    cdUIT_frmColorDialog_MouseWheel_ToolTipStart = 4600
    cdUIT_frmColorDialog_MouseWheel_ToolTipEnd = 4700
    cdUIT_frmColorDialog_EyeDropper_ToolTip = 4800
    cdUIT_frmColorDialog_picEyeDropper_ToolTipText = 4900
    cdUIT_frmColorDialog_lblRed_Caption = 5000
    cdUIT_frmColorDialog_lblGreen_Caption = 5100
    cdUIT_frmColorDialog_lblBlue_Caption = 5200
    cdUIT_frmColorDialog_lblHex_Caption = 5300
    cdUIT_frmColorDialog_lblHue_Caption = 5400
    cdUIT_frmColorDialog_lblSaturation_Caption = 5500
    cdUIT_frmColorDialog_Luminance_Caption = 5600
    cdUIT_frmColorDialog_Value_Caption = 5700
    cdUIT_frmColorDialog_lblColorSystem_Caption = 5800
    cdUIT_frmColorDialog_lblPalette_Caption = 5900
    cdUIT_frmColorDialog_cboPalette_ListItem1 = 6000
    cdUIT_frmColorDialog_cboPalette_ListItem2 = 6100
    cdUIT_frmColorDialog_cboPalette_ListItem3 = 6200
    cdUIT_frmColorDialog_cboPalette_ListItem4 = 6300
    cdUIT_frmColorDialog_InvalidColorMessage = 6400
    cdUIT_frmColorDialog_ParameterFullName_Hue = 6500
    cdUIT_frmColorDialog_ParameterFullName_Luminance = 6501
    cdUIT_frmColorDialog_ParameterFullName_Saturation = 6502
    cdUIT_frmColorDialog_ParameterFullName_Red = 6503
    cdUIT_frmColorDialog_ParameterFullName_Green = 6504
    cdUIT_frmColorDialog_ParameterFullName_Blue = 6505
    cdUIT_frmColorDialog_ParameterFullName_Value = 6506
    cdUIT_frmColorDialog_OK = 6600
    cdUIT_frmColorDialog_Cancel = 6700
    cdUIT_frmColorDialog_Close = 6800
    
End Enum

Private mUILanguage As BSLanguageIDConstants
Private mUISubLanguage As Long
Private mLanguageWindowsUI As Long
Private mSubLanguageWindowsUI As Long
Private Const cDefaultLanguage = bsLang_ENGLISH
Private mUIRightToLeft As Boolean
Private mStringsCache As New Collection

Public Function GetLocalizedString(TextID As CDUserInterfaceTextIDConstants) As String
    If mUILanguage = bsLang_AUTO_SELECT Then SetUILanguageToWindowsUILanguage
    
    On Error Resume Next
    GetLocalizedString = mStringsCache(CStr(TextID))
    If Err.Number = 0 Then Exit Function
    On Error GoTo 0
    
    Select Case mUILanguage
        Case bsLang_CHINESE_SIMPLIFIED
            Do_CHINESE_SIMPLIFIED TextID, GetLocalizedString
        Case bsLang_GERMAN
            Do_GERMAN TextID, GetLocalizedString
        Case bsLang_SPANISH
            Do_SPANISH TextID, GetLocalizedString
        Case bsLang_FRENCH
            Do_FRENCH TextID, GetLocalizedString
        Case bsLang_ITALIAN
            Do_ITALIAN TextID, GetLocalizedString
        Case bsLang_GREEK
            Do_GREEK TextID, GetLocalizedString
        Case Else ' ENGLISH
            Do_ENGLISH TextID, GetLocalizedString
    End Select
    mStringsCache.Add GetLocalizedString, CStr(TextID)
End Function

Private Sub Do_ENGLISH(ByRef TextID As Long, ByRef Text As String)
    ' Microsoft teminology search: https://www.microsoft.com/en-us/language/search
    Select Case TextID
        Case cdUIT_frmColorDialog_Form_Caption ' English: "Color selection"
            If mUISubLanguage = SUBLANG_ENGLISH_UK Then
                Text = "Colour selection"
            Else
                Text = "Color selection"
            End If
        Case cdUIT_ColorSelector_chkFixedPalette_Caption ' English: "Fixed"
            Text = "Fixed"
        Case cdUIT_ColorSelector_chkFixedPalette_ToolTipText ' English: "Reflects color changes visually in the palette or not"
            If mUISubLanguage = SUBLANG_ENGLISH_UK Then
                Text = "Reflects colour changes visually in the palette or not"
            Else
                Text = "Reflects color changes visually in the palette or not"
            End If
        Case cdUIT_ColorSelector_cboSliderParameter_ToolTipText ' English: "Slider control parameter"
            Text = "Slider control parameter"
        Case cdUIT_ColorSelector_lblMode_Caption ' English: "Mode:"
            Text = "Mode:"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Hue ' English: "Hue"
            Text = "Hue"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Luminance ' English: "Lum."
            Text = "Lum."
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Value ' English: "Value"
            Text = "Value"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Saturation ' English: "Sat."
            Text = "Sat."
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Red ' English: "Red"
            Text = "Red"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Green ' English: "Green"
            Text = "Green"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Blue ' English: "Blue"
            Text = "Blue"
        Case cdUIT_ColorSelector_cboColorSystem_ListItem_HSV ' English: "HSV"
            Text = "HSV"
        Case cdUIT_ColorSelector_cboColorSystem_ListItem_HSL ' English: "HSL"
            Text = "HSL"
        Case cdUIT_frmColorDialog_lblBasicColors_Caption ' English: "Basic colors:"
            Text = "Basic colors:"
        Case cdUIT_frmColorDialog_lblRecent_Caption ' English: "Recent:"
            Text = "Recent:"
        Case cdUIT_frmColorDialog_Color_Caption ' English: "color"
            If mUISubLanguage = SUBLANG_ENGLISH_UK Then
                Text = "colour"
            Else
                Text = "color"
            End If
        Case cdUIT_frmColorDialog_ColorNew_Caption ' English: "new"
            Text = "new"
        Case cdUIT_frmColorDialog_ColorPrevious_Caption ' English: "previous"
            Text = "previous"
        Case cdUIT_frmColorDialog_MouseWheel_ToolTipStart ' English: "Hold down the Control key to change"
            Text = "Hold down the Control key to change"
        Case cdUIT_frmColorDialog_MouseWheel_ToolTipEnd ' English: "with the mouse wheel, press Shift key to go slowly"
            Text = "with the mouse wheel, press Shift key to go slowly"
        Case cdUIT_frmColorDialog_EyeDropper_ToolTip ' English: "Choose a color from anywhere on the screen or press the Escape key to cancel"
            If mUISubLanguage = SUBLANG_ENGLISH_UK Then
                Text = "Choose a colour from anywhere on the screen or press the Escape key to cancel"
            Else
                Text = "Choose a color from anywhere on the screen or press the Escape key to cancel"
            End If
        Case cdUIT_frmColorDialog_picEyeDropper_ToolTipText ' English: "Choose a color from the screen"
            If mUISubLanguage = SUBLANG_ENGLISH_UK Then
                Text = "Choose a colour from the screen"
            Else
                Text = "Choose a color from the screen"
            End If
        Case cdUIT_frmColorDialog_lblRed_Caption ' English: "Red:"
            Text = "Red:"
        Case cdUIT_frmColorDialog_lblGreen_Caption ' English: "Green:"
            Text = "Green:"
        Case cdUIT_frmColorDialog_lblBlue_Caption ' English: "Blue:"
            Text = "Blue:"
        Case cdUIT_frmColorDialog_lblHex_Caption ' English: "Hex:"
            Text = "Hex:"
        Case cdUIT_frmColorDialog_lblHue_Caption ' English: "Hue:"
            Text = "Hue:"
        Case cdUIT_frmColorDialog_lblSaturation_Caption ' English: "Sat.:"
            Text = "Sat.:"
        Case cdUIT_frmColorDialog_Luminance_Caption ' English: "Lum.:"
            Text = "Lum.:"
        Case cdUIT_frmColorDialog_Value_Caption ' English: "Value:"
            Text = "Value:"
        Case cdUIT_frmColorDialog_lblColorSystem_Caption ' English: "Mode:"
            Text = "Mode:"
        Case cdUIT_frmColorDialog_lblPalette_Caption ' English: "Palette:"
            Text = "Palette:"
        Case cdUIT_frmColorDialog_cboPalette_ListItem1 ' English: "Wheel, fixed"
            Text = "Wheel, fixed"
        Case cdUIT_frmColorDialog_cboPalette_ListItem2 ' English: "Box, fixed"
            Text = "Box, fixed"
        Case cdUIT_frmColorDialog_cboPalette_ListItem3 ' English: "Wheel, dynamic"
            Text = "Wheel, dynamic"
        Case cdUIT_frmColorDialog_cboPalette_ListItem4 ' English: "Box, dynamic"
            Text = "Box, dynamic"
        Case cdUIT_frmColorDialog_InvalidColorMessage ' English: "The color is not valid."
            If mUISubLanguage = SUBLANG_ENGLISH_UK Then
                Text = "The colour is not valid."
            Else
                Text = "The color is not valid."
            End If
        Case cdUIT_frmColorDialog_ParameterFullName_Hue ' English: "Hue"
            Text = "Hue"
        Case cdUIT_frmColorDialog_ParameterFullName_Luminance ' English: "Luminance"
            Text = "Luminance"
        Case cdUIT_frmColorDialog_ParameterFullName_Saturation ' English: "Saturation"
            Text = "Saturation"
        Case cdUIT_frmColorDialog_ParameterFullName_Red ' English: "Red"
            Text = "Red"
        Case cdUIT_frmColorDialog_ParameterFullName_Green ' English: "Green"
            Text = "Green"
        Case cdUIT_frmColorDialog_ParameterFullName_Blue ' English: "Blue"
            Text = "Blue"
        Case cdUIT_frmColorDialog_ParameterFullName_Value ' English: "Value"
            Text = "Value"
        Case cdUIT_frmColorDialog_OK ' English: "OK"
            Text = "OK"
        Case cdUIT_frmColorDialog_Cancel ' English: "Cancel"
            Text = "Cancel"
        Case cdUIT_frmColorDialog_Close ' English: "Close"
            Text = "Close"
    End Select
End Sub

Private Sub Do_SPANISH(ByRef TextID As Long, ByRef Text As String)
    ' Microsoft teminology search: https://www.microsoft.com/en-us/language/search
    Select Case TextID
        Case cdUIT_frmColorDialog_Form_Caption ' English: "Color selection"
            Text = "Selección de color"
        Case cdUIT_ColorSelector_chkFixedPalette_Caption ' English: "Fixed"
            Text = "Fija"
        Case cdUIT_ColorSelector_chkFixedPalette_ToolTipText ' English: "Reflects color changes visually in the palette or not"
            Text = "Refleja o no los cambios de colores en la paleta"
        Case cdUIT_ColorSelector_cboSliderParameter_ToolTipText ' English: "Slider control parameter"
            Text = "Parámetro del control deslizante"
        Case cdUIT_ColorSelector_lblMode_Caption ' English: "Mode:"
            Text = "Modo:"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Hue ' English: "Hue"
            Text = "Matiz"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Luminance ' English: "Lum."
            Text = "Lum."
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Value ' English: "Value"
            Text = "Valor"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Saturation ' English: "Sat."
            Text = "Sat."
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Red ' English: "Red"
            Text = "Rojo"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Green ' English: "Green"
            Text = "Verde"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Blue ' English: "Blue"
            Text = "Azul"
        Case cdUIT_ColorSelector_cboColorSystem_ListItem_HSV ' English: "HSV"
            Text = "HSV"
        Case cdUIT_ColorSelector_cboColorSystem_ListItem_HSL ' English: "HSL"
            Text = "HSL"
        Case cdUIT_frmColorDialog_lblBasicColors_Caption ' English: "Basic colors:"
            Text = "Colores básicos:"
        Case cdUIT_frmColorDialog_lblRecent_Caption ' English: "Recent:"
            Text = "Recientes:"
        Case cdUIT_frmColorDialog_Color_Caption ' English: "color"
            Text = "color"
        Case cdUIT_frmColorDialog_ColorNew_Caption ' English: "new"
            Text = "nuevo"
        Case cdUIT_frmColorDialog_ColorPrevious_Caption ' English: "previous"
            Text = "previo"
        Case cdUIT_frmColorDialog_MouseWheel_ToolTipStart ' English: "Hold down the Control key to change"
            Text = "Mantenga presionada la tecla Control para cambiar"
        Case cdUIT_frmColorDialog_MouseWheel_ToolTipEnd ' English: "with the mouse wheel, press Shift key to go slowly"
            Text = "con la rueda del mouse, presione la tecla Mayúsculas para ir lento"
        Case cdUIT_frmColorDialog_EyeDropper_ToolTip ' English: "Choose a color from anywhere on the screen or press the Escape key to cancel"
            Text = "Elija un color en cualquier lugar de la pantalla o presione la tecla Escape para cancelar"
        Case cdUIT_frmColorDialog_picEyeDropper_ToolTipText ' English: "Choose a color from the screen"
            Text = "Elija un color en la pantalla"
        Case cdUIT_frmColorDialog_lblRed_Caption ' English: "Red:"
            Text = "Rojo:"
        Case cdUIT_frmColorDialog_lblGreen_Caption ' English: "Green:"
            Text = "Verde:"
        Case cdUIT_frmColorDialog_lblBlue_Caption ' English: "Blue:"
            Text = "Azul:"
        Case cdUIT_frmColorDialog_lblHex_Caption ' English: "Hex:"
            Text = "Hex:"
        Case cdUIT_frmColorDialog_lblHue_Caption ' English: "Hue:"
            Text = "Matiz:"
        Case cdUIT_frmColorDialog_lblSaturation_Caption ' English: "Sat.:"
            Text = "Sat.:"
        Case cdUIT_frmColorDialog_Luminance_Caption ' English: "Lum.:"
            Text = "Lum.:"
        Case cdUIT_frmColorDialog_Value_Caption ' English: "Value:"
            Text = "Valor:"
        Case cdUIT_frmColorDialog_lblColorSystem_Caption ' English: "Mode:"
            Text = "Modo:"
        Case cdUIT_frmColorDialog_lblPalette_Caption ' English: "Palette:"
            Text = "Paleta:"
        Case cdUIT_frmColorDialog_cboPalette_ListItem1 ' English: "Wheel, fixed"
            Text = "Rueda, fija"
        Case cdUIT_frmColorDialog_cboPalette_ListItem2 ' English: "Box, fixed"
            Text = "Caja, fija"
        Case cdUIT_frmColorDialog_cboPalette_ListItem3 ' English: "Wheel, dynamic"
            Text = "Rueda, dinámica"
        Case cdUIT_frmColorDialog_cboPalette_ListItem4 ' English: "Box, dynamic"
            Text = "Caja, dinámica"
        Case cdUIT_frmColorDialog_InvalidColorMessage ' English: "The color is not valid."
            Text = "El color no es válido."
        Case cdUIT_frmColorDialog_ParameterFullName_Hue ' English: "Hue"
            Text = "Matiz"
        Case cdUIT_frmColorDialog_ParameterFullName_Luminance ' English: "Luminance"
            Text = "Luminancia"
        Case cdUIT_frmColorDialog_ParameterFullName_Saturation ' English: "Saturation"
            Text = "Saturación"
        Case cdUIT_frmColorDialog_ParameterFullName_Red ' English: "Red"
            Text = "Rojo"
        Case cdUIT_frmColorDialog_ParameterFullName_Green ' English: "Green"
            Text = "Verde"
        Case cdUIT_frmColorDialog_ParameterFullName_Blue ' English: "Blue"
            Text = "Azul"
        Case cdUIT_frmColorDialog_ParameterFullName_Value ' English: "Value"
            Text = "Valor"
        Case cdUIT_frmColorDialog_OK ' English: "OK"
            Text = "Aceptar"
        Case cdUIT_frmColorDialog_Cancel ' English: "Cancel"
            Text = "Cancelar"
        Case cdUIT_frmColorDialog_Close ' English: "Close"
            Text = "Cerrar"
    End Select
End Sub

Private Sub Do_FRENCH(ByRef TextID As Long, ByRef Text As String)
    Select Case TextID
        Case cdUIT_frmColorDialog_Form_Caption ' English: "Color selection"
            Text = "Sélection couleur"
        Case cdUIT_ColorSelector_chkFixedPalette_Caption ' English: "Fixed"
            Text = "Fixe"
        Case cdUIT_ColorSelector_chkFixedPalette_ToolTipText ' English: "Reflects color changes visually in the palette or not"
            Text = "Reflète visuellement les changements de couleur dans la palette ou non"
        Case cdUIT_ColorSelector_cboSliderParameter_ToolTipText ' English: "Slider control parameter"
            Text = "Paramètre du control ascenseur"
        Case cdUIT_ColorSelector_lblMode_Caption ' English: "Mode:"
            Text = "Mode:"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Hue ' English: "Hue"
            Text = "Hue"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Luminance ' English: "Lum."
            Text = "Lum."
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Value ' English: "Value"
            Text = "Valeur"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Saturation ' English: "Sat."
            Text = "Sat."
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Red ' English: "Red"
            Text = "Rouge"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Green ' English: "Green"
            Text = "Vert"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Blue ' English: "Blue"
            Text = "Bleu"
        Case cdUIT_ColorSelector_cboColorSystem_ListItem_HSV ' English: "HSV"
            Text = "HSV"
        Case cdUIT_ColorSelector_cboColorSystem_ListItem_HSL ' English: "HSL"
            Text = "HSL"
        Case cdUIT_frmColorDialog_lblBasicColors_Caption ' English: "Basic colors:"
            Text = "Couleurs de base:"
        Case cdUIT_frmColorDialog_lblRecent_Caption ' English: "Recent:"
            Text = "Récent:"
        Case cdUIT_frmColorDialog_Color_Caption ' English: "color"
            Text = "couleur"
        Case cdUIT_frmColorDialog_ColorNew_Caption ' English: "new"
            Text = "nouvelle"
        Case cdUIT_frmColorDialog_ColorPrevious_Caption ' English: "previous"
            Text = "précédente"
        Case cdUIT_frmColorDialog_MouseWheel_ToolTipStart ' English: "Hold down the Control key to change"
            Text = "Maintenez la touche Control appuyée pour modifier"
        Case cdUIT_frmColorDialog_MouseWheel_ToolTipEnd ' English: "with the mouse wheel, press Shift key to go slowly"
            Text = "avec la molette de la souris, pressez Shift pour defiler doucement"
        Case cdUIT_frmColorDialog_EyeDropper_ToolTip ' English: "Choose a color from anywhere on the screen or press the Escape key to cancel"
            Text = "Choisir une couleur n'importe-où à l'écran ou presser Échap pour annuler"
        Case cdUIT_frmColorDialog_picEyeDropper_ToolTipText ' English: "Choose a color from the screen"
            Text = "Choisir une couleur sur l'écran"
        Case cdUIT_frmColorDialog_lblRed_Caption ' English: "Red:"
            Text = "Rouge:"
        Case cdUIT_frmColorDialog_lblGreen_Caption ' English: "Green:"
            Text = "Vert:"
        Case cdUIT_frmColorDialog_lblBlue_Caption ' English: "Blue:"
            Text = "Bleu:"
        Case cdUIT_frmColorDialog_lblHex_Caption ' English: "Hex:"
            Text = "Hex:"
        Case cdUIT_frmColorDialog_lblHue_Caption ' English: "Hue:"
            Text = "Hue:"
        Case cdUIT_frmColorDialog_lblSaturation_Caption ' English: "Sat.:"
            Text = "Sat.:"
        Case cdUIT_frmColorDialog_Luminance_Caption ' English: "Lum.:"
            Text = "Lum.:"
        Case cdUIT_frmColorDialog_Value_Caption ' English: "Value:"
            Text = "Valeur:"
        Case cdUIT_frmColorDialog_lblColorSystem_Caption ' English: "Mode:"
            Text = "Mode:"
        Case cdUIT_frmColorDialog_lblPalette_Caption ' English: "Palette:"
            Text = "Palette:"
        Case cdUIT_frmColorDialog_cboPalette_ListItem1 ' English: "Wheel, fixed"
            Text = "Molette, fixe"
        Case cdUIT_frmColorDialog_cboPalette_ListItem2 ' English: "Box, fixed"
            Text = "Boite, fixe"
        Case cdUIT_frmColorDialog_cboPalette_ListItem3 ' English: "Wheel, dynamic"
            Text = "Molette, dynamique"
        Case cdUIT_frmColorDialog_cboPalette_ListItem4 ' English: "Box, dynamic"
            Text = "Boite, dynamique"
        Case cdUIT_frmColorDialog_InvalidColorMessage ' English: "The color is not valid."
            Text = "La couleur n'est pas valide."
        Case cdUIT_frmColorDialog_ParameterFullName_Hue ' English: "Hue"
            Text = "Hue"
        Case cdUIT_frmColorDialog_ParameterFullName_Luminance ' English: "Luminance"
            Text = "Luminance"
        Case cdUIT_frmColorDialog_ParameterFullName_Saturation ' English: "Saturation"
            Text = "Saturation"
        Case cdUIT_frmColorDialog_ParameterFullName_Red ' English: "Red"
            Text = "Rouge"
        Case cdUIT_frmColorDialog_ParameterFullName_Green ' English: "Green"
            Text = "Vert"
        Case cdUIT_frmColorDialog_ParameterFullName_Blue ' English: "Blue"
            Text = "Bleu"
        Case cdUIT_frmColorDialog_ParameterFullName_Value ' English: "Value"
            Text = "Valeur"
        Case cdUIT_frmColorDialog_OK ' English: "OK"
            Text = "OK"
        Case cdUIT_frmColorDialog_Cancel ' English: "Cancel"
            Text = "Annuler"
        Case cdUIT_frmColorDialog_Close ' English: "Close"
            Text = "Fermer"
    End Select
End Sub

Private Sub Do_CHINESE_SIMPLIFIED(ByRef TextID As Long, ByRef Text As String)
    ' Microsoft teminology search: https://www.microsoft.com/en-us/language/search
    Select Case TextID
        Case cdUIT_frmColorDialog_Form_Caption ' English: "Color selection"
            Text = StrFromArray(156, 152, 114, 130, 9, 144, 233, 98) ' ANSI: "ÑÕÉ«Ñ¡Ôñ"
        Case cdUIT_ColorSelector_chkFixedPalette_Caption ' English: "Fixed"
            Text = StrFromArray(250, 86, 154, 91) ' ANSI: "¹Ì¶¨"
        Case cdUIT_ColorSelector_chkFixedPalette_ToolTipText ' English: "Reflects color changes visually in the palette or not"
            Text = StrFromArray(47, 102, 38, 84, 40, 87, 3, 140, 114, 130, 127, 103, 45, 78, 244, 118, 194, 137, 48, 87, 205, 83, 32, 102, 156, 152, 114, 130, 216, 83, 22, 83) ' ANSI: "ÊÇ·ñÔÚµ÷É«°åÖÐÖ±¹ÛµØ·´Ó³ÑÕÉ«±ä»¯"
        Case cdUIT_ColorSelector_cboSliderParameter_ToolTipText ' English: "Slider control parameter"
            Text = StrFromArray(209, 110, 87, 87, 167, 99, 246, 78, 132, 118, 194, 83, 112, 101) ' ANSI: "»¬¿é¿Ø¼þµÄ²ÎÊý"
        Case cdUIT_ColorSelector_lblMode_Caption ' English: "Mode:"
            Text = StrFromArray(33, 106, 15, 95) ' ANSI: "Ä£Ê½"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Hue ' English: "Hue"
            Text = StrFromArray(114, 130, 3, 140) ' ANSI: "É«µ÷"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Luminance ' English: "Lum."
            Text = StrFromArray(174, 78, 166, 94) ' ANSI: "ÁÁ¶È"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Value ' English: "Value"
            Text = StrFromArray(60, 80) ' ANSI: "Öµ"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Saturation ' English: "Sat."
            Text = StrFromArray(113, 153, 140, 84, 166, 94) ' ANSI: "±¥ºÍ¶È"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Red ' English: "Red"
            Text = StrFromArray(162, 126, 114, 130) ' ANSI: "ºìÉ«"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Green ' English: "Green"
            Text = StrFromArray(255, 126, 114, 130) ' ANSI: "ÂÌÉ«"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Blue ' English: "Blue"
            Text = StrFromArray(221, 132, 114, 130) ' ANSI: "À¶É«"
        Case cdUIT_ColorSelector_cboColorSystem_ListItem_HSV ' English: "HSV"
            Text = "HSV"
        Case cdUIT_ColorSelector_cboColorSystem_ListItem_HSL ' English: "HSL"
            Text = "HSL"
        Case cdUIT_frmColorDialog_lblBasicColors_Caption ' English: "Basic colors:"
            Text = StrFromArray(250, 87, 44, 103, 156, 152, 114, 130, 58, 0) ' ANSI: "»ù±¾ÑÕÉ«:"
        Case cdUIT_frmColorDialog_lblRecent_Caption ' English: "Recent:"
            Text = StrFromArray(0, 103, 209, 143, 58, 0) ' ANSI: "×î½ü:"
        Case cdUIT_frmColorDialog_Color_Caption ' English: "color"
            Text = StrFromArray(156, 152, 114, 130) ' ANSI: "ÑÕÉ«"
        Case cdUIT_frmColorDialog_ColorNew_Caption ' English: "new"
            Text = StrFromArray(176, 101, 250, 94) ' ANSI: "ÐÂ½¨"
        Case cdUIT_frmColorDialog_ColorPrevious_Caption ' English: "previous"
            Text = StrFromArray(10, 78, 0, 78, 42, 78) ' ANSI: "ÉÏÒ»¸ö"
        Case cdUIT_frmColorDialog_MouseWheel_ToolTipStart ' English: "Hold down the Control key to change"
            Text = StrFromArray(9, 99, 79, 79, 32, 0, 67, 0, 116, 0, 114, 0, 108, 0, 32, 0, 46, 149, 219, 143, 76, 136, 252, 91, 42, 130) ' ANSI: "°´×¡ Ctrl ¼ü½øÐÐµ¼º½"
        Case cdUIT_frmColorDialog_MouseWheel_ToolTipEnd ' English: "with the mouse wheel, press Shift key to go slowly"
            Text = StrFromArray(127, 79, 40, 117, 32, 159, 7, 104, 218, 110, 110, 143, 22, 98, 9, 99, 32, 0, 83, 0, 104, 0, 105, 0, 102, 0, 116, 0, 32, 0, 46, 149, 19, 127, 98, 97, 251, 121, 168, 82) ' ANSI: "Ê¹ÓÃÊó±ê¹öÂÖ»ò°´ Shift ¼ü»ºÂýÒÆ¶¯"
        Case cdUIT_frmColorDialog_EyeDropper_ToolTip ' English: "Choose a color from anywhere on the screen or press the Escape key to cancel"
            Text = StrFromArray(206, 78, 79, 92, 85, 94, 10, 78, 132, 118, 251, 78, 15, 97, 77, 79, 110, 127, 9, 144, 233, 98, 0, 78, 205, 121, 156, 152, 114, 130, 22, 98, 9, 99, 32, 0, 69, 0, 115, 0, 99, 0, 32, 0, 46, 149, 214, 83, 136, 109) ' ANSI: "´ÓÆÁÄ»ÉÏµÄÈÎÒâÎ»ÖÃÑ¡ÔñÒ»ÖÖÑÕÉ«»ò°´ Esc ¼üÈ¡Ïû"
        Case cdUIT_frmColorDialog_picEyeDropper_ToolTipText ' English: "Choose a color from the screen"
            Text = StrFromArray(206, 78, 79, 92, 85, 94, 45, 78, 9, 144, 233, 98, 0, 78, 205, 121, 156, 152, 114, 130) ' ANSI: "´ÓÆÁÄ»ÖÐÑ¡ÔñÒ»ÖÖÑÕÉ«"
        Case cdUIT_frmColorDialog_lblRed_Caption ' English: "Red:"
            Text = StrFromArray(162, 126, 114, 130, 58, 0) ' ANSI: "ºìÉ«:"
        Case cdUIT_frmColorDialog_lblGreen_Caption ' English: "Green:"
            Text = StrFromArray(255, 126, 114, 130, 58, 0) ' ANSI: "ÂÌÉ«:"
        Case cdUIT_frmColorDialog_lblBlue_Caption ' English: "Blue:"
            Text = StrFromArray(221, 132, 114, 130, 58, 0) ' ANSI: "À¶É«:"
        Case cdUIT_frmColorDialog_lblHex_Caption ' English: "Hex:"
            Text = StrFromArray(65, 83, 109, 81, 219, 143, 54, 82, 58, 0) ' ANSI: "Ê®Áù½øÖÆ:"
        Case cdUIT_frmColorDialog_lblHue_Caption ' English: "Hue:"
            Text = StrFromArray(114, 130, 3, 140, 58, 0) ' ANSI: "É«µ÷:"
        Case cdUIT_frmColorDialog_lblSaturation_Caption ' English: "Sat.:"
            Text = StrFromArray(113, 153, 140, 84, 166, 94, 58, 0) ' ANSI: "±¥ºÍ¶È:"
        Case cdUIT_frmColorDialog_Luminance_Caption ' English: "Lum.:"
            Text = StrFromArray(174, 78, 166, 94, 58, 0) ' ANSI: "ÁÁ¶È:"
        Case cdUIT_frmColorDialog_Value_Caption ' English: "Value:"
            Text = StrFromArray(60, 80, 58, 0) ' ANSI: "Öµ:"
        Case cdUIT_frmColorDialog_lblColorSystem_Caption ' English: "Mode:"
            Text = StrFromArray(33, 106, 15, 95, 58, 0) ' ANSI: "Ä£Ê½:"
        Case cdUIT_frmColorDialog_lblPalette_Caption ' English: "Palette:"
            Text = StrFromArray(3, 140, 114, 130, 127, 103, 58, 0) ' ANSI: "µ÷É«°å:"
        Case cdUIT_frmColorDialog_cboPalette_ListItem1 ' English: "Wheel, fixed"
            Text = StrFromArray(216, 118, 44, 0, 32, 0, 250, 86, 154, 91) ' ANSI: "ÅÌ, ¹Ì¶¨"
        Case cdUIT_frmColorDialog_cboPalette_ListItem2 ' English: "Box, fixed"
            Text = StrFromArray(70, 104, 44, 0, 32, 0, 250, 86, 154, 91) ' ANSI: "¿ò, ¹Ì¶¨"
        Case cdUIT_frmColorDialog_cboPalette_ListItem3 ' English: "Wheel, dynamic"
            Text = StrFromArray(216, 118, 44, 0, 32, 0, 168, 82, 1, 96, 132, 118) ' ANSI: "ÅÌ, ¶¯Ì¬µÄ"
        Case cdUIT_frmColorDialog_cboPalette_ListItem4 ' English: "Box, dynamic"
            Text = StrFromArray(70, 104, 44, 0, 32, 0, 168, 82, 1, 96, 132, 118) ' ANSI: "¿ò, ¶¯Ì¬µÄ"
        Case cdUIT_frmColorDialog_InvalidColorMessage ' English: "The color is not valid."
            Text = StrFromArray(156, 152, 114, 130, 224, 101, 72, 101) ' ANSI: "ÑÕÉ«ÎÞÐ§"
        Case cdUIT_frmColorDialog_ParameterFullName_Hue ' English: "Hue"
            Text = StrFromArray(114, 130, 3, 140) ' ANSI: "É«µ÷"
        Case cdUIT_frmColorDialog_ParameterFullName_Luminance ' English: "Luminance"
            Text = StrFromArray(174, 78, 166, 94) ' ANSI: "ÁÁ¶È"
        Case cdUIT_frmColorDialog_ParameterFullName_Saturation ' English: "Saturation"
            Text = StrFromArray(113, 153, 140, 84, 166, 94) ' ANSI: "±¥ºÍ¶È"
        Case cdUIT_frmColorDialog_ParameterFullName_Red ' English: "Red"
            Text = StrFromArray(162, 126, 114, 130) ' ANSI: "ºìÉ«"
        Case cdUIT_frmColorDialog_ParameterFullName_Green ' English: "Green"
            Text = StrFromArray(255, 126, 114, 130) ' ANSI: "ÂÌÉ«"
        Case cdUIT_frmColorDialog_ParameterFullName_Blue ' English: "Blue"
            Text = StrFromArray(221, 132, 114, 130) ' ANSI: "À¶É«"
        Case cdUIT_frmColorDialog_ParameterFullName_Value ' English: "Value"
            Text = StrFromArray(60, 80) ' ANSI: "Öµ"
        Case cdUIT_frmColorDialog_OK ' English: "OK"
            Text = StrFromArray(110, 120, 154, 91) ' ANSI: "È·¶¨"
        Case cdUIT_frmColorDialog_Cancel ' English: "Cancel"
            Text = StrFromArray(214, 83, 136, 109) ' ANSI: "È¡Ïû"
        Case cdUIT_frmColorDialog_Close ' English: "Close"
            Text = StrFromArray(115, 81, 237, 149) ' ANSI: "¹Ø±Õ"
    End Select
End Sub

Private Sub Do_ITALIAN(ByRef TextID As Long, ByRef Text As String)
    ' Microsoft teminology search: https://www.microsoft.com/en-us/language/search
    Select Case TextID
        Case cdUIT_frmColorDialog_Form_Caption ' English: "Color selection"
            Text = "Selezione colori"
        Case cdUIT_ColorSelector_chkFixedPalette_Caption ' English: "Fixed"
            Text = "Fissa"
        Case cdUIT_ColorSelector_chkFixedPalette_ToolTipText ' English: "Reflects color changes visually in the palette or not"
            Text = "Riflette visivamente i cambiamenti di colore nella tavolozza o meno"
        Case cdUIT_ColorSelector_cboSliderParameter_ToolTipText ' English: "Slider control parameter"
            Text = "Parametro del dispositivo di scorrimento"
        Case cdUIT_ColorSelector_lblMode_Caption ' English: "Mode:"
            Text = "Modalità:"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Hue ' English: "Hue"
            Text = "Ton."
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Luminance ' English: "Lum."
            Text = "Lum."
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Value ' English: "Value"
            Text = "Valore"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Saturation ' English: "Sat."
            Text = "Sat."
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Red ' English: "Red"
            Text = "Rosso"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Green ' English: "Green"
            Text = "Verde"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Blue ' English: "Blue"
            Text = "Blu"
        Case cdUIT_ColorSelector_cboColorSystem_ListItem_HSV ' English: "HSV"
            Text = "HSV"
        Case cdUIT_ColorSelector_cboColorSystem_ListItem_HSL ' English: "HSL"
            Text = "HSL"
        Case cdUIT_frmColorDialog_lblBasicColors_Caption ' English: "Basic colors:"
            Text = "Colori di base:"
        Case cdUIT_frmColorDialog_lblRecent_Caption ' English: "Recent:"
            Text = "Recenti:"
        Case cdUIT_frmColorDialog_Color_Caption ' English: "color"
            Text = "colore"
        Case cdUIT_frmColorDialog_ColorNew_Caption ' English: "new"
            Text = "nuovo"
        Case cdUIT_frmColorDialog_ColorPrevious_Caption ' English: "previous"
            Text = "precedente"
        Case cdUIT_frmColorDialog_MouseWheel_ToolTipStart ' English: "Hold down the Control key to change"
            Text = "Tieni premuto il tasto CTRL per modificare"
        Case cdUIT_frmColorDialog_MouseWheel_ToolTipEnd ' English: "with the mouse wheel, press Shift key to go slowly"
            Text = "con la rotellina del mouse, premere il tasto MAIUSC per procedere lentamente"
        Case cdUIT_frmColorDialog_EyeDropper_ToolTip ' English: "Choose a color from anywhere on the screen or press the Escape key to cancel"
            Text = "Scegli un colore da qualsiasi punto dello schermo o premi il tasto ESC per annullare"
        Case cdUIT_frmColorDialog_picEyeDropper_ToolTipText ' English: "Choose a color from the screen"
            Text = "Scegli un colore dallo schermo"
        Case cdUIT_frmColorDialog_lblRed_Caption ' English: "Red:"
            Text = "Rosso:"
        Case cdUIT_frmColorDialog_lblGreen_Caption ' English: "Green:"
            Text = "Verde:"
        Case cdUIT_frmColorDialog_lblBlue_Caption ' English: "Blue:"
            Text = "Blu:"
        Case cdUIT_frmColorDialog_lblHex_Caption ' English: "Hex:"
            Text = "Esa.:"
        Case cdUIT_frmColorDialog_lblHue_Caption ' English: "Hue:"
            Text = "Tonalità:"
        Case cdUIT_frmColorDialog_lblSaturation_Caption ' English: "Sat.:"
            Text = "Sat.:"
        Case cdUIT_frmColorDialog_Luminance_Caption ' English: "Lum.:"
            Text = "Lum.:"
        Case cdUIT_frmColorDialog_Value_Caption ' English: "Value:"
            Text = "Valore:"
        Case cdUIT_frmColorDialog_lblColorSystem_Caption ' English: "Mode:"
            Text = "Modalità:"
        Case cdUIT_frmColorDialog_lblPalette_Caption ' English: "Palette:"
            Text = "Tavolozza:"
        Case cdUIT_frmColorDialog_cboPalette_ListItem1 ' English: "Wheel, fixed"
            Text = "Ruota, fissa"
        Case cdUIT_frmColorDialog_cboPalette_ListItem2 ' English: "Box, fixed"
            Text = "Casella, fissa"
        Case cdUIT_frmColorDialog_cboPalette_ListItem3 ' English: "Wheel, dynamic"
            Text = "Ruota, dinamica"
        Case cdUIT_frmColorDialog_cboPalette_ListItem4 ' English: "Box, dynamic"
            Text = "Casella, dinamica"
        Case cdUIT_frmColorDialog_InvalidColorMessage ' English: "The color is not valid."
            Text = "Il colore non è valido."
        Case cdUIT_frmColorDialog_ParameterFullName_Hue ' English: "Hue"
            Text = "Tonalità"
        Case cdUIT_frmColorDialog_ParameterFullName_Luminance ' English: "Luminance"
            Text = "Luminosità"
        Case cdUIT_frmColorDialog_ParameterFullName_Saturation ' English: "Saturation"
            Text = "Saturazione"
        Case cdUIT_frmColorDialog_ParameterFullName_Red ' English: "Red"
            Text = "Rosso"
        Case cdUIT_frmColorDialog_ParameterFullName_Green ' English: "Green"
            Text = "Verde"
        Case cdUIT_frmColorDialog_ParameterFullName_Blue ' English: "Blue"
            Text = "Blu"
        Case cdUIT_frmColorDialog_ParameterFullName_Value ' English: "Value"
            Text = "Valore"
        Case cdUIT_frmColorDialog_OK ' English: "OK"
            Text = "OK"
        Case cdUIT_frmColorDialog_Cancel ' English: "Cancel"
            Text = "Annulla"
        Case cdUIT_frmColorDialog_Close ' English: "Close"
            Text = "Chiudi"
    End Select
End Sub

Private Sub Do_GREEK(ByRef TextID As Long, ByRef Text As String)
    ' Microsoft teminology search: https://www.microsoft.com/en-us/language/search
    Select Case TextID
        Case cdUIT_frmColorDialog_Form_Caption ' English: "Color selection"
            Text = StrFromArray(149, 3, 192, 3, 185, 3, 187, 3, 191, 3, 179, 3, 174, 3, 194, 3, 32, 0, 167, 3, 193, 3, 201, 3, 188, 3, 172, 3, 196, 3, 201, 3, 189, 3) ' ANSI: "ÅðéëïãÞò ×ñùìÜôùí"
        Case cdUIT_ColorSelector_chkFixedPalette_Caption ' English: "Fixed"
            Text = StrFromArray(163, 3, 196, 3, 177, 3, 184, 3, 181, 3, 193, 3, 204, 3) ' ANSI: "Óôáèåñü"
        Case cdUIT_ColorSelector_chkFixedPalette_ToolTipText ' English: "Reflects color changes visually in the palette or not"
            Text = StrFromArray(145, 3, 189, 3, 196, 3, 177, 3, 189, 3, 177, 3, 186, 3, 187, 3, 172, 3, 32, 0, 196, 3, 185, 3, 194, 3, 32, 0, 177, 3, 187, 3, 187, 3, 177, 3, 179, 3, 173, 3, 194, 3, 32, 0, 199, 3, 193, 3, 206, 3, 188, 3, 177, 3, 196, 3, 191, 3, 194, 3, 32, 0, 191, 3, 192, 3, 196, 3, 185, 3, 186, 3, 172, 3, 32, 0, 195, 3, 196, 3, 183, 3, 189, 3, 32, 0, 192, 3, 177, 3, 187, 3, 173, 3, 196, 3, 177, 3, 32, 0, 174, 3, 32, 0, 204, 3, 199, 3, 185, 3) ' ANSI: "ÁíôáíáêëÜ ôéò áëëáãÝò ÷ñþìáôïò ïðôéêÜ óôçí ðáëÝôá Þ ü÷é"
        Case cdUIT_ColorSelector_cboSliderParameter_ToolTipText ' English: "Slider control parameter"
            Text = StrFromArray(160, 3, 177, 3, 193, 3, 172, 3, 188, 3, 181, 3, 196, 3, 193, 3, 191, 3, 194, 3, 32, 0, 195, 3, 196, 3, 191, 3, 185, 3, 199, 3, 181, 3, 175, 3, 191, 3, 32, 0, 181, 3, 187, 3, 173, 3, 179, 3, 199, 3, 191, 3, 197, 3, 32, 0, 193, 3, 197, 3, 184, 3, 188, 3, 185, 3, 195, 3, 196, 3, 185, 3, 186, 3, 191, 3, 205, 3) ' ANSI: "ÐáñÜìåôñïò óôïé÷åßï åëÝã÷ïõ ñõèìéóôéêïý"
        Case cdUIT_ColorSelector_lblMode_Caption ' English: "Mode:"
            Text = StrFromArray(164, 3, 205, 3, 192, 3, 191, 3, 194, 3, 58, 0) ' ANSI: "Ôýðïò:"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Hue ' English: "Hue"
            Text = StrFromArray(145, 3, 192, 3, 191, 3, 199, 3, 46, 0) ' ANSI: "Áðï÷."
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Luminance ' English: "Lum."
            Text = StrFromArray(166, 3, 201, 3, 196, 3, 46, 0) ' ANSI: "Öùô."
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Value ' English: "Value"
            Text = StrFromArray(145, 3, 190, 3, 175, 3, 177, 3) ' ANSI: "Áîßá"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Saturation ' English: "Sat."
            Text = StrFromArray(154, 3, 191, 3, 193, 3, 46, 0) ' ANSI: "Êïñ."
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Red ' English: "Red"
            Text = StrFromArray(154, 3, 204, 3, 186, 3, 186, 3, 46, 0) ' ANSI: "Êüêê."
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Green ' English: "Green"
            Text = StrFromArray(160, 3, 193, 3, 172, 3, 195, 3, 46, 0) ' ANSI: "ÐñÜó."
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Blue ' English: "Blue"
            Text = StrFromArray(156, 3, 192, 3, 187, 3, 181, 3) ' ANSI: "Ìðëå"
        Case cdUIT_ColorSelector_cboColorSystem_ListItem_HSV ' English: "HSV"
            Text = "HSV"
        Case cdUIT_ColorSelector_cboColorSystem_ListItem_HSL ' English: "HSL"
            Text = "HSL"
        Case cdUIT_frmColorDialog_lblBasicColors_Caption ' English: "Basic colors:"
            Text = StrFromArray(146, 3, 177, 3, 195, 3, 185, 3, 186, 3, 172, 3, 32, 0, 199, 3, 193, 3, 206, 3, 188, 3, 177, 3, 196, 3, 177, 3, 58, 0) ' ANSI: "ÂáóéêÜ ÷ñþìáôá:"
        Case cdUIT_frmColorDialog_lblRecent_Caption ' English: "Recent:"
            Text = StrFromArray(160, 3, 193, 3, 204, 3, 195, 3, 198, 3, 177, 3, 196, 3, 177, 3, 58, 0) ' ANSI: "Ðñüóöáôá:"
        Case cdUIT_frmColorDialog_Color_Caption ' English: "color"
            Text = StrFromArray(199, 3, 193, 3, 206, 3, 188, 3, 177, 3) ' ANSI: "÷ñþìá"
        Case cdUIT_frmColorDialog_ColorNew_Caption ' English: "new"
            Text = StrFromArray(189, 3, 173, 3, 191, 3) ' ANSI: "íÝï"
        Case cdUIT_frmColorDialog_ColorPrevious_Caption ' English: "previous"
            Text = StrFromArray(192, 3, 193, 3, 191, 3, 183, 3, 179, 3, 191, 3, 205, 3, 188, 3, 181, 3, 189, 3, 191, 3) ' ANSI: "ðñïçãïýìåíï"
        Case cdUIT_frmColorDialog_MouseWheel_ToolTipStart ' English: "Hold down the Control key to change"
            Text = StrFromArray(154, 3, 193, 3, 177, 3, 196, 3, 174, 3, 195, 3, 196, 3, 181, 3, 32, 0, 192, 3, 177, 3, 196, 3, 183, 3, 188, 3, 173, 3, 189, 3, 191, 3, 32, 0, 196, 3, 191, 3, 32, 0, 192, 3, 187, 3, 174, 3, 186, 3, 196, 3, 193, 3, 191, 3, 32, 0, 67, 0, 111, 0, 110, 0, 116, 0, 114, 0, 111, 0, 108, 0, 32, 0, 179, 3, 185, 3, 177, 3, 32, 0, 177, 3, 187, 3, 187, 3, 177, 3, 179, 3, 174, 3) ' ANSI: "ÊñáôÞóôå ðáôçìÝíï ôï ðëÞêôñï Control ãéá áëëáãÞ"
        Case cdUIT_frmColorDialog_MouseWheel_ToolTipEnd ' English: "with the mouse wheel, press Shift key to go slowly"
            Text = StrFromArray(188, 3, 181, 3, 32, 0, 196, 3, 191, 3, 189, 3, 32, 0, 196, 3, 193, 3, 191, 3, 199, 3, 204, 3, 32, 0, 196, 3, 191, 3, 197, 3, 32, 0, 192, 3, 191, 3, 189, 3, 196, 3, 185, 3, 186, 3, 185, 3, 191, 3, 205, 3, 44, 0, 32, 0, 192, 3, 177, 3, 196, 3, 174, 3, 195, 3, 196, 3, 181, 3, 32, 0, 196, 3, 191, 3, 32, 0, 192, 3, 187, 3, 174, 3, 186, 3, 196, 3, 193, 3, 191, 3, 32, 0, 83, 0, 104, 0, 105, 0, 102, 0, 116, 0, 32, 0, 179, 3, 185, 3, 177, 3, 32, 0, 189, 3, 177, 3, 32, 0, 192, 3, 193, 3, 191, 3, 199, 3, 201, 3, 193, 3, 174, 3, 195, 3, 181, 3, 196, 3, 181, 3, 32, 0, 177, 3, 193, 3, 179, 3, 172, 3) ' ANSI: "ìå ôïí ôñï÷ü ôïõ ðïíôéêéïý, ðáôÞóôå ôï ðëÞêôñï Shift ãéá íá ðñï÷ùñÞóåôå áñãÜ"
        Case cdUIT_frmColorDialog_EyeDropper_ToolTip ' English: "Choose a color from anywhere on the screen or press the Escape key to cancel"
            Text = StrFromArray(149, 3, 192, 3, 185, 3, 187, 3, 173, 3, 190, 3, 196, 3, 181, 3, 32, 0, 173, 3, 189, 3, 177, 3, 32, 0, 199, 3, 193, 3, 206, 3, 188, 3, 177, 3, 32, 0, 177, 3, 192, 3, 204, 3, 32, 0, 191, 3, 192, 3, 191, 3, 197, 3, 180, 3, 174, 3, 192, 3, 191, 3, 196, 3, 181, 3, 32, 0, 195, 3, 196, 3, 183, 3, 189, 3, 32, 0, 191, 3, 184, 3, 204, 3, 189, 3, 183, 3, 32, 0, 174, 3, 32, 0, 192, 3, 177, 3, 196, 3, 174, 3, 195, 3, 196, 3, 181, 3, 32, 0, 196, 3, 191, 3, 32, 0, 192, 3, 187, 3, 174, 3, 186, 3, 196, 3, 193, 3, 191, 3, 32, 0, 69, 0, 115, 0, 99, 0, 97, 0, 112, 0, 101, 0, 32, 0, 179, 3, 185, 3, 177, 3, 32, 0, 177, 3, 186, 3, 205, 3, 193, 3, 201, 3, 195, 3, 183, 3) ' ANSI: "ÅðéëÝîôå Ýíá ÷ñþìá áðü ïðïõäÞðïôå óôçí ïèüíç Þ ðáôÞóôå ôï ðëÞêôñï Escape ãéá áêýñùóç"
        Case cdUIT_frmColorDialog_picEyeDropper_ToolTipText ' English: "Choose a color from the screen"
            Text = StrFromArray(149, 3, 192, 3, 185, 3, 187, 3, 173, 3, 190, 3, 196, 3, 181, 3, 32, 0, 173, 3, 189, 3, 177, 3, 32, 0, 199, 3, 193, 3, 206, 3, 188, 3, 177, 3, 32, 0, 177, 3, 192, 3, 204, 3, 32, 0, 196, 3, 183, 3, 189, 3, 32, 0, 191, 3, 184, 3, 204, 3, 189, 3, 183, 3) ' ANSI: "ÅðéëÝîôå Ýíá ÷ñþìá áðü ôçí ïèüíç"
        Case cdUIT_frmColorDialog_lblRed_Caption ' English: "Red:"
            Text = StrFromArray(154, 3, 204, 3, 186, 3, 186, 3, 185, 3, 189, 3, 191, 3, 58, 0) ' ANSI: "Êüêêéíï:"
        Case cdUIT_frmColorDialog_lblGreen_Caption ' English: "Green:"
            Text = StrFromArray(160, 3, 193, 3, 172, 3, 195, 3, 185, 3, 189, 3, 191, 3, 58, 0) ' ANSI: "ÐñÜóéíï:"
        Case cdUIT_frmColorDialog_lblBlue_Caption ' English: "Blue:"
            Text = StrFromArray(156, 3, 192, 3, 187, 3, 181, 3, 58, 0) ' ANSI: "Ìðëå:"
        Case cdUIT_frmColorDialog_lblHex_Caption ' English: "Hex:"
            Text = "Hex:"
        Case cdUIT_frmColorDialog_lblHue_Caption ' English: "Hue:"
            Text = StrFromArray(145, 3, 192, 3, 191, 3, 199, 3, 46, 0, 58, 0) ' ANSI: "Áðï÷.:"
        Case cdUIT_frmColorDialog_lblSaturation_Caption ' English: "Sat.:"
            Text = StrFromArray(154, 3, 191, 3, 193, 3, 46, 0, 58, 0) ' ANSI: "Êïñ.:"
        Case cdUIT_frmColorDialog_Luminance_Caption ' English: "Lum.:"
            Text = StrFromArray(166, 3, 201, 3, 196, 3, 46, 0, 58, 0) ' ANSI: "Öùô.:"
        Case cdUIT_frmColorDialog_Value_Caption ' English: "Value:"
            Text = StrFromArray(145, 3, 190, 3, 175, 3, 177, 3, 58, 0) ' ANSI: "Áîßá:"
        Case cdUIT_frmColorDialog_lblColorSystem_Caption ' English: "Mode:"
            Text = StrFromArray(164, 3, 205, 3, 192, 3, 191, 3, 194, 3, 58, 0) ' ANSI: "Ôýðïò:"
        Case cdUIT_frmColorDialog_lblPalette_Caption ' English: "Palette:"
            Text = StrFromArray(160, 3, 177, 3, 187, 3, 173, 3, 196, 3, 177, 3, 58, 0) ' ANSI: "ÐáëÝôá:"
        Case cdUIT_frmColorDialog_cboPalette_ListItem1 ' English: "Wheel, fixed"
            Text = StrFromArray(164, 3, 193, 3, 191, 3, 199, 3, 204, 3, 194, 3, 44, 0, 32, 0, 195, 3, 196, 3, 177, 3, 184, 3, 181, 3, 193, 3, 204, 3, 194, 3) ' ANSI: "Ôñï÷üò, óôáèåñüò"
        Case cdUIT_frmColorDialog_cboPalette_ListItem2 ' English: "Box, fixed"
            Text = StrFromArray(160, 3, 187, 3, 177, 3, 175, 3, 195, 3, 185, 3, 191, 3, 44, 0, 32, 0, 195, 3, 196, 3, 177, 3, 184, 3, 181, 3, 193, 3, 204, 3) ' ANSI: "Ðëáßóéï, óôáèåñü"
        Case cdUIT_frmColorDialog_cboPalette_ListItem3 ' English: "Wheel, dynamic"
            Text = StrFromArray(164, 3, 193, 3, 191, 3, 199, 3, 204, 3, 194, 3, 44, 0, 32, 0, 180, 3, 197, 3, 189, 3, 177, 3, 188, 3, 185, 3, 186, 3, 204, 3, 194, 3) ' ANSI: "Ôñï÷üò, äõíáìéêüò"
        Case cdUIT_frmColorDialog_cboPalette_ListItem4 ' English: "Box, dynamic"
            Text = StrFromArray(160, 3, 187, 3, 177, 3, 175, 3, 195, 3, 185, 3, 191, 3, 44, 0, 32, 0, 180, 3, 197, 3, 189, 3, 177, 3, 188, 3, 185, 3, 186, 3, 204, 3) ' ANSI: "Ðëáßóéï, äõíáìéêü"
        Case cdUIT_frmColorDialog_InvalidColorMessage ' English: "The color is not valid."
            Text = StrFromArray(164, 3, 191, 3, 32, 0, 199, 3, 193, 3, 206, 3, 188, 3, 177, 3, 32, 0, 180, 3, 181, 3, 189, 3, 32, 0, 181, 3, 175, 3, 189, 3, 177, 3, 185, 3, 32, 0, 173, 3, 179, 3, 186, 3, 197, 3, 193, 3, 191, 3, 46, 0) ' ANSI: "Ôï ÷ñþìá äåí åßíáé Ýãêõñï."
        Case cdUIT_frmColorDialog_ParameterFullName_Hue ' English: "Hue"
            Text = StrFromArray(145, 3, 192, 3, 204, 3, 199, 3, 193, 3, 201, 3, 195, 3, 183, 3) ' ANSI: "Áðü÷ñùóç"
        Case cdUIT_frmColorDialog_ParameterFullName_Luminance ' English: "Luminance"
            Text = StrFromArray(166, 3, 201, 3, 196, 3, 181, 3, 185, 3, 189, 3, 204, 3, 196, 3, 183, 3, 196, 3, 177, 3) ' ANSI: "Öùôåéíüôçôá"
        Case cdUIT_frmColorDialog_ParameterFullName_Saturation ' English: "Saturation"
            Text = StrFromArray(154, 3, 191, 3, 193, 3, 181, 3, 195, 3, 188, 3, 204, 3, 194, 3) ' ANSI: "Êïñåóìüò"
        Case cdUIT_frmColorDialog_ParameterFullName_Red ' English: "Red"
            Text = StrFromArray(154, 3, 204, 3, 186, 3, 186, 3, 185, 3, 189, 3, 191, 3) ' ANSI: "Êüêêéíï"
        Case cdUIT_frmColorDialog_ParameterFullName_Green ' English: "Green"
            Text = StrFromArray(160, 3, 193, 3, 172, 3, 195, 3, 185, 3, 189, 3, 191, 3) ' ANSI: "ÐñÜóéíï"
        Case cdUIT_frmColorDialog_ParameterFullName_Blue ' English: "Blue"
            Text = StrFromArray(156, 3, 192, 3, 187, 3, 181, 3) ' ANSI: "Ìðëå"
        Case cdUIT_frmColorDialog_ParameterFullName_Value ' English: "Value"
            Text = StrFromArray(145, 3, 190, 3, 175, 3, 177, 3, 58, 0) ' ANSI: "Áîßá:"
        Case cdUIT_frmColorDialog_OK ' English: "OK"
            Text = "OK"
        Case cdUIT_frmColorDialog_Cancel ' English: "Cancel"
            Text = StrFromArray(134, 3, 186, 3, 197, 3, 193, 3, 191, 3) ' ANSI: "¢êõñï"
        Case cdUIT_frmColorDialog_Close ' English: "Close"
            Text = StrFromArray(154, 3, 187, 3, 181, 3, 175, 3, 195, 3, 185, 3, 188, 3, 191, 3) ' ANSI: "Êëåßóéìï"
    End Select
End Sub

Private Sub Do_GERMAN(ByRef TextID As Long, ByRef Text As String)
    ' Microsoft teminology search: https://www.microsoft.com/en-us/language/search
    Select Case TextID
        Case cdUIT_frmColorDialog_Form_Caption ' English: "Color selection"
            Text = "Farbwahl"
        Case cdUIT_ColorSelector_chkFixedPalette_Caption ' English: "Fixed"
            Text = "Feste"
        Case cdUIT_ColorSelector_chkFixedPalette_ToolTipText ' English: "Reflects color changes visually in the palette or not"
            Text = "Reflektiert Farbänderungen visuell in der Palette oder nicht"
        Case cdUIT_ColorSelector_cboSliderParameter_ToolTipText ' English: "Slider control parameter"
            Text = "Schieberegler-Parameter"
        Case cdUIT_ColorSelector_lblMode_Caption ' English: "Mode:"
            Text = "Modus:"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Hue ' English: "Hue"
            Text = "Farbt."
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Luminance ' English: "Lum."
            Text = "Hell."
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Value ' English: "Value"
            Text = "Wert"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Saturation ' English: "Sat."
            Text = "Sätt."
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Red ' English: "Red"
            Text = "Rot"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Green ' English: "Green"
            Text = "Grün"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Blue ' English: "Blue"
            Text = "Blau"
        Case cdUIT_ColorSelector_cboColorSystem_ListItem_HSV ' English: "HSV"
            Text = "HSV"
        Case cdUIT_ColorSelector_cboColorSystem_ListItem_HSL ' English: "HSL"
            Text = "HSL"
        Case cdUIT_frmColorDialog_lblBasicColors_Caption ' English: "Basic colors:"
            Text = "Grundfarben:"
        Case cdUIT_frmColorDialog_lblRecent_Caption ' English: "Recent:"
            Text = "Letzte:"
        Case cdUIT_frmColorDialog_Color_Caption ' English: "color"
            Text = "Farbe"
        Case cdUIT_frmColorDialog_ColorNew_Caption ' English: "new"
            Text = "neue"
        Case cdUIT_frmColorDialog_ColorPrevious_Caption ' English: "previous"
            Text = "vorherige"
        Case cdUIT_frmColorDialog_MouseWheel_ToolTipStart ' English: "Hold down the Control key to change"
            Text = "Halten Sie zum Ändern die Strg-Taste gedrückt"
        Case cdUIT_frmColorDialog_MouseWheel_ToolTipEnd ' English: "with the mouse wheel, press Shift key to go slowly"
            Text = "Mit dem Mausrad drücken Sie die Umschalttaste, um langsam zu fahren"
        Case cdUIT_frmColorDialog_EyeDropper_ToolTip ' English: "Choose a color from anywhere on the screen or press the Escape key to cancel"
            Text = "Wählen Sie an einer beliebigen Stelle auf dem Bildschirm eine Farbe aus oder drücken Sie die Esc-Taste, um den Vorgang abzubrechen"
        Case cdUIT_frmColorDialog_picEyeDropper_ToolTipText ' English: "Choose a color from the screen"
            Text = "Wählen Sie eine Farbe auf dem Bildschirm aus"
        Case cdUIT_frmColorDialog_lblRed_Caption ' English: "Red:"
            Text = "Rot:"
        Case cdUIT_frmColorDialog_lblGreen_Caption ' English: "Green:"
            Text = "Grün:"
        Case cdUIT_frmColorDialog_lblBlue_Caption ' English: "Blue:"
            Text = "Blau:"
        Case cdUIT_frmColorDialog_lblHex_Caption ' English: "Hex:"
            Text = "Hex:"
        Case cdUIT_frmColorDialog_lblHue_Caption ' English: "Hue:"
            Text = "Farbton:"
        Case cdUIT_frmColorDialog_lblSaturation_Caption ' English: "Sat.:"
            Text = "Sätt:"
        Case cdUIT_frmColorDialog_Luminance_Caption ' English: "Lum.:"
            Text = "Hell.:"
        Case cdUIT_frmColorDialog_Value_Caption ' English: "Value:"
            Text = "Wert:"
        Case cdUIT_frmColorDialog_lblColorSystem_Caption ' English: "Mode:"
            Text = "Modus:"
        Case cdUIT_frmColorDialog_lblPalette_Caption ' English: "Palette:"
            Text = "Palette:"
        Case cdUIT_frmColorDialog_cboPalette_ListItem1 ' English: "Wheel, fixed"
            Text = "Rad, fest"
        Case cdUIT_frmColorDialog_cboPalette_ListItem2 ' English: "Box, fixed"
            Text = "Schachtel, fest"
        Case cdUIT_frmColorDialog_cboPalette_ListItem3 ' English: "Wheel, dynamic"
            Text = "Rad, dynamisch"
        Case cdUIT_frmColorDialog_cboPalette_ListItem4 ' English: "Box, dynamic"
            Text = "Schachtel, dynamisch"
        Case cdUIT_frmColorDialog_InvalidColorMessage ' English: "The color is not valid."
            Text = "The color is not valid."
        Case cdUIT_frmColorDialog_ParameterFullName_Hue ' English: "Hue"
            Text = "Farbton"
        Case cdUIT_frmColorDialog_ParameterFullName_Luminance ' English: "Luminance"
            Text = "Helligkeit"
        Case cdUIT_frmColorDialog_ParameterFullName_Saturation ' English: "Saturation"
            Text = "Saturation"
        Case cdUIT_frmColorDialog_ParameterFullName_Red ' English: "Red"
            Text = "Rot"
        Case cdUIT_frmColorDialog_ParameterFullName_Green ' English: "Green"
            Text = "Grün"
        Case cdUIT_frmColorDialog_ParameterFullName_Blue ' English: "Blue"
            Text = "Blau"
        Case cdUIT_frmColorDialog_ParameterFullName_Value ' English: "Value"
            Text = "Wert"
        Case cdUIT_frmColorDialog_OK ' English: "OK"
            Text = "OK"
        Case cdUIT_frmColorDialog_Cancel ' English: "Cancel"
            Text = "Abbrechen"
        Case cdUIT_frmColorDialog_Close ' English: "Close"
            Text = "Schließen"
    End Select
End Sub

'Private Sub Do_OTHER(ByRef TextID As Long, ByRef Text As String)
'    ' Microsoft teminology search: https://www.microsoft.com/en-us/language/search
'    Select Case TextID
'        Case cdUIT_frmColorDialog_Form_Caption ' English: "Color selection"
'            Text = "Color selection"
'        Case cdUIT_ColorSelector_chkFixedPalette_Caption ' English: "Fixed"
'            Text = "Fixed"
'        Case cdUIT_ColorSelector_chkFixedPalette_ToolTipText ' English: "Reflects color changes visually in the palette or not"
'            Text = "Reflects color changes visually in the palette or not"
'        Case cdUIT_ColorSelector_cboSliderParameter_ToolTipText ' English: "Slider control parameter"
'            Text = "Slider control parameter"
'        Case cdUIT_ColorSelector_lblMode_Caption ' English: "Mode:"
'            Text = "Mode:"
'        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Hue ' English: "Hue"
'            Text = "Hue"
'        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Luminance ' English: "Lum."
'            Text = "Lum."
'        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Value ' English: "Value"
'            Text = "Value"
'        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Saturation ' English: "Sat."
'            Text = "Sat."
'        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Red ' English: "Red"
'            Text = "Red"
'        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Green ' English: "Green"
'            Text = "Green"
'        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Blue ' English: "Blue"
'            Text = "Blue"
'        Case cdUIT_ColorSelector_cboColorSystem_ListItem_HSV ' English: "HSV"
'            Text = "HSV"
'        Case cdUIT_ColorSelector_cboColorSystem_ListItem_HSL ' English: "HSL"
'            Text = "HSL"
'        Case cdUIT_frmColorDialog_lblBasicColors_Caption ' English: "Basic colors:"
'            Text = "Basic colors:"
'        Case cdUIT_frmColorDialog_lblRecent_Caption ' English: "Recent:"
'            Text = "Recent:"
'        Case cdUIT_frmColorDialog_Color_Caption ' English: "color"
'            Text = "color"
'        Case cdUIT_frmColorDialog_ColorNew_Caption ' English: "new"
'            Text = "new"
'        Case cdUIT_frmColorDialog_ColorPrevious_Caption ' English: "previous"
'            Text = "previous"
'        Case cdUIT_frmColorDialog_MouseWheel_ToolTipStart ' English: "Hold down the Control key to change"
'            Text = "Hold down the Control key to change"
'        Case cdUIT_frmColorDialog_MouseWheel_ToolTipEnd ' English: "with the mouse wheel, press Shift key to go slowly"
'            Text = "with the mouse wheel, press Shift key to go slowly"
'        Case cdUIT_frmColorDialog_EyeDropper_ToolTip ' English: "Choose a color from anywhere on the screen or press the Escape key to cancel"
'            Text = "Choose a color from anywhere on the screen or press the Escape key to cancel"
'        Case cdUIT_frmColorDialog_picEyeDropper_ToolTipText ' English: "Choose a color from the screen"
'            Text = "Choose a color from the screen"
'        Case cdUIT_frmColorDialog_lblRed_Caption ' English: "Red:"
'            Text = "Red:"
'        Case cdUIT_frmColorDialog_lblGreen_Caption ' English: "Green:"
'            Text = "Green:"
'        Case cdUIT_frmColorDialog_lblBlue_Caption ' English: "Blue:"
'            Text = "Blue:"
'        Case cdUIT_frmColorDialog_lblHex_Caption ' English: "Hex:"
'            Text = "Hex:"
'        Case cdUIT_frmColorDialog_lblHue_Caption ' English: "Hue:"
'            Text = "Hue:"
'        Case cdUIT_frmColorDialog_lblSaturation_Caption ' English: "Sat.:"
'            Text = "Sat.:"
'        Case cdUIT_frmColorDialog_Luminance_Caption ' English: "Lum.:"
'            Text = "Lum.:"
'        Case cdUIT_frmColorDialog_Value_Caption ' English: "Value:"
'            Text = "Value:"
'        Case cdUIT_frmColorDialog_lblColorSystem_Caption ' English: "Mode:"
'            Text = "Mode:"
'        Case cdUIT_frmColorDialog_lblPalette_Caption ' English: "Palette:"
'            Text = "Palette:"
'        Case cdUIT_frmColorDialog_cboPalette_ListItem1 ' English: "Wheel, fixed"
'            Text = "Wheel, fixed"
'        Case cdUIT_frmColorDialog_cboPalette_ListItem2 ' English: "Box, fixed"
'            Text = "Box, fixed"
'        Case cdUIT_frmColorDialog_cboPalette_ListItem3 ' English: "Wheel, dynamic"
'            Text = "Wheel, dynamic"
'        Case cdUIT_frmColorDialog_cboPalette_ListItem4 ' English: "Box, dynamic"
'            Text = "Box, dynamic"
'        Case cdUIT_frmColorDialog_InvalidColorMessage ' English: "The color is not valid."
'            Text = "The color is not valid."
'        Case cdUIT_frmColorDialog_ParameterFullName_Hue ' English: "Hue"
'            Text = "Hue"
'        Case cdUIT_frmColorDialog_ParameterFullName_Luminance ' English: "Luminance"
'            Text = "Luminance"
'        Case cdUIT_frmColorDialog_ParameterFullName_Saturation ' English: "Saturation"
'            Text = "Saturation"
'        Case cdUIT_frmColorDialog_ParameterFullName_Red ' English: "Red"
'            Text = "Red"
'        Case cdUIT_frmColorDialog_ParameterFullName_Green ' English: "Green"
'            Text = "Green"
'        Case cdUIT_frmColorDialog_ParameterFullName_Blue ' English: "Blue"
'            Text = "Blue"
'        Case cdUIT_frmColorDialog_ParameterFullName_Value ' English: "Value"
'            Text = "Value"
'        Case cdUIT_frmColorDialog_OK ' English: "OK"
'            Text = "OK"
'        Case cdUIT_frmColorDialog_Cancel ' English: "Cancel"
'            Text = "Cancel"
'        Case cdUIT_frmColorDialog_Close ' English: "Close"
'            Text = "Close"
'    End Select
'End Sub

Public Property Get LanguageWindowsUI() As Variant
    If mLanguageWindowsUI = bsLang_AUTO_SELECT Then SetUILanguageToWindowsUILanguage
    LanguageWindowsUI = mLanguageWindowsUI
End Property

Public Property Get SubLanguageWindowsUI() As Variant
    If mSubLanguageWindowsUI = bsLang_AUTO_SELECT Then SetUILanguageToWindowsUILanguage
    SubLanguageWindowsUI = mSubLanguageWindowsUI
End Property

Private Property Get UILanguage() As Variant
    If mUILanguage = bsLang_AUTO_SELECT Then SetUILanguageToWindowsUILanguage
    UILanguage = mUILanguage
End Property

Private Property Get UISubLanguage() As Variant
    If mUILanguage = bsLang_AUTO_SELECT Then SetUILanguageToWindowsUILanguage
    UISubLanguage = mUISubLanguage
End Property

Private Sub SetUILanguageToWindowsUILanguage()
    mLanguageWindowsUI = CLng(GetUserDefaultUILanguage And &HFF)
    mSubLanguageWindowsUI = (GetUserDefaultUILanguage And Not mLanguageWindowsUI) / 1024
    mUILanguage = mLanguageWindowsUI
    mUISubLanguage = mSubLanguageWindowsUI
    If Not LanguageIsSupported(mUILanguage) Then
        mUILanguage = cDefaultLanguage
        mUISubLanguage = 0
    End If
    mUIRightToLeft = IsLanguagueRightToLeft(MAKELANGID(mUILanguage, mUISubLanguage))
    
#Const TestingLanguages = 0
#If TestingLanguages Then
    If InIDE Then
        mUILanguage = bsLang_GERMAN ' bsLang_GREEK ' bsLang_ITALIAN ' bsLang_CHINESE_SIMPLIFIED ' bsLang_FRENCH 'bsLang_ENGLISH
    End If
#End If
End Sub

Private Function LanguageIsSupported(nLang As Long) As Boolean
    Dim c As Long
    
    Select Case nLang
        Case bsLang_CHINESE_SIMPLIFIED, bsLang_GERMAN, bsLang_GREEK, bsLang_ENGLISH, bsLang_SPANISH, bsLang_FRENCH, bsLang_ITALIAN
            LanguageIsSupported = True
    End Select
End Function

Public Property Get LanguageIsUnicode() As Boolean
    Select Case mUILanguage
        'Case bsLang_CHINESE_SIMPLIFIED, bsLang_HEBREW, bsLang_ARABIC, bsLang_GREEK
        '    LanguageIsUnicode = True
        Case bsLang_CHINESE_SIMPLIFIED, bsLang_GREEK
            LanguageIsUnicode = True
    End Select
End Property

Private Function MAKELANGID(nPrimaryLanguage As Long, nSublanguage As Long) As Long
     MAKELANGID = nSublanguage * 1024 Or nPrimaryLanguage
End Function

Private Function IsLanguagueRightToLeft(nLCID As Long) As Boolean
    Const LOCALE_FONTSIGNATURE As Long = &H58
    Dim iLocaleSig As TLOCALESIGNATURE
    
    If GetLocaleInfoW(nLCID, LOCALE_FONTSIGNATURE, VarPtr(iLocaleSig), (LenB(iLocaleSig) / 2)) <> 0 Then
        IsLanguagueRightToLeft = (iLocaleSig.lsUsb(15) And 8) = 8 ' Unicode Subset Bitfield 123
    End If
End Function

Public Property Get UIRightToLeft() As Boolean
    If mUILanguage = bsLang_AUTO_SELECT Then SetUILanguageToWindowsUILanguage
    UIRightToLeft = mUIRightToLeft
End Property

Private Function StrFromArray(ParamArray ByteCodes() As Variant) As String
    Dim c As Long
    Dim u As Long
    Dim iByteArray() As Byte
    
    u = UBound(ByteCodes)
    ReDim iByteArray(0 To u)
    For c = 0 To u
        iByteArray(c) = ByteCodes(c)
    Next c
    StrFromArray = iByteArray
End Function

Public Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
    Dim hMod As Long, Loaded As Boolean
    
    ' Handle der DLL erhalten
    hMod = GetModuleHandle(sModule)
    If hMod = 0 Then ' Falls DLL nicht registriert ...
        hMod = LoadLibrary(sModule) ' DLL in den Speicher laden.
        If hMod Then Loaded = True
    End If
    If hMod Then
        If GetProcAddress(hMod, sFunction) Then IsFunctionExported = True
    End If
    If Loaded Then Call FreeLibrary(hMod)
End Function

Private Property Get InIDE() As Boolean
    Debug.Assert MakeTrue(InIDE)
End Property
 
Private Function MakeTrue(bValue As Boolean) As Boolean
    bValue = True
    MakeTrue = True
End Function

