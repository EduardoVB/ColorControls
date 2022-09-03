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
    bsLang_ENGLISH = 9
    bsLang_SPANISH = 10
    bsLang_FRENCH = 12
    
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
    cdUIT_frmColorDialog_lblRed_Caption = 4900
    cdUIT_frmColorDialog_lblGreen_Caption = 5000
    cdUIT_frmColorDialog_lblBlue_Caption = 5100
    cdUIT_frmColorDialog_lblHex_Caption = 5200
    cdUIT_frmColorDialog_lblHue_Caption = 5300
    cdUIT_frmColorDialog_lblSaturation_Caption = 5400
    cdUIT_frmColorDialog_Luminance_Caption = 5500
    cdUIT_frmColorDialog_Value_Caption = 5600
    cdUIT_frmColorDialog_lblColorSystem_Caption = 5700
    cdUIT_frmColorDialog_lblPalette_Caption = 5800
    cdUIT_frmColorDialog_cboPalette_ListItem1 = 5900
    cdUIT_frmColorDialog_cboPalette_ListItem2 = 6000
    cdUIT_frmColorDialog_cboPalette_ListItem3 = 6100
    cdUIT_frmColorDialog_cboPalette_ListItem4 = 6200
    cdUIT_frmColorDialog_InvalidColorMessage = 6300
    cdUIT_frmColorDialog_ParameterFullName_Hue = 6400
    cdUIT_frmColorDialog_ParameterFullName_Luminance = 6401
    cdUIT_frmColorDialog_ParameterFullName_Saturation = 6402
    cdUIT_frmColorDialog_ParameterFullName_Red = 6403
    cdUIT_frmColorDialog_ParameterFullName_Green = 6404
    cdUIT_frmColorDialog_ParameterFullName_Blue = 6405
    cdUIT_frmColorDialog_ParameterFullName_Value = 6406
    
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
        Case bsLang_SPANISH
            Do_SPANISH TextID, GetLocalizedString
        Case bsLang_FRENCH
            Do_FRENCH TextID, GetLocalizedString
        Case Else ' ENGLISH
            Do_ENGLISH TextID, GetLocalizedString
    End Select
    mStringsCache.Add GetLocalizedString, CStr(TextID)
End Function

Private Sub Do_ENGLISH(ByRef TextID As Long, ByRef Text As String)
    ' Microsoft teminology search: https://www.microsoft.com/en-us/language/search
    Select Case TextID
        Case cdUIT_frmColorDialog_Form_Caption
            Text = "Color selection"
        Case cdUIT_ColorSelector_chkFixedPalette_Caption
            Text = "Fixed"
        Case cdUIT_ColorSelector_chkFixedPalette_ToolTipText
            Text = "Reflects color changes visually in the palette or not"
        Case cdUIT_ColorSelector_cboSliderParameter_ToolTipText
            Text = "Parameter of the slidder control"
        Case cdUIT_ColorSelector_lblMode_Caption
            Text = "Mode:"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Hue
            Text = "Hue"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Luminance
            Text = "Lum."
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Value
            Text = "Value"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Saturation
            Text = "Sat."
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Red
            Text = "Red"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Green
            Text = "Green"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Blue
            Text = "Blue"
        Case cdUIT_frmColorDialog_lblBasicColors_Caption
            Text = "Basic colors:"
        Case cdUIT_frmColorDialog_lblRecent_Caption
            Text = "Recent:"
        Case cdUIT_frmColorDialog_Color_Caption
            Text = "color"
        Case cdUIT_frmColorDialog_ColorNew_Caption
            Text = "new"
        Case cdUIT_frmColorDialog_ColorPrevious_Caption
            Text = "previous"
        Case cdUIT_frmColorDialog_MouseWheel_ToolTipStart
            Text = "Hold the Control key down to navigate"
        Case cdUIT_frmColorDialog_MouseWheel_ToolTipEnd
            Text = "with the mouse wheel, press Shift key to go slowly"
        Case cdUIT_frmColorDialog_EyeDropper_ToolTip
            Text = "Choose a color from anywhere on the screen or press the Escape key to cancel"
        Case cdUIT_frmColorDialog_lblRed_Caption
            Text = "Red:"
        Case cdUIT_frmColorDialog_lblGreen_Caption
            Text = "Green:"
        Case cdUIT_frmColorDialog_lblBlue_Caption
            Text = "Blue:"
        Case cdUIT_frmColorDialog_lblHex_Caption
            Text = "Hex:"
        Case cdUIT_frmColorDialog_lblHue_Caption
            Text = "Hue:"
        Case cdUIT_frmColorDialog_lblSaturation_Caption
            Text = "Sat.:"
        Case cdUIT_frmColorDialog_Luminance_Caption
            Text = "Lum.:"
        Case cdUIT_frmColorDialog_Value_Caption
            Text = "Value:"
        Case cdUIT_frmColorDialog_lblColorSystem_Caption
            Text = "Mode:"
        Case cdUIT_frmColorDialog_lblPalette_Caption
            Text = "Palette:"
        Case cdUIT_frmColorDialog_cboPalette_ListItem1
            Text = "Wheel, fixed"
        Case cdUIT_frmColorDialog_cboPalette_ListItem2
            Text = "Box, fixed"
        Case cdUIT_frmColorDialog_cboPalette_ListItem3
            Text = "Wheel, dynamic"
        Case cdUIT_frmColorDialog_cboPalette_ListItem4
            Text = "Box, dynamic"
        Case cdUIT_frmColorDialog_InvalidColorMessage
            Text = "The color is not valid."
        Case cdUIT_frmColorDialog_ParameterFullName_Hue
            Text = "Hue"
        Case cdUIT_frmColorDialog_ParameterFullName_Luminance
            Text = "Luminance"
        Case cdUIT_frmColorDialog_ParameterFullName_Saturation
            Text = "Saturation"
        Case cdUIT_frmColorDialog_ParameterFullName_Red
            Text = "Red"
        Case cdUIT_frmColorDialog_ParameterFullName_Green
            Text = "Green"
        Case cdUIT_frmColorDialog_ParameterFullName_Blue
            Text = "Blue"
        Case cdUIT_frmColorDialog_ParameterFullName_Value
            Text = "Value"
    End Select
End Sub

Private Sub Do_SPANISH(ByRef TextID As Long, ByRef Text As String)
    ' Microsoft teminology search: https://www.microsoft.com/en-us/language/search
    Select Case TextID
        Case cdUIT_frmColorDialog_Form_Caption
            Text = "Selección de color"
        Case cdUIT_ColorSelector_chkFixedPalette_Caption
            Text = "Fija"
        Case cdUIT_ColorSelector_chkFixedPalette_ToolTipText
            Text = "Refleja o no los cambios de colores en la paleta"
        Case cdUIT_ColorSelector_cboSliderParameter_ToolTipText
            Text = "Parámetro del control deslizante"
        Case cdUIT_ColorSelector_lblMode_Caption
            Text = "Modo:"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Hue
            Text = "Matiz"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Luminance
            Text = "Lum."
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Value
            Text = "Valor"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Saturation
            Text = "Sat."
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Red
            Text = "Rojo"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Green
            Text = "Verde"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Blue
            Text = "Azul"
        Case cdUIT_frmColorDialog_lblBasicColors_Caption
            Text = "Colores básicos:"
        Case cdUIT_frmColorDialog_lblRecent_Caption
            Text = "Recientes:"
        Case cdUIT_frmColorDialog_Color_Caption
            Text = "color"
        Case cdUIT_frmColorDialog_ColorNew_Caption
            Text = "nuevo"
        Case cdUIT_frmColorDialog_ColorPrevious_Caption
            Text = "previo"
        Case cdUIT_frmColorDialog_MouseWheel_ToolTipStart
            Text = "Mantenga presionada la tecla Control para cambiar"
        Case cdUIT_frmColorDialog_MouseWheel_ToolTipEnd
            Text = "con la rueda del mouse, presione la tecla Mayúsculas para ir lento"
        Case cdUIT_frmColorDialog_EyeDropper_ToolTip
            Text = "Elija un color en cualquier lugar de la pantalla o presione la tecla Escape para cancelar"
        Case cdUIT_frmColorDialog_lblRed_Caption
            Text = "Rojo:"
        Case cdUIT_frmColorDialog_lblGreen_Caption
            Text = "Verde:"
        Case cdUIT_frmColorDialog_lblBlue_Caption
            Text = "Azul:"
        Case cdUIT_frmColorDialog_lblHex_Caption
            Text = "Hex:"
        Case cdUIT_frmColorDialog_lblHue_Caption
            Text = "Matiz:"
        Case cdUIT_frmColorDialog_lblSaturation_Caption
            Text = "Sat.:"
        Case cdUIT_frmColorDialog_Luminance_Caption
            Text = "Lum.:"
        Case cdUIT_frmColorDialog_Value_Caption
            Text = "Valor:"
        Case cdUIT_frmColorDialog_lblColorSystem_Caption
            Text = "Modo:"
        Case cdUIT_frmColorDialog_lblPalette_Caption
            Text = "Paleta:"
        Case cdUIT_frmColorDialog_cboPalette_ListItem1
            Text = "Rueda, fijo"
        Case cdUIT_frmColorDialog_cboPalette_ListItem2
            Text = "Caja, fijo"
        Case cdUIT_frmColorDialog_cboPalette_ListItem3
            Text = "Rueda, dinámico"
        Case cdUIT_frmColorDialog_cboPalette_ListItem4
            Text = "Caja, dinámico"
        Case cdUIT_frmColorDialog_InvalidColorMessage
            Text = "El color no es válido."
        Case cdUIT_frmColorDialog_ParameterFullName_Hue
            Text = "Matiz"
        Case cdUIT_frmColorDialog_ParameterFullName_Luminance
            Text = "Luminancia"
        Case cdUIT_frmColorDialog_ParameterFullName_Saturation
            Text = "Saturación"
        Case cdUIT_frmColorDialog_ParameterFullName_Red
            Text = "Rojo"
        Case cdUIT_frmColorDialog_ParameterFullName_Green
            Text = "Verde"
        Case cdUIT_frmColorDialog_ParameterFullName_Blue
            Text = "Azul"
        Case cdUIT_frmColorDialog_ParameterFullName_Value
            Text = "Valor"
    End Select
End Sub

Private Sub Do_FRENCH(ByRef TextID As Long, ByRef Text As String)
    Select Case TextID
        Case cdUIT_frmColorDialog_Form_Caption
            Text = "Sélection couleur"
        Case cdUIT_ColorSelector_chkFixedPalette_Caption
            Text = "Fixe"
        Case cdUIT_ColorSelector_chkFixedPalette_ToolTipText
            Text = "Reflète visuellement les changements de couleur dans la palette ou non"
        Case cdUIT_ColorSelector_cboSliderParameter_ToolTipText
            Text = "Paramètre du control ascenseur"
        Case cdUIT_ColorSelector_lblMode_Caption
            Text = "Mode:"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Hue
            Text = "Hue"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Luminance
            Text = "Lum."
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Value
            Text = "Valeur"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Saturation
            Text = "Sat."
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Red
            Text = "Rouge"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Green
            Text = "Vert"
        Case cdUIT_ColorSelector_cboSliderParameter_ListItem_Blue
            Text = "Bleu"
        Case cdUIT_frmColorDialog_lblBasicColors_Caption
            Text = "Couleurs de base:"
        Case cdUIT_frmColorDialog_lblRecent_Caption
            Text = "Récent:"
        Case cdUIT_frmColorDialog_Color_Caption
            Text = "couleur"
        Case cdUIT_frmColorDialog_ColorNew_Caption
            Text = "nouvelle"
        Case cdUIT_frmColorDialog_ColorPrevious_Caption
            Text = "précédente"
        Case cdUIT_frmColorDialog_MouseWheel_ToolTipStart
            Text = "Maintenez la touche Control appuyée pour naviguer"
        Case cdUIT_frmColorDialog_MouseWheel_ToolTipEnd
            Text = "avec la molette de la souris, pressez Shift pour defiler doucement"
        Case cdUIT_frmColorDialog_EyeDropper_ToolTip
            Text = "Choisir une couleur n'importe-où à l'écran ou presser Échap pour annuler"
        Case cdUIT_frmColorDialog_lblRed_Caption
            Text = "Rouge:"
        Case cdUIT_frmColorDialog_lblGreen_Caption
            Text = "Vert:"
        Case cdUIT_frmColorDialog_lblBlue_Caption
            Text = "Bleu:"
        Case cdUIT_frmColorDialog_lblHex_Caption
            Text = "Hex:"
        Case cdUIT_frmColorDialog_lblHue_Caption
            Text = "Hue:"
        Case cdUIT_frmColorDialog_lblSaturation_Caption
            Text = "Sat.:"
        Case cdUIT_frmColorDialog_Luminance_Caption
            Text = "Lum.:"
        Case cdUIT_frmColorDialog_Value_Caption
            Text = "Valeur:"
        Case cdUIT_frmColorDialog_lblColorSystem_Caption
            Text = "Mode:"
        Case cdUIT_frmColorDialog_lblPalette_Caption
            Text = "Palette:"
        Case cdUIT_frmColorDialog_cboPalette_ListItem1
            Text = "Molette, fixe"
        Case cdUIT_frmColorDialog_cboPalette_ListItem2
            Text = "Boite, fixe"
        Case cdUIT_frmColorDialog_cboPalette_ListItem3
            Text = "Molette, dynamique"
        Case cdUIT_frmColorDialog_cboPalette_ListItem4
            Text = "Boite, dynamique"
        Case cdUIT_frmColorDialog_InvalidColorMessage
            Text = "La couleur n'est pas valide."
        Case cdUIT_frmColorDialog_ParameterFullName_Hue
            Text = "Hue"
        Case cdUIT_frmColorDialog_ParameterFullName_Luminance
            Text = "Luminance"
        Case cdUIT_frmColorDialog_ParameterFullName_Saturation
            Text = "Saturation"
        Case cdUIT_frmColorDialog_ParameterFullName_Red
            Text = "Rouge"
        Case cdUIT_frmColorDialog_ParameterFullName_Green
            Text = "Vert"
        Case cdUIT_frmColorDialog_ParameterFullName_Blue
            Text = "Bleu"
        Case cdUIT_frmColorDialog_ParameterFullName_Value
            Text = "Valeur"
    End Select
End Sub

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
    mUILanguage = bsLang_FRENCH 'bsLang_ENGLISH
#End If
End Sub

Private Function LanguageIsSupported(nLang As Long) As Boolean
    Dim c As Long
    
    Select Case nLang
        Case bsLang_ENGLISH, bsLang_SPANISH, bsLang_FRENCH
            LanguageIsSupported = True
    End Select
End Function

Public Property Get LanguageIsUnicode() As Boolean
    Select Case mUILanguage
        'Case bsLang_CHINESE_SIMPLIFIED, bsLang_HEBREW, bsLang_ARABIC, bsLang_GREEK
        '    LanguageIsUnicode = True
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

