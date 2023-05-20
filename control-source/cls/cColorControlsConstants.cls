VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColorControlsConstants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum CDYesNoAutoConstants
    cdYNNo = 0
    cdYNYes = 1
    cdYNAuto = 2
End Enum

Public Enum CDStyleConstants
    cdStyleWheel = 0
    cdStyleBox = 1
End Enum

Public Enum CDColorSystemConstants
    cdColorSystemHSV
    cdColorSystemHSL
End Enum

Public Enum CDSliderParameterConstants
    cdParameterHue = 0
    cdParameterLuminance = 1
    cdParameterValue = 1
    cdParameterSaturation = 2
    cdParameterRed = 3
    cdParameterGreen = 4
    cdParameterBlue = 5
End Enum

Public Enum CDMouseWheelScrollConstants
    cdMouseWheelNavigatingSlider
    cdMouseWheelNavigatingAxial
    cdMouseWheelNavigatingRadial
End Enum

Public Enum CDSliderOptionsAvailableConstants
    cdSliderOptionsNone
    cdSliderOptionsLumAndSat
    cdSliderOptionsHueLumAndSat
    cdSliderOptionsAll
End Enum

Public Enum CDPointerTypeConstants
    cdPointerNone = 0
    cdPointerCrosshair = 1
    cdPointerCircle = 2
End Enum
