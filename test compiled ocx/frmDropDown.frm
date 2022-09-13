VERSION 5.00
Object = "{1E12B506-B2AA-40BC-A8FB-B6EF1CE3990F}#2.0#0"; "ClrCtrl2.ocx"
Begin VB.Form frmDropDown 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Color"
   ClientHeight    =   2196
   ClientLeft      =   5748
   ClientTop       =   2196
   ClientWidth     =   2796
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2196
   ScaleWidth      =   2796
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   272
      Left            =   2040
      TabIndex        =   2
      Top             =   1860
      Width           =   492
   End
   Begin VB.Timer tmrOpacity 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2040
      Top             =   1800
   End
   Begin ColorControls.ColorSelector ColorSelector1 
      Height          =   1680
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2352
      _ExtentX        =   4149
      _ExtentY        =   2963
      Style           =   1
      BackColor       =   16777215
   End
   Begin VB.Label lblTT 
      BackStyle       =   0  'Transparent
      Caption         =   "Escape key cancels"
      Height          =   252
      Left            =   60
      TabIndex        =   1
      Top             =   1740
      Width           =   2412
   End
End
Attribute VB_Name = "frmDropDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Const LWA_ALPHA = &H2

Public ColorSet As Boolean
Public Color As Long
Private mOpacity As Long

Public Sub SetTransparency()
    Const GWL_EXSTYLE = (-20)
    Const WS_EX_LAYERED = &H80000
    
    SetWindowLong Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    mOpacity = 0 '10
    SetLayeredWindowAttributes Me.hWnd, 0, mOpacity, LWA_ALPHA
    tmrOpacity.Enabled = True
End Sub

Private Sub cmdOK_Click()
    TakeColor
End Sub

Private Sub Form_Load()
    ColorSelector1.Color = Color
    ColorSelector1.Move Screen.TwipsPerPixelX * 3, Screen.TwipsPerPixelY * 3
    lblTT.Top = ColorSelector1.Top + ColorSelector1.Height
    Me.Move Me.Left, Me.Top, ColorSelector1.Width + 6 * Screen.TwipsPerPixelX, lblTT.Top + lblTT.Height + 6 * Screen.TwipsPerPixelY
End Sub

Private Sub Form_Resize()
    Me.Line (0, 0)-(Me.ScaleWidth - Screen.TwipsPerPixelX, Me.ScaleHeight - Screen.TwipsPerPixelY), &HC0C0C0, B
    cmdOK.Move Me.ScaleWidth - cmdOK.Width - 160, Me.ScaleHeight - cmdOK.Height - 60
End Sub

Private Sub Form_DblClick()
    TakeColor
End Sub

Private Sub ColorSelector1_DblClick()
    TakeColor
End Sub

Private Sub lblTT_DblClick()
    TakeColor
End Sub

Private Sub TakeColor()
    Color = ColorSelector1.Color
    ColorSet = True
    Unload Me
End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub


Private Sub tmrOpacity_Timer()
    mOpacity = mOpacity + 30
    If mOpacity > 255 Then
        mOpacity = 255
        tmrOpacity.Enabled = False
    End If
    SetLayeredWindowAttributes Me.hWnd, 0, mOpacity, LWA_ALPHA
End Sub
