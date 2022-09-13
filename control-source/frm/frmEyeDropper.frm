VERSION 5.00
Begin VB.Form frmEyeDropper 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1908
   ClientLeft      =   2208
   ClientTop       =   2160
   ClientWidth     =   2016
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
   MouseIcon       =   "frmEyeDropper.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   1908
   ScaleWidth      =   2016
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrRestoreWP 
      Interval        =   1000
      Left            =   420
      Top             =   1380
   End
   Begin VB.Timer tmrColorUnderMouse 
      Interval        =   100
      Left            =   420
      Top             =   960
   End
End
Attribute VB_Name = "frmEyeDropper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST = -1
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOMOVE As Long = &H2

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal HDC As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private mFormHwnd As Long
Private mEyeDropperControl As EyeDropper
Private mX As Long
Private mY As Long

Public Property Set EyeDropperControl(nObj As Object)
    Set mEyeDropperControl = nObj
End Property

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Const GWL_EXSTYLE = (-20)
    Const WS_EX_LAYERED = &H80000
    Const LWA_ALPHA = &H2
    
    mFormHwnd = Me.hWnd
    SetWindowPos mFormHwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    HookMouse
    SetWindowLong mFormHwnd, GWL_EXSTYLE, GetWindowLong(mFormHwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes mFormHwnd, 0, 1, LWA_ALPHA
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not mEyeDropperControl Is Nothing Then
        mEyeDropperControl.FormClosed
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnHookMouse
    Set mEyeDropperControl = Nothing
End Sub

Public Sub MousePos(ByVal X As Long, ByVal Y As Long)
    MoveWindow mFormHwnd, X - 150, Y - 150, 300, 300, 0
    Me.Refresh
    mX = X
    mY = Y
End Sub

Public Sub MouseDown(ByVal X As Long, ByVal Y As Long)
    Me.Hide
    If Not mEyeDropperControl Is Nothing Then
        mEyeDropperControl.SetColor GetCurrentColor
    End If
End Sub

Private Sub tmrColorUnderMouse_Timer()
    If Not mEyeDropperControl Is Nothing Then
        mEyeDropperControl.SetUnderMouseColor GetCurrentColor
    End If
End Sub

Friend Function GetCurrentColor() As Long
    Dim iDC As Long
    
    iDC = GetDC(0)
    GetCurrentColor = GetPixel(iDC, mX, mY)
    ReleaseDC 0, iDC
    If GetCurrentColor = -1 Then GetCurrentColor = 0
End Function
    
Private Sub tmrRestoreWP_Timer()
    SetWindowPos mFormHwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub
