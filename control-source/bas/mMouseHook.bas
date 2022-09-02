Attribute VB_Name = "mMouseHook"
Option Explicit

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type MSLLHOOKSTRUCT
    pt As POINTAPI
    mouseData As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Integer, lParam As Any) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Private mMouseHook As Long

Public Sub HookMouse()
    Const WH_MOUSE_LL = &HE&
    mMouseHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseProc, App.hInstance, 0)
End Sub
    
Public Sub UnHookMouse()
    If mMouseHook Then
        UnhookWindowsHookEx (mMouseHook)
        mMouseHook = 0
    End If
End Sub

Private Function MouseProc(ByVal uCode As Long, ByVal wParam As Long, lParam As MSLLHOOKSTRUCT) As Long
    Const HC_ACTION = 0
    Const WM_MOUSEMOVE As Long = &H200
    Const WM_LBUTTONDOWN As Long = &H201
    
    If uCode = HC_ACTION Then
        Select Case wParam
            Case WM_MOUSEMOVE
                frmEyeDropper.MousePos lParam.pt.X, lParam.pt.Y
            Case WM_LBUTTONDOWN
                frmEyeDropper.MouseDown lParam.pt.X, lParam.pt.Y
        End Select
    End If
End Function
