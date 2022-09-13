VERSION 5.00
Begin VB.UserControl EyeDropper 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   InvisibleAtRuntime=   -1  'True
   PropertyPages   =   "ctlEyeDropper.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ctlEyeDropper.ctx":0014
   Begin VB.Image imgAux 
      Height          =   384
      Left            =   180
      Picture         =   "ctlEyeDropper.ctx":0326
      Top             =   1080
      Width           =   384
   End
End
Attribute VB_Name = "EyeDropper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Control to choose a color from the screen."
Option Explicit

Public Event Click()
Attribute Click.VB_Description = "Occurs when the user chooses a color from the screen by clicking."

Private mColor As Long
Private mAColorWasChosen As Boolean
Private mEyeDropping As Boolean

Private Sub UserControl_Initialize()
    mColor = -1
    UserControl.PaintPicture imgAux.Picture, 2, 2
End Sub

Private Sub UserControl_Resize()
    Dim iH As Long
    Dim iW As Long
    Const cSize As Long = 36
    
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

Public Function StartDropper() As Boolean
Attribute StartDropper.VB_Description = "Starts the eye dropper."
    mEyeDropping = True
    mAColorWasChosen = False
    frmEyeDropper.Width = Screen.TwipsPerPixelX * 10
    frmEyeDropper.Height = Screen.TwipsPerPixelY * 10
    Set frmEyeDropper.EyeDropperControl = Me
    frmEyeDropper.Show vbModal
    mEyeDropping = False
    StartDropper = mAColorWasChosen
    On Error Resume Next
    UserControl.Parent.SetFocus
End Function

Friend Sub SetUnderMouseColor(ByVal nColor As Long)
    Dim iEdn As IEyeDropperNotification
    
    mColor = nColor
    
    On Error Resume Next
    Set iEdn = UserControl.Parent
    On Error GoTo 0
    If Not iEdn Is Nothing Then
        iEdn.ColorUnderMouseChange mColor
    End If
End Sub
    
Friend Sub SetColor(ByVal nColor As Long)
    mColor = nColor
    ClosefrmEyeDropper
    mAColorWasChosen = True
    RaiseEvent Click
End Sub

Friend Sub FormClosed()
    ClosefrmEyeDropper
End Sub

Private Sub ClosefrmEyeDropper()
    If IsFormLoaded(frmEyeDropper) Then
        Unload frmEyeDropper
    End If
    Set frmEyeDropper = Nothing
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

Public Property Get Color() As OLE_COLOR
Attribute Color.VB_Description = "Returns/sets the current color."
Attribute Color.VB_UserMemId = 0
Attribute Color.VB_MemberFlags = "200"
    If mColor = -1 Then
        If mEyeDropping Then
            mColor = frmEyeDropper.GetCurrentColor
        End If
    End If
    If mColor <> -1 Then
        Color = mColor
    End If
End Property

Public Property Get Canceled() As Boolean
Attribute Canceled.VB_Description = "Returns True if the action was canceled."
    Canceled = Not mAColorWasChosen
End Property

