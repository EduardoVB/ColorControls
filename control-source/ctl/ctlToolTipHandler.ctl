VERSION 5.00
Begin VB.UserControl ToolTipHandler 
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ctlToolTipHandler.ctx":0000
   ScaleHeight     =   2880
   ScaleWidth      =   3840
   ToolboxBitmap   =   "ctlToolTipHandler.ctx":0E14
   Begin VB.Timer tmrShowTT 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   180
      Top             =   1260
   End
   Begin VB.Timer tmrCheck 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   180
      Top             =   720
   End
End
Attribute VB_Name = "ToolTipHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hWndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function GetAncestor Lib "user32.dll" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Const GA_ROOT = 2

Private mTTData As Collection
Private WithEvents mToolTipEx As cToolTipEx
Attribute mToolTipEx.VB_VarHelpID = -1
Private mRTLForm As Boolean
Private mRTLFormChecked As Boolean
Private mFormWidth As Long
Private mRightToLeft As Boolean

Public Sub Add(ByVal nControlName As String, ByVal nToolTipText As String)
    Dim iTTD As cToolTipHandlerItem
    
    Remove nControlName
    Set iTTD = New cToolTipHandlerItem
    iTTD.ControlName = nControlName
    iTTD.ToolTipText = nToolTipText
    mTTData.Add iTTD, nControlName
End Sub
    
Public Sub Remove(ByVal nControlName As String)
    On Error Resume Next
    mTTData.Remove nControlName
    On Error GoTo 0
End Sub
    
Private Sub tmrShowTT_Timer()
    If MouseIsOverControl(CStr(tmrShowTT.Tag)) Then
        On Error Resume Next
        Set mToolTipEx = ShowToolTipEx(mTTData(CStr(tmrShowTT.Tag)).ToolTipText, , vxTTStandard, , , 0, , , , , , , , , 200, mRightToLeft)
        If Not mToolTipEx Is Nothing Then
            tmrCheck.Enabled = False
        End If
        On Error GoTo 0
    End If
    tmrShowTT.Enabled = False
    tmrShowTT.Tag = ""
End Sub

Private Sub mToolTipEx_Closed()
    Set mToolTipEx = Nothing
    tmrCheck.Enabled = True
End Sub

Private Sub UserControl_Initialize()
    Set mTTData = New Collection
End Sub

Private Sub UserControl_InitProperties()
    If Ambient.UserMode Then
        tmrCheck.Enabled = True
    End If
    mRightToLeft = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If Ambient.UserMode Then
        tmrCheck.Enabled = True
    End If
    mRightToLeft = PropBag.ReadProperty("RightToLeft", False)
End Sub

Private Sub UserControl_Resize()
    Dim iH As Long
    Dim iW As Long
    
    iH = UserControl.ScaleY(UserControl.ScaleHeight, vbTwips, vbPixels)
    iW = UserControl.ScaleX(UserControl.ScaleWidth, vbTwips, vbPixels)
    
    If (iH <> 34) Or (iW <> 34) Then
        If (iH <> 34) Then
            iH = 34
        End If
        If (iW <> 34) Then
            iW = 34
        End If
        UserControl.Size UserControl.ScaleX(iW, vbPixels, vbTwips), UserControl.ScaleY(iH, vbPixels, vbTwips)
    End If
End Sub

Private Sub UserControl_Terminate()
    Set mTTData = Nothing
    tmrCheck.Enabled = False
    tmrShowTT.Enabled = False
    If Not mToolTipEx Is Nothing Then
        On Error Resume Next
        mToolTipEx.CloseTip
        Set mToolTipEx = Nothing
        On Error GoTo 0
    End If
End Sub

Private Sub tmrCheck_Timer()
    Dim iVar As Variant
    Dim iTTD As cToolTipHandlerItem
    Dim iCP As POINTAPI
    Dim iHwnd As Long
    
    On Error GoTo TheExit
    GetCursorPos iCP
    iHwnd = GetAncestor(UserControl.ContainerHwnd, GA_ROOT)
    ScreenToClient iHwnd, iCP
    If Not TypeOf UserControl.Parent Is Form Then
        MapWindowPoints iHwnd, UserControl.Parent.hWnd, iCP, 1
    End If
    For Each iVar In mTTData
        Set iTTD = iVar
        If MouseIsOverControl(iTTD.ControlName, iCP.X, iCP.Y) Then
            If tmrShowTT.Tag = iTTD.ControlName Then
                Exit Sub
            End If
            tmrShowTT.Enabled = True
            tmrShowTT.Tag = iTTD.ControlName
            Exit For
        End If
    Next
    
TheExit:
    Err.Clear
End Sub

Private Function MouseIsOverControl(ByVal nControlName As String, Optional ByRef nMousePositionX As Long = -1, Optional ByRef nMousePositionY As Long = -1) As Boolean
    Dim iCP As POINTAPI
    Dim iLeft As Long
    Dim iTop As Long
    Dim iRight As Long
    Dim iBottom As Long
    Dim iCtl As Control
    Dim iContainer As Object
    Dim iAuxLng As Long
    Dim iHwnd As Long
    Dim iAux As Object
    Dim iRC As RECT
    Const GWL_EXSTYLE = (-20)
    Const WS_EX_LAYOUTRTL = &H400000
    
    If (nMousePositionX = -1) Or (nMousePositionY = -1) Then
        GetCursorPos iCP
        iHwnd = GetAncestor(UserControl.ContainerHwnd, GA_ROOT)
        ScreenToClient iHwnd, iCP
        If Not TypeOf UserControl.Parent Is Form Then
            MapWindowPoints iHwnd, UserControl.Parent.hWnd, iCP, 1
        Else
            mRTLForm = (GetWindowLong(iHwnd, GWL_EXSTYLE) And WS_EX_LAYOUTRTL) <> 0
        End If
        mRTLFormChecked = True
        If mRTLForm Then
            GetWindowRect iHwnd, iRC
            mFormWidth = iRC.Right - iRC.Left
        End If
    Else
        If Not mRTLFormChecked Then
            iHwnd = GetAncestor(UserControl.ContainerHwnd, GA_ROOT)
            If TypeOf UserControl.Parent Is Form Then
                mRTLForm = (GetWindowLong(iHwnd, GWL_EXSTYLE) And WS_EX_LAYOUTRTL) <> 0
            End If
            mRTLFormChecked = True
            If mRTLForm Then
                GetWindowRect iHwnd, iRC
                mFormWidth = iRC.Right - iRC.Left
            End If
        End If
        iCP.X = nMousePositionX
        iCP.Y = nMousePositionY
    End If
    
    Set iCtl = UserControl.Parent.Controls(nControlName)
    If iCtl.Visible Then
        On Error Resume Next
        iLeft = UserControl.ScaleX(iCtl.Left, ContainerScaleMode(iCtl), vbPixels)
        iTop = UserControl.ScaleY(iCtl.Top, ContainerScaleMode(iCtl), vbPixels)
        iRight = iLeft + UserControl.ScaleX(iCtl.Width, ContainerScaleMode(iCtl), vbPixels)
        iBottom = iTop + UserControl.ScaleY(iCtl.Height, ContainerScaleMode(iCtl), vbPixels)
        Set iContainer = iCtl.Container
        Do Until (iContainer Is Nothing)
            If (TypeOf iContainer Is Form) Then Exit Do
            iAuxLng = UserControl.ScaleX(iContainer.Left, ContainerScaleMode(iContainer), vbPixels)
            iLeft = iLeft + iAuxLng
            iRight = iRight + iAuxLng
            iAuxLng = UserControl.ScaleY(iContainer.Top, ContainerScaleMode(iContainer), vbPixels)
            iTop = iTop + iAuxLng
            iBottom = iBottom + iAuxLng
            Set iAux = iContainer
            Set iContainer = Nothing
            Set iContainer = iAux.Container
            Set iAux = Nothing
        Loop
        
        On Error GoTo TheExit
        If mRTLForm Then
            iCP.X = mFormWidth - iCP.X
        End If
        If iCP.X >= iLeft Then
            If iCP.X <= iRight Then
                If iCP.Y >= iTop Then
                    If iCP.Y <= iBottom Then
                        MouseIsOverControl = True
                    End If
                End If
            End If
        End If
    End If
    
TheExit:
    Err.Clear
End Function

Private Function ContainerScaleMode(nControl As Object) As ScaleModeConstants
    ContainerScaleMode = vbTwips
    On Error Resume Next
    ContainerScaleMode = nControl.Container.ScaleMode
    On Error GoTo 0
End Function


Public Property Get RightToLeft() As Boolean
    RightToLeft = mRightToLeft
End Property

Public Property Let RightToLeft(nValue As Boolean)
    If nValue <> mRightToLeft Then
        mRightToLeft = nValue
        PropertyChanged "RightToLeft"
    End If
End Property


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "RightToLeft", mRightToLeft, False
End Sub
