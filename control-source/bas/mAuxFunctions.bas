Attribute VB_Name = "mAuxFunctions"
Option Explicit

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
    szCSDVersion As String * 128
End Type

Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As Any) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Private Const MAX_PATH = 260

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Private Const TH32CS_SNAPHEAPLIST = &H1
Private Const TH32CS_SNAPPROCESS = &H2
Private Const TH32CS_SNAPTHREAD = &H4
Private Const TH32CS_SNAPMODULE = &H8
Private Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST + TH32CS_SNAPPROCESS + TH32CS_SNAPTHREAD + TH32CS_SNAPMODULE)

Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16

Private Const gstrSEP_DIR$ = "\"                         ' Directory separator character
'Private Const gstrAT$ = "@"
Private Const gstrSEP_DRIVE$ = ":"                       ' Driver separater character, e.g., C:\
Private Const gstrSEP_DIRALT$ = "/"                      ' Alternate directory separator character
'Private Const gstrSEP_EXT$ = "."                         ' Filename extension separator character
Private Const gstrSEP_URLDIR$ = "/"                      ' Separator for dividing directories in URL addresses.

Private mToolTipExCollection As New cToolTipExCollection

Public Property Get ClientProductName() As String
    Static sAlreadySet As Boolean
    Static sValue As String
    
    If Not sAlreadySet Then
        sValue = GetProductNameFromExeFile(ClientExeFile)
        sAlreadySet = True
    End If
    ClientProductName = sValue
End Property

Private Function GetProductNameFromExeFile(strFileName As String) As String
    Dim sInfo As String, lSizeof As Long
    Dim lResult As Long, intDel As Integer
    Dim lHandle As Long
    Dim intStrip As Integer
    Dim iIsNT As Boolean
    Dim iGetProductNameFromExeFile As String
    Dim iAuxStr As String
    
    If strFileName <> "" Then
        lHandle = 0
        lSizeof = GetFileVersionInfoSize(strFileName, lHandle)
        If lSizeof > 0 Then
            sInfo = String$(lSizeof, 0)
            lResult = GetFileVersionInfo(ByVal strFileName, 0&, ByVal lSizeof, ByVal sInfo)
            If lResult Then
                iIsNT = IsWindowsNT
                If iIsNT Then
                    sInfo = StrConv(sInfo, vbFromUnicode)
                End If
                intDel = InStr(sInfo, "ProductName")
                If intDel > 0 Then
                    If iIsNT Then
                        intDel = intDel + 13
                    Else
                        intDel = intDel + 12
                    End If
                    intStrip = InStr(intDel, sInfo, vbNullChar)
                    iGetProductNameFromExeFile = Trim$(Mid$(sInfo, intDel, intStrip - intDel))
                End If
                If Len(iGetProductNameFromExeFile) > 30 Or iGetProductNameFromExeFile = "" Then
                    intDel = InStr(sInfo, "Description")
                    If intDel > 0 Then
                        If iIsNT Then
                            intDel = intDel + 13
                        Else
                            intDel = intDel + 12
                        End If
                        intStrip = InStr(intDel, sInfo, vbNullChar)
                        iAuxStr = Trim$(Mid$(sInfo, intDel, intStrip - intDel))
                        If iAuxStr <> "" Then
                            iGetProductNameFromExeFile = iAuxStr
                        End If
                    End If
                End If
                If Len(iGetProductNameFromExeFile) > 40 Then
                    iGetProductNameFromExeFile = Left$(iGetProductNameFromExeFile, 40) & "..."
                End If
            End If
        End If
    End If
    iGetProductNameFromExeFile = Trim$(iGetProductNameFromExeFile)
    If iGetProductNameFromExeFile = "" Then
        iGetProductNameFromExeFile = GetFileName(strFileName)
    End If
    GetProductNameFromExeFile = iGetProductNameFromExeFile
End Function

Private Property Get ClientExeFile() As String
    Static sAlreadySet As Boolean
    Static sValue As String
    
    If Not sAlreadySet Then
        sValue = GetClientExe
        sAlreadySet = True
    End If
    ClientExeFile = sValue
End Property

Private Function GetClientExe() As String
    Dim cbNeeded2 As Long
    Dim Modules(1 To 200) As Long
    Dim ModuleName As String
    Dim nSize As Long
    Dim hProcess As Long
    Dim hSnapshot As Long, LRet As Long, P As PROCESSENTRY32
    Dim iProcessID As Long
    
    iProcessID = GetCurrentProcessId
    P.dwSize = Len(P)
    
    If IsWindowsNT Then
        ' NT
        'Get a handle to the Process
        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, iProcessID)
        'Got a Process handle
        If hProcess <> 0 Then
            'Get an array of the module handles for the specified
            'process
            LRet = EnumProcessModules(hProcess, Modules(1), 200, _
                                         cbNeeded2)
            'If the Module Array is retrieved, Get the ModuleFileName
            If LRet <> 0 Then
               ModuleName = Space$(MAX_PATH)
               nSize = 500
               LRet = GetModuleFileNameExA(hProcess, Modules(1), _
                               ModuleName, nSize)
               
               GetClientExe = Left$(ModuleName, LRet)
            End If
            CloseHandle hProcess
        End If
    Else
        'Windows 95/98
        hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, ByVal 0)
        If hSnapshot Then
            LRet = Process32First(hSnapshot, P)
            Do While LRet
                If P.th32ProcessID = iProcessID Then
                    GetClientExe = Left$(P.szExeFile, InStr(P.szExeFile, Chr$(0)) - 1)
                    Exit Do
                End If
                LRet = Process32Next(hSnapshot, P)
            Loop
            LRet = CloseHandle(hSnapshot)
        End If
    End If
End Function

Private Function IsWindowsNT() As Boolean
    Static sValue As Long
    
    If sValue = 0 Then
        Dim osinfo As OSVERSIONINFO
        Dim retvalue As Integer
        
        osinfo.dwOSVersionInfoSize = 148
        osinfo.szCSDVersion = Space$(128)
        retvalue = GetVersionEx(osinfo)
        If osinfo.dwPlatformID = 2 Then
            sValue = 2
        Else
            sValue = 1
        End If
    End If
    
    IsWindowsNT = (sValue = 2)
End Function

Private Function GetFileName(nFileFullPath As String) As String
    Dim iFileName As String
    
    SeparatePathAndFileName nFileFullPath, , iFileName
    GetFileName = iFileName
End Function

'Given a fully qualified filename, returns the path portion and the file
'   portion.
Private Sub SeparatePathAndFileName(FullPath As String, _
    Optional ByRef Path As String, _
    Optional ByRef FileName As String)

    Dim nSepPos As Long
    Dim nSepPos2 As Long
    Dim fUsingDriveSep As Boolean

    nSepPos = InStrRev(FullPath, gstrSEP_DIR)
    nSepPos2 = InStrRev(FullPath, gstrSEP_DIRALT)
    If nSepPos2 > nSepPos Then
        nSepPos = nSepPos2
    End If
    nSepPos2 = InStrRev(FullPath, gstrSEP_DRIVE)
    If nSepPos2 > nSepPos Then
        nSepPos = nSepPos2
        fUsingDriveSep = True
    End If

    If nSepPos = 0 Then
        'Separator was not found.
        Path = CurDir$
        FileName = FullPath
    Else
        If fUsingDriveSep Then
            Path = Left$(FullPath, nSepPos)
        Else
            Path = Left$(FullPath, nSepPos - 1)
        End If
        FileName = Mid$(FullPath, nSepPos + 1)
    End If
End Sub

Public Function ShowToolTipEx(nTipText As String, Optional nTitle As String, Optional nStyle As vbExBalloonTooltipStyleConstants = vxTTBalloon, Optional nCloseButton As Boolean, Optional nIcon As vbExBalloonTooltipIconConstants = vxTTNoIcon, Optional nDelayTimeSeconds As Variant, Optional nVisibleTimeSeconds As Variant, Optional nPositionX As Variant, Optional nPositionY As Variant, Optional nPositionIsRelative As Boolean, Optional nWidth As Variant, Optional nBackColor As Variant, Optional nForeColor As Variant, Optional nClosePrevious As Boolean = True, Optional nRestrictMouseMoveToTwips As Long = 300, Optional nRightToLeft As Boolean) As cToolTipEx
    Dim iPt As POINTAPI
    Dim iDelayTimeSeconds As Variant
    Dim iVisibleTimeSeconds As Variant
    Dim iPositionX As Variant
    Dim iPositionY As Variant
    Dim iWidth As Variant
    Dim iBackColor As Variant
    Dim iForeColor As Variant
    Dim iCBT As cToolTipEx
    Dim iParentHwnd As Long
    
    iParentHwnd = GetFormUnderMouseHwnd
    If (iParentHwnd = 0) Then
        iParentHwnd = GetActiveWindow
    ElseIf Not IsWindowLocal(iParentHwnd) Then
        iParentHwnd = GetActiveWindow
    End If
    If iParentHwnd = 0 Then Exit Function
    
    If nPositionIsRelative Then
        GetCursorPos iPt
        ScreenToClient iParentHwnd, iPt
        
        If Not IsMissing(nPositionX) Then
            iPositionX = (iPt.X * Screen.TwipsPerPixelX) + nPositionX
        End If
        If Not IsMissing(nPositionY) Then
            iPositionY = (iPt.Y * Screen.TwipsPerPixelY) + nPositionY
        End If
    Else
        If Not IsMissing(nPositionX) Then
            iPositionX = nPositionX
        End If
        If Not IsMissing(nPositionY) Then
            iPositionY = nPositionY
        End If
    End If
    If Not IsMissing(nDelayTimeSeconds) Then
        iDelayTimeSeconds = nDelayTimeSeconds
    End If
    If Not IsMissing(nVisibleTimeSeconds) Then
        iVisibleTimeSeconds = nVisibleTimeSeconds
    End If
    If Not IsMissing(nWidth) Then
        iWidth = nWidth
    End If
    If Not IsMissing(nBackColor) Then
        iBackColor = nBackColor
    End If
    If Not IsMissing(nForeColor) Then
        iForeColor = nForeColor
    End If
    
    For Each iCBT In mToolTipExCollection.GetCollection
        If iCBT.ParentHwnd = iParentHwnd Then
            If iCBT.TipText = nTipText Then
                If iCBT.Title = nTitle Then
                    If iCBT.BackColor = iBackColor Then
                        If iCBT.ForeColor = iForeColor Then
                            If iCBT.CloseButton = nCloseButton Then
                                If iCBT.DelayTimeSeconds = iDelayTimeSeconds Then
                                    If iCBT.VisibleTimeSeconds = iVisibleTimeSeconds Then
                                        If iCBT.Icon = nIcon Then
                                            If iCBT.PositionX = iPositionX Then
                                                If iCBT.PositionY = iPositionY Then
                                                    If iCBT.Style = nStyle Then
                                                        If iCBT.RightToLeft = nRightToLeft Then
                                                            If iCBT.Width = iWidth Then
                                                                If iCBT.RestrictMouseMoveToTwips = nRestrictMouseMoveToTwips Then
                                                                    iCBT.Reset
                                                                    Set ShowToolTipEx = iCBT
                                                                    Exit For
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    If ShowToolTipEx Is Nothing Then
        If nClosePrevious Then
            For Each iCBT In mToolTipExCollection.GetCollection
                If iCBT.ParentHwnd = iParentHwnd Then
                    iCBT.CloseTip
                End If
            Next
        End If
        Set iCBT = New cToolTipEx
        iCBT.TipText = nTipText
        iCBT.Title = nTitle
        iCBT.BackColor = iBackColor
        iCBT.ForeColor = iForeColor
        iCBT.CloseButton = nCloseButton
        iCBT.DelayTimeSeconds = iDelayTimeSeconds
        iCBT.VisibleTimeSeconds = iVisibleTimeSeconds
        iCBT.Icon = nIcon
        iCBT.PositionX = iPositionX
        iCBT.PositionY = iPositionY
        iCBT.Style = nStyle
        iCBT.Width = iWidth
        iCBT.RightToLeft = nRightToLeft
        iCBT.RestrictMouseMoveToTwips = nRestrictMouseMoveToTwips
        Set iCBT.TTCollection = mToolTipExCollection
        iCBT.Create iParentHwnd
        
        mToolTipExCollection.Add iCBT, iCBT.ToolTipHwnd
        
        Set ShowToolTipEx = iCBT
    End If
End Function

Public Function GetFormUnderMouseHwnd() As Long
    GetFormUnderMouseHwnd = WindowUnderMouseHwnd
    If Not IsWindowAForm(GetFormUnderMouseHwnd) Then
        GetFormUnderMouseHwnd = GetParentFormHwnd(GetFormUnderMouseHwnd)
        If Not IsWindowAForm(GetFormUnderMouseHwnd) Then
            GetFormUnderMouseHwnd = 0
        End If
    End If
End Function

Public Function WindowUnderMouseHwnd() As Long
    Dim iP As POINTAPI
    
    GetCursorPos iP
    WindowUnderMouseHwnd = WindowFromPoint(iP.X, iP.Y)
    
End Function

Public Function IsWindowLocal(ByVal hWnd As Long) As Boolean
    Dim idWnd As Long
    Call GetWindowThreadProcessId(hWnd, idWnd)
    IsWindowLocal = (idWnd = GetCurrentProcessId())
End Function

Public Function IsWindowAForm(nHwnd As Long) As Boolean
    Dim iClassname As String
    
    If nHwnd = 0 Then Exit Function
    
    iClassname = GetWindowClassName(nHwnd)
    IsWindowAForm = (iClassname = "ThunderRT6FormDC") Or (iClassname = "ThunderFormDC") Or (iClassname = "ThunderForm")
    
End Function

Public Function GetWindowClassName(nHwnd As Long) As String
    Dim iClassname As String
    Dim iSize As Long

    If nHwnd = 0 Then Exit Function

    iClassname = Space$(64)
    iSize = GetClassName(nHwnd, iClassname, Len(iClassname))
    GetWindowClassName = Left$(iClassname, iSize)

End Function

Public Function GetParentFormHwnd(nControlHwnd As Long) As Long
    Dim lPar As Long
    Dim iHwnd As Long
    
    iHwnd = nControlHwnd
    lPar = GetParent(iHwnd)
    While lPar <> 0
        
        If IsWindowAForm(lPar) Then
            iHwnd = lPar
        End If
        lPar = GetParent(lPar)
    Wend
    GetParentFormHwnd = iHwnd
End Function

Public Function IsValidOLE_COLOR(ByVal nColor As Long) As Boolean
    Const S_OK As Long = 0
    IsValidOLE_COLOR = (TranslateColor(nColor, 0, nColor) = S_OK)
End Function

