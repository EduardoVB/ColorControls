Attribute VB_Name = "mAuxFunctions"
Option Explicit

Private Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

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
        If osinfo.dwPlatformId = 2 Then
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

