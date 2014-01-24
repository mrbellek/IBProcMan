VERSION 5.00
Begin VB.Form frmMain 
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   8610
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraProcessManager 
      Caption         =   "Itty Bitty Process Manager"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      Begin VB.CommandButton cmdProcManRefresh 
         Caption         =   "Re&fresh"
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton cmdProcManBack 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   4440
         TabIndex        =   6
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton cmdProcManRun 
         Caption         =   "&Run..."
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton cmdProcManKill 
         Caption         =   "&Kill process"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   3360
         Width           =   1215
      End
      Begin VB.ListBox lstProcessManager 
         Height          =   1185
         IntegralHeight  =   0   'False
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   3
         Top             =   600
         Width           =   8175
      End
      Begin VB.CheckBox chkProcManShowDLLs 
         Alignment       =   1  'Right Justify
         Caption         =   "Show &DLLs"
         Height          =   255
         Left            =   4440
         TabIndex        =   2
         Top             =   330
         Width           =   1215
      End
      Begin VB.ListBox lstProcManDLLs 
         Height          =   1140
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   1
         Top             =   2040
         Visible         =   0   'False
         Width           =   8175
      End
      Begin VB.Label lblConfigInfo 
         AutoSize        =   -1  'True
         Caption         =   "Loaded DLL libraries by selected process:"
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   10
         Top             =   1800
         Width           =   2955
      End
      Begin VB.Image imgProcManCopy 
         Height          =   240
         Left            =   3720
         Picture         =   "frmMain.frx":1CFA
         ToolTipText     =   "Copy process list to clipboard"
         Top             =   330
         Width           =   240
      End
      Begin VB.Label lblProcManDblClick 
         Caption         =   "Double-click a file to view its properties."
         Height          =   390
         Left            =   5760
         TabIndex        =   9
         Top             =   3330
         Width           =   1575
      End
      Begin VB.Label lblConfigInfo 
         AutoSize        =   -1  'True
         Caption         =   "Running processes:"
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1410
      End
      Begin VB.Image imgProcManSave 
         Height          =   240
         Left            =   4080
         Picture         =   "frmMain.frx":1E44
         ToolTipText     =   "Save process list to file.."
         Top             =   330
         Width           =   240
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Kid Radd rules!"
      Visible         =   0   'False
      Begin VB.Menu mnuMainKill 
         Caption         =   "Kill process(es)"
      End
      Begin VB.Menu mnuMainStr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMainCopy 
         Caption         =   "Copy list to clipboard"
      End
      Begin VB.Menu mnuMainSave 
         Caption         =   "Save list to disk..."
      End
      Begin VB.Menu mnuMainStr2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMainProps 
         Caption         =   "File properties"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'v1.00 - original release, later added copy to clipboard button
'v1.01 - added label for dlls, keyboard shortcuts
'v1.01.1 - fixed crash bug in form_resize, added version number to frame
'v1.02 - added PID numbers to process list
'v1.03 - fixed killing multiple processes (it works now.. typos suck)
'        also added PauseProcess to the killing subs :D (excludes self)
'        added right-click menu to listboxes
'        fixed a crash bug with the CompanyName property of RAdmin.exe
'v1.04 - processes that fail to be killed are now resumed again
'--
'v1.05 - dll list is updated when browsing process list with keyboard
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Module32First Lib "kernel32" (ByVal hSnapshot As Long, uProcess As MODULEENTRY32) As Long
Private Declare Function Module32Next Lib "kernel32" (ByVal hSnapshot As Long, uProcess As MODULEENTRY32) As Long
Private Declare Function Thread32First Lib "kernel32" (ByVal hSnapshot As Long, uThread As THREADENTRY32) As Long
Private Declare Function Thread32Next Lib "kernel32" (ByVal hSnapshot As Long, ByRef ThreadEntry As THREADENTRY32) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function SuspendThread Lib "kernel32" (ByVal hThread As Long) As Long
Private Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
Private Declare Function OpenThread Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwThreadId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Private Declare Function EnumProcesses Lib "PSAPI.DLL" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long

Private Declare Function SHRunDialog Lib "shell32" Alias "#61" (ByVal hOwner As Long, ByVal Unknown1 As Long, ByVal Unknown2 As Long, ByVal szTitle As String, ByVal szPrompt As String, ByVal uFlags As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function ShellExecuteEx Lib "shell32.dll" (SEI As SHELLEXECUTEINFO) As Long
Private Declare Function SHFileExists Lib "shell32" Alias "#45" (ByVal szPath As String) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersionl As Integer
    dwStrucVersionh As Integer
    dwFileVersionMSl As Integer
    dwFileVersionMSh As Integer
    dwFileVersionLSl As Integer
    dwFileVersionLSh As Integer
    dwProductVersionMSl As Integer
    dwProductVersionMSh As Integer
    dwProductVersionLSl As Integer
    dwProductVersionLSh As Integer
    dwFileFlagsMask As Long
    dwFileFlags As Long
    dwFileOS As Long
    dwFileType As Long
    dwFileSubtype As Long
    dwFileDateMS As Long
    dwFileDateLS As Long
End Type

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
    szExeFile As String * 260
End Type

Private Type MODULEENTRY32
    dwSize As Long
    th32ModuleID As Long
    th32ProcessID As Long
    GlblcntUsage As Long
    ProccntUsage As Long
    modBaseAddr As Long
    modBaseSize As Long
    hModule As Long
    szModule As String * 256
    szExePath As String * 260
End Type

Private Type THREADENTRY32
    dwSize As Long
    dwRefCount As Long
    th32ThreadID As Long
    th32ProcessID As Long
    dwBasePriority As Long
    dwCurrentPriority As Long
    dwFlags As Long
End Type

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_OVERWRITEPROMPT = &H2

Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Const TH32CS_SNAPPROCESS = &H2
Private Const TH32CS_SNAPMODULE = &H8
Private Const TH32CS_SNAPTHREAD = &H4
Private Const PROCESS_TERMINATE = &H1
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
Private Const THREAD_SUSPEND_RESUME = &H2

Private Const LB_SETTABSTOPS = &H192

Private bIsWinNT As Boolean, bIsWinME As Boolean
Private sWinDir$, sWinVersion$

Private Sub RefreshProcessList(objList As ListBox)
    Dim hSnap&, uPE32 As PROCESSENTRY32, i&
    Dim sExeFile$, hProcess&

    hSnap = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0)
    
    uPE32.dwSize = Len(uPE32)
    If ProcessFirst(hSnap, uPE32) = 0 Then
        CloseHandle hSnap
        Exit Sub
    End If
    
    objList.Clear
    Do
        sExeFile = TrimNull(uPE32.szExeFile)
        objList.AddItem uPE32.th32ProcessID & vbTab & sExeFile
    Loop Until ProcessNext(hSnap, uPE32) = 0
    CloseHandle hSnap
End Sub

Private Sub RefreshProcessListNT(objList As ListBox)
    Dim lProcesses&(1 To 1024), lNeeded&, lNumProcesses&
    Dim hProc&, sProcessName$, lModules&(1 To 1024), i%
    On Error Resume Next

    If EnumProcesses(lProcesses(1), CLng(1024) * 4, lNeeded) = 0 Then
        'no PSAPI.DLL file or wrong version
        Exit Sub
    End If
    
    objList.Clear
    lNumProcesses = lNeeded / 4
    For i = 1 To lNumProcesses
        'hProc = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ Or PROCESS_TERMINATE, 0, lProcesses(i))
        hProc = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lProcesses(i))
        If hProc <> 0 Then
            'Openprocess can return 0 but we ignore this since
            'system processes are somehow protected, further
            'processes CAN be opened.... silly windows
        
            lNeeded = 0
            sProcessName = String(260, 0)
            If EnumProcessModules(hProc, lModules(1), CLng(1024) * 4, lNeeded) <> 0 Then
                GetModuleFileNameExA hProc, lModules(1), sProcessName, Len(sProcessName)
                sProcessName = TrimNull(sProcessName)
                If sProcessName <> vbNullString Then
                    If Left(sProcessName, 1) = "\" Then sProcessName = Mid(sProcessName, 2)
                    If Left(sProcessName, 3) = "??\" Then sProcessName = Mid(sProcessName, 4)
                    If InStr(1, sProcessName, "%Systemroot%", vbTextCompare) > 0 Then sProcessName = Replace(sProcessName, "%Systemroot%", sWinDir, , , vbTextCompare)
                    If InStr(1, sProcessName, "Systemroot", vbTextCompare) > 0 Then sProcessName = Replace(sProcessName, "Systemroot", sWinDir, , , vbTextCompare)
                    
                    objList.AddItem lProcesses(i) & vbTab & sProcessName
                End If
            End If
        End If
        CloseHandle hProc
    Next i
End Sub

Private Sub KillProcess(lPID&)
    Dim hProcess&
    If lPID = 0 Then Exit Sub
    hProcess = OpenProcess(PROCESS_TERMINATE, 0, lPID)
    If hProcess = 0 Then
        MsgBox "The selected process could not be killed." & _
               " It may have already closed, or it may be protected by Windows.", vbCritical
    Else
        If TerminateProcess(hProcess, 0) = 0 Then
            MsgBox "The selected process could not be killed." & _
                   " It may be protected by Windows.", vbCritical
        Else
            CloseHandle hProcess
            DoEvents
        End If
    End If
End Sub

Private Sub KillProcessNT(lPID&)
    Dim hProc&
    On Error Resume Next
    If lPID = 0 Then Exit Sub
    hProc = OpenProcess(PROCESS_TERMINATE, 0, lPID)
    If hProc <> 0 Then
        If TerminateProcess(hProc, 0) = 0 Then
            MsgBox "The selected process could not be killed." & _
                   " It may be protected by Windows.", vbCritical
        Else
            CloseHandle hProc
            DoEvents
        End If
    Else
        MsgBox "The selected process could not be killed." & _
               " It may have already closed, or it may be protected by Windows." & vbCrLf & vbCrLf & _
               "This process might be a service, which you can " & _
               "stop from the Services applet in Admin Tools." & vbCrLf & _
               "(To load this window, click Start, Run and enter 'services.msc')", vbCritical
    End If
End Sub

Private Sub RefreshDLLListNT(lPID&, objList As ListBox)
    Dim lProcesses&(1 To 1024), lNeeded&, lNumProcesses&
    Dim hProc&, sProcessName$, lModules&(1 To 1024), i%
    Dim sModuleName$, j&
    On Error Resume Next
    objList.Clear
    
    hProc = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lPID)
    If hProc <> 0 Then
        lNeeded = 0
        If EnumProcessModules(hProc, lModules(1), CLng(1024) * 4, lNeeded) <> 0 Then
            For j = 2 To 1024
                If lModules(j) = 0 Then Exit For
                sModuleName = String(260, 0)
                GetModuleFileNameExA hProc, lModules(j), sModuleName, Len(sModuleName)
                sModuleName = TrimNull(sModuleName)
                If sModuleName <> vbNullString And _
                   sModuleName <> "?" Then
                    objList.AddItem sModuleName
                End If
            Next j
        End If
        CloseHandle hProc
    End If
End Sub

Private Sub RefreshDLLList(lPID&, objList As ListBox)
    Dim hSnap&, uME32 As MODULEENTRY32
    Dim sDllFile$
    objList.Clear
    If lPID = 0 Then Exit Sub
    
    hSnap = CreateToolhelpSnapshot(TH32CS_SNAPMODULE, lPID)
    uME32.dwSize = Len(uME32)
    If Module32First(hSnap, uME32) = 0 Then
        CloseHandle hSnap
        Exit Sub
    End If

    Do
        sDllFile = TrimNull(uME32.szExePath)
        objList.AddItem sDllFile
    Loop Until Module32Next(hSnap, uME32) = 0
    CloseHandle hSnap
End Sub

Private Sub PauseProcess(lPID&, Optional bPauseOrResume As Boolean = True)
    Dim hSnap&, uTE32 As THREADENTRY32, hThread&
    If Not bIsWinNT And Not bIsWinME Then Exit Sub
    If lPID = GetCurrentProcessId Then Exit Sub
    
    hSnap = CreateToolhelpSnapshot(TH32CS_SNAPTHREAD, lPID)
    If hSnap = -1 Then Exit Sub
    
    uTE32.dwSize = Len(uTE32)
    If Thread32First(hSnap, uTE32) = 0 Then
        CloseHandle hSnap
        Exit Sub
    End If
    
    Do
        If uTE32.th32ProcessID = lPID Then
            hThread = OpenThread(THREAD_SUSPEND_RESUME, False, uTE32.th32ThreadID)
            If bPauseOrResume Then
                SuspendThread hThread
            Else
                ResumeThread hThread
            End If
            CloseHandle hThread
        End If
    Loop Until Thread32Next(hSnap, uTE32) = 0
    CloseHandle hSnap
End Sub

Private Sub SaveProcessList(objProcess As ListBox, objDLL As ListBox, Optional bDoDLLs As Boolean = False)
    Dim sFileName$, i%, sProcess$, sModule$
    sFileName = CmnDlgSaveFile("Save process list to file..", "Text files (*.txt)|*.txt|All files (*.*)|*.*", "processlist.txt")
    If sFileName = vbNullString Then Exit Sub
    
    On Error Resume Next
    Open sFileName For Output As #1
        Print #1, "Process list saved on " & Format(Time, "Long Time") & ", on " & Format(Date, "Short Date")
        Print #1, "Platform: " & sWinVersion & vbCrLf
        Print #1, "[pid]" & vbTab & "[full path to filename]" & vbTab & vbTab & "[file version]" & vbTab & "[company name]"
        For i = 0 To objProcess.ListCount - 1
            sProcess = objProcess.List(i)
            Print #1, sProcess & vbTab & vbTab & _
                      GetFilePropVersion(Mid(sProcess, InStr(sProcess, vbTab) + 1)) & vbTab & _
                      GetFilePropCompany(Mid(sProcess, InStr(sProcess, vbTab) + 1))
        Next i
    
        If bDoDLLs Then
            sProcess = objProcess.List(objProcess.ListIndex)
            sProcess = Mid(sProcess, InStr(sProcess, vbTab) + 1)
            Print #1, vbCrLf & vbCrLf & "DLLs loaded by process " & sProcess & ":" & vbCrLf
            Print #1, "[full path to filename]" & vbTab & vbTab & "[file version]" & vbTab & "[company name]"
            For i = 0 To objDLL.ListCount - 1
                sModule = objDLL.List(i)
                Print #1, sModule & vbTab & vbTab & GetFilePropVersion(sModule) & vbTab & GetFilePropCompany(sModule)
            Next i
        End If
    
    Close #1
    
    ShellExecute 0, "open", sFileName, vbNullString, vbNullString, 1
End Sub

Private Sub CopyProcessList(objProcess As ListBox, objDLL As ListBox, Optional bDoDLLs As Boolean = False)
    Dim i%, sList$, sProcess$, sModule$
    
    On Error Resume Next
    sList = "Process list saved on " & Format(Time, "Long Time") & ", on " & Format(Date, "Short Date") & vbCrLf & _
            "Platform: " & sWinVersion & vbCrLf & vbCrLf & _
            "[pid]" & vbTab & "[full path to filename]" & vbTab & vbTab & "[file version]" & vbTab & "[company name]" & vbCrLf
    For i = 0 To objProcess.ListCount - 1
        sProcess = objProcess.List(i)
        sList = sList & sProcess & vbTab & vbTab & _
                GetFilePropVersion(Mid(sProcess, InStr(sProcess, vbTab) + 1)) & vbTab & _
                GetFilePropCompany(Mid(sProcess, InStr(sProcess, vbTab) + 1)) & vbCrLf
    Next i
    
    If bDoDLLs Then
        sProcess = objProcess.List(objProcess.ListIndex)
        sProcess = Mid(sProcess, InStr(sProcess, vbTab) + 1)
        sList = sList & vbCrLf & vbCrLf & "DLLs loaded by process " & sProcess & ":" & vbCrLf & vbCrLf & _
                "[full path to filename]" & vbTab & vbTab & "[file version]" & vbTab & "[company name]" & vbCrLf
        For i = 0 To objDLL.ListCount - 1
            sModule = objDLL.List(i)
            sList = sList & sModule & vbTab & vbTab & GetFilePropVersion(sModule) & vbTab & GetFilePropCompany(sModule) & vbCrLf
        Next i
    End If
    
    Clipboard.Clear
    Clipboard.SetText sList
    If bDoDLLs Then
        MsgBox "The process list and dll list have been copied to your clipboard.", vbInformation
    Else
        MsgBox "The process list has been copied to your clipboard.", vbInformation
    End If
End Sub

Private Function GetFilePropVersion$(sFileName$)
    Dim hData&, lDataLen&, uBuf() As Byte, uCodePage(0 To 3) As Byte
    Dim sCodePage$, sCompanyName$, uVFFI As VS_FIXEDFILEINFO, sVersion$
    If Not FileExists(sFileName) Then Exit Function
    
    lDataLen = GetFileVersionInfoSize(sFileName, ByVal 0)
    If lDataLen = 0 Then Exit Function
        
    ReDim uBuf(0 To lDataLen - 1)
    GetFileVersionInfo sFileName, 0, lDataLen, uBuf(0)
    VerQueryValue uBuf(0), "\", hData, lDataLen
    CopyMemory uVFFI, ByVal hData, Len(uVFFI)
    
    With uVFFI
        sVersion = .dwFileVersionMSh & "." & _
                   .dwFileVersionMSl & "." & _
                   .dwFileVersionLSh & "." & _
                   .dwFileVersionLSl
    End With
    GetFilePropVersion = sVersion
    DoEvents
End Function

Private Function GetFilePropCompany$(sFileName$)
    Dim hData&, lDataLen&, uBuf() As Byte, uCodePage(0 To 3) As Byte
    Dim sCodePage$, sCompanyName$
    If Not FileExists(sFileName) Then Exit Function
    
    lDataLen = GetFileVersionInfoSize(sFileName, ByVal 0)
    If lDataLen = 0 Then Exit Function
        
    ReDim uBuf(0 To lDataLen - 1)
    GetFileVersionInfo sFileName, 0, lDataLen, uBuf(0)
    VerQueryValue uBuf(0), "\VarFileInfo\Translation", hData, lDataLen
    If lDataLen = 0 Then Exit Function
    
    CopyMemory uCodePage(0), ByVal hData, 4
    sCodePage = Format(Hex(uCodePage(1)), "00") & _
                Format(Hex(uCodePage(0)), "00") & _
                Format(Hex(uCodePage(3)), "00") & _
                Format(Hex(uCodePage(2)), "00")
        
    'get CompanyName string
    If VerQueryValue(uBuf(0), "\StringFileInfo\" & sCodePage & "\CompanyName", hData, lDataLen) = 0 Then Exit Function
    sCompanyName = String(lDataLen, 0)
    lstrcpy sCompanyName, hData
    GetFilePropCompany = TrimNull(sCompanyName)
    DoEvents
End Function

Private Function CmnDlgSaveFile(sTitle$, sFilter$, Optional sDefFile$)
    Dim uOFN As OPENFILENAME, sFile$
    On Error Resume Next
    
    sFile = sDefFile & String(256 - Len(sDefFile), 0)
    With uOFN
        .lStructSize = Len(uOFN)
        If InStr(sFilter, "|") > 0 Then sFilter = Replace(sFilter, "|", Chr(0))
        If Right(sFilter, 2) <> Chr(0) & Chr(0) Then sFilter = sFilter & Chr(0) & Chr(0)
        .lpstrFilter = sFilter
        .lpstrFile = sFile
        .lpstrTitle = sTitle
        .nMaxFile = 256
        .flags = OFN_HIDEREADONLY Or OFN_NONETWORKBUTTON Or OFN_OVERWRITEPROMPT
    End With
    If GetSaveFileName(uOFN) = 0 Then Exit Function
    sFile = Left(uOFN.lpstrFile, InStr(uOFN.lpstrFile, Chr(0)) - 1)
    CmnDlgSaveFile = sFile
End Function

Private Function TrimNull$(s$)
    If InStr(s, Chr(0)) = 0 Then
        TrimNull = s
    Else
        TrimNull = Left(s, InStr(s, Chr(0)) - 1)
    End If
End Function

Private Sub GetWinVersion()
    Dim uOVI As OSVERSIONINFO, sCSD$
    On Error Resume Next
    
    uOVI.dwOSVersionInfoSize = Len(uOVI)
    GetVersionEx uOVI
    With uOVI
        Select Case .dwPlatformId
            Case VER_PLATFORM_WIN32s: End
            Case VER_PLATFORM_WIN32_WINDOWS
                sWinVersion = "Win9x "
            Case VER_PLATFORM_WIN32_NT
                bIsWinNT = True
                sWinVersion = "WinNT "
        End Select
        sWinVersion = sWinVersion & CStr(.dwMajorVersion) & "." & _
                      String(2 - Len(CStr(.dwMinorVersion)), "0") & _
                      CStr(.dwMinorVersion) & "." & _
                      CStr(.dwBuildNumber And &HFFF)
        If Not bIsWinNT And _
           .dwMajorVersion = 4 And _
           .dwMinorVersion = 90 Then bIsWinME = True
        sCSD = Trim(TrimNull(.szCSDVersion))
        sCSD = Replace(sCSD, "Service Pack ", "SP", , , vbTextCompare)
        sCSD = Replace(sCSD, "ServicePack ", "SP", , , vbTextCompare)
        sCSD = Replace(sCSD, "Service Pack", "SP", , , vbTextCompare)
        sCSD = Replace(sCSD, "ServicePack", "SP", , , vbTextCompare)
        sWinVersion = sWinVersion & " " & sCSD
    End With
    
    sWinDir = String(260, 0)
    sWinDir = Left(sWinDir, GetWindowsDirectory(sWinDir, Len(sWinDir)))
End Sub

Private Function FileExists(sFile$) As Boolean
    On Error Resume Next
    Dim sDummy$
    sDummy = Replace(sFile, "\\", "\")
    If bIsWinNT Then
        'FileExists = IIf(Dir(sDummy, vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString, True, False)
        FileExists = IIf(SHFileExists(StrConv(sDummy, vbUnicode)) = 1, True, False)
    Else
        FileExists = IIf(SHFileExists(sDummy) = 1, True, False)
    End If
End Function

Private Sub ShowFileProperties(sFile$)
    Dim uSEI As SHELLEXECUTEINFO
    With uSEI
        .cbSize = Len(uSEI)
        .fMask = SEE_MASK_INVOKEIDLIST Or SEE_MASK_NOCLOSEPROCESS
        .hWnd = frmMain.hWnd
        .lpFile = sFile
        .lpVerb = "properties"
        .nShow = 1
    End With
    ShellExecuteEx uSEI
End Sub

Private Sub SetListBoxColumns(objListBox As ListBox)
    Dim lTabStop&(1)
    On Error GoTo 0:
    lTabStop(0) = 70
    lTabStop(1) = 0
    SendMessage objListBox.hWnd, LB_SETTABSTOPS, UBound(lTabStop), lTabStop(0)
End Sub

Private Sub chkProcManShowDLLs_Click()
    lstProcManDLLs.Visible = CBool(chkProcManShowDLLs.Value)
    On Error Resume Next
    'lstProcessManager.ListIndex = 0
    lstProcessManager_MouseUp 1, 0, 0, 0
    lstProcessManager.SetFocus
    Form_Resize
End Sub

Private Sub cmdProcManBack_Click()
    End
End Sub

Private Sub cmdProcManKill_Click()
    Dim sMsg$, i%
    sMsg = "Are you sure you want to close the selected processes?" & vbCrLf
    For i = 0 To lstProcessManager.ListCount - 1
        If lstProcessManager.Selected(i) Then
            sMsg = sMsg & lstProcessManager.List(i) & vbCrLf
        End If
    Next i
    sMsg = sMsg & vbCrLf & "Any unsaved data in it will be lost."
    If MsgBox(sMsg, vbExclamation + vbYesNo) = vbNo Then Exit Sub
    
    Dim s$
    'new since 1.03 - pauseprocess!
    For i = 0 To lstProcessManager.ListCount - 1
        If lstProcessManager.Selected(i) Then
            s = lstProcessManager.List(i)
            s = Left(s, InStr(s, vbTab) - 1)
            PauseProcess CLng(s)
        End If
    Next i
    For i = 0 To lstProcessManager.ListCount - 1
        If lstProcessManager.Selected(i) Then
            s = lstProcessManager.List(i)
            s = Left(s, InStr(s, vbTab) - 1)
            If Not bIsWinNT Then
                KillProcess CLng(s)
            Else
                KillProcessNT CLng(s)
            End If
        End If
    Next i
    Sleep 1000
    'resume any processes still alive
    For i = 0 To lstProcessManager.ListCount - 1
        If lstProcessManager.Selected(i) Then
            s = lstProcessManager.List(i)
            s = Left(s, InStr(s, vbTab) - 1)
            PauseProcess CLng(s), False
        End If
    Next i
    
    cmdProcManRefresh_Click
End Sub

Private Sub cmdProcManRefresh_Click()
    lstProcessManager.Clear
    If Not bIsWinNT Then
        RefreshProcessList lstProcessManager
    Else
        RefreshProcessListNT lstProcessManager
        lstProcessManager.ListIndex = 0
        If lstProcManDLLs.Visible Then
            Dim s$
            s = lstProcessManager.List(lstProcessManager.ListIndex)
            s = Left(s, InStr(s, vbTab) - 1)
            If Not bIsWinNT Then
                RefreshDLLList CLng(s), lstProcManDLLs
            Else
                RefreshDLLListNT CLng(s), lstProcManDLLs
            End If
        End If
    End If
    lblConfigInfo(8).Caption = "Running processes: (" & lstProcessManager.ListCount & ")"
    lblConfigInfo(9).Caption = "Loaded DLL libraries by selected process: (" & lstProcManDLLs.ListCount & ")"
End Sub

Private Sub cmdProcManRun_Click()
    If Not bIsWinNT Then
        SHRunDialog Me.hWnd, 0, 0, "Run", "Type the name of a program, folder, document or Internet resource, and Windows will open it for you.", 0
    Else
        SHRunDialog Me.hWnd, 0, 0, StrConv("Run", vbUnicode), StrConv("Type the name of a program, folder, document or Internet resource, and Windows will open it for you.", vbUnicode), 0
    End If
    Sleep 1000
    cmdProcManRefresh_Click
End Sub

Private Sub Form_Load()
    Me.Height = 7215
    If RunningInIDE Then Me.Caption = "IBProcMan"
    GetWinVersion
    cmdProcManRefresh_Click
    fraProcessManager.Caption = "Itty Bitty Process Manager - v" & App.Major & "." & Format(App.Minor, "00") & "." & App.Revision
    SetListBoxColumns lstProcessManager
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    fraProcessManager.Width = Me.ScaleWidth - 240
    lstProcessManager.Width = Me.ScaleWidth - 480
    lstProcManDLLs.Width = Me.ScaleWidth - 480
    chkProcManShowDLLs.Left = Me.ScaleWidth - 1575
    imgProcManSave.Left = Me.ScaleWidth - 2055
    imgProcManCopy.Left = Me.ScaleWidth - 2415

    fraProcessManager.Height = Me.ScaleHeight - 225
    If chkProcManShowDLLs.Value = 0 Then
        lstProcessManager.Height = Me.ScaleHeight - 1470
    Else
        lstProcessManager.Height = (Me.ScaleHeight - 1470) / 2 - 120
        lblConfigInfo(9).Top = (Me.ScaleHeight - 1470) / 2 + 480
        lstProcManDLLs.Top = (Me.ScaleHeight - 1470) / 2 + 720
        lstProcManDLLs.Height = Me.ScaleHeight - 1590 - (Me.ScaleHeight - 1470) / 2
    End If
    cmdProcManKill.Top = Me.ScaleHeight - 720
    cmdProcManRefresh.Top = Me.ScaleHeight - 720
    cmdProcManRun.Top = Me.ScaleHeight - 720
    cmdProcManBack.Top = Me.ScaleHeight - 720
    lblProcManDblClick.Top = Me.ScaleHeight - 720
End Sub

Private Sub fraProcessManager_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub

Private Sub imgProcManCopy_Click()
    imgProcManCopy.BorderStyle = 1
    DoEvents
    If chkProcManShowDLLs.Value = 1 Then
        CopyProcessList lstProcessManager, lstProcManDLLs, True
    Else
        CopyProcessList lstProcessManager, lstProcManDLLs, False
    End If
    imgProcManCopy.BorderStyle = 0
End Sub

Private Sub imgProcManSave_Click()
    imgProcManSave.BorderStyle = 1
    DoEvents
    If chkProcManShowDLLs.Value = 1 Then
        SaveProcessList lstProcessManager, lstProcManDLLs, True
    Else
        SaveProcessList lstProcessManager, lstProcManDLLs, False
    End If
    imgProcManSave.BorderStyle = 0
End Sub

Private Sub lstProcessManager_DblClick()
    Dim s$
    s = lstProcessManager.List(lstProcessManager.ListIndex)
    s = Mid(s, InStr(s, vbTab) + 1)
    ShowFileProperties s
End Sub

Private Sub lstProcessManager_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13: lstProcessManager_DblClick
        Case 33, 34, 35, 36, 37, 38, 40: lstProcessManager_MouseUp 1, 0, 0, 0
    End Select
End Sub

Private Sub lstProcessManager_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If lstProcManDLLs.Visible = False Then Exit Sub
        Dim s$
        s = lstProcessManager.List(lstProcessManager.ListIndex)
        s = Left(s, InStr(s, vbTab) - 1)
        If Not bIsWinNT Then
            RefreshDLLList CLng(s), lstProcManDLLs
        Else
            RefreshDLLListNT CLng(s), lstProcManDLLs
        End If
        lblConfigInfo(9).Caption = "Loaded DLL libraries by selected process: (" & lstProcManDLLs.ListCount & ")"
    ElseIf Button = 2 Then
        PopupMenu mnuMain, , , , mnuMainProps
    End If
End Sub

Private Sub lstProcManDLLs_DblClick()
    ShowFileProperties lstProcManDLLs.List(lstProcManDLLs.ListIndex)
End Sub

Private Sub lstProcManDLLs_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then lstProcManDLLs_DblClick
End Sub

Private Function RunningInIDE() As Boolean
    On Error Resume Next
    Debug.Print 1 / 0
    If Err Then RunningInIDE = True
End Function

Private Sub lstProcManDLLs_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        mnuMainKill.Enabled = False
        PopupMenu mnuMain, , , , mnuMainProps
        mnuMainKill.Enabled = True
    End If
End Sub

Private Sub mnuMainCopy_Click()
    imgProcManCopy_Click
End Sub

Private Sub mnuMainKill_Click()
    cmdProcManKill_Click
End Sub

Private Sub mnuMainProps_Click()
    lstProcessManager_DblClick
End Sub

Private Sub mnuMainSave_Click()
    imgProcManSave_Click
End Sub
