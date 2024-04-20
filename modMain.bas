Attribute VB_Name = "modMain"
Option Explicit

Public Locked As Boolean, TimerShowed As Boolean, lCdTime As Long

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Declare Function GetVolumeInformation Lib "kernel32" _
   Alias "GetVolumeInformationA" _
   (ByVal lpRootPathName As String, _
    ByVal lpVolumeNameBuffer As String, _
    ByVal nVolumeNameSize As Long, _
    lpVolumeSerialNumber As Long, _
    lpMaximumComponentLength As Long, _
    lpFileSystemFlags As Long, _
    ByVal lpFileSystemNameBuffer As String, _
    ByVal nFileSystemNameSize As Long) As Long

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dX As Long, ByVal dY As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Type POINTAPI
    x As Long
    y As Long
End Type

'---------------------------------------------------------------
'-注册表 API 声明...
'---------------------------------------------------------------
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
'---------------------------------------------------------------
'- 注册表 Api 常数...
'---------------------------------------------------------------
' Reg Data Types...
Const REG_SZ = 1                         ' Unicode空终结字符串
Const REG_EXPAND_SZ = 2                  ' Unicode空终结字符串
Const REG_DWORD = 4                      ' 32-bit 数字
' 注册表创建类型值...
Const REG_OPTION_NON_VOLATILE = 0       ' 当系统重新启动时，关键字被保留
' 注册表关键字安全选项...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Const KEY_EXECUTE = KEY_READ
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' 注册表关键字根类型...
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
' 返回值...
Const ERROR_NONE = 0
Const ERROR_BADKEY = 2
Const ERROR_ACCESS_DENIED = 8
Const ERROR_SUCCESS = 0
'---------------------------------------------------------------
'- 注册表安全属性类型...
'---------------------------------------------------------------
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

'====================SetCtrlsBrdClr====================
Private Type RECTW
    Left                As Long
    Top                 As Long
    Right               As Long
    Bottom              As Long
    Width               As Long
    Height              As Long
End Type

Private Type RECT
    Left        As Long
    Top         As Long
    Right       As Long
    Bottom      As Long
End Type

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long

Private Const WM_DESTROY        As Long = &H2
Private Const WM_PAINT          As Long = &HF
Private Const WM_NCPAINT        As Integer = &H85
Private Const GWL_WNDPROC = (-4)
Private Color As Long
'====================SetCtrlsBrdClr====================


Public Declare Function RtlSetProcessIsCritical Lib "ntdll.dll" (ByVal NewValue&, ByVal OldValue&, ByVal WinLogon&)
Public Declare Function RtlAdjustPrivilege& Lib "ntdll" (ByVal Privilege&, ByVal NewValue&, ByVal NewThread&, OldValue&)
Public Declare Function NtShutdownSystem& Lib "ntdll" (ByVal ShutdownAction&)
Public Const SE_SHUTDOWN_PRIVILEGE& = 19
'Public Const SHUTDOWN& = 0
'Public Const RESTART& = 1

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Const IDC_HAND As Long = 32649&

Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONONFO) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Type OSVERSIONONFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformld As Long
    dwCSDVersion As String * 128
End Type

'Hook
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Const VK_LWINDOWS = &H5B
Private Const VK_RWINDOWS = &H5C
Public Const WH_KEYBOARD_LL = 13
Public Type KBDLLHOOKSTRUCT '挂个低级钩子
    vkCode As Long '这里就是我们需要的键盘虚拟码了
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type
Dim jian As KBDLLHOOKSTRUCT
'End

'======Config======
Public Function ReadString(ByVal Caption As String, ByVal item As String, ByVal Path As String) As String
    On Error Resume Next
    Dim sBuffer As String
    sBuffer = Space(256)
    GetPrivateProfileString Caption, item, vbNullString, sBuffer, 256, Path
    
    ReadString = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
End Function

Public Function MyPath() As String
    Dim sPath As String
    sPath = App.Path
    
    If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
    
    MyPath = sPath
End Function

Public Function WriteString(ByVal Caption As String, ByVal item As String, ByVal ItemValue As String, ByVal Path As String) As Long
    Dim sBuffer As String
    sBuffer = Space(256)
    
    sBuffer = ItemValue & vbNullChar
    WriteString = WritePrivateProfileString(Caption, item, sBuffer, Path)
End Function

Public Function SaveCon(item As String, Txt As String) As Long
    WriteString "Settings", item, Txt, MyPath & "LockProCfg.ini"
End Function

Public Function ReadCon(item As String) As String
    ReadCon = ReadString("Settings", item, MyPath & "LockProCfg.ini")
End Function

Public Function SavePsw(item As String, Txt As String) As Long
    WriteString "Psws", item, Txt, MyPath & "LockPro.xm5"
End Function

Public Function ReadPsw(item As String) As String
    ReadPsw = ReadString("Psws", item, MyPath & "LockPro.xm5")
End Function
'=======End========

'====================SetCtrlsBrdClr====================
Public Sub setBorderColor(hwnd As Long, Color_ As Long)
    Color = Color_
    If GetProp(hwnd, "OrigProcAddr") = 0 Then
        SetProp hwnd, "OrigProcAddr", SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
    End If
End Sub

Public Sub UnHook(hwnd As Long)
    Dim OrigProc As Long
    OrigProc = GetProp(hwnd, "OrigProcAddr")
    If Not OrigProc = 0 Then
        SetWindowLong hwnd, GWL_WNDPROC, OrigProc
        OrigProc = SetWindowLong(hwnd, GWL_WNDPROC, OrigProc)
        RemoveProp hwnd, "OrigProcAddr"
    End If
End Sub

Private Function OnPaint(OrigProc As Long, hwnd As Long, uMsg As Long, wParam As Long, lParam As Long) As Long
    Dim m_hDC       As Long
    Dim m_wRect     As RECTW
    OnPaint = CallWindowProc(OrigProc, hwnd, uMsg, wParam, lParam)
    Call pGetWindowRectW(hwnd, m_wRect)
    m_hDC = GetWindowDC(hwnd)
    Call pFrameRect(m_hDC, 0, 0, m_wRect.Width, m_wRect.Height)
    Call ReleaseDC(hwnd, m_hDC)
End Function

Private Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim OrigProc As Long
    Dim ClassName As String
    If hwnd = 0 Then Exit Function
    OrigProc = GetProp(hwnd, "OrigProcAddr")
    If Not OrigProc = 0 Then
        If uMsg = WM_DESTROY Then
            SetWindowLong hwnd, GWL_WNDPROC, OrigProc
            WindowProc = CallWindowProc(OrigProc, hwnd, uMsg, wParam, lParam)
            RemoveProp hwnd, "OrigProcAddr"
        Else
            If uMsg = WM_PAINT Or WM_NCPAINT Then

                WindowProc = OnPaint(OrigProc, hwnd, uMsg, wParam, lParam)
            Else
                WindowProc = CallWindowProc(OrigProc, hwnd, uMsg, wParam, lParam)
            End If
        End If
    Else
        WindowProc = DefWindowProc(hwnd, uMsg, wParam, lParam)
    End If
End Function

Private Function pGetWindowRectW(ByVal hwnd As Long, lpRectW As RECTW) As Long
    Dim TmpRect As RECT
    Dim rtn     As Long
    rtn = GetWindowRect(hwnd, TmpRect)
    With lpRectW
        .Left = TmpRect.Left
        .Top = TmpRect.Top
        .Right = TmpRect.Right
        .Bottom = TmpRect.Bottom
        .Width = TmpRect.Right - TmpRect.Left
        .Height = TmpRect.Bottom - TmpRect.Top
    End With
    pGetWindowRectW = rtn
End Function

Private Function pFrameRect(ByVal hDC As Long, ByVal x As Long, y As Long, ByVal Width As Long, ByVal Height As Long) As Long
    Dim TmpRect     As RECT
    Dim m_hBrush    As Long
    With TmpRect
        .Left = x
        .Top = y
        .Right = x + Width
        .Bottom = y + Height
    End With
    m_hBrush = CreateSolidBrush(Color)
    pFrameRect = FrameRect(hDC, TmpRect, m_hBrush)
    DeleteObject m_hBrush
End Function
'====================SetCtrlsBrdClr====================

Sub Main()
    If LCase(Left(Command, 3)) = "reg" Then
        Dim sReg() As String
        sReg = Split(Command, "@@")
        UpdateKey CLng(sReg(1)), sReg(2), sReg(3), sReg(4)
        End
    End If
    
    If Dir(MyPath & "LockProCfg.ini") = "" Or Dir(MyPath & "LockPro.xm5") = "" Then
        MsgBox "未发现 Lock Pro 的配置文件，请检查文件是否存在！", 48, "Lock Pro 无法启动"
        End
    End If
    
    If Not CheckWinsockOCX Then
        MsgBox "Lock Pro 运行所依赖的必需组件未正确注册，请运行安装程序进行修复！", 48, "Lock Pro 无法启动"
        End
    End If
    
    Load frmTray
End Sub

Public Sub FormOnTop(frm As Form, Optional isFull As Boolean = False)
    On Error Resume Next
    If isFull Then
        SetWindowPos frm.hwnd, -1, 0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY, 0
    Else
        SetWindowPos frm.hwnd, -1, 0, 0, 0, 0, 3
    End If
End Sub

Public Function CheckUSB() As Long
    On Error Resume Next
    Dim fso As Object, disks As Object, disk As Object, ID As Object
    Dim sUSB As String
    sUSB = ReadPsw("USB")
    
    Set fso = CreateObject("scripting.filesystemobject")
    Set disks = fso.Drives
    For Each disk In disks
        Set ID = fso.GetDrive(fso.GetDriveName(disk))
        If ID.drivetype = 1 And disk.IsReady = True Then
            If XMD5(CStr(GetUSBSerial(ID.DriveLetter & ":\"))) = sUSB Then CheckUSB = 1: Exit For
        End If
    Next
End Function

Public Function GetUSBSerial(USB As String) As String
    Dim lpName As String, nSize As Long, nSerial As Long, lpMaxComp As Long, nFileFlags As Long, lpFileName As String
    GetVolumeInformation USB, lpName, nSize, nSerial, lpMaxComp, nFileFlags, lpFileName, nSize
    GetUSBSerial = nSerial
End Function

Public Sub mShellLnk(ByVal LnkName As String, IconFileIconIndex As String, ByVal FilePath As String, Optional ByVal FileName As String, Optional ByVal StrArg As String, Optional ByVal HookKey As String = "", Optional ByVal StrRemark As String = "", Optional ByVal strDesktop As String = "")
    Dim WshShell As Object, WScript As Object, oShellLink As Object
    Set WshShell = CreateObject("WScript.Shell")
    If strDesktop = "" Then strDesktop = WshShell.SpecialFolders("Desktop")   '桌面路径
    If UCase(Right(LnkName, 4)) = ".LNK" Then
        Set oShellLink = WshShell.CreateShortcut(strDesktop & "\" & LnkName) '创建快捷方式,参数为路径和名称
    Else
        Set oShellLink = WshShell.CreateShortcut(strDesktop & "\" & LnkName & ".lnk")
    End If
    With oShellLink
        .TargetPath = FilePath & "\" & FileName
        .Arguments = StrArg
        .WindowStyle = 1 '风格
        .Hotkey = HookKey '热键
        .IconLocation = IconFileIconIndex '图标
        .Description = StrRemark '快捷方式备注内容
        .WorkingDirectory = FilePath '源文件所在目录
        .save   '保存创建的快捷方式
    End With
    Set WshShell = Nothing
    Set oShellLink = Nothing
End Sub

'-------------------------------------------------------------------------------------------------
'sample usage - Debug.Print UpodateKey(HKEY_CLASSES_ROOT, "keyname", "newvalue")
'-------------------------------------------------------------------------------------------------
Public Function UpdateKey(KeyRoot As Long, KeyName As String, SubKeyName As String, SubKeyValue As String) As Boolean
    Dim rc As Long                                      ' 返回代码
    Dim hKey As Long                                    ' 处理一个注册表关键字
    Dim hDepth As Long                                  '
    Dim lpAttr As SECURITY_ATTRIBUTES                   ' 注册表安全类型
    
    lpAttr.nLength = 50                                 ' 设置安全属性为缺省值...
    lpAttr.lpSecurityDescriptor = 0                     ' ...
    lpAttr.bInheritHandle = True                        ' ...
    '------------------------------------------------------------
    '- 创建/打开注册表关键字...
    '------------------------------------------------------------
    rc = RegCreateKeyEx(KeyRoot, KeyName, _
                        0, REG_SZ, _
                        REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS + &H100, lpAttr, _
                        hKey, hDepth)                   ' 创建/打开//KeyRoot//KeyName
    
    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' 错误处理...
    
    '------------------------------------------------------------
    '- 创建/修改关键字值...
    '------------------------------------------------------------
    If (SubKeyValue = "") Then SubKeyValue = " "        ' 要让RegSetValueEx() 工作需要输入一个空格...
    
    ' 创建/修改关键字值
    rc = RegSetValueEx(hKey, SubKeyName, _
                       0, REG_SZ, _
                       SubKeyValue, LenB(StrConv(SubKeyValue, vbFromUnicode)))
                       
    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' 错误处理
    '------------------------------------------------------------
    '- 关闭注册表关键字...
    '------------------------------------------------------------
    rc = RegCloseKey(hKey)                              ' 关闭关键字
    
    UpdateKey = True                                    ' 返回成功
    Exit Function                                       ' 退出
CreateKeyError:
    UpdateKey = False                                   ' 设置错误返回代码
    rc = RegCloseKey(hKey)                              ' 试图关闭关键字
End Function

Public Function GetMoveNum(sToNum As Single, sNowNum As Single, lSpeed As Long, Optional lMode As Long = 0) As Long
    On Error Resume Next
    Select Case lMode
        Case 0
            Dim sTmp As Single
            sTmp = (sToNum - sNowNum) / lSpeed
            If Round(sTmp) = 0 Then sTmp = 0
            GetMoveNum = CLng(sTmp)
        Case 1
            If sNowNum < sToNum Then
                If sNowNum + lSpeed < sToNum Then
                    GetMoveNum = sNowNum + lSpeed
                Else
                    GetMoveNum = sToNum
                End If
            Else
                If sNowNum - lSpeed > sToNum Then
                    GetMoveNum = sNowNum - lSpeed
                Else
                    GetMoveNum = sToNum
                End If
            End If
    End Select
End Function

Public Sub Sleep(ByVal dwMilliseconds As Long)
    Dim SaveTime As Long
    Dim NowTime As Long
    Dim IsWait As Long
    IsWait = 0
    SaveTime = GetTickCount
    Do
       DoEvents
       NowTime = GetTickCount
       If NowTime - SaveTime >= dwMilliseconds Then
          IsWait = 1
       End If
    Loop While IsWait = 0
End Sub

Public Function HookKeyboard(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long   'Hook Win Key
    CopyMemory jian, ByVal lParam, Len(jian)
    If jian.vkCode = &H5B Or jian.vkCode = &H5C Then HookKeyboard = -1 '把左边和右边的win键全部过滤掉
End Function

Public Sub NtShutdown(Optional isReboot As Long = 0)
    RtlAdjustPrivilege SE_SHUTDOWN_PRIVILEGE, 1, 0, 0
    NtShutdownSystem isReboot
End Sub

Public Sub KillTaskMgr()
    Dim colProcessList As Object, objProcess As Object
    Set colProcessList = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2").ExecQuery _
                        ("Select * from Win32_Process Where Name='taskmgr.exe'")
    For Each objProcess In colProcessList
        objProcess.Terminate
    Next
    Set objProcess = Nothing
    Set colProcessList = Nothing
End Sub

Public Function CheckWinsockOCX() As Boolean
    On Error GoTo CWErr
    
    Dim oTestSck As Object
    Set oTestSck = CreateObject("MSWinsock.Winsock")
    CheckWinsockOCX = True
    Set oTestSck = Nothing
    
    Exit Function
CWErr:
    CheckWinsockOCX = False
End Function
