Attribute VB_Name = "modNet"
Option Explicit

'Dim bBuf() As Byte
Dim bWorking As Boolean

Private Declare Function MsgBoxTimeout Lib "user32" Alias "MessageBoxTimeoutA" _
    (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long, ByVal wlange As Long, ByVal dwTimeout As Long) As Long

Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Private Type STARTUPINFO
    cb                              As Long
    lpReserved                      As String
    lpDesktop                       As String
    lpTitle                         As String
    dwX                             As Long
    dwY                             As Long
    dwXSize                         As Long
    dwYSize                         As Long
    dwXCountChars                   As Long
    dwYCountChars                   As Long
    dwFillAttribute                 As Long
    dwFlags                         As Long
    wShowWindow                     As Integer
    cbReserved2                     As Integer
    lpReserved2                     As Long
    hStdInput                       As Long
    hStdOutput                      As Long
    hStdError                       As Long
End Type
Private Type PROCESS_INFORMATION
    hProcess                        As Long
    hThread                         As Long
    dwProcessId                     As Long
    dwThreadId                      As Long
End Type
Private Type SECURITY_ATTRIBUTES
    nLength                         As Long
    lpSecurityDescriptor            As Long
    bInheritHandle                  As Long
End Type
Private Const NORMAL_PRIORITY_CLASS  As Long = &H20&
Private Const STARTF_USESTDHANDLES   As Long = &H100&
Private Const STARTF_USESHOWWINDOW   As Long = &H1&
Private Const SW_HIDE                As Long = 0&
Private Const INFINITE               As Long = &HFFFF&

Private Function RunCommand(commandline As String) As String
    Dim si As STARTUPINFO                                                       'used to send info the CreateProcess
    Dim pi As PROCESS_INFORMATION                                               'used to receive info about the created process
    Dim retval As Long                                                          'return value
    Dim hRead As Long                                                           'the handle to the read end of the pipe
    Dim hWrite As Long                                                          'the handle to the write end of the pipe
    Dim sBuffer(0 To 63) As Byte                                                'the buffer to store data as we read it from the pipe
    Dim lgSize As Long                                                          'returned number of bytes read by readfile
    Dim sa As SECURITY_ATTRIBUTES
    Dim strResult As String                                                     'returned results of the command line
    'set up security attributes structure
    With sa
        .nLength = Len(sa)
        .bInheritHandle = 1&                                                    'inherit, needed for this to work
        .lpSecurityDescriptor = 0&
    End With
    'create our anonymous pipe an check for success
    ' note we use the default buffer size
    ' this could cause problems if the process tries to write more than this buffer size
    retval = CreatePipe(hRead, hWrite, sa, 0&)
    If retval = 0 Then
        'MsgBox "错误提示:创建管道失败!"
        RunCommand = ""
        Exit Function
    End If
    'set up startup info
    With si
        .cb = Len(si)
        .dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW                 'tell it to use (not ignore) the values below
        .wShowWindow = SW_HIDE
        .hStdOutput = hWrite                                                    'pass the write end of the pipe as the processes standard output
    End With
    'run the command line and check for success
    retval = CreateProcess(vbNullString, commandline & vbNullChar, sa, sa, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, vbNullString, si, pi)
    If retval Then
        'wait until the command line finishes
        ' trouble if the app doesn't end, or waits for user input, etc
        WaitForSingleObject pi.hProcess, INFINITE
        'read from the pipe until there's no more (bytes actually read is less than what we told it to)
        Do While ReadFile(hRead, sBuffer(0), 64, lgSize, ByVal 0&)
            'convert byte array to string and append to our result
            strResult = strResult & StrConv(sBuffer(), vbUnicode)
            'TODO = what's in the tail end of the byte array when lgSize is less than 64???
            Erase sBuffer()
            If lgSize <> 64 Then Exit Do
            DoEvents
        Loop
        'close the handles of the process
        CloseHandle pi.hProcess
        CloseHandle pi.hThread
    Else
        'MsgBox "错误提示:创建进程失败!" & vbCrLf
    End If
    'close pipe handles
    CloseHandle hRead
    CloseHandle hWrite
    'return the command line output
    RunCommand = Replace(strResult, vbNullChar, "")
End Function

Public Function GetMACAddr(sIP As String) As String
    If sIP = GetMyIP Then
        GetMACAddr = GetMyIP(True)
    Else
        Dim sTmp As String, i As Long
        sTmp = RunCommand("arp -a")
        sTmp = Replace(sTmp, " ", "")
        i = InStr(sTmp, sIP)
        If i <> 0 Then
            GetMACAddr = UCase(Replace(Mid(sTmp, i + Len(sIP), 17), "-", ":"))
        Else
            GetMACAddr = ""
        End If
    End If
End Function

Private Function CheckLicenseState(sIP As String) As Boolean
    Dim sTmp As String, sLc As String, sLcs() As String, i As Long
    sTmp = XMD5(GetMACAddr(sIP))
    sLc = ReadPsw("Remolock")
    CheckLicenseState = False
    If sLc <> "" Then
        If InStr(sLc, ",") = 0 Then
            CheckLicenseState = (sLc = sTmp)
        Else
            sLcs = Split(sLc, ",")
            For i = 0 To UBound(sLcs)
                If sLcs(i) = sTmp Then
                    CheckLicenseState = True
                End If
            Next i
        End If
    End If
End Function

Private Function AddLicense(sIP As String) As Long
    On Error GoTo ALErr
    
    Dim sTmp As String, sMAC As String
    sMAC = GetMACAddr(sIP)
    sTmp = ReadPsw("Remolock")
    If Len(Replace(sMAC, ":", "")) <> Len(sMAC) - 5 Then
        AddLicense = -1
        Exit Function
    ElseIf Len(sTmp) = 32 * 7 + 7 Then
        AddLicense = -2
        Exit Function
    End If
    SavePsw "Remolock", sTmp & XMD5(sMAC) & ","
    AddLicense = 0
    
    Exit Function
ALErr:
    AddLicense = Err.Number
End Function

Public Function GetLicenseNum() As Long
    Dim sTmp As String, sLcs() As String, lLNum As Long, i As Long
    sTmp = ReadPsw("Remolock")
    lLNum = 0
    If sTmp <> "" Then
        If InStr(sTmp, ",") = 0 Then
            lLNum = 1
        Else
            sLcs = Split(sTmp, ",")
            For i = 0 To UBound(sLcs)
                If sLcs(i) <> "" Then lLNum = lLNum + 1
            Next i
        End If
    End If
    GetLicenseNum = lLNum
End Function

Public Function GetMyIP(Optional bGetMAC As Boolean = False) As String
    Dim strComputer As String
    Dim objWMI As Object
    Dim colIP As Object
    Dim IP As Object
    Dim i As Integer
    strComputer = "."
    Set objWMI = GetObject("winmgmts://" & strComputer & "/root/cimv2")
    Set colIP = objWMI.ExecQuery _
                ("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
    For Each IP In colIP
        If Not IsNull(IP.IpAddress) Then
                GetMyIP = IIf(bGetMAC, IP.MacAddress(LBound(IP.IpAddress)), IP.IpAddress(LBound(IP.IpAddress)))
                Exit For
        End If
    Next
End Function

Private Function GetDeviceByUA(sUA As String) As String
    Dim sTmp As String, sKeys() As Variant, sDevices() As Variant, i As Long
    sKeys = Array("windows", "windows phone", "macintosh", "ipad", "ipod", "iphone", "android", "linux")
    sDevices = Array("Windows 设备", "Windows Phone 移动设备", "Mac 计算机", "iPad", "iPod", "iPhone", "Android 移动设备", "Linux 设备")
    For i = 0 To UBound(sKeys)
        If InStr(LCase(sUA), sKeys(i)) <> 0 Then
            sTmp = sDevices(i)
            Exit For
        End If
    Next i
    GetDeviceByUA = IIf(sTmp = "", "未知设备", sTmp)
End Function

Private Function HandleHTML(sHTML As String, sIP As String) As String
    Dim sTmp As String, sInfos() As Variant
    sTmp = sHTML
    sInfos = Array("此刻，世界只剩下按钮和你", "至繁归于至简", "择善而从", "知识最大的敌人不是无知，而是错觉", _
                    "知道得越多，越意识到自己无知", "‘任意’键在哪里？", "预测未来最好的方式是实现它", _
                    "让一切尽可能简单，但不要更简单", "Follow your heart.")
    Randomize
    sTmp = Replace(sTmp, "%INFO%", sInfos(Int(Rnd() * (UBound(sInfos) + 1))))
    'sTmp = Replace(sTmp, "%YEAR%", Year(Now))
    If CheckLicenseState(sIP) Then
        If Locked Then
            sTmp = Replace(sTmp, "%STATE%", "您的计算机已经锁定")
            sTmp = Replace(sTmp, "%BTNTXT%", "计算机已锁定")
            sTmp = Replace(sTmp, "%BTNLINK%", "/")
        Else
            sTmp = Replace(sTmp, "%STATE%", "按下按钮来锁定目标计算机")
            sTmp = Replace(sTmp, "%BTNTXT%", "立即锁定计算机")
            sTmp = Replace(sTmp, "%BTNLINK%", "/lock")
        End If
    Else
        sTmp = Replace(sTmp, "%STATE%", "您还未经过目标计算机的授权<br />在获得授权后方可使用该功能")
        sTmp = Replace(sTmp, "%BTNTXT%", "重新获取授权")
        sTmp = Replace(sTmp, "%BTNLINK%", "/")
    End If
    HandleHTML = sTmp
End Function

Public Sub HTTPRespond(sckReceive As Winsock, sData As String, sHTMLPath As String)      'Server 端响应请求用代码
    On Error GoTo HTTPErr
        
    Dim sTmp As String, sCmd() As String, sUA As String
    sTmp = sData
    sUA = ""
    With sckReceive
        sCmd = Split(sTmp, vbCrLf)
        Dim i As Long
        For i = 0 To UBound(sCmd)
            If Left(LCase(sCmd(i)), 11) = "user-agent:" Then
                sUA = Trim(Right(sCmd(i), Len(sCmd(i)) - 11))                   '载入 User-Agent 数据
            End If
        Next i
        
        sCmd = Split(sCmd(0), " ")
        sTmp = ""
        sCmd(1) = Replace(sCmd(1), "/", "\")                                    'sCmd(1) 存储了 URL 后的 Web 路径
HTTPDirect:
        Select Case LCase(sCmd(1))
            Case "\favicon.ico"
                .Tag = 0
                GoTo HTTPDone
            Case "\"
                sCmd(1) = MyPath & sHTMLPath & IIf(InStr(LCase(sUA), "mobile") <> 0, "\index_m.html", "\index.html")
                                                                                'User-Agent 中包含 "Mobile" 则为移动端浏览器
            Case Else
                If LCase(sCmd(1)) = "\lock" Then
                    If Locked Then
                        sCmd(1) = "\"
                        GoTo HTTPDirect
                    Else
                        frmLock.Show
                        GoTo HTTPDirect
                    End If
                'Else
                    'sCmd(1) = MyPath & sHTMLPath & sCmd(1)
                End If
        End Select
        
        If Not CheckLicenseState(.RemoteHostIP) And Not bWorking And Not Locked Then                            '授权验证
            Dim lMsg As Long
            bWorking = True
            lMsg = MsgBoxTimeout(frmTray.hwnd, "一个设备通过 Remolock 请求获得授权，若授权，此后该设备将可通过 Remolock 锁定该计算机。" & vbCrLf & _
                                "您可以通过“设置 - Remolock”管理或清除授权。" & vbCrLf & "设备类型：" & GetDeviceByUA(sUA) & vbCrLf & _
                                "是否授权？（20秒后将自动关闭此消息框并终止授权）", "Remolock 授权验证", 64 + vbYesNo, 0, 20000)
            If lMsg = vbYes Then
                Select Case AddLicense(.RemoteHostIP)
                    Case -1
                        MsgBox "无法授权该设备，获取设备 ID 时出错！", 48, "添加授权错误"
                    Case -2
                        MsgBox "无法授权该设备，授权的设备数已达上限（7个）！", 48, "添加授权错误"
                End Select
            End If
            bWorking = False
        End If
        
        Dim sType() As String, lFreeNum As Long
        sType = Split(sCmd(1), ".")                                             '后缀判断处理
        lFreeNum = FreeFile
        If (LCase(sType(UBound(sType))) = "html") Or (LCase(sType(UBound(sType))) = "htm") Then
            Open sCmd(1) For Input As lFreeNum
                sTmp = StrConv(InputB$(LOF(lFreeNum), lFreeNum), vbUnicode)
            Close lFreeNum
            sTmp = HandleHTML(sTmp, .RemoteHostIP)
            '.Tag = 0
        Else
            'Open sCmd(1) For Binary As lFreeNum
                'ReDim bBuf(LOF(lFreeNum))
                'sTmp = ""
                'Get lFreeNum, , bBuf
            'Close lFreeNum
            '.Tag = 1
            sTmp = ""
        End If
HTTPDone:
        .SendData "HTTP/1.1 200 OK" & vbCrLf & vbCrLf & sTmp
    
    Exit Sub
HTTPErr:
        .SendData "HTTP/1.1 500 Internal Server Error" & vbCrLf & vbCrLf
        '.Tag = 0
    End With
End Sub

Public Sub HTTPSendCheck(sckReceive As Winsock)
    With sckReceive
        'If .Tag = 0 Then
            .Close
            Unload sckReceive
        'Else
            '.SendData bBuf
            '.Tag = 0
        'End If
    End With
End Sub
