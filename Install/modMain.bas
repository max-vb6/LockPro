Attribute VB_Name = "modMain"
Option Explicit

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal csidl As Long, ByVal fCreate As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
    ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" _
    Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Sub SHChangeNotify Lib "shell32.dll" (ByVal wEventId As Long, ByVal uFlags As Long, ByVal dwItem1 As Long, ByVal dwItem2 As Long)
Private Const SHCNE_ASSOCCHANGED = &H8000000
Private Const SHCNF_IDLIST = &H0

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Function ReadString(ByVal Caption As String, ByVal item As String, ByVal Path As String) As String
    On Error Resume Next
    Dim sBuffer As String
    sBuffer = Space(128)
    GetPrivateProfileString Caption, item, vbNullString, sBuffer, 128, Path
    
    ReadString = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
End Function

Public Function WriteString(ByVal Caption As String, ByVal item As String, ByVal ItemValue As String, ByVal Path As String) As Long
    Dim sBuffer As String
    sBuffer = Space(128)
    
    sBuffer = ItemValue & vbNullChar
    WriteString = WritePrivateProfileString(Caption, item, sBuffer, Path)
End Function

Public Sub mShellLnk(ByVal LnkName As String, IconFileIconIndex As String, ByVal FilePath As String, Optional ByVal FileName As String, Optional ByVal StrArg As String, Optional ByVal HookKey As String = "", Optional ByVal StrRemark As String = "", Optional ByVal strDesktop As String = "")
    Dim WshShell As Object, WScript As Object, oShellLink As Object
    Set WshShell = CreateObject("WScript.Shell")
    If strDesktop = "" Then strDesktop = WshShell.SpecialFolders("Desktop")   '����·��
    If UCase(Right(LnkName, 4)) = ".LNK" Then
        Set oShellLink = WshShell.CreateShortcut(strDesktop & "\" & LnkName) '������ݷ�ʽ,����Ϊ·��������
    Else
        Set oShellLink = WshShell.CreateShortcut(strDesktop & "\" & LnkName & ".lnk")
    End If
    With oShellLink
        .TargetPath = FilePath & "\" & FileName
        .Arguments = StrArg
        .WindowStyle = 1 '���
        .Hotkey = HookKey '�ȼ�
        .IconLocation = IconFileIconIndex 'ͼ��
        .Description = StrRemark '��ݷ�ʽ��ע����
        .WorkingDirectory = FilePath 'Դ�ļ�����Ŀ¼
        .save   '���洴���Ŀ�ݷ�ʽ
    End With
    Set WshShell = Nothing
    Set oShellLink = Nothing
End Sub

Public Function GetMoveNum(sToNum As Single, sNowNum As Single, lSpeed As Long) As Long
    On Error Resume Next
    Dim sTmp As Single
    sTmp = (sToNum - sNowNum) / lSpeed
    If Round(sTmp) = 0 Then sTmp = 0
    GetMoveNum = CLng(sTmp)
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

Public Sub DeleteFolder(sDeleteFolder As String)
    Shell "cmd.exe /c rd /s /q """ & sDeleteFolder & """", 0
End Sub

Public Function GetStartMenuPath(Optional lPath As Long = 23) As String
    Dim sPath As String
    sPath = Space(260) & Chr(0)
    SHGetSpecialFolderPath 0, sPath, lPath, 0
    GetStartMenuPath = Trim(Replace(sPath, Chr(0), "")) & "\"
End Function

Public Sub RefreshShell()
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
End Sub

Public Function GetDirectory(Optional lhWnd As Long, Optional Msg) As String
    Dim bInfo As BROWSEINFO
    Dim Path As String
    Dim R As Long, x As Long, Pos As Integer
    ' Root folder = Desktop
    bInfo.pidlRoot = 0&
    If Not IsMissing(lhWnd) Then bInfo.hOwner = lhWnd
    ' Title in the dialog
    If IsMissing(Msg) Then
        bInfo.lpszTitle = "Select a folder."
    Else
        bInfo.lpszTitle = Msg
    End If
    ' Type of directory to return
    bInfo.ulFlags = &H1
    ' Display the dialog
    x = SHBrowseForFolder(bInfo)
    ' Parse the result
    Path = Space$(512)
    R = SHGetPathFromIDList(ByVal x, ByVal Path)
    If R Then
        Pos = InStr(Path, Chr$(0))
        GetDirectory = Left(Path, Pos - 1)
        If Mid(GetDirectory, Len(GetDirectory), 1) <> "\" Then GetDirectory = GetDirectory & "\"
    Else
        GetDirectory = ""
    End If
End Function

Public Function SaveFileFromRes(vntResourceID As Variant, sType As String, sFileName As String) As Boolean
    Dim bytImage() As Byte     'Always store binary data in byte arrays!
    Dim iFileNum As Integer     'Free File Handle
    On Error GoTo SaveFileFromRes_Err
    SaveFileFromRes = True
    'Load Binary Data from Resource file
    bytImage = LoadResData(vntResourceID, sType)
    'Get Free File Handle
    iFileNum = FreeFile
    'Open the file and save the data
    Open sFileName For Binary As iFileNum
        Put #iFileNum, , bytImage
    Close iFileNum
    Exit Function
SaveFileFromRes_Err:
    SaveFileFromRes = False: Exit Function
End Function

Public Function GetDriveSpaceString(sPath As String, Optional lLeastMB As Long) As String
    If sPath = "" Then Exit Function
    On Error GoTo GDSErr
    Dim fso As Object, drv As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set drv = fso.GetDrive(fso.GetDriveName(sPath))
    GetDriveSpaceString = "���� " & Left(sPath, 2) & " ���ÿռ� " & _
        Format(Round(drv.FreeSpace / 1024 ^ 3, 1), "0.0") & "GB �ܴ�С " & _
        Format(Round(drv.TotalSize / 1024 ^ 3, 1), "0.0") & "GB"
    If Not IsMissing(lLeastMB) Then
        If drv.FreeSpace / 1024 ^ 2 < lLeastMB Then
            GetDriveSpaceString = "���� " & Left(sPath, 2) & " �Ŀռ䲻�㣬������Ҫ " & _
                lLeastMB & "MB �Ŀռ�"
        End If
    End If
    Exit Function
GDSErr:
    GetDriveSpaceString = ""
End Function

Public Function CheckExeIsRun(exeName As String) As Boolean
    On Error GoTo CEErr
    Dim WMI As Object, Obj As Object, Objs As Object
    CheckExeIsRun = False
    Set WMI = GetObject("WinMgmts:")
    Set Objs = WMI.InstancesOf("Win32_Process")
    For Each Obj In Objs
      If (InStr(UCase(exeName), UCase(Obj.Description)) <> 0) Then
            CheckExeIsRun = True
            GoTo CEErr
      End If
    Next
CEErr:
    If Not Objs Is Nothing Then Set Objs = Nothing
    If Not WMI Is Nothing Then Set WMI = Nothing
End Function

Public Sub MoveCfg(cfgSrc As String, cfgTo As String)
    Dim sCfgs As Variant, sPsws As Variant, i As Long, sItm As String
    sCfgs = Array("First", "BGPic", "Psw", "PswLarge", "PswWait", _
        "PswErr", "Key", "Txt", "Scr", "ScrWait", "ExitPsw", "UNR", "FormLeft")
    sPsws = Array("Psw", "USB")
    For i = 0 To UBound(sCfgs)
        sItm = ReadString("Settings", sCfgs(i), cfgSrc)
        If sItm <> "" Then
            WriteString "Settings", sCfgs(i), sItm, cfgTo
        End If
    Next i
    For i = 0 To UBound(sPsws)
        sItm = ReadString("Psws", sPsws(i), cfgSrc)
        If sItm <> "" Then
            WriteString "Psws", sPsws(i), sItm, cfgTo
        End If
    Next i
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
