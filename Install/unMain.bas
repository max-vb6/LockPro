Attribute VB_Name = "modMain"
Option Explicit

Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal csidl As Long, ByVal fCreate As Long) As Long

'Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
'Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'Public Const HKEY_LOCAL_MACHINE = &H80000002
'Public Const REG_SZ = 1

Public Declare Function GetTickCount Lib "kernel32" () As Long

'Public Sub SaveString(hKey As Long, strPath As String, strValue As String, strData As String)
'On Error Resume Next
'Dim keyHand As Long
'Dim R As Long
'R = RegCreateKey(hKey, strPath, keyHand)
'R = RegSetValueEx(keyHand, strValue, 0, REG_SZ, ByVal strData, LenB(StrConv(strData, vbFromUnicode)))
'R = RegCloseKey(keyHand)
'End Sub

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
