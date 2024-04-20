Attribute VB_Name = "modMain"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Function MyPath() As String
    Dim sPath As String
    sPath = App.Path
    
    If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
    
    MyPath = sPath
End Function

Public Function NumToByte(lByt As Long, Optional lLen As Long) As String
    If lByt < 2 ^ 20 Then
        NumToByte = Round(lByt / 2 ^ 10, lLen) & "KB"
    Else
        NumToByte = Round(lByt / 2 ^ 20, lLen) & "MB"
    End If
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

