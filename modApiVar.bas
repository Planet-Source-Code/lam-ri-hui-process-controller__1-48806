Attribute VB_Name = "modApiVar"
Option Explicit


Public Const VK_F9 = &H78
Public Const VK_F10 = &H79
Public Const VK_F1 = &H70

Public Const WM_QUIT As Long = &H12
Public Const WM_CLOSE As Long = &H10

Public bStart As Boolean


Public Declare Function EnumWindows Lib "user32" ( _
ByVal lpEnumFunc As Long, _
ByVal lParam As Long) As Boolean

Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" ( _
ByVal hwnd As Long, _
ByVal lpString As String, _
ByVal cch As Long) As Long

Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" ( _
ByVal hwnd As Long) As Long

Public Declare Function GetAsyncKeyState Lib "user32.dll" ( _
     ByVal vKey As Long) As Integer


Public Declare Function DestroyWindow Lib "user32.dll" ( _
     ByVal hwnd As Long) As Long

Public Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" ( _
     ByVal hwnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     ByVal lParam As Long) As Long

Public Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
On Error GoTo err
    Dim sSave As String, Ret As Long
    Dim lCnt As Long
    
    Ret = GetWindowTextLength(hwnd)
    sSave = Space(Ret)
    GetWindowText hwnd, sSave, Ret + 1
    
    For lCnt = 0 To frmProcess.lstBlocked.ListCount - 1
        
        If InStr(1, sSave, frmProcess.lstBlocked.List(lCnt), vbTextCompare) > 0 Then
            
           If hwnd <> frmProcess.lstBlocked.hwnd Then
                PostMessage hwnd, WM_CLOSE, ByVal 0&, ByVal 0&
           End If
        End If
        
    Next
       
    If bStart = False Then
        frmProcess.lstList.AddItem hwnd & Chr$(134) & sSave
    End If
    
    
    EnumWindowsProc = True
    Exit Function
err:
    Debug.Print err.Description
End Function


