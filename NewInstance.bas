Attribute VB_Name = "NewInstance"
Option Explicit

Declare Function EnumWindows Lib "user32" (ByVal wndenmprc As Long, ByVal lParam As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Const WM_SYSCOMMAND = &H112
Public Const SC_MAXIMIZE = &HF030&

Private Target As String
' Check a returned task to see if we should
' maximize it.
Public Function EnumCallback(ByVal app_hWnd As Long, ByVal param As Long) As Long
Dim buf As String * 256
Dim title As String
Dim length As Long

    ' Get the window's title.
    length = GetWindowText(app_hWnd, buf, Len(buf))
    title = Left$(buf, length)

    ' See if this is the target window.
    If InStr(title, Target) <> 0 Then
        ' Maximize the window.
        SendMessage app_hWnd, WM_SYSCOMMAND, SC_MAXIMIZE, 0
        BringWindowToTop app_hWnd
    
        ' Stop searching.
        EnumCallback = 0
    Else
        ' Continue searching.
        EnumCallback = 1
    End If
End Function

' Ask Windows for the list of tasks.
Public Sub MaximizeTask(app_name As String)
    Target = app_name
    EnumWindows AddressOf EnumCallback, 0
End Sub

