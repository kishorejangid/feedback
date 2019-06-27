Attribute VB_Name = "FormMove"
Public Declare Function ReleaseCapture Lib "User32" () As Long
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
