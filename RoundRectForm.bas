Attribute VB_Name = "RoundRectForm"
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "User32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

Public Sub CreateRoundRectFromWindow(ByRef oWindow As Object, ByVal x As Integer, ByVal y As Integer)
    Dim lRight As Long
    Dim lBottom As Long
    Dim hRgn As Long
    With oWindow
        lRight = .Width / Screen.TwipsPerPixelX
        lBottom = .Height / Screen.TwipsPerPixelY
        hRgn = CreateRoundRectRgn(0, 0, lRight, lBottom, x, y)
        SetWindowRgn .hWnd, hRgn, True
    End With
    DeleteObject hRgn
End Sub
