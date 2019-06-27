Attribute VB_Name = "Update"
Private mFTP As cFTP
Private BeginTransfer                   As Single
Private TransferRate                    As Single
Private Declare Function ClipCursor Lib "User32" (lpRect As Any) As Long

Private Declare Function OSGetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function OSGetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function OSGetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function OSWritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function OSWritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Declare Function OSGetProfileInt Lib "kernel32" Alias "GetProfileIntA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Long) As Long
Private Declare Function OSGetProfileSection Lib "kernel32" Alias "GetProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function OSGetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long

Private Declare Function OSWriteProfileSection Lib "kernel32" Alias "WriteProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String) As Long
Private Declare Function OSWriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long

Private Const nBUFSIZEINI = 1024
Private Const nBUFSIZEINIALL = 4096
Private NewVersion As String
Private OldVersion As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public FilePathName As String
Public Const SERVER = "ftp.vosi.biz" '"www.dignaj.com"
Public Const SERVERUSER = "kishorejangid" '"dignajc"
Public Const SERVERPASSWORD = "gopalvarma@123" '"gopalvarma"
Public Const SERVERFOLDERPATH = "/feedback/" '"/httpdocs/downloads/feedback/"
Public CurrentAppVer As String
Public UpdateData(1 To 6) As String
Public fUpdate As Boolean

Public Function ConvertTime(ByVal TheTime As Single) As String
    Dim NewTime                         As String
    Dim Sec                             As Single
    Dim Min                             As Single
    Dim h                               As Single
    If TheTime > 60 Then
        Sec = TheTime
        Min = Sec / 60
        Min = Int(Min)
        Sec = Sec - Min * 60
        h = Int(Min / 60)
        Min = Min - h * 60
        NewTime = h & ":" & Min & ":" & Sec
        If h < 0 Then h = 0
        If Min < 0 Then Min = 0
        If Sec < 0 Then Sec = 0
        NewTime = Format(NewTime, "HH:MM:SS")
        ConvertTime = NewTime
    End If
    If TheTime < 60 Then
        NewTime = "00:00:" & TheTime
        NewTime = Format(NewTime, "HH:MM:SS")
        ConvertTime = NewTime
    End If
End Function
Public Function RunUpdate(UpdateURL As String)
    HyperJump UpdateURL
End Function
Private Function HyperJump(ByVal url As String) As Long
    HyperJump = ShellExecute(0&, vbNullString, url, vbNullString, vbNullString, vbNormalFocus)
End Function
Private Function GetPrivateProfileString(ByVal szSection As String, ByVal szEntry As Variant, ByVal szDefault As String, ByVal szFileName As String) As String
   ' *** Get an entry in the inifile ***
   Dim szTmp                     As String
   Dim nRet                      As Long
   If (IsNull(szEntry)) Then
      ' *** Get names of all entries in the named Section ***
      szTmp = String$(nBUFSIZEINIALL, 0)
      nRet = OSGetPrivateProfileString(szSection, 0&, szDefault, szTmp, nBUFSIZEINIALL, szFileName)
   Else
      ' *** Get the value of the named Entry ***
      szTmp = String$(nBUFSIZEINI, 0)
      nRet = OSGetPrivateProfileString(szSection, CStr(szEntry), szDefault, szTmp, nBUFSIZEINI, szFileName)
   End If
   GetPrivateProfileString = Left$(szTmp, nRet)
End Function
Private Function GetProfileString(ByVal szSection As String, ByVal szEntry As Variant, ByVal szDefault As String) As String
   ' *** Get an entry in the WIN inifile ***
   Dim szTmp                    As String
   Dim nRet                     As Long
   If (IsNull(szEntry)) Then
      ' *** Get names of all entries in the named Section ***
      szTmp = String$(nBUFSIZEINIALL, 0)
      nRet = OSGetProfileString(szSection, 0&, szDefault, szTmp, nBUFSIZEINIALL)
   Else
      ' *** Get the value of the named Entry ***
      szTmp = String$(nBUFSIZEINI, 0)
      nRet = OSGetProfileString(szSection, CStr(szEntry), szDefault, szTmp, nBUFSIZEINI)
   End If
   GetProfileString = Left$(szTmp, nRet)
End Function
Public Sub CheckForUpdate()
    On eror GoTo errHan
    Set mFTP = New cFTP
    mFTP.SetModePassive
    mFTP.SetTransferBinary
    
    Dim lTimer As Long
    Dim strRemote As String
    Dim strLocal As String
    Dim NewVer As String
    Dim Oldver As String
    Dim url As String
    Dim DOR As String
    Dim FileSize As String
    Dim WhatsNew As String
    NewVer = "none"
    Oldver = "none"
        
    strRemote = "feedback.inf"
    strLocal = App.Path & "\feedback.inf"
    lTimer = Timer
    

    If mFTP.OpenConnection(SERVER, SERVERUSER, SERVERPASSWORD) Then
        mFTP.SetFTPDirectory SERVERFOLDERPATH
        If Not mFTP.FTPDownloadFile(strLocal, strRemote) Then
            fUpdate = False
        End If
        DoEvents
    End If
    mFTP.CloseConnection
    
    'Gets your Version
    Oldver = CurrentAppVer

    'State & Access 'feedback.inf' file
    FilePathName = App.Path + "\feedback.inf"
        
    NewVer = GetPrivateProfileString("Version", "Version", "", FilePathName)
    DOR = GetPrivateProfileString("Version", "DOR", "", FilePathName)
    FileSize = GetPrivateProfileString("Version", "Filesize", "", FilePathName)
    WhatsNew = GetPrivateProfileString("Version", "Whatsnew", "", FilePathName)
    FileName = GetPrivateProfileString("Version", "FileName", "", FilePathName)
    
    If CInt(Mid(Oldver, 1, 1)) >= CInt(Mid(NewVer, 1, 1)) Then
        If CInt(Mid(Oldver, 3, 1)) >= CInt(Mid(NewVer, 3, 1)) Then
            If CInt(Mid(Oldver, 5, Len(Oldver) - 4)) >= CInt(Mid(NewVer, 5, Len(NewVer) - 4)) Then
                UpdateData(1) = NewVer
                UpdateData(2) = ""
                UpdateData(3) = ""
                UpdateData(4) = ""
                UpdateData(5) = ""
                fUpdate = False
            Else
                UpdateData(1) = NewVer
                UpdateData(2) = DOR
                UpdateData(3) = FileSize
                UpdateData(4) = WhatsNew
                UpdateData(5) = FileName
                fUpdate = True
            End If
        Else
            UpdateData(1) = NewVer
            UpdateData(2) = DOR
            UpdateData(3) = FileSize
            UpdateData(4) = WhatsNew
            UpdateData(5) = FileName
            fUpdate = True
        End If
    Else
        UpdateData(1) = NewVer
        UpdateData(2) = DOR
        UpdateData(3) = FileSize
        UpdateData(4) = WhatsNew
        UpdateData(5) = FileName
        fUpdate = True
    End If
errHan:
    If mFTP.GetLastErrorMessage = "12007" Then
        fUpdate = False
    ElseIf mFTP.GetLastErrorMessage = "12014" Then
        fUpdate = False
    ElseIf mFTP.GetLastErrorMessage = "12002" Then
        fUpdate = False
    ElseIf mFTP.GetLastErrorMessage = "12029" Then
        fUpdate = False
    ElseIf mFTP.GetLastErrorMessage = "12015" Then
        fUpdate = False
    ElseIf mFTP.GetLastErrorMessage = "12002" Then
        fUpdate = False
    End If
End Sub

