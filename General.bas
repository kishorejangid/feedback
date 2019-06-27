Attribute VB_Name = "General"
'Project Name: feedback
'Project Developer: Kishore Kumar H
'Project Date: 4 Nov 2010
'------------------------------------------------------------------------------------------------------
Public iDept As Integer
Public iBatch As Integer
Public iSem As Integer
Public strSubj As String
Public strSec As String
Public strStaffID As String
Public conn As ADODB.Connection
Public strAdobePath As String
Public strDataSource As String
Public strDataUser As String
Public strDataPassword As String
Public bAutoStart As Boolean
Public bShutDown As Boolean
Public bAutoUpdate As Boolean
Public fConnSuccess As Boolean
Public strCollegeName As String
Public strCity As String


Private LoginTable As Boolean
Private DeptTable As Boolean
Private SubjTable As Boolean
Private StaffTable As Boolean
Private StaffHandleTable As Boolean
Private MasterTable As Boolean

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub OpenDataBase()
    On Error GoTo errHandler
    Set conn = New ADODB.Connection
    conn.Open "Provider=MSDAORA.1;User ID=" & strDataUser & ";Password=" & strDataPassword & ";Data Source=" & strDataSource & ";Persist Security Info=False;"
    fConnSuccess = True
    Exit Sub
errHandler:
    If Err.Number = -2147217843 Then
        fConnSuccess = False
        MsgBox "Connection to Oracle Database failed." & vbCrLf & vbCrLf & "Change Oracle Settings in the application", vbCritical, "Jangid Corporation"
    ElseIf Err.Number <> 0 Then
        MsgBox "Error Number: " & Err.Number & vbCrLf & vbCrLf & "Error Description: " & Err.Description & vbCrLf & "Error Location : OpenDataBase()", vbCritical
    End If
End Sub

Sub Main()
    On Error Resume Next
    If App.PrevInstance Then
        MaximizeTask "Feedback"
        End
    End If
    strCollegeName = GetSetting(App.CompanyName & "\\feedback", "Settings", "College Name", "ASIATIC TECHNICAL RESEARCH AND DEVELOPMENT CENTRE, RAJAPALAYAM - 626117..")
    strCity = GetSetting(App.CompanyName & "\\feedback", "Settings", "City", "Tirunelveli")
    strAdobePath = GetSetting(App.CompanyName & "\\feedback", "Paths", "Adobe", "C:\Program Files\Adobe\Acrobat 5.0\Reader\AcroRd32.exe")
    strDataSource = GetSetting(App.CompanyName & "\\feedback", "Oracle", "Data Source", "student")
    strDataUser = GetSetting(App.CompanyName & "\\feedback", "Oracle", "Data User", "kishore")
    strDataPassword = GetSetting(App.CompanyName & "\\feedback", "Oracle", "Data Password", "kishore")
    bAutoStart = GetSetting(App.CompanyName & "\\feedback", "Settings", "Auto Start", False)
    bShutDown = GetSetting(App.CompanyName & "\\feedback", "Settings", "Shut Down", False)
    bAutoUpdate = GetSetting(App.CompanyName & "\\feedback", "Settings", "Auto Update", False)
    If bAutoStart Then
        Set reg = CreateObject("Wscript.Shell")
        reg.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\feedback", App.Path & "\" & App.EXEName & ".exe", "REG_SZ"
    End If
    CurrentAppVer = App.Major & "." & App.Minor & "." & App.Revision
    If bAutoUpdate Then
        frmPop.Show
        DoEvents
        CheckForUpdate
        Unload frmPop
        If fUpdate = True Then
            frmUpdate.Show
        Else
            Call OpenDataBase
            Call CheckTables
            If fConnSuccess Then
                frmMain.Show
                'frmFeed.Show
            Else
                frmSettings.Show
            End If
        End If
    Else
        Call OpenDataBase
        Call CheckTables
        If fConnSuccess Then
            frmMain.Show
            'frmFeed.Show
        Else
            frmSettings.Show
        End If
    End If
End Sub
Private Sub CheckTables()
    'Fetch DataBase Settings from Registry
    LoginTable = GetSetting(App.CompanyName, "DataBase", "LoginTable", False)
    DeptTable = GetSetting(App.CompanyName, "DataBase", "DeptTable", False)
    SubjTable = GetSetting(App.CompanyName, "DataBase", "SubjTable", False)
    StaffTable = GetSetting(App.CompanyName, "DataBase", "StaffTable", False)
    StaffHandleTable = GetSetting(App.CompanyName, "DataBase", "StaffHandleTable", False)
    MasterTable = GetSetting(App.CompanyName, "DataBase", "MasterTable", False)
        
    
    If LoginTable = False Then
        Call CreateLoginTable
    End If
    If DeptTable = False Then
        Call CreateDeptTable
    End If
    If SubjTable = False Then
        Call CreateSubjTable
    End If
    If StaffTable = False Then
        Call CreateStaffTable
    End If
    If StaffHandleTable = False Then
        Call CreateStaffHandleTable
    End If
    If MasterTable = False Then
        Call CreateMasterTable
    End If
     
    If Dir(App.Path & "\Reports", vbDirectory) = vbNullString Then
        MkDir App.Path & "\Reports"
    End If
End Sub

Public Function RoundUp(ByVal x As Double, Optional ByVal Factor As Double = 1) As Double
    Dim Temp As Double
    Temp = Int(x * Factor)
    RoundUp = (Temp + IIf(x = Temp, 0, 1)) / Factor
End Function
Public Sub cmbSec_Load(ComboBox As ComboBox)
    ComboBox.Clear
    ComboBox.AddItem ("A")
    ComboBox.AddItem ("B")
    ComboBox.AddItem ("C")
    ComboBox.AddItem ("D")
    ComboBox.AddItem ("E")
    ComboBox.AddItem ("F")
    ComboBox.AddItem ("G")
    ComboBox.AddItem ("H")
    strSec = ComboBox.Text
End Sub
Public Function Department(ComboBox As ComboBox) As Integer
    On Error Resume Next
    Dim rsDeptCode As New ADODB.Recordset
    Dim strSql As String
    strSql = "select deptcode from dept where deptshort='" & UCase(ComboBox.Text) & "'"
    rsDeptCode.Open strSql, conn, adOpenDynamic, adLockOptimistic, -1
    Department = rsDeptCode.Fields(0)
End Function
Public Function getDeptCode(strDeptShort As String) As Integer
    On Error Resume Next
    Dim rsDeptCode As New ADODB.Recordset
    Dim strSql As String
    strSql = "select deptcode from dept where deptshort='" & UCase(strDeptShort) & "'"
    rsDeptCode.Open strSql, conn, adOpenDynamic, adLockOptimistic, -1
    getDeptCode = rsDeptCode.Fields(0)
End Function
Public Sub cmbDept_Load(ComboBox As ComboBox)
    On Error GoTo errHan
    ComboBox.Clear
    Dim rsDept As New ADODB.Recordset
    Dim strSql As String
    strSql = "select deptshort from dept order by deptshort"
    rsDept.Open strSql, conn, adOpenDynamic, adLockOptimistic, -1
    rsDept.MoveFirst
    ComboBox.FontSize = 8
    While Not rsDept.EOF
        ComboBox.AddItem rsDept.Fields(0)
        rsDept.MoveNext
    Wend
    iDept = Department(ComboBox)
    Exit Sub
errHan:
    If Err.Number = 3021 Then
        MsgBox "No Departments are Created yet.", vbInformation, "feedback"
    End If
End Sub
Public Sub cmbBatch_Load(ComboBox As ComboBox)
    ComboBox.Clear
    For i = 2007 To CInt(DateTime.Year(Date))
        ComboBox.AddItem (i)
    Next
    iBatch = Val(ComboBox.Text)
End Sub
Public Sub cmbSem_Load(ComboBox As ComboBox)
    ComboBox.Clear
    ComboBox.AddItem (1)
    ComboBox.AddItem (2)
    ComboBox.AddItem (3)
    ComboBox.AddItem (4)
    ComboBox.AddItem (5)
    ComboBox.AddItem (6)
    ComboBox.AddItem (7)
    ComboBox.AddItem (8)
    iSem = Val(ComboBox.Text)
End Sub
Public Function JangidFormat(str As String) As String
    Dim i As Integer
    Dim newstr As String
    newstr = UCase(Mid$(str, 1, 1))
    For i = 2 To Len(str)
        newstr = newstr & LCase(Mid$(str, i, 1))
        If Mid$(str, i, 1) = " " Or Mid$(str, i, 1) = "," Or Mid$(str, i, 1) = "." Then
            newstr = newstr & UCase(Mid$(str, i + 1, 1))
            i = i + 1
        End If
    Next
    JangidFormat = newstr
End Function

Public Function FID() As Integer
    On Error Resume Next
    Dim rsFID As New ADODB.Recordset
    Dim strSql As String
    strSql = "Select nvl(max(fid),0) from master where dept='" & iDept & "' and batch='" & Mid(iBatch, 3, 2) & "' and sem='" & iSem & "' and sec='" & strSec & "'"
    rsFID.Open strSql, conn, adOpenDynamic, adLockOptimistic, -1
    FID = rsFID.Fields(0) + 1
    rsFID.Close
End Function
Public Function getNewStaffID() As String
    On Error Resume Next
    Dim strStaffID As String
    Dim rsStaffID As New ADODB.Recordset
    strStaffID = "select nvl(max(staffid),0) from staff where dept='" & iDept & "'"
    rsStaffID.Open strStaffID, conn, adOpenDynamic, adLockOptimistic, -1
    If rsStaffID.Fields(0) = "0" Then
        getNewStaffID = iDept & "001"
    Else
        getNewStaffID = rsStaffID.Fields(0) + 1
    End If
End Function
Public Function getStaffID(strStaffName As String, strStaffDept As String) As String
    On Error Resume Next
    Dim strStaffID As String
    Dim rsStaffID As New ADODB.Recordset
    rsStaffID.CursorLocation = adUseClient
    strStaffID = "select staffid from staff where dept='" & strStaffDept & "' and staffname= '" & UCase(Trim(strStaffName)) & "'"
    rsStaffID.Open strStaffID, conn, adOpenDynamic, adLockOptimistic, -1
    getStaffID = rsStaffID.Fields(0)
End Function
Public Sub cmbSubj_Load(ComboBox As ComboBox)
    On Error Resume Next
    Dim rsSubj As New ADODB.Recordset
    Dim strSubj As String
    strSql = "select subjcode from subj where dept='" & iDept & "' and semno='" & iSem & "' and batch='" & Mid(iBatch, 3, 2) & "' order by subjcode"
    rsSubj.Open strSql, conn, adOpenDynamic, adLockOptimistic, -1
    rsSubj.MoveFirst
    ComboBox.Clear
    While Not rsSubj.EOF
        ComboBox.AddItem rsSubj.Fields(0)
        rsSubj.MoveNext
    Wend
    ComboBox.ListIndex = 0
End Sub
Public Sub cmbStaff_Load(ComboBox As ComboBox)
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    Dim sql As String
    cmbStaff.Clear
    sql = "select staffname from staff where dept='" & iDept & "' order by staffname"
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic, -1
    rs.MoveFirst
    ComboBox.Clear
    ComboBox.FontSize = 8
    Do While Not rs.EOF
        ComboBox.AddItem rs.Fields(0)
        rs.MoveNext
    Loop
End Sub
