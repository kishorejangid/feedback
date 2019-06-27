Attribute VB_Name = "modTables"
Public Sub CreateLoginTable()
    On Error Resume Next
    Dim strSql As String
    Dim rsTable As New ADODB.Recordset
    rsTable.CursorLocation = adUseClient
    strSql = "CREATE TABLE LOGIN(LOGINTYPE VARCHAR2(20) NOT NULL ENABLE,LOGINID VARCHAR2(20) NOT NULL ENABLE,LOGINPASSWORD VARCHAR2(20) NOT NULL ENABLE,CONSTRAINT LOGIN_CON UNIQUE (LOGINID) ENABLE)"
    rsTable.Open strSql, conn, adOpenDynamic, adLockOptimistic, -1
    
    If Err.Number <> 0 Then
        If Err.Number = -2147217900 Then
            SaveSetting App.CompanyName, "DataBase", "LoginTable", True
        Else
            MsgBox "Error:" & Err.Description & vbCrLf & "Error Number:" & Err.Number
        End If
    Else
        SaveSetting App.CompanyName, "DataBase", "LoginTable", True
        MsgBox "Table Login Created Sucessfully"
    End If
End Sub

Public Sub CreateDeptTable()
    On Error Resume Next
    Dim strSql As String
    Dim rsTable As New ADODB.Recordset
    rsTable.CursorLocation = adUseClient
    strSql = "CREATE TABLE DEPT ( DEPTCODE NUMBER NOT NULL ENABLE,DEPTNAME VARCHAR2(100) NOT NULL ENABLE,DEPTSHORT VARCHAR2(10) NOT NULL ENABLE, CONSTRAINT DEPT_CON UNIQUE (DEPTNAME) ENABLE, PRIMARY KEY (DEPTCODE) ENABLE )"
    rsTable.Open strSql, conn, adOpenDynamic, adLockOptimistic, -1
    
    If Err.Number <> 0 Then
        If Err.Number = -2147217900 Then
            SaveSetting App.CompanyName, "DataBase", "DeptTable", True
        Else
            MsgBox "Error:" & Err.Description & vbCrLf & "Error Number:" & Err.Number
        End If
    Else
        SaveSetting App.CompanyName, "DataBase", "DeptTable", True
        MsgBox "Table Dept Created Sucessfully"
    End If
End Sub

Public Sub CreateMasterTable()
    On Error Resume Next
    Dim strSql As String
    Dim rsTable As New ADODB.Recordset
    rsTable.CursorLocation = adUseClient
    strSql = "CREATE TABLE  MASTER (FID NUMBER NOT NULL ENABLE,DEPT NUMBER,BATCH NUMBER,SEM NUMBER,SEC VARCHAR2(1),SUBJCODE VARCHAR2(10),STAFFID VARCHAR2(10),Q1 VARCHAR2(1),Q2 VARCHAR2(1),Q3 VARCHAR2(1),Q4 VARCHAR2(1),Q5 VARCHAR2(1),Q6 VARCHAR2(1),Q7 VARCHAR2(1),Q8 VARCHAR2(1),Q9 VARCHAR2(1),Q10 VARCHAR2(1),Q11 VARCHAR2(1),Q12 VARCHAR2(1),Q13 VARCHAR2(1),Q14 VARCHAR2(1),Q15 VARCHAR2(1),Q16 VARCHAR2(1),Q17 VARCHAR2(1),Q18 VARCHAR2(1),Q19 VARCHAR2(1),Q20 VARCHAR2(1))"
    rsTable.Open strSql, conn, adOpenDynamic, adLockOptimistic, -1
    
    If Err.Number <> 0 Then
        If Err.Number = -2147217900 Then
            SaveSetting App.CompanyName, "DataBase", "MasterTable", True
        Else
            MsgBox "Error:" & Err.Description & vbCrLf & "Error Number:" & Err.Number
        End If
    Else
        SaveSetting App.CompanyName, "DataBase", "MasterTable", True
        MsgBox "Table Master Created Sucessfully"
    End If
End Sub
Public Sub CreateStaffTable()
    On Error Resume Next
    Dim strSql As String
    Dim rsTable As New ADODB.Recordset
    rsTable.CursorLocation = adUseClient
    strSql = "CREATE TABLE STAFF(STAFFID VARCHAR2(10) NOT NULL ENABLE,STAFFNAME VARCHAR2(50) NOT NULL ENABLE,DEPT NUMBER NOT NULL ENABLE,CONSTRAINT STAFF_CON PRIMARY KEY (STAFFID) ENABLE,CONSTRAINT STAFF_CONUNI UNIQUE (STAFFNAME) ENABLE)"
    rsTable.Open strSql, conn, adOpenDynamic, adLockOptimistic, -1
    
    If Err.Number <> 0 Then
        If Err.Number = -2147217900 Then
            SaveSetting App.CompanyName, "DataBase", "StaffTable", True
        Else
            MsgBox "Error:" & Err.Description & vbCrLf & "Error Number:" & Err.Number
        End If
    Else
        SaveSetting App.CompanyName, "DataBase", "StaffTable", True
        MsgBox "Table Staff Created Sucessfully"
    End If
End Sub

Public Sub CreateStaffHandleTable()
    On Error Resume Next
    Dim strSql As String
    Dim rsTable As New ADODB.Recordset
    rsTable.CursorLocation = adUseClient
    strSql = "CREATE TABLE STAFFHANDLE(STAFFID VARCHAR2(10) NOT NULL ENABLE,DEPT NUMBER NOT NULL ENABLE,BATCH NUMBER NOT NULL ENABLE,SEM NUMBER NOT NULL ENABLE,SEC VARCHAR2(1) NOT NULL ENABLE,SUBJCODE VARCHAR2(10) NOT NULL ENABLE)"
    rsTable.Open strSql, conn, adOpenDynamic, adLockOptimistic, -1
    
    If Err.Number <> 0 Then
        If Err.Number = -2147217900 Then
            SaveSetting App.CompanyName, "DataBase", "StaffHandleTable", True
        Else
            MsgBox "Error:" & Err.Description & vbCrLf & "Error Number:" & Err.Number
        End If
    Else
        SaveSetting App.CompanyName, "DataBase", "StaffHandleTable", True
        MsgBox "Table StaffHandle Created Sucessfully"
    End If
End Sub
Public Sub CreateSubjTable()
    On Error Resume Next
    Dim strSql As String
    Dim rsTable As New ADODB.Recordset
    rsTable.CursorLocation = adUseClient
    strSql = "CREATE TABLE  SUBJ(SUBJCODE VARCHAR2(10),SUBJNAME VARCHAR2(50),SEMNO NUMBER,DEPT NUMBER,BATCH NUMBER)"
    rsTable.Open strSql, conn, adOpenDynamic, adLockOptimistic, -1
    
    If Err.Number <> 0 Then
        If Err.Number = -2147217900 Then
            SaveSetting App.CompanyName, "DataBase", "SubjTable", True
        Else
            MsgBox "Error:" & Err.Description & vbCrLf & "Error Number:" & Err.Number
        End If
    Else
        SaveSetting App.CompanyName, "DataBase", "SubjTable", True
        MsgBox "Table Subj Created Sucessfully"
    End If
End Sub
