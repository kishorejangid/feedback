VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMsg 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "feedback"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7905
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmMsg.frx":0000
   LinkTopic       =   "frmNew"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   7905
   StartUpPosition =   1  'CenterOwner
   Begin feedback.StylerButton cmdOK 
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   4800
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "OK"
      CaptionDisableColor=   12236471
      CaptionEffectColor=   16777215
      CaptionEffect   =   3
      FocusDottedRect =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshGrid 
      Height          =   4335
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   7400
      _ExtentX        =   13044
      _ExtentY        =   7646
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   12632256
      BackColorBkg    =   16777215
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "The following subjects have not assigned a staff."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   960
      TabIndex        =   0
      Top             =   0
      Width           =   5865
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim rsUn As New ADODB.Recordset
    rsUn.CursorLocation = adUseClient
    Dim strSql As String
    strSql = "select * from subj where subjcode not in (select subjcode from staffhandle) order by dept,batch,semno"
    rsUn.Open strSql, conn, adOpenDynamic, adLockOptimistic, -1
    If rsUn.RecordCount > 0 Then
        Set mshGrid.DataSource = rsUn
        mshGrid.ColWidth(0) = 1000
        mshGrid.ColWidth(1) = 3700
        mshGrid.ColWidth(2) = 750
        mshGrid.ColWidth(3) = 750
        mshGrid.ColWidth(4) = 750
            
        mshGrid.ColAlignment(2) = flexAlignCenterCenter
        mshGrid.ColAlignment(3) = flexAlignCenterCenter
        mshGrid.ColAlignment(4) = flexAlignCenterCenter
    
        mshGrid.ColAlignmentFixed(0) = flexAlignCenterCenter
        mshGrid.ColAlignmentFixed(1) = flexAlignCenterCenter
        mshGrid.ColAlignmentFixed(2) = flexAlignCenterCenter
        mshGrid.ColAlignmentFixed(3) = flexAlignCenterCenter
        mshGrid.ColAlignmentFixed(4) = flexAlignCenterCenter
        mshGrid.RowHeightMin = 300
    End If
End Sub

Private Sub Form_Paint()
    mshGrid.Left = 240
    mshGrid.Width = Me.Width - mshGrid.Left - 330
    cmdOK.SetFocus
End Sub

