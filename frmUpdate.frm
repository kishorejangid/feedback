VERSION 5.00
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkusercontrolsxp.ocx"
Begin VB.Form frmUpdate 
   BorderStyle     =   0  'None
   ClientHeight    =   3630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5520
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin vkUserContolsXP.vkFrame frameUpdate 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   6376
      Caption         =   "Feedback Update"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleColor1     =   8421504
      TitleColor2     =   4210752
      TitleGradient   =   2
      TitleHeight     =   360
      BorderColor     =   4210752
      Begin vkUserContolsXP.vkCommand cmdCancel 
         Height          =   375
         Left            =   4080
         TabIndex        =   15
         Top             =   2760
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Cancel"
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
      Begin vkUserContolsXP.vkLabel lblSpeed 
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2520
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Speed:"
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
      Begin vkUserContolsXP.vkLabel lblUpdateSize 
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         Top             =   1680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   ""
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
      Begin vkUserContolsXP.vkLabel lblUpdateDate 
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   ""
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
      Begin vkUserContolsXP.vkLabel lblUpdateVer 
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   ""
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
      Begin vkUserContolsXP.vkCommand cmdUpdate 
         Default         =   -1  'True
         Height          =   375
         Left            =   2640
         TabIndex        =   10
         Top             =   2760
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Update"
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
      Begin vkUserContolsXP.vkLabel lblTime 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2880
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Time Left :"
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
      Begin vkUserContolsXP.vkBar prgUpdate 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   3240
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   450
         BorderColor     =   4210752
         LeftColor       =   12632256
         RightColor      =   4210752
         Value           =   1
         GradientMode    =   1
         ForeColor       =   255
         BackPicture     =   "frmUpdate.frx":0000
         FrontPicture    =   "frmUpdate.frx":001C
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
      Begin vkUserContolsXP.vkLabel lblInfo 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Update Info :"
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
      Begin vkUserContolsXP.vkLabel lblSize 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Update Size :"
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
      Begin vkUserContolsXP.vkLabel lblVer 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Update Version :"
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
      Begin vkUserContolsXP.vkLabel lblDate 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Update Date :"
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
      Begin vkUserContolsXP.vkLabel lblCurrent 
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   344
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Current Version:"
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
      Begin feedback.StylerButton cmdClose 
         Height          =   255
         Left            =   4920
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         Caption         =   "x"
         ForeColor       =   255
         CaptionDisableColor=   12236471
         CaptionEffectColor=   16777215
         CaptionEffect   =   3
         FocusDottedRect =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedValue    =   1
      End
      Begin VB.Label lblFile 
         BackStyle       =   0  'Transparent
         Caption         =   "File Name:"
         Height          =   495
         Left            =   2640
         TabIndex        =   16
         Top             =   960
         Width           =   2775
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblUpdateInfo 
         BackStyle       =   0  'Transparent
         Height          =   735
         Left            =   1440
         TabIndex        =   7
         Top             =   2040
         Width           =   3855
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000006&
         BorderWidth     =   2
         X1              =   0
         X2              =   5500
         Y1              =   840
         Y2              =   840
      End
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mFTP As cFTP
Attribute mFTP.VB_VarHelpID = -1
Private FileName As String
Private Sub cmdCancel_Click()
    CloseUpdate
End Sub
Private Sub cmdClose_Click()
    CloseUpdate
End Sub
Private Sub cmdUpdate_Click()
    On eror GoTo errHan
    Dim lTimer As Long
    Dim strRemote As String
    Dim strLocal As String
    BeginTransfer = Timer
    strRemote = FileName
    strLocal = "C:\feedback_setup.exe"
    lTimer = Timer
    
    Set mFTP = New cFTP
    mFTP.SetModePassive
    mFTP.SetTransferBinary
    
    If mFTP.OpenConnection(SERVER, SERVERUSER, SERVERPASSWORD) Then
        mFTP.SetFTPDirectory SERVERFOLDERPATH
        DoEvents
        frameUpdate.Caption = "Downloading Updates..."
        If Not mFTP.FTPDownloadFile(strLocal, strRemote) Then
            DoEvents
            frameUpdate.Caption = "Error in Updating"
        Else
            frameUpdate.Caption = "Download Complete"
            DoEvents
            RunUpdate "C:\feedback_setup.exe"
            End
        End If
        DoEvents
    End If
    DoEvents
    mFTP.CloseConnection
errHan:
    If mFTP.GetLastErrorMessage = "12007" Then
        frameUpdate.Caption = "Error : Internet Connection not available"
    ElseIf mFTP.GetLastErrorMessage = "12014" Then
        frameUpdate.Caption = "Error : Connection to server failed"
    ElseIf mFTP.GetLastErrorMessage = "12029" Then
        frameUpdate.Caption = "Error : Connection to server failed"
    ElseIf mFTP.GetLastErrorMessage = "12015" Then
        frameUpdate.Caption = "Error : The Login request was denied."
    Else
        frameUpdate.Caption = "Error in Updating"
    End If
End Sub

Public Sub mFTP_FileTransferProgress(lCurrentBytes As Long, lTotalBytes As Long)
    On Error Resume Next
    TransferRate = Format(Int(lCurrentBytes / (Timer - BeginTransfer)) / 1000, "####.00")
    DoEvents
    prgUpdate.Max = lTotalBytes
    prgUpdate.Min = 0
    prgUpdate.Value = lCurrentBytes
    DoEvents
    lblTime.Caption = "Time Left : " & ConvertTime(Int(((prgUpdate.Max - prgUpdate.Value) / 1024) / TransferRate))
    DoEvents
    lblSpeed.Caption = "Transfer Speed :    " & Format(TransferRate, "##.#0#") & " Kbps"
    DoEvents
    prgUpdate.ToolTipText = prgUpdate.Value & " Bytes of " & prgUpdate.Max & " Bytes Transfered"
End Sub
Private Sub Form_Load()
    frameUpdate.Top = 0
    frameUpdate.Left = 0
    Me.Width = frameUpdate.Width
    Me.Height = frameUpdate.Height
    CreateRoundRectFromWindow Me, 7, 7
    frameUpdate.Visible = True
    CurrentAppVer = App.Major & "." & App.Minor & "." & App.Revision
    lblCurrent.Caption = "Current Version :   " & CurrentAppVer
    lblUpdateVer.Caption = UpdateData(1)
    lblUpdateDate.Caption = UpdateData(2)
    lblUpdateSize.Caption = UpdateData(3)
    lblUpdateInfo.Caption = UpdateData(4)
    lblFile.Caption = "File Name : " & UpdateData(5)
    FileName = UpdateData(5)
    frameUpdate.Caption = "Feedback Update Availabe"
End Sub
Private Sub frameUpdate_MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
Private Sub CloseUpdate()
    Unload Me
    Call OpenDataBase
    If fConnSuccess Then
        frmMain.Show
    Else
        frmSettings.Show
    End If
End Sub
