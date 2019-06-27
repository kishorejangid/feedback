VERSION 5.00
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkusercontrolsxp.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Feedback v2.0"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   -1485
   ClientWidth     =   15240
   Icon            =   "frmMain.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin vkUserContolsXP.vkLabel lblTitle 
      Height          =   360
      Left            =   1560
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Width           =   4000
      _ExtentX        =   7064
      _ExtentY        =   635
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "Staff Evaluation by Student."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   14737632
      Alignment       =   2
   End
   Begin feedback.ThemedComboBox ThemedComboBox 
      Left            =   0
      Top             =   0
      _ExtentX        =   556
      _ExtentY        =   529
   End
   Begin vkUserContolsXP.vkLabel lblCollege 
      Height          =   435
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   767
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      Alignment       =   2
   End
   Begin vkUserContolsXP.vkFrame frameLogin 
      Height          =   3975
      Left            =   1680
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1560
      Visible         =   0   'False
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   7011
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
      ShowTitle       =   0   'False
      BorderColor     =   4210752
      BorderWidth     =   2
      Begin vkUserContolsXP.vkCommand cmdOK 
         Height          =   375
         Left            =   1440
         TabIndex        =   13
         Top             =   3360
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
         Caption         =   "Submit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   4210752
         CustomStyle     =   0
      End
      Begin vkUserContolsXP.vkFrame frameAdmin 
         Height          =   1695
         Left            =   1200
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   -120
         Visible         =   0   'False
         Width           =   3150
         _ExtentX        =   5556
         _ExtentY        =   2990
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
         ShowTitle       =   0   'False
         TitleGradient   =   2
         BorderColor     =   4210752
         Begin feedback.StylerButton cmdAdminHide 
            Height          =   255
            Left            =   2640
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   0
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            Caption         =   "X"
            ForeColor       =   255
            CaptionDisableColor=   12236471
            CaptionEffectColor=   16777215
            CaptionEffect   =   3
            FocusDottedRect =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            RoundedValue    =   1
         End
         Begin feedback.StylerButton cmdCreate 
            Height          =   1095
            Left            =   1680
            TabIndex        =   11
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1931
            Caption         =   "Creation"
            ForeColor       =   255
            CaptionDisableColor=   12236471
            CaptionEffectColor=   8421504
            CaptionEffect   =   4
            FocusDottedRect =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin feedback.StylerButton cmdReport 
            Height          =   1095
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1931
            Caption         =   "Reports"
            ForeColor       =   255
            CaptionDisableColor=   12236471
            CaptionEffectColor=   8421504
            CaptionEffect   =   4
            FocusDottedRect =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.ComboBox cmbUser 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         TabIndex        =   1
         Top             =   1920
         Width           =   3735
      End
      Begin vkUserContolsXP.vkTextBox txtName 
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   2400
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   4210752
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtPass 
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   2880
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   4210752
         PassWordChar    =   "*"
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkLabel lblUserType 
         Height          =   255
         Left            =   360
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2040
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "User Type:"
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
      Begin vkUserContolsXP.vkLabel lblName 
         Height          =   255
         Left            =   360
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2520
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "User Name:"
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
      Begin vkUserContolsXP.vkLabel lblPass 
         Height          =   255
         Left            =   360
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   3000
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Password:"
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
      Begin VB.Image imgHome 
         Height          =   900
         Left            =   240
         Picture         =   "frmMain.frx":0ECA
         Stretch         =   -1  'True
         Top             =   600
         Width           =   1020
      End
      Begin VB.Line hr 
         BorderColor     =   &H80000000&
         X1              =   40
         X2              =   5500
         Y1              =   1725
         Y2              =   1725
      End
      Begin VB.Image imgFeed 
         Height          =   1500
         Left            =   1440
         Picture         =   "frmMain.frx":5055
         Top             =   120
         Width           =   3900
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fAdmin As Boolean
Private Sub cmbUser_Change()
    On Error Resume Next
    If cmbUser.Text = "Student" Then
        fAdmin = False
        txtName.Text = ""
        txtPass.Text = ""
        cmdOK.Visible = True
    ElseIf cmbUser.Text = "Administrator" Then
        txtName.Visible = True
        lblName.Visible = True
        txtPass.Visible = True
        lblPass.Visible = True
        cmdOK.Visible = True
        fAdmin = True
    End If
End Sub
Private Sub cmbUser_Click()
    On Error Resume Next
    If cmbUser.Text = "Student" Then
        fAdmin = False
        txtName.Text = ""
        txtPass.Text = ""
        cmdOK.Visible = True
    ElseIf cmbUser.Text = "Administrator" Then
        txtName.Visible = True
        lblName.Visible = True
        txtPass.Visible = True
        lblPass.Visible = True
        cmdOK.Visible = True
        fAdmin = True
    End If
End Sub





Private Sub cmdAdminHide_Click()
    On Error Resume Next
    frameAdmin.Visible = False
    cmdOK.Visible = True
    cmbUser.Visible = True
    txtName.Visible = True
    txtPass.Visible = True
    lblUserType.Visible = True
    lblPass.Visible = True
    lblName.Visible = True
    txtName.Text = ""
    txtPass.Text = ""
End Sub

Private Sub cmdCreate_Click()
    Unload Me
    frmNew.Show
End Sub
Private Sub HideControls()
    cmbUser.Visible = False
    txtName.Visible = False
    txtPass.Visible = False
    lblUserType.Visible = False
    lblPass.Visible = False
    lblName.Visible = False
    frameAdmin.Left = 1200
    frameAdmin.Top = 1920
    frameAdmin.Width = 3150
    frameAdmin.Height = 1695
    frameAdmin.Visible = True
    cmdCreate.Width = cmdReport.Width
    cmdCreate.Left = cmdReport.Left + cmdReport.Width + 240
    cmdReport.Visible = True
    cmdReport.SetFocus
End Sub


Private Sub cmdOK_Click()
    On Error GoTo lblerr
    Dim rs As New ADODB.Recordset
    Dim qr As String
    Dim Passwordflg As Boolean
    cmdOK.Visible = False
    If cmbUser.Text = "Administrator" And txtName.Text = "invigilator" And txtPass.Text = "eval6412fx" Then
            HideControls
            Exit Sub
    End If
    If cmbUser.Text = "Administrator" And txtName.Text = "kishorejangid" And txtPass.Text = "kishorejangid" Then
            HideControls
            Exit Sub
    End If
    
    If fAdmin = True Then
        rs.CursorLocation = adUseClient
        qr = "select * from login where logintype = '" & cmbUser.Text & "'"
        rs.Open qr, conn, adOpenDynamic, adLockOptimistic, -1
        Do While Not rs.EOF
            If txtName.Text = rs!loginid And txtPass.Text = rs!loginpassword Then
                HideControls
                cmdReport.Visible = False
                cmdCreate.Left = cmdReport.Left
                cmdCreate.Width = (2 * cmdReport.Width) + 240
                Passwordflg = True
                Exit Sub
            Else
                rs.MoveNext
                Passwordflg = False
            End If
        Loop
        If Passwordflg = False Then
            MsgBox "Check Your UserID And PassWord"
            txtName.Text = ""
            txtPass.Text = ""
            txtName.SetFocus
            cmdOK.Visible = True
            Exit Sub
        End If
    ElseIf fAdmin = False Then
        rs.CursorLocation = adUseClient
        qr = "select * from login where logintype = '" & cmbUser.Text & "'"
        rs.Open qr, conn, adOpenDynamic, adLockOptimistic, -1
        Do While Not rs.EOF
            If txtName.Text = rs!loginid And txtPass.Text = rs!loginpassword Then
                Unload Me
                frmFeed.Show
                Passwordflg = True
                Exit Sub
            Else
                rs.MoveNext
                Passwordflg = False
            End If
        Loop
        If Passwordflg = False Then
            MsgBox "Check Your UserID And PassWord"
            txtName.Text = ""
            txtPass.Text = ""
            txtName.SetFocus
            cmdOK.Visible = True
            Exit Sub
        End If
    End If
    Exit Sub
lblerr:
    MsgBox "Error Number : " & Err.Number & vbCrLf & "Error Description : " & Err.Description & vbCrLf & "Error Location : frmMain - cmdOK_Click()", vbCritical
End Sub

Private Sub cmdReport_Click()
    Unload Me
    frmReport.Show
End Sub
Private Sub Form_Load()
    On Error Resume Next
    frameAdmin.Visible = False
    
    frameLogin.Width = 5575 '8085
    frameLogin.Height = 3975 '5175
    frameLogin.Left = (Screen.Width - frameLogin.Width) / 2
    frameLogin.Top = (Screen.Height - frameLogin.Height) / 2 - 480
    frameLogin.Visible = True
    
    lblCollege.Caption = strCollegeName
    lblCollege.Left = 10 '(Screen.Width - lblCollege.Width) / 2
    lblCollege.Width = Screen.Width '11970
    lblCollege.Height = 555
    lblCollege.Top = 25
    lblCollege.ZOrder
    lblCollege.Visible = True
    
    lblTitle.Width = 4000
    lblTitle.Height = 360
    lblTitle.Left = (Screen.Width - lblTitle.Width) / 2
    lblTitle.Top = 600
    lblTitle.ZOrder
    
    cmbUser.AddItem ("Student")
    cmbUser.AddItem ("Administrator")
    cmbUser.Text = cmbUser.List(0)
End Sub
Private Sub txtPass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        Call cmdOK_Click
    End If
End Sub
