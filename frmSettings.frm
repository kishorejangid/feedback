VERSION 5.00
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkusercontrolsxp.ocx"
Begin VB.Form frmSettings 
   BorderStyle     =   0  'None
   Caption         =   "Settings"
   ClientHeight    =   4560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5520
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin vkUserContolsXP.vkFrame frameSettings 
      Height          =   4575
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   8070
      Caption         =   "Database Settings"
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
      Begin vkUserContolsXP.vkTextBox txtCity 
         Height          =   375
         Left            =   1320
         TabIndex        =   17
         Top             =   2880
         Width           =   3975
         _ExtentX        =   7011
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
         BorderColor     =   8421504
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkLabel lblAddress 
         Height          =   255
         Left            =   240
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3000
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "City:"
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
      Begin vkUserContolsXP.vkLabel lblCollege 
         Height          =   255
         Left            =   240
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2520
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "College:"
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
      Begin vkUserContolsXP.vkTextBox txtCollege 
         Height          =   375
         Left            =   1320
         TabIndex        =   14
         Top             =   2400
         Width           =   3975
         _ExtentX        =   7011
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
         BorderColor     =   8421504
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkCheck cbUpdate 
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   4200
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Auto Check for Update on StartUp"
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
         TabIndex        =   12
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
      Begin vkUserContolsXP.vkLabel lblDataPassword 
         Height          =   255
         Left            =   240
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1560
         Width           =   975
         _ExtentX        =   1720
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
      Begin vkUserContolsXP.vkTextBox txtDataPassword 
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   1440
         Width           =   3975
         _ExtentX        =   7011
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
         BorderColor     =   8421504
         PassWordChar    =   "*"
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkLabel lblDataUser 
         Height          =   255
         Left            =   240
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
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
      Begin vkUserContolsXP.vkTextBox txtDataUser 
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         Top             =   960
         Width           =   3975
         _ExtentX        =   7011
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
         BorderColor     =   8421504
         LegendForeColor =   16750899
      End
      Begin feedback.StylerButton cmdSave 
         Height          =   375
         Left            =   3960
         TabIndex        =   6
         Top             =   4080
         Width           =   1335
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "Save"
         CaptionDisableColor=   12236471
         CaptionEffectColor=   16777215
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
         RoundedValue    =   25
      End
      Begin vkUserContolsXP.vkLabel lblSource 
         Height          =   255
         Left            =   240
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Data Source:"
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
      Begin vkUserContolsXP.vkTextBox txtDataSource 
         Height          =   375
         Left            =   1320
         TabIndex        =   0
         Top             =   480
         Width           =   3975
         _ExtentX        =   7011
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
         BorderColor     =   8421504
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkTextBox txtAdobe 
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   1920
         Width           =   3975
         _ExtentX        =   7011
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
         BorderColor     =   8421504
         LegendForeColor =   16750899
      End
      Begin vkUserContolsXP.vkLabel lblAdobe 
         Height          =   255
         Left            =   240
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2040
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Adobe Path:"
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
      Begin vkUserContolsXP.vkCheck cbAuto 
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   3480
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Auto Start application on System Start."
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
      Begin vkUserContolsXP.vkCheck cbShutDown 
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   3840
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Shut down computer after each feedback."
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
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    On Error Resume Next
    If txtDataSource.Text <> "" And txtDataUser.Text <> "" And txtDataPassword.Text <> "" Then
        SaveSetting App.CompanyName & "\\feedback", "Oracle", "Data Source", txtDataSource.Text
        SaveSetting App.CompanyName & "\\feedback", "Oracle", "Data User", txtDataUser.Text
        SaveSetting App.CompanyName & "\\feedback", "Oracle", "Data Password", txtDataPassword.Text
        SaveSetting App.CompanyName & "\\feedback", "Paths", "Adobe", txtAdobe.Text
        
        strAdobePath = txtAdobe.Text
        strCollegeName = txtCollege.Text
        strCity = txtCity.Text
        
        SaveSetting App.CompanyName & "\\feedback", "Settings", "College Name", strCollegeName
        SaveSetting App.CompanyName & "\\feedback", "Settings", "City", strCity
        If cbAuto.Value = vbChecked Then
            SaveSetting App.CompanyName & "\\feedback", "Settings", "Auto Start", True
            Set reg = CreateObject("Wscript.Shell")
            reg.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\feedback", App.Path & "\" & App.EXEName & ".exe", "REG_SZ"
        Else
            SaveSetting App.CompanyName & "\\feedback", "Settings", "Auto Start", False
            Set reg = CreateObject("Wscript.Shell")
            reg.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\feedback"
        End If
        If cbShutDown.Value = vbChecked Then
            SaveSetting App.CompanyName & "\\feedback", "Settings", "Shut Down", True
        Else
            SaveSetting App.CompanyName & "\\feedback", "Settings", "Shut Down", False
        End If
        If cbUpdate.Value = vbChecked Then
            SaveSetting App.CompanyName & "\\feedback", "Settings", "Auto Update", True
        Else
            SaveSetting App.CompanyName & "\\feedback", "Settings", "Auto Update", False
        End If
        MsgBox "Settings Saved", vbInformation, "Jangid Corporation"
    End If
End Sub

Private Sub Form_Load()
    frameSettings.Top = 0
    frameSettings.Left = 0
    Me.Width = frameSettings.Width
    Me.Height = frameSettings.Height
    CreateRoundRectFromWindow Me, 7, 7
    frameSettings.Visible = True
    txtDataSource.Text = GetSetting(App.CompanyName & "\\feedback", "Oracle", "Data Source", "student")
    txtDataUser.Text = GetSetting(App.CompanyName & "\\feedback", "Oracle", "Data User", "kishore")
    txtDataPassword.Text = GetSetting(App.CompanyName & "\\feedback", "Oracle", "Data Password", "kishore")
    txtAdobe.Text = GetSetting(App.CompanyName & "\\feedback", "Paths", "Adobe")
    txtCollege.Text = GetSetting(App.CompanyName & "\\feedback", "Settings", "College Name", "ASIATIC TECHNICAL RESEARCH AND DEVELOPMENT CENTRE, RAJAPALAYAM-626117.")
    txtCity.Text = GetSetting(App.CompanyName & "\\feedback", "Settings", "City", "Tirunelveli")
    bAutoStart = GetSetting(App.CompanyName & "\\feedback", "Settings", "Auto Start", False)
    bShutDown = GetSetting(App.CompanyName & "\\feedback", "Settings", "Shut Down", False)
    bAutoUpdate = GetSetting(App.CompanyName & "\\feedback", "Settings", "Auto Update", False)
    If bAutoStart Then
        cbAuto.Value = vbChecked
    End If
    If bShutDown Then
        cbShutDown.Value = vbChecked
    End If
    If bAutoUpdate Then
        cbUpdate.Value = vbChecked
    End If
End Sub
Private Sub frameSettings_MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
