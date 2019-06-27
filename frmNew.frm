VERSION 5.00
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkusercontrolsxp.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmNew 
   BackColor       =   &H8000000C&
   Caption         =   "Feedback - New"
   ClientHeight    =   8100
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13620
   Icon            =   "frmNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   13620
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin vkUserContolsXP.vkFrame frameNew 
      Height          =   7695
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   13573
      Caption         =   "Feedback - New"
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
      BorderWidth     =   3
      Begin feedback.StylerButton cmdSettings 
         Height          =   375
         Left            =   30
         TabIndex        =   71
         Top             =   3480
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         Caption         =   "Settings"
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
         RoundedValue    =   1
      End
      Begin feedback.StylerButton cmdUninstall 
         Height          =   375
         Left            =   30
         TabIndex        =   7
         Top             =   3960
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         Caption         =   "UnInstall"
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
         RoundedValue    =   1
      End
      Begin vkUserContolsXP.vkFrame frameDept 
         Height          =   390
         Left            =   11040
         TabIndex        =   55
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   688
         Caption         =   "New Department"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TitleColor1     =   4210752
         TitleColor2     =   14737632
         TitleGradient   =   2
         TitleHeight     =   360
         BorderColor     =   4210752
         RoundAngle      =   2
         BorderWidth     =   2
         Begin feedback.StylerButton cmdDelDept 
            Height          =   375
            Left            =   4080
            TabIndex        =   66
            Top             =   2400
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "Delete"
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
         End
         Begin vkUserContolsXP.vkLabel lblDeptShort 
            Height          =   255
            Left            =   360
            TabIndex        =   62
            Top             =   1920
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Dept Short:"
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
         Begin vkUserContolsXP.vkTextBox txtDeptShort 
            Height          =   375
            Left            =   1440
            TabIndex        =   58
            Top             =   1800
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
            BorderColor     =   4210752
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkLabel lblDeptName 
            Height          =   255
            Left            =   360
            TabIndex        =   61
            Top             =   1320
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Dept Name:"
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
         Begin vkUserContolsXP.vkTextBox txtDeptName 
            Height          =   375
            Left            =   1440
            TabIndex        =   57
            Top             =   1200
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
            BorderColor     =   4210752
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkLabel lblDeptCode 
            Height          =   255
            Left            =   360
            TabIndex        =   60
            Top             =   720
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Dept Code:"
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
         Begin vkUserContolsXP.vkTextBox txtDeptCode 
            Height          =   375
            Left            =   1440
            TabIndex        =   56
            Top             =   600
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
            BorderColor     =   4210752
            LegendForeColor =   16750899
         End
         Begin feedback.StylerButton cmdCreateDept 
            Height          =   375
            Left            =   2640
            TabIndex        =   59
            Top             =   2400
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "Create"
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
         End
         Begin VB.Label lblNote 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Note: Create a seperate department for P.G."
            Height          =   195
            Left            =   360
            TabIndex        =   69
            Top             =   2880
            Width           =   3120
         End
      End
      Begin feedback.StylerButton cmdNewDept 
         Height          =   375
         Left            =   25
         TabIndex        =   1
         Top             =   600
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         Caption         =   "Department"
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
         RoundedValue    =   1
      End
      Begin vkUserContolsXP.vkFrame frameNewUser 
         Height          =   390
         Left            =   9600
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   688
         Caption         =   "New User"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TitleColor1     =   4210752
         TitleColor2     =   14737632
         TitleGradient   =   2
         TitleHeight     =   360
         BorderColor     =   4210752
         RoundAngle      =   2
         BorderWidth     =   2
         Begin vkUserContolsXP.vkTextBox txtRePass 
            Height          =   375
            Left            =   1440
            TabIndex        =   51
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
            BorderColor     =   4210752
            PassWordChar    =   "*"
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkLabel lblRePass 
            Height          =   255
            Left            =   360
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   2520
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Re Enter:"
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
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   1920
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
         Begin vkUserContolsXP.vkTextBox txtPass 
            Height          =   375
            Left            =   1440
            TabIndex        =   50
            Top             =   1800
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
            BorderColor     =   4210752
            PassWordChar    =   "*"
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkLabel lblUserType 
            Height          =   255
            Left            =   360
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   720
            Width           =   975
            _ExtentX        =   1720
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
         Begin VB.ComboBox cmbUserType 
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
            TabIndex        =   48
            Top             =   600
            Width           =   3975
         End
         Begin vkUserContolsXP.vkLabel lblUserName 
            Height          =   255
            Left            =   360
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   1320
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
         Begin vkUserContolsXP.vkTextBox txtUserName 
            Height          =   375
            Left            =   1440
            TabIndex        =   49
            Top             =   1200
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
            BorderColor     =   4210752
            LegendForeColor =   16750899
         End
         Begin feedback.StylerButton cmdCreateUser 
            Height          =   375
            Left            =   4080
            TabIndex        =   52
            Top             =   3000
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "Create"
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
         End
      End
      Begin feedback.StylerButton cmdNewUser 
         Height          =   375
         Left            =   30
         TabIndex        =   6
         Top             =   3000
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         Caption         =   "New User"
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
         RoundedValue    =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshData 
         Height          =   255
         Left            =   120
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   4800
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         FixedCols       =   0
         WordWrap        =   -1  'True
         Appearance      =   0
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin vkUserContolsXP.vkFrame frameProfile 
         Height          =   375
         Left            =   6480
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         Caption         =   "Department Faculty Profile"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TitleColor1     =   4210752
         TitleColor2     =   14737632
         TitleGradient   =   2
         TitleHeight     =   360
         BorderColor     =   4210752
         RoundAngle      =   2
         BorderWidth     =   2
         Begin feedback.StylerButton cmdLoad 
            Height          =   255
            Left            =   240
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   360
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            Caption         =   "Load"
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
            RoundedValue    =   1
         End
         Begin feedback.StylerButton cmdEx 
            Height          =   255
            Left            =   960
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            Caption         =   "+"
            ForeColor       =   33023
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
         Begin MSComctlLib.TreeView tvProfile 
            Height          =   5895
            Left            =   240
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   600
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   10398
            _Version        =   393217
            Style           =   7
            Appearance      =   1
         End
      End
      Begin vkUserContolsXP.vkFrame frameHandle 
         Height          =   375
         Left            =   3360
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         Caption         =   "Assigning Subjects to Staff"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TitleColor1     =   4210752
         TitleColor2     =   12632256
         TitleGradient   =   2
         TitleHeight     =   360
         BorderColor     =   4210752
         RoundAngle      =   2
         BorderWidth     =   2
         Begin feedback.StylerButton cmdHandleDelete 
            Height          =   375
            Left            =   4080
            TabIndex        =   67
            Top             =   4680
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "Delete"
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
         End
         Begin vkUserContolsXP.vkLabel lblHandleSec 
            Height          =   255
            Left            =   360
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   4320
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Section:"
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
         Begin VB.ComboBox cmbHandleSec 
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
            TabIndex        =   36
            Top             =   4200
            Width           =   1935
         End
         Begin vkUserContolsXP.vkLabel lblHandleSubj 
            Height          =   255
            Left            =   360
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   3720
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Subject:"
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
         Begin VB.ComboBox cmbHandleSubj 
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
            TabIndex        =   35
            Top             =   3600
            Width           =   1935
         End
         Begin VB.ComboBox cmbHandleSubjDept 
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
            TabIndex        =   32
            Top             =   1800
            Width           =   3975
         End
         Begin vkUserContolsXP.vkLabel lblHandleSubjDept 
            Height          =   255
            Left            =   360
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   1920
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Subject Dept:"
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
         Begin VB.ComboBox cmbHandleStaff 
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
            TabIndex        =   31
            Top             =   1200
            Width           =   3975
         End
         Begin vkUserContolsXP.vkLabel lblHandleStaff 
            Height          =   255
            Left            =   360
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   1320
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Staff Name:"
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
         Begin feedback.StylerButton cmdCreateHandle 
            Height          =   375
            Left            =   2640
            TabIndex        =   37
            Top             =   4680
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "Create"
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
         End
         Begin VB.ComboBox cmbHandleStaffDept 
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
            TabIndex        =   30
            Top             =   600
            Width           =   3975
         End
         Begin vkUserContolsXP.vkLabel lblHandleStaffDept 
            Height          =   255
            Left            =   360
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   720
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Staff Dept:"
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
         Begin vkUserContolsXP.vkLabel lblHandleSubjBatch 
            Height          =   255
            Left            =   360
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   2520
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Subj Batch:"
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
         Begin VB.ComboBox cmbHandleSubjBatch 
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
            TabIndex        =   33
            Top             =   2400
            Width           =   1935
         End
         Begin vkUserContolsXP.vkLabel lblHandleSubjSem 
            Height          =   255
            Left            =   360
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   3120
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Subj Semester:"
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
         Begin VB.ComboBox cmbHandleSubjSem 
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
            TabIndex        =   34
            Top             =   3000
            Width           =   1935
         End
      End
      Begin vkUserContolsXP.vkFrame frameSubject 
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "New Subject"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TitleColor1     =   4210752
         TitleColor2     =   14737632
         TitleGradient   =   2
         TitleHeight     =   360
         BorderColor     =   4210752
         RoundAngle      =   2
         BorderWidth     =   2
         Begin feedback.StylerButton cmdDeleteSubj 
            Height          =   375
            Left            =   4080
            TabIndex        =   65
            Top             =   3600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "Delete"
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
         End
         Begin vkUserContolsXP.vkTextBox txtSubjName 
            Height          =   375
            Left            =   1440
            TabIndex        =   19
            Top             =   3000
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
            BorderColor     =   4210752
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtSubjCode 
            Height          =   375
            Left            =   1440
            TabIndex        =   18
            Top             =   2400
            Width           =   1935
            _ExtentX        =   3413
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
         Begin vkUserContolsXP.vkLabel lblSubjSubjName 
            Height          =   255
            Left            =   360
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   3120
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Subject Name:"
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
         Begin vkUserContolsXP.vkLabel lblSubjSubjCode 
            Height          =   255
            Left            =   360
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   2520
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Subject Code:"
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
         Begin VB.ComboBox cmbSubjSem 
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
            TabIndex        =   17
            Top             =   1800
            Width           =   1935
         End
         Begin vkUserContolsXP.vkLabel lblSubjSem 
            Height          =   255
            Left            =   360
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   1920
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Semester:"
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
         Begin VB.ComboBox cmbSubjBatch 
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
            TabIndex        =   16
            Top             =   1200
            Width           =   1935
         End
         Begin vkUserContolsXP.vkLabel lblSubjBatch 
            Height          =   255
            Left            =   360
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   1320
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Batch:"
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
         Begin vkUserContolsXP.vkLabel lblSubjDept 
            Height          =   255
            Left            =   360
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   720
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Department:"
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
         Begin VB.ComboBox cmbSubjDept 
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
            TabIndex        =   15
            Top             =   600
            Width           =   3975
         End
         Begin feedback.StylerButton cmdCreateSubject 
            Height          =   375
            Left            =   2640
            TabIndex        =   20
            Top             =   3600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "Create"
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
         End
         Begin VB.Label lblSubNote 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Note:  You are requested not to add Lab Courses for this feedback."
            Height          =   195
            Left            =   360
            TabIndex        =   70
            Top             =   4080
            Width           =   4755
         End
      End
      Begin vkUserContolsXP.vkFrame frameStaff 
         Height          =   390
         Left            =   120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   688
         Caption         =   "New Faculty"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TitleColor1     =   4210752
         TitleColor2     =   14737632
         TitleGradient   =   2
         TitleHeight     =   360
         BorderColor     =   4210752
         RoundAngle      =   2
         BorderWidth     =   2
         Begin feedback.StylerButton cmdDeleteStaff 
            Height          =   375
            Left            =   4080
            TabIndex        =   64
            Top             =   1800
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "Delete"
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
         End
         Begin feedback.StylerButton cmdCreateStaff 
            Height          =   375
            Left            =   2640
            TabIndex        =   12
            Top             =   1800
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "Create"
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
         End
         Begin vkUserContolsXP.vkTextBox txtStaffName 
            Height          =   375
            Left            =   1440
            TabIndex        =   11
            Top             =   1200
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
            BorderColor     =   4210752
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkLabel lblStaffName 
            Height          =   255
            Left            =   360
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1320
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Staff Name:"
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
         Begin VB.ComboBox cmbStaffDept 
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
            TabIndex        =   10
            Top             =   600
            Width           =   3975
         End
         Begin vkUserContolsXP.vkLabel lblStaffDept 
            Height          =   255
            Left            =   360
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   720
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Department:"
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
      Begin feedback.StylerButton cmdFaculty 
         Height          =   375
         Left            =   30
         TabIndex        =   5
         Top             =   2520
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         Caption         =   "Faculty Profile"
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
         RoundedValue    =   1
      End
      Begin feedback.StylerButton cmdHandle 
         Height          =   375
         Left            =   30
         TabIndex        =   4
         Top             =   2040
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         Caption         =   "Staff Handle"
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
         RoundedValue    =   1
      End
      Begin feedback.StylerButton cmdStaff 
         Height          =   375
         Left            =   30
         TabIndex        =   3
         Top             =   1560
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         Caption         =   "Staff"
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
         RoundedValue    =   1
      End
      Begin feedback.StylerButton cmdSubject 
         Height          =   375
         Left            =   30
         TabIndex        =   2
         Top             =   1080
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         Caption         =   "Subject"
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
         RoundedValue    =   1
      End
   End
   Begin feedback.ThemedComboBox ThemedComboBox1 
      Left            =   0
      Top             =   0
      _ExtentX        =   556
      _ExtentY        =   529
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tvExpanded As Boolean
Dim strSubjCode(10) As String
Private Sub cmbHandleStaffDept_Change()
    On Error Resume Next
    Dim rsStaffName As New ADODB.Recordset
    Dim strSql As String
    strSql = "select staffname from staff where dept='" & Department(cmbHandleStaffDept) & "'"
    rsStaffName.Open strSql, conn, adOpenDynamic, adLockOptimistic, -1
    rsStaffName.MoveFirst
    cmbHandleStaff.Clear
    cmbHandleStaff.FontSize = 8
    While Not rsStaffName.EOF
        cmbHandleStaff.AddItem rsStaffName.Fields(0)
        rsStaffName.MoveNext
    Wend
    cmbHandleStaff.Text = cmbHandleStaff.List(0)
    LoadGridForHandle
End Sub
Private Sub cmbHandleStaffDept_Click()
    On Error Resume Next
    Dim rsStaffName As New ADODB.Recordset
    Dim strSql As String
    strSql = "select staffname from staff where dept='" & Department(cmbHandleStaffDept) & "'"
    rsStaffName.Open strSql, conn, adOpenDynamic, adLockOptimistic, -1
    rsStaffName.MoveFirst
    cmbHandleStaff.Clear
    cmbHandleStaff.FontSize = 8
    While Not rsStaffName.EOF
        cmbHandleStaff.AddItem rsStaffName.Fields(0)
        rsStaffName.MoveNext
    Wend
    cmbHandleStaff.Text = cmbHandleStaff.List(0)
    LoadGridForHandle
End Sub

Private Sub cmbHandleSubjBatch_Change()
    On Error Resume Next
    iBatch = cmbHandleSubjBatch.Text
    cmbSubj_Load cmbHandleSubj
End Sub
Private Sub cmbHandleSubjBatch_Click()
    On Error Resume Next
    iBatch = cmbHandleSubjBatch.Text
    cmbSubj_Load cmbHandleSubj
End Sub
Private Sub cmbHandleSubjDept_Change()
    On Error Resume Next
    iDept = Department(cmbHandleSubjDept)
    cmbSubj_Load cmbHandleSubj
End Sub
Private Sub cmbHandleSubjDept_Click()
    On Error Resume Next
    iDept = Department(cmbHandleSubjDept)
    cmbSubj_Load cmbHandleSubj
End Sub

Private Sub cmbHandleSubjSem_Change()
    On Error Resume Next
    iSem = cmbHandleSubjSem.Text
    cmbSubj_Load cmbHandleSubj
End Sub
Private Sub cmbHandleSubjSem_Click()
    On Error Resume Next
    iSem = cmbHandleSubjSem.Text
    cmbSubj_Load cmbHandleSubj
End Sub

Private Sub cmbStaffDept_Change()
    On Error Resume Next
    iDept = Department(cmbStaffDept)
    LoadGridForStaff
End Sub
Private Sub cmbStaffDept_Click()
    On Error Resume Next
    iDept = Department(cmbStaffDept)
    LoadGridForStaff
End Sub

Private Sub cmbSubjDept_Change()
    LoadGridForSubject
End Sub
Private Sub cmbSubjDept_Click()
    LoadGridForSubject
End Sub

Private Sub cmdCreateDept_Click()
    On Error GoTo errHan
    If txtDeptCode.Text = "" Then
        MsgBox "Enter Department code.", vbInformation, "Jangid Corporation"
        txtDeptCode.SetFocus
        Exit Sub
    End If
    If txtDeptName.Text = "" Then
        MsgBox "Enter Department name.", vbInformation, "Jangid Corporation"
        txtDeptName.SetFocus
        Exit Sub
    End If
    If txtDeptShort.Text = "" Then
        MsgBox "Enter Department short name.", vbInformation, "Jangid Corporation"
        txtDeptShort.SetFocus
        Exit Sub
    End If
    Dim rsNewdept As New ADODB.Recordset
    Dim strsqlDept As String
    strsqlDept = "insert into dept values ('" & txtDeptCode.Text & "','" & Trim(UCase(txtDeptName.Text)) & "','" & Trim(UCase(txtDeptShort.Text)) & "')"
    rsNewdept.Open strsqlDept, conn, adOpenDynamic, adLockOptimistic, -1
    If Err.Number = 0 Then
        MsgBox "Department added successfully", vbInformation, "Jangid Corporation"
        txtDeptCode.Text = ""
        txtDeptName.Text = ""
        txtDeptShort.Text = ""
        txtDeptCode.SetFocus
        LoadGridForDept
    End If
    Exit Sub
errHan:
    If Err.Number <> 0 Then
        If Err.Number = -2147217873 Then
            MsgBox "Deparment already exists", vbExclamation, "Jangid Corporation"
            txtDeptCode.Text = ""
            txtDeptName.Text = ""
            txtDeptShort.Text = ""
            txtDeptCode.SetFocus
        Else
            MsgBox "Error Number: " & Err.Number & vbCrLf & "Error Description : " & Err.Description & vbCrLf & "Error Location : frmNew-cmdCreateDept_Click()", vbExclamation, "Jangid Corporation"
        End If
    End If
End Sub

Private Sub cmdCreateHandle_Click()
    On Error GoTo errHandle
    If cmbHandleStaffDept.Text = "" Or cmbHandleStaff.Text = "" Or cmbHandleSubj.Text = "" Or cmbHandleSubjDept.Text = "" Or cmbHandleSubjBatch.Text = "" Or cmbHandleSubjSem.Text = "" Or cmbHandleSubj.Text = "" Then
        MsgBox "Data missing in some fields.", vbExclamation, "Jangid Corporation"
        Exit Sub
    End If
    Dim strSql As String
    Dim rsHandle As New ADODB.Recordset
    strSql = "insert into staffhandle (staffid,dept,batch,sem,sec,subjcode) values ('" & getStaffID(cmbHandleStaff.Text, Department(cmbHandleStaffDept)) & "','" & iDept & "','" & Mid(iBatch, 3, 2) & "','" & iSem & "','" & cmbHandleSec.Text & "','" & cmbHandleSubj.Text & "')"
    rsHandle.Open strSql, conn, adOpenDynamic, adLockOptimistic, -1
    If Err.Number = 0 Then
       MsgBox "Subject assigned to the staff successfully.", vbInformation, "Feedback"
       LoadGridForHandle
       cmbHandleStaffDept.SetFocus
    End If
    Exit Sub
errHandle:
    If Err.Number <> 0 Then
        MsgBox "Error Number: " & Err.Number & vbCrLf & "Error Description : " & Err.Description & vbCrLf & "Error Location : frmNew-cmdCreateHandle_Click()", vbExclamation, "Jangid Corporation"
    End If
End Sub

Private Sub cmdCreateStaff_Click()
    On Error GoTo errHan
    If txtStaffName.Text = "" Then
        MsgBox "Enter Staff Name", vbInformation, "Feedback"
        txtStaffName.SetFocus
        Exit Sub
    End If
    Dim rsStaff As New ADODB.Recordset
    Dim strSql As String
    strSql = "insert into staff values('" & getNewStaffID & "','" & Trim(UCase(txtStaffName.Text)) & "','" & iDept & "')"
    rsStaff.Open strSql, conn, adOpenDynamic, adLockOptimistic, -1
    If Err.Number = 0 Then
        MsgBox "Faculty Added Succesfully.", vbInformation, "Feedback"
        txtStaffName.Text = ""
        cmbStaffDept.SetFocus
        LoadGridForStaff
    End If
    Exit Sub
errHan:
    If Err.Number <> 0 Then
        If Err.Number = -2147217873 Then
            MsgBox "Staff already exists", vbInformation, "Jangid Corporation"
            txtStaffName.Text = ""
            cmbStaffDept.SetFocus
        Else
            MsgBox "Error Number: " & Err.Number & vbCrLf & "Error Description : " & Err.Description & vbCrLf & "Error Location : frmNew-cmdCreateStaff_Click()", vbExclamation, "Jangid Corporation"
        End If
    End If
End Sub

Private Sub cmdCreateSubject_Click()
    On Error GoTo errHan
    If cmbSubjDept.Text = "" Or cmbSubjBatch.Text = "" Or cmbSubjSem.Text = "" Then
        MsgBox "Some Fields are missing.", vbInformation, "Jangid Corporation"
        Exit Sub
    End If
    If txtSubjCode.Text = "" Then
        MsgBox "Enter Subject Code.", vbInformation, "Feedback"
        txtSubjCode.SetFocus
        Exit Sub
    End If
    If txtSubjName.Text = "" Then
        MsgBox "Enter Subject Name.", vbInformation, "Feedback"
        txtSubjName.SetFocus
        Exit Sub
    End If
    Dim rsSubject As New ADODB.Recordset
    Dim strSql As String
    strSql = "insert into subj values('" & Trim(UCase(txtSubjCode.Text)) & "','" & Trim(UCase(txtSubjName.Text)) & "','" & Trim(cmbSubjSem.Text) & "','" & Department(cmbSubjDept) & "','" & Mid(cmbSubjBatch.Text, 3, 2) & "')"
    rsSubject.Open strSql, conn, adOpenDynamic, adLockOptimistic, -1
    If Err.Number = 0 Then
        MsgBox "New subject added successfully.", vbInformation, "Feedback"
        txtSubjCode.Text = ""
        txtSubjName.Text = ""
        cmbSubjDept.SetFocus
        LoadGridForSubject
    End If
errHan:
    If Err.Number <> 0 Then
        If Err.Number = -2147217873 Then
            MsgBox "The given subject already exist in the database", vbInformation, "Jangid Corporation"
            txtSubjCode.Text = ""
            txtSubjName.Text = ""
            cmbSubjDept.SetFocus
        Else
            MsgBox "Error Number: " & Err.Number & vbCrLf & "Error Description : " & Err.Description & vbCrLf & "Error Location : frmNew-cmdCreateSubject_Click()", vbExclamation, "Jangid Corporation"
        End If
    End If
End Sub

Private Sub cmdCreateUser_Click()
    On Error GoTo errHan
    If cmbUserType.Text = "" Then
        MsgBox "Select User Type from the Combo box.", vbInformation, "Jangid Corporation"
        cmbUserType.SetFocus
        Exit Sub
    End If
    If txtUserName.Text = "" Then
        MsgBox "Enter user name.", vbInformation, "Jangid Corporation"
        txtUserName.SetFocus
        Exit Sub
    End If
    If txtPass.Text = "" Then
        MsgBox "Enter the password.", vbInformation, "Jangid Corporation"
        txtPass.SetFocus
        Exit Sub
    End If
    If txtRePass.Text = "" Then
        MsgBox "Re Enter the password.", vbInformation, "Jangid Corporation"
        txtRePass.SetFocus
        Exit Sub
    End If
    If txtPass.Text = txtRePass.Text Then
        Dim rsUser As New ADODB.Recordset
        Dim strSql As String
        strSql = "insert into login values('" & cmbUserType.Text & "','" & Trim(txtUserName.Text) & "','" & Trim(txtPass.Text) & "')"
        rsUser.Open strSql, conn, adOpenDynamic, adLockOptimistic, -1
    Else
        MsgBox "Password doesn't match. Retry", vbCritical, "Jangid Corporation"
        Exit Sub
    End If
errHan:
    If Err.Number = 0 Then
        MsgBox "User Created Successfully"
        txtUserName.Text = ""
        txtPass.Text = ""
        cmbUserType.SetFocus
    Else
        If Err.Number = -2147217873 Then
            MsgBox "User already exist.", vbCritical, "Jangid Corporation"
        Else
            MsgBox "Error Number: " & Err.Number & vbCrLf & "Error Description : " & Err.Description & vbCrLf & "Error Location : frmNew-cmdCreateUser_Click()", vbExclamation, "Jangid Corporation"
        End If
    End If
End Sub

Private Sub cmdDelDept_Click()
    On Error Resume Next
    If txtDeptCode.Text = "" Then
        MsgBox "Enter Department Code as you entered during creation.", vbInformation, "Jangid Corporation - Delete Department"
        Exit Sub
    End If
    Dim rsDelDept As New ADODB.Recordset
    rsDelDept.Open "delete from dept where deptcode='" & Val(txtDeptCode.Text) & "'", conn, adOpenDynamic, adLockOptimistic, -1
    If Err.Number = 0 Then
        MsgBox "Department deleted successfully.", vbInformation, "Jangid Corporation"
        txtDeptCode.Text = ""
        txtDeptName.Text = ""
        txtDeptShort.Text = ""
        LoadGridForDept
    Else
        MsgBox "Error Number: " & Err.Number & vbCrLf & "Error Description : " & Err.Description & vbCrLf & "Error Location : frmNew-cmdDelDept_Click()", vbExclamation, "Jangid Corporation"
    End If
End Sub

Private Sub cmdDeleteStaff_Click()
    On Error Resume Next
    If txtStaffName.Text = "" Then
        MsgBox "Enter Staff name", vbInformation, "Jangid Corporation"
        Exit Sub
    End If
    Dim rsDel As New ADODB.Recordset
    rsDel.Open "delete from staff where staffname='" & Trim(UCase(txtStaffName.Text)) & "'", conn, adOpenDynamic, adLockOptimistic, -1
    If Err.Number = 0 Then
        MsgBox "Staff " & txtStaffName.Text & " deleted successfully"
        txtStaffName.Text = ""
        LoadGridForStaff
    Else
        MsgBox "Error Number: " & Err.Number & vbCrLf & "Error Description : " & Err.Description & vbCrLf & "Error Location : frmNew-cmdDeleteStaff_Click()", vbExclamation, "Jangid Corporation"
    End If
End Sub

Private Sub cmdDeleteSubj_Click()
    On Error Resume Next
    If txtSubjCode.Text = "" Then
        MsgBox "Enter subject Code.", vbInformation
        Exit Sub
    End If
    Dim rsDelSubj As New ADODB.Recordset
    Dim strDelSql As String
    strDelSql = "delete from subj where dept='" & Department(cmbSubjDept) & "' and batch='" & Mid(cmbSubjBatch.Text, 3, 2) & "' and semno='" & cmbSubjSem.Text & "' and subjcode='" & UCase(txtSubjCode.Text) & "'"
    rsDelSubj.Open strDelSql, conn, adOpenDynamic, adLockOptimistic, -1
    If Err.Number = 0 Then
        MsgBox "Subject deleted sucessfully"
        txtSubjCode.Text = ""
        txtSubjName.Text = ""
        LoadGridForSubject
    Else
        MsgBox "Error Number: " & Err.Number & vbCrLf & "Error Description : " & Err.Description & vbCrLf & "Error Location : frmNew-cmdDeleteSubj_Click()", vbExclamation, "Jangid Corporation"
    End If
End Sub


Private Sub cmdEx_Click()
    On Error Resume Next
    Dim nodExpand As Node
    If tvExpanded = False Then
        cmdEx.Caption = "--"
        tvExpanded = True
        For Each nodExpand In tvProfile.Nodes
            nodExpand.Expanded = True
        Next
    Else
        cmdEx.Caption = "+"
        tvExpanded = False
        For Each nodExpand In tvProfile.Nodes
            nodExpand.Expanded = False
        Next
    End If
End Sub

Private Sub cmdFaculty_Click()
    On Error Resume Next
    mshData.Visible = False
    
    cmdSubject.Width = 1200
    cmdHandle.Width = 1200
    cmdStaff.Width = 1200
    cmdNewUser.Width = 1200
    cmdNewDept.Width = 1200
    cmdFaculty.Width = 1500
        
    frameHandle.Visible = False
    frameStaff.Visible = False
    frameSubject.Visible = False
    frameNewUser.Visible = False
    frameDept.Visible = False
    frameProfile.Visible = True
    
    cmdEx.Caption = "+"
    tvExpanded = False
End Sub
Private Sub FacultyProfile()
    On Error Resume Next
    tvProfile.Nodes.Clear
    tvProfile.Nodes.Add , , "Root", "Departments"
    
    Dim rsDept As New ADODB.Recordset
    Dim rsStaff As New ADODB.Recordset
    Dim rsSubj As New ADODB.Recordset
    Dim rsSec As New ADODB.Recordset
    
    rsDept.CursorLocation = adUseClient
    rsStaff.CursorLocation = adUseClient
    rsSubj.CursorLocation = adUseClient
    rsSec.CursorLocation = adUseClient
    
    rsDept.Open "select deptshort,deptcode from dept", conn, adOpenDynamic, adLockOptimistic, -1
    rsDept.MoveFirst
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    For i = 0 To rsDept.RecordCount - 1
        tvProfile.Nodes.Add "Root", tvwChild, rsDept.Fields(0), rsDept.Fields(0)
        rsStaff.Open "select staffname,staffid from staff where dept='" & rsDept.Fields(1) & "'", conn, adOpenDynamic, adLockOptimistic, -1
        rsStaff.MoveFirst
        
        For j = 0 To rsStaff.RecordCount - 1
            tvProfile.Nodes.Add CStr(rsDept.Fields(0)), tvwChild, rsStaff.Fields(0), rsStaff.Fields(0)
            rsSubj.Open "select h.subjcode,d.deptshort,h.sem,h.sec from staffhandle h,dept d where d.deptcode = h.dept and h.staffid='" & rsStaff.Fields(1) & "'", conn, adOpenDynamic, adLockOptimistic, -1
            
            rsSubj.MoveFirst
            
            For k = 0 To rsSubj.RecordCount - 1
                tvProfile.Nodes.Add CStr(rsStaff.Fields(0)), tvwChild, rsSubj.Fields(0) & j & k, Format(rsSubj.Fields(0), "!@@@@@@@@@@") & Format(rsSubj.Fields(1), "!@@@@@@@@@@") & Format("Sem " & rsSubj.Fields(2), "!@@@@@@@@@@") & "Sec " & rsSubj.Fields(3)
                rsSubj.MoveNext
            Next
            rsStaff.MoveNext
            rsSubj.Close
            
        Next
        rsDept.MoveNext
        rsStaff.Close
        
    Next
    If Err.Number <> 0 Then
        If Err.Number = 3021 Then
            MsgBox "Insufficient data in the database." & vbCrLf & "Check whether department,staff and subject are created.", vbInformation, "Jangid Corporation"
        Else
            MsgBox "Error Number: " & Err.Number & vbCrLf & "Error Description : " & Err.Description & vbCrLf & "Error Location : frmNew-FacultyProfile()", vbExclamation, "Jangid Corporation"
        End If
    End If
End Sub

Private Sub cmdHandle_Click()
    On Error Resume Next
    cmdSubject.Width = 1200
    cmdFaculty.Width = 1200
    cmdStaff.Width = 1200
    cmdNewUser.Width = 1200
    cmdNewDept.Width = 1200
    cmdHandle.Width = 1500
    
    frameStaff.Visible = False
    frameSubject.Visible = False
    frameProfile.Visible = False
    frameNewUser.Visible = False
    frameDept.Visible = False
    frameHandle.Visible = True
    
    cmbDept_Load cmbHandleStaffDept
    cmbDept_Load cmbHandleSubjDept
    cmbBatch_Load cmbHandleSubjBatch
    cmbSem_Load cmbHandleSubjSem
    cmbSec_Load cmbHandleSec
    cmbSubj_Load cmbHandleSubj
    
    LoadGridForHandle
    cmbHandleStaffDept.SetFocus
    CheckUnHandled
End Sub
Private Sub CheckUnHandled()
    On Error Resume Next
    Dim rsUn As New ADODB.Recordset
    rsUn.CursorLocation = adUseClient
    Dim strSql As String
    strSql = "select * from subj where subjcode not in (select subjcode from staffhandle)"
    rsUn.Open strSql, conn, adOpenDynamic, adLockOptimistic, -1
    If rsUn.RecordCount > 0 Then
        frmMsg.Show
    End If
    If Err.Number <> 0 Then
        MsgBox "Error Number: " & Err.Number & vbCrLf & "Error Description : " & Err.Description & vbCrLf & "Error Location : frmNew-CheckUnHandledSubj()", vbExclamation, "Jangid Corporation"
    End If
End Sub


Private Sub cmdHandleDelete_Click()
    On Error GoTo errHan
    Dim rsHandleDelete As New ADODB.Recordset
    Dim strSql As String
        If cmbHandleStaffDept.Text = "" Or cmbHandleStaff.Text = "" Or cmbHandleSubj.Text = "" Or cmbHandleSubjDept.Text = "" Or cmbHandleSubjBatch.Text = "" Or cmbHandleSubjSem.Text = "" Or cmbHandleSubj.Text = "" Then
        MsgBox "Data missing in some fields.", vbExclamation, "Jangid Corporation"
        Exit Sub
    End If
    Dim response As VbMsgBoxResult
    confirm = MsgBox("Do you want to delete the Subject Handle to the Staff", vbInformation + vbYesNo, "Jangid Corporation")
    strSql = "delete from staffhandle where dept='" & Department(cmbHandleSubjDept) & "' and batch='" & Mid(cmbHandleSubjBatch.Text, 3, 2) & "' and sem='" & cmbHandleSubjSem.Text & "' and sec='" & cmbHandleSec.Text & "' and subjcode='" & cmbHandleSubj.Text & "' and staffid='" & getStaffID(cmbHandleStaff.Text, Department(cmbHandleStaffDept)) & "'"
    If confirm = vbYes Then
        rsHandleDelete.Open strSql, conn, adOpenDynamic, adLockOptimistic, -1
    Else
        Exit Sub
    End If
    If Err.Number = 0 Then
        MsgBox "Deleted Successfully."
        LoadGridForHandle
    End If
    Exit Sub
errHan:
    If Err.Number <> 0 Then
        MsgBox "Error Number: " & Err.Number & vbCrLf & "Error Description : " & Err.Description & vbCrLf & "Error Location : frmNew-cmdHandleUpdate()", vbExclamation, "Jangid Corporation"
    End If
End Sub



Private Sub cmdLoad_Click()
    FacultyProfile
End Sub

Private Sub cmdNewDept_Click()
    On Error Resume Next
    cmdSubject.Width = 1200
    cmdHandle.Width = 1200
    cmdFaculty.Width = 1200
    cmdStaff.Width = 1200
    cmdNewUser.Width = 1200
    cmdNewDept.Width = 1500
    
    frameHandle.Visible = False
    frameStaff.Visible = False
    frameSubject.Visible = False
    frameProfile.Visible = False
    frameNewUser.Visible = False
    frameDept.Visible = True
    
    txtDeptCode.SetFocus
    LoadGridForDept
    
    'MsgBox "Create a seperate department for Post Graduates (P.G)", vbInformation, "Feedback Tips"
End Sub
Private Sub LoadGridForDept()
    On Error Resume Next
    Dim rsDeptData As New ADODB.Recordset
    rsDeptData.Open "select * from dept order by deptcode", conn, adOpenDynamic, adLockOptimistic, -1
    mshData.Clear
    Set mshData.DataSource = rsDeptData
    mshData.Visible = True
    mshData.Left = frameDept.Left + frameDept.Width + 240
    mshData.Width = frameNew.Width - mshData.Left - 360
    mshData.Top = frameDept.Top
    mshData.Height = frameNew.Height - mshData.Top - 360
    
    mshData.ColWidth(0) = 1200
    mshData.ColWidth(1) = 4000
    mshData.ColWidth(2) = 1200
    
    
    mshData.ColAlignment(0) = flexAlignCenterCenter
    mshData.ColAlignment(2) = flexAlignCenterCenter
    
    
    mshData.ColAlignmentFixed(0) = flexAlignCenterCenter
    mshData.ColAlignmentFixed(1) = flexAlignCenterCenter
    mshData.ColAlignmentFixed(2) = flexAlignCenterCenter
    If Err.Number <> 0 Then
        If Err.Number = 30004 Then
            'Invalid Column alignment error
        ElseIf Err.Number = 30022 Then
            'The Hierarchical FlexGrid does not support the requested type of data binding.
        ElseIf Err.Number = 3021 Then
        Else
            MsgBox "Error Number: " & Err.Number & vbCrLf & "Error Description : " & Err.Description & vbCrLf & "Error Location : frmNew-LoadGridForDept()", vbExclamation, "Jangid Corporation"
        End If
    End If
End Sub


Private Sub cmdNewUser_Click()
    On Error Resume Next
    cmdSubject.Width = 1200
    cmdHandle.Width = 1200
    cmdFaculty.Width = 1200
    cmdStaff.Width = 1200
    cmdNewDept.Width = 1200
    cmdNewUser.Width = 1500
    
    frameDept.Visible = False
    frameHandle.Visible = False
    frameStaff.Visible = False
    frameSubject.Visible = False
    frameProfile.Visible = False
    frameNewUser.Visible = True
    
    cmbUserType.AddItem ("Administrator")
    cmbUserType.AddItem ("Student")
    cmbUserType.ListIndex = 0
    cmbUserType.SetFocus
End Sub




Private Sub cmdSettings_Click()
    frmSettings.Show Modal, frmNew
End Sub

Private Sub cmdStaff_Click()
    On Error Resume Next
    cmdSubject.Width = 1200
    cmdHandle.Width = 1200
    cmdFaculty.Width = 1200
    cmdNewUser.Width = 1200
    cmdNewDept.Width = 1200
    cmdStaff.Width = 1500
    
    
    frameSubject.Visible = False
    frameHandle.Visible = False
    frameProfile.Visible = False
    frameNewUser.Visible = False
    frameDept.Visible = False
    frameStaff.Visible = True
    
    cmbDept_Load cmbStaffDept
    
    LoadGridForStaff
    cmbStaffDept.SetFocus
End Sub

Private Sub cmdSubject_Click()
    On Error Resume Next
    cmdHandle.Width = 1200
    cmdFaculty.Width = 1200
    cmdStaff.Width = 1200
    cmdNewUser.Width = 1200
    cmdNewDept.Width = 1200
    cmdSubject.Width = 1500
        
    frameStaff.Visible = False
    frameHandle.Visible = False
    frameProfile.Visible = False
    frameNewUser.Visible = False
    frameDept.Visible = False
    frameSubject.Visible = True
    
    cmbDept_Load cmbSubjDept
    cmbBatch_Load cmbSubjBatch
    cmbSem_Load cmbSubjSem
    LoadGridForSubject
    cmbSubjDept.SetFocus
    'MsgBox "You are requested not to add Lab Courses for this feedback.", vbInformation, "Feedback Tips"
    End Sub
Private Sub LoadGridForSubject()
    On Error Resume Next
    Dim rsSubjData As New ADODB.Recordset
    rsSubjData.Open "select d.deptshort as dept,s.subjcode,s.subjname,s.batch,s.semno from subj s,dept d where s.dept=d.deptcode and d.deptcode='" & Department(cmbSubjDept) & "' order by d.deptshort,s.batch,s.semno,s.subjname  ", conn, adOpenDynamic, adLockOptimistic, -1
    mshData.Clear
    Set mshData.DataSource = rsSubjData
    mshData.Visible = True
    mshData.Left = frameSubject.Left + frameSubject.Width + 240
    mshData.Width = frameNew.Width - mshData.Left - 360
    mshData.Top = frameSubject.Top
    mshData.Height = frameNew.Height - mshData.Top - 360
    
    mshData.ColWidth(0) = 900
    mshData.ColWidth(1) = 1000
    mshData.ColWidth(2) = 3200
    mshData.ColWidth(3) = 750
    mshData.ColWidth(4) = 750
    
    mshData.ColAlignment(0) = flexAlignCenterCenter
    mshData.ColAlignment(1) = flexAlignCenterCenter
    mshData.ColAlignment(3) = flexAlignCenterCenter
    mshData.ColAlignment(4) = flexAlignCenterCenter
    
    mshData.ColAlignmentFixed(0) = flexAlignCenterCenter
    mshData.ColAlignmentFixed(1) = flexAlignCenterCenter
    mshData.ColAlignmentFixed(2) = flexAlignCenterCenter
    mshData.ColAlignmentFixed(3) = flexAlignCenterCenter
    mshData.ColAlignmentFixed(4) = flexAlignCenterCenter
    If Err.Number <> 0 Then
        If Err.Number = 30004 Then
            'Invalid Column alignment error
        ElseIf Err.Number = 30022 Then
            'The Hierarchical FlexGrid does not support the requested type of data binding.
        ElseIf Err.Number = 3021 Then
        Else
            MsgBox "Error Number: " & Err.Number & vbCrLf & "Error Description : " & Err.Description & vbCrLf & "Error Location : frmNew-LoadGridForSubject()", vbExclamation, "Jangid Corporation"
        End If
    End If
End Sub
Private Sub LoadGridForStaff()
    On Error Resume Next
    Dim rsStaffData As New ADODB.Recordset
    rsStaffData.Open "select d.deptshort as dept,s.staffid,s.staffname from staff s,dept d where s.dept=d.deptcode and s.dept = '" & Department(cmbStaffDept) & "' order by d.deptshort,s.staffid", conn, adOpenDynamic, adLockOptimistic, -1
    mshData.Clear
    Set mshData.DataSource = rsStaffData
    mshData.Visible = True
    mshData.Left = frameStaff.Left + frameStaff.Width + 240
    mshData.Width = frameNew.Width - mshData.Left - 360
    mshData.Top = frameStaff.Top
    mshData.Height = frameNew.Height - mshData.Top - 360
    
    mshData.ColWidth(0) = 1000
    mshData.ColWidth(1) = 1000
    mshData.ColWidth(2) = 3200
        
    mshData.ColAlignment(0) = flexAlignCenterCenter
    mshData.ColAlignment(1) = flexAlignCenterCenter
        
    mshData.ColAlignmentFixed(0) = flexAlignCenterCenter
    mshData.ColAlignmentFixed(1) = flexAlignCenterCenter
    mshData.ColAlignmentFixed(2) = flexAlignCenterCenter
    
    If Err.Number <> 0 Then
        If Err.Number = 30004 Then
            'Invalid Column alignment error
        ElseIf Err.Number = 30022 Then
            'The Hierarchical FlexGrid does not support the requested type of data binding.
        ElseIf Err.Number = 3021 Then
        Else
            MsgBox "Error Number: " & Err.Number & vbCrLf & "Error Description : " & Err.Description & vbCrLf & "Error Location : frmNew-LoadGridForStaff()", vbExclamation, "Jangid Corporation"
        End If
    End If
End Sub
Private Sub LoadGridForHandle()
    On Error Resume Next
    Dim rsHandleData As New ADODB.Recordset
    rsHandleData.Open "select d1.deptshort as staffdept,s.staffname,d.deptshort as subjdept,h.batch,h.sem,h.sec,h.subjcode from staffhandle h,dept d,dept d1,staff s where h.dept=d.deptcode and h.staffid=s.staffid and s.dept=d1.deptcode and s.dept='" & Department(cmbHandleStaffDept) & "' order by d1.deptshort,s.staffname,d.deptshort,h.batch,h.sem,h.sec,h.subjcode", conn, adOpenDynamic, adLockOptimistic, -1
    mshData.Clear
    Set mshData.DataSource = rsHandleData
    mshData.Visible = True
    mshData.Left = frameHandle.Left + frameHandle.Width + 240
    mshData.Width = frameNew.Width - mshData.Left - 360
    mshData.Top = frameHandle.Top
    mshData.Height = frameNew.Height - mshData.Top - 360
    
    mshData.ColWidth(0) = 1000
    mshData.ColWidth(1) = 1600
    mshData.ColWidth(2) = 1000
    mshData.ColWidth(3) = 750
    mshData.ColWidth(4) = 650
    mshData.ColWidth(5) = 650
    mshData.ColWidth(6) = 1000
    
    mshData.ColAlignment(0) = flexAlignCenterCenter
    mshData.ColAlignment(2) = flexAlignCenterCenter
    mshData.ColAlignment(3) = flexAlignCenterCenter
    mshData.ColAlignment(4) = flexAlignCenterCenter
    mshData.ColAlignment(5) = flexAlignCenterCenter
    mshData.ColAlignment(6) = flexAlignCenterCenter
    
    mshData.ColAlignmentFixed(0) = flexAlignCenterCenter
    mshData.ColAlignmentFixed(1) = flexAlignCenterCenter
    mshData.ColAlignmentFixed(2) = flexAlignCenterCenter
    mshData.ColAlignmentFixed(3) = flexAlignCenterCenter
    mshData.ColAlignmentFixed(4) = flexAlignCenterCenter
    mshData.ColAlignmentFixed(5) = flexAlignCenterCenter
    
    If Err.Number <> 0 Then
        If Err.Number = 30004 Then
            'Invalid Column alignment error
        ElseIf Err.Number = 30022 Then
            'The Hierarchical FlexGrid does not support the requested type of data binding.
        ElseIf Err.Number = 3021 Then
            'Dept Combo is left blank at the start. So EOF or BOF error.
        Else
            MsgBox "Error Number: " & Err.Number & vbCrLf & "Error Description : " & Err.Description & vbCrLf & "Error Location : frmNew-LoadGridForHandle()", vbExclamation, "Jangid Corporation"
        End If
    End If
End Sub

Private Sub cmdUninstall_Click()
    On Error GoTo errHan
    Set reg = CreateObject("Wscript.Shell")
    reg.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\feedback"
    Shell App.Path & "\unins000.exe"
    End
    Exit Sub
errHan:
    If Err.Number <> 0 Then
        If Err.Number = 53 Then
            MsgBox "The application is not installed using the feedback installation package.", vbCritical, "Jangid Corporation"
        End If
    End If
End Sub

Private Sub Form_Activate()
    frameNew.Top = 240
    frameNew.Left = 240
    frameNew.Width = Screen.Width - 480
    frameNew.Height = Screen.Height - 1440
    'cmdUninstall.Top = frameNew.Height - cmdUninstall.Height - 360
    
    frameDept.Top = 600
    frameDept.Left = 1500
    frameDept.Width = 5775
    frameDept.Height = 3270
    
    frameSubject.Top = 600
    frameSubject.Left = 1500
    frameSubject.Width = 5775
    frameSubject.Height = 4455
    
    frameStaff.Top = 600
    frameStaff.Left = 1500
    frameStaff.Width = 5775
    frameStaff.Height = 2415
    
    frameHandle.Top = 600
    frameHandle.Left = 1500
    frameHandle.Width = 5775
    frameHandle.Height = 5295
    
    frameProfile.Top = 600
    frameProfile.Left = 1500
    frameProfile.Width = frameNew.Width - frameProfile.Left - 360
    frameProfile.Height = frameNew.Height - frameProfile.Top - 360
    tvProfile.Left = 240
    tvProfile.Top = 600
    tvProfile.Width = frameProfile.Width - 480
    tvProfile.Height = frameProfile.Height - tvProfile.Top - 240
    
    frameNewUser.Top = 600
    frameNewUser.Left = 1500
    frameNewUser.Width = 5775
    frameNewUser.Height = 3630
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    frmMain.Show
End Sub





