VERSION 5.00
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkusercontrolsxp.ocx"
Begin VB.Form frmFeed 
   BackColor       =   &H8000000C&
   Caption         =   "Feedback"
   ClientHeight    =   6735
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11235
   Icon            =   "frmFeed.frx":0000
   ScaleHeight     =   6735
   ScaleWidth      =   11235
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin feedback.ThemedComboBox ThemedComboBox1 
      Left            =   8640
      Top             =   0
      _ExtentX        =   556
      _ExtentY        =   529
   End
   Begin VB.Frame frame 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      TabIndex        =   144
      Top             =   6480
      Visible         =   0   'False
      Width           =   3135
      Begin VB.CommandButton cmdTrunc 
         Caption         =   "Truncate"
         Enabled         =   0   'False
         Height          =   255
         Left            =   0
         TabIndex        =   151
         TabStop         =   0   'False
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmd5 
         Caption         =   "5"
         Height          =   255
         Left            =   2760
         TabIndex        =   150
         TabStop         =   0   'False
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmd4 
         Caption         =   "4"
         Height          =   255
         Left            =   2400
         TabIndex        =   149
         TabStop         =   0   'False
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmd3 
         Caption         =   "3"
         Height          =   255
         Left            =   2040
         TabIndex        =   148
         TabStop         =   0   'False
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmd2 
         Caption         =   "2"
         Height          =   255
         Left            =   1680
         TabIndex        =   147
         TabStop         =   0   'False
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "1"
         Height          =   255
         Left            =   1320
         TabIndex        =   146
         TabStop         =   0   'False
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdFill 
         Caption         =   "Fill"
         Height          =   255
         Left            =   960
         TabIndex        =   145
         TabStop         =   0   'False
         Top             =   0
         Width           =   375
      End
   End
   Begin vkUserContolsXP.vkFrame frameFeed 
      Height          =   375
      Left            =   9360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
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
      TitleColor1     =   8421504
      TitleColor2     =   4210752
      TitleGradient   =   2
      TitleHeight     =   400
      BorderColor     =   4210752
      RoundAngle      =   15
      BorderWidth     =   2
      Begin vkUserContolsXP.vkLabel lblRemaining 
         Height          =   435
         Left            =   360
         TabIndex        =   157
         Top             =   9000
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   767
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
      End
      Begin vkUserContolsXP.vkLabel lblPoints 
         Height          =   375
         Left            =   3360
         TabIndex        =   143
         Top             =   9000
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   661
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Poor : 1   |   Fair : 2    |   Good : 3    |   Very Good : 4   |   Excellent : 5 "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
      Begin vkUserContolsXP.vkFrame frameFour 
         Height          =   4095
         Left            =   7560
         TabIndex        =   112
         Top             =   4800
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   7223
         Caption         =   "Class Management / Assesement Of Students"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TitleColor1     =   8421504
         TitleColor2     =   12632256
         TitleGradient   =   2
         TitleHeight     =   360
         BorderColor     =   4210752
         BorderWidth     =   2
         Begin vkUserContolsXP.vkOptionButton ob1 
            Height          =   255
            Index           =   19
            Left            =   4560
            TabIndex        =   137
            Top             =   3360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   20
         End
         Begin vkUserContolsXP.vkOptionButton ob2 
            Height          =   255
            Index           =   19
            Left            =   5040
            TabIndex        =   136
            Top             =   3360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   20
         End
         Begin vkUserContolsXP.vkOptionButton ob3 
            Height          =   255
            Index           =   19
            Left            =   5520
            TabIndex        =   135
            Top             =   3360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "3"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   20
         End
         Begin vkUserContolsXP.vkOptionButton ob4 
            Height          =   255
            Index           =   19
            Left            =   6000
            TabIndex        =   134
            Top             =   3360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "4"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   20
         End
         Begin vkUserContolsXP.vkOptionButton ob5 
            Height          =   255
            Index           =   19
            Left            =   6480
            TabIndex        =   133
            Top             =   3360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   20
         End
         Begin vkUserContolsXP.vkOptionButton ob1 
            Height          =   255
            Index           =   18
            Left            =   4560
            TabIndex        =   132
            Top             =   2640
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   19
         End
         Begin vkUserContolsXP.vkOptionButton ob2 
            Height          =   255
            Index           =   18
            Left            =   5040
            TabIndex        =   131
            Top             =   2640
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   19
         End
         Begin vkUserContolsXP.vkOptionButton ob3 
            Height          =   255
            Index           =   18
            Left            =   5520
            TabIndex        =   130
            Top             =   2640
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "3"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   19
         End
         Begin vkUserContolsXP.vkOptionButton ob4 
            Height          =   255
            Index           =   18
            Left            =   6000
            TabIndex        =   129
            Top             =   2640
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "4"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   19
         End
         Begin vkUserContolsXP.vkOptionButton ob5 
            Height          =   255
            Index           =   18
            Left            =   6480
            TabIndex        =   128
            Top             =   2640
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   19
         End
         Begin vkUserContolsXP.vkOptionButton ob1 
            Height          =   255
            Index           =   17
            Left            =   4560
            TabIndex        =   127
            Top             =   1920
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   18
         End
         Begin vkUserContolsXP.vkOptionButton ob2 
            Height          =   255
            Index           =   17
            Left            =   5040
            TabIndex        =   126
            Top             =   1920
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   18
         End
         Begin vkUserContolsXP.vkOptionButton ob3 
            Height          =   255
            Index           =   17
            Left            =   5520
            TabIndex        =   125
            Top             =   1920
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "3"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   18
         End
         Begin vkUserContolsXP.vkOptionButton ob4 
            Height          =   255
            Index           =   17
            Left            =   6000
            TabIndex        =   124
            Top             =   1920
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "4"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   18
         End
         Begin vkUserContolsXP.vkOptionButton ob5 
            Height          =   255
            Index           =   17
            Left            =   6480
            TabIndex        =   123
            Top             =   1920
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   18
         End
         Begin vkUserContolsXP.vkOptionButton ob1 
            Height          =   255
            Index           =   16
            Left            =   4560
            TabIndex        =   122
            Top             =   1200
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   17
         End
         Begin vkUserContolsXP.vkOptionButton ob2 
            Height          =   255
            Index           =   16
            Left            =   5040
            TabIndex        =   121
            Top             =   1200
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   17
         End
         Begin vkUserContolsXP.vkOptionButton ob3 
            Height          =   255
            Index           =   16
            Left            =   5520
            TabIndex        =   120
            Top             =   1200
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "3"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   17
         End
         Begin vkUserContolsXP.vkOptionButton ob4 
            Height          =   255
            Index           =   16
            Left            =   6000
            TabIndex        =   119
            Top             =   1200
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "4"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   17
         End
         Begin vkUserContolsXP.vkOptionButton ob5 
            Height          =   255
            Index           =   16
            Left            =   6480
            TabIndex        =   118
            Top             =   1200
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   17
         End
         Begin vkUserContolsXP.vkOptionButton ob1 
            Height          =   255
            Index           =   15
            Left            =   4560
            TabIndex        =   117
            Top             =   480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   16
         End
         Begin vkUserContolsXP.vkOptionButton ob2 
            Height          =   255
            Index           =   15
            Left            =   5040
            TabIndex        =   116
            Top             =   480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   16
         End
         Begin vkUserContolsXP.vkOptionButton ob3 
            Height          =   255
            Index           =   15
            Left            =   5520
            TabIndex        =   115
            Top             =   480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "3"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   16
         End
         Begin vkUserContolsXP.vkOptionButton ob4 
            Height          =   255
            Index           =   15
            Left            =   6000
            TabIndex        =   114
            Top             =   480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "4"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   16
         End
         Begin vkUserContolsXP.vkOptionButton ob5 
            Height          =   255
            Index           =   15
            Left            =   6480
            TabIndex        =   113
            Top             =   480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   16
         End
         Begin VB.Label lblQues 
            BackStyle       =   0  'Transparent
            Caption         =   "4.5 Teacher is prompt in valuing and returning the answer scripts providing feedback on performance."
            Height          =   495
            Index           =   19
            Left            =   240
            TabIndex        =   142
            Top             =   3360
            Width           =   4200
         End
         Begin VB.Label lblQues 
            BackStyle       =   0  'Transparent
            Caption         =   "4.4 Teacher's making of scripts is fair and impartial."
            Height          =   255
            Index           =   18
            Left            =   240
            TabIndex        =   141
            Top             =   2640
            Width           =   4200
         End
         Begin VB.Label lblQues 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmFeed.frx":0ECA
            Height          =   735
            Index           =   17
            Left            =   240
            TabIndex        =   140
            Top             =   1920
            Width           =   4200
         End
         Begin VB.Label lblQues 
            BackStyle       =   0  'Transparent
            Caption         =   "4.2 Teacher covers the syllabus completely and at appropriate pace."
            Height          =   495
            Index           =   16
            Left            =   240
            TabIndex        =   139
            Top             =   1200
            Width           =   4200
         End
         Begin VB.Label lblQues 
            BackStyle       =   0  'Transparent
            Caption         =   "4.1 Teacher engages classes regularly and maintains discipline."
            Height          =   495
            Index           =   15
            Left            =   240
            TabIndex        =   138
            Top             =   480
            Width           =   4200
         End
      End
      Begin vkUserContolsXP.vkFrame frameThree 
         Height          =   4095
         Left            =   7560
         TabIndex        =   81
         Top             =   480
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   7223
         Caption         =   "Student's Participation"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TitleColor1     =   8421504
         TitleColor2     =   12632256
         TitleGradient   =   2
         TitleHeight     =   360
         BorderColor     =   4210752
         BorderWidth     =   2
         Begin vkUserContolsXP.vkOptionButton ob1 
            Height          =   255
            Index           =   14
            Left            =   4560
            TabIndex        =   106
            Top             =   3360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   15
         End
         Begin vkUserContolsXP.vkOptionButton ob2 
            Height          =   255
            Index           =   14
            Left            =   5040
            TabIndex        =   105
            Top             =   3360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   15
         End
         Begin vkUserContolsXP.vkOptionButton ob3 
            Height          =   255
            Index           =   14
            Left            =   5520
            TabIndex        =   104
            Top             =   3360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "3"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   15
         End
         Begin vkUserContolsXP.vkOptionButton ob4 
            Height          =   255
            Index           =   14
            Left            =   6000
            TabIndex        =   103
            Top             =   3360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "4"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   15
         End
         Begin vkUserContolsXP.vkOptionButton ob5 
            Height          =   255
            Index           =   14
            Left            =   6480
            TabIndex        =   102
            Top             =   3360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   15
         End
         Begin vkUserContolsXP.vkOptionButton ob1 
            Height          =   255
            Index           =   13
            Left            =   4560
            TabIndex        =   101
            Top             =   2640
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   14
         End
         Begin vkUserContolsXP.vkOptionButton ob2 
            Height          =   255
            Index           =   13
            Left            =   5040
            TabIndex        =   100
            Top             =   2640
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   14
         End
         Begin vkUserContolsXP.vkOptionButton ob3 
            Height          =   255
            Index           =   13
            Left            =   5520
            TabIndex        =   99
            Top             =   2640
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "3"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   14
         End
         Begin vkUserContolsXP.vkOptionButton ob4 
            Height          =   255
            Index           =   13
            Left            =   6000
            TabIndex        =   98
            Top             =   2640
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "4"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   14
         End
         Begin vkUserContolsXP.vkOptionButton ob5 
            Height          =   255
            Index           =   13
            Left            =   6480
            TabIndex        =   97
            Top             =   2640
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   14
         End
         Begin vkUserContolsXP.vkOptionButton ob1 
            Height          =   255
            Index           =   12
            Left            =   4560
            TabIndex        =   96
            Top             =   1920
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   13
         End
         Begin vkUserContolsXP.vkOptionButton ob2 
            Height          =   255
            Index           =   12
            Left            =   5040
            TabIndex        =   95
            Top             =   1920
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   13
         End
         Begin vkUserContolsXP.vkOptionButton ob3 
            Height          =   255
            Index           =   12
            Left            =   5520
            TabIndex        =   94
            Top             =   1920
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "3"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   13
         End
         Begin vkUserContolsXP.vkOptionButton ob4 
            Height          =   255
            Index           =   12
            Left            =   6000
            TabIndex        =   93
            Top             =   1920
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "4"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   13
         End
         Begin vkUserContolsXP.vkOptionButton ob5 
            Height          =   255
            Index           =   12
            Left            =   6480
            TabIndex        =   92
            Top             =   1920
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   13
         End
         Begin vkUserContolsXP.vkOptionButton ob1 
            Height          =   255
            Index           =   11
            Left            =   4560
            TabIndex        =   91
            Top             =   1200
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   12
         End
         Begin vkUserContolsXP.vkOptionButton ob2 
            Height          =   255
            Index           =   11
            Left            =   5040
            TabIndex        =   90
            Top             =   1200
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   12
         End
         Begin vkUserContolsXP.vkOptionButton ob3 
            Height          =   255
            Index           =   11
            Left            =   5520
            TabIndex        =   89
            Top             =   1200
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "3"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   12
         End
         Begin vkUserContolsXP.vkOptionButton ob4 
            Height          =   255
            Index           =   11
            Left            =   6000
            TabIndex        =   88
            Top             =   1200
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "4"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   12
         End
         Begin vkUserContolsXP.vkOptionButton ob5 
            Height          =   255
            Index           =   11
            Left            =   6480
            TabIndex        =   87
            Top             =   1200
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   12
         End
         Begin vkUserContolsXP.vkOptionButton ob1 
            Height          =   255
            Index           =   10
            Left            =   4560
            TabIndex        =   86
            Top             =   480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   11
         End
         Begin vkUserContolsXP.vkOptionButton ob2 
            Height          =   255
            Index           =   10
            Left            =   5040
            TabIndex        =   85
            Top             =   480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   11
         End
         Begin vkUserContolsXP.vkOptionButton ob3 
            Height          =   255
            Index           =   10
            Left            =   5520
            TabIndex        =   84
            Top             =   480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "3"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   11
         End
         Begin vkUserContolsXP.vkOptionButton ob4 
            Height          =   255
            Index           =   10
            Left            =   6000
            TabIndex        =   83
            Top             =   480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "4"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   11
         End
         Begin vkUserContolsXP.vkOptionButton ob5 
            Height          =   255
            Index           =   10
            Left            =   6480
            TabIndex        =   82
            Top             =   480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   11
         End
         Begin VB.Label lblQues 
            BackStyle       =   0  'Transparent
            Caption         =   "3.5 Teacher is courteous and impartial in dealing with the students."
            Height          =   495
            Index           =   14
            Left            =   240
            TabIndex        =   111
            Top             =   3360
            Width           =   4200
         End
         Begin VB.Label lblQues 
            BackStyle       =   0  'Transparent
            Caption         =   "3.4 Teacher encourages, compliments and praises originality and creativity displayed by the students."
            Height          =   495
            Index           =   13
            Left            =   240
            TabIndex        =   110
            Top             =   2640
            Width           =   4200
         End
         Begin VB.Label lblQues 
            BackStyle       =   0  'Transparent
            Caption         =   "3.3 Teacher ensures learner activity and problems solving ability in the class."
            Height          =   495
            Index           =   12
            Left            =   240
            TabIndex        =   109
            Top             =   1920
            Width           =   4200
         End
         Begin VB.Label lblQues 
            BackStyle       =   0  'Transparent
            Caption         =   "3.2 Teacher encourages questioning / raising doubts by students and answers them well."
            Height          =   495
            Index           =   11
            Left            =   240
            TabIndex        =   108
            Top             =   1200
            Width           =   4200
         End
         Begin VB.Label lblQues 
            BackStyle       =   0  'Transparent
            Caption         =   "3.1 Teacher asks questions to promote interaction and reflecctive thinking."
            Height          =   495
            Index           =   10
            Left            =   240
            TabIndex        =   107
            Top             =   480
            Width           =   4200
         End
      End
      Begin vkUserContolsXP.vkFrame frameTwo 
         Height          =   4095
         Left            =   240
         TabIndex        =   50
         Top             =   4800
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   7223
         Caption         =   "Presentation | Communication"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TitleColor1     =   8421504
         TitleColor2     =   12632256
         TitleGradient   =   2
         TitleHeight     =   360
         BorderColor     =   4210752
         BorderWidth     =   2
         Begin vkUserContolsXP.vkOptionButton ob5 
            Height          =   255
            Index           =   9
            Left            =   6480
            TabIndex        =   75
            Top             =   3360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   10
         End
         Begin vkUserContolsXP.vkOptionButton ob4 
            Height          =   255
            Index           =   9
            Left            =   6000
            TabIndex        =   74
            Top             =   3360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "4"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   10
         End
         Begin vkUserContolsXP.vkOptionButton ob3 
            Height          =   255
            Index           =   9
            Left            =   5520
            TabIndex        =   73
            Top             =   3360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "3"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   10
         End
         Begin vkUserContolsXP.vkOptionButton ob2 
            Height          =   255
            Index           =   9
            Left            =   5040
            TabIndex        =   72
            Top             =   3360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   10
         End
         Begin vkUserContolsXP.vkOptionButton ob1 
            Height          =   255
            Index           =   9
            Left            =   4560
            TabIndex        =   71
            Top             =   3360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   10
         End
         Begin vkUserContolsXP.vkOptionButton ob5 
            Height          =   255
            Index           =   8
            Left            =   6480
            TabIndex        =   70
            Top             =   2640
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   9
         End
         Begin vkUserContolsXP.vkOptionButton ob4 
            Height          =   255
            Index           =   8
            Left            =   6000
            TabIndex        =   69
            Top             =   2640
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "4"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   9
         End
         Begin vkUserContolsXP.vkOptionButton ob3 
            Height          =   255
            Index           =   8
            Left            =   5520
            TabIndex        =   68
            Top             =   2640
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "3"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   9
         End
         Begin vkUserContolsXP.vkOptionButton ob2 
            Height          =   255
            Index           =   8
            Left            =   5040
            TabIndex        =   67
            Top             =   2640
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   9
         End
         Begin vkUserContolsXP.vkOptionButton ob1 
            Height          =   255
            Index           =   8
            Left            =   4560
            TabIndex        =   66
            Top             =   2640
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   9
         End
         Begin vkUserContolsXP.vkOptionButton ob5 
            Height          =   255
            Index           =   7
            Left            =   6480
            TabIndex        =   65
            Top             =   1920
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   8
         End
         Begin vkUserContolsXP.vkOptionButton ob4 
            Height          =   255
            Index           =   7
            Left            =   6000
            TabIndex        =   64
            Top             =   1920
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "4"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   8
         End
         Begin vkUserContolsXP.vkOptionButton ob3 
            Height          =   255
            Index           =   7
            Left            =   5520
            TabIndex        =   63
            Top             =   1920
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "3"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   8
         End
         Begin vkUserContolsXP.vkOptionButton ob2 
            Height          =   255
            Index           =   7
            Left            =   5040
            TabIndex        =   62
            Top             =   1920
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   8
         End
         Begin vkUserContolsXP.vkOptionButton ob1 
            Height          =   255
            Index           =   7
            Left            =   4560
            TabIndex        =   61
            Top             =   1920
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   8
         End
         Begin vkUserContolsXP.vkOptionButton ob5 
            Height          =   255
            Index           =   6
            Left            =   6480
            TabIndex        =   60
            Top             =   1200
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   7
         End
         Begin vkUserContolsXP.vkOptionButton ob4 
            Height          =   255
            Index           =   6
            Left            =   6000
            TabIndex        =   59
            Top             =   1200
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "4"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   7
         End
         Begin vkUserContolsXP.vkOptionButton ob3 
            Height          =   255
            Index           =   6
            Left            =   5520
            TabIndex        =   58
            Top             =   1200
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "3"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   7
         End
         Begin vkUserContolsXP.vkOptionButton ob2 
            Height          =   255
            Index           =   6
            Left            =   5040
            TabIndex        =   57
            Top             =   1200
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   7
         End
         Begin vkUserContolsXP.vkOptionButton ob1 
            Height          =   255
            Index           =   6
            Left            =   4560
            TabIndex        =   56
            Top             =   1200
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   7
         End
         Begin vkUserContolsXP.vkOptionButton ob5 
            Height          =   255
            Index           =   5
            Left            =   6480
            TabIndex        =   55
            Top             =   480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   6
         End
         Begin vkUserContolsXP.vkOptionButton ob4 
            Height          =   255
            Index           =   5
            Left            =   6000
            TabIndex        =   54
            Top             =   480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "4"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   6
         End
         Begin vkUserContolsXP.vkOptionButton ob3 
            Height          =   255
            Index           =   5
            Left            =   5520
            TabIndex        =   53
            Top             =   480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "3"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   6
         End
         Begin vkUserContolsXP.vkOptionButton ob2 
            Height          =   255
            Index           =   5
            Left            =   5040
            TabIndex        =   52
            Top             =   480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   6
         End
         Begin vkUserContolsXP.vkOptionButton ob1 
            Height          =   255
            Index           =   5
            Left            =   4560
            TabIndex        =   51
            Top             =   480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   6
         End
         Begin VB.Label lblQues 
            BackStyle       =   0  'Transparent
            Caption         =   "2.5 Teacher offers assistance and counselling to the needy students."
            Height          =   495
            Index           =   9
            Left            =   240
            TabIndex        =   80
            Top             =   3360
            Width           =   4200
         End
         Begin VB.Label lblQues 
            BackStyle       =   0  'Transparent
            Caption         =   "2.4 Teacher's pace and level of instruction are suited to the attainment of students."
            Height          =   495
            Index           =   8
            Left            =   240
            TabIndex        =   79
            Top             =   2640
            Width           =   4200
         End
         Begin VB.Label lblQues 
            BackStyle       =   0  'Transparent
            Caption         =   "2.3 Teacher provides examples of concepts / principles. Explanations are clear and effective."
            Height          =   495
            Index           =   7
            Left            =   240
            TabIndex        =   78
            Top             =   1920
            Width           =   4200
         End
         Begin VB.Label lblQues 
            BackStyle       =   0  'Transparent
            Caption         =   "2.2 Teacher writes and draws legibly."
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   77
            Top             =   1200
            Width           =   4200
         End
         Begin VB.Label lblQues 
            BackStyle       =   0  'Transparent
            Caption         =   "2.1 Teacher speaks clearly and audibly."
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   76
            Top             =   480
            Width           =   4200
         End
      End
      Begin vkUserContolsXP.vkFrame frameOne 
         Height          =   4095
         Left            =   240
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   480
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   7223
         Caption         =   "Planning && Organisaton"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TitleColor1     =   8421504
         TitleColor2     =   12632256
         TitleGradient   =   2
         TitleHeight     =   360
         BorderColor     =   4210752
         BorderWidth     =   2
         Begin vkUserContolsXP.vkOptionButton ob5 
            Height          =   255
            Index           =   4
            Left            =   6480
            TabIndex        =   44
            Top             =   2640
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   4
         End
         Begin vkUserContolsXP.vkOptionButton ob4 
            Height          =   255
            Index           =   4
            Left            =   6000
            TabIndex        =   43
            Top             =   2640
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "4"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   4
         End
         Begin vkUserContolsXP.vkOptionButton ob3 
            Height          =   255
            Index           =   4
            Left            =   5520
            TabIndex        =   42
            Top             =   2640
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "3"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   4
         End
         Begin vkUserContolsXP.vkOptionButton ob2 
            Height          =   255
            Index           =   4
            Left            =   5040
            TabIndex        =   41
            Top             =   2640
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   4
         End
         Begin vkUserContolsXP.vkOptionButton ob1 
            Height          =   255
            Index           =   4
            Left            =   4560
            TabIndex        =   40
            Top             =   2640
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   4
         End
         Begin vkUserContolsXP.vkOptionButton ob5 
            Height          =   255
            Index           =   3
            Left            =   6480
            TabIndex        =   39
            Top             =   3360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   5
         End
         Begin vkUserContolsXP.vkOptionButton ob4 
            Height          =   255
            Index           =   3
            Left            =   6000
            TabIndex        =   38
            Top             =   3360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "4"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   5
         End
         Begin vkUserContolsXP.vkOptionButton ob3 
            Height          =   255
            Index           =   3
            Left            =   5520
            TabIndex        =   37
            Top             =   3360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "3"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   5
         End
         Begin vkUserContolsXP.vkOptionButton ob2 
            Height          =   255
            Index           =   3
            Left            =   5040
            TabIndex        =   36
            Top             =   3360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   5
         End
         Begin vkUserContolsXP.vkOptionButton ob1 
            Height          =   255
            Index           =   3
            Left            =   4560
            TabIndex        =   35
            Top             =   3360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   5
         End
         Begin vkUserContolsXP.vkOptionButton ob5 
            Height          =   255
            Index           =   2
            Left            =   6480
            TabIndex        =   34
            Top             =   1920
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   3
         End
         Begin vkUserContolsXP.vkOptionButton ob4 
            Height          =   255
            Index           =   2
            Left            =   6000
            TabIndex        =   33
            Top             =   1920
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "4"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   3
         End
         Begin vkUserContolsXP.vkOptionButton ob3 
            Height          =   255
            Index           =   2
            Left            =   5520
            TabIndex        =   32
            Top             =   1920
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "3"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   3
         End
         Begin vkUserContolsXP.vkOptionButton ob2 
            Height          =   255
            Index           =   2
            Left            =   5040
            TabIndex        =   31
            Top             =   1920
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   3
         End
         Begin vkUserContolsXP.vkOptionButton ob1 
            Height          =   255
            Index           =   2
            Left            =   4560
            TabIndex        =   30
            Top             =   1920
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   3
         End
         Begin vkUserContolsXP.vkOptionButton ob5 
            Height          =   255
            Index           =   1
            Left            =   6480
            TabIndex        =   29
            Top             =   1200
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   2
         End
         Begin vkUserContolsXP.vkOptionButton ob4 
            Height          =   255
            Index           =   1
            Left            =   6000
            TabIndex        =   28
            Top             =   1200
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "4"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   2
         End
         Begin vkUserContolsXP.vkOptionButton ob3 
            Height          =   255
            Index           =   1
            Left            =   5520
            TabIndex        =   27
            Top             =   1200
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "3"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   2
         End
         Begin vkUserContolsXP.vkOptionButton ob2 
            Height          =   255
            Index           =   1
            Left            =   5040
            TabIndex        =   26
            Top             =   1200
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   2
         End
         Begin vkUserContolsXP.vkOptionButton ob1 
            Height          =   255
            Index           =   1
            Left            =   4560
            TabIndex        =   25
            Top             =   1200
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   2
         End
         Begin vkUserContolsXP.vkOptionButton ob5 
            Height          =   255
            Index           =   0
            Left            =   6480
            TabIndex        =   24
            Top             =   480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   1
         End
         Begin vkUserContolsXP.vkOptionButton ob4 
            Height          =   255
            Index           =   0
            Left            =   6000
            TabIndex        =   23
            Top             =   480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "4"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   1
         End
         Begin vkUserContolsXP.vkOptionButton ob3 
            Height          =   255
            Index           =   0
            Left            =   5520
            TabIndex        =   22
            Top             =   480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "3"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   1
         End
         Begin vkUserContolsXP.vkOptionButton ob2 
            Height          =   255
            Index           =   0
            Left            =   5040
            TabIndex        =   21
            Top             =   480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   1
         End
         Begin vkUserContolsXP.vkOptionButton ob1 
            Height          =   255
            Index           =   0
            Left            =   4560
            TabIndex        =   20
            Top             =   480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Group           =   1
         End
         Begin VB.Label lblQues 
            BackStyle       =   0  'Transparent
            Caption         =   "1.5 Teacher comes well prepared in the subject."
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   49
            Top             =   3360
            Width           =   4200
         End
         Begin VB.Label lblQues 
            BackStyle       =   0  'Transparent
            Caption         =   "1.4 Subject matter organized in logical sequence."
            Height          =   495
            Index           =   3
            Left            =   240
            TabIndex        =   48
            Top             =   2640
            Width           =   4200
         End
         Begin VB.Label lblQues 
            BackStyle       =   0  'Transparent
            Caption         =   "1.3 Aims/Objectives made clear."
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   47
            Top             =   1920
            Width           =   4200
         End
         Begin VB.Label lblQues 
            BackStyle       =   0  'Transparent
            Caption         =   "1.2 Teacher is well planned."
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   46
            Top             =   1200
            Width           =   4200
         End
         Begin VB.Label lblQues 
            BackStyle       =   0  'Transparent
            Caption         =   "1.1 Teacher comes to class in time."
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   45
            Top             =   480
            Width           =   4200
         End
      End
      Begin feedback.StylerButton cmdGo 
         Height          =   375
         Left            =   12720
         TabIndex        =   18
         Top             =   9000
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Caption         =   "Submit"
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
   Begin vkUserContolsXP.vkFrame frameMain 
      Height          =   4095
      Left            =   840
      TabIndex        =   152
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7223
      Caption         =   "   Enter your information."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextPosition    =   0
      TitleColor1     =   8421504
      TitleColor2     =   4210752
      TitleGradient   =   2
      TitleHeight     =   360
      BorderColor     =   4210752
      BorderWidth     =   2
      Begin feedback.StylerButton cmdSubmit 
         Height          =   375
         Left            =   2280
         TabIndex        =   5
         Top             =   3600
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         Caption         =   "Submit"
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
      Begin feedback.StylerButton cmdTake 
         Height          =   375
         Left            =   6120
         TabIndex        =   16
         Top             =   3600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Take Feedback!"
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
      Begin vkUserContolsXP.vkCheck cbSubj 
         Height          =   255
         Index           =   9
         Left            =   3240
         TabIndex        =   15
         Top             =   3240
         Visible         =   0   'False
         Width           =   4215
         _ExtentX        =   7435
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
      Begin vkUserContolsXP.vkCheck cbSubj 
         Height          =   255
         Index           =   8
         Left            =   3240
         TabIndex        =   14
         Top             =   3000
         Visible         =   0   'False
         Width           =   4215
         _ExtentX        =   7435
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
      Begin vkUserContolsXP.vkCheck cbSubj 
         Height          =   255
         Index           =   7
         Left            =   3240
         TabIndex        =   13
         Top             =   2760
         Visible         =   0   'False
         Width           =   4215
         _ExtentX        =   7435
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
      Begin vkUserContolsXP.vkCheck cbSubj 
         Height          =   255
         Index           =   6
         Left            =   3240
         TabIndex        =   12
         Top             =   2520
         Visible         =   0   'False
         Width           =   4215
         _ExtentX        =   7435
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
      Begin vkUserContolsXP.vkCheck cbSubj 
         Height          =   255
         Index           =   5
         Left            =   3240
         TabIndex        =   17
         Top             =   2280
         Visible         =   0   'False
         Width           =   4215
         _ExtentX        =   7435
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
      Begin vkUserContolsXP.vkCheck cbSubj 
         Height          =   255
         Index           =   4
         Left            =   3240
         TabIndex        =   11
         Top             =   2040
         Visible         =   0   'False
         Width           =   4215
         _ExtentX        =   7435
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
      Begin vkUserContolsXP.vkCheck cbSubj 
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   10
         Top             =   1800
         Visible         =   0   'False
         Width           =   4215
         _ExtentX        =   7435
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
      Begin vkUserContolsXP.vkCheck cbSubj 
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   9
         Top             =   1560
         Visible         =   0   'False
         Width           =   4215
         _ExtentX        =   7435
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
      Begin vkUserContolsXP.vkCheck cbSubj 
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   4215
         _ExtentX        =   7435
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
      Begin vkUserContolsXP.vkCheck cbSelect 
         Height          =   255
         Left            =   3240
         TabIndex        =   6
         Top             =   600
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Select All"
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
      Begin vkUserContolsXP.vkCheck cbSubj 
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   7
         Top             =   1080
         Visible         =   0   'False
         Width           =   4215
         _ExtentX        =   7435
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
      Begin VB.ComboBox cmbBatch 
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
         Left            =   1320
         TabIndex        =   2
         Top             =   1200
         Width           =   1575
      End
      Begin VB.ComboBox cmbSem 
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
         Left            =   1320
         TabIndex        =   3
         Top             =   1800
         Width           =   1575
      End
      Begin vkUserContolsXP.vkLabel lblSem 
         Height          =   255
         Left            =   240
         TabIndex        =   156
         TabStop         =   0   'False
         Top             =   1920
         Width           =   855
         _ExtentX        =   1508
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
      Begin vkUserContolsXP.vkLabel lblBatch 
         Height          =   255
         Left            =   240
         TabIndex        =   155
         TabStop         =   0   'False
         Top             =   1320
         Width           =   855
         _ExtentX        =   1508
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
      Begin vkUserContolsXP.vkLabel lblDept 
         Height          =   255
         Left            =   240
         TabIndex        =   154
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
      Begin VB.ComboBox cmbDept 
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
         Left            =   1320
         TabIndex        =   1
         Top             =   600
         Width           =   1575
      End
      Begin vkUserContolsXP.vkLabel lblSec 
         Height          =   255
         Left            =   240
         TabIndex        =   153
         TabStop         =   0   'False
         Top             =   2520
         Width           =   975
         _ExtentX        =   1720
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
      Begin VB.ComboBox cmbSec 
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
         Left            =   1320
         TabIndex        =   4
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Line vLine 
         BorderWidth     =   2
         X1              =   3120
         X2              =   3120
         Y1              =   360
         Y2              =   4050
      End
   End
End
Attribute VB_Name = "frmFeed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iScore() As Integer
Dim strSubj() As String
Dim tmpSubj() As String
Dim iCount As Integer
Dim tmpCount As Integer
Dim iSPtr As Integer
Dim rsMaster As New ADODB.Recordset
Dim strMaster As String
Private Sub cbSelect_Change(Value As CheckBoxConstants)
    If Value = vbChecked Then
        For i = 0 To tmpCount
            cbSubj(i).Value = vbChecked
        Next
    Else
        For i = 0 To tmpCount
            cbSubj(i).Value = vbUnchecked
        Next
    End If
End Sub
Private Sub cmdTake_Click()
    Dim k As Integer
    k = 0
    
    For i = 0 To tmpCount
        If cbSubj(i).Value = vbChecked Then
            strSubj(k, 0) = tmpSubj(i, 0)
            strSubj(k, 1) = tmpSubj(i, 1)
            strSubj(k, 2) = tmpSubj(i, 2)
            strSubj(k, 3) = tmpSubj(i, 3)
            k = k + 1
        End If
    Next
    
    iCount = k - 1
    iSPtr = 0
    
    If iCount = -1 Then
        MsgBox "You have not selected any subject."
        Exit Sub
    Else
        frameMain.Visible = False
        frameFeed.Visible = True
    End If
    
    
    lblRemaining.Caption = "Total Subjects : " & iCount + 1 & vbCrLf & "Subject Completed : " & iSPtr
    frameFeed.Caption = "SUBJECT CODE :  " & strSubj(iSPtr, 0) & "    -    SUBJECT NAME :  " & strSubj(iSPtr, 1) & "    -    STAFF NAME :  " & strSubj(iSPtr, 3)
End Sub
Private Sub HideCheckBox()
    On Error Resume Next
    For i = 0 To tmpCount
        cbSubj(i).Value = vbUnchecked
        cbSubj(i).Visible = False
    Next
End Sub

Private Sub cmdSubmit_Click()
    On Error GoTo errHandler
    
    Dim rs As New ADODB.Recordset
    rs.CursorLocation = adUseClient
    Dim sql As String
    sql = "select s.subjcode,s.subjname,s1.staffid,s1.staffname from subj s,staffhandle h,staff s1 where s1.staffid=h.staffid and s.subjcode=h.subjcode and s.dept=h.dept and s.semno=h.sem and s.batch=h.batch and s.dept='" & iDept & "' and s.batch= '" & Mid(iBatch, 3, 2) & "' and s.semno= '" & iSem & "'  and h.sec='" & strSec & "'"
    rs.Open sql, conn, adOpenDynamic, adLockOptimistic, -1
    rs.MoveFirst
    
    tmpCount = rs.RecordCount - 1
    
    If tmpCount = -1 Then
        MsgBox "No data exist for the given information"
        rs.Close
        Exit Sub
    Else
        Dim o As Integer
        o = 0
        HideCheckBox
        frameMain.Refresh
        While frameMain.Width < 7575
            frameMain.Left = (Screen.Width - frameMain.Width) / 2
            frameMain.Width = frameMain.Width + (o * 100)
            o = o + 2
        Wend
        frameMain.Width = 7575
    End If
    
    ReDim tmpSubj(tmpCount, 3) As String
    ReDim strSubj(tmpCount, 3) As String
    ReDim iScore(tmpCount, 19) As Integer
    
    Dim i As Integer
    i = 0
    While Not rs.EOF
        tmpSubj(i, 0) = rs.Fields(0)
        tmpSubj(i, 1) = rs.Fields(1)
        tmpSubj(i, 2) = rs.Fields(2)
        tmpSubj(i, 3) = rs.Fields(3)
        cbSubj(i).Caption = rs!subjname
        cbSubj(i).Visible = True
        i = i + 1
        rs.MoveNext
    Wend
errHandler:
    If Err.Number <> 0 Then
        If Err.Number = 3021 Then
            MsgBox "Insufficient data in the database." & vbCrLf & "Check whether department,staff and subject are created.", vbInformation, "Jangid Corporation"
        Else
            MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & "Error Location : frmFeed:cmdSubmit_Click", vbInformation, "Jangid Corporation"
        End If
    End If
End Sub
Private Sub cmdGo_Click()
    On Error GoTo errHandler
        
    For i = 0 To 19
        If ob1(i).Value = vbChecked Then
            iScore(iSPtr, i) = 1
        ElseIf ob2(i).Value = vbChecked Then
            iScore(iSPtr, i) = 2
        ElseIf ob3(i).Value = vbChecked Then
            iScore(iSPtr, i) = 3
        ElseIf ob4(i).Value = vbChecked Then
            iScore(iSPtr, i) = 4
        ElseIf ob5(i).Value = vbChecked Then
            iScore(iSPtr, i) = 5
        Else
            MsgBox "Some questions are not answered." & vbCrLf & "Please answer all the question."
            Exit Sub
        End If
    Next
        
    Call rbClear
    
    If iSPtr = iCount Then
        iFID = FID
        For i = 0 To iCount
            strMaster = "insert into master values('" & iFID & "','" & iDept & "','" & Mid(iBatch, 3, 2) & "','" & iSem & "','" & strSec & "','" & strSubj(i, 0) & "','" & strSubj(i, 2) & "','" & iScore(i, 0) & "','" & iScore(i, 1) & "','" & iScore(i, 2) & "','" & iScore(i, 3) & "','" & iScore(i, 4) & "','" & iScore(i, 5) & "','" & iScore(i, 6) & "','" & iScore(i, 7) & "','" & iScore(i, 8) & "','" & iScore(i, 9) & "','" & iScore(i, 10) & "','" & iScore(i, 11) & "','" & iScore(i, 12) & "','" & iScore(i, 13) & "','" & iScore(i, 14) & "','" & iScore(i, 15) & "','" & iScore(i, 16) & "','" & iScore(i, 17) & "','" & iScore(i, 18) & "','" & iScore(i, 19) & "')"
            rsMaster.Open strMaster, conn, adOpenDynamic, adLockOptimistic
        Next
        MsgBox "Teaching Evaluation by student completed." & vbCrLf & "Thank you.", vbInformation, "Jangid Corporation"
        Unload Me
        frmMain.Show
        If bShutDown Then
            Shell "Shutdown.exe -s -t 0"
        End If
        Exit Sub
    End If
    
    iSPtr = iSPtr + 1
    
    lblRemaining.Caption = "Total Subjects : " & iCount + 1 & vbCrLf & "Subject Completed : " & iSPtr
    frameFeed.Caption = "SUBJECT CODE :  " & strSubj(iSPtr, 0) & "    -    SUBJECT NAME :  " & strSubj(iSPtr, 1) & "    -    STAFF NAME :  " & strSubj(iSPtr, 3)
    
    Exit Sub
errHandler:
    If Err.Number = 3021 Then
        MsgBox "Insufficient data in the database." & vbCrLf & "Check whether department,staff and subject are created.", vbInformation, "Jangid Corporation"
    ElseIf Err.Number = 0 Then
    Else
        MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & "Error Location : frmFeed:cmdGo_Click", vbInformation, "Jangid Corporation"
    End If
End Sub
Private Sub rbClear()
    Dim i As Integer
    For i = 0 To 19
        ob1(i).Value = vbUnchecked
        ob2(i).Value = vbUnchecked
        ob3(i).Value = vbUnchecked
        ob4(i).Value = vbUnchecked
        ob5(i).Value = vbUnchecked
    Next
End Sub
Private Sub Form_Load()
    cmbDept_Load cmbDept
    cmbSem_Load cmbSem
    cmbSec_Load cmbSec
    cmbBatch_Load cmbBatch
    
    frameMain.Width = 3120
    frameMain.Left = (Screen.Width - frameMain.Width) / 2
    frameMain.Height = 4095
    frameMain.Top = (Screen.Height - frameMain.Height) / 2 - 480
    frameMain.Visible = True
End Sub

Private Sub Form_Paint()
    frame.Top = Screen.Height - frame.Height - 1024
    frame.Left = frameFour.Left + frameFour.Width - frame.Width
            
    frameFeed.Top = 240
    frameFeed.Left = 240
    frameFeed.Width = Screen.Width - 480
    frameFeed.Height = Screen.Height - 1440
    frameOne.Left = (frameFeed.Width - (2 * frameOne.Width) - 225) / 2
    frameThree.Left = frameOne.Left + frameOne.Width + 225
    frameTwo.Left = (frameFeed.Width - (2 * frameTwo.Width) - 225) / 2
    frameFour.Left = frameTwo.Left + frameTwo.Width + 225
    cmdGo.Left = frameFour.Left + frameFour.Width - cmdGo.Width
    lblPoints.Left = (Screen.Width - lblPoints.Width) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Show
End Sub

Private Sub cmbDept_Change()
    On Error Resume Next
    iDept = Department(cmbDept)
    HideCheckBox
End Sub
Private Sub cmbDept_Click()
    On Error Resume Next
    iDept = Department(cmbDept)
    HideCheckBox
End Sub
Private Sub cmbBatch_Change()
    On Error Resume Next
    iBatch = cmbBatch.Text
    HideCheckBox
End Sub
Private Sub cmbBatch_Click()
    On Error Resume Next
    iBatch = cmbBatch.Text
    HideCheckBox
End Sub

Private Sub cmbSec_Change()
    On Error Resume Next
    strSec = cmbSec.Text
End Sub
Private Sub cmbSec_Click()
    On Error Resume Next
    strSec = cmbSec.Text
End Sub

Private Sub cmbSem_Change()
    On Error Resume Next
    iSem = Val(cmbSem.Text)
    HideCheckBox
End Sub

Private Sub cmbSem_Click()
    On Error Resume Next
    iSem = Val(cmbSem.Text)
    HideCheckBox
End Sub


'-----------------Automatic Answers-------------------------
Private Sub cmd1_Click()
    Dim i As Integer
    For i = 0 To 19
        ob1(i).Value = vbChecked
    Next
End Sub

Private Sub cmd2_Click()
    Dim i As Integer
    For i = 0 To 19
        ob2(i).Value = vbChecked
    Next
End Sub
Private Sub cmd3_Click()
    Dim i As Integer
    For i = 0 To 19
        ob3(i).Value = vbChecked
    Next
End Sub
Private Sub cmd4_Click()
    Dim i As Integer
    For i = 0 To 19
        ob4(i).Value = vbChecked
    Next
End Sub
Private Sub cmd5_Click()
    Dim i As Integer
    For i = 0 To 19
        ob5(i).Value = vbChecked
    Next
End Sub
Private Sub cmdFill_Click()
    Dim i As Integer
    Dim r As Integer
    For i = 0 To 19
        r = Round((Rnd() * 4) + 1)
        Select Case r
            Case 1
                ob1(i).Value = vbChecked
            Case 2
                ob2(i).Value = vbChecked
            Case 3
                ob3(i).Value = vbChecked
            Case 4
                ob4(i).Value = vbChecked
            Case 5
                ob5(i).Value = vbChecked
        End Select
    Next
End Sub

Private Sub cmdTrunc_Click()
    Dim rsTrunc As New ADODB.Recordset
    Dim strSql As String
    strSql = "truncate table master"
    rsTrunc.Open strSql, conn, adOpenDynamic, adLockOptimistic, -1
End Sub
'----------------------Automatic Ends Here-------------------------

