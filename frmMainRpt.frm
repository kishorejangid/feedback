VERSION 5.00
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkusercontrolsxp.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMainRpt 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Final Year Report Generator"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12480
   Icon            =   "frmMainRpt.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   12480
   StartUpPosition =   2  'CenterScreen
   Begin JURA2.StylerButton cmdSettings 
      Height          =   255
      Left            =   11160
      TabIndex        =   110
      Top             =   600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      Caption         =   "Settings"
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
   Begin JURA2.ucXTab tabMain 
      Height          =   7725
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   13626
      TabCount        =   5
      TabCaption(0)   =   "Tab 0"
      TabCaption(1)   =   "Tab 1"
      TabCaption(2)   =   "Tab 2"
      TabCaption(3)   =   "Tab 3"
      TabCaption(4)   =   "Tab 4"
      ActiveTab       =   3
      ActiveTabBackEndColor=   16514555
      ActiveTabBackStartColor=   16514555
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ActiveTabHeight =   40
      BackColor       =   16514555
      BottomRightInnerBorderColor=   10526880
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   10526880
      ForeColor       =   -2147483630
      InActiveTabBackEndColor=   15397104
      InActiveTabBackStartColor=   16777215
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      InActiveTabHeight=   25
      OuterBorderColor=   10198161
      PictureMaskColor=   16711935
      ShowFocusRect   =   0   'False
      TabStyle        =   1
      TabTheme        =   1
      TopLeftInnerBorderColor=   16777215
      UseMouseWheelScroll=   0   'False
      Begin vkUserContolsXP.vkFrame frameMarks 
         Height          =   7140
         Left            =   0
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   600
         Visible         =   0   'False
         Width           =   12000
         _ExtentX        =   21167
         _ExtentY        =   12594
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
         RoundAngle      =   0
         BorderWidth     =   2
         Begin vkUserContolsXP.vkLabel lblPassedOn 
            Height          =   195
            Left            =   6000
            TabIndex        =   109
            TabStop         =   0   'False
            Top             =   1440
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   344
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Passed On"
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
         Begin vkUserContolsXP.vkLabel lblMarks 
            Height          =   195
            Left            =   3600
            TabIndex        =   108
            TabStop         =   0   'False
            Top             =   1440
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   344
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Marks"
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
            Height          =   195
            Left            =   3480
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   840
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   344
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
         Begin vkUserContolsXP.vkCommand cmdSave 
            Height          =   495
            Left            =   10080
            TabIndex        =   51
            Top             =   6195
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            Caption         =   "Save"
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
         Begin vkUserContolsXP.vkCommand cmdGo 
            Height          =   375
            Left            =   8880
            TabIndex        =   30
            Top             =   720
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            Caption         =   "Go"
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
         Begin vkUserContolsXP.vkTextBox txtPassed 
            Height          =   360
            Index           =   9
            Left            =   5640
            TabIndex        =   50
            Top             =   6360
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   635
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtPassed 
            Height          =   360
            Index           =   8
            Left            =   5640
            TabIndex        =   48
            Top             =   5880
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   635
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtPassed 
            Height          =   360
            Index           =   7
            Left            =   5640
            TabIndex        =   46
            Top             =   5400
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   635
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtPassed 
            Height          =   360
            Index           =   6
            Left            =   5640
            TabIndex        =   44
            Top             =   4920
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   635
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtPassed 
            Height          =   360
            Index           =   5
            Left            =   5640
            TabIndex        =   42
            Top             =   4200
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   635
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtPassed 
            Height          =   360
            Index           =   4
            Left            =   5640
            TabIndex        =   40
            Top             =   3720
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   635
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtPassed 
            Height          =   360
            Index           =   3
            Left            =   5640
            TabIndex        =   38
            Top             =   3240
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   635
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtPassed 
            Height          =   360
            Index           =   2
            Left            =   5640
            TabIndex        =   36
            Top             =   2760
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   635
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtPassed 
            Height          =   360
            Index           =   1
            Left            =   5640
            TabIndex        =   34
            Top             =   2280
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   635
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtPassed 
            Height          =   360
            Index           =   0
            Left            =   5640
            TabIndex        =   32
            Top             =   1800
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   635
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtMarks 
            Height          =   360
            Index           =   9
            Left            =   3240
            TabIndex        =   49
            Top             =   6360
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   635
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtMarks 
            Height          =   360
            Index           =   8
            Left            =   3240
            TabIndex        =   47
            Top             =   5880
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   635
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtMarks 
            Height          =   360
            Index           =   7
            Left            =   3240
            TabIndex        =   45
            Top             =   5400
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   635
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtMarks 
            Height          =   360
            Index           =   6
            Left            =   3240
            TabIndex        =   43
            Top             =   4920
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   635
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtMarks 
            Height          =   360
            Index           =   5
            Left            =   3240
            TabIndex        =   41
            Top             =   4200
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   635
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtMarks 
            Height          =   360
            Index           =   4
            Left            =   3240
            TabIndex        =   39
            Top             =   3720
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   635
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtMarks 
            Height          =   360
            Index           =   3
            Left            =   3240
            TabIndex        =   37
            Top             =   3240
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   635
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtMarks 
            Height          =   360
            Index           =   2
            Left            =   3240
            TabIndex        =   35
            Top             =   2760
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   635
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtMarks 
            Height          =   360
            Index           =   1
            Left            =   3240
            TabIndex        =   33
            Top             =   2280
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   635
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtMarks 
            Height          =   360
            Index           =   0
            Left            =   3240
            TabIndex        =   31
            Top             =   1800
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   635
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkLabel lblSubj 
            Height          =   360
            Index           =   9
            Left            =   1320
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   6360
            Visible         =   0   'False
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   635
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
         Begin vkUserContolsXP.vkLabel lblSubj 
            Height          =   360
            Index           =   8
            Left            =   1320
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   5880
            Visible         =   0   'False
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   635
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
         Begin vkUserContolsXP.vkLabel lblSubj 
            Height          =   360
            Index           =   7
            Left            =   1320
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   5400
            Visible         =   0   'False
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   635
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
         Begin vkUserContolsXP.vkLabel lblSubj 
            Height          =   360
            Index           =   6
            Left            =   1320
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   4920
            Visible         =   0   'False
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   635
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
         Begin vkUserContolsXP.vkLabel lblPrac 
            Height          =   195
            Left            =   240
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   4560
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   344
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Practicals:"
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
         Begin vkUserContolsXP.vkLabel lblSubj 
            Height          =   360
            Index           =   5
            Left            =   1320
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   4200
            Visible         =   0   'False
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   635
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
         Begin vkUserContolsXP.vkLabel lblSubj 
            Height          =   360
            Index           =   4
            Left            =   1320
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   3720
            Visible         =   0   'False
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   635
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
         Begin vkUserContolsXP.vkLabel lblSubj 
            Height          =   360
            Index           =   3
            Left            =   1320
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   3240
            Visible         =   0   'False
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   635
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
         Begin vkUserContolsXP.vkLabel lblSubj 
            Height          =   360
            Index           =   2
            Left            =   1320
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   2760
            Visible         =   0   'False
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   635
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
         Begin vkUserContolsXP.vkLabel lblSubj 
            Height          =   360
            Index           =   1
            Left            =   1320
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   2280
            Visible         =   0   'False
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   635
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
         Begin vkUserContolsXP.vkLabel lblSubj 
            Height          =   360
            Index           =   0
            Left            =   1320
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   1800
            Visible         =   0   'False
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   635
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
         Begin vkUserContolsXP.vkLabel lblTheory 
            Height          =   315
            Left            =   240
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   1440
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   556
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Theory:"
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
            Left            =   6480
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   360
            Width           =   615
            _ExtentX        =   1085
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
         Begin vkUserContolsXP.vkLabel lblSem 
            Height          =   255
            Left            =   6480
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   840
            Width           =   735
            _ExtentX        =   1296
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
         Begin vkUserContolsXP.vkLabel lblRegNo 
            Height          =   315
            Left            =   240
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   840
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   556
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Reg No:"
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
            Height          =   315
            Left            =   240
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   360
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   556
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
            Left            =   7320
            TabIndex        =   29
            Top             =   720
            Width           =   1215
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
            Left            =   7320
            TabIndex        =   27
            Top             =   240
            Width           =   1215
         End
         Begin VB.ComboBox cmbRegNo 
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
            TabIndex        =   28
            Top             =   720
            Width           =   2055
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
            TabIndex        =   26
            Top             =   240
            Width           =   4815
         End
         Begin VB.ComboBox cmbPassed 
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
            Left            =   10200
            TabIndex        =   106
            Top             =   1680
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Line line 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   2
            X1              =   0
            X2              =   12000
            Y1              =   1320
            Y2              =   1320
         End
      End
      Begin vkUserContolsXP.vkFrame frameReport 
         Height          =   7095
         Left            =   -10000
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   600
         Visible         =   0   'False
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   12515
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
         RoundAngle      =   0
         BorderWidth     =   2
         Begin vkUserContolsXP.vkCommand cmdReport 
            Height          =   375
            Left            =   240
            TabIndex        =   55
            Top             =   2280
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   661
            Caption         =   "Generate Report"
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
         Begin vkUserContolsXP.vkLabel lblRptName 
            Height          =   195
            Left            =   3600
            TabIndex        =   79
            Top             =   1680
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   344
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
         Begin vkUserContolsXP.vkLabel lblRptBatch 
            Height          =   255
            Left            =   240
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   1080
            Width           =   615
            _ExtentX        =   1085
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
         Begin vkUserContolsXP.vkLabel lblRptRegNo 
            Height          =   315
            Left            =   240
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   1680
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   556
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Reg No:"
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
         Begin vkUserContolsXP.vkLabel lblRptDept 
            Height          =   315
            Left            =   240
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   480
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   556
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
         Begin VB.ComboBox cmbRptBatch 
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
            TabIndex        =   53
            Top             =   960
            Width           =   2055
         End
         Begin VB.ComboBox cmbRptRegNo 
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
            TabIndex        =   54
            Top             =   1560
            Width           =   2055
         End
         Begin VB.ComboBox cmbRptDept 
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
            TabIndex        =   52
            Top             =   360
            Width           =   4815
         End
      End
      Begin vkUserContolsXP.vkFrame frameSubj 
         Height          =   5175
         Left            =   10000
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   600
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   9128
         BackColor1      =   16777215
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ShowTitle       =   0   'False
         TitleColor1     =   8438015
         TitleColor2     =   33023
         TitleGradient   =   2
         TitleHeight     =   300
         RoundAngle      =   0
         BorderWidth     =   2
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
            Left            =   1680
            TabIndex        =   23
            Top             =   3120
            Width           =   4455
         End
         Begin vkUserContolsXP.vkLabel lblSubjBatch 
            Height          =   375
            Left            =   480
            TabIndex        =   85
            TabStop         =   0   'False
            Top             =   3240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
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
         Begin vkUserContolsXP.vkCommand cmdSubjInsert 
            Height          =   510
            Left            =   480
            TabIndex        =   25
            Top             =   4320
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   900
            Caption         =   "Insert"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16744576
            BorderColor     =   16744576
            CustomStyle     =   0
         End
         Begin vkUserContolsXP.vkTextBox txtSubjName 
            Height          =   375
            Left            =   1680
            TabIndex        =   20
            Top             =   1320
            Width           =   4455
            _ExtentX        =   7858
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkLabel lblSubjSem 
            Height          =   375
            Left            =   480
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   2160
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
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
         Begin vkUserContolsXP.vkLabel lblSubjDept 
            Height          =   375
            Left            =   480
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   2640
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
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
            Left            =   1680
            TabIndex        =   21
            Top             =   1920
            Width           =   4455
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
            Left            =   1680
            TabIndex        =   22
            Top             =   2520
            Width           =   4455
         End
         Begin vkUserContolsXP.vkLabel lblSubjSubjName 
            Height          =   375
            Left            =   480
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   1440
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
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
         Begin vkUserContolsXP.vkTextBox txtSubjCode 
            Height          =   375
            Left            =   1680
            TabIndex        =   19
            Top             =   720
            Width           =   4455
            _ExtentX        =   7858
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkLabel lblSubjSubjCode 
            Height          =   255
            Left            =   480
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   840
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
         Begin vkUserContolsXP.vkCheck cbIsLab 
            Height          =   375
            Left            =   1680
            TabIndex        =   24
            Top             =   3720
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   661
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Is Lab"
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
      Begin vkUserContolsXP.vkFrame frameContact 
         Height          =   7065
         Left            =   20000
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   600
         Width           =   11970
         _ExtentX        =   21114
         _ExtentY        =   12462
         Caption         =   "Contact"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ShowTitle       =   0   'False
         TitleColor1     =   33023
         TitleColor2     =   8438015
         TitleGradient   =   2
         TitleHeight     =   300
         RoundAngle      =   0
         BorderWidth     =   2
         Begin MSComDlg.CommonDialog Dialog1 
            Left            =   8640
            Top             =   600
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin vkUserContolsXP.vkCommand cmdImage 
            Height          =   495
            Left            =   9600
            TabIndex        =   17
            Top             =   2760
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   873
            Caption         =   "Add Image"
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
         Begin vkUserContolsXP.vkCommand cmdStudSave 
            Height          =   495
            Left            =   9600
            TabIndex        =   18
            Top             =   6240
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   873
            Caption         =   "Insert"
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
         Begin vkUserContolsXP.vkTextBox txtStudMobile 
            Height          =   375
            Left            =   5400
            TabIndex        =   16
            Top             =   6360
            Width           =   1815
            _ExtentX        =   3201
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtStudPhone 
            Height          =   375
            Left            =   1920
            TabIndex        =   15
            Top             =   6360
            Width           =   1815
            _ExtentX        =   3201
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtStudState 
            Height          =   375
            Left            =   1920
            TabIndex        =   14
            Top             =   5760
            Width           =   1815
            _ExtentX        =   3201
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtStudPinCode 
            Height          =   375
            Left            =   5400
            TabIndex        =   13
            Top             =   5160
            Width           =   1815
            _ExtentX        =   3201
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtStudCity 
            Height          =   375
            Left            =   1920
            TabIndex        =   12
            Top             =   5160
            Width           =   1815
            _ExtentX        =   3201
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtStudAddress 
            Height          =   375
            Left            =   1920
            TabIndex        =   11
            Top             =   4560
            Width           =   5295
            _ExtentX        =   9340
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtStudOccupation 
            Height          =   375
            Left            =   1920
            TabIndex        =   10
            Top             =   3960
            Width           =   1815
            _ExtentX        =   3201
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtStudParent 
            Height          =   375
            Left            =   1920
            TabIndex        =   9
            Top             =   3360
            Width           =   5295
            _ExtentX        =   9340
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtStudIncome 
            Height          =   375
            Left            =   5400
            TabIndex        =   8
            Top             =   2760
            Width           =   1815
            _ExtentX        =   3201
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtStudDoJ 
            Height          =   375
            Left            =   1920
            TabIndex        =   7
            Top             =   2760
            Width           =   1815
            _ExtentX        =   3201
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
            LegendForeColor =   16750899
         End
         Begin VB.ComboBox cmbStudCaste 
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
            Left            =   5400
            TabIndex        =   6
            Top             =   2160
            Width           =   1815
         End
         Begin vkUserContolsXP.vkTextBox txtStudDoB 
            Height          =   375
            Left            =   1920
            TabIndex        =   5
            Top             =   2160
            Width           =   1815
            _ExtentX        =   3201
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
            LegendForeColor =   16750899
         End
         Begin VB.ComboBox cmbStudGender 
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
            ItemData        =   "frmMainRpt.frx":32D2
            Left            =   5400
            List            =   "frmMainRpt.frx":32D4
            TabIndex        =   4
            Top             =   1560
            Width           =   1815
         End
         Begin VB.ComboBox cmbStudSec 
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
            Left            =   1920
            TabIndex        =   3
            Top             =   1560
            Width           =   1815
         End
         Begin vkUserContolsXP.vkTextBox txtStudName 
            Height          =   375
            Left            =   1920
            TabIndex        =   2
            Top             =   960
            Width           =   5295
            _ExtentX        =   9340
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkTextBox txtStudRegNo 
            Height          =   375
            Left            =   1920
            TabIndex        =   1
            Top             =   360
            Width           =   5295
            _ExtentX        =   9340
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
            LegendForeColor =   16750899
         End
         Begin vkUserContolsXP.vkLabel lblStudIncome 
            Height          =   195
            Left            =   4200
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   2880
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   344
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Annual Income:"
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
         Begin vkUserContolsXP.vkLabel lblStudCaste 
            Height          =   195
            Left            =   4200
            TabIndex        =   101
            TabStop         =   0   'False
            Top             =   2280
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   344
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Caste:"
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
         Begin vkUserContolsXP.vkLabel lblStudDoJ 
            Height          =   315
            Left            =   360
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   2880
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Date of Joining:"
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
         Begin vkUserContolsXP.vkLabel lblStudSec 
            Height          =   195
            Left            =   360
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   1680
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   344
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
         Begin VB.PictureBox Picture 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   2175
            Left            =   9600
            ScaleHeight     =   2145
            ScaleWidth      =   2025
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   360
            Width           =   2055
            Begin VB.Image imgStud 
               BorderStyle     =   1  'Fixed Single
               Height          =   2175
               Left            =   0
               Picture         =   "frmMainRpt.frx":32D6
               Stretch         =   -1  'True
               Top             =   0
               Width           =   2055
            End
         End
         Begin vkUserContolsXP.vkLabel lblStudGender 
            Height          =   195
            Left            =   4200
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   1680
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   344
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Gender:"
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
         Begin vkUserContolsXP.vkLabel lblStudMobile 
            Height          =   195
            Left            =   4200
            TabIndex        =   97
            TabStop         =   0   'False
            Top             =   6480
            Width           =   540
            _ExtentX        =   900
            _ExtentY        =   344
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Mobile:"
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
         Begin vkUserContolsXP.vkLabel lblStudPhone 
            Height          =   195
            Left            =   360
            TabIndex        =   96
            TabStop         =   0   'False
            Top             =   6480
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   344
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Phone:"
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
         Begin vkUserContolsXP.vkLabel lblStudState 
            Height          =   195
            Left            =   360
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   5880
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   344
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "State:"
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
         Begin vkUserContolsXP.vkLabel lblStudPincode 
            Height          =   195
            Left            =   4200
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   5280
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   344
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "PinCode:"
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
         Begin vkUserContolsXP.vkLabel lblStudCity 
            Height          =   195
            Left            =   360
            TabIndex        =   93
            TabStop         =   0   'False
            Top             =   5280
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   344
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
         Begin vkUserContolsXP.vkLabel lblStudAddress 
            Height          =   195
            Left            =   360
            TabIndex        =   92
            TabStop         =   0   'False
            Top             =   4680
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   344
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Address:"
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
         Begin vkUserContolsXP.vkLabel lblStudOccupation 
            Height          =   195
            Left            =   360
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   4080
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   344
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "&Occupation:"
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
         Begin vkUserContolsXP.vkLabel lblStudRegNo 
            Height          =   315
            Left            =   360
            TabIndex        =   90
            TabStop         =   0   'False
            Top             =   480
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   556
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Reg No:"
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
         Begin vkUserContolsXP.vkLabel lblStudParent 
            Height          =   315
            Left            =   360
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   3480
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Parent/Gaurdian:"
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
         Begin vkUserContolsXP.vkLabel lblStudDoB 
            Height          =   195
            Left            =   360
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   2280
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   344
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "D.o.B:"
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
         Begin vkUserContolsXP.vkLabel lblStudName 
            Height          =   195
            Left            =   360
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   1080
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   344
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "&Name:"
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
      Begin vkUserContolsXP.vkFrame frameIntro 
         Height          =   7095
         Left            =   30000
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   600
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   12515
         Caption         =   "MainForm"
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
         RoundAngle      =   0
         BorderWidth     =   2
         Begin vkUserContolsXP.vkLabel lblCollegeCity 
            Height          =   390
            Left            =   0
            TabIndex        =   105
            TabStop         =   0   'False
            Top             =   840
            Width           =   11895
            _ExtentX        =   20981
            _ExtentY        =   688
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Tirunelveli-3"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16744576
            Alignment       =   2
         End
         Begin vkUserContolsXP.vkLabel lblCollegeName 
            Height          =   615
            Left            =   0
            TabIndex        =   104
            TabStop         =   0   'False
            Top             =   240
            Width           =   11895
            _ExtentX        =   20981
            _ExtentY        =   1085
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Francis Xavier Engineering College"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16744576
            Alignment       =   2
         End
      End
   End
   Begin JURA2.ThemedComboBox ThemedComboBox 
      Left            =   0
      Top             =   0
      _ExtentX        =   556
      _ExtentY        =   529
      BorderColorStyle=   1
      ComboBoxBorderColor=   16744576
      DriveListBoxBorderColor=   16744576
      ImageComboBorderColor=   16744576
   End
   Begin vkUserContolsXP.vkLabel dignaj 
      Height          =   195
      Left            =   10920
      TabIndex        =   107
      Top             =   8000
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   344
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "Help && Support"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
End
Attribute VB_Name = "frmMainRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim itxtPassedIndex As Integer
Dim sImageName As String
Dim iSubCount As Integer

'Function to call the default browser
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOW = 5

Public Sub OpenWebsite(URL As String)
    Dim ret&
    ret& = ShellExecute(Me.hWnd, "Open", URL, vbNullString, vbNullString, SW_SHOW)
End Sub

Private Sub cmbBatch_Change()
    iBatch = cmbBatch.Text
    cmbRegNo_Load cmbRegNo
End Sub
Private Sub cmbBatch_Click()
    iBatch = cmbBatch.Text
    cmbRegNo_Load cmbRegNo
End Sub
Private Sub cmbSubjBatch_Change()
    iBatch = cmbSubjBatch.Text
End Sub
Private Sub cmbSubjBatch_Click()
    iBatch = cmbSubjBatch.Text
End Sub

Private Sub cmbRptBatch_Change()
    iBatch = cmbRptBatch.Text
    cmbRegNo_Load cmbRptRegNo
End Sub

Private Sub cmbRptBatch_Click()
    iBatch = cmbRptBatch.Text
    cmbRegNo_Load cmbRptRegNo
End Sub


Private Sub cmbDept_Change()
    iDept = Department(cmbDept)
    Clear
    cmbRegNo_Load cmbRegNo
End Sub
Private Sub cmbDept_Click()
    iDept = Department(cmbDept)
    Clear
    cmbRegNo_Load cmbRegNo
End Sub

Private Sub cmbSubjDept_Change()
    iDept = Department(cmbSubjDept)
End Sub
Private Sub cmbSubjDept_Click()
    iDept = Department(cmbSubjDept)
End Sub

Private Sub cmbRptDept_Change()
    iDept = Department(cmbRptDept)
    cmbRegNo_Load cmbRptRegNo
End Sub

Private Sub cmbRptDept_Click()
    iDept = Department(cmbRptDept)
    cmbRegNo_Load cmbRptRegNo
End Sub

Private Sub cmbSem_Change()
    iSem = Val(cmbSem.Text)
    HideControls
    Clear
End Sub

Private Sub cmbSem_Click()
    iSem = Val(cmbSem.Text)
    HideControls
    Clear
End Sub

Private Sub cmbSubjSem_Change()
    iSem = Val(cmbSubjSem.Text)
End Sub

Private Sub cmbSubjSem_Click()
    iSem = Val(cmbSubjSem.Text)
End Sub


Private Sub cmbRegNo_Change()
    lblName.Caption = GetStudName(cmbRegNo.Text)
End Sub
Private Sub cmbRegNo_Click()
    lblName.Caption = GetStudName(cmbRegNo.Text)
End Sub

Private Sub cmbRptRegNo_Change()
    lblRptName.Caption = GetStudName(cmbRptRegNo.Text)
End Sub
Private Sub cmbRptRegNo_Click()
    lblRptName.Caption = GetStudName(cmbRptRegNo.Text)
End Sub

Private Sub cmbPassed_LostFocus()
    On Error Resume Next
    txtPassed(itxtPassedIndex).Text = cmbPassed.Text
    cmbPassed.Visible = False
    If itxtPassedIndex = iSubCount - 1 Then
        cmdSave.SetFocus
    End If
    txtMarks(itxtPassedIndex + 1).SetFocus
End Sub


Private Sub cmdGo_Click()
    On Error Resume Next
    Dim rsTheory As New ADODB.Recordset
    Dim rsLab As New ADODB.Recordset
    Dim sqlTheory As String
    Dim sqlLab As String
    Dim iTheory, iLab As Integer
    iTheory = 0
    iLab = 6
    iSubCount = 0
    sqlTheory = "select subjcode from subj where dept='" & iDept & "' and batch='" & Mid(iBatch, 3, 2) & "' and semno='" & iSem & "' and lab=0 "
    sqlLab = "select subjcode from subj where dept='" & iDept & "' and batch='" & Mid(iBatch, 3, 2) & "' and semno='" & iSem & "' and lab=1 "
    rsTheory.Open sqlTheory, conn, adOpenDynamic, adLockOptimistic, -1
    rsLab.Open sqlLab, conn, adOpenDynamic, adLockOptimistic, -1
    Do While Not rsTheory.EOF
        lblSubj(iTheory).Caption = rsTheory.Fields(0)
        lblSubj(iTheory).Visible = True
        txtMarks(iTheory).Visible = True
        txtPassed(iTheory).Visible = True
        iTheory = iTheory + 1
        iSubCount = iSubCount + 1
        rsTheory.MoveNext
    Loop
    Do While Not rsLab.EOF
        lblSubj(iLab).Caption = rsLab.Fields(0)
        lblSubj(iLab).Visible = True
        txtMarks(iLab).Visible = True
        txtPassed(iLab).Visible = True
        iLab = iLab + 1
        iSubCount = iSubCount + 1
        rsLab.MoveNext
    Loop
    If Err.Number <> 0 Then
        MsgBox Error & vbCrLf & "Error Number: " & Err.Number
    End If
End Sub

Private Sub HideControls()
    Dim i As Integer
    For i = 0 To 9
        lblSubj(i).Visible = False
        txtMarks(i).Visible = False
        txtPassed(i).Visible = False
    Next
End Sub


Private Sub cmdImage_Click()
With Dialog1
       .InitDir = App.Path
       .Filter = "JPEG image|*.jpg|GIF image|*.gif|BITMAP image|*.bmp|Icon image|*.ico|Cursor image|*.cur|Panerio image|*.pan"
       .ShowOpen
          If .FileName <> "" Then
             sImageName = .FileName
             imgStud.Picture = LoadPicture(sImageName)
          End If
     End With
End Sub

Private Sub cmdReport_Click()
    On Error Resume Next
    
    If cmbRptRegNo.Text = "" Then
        Exit Sub
    End If
    
    
    Dim PDF As New clsPDF
    Dim strLeft As Double
    Dim l As Double
    Dim strCollegeName As String
    Dim strDept As String
    Dim iRectno As Integer
    Dim x As Double
    Dim i As Double
    Dim y As Double
    Dim strSem(1 To 8) As String
    Dim iTotal As Integer
    Dim iSubj As Integer
    Dim flgArrear As Boolean
    Dim dCumPercentage As Double
    
    flgArrear = False
    dCumPercentage = 0
    
    PDF.PDFTitle = "Report"
    PDF.PDFFileName = App.Path & "\Reports\" & "Report" & ".pdf"
    PDF.PDFAuthor = "JURA"
    PDF.PDFLoadAfm = App.Path & "\Fonts"
    PDF.PDFView = True
    
    PDF.PDFSetUnit = UNIT_CM
    PDF.PDFFormatPage = FORMAT_A3
    PDF.PDFOrientation = ORIENT_PAYSAGE
    
    PDF.PDFSetBorder = BORDER_ALL
    PDF.PDFSetTopMargin = 1
    PDF.PDFSetBottomMargin = 1
    PDF.PDFSetRightMargin = 1
    PDF.PDFSetLeftMargin = 1
    
    PDF.PDFBeginDoc
        
        'Main Border
        PDF.PDFDrawRectangle 0.5, 0.5, 41, 28.7
        
        'Header
        PDF.PDFSetTextColor = vbBlack
        PDF.PDFSetFont FONT_TIMES, 20, FONT_BOLD
        
        strCollegeName = "FRANCIS XAVIER ENGINEERING COLLEGE, TIRUNELVELI-3"
        strDept = "DEPARTMENT OF " & UCase(cmbRptDept.Text)
        
        l = PDF.PDFGetStringWidth(strCollegeName, "Times-Bold", 20)
        strLeft = (41 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut strCollegeName, strLeft, 1.25
        
        PDF.PDFSetFont FONT_TIMES, 18, FONT_BOLD
        
        l = PDF.PDFGetStringWidth(strDept, "Times-Bold", 18)
        strLeft = (41 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut strDept, strLeft, 1.9
        
        l = PDF.PDFGetStringWidth("PERSONAL DETAILS", "Times-Bold", 18)
        strLeft = (41 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut "PERSONAL DETAILS", strLeft, 2.6
        
        'Student Information Border
        PDF.PDFSetLineWidth = 0.01
        PDF.PDFDrawLineHor 1.5, 3, 0
        PDF.PDFDrawRectangle 1.5, 3, 34, 3.5
        PDF.PDFDrawRectangle 36.5, 2, 4, 4.5 'For Photo
        
        'Student Information Fetching
        PDF.PDFSetTextColor = vbBlack
        PDF.PDFSetFont FONT_TIMES, 12, FONT_NORMAL
        
        Dim rs As New ADODB.Recordset
        Dim sql As String
        sql = "select * from studdetails where regno='" & cmbRptRegNo.Text & "'"
        rs.Open sql, conn, adOpenDynamic, adLockOptimistic, -1
        
        PDF.PDFTextOut "Name:", 0.85, 3.6
        PDF.PDFTextOut "Class:", 0.85, 4.25
        PDF.PDFTextOut "D.o.B:", 0.85, 4.9
        PDF.PDFTextOut "Caste:", 0.85, 5.55
        
        PDF.PDFTextOut rs.Fields("STUDNAME"), 3.85, 3.6
        PDF.PDFTextOut rs.Fields("SEC"), 3.85, 4.25
        PDF.PDFTextOut rs.Fields("DOB"), 3.85, 4.9
        PDF.PDFTextOut rs.Fields("CASTE"), 3.85, 5.55
        
        PDF.PDFTextOut "Reg No:", 10.85, 3.6
        PDF.PDFTextOut "Date of Joining:", 10.85, 4.25
        PDF.PDFTextOut "Gender:", 10.85, 4.9
        PDF.PDFTextOut "Address:", 10.85, 5.55
        
        PDF.PDFTextOut rs.Fields("REGNO"), 14.85, 3.6
        PDF.PDFTextOut rs.Fields("DOJ"), 14.85, 4.25
        PDF.PDFTextOut rs.Fields("GENDER"), 14.85, 4.9
        PDF.PDFTextOut rs.Fields("ADDRESS"), 14.85, 5.55
        PDF.PDFTextOut rs.Fields("CITY") & "-" & rs.Fields("PINCODE") & "," & rs.Fields("STATE"), 14.85, 6.1
        
        
        PDF.PDFTextOut "Parent/Gaurdian Name:", 20.85, 3.6
        PDF.PDFTextOut "Occupation:", 20.85, 4.25
        PDF.PDFTextOut "Annual Income:", 20.85, 4.9
        PDF.PDFTextOut "Contact No:", 20.85, 5.55
        
        PDF.PDFTextOut rs.Fields("FATHER"), 26.85, 3.6
        PDF.PDFTextOut rs.Fields("OCCUPATION"), 26.85, 4.25
        PDF.PDFTextOut rs.Fields("INCOME"), 26.85, 4.9
        PDF.PDFTextOut rs.Fields("LANDLINE") & " / " & rs.Fields("MOBILE"), 26.85, 5.55
        
        PDF.PDFImage App.Path & "\Images\" & rs.Fields("IMAGE"), 36.6, 2.1, 3.8, 4.3
        
        x = 1.5
        strSem(1) = "I SEMESTER"
        strSem(2) = "II SEMESTER"
        strSem(3) = "III SEMESTER"
        strSem(4) = "IV SEMESTER"
        strSem(5) = "V SEMESTER"
        strSem(6) = "VI SEMESTER"
        strSem(7) = "VII SEMESTER"
        strSem(8) = "VIII SEMESTER"
        'Rectangle for semester
        For iRectno = 1 To 8
            If (iRectno Mod 2) = 1 Then
                y = 7
                
                PDF.PDFDrawRectangle x, y, 9.375, 9.625 'Semester Rectangle
                
                PDF.PDFDrawLineVer x + 1.5, y + 0.6875, 8.25 'VLine for subjcode
                PDF.PDFDrawLineVer x + 6, y + 0.6875, 8.25 'VLine for subjname
                PDF.PDFDrawLineVer x + 7, y + 0.6875, 8.25 'VLine for Marks
                PDF.PDFDrawLineVer x + 7.75, y + 0.6875, 8.25 'VLine for Pass/Fail
                
                'Hori Lines for Marks
                For i = y + 0.6875 To y + 9.375 Step 0.6875
                    PDF.PDFDrawLineHor x, i, 9.375
                Next
                
                PDF.PDFSetTextColor = vbBlack
                PDF.PDFSetFont FONT_TIMES, 10, FONT_BOLD
                
                PDF.PDFTextOut strSem(iRectno), x + 2.5, y + 0.5
                PDF.PDFTextOut "S Code", x - 0.85, y + 1.2
                PDF.PDFTextOut "Subj Name", x + 2, y + 1.2
                PDF.PDFSetFont FONT_TIMES, 8, FONT_BOLD
                PDF.PDFTextOut "Marks", x + 5.1, y + 1.2
                PDF.PDFSetFont FONT_TIMES, 10, FONT_BOLD
                PDF.PDFTextOut "P/F", x + 6.15, y + 1.2
                PDF.PDFTextOut "Result", x + 7.1, y + 1.2
                PDF.PDFTextOut "Lab", x + 2.4, y + 6
                PDF.PDFTextOut "Percentage:", x - 0.65, y + 9.5
                PDF.PDFTextOut "Total:", x + 4, y + 9.5
                
                'Fetching Data
                Dim rsTheoryOdd As New ADODB.Recordset
                Dim rsLabOdd As New ADODB.Recordset
                Dim sqlSemTheoryOdd As String
                Dim sqlSemLabOdd As String
                sqlSemTheoryOdd = "select s1.subjcode,s2.subjname,s1.marks,s1.passed from studmarks s1,subj s2 where s1.subjcode=s2.subjcode and s1.regno='" & cmbRptRegNo.Text & "' and s2.lab=0 and s1.semno=" & iRectno & "  "
                sqlSemLabOdd = "select s1.subjcode,s2.subjname,s1.marks,s1.passed from studmarks s1,subj s2 where s1.subjcode=s2.subjcode and s1.regno='" & cmbRptRegNo.Text & "' and s2.lab=1 and s1.semno=" & iRectno & " "
                rsTheoryOdd.Open sqlSemTheoryOdd, conn, adOpenDynamic, adLockOptimistic, -1
                rsLabOdd.Open sqlSemLabOdd, conn, adOpenDynamic, adLockOptimistic, -1
                
                Dim dConsOdd As Double
                dConsOdd = 0
                
                iTotal = 0
                iSubj = 0
                
                Do While Not rsTheoryOdd.EOF
                    PDF.PDFTextOut rsTheoryOdd.Fields(0), x - 0.9, y + 1.9 + dConsOdd
                    PDF.PDFSetFont FONT_TIMES, 8, FONT_BOLD
                    PDF.PDFTextOut Mid(rsTheoryOdd.Fields(1), 1, 29), x + 0.6, y + 1.9 + dConsOdd
                    PDF.PDFSetFont FONT_TIMES, 10, FONT_BOLD
                    
                    iTotal = iTotal + rsTheoryOdd.Fields(2)
                    iSubj = iSubj + 1
                    
                    If rsTheoryOdd.Fields(2) >= 50 Then
                        PDF.PDFTextOut rsTheoryOdd.Fields(2), x + 5.15, y + 1.9 + dConsOdd
                        PDF.PDFTextOut "P", x + 6.25, y + 1.9 + dConsOdd
                        PDF.PDFTextOut rsTheoryOdd.Fields(3), x + 6.8, y + 1.9 + dConsOdd
                    Else
                        flgArrear = True
                    End If
                    dConsOdd = dConsOdd + 0.6875
                    rsTheoryOdd.MoveNext
                Loop
                dConsOdd = dConsOdd + 0.6875
                Do While Not rsLabOdd.EOF
                    PDF.PDFTextOut rsLabOdd.Fields(0), x - 0.9, y + 1.9 + dConsOdd
                    PDF.PDFSetFont FONT_TIMES, 8, FONT_BOLD
                    PDF.PDFTextOut Mid(rsLabOdd.Fields(1), 1, 29), x + 0.6, y + 1.9 + dConsOdd
                    PDF.PDFSetFont FONT_TIMES, 10, FONT_BOLD
                    
                    iTotal = iTotal + rsLabOdd.Fields(2)
                    iSubj = iSubj + 1
                    
                    If rsLabOdd.Fields(2) >= 50 Then
                        PDF.PDFTextOut rsLabOdd.Fields(2), x + 5.15, y + 1.9 + dConsOdd
                        PDF.PDFTextOut "P", x + 6.25, y + 1.9 + dConsOdd
                        PDF.PDFTextOut rsLabOdd.Fields(3), x + 6.8, y + 1.9 + dConsOdd
                    Else
                        flgArrear = True
                    End If
                    dConsOdd = dConsOdd + 0.6875
                    rsLabOdd.MoveNext
                Loop
                rsLabOdd.Close
                rsTheoryOdd.Close
                
                If Not flgArrear Then
                    PDF.PDFTextOut CStr(iTotal), x + 5.15, y + 9.5
                    PDF.PDFTextOut CStr(Round(iTotal / iSubj, 2)) & "%", x + 1.5, y + 9.5
                    dCumPercentage = dCumPercentage + Round(iTotal / iSubj, 2)
                End If
            Else
                y = 17.125
                
                PDF.PDFDrawRectangle x, y, 9.375, 9.625 'Semester Rectangle
    
                PDF.PDFDrawLineVer x + 1.5, y + 0.6875, 8.25 'VLine for subjcode
                PDF.PDFDrawLineVer x + 6, y + 0.6875, 8.25 'VLine for subjname
                PDF.PDFDrawLineVer x + 7, y + 0.6875, 8.25 'VLine for Marks
                PDF.PDFDrawLineVer x + 7.75, y + 0.6875, 8.25 'VLine for Pass/Fail
                
                'Hori Lines for Marks
                For i = y + 0.6875 To y + 9.375 Step 0.6875
                    PDF.PDFDrawLineHor x, i, 9.375
                Next
                
                PDF.PDFSetTextColor = vbBlack
                PDF.PDFSetFont FONT_TIMES, 10, FONT_BOLD
                
                PDF.PDFTextOut strSem(iRectno), x + 2.5, y + 0.5
                PDF.PDFTextOut "S Code", x - 0.85, y + 1.2
                PDF.PDFTextOut "Subj Name", x + 2, y + 1.2
                PDF.PDFSetFont FONT_TIMES, 8, FONT_BOLD
                PDF.PDFTextOut "Marks", x + 5.1, y + 1.2
                PDF.PDFSetFont FONT_TIMES, 10, FONT_BOLD
                PDF.PDFTextOut "P/F", x + 6.15, y + 1.2
                PDF.PDFTextOut "Result", x + 7.1, y + 1.2
                PDF.PDFTextOut "Lab", x + 2.4, y + 6
                PDF.PDFTextOut "Percentage:", x - 0.65, y + 9.5
                PDF.PDFTextOut "Total:", x + 4, y + 9.5
                
                
                'Fetching Data
                Dim rsTheoryEven As New ADODB.Recordset
                Dim rsLabEven As New ADODB.Recordset
                Dim sqlSemTheoryEven As String
                Dim sqlSemLabEven As String
                sqlSemTheoryEven = "select s1.subjcode,s2.subjname,s1.marks,s1.passed from studmarks s1,subj s2 where s1.subjcode=s2.subjcode and s1.regno='" & cmbRptRegNo.Text & "' and s2.lab=0 and s1.semno=" & iRectno & "  "
                sqlSemLabEven = "select s1.subjcode,s2.subjname,s1.marks,s1.passed from studmarks s1,subj s2 where s1.subjcode=s2.subjcode and s1.regno='" & cmbRptRegNo.Text & "' and s2.lab=1 and s1.semno=" & iRectno & " "
                rsTheoryEven.Open sqlSemTheoryEven, conn, adOpenDynamic, adLockOptimistic, -1
                rsLabEven.Open sqlSemLabEven, conn, adOpenDynamic, adLockOptimistic, -1
                
                Dim dConsEven As Double
                dConsEven = 0
                
                iTotal = 0
                iSubj = 0
                
                Do While Not rsTheoryEven.EOF
                    PDF.PDFTextOut rsTheoryEven.Fields(0), x - 0.9, y + 1.9 + dConsEven
                    PDF.PDFSetFont FONT_TIMES, 8, FONT_BOLD
                    PDF.PDFTextOut Mid(rsTheoryEven.Fields(1), 1, 29), x + 0.6, y + 1.9 + dConsEven
                    PDF.PDFSetFont FONT_TIMES, 10, FONT_BOLD
                    
                    iTotal = iTotal + rsTheoryEven.Fields(2)
                    iSubj = iSubj + 1
                    
                    If rsTheoryEven.Fields(2) >= 50 Then
                        PDF.PDFTextOut rsTheoryEven.Fields(2), x + 5.15, y + 1.9 + dConsEven
                        PDF.PDFTextOut "P", x + 6.25, y + 1.9 + dConsEven
                        PDF.PDFTextOut rsTheoryEven.Fields(3), x + 6.8, y + 1.9 + dConsEven
                    Else
                        flgArrear = True
                    End If
                    dConsEven = dConsEven + 0.6875
                    rsTheoryEven.MoveNext
                Loop
                dConsEven = dConsEven + 0.6875
                Do While Not rsLabEven.EOF
                    PDF.PDFTextOut rsLabEven.Fields(0), x - 0.9, y + 1.9 + dConsEven
                    PDF.PDFSetFont FONT_TIMES, 8, FONT_BOLD
                    PDF.PDFTextOut Mid(rsLabEven.Fields(1), 1, 29), x + 0.6, y + 1.9 + dConsEven
                    PDF.PDFSetFont FONT_TIMES, 10, FONT_BOLD
                    
                    iTotal = iTotal + rsLabEven.Fields(2)
                    iSubj = iSubj + 1
                    
                    If rsLabEven.Fields(2) >= 50 Then
                        PDF.PDFTextOut rsLabEven.Fields(2), x + 5.15, y + 1.9 + dConsEven
                        PDF.PDFTextOut "P", x + 6.25, y + 1.9 + dConsEven
                        PDF.PDFTextOut rsLabEven.Fields(3), x + 6.8, y + 1.9 + dConsEven
                    Else
                        flgArrear = True
                    End If
                    dConsEven = dConsEven + 0.6875
                    rsLabEven.MoveNext
                Loop
                rsLabEven.Close
                rsTheoryEven.Close
                
                If Not flgArrear Then
                    PDF.PDFTextOut CStr(iTotal), x + 5.15, y + 9.5
                    PDF.PDFTextOut CStr(Round(iTotal / iSubj, 2)) & "%", x + 1.5, y + 9.5
                    dCumPercentage = dCumPercentage + Round(iTotal / iSubj, 2)
                End If
                
                x = x + 9.875
            End If
        Next
        PDF.PDFTextOut "Extra Curricular Activities:", 0.5, 27.75
        PDF.PDFTextOut "Sports:", 20.25, 27.75
        PDF.PDFTextOut "F.A. Sign", 3, 29.1
        PDF.PDFTextOut "H.o.D. Sign", 13, 29.1
        PDF.PDFTextOut "Pricipal Sign", 23, 29.1
        PDF.PDFTextOut "Cummulative Percentage:", 33.15, 28.7
        
        If Not flgArrear Then
            PDF.PDFTextOut Round(dCumPercentage / 8, 2) & "%", 37.9, 28.6
        End If
        
        PDF.PDFDrawRectangle 6, 27, 14.75, 0.75
        PDF.PDFDrawRectangle 22.65, 27, 17.85, 0.75
        PDF.PDFDrawRectangle 38.5, 28, 2, 0.75
    PDF.PDFEndDoc
End Sub
Private Sub cmdSave_Click()
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    Dim sql As String
    Dim i As Integer
    For i = 0 To iSubCount - 1
        If txtMarks(0).Text = "" Then
            MsgBox "Marks of " & lblSubj(i).Caption & " is not entered"
            Exit Sub
        End If
    Next
    For i = 0 To iSubCount - 1
    If lblSubj(i).Caption = "" Then
    Else
        sql = "insert into studmarks (regno,semno,dept,batch,subjcode,marks,passed) values ('" & cmbRegNo.Text & "','" & iSem & "','" & iDept & "','" & Mid(iBatch, 3, 2) & "','" & lblSubj(i).Caption & "','" & txtMarks(i).Text & "','" & txtPassed(i).Text & "')"
        rs.Open sql, conn, adOpenDynamic, adLockOptimistic, -1
    End If
    Next
    If Err.Number <> 0 Then
        MsgBox Error & vbCrLf & "Error Number: " & Err.Number
    Else
        Clear
        HideControls
    End If
    If Err.Number <> 0 Then
        MsgBox Error & vbCrLf & "Error Number: " & Err.Number
    End If
End Sub

Private Sub cmdSettings_Click()
    frmSettings.Show modal, frmMain
End Sub

Private Sub cmdStudSave_Click()
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    Dim qr As String
    rs.CursorLocation = adUseClient
    qr = "insert into studdetails(regno,studname,sec,dob,gender,caste,parent,occupation,address,city,pincode,state,landline,mobile,image,doj,income) values ('" & txtStudRegNo.Text & "','" & txtStudName.Text & "','" & cmbStudSec.Text & "',to_date('" & txtStudDoB.Text & "','DD-MM-YYYY'),'" & cmbStudGender.Text & "','" & cmbStudCaste.Text & "','" & txtStudParent.Text & "','" & txtStudOccupation.Text & "','" & txtStudAddress.Text & "','" & txtStudCity.Text & "','" & txtStudPinCode.Text & "','" & txtStudState.Text & "','" & txtStudPhone.Text & "','" & txtStudMobile.Text & "','" & sImageName & "',to_date('" & txtStudDoJ.Text & "','DD-MM-YYYY'),'" & txtStudIncome.Text & "')"
    MsgBox qr
    rs.Open qr, conn, adOpenDynamic, adLockOptimistic
    If Err.Number <> 0 Then
        MsgBox Error & vbCrLf & "Error Number: " & Err.Number
    Else
        MsgBox ("Student Created")
        txtStudName.Text = ""
        txtStudRegNo.Text = ""
        txtStudMobile.Text = ""
        txtStudPhone.Text = ""
        txtStudParent.Text = ""
        txtStudIncome.Text = ""
        txtStudOccupation.Text = ""
        txtStudAddress.Text = ""
        txtStudCity.Text = ""
        txtStudPinCode.Text = ""
        txtStudState.Text = ""
        txtStudDoB.Text = ""
        txtStudDoJ.Text = ""
        cmbStudCaste.Text = ""
        cmbStudGender.Text = ""
        cmbStudSec.Text = ""
    End If
End Sub

Private Sub cmdSubjInsert_Click()
    On Error Resume Next
    Dim rs As New ADODB.Recordset
    Dim qr As String
    Dim strSubjCode As String
    Dim strSubjName As String
    Dim iIsLab As Integer
    If txtSubjCode.Text = "" Or txtSubjName.Text = "" Then
        MsgBox "Enter Some Data"
        Exit Sub
    End If
    
    If cbIsLab.Value = vbChecked Then
        iIsLab = 1
    Else
        iIsLab = 0
    End If
    
    strSubjCode = Trim(txtSubjCode.Text)
    strSubjName = Trim(txtSubjName.Text)
    rs.CursorLocation = adUseClient
    qr = "insert into subj (subjcode,subjname,semno,dept,batch,lab) values('" & strSubjCode & "','" & strSubjName & "'," & iSem & "," & iDept & "," & Mid(iBatch, 3, 2) & "," & iIsLab & ")"
    rs.Open qr, conn, adOpenDynamic, adLockOptimistic, 1
    If Err.Number <> 0 Then
        MsgBox Error & vbCrLf & "Error Number: " & Err.Number
    ElseIf Err.Number = -2147217873 Then
        MsgBox "Subject Code all ready exist in the database for the same dept"
    Else
        MsgBox "Inserted"
        txtSubjCode.Text = ""
        txtSubjName.Text = ""
        cbIsLab.Value = vbUnchecked
        txtSubjCode.SetFocus
    End If
End Sub




Private Sub dignaj_MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
    OpenWebsite "http://search.dignaj.com/"
End Sub

Private Sub dignaj_MouseHover()
    dignaj.ForeColor = vbRed
End Sub


Private Sub dignaj_MouseLeave()
    dignaj.ForeColor = &HFF0000
End Sub

Private Sub Form_Load()

    tabMain.TabCaption(0) = "     Home     "
    tabMain.TabCaption(1) = "     Student     "
    tabMain.TabCaption(2) = "     Subject     "
    tabMain.TabCaption(3) = "     Marks     "
    tabMain.TabCaption(4) = "     Report     "
    
    frameMarks.Width = tabMain.Width + 10
    frameMarks.Height = tabMain.Height - 590
    
    frameReport.Width = tabMain.Width + 10
    frameReport.Height = tabMain.Height - 590
    
    frameSubj.Width = tabMain.Width + 10
    frameSubj.Height = tabMain.Height - 590
    
    frameContact.Width = tabMain.Width + 10
    frameContact.Height = tabMain.Height - 590
    
    frameIntro.Width = tabMain.Width + 10
    frameIntro.Height = tabMain.Height - 590
    
    tabMain.ActiveTab = 0
    tabMain.HoverColor = &HFF&
        
    
    cmbDept_Load cmbDept
    cmbBatch_Load cmbBatch
    cmbRegNo_Load cmbRegNo
    cmbSem_Load cmbSem
    
    cmbDept_Load cmbRptDept
    cmbBatch_Load cmbRptBatch
    cmbRegNo_Load cmbRptRegNo
    
    cmbDept_Load cmbSubjDept
    cmbSem_Load cmbSubjSem
    cmbBatch_Load cmbSubjBatch
    
    cmbPassed.AddItem "JAN 2008"
    cmbPassed.AddItem "MAY 2008"
    cmbPassed.AddItem "DEC 2008"
    cmbPassed.AddItem "MAY 2009"
    cmbPassed.AddItem "JAN 2009"
    cmbPassed.AddItem "AUG 2010"
    cmbPassed.AddItem "JAN 2011"
    cmbPassed.AddItem "MAY 2011"
    
    
    cmbStudCaste.AddItem "FC"
    cmbStudCaste.AddItem "BC"
    cmbStudCaste.AddItem "MBC"
    cmbStudCaste.AddItem "OBC"
    cmbStudCaste.AddItem "SC"
    
    cmbStudGender.AddItem "Male"
    cmbStudGender.AddItem "Female"
    
    cmbStudSec.AddItem "A"
    cmbStudSec.AddItem "B"
    cmbStudSec.AddItem "C"
    cmbStudSec.AddItem "D"
        
    frameIntro.Visible = True
    frameContact.Visible = False
    frameSubj.Visible = False
    frameMarks.Visible = False
    frameReport.Visible = False
End Sub



Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    conn.Close
    End
End Sub
Private Sub tabMain_BeforeTabSwitch(ByVal iNewActiveTab As Integer, bCancel As Boolean)
    Select Case iNewActiveTab
        Case 0
            frameIntro.Visible = True
            frameContact.Visible = False
            frameSubj.Visible = False
            frameMarks.Visible = False
            frameReport.Visible = False
        Case 1
            frameIntro.Visible = False
            frameContact.Visible = True
            frameSubj.Visible = False
            frameMarks.Visible = False
            frameReport.Visible = False
            txtStudRegNo.SetFocus
        Case 2
            frameIntro.Visible = False
            frameContact.Visible = False
            frameSubj.Visible = True
            frameMarks.Visible = False
            frameReport.Visible = False
            txtSubjCode.SetFocus
        Case 3
            frameIntro.Visible = False
            frameContact.Visible = False
            frameSubj.Visible = False
            frameMarks.Visible = True
            frameReport.Visible = False
            cmbDept.SetFocus
        Case 4
            frameIntro.Visible = False
            frameContact.Visible = False
            frameSubj.Visible = False
            frameMarks.Visible = False
            frameReport.Visible = True
            cmbRptDept.SetFocus
    End Select
End Sub
Private Sub txtPassed_GotFocus(Index As Integer)
    cmbPassed.Visible = True
    cmbPassed.ZOrder
    cmbPassed.SetFocus
    cmbPassed.Left = txtPassed(Index).Left
    cmbPassed.Top = txtPassed(Index).Top
    itxtPassedIndex = Index
End Sub
Private Sub Clear()
    Dim i As Integer
    For i = 0 To 9
        lblSubj(i).Caption = ""
        txtMarks(i).Text = ""
        txtPassed(i).Text = ""
    Next
End Sub


