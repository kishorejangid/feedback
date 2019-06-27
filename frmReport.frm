VERSION 5.00
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkusercontrolsxp.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmReport 
   BackColor       =   &H80000010&
   Caption         =   "Feedback - Reports"
   ClientHeight    =   8250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin feedback.ThemedComboBox ThemedComboBox1 
      Left            =   0
      Top             =   0
      _ExtentX        =   556
      _ExtentY        =   529
   End
   Begin vkUserContolsXP.vkFrame frameReport 
      Height          =   7935
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   13996
      Caption         =   "Report"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowTitle       =   0   'False
      TitleColor1     =   4210752
      TitleColor2     =   14737632
      TitleGradient   =   2
      TitleHeight     =   300
      BorderColor     =   4210752
      BorderWidth     =   2
      Begin feedback.StylerButton cmdSettings 
         Height          =   375
         Left            =   4920
         TabIndex        =   38
         Top             =   0
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
      Begin feedback.StylerButton cmdLoad 
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   3360
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         Caption         =   "Load"
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
      Begin feedback.StylerButton cmdEx 
         Height          =   255
         Left            =   1320
         TabIndex        =   32
         Top             =   3360
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         Caption         =   "Expand All"
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
      Begin feedback.StylerButton cmdTree 
         Height          =   375
         Left            =   6480
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "Tree View"
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
      Begin feedback.StylerButton cmdSubjRpt 
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   0
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "Subject Report"
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
      Begin feedback.StylerButton cmdClass 
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   0
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "Class Report"
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
      Begin feedback.StylerButton cmdStaff 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   0
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "Staff Report"
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
      Begin vkUserContolsXP.vkFrame frameControls 
         Height          =   2775
         Left            =   240
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   600
         Visible         =   0   'False
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   4895
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
         Begin vkUserContolsXP.vkFrame fSem 
            Height          =   855
            Left            =   120
            TabIndex        =   40
            Top             =   1800
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   1508
            Caption         =   "Semester"
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
            TitleColor2     =   12632256
            TitleGradient   =   2
            TitleHeight     =   300
            BorderColor     =   12632256
            Begin vkUserContolsXP.vkOptionButton rbEven 
               Height          =   255
               Left            =   1320
               TabIndex        =   42
               Top             =   480
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   450
               BackColor       =   16777215
               BackStyle       =   0
               Caption         =   "Even Sem"
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
            Begin vkUserContolsXP.vkOptionButton rbOdd 
               Height          =   255
               Left            =   120
               TabIndex        =   41
               Top             =   480
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   450
               BackColor       =   16777215
               BackStyle       =   0
               Caption         =   "Odd Sem"
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
         End
         Begin feedback.StylerButton cmdFolder 
            Height          =   375
            Left            =   8400
            TabIndex        =   39
            Top             =   1800
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            Caption         =   "Open Report Folder"
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
         Begin feedback.StylerButton cmdSub 
            Height          =   375
            Left            =   6840
            TabIndex        =   37
            Top             =   840
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            Caption         =   "Detailed PDF"
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
         Begin vkUserContolsXP.vkCheck cbStaff 
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   840
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Generate PDF for Individual Subject"
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
         Begin feedback.StylerButton cmdDetailedRpt 
            Height          =   375
            Left            =   8400
            TabIndex        =   35
            Top             =   1200
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            Caption         =   "Grouped PDF"
            CaptionDisableColor=   12236471
            CaptionEffectColor=   16777215
            FocusDottedRect =   0   'False
            Enabled         =   0   'False
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
         Begin feedback.StylerButton cmdPrint 
            Height          =   375
            Left            =   8400
            TabIndex        =   33
            Top             =   720
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            Caption         =   "Create PDF"
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
         Begin vkUserContolsXP.vkLabel lblStaffName 
            Height          =   255
            Left            =   4320
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   1560
            Visible         =   0   'False
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   450
            BorderColor     =   4210752
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
         Begin vkUserContolsXP.vkLabel lblStaff 
            Height          =   255
            Left            =   3240
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   1560
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Staff  Name:"
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
         Begin vkUserContolsXP.vkLabel lblSubjName 
            Height          =   255
            Left            =   4320
            TabIndex        =   29
            Top             =   1200
            Width           =   3255
            _ExtentX        =   5741
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
         Begin vkUserContolsXP.vkLabel lblSubj 
            Height          =   255
            Left            =   3240
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   960
            Visible         =   0   'False
            Width           =   615
            _ExtentX        =   1085
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
         Begin VB.ComboBox cmbSubj 
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
            Left            =   4320
            TabIndex        =   11
            Top             =   840
            Visible         =   0   'False
            Width           =   2295
         End
         Begin feedback.StylerButton cmdClassGo 
            Height          =   375
            Left            =   8400
            TabIndex        =   12
            Top             =   240
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            Caption         =   "Go"
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
         Begin VB.ComboBox cmbStaff 
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
            Left            =   4320
            TabIndex        =   7
            Top             =   240
            Visible         =   0   'False
            Width           =   3585
         End
         Begin vkUserContolsXP.vkLabel lblRptStaff 
            Height          =   255
            Left            =   3240
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   360
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Select Staff:"
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
            Height          =   255
            Left            =   120
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   360
            Visible         =   0   'False
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
            Left            =   1200
            TabIndex        =   6
            Top             =   240
            Visible         =   0   'False
            Width           =   1815
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
            Left            =   1200
            TabIndex        =   8
            Top             =   840
            Visible         =   0   'False
            Width           =   1815
         End
         Begin vkUserContolsXP.vkLabel lblRptBatch 
            Height          =   255
            Left            =   120
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   960
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
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
         Begin vkUserContolsXP.vkLabel lblRptSemester 
            Height          =   255
            Left            =   120
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   1560
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
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
            Left            =   1200
            TabIndex        =   10
            Top             =   2040
            Visible         =   0   'False
            Width           =   1815
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
            Left            =   1200
            TabIndex        =   9
            Top             =   1440
            Visible         =   0   'False
            Width           =   1815
         End
         Begin vkUserContolsXP.vkLabel lblRptSec 
            Height          =   255
            Left            =   120
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   2160
            Visible         =   0   'False
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
         Begin vkUserContolsXP.vkLabel lblStaffID 
            Height          =   255
            Left            =   7320
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   600
            Visible         =   0   'False
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Staff ID"
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
         Begin vkUserContolsXP.vkLabel lblStu 
            Height          =   195
            Left            =   8160
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   2280
            Visible         =   0   'False
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   344
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Student Count:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin MSComctlLib.TreeView tvFeed 
         Height          =   1110
         Left            =   240
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   3600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1958
         _Version        =   393217
         Style           =   7
         Appearance      =   1
      End
      Begin vkUserContolsXP.vkFrame frameGroup 
         Height          =   2775
         Left            =   10440
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   600
         Visible         =   0   'False
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   4895
         Caption         =   "Group Representation"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   4210752
         TitleColor1     =   8421504
         TitleColor2     =   14737632
         TitleGradient   =   2
         TitleHeight     =   360
         BorderColor     =   4210752
         BorderWidth     =   2
         Begin vkUserContolsXP.vkLabel lblG1 
            Height          =   255
            Left            =   240
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   480
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Group 1:          Planning and Organisation."
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
         Begin vkUserContolsXP.vkLabel lblG2 
            Height          =   255
            Left            =   240
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   960
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Group 2:          Presentation / Communication."
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
         Begin vkUserContolsXP.vkLabel lblG3 
            Height          =   255
            Left            =   240
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   1440
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Group 3:          Student's Participation."
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
         Begin vkUserContolsXP.vkLabel lblG4 
            Height          =   255
            Left            =   240
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   1920
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Group 4:          Class Management."
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
         Begin vkUserContolsXP.vkLabel lblG41 
            Height          =   255
            Left            =   1320
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   2160
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   450
            BackColor       =   16777215
            BackStyle       =   0
            Caption         =   "Assessment of Students."
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshGrid 
         Height          =   3975
         Left            =   240
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   3600
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   7011
         _Version        =   393216
         FixedCols       =   0
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tvExpanded As Boolean
Dim FSubject As Boolean
Dim strPointer As String

Private Sub cbStaff_Change(Value As CheckBoxConstants)
    LoadSubjDetailed
End Sub

Private Sub cbStaff_Click()
    LoadSubjDetailed
End Sub
Private Sub LoadSubjDetailed()
    If cmbStaff.Text = "" Then
        cbStaff.Value = vbUnchecked
        Exit Sub
    End If
    If cbStaff.Value = vbChecked Then
        Dim rs As New ADODB.Recordset
        rs.Open "select subjcode,dept,batch,sem,sec from staffhandle where staffid='" & lblStaffID.Caption & "'", conn, adOpenDynamic, adLockOptimistic, -1
        cmbSubj.Clear
        rs.MoveFirst
        While Not rs.EOF
            Dim rsDeptShort As New ADODB.Recordset
            rsDeptShort.Open "select deptshort from dept where deptcode= '" & rs.Fields(1) & "'", conn, adOpenDynamic, adLockOptimistic
            cmbSubj.AddItem rs.Fields(0) & "-" & rsDeptShort.Fields(0) & "-" & rs.Fields(2) & "-" & rs.Fields(3) & "-" & rs.Fields(4)
            rs.MoveNext
            rsDeptShort.Close
        Wend
        cmbSubj.ListIndex = 0
        lblSubj.Visible = True
        cmbSubj.Visible = True
        cmdSub.Visible = True
    ElseIf cbStaff.Value = vbUnchecked Then
        lblSubj.Visible = False
        cmbSubj.Visible = False
        cmdSub.Visible = False
    End If
End Sub

Private Sub cmbDept_Change()
    On Error Resume Next
    mshGrid.Clear
    iDept = Department(cmbDept)
    cmbStaff_Load cmbStaff
    lblStaffName.Caption = ""
    lblSubjName.Caption = ""
    cmbSubj_Load cmbSubj
End Sub
Private Sub cmbDept_Click()
    On Error Resume Next
    mshGrid.Clear
    iDept = Department(cmbDept)
    cmbStaff_Load cmbStaff
    lblStaffName.Caption = ""
    lblSubjName.Caption = ""
    cmbSubj_Load cmbSubj
End Sub
Private Sub cmbBatch_Change()
    On Error Resume Next
    iBatch = cmbBatch.Text
    lblStaffName.Caption = ""
    lblSubjName.Caption = ""
    cmbSubj_Load cmbSubj
End Sub
Private Sub cmbBatch_Click()
    On Error Resume Next
    iBatch = cmbBatch.Text
    lblStaffName.Caption = ""
    lblSubjName.Caption = ""
    cmbSubj_Load cmbSubj
End Sub
Private Sub cmbSec_Change()
    On Error Resume Next
    strSec = cmbSec.Text
    lblStaffName.Caption = ""
    lblSubjName.Caption = ""
    cmbSubj_Load cmbSubj
End Sub
Private Sub cmbSec_Click()
    On Error Resume Next
    strSec = cmbSec.Text
    lblStaffName.Caption = ""
    lblSubjName.Caption = ""
    cmbSubj_Load cmbSubj
End Sub

Private Sub cmbSem_Change()
    On Error Resume Next
    iSem = Val(cmbSem.Text)
    lblStaffName.Caption = ""
    lblSubjName.Caption = ""
    cmbSubj_Load cmbSubj
End Sub

Private Sub cmbSem_Click()
    On Error Resume Next
    iSem = Val(cmbSem.Text)
    lblStaffName.Caption = ""
    lblSubjName.Caption = ""
    cmbSubj_Load cmbSubj
End Sub

Private Sub cmbStaff_Change()
    On Error Resume Next
    Dim rsStaffID As New ADODB.Recordset
    Dim sqlStaffID As String
    sqlStaffID = "select staffid from staff where staffname='" & UCase(cmbStaff.Text) & "'"
    rsStaffID.Open sqlStaffID, conn, adOpenDynamic, adLockOptimistic, -1
    lblStaffID.Caption = rsStaffID.Fields(0)
    lblSubj.Visible = False
    cmbSubj.Visible = False
    cmdSub.Visible = False
    cbStaff.Value = vbUnchecked
    LoadStaff
    If Err.Number <> 0 Then
        MsgBox "Error Number : " & Err.Number & vbCrLf & "Error Description : " & Err.Description & vbCrLf & "Error Location : frmReport-cmbStaff_Change()", vbInformation, "Jangid Corporation"
    End If
End Sub
Private Sub cmbStaff_Click()
    On Error Resume Next
    Dim rsStaffID As New ADODB.Recordset
    Dim sqlStaffID As String
    sqlStaffID = "select staffid from staff where staffname='" & UCase(cmbStaff.Text) & "'"
    rsStaffID.Open sqlStaffID, conn, adOpenDynamic, adLockOptimistic, -1
    lblStaffID.Caption = rsStaffID.Fields(0)
    
    lblSubj.Visible = False
    cmbSubj.Visible = False
    cmdSub.Visible = False
    cbStaff.Value = vbUnchecked
    
    LoadStaff
    If Err.Number <> 0 Then
        MsgBox "Error Number : " & Err.Number & vbCrLf & "Error Description : " & Err.Description & vbCrLf & "Error Location : frmReport-cmbStaff_Click()", vbInformation, "Jangid Corporation"
    End If
End Sub

Private Sub LoadClass()
    On Error Resume Next
    Dim rsClass As New ADODB.Recordset
    Dim strSql As String
    strSql = "select m.subjcode,s.subjname,t.staffname, round(avg(m.q1+m.q2+m.q3+m.q4+m.q5)/5,2) as g1,round(avg(m.q6+m.q7+m.q8+m.q9+m.q10)/5,2) as g2,round(avg(m.q11+m.q12+m.q13+m.q14+m.q15)/5,2) as g3,round(avg(m.q16+m.q17+m.q18+m.q19+m.q20)/5,2) as g4,round(avg(m.q1+m.q2+m.q3+m.q4+m.q5+m.q6+m.q7+m.q8+m.q9+m.q10+m.q11+m.q12+m.q13+m.q14+m.q15+m.q16+m.q17+m.q18+m.q19+m.q20)/20,2) as op,count(m.fid) as Count from master m,subj s,staff t where s.subjcode=m.subjcode and s.dept=m.dept and t.staffid=m.staffid and m.dept='" & iDept & "' and m.sec='" & strSec & "' and m.sem='" & iSem & "' and m.batch='" & Mid(iBatch, 3, 2) & "' group by m.sec,m.subjcode,s.subjname,t.staffname order by m.subjcode"
    rsClass.Open strSql, conn, adOpenDynamic, adLockOptimistic, -1
    Set mshGrid.DataSource = rsClass
    
    tvFeed.Visible = False
    mshGrid.Visible = True
    mshGrid.ColWidth(0) = 1100
    mshGrid.ColWidth(1) = 4000
    mshGrid.ColWidth(2) = 2700
    mshGrid.ColWidth(3) = 750
    mshGrid.ColWidth(4) = 750
    mshGrid.ColWidth(5) = 750
    mshGrid.ColWidth(6) = 750
    mshGrid.ColWidth(7) = 1700
    mshGrid.ColWidth(8) = 600
    
    mshGrid.ColAlignment(0) = flexAlignCenterCenter
    mshGrid.ColAlignment(3) = flexAlignCenterCenter
    mshGrid.ColAlignment(4) = flexAlignCenterCenter
    mshGrid.ColAlignment(5) = flexAlignCenterCenter
    mshGrid.ColAlignment(6) = flexAlignCenterCenter
    mshGrid.ColAlignment(7) = flexAlignCenterCenter
    mshGrid.ColAlignment(8) = flexAlignCenterCenter
    mshGrid.ColAlignmentFixed(0) = flexAlignCenterCenter
    mshGrid.ColAlignmentFixed(1) = flexAlignCenterCenter
    mshGrid.ColAlignmentFixed(2) = flexAlignCenterCenter
    mshGrid.ColAlignmentFixed(3) = flexAlignCenterCenter
    mshGrid.ColAlignmentFixed(4) = flexAlignCenterCenter
    mshGrid.ColAlignmentFixed(5) = flexAlignCenterCenter
    mshGrid.ColAlignmentFixed(6) = flexAlignCenterCenter
    mshGrid.ColAlignmentFixed(7) = flexAlignCenterCenter
    mshGrid.ColAlignmentFixed(8) = flexAlignCenterCenter
    mshGrid.RowHeightMin = 350
    
    
    mshGrid.TextMatrix(0, 0) = "Subject Code"
    mshGrid.TextMatrix(0, 1) = "Subject Name"
    mshGrid.TextMatrix(0, 2) = "Staff Name"
    mshGrid.TextMatrix(0, 3) = "Group 1"
    mshGrid.TextMatrix(0, 4) = "Group 2"
    mshGrid.TextMatrix(0, 5) = "Group 3"
    mshGrid.TextMatrix(0, 6) = "Group 4"
    mshGrid.TextMatrix(0, 7) = "Overall Performance"
    mshGrid.TextMatrix(0, 8) = "Count"
End Sub
Private Sub cmbSubj_Change()
    On Error Resume Next
    Dim rsSubjName As New ADODB.Recordset
    rsSubjName.CursorLocation = adUseClient
    Dim strSql As String
    strSql = "select s1.subjname,s3.staffname from subj s1,staffhandle s2,staff s3 where s1.subjcode='" & cmbSubj.Text & "' and s1.subjcode=s2.subjcode and s2.staffid=s3.staffid and s2.sec='" & cmbSec.Text & "'"
    rsSubjName.Open strSql, conn, adOpenDynamic, adLockOptimistic, -1
    lblSubjName.Caption = rsSubjName.Fields(0)
    lblStaffName.Caption = rsSubjName.Fields(1)
End Sub
Private Sub cmbSubj_Click()
    On Error Resume Next
    Dim rsSubjName As New ADODB.Recordset
    rsSubjName.CursorLocation = adUseClient
    Dim strSql As String
    strSql = "select s1.subjname,s3.staffname from subj s1,staffhandle s2,staff s3 where s1.subjcode='" & cmbSubj.Text & "' and s1.subjcode=s2.subjcode and s2.staffid=s3.staffid and s2.sec='" & cmbSec.Text & "'"
    rsSubjName.Open strSql, conn, adOpenDynamic, adLockOptimistic, -1
    lblSubjName.Caption = rsSubjName.Fields(0)
    lblStaffName.Caption = rsSubjName.Fields(1)
End Sub

Private Sub cmdClass_Click()
    On Error Resume Next
    strPointer = "Class"
    FSubject = False
    frameControls.Visible = True
    frameGroup.Visible = True
    fSem.Visible = False
    
    cmdStaff.ForeColor = vbBlack
    cmdStaff.CaptionEffect = Normal
    cmdClass.ForeColor = vbRed
    cmdClass.CaptionEffect = outline
    cmdSubjRpt.ForeColor = vbBlack
    cmdSubjRpt.CaptionEffect = Normal
    cmdTree.ForeColor = vbBlack
    cmdTree.CaptionEffect = Normal
    cmdPrint.Visible = True
    
    lblRptDept.Visible = True
    lblRptSec.Visible = True
    lblRptBatch.Visible = True
    lblRptSemester.Visible = True
    
    cmbSec.Visible = True
    cmbSem.Visible = True
    cmbBatch.Visible = True
    cmbDept.Visible = True
    
    cmdClassGo.Visible = True
    
    lblRptStaff.Visible = False
    cmbStaff.Visible = False
    lblStu.Visible = False
    lblSubj.Visible = False
    lblStaff.Visible = False
    lblStaffName.Visible = False
    cmbSubj.Visible = False
    
    tvFeed.Top = mshGrid.Top
    tvFeed.Height = mshGrid.Height
    cmdEx.Visible = False
    cmdLoad.Visible = False
    tvFeed.Visible = False
    lblSubjName.Visible = False
    cmdDetailedRpt.Visible = False
    cbStaff.Visible = False
    cmdSub.Visible = False
    mshGrid.Visible = True
    mshGrid.ClearStructure
    cmbDept.SetFocus
End Sub

Private Sub cmdClassGo_Click()
    On Error Resume Next
    If FSubject = True Then
        LoadSubject
    ElseIf FSubject = False Then
        LoadClass
    End If
End Sub
Private Sub rptSubject()
    On Error Resume Next
    Dim strLeft As Double
    Dim PDF As New clsPDF
    Dim j As Integer
    PDF.PDFTitle = "Subject wise Report"
    PDF.PDFFileName = App.Path & "\Reports\" & cmbDept.Text & "_" & cmbBatch.Text & "_" & cmbSem.Text & "_" & cmbSec.Text & "_" & cmbSubj.Text & "_" & "Detailed" & ".pdf"
    PDF.PDFAuthor = "Jangid Corporation"
    PDF.PDFLoadAfm = App.Path
    PDF.PDFSetAdobePath = strAdobePath
    PDF.PDFView = True
    
    PDF.PDFSetUnit = UNIT_CM
    PDF.PDFFormatPage = FORMAT_A4
    PDF.PDFOrientation = ORIENT_PAYSAGE
    
    PDF.PDFSetBorder = BORDER_ALL
    PDF.PDFSetTopMargin = 1
    PDF.PDFSetBottomMargin = 1
    PDF.PDFSetRightMargin = 1
    PDF.PDFSetLeftMargin = 1
    PDF.PDFBeginDoc
        
        'Draw Border
        PDF.PDFDrawRectangle 0.5, 0.5, 28.7, 20
        'Header Part
        PDF.PDFSetTextColor = vbBlack
        PDF.PDFSetFont FONT_TIMES, 18, FONT_BOLD
        l = PDF.PDFGetStringWidth(UCase(strCollegeName), "Times-Bold", 18)
        strLeft = (28.7 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut UCase(strCollegeName), strLeft, 1.5
        
        PDF.PDFSetFont FONT_TIMES, 14, FONT_BOLD
        l = PDF.PDFGetStringWidth("STAFF EVALUATION BY STUDENTS", "Times-Bold", 14)
        strLeft = (28.7 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut "STAFF EVALUATION BY STUDENTS", strLeft, 2.1
        l = PDF.PDFGetStringWidth("DEPARTMENT OF " & UCase(cmbDept.Text), "Times-Bold", 14)
        strLeft = (28.7 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut "DEPARTMENT OF " & UCase(cmbDept.Text), strLeft, 2.75
        PDF.PDFSetFont FONT_TIMES, 12, FONT_BOLD
        PDF.PDFTextOut "Date:", 24.9, 3.75
        PDF.PDFTextOut Format(DateTime.Date, "dd/mm/yyyy"), 26, 3.75
        
        'Private for this module
        l = PDF.PDFGetStringWidth("SUBJECT REPORT - DETAILED", "Times-Bold", 14)
        strLeft = (28.7 - (l * 2.54) / 72) / 2
        PDF.PDFSetFont FONT_TIMES, 14, FONT_BOLD
        PDF.PDFTextOut "SUBJECT REPORT - DETAILED", strLeft, 3.75
        
        PDF.PDFSetFont 2, 12, FONT_BOLD
        PDF.PDFTextOut "SUBJECT:", -0.3, 4.5
        PDF.PDFTextOut cmbSubj.Text & " - " & lblSubjName.Caption, 2.15, 4.5
        PDF.PDFTextOut "STAFF:", 15, 4.5
        PDF.PDFTextOut lblStaffName.Caption, 16.75, 4.5
        
        
        'Divider for header
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLineHor 0.5, 4, 28.7
        PDF.PDFSetLineWidth = 0.02
        PDF.PDFDrawLineHor 0.5, 4.75, 28.7
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLineHor 0.5, 5.5, 28.7
                                                                
        PDF.PDFTextOut "FID", -0.3, 5.25
        PDF.PDFTextOut "1", 0.95, 5.25
        PDF.PDFTextOut "2", 1.85, 5.25
        PDF.PDFTextOut "3", 2.75, 5.25
        PDF.PDFTextOut "4", 3.65, 5.25
        PDF.PDFTextOut "5", 4.55, 5.25
        PDF.PDFTextOut "6", 5.45, 5.25
        PDF.PDFTextOut "7", 6.35, 5.25
        PDF.PDFTextOut "8", 7.25, 5.25
        PDF.PDFTextOut "9", 8.15, 5.25
        PDF.PDFTextOut "10", 8.95, 5.25
        PDF.PDFTextOut "11", 9.85, 5.25
        PDF.PDFTextOut "12", 10.75, 5.25
        PDF.PDFTextOut "13", 11.65, 5.25
        PDF.PDFTextOut "14", 12.55, 5.25
        PDF.PDFTextOut "15", 13.45, 5.25
        PDF.PDFTextOut "16", 14.35, 5.25
        PDF.PDFTextOut "17", 15.25, 5.25
        PDF.PDFTextOut "18", 16.15, 5.25
        PDF.PDFTextOut "19", 17.05, 5.25
        PDF.PDFTextOut "20", 17.95, 5.25
        PDF.PDFSetFont 2, 9, FONT_BOLD
        PDF.PDFTextOut "GROUP 1", 18.85, 5.25
        PDF.PDFTextOut "GROUP 2", 20.75, 5.25
        PDF.PDFTextOut "GROUP 3", 22.65, 5.25
        PDF.PDFTextOut "GROUP 4", 24.55, 5.25
        PDF.PDFTextOut "OVERALL", 26.4, 5.25

        
        Dim iTotalRows As Integer
        Dim iRowStart As Integer
        Dim iRowEnd As Integer
        Dim iNoOfExtraPages As Integer
        Dim fMultiPage As Boolean
        
        fMultiPage = False
        iTotalRows = mshGrid.Rows - 1
        
        If iTotalRows >= 24 Then
            iRowStart = 1
            iRowEnd = 24
            fMultiPage = True
            iNoOfExtraPages = RoundUp((iTotalRows - 24) / 30)
        Else
            iRowStart = 1
            iRowEnd = iTotalRows
            fMultiPage = False
        End If
        
        Dim dVH As Double
        dVH = (iRowEnd + 1) * 0.6275
        PDF.PDFDrawLineVer 1.6, 4.75, dVH
        PDF.PDFDrawLineVer 2.5, 4.75, dVH
        PDF.PDFDrawLineVer 3.4, 4.75, dVH
        PDF.PDFDrawLineVer 4.3, 4.75, dVH
        PDF.PDFDrawLineVer 5.2, 4.75, dVH
        PDF.PDFDrawLineVer 6.1, 4.75, dVH
        PDF.PDFDrawLineVer 7, 4.75, dVH
        PDF.PDFDrawLineVer 7.9, 4.75, dVH
        PDF.PDFDrawLineVer 8.8, 4.75, dVH
        PDF.PDFDrawLineVer 9.7, 4.75, dVH
        PDF.PDFDrawLineVer 10.6, 4.75, dVH
        PDF.PDFDrawLineVer 11.5, 4.75, dVH
        PDF.PDFDrawLineVer 12.4, 4.75, dVH
        PDF.PDFDrawLineVer 13.3, 4.75, dVH
        PDF.PDFDrawLineVer 14.2, 4.75, dVH
        PDF.PDFDrawLineVer 15.1, 4.75, dVH
        PDF.PDFDrawLineVer 16, 4.75, dVH
        PDF.PDFDrawLineVer 16.9, 4.75, dVH
        PDF.PDFDrawLineVer 17.8, 4.75, dVH
        PDF.PDFDrawLineVer 18.7, 4.75, dVH
        PDF.PDFDrawLineVer 19.6, 4.75, dVH
        PDF.PDFDrawLineVer 21.5, 4.75, dVH
        PDF.PDFDrawLineVer 23.4, 4.75, dVH
        PDF.PDFDrawLineVer 25.3, 4.75, dVH
        PDF.PDFDrawLineVer 27.2, 4.75, dVH
        PDF.PDFDrawLineHor 0.5, 4.75 + dVH, 28.7
        
        PDF.PDFSetFont 2, 10, FONT_NORMAL
        For i = iRowStart To iRowEnd
            PDF.PDFTextOut mshGrid.TextMatrix(i, 0), -0.01, 5.6 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 1), 0.95, 5.6 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 2), 1.85, 5.6 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 3), 2.75, 5.6 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 4), 3.65, 5.6 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 5), 4.55, 5.6 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 6), 5.45, 5.6 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 7), 6.35, 5.6 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 8), 7.25, 5.6 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 9), 8.15, 5.6 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 10), 9.05, 5.6 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 11), 9.95, 5.6 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 12), 10.85, 5.6 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 13), 11.75, 5.6 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 14), 12.65, 5.6 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 15), 13.55, 5.6 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 16), 14.45, 5.6 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 17), 15.35, 5.6 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 18), 16.25, 5.6 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 19), 17.15, 5.6 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 20), 18.05, 5.6 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 21), 19.35, 5.6 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 22), 21.25, 5.6 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 23), 23.15, 5.6 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 24), 25.05, 5.6 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 25), 26.95, 5.6 + i * 0.6
        Next
        'PDF.PDFSetFont FONT_TIMES, 12, FONT_BOLD
        'PDF.PDFTextOut "GROUP REPRESENTATION:", 0.25, 17
        'PDF.PDFSetFont 2, 10, FONT_NORMAL
        'PDF.PDFTextOut "GROUP 1:    Planning & Organisation", 1.25, 17.8
        'PDF.PDFTextOut "GROUP 2:    Presentation / Communication", 1.25, 18.4
        'PDF.PDFTextOut "GROUP 3:    Student's Participation", 1.25, 19
        'PDF.PDFTextOut "GROUP 4:    Class Management / Assessment Of Students", 1.25, 19.6
        
        'Second Page
    If fMultiPage Then
        For q = 1 To iNoOfExtraPages
        PDF.PDFNewPage
            PDF.PDFDrawRectangle 0.5, 0.5, 28.7, 20
            PDF.PDFSetLineWidth = 0.02
            PDF.PDFDrawLineHor 0.5, 1.25, 28.7
            PDF.PDFSetLineWidth = 0.03
            PDF.PDFDrawLineHor 0.5, 2, 28.7
            PDF.PDFSetTextColor = vbBlack
            
            PDF.PDFSetFont 2, 12, FONT_BOLD
            
            PDF.PDFTextOut "FID", -0.3, 1.75
            PDF.PDFTextOut "1", 0.95, 1.75
            PDF.PDFTextOut "2", 1.85, 1.75
            PDF.PDFTextOut "3", 2.75, 1.75
            PDF.PDFTextOut "4", 3.65, 1.75
            PDF.PDFTextOut "5", 4.55, 1.75
            PDF.PDFTextOut "6", 5.45, 1.75
            PDF.PDFTextOut "7", 6.35, 1.75
            PDF.PDFTextOut "8", 7.25, 1.75
            PDF.PDFTextOut "9", 8.15, 1.75
            PDF.PDFTextOut "10", 8.95, 1.75
            PDF.PDFTextOut "11", 9.85, 1.75
            PDF.PDFTextOut "12", 10.75, 1.75
            PDF.PDFTextOut "13", 11.65, 1.75
            PDF.PDFTextOut "14", 12.55, 1.75
            PDF.PDFTextOut "15", 13.45, 1.75
            PDF.PDFTextOut "16", 14.35, 1.75
            PDF.PDFTextOut "17", 15.25, 1.75
            PDF.PDFTextOut "18", 16.15, 1.75
            PDF.PDFTextOut "19", 17.05, 1.75
            PDF.PDFTextOut "20", 17.95, 1.75
            PDF.PDFSetFont 2, 9, FONT_BOLD
            PDF.PDFTextOut "GROUP 1", 18.85, 1.75
            PDF.PDFTextOut "GROUP 2", 20.75, 1.75
            PDF.PDFTextOut "GROUP 3", 22.65, 1.75
            PDF.PDFTextOut "GROUP 4", 24.55, 1.75
            PDF.PDFTextOut "OVERALL", 26.4, 1.75
            
            PDF.PDFSetFont 2, 12, FONT_BOLD
            PDF.PDFTextOut "SUBJECT:", -0.3, 1
            PDF.PDFTextOut cmbSubj.Text & " - " & lblSubjName.Caption, 2.15, 1
            PDF.PDFTextOut "STAFF:", 15, 1
            PDF.PDFTextOut lblStaffName.Caption, 16.75, 1
            
            iRowStart = iRowEnd + 1
            
            
            
            If (iTotalRows - iRowEnd) > 30 Then
                iRowEnd = iRowEnd + 30
            Else
                iRowEnd = iRowEnd + (iTotalRows - iRowEnd)
            End If
            
            Dim ii As Integer
            ii = 1
            PDF.PDFSetFont 2, 10, FONT_NORMAL
            For i = iRowStart To iRowEnd
                PDF.PDFTextOut mshGrid.TextMatrix(i, 0), -0.01, 2 + ii * 0.6
                PDF.PDFTextOut mshGrid.TextMatrix(i, 1), 0.95, 2 + ii * 0.6
                PDF.PDFTextOut mshGrid.TextMatrix(i, 2), 1.85, 2 + ii * 0.6
                PDF.PDFTextOut mshGrid.TextMatrix(i, 3), 2.75, 2 + ii * 0.6
                PDF.PDFTextOut mshGrid.TextMatrix(i, 4), 3.65, 2 + ii * 0.6
                PDF.PDFTextOut mshGrid.TextMatrix(i, 5), 4.55, 2 + ii * 0.6
                PDF.PDFTextOut mshGrid.TextMatrix(i, 6), 5.45, 2 + ii * 0.6
                PDF.PDFTextOut mshGrid.TextMatrix(i, 7), 6.35, 2 + ii * 0.6
                PDF.PDFTextOut mshGrid.TextMatrix(i, 8), 7.25, 2 + ii * 0.6
                PDF.PDFTextOut mshGrid.TextMatrix(i, 9), 8.15, 2 + ii * 0.6
                PDF.PDFTextOut mshGrid.TextMatrix(i, 10), 9.05, 2 + ii * 0.6
                PDF.PDFTextOut mshGrid.TextMatrix(i, 11), 9.95, 2 + ii * 0.6
                PDF.PDFTextOut mshGrid.TextMatrix(i, 12), 10.85, 2 + ii * 0.6
                PDF.PDFTextOut mshGrid.TextMatrix(i, 13), 11.75, 2 + ii * 0.6
                PDF.PDFTextOut mshGrid.TextMatrix(i, 14), 12.65, 2 + ii * 0.6
                PDF.PDFTextOut mshGrid.TextMatrix(i, 15), 13.55, 2 + ii * 0.6
                PDF.PDFTextOut mshGrid.TextMatrix(i, 16), 14.45, 2 + ii * 0.6
                PDF.PDFTextOut mshGrid.TextMatrix(i, 17), 15.35, 2 + ii * 0.6
                PDF.PDFTextOut mshGrid.TextMatrix(i, 18), 16.25, 2 + ii * 0.6
                PDF.PDFTextOut mshGrid.TextMatrix(i, 19), 17.15, 2 + ii * 0.6
                PDF.PDFTextOut mshGrid.TextMatrix(i, 20), 18.05, 2 + ii * 0.6
                PDF.PDFTextOut mshGrid.TextMatrix(i, 21), 19.35, 2 + ii * 0.6
                PDF.PDFTextOut mshGrid.TextMatrix(i, 22), 21.25, 2 + ii * 0.6
                PDF.PDFTextOut mshGrid.TextMatrix(i, 23), 23.15, 2 + ii * 0.6
                PDF.PDFTextOut mshGrid.TextMatrix(i, 24), 25.05, 2 + ii * 0.6
                PDF.PDFTextOut mshGrid.TextMatrix(i, 25), 26.95, 2 + ii * 0.6
                ii = ii + 1
            Next
            
        dVH = 19.25
        PDF.PDFDrawLineVer 1.6, 1.25, dVH
        PDF.PDFDrawLineVer 2.5, 1.25, dVH
        PDF.PDFDrawLineVer 3.4, 1.25, dVH
        PDF.PDFDrawLineVer 4.3, 1.25, dVH
        PDF.PDFDrawLineVer 5.2, 1.25, dVH
        PDF.PDFDrawLineVer 6.1, 1.25, dVH
        PDF.PDFDrawLineVer 7, 1.25, dVH
        PDF.PDFDrawLineVer 7.9, 1.25, dVH
        PDF.PDFDrawLineVer 8.8, 1.25, dVH
        PDF.PDFDrawLineVer 9.7, 1.25, dVH
        PDF.PDFDrawLineVer 10.6, 1.25, dVH
        PDF.PDFDrawLineVer 11.5, 1.25, dVH
        PDF.PDFDrawLineVer 12.4, 1.25, dVH
        PDF.PDFDrawLineVer 13.3, 1.25, dVH
        PDF.PDFDrawLineVer 14.2, 1.25, dVH
        PDF.PDFDrawLineVer 15.1, 1.25, dVH
        PDF.PDFDrawLineVer 16, 1.25, dVH
        PDF.PDFDrawLineVer 16.9, 1.25, dVH
        PDF.PDFDrawLineVer 17.8, 1.25, dVH
        PDF.PDFDrawLineVer 18.7, 1.25, dVH
        PDF.PDFDrawLineVer 19.6, 1.25, dVH
        PDF.PDFDrawLineVer 21.5, 1.25, dVH
        PDF.PDFDrawLineVer 23.4, 1.25, dVH
        PDF.PDFDrawLineVer 25.3, 1.25, dVH
        PDF.PDFDrawLineVer 27.2, 1.25, dVH
        PDF.PDFDrawLineHor 0.5, 1.25 + dVH, 28.7
        PDF.PDFEndPage
        Next
    End If
    
    PDF.PDFEndDoc
End Sub

Private Sub cmdDetailedRpt_Click()
    On Error Resume Next
    Dim strLeft As Double
    Dim PDF As New clsPDF
    Dim i, j As Integer
    PDF.PDFTitle = "Staff Report"
    PDF.PDFFileName = App.Path & "\Reports\" & cmbStaff.Text & "(Detailed)" & ".pdf"
    PDF.PDFAuthor = "Jangid Corporation"
    PDF.PDFLoadAfm = App.Path
    PDF.PDFSetAdobePath = strAdobePath
    PDF.PDFView = True
    
    PDF.PDFSetBorder = BORDER_ALL
    PDF.PDFSetTopMargin = 1
    PDF.PDFSetBottomMargin = 1
    PDF.PDFSetRightMargin = 1
    PDF.PDFSetLeftMargin = 1
    PDF.PDFBeginDoc
        
        'Draw Border
        PDF.PDFDrawRectangle 1, 1, 19, 27
        'Header Part
        PDF.PDFSetTextColor = vbBlack
        PDF.PDFSetFont FONT_TIMES, 18, FONT_BOLD
        l = PDF.PDFGetStringWidth(UCase(strCollegeName), "Times-Bold", 18)
        strLeft = (19 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut UCase(strCollegeName), strLeft, 2
        
        PDF.PDFSetFont FONT_TIMES, 16, FONT_BOLD
        l = PDF.PDFGetStringWidth(strCity, "Times-Bold", 16)
        strLeft = (19 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut strCity, strLeft, 2.75
                
        PDF.PDFSetFont FONT_TIMES, 14, FONT_BOLD
        l = PDF.PDFGetStringWidth("STAFF EVALUATION BY STUDENTS", "Times-Bold", 14)
        strLeft = (19 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut "STAFF EVALUATION BY STUDENTS", strLeft, 3.65
        l = PDF.PDFGetStringWidth("DEPARTMENT OF " & UCase(cmbDept.Text), "Times-Bold", 14)
        strLeft = (19 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut "DEPARTMENT OF " & UCase(cmbDept.Text), strLeft, 4.4
        PDF.PDFSetFont FONT_TIMES, 12, FONT_BOLD
        PDF.PDFTextOut "Date:", 15.6, 5.25
        PDF.PDFTextOut Format(DateTime.Date, "dd/mm/yyyy"), 16.65, 5.25
        
        'Divider for header
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLineHor 1, 5.5, 19
        PDF.PDFSetLineWidth = 0.02
        PDF.PDFDrawLineHor 1, 6.25, 19
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLineHor 1, 7, 19
        PDF.PDFDrawLineHor 1, 22.25, 19
        PDF.PDFDrawLineHor 1, 16.25, 19
        PDF.PDFSetFont FONT_TIMES, 12, FONT_BOLD
        PDF.PDFTextOut "H.O.D", 2.5, 25.75
        PDF.PDFTextOut "Principal", 15, 25.75
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 27.25, 20, 27.25
        PDF.PDFSetLineWidth = 0.025
        PDF.PDFDrawLine 1, 27.3, 20, 27.3
        
        'Private for this module
        l = PDF.PDFGetStringWidth("STAFF REPORT", "Times-Bold", 14)
        strLeft = (19 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut "STAFF REPORT", strLeft, 5.15
                        
        PDF.PDFSetFont 2, 12, FONT_BOLD
        PDF.PDFTextOut "STAFF NAME:", 0.3, 6
        PDF.PDFTextOut cmbStaff.Text, 3.5, 6
                                                        
        PDF.PDFTextOut "SUB CODE", 0.3, 6.75
        PDF.PDFTextOut "DEPT", 3.1, 6.75
        PDF.PDFTextOut "SEM", 4.625, 6.75
        PDF.PDFTextOut "SEC", 6.025, 6.75
        PDF.PDFTextOut "PERFORMANCE", 13.5, 6.75
        PDF.PDFTextOut "COUNT", 17.35, 6.75
        PDF.PDFSetFont 2, 9, FONT_BOLD
        PDF.PDFTextOut "GROUP 1", 7.15, 6.75
        PDF.PDFTextOut "GROUP 2", 8.65, 6.75
        PDF.PDFTextOut "GROUP 3", 10.15, 6.75
        PDF.PDFTextOut "GROUP 4", 11.65, 6.75
        
        
        PDF.PDFDrawLineVer 3.9, 6.25, 10
        PDF.PDFDrawLineVer 5.4, 6.25, 10
        PDF.PDFDrawLineVer 6.75, 6.25, 10
        PDF.PDFDrawLineVer 8.1, 6.25, 10
        PDF.PDFDrawLineVer 9.6, 6.25, 10
        PDF.PDFDrawLineVer 11.1, 6.25, 10
        PDF.PDFDrawLineVer 12.6, 6.25, 10
        PDF.PDFDrawLineVer 14.1, 6.25, 10
        PDF.PDFDrawLineVer 18.25, 6.25, 10
                
        PDF.PDFSetFont 2, 10, FONT_NORMAL
        For i = 1 To mshGrid.Rows - 1
            PDF.PDFTextOut mshGrid.TextMatrix(i, 0), 0.75, 7.25 + i * 0.6
            If Len(mshGrid.TextMatrix(i, 2)) > 6 Then
                PDF.PDFSetFont 2, 8, FONT_NORMAL
                PDF.PDFTextOut mshGrid.TextMatrix(i, 2), 3.1, 7.25 + i * 0.6
                PDF.PDFSetFont 2, 10, FONT_NORMAL
            Else
                PDF.PDFTextOut mshGrid.TextMatrix(i, 2), 3.25, 7.25 + i * 0.6
            End If
            PDF.PDFTextOut mshGrid.TextMatrix(i, 3), 5, 7.25 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 4), 6.25, 7.25 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 5), 7.55, 7.25 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 6), 9.05, 7.25 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 7), 10.55, 7.25 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 8), 12.05, 7.25 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 9), 14.85, 7.25 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 10), 17.9, 7.25 + i * 0.6
        Next
        PDF.PDFSetFont FONT_TIMES, 12, FONT_BOLD
        PDF.PDFTextOut "GROUP REPRESENTATION:", 0.25, 17
        PDF.PDFSetFont 2, 10, FONT_NORMAL
        PDF.PDFTextOut "GROUP 1:    Planning & Organisation", 1.25, 17.8
        PDF.PDFTextOut "GROUP 2:    Presentation / Communication", 1.25, 18.4
        PDF.PDFTextOut "GROUP 3:    Student's Participation", 1.25, 19
        PDF.PDFTextOut "GROUP 4:    Class Management / Assessment Of Students", 1.25, 19.6
        
    PDF.PDFEndDoc

End Sub

Private Sub cmdEx_Click()
    On Error Resume Next
    Dim nodExpand As Node
    If tvExpanded = False Then
        cmdEx.Caption = "Collapse"
        tvExpanded = True
        For Each nodExpand In tvFeed.Nodes
            nodExpand.Expanded = True
        Next
    Else
        cmdEx.Caption = "Expand All"
        tvExpanded = False
        For Each nodExpand In tvFeed.Nodes
            nodExpand.Expanded = False
        Next
    End If
End Sub



Private Sub rptStaff()
    On Error Resume Next
    Dim strLeft As Double
    Dim PDF As New clsPDF
    Dim i, j As Integer
    PDF.PDFTitle = "Staff Report"
    PDF.PDFFileName = App.Path & "\Reports\" & cmbStaff.Text & ".pdf"
    PDF.PDFAuthor = "JURA"
    PDF.PDFLoadAfm = App.Path
    PDF.PDFSetAdobePath = strAdobePath
    PDF.PDFView = True
    
    PDF.PDFSetBorder = BORDER_ALL
    PDF.PDFSetTopMargin = 1
    PDF.PDFSetBottomMargin = 1
    PDF.PDFSetRightMargin = 1
    PDF.PDFSetLeftMargin = 1
    PDF.PDFBeginDoc
        
        'Draw Border
        PDF.PDFDrawRectangle 1, 1, 19, 27
        'Header Part
        PDF.PDFSetTextColor = vbBlack
        PDF.PDFSetFont FONT_TIMES, 18, FONT_BOLD
        l = PDF.PDFGetStringWidth(UCase(strCollegeName), "Times-Bold", 18)
        strLeft = (19 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut UCase(strCollegeName), strLeft, 2
        
        PDF.PDFSetFont FONT_TIMES, 16, FONT_BOLD
        l = PDF.PDFGetStringWidth(strCity, "Times-Bold", 16)
        strLeft = (19 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut strCity, strLeft, 2.75
        
        PDF.PDFSetFont FONT_TIMES, 14, FONT_BOLD
        l = PDF.PDFGetStringWidth("STAFF EVALUATION BY STUDENTS", "Times-Bold", 14)
        strLeft = (19 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut "STAFF EVALUATION BY STUDENTS", strLeft, 3.65
        l = PDF.PDFGetStringWidth("DEPARTMENT OF " & UCase(cmbDept.Text), "Times-Bold", 14)
        strLeft = (19 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut "DEPARTMENT OF " & UCase(cmbDept.Text), strLeft, 4.4
        PDF.PDFSetFont FONT_TIMES, 12, FONT_BOLD
        PDF.PDFTextOut "Date:", 15.6, 5.25
        PDF.PDFTextOut Format(DateTime.Date, "dd/mm/yyyy"), 16.65, 5.25
        
        'Divider for header
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLineHor 1, 5.5, 19
        PDF.PDFSetLineWidth = 0.02
        PDF.PDFDrawLineHor 1, 6.25, 19
        'Divide for signature
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLineHor 1, 7, 19
        'Bottom
        PDF.PDFDrawLineHor 1, 22.25, 19
        PDF.PDFSetFont FONT_TIMES, 12, FONT_BOLD
        PDF.PDFTextOut "H.O.D", 2.5, 25.75
        PDF.PDFTextOut "Principal", 15, 25.75
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 27.25, 20, 27.25
        PDF.PDFSetLineWidth = 0.025
        PDF.PDFDrawLine 1, 27.3, 20, 27.3
        
        'Private for this module
        l = PDF.PDFGetStringWidth("STAFF REPORT", "Times-Bold", 14)
        strLeft = (19 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut "STAFF REPORT", strLeft, 5.15
                        
        PDF.PDFSetFont 2, 12, FONT_BOLD
        PDF.PDFTextOut "STAFF NAME:", 0.3, 6
        PDF.PDFTextOut cmbStaff.Text, 3.5, 6
                                                        
        PDF.PDFTextOut "SUBJECT", 0.3, 6.75
        PDF.PDFTextOut "SUBJECT NAME", 5, 6.75
        PDF.PDFTextOut "DEPT", 10.75, 6.75
        PDF.PDFTextOut "BATCH", 12.225, 6.75
        PDF.PDFTextOut "SEM", 14.05, 6.75
        PDF.PDFTextOut "SEC", 15.25, 6.75
        PDF.PDFSetFont 2, 8, FONT_BOLD
        PDF.PDFTextOut "PERFORMANCE", 16.5, 6.75
        PDF.PDFSetFont 2, 12, FONT_BOLD
        
        PDF.PDFDrawLineVer 3.4, 6.25, 16
        PDF.PDFDrawLineVer 11.55, 6.25, 16
        PDF.PDFDrawLineVer 13.05, 6.25, 16
        PDF.PDFDrawLineVer 14.9, 6.25, 16
        PDF.PDFDrawLineVer 16.1, 6.25, 16
        PDF.PDFDrawLineVer 17.25, 6.25, 16
                
        PDF.PDFSetFont 2, 10, FONT_NORMAL
        For i = 1 To mshGrid.Rows - 1
            PDF.PDFTextOut mshGrid.TextMatrix(i, 0), 0.65, 7.25 + i * 0.6
            PDF.PDFTextOut Mid(UCase(mshGrid.TextMatrix(i, 1)), 1, 33), 2.65, 7.25 + i * 0.6
            If Len(mshGrid.TextMatrix(i, 2)) > 6 Then
                PDF.PDFSetFont 2, 8, FONT_NORMAL
                PDF.PDFTextOut mshGrid.TextMatrix(i, 2), 10.65, 7.25 + i * 0.6
                PDF.PDFSetFont 2, 10, FONT_NORMAL
            Else
                PDF.PDFTextOut mshGrid.TextMatrix(i, 2), 10.85, 7.25 + i * 0.6
            End If
            PDF.PDFTextOut mshGrid.TextMatrix(i, 3), 12.9, 7.25 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 4), 14.4, 7.25 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 5), 15.55, 7.25 + i * 0.6
            PDF.PDFTextOut mshGrid.TextMatrix(i, 10), 17.25, 7.25 + i * 0.6
        Next
    PDF.PDFEndDoc

End Sub

Private Sub rptClass()
    On Error Resume Next
    Dim strLeft As Double
    Dim PDF As New clsPDF
    Dim i, j As Integer
    PDF.PDFTitle = "Class Report"
    PDF.PDFFileName = App.Path & "\Reports\" & cmbDept.Text & "-" & cmbBatch.Text & "-" & cmbSem.Text & "-" & cmbSec.Text & ".pdf"
    PDF.PDFAuthor = "Jangid Corporation"
    PDF.PDFLoadAfm = App.Path
    PDF.PDFView = True
    
    PDF.PDFSetBorder = BORDER_ALL
    PDF.PDFSetTopMargin = 1
    PDF.PDFSetBottomMargin = 1
    PDF.PDFSetRightMargin = 1
    PDF.PDFSetLeftMargin = 1
    PDF.PDFSetAdobePath = strAdobePath
    PDF.PDFBeginDoc
        
        'Draw Border
        PDF.PDFDrawRectangle 1, 1, 19, 27
        'Header Part
        PDF.PDFSetTextColor = vbBlack
        PDF.PDFSetFont FONT_TIMES, 18, FONT_BOLD
        l = PDF.PDFGetStringWidth(UCase(strCollegeName), "Times-Bold", 18)
        strLeft = (19 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut UCase(strCollegeName), strLeft, 2
        
        PDF.PDFSetFont FONT_TIMES, 16, FONT_BOLD
        l = PDF.PDFGetStringWidth(strCity, "Times-Bold", 16)
        strLeft = (19 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut strCity, strLeft, 2.75
                
        PDF.PDFSetFont FONT_TIMES, 14, FONT_BOLD
        l = PDF.PDFGetStringWidth("STAFF EVALUATION BY STUDENTS", "Times-Bold", 14)
        strLeft = (19 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut "STAFF EVALUATION BY STUDENTS", strLeft, 3.65
        l = PDF.PDFGetStringWidth("DEPARTMENT OF " & UCase(cmbDept.Text), "Times-Bold", 14)
        strLeft = (19 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut "DEPARTMENT OF " & UCase(cmbDept.Text), strLeft, 4.4
        PDF.PDFSetFont FONT_TIMES, 12, FONT_BOLD
        PDF.PDFTextOut "Date:", 15.6, 5.25
        PDF.PDFTextOut Format(DateTime.Date, "dd/mm/yyyy"), 16.65, 5.25
        PDF.PDFSetFont FONT_TIMES, 9, FONT_BOLD
        PDF.PDFTextOut "Students Appeared: " & mshGrid.TextMatrix(1, 8), 0.25, 22.75
        
        
        'Divider for header
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLineHor 1, 5.5, 19
        PDF.PDFSetLineWidth = 0.02
        PDF.PDFDrawLineHor 1, 6.25, 19
        'Divide for signature
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLineHor 1, 7, 19
        'Bottom
        PDF.PDFDrawLineHor 1, 22.25, 19
        PDF.PDFSetFont FONT_TIMES, 12, FONT_BOLD
        PDF.PDFTextOut "H.O.D", 2.5, 25.75
        PDF.PDFTextOut "Principal", 15, 25.75
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 27.25, 20, 27.25
        PDF.PDFSetLineWidth = 0.025
        PDF.PDFDrawLine 1, 27.3, 20, 27.3
        
        'Private for this module
        l = PDF.PDFGetStringWidth("CLASS REPORT", "Times-Bold", 14)
        strLeft = (19 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut "CLASS REPORT", strLeft, 5.15
                        
        PDF.PDFSetFont 2, 12, FONT_BOLD
        PDF.PDFTextOut "BATCH:", 1.75, 6
        PDF.PDFTextOut cmbBatch.Text, 3.6, 6
        PDF.PDFTextOut "SEMESTER:", 7.8, 6
        PDF.PDFTextOut cmbSem.Text, 10.45, 6
        PDF.PDFTextOut "SECTION:", 13.85, 6
        PDF.PDFTextOut cmbSec.Text, 16.15, 6
                                                        
        PDF.PDFTextOut "SUB CODE", 0.3, 6.75
        PDF.PDFTextOut "SUBJECT NAME", 5, 6.75
        PDF.PDFTextOut "STAFF", 11.5, 6.75
        PDF.PDFTextOut "PERFORMANCE", 15.2, 6.75
        
        PDF.PDFDrawLineVer 3.9, 6.25, 16
        PDF.PDFDrawLineVer 10.65, 6.25, 16
        PDF.PDFDrawLineVer 15.85, 6.25, 16
                
        PDF.PDFSetFont 2, 9, FONT_NORMAL
        For i = 1 To mshGrid.Rows - 1
            PDF.PDFTextOut mshGrid.TextMatrix(i, 0), 0.75, 7.25 + i * 0.6
            PDF.PDFTextOut Mid(UCase(mshGrid.TextMatrix(i, 1)), 1, 32), 3.15, 7.25 + i * 0.6
            If Len(mshGrid.TextMatrix(i, 2)) > 25 Then
                PDF.PDFSetFont 2, 7, FONT_NORMAL
                PDF.PDFTextOut mshGrid.TextMatrix(i, 2), 9.85, 7.25 + i * 0.6
                PDF.PDFSetFont 2, 9, FONT_NORMAL
            Else
                PDF.PDFTextOut mshGrid.TextMatrix(i, 2), 9.85, 7.25 + i * 0.6
            End If
            PDF.PDFTextOut mshGrid.TextMatrix(i, 7), 16.65, 7.25 + i * 0.6
        Next
    PDF.PDFEndDoc
errHan:
    If Err.Number <> 0 Then
        'MsgBox "Error Number : " & Err.Number & vbCrLf & "Error Description : " & Err.Description & vbCrLf & "Error Location : frmReport-rptStaff()", vbInformation, "Jangid Corporation"
    End If
End Sub

Private Sub cmdFolder_Click()
      Shell "Explorer.exe " & App.Path & "\Reports", vbMaximizedFocus
End Sub

Private Sub cmdLoad_Click()
    tvLoad
End Sub

Private Sub cmdPrint_Click()
    If mshGrid.Rows = 1 Then
        MsgBox "Data not found for the given information", vbInformation, "Jangid Corporation"
        Exit Sub
    ElseIf mshGrid.TextMatrix(1, 0) = "" Then
        MsgBox "No data availabe for Report." & vbCrLf & "First select values to generate report", vbInformation, "Jangid Corporation"
        Exit Sub
    End If
    If strPointer = "Staff" Then
        rptStaff
    ElseIf strPointer = "Class" Then
        LoadClass
        rptClass
    ElseIf strPointer = "Subject" Then
        rptSubject
    End If
End Sub

Private Sub cmdSettings_Click()
    frmSettings.Show Modal, frmReport
End Sub

Private Sub cmdSub_Click()
    On Error Resume Next
    Dim strLeft As Double
    Dim PDF As New clsPDF
    Dim i, j As Integer
    Dim aStaffData() As String
    aStaffData = Split(cmbSubj.Text, "-")
    Dim rsRpt As New ADODB.Recordset
        Dim sql As String
        sql = "select s2.subjcode,s1.subjname,s3.deptshort as dept,s2.batch,s2.sem,s2.sec," & _
              "avg(s2.q1) as q1,avg(s2.q2) as q2,avg(s2.q3) as q3,avg(s2.q4) as q4,avg(s2.q5) as q5," & _
              "avg(s2.q6) as q6,avg(s2.q7) as q7,avg(s2.q8) as q8,avg(s2.q9) as q9,avg(s2.q10) as q10," & _
              "avg(s2.q11) as q11,avg(s2.q12) as q12,avg(s2.q13) as q13,avg(s2.q14) as q14,avg(s2.q15) as q15," & _
              "avg(s2.q16) as q16,avg(s2.q17) as q17,avg(s2.q18) as q18,avg(s2.q19) as q19,avg(s2.q20) as q20," & _
              "round(avg(s2.q1+s2.q2+s2.q3+s2.q4+s2.q5)/5,2) as g1," & _
              "round(avg(s2.q6+s2.q7+s2.q8+s2.q9+s2.q10)/5,2) as g2," & _
              "round(avg(s2.q11+s2.q12+s2.q13+s2.q14+s2.q15)/5,2) as g3," & _
              "round(avg(s2.q16+s2.q17+s2.q18+s2.q19+s2.q20)/5,2) as g4," & _
              "round(avg(s2.q1+s2.q2+s2.q3+s2.q4+s2.q5+s2.q6+s2.q7+s2.q8+s2.q9+s2.q10+s2.q11+s2.q12+s2.q13+s2.q14+s2.q15+s2.q16+s2.q17+s2.q18+s2.q19+s2.q20)/20,2) as overall," & _
              "count(s2.fid) as Count" & _
              " from master s2,subj s1,dept s3" & _
              " where s1.dept=s2.dept and s2.dept=s3.deptcode and s1.subjcode=s2.subjcode" & _
              " and s2.staffid='" & lblStaffID.Caption & "'" & _
              " and s2.subjcode='" & aStaffData(0) & "'" & _
              " and s2.dept='" & getDeptCode(aStaffData(1)) & "'" & _
              " and s2.batch='" & aStaffData(2) & "'" & _
              " and s2.sem='" & aStaffData(3) & "'" & _
              " and s2.sec='" & aStaffData(4) & "'" & _
              " group by  s2.batch, s2.sec,s2.subjcode,s1.subjname,s2.sem,s3.deptshort" & _
              " order by s2.subjcode"
        rsRpt.Open sql, conn, adOpenDynamic, adLockOptimistic, -1
    PDF.PDFTitle = "Staff Report"
    PDF.PDFFileName = App.Path & "\Reports\" & cmbStaff.Text & "(" & cmbSubj.Text & ")" & ".pdf"
    PDF.PDFAuthor = "Jangid Corporation"
    PDF.PDFLoadAfm = App.Path
    PDF.PDFSetAdobePath = strAdobePath
    PDF.PDFView = True
    
    PDF.PDFSetBorder = BORDER_ALL
    PDF.PDFSetTopMargin = 1
    PDF.PDFSetBottomMargin = 1
    PDF.PDFSetRightMargin = 1
    PDF.PDFSetLeftMargin = 1
    PDF.PDFBeginDoc
        
        'Draw Border
        PDF.PDFDrawRectangle 1, 1, 19, 27
        'Header Part
        PDF.PDFSetTextColor = vbBlack
        
        PDF.PDFSetFont FONT_TIMES, 18, FONT_BOLD
        l = PDF.PDFGetStringWidth(UCase(strCollegeName), "Times-Bold", 18)
        strLeft = (19 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut UCase(strCollegeName), strLeft, 2
        
        PDF.PDFSetFont FONT_TIMES, 16, FONT_BOLD
        l = PDF.PDFGetStringWidth(strCity, "Times-Bold", 16)
        strLeft = (19 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut strCity, strLeft, 2.75
        
        PDF.PDFSetFont FONT_TIMES, 14, FONT_BOLD
        l = PDF.PDFGetStringWidth("STAFF EVALUATION BY STUDENTS", "Times-Bold", 14)
        strLeft = (19 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut "STAFF EVALUATION BY STUDENTS", strLeft, 3.65
        l = PDF.PDFGetStringWidth("DEPARTMENT OF " & UCase(cmbDept.Text), "Times-Bold", 14)
        strLeft = (19 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut "DEPARTMENT OF " & UCase(cmbDept.Text), strLeft, 4.4
        PDF.PDFSetFont FONT_TIMES, 12, FONT_BOLD
        PDF.PDFTextOut "Date:", 15.6, 5.25
        PDF.PDFTextOut Format(DateTime.Date, "dd/mm/yyyy"), 16.65, 5.25
        
        'Divider for header
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLineHor 1, 5.5, 19
    
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLineHor 1, 7.75, 19
        PDF.PDFDrawLineHor 1, 24.25, 19
        PDF.PDFSetFont FONT_TIMES, 12, FONT_BOLD
        PDF.PDFTextOut "H.O.D", 2.5, 26.75
        PDF.PDFTextOut "Principal", 15, 26.75
        PDF.PDFSetLineWidth = 0.03
        PDF.PDFDrawLine 1, 27.25, 20, 27.25
        PDF.PDFSetLineWidth = 0.025
        PDF.PDFDrawLine 1, 27.3, 20, 27.3
        
        'Private for this module
        l = PDF.PDFGetStringWidth("STAFF REPORT", "Times-Bold", 14)
        strLeft = (19 - (l * 2.54) / 72) / 2
        PDF.PDFTextOut "STAFF REPORT", strLeft, 5.15
                        
        PDF.PDFSetFont 2, 12, FONT_BOLD
        PDF.PDFTextOut "STAFF NAME:", 0.3, 6
        PDF.PDFTextOut cmbStaff.Text, 3.5, 6
        PDF.PDFTextOut "SUBJECT CODE : " & aStaffData(0), 0.3, 6.75
        PDF.PDFTextOut "SUBJECT NAME : " & Mid(UCase(rsRpt.Fields("SUBJNAME")), 1, 35), 6, 6.75
        PDF.PDFTextOut "SUBJECT DEPT  : " & rsRpt.Fields("DEPT"), 0.3, 7.5
        PDF.PDFTextOut "BATCH : " & rsRpt.Fields("BATCH"), 6, 7.5
        PDF.PDFTextOut "SEMESTER : " & rsRpt.Fields("SEM"), 9.8, 7.5
        PDF.PDFTextOut "SECTION : " & rsRpt.Fields("SEC"), 14.5, 7.5
        
        PDF.PDFSetFont 2, 12, FONT_NORMAL
        PDF.PDFTextOut UCase("1.Planning & Organisation"), 0.3, 8.5
        PDF.PDFTextOut rsRpt.Fields("G1"), 18, 8.5
        PDF.PDFSetFont 2, 10, FONT_NORMAL
        PDF.PDFTextOut "1.1 Teacher comes to class in time.", 0.8, 9
        PDF.PDFTextOut Round(rsRpt.Fields("Q1"), 2), 18.06, 9
        PDF.PDFTextOut "1.2 Teacher is well planned.", 0.8, 9.5
        PDF.PDFTextOut Round(rsRpt.Fields("Q2"), 2), 18.06, 9.5
        PDF.PDFTextOut "1.3 Aims/Objectives made clear.", 0.8, 10
        PDF.PDFTextOut Round(rsRpt.Fields("Q3"), 2), 18.06, 10
        PDF.PDFTextOut "1.4 Subject matter organized in logical sequence.", 0.8, 10.5
        PDF.PDFTextOut Round(rsRpt.Fields("Q4"), 2), 18.06, 10.5
        PDF.PDFTextOut "1.5 Teacher comes well prepared in the subject.", 0.8, 11
        PDF.PDFTextOut Round(rsRpt.Fields("Q5"), 2), 18.06, 11
        
        PDF.PDFSetFont 2, 12, FONT_NORMAL
        PDF.PDFTextOut UCase("2.Presentation / Communication"), 0.3, 12
        PDF.PDFTextOut rsRpt.Fields("G2"), 18, 12
        PDF.PDFSetFont 2, 10, FONT_NORMAL
        PDF.PDFTextOut "2.1 Teacher speaks clearly and audibly.", 0.8, 12.5
        PDF.PDFTextOut Round(rsRpt.Fields("Q6"), 2), 18.06, 12.5
        PDF.PDFTextOut "2.2 Teacher writes and draws legibly.", 0.8, 13
        PDF.PDFTextOut Round(rsRpt.Fields("Q7"), 2), 18.06, 13
        PDF.PDFTextOut "2.3 Teacher provides examples of concepts / principles. Explanations are clear and effective.", 0.8, 13.5
        PDF.PDFTextOut Round(rsRpt.Fields("Q8"), 2), 18.06, 13.5
        PDF.PDFTextOut "2.4 Teacher's pace and level of instruction are suited to the attainment of students.", 0.8, 14
        PDF.PDFTextOut Round(rsRpt.Fields("Q9"), 2), 18.06, 14
        PDF.PDFTextOut "2.5 Teacher offers assistance and counselling to the needy students.", 0.8, 14.5
        PDF.PDFTextOut Round(rsRpt.Fields("Q10"), 2), 18.06, 14.5
        
        PDF.PDFSetFont 2, 12, FONT_NORMAL
        PDF.PDFTextOut UCase("3.Student's Participation"), 0.3, 15.5
        PDF.PDFTextOut rsRpt.Fields("G3"), 18, 15.5
        PDF.PDFSetFont 2, 10, FONT_NORMAL
        PDF.PDFTextOut "3.1 Teacher asks questions to promote interaction and reflecctive thinking.", 0.8, 16
        PDF.PDFTextOut Round(rsRpt.Fields("Q11"), 2), 18.06, 16
        PDF.PDFTextOut "3.2 Teacher encourages questioning / raising doubts by students and answers them well.", 0.8, 16.5
        PDF.PDFTextOut Round(rsRpt.Fields("Q12"), 2), 18.06, 16.5
        PDF.PDFTextOut "3.3 Teacher ensures learner activity and problems solving ability in the class.", 0.8, 17
        PDF.PDFTextOut Round(rsRpt.Fields("Q13"), 2), 18.06, 17
        PDF.PDFTextOut "3.4 Teacher encourages, compliments and praises originality and creativity displayed by the students.", 0.8, 17.5
        PDF.PDFTextOut Round(rsRpt.Fields("Q14"), 2), 18.06, 17.5
        PDF.PDFTextOut "3.5 Teacher is courteous and impartial in dealing with the students.", 0.8, 18
        PDF.PDFTextOut Round(rsRpt.Fields("Q15"), 2), 18.06, 18
        
        PDF.PDFSetFont 2, 12, FONT_NORMAL
        PDF.PDFTextOut UCase("4.Class Management / Assessment Of Students"), 0.3, 19
        PDF.PDFTextOut rsRpt.Fields("G4"), 18, 19
        PDF.PDFSetFont 2, 10, FONT_NORMAL
        PDF.PDFTextOut "4.1 Teacher engages classes regularly and maintains discipline.", 0.8, 19.5
        PDF.PDFTextOut Round(rsRpt.Fields("Q16"), 2), 18.06, 19.5
        PDF.PDFTextOut "4.2 Teacher covers the syllabus completely and at appropriate pace.", 0.8, 20
        PDF.PDFTextOut Round(rsRpt.Fields("Q17"), 2), 18.06, 20
        PDF.PDFTextOut "4.3 Teacher holds tests regularly which are helpful to students in building up confidence in their acquisition and", 0.8, 20.5
        PDF.PDFTextOut "application of knowledge.", 1, 21
        PDF.PDFTextOut Round(rsRpt.Fields("Q18"), 2), 18.06, 21
        PDF.PDFTextOut "4.4 Teacher's making of scripts is fair and impartial.", 0.8, 21.5
        PDF.PDFTextOut Round(rsRpt.Fields("Q19"), 2), 18.06, 21.5
        PDF.PDFTextOut "4.5 Teacher is prompt in valuing and returning the answer scripts providing feedback on performance.", 0.8, 22
        PDF.PDFTextOut Round(rsRpt.Fields("Q20"), 2), 18.06, 22
        
        PDF.PDFSetFont 2, 12, FONT_BOLD
        PDF.PDFTextOut "OVERALL PERFORMANCE : " & rsRpt.Fields("OVERALL"), 0.3, 23
        PDF.PDFSetFont 2, 10, FONT_BOLD
        PDF.PDFTextOut "No of students appeared. : " & rsRpt.Fields("COUNT"), 0.3, 24
        
        
        PDF.PDFDrawLineVer 18.8, 7.75, 16.5
                            
    PDF.PDFEndDoc
End Sub

Private Sub cmdSubjRpt_Click()
    On Error Resume Next
    strPointer = "Subject"
    FSubject = True
    frameControls.Visible = True
    frameGroup.Visible = True
    fSem.Visible = False
    
    cmdStaff.ForeColor = vbBlack
    cmdStaff.CaptionEffect = Normal
    cmdClass.ForeColor = vbBlack
    cmdClass.CaptionEffect = Normal
    cmdSubjRpt.ForeColor = vbRed
    cmdSubjRpt.CaptionEffect = outline
    cmdTree.ForeColor = vbBlack
    cmdTree.CaptionEffect = Normal
    
    lblRptDept.Visible = True
    lblRptSec.Visible = True
    lblRptBatch.Visible = True
    lblRptSemester.Visible = True
    lblStu.Visible = True
    
    cmbSec.Visible = True
    cmbSem.Visible = True
    cmbBatch.Visible = True
    cmbDept.Visible = True
    
    cmdClassGo.Visible = True
    
    lblRptStaff.Visible = False
    cmbStaff.Visible = False
    lblSubj.Visible = True
    lblStaff.Visible = True
    lblStaffName.Visible = True
    cmbSubj.Visible = True
    cmbSubj.Clear
    cbStaff.Visible = False
    cmdSub.Visible = False
    
    tvFeed.Top = mshGrid.Top
    tvFeed.Height = mshGrid.Height
    cmdEx.Visible = False
    cmdLoad.Visible = False
    tvFeed.Visible = False
    lblSubjName.Visible = True
    cmdDetailedRpt.Visible = False
    
    mshGrid.Visible = True
    mshGrid.ClearStructure
    cmdPrint.Visible = True
    cmbDept.SetFocus
End Sub
Private Sub LoadSubject()
    On Error Resume Next
    Dim rsOverall As New ADODB.Recordset
    Dim sqlOverall As String
    sqlOverall = "select m.fid,m.q1,m.q2,m.q3,m.q4,m.q5,m.q6,m.q7,m.q8,m.q9,m.q10,m.q11,m.q12,m.q13,m.q14,m.q15,m.q16,m.q17,m.q18,m.q19,m.q20,((m.q1+m.q2+m.q3+m.q4+m.q5)/5),((m.q6+m.q7+m.q8+m.q9+m.q10)/5),((m.q11+m.q12+m.q13+m.q14+m.q15)/5),((m.q16+m.q17+m.q18+m.q19+m.q20)/5),((m.q1+m.q2+m.q3+m.q4+m.q5+m.q6+m.q7+m.q8+m.q9+m.q10+m.q11+m.q12+m.q13+m.q14+m.q15+m.q16+m.q17+m.q18+m.q19+m.q20)/20) from master m,staff s where m.staffid=s.staffid and m.dept='" & iDept & "' and m.sem='" & iSem & "' and m.batch='" & Mid(iBatch, 3, 2) & "' and m.sec='" & strSec & "' and m.subjcode='" & cmbSubj.Text & "' order by m.fid"
    rsOverall.Open sqlOverall, conn, adOpenDynamic, adLockOptimistic, -1
    Set mshGrid.DataSource = rsOverall
    
    Dim rsStudCount As New ADODB.Recordset
    Dim sqlCount As String
    sqlCount = "select count(unique m.fid) from master m where m.dept='" & iDept & "' and m.sem='" & iSem & "' and m.batch='" & Mid(iBatch, 3, 2) & "' and m.sec='" & strSec & "' and m.subjcode='" & cmbSubj.Text & "'"
    rsStudCount.CursorLocation = adUseClient
    rsStudCount.Open sqlCount, conn, adOpenDynamic, adLockOptimistic, -1
    lblStu.Caption = "Student Count : " & rsStudCount.Fields(0)
     
     tvFeed.Visible = False
     mshGrid.ColWidth(0) = 500
     mshGrid.ColAlignment(0) = flexAlignCenterCenter
     mshGrid.ColAlignment(1) = flexAlignCenterCenter
     For i = 1 To 20
        mshGrid.ColAlignment(i) = flexAlignCenterCenter
        mshGrid.TextMatrix(0, i) = i
        mshGrid.ColWidth(i) = 450
     Next
     For i = 1 To 20
        mshGrid.ColAlignmentFixed(i) = flexAlignCenterCenter
     Next
     For i = 21 To 25
        mshGrid.ColWidth(i) = 850
        mshGrid.ColAlignment(i) = flexAlignCenterCenter
        mshGrid.ColAlignmentFixed(i) = flexAlignCenterCenter
     Next
     mshGrid.TextMatrix(0, 21) = "Group 1"
     mshGrid.TextMatrix(0, 22) = "Group 2"
     mshGrid.TextMatrix(0, 23) = "Group 3"
     mshGrid.TextMatrix(0, 24) = "Group 4"
     mshGrid.TextMatrix(0, 25) = "Overall"
     mshGrid.Visible = True
     rsStudCount.Close
End Sub

Private Sub LoadStaff()
    'On Error Resume Next
    Dim rsStaff As New ADODB.Recordset
    Dim sqlStaff As String
    Dim iYear As Integer
    If rbOdd.Value = vbChecked Then
        iYear = Mid(DateTime.Year(DateTime.Date$), 3, 2) - 3
        sqlStaff = "select s2.subjcode,s1.subjname,s3.deptshort as dept,s2.batch,s2.sem,s2.sec,round(avg(s2.q1+s2.q2+s2.q3+s2.q4+s2.q5)/5,2) ,round(avg(s2.q6+s2.q7+s2.q8+s2.q9+s2.q10)/5,2),round(avg(s2.q11+s2.q12+s2.q13+s2.q14+s2.q15)/5,2),round(avg(s2.q16+s2.q17+s2.q18+s2.q19+s2.q20)/5,2),round(avg(s2.q1+s2.q2+s2.q3+s2.q4+s2.q5+s2.q6+s2.q7+s2.q8+s2.q9+s2.q10+s2.q11+s2.q12+s2.q13+s2.q14+s2.q15+s2.q16+s2.q17+s2.q18+s2.q19+s2.q20)/20,2),count(s2.fid) as Count  from master s2,subj s1,dept s3 where s1.dept=s2.dept and s2.dept=s3.deptcode and s1.subjcode=s2.subjcode and s2.staffid='" & lblStaffID.Caption & "' and s2.sem in (1,3,5,7) and s2.batch >= '" & iYear & "'  group by s2.sec,s2.subjcode,s1.subjname,s2.batch,s2.sem,s3.deptshort order by s2.subjcode"
    ElseIf rbEven.Value = vbChecked Then
        iYear = Mid(DateTime.Year(DateTime.Date$), 3, 2) - 4
        sqlStaff = "select s2.subjcode,s1.subjname,s3.deptshort as dept,s2.batch,s2.sem,s2.sec,round(avg(s2.q1+s2.q2+s2.q3+s2.q4+s2.q5)/5,2) ,round(avg(s2.q6+s2.q7+s2.q8+s2.q9+s2.q10)/5,2),round(avg(s2.q11+s2.q12+s2.q13+s2.q14+s2.q15)/5,2),round(avg(s2.q16+s2.q17+s2.q18+s2.q19+s2.q20)/5,2),round(avg(s2.q1+s2.q2+s2.q3+s2.q4+s2.q5+s2.q6+s2.q7+s2.q8+s2.q9+s2.q10+s2.q11+s2.q12+s2.q13+s2.q14+s2.q15+s2.q16+s2.q17+s2.q18+s2.q19+s2.q20)/20,2),count(s2.fid) as Count  from master s2,subj s1,dept s3 where s1.dept=s2.dept and s2.dept=s3.deptcode and s1.subjcode=s2.subjcode and s2.staffid='" & lblStaffID.Caption & "' and s2.sem in (2,4,6,8) and s2.batch >= '" & iYear & "'  group by s2.sec,s2.subjcode,s1.subjname,s2.batch,s2.sem,s3.deptshort order by s2.subjcode"
    End If
    rsStaff.Open sqlStaff, conn, adOpenDynamic, adLockOptimistic, -1
    Set mshGrid.DataSource = rsStaff
    
    tvFeed.Visible = False
    mshGrid.ColWidth(0) = 900
    mshGrid.ColWidth(1) = 4000
    mshGrid.ColWidth(2) = 800
    mshGrid.ColWidth(3) = 800
    mshGrid.ColWidth(4) = 800
    mshGrid.ColWidth(5) = 800
    mshGrid.ColWidth(6) = 750
    mshGrid.ColWidth(7) = 750
    mshGrid.ColWidth(8) = 750
    mshGrid.ColWidth(9) = 750
    mshGrid.ColWidth(10) = 1700
    mshGrid.ColWidth(11) = 600
    
    mshGrid.ColAlignment(0) = flexAlignCenterCenter
    mshGrid.ColAlignment(2) = flexAlignCenterCenter
    mshGrid.ColAlignment(3) = flexAlignCenterCenter
    mshGrid.ColAlignment(4) = flexAlignCenterCenter
    mshGrid.ColAlignment(5) = flexAlignCenterCenter
    mshGrid.ColAlignment(6) = flexAlignCenterCenter
    mshGrid.ColAlignment(7) = flexAlignCenterCenter
    mshGrid.ColAlignment(8) = flexAlignCenterCenter
    mshGrid.ColAlignment(9) = flexAlignCenterCenter
    mshGrid.ColAlignment(10) = flexAlignCenterCenter
    mshGrid.ColAlignment(11) = flexAlignCenterCenter
    mshGrid.ColAlignmentFixed(0) = flexAlignCenterCenter
    mshGrid.ColAlignmentFixed(1) = flexAlignCenterCenter
    mshGrid.ColAlignmentFixed(2) = flexAlignCenterCenter
    mshGrid.ColAlignmentFixed(3) = flexAlignCenterCenter
    mshGrid.ColAlignmentFixed(4) = flexAlignCenterCenter
    mshGrid.ColAlignmentFixed(5) = flexAlignCenterCenter
    mshGrid.ColAlignmentFixed(6) = flexAlignCenterCenter
    mshGrid.ColAlignmentFixed(7) = flexAlignCenterCenter
    mshGrid.ColAlignmentFixed(8) = flexAlignCenterCenter
    mshGrid.ColAlignmentFixed(9) = flexAlignCenterCenter
    mshGrid.ColAlignmentFixed(10) = flexAlignCenterCenter
    mshGrid.ColAlignmentFixed(11) = flexAlignCenterCenter
    
    mshGrid.TextMatrix(0, 0) = "Subject"
    mshGrid.TextMatrix(0, 1) = "Subject Name"
    mshGrid.TextMatrix(0, 2) = "Dept"
    mshGrid.TextMatrix(0, 3) = "Batch"
    mshGrid.TextMatrix(0, 4) = "Semester"
    mshGrid.TextMatrix(0, 5) = "Section"
    mshGrid.TextMatrix(0, 6) = "Group 1"
    mshGrid.TextMatrix(0, 7) = "Group 2"
    mshGrid.TextMatrix(0, 8) = "Group 3"
    mshGrid.TextMatrix(0, 9) = "Group 4"
    mshGrid.TextMatrix(0, 10) = "Overall Performance"
    mshGrid.TextMatrix(0, 11) = "Count"
    mshGrid.RowHeightMin = 350
End Sub

Private Sub cmdStaff_Click()
    On Error Resume Next
    strPointer = "Staff"
    frameControls.Visible = True
    frameGroup.Visible = True
    fSem.Visible = True
    rbOdd.Value = vbChecked
        
    cmdStaff.ForeColor = vbRed
    cmdStaff.CaptionEffect = outline
    cmdClass.ForeColor = vbBlack
    cmdClass.CaptionEffect = Normal
    cmdSubjRpt.ForeColor = vbBlack
    cmdSubjRpt.CaptionEffect = Normal
    cmdTree.ForeColor = vbBlack
    cmdTree.CaptionEffect = Normal
    
    lblRptDept.Visible = True
    cmbDept.Visible = True
    lblRptStaff.Visible = True
    cmbStaff.Visible = True
    cmbStaff.Text = ""
    
    lblRptSec.Visible = False
    lblRptBatch.Visible = False
    lblRptSemester.Visible = False
    cmbSec.Visible = False
    cmbSem.Visible = False
    cmbBatch.Visible = False
    lblStu.Visible = False
    lblSubj.Visible = False
    cmbSubj.Visible = False
    lblSubjName.Visible = False
    lblStaff.Visible = False
    lblStaffName.Visible = False
    cmdSub.Visible = False
    
    cmdClassGo.Visible = False
    cmdDetailedRpt.Visible = True
    cmdPrint.Visible = True
        
    tvFeed.Top = mshGrid.Top
    tvFeed.Height = mshGrid.Height
    cmdEx.Visible = False
    cmdLoad.Visible = False
    tvFeed.Visible = False
    cbStaff.Visible = True
    cbStaff.Value = vbUnchecked
    mshGrid.Visible = True
    mshGrid.ClearStructure
    cmbDept.SetFocus
End Sub

Private Sub cmdTree_Click()
    On Error Resume Next
    frameControls.Visible = False
    frameGroup.Visible = False
    mshGrid.Visible = False
    tvFeed.Top = 840
    cmdEx.Top = 600
    cmdLoad.Top = 600
    tvFeed.Height = frameReport.Height - 240 - tvFeed.Top
    cmdEx.Visible = True
    cmdLoad.Visible = True
    tvFeed.Visible = True
    cmdPrint.Visible = True
    
    cmdStaff.ForeColor = vbBlack
    cmdStaff.CaptionEffect = Normal
    cmdClass.ForeColor = vbBlack
    cmdClass.CaptionEffect = Normal
    cmdSubjRpt.ForeColor = vbBlack
    cmdSubjRpt.CaptionEffect = Normal
    cmdTree.ForeColor = vbRed
    cmdTree.CaptionEffect = outline
    cmdLoad.SetFocus
End Sub




Private Sub Form_Activate()
    On Error Resume Next
    tvExpanded = False
    cmbDept_Load cmbDept
    cmbStaff_Load cmbStaff
    cmbSec_Load cmbSec
    cmbBatch_Load cmbBatch
    cmbSem_Load cmbSem
End Sub

Private Sub Form_Load()
    frameReport.Left = 240
    frameReport.Top = 240
    frameReport.Width = Screen.Width - 480
    frameReport.Height = Screen.Height - 1440
    
    mshGrid.Left = 240
    mshGrid.Width = frameReport.Width - 480
    mshGrid.Height = frameReport.Height - mshGrid.Top - 240
    mshGrid.Visible = False
    
    tvFeed.Left = 240
    tvFeed.Width = frameReport.Width - 480
    tvFeed.Height = mshGrid.Height
    tvFeed.Visible = False
    
    cmdEx.Top = tvFeed.Top
    
    
    frameGroup.Left = frameReport.Width - frameGroup.Width - 240
End Sub

Private Sub tvLoad()
    On Error Resume Next

    Dim nod As Node
    Set nod = tvFeed.Nodes.Add(, , "Root", "Departments")
    
    Dim rsDept As New ADODB.Recordset
    Dim rsStaff As New ADODB.Recordset
    Dim rsSubj As New ADODB.Recordset
    Dim rsSec As New ADODB.Recordset
    Dim rsScore As New ADODB.Recordset
    Dim sqlScore As String
    rsDept.CursorLocation = adUseClient
    rsStaff.CursorLocation = adUseClient
    rsSubj.CursorLocation = adUseClient
    rsSec.CursorLocation = adUseClient
    rsScore.CursorLocation = adUseClient
    
    Dim tDeptShort As String
    Dim tDeptCode As String
    Dim tStaffID As String
    Dim tStaffName As String
    Dim tSubj As String
    Dim tSec As String
    
    rsDept.Open "select deptcode,deptshort from dept", conn, adOpenDynamic, adLockOptimistic, -1
    rsDept.MoveFirst
    
    For i = 0 To rsDept.RecordCount - 1
        
        tDeptCode = rsDept.Fields(0)
        tDeptShort = rsDept.Fields(1)
        
        Set nod = tvFeed.Nodes.Add("Root", tvwChild, tDeptShort, tDeptShort)
        rsStaff.Open "select staffid,staffname from staff where dept='" & tDeptCode & "'", conn, adOpenDynamic, adLockOptimistic, -1
        rsStaff.MoveFirst
        
        For j = 0 To rsStaff.RecordCount - 1
        
            tStaffID = rsStaff.Fields(0)
            tStaffName = rsStaff.Fields(1)
            
            tvFeed.Nodes.Add tDeptShort, tvwChild, tStaffName, tStaffName
            rsSubj.Open "select subjcode from staffhandle where staffid='" & tStaffID & "'", conn, adOpenDynamic, adLockOptimistic, -1
            rsSubj.MoveFirst
            
            For k = 0 To rsSubj.RecordCount - 1
                
                tSubj = rsSubj.Fields(0)
                
                tvFeed.Nodes.Add tStaffName, tvwChild, tSubj, tSubj
                rsSec.Open "select unique sec from staffhandle where staffid='" & tStaffID & "' and subjcode='" & tSubj & "' order by sec", conn, adOpenDynamic, adLockOptimistic, -1
                rsSec.MoveFirst
                
                For m = 0 To rsSec.RecordCount - 1
                
                    tSec = rsSec.Fields(0)
                    
                    tvFeed.Nodes.Add tSubj, tvwChild, tSec & j & k & m, tSec
                    
                    sqlScore = "select round(avg(m.q1),2),round(avg(m.q2),2),round(avg(m.q3),2),round(avg(m.q4),2),round(avg(m.q5),2),round(avg(m.q6),2),round(avg(m.q7),2),round(avg(m.q8),2),round(avg(m.q9),2),round(avg(m.q10),2),round(avg(m.q11),2),round(avg(m.q12),2),round(avg(m.q13),2),round(avg(m.q14),2),round(avg(m.q15),2),round(avg(m.q16),2),round(avg(m.q17),2),round(avg(m.q18),2),round(avg(m.q19),2),round(avg(m.q20),2), round(avg(m.q1+m.q2+m.q3+m.q4+m.q5)/5,2) as g1,round(avg(m.q6+m.q7+m.q8+m.q9+m.q10)/5,2) as g2,round(avg(m.q11+m.q12+m.q13+m.q14+m.q15)/5,2) as g3,round(avg(m.q16+m.q17+m.q18+m.q19+m.q20)/5,2) as g4,round(avg(m.q1+m.q2+m.q3+m.q4+m.q5+m.q6+m.q7+m.q8+m.q9+m.q10+m.q11+m.q12+m.q13+m.q14+m.q15+m.q16+m.q17+m.q18+m.q19+m.q20)/20,2) as op from master m where m.subjcode='" & tSubj & "' and m.staffid='" & tStaffID & "' and m.dept='" & tDeptCode & "' and m.sec='" & tSec & "' group by m.sec,m.subjcode"
                    rsScore.Open sqlScore, conn, adOpenDynamic, adLockOptimistic, -1
                    tvFeed.Nodes.Add CStr(tSec & j & k & m), tvwChild, "Overall" & j & k & m, "Overall Performance: " & rsScore.Fields(24)
                    
                    tvFeed.Nodes.Add CStr("Overall" & j & k & m), tvwChild, "g1" & j & k & m, "Planning and Organisation:    " & rsScore.Fields(20)
                    
                    tvFeed.Nodes.Add CStr("g1" & j & k & m), tvwChild, "q1" & j & k & m, "1.1 Teacher comes to class in time: " & rsScore.Fields(0)
                    tvFeed.Nodes.Add CStr("g1" & j & k & m), tvwChild, "q2" & j & k & m, "1.2 Teacher is well planned: " & rsScore.Fields(1)
                    tvFeed.Nodes.Add CStr("g1" & j & k & m), tvwChild, "q3" & j & k & m, "1.3 Aims/Objectives made clear: " & rsScore.Fields(2)
                    tvFeed.Nodes.Add CStr("g1" & j & k & m), tvwChild, "q4" & j & k & m, "1.4 Subject matter organized in logical sequence: " & rsScore.Fields(3)
                    tvFeed.Nodes.Add CStr("g1" & j & k & m), tvwChild, "q5" & j & k & m, "1.5 Teacher comes well prepared in the subject: " & rsScore.Fields(4)
                    
                    tvFeed.Nodes.Add CStr("Overall" & j & k & m), tvwChild, "g2" & j & k & m, "Presentation / Communication: " & rsScore.Fields(21)
                                        
                    tvFeed.Nodes.Add CStr("g2" & j & k & m), tvwChild, "q6" & j & k & m, "2.1 Teacher speaks clearly and audibly: " & rsScore.Fields(5)
                    tvFeed.Nodes.Add CStr("g2" & j & k & m), tvwChild, "q7" & j & k & m, "2.2 Teacher writes and draws legibly: " & rsScore.Fields(6)
                    tvFeed.Nodes.Add CStr("g2" & j & k & m), tvwChild, "q8" & j & k & m, "2.3 Teacher provides examples of concepts / principles. Explanations are clear and effective: " & rsScore.Fields(7)
                    tvFeed.Nodes.Add CStr("g2" & j & k & m), tvwChild, "q9" & j & k & m, "2.4 Teacher's pace and level of instruction are suited to the attainment of students: " & rsScore.Fields(8)
                    tvFeed.Nodes.Add CStr("g2" & j & k & m), tvwChild, "q10" & j & k & m, "2.5 Teacher offers assistance and counselling to the needy students: " & rsScore.Fields(9)
                                        
                    tvFeed.Nodes.Add CStr("Overall" & j & k & m), tvwChild, "g3" & j & k & m, "Student's Participation: " & rsScore.Fields(22)
                    
                    tvFeed.Nodes.Add CStr("g3" & j & k & m), tvwChild, "q11" & j & k & m, "3.1 Teacher asks questions to promote interaction and reflecctive thinking: " & rsScore.Fields(10)
                    tvFeed.Nodes.Add CStr("g3" & j & k & m), tvwChild, "q12" & j & k & m, "3.2 Teacher encourages questioning / raising doubts by students and answers them well: " & rsScore.Fields(11)
                    tvFeed.Nodes.Add CStr("g3" & j & k & m), tvwChild, "q13" & j & k & m, "3.3 Teacher ensures learner activity and problems solving ability in the class: " & rsScore.Fields(12)
                    tvFeed.Nodes.Add CStr("g3" & j & k & m), tvwChild, "q14" & j & k & m, "3.4 Teacher encourages, compliments and praises originality and creativity displayed by the students: " & rsScore.Fields(13)
                    tvFeed.Nodes.Add CStr("g3" & j & k & m), tvwChild, "q15" & j & k & m, "3.5 Teacher is courteous and impartial in dealing with the students: " & rsScore.Fields(14)
                    
                    tvFeed.Nodes.Add CStr("Overall" & j & k & m), tvwChild, "g4" & j & k & m, "Class Management: " & rsScore.Fields(23)
                    
                    tvFeed.Nodes.Add CStr("g4" & j & k & m), tvwChild, "q16" & j & k & m, "4.1 Teacher engages classes regularly and maintains discipline: " & rsScore.Fields(15)
                    tvFeed.Nodes.Add CStr("g4" & j & k & m), tvwChild, "q17" & j & k & m, "4.2 Teacher covers the syllabus completely and at appropriate pace: " & rsScore.Fields(16)
                    tvFeed.Nodes.Add CStr("g4" & j & k & m), tvwChild, "q18" & j & k & m, "4.3 Teacher holds tests regularly which are helpful to students in building up confidence in their acquisition and application of knowledge: " & rsScore.Fields(17)
                    tvFeed.Nodes.Add CStr("g4" & j & k & m), tvwChild, "q19" & j & k & m, "4.4 Teacher's making of scripts is fair and impartial.: " & rsScore.Fields(18)
                    tvFeed.Nodes.Add CStr("g4" & j & k & m), tvwChild, "q20" & j & k & m, "4.5 Teacher is prompt in valuing and returning the answer scripts providing feedback on performance: " & rsScore.Fields(19)
                    
                    rsSec.MoveNext
                Next
                
                rsSubj.MoveNext
                rsSec.Close
                rsScore.Close
                
            Next
            rsStaff.MoveNext
            rsSubj.Close
            
        Next
        rsDept.MoveNext
        rsStaff.Close
        
    Next
    tvFeed.Nodes(1).Expanded = True
    If Err.Number = 3021 Then
        MsgBox "Insufficient data in the database." & vbCrLf & "Check whether department,staff and subject are created.", vbInformation, "Jangid Corporation"
    ElseIf Err.Number = 0 Then
    Else
        MsgBox "Error Number : " & Err.Number & vbCrLf & "Error Description : " & Err.Description & vbCrLf & "Error Location : frmReport-tvLoad()", vbInformation, "Jangid Corporation"
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    frmMain.Show
End Sub
Private Sub rbEven_Change(Value As CheckBoxConstants)
    LoadStaff
End Sub

Private Sub rbOdd_Change(Value As CheckBoxConstants)
    LoadStaff
End Sub
