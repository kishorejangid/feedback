VERSION 5.00
Object = "{4C5605EA-720A-490B-820A-E3CDEE939855}#1.0#0"; "vkusercontrolsxp.ocx"
Begin VB.Form frmPop 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   ClientHeight    =   3120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin vkUserContolsXP.vkFrame framePop 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5530
      Caption         =   "     Feedback Update"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      BorderColor     =   8421504
      BorderWidth     =   2
      Begin vkUserContolsXP.vkLabel lblMsg 
         Height          =   435
         Left            =   120
         TabIndex        =   1
         Top             =   2040
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   767
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Checking For Updates..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: If you don't want to check for Updates in future, disable auto update in            settings."
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   2640
         Width           =   5535
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgFeed 
         Height          =   1500
         Left            =   960
         Picture         =   "frmPop.frx":0000
         Top             =   480
         Width           =   3900
      End
   End
End
Attribute VB_Name = "frmPop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    framePop.Top = 0
    framePop.Left = 0
    Me.Width = framePop.Width
    Me.Height = framePop.Height
    CreateRoundRectFromWindow Me, 7, 7
End Sub
