VERSION 5.00
Begin VB.Form frmMasterLogin 
   Caption         =   "Master Login"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9240
   Icon            =   "frmMasterLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   9240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmd_exit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   8160
      TabIndex        =   6
      Top             =   3240
      Width           =   855
   End
   Begin VB.Frame frm_security 
      Caption         =   "Security..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   1920
      TabIndex        =   0
      Top             =   1200
      Width           =   5415
      Begin VB.CommandButton Cmd_entre 
         Caption         =   "&Login..."
         Default         =   -1  'True
         Height          =   375
         Left            =   4080
         Picture         =   "frmMasterLogin.frx":27A2
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txt_user 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Text            =   "Administrator"
         Top             =   360
         Width           =   3975
      End
      Begin VB.TextBox txt_pwd 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   960
         Width           =   3975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   960
         Width           =   690
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "User name"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   765
      End
   End
   Begin VB.Shape Shape2 
      Height          =   2655
      Left            =   0
      Top             =   1080
      Width           =   9135
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   9135
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   120
      Picture         =   "frmMasterLogin.frx":4717
      Stretch         =   -1  'True
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "1.0"
      Height          =   195
      Left            =   8520
      TabIndex        =   8
      Top             =   720
      Width           =   225
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Call Center Management System"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   645
      Left            =   1200
      TabIndex        =   7
      Top             =   120
      Width           =   7545
   End
End
Attribute VB_Name = "frmMasterLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_entre_Click()
If txt_user = "Administrator" And txt_pwd = "hotline" Then
Unload Me
Load frmLogin
frmLogin.Show
Else
MsgBox "Invalid user name/password", vbCritical, "CCMS"
txt_pwd = ""
txt_pwd.SetFocus
End If

End Sub

Private Sub Cmd_exit_Click()
End
End Sub

Private Sub Form_Activate()
txt_pwd.SetFocus


End Sub




