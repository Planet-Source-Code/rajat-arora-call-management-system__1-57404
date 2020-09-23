VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1320
      Top             =   4560
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7680
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      X1              =   720
      X2              =   7920
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   960
      X2              =   7680
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   600
      X2              =   8040
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      Height          =   615
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   7695
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   4080
      Picture         =   "frmMain.frx":0000
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Connecting Database please wait..."
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   7440
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   7695
      Left            =   0
      Picture         =   "frmMain.frx":27A2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8655
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   4080
      Picture         =   "frmMain.frx":54831
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()

If pb.Value < 99 Then
pb.Value = pb.Value + 3
Else
Timer1.Enabled = False
Unload Me
Load frmMasterLogin
frmMasterLogin.Show
Exit Sub
End If

End Sub


