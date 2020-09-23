VERSION 5.00
Begin VB.Form frmReports 
   Caption         =   "Flexi Reports/Query"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11580
   Icon            =   "frmReports.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   11580
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display"
      Height          =   375
      Left            =   5820
      TabIndex        =   6
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtrptname 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5820
      MaxLength       =   255
      TabIndex        =   3
      Top             =   3600
      Width           =   2295
   End
   Begin VB.ComboBox cborptid 
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
      ItemData        =   "frmReports.frx":27A2
      Left            =   5820
      List            =   "frmReports.frx":27B2
      TabIndex        =   2
      Text            =   "Report ID"
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label lblrptname 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   4335
      TabIndex        =   5
      Top             =   3600
      Width           =   1305
   End
   Begin VB.Label lblrptid 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report ID :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   4680
      TabIndex        =   4
      Top             =   3000
      Width           =   945
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Select Report ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   1
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "FLEXI REPORTS AND QUERY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cborptid_Click()
Select Case cborptid.Text
Case 3215
    txtrptname.Text = "Customer Details"
Case 3216
    txtrptname.Text = "Notification Details"
Case 3217
    txtrptname.Text = "Employee Details"
Case 3218
    txtrptname.Text = "Department List"
End Select

End Sub

Private Sub cmdDisplay_Click()
Select Case cborptid.Text
Case 3215
    rptCustDetails.Show
Case 3216
    rptNotification.Show
Case 3217
    rptEmployee.Show
Case 3218
    rptDepartment.Show
Case Else
    MsgBox "Select Valid Report ID"
End Select

End Sub
