VERSION 5.00
Begin VB.Form frmusermanager 
   Caption         =   "User Manager"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6045
   Icon            =   "frmusermanager.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3525
   ScaleWidth      =   6045
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   375
      Left            =   1440
      Picture         =   "frmusermanager.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   5040
      Picture         =   "frmusermanager.frx":4717
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3000
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   5775
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   360
         Width           =   4335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   720
         Width           =   4335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1200
         Width           =   4335
      End
      Begin VB.Label Label4 
         Caption         =   "User Name"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "User ID"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Confirm ID"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   735
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   2880
      Picture         =   "frmusermanager.frx":6548
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Add User"
      Height          =   375
      Left            =   120
      Picture         =   "frmusermanager.frx":84BD
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   1560
      TabIndex        =   14
      Top             =   600
      Width           =   4155
   End
   Begin VB.Label Label2 
      Caption         =   "User ID"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select User"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   825
   End
End
Attribute VB_Name = "frmusermanager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim users As ADODB.Recordset
Dim i As Integer



Private Sub Combo1_Click()
users.MoveFirst
While users.EOF = False
    If Combo1.Text = users.Fields(0) Then
        Label3.Caption = users.Fields(1)
    
 Exit Sub
 Else
    users.MoveNext
    End If
Wend
End Sub

Private Sub Command1_Click()
If Text2 = Text3 Then
users.AddNew
users.Fields(0) = Text1.Text
users.Fields(1) = Text2.Text
users.Update
Frame1.Enabled = False
Command1.Enabled = False
Else
MsgBox "Confirm ID not match", vbInformation, "Confirm ID"
Text3 = ""
Text3.SetFocus
End If
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Frame1.Enabled = True
Command1.Enabled = True
End Sub

Private Sub Command3_Click()
'**** DELETE FROM THE ISSUE TABLE **********
        users.MoveFirst
        While users.EOF = False
        If Combo1.Text = users.Fields(0) Then
         If MsgBox("This will delete users from the list", vbYesNo, "Confirmation") = vbYes Then
        users.Delete
        Else
        End If
        Exit Sub
        Else
        users.MoveNext
        End If
        Wend
        '*******************************************
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Activate()

Combo1.Clear
i = 0
'***open table users*********
Set users = New ADODB.Recordset
'** open table users
users.Open "select * from users", con, adOpenDynamic, adLockOptimistic
'******************
users.MoveFirst
j = users.RecordCount
users.MoveFirst
While users.EOF = False
    For i = 0 To i = users.RecordCount - 1
        Combo1.AddItem (users.Fields(i))
        i = i + 1
        users.MoveNext
    Next i
        
Wend
frmusermanager.Width = 6285
frmusermanager.Height = 3990
frmusermanager.ZOrder

      
End Sub

Private Sub Form_Load()
Frame1.Enabled = False
Command1.Enabled = False
End Sub



