VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "User Login"
   ClientHeight    =   1515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Swtich"
      Default         =   -1  'True
      Height          =   375
      Left            =   4320
      Picture         =   "frmLogin.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "ô"
      TabIndex        =   1
      Top             =   720
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "User ID"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "User Name"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   795
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public users As ADODB.Recordset
Private Sub Command1_Click()
users.MoveFirst
While users.EOF = False
    If Text1.Text = users.Fields(0) And Text2.Text = users.Fields(1) Then
        mdDataConnection.us = Text1.Text
        Unload Me
        Load MdiParent
        MdiParent.Show
        'Load employee
        'employee.Show

        
 Exit Sub
 Else
    users.MoveNext
    End If
Wend
MsgBox "Invalid user name/ID", vbCritical, "CCMS"
Text1 = ""
Text2 = ""
Text1.SetFocus

End Sub

Private Sub Form_Activate()
Call data_connection
'***open table users*********
Set users = New ADODB.Recordset
'** open table users
users.Open "select * from users", mdDataConnection.con, adOpenDynamic, adLockOptimistic

'******************
End Sub




