VERSION 5.00
Begin VB.Form frmCallClose 
   Caption         =   "Call Close Form"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmCallClose.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame fraAdminNotes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "  Administrator Notes  "
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   480
      TabIndex        =   28
      Top             =   5400
      Width           =   10815
      Begin VB.TextBox txtAttendBy 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8685
         MaxLength       =   12
         TabIndex        =   13
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtDate 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8685
         MaxLength       =   12
         TabIndex        =   11
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtAdminDetails 
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
         Left            =   1875
         TabIndex        =   12
         Top             =   960
         Width           =   4335
      End
      Begin VB.TextBox txtStatus 
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
         Left            =   1875
         MaxLength       =   12
         TabIndex        =   10
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label lblAttendBy 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Attended By :"
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
         Left            =   7395
         TabIndex        =   34
         Top             =   960
         Width           =   1185
      End
      Begin VB.Label lblDate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date :"
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
         Index           =   2
         Left            =   8055
         TabIndex        =   31
         Top             =   480
         Width           =   525
      End
      Begin VB.Label lblDetails 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks :"
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
         Index           =   1
         Left            =   795
         TabIndex        =   30
         Top             =   960
         Width           =   915
      End
      Begin VB.Label lblStatus 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
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
         Index           =   2
         Left            =   1095
         TabIndex        =   29
         Top             =   480
         Width           =   645
      End
   End
   Begin VB.Frame fraNotification 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "  Notification Details  "
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   480
      TabIndex        =   15
      Top             =   600
      Width           =   10815
      Begin VB.CommandButton cmdGo 
         Caption         =   "GO"
         Height          =   375
         Left            =   3600
         TabIndex        =   2
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtMode 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8685
         TabIndex        =   25
         Text            =   "Telephone"
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtCallRecFrom 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8685
         TabIndex        =   24
         Text            =   "95120"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtTaker 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1905
         TabIndex        =   3
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtZone 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1905
         TabIndex        =   4
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtCallType 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6120
         MaxLength       =   4
         TabIndex        =   16
         Text            =   "Z1"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtNotificationNo 
         Alignment       =   1  'Right Justify
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
         Left            =   1905
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblMode 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Call Mode:"
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
         Index           =   1
         Left            =   7575
         TabIndex        =   27
         Top             =   960
         Width           =   960
      End
      Begin VB.Label lblCallRecFrom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Call Received From:"
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
         Index           =   1
         Left            =   6720
         TabIndex        =   26
         Top             =   480
         Width           =   1830
      End
      Begin VB.Label lblTaker 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Taker :"
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
         Left            =   1080
         TabIndex        =   21
         Top             =   960
         Width           =   630
      End
      Begin VB.Label lblZone 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zone :"
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
         Index           =   1
         Left            =   1155
         TabIndex        =   20
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label lblCallType 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Call Type:"
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
         Left            =   5160
         TabIndex        =   18
         Top             =   480
         Width           =   915
      End
      Begin VB.Label lblNotification 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notification No:"
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
         Index           =   1
         Left            =   315
         TabIndex        =   17
         Top             =   480
         Width           =   1350
      End
   End
   Begin VB.Frame fraCustDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "  Customer Details  "
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   480
      TabIndex        =   0
      Top             =   2760
      Width           =   10815
      Begin VB.TextBox txtregdate 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8670
         MaxLength       =   12
         TabIndex        =   35
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtSignBy 
         Alignment       =   1  'Right Justify
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
         Left            =   8685
         MaxLength       =   12
         TabIndex        =   9
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtCustName 
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
         Left            =   1875
         MaxLength       =   12
         TabIndex        =   6
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox txtCustNo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1875
         MaxLength       =   12
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtProduct 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1875
         TabIndex        =   7
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtDetails 
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
         Left            =   1875
         TabIndex        =   8
         Top             =   1800
         Width           =   4335
      End
      Begin VB.Label lblDate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reg Date :"
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
         Left            =   7605
         TabIndex        =   36
         Top             =   1200
         Width           =   960
      End
      Begin VB.Label lblSignBy 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Signed By :"
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
         Index           =   2
         Left            =   7560
         TabIndex        =   33
         Top             =   1800
         Width           =   1020
      End
      Begin VB.Label lblCustNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name :"
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
         Index           =   1
         Left            =   195
         TabIndex        =   23
         Top             =   840
         Width           =   1545
      End
      Begin VB.Label lblCustNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer No :"
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
         Left            =   480
         TabIndex        =   22
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label lblProduct 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product :"
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
         Left            =   840
         TabIndex        =   19
         Top             =   1320
         Width           =   780
      End
      Begin VB.Label lblDetails 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comp Details :"
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
         Left            =   405
         TabIndex        =   14
         Top             =   1800
         Width           =   1305
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Check / Close Complaint "
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
      TabIndex        =   32
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmCallClose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset

Private Sub cmdGo_Click()
rs.MoveFirst
    Do While Not rs.EOF
    If rs("notification_no") = txtNotificationNo.Text Then
    fill_details
    Exit Sub
    Else
    rs.MoveNext
    End If
    Loop
    clear_all
End Sub

Private Sub Form_Load()
    mdDataConnection.data_connection
    Set rs = New ADODB.Recordset
    rs.Open "select * from call_registration", mdDataConnection.con, adOpenDynamic, adLockOptimistic
    clear_all
End Sub

Public Sub fill_details()
    txtCustNo.Text = rs("cust_no")
    txtCustName = rs("name")
    txtStatus.Text = rs("status")
    txtregdate.Text = rs("date_reg")
    txtProduct.Text = rs("product")
    txtCallType.Text = rs("call_type")
    txtTaker.Text = rs("taker")
    txtDetails.Text = rs("service_request")
    txtZone.Text = rs("zone")
    txtDate.Text = Format(Now(), "DD-MM-YYYY")
End Sub

Public Sub clear_all()
    For Each Control In frmCallClose
        If TypeOf Control Is TextBox Then
        Control.Text = ""
        End If
    Next Control
    txtCallRecFrom = "95120"
    txtCallType = "Z1"
    txtMode = "Telephone"
End Sub

Public Sub save_data()
rs("remarks") = txtDetails.Text
rs("sign_by") = txtSignBy.Text
rs("date_close") = txtDate.Text
rs("attend_by") = txtAttendBy.Text
rs("status") = txtStatus.Text
rs.Update
MsgBox "Class Closed Successfully"
clear_all
End Sub
