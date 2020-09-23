VERSION 5.00
Begin VB.Form frmNotification 
   Caption         =   "Notification Creation Form"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11640
   Icon            =   "frmNotification.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8055
   ScaleWidth      =   11640
   WindowState     =   2  'Maximized
   Begin VB.Frame fraProcessing 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "  Processing  "
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   480
      TabIndex        =   28
      Top             =   5040
      Width           =   10695
      Begin VB.TextBox txtDate 
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
         Height          =   330
         Left            =   8445
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox chkBreakdown 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Breakdown : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   870
         Width           =   1575
      End
      Begin VB.TextBox txtPriority 
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
         TabIndex        =   7
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtPriorityText 
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
         Left            =   2625
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtInformation 
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
         TabIndex        =   11
         Top             =   1320
         Width           =   4335
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
         Index           =   0
         Left            =   7755
         TabIndex        =   32
         Top             =   360
         Width           =   525
      End
      Begin VB.Label lblPriority 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Priority :"
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
         Left            =   1005
         TabIndex        =   31
         Top             =   360
         Width           =   705
      End
      Begin VB.Label lblh 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   2400
         TabIndex        =   30
         Top             =   480
         Width           =   150
      End
      Begin VB.Label lblInformation 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Information :"
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
         Left            =   645
         TabIndex        =   29
         Top             =   1320
         Width           =   1065
      End
   End
   Begin VB.Frame fraDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "  Description  "
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   480
      TabIndex        =   20
      Top             =   2880
      Width           =   10695
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
         Height          =   330
         Left            =   1905
         TabIndex        =   6
         Top             =   1320
         Width           =   4335
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
         TabIndex        =   5
         Top             =   840
         Width           =   2295
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
         Left            =   8445
         TabIndex        =   22
         Text            =   "95120"
         Top             =   360
         Width           =   1215
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
         Left            =   8445
         TabIndex        =   21
         Text            =   "Telephone"
         Top             =   840
         Width           =   1935
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
         TabIndex        =   4
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblDetails 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Details :"
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
         Left            =   990
         TabIndex        =   27
         Top             =   1320
         Width           =   720
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
         TabIndex        =   26
         Top             =   840
         Width           =   555
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
         Left            =   6480
         TabIndex        =   25
         Top             =   360
         Width           =   1830
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
         Left            =   7335
         TabIndex        =   24
         Top             =   840
         Width           =   960
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
         TabIndex        =   23
         Top             =   360
         Width           =   630
      End
   End
   Begin VB.Frame fraNotification 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "  Notification Details  "
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   10695
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
         Height          =   315
         Left            =   1905
         TabIndex        =   3
         Top             =   1440
         Width           =   2295
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
         Height          =   330
         Left            =   1905
         TabIndex        =   2
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox txtNotificationNo 
         Alignment       =   1  'Right Justify
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
         Left            =   1905
         TabIndex        =   1
         Top             =   480
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
         Left            =   6240
         MaxLength       =   4
         TabIndex        =   13
         Text            =   "Z1"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtCustNo 
         Alignment       =   1  'Right Justify
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
         Left            =   8445
         MaxLength       =   12
         TabIndex        =   12
         Top             =   480
         Width           =   1695
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
         Left            =   885
         TabIndex        =   19
         Top             =   1440
         Width           =   780
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
         Index           =   1
         Left            =   1020
         TabIndex        =   17
         Top             =   960
         Width           =   645
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
         TabIndex        =   16
         Top             =   480
         Width           =   1350
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
         Left            =   5280
         TabIndex        =   15
         Top             =   480
         Width           =   915
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
         Left            =   7050
         TabIndex        =   14
         Top             =   480
         Width           =   1260
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Notification Details Form"
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
      Left            =   4320
      TabIndex        =   18
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmNotification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset

Public Sub fill_details()
    rs.MoveLast
    txtNotificationNo.Text = rs("notification_no") + 1
    txtCustNo.Text = mdDataConnection.cust.cust_no
    'txtCallRecFrom.Text = "Telephone"
    txtCallType.Text = "Z1"
    txtDate.Text = Format(Now(), "DD-MM-YYYY")
    txtDetails.Text = frmCompBooking.sr1 + " - " + frmCompBooking.sr2 + " - " + frmCompBooking.sr3
    'txtMode.Text = "95120"
    'txtNotificationNo.Text = "NEW"
    txtPriority.Text = frmCompBooking.prt
    txtProduct = mdDataConnection.cust.appliance
    txtZone.Text = mdDataConnection.cust.zone
    txtTaker.Text = frmCompBooking.taker
    txtStatus.Text = "OPEN"
    Select Case txtPriority.Text
    Case "N"
    txtPriorityText = "NORMAL"
    Case "L"
    txtPriorityText = "LOW"
    Case "H"
    txtPriorityText = "HIGH"
    End Select
End Sub

Private Sub chkBreakdown_Click()
    Select Case chkBreakdown.Value
    Case 1
    txtInformation.Enabled = True
    txtInformation.SetFocus
    txtInformation.Text = ""
    Case 0
    txtInformation.Enabled = False
    txtInformation.Text = 0
    End Select

End Sub

Private Sub Form_Load()
    mdDataConnection.data_connection
    'Dim clsEmpVal As New clsEmployee
    Set rs = New ADODB.Recordset
    rs.Open "select * from call_registration", mdDataConnection.con, adOpenDynamic, adLockOptimistic
    fill_details
End Sub

Public Sub save_data()
    
    rs.AddNew
    rs("notification_no") = txtNotificationNo.Text
    rs("cust_no") = txtCustNo.Text
    rs("name") = mdDataConnection.cust.fname + " " + mdDataConnection.cust.lname
    rs("status") = txtStatus.Text
    rs("date_reg") = txtDate.Text
    rs("product") = txtProduct.Text
    rs("call_type") = txtCallType.Text
    rs("taker") = txtTaker.Text
    rs("service_request") = txtDetails.Text
    rs("priority") = txtPriority.Text
    rs("bd") = chkBreakdown.Value
    rs("bd_text") = txtInformation.Text
    rs("zone") = txtZone.Text
    
    rs.Update
    
    MsgBox "Call Registered Successfully"
    clear_all
End Sub

Public Sub clear_all()
For Each Control In frmNotification
    If TypeOf Control Is TextBox Then
    Control.Text = ""
    End If
Next Control
    txtCallRecFrom = "95120"
    txtCallType = "Z1"
    txtMode = "Telephone"
End Sub
