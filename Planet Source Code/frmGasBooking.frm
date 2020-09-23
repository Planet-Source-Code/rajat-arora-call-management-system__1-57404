VERSION 5.00
Begin VB.Form frmCompBooking 
   Caption         =   "Complaint Booking Form"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmGasBooking.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame fraServiceDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "   Service  Details  "
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
      Height          =   4095
      Left            =   480
      TabIndex        =   34
      Top             =   2640
      Width           =   5295
      Begin VB.ComboBox cboPriority 
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
         ItemData        =   "frmGasBooking.frx":27A2
         Left            =   1320
         List            =   "frmGasBooking.frx":27AF
         TabIndex        =   38
         Text            =   "N"
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox txtSR1 
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
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Width           =   4575
      End
      Begin VB.TextBox txtSR2 
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
         Left            =   360
         TabIndex        =   4
         Top             =   1560
         Width           =   4575
      End
      Begin VB.TextBox txtSR3 
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
         Left            =   360
         MaxLength       =   12
         TabIndex        =   5
         Top             =   2040
         Width           =   4575
      End
      Begin VB.TextBox txtPrtText 
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
         Height          =   360
         Left            =   2160
         MaxLength       =   255
         TabIndex        =   6
         Text            =   "NORMAL"
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label lblServiceRequest 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Service Request:"
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
         Left            =   420
         TabIndex        =   36
         Top             =   600
         Width           =   1545
      End
      Begin VB.Label lblPriority 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Priority: "
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
         Left            =   540
         TabIndex        =   35
         Top             =   2760
         Width           =   705
      End
   End
   Begin VB.Frame fraCustomer 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "   Customer Details  "
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
      Height          =   4095
      Left            =   5880
      TabIndex        =   15
      Top             =   2640
      Width           =   5295
      Begin VB.TextBox txtPincode 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   24
         Top             =   3360
         Width           =   2295
      End
      Begin VB.TextBox txtCity 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         MaxLength       =   12
         TabIndex        =   23
         Top             =   3000
         Width           =   2295
      End
      Begin VB.TextBox txtLandmark 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         MaxLength       =   255
         TabIndex        =   22
         Top             =   2640
         Width           =   2295
      End
      Begin VB.TextBox txtLocality 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         MaxLength       =   12
         TabIndex        =   21
         Top             =   2280
         Width           =   2295
      End
      Begin VB.TextBox txtAddress2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         MaxLength       =   255
         TabIndex        =   20
         Top             =   1920
         Width           =   2295
      End
      Begin VB.TextBox txtAddress1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         MaxLength       =   255
         TabIndex        =   19
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox txtHouseNo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         MaxLength       =   12
         TabIndex        =   18
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtLastName 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         TabIndex        =   17
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txtFirstName 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         TabIndex        =   16
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblPincode 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PinCode :"
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
         Index           =   4
         Left            =   720
         TabIndex        =   33
         Top             =   3360
         Width           =   870
      End
      Begin VB.Label lblCity 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "City :"
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
         Index           =   3
         Left            =   720
         TabIndex        =   32
         Top             =   3000
         Width           =   420
      End
      Begin VB.Label lblLandmark 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Landmark :"
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
         Left            =   720
         TabIndex        =   31
         Top             =   2640
         Width           =   990
      End
      Begin VB.Label lblLocality 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Locality :"
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
         Left            =   720
         TabIndex        =   30
         Top             =   2280
         Width           =   780
      End
      Begin VB.Label lblAddress2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address 2 :"
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
         Left            =   720
         TabIndex        =   29
         Top             =   1920
         Width           =   1005
      End
      Begin VB.Label lblAddress1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address 1:"
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
         Left            =   720
         TabIndex        =   28
         Top             =   1560
         Width           =   960
      End
      Begin VB.Label lblHouseno 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "House No: "
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
         Left            =   720
         TabIndex        =   27
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label lblLastName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name :"
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
         Left            =   720
         TabIndex        =   26
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label lblFirstName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name :"
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
         Left            =   720
         TabIndex        =   25
         Top             =   480
         Width           =   1065
      End
   End
   Begin VB.Frame fraCallDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "  Call Details  "
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   10695
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
         Height          =   345
         Left            =   7365
         MaxLength       =   12
         TabIndex        =   2
         Top             =   960
         Width           =   1695
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
         Left            =   1665
         TabIndex        =   1
         Top             =   960
         Width           =   1335
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
         Height          =   345
         Left            =   7365
         TabIndex        =   11
         Text            =   "Telephone"
         Top             =   360
         Width           =   2175
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
         Height          =   360
         Left            =   4485
         TabIndex        =   9
         Text            =   "95120"
         Top             =   360
         Width           =   1215
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
         Height          =   345
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   7
         Text            =   "Z1"
         Top             =   360
         Width           =   375
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
         Left            =   5970
         TabIndex        =   14
         Top             =   960
         Width           =   1260
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
         Left            =   870
         TabIndex        =   13
         Top             =   960
         Width           =   630
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
         Left            =   6210
         TabIndex        =   12
         Top             =   360
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
         Left            =   2460
         TabIndex        =   10
         Top             =   360
         Width           =   1830
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
         Left            =   570
         TabIndex        =   8
         Top             =   360
         Width           =   915
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Complaint Booking Form"
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
      TabIndex        =   37
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmCompBooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cust_no As Double
Public sr1 As String
Public sr2 As String
Public sr3 As String
Public prt As String
Public taker As String

Private Sub cboPriority_Click()
'MsgBox cboPriority.Text
Select Case cboPriority.Text
Case "L"
txtPrtText.Text = "LOW"
Case "N"
txtPrtText.Text = "NORMAL"
Case "H"
txtPrtText.Text = "HIGH"
End Select
End Sub

Private Sub Form_Load()
    If mdDataConnection.searchCust = "yes" Then
    'MsgBox "Searched Customer"
    fill_details
    Else
    'MsgBox "New Complaint"
    End If
    txtTaker.Text = mdDataConnection.us
    txtTaker.Enabled = False
End Sub

Public Sub fill_details()
    txtAddress1.Text = mdDataConnection.cust.add1
    txtAddress2.Text = mdDataConnection.cust.add2
    txtCity.Text = mdDataConnection.cust.city
    txtCustNo.Text = mdDataConnection.cust.cust_no
    txtFirstName = mdDataConnection.cust.fname
    txtHouseNo.Text = mdDataConnection.cust.houseno
    txtLandmark = mdDataConnection.cust.landmark
    txtLastName = mdDataConnection.cust.lname
    txtLocality = mdDataConnection.cust.locality
    txtPincode = mdDataConnection.cust.pincode
End Sub


Public Sub search_data()
        Unload Me
        frmCustSearch.Show (1)
        frmCustSearch.Move 2000, 2000, 7200, 6000
        frmCustSearch.ZOrder
End Sub

Public Sub add_data()
    txtCustNo = ""
    txtAddress1.Text = ""
    txtAddress2.Text = ""
    txtCity.Text = ""
    txtCustNo.Text = ""
    txtFirstName = ""
    txtHouseNo.Text = ""
    txtLandmark = ""
    txtLastName = ""
    txtLocality = ""
    txtPincode = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
mdDataConnection.searchCust = "no"
End Sub

Public Sub save_data()
    
    If Not txtCustNo.Text = "" Then
    cust_no = txtCustNo.Text
    sr1 = txtSR1.Text
    sr2 = txtSR2.Text
    sr3 = txtSR3.Text
    prt = cboPriority.Text
    taker = txtTaker.Text
    Unload Me
    Load frmNotification
    frmNotification.Show
    Else
    MsgBox "Customer Not Selected"
    End If
End Sub
