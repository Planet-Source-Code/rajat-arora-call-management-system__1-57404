VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCustomerReg 
   Caption         =   "Customer Registration Form"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmCustomerReg.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmpComp 
      Caption         =   "Register Complaint"
      Height          =   375
      Left            =   3360
      TabIndex        =   42
      Top             =   480
      Width           =   1695
   End
   Begin VB.ComboBox cboCust_Type 
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
      ItemData        =   "frmCustomerReg.frx":27A2
      Left            =   4920
      List            =   "frmCustomerReg.frx":27B2
      TabIndex        =   15
      Text            =   "    ---- Select ----"
      Top             =   4200
      Width           =   2295
   End
   Begin VB.TextBox txtCustomerNo 
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
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.Frame fraServiceProvider 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "   Service Provider Details  "
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
      Height          =   1215
      Left            =   360
      TabIndex        =   37
      Top             =   5280
      Width           =   11055
      Begin VB.TextBox txtspCode 
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
         Left            =   8400
         MaxLength       =   3
         TabIndex        =   19
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtPlant 
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
         Left            =   1320
         MaxLength       =   150
         TabIndex        =   17
         Text            =   "4510"
         Top             =   480
         Width           =   735
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
         Height          =   360
         Left            =   4560
         MaxLength       =   255
         TabIndex        =   18
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblPlant 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plant :"
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
         Left            =   585
         TabIndex        =   40
         Top             =   480
         Width           =   540
      End
      Begin VB.Label lblZone 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Zone :"
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
         Left            =   3015
         TabIndex        =   39
         Top             =   480
         Width           =   1350
      End
      Begin VB.Label lblSp 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S.P Code :"
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
         Left            =   7260
         TabIndex        =   38
         Top             =   480
         Width           =   945
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
      Height          =   2535
      Left            =   360
      TabIndex        =   23
      Top             =   960
      Width           =   11055
      Begin VB.TextBox txtTitle 
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
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtFirstName 
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
         Left            =   4560
         TabIndex        =   3
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtLastName 
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
         Left            =   8400
         TabIndex        =   4
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtHouseNo 
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
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   5
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtAddress1 
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
         Left            =   4560
         MaxLength       =   255
         TabIndex        =   6
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtAddress2 
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
         Left            =   8400
         MaxLength       =   255
         TabIndex        =   7
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtLocality 
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
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   8
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtLandmark 
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
         Left            =   4560
         MaxLength       =   255
         TabIndex        =   9
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox txtCity 
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
         Left            =   8400
         MaxLength       =   12
         TabIndex        =   10
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox txtPincode 
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
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   11
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtPhoneNo 
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
         Left            =   4560
         MaxLength       =   12
         TabIndex        =   12
         Top             =   1920
         Width           =   2295
      End
      Begin VB.TextBox txtAddProof 
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
         Left            =   8400
         MaxLength       =   255
         TabIndex        =   13
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title : "
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
         TabIndex        =   35
         Top             =   480
         Width           =   645
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
         Left            =   3300
         TabIndex        =   34
         Top             =   480
         Width           =   1065
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
         Left            =   7140
         TabIndex        =   33
         Top             =   480
         Width           =   1065
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
         Left            =   120
         TabIndex        =   32
         Top             =   960
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
         Left            =   3405
         TabIndex        =   31
         Top             =   960
         Width           =   960
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
         Left            =   7200
         TabIndex        =   30
         Top             =   960
         Width           =   1005
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
         Left            =   345
         TabIndex        =   29
         Top             =   1440
         Width           =   780
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
         Left            =   3375
         TabIndex        =   28
         Top             =   1440
         Width           =   990
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
         Left            =   7785
         TabIndex        =   27
         Top             =   1440
         Width           =   420
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
         Left            =   255
         TabIndex        =   26
         Top             =   1920
         Width           =   870
      End
      Begin VB.Label lblPhoneNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No :"
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
         Index           =   5
         Left            =   3375
         TabIndex        =   25
         Top             =   1920
         Width           =   990
      End
      Begin VB.Label lblAddProof 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add Proof :"
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
         Index           =   6
         Left            =   7215
         TabIndex        =   24
         Top             =   1920
         Width           =   990
      End
   End
   Begin VB.Frame fraProduct 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "   Product  Details  "
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
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   3720
      Width           =   11055
      Begin VB.TextBox txtAppliance 
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
         Left            =   1320
         MaxLength       =   150
         TabIndex        =   14
         Top             =   480
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dtpDop 
         Height          =   375
         Left            =   8400
         TabIndex        =   16
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "  dd - MM - yyyy"
         Format          =   24641539
         CurrentDate     =   38127
      End
      Begin VB.Label lblDateofPurchase 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Pur. :"
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
         Left            =   7080
         TabIndex        =   22
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label lblType 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cust. Type :"
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
         Left            =   3315
         TabIndex        =   21
         Top             =   480
         Width           =   1050
      End
      Begin VB.Label lblProduct 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Applicace :"
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
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   1005
      End
   End
   Begin VB.Label lblCustomerNo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Cust No :  "
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
      Left            =   585
      TabIndex        =   41
      Top             =   600
      Width           =   930
   End
   Begin VB.Label lblMaintitle 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Customer Registration"
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
      TabIndex        =   36
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmCustomerReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nstat As String
Dim rs As ADODB.Recordset
Dim search As Double
'Dim cust As clsCustomer

Private Sub cmpComp_Click()
If txtCustomerNo.Text = "" Then
    MsgBox "Please Select valid Custoemr"
Else
'    MsgBox "Valid Customer"
    'Set cust = New clsCustomer
    mdDataConnection.searchCust = "yes"
    
    mdDataConnection.cust.add1 = rs("add1")
    mdDataConnection.cust.add2 = rs("add2")
    mdDataConnection.cust.addproof = rs("addproof")
    mdDataConnection.cust.appliance = rs("appliance")
    mdDataConnection.cust.city = rs("city")
    'cust.cust_name  =
    mdDataConnection.cust.cust_no = rs("cust_no")
    mdDataConnection.cust.cust_type = rs("cust_type")
    mdDataConnection.cust.dop = rs("dop")
    mdDataConnection.cust.fname = rs("fname")
    mdDataConnection.cust.houseno = rs("houseno")
    mdDataConnection.cust.landmark = rs("landmark")
    mdDataConnection.cust.lname = rs("lname")
    mdDataConnection.cust.locality = rs("locality")
    mdDataConnection.cust.phone = rs("phone")
    mdDataConnection.cust.pincode = rs("pincode")
    mdDataConnection.cust.plant = rs("plant")
    mdDataConnection.cust.sp = rs("sp_code")
    mdDataConnection.cust.zone = rs("d_zone")
    
    Load frmCompBooking
    frmCompBooking.Show
    Unload Me

End If

End Sub

Private Sub Form_Activate()
    Me.KeyPreview = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys (vbTab)
    End If
End Sub

Private Sub Form_Load()
    mdDataConnection.data_connection
    'Dim clsEmpVal As New clsEmployee
    Set rs = New ADODB.Recordset
    Set rsdept = New ADODB.Recordset
    'rs.Open "select empid,fname,lname,department from employee_details", mdDataConnection.con, adOpenDynamic, adLockOptimistic
    rs.Open "select * from customer_details", mdDataConnection.con, adOpenDynamic, adLockPessimistic
    'rsdept.Open "select * from department", mdDataConnection.con, adOpenDynamic, adLockPessimistic
    
    If mdDataConnection.searchStatus = "yes" Then
    'MsgBox "search details"
    search_CustDetails
    Else
    'MsgBox "new user"
    End If
    
    dtpDop.Value = Now()
End Sub
Public Sub clear_all()
    For Each Control In frmCustomerReg
        If TypeOf Control Is TextBox Then
            Control.Text = ""
        End If
    Next Control
    txtPlant.Text = "4510"
    txtPlant.Enabled = False
End Sub
Public Sub add_data()
        clear_all
        Enable_all
        'cboDept.Enabled = True
        'fraDept.Enabled = False
        'txtEmpCode.Enabled = False
        txtCustomerNo.Enabled = False
        rs.MoveLast
        txtCustomerNo.Text = rs("cust_no") + 1
        txtTitle.SetFocus
        nstat = "add"
End Sub

Public Sub save_data()
        
        
            strmsg = ""
        '    validation_rules
            If strmsg = "" Then
               If nstat = "add" Then
                    rs.AddNew
                    rs("cust_no") = txtCustomerNo.Text
                    rs("title") = txtTitle.Text
                    rs("fname") = txtFirstName.Text
                    rs("lname") = txtLastName.Text
                    rs("houseno") = txtHouseNo.Text
                    rs("add1") = txtAddress1.Text
                    rs("add2") = txtAddress2.Text
                    rs("locality") = txtLocality.Text
                    rs("landmark") = txtLandmark.Text
                    rs("city") = txtCity.Text
                    rs("pincode") = txtPincode.Text
                    rs("phone") = txtPhoneNo.Text
                    rs("addproof") = txtAddProof.Text
                    'rs("designation") = txtDesignation.Text
                    'rs("salary") = txtSalary.Text
                    rs("appliance") = txtAppliance.Text
                    rs("cust_type") = cboCust_Type.Text
                    rs("dop") = dtpDop.Value
                    rs("plant") = txtPlant.Text
                    rs("d_zone") = txtZone.Text
                    rs("sp_code") = txtspCode.Text
                    rs.Update
                    
                   MsgBox "....Record Saved Succussfully...."
                   ' clear_all
                   ' disable_all
                   ' fraDept.Enabled = True
                   ' refresh_data
                ElseIf nstat = "modify" Then
                    rs("cust_no") = txtCustomerNo.Text
                    rs("title") = txtTitle.Text
                    rs("fname") = txtFirstName.Text
                    rs("lname") = txtLastName.Text
                    rs("houseno") = txtHouseNo.Text
                    rs("add1") = txtAddress1.Text
                    rs("add2") = txtAddress2.Text
                    rs("locality") = txtLocality.Text
                    rs("landmark") = txtLandmark.Text
                    rs("city") = txtCity.Text
                    rs("pincode") = txtPincode.Text
                    rs("phone") = txtPhoneNo.Text
                    rs("addproof") = txtAddProof.Text
                    'rs("designation") = txtDesignation.Text
                    'rs("salary") = txtSalary.Text
                    rs("appliance") = txtAppliance.Text
                    rs("cust_type") = cboCust_Type.Text
                    rs("dop") = dtpDop.Value
                    rs("plant") = txtPlant.Text
                    rs("d_zone") = txtZone.Text
                    rs("sp_code") = txtspCode.Text
                    rs.Update
                    'disable_all
                    'refresh_data
                    MsgBox "Record Updated Successfully"
                    Disable_all
                End If
                
               Else
        '            MsgBox strmsg
        '            txtFirstName.SetFocus
               End If
            'Else
                    MsgBox "Nothing Selected"
            'End If
End Sub


Public Sub fill_CustDetails()
MsgBox "fill customer details"
    txtCustomerNo.Text = rs("cust_no")
    txtTitle.Text = rs("title")
    txtFirstName.Text = rs("fname")
    txtLastName.Text = rs("lname")
    txtHouseNo.Text = rs("houseno")
    txtAddress1.Text = rs("add1")
    txtAddress2.Text = rs("add2")
    txtLocality.Text = rs("locality")
    txtLandmark.Text = rs("landmark")
    txtCity.Text = rs("city")
    txtPincode.Text = rs("pincode")
    txtPhoneNo.Text = rs("phone")
    txtAddProof.Text = rs("addproof")
    'txtDesignation.Text =   'rs("designation")
    'txtSalary.Text  =   'rs("salary")
    txtAppliance.Text = rs("appliance")
    cboCust_Type.Text = rs("cust_type")
    dtpDop.Value = rs("dop")
    txtPlant.Text = rs("plant")
    txtZone.Text = rs("d_zone")
    txtspCode.Text = rs("sp_code")
End Sub

Public Sub search_CustDetails()
search = Trim(mdDataConnection.cust_no)
rs.MoveFirst

Do While Not rs.EOF

    If rs("cust_no") = search Then
    fill_CustDetails
    Disable_all
    Exit Do
    Else
    rs.MoveNext
    End If
Loop

End Sub

Private Sub Form_Unload(Cancel As Integer)
mdDataConnection.searchStatus = ""
mdDataConnection.cust_no = ""
End Sub
Public Sub Enable_all()
    For Each Control In frmCustomerReg
            If TypeOf Control Is TextBox Then
                Control.Enabled = True
            End If
        Next Control
        txtPlant.Enabled = False
        dtpDop.Enabled = True
        cboCust_Type.Enabled = True
End Sub

Public Sub Disable_all()
    For Each Control In frmCustomerReg
            If TypeOf Control Is TextBox Then
                Control.Enabled = False
            End If
        Next Control
        txtPlant.Enabled = False
        dtpDop.Enabled = False
        cboCust_Type.Enabled = False
End Sub
Public Sub modify_data()
        If txtCustomerNo.Text = "" Then
            MsgBox "Nothing Selected"
        Else
        Enable_all
        txtCustomerNo.Enabled = False
        nstat = "modify"
        End If
End Sub


Public Sub search_data()
        Unload Me
        frmCustSearch.Show (1)
        frmCustSearch.Move 2000, 2000, 7200, 6000
        frmCustSearch.ZOrder
End Sub

Public Sub cancel_data()
        clear_all
        Disable_all
End Sub

Public Sub delete_data()

        If txtCustomerNo.Text = "" Or rs("cust_no") = "" Or rs.EOF Or rs.BOF Then
            MsgBox "Nothing Selected"
        ElseIf rs("cust_no") = 12346 Then
        MsgBox "Test Record Can not be Deleted"
        Else
            rs.Delete
            clear_all
            Disable_all
        '    refresh_data
        End If
End Sub

