VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCheckStatus 
   Caption         =   "Check Status -- Customer NO"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmCheckStatus.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc adoSearch 
      Height          =   330
      Left            =   4920
      Top             =   6960
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Search"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame fraNotificationDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "   Notification Details  "
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
      Height          =   2775
      Left            =   480
      TabIndex        =   22
      Top             =   4080
      Width           =   11055
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmCheckStatus.frx":27A2
         Height          =   1815
         Left            =   480
         TabIndex        =   11
         Top             =   600
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   3201
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
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
      Height          =   2895
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   11055
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   375
         Left            =   4320
         TabIndex        =   2
         Top             =   425
         Width           =   855
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
         Left            =   2160
         MaxLength       =   12
         TabIndex        =   1
         Top             =   480
         Width           =   1695
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
         Height          =   315
         Left            =   6240
         MaxLength       =   12
         TabIndex        =   10
         Top             =   2400
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
         Height          =   300
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   9
         Top             =   2400
         Width           =   1095
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
         Height          =   330
         Left            =   6240
         MaxLength       =   12
         TabIndex        =   8
         Top             =   1920
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
         Height          =   300
         Left            =   2160
         MaxLength       =   12
         TabIndex        =   7
         Top             =   1920
         Width           =   1695
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
         Height          =   300
         Left            =   6240
         MaxLength       =   255
         TabIndex        =   6
         Top             =   1440
         Width           =   2295
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
         Height          =   315
         Left            =   6240
         MaxLength       =   255
         TabIndex        =   4
         Top             =   960
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
         Height          =   300
         Left            =   2160
         MaxLength       =   12
         TabIndex        =   5
         Top             =   1440
         Width           =   1695
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
         Height          =   300
         Left            =   2160
         TabIndex        =   3
         Top             =   960
         Width           =   2295
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
         Left            =   720
         TabIndex        =   21
         Top             =   480
         Width           =   1260
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
         Left            =   5055
         TabIndex        =   19
         Top             =   2400
         Width           =   990
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
         Left            =   1095
         TabIndex        =   18
         Top             =   2400
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
         Left            =   5625
         TabIndex        =   17
         Top             =   1920
         Width           =   420
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
         Left            =   1185
         TabIndex        =   16
         Top             =   1920
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
         Left            =   5040
         TabIndex        =   15
         Top             =   1440
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
         Left            =   5085
         TabIndex        =   14
         Top             =   960
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
         Left            =   960
         TabIndex        =   13
         Top             =   1440
         Width           =   1005
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
         Left            =   900
         TabIndex        =   12
         Top             =   960
         Width           =   1065
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Check Complaint Status"
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
      TabIndex        =   20
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmCheckStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rsnot As ADODB.Recordset

Private Sub cmdSearch_Click()
rs.MoveFirst
Do While Not rs.EOF
    If rs("cust_no") = txtCustNo.Text Then
    txtAddress1.Text = rs("add1")
    txtAddress2.Text = rs("add2")
    txtCity.Text = rs("city")
    txtFirstName.Text = rs("fname")
    txtHouseNo.Text = rs("houseno")
    txtLocality.Text = rs("locality")
    txtPincode.Text = rs("pincode")
    txtPhoneNo.Text = rs("phone")
    fill_grid
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
        Set rsnot = New ADODB.Recordset
        
        adoSearch.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=ccms"
        rs.Open "select * from customer_details", mdDataConnection.con, adOpenDynamic, adLockOptimistic
        rsnot.Open "select * from call_registration", mdDataConnection.con, adOpenDynamic, adLockOptimistic
        clear_all
End Sub

Private Sub txtPhone_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Not Trim(txtPhone.Text) = "" Then
    adoSearch.RecordSource = "select cust_no,appliance,fname,lname from customer_details where phone =" & Trim(txtPhone.Text)
    adoSearch.CommandType = adCmdText
    adoSearch.Refresh
    Else
    MsgBox "Nothing to Search" & vbCrLf & "Enter Valid Phone Number"
    End If
End If

End Sub

Public Sub fill_grid()
If Not Trim(txtCustNo.Text) = "" Then
    adoSearch.RecordSource = "select notification_no,date_reg,status,date_close,product,taker from call_registration where cust_no =" & Trim(txtCustNo.Text)
    adoSearch.CommandType = adCmdText
    adoSearch.Refresh
    Else
    MsgBox "Nothing to Search" & vbCrLf & "Enter Valid Phone Number"
End If
End Sub
Public Sub clear_all()
Dim r As Integer
Dim c As Integer
For Each Control In frmCheckStatus
    If TypeOf Control Is TextBox Then
    Control.Text = ""
    End If
Next Control

    adoSearch.RecordSource = "select notification_no,date_reg,status,date_close,product,taker from call_registration where cust_no = 0"
    adoSearch.CommandType = adCmdText
    adoSearch.Refresh

'MsgBox DataGrid1.Col
'MsgBox DataGrid1.Row
'MsgBox DataGrid1.Columns.Count
'MsgBox DataGrid1.VisibleRows
End Sub
