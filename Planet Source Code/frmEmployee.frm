VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEmployee 
   Caption         =   "Employes Information Form"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmEmployee.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8535
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc adoEmployee 
      Height          =   330
      Left            =   7920
      Top             =   6600
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Caption         =   "adoEmployee"
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
   Begin VB.Frame fraEmployee 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "   Employe's Details  "
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
      Height          =   6135
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   11055
      Begin VB.CommandButton cmdgo 
         Caption         =   "GO"
         Height          =   375
         Left            =   4080
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Frame fraDept 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   5655
         Left            =   5880
         TabIndex        =   28
         Top             =   240
         Width           =   4815
         Begin MSDataGridLib.DataGrid dgEmp 
            Bindings        =   "frmEmployee.frx":27A2
            Height          =   4335
            Left            =   0
            TabIndex        =   14
            Top             =   1080
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   7646
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   0   'False
            Appearance      =   0
            Enabled         =   -1  'True
            ColumnHeaders   =   -1  'True
            HeadLines       =   1
            RowHeight       =   17
            RowDividerStyle =   3
            FormatLocked    =   -1  'True
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
            Caption         =   "Employee Search"
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "emp_no"
               Caption         =   "Emp ID"
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
               DataField       =   "fname"
               Caption         =   "First Name"
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
            BeginProperty Column02 
               DataField       =   "lname"
               Caption         =   "Last Name"
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
                  Locked          =   -1  'True
                  ColumnWidth     =   915.024
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1739.906
               EndProperty
            EndProperty
         End
         Begin VB.ComboBox cboDepart 
            Height          =   315
            Left            =   2145
            TabIndex        =   13
            Text            =   " --- All ----"
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label lblDepartment 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Department :"
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
            Left            =   840
            TabIndex        =   29
            Top             =   360
            Width           =   1140
         End
      End
      Begin VB.ComboBox cboDept 
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
         ItemData        =   "frmEmployee.frx":27BC
         Left            =   2640
         List            =   "frmEmployee.frx":27BE
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   4440
         Width           =   2295
      End
      Begin VB.TextBox txtSalary 
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
         Height          =   360
         Left            =   2625
         MaxLength       =   12
         TabIndex        =   12
         Top             =   5400
         Width           =   1575
      End
      Begin VB.TextBox txtDesignation 
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
         Left            =   2625
         TabIndex        =   11
         Top             =   4920
         Width           =   2295
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
         Left            =   2640
         MaxLength       =   12
         TabIndex        =   9
         Top             =   3960
         Width           =   1335
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
         Left            =   2640
         MaxLength       =   6
         TabIndex        =   8
         Top             =   3480
         Width           =   1335
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
         Left            =   2640
         MaxLength       =   12
         TabIndex        =   7
         Top             =   3000
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
         Left            =   2640
         MaxLength       =   255
         TabIndex        =   6
         Top             =   2520
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
         Height          =   360
         Left            =   2640
         MaxLength       =   255
         TabIndex        =   5
         Top             =   2040
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
         Left            =   2640
         TabIndex        =   4
         Top             =   1560
         Width           =   2295
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
         Left            =   2640
         TabIndex        =   3
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtEmpCode 
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
         Left            =   2640
         MaxLength       =   8
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblDepartment 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department :"
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
         Left            =   1335
         TabIndex        =   26
         Top             =   4440
         Width           =   1140
      End
      Begin VB.Label lblRs 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rs./Mon"
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
         Left            =   4365
         TabIndex        =   25
         Top             =   5520
         Width           =   750
      End
      Begin VB.Label lblSalary 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salary :"
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
         Left            =   1755
         TabIndex        =   24
         Top             =   5400
         Width           =   675
      End
      Begin VB.Label lblDesignation 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Designation :"
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
         Left            =   1260
         TabIndex        =   23
         Top             =   4920
         Width           =   1170
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
         Left            =   1455
         TabIndex        =   22
         Top             =   3960
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
         Left            =   1575
         TabIndex        =   21
         Top             =   3480
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
         Left            =   2025
         TabIndex        =   20
         Top             =   3000
         Width           =   420
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
         Left            =   1440
         TabIndex        =   19
         Top             =   2520
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
         Left            =   1485
         TabIndex        =   18
         Top             =   2040
         Width           =   960
      End
      Begin VB.Label lblLastName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Emp Last Name :"
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
         Left            =   915
         TabIndex        =   17
         Top             =   1560
         Width           =   1530
      End
      Begin VB.Label lblFirstName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Emp First Name :"
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
         Left            =   915
         TabIndex        =   16
         Top             =   1080
         Width           =   1530
      End
      Begin VB.Label lblEmpCode 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Code :"
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
         TabIndex        =   15
         Top             =   600
         Width           =   1560
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Employee Details"
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
      TabIndex        =   27
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rsdept As ADODB.Recordset
Dim nstat As String
Dim strmsg As String


Private Sub cboDepart_Click()
    strdept = Trim(cboDepart.Text)
        If strdept = "All" Then
        strsql = "select emp_no,fname,lname from employee_details"
        Else
        strsql = "select emp_no,fname,lname from employee_details where department = '" & strdept & "'"
        End If
    adoEmployee.RecordSource = strsql
    adoEmployee.CommandType = adCmdText
    adoEmployee.Refresh
    clear_all
End Sub

Private Sub cmdGo_Click()
Dim nst As Integer
    nst = 0
    rs.MoveFirst
    While Not rs.EOF
        If rs("emp_no") = Val(Trim(txtEmpCode.Text)) Then
            fill_EmpDetails
            nst = 1
            Disable_all
            fraDept.Enabled = True
            Exit Sub
        Else
            rs.MoveNext
            nst = 0
        End If
    Wend
        If nst = 0 Then
        clear_all
        Disable_all
        End If
End Sub

Private Sub dgEmp_Click()
    Disable_all
    dgEmp.Col = 0
    Dim strEmpno As String
    strEmpno = Trim(dgEmp.Text)
    rs.MoveFirst
    While Not rs.EOF
        If rs("emp_no") = strEmpno Then
            fill_EmpDetails
            Exit Sub
        Else
            rs.MoveNext
        End If
    Wend
    
End Sub

Private Sub Form_Activate()
Call Disable_all
adoEmployee.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=ccms"
adoEmployee.RecordSource = "select * from employee_details"
adoEmployee.CommandType = adCmdText
adoEmployee.Refresh
End Sub

Private Sub Form_Load()
    mdDataConnection.data_connection
    Dim clsEmpVal As New clsEmployee
    Set rs = New ADODB.Recordset
    Set rsdept = New ADODB.Recordset
    'rs.Open "select empid,fname,lname,department from employee_details", mdDataConnection.con, adOpenDynamic, adLockOptimistic
    rs.Open "select * from employee_details", mdDataConnection.con, adOpenDynamic, adLockPessimistic
    rsdept.Open "select * from department", mdDataConnection.con, adOpenDynamic, adLockPessimistic

    fill_deptlist
End Sub

Public Sub fill_deptlist()
On Error GoTo np:
    rsdept.MoveFirst
    cboDepart.AddItem ("All")
    While Not rsdept.EOF
        cboDept.AddItem (rsdept(0))
        cboDepart.AddItem (rsdept(0))
        rsdept.MoveNext
    Wend
    Exit Sub
np:
    MsgBox "There are No Departments"
End Sub

Public Sub Disable_all()
    For Each Control In frmEmployee
        If TypeOf Control Is TextBox Then
            Control.Enabled = False
        End If
    Next Control
    cboDept.Enabled = False
    cmdGo.Visible = False
End Sub

Public Sub fill_EmpDetails()
    txtEmpCode.Text = rs("emp_no")
    txtFirstName.Text = rs("fname")
    txtLastName.Text = rs("lname")
    txtAddress1.Text = rs("add1")
    txtAddress2.Text = rs("add2")
    txtCity.Text = rs("city")
    txtPincode.Text = rs("pincode")
    txtPhoneNo.Text = rs("phone")
    cboDept.Text = rs("department")
    txtDesignation.Text = rs("designation")
    txtSalary.Text = rs("salary")
End Sub

Public Sub clear_all()
    For Each Control In frmEmployee
        If TypeOf Control Is TextBox Then
            Control.Text = ""
        End If
    Next Control
    cmdGo.Visible = False
'    cboDept.Text = ""
End Sub

Public Sub add_data()
        clear_all
        For Each Control In frmEmployee
            If TypeOf Control Is TextBox Then
                Control.Enabled = True
            End If
        Next Control
        cboDept.Enabled = True
        fraDept.Enabled = False
        txtEmpCode.Enabled = False
        rs.MoveLast
        txtEmpCode.Text = rs("Emp_no") + 1
        txtFirstName.SetFocus
        nstat = "add"
End Sub

Private Sub txtPhoneNo_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
        KeyAscii = 0
        End If
End Sub

Private Sub txtPincode_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
        KeyAscii = 0
        End If
End Sub

Public Sub save_data()
        If Not txtEmpCode.Text = "" Then
            strmsg = ""
            validation_rules
            If strmsg = "" Then
               If nstat = "add" Then
                    rs.AddNew
                    rs("emp_no") = txtEmpCode.Text
                    rs("fname") = txtFirstName.Text
                    rs("lname") = txtLastName.Text
                    rs("add1") = txtAddress1.Text
                    rs("add2") = txtAddress2.Text
                    rs("city") = txtCity.Text
                    rs("pincode") = txtPincode.Text
                    rs("phone") = txtPhoneNo.Text
                    rs("department") = cboDept.Text
                    rs("designation") = txtDesignation.Text
                    rs("salary") = txtSalary.Text
                    rs.Update
                    
                    clear_all
                    Disable_all
                    fraDept.Enabled = True
                    refresh_data
                ElseIf nstat = "modify" Then
                    rs("emp_no") = txtEmpCode.Text
                    rs("fname") = txtFirstName.Text
                    rs("lname") = txtLastName.Text
                    rs("add1") = txtAddress1.Text
                    rs("add2") = txtAddress2.Text
                    rs("city") = txtCity.Text
                    rs("pincode") = txtPincode.Text
                    rs("phone") = txtPhoneNo.Text
                    rs("department") = cboDept.Text
                    rs("designation") = txtDesignation.Text
                    rs("salary") = txtSalary.Text
                    rs.Update
                    
                    clear_all
                    Disable_all
                    fraDept.Enabled = True
                    refresh_data
                End If
                
               Else
                    MsgBox strmsg
                    txtFirstName.SetFocus
               End If
            Else
                    MsgBox "Nothing Selected"
            End If
End Sub

Public Sub validation_rules()
        'strmsg This variable holds error message
        
        If Trim(txtFirstName.Text) = "" Then
            If strmsg = "" Then
                strmsg = "Following Fields are Required " & vbCrLf
            End If
            strmsg = strmsg & Space(10) & "- " & "First Name" & vbCrLf
        End If
        
        If Trim(txtLastName.Text) = "" Then
            If strmsg = "" Then
                strmsg = "Following Fields are Required " & vbCrLf
            End If
            strmsg = strmsg & Space(10) & "- " & "Last Name" & vbCrLf
        End If
        
        If Trim(txtAddress1.Text) = "" Then
            If strmsg = "" Then
                strmsg = "Following Fields are Required " & vbCrLf
            End If
            strmsg = strmsg & Space(10) & "- " & "Address 1" & vbCrLf
        End If
        
        If Trim(txtAddress2.Text) = "" Then
            If strmsg = "" Then
                strmsg = "Following Fields are Required " & vbCrLf
            End If
            strmsg = strmsg & Space(10) & "- " & "Address 2" & vbCrLf
        End If
        
        If Trim(txtCity.Text) = "" Then
            If strmsg = "" Then
                strmsg = "Following Fields are Required " & vbCrLf
            End If
            strmsg = strmsg & Space(10) & "- " & "City" & vbCrLf
        End If
        
        If Trim(txtPincode.Text) = "" Then
            If strmsg = "" Then
                strmsg = "Following Fields are Required " & vbCrLf
            End If
            strmsg = strmsg & Space(10) & "- " & "PinCode" & vbCrLf
        End If
        
        If Trim(txtPhoneNo.Text) = "" Then
            If strmsg = "" Then
                strmsg = "Following Fields are Required " & vbCrLf
            End If
            strmsg = strmsg & Space(10) & "- " & "Phone Number" & vbCrLf
        End If
        
        If Trim(cboDept.Text) = "" Then
            If strmsg = "" Then
                strmsg = "Following Fields are Required " & vbCrLf
            End If
            strmsg = strmsg & Space(10) & "- " & "Department" & vbCrLf
        End If
        If Trim(txtDesignation.Text) = "" Then
            If strmsg = "" Then
                strmsg = "Following Fields are Required " & vbCrLf
            End If
            strmsg = strmsg & Space(10) & "- " & "Designation" & vbCrLf
        End If
        If Trim(txtSalary.Text) = "" Then
            If strmsg = "" Then
                strmsg = "Following Fields are Required " & vbCrLf
            End If
            strmsg = strmsg & Space(10) & "- " & "Phone Number" & vbCrLf
        End If
End Sub

Private Sub txtSalary_KeyPress(KeyAscii As Integer)
        If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
        KeyAscii = 0
        End If
End Sub

Public Sub delete_data()
        If txtEmpCode.Text = "" Or rs("emp_no") = "" Or rs.EOF Or rs.BOF Then
            MsgBox "Nothing Selected"
        ElseIf rs("emp_no") = "1101" Then
        MsgBox "Test Record Can not be Deleted"
        Else
            rs.Delete
            clear_all
            Disable_all
            refresh_data
        End If
End Sub

Public Sub refresh_data()
        strsql = "select emp_no,fname,lname from employee_details"
        adoEmployee.RecordSource = strsql
        adoEmployee.CommandType = adCmdText
        adoEmployee.Refresh
End Sub

Public Sub modify_data()
        If txtEmpCode.Text = "" Then
            MsgBox "Nothing Selected"
        Else
        For Each Control In frmEmployee
            If TypeOf Control Is TextBox Then
                Control.Enabled = True
            End If
        Next Control
        cboDept.Enabled = True
        fraDept.Enabled = False
        txtEmpCode.Enabled = False
        nstat = "modify"
        End If
End Sub
Public Sub cancel_data()
        clear_all
        Disable_all
        fraDept.Enabled = True
        cboDepart.Enabled = True
End Sub

Public Sub search_data()
        clear_all
        Disable_all
        cmdGo.Visible = True
        txtEmpCode.Enabled = True
        txtEmpCode.SetFocus
End Sub
