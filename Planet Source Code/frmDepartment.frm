VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmDepartment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Department Form"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3720
   Icon            =   "frmDepartment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraDepartment 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3495
      Begin VB.CommandButton cmdGo 
         Caption         =   "GO"
         Height          =   375
         Left            =   2880
         TabIndex        =   6
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSAdodcLib.Adodc adoDept 
         Height          =   330
         Left            =   1200
         Top             =   4440
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
         Caption         =   "Adodc1"
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
      Begin VB.TextBox txtSearch 
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
         Left            =   1155
         MaxLength       =   25
         TabIndex        =   3
         Top             =   4920
         Width           =   1695
      End
      Begin MSDataGridLib.DataGrid dgDepartmentList 
         Bindings        =   "frmDepartment.frx":27A2
         Height          =   3975
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   7011
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         Appearance      =   0
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
         ColumnCount     =   1
         BeginProperty Column00 
            DataField       =   "department"
            Caption         =   "     DEPARTMENT"
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
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
      Begin VB.Label lblDepartment 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
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
         Left            =   45
         TabIndex        =   5
         Top             =   5040
         Width           =   1050
      End
      Begin VB.Label lblSearch 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search :"
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
         TabIndex        =   4
         Top             =   4560
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Department List  "
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
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "frmDepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rsEmp As ADODB.Recordset
Dim strDel As String
Dim err As ErrObject

Private Sub cmdGo_Click()
        Dim nstats As Integer
        dgDepartmentList.Col = 0
               
        For nRow = 0 To dgDepartmentList.VisibleRows - 1 Step 1
            dgDepartmentList.Row = nRow
            'MsgBox dgDepartmentList.Text
            If dgDepartmentList.Text = UCase(Trim(txtSearch.Text)) Then
                dgDepartmentList.Columns(0).Button = True
                'MsgBox "found"
                nstats = 1
                Exit For
            Else
                'MsgBox "Not found"
                nstats = 0
            End If
        Next
        
        If nstats = 0 Then
        txtSearch.Text = ""
        txtSearch.Enabled = False
        dgDepartmentList.Row = 0
        MsgBox "Department Not Exist"
        End If
        cmdGo.Visible = False
End Sub

Private Sub dgDepartmentList_Click()
    On Error GoTo ErrHd
        'dgDepartmentList.Col (0)
        strDel = dgDepartmentList.Text
        txtSearch.Text = strDel
        Exit Sub
ErrHd:
    MsgBox "Data Not Present"
End Sub

Private Sub Form_Activate()
        Set rs = New ADODB.Recordset
        Set rsEmp = New ADODB.Recordset

        mdDataConnection.data_connection
        adoDept.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=ccms"
        adoDept.RecordSource = "select department from department"
        adoDept.CommandType = adCmdText
        adoDept.Refresh
        
        
        rs.Open "select * from department", mdDataConnection.con, adOpenDynamic, adLockOptimistic
        rsEmp.Open "select emp_no,fname,department from employee_details", mdDataConnection.con, adOpenDynamic, adLockOptimistic
End Sub

Public Sub add_data()
       txtSearch.Enabled = True
       lblSearch(0).Caption = "Add Department"
       lblSearch(0).Visible = True
       txtSearch.Text = ""
       txtSearch.SetFocus
End Sub

Public Sub delete_data()
        Dim nstatus As Integer
        If strDel = "" Then
            MsgBox "Nothing Selected"
            Exit Sub
        End If
                
        ' Check for existing of Employees in the Department
        rsEmp.MoveFirst
        Do While Not rsEmp.EOF
            If rsEmp("department") = strDel Then
                nstatus = 1
            Exit Do
            Else
                nstatus = 0
                rsEmp.MoveNext
            End If
        Loop
        
np:
        If nstatus = 0 Then
        rs.MoveFirst
        While Not rs.EOF
            If rs("department") = strDel Then
                rs.Delete
                txtSearch.Text = ""
                MsgBox "Department Successfully Deleted"
                refresh_data
                Exit Sub
            Else
                rs.MoveNext
            End If
        Wend
        Else
            MsgBox "There are some Employees Associated With this Department"
        End If
End Sub

Public Sub refresh_data()

        adoDept.RecordSource = "select * from department"
        adoDept.CommandType = adCmdText
        adoDept.Refresh
End Sub

Public Sub save_data()
        If Trim(txtSearch) = "" Then
            MsgBox "Value Can't be Null"
        Else
        rs.MoveFirst
        Do While Not rs.EOF
            If rs("department") = UCase(Trim(txtSearch)) Then
                MsgBox "Department Already Exist"
                txtSearch.Text = ""
                Exit Sub
            Else
                rs.MoveNext
            End If
        Loop
        
        rs.AddNew
        rs("department") = UCase(Trim(txtSearch))
        rs.Update
        MsgBox "Department Saved Successfully "
        adoDept.Refresh
        lblSearch(0).Visible = False
        txtSearch.Enabled = False
        End If
End Sub

Public Sub cancel_data()
        txtSearch.Text = ""
        txtSearch.Enabled = False
        lblSearch(0).Visible = False
        cmdGo.Visible = False
End Sub

Public Sub search_data()
        Dim nRow As Integer
        lblSearch(0).Caption = "Search Department"
        lblSearch(0).Visible = True
        txtSearch.Text = ""
        txtSearch.Enabled = True
        txtSearch.SetFocus
        cmdGo.Visible = True
End Sub

