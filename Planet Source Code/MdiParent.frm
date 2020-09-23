VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MdiParent 
   BackColor       =   &H8000000C&
   Caption         =   " Complaints Management"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   Icon            =   "MdiParent.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList Imglst_toolbar 
      Left            =   1800
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiParent.frx":27A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiParent.frx":4F56
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiParent.frx":53AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiParent.frx":7B5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiParent.frx":A312
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiParent.frx":A62E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MdiParent.frx":A982
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbMain 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1111
      ButtonWidth     =   1085
      ButtonHeight    =   953
      ImageList       =   "Imglst_toolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modify"
            Key             =   "Modify"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "sep1"
            Key             =   "sep1"
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            Key             =   "Search"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "sep2"
            Key             =   "sep2"
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancel"
            Key             =   "Cancel"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "Delete"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "sep3"
            Key             =   "sep3"
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            Key             =   "Close"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu mnuum 
      Caption         =   "User Manager"
      Begin VB.Menu mnumascontrol 
         Caption         =   "Master Control"
      End
      Begin VB.Menu mnulogoff 
         Caption         =   "Log Off .."
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuCust 
      Caption         =   "Customer"
      Begin VB.Menu mnuCustomer 
         Caption         =   "Customer Registration"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Customer Search"
      End
   End
   Begin VB.Menu mnuEnt 
      Caption         =   "Enterprise Record"
      Begin VB.Menu mnuDepartment 
         Caption         =   "Department"
      End
      Begin VB.Menu mnuEmployees 
         Caption         =   "Employees"
      End
   End
   Begin VB.Menu mnuComplaint 
      Caption         =   "Complaint"
      Begin VB.Menu mnuGasBooking 
         Caption         =   "Complaint Booking"
      End
      Begin VB.Menu mnuNotification 
         Caption         =   "Notification"
      End
      Begin VB.Menu mnuCloseCall 
         Caption         =   "CloseCall"
      End
   End
   Begin VB.Menu mnu_chk_status 
      Caption         =   "Check Status"
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Flexi Reports/Query"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About CCMS"
   End
End
Attribute VB_Name = "MdiParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnu_chk_status_Click()
Load frmCheckStatus
frmCheckStatus.Show
frmCheckStatus.ZOrder
End Sub

Private Sub mnuAbout_Click()
Load frmAbout
frmAbout.Show
frmAbout.ZOrder
End Sub

Private Sub mnuCloseCall_Click()
frmCallClose.Show
frmCallClose.ZOrder
End Sub

Private Sub mnuCustomer_Click()
frmCustomerReg.Show
frmCustomerReg.ZOrder
End Sub

Private Sub mnuDepartment_Click()
frmDepartment.Show
frmDepartment.ZOrder
End Sub

Private Sub mnuEmployees_Click()
frmEmployee.Show
frmEmployee.ZOrder
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuGasBooking_Click()
frmCompBooking.Show
frmCompBooking.ZOrder
End Sub

Private Sub mnulogoff_Click()
Load frmLogin
frmLogin.Show
frmLogin.ZOrder
End Sub

Private Sub mnumascontrol_Click()
Load frmusermanager
frmusermanager.Show
frmusermanager.ZOrder

End Sub

Private Sub mnuNotification_Click()
frmNotification.Show
frmNotification.ZOrder
End Sub

Private Sub mnuReports_Click()
Load frmReports
frmReports.Show
frmReports.ZOrder
End Sub

Private Sub mnuSearch_Click()
frmCustSearch.Show (1)
frmCustSearch.Move 2000, 2000, 7200, 6000
frmCustSearch.ZOrder
End Sub


Private Sub tbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strval As String
strval = Button.Key

On Error GoTo e:

Select Case strval
    Case "New"
            'MsgBox "New Is Pressed"
            MdiParent.ActiveForm.add_data
    Case "Modify"
            'MsgBox "Modify"
            MdiParent.ActiveForm.modify_data
    Case "Save"
            'MsgBox "Save"
            MdiParent.ActiveForm.save_data
    Case "Search"
            'MsgBox "Search"
            MdiParent.ActiveForm.search_data
    Case "Cancel"
            'MsgBox "Cancel"
            MdiParent.ActiveForm.cancel_data
    Case "Delete"
            'MsgBox "Delete"
            MdiParent.ActiveForm.delete_data
    Case "Close"
            Unload MdiParent.ActiveForm
            
End Select
Exit Sub

e:
    MsgBox "Frame Not Available"
End Sub
