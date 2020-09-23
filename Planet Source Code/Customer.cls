VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarcust_no As Single 'local copy
Private mvarcust_name As String 'local copy
'local variable(s) to hold property value(s)
Private mvarfname As String 'local copy
Private mvarlname As String 'local copy
Private mvarhouseno As String 'local copy
Private mvaradd1 As String 'local copy
Private mvaradd2 As String 'local copy
Private mvarlocality As String 'local copy
Private mvarlandmark As String 'local copy
Private mvarcity As String 'local copy
Private mvarpincode As Double 'local copy
Private mvarphone As String 'local copy
Private mvaraddproof As String 'local copy
Private mvarappliance As String 'local copy
Private mvarcust_type As String 'local copy
Private mvardop As String 'local copy
Private mvarplant As String 'local copy
Private mvarzone As String 'local copy
Private mvarsp As String 'local copy
Public Property Let sp(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sp = 5
    mvarsp = vData
End Property


Public Property Get sp() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sp
    sp = mvarsp
End Property



Public Property Let zone(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.zone = 5
    mvarzone = vData
End Property


Public Property Get zone() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.zone
    zone = mvarzone
End Property



Public Property Let plant(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.plant = 5
    mvarplant = vData
End Property


Public Property Get plant() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.plant
    plant = mvarplant
End Property



Public Property Let dop(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dop = 5
    mvardop = vData
End Property


Public Property Get dop() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dop
    dop = mvardop
End Property



Public Property Let cust_type(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.cust_type = 5
    mvarcust_type = vData
End Property


Public Property Get cust_type() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cust_type
    cust_type = mvarcust_type
End Property



Public Property Let appliance(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.appliance = 5
    mvarappliance = vData
End Property


Public Property Get appliance() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.appliance
    appliance = mvarappliance
End Property



Public Property Let addproof(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.addproof = 5
    mvaraddproof = vData
End Property


Public Property Get addproof() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.addproof
    addproof = mvaraddproof
End Property



Public Property Let phone(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.phone = 5
    mvarphone = vData
End Property


Public Property Get phone() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.phone
    phone = mvarphone
End Property



Public Property Let pincode(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.pincode = 5
    mvarpincode = vData
End Property


Public Property Get pincode() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.pincode
    pincode = mvarpincode
End Property



Public Property Let city(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.city = 5
    mvarcity = vData
End Property


Public Property Get city() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.city
    city = mvarcity
End Property



Public Property Let landmark(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.landmark = 5
    mvarlandmark = vData
End Property


Public Property Get landmark() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.landmark
    landmark = mvarlandmark
End Property



Public Property Let locality(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.locality = 5
    mvarlocality = vData
End Property


Public Property Get locality() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.locality
    locality = mvarlocality
End Property



Public Property Let add2(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.add2 = 5
    mvaradd2 = vData
End Property


Public Property Get add2() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.add2
    add2 = mvaradd2
End Property



Public Property Let add1(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.add1 = 5
    mvaradd1 = vData
End Property


Public Property Get add1() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.add1
    add1 = mvaradd1
End Property



Public Property Let houseno(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.houseno = 5
    mvarhouseno = vData
End Property


Public Property Get houseno() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.houseno
    houseno = mvarhouseno
End Property



Public Property Let lname(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lname = 5
    mvarlname = vData
End Property


Public Property Get lname() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lname
    lname = mvarlname
End Property



Public Property Let fname(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fname = 5
    mvarfname = vData
End Property


Public Property Get fname() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fname
    fname = mvarfname
End Property



Public Property Let cust_name(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.cust_name = 5
    mvarcust_name = vData
End Property


Public Property Get cust_name() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cust_name
    cust_name = mvarcust_name
End Property



Public Property Let cust_no(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.cust_no = 5
    mvarcust_no = vData
End Property


Public Property Get cust_no() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cust_no
    cust_no = mvarcust_no
End Property



