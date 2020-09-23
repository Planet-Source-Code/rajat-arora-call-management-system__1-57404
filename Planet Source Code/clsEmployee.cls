VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Dim emp_no As String
Dim fname As String
Dim lname As String
Dim add1 As String
Dim add2 As String
Dim city As String
Dim pincode As String
Dim phone As String
Dim department As String
Dim desig As String
Dim salary As String

Public Property Get employee_no() As Variant
        employee_no = emp_no
End Property

Public Property Let employee_no(ByVal vNewValue As Variant)
        emp_no = vNewValue
End Property

Public Property Get first_name() As Variant
        first_name = fname
End Property

Public Property Let first_name(ByVal vNewValue As Variant)
        fname = vNewValue
End Property

Public Property Get last_name() As Variant
        last_name = lname
End Property

Public Property Let last_name(ByVal vNewValue As Variant)
        lname = vNewValue
End Property

Public Property Get address1() As Variant
        address1 = add1
End Property

Public Property Let address1(ByVal vNewValue As Variant)
        add1 = vNewValue
End Property

Public Property Get address2() As Variant
        address2 = add2
End Property

Public Property Let address2(ByVal vNewValue As Variant)
        add2 = address2
End Property

Public Property Get pcity() As Variant
        pcity = city
End Property

Public Property Let pcity(ByVal vNewValue As Variant)
        city = vNewValue
End Property

Public Property Get pin() As Variant
        pin = pincode
End Property

Public Property Let pin(ByVal vNewValue As Variant)
        pincode = vNewValue
End Property

Public Property Get phoneno() As Variant
        phoneno = phone
End Property

Public Property Let phoneno(ByVal vNewValue As Variant)
        phone = vNewValue
End Property

Public Property Get dpt() As Variant
        dpt = department
End Property

Public Property Let dpt(ByVal vNewValue As Variant)
        department = vNewValue
End Property

Public Property Get designation() As Variant
        designation = desig
End Property

Public Property Let designation(ByVal vNewValue As Variant)
        desig = vNewValue
End Property

Public Property Get sal() As Variant
        sal = salary
End Property

Public Property Let sal(ByVal vNewValue As Variant)
        salary = vNewValue
End Property
