Attribute VB_Name = "mdDataConnection"
Public con As ADODB.Connection
Public cust_no As String
Public searchStatus As String
Public searchCust As String
Public cust As New clsCustomer
Public us As String


Public Sub data_connection()
Set con = New ADODB.Connection
'Set cust = New clsCustomer

'con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=CallManagement.mdb;Persist Security Info=False"
con.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=ccms"

End Sub

