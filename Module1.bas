Attribute VB_Name = "Module1"
Public con As New ADODB.Connection
Public rs As New ADODB.Recordset
Public Sub getconnection()
If con.state = adstateopen Then
con.Close
End If
con.connectionstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\MOCHA.MDB;Persist Security Info=False"
'con.connectionstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\user\Documents\MOCHA.MDB;Persist Security Info=False"
con.Mode = adModeShareDenyNone
con.Open
End Sub
Public Sub writedata(ByVal sql As String)
getconnection
con.Execute sql
End Sub
Public Sub readdata(ByVal sql As String)
getconnection
rs.Open sql, con
End Sub
