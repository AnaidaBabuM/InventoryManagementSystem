<% 
Option Explicit

' Retrieve all suppliers
Sub GetSuppliers()
  Dim conn, rs
  Set conn = GetDBConnection()
  Set rs = Server.CreateObject("ADODB.Recordset")
  
  On Error Resume Next
  rs.Open "SELECT SupplierID, SupplierName FROM Suppliers ORDER BY SupplierName", conn
  If Err.Number <> 0 Then
    Response.Write "Error retrieving suppliers: " & Err.Description
    Response.End
  End If
  On Error GoTo 0
%>
  <table>
    <tr>
      <th>Name</th>
      <th>Actions</th>
    </tr>
<%
  Do While Not rs.EOF
%>
    <tr>
      <td><%= Server.HTMLEncode(rs("SupplierName")) %></td>
      <td>
        <a href="suppliers.asp?action=edit&id=<%= rs("SupplierID") %>">Edit</a>
        <a href="suppliers.asp?action=delete&id=<%= rs("SupplierID") %>" onclick="return confirmDelete()">Delete</a>
      </td>
    </tr>
<%
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing
  conn.Close
  Set conn = Nothing
%>
  </table>
<%
End Sub

' Save or update supplier
Sub SaveSupplier(id, name)
  Dim conn, cmd
  Set conn = GetDBConnection()
  Set cmd = Server.CreateObject("ADODB.Command")
  Set cmd.ActiveConnection = conn
  
  If id = "" Then
    cmd.CommandText = "INSERT INTO Suppliers (SupplierName) VALUES (?)"
  Else
    cmd.CommandText = "UPDATE Suppliers SET SupplierName = ? WHERE SupplierID = ?"
  End If
  
  cmd.Prepared = True
  cmd.Parameters.Append cmd.CreateParameter("name", 200, 1, 255, name)
  If id <> "" Then
    cmd.Parameters.Append cmd.CreateParameter("id", 3, 1, , CInt(id))
  End If
  
  On Error Resume Next
  cmd.Execute
  If Err.Number <> 0 Then
    Response.Write "Error saving supplier: " & Err.Description
    Response.End
  End If
  On Error GoTo 0
  
  conn.Close
  Set conn = Nothing
  Set cmd = Nothing
End Sub
%>
