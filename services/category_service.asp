<% 
Option Explicit

' Retrieve all categories
Sub GetCategories()
  Dim conn, rs
  Set conn = GetDBConnection()
  Set rs = Server.CreateObject("ADODB.Recordset")
  
  On Error Resume Next
  rs.Open "SELECT CategoryID, CategoryName FROM Categories ORDER BY CategoryName", conn
  If Err.Number <> 0 Then
    Response.Write "Error retrieving categories: " & Err.Description
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
      <td><%= Server.HTMLEncode(rs("CategoryName")) %></td>
      <td>
        <a href="categories.asp?action=edit&id=<%= rs("CategoryID") %>">Edit</a>
        <a href="categories.asp?action=delete&id=<%= rs("CategoryID") %>" onclick="return confirmDelete()">Delete</a>
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

' Save or update category
Sub SaveCategory(id, name)
  Dim conn, cmd
  Set conn = GetDBConnection()
  Set cmd = Server.CreateObject("ADODB.Command")
  Set cmd.ActiveConnection = conn
  
  If id = "" Then
    cmd.CommandText = "INSERT INTO Categories (CategoryName) VALUES (?)"
  Else
    cmd.CommandText = "UPDATE Categories SET CategoryName = ? WHERE CategoryID = ?"
  End If
  
  cmd.Prepared = True
  cmd.Parameters.Append cmd.CreateParameter("name", 200, 1, 255, name)
  If id <> "" Then
    cmd.Parameters.Append cmd.CreateParameter("id", 3, 1, , CInt(id))
  End If
  
  On Error Resume Next
  cmd.Execute
  If Err.Number <> 0 Then
    Response.Write "Error saving category: " & Err.Description
    Response.End
  End If
  On Error GoTo 0
  
  conn.Close
  Set conn = Nothing
  Set cmd = Nothing
End Sub
%>
