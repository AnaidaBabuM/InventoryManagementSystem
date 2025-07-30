<% Option Explicit %>
<!-- #include file="includes/config.asp" -->
<!-- #include file="includes/db_connect.asp" -->
<!-- #include file="includes/header.asp" -->
<!-- #include file="services/supplier_service.asp" -->
<h1>Suppliers</h1>
<%
  Dim action, id, conn, rs, supplierName
  action = Request.QueryString("action")
  id = Request.QueryString("id")

  If action = "delete" And id <> "" Then
    Set conn = GetDBConnection()
    On Error Resume Next
    conn.Execute "DELETE FROM Suppliers WHERE SupplierID = " & CInt(id)
    If Err.Number <> 0 Then
      Response.Write "Error deleting supplier: " & Err.Description
      Response.End
    End If
    On Error GoTo 0
    conn.Close
    Response.Redirect "suppliers.asp"
  End If

  If Request.Form("submit") <> "" Then
    Call SaveSupplier(Request.Form("id"), Request.Form("name"))
    Response.Redirect "suppliers.asp"
  End If

  If action = "edit" And id <> "" Then
    Set conn = GetDBConnection()
    Set rs = Server.CreateObject("ADODB.Recordset")
    On Error Resume Next
    rs.Open "SELECT SupplierName FROM Suppliers WHERE SupplierID = " & CInt(id), conn
    If Err.Number <> 0 Then
      Response.Write "Error retrieving supplier: " & Err.Description
      Response.End
    End If
    If Not rs.EOF Then
      supplierName = rs("SupplierName")
    End If
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
  End If
%>
<div class="form-group">
  <h2><%= IIf(action = "edit", "Edit Supplier", "Add Supplier") %></h2>
  <form method="post" action="suppliers.asp">
    <input type="hidden" name="id" value="<%= id %>">
    <div class="form-group">
      <label for="name">Name:</label>
      <input type="text" id="name" name="name" value="<%= Server.HTMLEncode(supplierName & "") %>">
    </div>
    <button type="submit" name="submit">Save</button>
  </form>
</div>
<h2>Supplier List</h2>
<% Call GetSuppliers() %>
<!-- #include file="includes/footer.asp" -->
