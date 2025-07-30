<% Option Explicit %>
<!-- #include file="includes/config.asp" -->
<!-- #include file="includes/db_connect.asp" -->
<!-- #include file="includes/header.asp" -->
<h1>Dashboard</h1>
<%
  Dim conn, rs, productCount, categoryCount, supplierCount
  Set conn = GetDBConnection()
  Set rs = Server.CreateObject("ADODB.Recordset")
  
  On Error Resume Next
  
  ' Get product count
  rs.Open "SELECT COUNT(*) as ProductCount FROM Products", conn
  If Err.Number <> 0 Then
    Response.Write "Error retrieving product count: " & Err.Description
    Response.End
  End If
  productCount = rs("ProductCount")
  rs.Close
  
  ' Get category count
  rs.Open "SELECT COUNT(*) as CategoryCount FROM Categories", conn
  If Err.Number <> 0 Then
    Response.Write "Error retrieving category count: " & Err.Description
    Response.End
  End If
  categoryCount = rs("CategoryCount")
  rs.Close
  
  ' Get supplier count
  rs.Open "SELECT COUNT(*) as SupplierCount FROM Suppliers", conn
  If Err.Number <> 0 Then
    Response.Write "Error retrieving supplier count: " & Err.Description
    Response.End
  End If
  supplierCount = rs("SupplierCount")
  rs.Close
  
  ' Get low stock products
  rs.Open "SELECT TOP 5 ProductName, UnitsInStock FROM Products WHERE UnitsInStock < 10 ORDER BY UnitsInStock", conn
  If Err.Number <> 0 Then
    Response.Write "Error retrieving low stock products: " & Err.Description
    Response.End
  End If
%>
<div class="stats">
  <div>Total Products: <%= productCount %></div>
  <div>Total Categories: <%= categoryCount %></div>
  <div>Total Suppliers: <%= supplierCount %></div>
</div>
<h2>Low Stock Products</h2>
<table>
  <tr>
    <th>Product Name</th>
    <th>Units in Stock</th>
  </tr>
<%
  Do While Not rs.EOF
%>
  <tr>
    <td><%= Server.HTMLEncode(rs("ProductName")) %></td>
    <td><%= rs("UnitsInStock") %></td>
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
<!-- #include file="includes/footer.asp" -->
