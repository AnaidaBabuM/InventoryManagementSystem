<% Option Explicit %>
<!-- #include file="includes/config.asp" -->
<!-- #include file="includes/db_connect.asp" -->
<!-- #include file="includes/header.asp" -->
<!-- #include file="services/product_service.asp" -->
<h1>Products</h1>
<%
  Dim action, id, conn, rs
  Dim productName, unitPrice, unitsInStock, categoryID, supplierID
  action = Request.QueryString("action")
  id = Request.QueryString("id")

  If action = "delete" And id <> "" Then
    Set conn = GetDBConnection()
    On Error Resume Next
    conn.Execute "DELETE FROM Products WHERE ProductID = " & CInt(id)
    If Err.Number <> 0 Then
      Response.Write "Error deleting product: " & Err.Description
      Response.End
    End If
    On Error GoTo 0
    conn.Close
    Response.Redirect "products.asp"
  End If

  If Request.Form("submit") <> "" Then
    Call SaveProduct(Request.Form("id"), Request.Form("name"), Request.Form("price"), Request.Form("stock"), Request.Form("category"), Request.Form("supplier"))
    Response.Redirect "products.asp"
  End If

  If action = "edit" And id <> "" Then
    Set conn = GetDBConnection()
    Set rs = Server.CreateObject("ADODB.Recordset")
    On Error Resume Next
    rs.Open "SELECT * FROM Products WHERE ProductID = " & CInt(id), conn
    If Err.Number <> 0 Then
      Response.Write "Error retrieving product: " & Err.Description
      Response.End
    End If
    If Not rs.EOF Then
      productName = rs("ProductName")
      unitPrice = rs("UnitPrice")
      unitsInStock = rs("UnitsInStock")
      categoryID = rs("CategoryID")
      supplierID = rs("SupplierID")
    End If
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
  End If
%>
<div class="form-group">
  <h2><%= IIf(action = "edit", "Edit Product", "Add Product") %></h2>
  <form method="post" action="products.asp" name="productForm" onsubmit="return validateProductForm()">
    <input type="hidden" name="id" value="<%= id %>">
    <div class="form-group">
      <label for="name">Name:</label>
      <input type="text" id="name" name="name" value="<%= Server.HTMLEncode(productName & "") %>">
    </div>
    <div class="form-group">
      <label for="price">Price:</label>
      <input type="number" id="price" name="price" value="<%= unitPrice %>" step="0.01" min="0">
    </div>
    <div class="form-group">
      <label for="stock">Stock:</label>
      <input type="number" id="stock" name="stock" value="<%= unitsInStock %>" min="0">
    </div>
    <div class="form-group">
      <label for="category">Category:</label>
      <select id="category" name="category">
        <%
          Set conn = GetDBConnection()
          Set rs = Server.CreateObject("ADODB.Recordset")
          On Error Resume Next
          rs.Open "SELECT CategoryID, CategoryName FROM Categories ORDER BY CategoryName", conn
          If Err.Number <> 0 Then
            Response.Write "Error retrieving categories: " & Err.Description
            Response.End
          End If
          Do While Not rs.EOF
        %>
          <option value="<%= rs("CategoryID") %>" <%= IIf(categoryID = rs("CategoryID"), "selected", "") %>><%= Server.HTMLEncode(rs("CategoryName")) %></option>
        <%
            rs.MoveNext
          Loop
          rs.Close
        %>
      </select>
    </div>
    <div class="form-group">
      <label for="supplier">Supplier:</label>
      <select id="supplier" name="supplier">
        <%
          rs.Open "SELECT SupplierID, SupplierName FROM Suppliers ORDER BY SupplierName", conn
          If Err.Number <> 0 Then
            Response.Write "Error retrieving suppliers: " & Err.Description
            Response.End
          End If
          Do While Not rs.EOF
        %>
          <option value="<%= rs("SupplierID") %>" <%= IIf(supplierID = rs("SupplierID"), "selected", "") %>><%= Server.HTMLEncode(rs("SupplierName")) %></option>
        <%
            rs.MoveNext
          Loop
          rs.Close
          Set rs = Nothing
          conn.Close
          Set conn = Nothing
        %>
      </select>
    </div>
    <button type="submit" name="submit">Save</button>
  </form>
</div>
<h2>Product List</h2>
<% Call GetProducts(1) %>
<!-- #include file="includes/footer.asp" -->
