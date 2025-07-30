<%
Sub GetProducts(page)
    Dim conn, rs, sql
    Set conn = GetDBConnection()
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT p.ProductID, p.ProductName, p.UnitPrice, p.UnitsInStock, c.CategoryName, s.SupplierName " & _
          "FROM Products p " & _
          "INNER JOIN Categories c ON p.CategoryID = c.CategoryID " & _
          "INNER JOIN Suppliers s ON p.SupplierID = s.SupplierID " & _
          "ORDER BY p.ProductName " & _
          "OFFSET " & ((page - 1) * ITEMS_PER_PAGE) & " ROWS FETCH NEXT " & ITEMS_PER_PAGE & " ROWS ONLY"
    rs.Open sql, conn
%>
    <table>
        <tr>
            <th>Name</th>
            <th>Price</th>
            <th>Stock</th>
            <th>Category</th>
            <th>Supplier</th>
            <th>Actions</th>
        </tr>
<%
    Do While Not rs.EOF
%>
        <tr>
            <td><%= rs("ProductName") %></td>
            <td><%= FormatCurrency(rs("UnitPrice")) %></td>
            <td><%= rs("UnitsInStock") %></td>
            <td><%= rs("CategoryName") %></td>
            <td><%= rs("SupplierName") %></td>
            <td>
                <a href="products.asp?action=edit&id=<%= rs("ProductID") %>">Edit</a>
                <a href="products.asp?action=delete&id=<%= rs("ProductID") %>" onclick="return confirmDelete()">Delete</a>
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

Sub SaveProduct(id, name, price, stock, categoryId, supplierId)
    Dim conn, cmd
    Set conn = GetDBConnection()
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    
    If id = "" Then
        cmd.CommandText = "INSERT INTO Products (ProductName, UnitPrice, UnitsInStock, CategoryID, SupplierID) VALUES (?,?,?,?,?)"
    Else
        cmd.CommandText = "UPDATE Products SET ProductName = ?, UnitPrice = ?, UnitsInStock = ?, CategoryID = ?, SupplierID = ? WHERE ProductID = ?"
    End If
    
    cmd.Prepared = True
    cmd.Parameters.Append cmd.CreateParameter("name", 200, 1, 255, name)
    cmd.Parameters.Append cmd.CreateParameter("price", 5, 1, 0, CDbl(price))
    cmd.Parameters.Append cmd.CreateParameter("stock", 3, 1, 0, CInt(stock))
    cmd.Parameters.Append cmd.CreateParameter("category", 3, 1, 0, CInt(categoryId))
    cmd.Parameters.Append cmd.CreateParameter("supplier", 3, 1, 0, CInt(supplierId))
    If id <> "" Then
        cmd.Parameters.Append cmd.CreateParameter("id", 3, 1, 0, CInt(id))
    End If
    
    cmd.Execute
    conn.Close
End Sub
%>

/services/category_service.asp
