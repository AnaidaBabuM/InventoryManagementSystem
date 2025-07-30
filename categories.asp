<% Option Explicit %>
<!-- #include file="includes/config.asp" -->
<!-- #include file="includes/db_connect.asp" -->
<!-- #include file="includes/header.asp" -->
<!-- #include file="services/category_service.asp" -->
<h1>Categories</h1>
<%
  Dim action, id, conn, rs, categoryName
  action = Request.QueryString("action")
  id = Request.QueryString("id")

  If action = "delete" And id <> "" Then
    Set conn = GetDBConnection()
    On Error Resume Next
    conn.Execute "DELETE FROM Categories WHERE CategoryID = " & CInt(id)
    If Err.Number <> 0 Then
      Response.Write "Error deleting category: " & Err.Description
      Response.End
    End If
    On Error GoTo 0
    conn.Close
    Response.Redirect "categories.asp"
  End If

  If Request.Form("submit") <> "" Then
    Call SaveCategory(Request.Form("id"), Request.Form("name"))
    Response.Redirect "categories.asp"
  End If

  If action = "edit" And id <> "" Then
    Set conn = GetDBConnection()
    Set rs = Server.CreateObject("ADODB.Recordset")
    On Error Resume Next
    rs.Open "SELECT CategoryName FROM Categories WHERE CategoryID = " & CInt(id), conn
    If Err.Number <> 0 Then
      Response.Write "Error retrieving category: " & Err.Description
      Response.End
    End If
    If Not rs.EOF Then
      categoryName = rs("CategoryName")
    End If
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
  End If
%>
<div class="form-group">
  <h2><%= IIf(action = "edit", "Edit Category", "Add Category") %></h2>
  <form method="post" action="categories.asp">
    <input type="hidden" name="id" value="<%= id %>">
    <div class="form-group">
      <label for="name">Name:</label>
      <input type="text" id="name" name="name" value="<%= Server.HTMLEncode(categoryName & "") %>">
    </div>
    <button type="submit" name="submit">Save</button>
  </form>
</div>
<h2>Category List</h2>
<% Call GetCategories() %>
<!-- #include file="includes/footer.asp" -->
