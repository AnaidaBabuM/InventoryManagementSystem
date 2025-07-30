<%
Function GetDBConnection()
    Dim conn
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open DB_CONNECTION_STRING
    Set GetDBConnection = conn
End Function
%>
