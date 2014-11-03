<%
connection = "DBQ=" & Server.MapPath("../../data/base_perm.mdb")&";DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}"
Set conn = Server.CreateObject("ADODB.Connection")
conn.open connection
%>