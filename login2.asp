<!--#include file="connexion_perm.asp"-->
<%
Zlogin=request.form("login")
Zpassword=request.form("password")
if Zlogin="" or Zpassword="" then
response.redirect("login.asp")
else
SQLlogin="SELECT * from [admin] where login='"&Zlogin&"' and password='"&Zpassword&"'"
Set rslogin=server.Createobject("adodb.recordset")
rslogin.open SQLlogin,conn,3,3
nbre=rslogin.recordcount
if nbre<>1 then
response.redirect("default.asp")
else
Response.Cookies("adminpp")="ok"
Response.Cookies("idadmin")=rslogin("id_admin")
response.redirect("default.asp")
end if
end if
rslogin.close
Set rslogin=nothing
conn.close
Set conn=nothing
%>