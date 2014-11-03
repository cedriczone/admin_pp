<!--#include file="connexion2.asp"-->
<%
Zlogin=request.form("login")
Zpwd1=request.form("pwd1")
Zpwd2=request.form("pwd2")
Zemail=request.form("email")
Zavo=request.Form("avo")

SQLadd_user="Insert Into [login](login,password,email,avo) Values('"&Zlogin&"','"&Zpwd1&"','"&Zemail&"',"&Zavo&")"
Set saisie= Server.CreateObject("ADODB.RecordSet")
saisie.open SQLadd_user,conn2

conn2.close
Set conn2=nothing

response.redirect("gestion_login.asp?m=1")
%>