<!--#include file="connexion_perm.asp"-->
<%
SQLpwd="SELECT * from [admin] where id_admin="&cint(request.Cookies("idadmin"))
Set rspwd=server.Createobject("adodb.recordset")
rspwd.open SQLpwd,conn,3,3

Zoldpwd=request.form("old_pwd")
Znew_pwd=request.form("new_pwd")
Znew_pwd2=request.form("new_pwd2")

if Znew_pwd<>Znew_pwd2 then response.redirect("modif_admin.asp?m=2")
if Zoldpwd<>rspwd("password") then response.redirect("modif_admin.asp?m=3")

SQLmodif="UPDATE [admin] set password='"&Znew_pwd&"' WHERE id_admin="&request.Cookies("idadmin")
Set modif= Server.CreateObject("ADODB.RecordSet")
modif.open SQLmodif,conn

response.redirect("modif_admin.asp?m=1")
%>