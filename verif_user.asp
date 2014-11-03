<!--#include file="connexion_perm.asp"-->
<%
Zlogin=request.form("login")
SQLuser="SELECT * from [login] where login='"&Zlogin&"'"
Set rsuser=server.Createobject("adodb.recordset")
rsuser.open SQLuser,conn,3,3
nbre_user=rsuser.recordcount
if nbre_user<>0 then
%>
<span id="verifuser" style="color:#FF0000; margin-left:5px"><strong>login existant</strong></span>
<%end if%>