<!--#include file="connexion_perm.asp"-->
<%
Zid=request.querystring("id")
SQLsuppr="DELETE * from [vacances] where id_plage="&Zid
Set suppr= Server.CreateObject("ADODB.RecordSet")
suppr.open SQLsuppr,conn
conn.close
Set conn=nothing
response.Redirect("vacances_jud.asp?m=4")
%>