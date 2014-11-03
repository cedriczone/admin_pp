<!--#include file="connexion2.asp"-->
<%
Zid=request.form("id")
SQLsuppr="DELETE * from [infojour] where id_infojour="&Zid
Set suppr= Server.CreateObject("ADODB.RecordSet")
suppr.open SQLsuppr,conn2
conn2.close
Set conn2=nothing
response.redirect("infojour.asp")
%>