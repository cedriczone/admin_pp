<!--#include file="connexion2.asp"-->
<%
Zid=request.form("id")
SQLsuppr="DELETE * from [defilantes] where id_defilante="&Zid
Set suppr= Server.CreateObject("ADODB.RecordSet")
suppr.open SQLsuppr,conn2
conn2.close
Set conn2=nothing
%>