<!--#include file="connexion_perm.asp"-->
<%
Zcat=request.Form("cat")

SQLsuppr="DELETE * from [cat] where id_cat="&Zcat
Set suppr= Server.CreateObject("ADODB.RecordSet")
suppr.open SQLsuppr,conn
conn.close
Set conn=nothing

response.Redirect("documents.asp")
%>