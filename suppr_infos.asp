<!--#include file="verif_ident.asp"-->
<!--#include file="connexion_perm.asp"-->
<%
Zid=request.Querystring("id")

SQLsuppr="DELETE * FROM [messagerie] WHERE id_message="&Zid
Set suppr= Server.CreateObject("ADODB.RecordSet")
suppr.open SQLsuppr,conn

response.redirect("news.asp?m=3")
%>