<!--#include file="verif_ident.asp"-->
<!--#include file="connexion2.asp"-->
<%
Zid=request.QueryString("id")
if Zid="" then response.Redirect("dispos_gav_coord.asp")

SQLsuppr="DELETE * from [dispos_coord_gav] where id_dispo="&Zid
Set suppr= Server.CreateObject("ADODB.RecordSet")
suppr.open SQLsuppr,conn2
conn2.close
Set conn2=nothing

response.Redirect("dispos_gav_coord.asp?m=3")
%>