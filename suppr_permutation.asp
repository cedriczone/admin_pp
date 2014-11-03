<!--#include file="verif_ident.asp"-->
<!--#include file="connexion2.asp"-->

<%

Zid=request.queryString("id")

Zid=request.QueryString("id")
if Zid="" then response.Redirect("permutations.asp")

SQLsuppr="DELETE * from [permutation] where id_permutation="&Zid
Set suppr= Server.CreateObject("ADODB.RecordSet")
suppr.open SQLsuppr,conn2
conn2.close
Set conn2=nothing

response.redirect("permutations.asp")
%>