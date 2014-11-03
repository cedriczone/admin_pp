<!--#include file="connexion_perm.asp"-->
<%
Zcat=request.Form("cat")
Znveau_nom_cat=request.Form("nveau_nom_cat")
Znveau_nom_cat=replace(Znveau_nom_cat,"'","''")

SQLmodif="UPDATE [cat] set nom_cat='"&Znveau_nom_cat&"' WHERE id_cat="&Zcat
Set modif= Server.CreateObject("ADODB.RecordSet")
modif.open SQLmodif,conn

response.redirect("documents.asp")
%>