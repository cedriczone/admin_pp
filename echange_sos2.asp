<!--#include file="verif_ident.asp"-->
<!--#include file="connexion2.asp"-->
<%
Zdate_ech=request.form("date_ech")
Zech=request.querystring("ech")
if Zech="" then response.redirect("default.asp")
if Zech="tit" then
Zancien=request.form("ancien_tit")
Znveau=request.form("nveau_tit")
SQLmodif="UPDATE [planning_sos] set titulaire="&Znveau&" WHERE titulaire="&Zancien&" and jour=#"&Zdate_ech&"#"

elseif zech="sup" then
Zancien=request.form("ancien_sup")
Znveau=request.form("nveau_sup")
SQLmodif="UPDATE [planning_sos] set suppleant="&Znveau&" WHERE suppleant="&Zancien&" and jour=#"&Zdate_ech&"#"

else
response.redirect("default.asp")
end if

Set modif= Server.CreateObject("ADODB.RecordSet")
modif.open SQLmodif,conn2

response.redirect("gene_planning_sos.asp")
%>
