<!--#include file="verif_ident.asp"-->
<!--#include file="connexion2.asp"-->
<%
Zdate_ech=request.form("date_ech")
Zremplace=request.form("remplace")
Zremplacant=request.form("remplacant")

position=left(Zremplace,1)
Zzremplace=mid(Zremplace,2)

Select case position
case 1
poste="num_coordinateur"
case 2
poste="num_inter_chev"
case 3
poste="num_inter1"
case 4
poste="num_inter2"
case 5
poste="num_inter3"
case 6
poste="num_inter4"
case 7
poste="num_inter5"
case 8
poste="num_inter6"
end Select

SQLmodif="UPDATE [planning_gav] set "&poste&"="&Zremplacant&" WHERE "&poste&"="&Zzremplace&" and date_gav=#"&Zdate_ech&"#"

Set modif= Server.CreateObject("ADODB.RecordSet")
modif.open SQLmodif,conn2

response.redirect("planning_gav.asp")
%>
