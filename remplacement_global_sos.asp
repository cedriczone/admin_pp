<!--#include file="verif_ident.asp"-->
<!--#include file="connexion2.asp"-->
<%
Zdate_debut=request.form("date_debut")
Zdate_fin=request.form("date_fin")
Zavocat=request.form("avocat")
Zavocat2=request.form("avocat2")
if Zavocat="" then response.redirect("gene_planning_sos.asp")
if Zavocat2="" then response.redirect("gene_planning_sos.asp")

if Zdate_debut="" then Zdate_debut=date()
if Zdate_fin="" then
SQLdate="SELECT * from [planning_sos] order by jour DESC"
Set rsdate=server.Createobject("adodb.recordset")
rsdate.open SQLdate,conn2,3,3
rsdate.movefirst
Zdate_fin=rsdate("jour")
end if

Zzdate_debut=month(Zdate_debut)&"/"&day(Zdate_debut)&"/"&year(Zdate_debut)
Zzdate_fin=month(Zdate_fin)&"/"&day(Zdate_fin)&"/"&year(Zdate_fin)

SQLmodif="UPDATE [planning_sos] set titulaire="&Zavocat2&" WHERE titulaire="&Zavocat&" and jour>=#"&Zzdate_debut&"# and jour<=#"&Zzdate_fin&"#"
Set modif= Server.CreateObject("ADODB.RecordSet")
modif.open SQLmodif,conn2

SQLmodif2="UPDATE [planning_sos] set suppleant="&Zavocat2&" WHERE suppleant="&Zavocat&" and jour>=#"&Zzdate_debut&"# and jour<=#"&Zzdate_fin&"#"
Set modif2= Server.CreateObject("ADODB.RecordSet")
modif2.open SQLmodif2,conn2

response.redirect("gene_planning_sos.asp")
%>
