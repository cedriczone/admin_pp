<!--#include file="verif_ident.asp"-->
<!--#include file="connexion2.asp"-->
<%
datedujour=month(now())&"/"&day(now())&"/"&year(now())
SQLmodif="UPDATE [repartition_gav] set archive=1,date_archive=#"&datedujour&"#,heure_archive=#"&time()&"# WHERE archive=0"
Set modif= Server.CreateObject("ADODB.RecordSet")
modif.open SQLmodif,conn2

response.redirect("raz.asp")
%>