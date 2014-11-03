<!--#include file="verif_ident.asp"-->
<!--#include file="connexion2.asp"-->
<%
Zcoord=request.form("coord")
Zmois=request.form("mois")
Zannee=request.form("annee")

SQLexist="SELECT * from [dispos_coord_gav] where avo_code="&Zcoord&" and mois_dispo="&Zmois&" and annee_dispo="&Zannee
Set rsexist=server.Createobject("adodb.recordset")
rsexist.open SQLexist,conn2,3,3
nbre_exist=rsexist.recordcount

if nbre_exist>0 then
response.redirect("dispos_gav_coord.asp?m=1")
else

SQLajoutdispo="Insert Into [dispos_coord_gav](avo_code,mois_dispo,annee_dispo) Values("&Zcoord&","&Zmois&","&Zannee&")"
Set saisiedispo= Server.CreateObject("ADODB.RecordSet")
saisiedispo.open SQLajoutdispo,conn2

response.redirect("dispos_gav_coord.asp?m=2")

end if
%>