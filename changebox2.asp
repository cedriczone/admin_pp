<!--#include file="connexion2.asp"-->
<%
Zzid=request.querystring("valeur")
Zaction=request.QueryString("action")
if Zzid<>"" then
lng=len(Zzid)-6
Zid=right(Zzid,lng)
Zmois=left(Zzid,2)
Zannee=mid(Zzid,3,4)

if Zaction=1 then

SQLadd="Insert Into [dispos_coord_gav](avo_code,mois_dispo,annee_dispo) Values("&Zid&","&Zmois&","&Zannee&")"
Set saisie= Server.CreateObject("ADODB.RecordSet")
saisie.open SQLadd,conn2

elseif Zaction=0 then

SQLsuppr="DELETE * from [dispos_coord_gav] where avo_code="&Zid&" and mois_dispo="&Zmois&" and annee_dispo="&Zannee
Set suppr= Server.CreateObject("ADODB.RecordSet")
suppr.open SQLsuppr,conn2

end if
end if
%>