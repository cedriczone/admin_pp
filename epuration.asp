<!--#include file="connexion.asp"-->
<!--#include file="connexion2.asp"-->
<%
'VERIF DISPOS
'/////////////////////
SQLverifdispos="SELECT * from [dispos_gav] order by id_dispo"
Set rsverifdispos=server.Createobject("adodb.recordset")
rsverifdispos.open SQLverifdispos,conn2,3,3
nbre_verifdispos=rsverifdispos.recordcount
if nbre_verifdispos>0 then
rsverifdispos.movefirst

do while not rsverifdispos.eof

SQLliste="SELECT * from [Intervenants_GAV] where avo_code="&rsverifdispos("num_avo")
Set rsliste=server.Createobject("adodb.recordset")
rsliste.open SQLliste,conn,3,3
nbre_liste=rsliste.recordcount

if nbre_liste<1 then

SQLsuppr="DELETE * from [dispos_gav] where num_avo="&rsverifdispos("num_avo")
Set suppr= Server.CreateObject("ADODB.RecordSet")
suppr.open SQLsuppr,conn2

end if

rsverifdispos.movenext
loop
end if
%>