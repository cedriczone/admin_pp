<!--#include file="verif_ident.asp"-->
<!--#include file="connexion.asp"-->
<!--#include file="test_connexion2.asp"-->
<%
Zmois=request.form("mois")
if cint(Zmois)<month(date()) then
Zannee=year(date())+1
else
Zannee=year(date())
end if
if Zmois<month(date()) then Zannee=Zannee+1

'nombre de jours du mois
d1 = request.form("date1")
d2 = request.form("date2")
joursmois = datediff("d",d1,d2)
premier_jour = weekday(d1,2)
if premier_jour<>2 and premier_jour<>4 and premier_jour<>6 then response.Redirect("gene_planning_gav.asp?err="&premier_jour)

dim coord(99)
coord(1)=request.form("coord1")
coord(2)=request.form("coord2")
coord(3)=request.form("coord3")

'//////////////////////
'//////////////////////
'FONCTION DE VERIFICATION DES DISPOS
function selection(avocat,periode)
if avocat>0 and periode<>"" then
SQLdispo="SELECT * from [dispos_gav] where num_avo="&avocat&" and "&periode&"=1"
Set rsdispo=server.Createobject("adodb.recordset")
rsdispo.open SQLdispo,conn2,3,3
nbre_dispo=rsdispo.recordcount
if nbre_dispo=1 then
selection=1
else
selection=0
end if
else
selection=0
end if
end function

'//////////////////////
'//////////////////////

'//////////////////////////////////////////////////////
'ON SELECTIONNE TOUS LES AVOCATS  DE LA GAV
SQLliste="SELECT * from [Intervenants_GAV] order by avo_libelle"
Set rsliste=server.Createobject("adodb.recordset")
rsliste.open SQLliste,conn,3,3
nbre_avo=rsliste.recordcount

'on verifie on en est dans la liste alphabetique
SQLverifdernier="SELECT * from [params_gav]"
Set rsverifdernier=server.Createobject("adodb.recordset")
rsverifdernier.open SQLverifdernier,conn2,3,3
nbre_dernier=rsverifdernier.recordcount
'si il n'y a rien on part du debut
rsliste.movefirst
'sinon on se place a la suite du dernier avocat tiré
if nbre_dernier>0 then
rsverifdernier.movefirst
do while not rsliste.eof
if rsliste("avo_code")=rsverifdernier("dernier") then exit do
rsliste.movenext
loop
end if

'/////////////////////////////////////
'CHOIX DES INTERVENANTS DE LA JOURNEE
'/////////////////////////////////////

j=0
k=0
l=0

for i=1 to joursmois+1
datecours=dateadd("d",k,d1)
k=k+1
l=l+1
'le jour en cours dans la boucle
dateencours=month(datecours)&"/"&day(datecours)&"/"&year(datecours)
jourencours=Weekday(datecours,2)
jourencours2=Weekday(datecours,2)

Select case jourencours
case 2,4,6
j=j+1
if j=4 then j=1
if l>=7 then
j=j+1
l=0
end if
if j=4 then j=1
end select

if j=4 then j=1
Zcoord=coord(j)

Select case jourencours2
case 6,7
we=1
case else
we=0
end select

'on vérifie s'il y en a dans la table temp

SQLtemp="SELECT * from [temp] where avo_code<>"&Zcoord&" order by avo_nom"
Set rstemp=server.Createobject("adodb.recordset")
rstemp.open SQLtemp,conn2,3,3
nbre_temp=rstemp.recordcount

t=0

'si la table temp est remplie
if nbre_temp>0 then

'on verifie au fur et a mesure les dispos
do while not rstemp.eof

if we=0 then
resultat=selection(rstemp("avo_code"),"sej")
elseif we=1 then
resultat=selection(rstemp("avo_code"),"wej")
end if

if resultat=1 then
t=t+1
'on supprime l'intervenant de la base temp
SQLsuppr="DELETE * from [temp] where avo_code="&rstemp("avo_code")
Set suppr= Server.CreateObject("ADODB.RecordSet")
suppr.open SQLsuppr,conn2

if t=1 then
response.write("temp - intermatin1 : "&rstemp("avo_nom")&"<br />")
intermatin1=rstemp("avo_code")
elseif t=2 then
response.write("temp - intermatin2 : "&rstemp("avo_nom")&"<br />")
intermatin2=rstemp("avo_code")
end if

end if

rstemp.movenext
if t=2 then exit do

loop

end if

'si on a pas trouvé dans la base temp alors on continue la liste des avocats
if t<2 then
do while not rsliste.eof
if rsliste("avo_code")<>Zcoord then
if we=0 then
resultat=selection(rsliste("avo_code"),"sej")
elseif we=1 then
resultat=selection(rsliste("avo_code"),"wej")
end if

if resultat=1 then
t=t+1

if t=1 then
response.write("liste - intermatin1 : "&rsliste("avo_libelle")&"<br />")
intermatin1=rsliste("avo_code")
elseif t=2 then
response.write("liste - intermatin2 : "&rsliste("avo_libelle")&"<br />")
intermatin2=rsliste("avo_code")
end if
else
's'il n'est pas dispo on l'ajoute dans la base temp
'on verifie qu'il n'y soit pas déjà
SQLexistetemp="SELECT * from [temp] where avo_code="&rsliste("avo_code")
Set rsexistetemp=server.Createobject("adodb.recordset")
rsexistetemp.open SQLexistetemp,conn2,3,3
nbre_existetemp=rsexistetemp.recordcount

if nbre_existetemp=0 then
Znom=replace(rsliste("avo_libelle"),"'","''")
SQLadd_temp="Insert Into [temp](avo_code,avo_nom) Values("&rsliste("avo_code")&",'"&Znom&"')"
Set saisie_temp= Server.CreateObject("ADODB.RecordSet")
saisie_temp.open SQLadd_temp,conn2
end if
end if
end if
rsliste.movenext
if rsliste.eof then rsliste.movefirst
if t=2 then exit do

loop

end if

'////////////////////////////////////////////
'INTERVENANTS SOIR
'////////////////////////////////////////////
'on vérifie s'il y en a dans la table temp

t=0

'si la table temp est remplie
SQLtemp="SELECT * from [temp] where avo_code<>"&Zcoord&" order by avo_nom"
Set rstemp=server.Createobject("adodb.recordset")
rstemp.open SQLtemp,conn2,3,3
nbre_temp=rstemp.recordcount
if nbre_temp>0 then
rstemp.movefirst

'on verifie au fur et a mesure les dispos
do while not rstemp.eof

if we=0 then
resultat=selection(rstemp("avo_code"),"sen")
elseif we=1 then
resultat=selection(rstemp("avo_code"),"wen")
end if

if resultat=1 then
t=t+1
'on supprime l'intervenant de la base temp
SQLsuppr="DELETE * from [temp] where avo_code="&rstemp("avo_code")
Set suppr= Server.CreateObject("ADODB.RecordSet")
suppr.open SQLsuppr,conn2

if t=1 then
response.write("temp - intersoir1 : "&rstemp("avo_nom")&"<br />")
intersoir1=rstemp("avo_code")
elseif t=2 then
response.write("temp - intersoir2 : "&rstemp("avo_nom")&"<br />")
intersoir2=rstemp("avo_code")
end if

end if

rstemp.movenext
if t=2 then exit do

loop

end if

'si on a pas trouvé dans la base temp alors on continue la liste des avocats
if t<2 then
do while not rsliste.eof
if rsliste("avo_code")<>Zcoord then
if we=0 then
resultat=selection(rsliste("avo_code"),"sen")
elseif we=1 then
resultat=selection(rsliste("avo_code"),"wen")
end if

if resultat=1 then
t=t+1

if t=1 then
response.write("liste - intersoir1 : "&rsliste("avo_libelle")&"<br />")
intersoir1=rsliste("avo_code")
elseif t=2 then
response.write("liste - intersoir2 : "&rsliste("avo_libelle")&"<br />")
intersoir2=rsliste("avo_code")
end if

else
's'il n'est pas dispo on l'ajoute dans la base temp
'on verifie qu'il n'y soit pas déjà
SQLexistetemp="SELECT * from [temp] where avo_code="&rsliste("avo_code")
Set rsexistetemp=server.Createobject("adodb.recordset")
rsexistetemp.open SQLexistetemp,conn2,3,3
nbre_existetemp=rsexistetemp.recordcount

if nbre_existetemp=0 then
Znom=replace(rsliste("avo_libelle"),"'","''")
SQLadd_temp="Insert Into [temp](avo_code,avo_nom) Values("&rsliste("avo_code")&",'"&Znom&"')"
Set saisie_temp= Server.CreateObject("ADODB.RecordSet")
saisie_temp.open SQLadd_temp,conn2
end if
end if
end if
rsliste.movenext
if rsliste.eof then rsliste.movefirst
if t=2 then exit do

loop
end if

'///////////////////////////////////////
'INSERTION DANS LA BASE DE DONNEES
'///////////////////////////////////////
'enregistrement du dernier avocat selectionné
SQLverifdernier="SELECT * from [params_gav]"
Set rsverifdernier=server.Createobject("adodb.recordset")
rsverifdernier.open SQLverifdernier,conn2,3,3
nbre_dernier=rsverifdernier.recordcount

if nbre_dernier=0 then
SQLdernier="Insert Into [params_gav](dernier) Values("&rsliste("avo_code")&")"
else
SQLdernier="UPDATE [params_gav] set dernier="&rsliste("avo_code")&" WHERE id_param=1"
end if

Set modifdernier= Server.CreateObject("ADODB.RecordSet")
modifdernier.open SQLdernier,conn2

SQLaddgav="Insert Into [planning_gav](date_gav,num_coordinateur,num_inter_jour1,num_inter_jour2,num_inter_nuit1,num_inter_nuit2) Values(#"&dateencours&"#,"&Zcoord&","&intermatin1&","&intermatin2&","&intersoir1&","&intersoir2&")"
response.write(SQLaddgav&"<br />")
Set saisiegav= Server.CreateObject("ADODB.RecordSet")
saisiegav.open SQLaddgav,conn2

'///////////////////////////////////////////////////
'ON INCREMENTE POUR CHACUN LEUR COMPTEUR DE TIRAGE
'///////////////////////////////////////////////////
' On vérifie si la personne a déja eu des tirages sinon on la créee 
' Pour le coordinateur:
SQLrepartcoord="SELECT * FROM [repartition_coord_gav] where avo_code="&Zcoord&" and archive=0"
Set rsrepartcoord=server.Createobject("adodb.recordset")
rsrepartcoord.open SQLrepartcoord,conn2,3,3
nbre_repartcoord=rsrepartcoord.recordcount

if nbre_repartcoord>0 then
incremente0=rsrepartcoord("nbre_coord_gav")+1
SQLmodifcoord="UPDATE [repartition_coord_gav] set nbre_coord_gav="&incremente0&" WHERE avo_code="&Zcoord
else
SQLmodifcoord="Insert Into [repartition_coord_gav](avo_code,nbre_coord_gav,archive) Values("&Zcoord&",1,0)"
end if

Set modifcoord= Server.CreateObject("ADODB.RecordSet")
modifcoord.open SQLmodifcoord,conn2

' pour les intervenants
' inter jour1

SQLrepartjour1="SELECT * FROM [repartition_gav] where avo_code="&intermatin1&" and archive=0"
Set rsrepartjour1=server.Createobject("adodb.recordset")
rsrepartjour1.open SQLrepartjour1,conn2,3,3
nbre_repartjour1=rsrepartjour1.recordcount

if nbre_repartjour1>0 then
incremente1=rsrepartjour1("nbre_gav")+1
SQLmodifjour1="UPDATE [repartition_gav] set nbre_gav="&incremente1&" WHERE avo_code="&intermatin1
else
SQLmodifjour1="Insert Into [repartition_gav](avo_code,nbre_gav,archive) Values("&intermatin1&",1,0)"
end if
Set modifjour1= Server.CreateObject("ADODB.RecordSet")
modifjour1.open SQLmodifjour1,conn2

' inter jour2

SQLrepartjour2="SELECT * FROM [repartition_gav] where avo_code="&intermatin2&" and archive=0"
Set rsrepartjour2=server.Createobject("adodb.recordset")
rsrepartjour2.open SQLrepartjour2,conn2,3,3
nbre_repartjour2=rsrepartjour2.recordcount

if nbre_repartjour2>0 then
incremente2=rsrepartjour2("nbre_gav")+1
SQLmodifjour2="UPDATE [repartition_gav] set nbre_gav="&incremente2&" WHERE avo_code="&intermatin2
else
SQLmodifjour2="Insert Into [repartition_gav](avo_code,nbre_gav,archive) Values("&intermatin2&",1,0)"
end if
Set modifjour2= Server.CreateObject("ADODB.RecordSet")
modifjour2.open SQLmodifjour2,conn2

' inter nuit 1

SQLrepartjour3="SELECT * FROM [repartition_gav] where avo_code="&intersoir1&" and archive=0"
Set rsrepartjour3=server.Createobject("adodb.recordset")
rsrepartjour3.open SQLrepartjour3,conn2,3,3
nbre_repartjour3=rsrepartjour3.recordcount

if nbre_repartjour3>0 then
incremente3=rsrepartjour3("nbre_gav")+1
SQLmodifjour3="UPDATE [repartition_gav] set nbre_gav="&incremente3&" WHERE avo_code="&intersoir1
else
SQLmodifjour3="Insert Into [repartition_gav](avo_code,nbre_gav,archive) Values("&intersoir1&",1,0)"
end if
Set modifjour3= Server.CreateObject("ADODB.RecordSet")
modifjour3.open SQLmodifjour3,conn2

' inter nuit 2

SQLrepartjour4="SELECT * FROM [repartition_gav] where avo_code="&intersoir2&" and archive=0"
Set rsrepartjour4=server.Createobject("adodb.recordset")
rsrepartjour4.open SQLrepartjour4,conn2,3,3
nbre_repartjour4=rsrepartjour4.recordcount

if nbre_repartjour4>0 then
incremente4=rsrepartjour4("nbre_gav")+1
SQLmodifjour4="UPDATE [repartition_gav] set nbre_gav="&incremente4&" WHERE avo_code="&intersoir2
else
SQLmodifjour4="Insert Into [repartition_gav](avo_code,nbre_gav,archive) Values("&intersoir2&",1,0)"
end if
Set modifjour4= Server.CreateObject("ADODB.RecordSet")
modifjour4.open SQLmodifjour4,conn2

'///////////////////
'FIN DE BOUCLE
'///////////////////
next
'response.redirect("test_gene_planning_gav.asp")
%>