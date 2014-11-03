<!--#include file="connexion.asp"-->
<!--#include file="connexion2.asp"-->
<%
Set suppr= Server.CreateObject("ADODB.RecordSet")
SQLsuppr="DELETE * FROM [planning_sos]"
suppr.open SQLsuppr,conn2

Znum=request.form("num")
Zjour=request.form("jour")

if Znum="" then Znum = 1
if Zjour="" then Zjour=date()

SQLavocats="SELECT * from [Intervenants_SOS] order by avo_nom"
Set rsavocats=server.Createobject("adodb.recordset")
rsavocats.open SQLavocats,conn,3,3
nbre_avocats=rsavocats.recordcount
rsavocats.movefirst

if Znum<>1 then
do while not rsavocats("avo_code")=cint(Znum)
rsavocats.movenext
loop
end if

for i=1 to 365
lejour=Weekday(Zjour,7)

Ztitulaire=rsavocats("avo_code")
rsavocats.movenext
if rsavocats.eof then
rsavocats.movefirst
Zsuppleant=rsavocats("avo_code")
rsavocats.movelast
else
Zsuppleant=rsavocats("avo_code")
rsavocats.moveprevious
end if

jour_d=DatePart("d", Zjour)
mois_d=DatePart("m", Zjour)
annee_d=DatePart("yyyy", Zjour)
Zjour2=mois_d&"/"&jour_d&"/"&annee_d

SQLaddjour="Insert Into [planning_sos](jour,titulaire,suppleant) Values(#"&Zjour2&"#,"&Ztitulaire&","&Zsuppleant&")"
Set saisie= Server.CreateObject("ADODB.RecordSet")
saisie.open SQLaddjour,conn2

if lejour=3 or lejour=5 or lejour=7 then
rsavocats.movenext
if rsavocats.eof then rsavocats.movefirst
end if

Zjour=dateAdd("d",1,Zjour)
lejour=lejour+1
i=i+1

next

response.redirect("gene_planning_sos.asp?m=1")
%>