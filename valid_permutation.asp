<!--#include file="verif_ident.asp"-->
<!--#include file="connexion2.asp"-->
<!--#include file="connexion.asp"-->
<%
Zid=request.querystring("id")

SQLpermute="SELECT * FROM [permutation] where id_permutation="&Zid
Set rspermute=server.Createobject("adodb.recordset")
rspermute.open SQLpermute,conn2,3,3
rspermute.movefirst

Ztype=rspermute("type_planning")

Select Case Ztype
Case "PP","ETR","MIN"
'rien

Case "SOS"
'REMPLACE
Zdate=month(rspermute("date_perm"))&"/"&day(rspermute("date_perm"))&"/"&year(rspermute("date_perm"))
SQLsos="SELECT * FROM [planning_sos] where jour=#"&Zdate&"#"
Set rssos=server.Createobject("adodb.recordset")
rssos.open SQLsos,conn2,3,3
rssos.movefirst
if rssos("titulaire")=rspermute("num_avocat") then
SQLmodif="UPDATE [planning_sos] set titulaire="&rspermute("num_permute")&" WHERE num_planning="&rssos("num_planning")
elseif rssos("suppleant")=rspermute("num_avocat") then
SQLmodif="UPDATE [planning_sos] set suppleant="&rspermute("num_permute")&" WHERE num_planning="&rssos("num_planning")
end if

Set modif= Server.CreateObject("ADODB.RecordSet")
modif.open SQLmodif,conn2

'REMPLACANT

Zdate2=month(rspermute("date_permute"))&"/"&day(rspermute("date_permute"))&"/"&year(rspermute("date_permute"))
SQLsos2="SELECT * FROM [planning_sos] where jour=#"&Zdate2&"#"
Set rssos2=server.Createobject("adodb.recordset")
rssos2.open SQLsos2,conn2,3,3
rssos2.movefirst
if rssos2("titulaire")=rspermute("num_permute") then
SQLmodif2="UPDATE [planning_sos] set titulaire="&rspermute("num_avocat")&" WHERE num_planning="&rssos2("num_planning")
elseif rssos2("suppleant")=rspermute("num_permute") then
SQLmodif2="UPDATE [planning_sos] set suppleant="&rspermute("num_avocat")&" WHERE num_planning="&rssos2("num_planning")
end if

Set modif2= Server.CreateObject("ADODB.RecordSet")
modif2.open SQLmodif2,conn2

Case "GAV"
'REMPLACE
Zdate=month(rspermute("date_perm"))&"/"&day(rspermute("date_perm"))&"/"&year(rspermute("date_perm"))
SQLgav="SELECT * FROM [planning_gav] where date_gav=#"&Zdate&"#"
Set rsgav=server.Createobject("adodb.recordset")
rsgav.open SQLgav,conn2,3,3
rsgav.movefirst

if rsgav("num_inter_chev")=rspermute("num_avocat") then
SQLmodif="UPDATE [planning_gav] set num_inter_chev="&rspermute("num_permute")&" WHERE num_ligne="&rsgav("num_ligne")
elseif rsgav("num_inter1")=rspermute("num_avocat") then
SQLmodif="UPDATE [planning_gav] set num_inter1="&rspermute("num_permute")&" WHERE num_ligne="&rsgav("num_ligne")
elseif rsgav("num_inter2")=rspermute("num_avocat") then
SQLmodif="UPDATE [planning_gav] set num_inter2="&rspermute("num_permute")&" WHERE num_ligne="&rsgav("num_ligne")
elseif rsgav("num_inter3")=rspermute("num_avocat") then
SQLmodif="UPDATE [planning_gav] set num_inter3="&rspermute("num_permute")&" WHERE num_ligne="&rsgav("num_ligne")
elseif rsgav("num_inter4")=rspermute("num_avocat") then
SQLmodif="UPDATE [planning_gav] set num_inter4="&rspermute("num_permute")&" WHERE num_ligne="&rsgav("num_ligne")
elseif rsgav("num_inter5")=rspermute("num_avocat") then
SQLmodif="UPDATE [planning_gav] set num_inter5="&rspermute("num_permute")&" WHERE num_ligne="&rsgav("num_ligne")
elseif rsgav("num_inter6")=rspermute("num_avocat") then
SQLmodif="UPDATE [planning_gav] set num_inter6="&rspermute("num_permute")&" WHERE num_ligne="&rsgav("num_ligne")
end if
response.write(SQLmodif)
Set modif= Server.CreateObject("ADODB.RecordSet")
modif.open SQLmodif,conn2

'REMPLACANT
Zdate2=month(rspermute("date_permute"))&"/"&day(rspermute("date_permute"))&"/"&year(rspermute("date_permute"))
SQLgav2="SELECT * FROM [planning_gav] where date_gav=#"&Zdate2&"#"
Set rsgav2=server.Createobject("adodb.recordset")
rsgav2.open SQLgav2,conn2,3,3
rsgav2.movefirst

if rsgav2("num_inter_chev")=rspermute("num_permute") then
SQLmodif2="UPDATE [planning_gav] set num_inter_chev="&rspermute("num_avocat")&" WHERE num_ligne="&rsgav2("num_ligne")
elseif rsgav2("num_inter1")=rspermute("num_permute") then
SQLmodif2="UPDATE [planning_gav] set num_inter1="&rspermute("num_avocat")&" WHERE num_ligne="&rsgav2("num_ligne")
elseif rsgav2("num_inter2")=rspermute("num_permute") then
SQLmodif2="UPDATE [planning_gav] set num_inter2="&rspermute("num_avocat")&" WHERE num_ligne="&rsgav2("num_ligne")
elseif rsgav2("num_inter3")=rspermute("num_permute") then
SQLmodif2="UPDATE [planning_gav] set num_inter3="&rspermute("num_avocat")&" WHERE num_ligne="&rsgav2("num_ligne")
elseif rsgav2("num_inter4")=rspermute("num_permute") then
SQLmodif2="UPDATE [planning_gav] set num_inter4="&rspermute("num_avocat")&" WHERE num_ligne="&rsgav2("num_ligne")
elseif rsgav2("num_inter5")=rspermute("num_permute") then
SQLmodif2="UPDATE [planning_gav] set num_inter5="&rspermute("num_avocat")&" WHERE num_ligne="&rsgav2("num_ligne")
elseif rsgav2("num_inter6")=rspermute("num_permute") then
SQLmodif2="UPDATE [planning_gav] set num_inter6="&rspermute("num_avocat")&" WHERE num_ligne="&rsgav2("num_ligne")
end if

Set modif2= Server.CreateObject("ADODB.RecordSet")
modif2.open SQLmodif2,conn2

End Select

SQLmodif0="UPDATE [permutation] set validee=1 WHERE id_permutation="&Zid
Set modif0= Server.CreateObject("ADODB.RecordSet")
modif0.open SQLmodif0,conn2
response.redirect("permutations.asp")
%>