<!--#include file="verif_ident.asp"-->
<!--#include file="connexion2.asp"-->
<!--#include file="connexion.asp"-->
<%
Zid=request.querystring("id")

SQLremplace="SELECT * FROM [remplacement] where id_remplace="&Zid
Set rsremplace=server.Createobject("adodb.recordset")
rsremplace.open SQLremplace,conn2,3,3
rsremplace.movefirst

Ztype=rsremplace("type_planning")

Select Case Ztype
Case "PP","ETR","MIN"
'rien

Case "SOS"
Zdate=month(rsremplace("date_remplace"))&"/"&day(rsremplace("date_remplace"))&"/"&year(rsremplace("date_remplace"))
SQLsos="SELECT * FROM [planning_sos] where jour=#"&Zdate&"#"
Set rssos=server.Createobject("adodb.recordset")
rssos.open SQLsos,conn2,3,3
rssos.movefirst
if rssos("titulaire")=rsremplace("remplace") then
SQLmodif="UPDATE [planning_sos] set titulaire="&rsremplace("remplacant")&" WHERE num_planning="&rssos("num_planning")
elseif rssos("suppleant")=rsremplace("remplace") then
SQLmodif="UPDATE [planning_sos] set suppleant="&sremplace("remplacant")&" WHERE num_planning="&rssos("num_planning")
end if

Set modif= Server.CreateObject("ADODB.RecordSet")
modif.open SQLmodif,conn2

Case "GAV"
Zdate=month(rsremplace("date_remplace"))&"/"&day(rsremplace("date_remplace"))&"/"&year(rsremplace("date_remplace"))
SQLgav="SELECT * FROM [planning_gav] where date_gav=#"&Zdate&"#"
response.Write(SQLgav)
Set rsgav=server.Createobject("adodb.recordset")
rsgav.open SQLgav,conn2,3,3
rsgav.movefirst
if rsgav("num_inter_chev")=rsremplace("remplace") then
SQLmodif="UPDATE [planning_gav] set num_inter_chev="&rsremplace("remplacant")&" WHERE num_ligne="&rsgav("num_ligne")
elseif rsgav("num_inter1")=rsremplace("remplace") then
SQLmodif="UPDATE [planning_gav] set num_inter1="&rsremplace("remplacant")&" WHERE num_ligne="&rsgav("num_ligne")
elseif rsgav("num_inter2")=rsremplace("remplace") then
SQLmodif="UPDATE [planning_gav] set num_inter2="&rsremplace("remplacant")&" WHERE num_ligne="&rsgav("num_ligne")
elseif rsgav("num_inter3")=rsremplace("remplace") then
SQLmodif="UPDATE [planning_gav] set num_inter3="&rsremplace("remplacant")&" WHERE num_ligne="&rsgav("num_ligne")
elseif rsgav("num_inter4")=rsremplace("remplace") then
SQLmodif="UPDATE [planning_gav] set num_inter4="&rsremplace("remplacant")&" WHERE num_ligne="&rsgav("num_ligne")
elseif rsgav("num_inter5")=rsremplace("remplace") then
SQLmodif="UPDATE [planning_gav] set num_inter5="&rsremplace("remplacant")&" WHERE num_ligne="&rsgav("num_ligne")
elseif rsgav("num_inter6")=rsremplace("remplace") then
SQLmodif="UPDATE [planning_gav] set num_inter6="&rsremplace("remplacant")&" WHERE num_ligne="&rsgav("num_ligne")
end if

Set modif= Server.CreateObject("ADODB.RecordSet")
modif.open SQLmodif,conn2

End Select

SQLmodif0="UPDATE [remplacement] set validee=1 WHERE id_remplace="&Zid
Set modif0= Server.CreateObject("ADODB.RecordSet")
modif0.open SQLmodif0,conn2
response.redirect("remplacements.asp")
%>