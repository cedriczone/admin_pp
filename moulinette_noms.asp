<!--#include file="verif_ident.asp"-->
<!--#include file="connexion.asp"-->
<!--#include file="connexion2.asp"-->
<%
SQLmoulinette="SELECT * from [avocats_pp] order by avo_code"
Set rsmoulinette=server.Createobject("adodb.recordset")
rsmoulinette.open SQLmoulinette,conn,3,3

rsmoulinette.movefirst
do while not rsmoulinette.eof
le_nom = replace(rsmoulinette("avo_libelle"),"'","")
SQLmodif="UPDATE [dispos_gav] set avo_nom='"&le_nom&"' WHERE num_avo="&rsmoulinette("avo_code")
Set modif= Server.CreateObject("ADODB.RecordSet")
modif.open SQLmodif,conn2

rsmoulinette.movenext
loop
response.write("termin&eacute;")
%>