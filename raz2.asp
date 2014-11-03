<!--#include file="verif_ident.asp"-->
<!--#include file="connexion2.asp"-->
<%
SQLmodif="UPDATE [dispos_gav] set cpt=0"
Set modif= Server.CreateObject("ADODB.RecordSet")
modif.open SQLmodif,conn2
%>
