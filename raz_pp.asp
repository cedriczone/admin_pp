<!--#include file="verif_ident.asp"-->
<!--#include file="connexion2.asp"-->
<%
SQLmodif="UPDATE [cpt_pp] set cpt=0, type=''"
Set modif= Server.CreateObject("ADODB.RecordSet")
modif.open SQLmodif,conn2
%>
