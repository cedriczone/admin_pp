<!--#include file="verif_ident.asp"-->
<!--#include file="connexion_perm.asp"-->
<%
Zid=request.form("id")
Ztitre=Request.form("titre")
Ztitre=replace(Ztitre,"'","''")
Ztexte=Request.Form("message")
Ztexte=replace(Ztexte,"'","''")

SQLmodif="UPDATE [messagerie] set titre='"&Ztitre&"',texte='"&Ztexte&"' WHERE id_message="&Zid
Set modif= Server.CreateObject("ADODB.RecordSet")
modif.open SQLmodif,conn

response.redirect("news.asp?m=2")
%>