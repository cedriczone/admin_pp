<!--#include file="verif_ident.asp"-->
<!--#include file="connexion_perm.asp"-->
<%
Zdate=date()
Zdate=month(Zdate)&"/"&day(Zdate)&"/"&year(Zdate)
Ztitre=Request.form("titre")
Ztitre=replace(Ztitre,"'","''")
Ztexte=Request.Form("message")
Ztexte=replace(Ztexte,"'","''")

SQLaddnews="Insert Into [messagerie](date_message,titre,texte) Values(#"&Zdate&"#,'"&Ztitre&"','"&Ztexte&"')"
Set addnews= Server.CreateObject("ADODB.RecordSet")
addnews.open SQLaddnews,conn

response.redirect("news.asp?m=1")
%>