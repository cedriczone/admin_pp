<!--#include file="verif_ident.asp"-->
<!--#include file="connexion_perm.asp"-->
<%
Zdebut=request.form("debut")
Zfin=request.form("fin")
if Zdebut="" or Zfin="" then response.Redirect("vacances_jud.asp?m=1")

delta=dateDiff("d",Zdebut,Zfin)
if delta<1 then response.Redirect("vacances_jud.asp?m=2")

SQLajoutplage="Insert Into [vacances](debut,fin) Values(#"&Zdebut&"#,#"&Zfin&"#)"
Set saisieplage= Server.CreateObject("ADODB.RecordSet")
saisieplage.open SQLajoutplage,conn

response.redirect("vacances_jud.asp?m=3")
%>