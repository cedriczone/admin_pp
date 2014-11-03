<!--#include file="connexion_perm.asp"-->
<%
Znbre_box=request.form("nbre_box")
if Znbre_box<1 then response.redirect("gestion_login.asp")
Dim Zbox(9999)
for i=1 to Znbre_box
Zbox(i)=request.form("checkbox"&i)
if Zbox(i)<>"" and Zbox(i)>0 then
Zid=cint(Zbox(i))
SQLsuppr="DELETE * from [login] where id_login="&Zid
Set suppr= Server.CreateObject("ADODB.RecordSet")
suppr.open SQLsuppr,conn
end if
next
response.redirect("gestion_login.asp")
%>