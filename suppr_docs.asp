<!--#include file="connexion_perm.asp"-->
<%
On Error resume next

Znbre_box=request.form("nbre_box")
if Znbre_box<1 then response.redirect("documents.asp")
Dim Zbox(999)
for i=1 to Znbre_box
Zbox(i)=request.form("checkbox"&i)
if Zbox(i)<>"" and Zbox(i)>0 then
Zid=cint(Zbox(i))
SQLsuppr="DELETE * from [docs] where id_doc="&Zid
Set suppr= Server.CreateObject("ADODB.RecordSet")
suppr.open SQLsuppr,conn
end if
next
response.redirect("documents.asp")
%>