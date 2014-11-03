<!--#include file="verif_ident.asp"-->
<!--#include file="connexion_perm.asp"-->
<%
Zid=request.querystring("id")
function tarea(text)
         tarea=replace(text,"&","&amp;")
         tarea=replace(tarea,"<","&lt;")
		 tarea=replace(tarea,">","&gt;")
		 tarea=replace(tarea,VbCrLf,"<br>")
		 tarea=replace(tarea,"'","''")
end function

Ztitre=request.form("titre")
Ztexte=request.form("texte")

Ztitre=tarea(Ztitre)
Ztexte=tarea(Ztexte)

SQLmodif="UPDATE [news] set titre='"&Ztitre&"',texte='"&Ztexte&"' WHERE id_news="&Zid
Set modif= Server.CreateObject("ADODB.RecordSet")
modif.open SQLmodif,conn

response.redirect("infos.asp")
%>