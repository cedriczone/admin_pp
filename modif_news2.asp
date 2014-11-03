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

Ztexte=request.form("texte")

Ztexte=tarea(Ztexte)

SQLmodif="UPDATE [defilantes] set texte_defilante='"&Ztexte&"' WHERE id_defilante="&Zid
Set modif= Server.CreateObject("ADODB.RecordSet")
modif.open SQLmodif,conn

response.redirect("news.asp")
%>