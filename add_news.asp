<!--#include file="connexion2.asp"-->
<%
Ztexte=request.form("texte")
response.write(Ztexte)
function tarea(text)
         tarea=replace(text,"&","&amp;")
         tarea=replace(tarea,"<","&lt;")
		 tarea=replace(tarea,">","&gt;")
		 tarea=replace(tarea,"'","''")
end function

Ztexte=tarea(Ztexte)

SQLadd="Insert Into [defilantes](texte_defilante) Values('"&Ztexte&"')"
Set saisie= Server.CreateObject("ADODB.RecordSet")
saisie.open SQLadd,conn2

conn2.close
Set conn2=nothing

response.redirect("news.asp?m=1")
%>