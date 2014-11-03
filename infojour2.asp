<!--#include file="connexion2.asp"-->
<%
Zinfo=request.form("info")
response.write(Zinfo)
function tarea(text)
         tarea=replace(text,"&","&amp;")
         tarea=replace(tarea,"<","&lt;")
		 tarea=replace(tarea,">","&gt;")
		 tarea=replace(tarea,"'","''")
end function

Zinfo=tarea(Zinfo)

Zdate_info=request.form("date_info")
ZZdate_info=month(Zdate_info)&"/"&day(Zdate_info)&"/"&year(Zdate_info)

SQLadd="Insert Into [infojour](date_info,texte_info) Values(#"&ZZdate_info&"#,'"&Zinfo&"')"
Set saisie= Server.CreateObject("ADODB.RecordSet")
saisie.open SQLadd,conn2

conn2.close
Set conn2=nothing

response.redirect("infojour.asp")
%>