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

Zdate=request.form("date_info")
Zzdate=month(Zdate)&"/"&day(Zdate)&"/"&year(Zdate)
Ztexte=request.form("info")

Ztexte=tarea(Ztexte)

SQLmodif="UPDATE [infojour] set date_info=#"&Zzdate&"#,texte_info='"&Ztexte&"' WHERE id_infojour="&Zid
Set modif= Server.CreateObject("ADODB.RecordSet")
modif.open SQLmodif,conn

response.redirect("infojour.asp")
%>