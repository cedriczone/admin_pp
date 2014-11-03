<!--#include file="verif_ident.asp"-->
<!--#include file="connexion_perm.asp"-->
<%
Zid=request.form("id")
Zid=cint(Zid)
SQLinfo="SELECT * from [infojour] where id_infojour="&Zid
Set rsinfo=server.Createobject("adodb.recordset")
rsinfo.open SQLinfo,conn,3,3
rsinfo.movefirst
function tarea(text)
         tarea=replace(text,"&amp;","&")
         tarea=replace(tarea,"&lt;","<")
		 tarea=replace(tarea,"&gt;",">")
		 tarea=replace(tarea,"<br>",VbCrLf)
		 tarea=replace(tarea,"''","'")
end function
texte=tarea(rsinfo("texte_info"))
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<link href="css/main.css" rel="stylesheet" type="text/css" />
</head>
<body>
<div id="header"><p class="titre_principal">Administration</p></div>
<div>
<div>
 <div>
   <p><strong>Modifier une info pour une journ&eacute;e du planning perm classique</strong></p><form id="form2" name="form2" method="post" action="modif_infojour2.asp?id=<%=Zid%>">
  <p>Date (jj/mm/aaaa) : 
    <input name="date_info" type="text" id="date_info" size="12" value="<%=rsinfo("date_info")%>"/>
  </p>
  <p>Info :
    <textarea name="info" id="info" cols="60" rows="5"><%=texte%></textarea>
    <br />
    <br />
    <input type="submit" name="button" id="button" value="Modifier" />
  </p>
  </form>
   </div>
</div>
</div>
<div id="footer">&nbsp;</div>
</body>
</html>
