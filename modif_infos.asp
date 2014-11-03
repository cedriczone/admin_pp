<!--#include file="verif_ident.asp"-->
<!--#include file="connexion_perm.asp"-->
<%
Zid=request.Querystring("id")
SQLinfos="SELECT * from [messagerie] WHERE id_message="&Zid
Set rsinfos=server.Createobject("adodb.recordset")
rsinfos.open SQLinfos,conn,3,3

function tarea(ch)
   tarea = replace(ch,"<br />",VbCrLf)
   tarea = replace(ch,"&","&amp;")
   tarea = replace(tarea,"<","&lt;")
   tarea = replace(tarea,"""","&quote;")
end function
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Messagerie</title>
<link href="css/main.css" rel="stylesheet" type="text/css" />
</head>
<body>
<div id="header"><p class="titre_principal">Administration</p></div>
<!--#include file="menu.asp"-->
<div id="main">
<!--FORMULAIRE AJOUT infos -->
	<form action="modif_infos2.asp" method="post" name="form_ajout_infos" id="form_ajout_infos">          
		<fieldset>
		  <p>
			<input type="hidden" id="id" name="id" value="<%=Zid%>" />
			<label for="titre">Titre</label>
			<input type="text" name="titre" id="titre" tabindex="30" value="<%=rsinfos("titre")%>" />
			</p>
			<p>
			<label for="message">Message</label>
			<textarea name="message" id="message" cols="45" rows="5" tabindex="40"><%=tarea(rsinfos("texte"))%></textarea>
			</p>
		</fieldset>		
		<p><input type="submit" name="btn_envoi" id="btn_envoi" value="Modifier" tabindex="150" /></p>
	</form>
</div>
<div id="footer">&nbsp;</div>
</body>
</html>
